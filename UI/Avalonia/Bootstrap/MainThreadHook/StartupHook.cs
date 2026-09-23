// Runs the PowerShell startup script on the process MAIN thread, inside the
// user's own pwsh.
//
// Why: Cocoa (macOS) only allows the GUI - NSApplication, windows, dialogs - on
// the first thread of the process. A normal pwsh pipeline runs on a worker
// thread (ConsoleHost's ReuseThread runspace), so Avalonia cannot be started
// from an ordinary `pwsh -File`. Windows (STA relaunch) and Linux (X11/Wayland)
// have no such rule; this hook is only engaged on macOS, though it works
// anywhere for testing.
//
// How: .NET calls StartupHook.Initialize() on the main thread BEFORE the app's
// Main when DOTNET_STARTUP_HOOKS names this DLL. If IM_MAIN_THREAD_HOOK=1 is
// also set and pwsh was started with -File, Initialize opens a UseCurrentThread
// runspace right here, runs the script and terminates the process - pwsh's own
// console Main never executes; the process is just the engine + runtime host.
// Without the env var (child processes, ad-hoc pwsh) it returns immediately and
// pwsh starts normally.
//
// Because it executes inside pwsh, the hook needs no PowerShell SDK, no separate
// .NET install and no version matching: whatever pwsh the user has (7.4 or
// newer, on .NET 8/9/10/...) brings its own engine, runtime and $PSHOME/ref
// compile references, all consistent with each other by construction.
using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;

namespace IntuneManagement.MainThreadHook
{
    // Contract used by the PowerShell side (Start-Avalonia.ps1, CoreUIAvalonia.ps1,
    // tests): the type exists only when the hook engaged, and Verify() throws
    // if the caller is not on the thread the script was started on.
    public static class MainThread
    {
        public static int ManagedId { get; internal set; }

        [DllImport("/usr/lib/libSystem.B.dylib")]
        private static extern int pthread_main_np();

        public static bool IsNativeMainThread => !OperatingSystem.IsMacOS() || pthread_main_np() != 0;

        public static void Verify()
        {
            if (Environment.CurrentManagedThreadId != ManagedId || !IsNativeMainThread)
                throw new InvalidOperationException("The macOS GUI must run on the process main thread.");
        }
    }

    // This is a GUI script host, not an interactive console: output is forwarded
    // through the streams, nested prompts are deliberately unsupported, and
    // SetShouldExit captures a script-level `exit N`.
    internal sealed class HookHost : PSHost
    {
        public int? ExitCode { get; private set; }
        public override Guid InstanceId { get; } = Guid.NewGuid();
        public override string Name => "IntuneManagement.MainThreadHook";
        public override Version Version => new(1, 0);
        public override PSHostUserInterface UI => null!;
        public override CultureInfo CurrentCulture => CultureInfo.CurrentCulture;
        public override CultureInfo CurrentUICulture => CultureInfo.CurrentUICulture;
        public override void SetShouldExit(int exitCode) => ExitCode = exitCode;
        public override void EnterNestedPrompt() => throw new NotSupportedException("Nested prompts are not supported by the main-thread hook.");
        public override void ExitNestedPrompt() => throw new NotSupportedException("Nested prompts are not supported by the main-thread hook.");
        public override void NotifyBeginApplication() { }
        public override void NotifyEndApplication() { }
    }

    internal static class ScriptRunner
    {
        public const string EngageVariable = "IM_MAIN_THREAD_HOOK";

        // Returns null when this pwsh was not started with -File <script>.
        internal static (string Script, List<string> ScriptArgs)? ParseCommandLine(string[] args)
        {
            for (int i = 1; i < args.Length - 1; i++)
            {
                if (args[i].Equals("-File", StringComparison.OrdinalIgnoreCase) ||
                    args[i].Equals("-f", StringComparison.OrdinalIgnoreCase))
                {
                    return (Path.GetFullPath(args[i + 1]), args.Skip(i + 2).ToList());
                }
            }
            return null;
        }

        // Do not make this async: every script and UI callback must stay on this thread.
        internal static int Run(string script, List<string> rest)
        {
            MainThread.ManagedId = Environment.CurrentManagedThreadId;
            try
            {
                MainThread.Verify();
                if (!File.Exists(script)) throw new FileNotFoundException("Startup script not found", script);

                var host = new HookHost();
                using var runspace = RunspaceFactory.CreateRunspace(host, InitialSessionState.CreateDefault());
                runspace.ThreadOptions = PSThreadOptions.UseCurrentThread;
                runspace.Open();
                Runspace.DefaultRunspace = runspace;

                using var ps = PowerShell.Create();
                ps.Runspace = runspace;
                ps.Streams.Error.DataAdded += (_, e) => Console.Error.WriteLine(ps.Streams.Error[e.Index]);
                ps.Streams.Warning.DataAdded += (_, e) => Console.Error.WriteLine("WARNING: " + ps.Streams.Warning[e.Index]);
                ps.Streams.Information.DataAdded += (_, e) => Console.WriteLine(ps.Streams.Information[e.Index].MessageData);

                ps.AddCommand(script);
                // Tokens after the script path follow pwsh -File conventions:
                // "-Name value" binds a named parameter, a bare "-Switch" a switch,
                // anything else is positional.
                for (int i = 0; i < rest.Count; i++)
                {
                    if (rest[i].Length > 1 && rest[i][0] == '-')
                    {
                        string name = rest[i].TrimStart('-');
                        if (i + 1 < rest.Count && !(rest[i + 1].Length > 1 && rest[i + 1][0] == '-'))
                            ps.AddParameter(name, rest[++i]);
                        else
                            ps.AddParameter(name);
                    }
                    else ps.AddArgument(rest[i]);
                }

                foreach (var value in ps.Invoke()) Console.WriteLine(value);

                // Only a TERMINATING failure, or an explicit exit code, is a failure.
                //
                // What must NOT reach the caller is ps.HadErrors on its own: it is set
                // by ANY error record the pipeline ever saw, including ones the script
                // itself caught in a try/catch or silenced with -ErrorAction
                // SilentlyContinue. This application does that hundreds of times in a
                // normal run (probing for an optional file, a registry value that may
                // not exist, ...), so returning 1 for it made every healthy macOS
                // launch end in
                //     Main-thread launch failed (1). See the error above.
                // from Start-Avalonia.ps1 - and with no error above, because an error
                // the script already handled never prints.
                //
                // A real terminating error still reports: ps.Invoke() throws and the
                // catch below returns 1.
                if (host.ExitCode is int hostExit) return hostExit;

                // A script-file `exit N` does NOT reach SetShouldExit through the
                // hosting API. It ends the pipeline with HadErrors set, no error
                // record, and $LASTEXITCODE holding N - that combination is the
                // signature of an explicit exit, which is why all three are required
                // here. A handled error matches the first two but leaves
                // $LASTEXITCODE alone, so it correctly falls through to 0.
                if (ps.HadErrors && ps.Streams.Error.Count == 0 &&
                    runspace.SessionStateProxy.GetVariable("LASTEXITCODE") is int exitCode && exitCode != 0)
                    return exitCode;

                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 1;
            }
        }
    }
}

// The runtime's startup-hook contract: a class named StartupHook in the global
// namespace with a static, parameterless void Initialize().
internal static class StartupHook
{
    public static void Initialize()
    {
        // Never let child processes (browser launch, nested pwsh, dotnet) inherit
        // the hook. The runtime has already read this variable for this process.
        Environment.SetEnvironmentVariable("DOTNET_STARTUP_HOOKS", null);

        if (Environment.GetEnvironmentVariable(IntuneManagement.MainThreadHook.ScriptRunner.EngageVariable) != "1")
            return;
        Environment.SetEnvironmentVariable(IntuneManagement.MainThreadHook.ScriptRunner.EngageVariable, null);

        var launch = IntuneManagement.MainThreadHook.ScriptRunner.ParseCommandLine(Environment.GetCommandLineArgs());
        if (launch == null) return;   // not `pwsh -File ...`: let pwsh start normally

        int code = IntuneManagement.MainThreadHook.ScriptRunner.Run(launch.Value.Script, launch.Value.ScriptArgs);
        Environment.Exit(code);       // pwsh's own Main must never run
    }
}
