using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Data.Converters;
using Avalonia.Markup.Xaml;
using Avalonia.Markup.Xaml.Styling;
using Avalonia.Platform.Storage;
using Avalonia.Themes.Fluent;
using Avalonia.Threading;

namespace IntuneManagement.AvaloniaHost;

public sealed class IntuneManagementApp : Application
{
    public static string ThemeVariant { get; set; } = "Default";

    public override void Initialize()
    {
        Styles.Add(new FluentTheme());

        // DataGrid ships in a separate package with its own theme XAML. Add
        // the StyleInclude here so DataGrid controls render with Fluent
        // styling instead of as unstyled placeholders. Wrapped in try/catch
        // so a missing Avalonia.Controls.DataGrid.dll (e.g. older restore)
        // doesn't kill app startup — DataGrid just won't be usable.
        try
        {
            var dgStylesUri = new Uri("avares://Avalonia.Controls.DataGrid/Themes/Fluent.xaml");
            Styles.Add(new StyleInclude(dgStylesUri) { Source = dgStylesUri });
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.WriteLine("DataGrid theme include failed: " + ex.Message);
        }

        RequestedThemeVariant = string.Equals(ThemeVariant, "Dark", StringComparison.OrdinalIgnoreCase)
            ? Avalonia.Styling.ThemeVariant.Dark
            : Avalonia.Styling.ThemeVariant.Light;
    }
}

// AvaloniaHost: PowerShell-driven Avalonia host running on the calling thread.
//
// Model: Setup runs on the PowerShell thread. That thread becomes Avalonia's
// UI thread. RunMainWindow blocks running the dispatcher loop until the last
// window closes. Because PS thread == UI thread, no cross-thread marshaling
// is needed for UI work, and PowerShell scriptblocks attached to events fire
// on the same thread that subscribed them — exactly like WPF in PowerShell.
public static class Host
{
    private static ClassicDesktopStyleApplicationLifetime _lifetime;
    private static bool _setup;
    private static readonly object _gate = new();
    private static string[] _droppedAssemblies;

    [DllImport("/usr/lib/libSystem.B.dylib")]
    private static extern int pthread_main_np();

    public static void Initialize(string themeVariant = "Default")
    {
        if (OperatingSystem.IsMacOS() && pthread_main_np() == 0)
            throw new InvalidOperationException(
                "Avalonia on macOS requires the process main thread. Launch with Start-Avalonia.command.");
        lock (_gate)
        {
            if (_setup) return;
            IntuneManagementApp.ThemeVariant = themeVariant ?? "Default";

            _lifetime = new ClassicDesktopStyleApplicationLifetime
            {
                Args = Array.Empty<string>(),
                ShutdownMode = ShutdownMode.OnLastWindowClose
            };

            AppBuilder.Configure<IntuneManagementApp>()
                .UsePlatformDetect()
                .WithInterFont()
                .LogToTrace()
                .SetupWithLifetime(_lifetime);

            SanitizeXamlTypeSystem();

            _setup = true;
        }
    }

    // Avalonia's runtime XAML loader resolves types by walking every assembly in
    // the AppDomain and calling GetExportedTypes() on each. That call throws for
    // an assembly whose type references cannot be resolved, and the exception
    // propagates out of the loader, so a single unreadable assembly stops every
    // XAML file in the app from parsing.
    //
    // On Windows, PowerShell's assembly resolver answers some requests out of the
    // .NET Framework GAC (System.Web.Extensions v4.0 is loaded during pwsh
    // startup). Those assemblies reference .NET Framework types that do not exist
    // on modern .NET - System.Web ships as a facade with no types at all - so
    // enumerating them throws TypeLoadException. Office interop assemblies, which
    // the documentation subsystem loads for Word output, fail the same way.
    //
    // Drop those assemblies from the loader's type system before the first parse.
    // The XAML in this app only references Avalonia and our own types, so nothing
    // that can be dropped here was ever reachable from markup. On Linux and macOS
    // no such assembly is loaded and this is a no-op.
    //
    // Returns the names of the assemblies that were dropped. Reflection is used
    // against internal Avalonia members, so every failure path degrades to
    // "changed nothing" rather than breaking startup.
    public static string[] SanitizeXamlTypeSystem()
    {
        if (_droppedAssemblies != null) return _droppedAssemblies;

        const BindingFlags Flags = BindingFlags.Public | BindingFlags.NonPublic |
                                   BindingFlags.Static | BindingFlags.Instance;
        var dropped = new List<string>();
        try
        {
            var loaderAsm = typeof(AvaloniaRuntimeXamlLoader).Assembly;
            var compiler = loaderAsm.GetType("Avalonia.Markup.Xaml.XamlIl.AvaloniaXamlIlRuntimeCompiler");
            var initSre = compiler?.GetMethod("InitializeSre", Flags);
            var tsField = compiler?.GetField("_sreTypeSystem", Flags);
            if (initSre == null || tsField == null) return _droppedAssemblies = Array.Empty<string>();

            // Build the type system so it can be inspected before it is used.
            // This same walk is what fails, but the type system field is assigned
            // before the failing step, so the object is available afterwards
            // either way - the throw is expected here and is not an error.
            try { initSre.Invoke(null, null); } catch { }

            var typeSystem = tsField.GetValue(null);
            var listField = typeSystem?.GetType().GetField("_assemblies", Flags);
            var assemblies = listField?.GetValue(typeSystem) as IList;
            if (assemblies == null) return _droppedAssemblies = Array.Empty<string>();

            var doomed = new List<object>();
            foreach (var entry in assemblies)
            {
                if (entry == null) continue;
                var clr = entry.GetType().GetProperty("Assembly", Flags)?.GetValue(entry) as Assembly;
                if (clr == null) continue;
                try
                {
                    clr.GetExportedTypes();
                }
                catch
                {
                    doomed.Add(entry);
                    dropped.Add(clr.GetName().Name);
                }
            }

            foreach (var entry in doomed) assemblies.Remove(entry);

            // Re-run initialization now that the type system can be walked end to
            // end. The first attempt aborted partway, leaving the xmlns mappings
            // and emit context unset; this pass completes them.
            if (doomed.Count > 0) initSre.Invoke(null, null);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.WriteLine("SanitizeXamlTypeSystem skipped: " + ex.Message);
        }

        return _droppedAssemblies = dropped.ToArray();
    }

    public static object LoadXaml(string xaml, string baseUri = null)
    {
        EnsureSetup();
        Uri uri = baseUri != null ? new Uri(baseUri) : null;
        return AvaloniaRuntimeXamlLoader.Load(xaml, localAssembly: null, rootInstance: null, uri: uri);
    }

    // Parse a Styles or ResourceDictionary file and merge it into Application.Current.
    // Used to load theme files so XAML across the app can resolve {DynamicResource ...}
    // lookups for brushes and pick up control style setters.
    public static void LoadStyles(string xamlPath)
    {
        EnsureSetup();
        var text = File.ReadAllText(xamlPath);
        var uri = new Uri(new Uri("file:///" + xamlPath.Replace('\\', '/')).AbsoluteUri);
        var obj = AvaloniaRuntimeXamlLoader.Load(text, localAssembly: null, rootInstance: null, uri: uri);
        switch (obj)
        {
            case Avalonia.Styling.Styles styles:
                Application.Current.Styles.Add(styles);
                break;
            case Avalonia.Styling.Style style:
                Application.Current.Styles.Add(style);
                break;
            case Avalonia.Controls.ResourceDictionary rd:
                Application.Current.Resources.MergedDictionaries.Add(rd);
                break;
            default:
                throw new InvalidOperationException(
                    "LoadStyles: expected Styles, Style, or ResourceDictionary in " + xamlPath +
                    "; got " + (obj?.GetType().FullName ?? "null"));
        }
    }

    public static object FindByName(object root, string name)
    {
        if (root is not Control control) return null;
        return (object)control.FindControl<Control>(name) ?? control.FindNameScope()?.Find(name);
    }

    [DllImport("user32.dll")]
    private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    private static IntPtr GetHwnd(Window w) =>
        w?.TryGetPlatformHandle() is { } h ? h.Handle : IntPtr.Zero;

    // Run the main window: assigns it as MainWindow, calls lifetime.Start which
    // blocks pumping the dispatcher loop until the last window closes.
    public static int RunMainWindow(Window window)
    {
        EnsureSetup();
        _lifetime.MainWindow = window;
        return _lifetime.Start(Array.Empty<string>());
    }

    // Run a short, bounded nested dispatcher loop. Dispatcher.RunJobs() only
    // drains work that has already reached Avalonia's managed queue; it does not
    // provide the platform event loop needed to turn native mouse/keyboard input
    // into routed events while synchronous PowerShell code owns the UI thread.
    // Authentication waits call this repeatedly so their status-overlay Cancel
    // button remains genuinely clickable without allowing an unbounded nested
    // loop or moving PowerShell state onto another thread.
    public static void PumpEventsOnce(int maximumMilliseconds = 10)
    {
        EnsureSetup();
        if (!Dispatcher.UIThread.CheckAccess())
            throw new InvalidOperationException("PumpEventsOnce must run on the Avalonia UI thread.");

        maximumMilliseconds = Math.Clamp(maximumMilliseconds, 1, 50);
        var frame = new DispatcherFrame();
        var timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(maximumMilliseconds)
        };

        EventHandler tick = null;
        tick = (_, _) =>
        {
            timer.Stop();
            timer.Tick -= tick;
            frame.Continue = false;
        };

        timer.Tick += tick;
        timer.Start();
        try
        {
            Dispatcher.UIThread.PushFrame(frame);
        }
        finally
        {
            timer.Stop();
            timer.Tick -= tick;
            frame.Continue = false;
        }
    }

    // Modal-style dialog: use Avalonia's owned ShowDialog so the owner remains
    // visually stable, then pump a nested DispatcherFrame until the dialog closes.
    // Avalonia's Window.ShowDialog<T>(owner) returns a Task and is awkward to
    // await synchronously; this approach gives sync semantics matching WPF.
    public static void ShowDialog(Window dialog, Window owner = null)
    {
        EnsureSetup();

        var frame = new DispatcherFrame();

        if (owner != null)
        {
            dialog.WindowStartupLocation = WindowStartupLocation.Manual;
            dialog.Position = new PixelPoint(
                owner.Position.X + 40,
                owner.Position.Y + 40);
        }

        if (owner != null)
        {
            dialog.ShowDialog(owner).ContinueWith(_ =>
                Dispatcher.UIThread.Post(() => frame.Continue = false));
        }
        else
        {
            dialog.Closed += (_, _) => frame.Continue = false;
            dialog.Show();
        }

        Dispatcher.UIThread.PushFrame(frame);
    }

    // Synchronous folder picker. The async StorageProvider API is awkward to
    // bridge to a PowerShell caller that needs a return value; same approach
    // as ShowDialog — push a nested DispatcherFrame, set Continue=false on the
    // UI thread once the task completes.
    public static string OpenFolderPicker(Window owner, string title)
    {
        return OpenFolderPicker(owner, title, null);
    }

    public static string OpenFolderPicker(Window owner, string title, string suggestedStartFolder)
    {
        EnsureSetup();
        var topLevel = owner != null ? TopLevel.GetTopLevel(owner) : null;
        if (topLevel?.StorageProvider == null) return null;

        var options = new FolderPickerOpenOptions
        {
            AllowMultiple = false,
            Title = string.IsNullOrEmpty(title) ? "Select folder" : title
        };

        // Folder resolution and the picker run in one async flow - blocking on
        // TryGetFolderFromPathAsync before the frame is pushed would deadlock the
        // dispatcher thread (see SaveFilePicker).
        async Task<string> PickAsync()
        {
            if (!string.IsNullOrEmpty(suggestedStartFolder) && Directory.Exists(suggestedStartFolder))
            {
                try
                {
                    var folder = await topLevel.StorageProvider.TryGetFolderFromPathAsync(new Uri(suggestedStartFolder));
                    if (folder != null) options.SuggestedStartLocation = folder;
                }
                catch { /* swallow - picker just opens in default location */ }
            }

            var picked = await topLevel.StorageProvider.OpenFolderPickerAsync(options);
            return (picked != null && picked.Count > 0) ? picked[0].Path.LocalPath : null;
        }

        string result = null;
        var frame = new DispatcherFrame();

        PickAsync().ContinueWith(t =>
        {
            string r = (t.Status == TaskStatus.RanToCompletion) ? t.Result : null;
            Dispatcher.UIThread.Post(() =>
            {
                result = r;
                frame.Continue = false;
            });
        });

        Dispatcher.UIThread.PushFrame(frame);
        return result;
    }

    // Copy text to the system clipboard. Called from the UI thread while the
    // auth flow surfaces a device-code; like the pickers we run the async
    // Clipboard API under a nested DispatcherFrame rather than blocking on it
    // (a GetResult() on the UI thread would deadlock the continuation). Uses
    // Avalonia's clipboard (X11/Wayland direct) so it works on a bare Linux box
    // where PowerShell's Set-Clipboard is a no-op.
    public static void SetClipboardText(Window owner, string text)
    {
        EnsureSetup();
        var topLevel = owner != null ? TopLevel.GetTopLevel(owner) : null;
        var clipboard = topLevel?.Clipboard;
        if (clipboard == null) return;

        var frame = new DispatcherFrame();
        clipboard.SetTextAsync(text ?? string.Empty).ContinueWith(_ =>
            Dispatcher.UIThread.Post(() => frame.Continue = false));
        Dispatcher.UIThread.PushFrame(frame);
    }

    // Synchronous open-file picker. Single-select. filterName/filterPattern
    // mirror SaveFilePicker — display name + semicolon-separated patterns.
    // Returns null on cancel.
    public static string OpenFilePicker(Window owner, string title, string suggestedStartFolder, string filterName, string filterPattern)
    {
        EnsureSetup();
        var topLevel = owner != null ? TopLevel.GetTopLevel(owner) : null;
        if (topLevel?.StorageProvider == null) return null;

        var options = new FilePickerOpenOptions
        {
            Title = string.IsNullOrEmpty(title) ? "Open file" : title,
            AllowMultiple = false
        };

        if (!string.IsNullOrEmpty(filterName) && !string.IsNullOrEmpty(filterPattern))
        {
            options.FileTypeFilter = new[]
            {
                new FilePickerFileType(filterName)
                {
                    Patterns = filterPattern.Split(';', StringSplitOptions.RemoveEmptyEntries)
                }
            };
        }

        // One async flow - see SaveFilePicker for why this must not block.
        async Task<string> PickAsync()
        {
            if (!string.IsNullOrEmpty(suggestedStartFolder) && Directory.Exists(suggestedStartFolder))
            {
                try
                {
                    var folder = await topLevel.StorageProvider.TryGetFolderFromPathAsync(new Uri(suggestedStartFolder));
                    if (folder != null) options.SuggestedStartLocation = folder;
                }
                catch { /* swallow - picker just opens in default location */ }
            }

            var picked = await topLevel.StorageProvider.OpenFilePickerAsync(options);
            return (picked != null && picked.Count > 0) ? picked[0].Path.LocalPath : null;
        }

        string result = null;
        var frame = new DispatcherFrame();

        PickAsync().ContinueWith(t =>
        {
            string r = (t.Status == TaskStatus.RanToCompletion) ? t.Result : null;
            Dispatcher.UIThread.Post(() =>
            {
                result = r;
                frame.Continue = false;
            });
        });

        Dispatcher.UIThread.PushFrame(frame);
        return result;
    }

    // Synchronous save-file picker. filterName/filterPattern model the WPF
    // "Json (*.json)|*.json" pairs: callers pass a display name and a
    // semicolon-separated list of patterns ("*.json;*.txt").
    // Keeps the original 6-argument signature working; forwards with no start folder.
    public static string SaveFilePicker(Window owner, string title, string suggestedFileName, string defaultExtension, string filterName, string filterPattern)
    {
        return SaveFilePicker(owner, title, suggestedFileName, defaultExtension, filterName, filterPattern, null);
    }

    public static string SaveFilePicker(Window owner, string title, string suggestedFileName, string defaultExtension, string filterName, string filterPattern, string suggestedStartFolder)
    {
        EnsureSetup();
        var topLevel = owner != null ? TopLevel.GetTopLevel(owner) : null;
        if (topLevel?.StorageProvider == null) return null;

        var options = new FilePickerSaveOptions
        {
            Title = string.IsNullOrEmpty(title) ? "Save" : title,
            SuggestedFileName = suggestedFileName,
            DefaultExtension = defaultExtension
        };

        if (!string.IsNullOrEmpty(filterName) && !string.IsNullOrEmpty(filterPattern))
        {
            options.FileTypeChoices = new[]
            {
                new FilePickerFileType(filterName)
                {
                    Patterns = filterPattern.Split(';', StringSplitOptions.RemoveEmptyEntries)
                }
            };
        }

        // Resolve the start folder (WPF's InitialDirectory) and show the picker in
        // ONE async flow. Blocking on TryGetFolderFromPathAsync before the frame is
        // pushed would deadlock: this runs on the dispatcher thread, so any
        // continuation the storage provider posts back to it could never run.
        async Task<string> PickAsync()
        {
            if (!string.IsNullOrEmpty(suggestedStartFolder) && Directory.Exists(suggestedStartFolder))
            {
                try
                {
                    var startFolder = await topLevel.StorageProvider.TryGetFolderFromPathAsync(new Uri(suggestedStartFolder));
                    if (startFolder != null) options.SuggestedStartLocation = startFolder;
                }
                catch { /* swallow - picker just opens in the default location */ }
            }

            var file = await topLevel.StorageProvider.SaveFilePickerAsync(options);
            return file?.Path.LocalPath;
        }

        string result = null;
        var frame = new DispatcherFrame();

        PickAsync().ContinueWith(t =>
        {
            string r = (t.Status == TaskStatus.RanToCompletion) ? t.Result : null;
            Dispatcher.UIThread.Post(() =>
            {
                result = r;
                frame.Continue = false;
            });
        });

        Dispatcher.UIThread.PushFrame(frame);
        return result;
    }

    public static void Shutdown()
    {
        if (!_setup) return;
        _lifetime?.Shutdown();
    }

    // Factory for IntunePolicyPathConverter: PowerShell's [FullName]::new()
    // resolution walks every loaded assembly and calls GetTypes() on each.
    // After Avalonia's runtime XAML loader has compiled .axaml files into
    // dynamic assemblies, GetTypes() on those throws ReflectionTypeLoadException
    // and the PS resolution fails. Routing through this factory reuses the
    // already-resolved Host type (assigned to $script:AvaloniaHostType at
    // init time, before any XAML compilation), so no fresh scan happens.
    public static IValueConverter CreatePathConverter() => new IntunePolicyPathConverter();

    private static void EnsureSetup()
    {
        if (!_setup)
        {
            throw new InvalidOperationException(
                "AvaloniaHost is not initialized. Call IntuneManagement.AvaloniaHost.Host.Initialize() first.");
        }
    }
}

// Walks a dotted property path on an arbitrary value (CLR object or PSObject).
// Used by the Intune Manager DataGrid for the user-defined "ObjectColumns"
// override — those paths reach into the original IntunePolicyBase / its
// underlying PSCustomObject (`Object`) and so do not resolve against the
// canonical fields on IntuneObjectRowItem. The binding points at the row's
// `Source` and passes the path string as ConverterParameter; this converter
// steps through PSObject.Properties at each segment, which catches both CLR
// properties and PSObject NoteProperties (Avalonia's binder ignores the
// latter on its own — see avalonia-binding-needs-clr-types).
//
// Lives in C# because PowerShell `class X : SomeAvaloniaType` resolves the
// base type at parse time, before Initialize-AvaloniaRuntime has loaded the
// Avalonia assemblies.
//
// PSObject access is via late-bound reflection rather than a compile-time
// reference to System.Management.Automation: that assembly's transitive
// references (e.g. System.Linq.Expressions v10) outpace the .NET ref-pack
// shipped with the SDK (v8), and Roslyn rejects the cross-version mismatch
// with CS1705. Reflection sidesteps the compile-time check entirely.
public sealed class IntunePolicyPathConverter : IValueConverter
{
    private static readonly Type _psObjectType =
        Type.GetType("System.Management.Automation.PSObject, System.Management.Automation");
    private static readonly MethodInfo _asPSObject =
        _psObjectType?.GetMethod("AsPSObject", new[] { typeof(object) });
    private static readonly PropertyInfo _propertiesProp =
        _psObjectType?.GetProperty("Properties");

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value == null || parameter == null) return null;
        var path = parameter.ToString();
        if (string.IsNullOrWhiteSpace(path)) return null;

        object current = value;
        foreach (var segment in path.Split('.'))
        {
            if (current == null) return null;

            object next = null;
            try
            {
                if (TryGetPSProperty(current, segment, out var psValue))
                {
                    next = psValue;
                }
                else
                {
                    var clrType = current.GetType();
                    var clrProp = clrType.GetProperty(segment);
                    if (clrProp != null)
                    {
                        next = clrProp.GetValue(current);
                    }
                    else
                    {
                        var clrField = clrType.GetField(segment);
                        if (clrField != null)
                        {
                            next = clrField.GetValue(current);
                        }
                    }
                }
            }
            catch
            {
                return null;
            }
            current = next;
        }

        if (current == null) return null;
        if (current is string || current.GetType().IsPrimitive) return current;
        return current.ToString();
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return null;
    }

    private static bool TryGetPSProperty(object instance, string name, out object value)
    {
        value = null;
        if (_psObjectType == null || _asPSObject == null || _propertiesProp == null)
            return false;

        var psObj = _asPSObject.Invoke(null, new[] { instance });
        if (psObj == null) return false;

        var properties = _propertiesProp.GetValue(psObj);
        if (properties == null) return false;

        var indexer = properties.GetType().GetProperty("Item", new[] { typeof(string) });
        var prop = indexer?.GetValue(properties, new object[] { name });
        if (prop == null) return false;

        var valueProp = prop.GetType().GetProperty("Value");
        if (valueProp == null) return false;

        value = valueProp.GetValue(prop);
        return true;
    }
}
