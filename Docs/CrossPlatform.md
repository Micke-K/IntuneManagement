# Running on macOS and Linux (experimental)

IntuneManagement.Next runs outside Windows through the **Avalonia** UI backend. The
engine is the same one the Windows build uses; only the presentation layer differs.

> **Experimental.** Linux has been exercised during development. macOS has been
> started end to end (module import, sign-in, policy browsing) on Apple Silicon with
> PowerShell 7.6.6, but the native GUI paths have had far less mileage than Windows.
> Treat a macOS run as a bug hunt.

The app says so itself: on the first non-Windows launch it shows a notice with a
**Do not show this message again** checkbox. See [Turning the notice back on](#turning-the-notice-back-on).

## Requirements

| | |
|---|---|
| PowerShell | **7.4 or newer** (`pwsh`) on Linux and macOS - 7.6 / 7.7 included; nothing is pinned to a particular 7.x. pwsh bundles its own .NET runtime, so **no separate .NET install is needed** on either platform. Windows PowerShell 5.1 remains supported by WPF, not Avalonia. |
| Display | A desktop session. X11 or Wayland on Linux, Aqua on macOS. |

## Launching

```sh
./Start-Avalonia.command          # macOS (also double-clickable in Finder) and Linux
```

or directly:

```sh
pwsh -NoProfile -File ./UI/Avalonia/Start-Avalonia.ps1
```

`-ThemeVariant Dark` and `-Provider OAuth` (or `MSAL`, `MgGraph`: the authentication
provider for this session only, without touching the saved setting) are supported by
both entry points.

On macOS both entry points run the script through the **main-thread hook**
(`Bin/MainThreadHook/IntuneManagement.MainThreadHook.dll`, ~10 KB). Cocoa only allows
the GUI on the process's first thread, and a normal `pwsh` pipeline runs on a worker
thread - even when single-threaded. The hook is a .NET *startup hook*: `pwsh` loads it
before its own `Main` runs, on the main thread, and it opens a `UseCurrentThread`
runspace there and runs `Start-Avalonia.ps1` inside the very same `pwsh`. No second
engine, no second runtime: whatever pwsh you have brings its engine, its .NET and its
`$PSHOME/ref` compile references, all consistent with each other. Direct script
invocation re-launches itself through the hook when it notices it is not on the main
thread. See [The macOS main-thread hook](#the-macos-main-thread-hook).

`Start-Avalonia.ps1` sets `IM_UI_BACKEND=Avalonia` for you. On **Linux or Windows**,
setting that variable and importing the module by hand works too (on Windows use
`pwsh -STA`):

```powershell
$env:IM_UI_BACKEND = 'Avalonia'
Import-Module ./IntuneManagement.psd1 -Force
Show-IMMainWindow -View 'IntuneManagement'
```

On macOS do not import the GUI manually in an ordinary `pwsh` session: the native
host rejects initialization off the main thread. Use one of the entry points above.
Headless imports with the `None` backend are unaffected.

Without `IM_UI_BACKEND`, the module defaults to the headless `None` backend off
Windows - useful for automation, and the reason `Connect-IMIntuneManagement` and the
bulk cmdlets work fine on a Mac or a Linux box with no display at all.

## What does not work off Windows

These degrade with a log message rather than an error:

| Feature | Why |
|---|---|
| **Word documentation output** | Needs `Microsoft.Office.Interop.Word` COM automation. HTML, Markdown, CSV and JSON output are unaffected. |
| **MSI property extraction** on app import | Needs the `WindowsInstaller.Installer` COM object. Other app types import normally. |
| **WAM / broker sign-in** | Windows-only. Authentication falls back to the system browser. |

The `Default` theme follows the OS on all three platforms: the Windows app theme, the
macOS appearance setting, and the GNOME colour scheme on Linux. Linux desktops that are
not GNOME have no common way to report a preference and resolve to Light; pick Light or
Dark explicitly in Settings there.

Token cache persistence *is* implemented on both platforms - macOS uses the Keychain
and Linux uses libsecret (gnome-keyring / KWallet), via `MsalCacheHelper`. On a Linux
box with no keyring daemon, expect to sign in every session.

## Turning the notice back on

The startup notice is a normal setting. Clear **Settings -> General -> Hide
experimental platform notice** to see it again.

To preview it on Windows, where it never appears on its own:

```powershell
$env:IM_EXPERIMENTAL_NOTICE = '1'
./UI/Avalonia/Start-Avalonia.ps1
```

## Rebuilding the Avalonia binaries

`Bin/Avalonia` is committed and holds the natives for all three platforms side by
side - `.dll` for Windows, `.so` for Linux, `.dylib` for macOS - plus the managed
Avalonia assemblies, which are platform-neutral.

To rebuild for the machine you are on:

```powershell
./UI/Avalonia/Bootstrap/Restore-AvaloniaBinaries.ps1
```

To build another platform's natives **without that platform** - how the committed
macOS binaries were produced, from Windows:

```powershell
./UI/Avalonia/Bootstrap/Restore-AvaloniaBinaries.ps1 -RuntimeIdentifier osx-arm64
```

NuGet serves the RID-specific native packages regardless of the host OS. The macOS
dylibs are **universal binaries** (x86_64 + arm64 slices), so `osx-arm64` and
`osx-x64` produce identical files and either one covers every Mac.

`dotnet publish` does not clean its output directory, which is what lets one
`Bin/Avalonia` hold all three platforms at once. It does overwrite
`AvaloniaPayload.deps.json` with the last RID published; that file is inert here,
because the module loads the assemblies with `Add-Type -Path` rather than through
the `dotnet` host.

## The macOS main-thread hook

Source: [`UI/Avalonia/Bootstrap/MainThreadHook/StartupHook.cs`](../UI/Avalonia/Bootstrap/MainThreadHook/StartupHook.cs).
Binary: `Bin/MainThreadHook/IntuneManagement.MainThreadHook.dll` - **committed**, so a
plain clone or source ZIP is a complete, runnable download.

`Start-Avalonia.command` sets two environment variables and runs `pwsh -File`:

| Variable | Purpose |
|---|---|
| `DOTNET_STARTUP_HOOKS` | Path of the hook DLL. The .NET runtime calls its `StartupHook.Initialize()` on the main thread before `pwsh`'s own `Main`. |
| `IM_MAIN_THREAD_HOOK=1` | Engages the hook. Without it `Initialize` returns at once and pwsh starts normally, so child processes are unaffected (the hook also clears both variables from its own environment). |

`Initialize` then parses `-File <script> [args]` from the command line, opens a
`UseCurrentThread` runspace, runs the script and terminates the process with the
script's exit code. The PowerShell side recognises an engaged hook by the presence of
the `[IntuneManagement.MainThreadHook.MainThread]` type; `::Verify()` throws when called
off the main thread, which `Tests/MainThreadHook.Tests.ps1` and the UI smoke test use.

Compatibility comes from two choices in the project file: it targets `net8.0` (the
runtime of pwsh 7.4, the oldest supported engine) and references
`Microsoft.PowerShell.SDK` **7.4.x at compile time only**. .NET binds a reference to a
lower `System.Management.Automation` version against whatever newer one pwsh loaded,
so the one DLL works in 7.4, 7.5, 7.6, 7.7 and later on .NET 8, 9, 10 and later
without a rebuild. Keep it that way: do not bump the SDK reference to "latest".

Rebuild only when `StartupHook.cs` changes (needs the .NET 8+ SDK and NuGet access),
then commit the DLL:

```powershell
./UI/Avalonia/Bootstrap/Publish-MainThreadHook.ps1
```

The project's own tests then run the hook inside a child `pwsh` on every OS (the
mechanism is not macOS-specific; only Cocoa needs it) and check thread ownership,
parameter forwarding, exit codes, module import and that an un-engaged hook is inert,
plus a native-backend smoke test on a Mac desktop session that initializes Avalonia and
exercises a dispatcher callback and button event.

Before calling macOS supported, test on Intel and Apple Silicon: startup, sign-in
and browser return, message boxes, Bulk Compare pickers, file dialogs, clipboard,
closing child windows and quitting/relaunching.

## Reporting problems

Include the platform and architecture (`[System.Runtime.InteropServices.RuntimeInformation]::OSDescription`
and `::ProcessArchitecture`), the PowerShell version, and `IntuneManagement.log` from
the app data folder.
