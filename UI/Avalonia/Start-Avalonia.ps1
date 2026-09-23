[CmdletBinding()]
param(
    [switch]$NoUI,
    [ValidateSet('Default','Light','Dark')]
    [string]$ThemeVariant,
    # Authentication provider for this session only (MSAL, OAuth, MgGraph, ...),
    # overriding the saved "Active authentication provider" setting without
    # changing it. Surfaced as IM_AUTH_PROVIDER for the module; forwarded as an
    # argument through the macOS main-thread and Windows STA re-launches.
    [string]$Provider
)

if ($PSVersionTable.PSEdition -ne 'Core') {
    Write-Error "The Avalonia backend requires PowerShell 7+. Launch with pwsh.exe."
    return
}

if ($Provider) {
    $env:IM_AUTH_PROVIDER = $Provider
}
else {
    # Same cleanup as IM_THEME_VARIANT below: a previous -Provider launch from
    # this shell must not silently stick to later plain launches.
    Remove-Item Env:IM_AUTH_PROVIDER -ErrorAction SilentlyContinue
}

# Cocoa only allows the GUI on the process main thread, and a normal pwsh
# pipeline is not that thread. Start-Avalonia.command engages the main-thread
# startup hook up front; when this script is run directly (pwsh -File ...) and
# the hook is absent, re-launch this same pwsh with the hook so the script runs
# on the main thread. See UI/Avalonia/Bootstrap/MainThreadHook/StartupHook.cs.
if ($IsMacOS -and -not $NoUI -and -not ('IntuneManagement.MainThreadHook.MainThread' -as [type])) {
    $root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $hook = Join-Path $root 'Bin/MainThreadHook/IntuneManagement.MainThreadHook.dll'
    if (-not (Test-Path -LiteralPath $hook)) { throw "Main-thread hook missing: $hook. Re-download the application or rebuild it with UI/Avalonia/Bootstrap/Publish-MainThreadHook.ps1." }
    $pwsh = (Get-Process -Id $PID).Path
    $env:DOTNET_STARTUP_HOOKS = if ($env:DOTNET_STARTUP_HOOKS) { $hook + [IO.Path]::PathSeparator + $env:DOTNET_STARTUP_HOOKS } else { $hook }
    $env:IM_MAIN_THREAD_HOOK = '1'
    try {
        $argList = @('-NoLogo', '-NoProfile', '-File', $PSCommandPath)
        if ($ThemeVariant) { $argList += @('-ThemeVariant', $ThemeVariant) }
        if ($Provider) { $argList += @('-Provider', $Provider) }
        & $pwsh @argList
        if ($LASTEXITCODE -ne 0) { throw "Main-thread launch failed ($LASTEXITCODE). See the error above." }
    }
    finally {
        Remove-Item Env:DOTNET_STARTUP_HOOKS, Env:IM_MAIN_THREAD_HOOK -ErrorAction SilentlyContinue
    }
    return
}

# Surface -ThemeVariant as the env var Initialize-AvaloniaRuntime reads. Setting
# it before the STA self-relaunch makes it inherited by the child pwsh process.
if ($ThemeVariant) {
    $env:IM_THEME_VARIANT = $ThemeVariant
}
else {
    # The env var outranks the AppTheme setting, and a previous -ThemeVariant
    # launch leaves it behind in this shell - without this cleanup every later
    # plain launch from the same terminal silently kept the old variant and
    # the Settings theme appeared to have no effect.
    Remove-Item Env:IM_THEME_VARIANT -ErrorAction SilentlyContinue
}

# Avalonia hosts on the calling thread. Win32 platform features (clipboard,
# drag/drop, OLE) want STA. PowerShell 7 defaults to MTA, so re-launch self
# with -STA when needed. Detect the re-launch via $env:IM_AVALONIA_STA so we
# don't recurse. Apartment state is a Win32/OLE concept — pwsh rejects -STA on
# Linux/macOS, where Avalonia uses X11/Cocoa backends that don't need it, so
# only attempt the relaunch on Windows.
if ($IsWindows -and -not $env:IM_AVALONIA_STA -and [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $env:IM_AVALONIA_STA = '1'
    $argList = @('-NoLogo', '-NoProfile', '-STA', '-File', $PSCommandPath)
    if ($NoUI) { $argList += '-NoUI' }
    if ($ThemeVariant) { $argList += @('-ThemeVariant', $ThemeVariant) }
    if ($Provider) { $argList += @('-Provider', $Provider) }
    & pwsh @argList
    exit $LASTEXITCODE
}

$env:IM_UI_BACKEND = 'Avalonia'

# Linux: Avalonia's file dialogs use xdg-desktop-portal (the native, correctly
# sized picker) only when it can reach the session bus. xrdp/XFCE sessions often
# leave XDG_RUNTIME_DIR (and sometimes DBUS_SESSION_BUS_ADDRESS) unset, which
# forces Avalonia's oversized managed dialog and also breaks "open in browser".
# Point them at the standard runtime dir / bus when they exist so the native
# picker (and Open-ExternalUri) work.
if (-not $IsWindows) {
    try {
        $uid = (& id -u).Trim()
        if ([string]::IsNullOrEmpty($env:XDG_RUNTIME_DIR) -and (Test-Path "/run/user/$uid")) {
            $env:XDG_RUNTIME_DIR = "/run/user/$uid"
        }
        if ([string]::IsNullOrEmpty($env:DBUS_SESSION_BUS_ADDRESS) -and $env:XDG_RUNTIME_DIR -and (Test-Path "$($env:XDG_RUNTIME_DIR)/bus")) {
            $env:DBUS_SESSION_BUS_ADDRESS = "unix:path=$($env:XDG_RUNTIME_DIR)/bus"
        }
    }
    catch { }
}

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$avaloniaBin = Join-Path $projectRoot 'Bin/Avalonia'

if (-not (Test-Path $avaloniaBin)) {
    Write-Host "Avalonia binaries missing. Running Bootstrap/Restore-AvaloniaBinaries.ps1..." -ForegroundColor Yellow
    & (Join-Path $PSScriptRoot 'Bootstrap/Restore-AvaloniaBinaries.ps1')
}

Import-Module (Join-Path $projectRoot 'IntuneManagement.psd1') -Force -ErrorAction Stop

if (-not $NoUI) {
    Show-IMMainWindow -View 'IntuneManagement'
}
