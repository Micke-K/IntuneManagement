# Avalonia counterpart of UI/WPF/Extensions/CoreUI.ps1.
#
# Threading model: Avalonia is set up on the PowerShell thread. That thread
# IS the UI thread, so UI access is direct (no Dispatcher.Invoke marshaling)
# and PS scriptblocks attached to events fire on the same thread that
# subscribed them — same as WPF in PowerShell.

$script:AvaloniaBinDir = Join-Path $script:AppRootFolder 'Bin/Avalonia'
$script:AvaloniaInitialized = $false
$script:AvaloniaHostType = $null
$script:Window = $null

# View registry: ports the WPF CoreUI's $script:viewObjects + $script:ActiveView
# pair. Show-MainWindow populates the registry by discovering ViewObjectBase
# subclasses; Show-View flips ActiveView and reflects the change in the chrome.
$script:viewObjects = @()
$script:ActiveView = $null

function Resolve-NetCoreRefPack {
    # pwsh ships the reference assemblies for its own runtime in $PSHOME/ref (it
    # is what Add-Type compiles against by default), and they match the runtime
    # the script is executing on by construction - including on macOS, where the
    # main-thread hook runs this code inside the user's pwsh. Mixing them with
    # a different SDK's references causes Roslyn CS1705 during host compilation,
    # so this is the first and normal choice on every platform.
    $bundledRef = Join-Path $PSHOME 'ref'
    if (Test-Path (Join-Path $bundledRef 'System.Runtime.dll')) {
        return $bundledRef
    }

    # Locate the highest installed Microsoft.NETCore.App.Ref pack so Roslyn
    # has reference assemblies for compiling the C# host.
    $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
    if (-not $dotnet) {
        throw "No reference assemblies: $PSHOME/ref is missing and no dotnet SDK was found in PATH."
    }

    # Resolve the real dotnet root. On Linux the launcher is usually a symlink
    # (e.g. /usr/bin/dotnet -> /usr/lib/dotnet/dotnet), so Split-Path on the
    # symlink path points at /usr/bin, which has no packs/ folder. Build a list
    # of candidate roots (DOTNET_ROOT, the symlink target dir, the launcher dir)
    # and pick the first that actually contains the ref pack.
    $candidateRoots = @()
    if ($env:DOTNET_ROOT) { $candidateRoots += $env:DOTNET_ROOT }
    try {
        $resolved = (Get-Item -LiteralPath $dotnet.Path).ResolveLinkTarget($true)
        if ($resolved) { $candidateRoots += (Split-Path -Parent $resolved.FullName) }
    } catch { }
    $candidateRoots += (Split-Path -Parent $dotnet.Path)

    $refRoot = $null
    foreach ($root in ($candidateRoots | Where-Object { $_ } | Select-Object -Unique)) {
        $candidate = Join-Path $root 'packs/Microsoft.NETCore.App.Ref'
        if (Test-Path $candidate) { $refRoot = $candidate; break }
    }

    if (-not $refRoot) {
        throw "Microsoft.NETCore.App.Ref pack not found under any of: $($candidateRoots -join ', '). Install the .NET 8+ SDK."
    }

    $versionDirs = Get-ChildItem -Path $refRoot -Directory |
        Where-Object { $_.Name -match '^\d+\.\d+\.\d+' } |
        Sort-Object { [Version]($_.Name -replace '-.*$','') } -Descending

    foreach ($vd in $versionDirs) {
        $refDir = Join-Path $vd.FullName 'ref'
        if (-not (Test-Path $refDir)) { continue }
        $tfms = Get-ChildItem -Path $refDir -Directory |
            Where-Object { $_.Name -match '^net\d+\.\d+$' } |
            Sort-Object { [Version](($_.Name -replace '^net','')) } -Descending
        if ($tfms) { return $tfms[0].FullName }
    }

    throw "No usable .NET reference assemblies found under $refRoot."
}

function Initialize-AvaloniaRuntime {
    if ($script:AvaloniaInitialized) { return }

    if (-not (Test-Path $script:AvaloniaBinDir)) {
        throw "Avalonia binaries not found at $($script:AvaloniaBinDir). Run .\UI\Avalonia\Bootstrap/Restore-AvaloniaBinaries.ps1 first."
    }

    if ($PSVersionTable.PSEdition -ne 'Core') {
        throw "Avalonia backend requires PowerShell 7+ (PSEdition=Core). Current edition: $($PSVersionTable.PSEdition)."
    }

    $coreAssemblies = @(
        'Avalonia.Base.dll',
        'Avalonia.Controls.dll',
        'Avalonia.Controls.DataGrid.dll',
        'Avalonia.Desktop.dll',
        'Avalonia.DesignerSupport.dll',
        'Avalonia.Dialogs.dll',
        'Avalonia.Markup.dll',
        'Avalonia.Markup.Xaml.dll',
        'Avalonia.Markup.Xaml.Loader.dll',
        'Avalonia.Themes.Fluent.dll',
        'Avalonia.Fonts.Inter.dll',
        'Avalonia.Win32.dll',
        'Avalonia.Skia.dll',
        'Avalonia.dll'
    )

    $loaded = @()
    foreach ($name in $coreAssemblies) {
        $path = Join-Path $script:AvaloniaBinDir $name
        if (Test-Path $path) {
            try {
                Add-Type -Path $path -ErrorAction Stop
                $loaded += $path
            } catch {
                Write-LogDebug "Skipping $name (already loaded or not needed): $($_.Exception.Message)"
            }
        }
    }

    if (-not $loaded) {
        throw "Failed to load any Avalonia assemblies from $($script:AvaloniaBinDir)."
    }

    $hostSource = [IO.File]::ReadAllText((Join-Path $script:AppUIRootFolder 'Bootstrap/AvaloniaHost.cs'))

    $bclRefDir = Resolve-NetCoreRefPack
    $bclRefs = Get-ChildItem -Path $bclRefDir -Filter '*.dll' -File | ForEach-Object FullName

    $avaloniaRefs = @(
        'Avalonia.Base.dll',
        'Avalonia.Controls.dll',
        'Avalonia.Controls.DataGrid.dll',
        'Avalonia.Desktop.dll',
        'Avalonia.Markup.Xaml.dll',
        'Avalonia.Markup.Xaml.Loader.dll',
        'Avalonia.Themes.Fluent.dll',
        'Avalonia.dll'
    ) | ForEach-Object { Join-Path $script:AvaloniaBinDir $_ } |
        Where-Object { Test-Path $_ }

    $referenced = ($bclRefs + $avaloniaRefs + $loaded) | Sort-Object -Unique

    # -WarningAction SilentlyContinue: when the installed .NET SDK ref-pack is
    # a newer major than what Avalonia.dll references (e.g. SDK 9 vs Avalonia
    # built against runtime 8), Roslyn emits CS1701 "assuming assembly reference
    # matches identity" for every Avalonia.* reference. Binding redirects handle
    # it at runtime; the noise just clutters startup.
    Add-Type -TypeDefinition $hostSource `
        -ReferencedAssemblies $referenced `
        -Language CSharp -IgnoreWarnings -ErrorAction Stop `
        -WarningAction SilentlyContinue

    $script:AvaloniaHostType = [IntuneManagement.AvaloniaHost.Host]
    # IM_THEME_VARIANT env var lets you flip dark/light without editing the
    # settings JSON. Otherwise prefer the AppTheme setting (the key the
    # Settings dialog writes) so theme picks selected via the UI survive a
    # restart. ThemeVariant is the older Avalonia-specific key kept as a
    # fallback for users who set it directly before AppTheme existed.
    $themeVariant = if ($env:IM_THEME_VARIANT) {
        $env:IM_THEME_VARIANT
    } else {
        # Get-SettingValue (not raw Get-SettingStoreValue) so tenant-scoped overrides and the
        # registered default apply - matches how WPF reads AppTheme.
        $appTheme = Get-SettingValue "AppTheme" ""
        if ($appTheme) { $appTheme } else { Get-SettingStoreValue "" "ThemeVariant" "Default" }
    }
    # Resolve before handing it to the host: "Default" means "follow Windows", and
    # IntuneManagementApp.Initialize only understands concrete Light/Dark.
    $script:AvaloniaHostType::Initialize((Resolve-AppTheme $themeVariant))

    # Initialize() drops assemblies that Avalonia's XAML type-system walk cannot
    # enumerate (see Host.SanitizeXamlTypeSystem). Surface what was dropped so a
    # future "type not found in XAML" report has the list in the log.
    $dropped = $script:AvaloniaHostType::SanitizeXamlTypeSystem()
    if ($dropped -and $dropped.Count -gt 0) {
        Write-LogDebug "XAML type system: skipped $($dropped.Count) unreadable assemblies ($($dropped -join ', '))"
    }

    # Theme load order:
    #   1. Brushes.axaml  — ThemeDictionaries (Light/Dark) -> App.Resources.MergedDictionaries
    #   2. Default.axaml  — Style selectors that DynamicResource against those keys
    #   3. Styles.axaml   — Named/class styles
    # Variant selection is automatic via Application.RequestedThemeVariant; no
    # conditional Dark.axaml load needed.
    $themesDir = Join-Path $script:AppUIRootFolder 'Themes'

    foreach ($leaf in 'Brushes.axaml','Default.axaml','Styles.axaml') {
        $path = Join-Path $themesDir $leaf
        if (Test-Path $path) {
            $script:AvaloniaHostType::LoadStyles($path)
        } else {
            Write-LogDebug "Theme file not found: $path"
        }
    }

    $script:AvaloniaInitialized = $true
}

function Get-AvaloniaHost {
    if (-not $script:AvaloniaInitialized) { Initialize-AvaloniaRuntime }
    return $script:AvaloniaHostType
}

#region XAML helpers (Avalonia)

function Get-XamlObject {
    param($FileName, [switch]$AddVariables, [switch]$AddStyles)

    $ui = $script:UIProvider
    if (-not [IO.File]::Exists($FileName)) {
        Write-Log "Failed to open Xaml file. File not found: $FileName" 3
        return $null
    }

    $xamlText = [IO.File]::ReadAllText($FileName)
    $hostType = Get-AvaloniaHost
    try {
        $obj = $hostType::LoadXaml($xamlText, ([Uri]("file:///" + ($FileName -replace '\\','/'))).AbsoluteUri)
    } catch {
        $inner = $_.Exception
        while ($inner.InnerException) { $inner = $inner.InnerException }
        Write-LogError "Failed to load Avalonia XAML $FileName. Error:" $_.Exception
        throw "Avalonia XAML load failed for $FileName`: $($inner.GetType().FullName): $($inner.Message)"
    }

    if ($obj -and $AddVariables) {
        $ui.AddXamlVariables($xamlText, $obj)
    }

    return $obj
}

function Add-XamlVariables {
    param($XamlText, $Obj)

    try {
        [xml]$Xaml = $XamlText
        $Xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
            $name = $_.Name
            if (-not $name) { return }
            $found = (Get-AvaloniaHost)::FindByName($Obj, $name)
            New-Variable -Name $name -Value $found -Force -Scope Script
        }
    } catch {
        Write-LogDebug "Add-XamlVariables (Avalonia) skipped: $($_.Exception.Message)"
    }
}

function Set-XamlProperty {
    param($XamlObj, $ControlName, $PropertyName, $Value)

    try {
        $Obj = (Get-AvaloniaHost)::FindByName($XamlObj, $ControlName)
        if ($Obj) {
            $Obj.$PropertyName = $Value
        } else {
            Write-Log "Could not find object with name $ControlName" 3
        }
    } catch {
        Write-LogError "Failed to set Avalonia property. Control: $ControlName Property: $PropertyName Error:" $_.Exception
    }
}

function Get-XamlProperty {
    param($XamlObj, $ControlName, $PropertyName, $DefaultValue = $null)

    try {
        $Obj = (Get-AvaloniaHost)::FindByName($XamlObj, $ControlName)
        if (-not $Obj) {
            Write-Log "Could not find object with name $ControlName" 3
            return $DefaultValue
        }
        $val = $Obj.$PropertyName
        if ([String]::IsNullOrEmpty($val)) { return $DefaultValue }
        return $val
    } catch {
        Write-LogError "Failed to read Avalonia property. Control: $ControlName Property: $PropertyName Error:" $_.Exception
        return $DefaultValue
    }
}

function Add-XamlEvent {
    param($XamlObj, [string]$ControlName, [string]$EventName, [scriptblock]$ScriptBlock)

    try {
        $Obj = (Get-AvaloniaHost)::FindByName($XamlObj, $ControlName)
        if (-not $Obj) {
            Write-Log "Failed to add event $EventName to $ControlName. Control not found" 3
            return
        }
        # WPF tolerates `$Obj."Add_Click"($scriptBlock)` — the dynamic member
        # invoke binds the scriptblock as a RoutedEventHandler via PowerShell's
        # auto-conversion. In Avalonia the same path silently no-ops: the call
        # returns without subscribing, so click handlers attached this way
        # never fire. Resolve the accessor explicitly and cast the scriptblock
        # to the event delegate type before invoking so the subscription
        # actually lands. Direct add_X({...}) calls elsewhere keep working
        # untouched because they go through a different binder path.
        $accessorName = $EventName
        if ($accessorName -match '^Add_') { $accessorName = 'add_' + $accessorName.Substring(4) }
        $method = $Obj.GetType().GetMethod($accessorName, [Reflection.BindingFlags]'Public,Instance')
        if ($method) {
            # Bind the scriptblock to this module's session state before casting
            # to a delegate. The PowerShell -> delegate cast otherwise runs the
            # scriptblock against the runspace's CURRENT session state at click
            # time (not the module's), which makes $script:* lookups return $null
            # AND breaks bare module-private function calls like Save-SettingStoreValue,
            # Get-PoliciesFromFolder, Close-TopModalObject, Write-Log, etc.
            # NewBoundScriptBlock rebinds without disturbing local-variable
            # captures from GetNewClosure().
            $ScriptBlock = ConvertTo-AvaloniaEventScriptBlock $ScriptBlock
            $delegateType = $method.GetParameters()[0].ParameterType
            $delegate = $ScriptBlock -as $delegateType
            if (-not $delegate) {
                Write-Log "Failed to bind $EventName scriptblock to $($delegateType.FullName) on $ControlName" 3
                return
            }
            [void]$method.Invoke($Obj, @($delegate))
        } else {
            $Obj."$EventName"($ScriptBlock)
        }
    } catch {
        Write-LogError "Failed to add Avalonia event $EventName to $ControlName. Error:" $_.Exception
    }
}

# Wraps a scriptblock so it runs against IntuneManagement's session state when
# invoked as an Avalonia event delegate. Use at every Avalonia event-subscription
# call site — Add-XamlEvent calls it automatically, direct .add_X({...}) sites
# should wrap their scriptblock with this helper:
#
#     $btn.add_Click( (ConvertTo-AvaloniaEventScriptBlock {
#         ...
#     }.GetNewClosure()) )
#
# Without the wrap, $script:UIProvider, $script:Window, and bare module-private
# function calls (Save-SettingStoreValue, Close-TopModalObject, Get-PoliciesFromFolder,
# Get-MigrationTableInfo, ...) all fail at click time because the scriptblock
# is invoked against the runspace's CURRENT session state, not the module's.
function ConvertTo-AvaloniaEventScriptBlock {
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock)

    $module = Get-Module IntuneManagement
    if (-not $module) { return $ScriptBlock }
    return $module.NewBoundScriptBlock($ScriptBlock)
}

#endregion

#region App lifecycle (Avalonia)

function Initialize-UI {
    Initialize-AvaloniaRuntime
}

function Show-MainWindow {
    param($View)

    $ui = $script:UIProvider
    Initialize-AvaloniaRuntime

    $mainXaml = Join-Path $script:AppUIRootFolder 'XAML/MainWindow.axaml'
    $script:Window = $ui.GetXamlObject($mainXaml, $true)

    if (-not $script:Window) {
        throw "Failed to load Avalonia MainWindow.axaml"
    }

    # The XAML is authored for the Windows chrome (extended client area + system
    # caption buttons). Linux/macOS need that adapted before the window is shown.
    Set-MainWindowChrome $script:Window

    # Maximized windows with ExtendClientAreaToDecorationsHint hang over the
    # screen edge by the resize-border thickness plus the invisible "padded
    # border" (standard Win32 behaviour - normally the OS chrome absorbs it,
    # but with extended client area the CONTENT gets clipped a few pixels all
    # around). Window.OffScreenMargin under-reports by the padded border, so
    # compute the real overhang from SM_CXSIZEFRAME(32) + SM_CXPADDEDBORDER(92)
    # at the window's DPI and use whichever is larger.
    if ($script:IsWindowsOS -and -not ("IMMaximizeMetrics" -as [type])) {
        Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public static class IMMaximizeMetrics
{
    [DllImport("user32.dll")] private static extern int GetSystemMetricsForDpi(int nIndex, uint dpi);
    [DllImport("user32.dll")] private static extern int GetSystemMetrics(int nIndex);
    public static int GetOverhangPixels(double scaling)
    {
        try { uint dpi = (uint)(96.0 * scaling); return GetSystemMetricsForDpi(32, dpi) + GetSystemMetricsForDpi(92, dpi); }
        catch { return GetSystemMetrics(32) + GetSystemMetrics(92); }
    }
}
'@
    }
    $script:Window.add_PropertyChanged({
        param($S, $E)
        # Win32-only: IMMaximizeMetrics P/Invokes user32, and off-Windows there is
        # no extended client area to correct for (Set-MainWindowChrome disabled it).
        if (-not $script:IsWindowsOS) { return }
        if ($E.Property.Name -notin @('WindowState', 'OffScreenMargin')) { return }
        try {
            if (-not $S.Content) { return }
            if ($S.WindowState -eq [Avalonia.Controls.WindowState]::Maximized) {
                $dip = [IMMaximizeMetrics]::GetOverhangPixels($S.RenderScaling) / $S.RenderScaling
                $dip = [Math]::Max($dip, $S.OffScreenMargin.Left)
                $S.Content.Margin = [Avalonia.Thickness]::new($dip)
            }
            else {
                $S.Content.Margin = $S.OffScreenMargin
            }
        } catch { }
    })

    # Capture overlay refs that Update-UIStatus / Show-Popup poke at by
    # script-scoped name. Add-XamlVariables (the -AddVariables sweep) only
    # synthesizes script vars when PowerShell's XML adapter projects the
    # Name attribute as a property, which it doesn't on every node — keep
    # the explicit assignments to mirror the WPF Show-MainWindow flow.
    $hostType                 = Get-AvaloniaHost
    $script:grdStatus         = $hostType::FindByName($script:Window, 'grdStatus')
    $script:txtInfo           = $hostType::FindByName($script:Window, 'txtInfo')
    $script:txtInfoDetail     = $hostType::FindByName($script:Window, 'txtInfoDetail')
    # Wired once here rather than per Write-Status call: the handler is fixed, only
    # the armed action changes (Internal/StatusCancel.ps1 holds that). Calling the
    # module function BY NAME is also what keeps this legal - a handler built by
    # ConvertTo-AvaloniaEventScriptBlock loses function-local captures.
    $script:btnStatusCancel   = $hostType::FindByName($script:Window, 'btnStatusCancel')
    if($script:btnStatusCancel) {
        $script:btnStatusCancel.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            try {
                Write-Log "Avalonia status Cancel click dispatched"
                $S.IsEnabled = $false
                $S.Content = "Cancelling..."
                Request-StatusCancel
            }
            catch {
                Write-LogError "Avalonia status Cancel handler failed" $_.Exception
            }
        }))
    }

    Set-CacheObject "ShowUI" $true

    # Discover ViewObjectBase subclasses, instantiate as singletons, register
    # the ones that opt into the menu. Mirrors the WPF Show-MainWindow flow.
    $allViews = @(Get-SubClasses ([ViewObjectBase]) | ForEach-Object {
        try { Get-SingletonObject $_ } catch { Write-LogError "Failed to instantiate view $($_.Name):" $_.Exception }
    } | Where-Object { $null -ne $_ })

    foreach ($vo in ($allViews | Where-Object { $_.AddToMenu })) {
        # The Avalonia Preview sandbox view is a porting aid, not a product
        # view - keep it out of the Views menu unless explicitly requested.
        if ($vo -is [PreviewViewObject] -and $env:IM_AVALONIA_PREVIEW -ne '1') { continue }
        Add-ViewObject $vo
    }

    Build-ViewsMenu

    # Wire chrome handlers that need no view context: Exit closes the window
    # and ends the dispatcher loop. Other File-menu items remain stubs until
    # their dialogs are ported.
    $ui.AddXamlEvent($script:Window, 'mnuExit', 'Add_Click', ({
        # WPF confirms before exiting; Avalonia closed immediately.
        if ((Show-MessageBox -Text "Are you sure you want to exit?" -Caption "Exit?" -Button "YesNo" -Icon "Question") -eq "Yes") {
            $script:Window.Close()
        }
    }))

    $ui.AddXamlEvent($script:Window, 'mnuAbout', 'Add_Click', ({
        try { $ui.ShowAboutDialog() } catch { Write-LogError "Show-AboutDialog failed" $_.Exception }
    }))

    $ui.AddXamlEvent($script:Window, 'mnuSettings', 'Add_Click', ({
        try { Show-SettingsForm } catch { Write-LogError "Show-SettingsForm failed" $_.Exception }
    }))

    $ui.AddXamlEvent($script:Window, 'mnuTenantSettings', 'Add_Click', ({
        try { Show-SettingsForm -Tenant } catch { Write-LogError "Show-SettingsForm -Tenant failed" $_.Exception }
    }))

    $ui.AddXamlEvent($script:Window, 'mnuUpdates', 'Add_Click', ({
        try { Show-ReleaseNotes } catch { Write-LogError 'Show-ReleaseNotes failed' $_.Exception }
    }))

    # Tenant Settings depends on an active session — disable it when signed
    # out. Re-evaluate on File-menu open (covers the steady state) and on auth
    # events (so the user doesn't have to close + reopen File after sign-in
    # to see the state flip). Mirrors WPF Update-AuthDependentMenuState wire.
    $ui.AddXamlEvent($script:Window, 'mnuFile', 'Add_SubmenuOpened', ({
        try { Update-AuthDependentMenuState } catch { Write-LogError 'Update-AuthDependentMenuState failed on submenu open' $_.Exception }
    }))
    Add-AppEventHandler 'AuthenticatedNewToken'          'Update-AuthDependentMenuState'
    Add-AppEventHandler 'AuthenticationUserDisconnected' 'Update-AuthDependentMenuState'

    # Title-bar + taskbar icon. Asset is copied locally (UI/Avalonia/Assets/)
    # rather than referenced from UI/WPF/ — keeps the port self-contained.
    Set-MainWindowIcon

    # Apply the immersive-dark-mode title bar to match the saved AppTheme.
    # Window.Opened fires after the HWND exists; calling Set-WindowTitleBarTheme
    # before that is a no-op (TryGetPlatformHandle returns null). Window.Opened
    # is also where the first-run welcome modal goes — before then the modal
    # host (grdModal) hasn't been laid out so the overlay can't display.
    $script:Window.add_Opened((ConvertTo-AvaloniaEventScriptBlock {
        param($S, $E)
        $theme = Get-SettingValue "AppTheme" "Default"
        Set-WindowTitleBarTheme $theme

        if ((Get-SettingStoreValue "" "FirstTimeRunning" "true") -eq "true") {
            try { Show-WelcomeDialog } catch { Write-LogError 'Show-WelcomeDialog failed' $_.Exception }
        }

        # Stacks on top of the welcome modal on a first run off Windows, so the
        # experimental warning is the first thing read and the first dismissed.
        try { Show-ExperimentalPlatformNotice } catch { Write-LogError 'Show-ExperimentalPlatformNotice failed' $_.Exception }
    }))

    # Tell Windows this process is the Intune Management app so the taskbar
    # button gets grouped under its own AUMID instead of being merged with
    # "Windows PowerShell". Must run before the window is shown.
    Set-TaskbarAppId

    $ui.AddXamlEvent($script:Window, 'borderTitleBar', 'Add_PointerPressed', ({
        param($s, $e)
        # Skip drag/maximize when the click originated on an interactive control
        # inside the title bar. Avalonia's Menu/MenuItem don't mark
        # PointerPressed as handled when bubbling, so without this filter the
        # File-menu submenu would close immediately as soon as BeginMoveDrag
        # starts a window drag.
        $src = $e.Source
        while ($src -and $src -ne $s) {
            if ($src -is [Avalonia.Controls.Button] -or
                $src -is [Avalonia.Controls.MenuItem] -or
                $src -is [Avalonia.Controls.Menu]) {
                return
            }
            $src = $src.Parent
        }

        if ($e.ClickCount -eq 2) {
            $script:Window.WindowState = if ($script:Window.WindowState -eq [Avalonia.Controls.WindowState]::Maximized) {
                [Avalonia.Controls.WindowState]::Normal
            } else {
                [Avalonia.Controls.WindowState]::Maximized
            }
        } else {
            $script:Window.BeginMoveDrag($e)
        }
    }))

    # Selecting a left-nav item routes through the active view, mirroring the
    # WPF wire-up. Views that don't override OnItemChanged inherit the
    # ViewObjectBase no-op, so this is safe even before specific views are
    # ported.
    $ui.AddXamlEvent($script:Window, 'lstMenuItems', 'Add_SelectionChanged', ({
        param($S, $E)
        # Suppressed while the Intune view's count-label refresh rebinds the
        # ItemsSource and restores the selection — without this guard, the
        # restore would retrigger a full Graph reload of the current type.
        if ($script:_suppressMenuSelectionEvents) { return }
        # Show-ViewMenu interleaves Category-header rows in the items list to
        # mimic WPF's <ListBox.GroupStyle>. Skip them — the ListBoxItem style
        # already disables hit-test + focus on headers, but a programmatic
        # SelectedItem assignment (e.g., Add-ViewMenu's "select by Id" loop)
        # could still land on one.
        $sel = $S.SelectedItem
        if ($sel -and $sel.PSObject.Properties['IsHeader'] -and $sel.IsHeader) { return }
        if ($script:ActiveView) {
            try {
                $script:ActiveView.OnItemChanged($sel) | Out-Null
            } catch {
                Write-LogError "OnItemChanged threw on view $($script:ActiveView.Id):" $_.Exception
            }
        }
    }))

    Show-View $View

    # Startup update check, matching the WPF backend. The version comparison
    # itself lives in Internal/AppUpdateCheck.ps1 so both backends share one
    # implementation; only the notification differs. Wrapped because a failed
    # network call must never stop the app from starting.
    if ((Get-SettingValue 'CheckForUpdates') -eq $true) {
        try {
            $updateInfo = Get-AppUpdateInfo
            if ($updateInfo.IsOutdated) {
                Show-MessageBox ("There is a new version available on GitHub $($updateInfo.RemoteVersion.ToString())" +
                    [Environment]::NewLine + [Environment]::NewLine +
                    "Current version is $($updateInfo.LocalVersion.ToString())") 'Old version!' 'OK' 'Warning' | Out-Null
            }
        }
        catch { Write-LogError 'Update check failed' $_.Exception }
    }

    # Populate the title-bar profile slot before showing the window. WPF does
    # this in Window.Loaded; we call it pre-RunMainWindow so the avatar /
    # Sign-in button is already in grdMenu when the window appears (no flash
    # of empty title bar). Get-MSALUserInfo handles all three startup states
    # and internally calls Show-AuthenticationInfo + Set-EnvironmentInfo.
    if (Get-Command Get-MSALUserInfo -ErrorAction SilentlyContinue) {
        try { Get-MSALUserInfo } catch { Write-LogError 'Initial Get-MSALUserInfo failed' $_.Exception }
    } else {
        # Fallback in case the auth UI file didn't load — at least render the
        # Sign-in button so the user has *something* to click.
        try { $ui.ShowAuthenticationInfo() } catch { Write-LogError 'Initial Show-AuthenticationInfo failed' $_.Exception }
    }

    # Blocks running the dispatcher loop on this thread until the last
    # window closes. This is Avalonia's normal app entry point.
    (Get-AvaloniaHost)::RunMainWindow($script:Window) | Out-Null
}

#region View system (Avalonia)

function Get-ViewObject {
    param([string]$ViewId)

    $viewObject = $script:viewObjects | Where-Object { $_.Id -eq $ViewId }
    if (-not $viewObject) {
        Write-Log "Could not find View with id $ViewId" 3
        return
    }
    $viewObject
}

function Add-ViewObject {
    param([ViewObjectBase]$ViewObject)

    if ($ViewObject) {
        $script:viewObjects += $ViewObject
    } else {
        Write-Log "Add-ViewObject called with empty ViewObject"
    }
}

function Get-CurrentViewObject {
    $script:ActiveView
}

function Add-MenuItem {
    param($MenuItem, $Index)

    if (-not $script:mnuMain) {
        Write-Log "Add-MenuItem: mnuMain not available yet" 3
        return
    }
    $script:mnuMain.Items.Insert($Index, $MenuItem) | Out-Null
}

function Show-View {
    param($ViewId = "IntuneManagement")

    $ui = $script:UIProvider
    if (($script:viewObjects | Measure-Object).Count -eq 0) {
        Write-Log "No View Objects loaded!" 3
        return
    }

    if (-not $ViewId) {
        $ViewId = $script:viewObjects[0].Id
    }

    if ($script:ActiveView -and $script:ActiveView.ID -eq $ViewId) { return }

    $viewObject = Get-ViewObject $ViewId
    if (-not $viewObject) {
        # Common during the port: callers pass the WPF default
        # "IntuneManagement" which Avalonia hasn't ported yet. Prefer
        # IntuneTools (real Intune surface) over PreviewViewObject (sandbox
        # demo) so the first impression is the actual app, not the demo.
        $preferred = @('IntuneTools')
        foreach ($pref in $preferred) {
            $viewObject = $script:viewObjects | Where-Object { $_.Id -eq $pref } | Select-Object -First 1
            if ($viewObject) { break }
        }
        if (-not $viewObject) { $viewObject = $script:viewObjects[0] }
        Write-Log "Falling back to view: $($viewObject.Id)" 2
    }

    Write-Log "Change view to $($viewObject.Title)"

    if ($script:ActiveView) {
        Write-LogDebug "Deactivating View $($script:ActiveView.Title)"
        $script:ActiveView.OnDeactivating($viewObject)
    }

    $previousView = $script:ActiveView
    $script:ActiveView = $viewObject

    Show-ViewMenu

    if ($script:lblMenuTitle) {
        $script:lblMenuTitle.Content = $viewObject.Title
    }

    if ($script:grdViewPanel) {
        $script:grdViewPanel.Children.Clear()
    }

    Write-LogDebug "Activating View $($viewObject.Title)"
    $viewObject.OnActivating($previousView)

    $panel = $viewObject.ViewPanel
    if ($panel -and $script:grdViewPanel) {
        $script:grdViewPanel.Children.Add($panel) | Out-Null
    }

    $ui.SetMainTitle()

    # WPF Visibility -> Avalonia IsVisible (boolean).
    if ($script:grdViewItemMenu) {
        $script:grdViewItemMenu.IsVisible = -not $viewObject.HideMenu
    }

    $viewObject.OnActivated()
}

function Build-ViewsMenu {
    # Populates the title-bar Views submenu with one MenuItem per registered
    # view (clicking → Show-View). Views with ExpandInViewsMenu also surface
    # their items as nested children (click → activate view + select the row
    # in the left nav). Direct port of the WPF builder in
    # UI/WPF/Extensions/CoreUI.ps1 ~1011-1098.
    if (-not $script:mnuViews) {
        $script:mnuViews = (Get-AvaloniaHost)::FindByName($script:Window, 'mnuViews')
    }
    if (-not $script:mnuViews) { return }

    $script:mnuViews.Items.Clear()

    $pinTopId     = 'IntuneManagement'
    $pinBottomIds = @('CoreLog','GraphCalls','CoreCachedObjects')

    $top    = @($script:viewObjects | Where-Object { $_.Id -eq $pinTopId })
    $middle = @($script:viewObjects | Where-Object { $_.Id -ne $pinTopId -and $_.Id -notin $pinBottomIds } | Sort-Object Title)
    $bottom = @(
        foreach ($id in $pinBottomIds) { $script:viewObjects | Where-Object { $_.Id -eq $id } }
    ) | Where-Object { $_ }

    $addViewMenuItem = {
        param($view)
        $subItem = [Avalonia.Controls.MenuItem]::new()
        $subItem.Header = $view.Title
        $subItem.Tag    = $view.Id
        $subItem.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            if ($S.Tag) { Show-View $S.Tag }
        }))

        if ($view.ExpandInViewsMenu) {
            try {
                $childItems = @($view.GetViewItems())
                if ($childItems.Count -gt 0) {
                    # Sort by Category then Title so the visual layout matches
                    # the left-nav (which Show-ViewMenu also sorts the same way).
                    $sorted = $childItems | Sort-Object @{Expression='Category'}, @{Expression='Title'}
                    $lastCat = $null
                    foreach ($child in $sorted) {
                        $cat = if ($child.Category) { [string]$child.Category } else { $null }
                        if ($null -ne $lastCat -and $cat -ne $lastCat) {
                            [void]$subItem.Items.Add([Avalonia.Controls.Separator]::new())
                        }
                        $lastCat = $cat

                        $childMenu = [Avalonia.Controls.MenuItem]::new()
                        $childMenu.Header = $child.Title
                        $childMenu.Tag    = [PSCustomObject]@{ ViewId = $view.Id; ItemId = $child.Id }
                        $childMenu.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                            param($S, $E)
                            $info = $S.Tag
                            if (-not $info) { return }
                            Show-View $info.ViewId
                            if ($script:lstMenuItems) {
                                foreach ($it in @($script:lstMenuItems.Items)) {
                                    if ($it -and $it.Id -eq $info.ItemId) {
                                        $script:lstMenuItems.SelectedItem = $it
                                        break
                                    }
                                }
                            }
                        }))
                        [void]$subItem.Items.Add($childMenu)
                    }
                }
            } catch {
                Write-LogDebug "Views menu sub-item build failed for $($view.Id): $($_.Exception.Message)"
            }
        }

        $script:mnuViews.Items.Add($subItem) | Out-Null
    }

    foreach ($view in $top)    { & $addViewMenuItem $view }
    foreach ($view in $middle) { & $addViewMenuItem $view }

    if ($bottom.Count -gt 0) {
        $hasAbove = ($top.Count -gt 0) -or ($middle.Count -gt 0)
        if ($hasAbove) {
            $script:mnuViews.Items.Add([Avalonia.Controls.Separator]::new()) | Out-Null
        }
        foreach ($view in $bottom) { & $addViewMenuItem $view }
    }
}

function Show-ViewMenu {
    if (-not $script:lstMenuItems) { return }

    $items = @($script:ActiveView.GetViewItems())

    # WPF used ListCollectionView + PropertyGroupDescription("Category") to
    # render an expandable-group tree under <ListBox.GroupStyle>. Avalonia's
    # ListBox has no GroupStyle equivalent (Avalonia 11), so we emulate the
    # WPF appearance by interleaving non-selectable [ViewMenuItem] header
    # rows (IsHeader=$true) between the leaves of each category. The
    # MainWindow.axaml DataTemplate branches on IsHeader to render either
    # the SemiBold category title or the icon+MenuLabel leaf. We lose the
    # WPF expand/collapse affordance, but WPF defaulted to IsExpanded=True
    # so the visual at rest is the same.
    $hasCategory = $false
    foreach ($it in $items) {
        if ($it -and $it.Category) { $hasCategory = $true; break }
    }

    if ($hasCategory) {
        # Build header+leaf interleaved list. Stable category ordering:
        # categories appear in the order their first item was returned by
        # GetViewItems, then alphabetical within each. Empty Category items
        # land in a leading "(Other)" group so nothing disappears silently.
        $grouped = [ordered]@{}
        foreach ($it in $items) {
            if (-not $it) { continue }
            $cat = if ($it.Category) { [string]$it.Category } else { '(Other)' }
            if (-not $grouped.Contains($cat)) {
                $grouped[$cat] = New-Object 'System.Collections.Generic.List[object]'
            }
            $grouped[$cat].Add($it)
        }
        $flat = New-Object 'System.Collections.Generic.List[object]'
        # Snapshot keys so we don't iterate the live OrderedDictionary key
        # collection. Use ::new() + property assignment instead of the
        # [ViewMenuItem]@{...} hashtable cast — the cast resolves overloads
        # dynamically on every call and intermittently throws
        # "Argument types do not match" when a property has a default value
        # ([bool]$IsHeader = $false) but the hashtable supplies a different
        # value. Explicit assignment sidesteps that.
        foreach ($cat in @($grouped.Keys)) {
            $header = [ViewMenuItem]::new()
            $header.Title    = [string]$cat
            $header.IsHeader = $true
            [void]$flat.Add($header)

            $sortedItems = $grouped[$cat] | Sort-Object -Property Title
            foreach ($it in $sortedItems) {
                [void]$flat.Add($it)
            }
        }
        $items = $flat.ToArray()
    }

    $script:lstMenuItems.ItemsSource = $items

    # Per-view title-config button. Default to hidden; opt-in views show +
    # wire it. Kept here (rather than in each view) so a view that doesn't
    # know about the button can't accidentally leave it visible from a
    # previous active view. Mirrors CoreUIWPF's Show-ViewMenu hook.
    if ($script:btnMenuTitleConfig) {
        $script:btnMenuTitleConfig.IsVisible = $false
    }
    if ($script:ActiveView -and $script:ActiveView.ID -eq 'IntuneManagement') {
        try { Update-AvaloniaMenuTitleConfigForIntuneView }
        catch { Write-LogDebug "Update-AvaloniaMenuTitleConfigForIntuneView failed: $($_.Exception.Message)" }
    }
}

function Show-WelcomeDialog {
    # First-run license-accept gate. Loads Welcome.axaml as a modal with hidden
    # default Close button (the form supplies its own OK / Cancel pair). The
    # checkbox toggles the OK button; OK saves LicenseAccepted=True and
    # FirstTimeRunning=False then dismisses; Cancel prompts to close the app.
    $ui = $script:UIProvider
    $welcome = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/Welcome.axaml'))
    if (-not $welcome) { return }

    # Open-ExternalUri by name - a $launchUri closure here would be stripped
    # by ConvertTo-AvaloniaEventScriptBlock and resolve to nothing at click
    # time (the documented .GetNewClosure() trap).
    $ui.AddXamlEvent($welcome, 'gitHubLink', 'Add_Click', ((ConvertTo-AvaloniaEventScriptBlock { Open-ExternalUri 'https://github.com/Micke-K/IntuneManagement' })))
    $ui.AddXamlEvent($welcome, 'licenseLink', 'Add_Click', ((ConvertTo-AvaloniaEventScriptBlock { Open-ExternalUri 'https://github.com/Micke-K/IntuneManagement/blob/master/LICENSE' })))

    $chkAccept = (Get-AvaloniaHost)::FindByName($welcome, 'chkAcceptConditions')
    $btnOk     = (Get-AvaloniaHost)::FindByName($welcome, 'btnAcceptConditions')
    if ($chkAccept -and $btnOk) {
        # Avalonia's CheckBox raises IsCheckedChanged (StyledProperty change);
        # Click also fires on toggle and is simpler from PS.
        # The OK button rides on the checkbox's Tag: a captured $btnOk local is
        # stripped by ConvertTo-AvaloniaEventScriptBlock, which left the accept
        # checkbox inert and the OK button permanently disabled - the first-run
        # dialog could never be completed.
        $chkAccept.Tag = $btnOk
        $chkAccept.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            if ($S.Tag) { $S.Tag.IsEnabled = [bool]$S.IsChecked }
        }))
    }

    $ui.AddXamlEvent($welcome, 'btnAcceptConditions', 'Add_Click', ({
        Save-SettingStoreValue "" "FirstTimeRunning" "False"
        Show-ModalObject
    }))

    $ui.AddXamlEvent($welcome, 'btnCancel', 'Add_Click', ({
        if (($ui.ShowMessageBox("Conditions not accepted`n`nDo you want to close the application?", 'Close App?', 'YesNo', 'Warning')) -eq 'Yes') {
            $script:Window.Close()
        }
    }))

    $ui.ShowModalForm($script:Window.Title, $welcome, $true)
}

function Show-ReleaseNotes {
    # Avalonia port of the WPF Show-UpdatesDialog. Prefers the published notes so
    # the dialog answers "what changed upstream", falling back to the local copy
    # when GitHub is unreachable. Fetch + version comparison live in
    # Internal/AppUpdateCheck.ps1, shared with WPF (R10). WPF renders the version
    # banner in a separate tab; here it is prepended as a header line, since the
    # Avalonia dialog is a single TextBox.
    $ui = $script:UIProvider
    $releaseNotesPath = Join-Path $script:AppRootFolder 'ReleaseNotes.md'
    $localNotes = if (Test-Path -LiteralPath $releaseNotesPath) {
        [IO.File]::ReadAllText($releaseNotesPath)
    } else {
        "ReleaseNotes.md not found at $releaseNotesPath."
    }

    $remote  = $null
    $header  = $null
    try {
        $remote = Get-AppRemoteReleaseNotes
        $info   = Get-AppUpdateInfo
        if ($info.IsOutdated) {
            $header = "A newer version is available on GitHub: $($info.RemoteVersion.ToString()) (installed: $($info.LocalVersion.ToString()))"
        }
        elseif ($info.Resolved) {
            $header = "Running the latest version: $($info.LocalVersion.ToString())"
        }
    }
    catch { Write-LogError 'Failed to get release notes from GitHub' $_.Exception }

    $body = if ($remote -and $remote.Text) { $remote.Text } else { $localNotes }
    $content = if ($header) {
        $header + [Environment]::NewLine + ('-' * 72) + [Environment]::NewLine + [Environment]::NewLine + $body
    } else {
        $body
    }

    # Rendered rather than shown as raw text: Set-AvaloniaMarkdownText turns the
    # headings, bullets and **bold** spans into TextBlock inlines. See
    # UI/Avalonia/Extensions/MarkdownRenderAvalonia.ps1. A TextBlock (not TextBox)
    # is required - only TextBlock exposes Inlines - so it needs an explicit
    # ScrollViewer around it, unlike the TextBox this replaced.
    $tb = [Avalonia.Controls.TextBlock]::new()
    $tb.TextWrapping = [Avalonia.Media.TextWrapping]::Wrap
    $tb.MaxWidth     = 900
    Set-AvaloniaMarkdownText $tb $content

    $scroll = [Avalonia.Controls.ScrollViewer]::new()
    $scroll.Content   = $tb
    $scroll.MinWidth  = 700
    $scroll.MinHeight = 500
    $scroll.VerticalScrollBarVisibility   = [Avalonia.Controls.Primitives.ScrollBarVisibility]::Auto
    $scroll.HorizontalScrollBarVisibility = [Avalonia.Controls.Primitives.ScrollBarVisibility]::Disabled

    $ui.ShowModalForm('Release Notes', $scroll)
}

function Set-MainWindowIcon {
    if (-not $script:Window) { return }
    $iconPath = Join-Path $script:AppUIRootFolder 'Assets/intune.png'
    if (-not (Test-Path -LiteralPath $iconPath)) {
        Write-LogDebug "Window icon not found at $iconPath"
        return
    }
    try {
        $bitmap = [Avalonia.Media.Imaging.Bitmap]::new($iconPath)
        # Window.Icon expects an Avalonia.Controls.WindowIcon (constructed from a Bitmap).
        $script:Window.Icon = [Avalonia.Controls.WindowIcon]::new($bitmap)

        if ($script:imgTitleIcon) {
            $script:imgTitleIcon.Source = $bitmap
        }
    } catch {
        Write-LogError 'Failed to load window icon' $_.Exception
    }
}

function Update-AuthDependentMenuState {
    # Toggles Tenant Settings IsEnabled based on whether any auth provider has
    # a user. Called from File-menu open and from AuthenticatedNewToken /
    # AuthenticationUserDisconnected so the state stays current. EventArg is
    # the AppEvent payload — unused, but the handler signature accepts it.
    [CmdletBinding()]
    param($EventArg)

    $ui = $script:UIProvider
    if (-not $script:Window) { return }

    $signedIn = $false
    try {
        $provider = Get-AuthProvider
        if ($provider -and $provider.GetUserInfo(0)) { $signedIn = $true }
    } catch { }

    try { $ui.SetXamlProperty($script:Window, 'mnuTenantSettings', 'IsEnabled', $signedIn) }
    catch { Write-LogDebug "Update-AuthDependentMenuState: $($_.Exception.Message)" }
}

function Set-TaskbarAppId {
    # Direct port of the WPF AUMID logic. Without this the taskbar groups the
    # window under powershell.exe and uses the PowerShell icon regardless of
    # WM_SETICON. SetCurrentProcessExplicitAppUserModelID must be called
    # before the window's HWND is created (i.e. before RunMainWindow).
    # shell32.dll is Windows-only; the X11/Cocoa backends group windows by their
    # own .desktop/app identity, so this is a no-op elsewhere.
    if (-not $script:IsWindowsOS) { return }
    if (-not ([System.Management.Automation.PSTypeName]'AppUserModelHelper').Type) {
        try {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class AppUserModelHelper {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern int SetCurrentProcessExplicitAppUserModelID(string AppID);
}
"@ -ErrorAction Stop
        } catch {
            Write-LogDebug "AUMID type compile failed: $($_.Exception.Message)"
            return
        }
    }
    try {
        [void][AppUserModelHelper]::SetCurrentProcessExplicitAppUserModelID('Intune.IntuneManagement')
    } catch {
        Write-LogDebug "AUMID set failed: $($_.Exception.Message)"
    }
}

function Set-AppTheme {
    # Runtime theme variant switch — called from the Settings dialog when the
    # user changes the AppTheme preference. Avalonia's theme system reacts to
    # Application.RequestedThemeVariant; setting it re-resolves every
    # {DynamicResource ...} that points at Brushes.axaml's ThemeDictionaries.
    # Also re-applies the Win32 immersive-dark-mode title bar so the OS-painted
    # caption buttons match (without this they stay light on the dark title strip).
    param([string]$ThemeName)

    # An empty name is NOT a no-op: fall through so Resolve-AppTheme reads the
    # stored AppTheme setting. Returning here used to swallow the call silently
    # whenever a caller lost the argument, applying nothing and logging nothing.
    if (-not $ThemeName) {
        Write-LogDebug "Set-AppTheme called with no theme name, falling back to the AppTheme setting"
    }

    $app = [Avalonia.Application]::Current
    if (-not $app) {
        Write-LogDebug "Set-AppTheme: no Avalonia Application (headless), nothing to theme"
        return
    }

    # Resolve "Default" to a concrete Light/Dark from the Windows app theme rather
    # than handing Avalonia ThemeVariant::Default. Brushes.axaml only defines Light
    # and Dark dictionaries, so ThemeVariant::Default depends on toolkit fallback
    # behaviour instead of the user's OS setting - and would disagree with WPF,
    # which resolves the same setting through Resolve-AppTheme.
    $resolved = Resolve-AppTheme $ThemeName

    $variant = if ($resolved -ieq 'Dark') {
        [Avalonia.Styling.ThemeVariant]::Dark
    } else {
        [Avalonia.Styling.ThemeVariant]::Light
    }

    try { $app.RequestedThemeVariant = $variant }
    catch { Write-LogError "Set-AppTheme failed for '$ThemeName' (resolved '$resolved')" $_.Exception }

    # The Win32 caption buttons follow the RESOLVED theme, not the setting name -
    # passing "Default" here would leave light caption buttons on a dark strip.
    Set-WindowTitleBarTheme $resolved
}

function Set-WindowTitleBarTheme {
    # Toggle the Win32 immersive-dark-mode flag on the window's title bar so the
    # OS-painted caption buttons (min/max/close — kept via PreferSystemChrome)
    # render in the right palette for our theme. Silent no-op on pre-1903 OS
    # builds (DwmSetWindowAttribute returns non-zero; we ignore it).
    param([string]$ThemeName)

    if (-not $script:Window) { return }
    # dwmapi.dll is Windows-only. On X11/Cocoa the caption buttons aren't
    # OS-painted via DWM, so there's nothing to toggle.
    if (-not $script:IsWindowsOS) { return }

    try {
        if (-not ([System.Management.Automation.PSTypeName]'DwmTitleBarHelper').Type) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DwmTitleBarHelper {
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
}
"@ -ErrorAction Stop
        }

        # HWND is only valid after Window.Opened. Caller wires the initial
        # invocation via the Opened event; subsequent Set-AppTheme calls land
        # here while the window is alive so the handle is always populated.
        $handle = $script:Window.TryGetPlatformHandle()
        if (-not $handle) { return }
        $hwnd = $handle.Handle
        if ($hwnd -eq [IntPtr]::Zero) { return }

        $darkMode = if ($ThemeName -ieq 'Dark') { 1 } else { 0 }
        [void][DwmTitleBarHelper]::DwmSetWindowAttribute(
            $hwnd,
            [DwmTitleBarHelper]::DWMWA_USE_IMMERSIVE_DARK_MODE,
            [ref]$darkMode,
            4)
    } catch {
        Write-LogDebug "Set-WindowTitleBarTheme failed: $($_.Exception.Message)"
    }
}

function Set-MainTitle {
    param($Title)

    if (-not $Title -and $script:ActiveView) {
        $Title = $script:ActiveView.Title
    }
    if (-not $Title) { return }

    # Set both the centered title-bar label (visible) and the OS window title
    # (taskbar caption / Alt+Tab text). WPF Set-MainTitle keeps these in sync.
    if ($script:Window)            { $script:Window.Title = [string]$Title }
    if ($script:txtTitleViewName)  { $script:txtTitleViewName.Text = [string]$Title }
}

function Get-MainWindow {
    $script:Window
}

#endregion

function Show-InputDialog {
    param(
        $FormTitle = "Input",
        $FormText,
        $DefaultValue
    )

    $ui = $script:UIProvider
    Initialize-AvaloniaRuntime

    $dialogXaml = Join-Path $script:AppUIRootFolder 'XAML/InputDialog.axaml'
    $dialog = $ui.GetXamlObject($dialogXaml)
    if (-not $dialog) { return }

    $dialog.Title = $FormTitle
    $ui.SetXamlProperty($dialog, 'txtLabel', 'Content', $FormText)
    $ui.SetXamlProperty($dialog, 'txtValue', 'Text', $DefaultValue)

    $txtValue = (Get-AvaloniaHost)::FindByName($dialog, 'txtValue')

    # Module scope: handler captures are stripped, so OK/Cancel could not
    # close the dialog and Cancel could not blank the value.
    $script:_inputDialog    = $dialog
    $script:_inputDialogTxt = $txtValue

    $ui.AddXamlEvent($dialog, 'btnOk', 'Add_Click', ({
        if ($script:_inputDialog) { $script:_inputDialog.Close() }
    }))
    $ui.AddXamlEvent($dialog, 'btnCancel', 'Add_Click', ({
        if ($script:_inputDialogTxt) { $script:_inputDialogTxt.Text = '' }
        if ($script:_inputDialog) { $script:_inputDialog.Close() }
    }))

    (Get-AvaloniaHost)::ShowDialog($dialog, $script:Window)
    $value = if ($txtValue) { $txtValue.Text } else { '' }
    $script:_inputDialog    = $null
    $script:_inputDialogTxt = $null
    return $value
}

#endregion

#region Modal host (Avalonia)

# WPF used WindowsForms Application.DoEvents() to flush the message pump after
# mounting/unmounting modal content. Avalonia's dispatcher loop runs
# continuously on this thread (see Initialize-AvaloniaRuntime), so the visual
# tree updates without an explicit pump call.

function Show-ModalForm
{
    param(
        $FormTitle = "",
        $FormObject,
        [switch]$HideButtons)

    $ui = $script:UIProvider
    $modalForm = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/ModalForm.axaml'))
    if (-not $modalForm) { return }

    if ($HideButtons) {
        $ui.SetXamlProperty($modalForm, 'spButtons', 'IsVisible', $false)
    } else {
        $closeButton = (Get-AvaloniaHost)::FindByName($modalForm, 'btnClose')
        if ($closeButton) { $closeButton.Tag = $FormObject }
        $ui.AddXamlEvent($modalForm, 'btnClose', 'Add_Click', ({
            param($S, $E)
            $formObject = $S.Tag
            if ($formObject -and $formObject.PSObject.Methods['ConfirmClose']) {
                try {
                    if ($formObject.ConfirmClose() -ne $true) { return }
                } catch {
                    Write-LogError "Modal close confirmation failed" $_.Exception
                    return
                }
            }
            Show-ModalObject
        }))
    }

    $ui.SetXamlProperty($modalForm, 'txtTitle', 'Text', $FormTitle)

    $grdModalContainer = (Get-AvaloniaHost)::FindByName($modalForm, 'grdModalContainer')
    if ($grdModalContainer -and $FormObject) {
        [Avalonia.Controls.Grid]::SetRow($FormObject, 1)
        $grdModalContainer.Children.Add($FormObject) | Out-Null
    }
    $ui.ShowModalObject($modalForm)
}

function Show-ModalObject
{
    param( $Obj )

    if (-not $script:grdModal) { return }

    if ($Obj) {
        [Avalonia.Controls.Grid]::SetRow($Obj, 1)
        [Avalonia.Controls.Grid]::SetColumn($Obj, 1)
        $script:grdModal.Children.Add($Obj) | Out-Null
        $script:grdModal.IsVisible = $true
    } else {
        $script:grdModal.Children.Clear()
        $script:grdModal.IsVisible = $false
    }
}

function Close-TopModalObject
{
    if ($script:grdModal -and $script:grdModal.Children.Count -gt 0) {
        $script:grdModal.Children.RemoveAt($script:grdModal.Children.Count - 1)
        $script:grdModal.IsVisible = ($script:grdModal.Children.Count -gt 0)
    }
}

#endregion

#region Status functions

function Update-UIStatus
{
    param($Text, $Detail, [switch]$SkipLog, [switch]$Block, [switch]$Force, $CancelText)

    $ui = $script:UIProvider
    $hasText   = $PSBoundParameters.ContainsKey('Text')
    $hasDetail = $PSBoundParameters.ContainsKey('Detail')

    if($hasText -and -not $Text) { $script:BlockStatusUpdates = $false }
    elseif($script:BlockStatusUpdates -eq $true -and $Force -ne $true) {
        return $true
    }
    elseif($Block -eq $true) { $script:BlockStatusUpdates = $true }

    if($hasText -and $script:txtInfo)
    {
        $script:txtInfo.Text = $Text
        if(-not $hasDetail -and $script:txtInfoDetail) {
            $script:txtInfoDetail.Text = ""
            $script:txtInfoDetail.IsVisible = $false
        }
    }

    if($hasDetail -and $script:txtInfoDetail)
    {
        $script:txtInfoDetail.Text = $Detail
        $script:txtInfoDetail.IsVisible = [bool]$Detail
    }

    # The cancel button is per-scope: a caller that passes -CancelText gets it, and
    # any later -Text without one takes it away, so it can never outlive the wait
    # it belonged to.
    if($script:btnStatusCancel -and ($PSBoundParameters.ContainsKey('CancelText') -or $hasText))
    {
        if($CancelText) {
            $script:btnStatusCancel.Content = $CancelText
            $script:btnStatusCancel.IsEnabled = $true
            $script:btnStatusCancel.IsVisible = $true
        }
        else {
            $script:btnStatusCancel.IsVisible = $false
        }
    }

    if($script:grdStatus)
    {
        if(($hasText -and $Text) -or (-not $hasText -and $script:txtInfo -and $script:txtInfo.Text))
        {
            $script:grdStatus.IsVisible = $true
            if($SkipLog -ne $true) {
                if($hasText -and $Text) { Write-Log $Text }
                if($hasDetail -and $Detail) { Write-Log $Detail }
            }
            # Avalonia coalesces invalidations and only paints when the
            # PowerShell UI thread returns to the dispatcher. Long-running
            # synchronous work (Graph PATCH, file I/O) blocks that return,
            # and any later Write-Status "" toggles IsVisible back off
            # before a single paint happens — net effect: dimming overlay
            # never appears. RunJobs drains queued layout/render passes on
            # this thread (Avalonia's nested-pump equivalent of WPF's
            # DoEvents) so the overlay actually paints before the caller
            # resumes its blocking work.
            $ui.InvokeUIMessagePump()
        }
        elseif($hasText -and -not $Text)
        {
            $script:grdStatus.IsVisible = $false
            if($script:txtInfoDetail) {
                $script:txtInfoDetail.Text = ""
                $script:txtInfoDetail.IsVisible = $false
            }
            if($script:btnStatusCancel) { $script:btnStatusCancel.IsVisible = $false }
            $ui.InvokeUIMessagePump()
        }

        return $true
    }

    return $false
}

function Invoke-UIMessagePump
{
    # Enter a short nested platform/dispatcher loop. RunJobs() was sufficient to
    # repaint but did not reliably ingest native pointer input while synchronous
    # PowerShell owned the UI thread, leaving the visible status Cancel button dead.
    try {
        (Get-AvaloniaHost)::PumpEventsOnce(10)
    } catch {
        # Pre-init or shutdown — pump unavailable. Keep status updates non-fatal,
        # but leave evidence for any unexpected runtime/API failure.
        Write-LogDebug "Avalonia event pump unavailable: $($_.Exception.Message)"
    }
}

function Request-UIConfirmation
{
    param([string]$Message, [string]$Caption = "Confirm")

    $ui = $script:UIProvider
    return (($ui.ShowMessageBox($Message, $Caption, "YesNo", "Question")) -eq "Yes")
}

#endregion

#region About dialog (Avalonia)

function Show-AboutDialog
{
    $ui = $script:UIProvider
    $dlgAbout = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/AboutDialog.axaml'))
    if (-not $dlgAbout) { return }

    $loadedItems = @()
    $externalModules    = @('MSAL.PS','Az.Account')
    $externalAssemblies = @('Microsoft.Identity.Client.dll')

    foreach ($module in (Get-Module | Where-Object {
                $_.ModuleBase -like "$($script:AppRootFolder)*" -or $_.Name -in $externalModules
            })) {
        $ver = $module.Version
        if ($module.Version.Major -eq 0 -and $module.Version.Minor -eq 0) {
            $cmd = $module.ExportedFunctions['Get-ModuleVersion']
            if ($cmd) {
                $tmpVer = Invoke-Command -ScriptBlock $cmd.ScriptBlock
                $ver = ?? $tmpVer $ver
            }
        }
        $loadedItems += [AboutModuleEntry]@{
            Name    = $module.Name
            Version = "$ver"
            Type    = 'PSModule'
        }
    }

    $assms = [System.AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GlobalAssemblyCache -eq $false -and -not [String]::IsNullOrEmpty($_.Location) }
    foreach ($assmName in $externalAssemblies) {
        $assmObjs = $assms | Where-Object { $_.Location -like "*\$assmName" }
        foreach ($assmObj in $assmObjs) {
            try {
                $fi = [IO.FileInfo]"$($assmObj.Location)"
                $loadedItems += [AboutModuleEntry]@{
                    Name    = $fi.Name
                    Version = $fi.VersionInfo.FileVersion
                    Type    = 'Assembly'
                }
            } catch {}
        }
    }

    $ui.SetXamlProperty($dlgAbout, 'txtTitle', 'Text', 'Intune Management')
    if ($script:ActiveView) {
        $ui.SetXamlProperty($dlgAbout, 'txtViewTitle', 'Text', ("Current view: " + $script:ActiveView.Title))
        if ($script:ActiveView.Description) {
            $ui.SetXamlProperty($dlgAbout, 'txtViewDescription', 'Text', $script:ActiveView.Description)
        }
    }
    $ui.SetXamlProperty($dlgAbout, 'lstModules', 'ItemsSource', $loadedItems)

    $ui.AddXamlEvent($dlgAbout, 'linkSource', 'Add_Click', ({
        Open-ExternalUri 'https://github.com/Micke-K/IntuneManagement'
    }))
    $ui.AddXamlEvent($dlgAbout, 'linkCoffee', 'Add_Click', ({
        Open-ExternalUri 'https://buymeacoffee.com/MickeK'
    }))

    $ui.ShowModalForm('About', $dlgAbout)
}

#endregion

#region Message box (Avalonia)

function Show-MessageBox {
    # Avalonia counterpart of UI/WPF/Extensions/CoreUI.ps1 Show-MessageBox.
    # Differences vs WPF:
    #  - Params/return are strings (no [System.Windows.MessageBoxButton] enums)
    #    because no caller in this repo actually passes enum literals; they all
    #    pass strings like "OK", "YesNo", "Error".
    #  - Escape key handled in code (Avalonia has no IsCancel on Button).
    #  - Closed event derives a default result when the user dismisses via X /
    #    Alt-F4 with no button click, matching WPF's Closing-event logic.
    param(
        [Parameter(Mandatory)]
        [string]$Text,
        [string]$Caption = "",
        [ValidateSet("OK","OKCancel","YesNo","YesNoCancel")]
        [string]$Button = "OK",
        [ValidateSet("None","Information","Question","Warning","Error","Asterisk","Exclamation","Hand","Stop")]
        [string]$Icon = "None"
    )

    $ui = $script:UIProvider
    Initialize-AvaloniaRuntime

    $dialog = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/MessageBox.axaml'))
    if (-not $dialog) { return "None" }

    $captionText = if ($Caption) { $Caption } elseif ($script:Window) { $script:Window.Title } else { "" }
    $dialog.Title = $captionText
    $ui.SetXamlProperty($dialog, 'txtCaption', 'Text', $captionText)
    $ui.SetXamlProperty($dialog, 'txtMessage', 'Text', $Text)

    if ($Icon -ne 'None') {
        $iconBorder = (Get-AvaloniaHost)::FindByName($dialog, 'iconBorder')
        $iconSymbol = (Get-AvaloniaHost)::FindByName($dialog, 'iconSymbol')
        switch ($Icon) {
            'Question'                              { $iconColor = '#FF0078D4'; $iconGlyph = '?' }
            { $_ -in 'Warning','Exclamation' }      { $iconColor = '#FFE8A000'; $iconGlyph = '!' }
            { $_ -in 'Error','Hand','Stop' }        { $iconColor = '#FFD92D20'; $iconGlyph = 'X' }
            default                                 { $iconColor = '#FF0078D4'; $iconGlyph = 'i' }
        }
        $iconBorder.Background = [Avalonia.Media.SolidColorBrush]::new([Avalonia.Media.Color]::Parse($iconColor))
        $iconSymbol.Text       = $iconGlyph
        $iconBorder.IsVisible  = $true
    }

    $showOK     = $Button -in 'OK','OKCancel'
    $showYesNo  = $Button -in 'YesNo','YesNoCancel'
    $showCancel = $Button -in 'OKCancel','YesNoCancel'

    # MODULE scope, not a captured local: handler captures are stripped by
    # ConvertTo-AvaloniaEventScriptBlock, which left every button unable to
    # record its answer or close the dialog. Only one message box runs at a
    # time (ShowDialog blocks), so a single module slot is safe.
    $script:_msgBoxDialog    = $dialog
    $script:_msgBoxResult    = 'None'
    $script:_msgBoxShowCancel = $showCancel
    $script:_msgBoxShowYesNo  = $showYesNo

    if ($showOK) {
        $btnOK = (Get-AvaloniaHost)::FindByName($dialog, 'btnOK')
        $btnOK.IsVisible = $true
        $btnOK.IsDefault = $true
        $ui.AddXamlEvent($dialog, 'btnOK', 'Add_Click', ({
            $script:_msgBoxResult = 'OK'
            if ($script:_msgBoxDialog) { $script:_msgBoxDialog.Close() }
        }))
    }

    if ($showYesNo) {
        $btnYes = (Get-AvaloniaHost)::FindByName($dialog, 'btnYes')
        $btnNo  = (Get-AvaloniaHost)::FindByName($dialog, 'btnNo')
        $btnYes.IsVisible = $true
        $btnYes.IsDefault = $true
        $btnNo.IsVisible  = $true
        $ui.AddXamlEvent($dialog, 'btnYes', 'Add_Click', ({
            $script:_msgBoxResult = 'Yes'
            if ($script:_msgBoxDialog) { $script:_msgBoxDialog.Close() }
        }))
        $ui.AddXamlEvent($dialog, 'btnNo', 'Add_Click', ({
            $script:_msgBoxResult = 'No'
            if ($script:_msgBoxDialog) { $script:_msgBoxDialog.Close() }
        }))
    }

    if ($showCancel) {
        $btnCancel = (Get-AvaloniaHost)::FindByName($dialog, 'btnCancel')
        $btnCancel.IsVisible = $true
        $ui.AddXamlEvent($dialog, 'btnCancel', 'Add_Click', ({
            $script:_msgBoxResult = 'Cancel'
            if ($script:_msgBoxDialog) { $script:_msgBoxDialog.Close() }
        }))
    }

    $ui.AddXamlEvent($dialog, 'btnTitleClose', 'Add_Click', ({
        if ($script:_msgBoxDialog) { $script:_msgBoxDialog.Close() }
    }))

    $dialog.Add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
        param($S, $E)
        if ($E.Key -eq [Avalonia.Input.Key]::Escape) {
            if ($script:_msgBoxDialog) { $script:_msgBoxDialog.Close() }
        }
    }))

    $dialog.Add_Closed((ConvertTo-AvaloniaEventScriptBlock {
        if ($script:_msgBoxResult -eq 'None') {
            if ($script:_msgBoxShowCancel)    { $script:_msgBoxResult = 'Cancel' }
            elseif ($script:_msgBoxShowYesNo) { $script:_msgBoxResult = 'No' }
            else                              { $script:_msgBoxResult = 'OK' }
        }
    }))

    # WindowStartupLocation=CenterOwner positions BEFORE SizeToContent has
    # produced the final size (and misbehaves with custom decorations), which
    # is why the box landed in a corner. Center manually once the real size is
    # known; reads only the sender + $script:Window so no captures are needed.
    $dialog.add_Opened({
        param($S, $E)
        try {
            $owner = $script:Window
            if (-not $owner) { return }
            $ownerSize = if ($owner.FrameSize.HasValue) { $owner.FrameSize.Value } else { $owner.ClientSize }
            $dlgSize   = if ($S.FrameSize.HasValue) { $S.FrameSize.Value } else { $S.ClientSize }
            $x = $owner.Position.X + [int](($ownerSize.Width  * $owner.RenderScaling - $dlgSize.Width  * $S.RenderScaling) / 2)
            $y = $owner.Position.Y + [int](($ownerSize.Height * $owner.RenderScaling - $dlgSize.Height * $S.RenderScaling) / 2)
            $S.Position = [Avalonia.PixelPoint]::new($x, $y)
        } catch { Write-LogDebug "MessageBox centering failed: $($_.Exception.Message)" }
    })

    (Get-AvaloniaHost)::ShowDialog($dialog, $script:Window)
    $answer = $script:_msgBoxResult
    $script:_msgBoxDialog = $null
    return $answer
}

#endregion

#region Settings dialog (Avalonia)

# Port of UI/WPF/Extensions/CoreUI.ps1 Settings section. Differences vs WPF:
#  - Setting helpers (Add-SettingTextBox/CheckBox/ComboBox/Folder) build the
#    controls programmatically and return a PSCustomObject @{ Visual; Control }.
#    WPF parsed XAML strings and used FindName($Id) to get the inner control;
#    Avalonia doesn't auto-register names for programmatically-built trees,
#    so we hand the inner control back explicitly.
#  - ComboBox ItemsSource items are converted from PSCustomObject (which
#    Avalonia's binder can't traverse) to [SettingsListItem] CLR instances.
#    See [[avalonia-binding-needs-clr-types]].
#  - Folder picker uses (Get-AvaloniaHost)::OpenFolderPicker instead of
#    WindowsAPICodePack CommonOpenFileDialog.
#  - DynamicResource Foreground / InfoIcon style come from inline XAML snippets
#    parsed via Get-AvaloniaHost::LoadXaml so theme variant swap still works.

function Add-SettingTextBox {
    param($Value)
    $tb = [Avalonia.Controls.TextBox]::new()
    if ($null -ne $Value) { $tb.Text = [string]$Value }
    [PSCustomObject]@{ Visual = $tb; Control = $tb }
}

function Add-SettingCheckBox {
    param($Value)
    $cb = [Avalonia.Controls.CheckBox]::new()
    $cb.IsChecked = ($Value -eq $true -or $Value -eq 'true')
    [PSCustomObject]@{ Visual = $cb; Control = $cb }
}

function Add-SettingComboBox {
    param($Value, $SettingObj)

    $nameProp  = if ($SettingObj.DisplayMemberPath) { $SettingObj.DisplayMemberPath } else { 'Name' }
    $valueProp = if ($SettingObj.SelectedValuePath) { $SettingObj.SelectedValuePath } else { 'Value' }

    $items = New-Object 'System.Collections.Generic.List[SettingsListItem]'
    foreach ($src in @($SettingObj.ItemsSource)) {
        if (-not $src) { continue }
        $items.Add([SettingsListItem]@{
            Name  = [string]($src.$nameProp)
            Value = $src.$valueProp
        })
    }

    $combo = [Avalonia.Controls.ComboBox]::new()
    $combo.ItemsSource = $items
    # Built in code, so the display field is set here rather than in .axaml -
    # same mechanism, just assigned imperatively.
    $combo.DisplayMemberBinding = [Avalonia.Data.Binding]::new('Name')
    # Avalonia's ComboBox defaults to HorizontalAlignment=Left, unlike TextBox
    # and friends - without this it collapses to its content width instead of
    # filling the settings column.
    $combo.HorizontalAlignment = [Avalonia.Layout.HorizontalAlignment]::Stretch

    if ($null -ne $Value) {
        $selected = $items | Where-Object { "$($_.Value)" -eq "$Value" } | Select-Object -First 1
        if ($selected) { $combo.SelectedItem = $selected }
    }

    [PSCustomObject]@{ Visual = $combo; Control = $combo }
}

function Add-SettingFolder {
    param($Value)

    $ui = $script:UIProvider
    $grid = [Avalonia.Controls.Grid]::new()
    $grid.HorizontalAlignment = [Avalonia.Layout.HorizontalAlignment]::Stretch
    $grid.Margin = [Avalonia.Thickness]::new(0, 5, 0, 0)
    $grid.ColumnDefinitions.Add([Avalonia.Controls.ColumnDefinition]::new([Avalonia.Controls.GridLength]::new(1, [Avalonia.Controls.GridUnitType]::Star)))
    $grid.ColumnDefinitions.Add([Avalonia.Controls.ColumnDefinition]::new([Avalonia.Controls.GridLength]::new(5)))
    $grid.ColumnDefinitions.Add([Avalonia.Controls.ColumnDefinition]::new([Avalonia.Controls.GridLength]::Auto))

    $tb = [Avalonia.Controls.TextBox]::new()
    if ($null -ne $Value) { $tb.Text = [string]$Value }
    $grid.Children.Add($tb)

    $btn = [Avalonia.Controls.Button]::new()
    $btn.Content = '...'
    $btn.Width = 50
    $btn.Padding = [Avalonia.Thickness]::new(5, 0, 5, 0)
    [Avalonia.Controls.Grid]::SetColumn($btn, 2)
    $btn.Tag = $tb
    $btn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
        param($S, $E)
        $textBox = $S.Tag
        $picked = $ui.ShowFolderPicker('Select folder')
        if ($picked) { $textBox.Text = $picked }
    }))
    $grid.Children.Add($btn)

    [PSCustomObject]@{ Visual = $grid; Control = $tb }
}

function Add-SettingsItem {
    param($SettingItem, $SettingValue)

    $rd = [Avalonia.Controls.RowDefinition]::new([Avalonia.Controls.GridLength]::Auto)
    $script:spSettings.RowDefinitions.Add($rd)
    $rowIndex = $script:spSettings.RowDefinitions.Count - 1
    [Avalonia.Controls.Grid]::SetRow($SettingItem, $rowIndex)

    if (-not $SettingValue) {
        [Avalonia.Controls.Grid]::SetColumnSpan($SettingItem, 99)
        $script:spSettings.Children.Add($SettingItem)
        return
    }

    # Title strip: developer-supplied $Title / $Description are assigned via
    # property setters after the parse so a stray <, &, " in the JSON config
    # can't break the XAML parser. Same approach the WPF version takes.
    $titleXaml = @'
<StackPanel xmlns="https://github.com/avaloniaui" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Orientation="Horizontal" Margin="5,5,5,0">
    <TextBlock x:Name="settingTitleText" Foreground="{DynamicResource TitleForegroundColor}" VerticalAlignment="Center"/>
    <TextBlock x:Name="settingDescIcon" Text="&#x24D8;" Margin="5,0,0,0" IsVisible="False"
               Foreground="{DynamicResource TitleForegroundColor}" VerticalAlignment="Center" FontSize="12"/>
</StackPanel>
'@
    $titleStack = (Get-AvaloniaHost)::LoadXaml($titleXaml)
    $titleText = (Get-AvaloniaHost)::FindByName($titleStack, 'settingTitleText')
    $descIcon  = (Get-AvaloniaHost)::FindByName($titleStack, 'settingDescIcon')
    if ($titleText) { $titleText.Text = [string]$SettingValue.Title }
    if ($SettingValue.Description -and $descIcon) {
        [Avalonia.Controls.ToolTip]::SetTip($descIcon, [string]$SettingValue.Description)
        $descIcon.IsVisible = $true
    }

    if ($script:tenantSettings) {
        $tenantConfig = [Avalonia.Controls.CheckBox]::new()
        [Avalonia.Controls.ToolTip]::SetTip($tenantConfig, 'Enable tenant specific setting')
        [Avalonia.Controls.Grid]::SetRow($tenantConfig, $rowIndex)
        [Avalonia.Controls.Grid]::SetColumn($tenantConfig, 0)
        $tenantConfig.Margin = [Avalonia.Thickness]::new(0, 5, 0, 0)
        $tenantConfig.Tag = $SettingValue
        $tenantConfig.IsChecked = (Test-SettingValueConfigured -Key $SettingValue.Key -Tenant)
        $SettingItem.IsEnabled = [bool]$tenantConfig.IsChecked
        $tenantConfig.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            if ($S.Tag -and $S.Tag.Control) { $S.Tag.Control.IsEnabled = [bool]$S.IsChecked }
        }))
        $script:spSettings.Children.Add($tenantConfig)
    }

    [Avalonia.Controls.Grid]::SetRow($titleStack, $rowIndex)
    [Avalonia.Controls.Grid]::SetColumn($titleStack, 1)
    $script:spSettings.Children.Add($titleStack)

    [Avalonia.Controls.Grid]::SetColumn($SettingItem, 2)
    $SettingItem.Margin = [Avalonia.Thickness]::new(0, 5, 0, 0)
    $script:spSettings.Children.Add($SettingItem)
}

function Add-SettingTitle {
    param($Title, $MarginTop = 0)

    # Parse from inline XAML so the {DynamicResource ...} brush refs survive
    # theme-variant switches. The previous version tried
    # Application.Current.TryFindResource — that's an extension method on
    # IResourceHost in Avalonia 11, not an instance method, so it throws
    # "does not contain a method named TryFindResource" when invoked from
    # PowerShell. Title text + margin are set after parse to keep developer-
    # supplied $Title out of the XAML literal (safe against stray <, &, ").
    $xaml = @'
<TextBlock xmlns="https://github.com/avaloniaui"
           FontWeight="Bold" Padding="5"
           Background="{DynamicResource SettingsSectionBackgroundColor}"
           Foreground="{DynamicResource TitleForegroundColor}"/>
'@
    $tb = (Get-AvaloniaHost)::LoadXaml($xaml)
    $tb.Text   = [string]$Title
    $tb.Margin = [Avalonia.Thickness]::new(0, [double]$MarginTop, 0, 0)

    Add-SettingsItem $tb
}

function Add-SettingValue {
    param($SettingValue)

    if ($SettingValue.TenantSettings -eq $false -and $script:tenantSettings) { return }
    if ($SettingValue.GlobalSettings -eq $false -and -not $script:tenantSettings) { return }

    $Value = Get-SettingValue $SettingValue.Key -GlobalOnly:(-not $script:tenantSettings)

    $built = switch -Regex ("$($SettingValue.Type)") {
        '^(?i)folder$'  { Add-SettingFolder $Value;             break }
        '^(?i)boolean$' { Add-SettingCheckBox $Value;           break }
        '^(?i)list$'    { Add-SettingComboBox $Value $SettingValue; break }
        default         { Add-SettingTextBox $Value }
    }

    if ($built) {
        Add-SettingsItem $built.Visual $SettingValue
        if ($SettingValue | Get-Member -MemberType NoteProperty -Name 'Control') {
            $SettingValue.Control = $built.Control
        } else {
            $SettingValue | Add-Member -MemberType NoteProperty -Name 'Control' -Value $built.Control -Force
        }
    }
}

function Show-SettingsForm {
    param([switch]$Tenant)

    $ui = $script:UIProvider
    Initialize-AvaloniaRuntime

    $settingsForm = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/SettingsForm.axaml'))
    if (-not $settingsForm) { return }

    $script:spSettings = (Get-AvaloniaHost)::FindByName($settingsForm, 'spSettings')
    $script:tenantSettings = [bool]$Tenant

    $ui.AddXamlEvent($settingsForm, 'btnSave', 'Add_Click', ({
        Save-AllSettings
    }))

    $ui.AddXamlEvent($settingsForm, 'btnClose', 'Add_Click', ({
        $script:tenantSettings = $false
        $script:spSettings = $null
        Show-ModalObject
    }))

    if ($JsonSettingsObj -or $script:tenantSettings) {
        $ui.SetXamlProperty($settingsForm, 'btnExport', 'IsVisible', $false)
    } else {
        $ui.AddXamlEvent($settingsForm, 'btnExport', 'Add_Click', ({
            $picked = (Get-AvaloniaHost)::SaveFilePicker(
                $script:Window,
                'Export settings',
                'IntuneManagementSettings',
                'json',
                'Json files',
                '*.json')
            if ($picked) { Export-Settings $picked }
        }))
    }

    $tmp = Get-SettingsSection 'General'
    if ($tmp -and $tmp.Values.Count -gt 0) {
        Add-SettingTitle $tmp.Title
        foreach ($SettingObj in $tmp.Values) { Add-SettingValue $SettingObj }
    }

    foreach ($section in ((Get-SettingsSections) | Where-Object Id -ne 'General' | Sort-Object -Property Order, Title)) {
        if ($section.Values.Count -eq 0) { continue }
        Add-SettingTitle $section.Title 5
        foreach ($SettingObj in $section.Values) { Add-SettingValue $SettingObj }
    }

    $ui.ShowModalForm('Settings', $settingsForm)
}

function Save-AllSettings {
    Write-Status 'Save settings'
    $dt1 = Get-Date
    $curHideNoAccess = Get-SettingValue 'HideNoAccess'

    foreach ($section in (Get-SettingsSections)) {
        foreach ($SettingObj in $section.Values) {
            if (-not $SettingObj.Control) { continue }
            if ($SettingObj.Control.IsEnabled -eq $false -and $script:tenantSettings) {
                Remove-SettingValue -Key $SettingObj.Key -Tenant
                continue
            }

            $valueFound = $false
            $Value = $null
            $typeName = $SettingObj.Control.GetType().Name
            if ($typeName -eq 'TextBox') {
                $Value = $SettingObj.Control.Text
                if ($SettingObj.Type -eq 'Int') {
                    try { $Value = [int]$Value } catch { $Value = $SettingObj.Value }
                }
                $valueFound = $true
            } elseif ($typeName -eq 'CheckBox') {
                $Value = [bool]$SettingObj.Control.IsChecked
                $valueFound = $true
            } elseif ($typeName -eq 'ComboBox') {
                $selected = $SettingObj.Control.SelectedItem
                if ($selected) { $Value = $selected.Value }
                $valueFound = $true
            }

            if (-not $valueFound) { continue }

            # The path comes from the resolver, not from string concatenation here.
            # The old code composed "$tenantId\$($SettingObj.SubPath)" and only
            # assigned $subPath INSIDE `if ($tenantId)`, with no else - so with the
            # tenant settings form open and no resolvable tenant, this setting was
            # written to the PREVIOUS setting's path.
            $isTenant = ($script:tenantSettings -eq $true)
            $subPath = Resolve-SettingStorePath -Key $SettingObj.Key -Definition $SettingObj -Tenant:$isTenant
            if ($null -eq $subPath) { continue }

            # The raw stored value at THIS scope, not the effective one: it only feeds
            # the change event below, and an unset tenant value has to look unset.
            $currentValue = Get-SettingStoreValue -SubPath $subPath -Key $SettingObj.Key
            Set-SettingValue -Key $SettingObj.Key -Value $Value -Tenant:$isTenant

            $stringValue = if ($null -ne $Value -and $Value -isnot [string]) { $Value.ToString() } else { $Value }
            if ($stringValue -cne $currentValue) {
                # Fire on EVERY change, including from/to empty - the old
                # non-empty guards suppressed the live-apply (theme, env badge)
                # for a setting's very first save.
                Write-Log "Setting saved: $($SettingObj.Key) = '$stringValue' (was: '$currentValue')"
                # Invoke-AppEvent takes ONE argument payload and splats it into the
                # handler, so a multi-argument event has to pass a hashtable. Three
                # positional args bound only $SettingObj (as $EventArguments) and
                # dropped the rest into $args, so handlers saw $NewValue = $null -
                # which is why saving a new theme persisted it but never applied it.
                Invoke-AppEvent 'SettingValueUpdated' @{
                    SettingInfo = $SettingObj
                    NewValue    = $Value
                    OldValue    = $currentValue
                }
            }
        }
    }

    Invoke-AppEvent 'SettingsUpdated'

    $newHideNoAccess = Get-SettingValue 'HideNoAccess'
    if ($curHideNoAccess -ne $newHideNoAccess) {
        Show-ViewMenu
    }

    if ($dt1.AddSeconds(1) -lt (Get-Date)) {
        Start-Sleep -Seconds 1
    }
    Write-Status ''
}

#endregion

#region Popup overlay (Avalonia)

# Convert a point in an anchor visual's local coords to window-client coords.
# Uses PointToScreen -> PointToClient (rather than TranslatePoint) because the
# latter returned null in practice when both visuals were attached but Avalonia
# hadn't run a full layout pass on the popup-host yet. Returns $null on
# failure so callers can fall back to a sensible default.
function Get-WindowRelativePoint {
    param($AnchorVisual, [Avalonia.Point]$AnchorPoint)

    if (-not $AnchorVisual -or -not $script:Window) { return $null }

    # Visual-tree transform first. It stays entirely inside the window's own
    # coordinate space, so window decorations cannot affect it. The
    # PointToScreen/PointToClient round-trip below leaves that space and needs
    # the platform to report a consistent client origin - which it does not on
    # X11 once the WM owns the title bar, so it failed there and callers
    # silently fell back to a fixed guess.
    # TranslatePoint is an EXTENSION method (Avalonia.VisualExtensions), not an
    # instance method, so PowerShell cannot reach it with dot syntax -
    # $visual.TranslatePoint(...) throws "does not contain a method named".
    # It has to be invoked statically with the visual as the first argument.
    try {
        $translated = [Avalonia.VisualExtensions]::TranslatePoint($AnchorVisual, $AnchorPoint, $script:Window)
        if ($null -ne $translated) {
            # PowerShell usually unwraps Nullable[Point]; handle both shapes.
            if ($translated -is [Avalonia.Point]) { return $translated }
            if ($translated.HasValue) { return $translated.Value }
        }
    } catch {
        Write-LogDebug "Get-WindowRelativePoint: TranslatePoint failed: $($_.Exception.Message)"
    }

    try {
        $screenPt = $AnchorVisual.PointToScreen($AnchorPoint)
        return $script:Window.PointToClient($screenPt)
    } catch {
        Write-LogDebug "Get-WindowRelativePoint failed: $($_.Exception.Message)"
        return $null
    }
}

# Avalonia port of UI/WPF/Extensions/CoreUI.ps1 Show-Popup / Hide-Popup. The grdPopup
# overlay + cvsPopup Canvas are already in MainWindow.axaml; helpers mount /
# unmount the popup content and toggle visibility. Click-outside (on the dim
# layer) and Escape both dismiss.

function Show-Popup {
    param($Popup)

    $ui = $script:UIProvider
    if (-not $script:grdPopup -or -not $script:cvsPopup) { return }

    $script:cvsPopup.Children.Add($Popup) | Out-Null
    $script:grdPopup.IsVisible = $true

    if (-not $script:popupHandlersInstalled) {
        # Escape dismisses. KeyDown on the window catches it regardless of focus.
        $script:Window.add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            if ($E.Key -eq [Avalonia.Input.Key]::Escape -and
                $script:grdPopup -and $script:grdPopup.IsVisible) {
                $ui.HidePopup()
                $E.Handled = $true
            }
        }))
        # Click on the dim layer (i.e. anywhere outside the popup content)
        # dismisses. Children of cvsPopup mark their own PointerPressed handled,
        # so a click on the popup body doesn't bubble back here.
        $script:grdPopup.add_PointerPressed((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            if ($E.Source -eq $S -or $E.Source -eq $script:cvsPopup) {
                $ui.HidePopup()
            }
        }))
        $script:popupHandlersInstalled = $true
    }
}

function Hide-Popup {
    if (-not $script:grdPopup -or -not $script:cvsPopup) { return }
    $script:cvsPopup.Children.Clear()
    $script:grdPopup.IsVisible = $false
}

# Position a popup below an anchor visual, right-aligned with the anchor's
# right edge. Fallback: if PointToScreen / PointToClient fail (anchor not yet
# rendered, or other reason), position near top-right of the window so the
# popup is at least visible and approximately where the user expects.
function Set-PopupPosition {
    param($Popup, $Anchor, [double]$PopupWidth = 320)

    if (-not $Popup) { return }

    if ($Popup.Width -gt 0) { $PopupWidth = $Popup.Width }

    $placed = $false
    if ($Anchor -and $script:Window) {
        $clientPt = Get-WindowRelativePoint $Anchor ([Avalonia.Point]::new(0, $Anchor.Bounds.Height))
        if ($clientPt) {
            [Avalonia.Controls.Canvas]::SetLeft($Popup, $clientPt.X + $Anchor.Bounds.Width - $PopupWidth)
            [Avalonia.Controls.Canvas]::SetTop($Popup, $clientPt.Y + 4)
            $placed = $true
        }
    }

    if (-not $placed -and $script:Window) {
        # Fallback: anchor to the right edge of the window client area, just
        # below the title strip (32px tall).
        #
        # Only Windows paints caption buttons over our client area
        # (ExtendClientAreaChromeHints=PreferSystemChrome), and only there does
        # grdTitleBarContent keep its reserve - Set-MainWindowChrome zeroes it
        # everywhere else. Subtracting it unconditionally put the popup a full
        # caption-width left of the avatar it is meant to hang below, which is
        # exactly what showed on Linux.
        $reserve = 0
        if ($script:IsWindowsOS) {
            $reserve = if ($script:IMWindowsCaptionWidth) { $script:IMWindowsCaptionWidth } else { 140 }
        }
        $clientWidth = $script:Window.Bounds.Width
        if ($clientWidth -le 0) { $clientWidth = $script:Window.Width }
        [Avalonia.Controls.Canvas]::SetLeft($Popup, [Math]::Max(0, $clientWidth - $reserve - $PopupWidth))
        [Avalonia.Controls.Canvas]::SetTop($Popup, 36)
    }
}

#endregion

#region Authentication chrome (Avalonia)

# WPF Show-AuthenticationInfo defers to Get-MSALUserProfile (~400 lines: photo
# ellipse, sign-in button, cached-account popup, "Sign in to a different cloud"
# button). The Avalonia port currently implements only the two terminal states:
#   - Signed out: a "Sign in" Button that invokes Invoke-AuthProviderInteractiveLogin.
#   - Signed in:  a circular Border with the user's photo or initials and a
#                 tooltip showing display name + UPN. Click opens the profile
#                 popup. Cached-account picker shown when cached rows exist.
# When more of Get-MSALUserProfile is ported (session info, token viewers,
# cloud picker, etc.) consider refactoring the WPF version to share.

# Fills a Border with the user's avatar: ImageBrush background from
# $script:CurrentProfilePhoto if a non-empty file exists, otherwise solid blue
# with the user's initials. Used both for the small title-bar avatar (28x28,
# FontSize 11) and the larger popup avatar (60x60, FontSize 24).
function Set-AvatarVisual {
    param(
        [Avalonia.Controls.Border]$Border,
        [int]$FontSize = 11
    )

    # App-login detection is provider-agnostic (GetUserInfo().AuthType), so the 'APP'
    # initial works for any provider, not just MSAL.
    $isAppLogin = $false
    try {
        $p = Get-AuthProvider
        if ($p) { $u = $p.GetUserInfo((Get-DefaultAuthTokenId)); $isAppLogin = [bool]($u -and $u.AuthType -in @('Confidential','ClientCredential','ManagedIdentity','WorkloadFederation')) }
    } catch { }

    $initials = if ($script:CurrentUser.givenName -and $script:CurrentUser.surname) {
        "$($script:CurrentUser.givenName[0])$($script:CurrentUser.surname[0])".ToUpper()
    } elseif ($script:CurrentUser.userPrincipalName) {
        "$($script:CurrentUser.userPrincipalName[0])".ToUpper()
    } elseif ($isAppLogin) {
        'APP'
    } else {
        '?'
    }

    $photoLoaded = $false
    if ($script:CurrentProfilePhoto -and (Test-Path -LiteralPath $script:CurrentProfilePhoto)) {
        try {
            $fi = Get-Item -LiteralPath $script:CurrentProfilePhoto
            if ($fi.Length -gt 0) {
                $bitmap = [Avalonia.Media.Imaging.Bitmap]::new($script:CurrentProfilePhoto)
                $brush  = [Avalonia.Media.ImageBrush]::new($bitmap)
                $brush.Stretch = [Avalonia.Media.Stretch]::UniformToFill
                $Border.Background = $brush
                $Border.Child = $null
                $photoLoaded = $true
            }
        } catch {
            Write-LogDebug "Profile photo load failed, falling back to initials: $($_.Exception.Message)"
        }
    }

    if (-not $photoLoaded) {
        $Border.Background = [Avalonia.Media.SolidColorBrush]::new([Avalonia.Media.Color]::Parse('#FF0078D4'))
        $tb = [Avalonia.Controls.TextBlock]::new()
        $tb.Text = $initials
        $tb.Foreground = [Avalonia.Media.Brushes]::White
        $tb.FontWeight = [Avalonia.Media.FontWeight]::Bold
        $tb.HorizontalAlignment = [Avalonia.Layout.HorizontalAlignment]::Center
        $tb.VerticalAlignment   = [Avalonia.Layout.VerticalAlignment]::Center
        $tb.FontSize = $FontSize
        $Border.Child = $tb
    }
}

function Get-AvaloniaUserProfile {
    # An expired default token counts as "not signed in" so the title-bar reverts to
    # the Sign-in button (which prompts interactive login on click). Provider-agnostic;
    # returns $false when expiry is unknown/SDK-managed.
    $tokenExpired = $false
    if(Get-Command Test-DefaultTokenExpired -ErrorAction SilentlyContinue) {
        try { $tokenExpired = Test-DefaultTokenExpired } catch { }
    }
    # Signed-in state is provider-agnostic: $script:CurrentUser is set by
    # Sync-AuthContextFromProvider for every provider, while $script:MSALDefaultToken
    # is MSAL-only (null for an OAuth/MgGraph session). Gating on it kept OAuth logins
    # showing the Sign-in button.
    if (-not $script:CurrentUser -or $tokenExpired) {
        $btn = [Avalonia.Controls.Button]::new()
        $btn.Content = 'Sign in'
        $btn.Padding = [Avalonia.Thickness]::new(8, 2, 8, 2)
        $btn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            try {
                # If the active provider exposes cached accounts, show the
                # picker popup so the user can re-auth as a previously-signed-in
                # account without re-typing their UPN. Empty list (or provider
                # doesn't support caching, e.g. MgGraph today) falls through to
                # the interactive flow.
                $authProvider = $null
                try { $authProvider = Get-AuthProvider } catch { Write-LogError 'Get-AuthProvider failed in login-click handler' $_.Exception }

                # Bootstrap an MSAL public-client app + refresh $script:MSALAccounts
                # from the on-disk cache before reading cached accounts. On a fresh
                # session (or pure BYO/CC session) $script:MSALApps is empty, so
                # AuthenticationMSAL.GetCachedAccounts() would return [] even when
                # the persistent cache has interactive accounts from earlier
                # sessions. Mirrors WPF Get-MSALUserProfile (~line 117). No network
                # call — GetAccountsAsync just deserialises the cache file.
                try {
                    $publicClientApp = $script:MSALApps | Select-Object -First 1
                    if (-not $publicClientApp -and (Get-Command New-MSALApp -ErrorAction SilentlyContinue)) {
                        $publicClientApp = New-MSALApp
                    }
                    if ($publicClientApp) {
                        $script:MSALAccounts = $publicClientApp.GetAccountsAsync().GetAwaiter().GetResult()
                    }
                } catch {
                    Write-LogDebug "Failed to refresh cached account list: $($_.Exception.Message)"
                }

                $cachedRows = @()
                if ($authProvider -and $authProvider.SupportsCachedUsers) {
                    try { $cachedRows = @($authProvider.GetCachedAccounts()) } catch { $cachedRows = @() }
                }

                if ($cachedRows.Count -gt 0) {
                    Show-CachedAccountPicker $S $cachedRows
                    return
                }

                # Do not start a synchronous wait before this Click event returns:
                # Avalonia still owns pointer capture for the Sign-in button until
                # then, which makes the first later click on Cancel disappear.
                Invoke-AvaloniaDeferredAction -Action {
                    param($unused)
                    Invoke-AvaloniaInteractiveSignIn
                }
            } catch {
                Write-LogError 'Sign-in click handler failed' $_.Exception
                Write-Status ''
            }
        }))
        return $btn
    }

    $border = [Avalonia.Controls.Border]::new()
    $border.Width = 28
    $border.Height = 28
    $border.CornerRadius = [Avalonia.CornerRadius]::new(14)
    $border.Cursor = [Avalonia.Input.Cursor]::new([Avalonia.Input.StandardCursorType]::Hand)

    $tipText = if ($script:CurrentUser.userPrincipalName) {
        "$($script:CurrentUser.displayName)`n$($script:CurrentUser.userPrincipalName)"
    } else {
        [string]$script:CurrentUser.displayName
    }
    [Avalonia.Controls.ToolTip]::SetTip($border, $tipText)

    Set-AvatarVisual -Border $border -FontSize 11

    $border.add_PointerPressed((ConvertTo-AvaloniaEventScriptBlock {
        param($S, $E)
        # Mark handled so the title bar's PointerPressed does not also start a
        # window drag - the drag filter only skips Button/MenuItem/Menu, and
        # this avatar is a Border, so clicking it dragged the window too.
        $E.Handled = $true
        try { Show-ProfilePopup $S } catch { Write-LogError 'Show-ProfilePopup failed' $_.Exception }
    }))

    return $border
}

function Show-ProfilePopup {
    param($Anchor)

    $ui = $script:UIProvider
    if (-not $script:CurrentUser) { return }

    $popup = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/ProfileInfo.axaml'))
    if (-not $popup) { return }

    # Provider-agnostic user view drives org / auth / expiry / app / app-login below.
    # Identity name + photo still come from $script:CurrentUser / $script:CurrentProfilePhoto.
    $provider = $null
    try { $provider = Get-AuthProvider } catch { }
    $userInfo = $null
    if ($provider) { try { $userInfo = $provider.GetUserInfo((Get-DefaultAuthTokenId)) } catch { } }
    $isAppLogin = [bool]($userInfo -and $userInfo.AuthType -in @('Confidential','ClientCredential','ManagedIdentity','WorkloadFederation'))

    $orgName = if ($userInfo -and $userInfo.TenantName) { $userInfo.TenantName } else { $script:OrganizationName }
    $ui.SetXamlProperty($popup, 'txtOrganization', 'Text', ([string]$orgName))

    # An app-only login has no human account, so the provider fills DisplayName /
    # UPN from the token's account username - an identifier (the client id, or the
    # service principal oid on a BYO token), never a name. Lead with the resolved
    # app name instead; the identifier is still shown by pnlAppInfo below. WPF
    # prints the fixed string "App Login" here, which is the fallback when the
    # app name could not be resolved.
    $displayName = if ($isAppLogin) {
                       if ($userInfo.AppName) { [string]$userInfo.AppName } else { 'App Login' }
                   }
                   elseif ($script:CurrentUser.displayName) { [string]$script:CurrentUser.displayName }
                   elseif ($script:CurrentUser.userPrincipalName) { [string]$script:CurrentUser.userPrincipalName }
                   else { '' }
    $ui.SetXamlProperty($popup, 'txtUsername', 'Text', $displayName)
    $ui.SetXamlProperty($popup, 'txtLogonName', 'Text', ([string]$script:CurrentUser.userPrincipalName))

    # Auth method + token expiry from the provider-agnostic user view.
    $authLabel   = $null
    $expiryLabel = $null
    if ($userInfo) {
        $authTypeRaw = if ($userInfo.AuthType) { $userInfo.AuthType } else { 'Interactive' }
        $modeName = switch ($authTypeRaw) {
            'BYO'              { 'Bring-your-own token' }
            'Confidential'     { 'Client credentials' }
            'ClientCredential' { 'Client credentials' }
            'ManagedIdentity'  { 'Managed identity' }
            default            { 'Interactive' }
        }
        $providerDisplay = if ($provider) { $provider.DisplayName } else { '' }
        $authLabel = "$providerDisplay - $modeName"
        if ($userInfo.ExpiresOn) { $expiryLabel = "Token expires $($userInfo.ExpiresOn.ToString('g'))" }
    }
    if ($authLabel) {
        $ui.SetXamlProperty($popup, 'txtAuthMethod', 'Text', $authLabel)
        $ui.SetXamlProperty($popup, 'txtAuthMethod', 'IsVisible', $true)
    }
    if ($expiryLabel) {
        $ui.SetXamlProperty($popup, 'txtAuthExpiry', 'Text', $expiryLabel)
        $ui.SetXamlProperty($popup, 'txtAuthExpiry', 'IsVisible', $true)
    }

    $bigAvatar = (Get-AvaloniaHost)::FindByName($popup, 'bigAvatar')
    if ($bigAvatar) { Set-AvatarVisual -Border $bigAvatar -FontSize 24 }

    # App info from the provider-agnostic user view (GetUserInfo resolves it per provider).
    $appName = if ($userInfo) { $userInfo.AppName } else { $null }
    $appId   = if ($userInfo) { $userInfo.AppId }   else { $null }
    if ($appName -or $appId) {
        $ui.SetXamlProperty($popup, 'txtAppName', 'Text', ([string]$appName))
        $ui.SetXamlProperty($popup, 'txtAppId', 'Text', ([string]$appId))
        $ui.SetXamlProperty($popup, 'pnlAppInfo', 'IsVisible', $true)
    }

    # Request Consent - only providers that support it, skip on app-token logins
    # (the consent flow targets a signed-in user, N/A for client-cred), and -
    # matching WPF - only when permissions are actually missing. Avalonia showed
    # it on every MSAL session.
    $hasMissingPermissions = (($script:missingPermissions | Measure-Object).Count -gt 0)
    if ($provider -and $provider.SupportsConsentPrompt -and -not $isAppLogin -and $hasMissingPermissions) {
        $ui.SetXamlProperty($popup, 'lnkRequestConsent', 'IsVisible', $true)
        $ui.AddXamlEvent($popup, 'lnkRequestConsent', 'Add_Click', ({
            $ui.HidePopup()
            Invoke-AvaloniaDeferredAction -Action {
                param($unused)
                Invoke-AvaloniaConsentPrompt
            }
        }))
    }

    $ui.AddXamlEvent($popup, 'btnProfileClose', 'Add_Click', ({ $ui.HidePopup() }))

    # Refresh: same logic as WPF lnkForceRefresh — MSAL fast-path via
    # Connect-EntraEnvironment -ForceRefresh, generic providers via Refresh(0).
    # Gating: providers that can't refresh (e.g. MgGraph today) and BYO tokens
    # disable the button entirely.
    $lnkRefresh = (Get-AvaloniaHost)::FindByName($popup, 'lnkRefresh')
    if ($lnkRefresh) {
        $activeProvider = $null
        try { $activeProvider = Get-AuthProvider } catch { }

        if ($activeProvider -and -not $activeProvider.SupportsRefresh) {
            $lnkRefresh.IsEnabled = $false
            [Avalonia.Controls.ToolTip]::SetTip($lnkRefresh, "$($activeProvider.DisplayName) does not support manual refresh.")
        } elseif ($userInfo -and $userInfo.AuthType -eq 'BYO') {
            $lnkRefresh.IsEnabled = $false
            [Avalonia.Controls.ToolTip]::SetTip($lnkRefresh, 'BYO tokens cannot be refreshed. Re-run Connect-IntuneManagement with a new token.')
        }

        $ui.AddXamlEvent($popup, 'lnkRefresh', 'Add_Click', ({
            Invoke-AvaloniaDeferredAction -Action {
                param($unused)
                Invoke-AvaloniaProfileRefresh
            }
        }))
    }

    # Session Info — provider-agnostic. MSAL: AuthenticationResult fields.
    # MgGraph: Get-MgContext properties. Other providers: empty.
    $ui.AddXamlEvent($popup, 'lnkTokeninfo', 'Add_Click', ({
        $tokenArr = @()
        $providerNow = $null
        try { $providerNow = Get-AuthProvider } catch { }

        # Provider-agnostic: each provider returns its own native session rows (MSAL
        # AuthenticationResult fields, MgGraph Get-MgContext, ...) as Name/Value pairs;
        # wrap into the CLR row type the Avalonia DataGrid binds to.
        if ($providerNow) {
            foreach ($r in $providerNow.GetSessionInfoRows()) {
                $tokenArr += [TokenInfoRow]@{ Name = $r.Name; Value = $r.Value }
            }
        }

        $dg = [Avalonia.Controls.DataGrid]::new()
        $dg.AutoGenerateColumns = $true
        $dg.IsReadOnly          = $true
        $dg.CanUserSortColumns  = $true
        $dg.MinWidth            = 600
        $dg.MinHeight           = 400
        $dg.ItemsSource         = $tokenArr
        $ui.ShowModalForm('Session Info', $dg)
    }))

    $ui.AddXamlEvent($popup, 'lnkAccessTokenInfo', 'Add_Click', ({
        # Decode the active provider's access token on demand (works for any provider
        # whose access token is a JWT - MSAL, OAuth).
        $providerNow = $null
        try { $providerNow = Get-AuthProvider } catch { }
        if ($providerNow) {
            $raw = $providerNow.GetAccessToken(0, "https://$(Get-GraphDomain)")
            if ($raw) {
                $jwt = Get-JWTtoken $raw
                Show-MSALDecodedToken $jwt 'Access Token Info'
            } else {
                $ui.ShowMessageBox("No access token available from provider '$($providerNow.Id)'.", 'Access Token Info', 'OK', 'Information') | Out-Null
            }
        }
    }))

    $ui.AddXamlEvent($popup, 'lnkIdTokenInfo', 'Add_Click', ({
        # ID token decode via the provider (only providers that surface one return it).
        $providerNow = $null
        try { $providerNow = Get-AuthProvider } catch { }
        $idJwt = if ($providerNow) { $providerNow.GetIdTokenJwt((Get-DefaultAuthTokenId)) } else { $null }
        if ($idJwt) {
            Show-MSALDecodedToken $idJwt 'Id Token Info'
        } else {
            $ui.ShowMessageBox('ID token is not available from the active provider.', 'Id Token Info', 'OK', 'Information') | Out-Null
        }
    }))

    # Permissions: token scopes x Intune role, per policy type (EffectivePermissionsUIAvalonia.ps1).
    $ui.AddXamlEvent($popup, 'lnkEffectivePermissions', 'Add_Click', ({ Show-EffectivePermissionsDialog }))

    # Inspector button visibility, self-describing per provider (no MSAL knowledge):
    #   Session Info - only when the provider returns session rows (OAuth has none).
    #   Id Token     - only when the provider surfaces an id token (MSAL).
    #   Access Token - always (decoded on demand from GetAccessToken).
    $lnkSession = (Get-AvaloniaHost)::FindByName($popup, 'lnkTokeninfo')
    if ($lnkSession -and (-not $provider -or @($provider.GetSessionInfoRows()).Count -eq 0)) { $lnkSession.IsVisible = $false }
    $lnkIdToken = (Get-AvaloniaHost)::FindByName($popup, 'lnkIdTokenInfo')
    if ($lnkIdToken -and (-not $provider -or $null -eq $provider.GetIdTokenJwt((Get-DefaultAuthTokenId)))) { $lnkIdToken.IsVisible = $false }

    $ui.AddXamlEvent($popup, 'lnkLogout', 'Add_Click', ({
        $ui.HidePopup()
        try {
            if (Disconnect-EntraEnvironment) {
                # AuthenticationUserDisconnected event triggers Get-MSALUserInfo
                # → Show-AuthenticationInfo, which swaps the avatar back to a
                # Sign-in button. Nothing to do here.
            }
        } catch {
            Write-LogError 'Sign-out failed' $_.Exception
        }
    }))

    # Cached accounts — other MSAL identities the user has signed in with.
    # Skip the current user and personal MSA accounts. Mirrors the WPF block
    # in MSGraphAuthenticationUIWPF.ps1 (Add-CachedUser per row).
    $grdCached = (Get-AvaloniaHost)::FindByName($popup, 'grdCachedAccounts')
    if ($grdCached) {
        $activeProviderForCached = $null
        try { $activeProviderForCached = Get-AuthProvider } catch { }

        $cachedRows = @()
        if ($activeProviderForCached -and $activeProviderForCached.SupportsCachedUsers) {
            try { $cachedRows = @($activeProviderForCached.GetCachedAccounts()) } catch { $cachedRows = @() }
        }
        if ((Get-SettingValue 'SortAccountList') -eq $true) {
            $cachedRows = @($cachedRows | Sort-Object -Property Username)
        }

        foreach ($row in $cachedRows) {
            if (-not $row.Native) { continue }
            $acct = $row.Native

            # Skip the currently signed-in user (matched on UPN / user id from the
            # provider-agnostic user view).
            if ($script:CurrentUser.userPrincipalName -eq $acct.Username -or
                ($userInfo -and $userInfo.UserId -eq $acct.HomeAccountId.ObjectId)) {
                continue
            }
            if (Test-IsPersonalMSAAccount $acct) { continue }

            Add-CachedUser $acct $grdCached
        }
    }

    # "Sign in with a different account" + optional cloud picker — appended to
    # grdLoginAccount footer. Mirrors WPF rows after grdCachedAccounts.
    $grdLoginAccount = (Get-AvaloniaHost)::FindByName($popup, 'grdLoginAccount')
    if ($grdLoginAccount) {
        $altBtn = [Avalonia.Controls.Button]::new()
        $altBtn.Content = 'Sign in with a different account'
        $altBtn.Margin  = [Avalonia.Thickness]::new(0, 5, 0, 0)
        $altBtn.HorizontalAlignment = [Avalonia.Layout.HorizontalAlignment]::Stretch
        $altBtn.HorizontalContentAlignment = [Avalonia.Layout.HorizontalAlignment]::Left
        $altBtn.Cursor = [Avalonia.Input.Cursor]::new([Avalonia.Input.StandardCursorType]::Hand)
        $altBtn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $ui.HidePopup()
            Invoke-AvaloniaDeferredAction -Action {
                param($unused)
                Invoke-AvaloniaInteractiveSignIn
            }
        }))
        $grdLoginAccount.Children.Add($altBtn) | Out-Null

        if (@($script:Clouds).Count -gt 1) {
            $cloudBtn = [Avalonia.Controls.Button]::new()
            $cloudBtn.Content = 'Sign in to a different cloud...'
            $cloudBtn.Margin  = [Avalonia.Thickness]::new(0, 5, 0, 0)
            $cloudBtn.HorizontalAlignment = [Avalonia.Layout.HorizontalAlignment]::Stretch
            $cloudBtn.HorizontalContentAlignment = [Avalonia.Layout.HorizontalAlignment]::Left
            $cloudBtn.Cursor = [Avalonia.Input.Cursor]::new([Avalonia.Input.StandardCursorType]::Hand)
            [Avalonia.Controls.ToolTip]::SetTip($cloudBtn, 'Sign in to US Government (GCC High / DoD) or China cloud. To make a choice permanent, change Default cloud in Settings.')
            $cloudBtn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                $ui.HidePopup()
                try {
                    $picked = Show-CloudPickerMenu
                    if ($picked) {
                        Invoke-AvaloniaDeferredAction -Argument $picked -Action {
                            param($cloud)
                            Invoke-AvaloniaInteractiveSignIn -Cloud $cloud
                        }
                    }
                } catch {
                    Write-LogError 'Cloud picker sign-in failed' $_.Exception
                    Write-Status ''
                }
            }))
            $grdLoginAccount.Children.Add($cloudBtn) | Out-Null
        }
    }

    # Tenant accounts grid — visible only when $script:AccessibleTenants has
    # >1 entry (MSAL populates this from /organization?$select=... when the
    # GetTenantList setting is on). Each non-current tenant is a button that
    # re-acquires via Connect-EntraEnvironment -TenantId; the current tenant
    # is shown as a highlighted text block.
    $grdTenants = (Get-AvaloniaHost)::FindByName($popup, 'grdTenantAccounts')
    if ($grdTenants -and @($script:AccessibleTenants).Count -gt 1) {
        $header = [Avalonia.Controls.TextBlock]::new()
        $header.Text       = 'Tenants:'
        $header.FontWeight = [Avalonia.Media.FontWeight]::Bold
        $header.Margin     = [Avalonia.Thickness]::new(0, 5, 0, 0)
        $grdTenants.Children.Add($header) | Out-Null

        $tenants = if ((Get-SettingValue 'SortTenantList') -eq $true) {
            $script:AccessibleTenants | Sort-Object -Property DisplayName
        } else {
            $script:AccessibleTenants
        }

        $currentTenantId = if ($userInfo) { $userInfo.TenantId } else { $null }

        foreach ($tenant in $tenants) {
            try {
                # Build a TextBlock with three Runs (name bold, then domain, then id).
                $tb = [Avalonia.Controls.TextBlock]::new()
                $tb.HorizontalAlignment = [Avalonia.Layout.HorizontalAlignment]::Stretch

                $nameRun = [Avalonia.Controls.Documents.Run]::new([string]$tenant.DisplayName)
                $nameRun.FontWeight = [Avalonia.Media.FontWeight]::Bold
                $tb.Inlines.Add($nameRun) | Out-Null
                $tb.Inlines.Add([Avalonia.Controls.Documents.LineBreak]::new()) | Out-Null
                $tb.Inlines.Add([Avalonia.Controls.Documents.Run]::new([string]$tenant.defaultDomain)) | Out-Null
                $tb.Inlines.Add([Avalonia.Controls.Documents.LineBreak]::new()) | Out-Null
                $tb.Inlines.Add([Avalonia.Controls.Documents.Run]::new([string]$tenant.tenantId)) | Out-Null

                if ($tenant.tenantId -ne $currentTenantId) {
                    # Other tenant — clickable Button switches Entra context.
                    $btn = [Avalonia.Controls.Button]::new()
                    $btn.Content                    = $tb
                    $btn.HorizontalAlignment        = [Avalonia.Layout.HorizontalAlignment]::Stretch
                    $btn.HorizontalContentAlignment = [Avalonia.Layout.HorizontalAlignment]::Left
                    $btn.Margin                     = [Avalonia.Thickness]::new(0, 5, 0, 0)
                    $btn.Cursor                     = [Avalonia.Input.Cursor]::new([Avalonia.Input.StandardCursorType]::Hand)
                    $btn.Tag                        = $tenant
                    $btn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                        param($S, $E)
                        $t = $S.Tag
                        $ui.HidePopup()
                        Invoke-AvaloniaDeferredAction -Argument $t -Action {
                            param($selectedTenant)
                            Invoke-AvaloniaTenantSwitch $selectedTenant
                        }
                    }))
                    $grdTenants.Children.Add($btn) | Out-Null
                } else {
                    # Current tenant — non-clickable highlighted row.
                    $border = [Avalonia.Controls.Border]::new()
                    $border.Padding = [Avalonia.Thickness]::new(5, 2, 5, 2)
                    $border.Margin  = [Avalonia.Thickness]::new(0, 5, 0, 0)
                    $border.HorizontalAlignment = [Avalonia.Layout.HorizontalAlignment]::Stretch
                    # Try the theme's selected-row brush; fall back to a subtle white tint.
                    $bg = $null
                    [void][Avalonia.Application]::Current.Resources.TryGetResource('SelectedRowBackgroundColor', $null, [ref]$bg)
                    if ($bg) {
                        $border.Background = $bg
                    } else {
                        $border.Background = [Avalonia.Media.SolidColorBrush]::new([Avalonia.Media.Color]::Parse('#22FFFFFF'))
                    }
                    $border.Child = $tb
                    $grdTenants.Children.Add($border) | Out-Null
                }
            } catch {
                Write-LogDebug "Tenant row build failed for $($tenant.DisplayName): $($_.Exception.Message)"
            }
        }
    }

    Set-PopupPosition $popup $Anchor 320

    # Stop dim-layer click-outside dismissal from firing when the user clicks
    # inside the popup body. Show-Popup checks $e.Source against grdPopup /
    # cvsPopup; marking handled at the popup root keeps it from reaching either.
    $popup.add_PointerPressed((ConvertTo-AvaloniaEventScriptBlock { param($S, $E) $E.Handled = $true }))

    $ui.ShowPopup($popup)
}

function Show-CachedAccountPicker {
    param($Anchor, $CachedRows)

    $ui = $script:UIProvider
    $popup = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/LoginPanel.axaml'))
    if (-not $popup) { return }

    $grdAccounts = (Get-AvaloniaHost)::FindByName($popup, 'grdAccounts')
    if (-not $grdAccounts) { return }

    foreach ($row in $CachedRows) {
        if ($row.Native -and (Test-IsPersonalMSAAccount $row.Native)) { continue }
        Add-CachedUser $row.Native $grdAccounts
    }

    # Footer: "Sign in with a different account" — closes the picker and goes
    # straight to the active provider's interactive flow.
    $altBtn = [Avalonia.Controls.Button]::new()
    $altBtn.Content = 'Sign in with a different account'
    $altBtn.Margin  = [Avalonia.Thickness]::new(0, 10, 0, 0)
    $altBtn.HorizontalAlignment = [Avalonia.Layout.HorizontalAlignment]::Stretch
    $altBtn.Cursor = [Avalonia.Input.Cursor]::new([Avalonia.Input.StandardCursorType]::Hand)
    $altBtn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
        $ui.HidePopup()
        Invoke-AvaloniaDeferredAction -Action {
            param($unused)
            Invoke-AvaloniaInteractiveSignIn
        }
    }))
    $grdAccounts.Children.Add($altBtn) | Out-Null

    # Cloud picker: pick a non-default cloud (GCC High, DoD, China) before
    # sign-in. Only useful when $script:Clouds has >1 entry; otherwise hide.
    if (@($script:Clouds).Count -gt 1) {
        $cloudBtn = [Avalonia.Controls.Button]::new()
        $cloudBtn.Content = 'Sign in to a different cloud...'
        $cloudBtn.Margin  = [Avalonia.Thickness]::new(0, 5, 0, 0)
        $cloudBtn.HorizontalAlignment = [Avalonia.Layout.HorizontalAlignment]::Stretch
        $cloudBtn.Cursor = [Avalonia.Input.Cursor]::new([Avalonia.Input.StandardCursorType]::Hand)
        $cloudBtn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $ui.HidePopup()
            try {
                $picked = Show-CloudPickerMenu
                if ($picked) {
                    Invoke-AvaloniaDeferredAction -Argument $picked -Action {
                        param($cloud)
                        Invoke-AvaloniaInteractiveSignIn -Cloud $cloud
                    }
                }
            } catch {
                Write-LogError 'Cloud picker sign-in failed' $_.Exception
                Write-Status ''
            }
        }))
        $grdAccounts.Children.Add($cloudBtn) | Out-Null
    }

    Set-PopupPosition $popup $Anchor 340

    $popup.add_PointerPressed((ConvertTo-AvaloniaEventScriptBlock { param($S, $E) $E.Handled = $true }))

    $ui.ShowPopup($popup)
}

function Show-AuthenticationInfo {
    if (-not $script:grdMenu) { return }

    # Snapshot then remove — Avalonia Children is a live collection and removing
    # while iterating misses elements.
    $existing = @()
    foreach ($child in $script:grdMenu.Children) {
        if ($child.Tag -eq 'ProfilePicture') { $existing += $child }
    }
    foreach ($child in $existing) { [void]$script:grdMenu.Children.Remove($child) }

    $profileObj = Get-AvaloniaUserProfile
    if ($profileObj) {
        $profileObj.Tag = 'ProfilePicture'
        [Avalonia.Controls.Grid]::SetColumn($profileObj, 2)
        $script:grdMenu.Children.Add($profileObj) | Out-Null
    }
}

function Set-EnvironmentInfo {
    param([string]$TenantName)

    if ([string]::IsNullOrWhiteSpace($TenantName)) {
        $TenantName = $script:OrganizationName
    }

    if (-not $script:borderEnvBadge -or -not $script:txtEnvBadge) { return }

    $signedIn = $false
    try {
        $activeProvider = Get-AuthProvider
        if ($activeProvider) {
            $defId = Get-DefaultTokenId
            if ($defId) { $signedIn = ($null -ne $activeProvider.GetUserInfo($defId)) }
            if (-not $signedIn) { $signedIn = ($null -ne $activeProvider.GetUserInfo(0)) }
        }
    } catch { $signedIn = $false }

    # An expired-but-still-registered token still returns GetUserInfo, so treat an
    # expired default token as "not signed in" - keeps the badge in step with the
    # avatar, which also reverts to the Sign-in button on expiry.
    if ($signedIn -and (Get-Command Test-DefaultTokenExpired -ErrorAction SilentlyContinue)) {
        try { if (Test-DefaultTokenExpired) { $signedIn = $false } } catch { }
    }

    if (-not $signedIn) {
        $script:borderEnvBadge.IsVisible = $false
        $script:txtEnvBadge.Text = ''
        return
    }

    $envText  = Get-SettingValue 'EnvironmentText'
    $envColor = Get-SettingValue 'EnvironmentColor'
    $showOrg  = (Get-SettingValue 'MenuShowOrganizationName') -eq $true

    $parts = @()
    if ($envText)                  { $parts += $envText }
    if ($TenantName -and $showOrg) { $parts += $TenantName }
    $badgeText = $parts -join ' - '

    if ([string]::IsNullOrWhiteSpace($badgeText)) {
        $script:borderEnvBadge.IsVisible = $false
        $script:txtEnvBadge.Text = ''
        return
    }

    $script:txtEnvBadge.Text = $badgeText

    if ($envText -and $envColor) {
        try {
            $c = [Avalonia.Media.Color]::Parse($envColor)
            $script:borderEnvBadge.Background = [Avalonia.Media.SolidColorBrush]::new($c)
            $luminance = 0.299 * $c.R + 0.587 * $c.G + 0.114 * $c.B
            $fgHex = if ($luminance -gt 128) { '#FF1A1A1A' } else { '#FFEEEEEE' }
            $script:txtEnvBadge.Foreground = [Avalonia.Media.SolidColorBrush]::new([Avalonia.Media.Color]::Parse($fgHex))
        } catch { }
    } else {
        $script:borderEnvBadge.Background = [Avalonia.Media.Brushes]::Transparent
        $script:txtEnvBadge.ClearValue([Avalonia.Controls.TextBlock]::ForegroundProperty)
    }

    $script:borderEnvBadge.IsVisible = $true
}

#endregion

#region Debug views — Log / Cached Objects / Graph Calls

# Avalonia counterparts of Get-LogViewPanel / Get-CachedObjectsViewPanel /
# Get-GraphCallsViewPanel from UI/WPF/Extensions/CoreUIWPF.ps1. Hosted by the
# three ViewObject subclasses in Classes/CoreUIBaseClassesAvalonia.ps1; the
# binder needs CLR row classes (LogRowItem / CacheRowItem / GraphCallRowItem /
# GraphBatchRequestRowItem in Classes/) because $script:LogItems /
# $script:cacheObjects / $script:AllGraphCalls hold PSCustomObjects that the
# Avalonia DataTemplates can't reach (see [[avalonia-binding-needs-clr-types]]).

function Get-LogViewPanel {
    $ui = $script:UIProvider
    $viewPanel = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/LogInfo.axaml'))
    if (-not $viewPanel) { return $null }

    Update-LogView -ViewPanel $viewPanel
    return $viewPanel
}

function Update-LogView {
    param($ViewPanel)

    if (-not $ViewPanel) { return }

    # Row-color brushes match the WPF DataTrigger palette: Type 2 -> Orange
    # (warning), Type 3 -> Red (error). Computed once per refresh and shared
    # across every projected row so the binder only sees a couple of brush
    # instances. Foreground=$null falls back to the inherited theme color.
    $orange = [Avalonia.Media.SolidColorBrush]::new([Avalonia.Media.Color]::FromRgb(255, 165, 0))
    $red    = [Avalonia.Media.SolidColorBrush]::new([Avalonia.Media.Colors]::Red)

    $rows = @()
    if ($script:LogItems) {
        $rows = @($script:LogItems | ForEach-Object {
            $row = [LogRowItem]::new()
            $row.ID       = $_.ID
            $row.DateTime = $_.DateTime
            $row.Type     = $_.Type
            $row.TypeText = $_.TypeText
            $row.Text     = $_.Text
            $row.RowForeground = switch ($_.Type) {
                2       { $orange }
                3       { $red }
                default { $null }
            }
            $row
        })
    }

    $dg = (Get-AvaloniaHost)::FindByName($ViewPanel, 'dgLogInfo')
    if ($dg) {
        $dg.ItemsSource = $null
        $dg.ItemsSource = [System.Collections.IEnumerable]$rows
    }
}

function Format-CacheByteSize {
    param([long]$Bytes)

    if ($Bytes -lt 1KB) { return ("{0:N0} B"  -f $Bytes) }
    if ($Bytes -lt 1MB) { return ("{0:N1} KB" -f ($Bytes / 1KB)) }
    if ($Bytes -lt 1GB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    return ("{0:N2} GB" -f ($Bytes / 1GB))
}

function Get-CachedObjectsViewPanel {
    $ui = $script:UIProvider
    $viewPanel = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/CachedObjects.axaml'))
    if (-not $viewPanel) { return $null }

    Update-CachedObjectsView -ViewPanel $viewPanel

    # SelectionChanged on dgCachedObjects updates the detail TextBox. The
    # detail Grid binds DataContext to SelectedItem so most fields refresh
    # automatically; the value pane gets special-case formatting for
    # Hashtable values (one entry per line, "key - count objects").
    $dgCachedObjects = (Get-AvaloniaHost)::FindByName($viewPanel, 'dgCachedObjects')
    $txtCacheObjectInfo = (Get-AvaloniaHost)::FindByName($viewPanel, 'txtCacheObjectInfo')
    if ($dgCachedObjects -and $txtCacheObjectInfo) {
        # Module-bound handlers lose function-local captures, so the target
        # TextBox travels on the grid's Tag (same pattern as btnClearCachedObject).
        $dgCachedObjects.Tag = $txtCacheObjectInfo
        $dgCachedObjects.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $txt = $s.Tag
            if (-not $txt) { return }
            $sel = $s.SelectedItem
            if (-not $sel) { $txt.Text = ''; return }
            if ($sel.Value -is [Hashtable]) {
                $txt.Text = ($sel.Value.Keys | ForEach-Object {
                    "$_ - $($sel.Value[$_].Count) objects"
                }) -join "`n"
            } else {
                $txt.Text = "$($sel.Value)"
            }
        }))
    }

    $btnClear = (Get-AvaloniaHost)::FindByName($viewPanel, 'btnClearCachedObject')
    if ($btnClear) {
        $btnClear.Tag = $viewPanel
        $btnClear.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $rootPanel = $s.Tag
            $dg = (Get-AvaloniaHost)::FindByName($rootPanel, 'dgCachedObjects')
            $sel = if ($dg) { $dg.SelectedItem } else { $null }
            if ($sel -and $sel.Persistent -ne $true) {
                $answer = Show-MessageBox -Text "Are you sure you want to remove $($sel.Name)?" -Caption 'Clear object from cache?' -Button 'YesNo' -Icon 'Question'
                if ($answer -eq 'Yes') {
                    Clear-CacheObject -Name $sel.Name
                    Update-CachedObjectsView -ViewPanel $rootPanel
                }
            }
        }))
    }

    $btnRefresh = (Get-AvaloniaHost)::FindByName($viewPanel, 'btnRefreshCachedObjects')
    if ($btnRefresh) {
        $btnRefresh.Tag = $viewPanel
        $btnRefresh.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            Update-CachedObjectsView -ViewPanel $s.Tag
        }))
    }

    return $viewPanel
}

function Update-CachedObjectsView {
    param($ViewPanel)

    $ui = $script:UIProvider
    if (-not $ViewPanel) { return }

    $totalBytes = [long]0

    $rows = @()
    if ($script:cacheObjects -and $script:cacheObjects.Values) {
        $rows = @($script:cacheObjects.Values | ForEach-Object {
            $sz = Get-CacheObjectSize $_.Value
            $totalBytes += $sz
            $tagsText = if ($_.Tags) { ($_.Tags -join ', ') } else { '' }

            $valueText =
                if ($null -eq $_.Value) { '' }
                elseif ($_.Value -is [Hashtable]) { "Hashtable ($($_.Value.Count) keys)" }
                else { "$($_.Value)" }

            $row = [CacheRowItem]::new()
            $row.Name       = $_.Name
            $row.Tags       = $_.Tags
            $row.TagsText   = $tagsText
            $row.Value      = $_.Value
            $row.ValueText  = $valueText
            $row.Persistent = [bool]$_.Persistent
            if ($_.TimeOut) { $row.TimeOut = $_.TimeOut }
            $row.Size       = $sz
            $row.SizeText   = (Format-CacheByteSize $sz)
            $row
        })
    }

    $dg = (Get-AvaloniaHost)::FindByName($ViewPanel, 'dgCachedObjects')
    if ($dg) {
        $dg.ItemsSource = $null
        $dg.ItemsSource = [System.Collections.IEnumerable]$rows
    }

    $stats = Get-CacheStats
    $ui.SetXamlProperty($ViewPanel, 'txtCacheEntries', 'Text', ('{0:N0}' -f $stats.Entries))
    $ui.SetXamlProperty($ViewPanel, 'txtCacheHits',    'Text', ('{0:N0}' -f $stats.Hits))
    $ui.SetXamlProperty($ViewPanel, 'txtCacheMisses',  'Text', ('{0:N0}' -f $stats.Misses))
    $ui.SetXamlProperty($ViewPanel, 'txtCacheBytes',   'Text', (Format-CacheByteSize $totalBytes))
}

function Format-GraphTotalDuration {
    param([double]$Milliseconds)

    if ($Milliseconds -lt 1000) { return ('{0:N0} ms' -f $Milliseconds) }

    $totalSeconds = [int]($Milliseconds / 1000)
    if ($totalSeconds -lt 60) { return ('{0:N1} s' -f ($Milliseconds / 1000)) }

    $h = [int]($totalSeconds / 3600)
    $m = [int]((($totalSeconds) % 3600) / 60)
    $s = $totalSeconds % 60
    if ($h -gt 0) { return ('{0}h {1:D2}m {2:D2}s' -f $h, $m, $s) }
    return ('{0}m {1:D2}s' -f $m, $s)
}

function Get-HttpStatusDescription {
    param([string]$Code)

    # Verbatim port from UI/WPF/Extensions/CoreUIWPF.ps1 — descriptions
    # oriented toward what each code means specifically for Microsoft Graph.
    switch ($Code)
    {
        '200' { 'OK - Request succeeded' }
        '201' { 'Created - New resource was created' }
        '202' { 'Accepted - Request accepted; processing is async' }
        '204' { 'No Content - Request succeeded; nothing to return (typical for DELETE, PATCH)' }
        '301' { 'Moved Permanently - Resource has a new permanent URL' }
        '302' { 'Found - Temporary redirect' }
        '304' { 'Not Modified - Cached copy is still valid (conditional GET)' }
        '307' { 'Temporary Redirect - Use new URL for this request only' }
        '308' { 'Permanent Redirect - Use new URL going forward' }
        '400' { 'Bad Request - Malformed request (invalid filter/expand/JSON)' }
        '401' { 'Unauthorized - Token missing, expired, or invalid' }
        '402' { 'Payment Required - Tenant license issue' }
        '403' { 'Forbidden - Caller lacks permission for this resource' }
        '404' { "Not Found - Resource (policy, user, group) doesn't exist" }
        '405' { 'Method Not Allowed - HTTP verb not supported on this endpoint' }
        '406' { 'Not Acceptable - Server cannot produce the requested format' }
        '408' { 'Request Timeout - Client took too long to send the request' }
        '409' { 'Conflict - State conflict (e.g. duplicate name, concurrent modification)' }
        '410' { 'Gone - Resource was deleted' }
        '411' { 'Length Required - Missing Content-Length header' }
        '412' { "Precondition Failed - If-Match/If-None-Match header didn't match" }
        '413' { 'Payload Too Large - Request body exceeds the limit' }
        '415' { 'Unsupported Media Type - Content-Type not accepted' }
        '416' { 'Range Not Satisfiable - Requested byte range is invalid' }
        '422' { 'Unprocessable Entity - Request is well-formed but semantically wrong' }
        '423' { 'Locked - Resource is locked by another operation' }
        '429' { 'Too Many Requests - Graph throttling; back off and retry' }
        '500' { 'Internal Server Error - Graph backend failure' }
        '501' { 'Not Implemented - Endpoint or feature not supported' }
        '502' { 'Bad Gateway - Upstream service error' }
        '503' { 'Service Unavailable - Graph or downstream service is unavailable' }
        '504' { 'Gateway Timeout - Upstream service timed out' }
        '507' { 'Insufficient Storage - Service-side storage limit exceeded' }
        '509' { 'Bandwidth Limit Exceeded - Tenant or app bandwidth quota hit' }
        'n/a' { 'No response - Request failed before a status code was returned' }
        default { "HTTP $Code" }
    }
}

function Build-GraphStatusCodePanel {
    # Populates a WrapPanel with one clickable Button per status code. Click
    # sets $script:_graphCallsFilter and re-runs Update-GraphCallsView so the
    # grid + indicator both refresh in one pass. WPF used a LinkButton style;
    # Avalonia uses Themes/Styles.axaml's Button.link class for the same look.
    param(
        $ViewPanel,
        [string]$PanelName,
        [System.Collections.Generic.Dictionary[string,int]]$Counts,
        [string]$ActiveCode
    )

    $panel = (Get-AvaloniaHost)::FindByName($ViewPanel, $PanelName)
    if (-not $panel) { return }
    $panel.Children.Clear()

    if (-not $Counts -or $Counts.Count -eq 0) {
        $tb = [Avalonia.Controls.TextBlock]::new()
        $tb.Text = '-'
        $tb.Margin = [Avalonia.Thickness]::new(0, 0, 4, 0)
        [void]$panel.Children.Add($tb)
        return
    }

    foreach ($key in ($Counts.Keys | Sort-Object)) {
        $btn = [Avalonia.Controls.Button]::new()
        $btn.Content = ('{0}: {1:N0}' -f $key, $Counts[$key])
        $btn.Margin  = [Avalonia.Thickness]::new(0, 0, 12, 0)
        $btn.Padding = [Avalonia.Thickness]::new(0)
        [void]$btn.Classes.Add('link')

        $tt = (Get-HttpStatusDescription $key) + "`n`nClick to filter the grid by this status code (includes batch jobs containing a matching sub-request)."
        [Avalonia.Controls.ToolTip]::SetTip($btn, $tt)

        if ($ActiveCode -and $key -eq $ActiveCode) {
            $btn.FontWeight = [Avalonia.Media.FontWeight]::Bold
        }

        $btn.Tag = [PSCustomObject]@{ Code = $key; ViewPanel = $ViewPanel }
        $btn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $script:_graphCallsFilter = $s.Tag.Code
            Update-GraphCallsView -ViewPanel $s.Tag.ViewPanel
        }))

        [void]$panel.Children.Add($btn)
    }
}

function Get-GraphCallsViewPanel {
    $ui = $script:UIProvider
    $xamlPath = Join-Path $script:AppUIRootFolder 'XAML/GraphCallsPanel.axaml'
    $viewPanel = $ui.GetXamlObject($xamlPath)
    if (-not $viewPanel) {
        Write-LogError "Get-GraphCallsViewPanel: XAML failed to load from '$xamlPath'." $null
        return $null
    }

    try {
        Update-GraphCallsView -ViewPanel $viewPanel
    }
    catch {
        Write-LogError 'Get-GraphCallsViewPanel: Update-GraphCallsView threw' $_.Exception
    }

    $btnRefresh = (Get-AvaloniaHost)::FindByName($viewPanel, 'btnRefreshGraphCalls')
    if ($btnRefresh) {
        $btnRefresh.Tag = $viewPanel
        $btnRefresh.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $script:_graphCallsFilter = $null
            Update-GraphCallsView -ViewPanel $s.Tag
        }))
    }

    $btnClearFilter = (Get-AvaloniaHost)::FindByName($viewPanel, 'btnClearGraphCallsFilter')
    if ($btnClearFilter) {
        $btnClearFilter.Tag = $viewPanel
        $btnClearFilter.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $script:_graphCallsFilter = $null
            Update-GraphCallsView -ViewPanel $s.Tag
        }))
    }

    return $viewPanel
}

function Update-GraphCallsView {
    param($ViewPanel)

    $ui = $script:UIProvider
    if (-not $ViewPanel) { return }

    $calls = if ($script:AllGraphCalls) { @($script:AllGraphCalls) } else { @() }

    # Apply the active code filter (if any). A call matches when the top-level
    # StatusCode equals the code OR any of its batch sub-requests does. Stays
    # outside the totals/breakdown loop so the breakdown always shows the full
    # distribution — that's how the user discovers what to filter on.
    $filterCode = $script:_graphCallsFilter
    $callsForGrid = $calls
    if ($filterCode) {
        $callsForGrid = @($calls | Where-Object {
            if ("$($_.StatusCode)" -eq $filterCode) { return $true }
            if ($_.IsBatch -and $_.BatchRequests) {
                foreach ($br in $_.BatchRequests) {
                    if ("$($br.Response.StatusCode)" -eq $filterCode) { return $true }
                }
            }
            return $false
        })
    }

    $totalBytes      = [long]0
    $totalObjects    = [long]0
    $totalDurationMs = [double]0
    $httpStatuses    = [System.Collections.Generic.Dictionary[string,int]]::new()
    $batchStatuses   = [System.Collections.Generic.Dictionary[string,int]]::new()

    # Single pass: aggregate totals + per-code counts across the FULL call
    # list, project filtered calls into GraphCallRowItem instances for the
    # grid. Two collections so totals reflect the unfiltered population
    # while the grid honors the active filter.
    $rowsForGrid = [System.Collections.Generic.List[object]]::new()

    foreach ($call in $calls) {
        if ($call.KB)          { $totalBytes      += [long]([double]$call.KB * 1024) }
        if ($call.ObjectCount) { $totalObjects   += [long]$call.ObjectCount }
        if ($call.Duration)    { $totalDurationMs += [double]$call.Duration }

        $code = if ($null -ne $call.StatusCode -and "$($call.StatusCode)" -ne '') { "$($call.StatusCode)" } else { 'n/a' }
        if (-not $httpStatuses.ContainsKey($code)) { $httpStatuses[$code] = 0 }
        $httpStatuses[$code]++

        $errSummary = ''
        $batchRows  = [System.Collections.Generic.List[GraphBatchRequestRowItem]]::new()
        if ($call.IsBatch -and $call.BatchRequests) {
            $errorCounts = [System.Collections.Generic.Dictionary[string,int]]::new()
            foreach ($item in $call.BatchRequests) {
                $bcode = if ($null -ne $item.Response.StatusCode) { "$($item.Response.StatusCode)" } else { 'n/a' }
                if (-not $batchStatuses.ContainsKey($bcode)) { $batchStatuses[$bcode] = 0 }
                $batchStatuses[$bcode]++

                $isErr = $false
                if ($bcode -eq 'n/a') { $isErr = $true }
                else {
                    $n = 0
                    if ([int]::TryParse($bcode, [ref]$n) -and $n -ge 400) { $isErr = $true }
                }
                if ($isErr) {
                    if (-not $errorCounts.ContainsKey($bcode)) { $errorCounts[$bcode] = 0 }
                    $errorCounts[$bcode]++
                }

                # Project each batch sub-request into its CLR row class so the
                # nested DataGrid (bound via SelectedItem.BatchRequests) can
                # resolve column paths. The WPF version walked through PSObject
                # via {Binding Response.StatusCode}; Avalonia's binder doesn't.
                $br = [GraphBatchRequestRowItem]::new()
                $br.Id          = "$($item.Id)"
                if ($item.Response -and $null -ne $item.Response.KB)          { $br.KB          = [double]$item.Response.KB }
                if ($item.Response -and $null -ne $item.Response.ObjectCount) { $br.ObjectCount = [int]$item.Response.ObjectCount }
                if ($item.Response -and $null -ne $item.Response.PageCount)   { $br.PageCount   = [int]$item.Response.PageCount }
                $br.StatusCode  = if ($item.Response) { "$($item.Response.StatusCode)" } else { '' }
                $br.Method      = "$($item.Method)"
                $br.URL         = "$($item.URL)"
                [void]$batchRows.Add($br)
            }
            if ($errorCounts.Count -gt 0) {
                $errSummary = (($errorCounts.Keys | Sort-Object | ForEach-Object { "$($_):$($errorCounts[$_])" }) -join ', ')
            }
        }

        # Stash on the source PSCustomObject so external WPF callers (and any
        # legacy code that still reads the Note) keep seeing the same shape.
        $call | Add-Member -MemberType NoteProperty -Name 'BatchErrorSummary' -Value $errSummary -Force

        # Filter happens here, after totals are aggregated.
        $passes = $true
        if ($filterCode) {
            $passes = $false
            if ($code -eq $filterCode) { $passes = $true }
            elseif ($call.IsBatch -and $call.BatchRequests) {
                foreach ($br2 in $call.BatchRequests) {
                    if ("$($br2.Response.StatusCode)" -eq $filterCode) { $passes = $true; break }
                }
            }
        }
        if (-not $passes) { continue }

        $row = [GraphCallRowItem]::new()
        $row.Provider          = "$($call.Provider)"
        if ($call.Time)        { $row.Time = $call.Time }
        if ($null -ne $call.Duration)    { $row.Duration    = [double]$call.Duration }
        if ($null -ne $call.KB)          { $row.KB          = [double]$call.KB }
        if ($null -ne $call.ObjectCount) { $row.ObjectCount = [int]$call.ObjectCount }
        if ($null -ne $call.PageCount)   { $row.PageCount   = [int]$call.PageCount }
        $row.StatusCode        = $code
        $row.BatchErrorSummary = $errSummary
        $row.Method            = "$($call.Method)"
        $row.URL               = "$($call.URL)"
        $row.IsBatch           = [bool]$call.IsBatch
        foreach ($br in $batchRows) { [void]$row.BatchRequests.Add($br) }
        [void]$rowsForGrid.Add($row)
    }

    $dgGraphCalls = (Get-AvaloniaHost)::FindByName($ViewPanel, 'dgGraphCalls')
    if ($dgGraphCalls) {
        $dgGraphCalls.ItemsSource = $null
        $dgGraphCalls.ItemsSource = [System.Collections.IEnumerable]$rowsForGrid
    }

    $ui.SetXamlProperty($ViewPanel, 'txtTotalCalls',   'Text', ('{0:N0}' -f $calls.Count))
    $ui.SetXamlProperty($ViewPanel, 'txtTotalMB',      'Text', ('{0:N2} MB' -f ($totalBytes / 1MB)))
    $ui.SetXamlProperty($ViewPanel, 'txtTotalObjects', 'Text', ('{0:N0}' -f $totalObjects))
    $ui.SetXamlProperty($ViewPanel, 'txtTotalTime',    'Text', (Format-GraphTotalDuration $totalDurationMs))

    Build-GraphStatusCodePanel -ViewPanel $ViewPanel -PanelName 'pnlHttpStatuses'  -Counts $httpStatuses  -ActiveCode $filterCode
    Build-GraphStatusCodePanel -ViewPanel $ViewPanel -PanelName 'pnlBatchStatuses' -Counts $batchStatuses -ActiveCode $filterCode

    $indicator = (Get-AvaloniaHost)::FindByName($ViewPanel, 'pnlFilterIndicator')
    if ($indicator) {
        if ($filterCode) {
            $shown = $rowsForGrid.Count
            $ui.SetXamlProperty($ViewPanel, 'txtActiveFilter', 'Text', ('status code {0} - showing {1:N0} of {2:N0} call(s)' -f $filterCode, $shown, $calls.Count))
            $indicator.IsVisible = $true
        } else {
            $ui.SetXamlProperty($ViewPanel, 'txtActiveFilter', 'Text', '')
            $indicator.IsVisible = $false
        }
    }
}

#endregion

#region DataGrid select-all header
#
# WPF parity for the select/deselect-all column header. Avalonia's
# DataGridCheckBoxColumn doesn't expose a bindable header CheckBox, so the
# bulk forms used to carry a separate `chkSelectAll` on the toolbar — the user
# called this out as a regression vs. the WPF UX. The pattern is:
#   1. AXAML declares a DataGridTemplateColumn whose Header is a CheckBox
#      (Header object can be any Avalonia Control) and whose CellTemplate is
#      a CheckBox bound to <BindingProperty> with TwoWay mode.
#   2. ps1 calls Initialize-AvaloniaGridSelectAllHeader after the form's
#      DataGrid is in the visual tree; the helper locates the header CheckBox
#      and wires IsCheckedChanged -> Invoke-AvaloniaGridSelectAllToggle.
#
# Per-grid state lives in $script:_avaloniaGridSelectAllMap so the click
# handler can re-resolve everything by reference. Closures in
# .GetNewClosure() scriptblocks get scrubbed by
# ConvertTo-AvaloniaEventScriptBlock (see feedback_avalonia_closure_dynamic_module),
# so we route through a module-scope function + state instead.

function Initialize-AvaloniaGridSelectAllHeader {
    param(
        [Parameter(Mandatory)]$Grid,
        [string]$BindingProperty = 'Selected',
        [Nullable[bool]]$InitiallyChecked = $true
    )

    if (-not $Grid) { return $null }

    # AXAML carries a DataGridTemplateColumn whose Header is a CheckBox. Pull
    # it out without depending on names — Avalonia DataGrid header names live
    # in a separate scope and aren't reliably findable from the form's root.
    $headerCb = $null
    foreach ($col in $Grid.Columns) {
        if ($col.Header -is [Avalonia.Controls.CheckBox]) { $headerCb = $col.Header; break }
    }
    if (-not $headerCb) { return $null }

    if ($null -ne $InitiallyChecked) { $headerCb.IsChecked = [bool]$InitiallyChecked }

    if (-not $script:_avaloniaGridSelectAllMap) { $script:_avaloniaGridSelectAllMap = @{} }
    # Hashtable keys for reference-typed objects use ReferenceEqualityComparer
    # by default, so the same CheckBox always hits the same entry.
    $script:_avaloniaGridSelectAllMap[$headerCb] = [PSCustomObject]@{
        Grid            = $Grid
        BindingProperty = $BindingProperty
    }

    $headerCb.add_IsCheckedChanged((ConvertTo-AvaloniaEventScriptBlock {
        # $this is the sender control (the header CheckBox) inside an Avalonia
        # event handler — see CoreUIAvalonia's ConvertTo-AvaloniaEventScriptBlock.
        Invoke-AvaloniaGridSelectAllToggle -HeaderCheckBox $this
    }))

    return $headerCb
}

function Invoke-AvaloniaGridSelectAllToggle {
    param([Parameter(Mandatory)]$HeaderCheckBox)
    if (-not $script:_avaloniaGridSelectAllMap -or
        -not $script:_avaloniaGridSelectAllMap.ContainsKey($HeaderCheckBox)) { return }

    $entry = $script:_avaloniaGridSelectAllMap[$HeaderCheckBox]
    $grid  = $entry.Grid
    $prop  = $entry.BindingProperty
    if (-not $grid -or -not $grid.ItemsSource) { return }

    $newVal = [bool]$HeaderCheckBox.IsChecked
    $touched = $false
    foreach ($row in @($grid.ItemsSource)) {
        if (-not $row) { continue }
        if ($row.PSObject.Properties[$prop]) {
            $row.$prop = $newVal
            $touched = $true
        }
    }
    if (-not $touched) { return }

    # PSCustomObject rows don't raise PropertyChanged, so the cell CheckBoxes
    # don't repaint automatically. Re-binding the ItemsSource forces the grid
    # to re-realise its cells — same pattern as the per-form chkSelectAll
    # workarounds we're replacing.
    $current = @($grid.ItemsSource)
    $grid.ItemsSource = $null
    $grid.ItemsSource = $current
}

#endregion

#region UI settings registration


function Invoke-CoreUIEventSettingValueUpdated {
    # Apply UI-affecting setting changes immediately when the user saves them
    # instead of waiting for an app restart. Matches WPF dispatch.
    param($SettingInfo, $NewValue, $OldValue)

    $ui = $script:UIProvider
    if ($SettingInfo.Key -eq 'EnvironmentText' -or $SettingInfo.Key -eq 'EnvironmentColor') {
        try { $ui.SetEnvironmentInfo() } catch { Write-LogError 'Set-EnvironmentInfo failed in SettingValueUpdated' $_.Exception }
    }
    elseif ($SettingInfo.Key -eq 'AppTheme') {
        try { $ui.SetAppTheme($NewValue) } catch { Write-LogError 'Set-AppTheme failed in SettingValueUpdated' $_.Exception }
    }
}

#endregion

# Top-level registration — runs when the module loads the file. Settings have
# to be registered before the Settings dialog is opened; the event + handler
# pairing is what makes "Save" auto-apply theme / env-badge changes.
# Shared UI settings (AppTheme, HideNoAccess, environment badge, ...) are
# registered once for both backends in UI/Classes/UICommonSettings.ps1.
Add-AppEvent "SettingValueUpdated"
Add-AppEventHandler "SettingValueUpdated" "Invoke-CoreUIEventSettingValueUpdated"
