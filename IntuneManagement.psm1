
#Set-StrictMode -Version Latest
$script:AppRootFolder = $PSScriptRoot
$script:SessionID = [Guid]::NewGuid()
$PSModuleAutoloadingPreference = "none"

# Unblock-File clears the NTFS mark-of-the-web (Zone.Identifier) stream, which
# only exists on Windows. On Linux/macOS the cmdlet has nothing to do and errors.
# $IsWindows is defined only on PowerShell Core; Windows PowerShell 5.1 is always
# Windows (Desktop edition), so treat that as Windows too.
$script:IsWindowsOS = $IsWindows -or $PSVersionTable.PSEdition -eq 'Desktop'

# Windows PowerShell 5.1 does not load System.Net.Http into the runspace by
# default, but class files (AuthenticationMgGraph) reference its types at class
# compile time. PS7 has it loaded already; the call is a cheap no-op there.
try { Add-Type -AssemblyName System.Net.Http -ErrorAction Stop } catch { }

$dependencyFiles = @()
$ImportOrderStr = "#ImportOrder "
$useDotSoure = $true
$loadedFiles = @()

$preloadFiles = @()

# WPF is the default backend on Windows. Off Windows it can't load, so default to
# the headless 'None' backend unless the caller explicitly opts into a UI (e.g.
# the cross-platform Avalonia backend) via IM_UI_BACKEND.
$script:UIBackend = if ($env:IM_UI_BACKEND) { $env:IM_UI_BACKEND } elseif ($script:IsWindowsOS) { 'WPF' } else { 'None' }
# 'None' runs the module headless (CLI / automation): no UI files are scanned or
# loaded and no UI provider is instantiated. Useful on Linux/macOS where the WPF
# backend isn't available.
$script:AppUIRootFolder = if ($script:UIBackend -ieq 'None') {
    $null
} elseif ($script:UIBackend -ieq 'Avalonia') {
    "$script:AppRootFolder/UI/Avalonia"
} else {
    "$script:AppRootFolder/UI/WPF"
}
$script:AppUIFolderExists = $script:AppUIRootFolder -and [IO.Directory]::Exists($script:AppUIRootFolder)

$preloadFiles += "$script:AppRootFolder/Internal/Core.ps1"

if ($script:AppUIFolderExists) {
    # Glob so provider-suffixed names match (CoreUIWPF.ps1 / CoreUIAvalonia.ps1)
    # alongside the legacy CoreUI.ps1 during the rename window.
    $preloadFiles += @(Get-ChildItem -Path "$script:AppUIRootFolder/Extensions" `
        -Filter 'CoreUI*.ps1' -ErrorAction SilentlyContinue |
        ForEach-Object { $_.FullName })
}

foreach ($preloadFile in $preloadFiles) {
    $file = [IO.FileInfo]$preloadFile
    if ($file.Exists) {
        Write-Host "Import file $($file.Name)"
        if ($script:IsWindowsOS) { Unblock-File -Path $file.FullName }
        if ($useDotSoure) {
            . $file.FullName 
        }
        else {
            $ExecutionContext.InvokeCommand.InvokeScript($false, ([scriptblock]::Create([IO.File]::ReadAllText($File.FullName))), $null, $null)
        }
        $loadedFiles += $File.FullName
    }
    else {
        if ($file.DirectoryName.StartsWith($Script:AppUIRootFolder, "CurrentCultureIgnoreCase")) {
            Write-Warning "UI file $($file.Name) not found. Ignoring"
        }
        else {
            # A missing core preload (Core.ps1) is fatal: everything after this
            # loop calls functions defined there. Throw so module import fails
            # with one clear error instead of a cascade of "function not found".
            throw "Preload file '$($file.FullName)' does not exist"
        }
    }
}

# Everything from here on (Write-Log, Set-CacheObject, Get-CacheObject, ...) is
# defined in Internal/Core.ps1, which the preload loop above is responsible for
# loading first. Make that dependency explicit: if the preload didn't deliver
# the primitives (Core renamed/reordered out of the preload set, a function
# removed), fail with one clear error instead of a cascade of "command not
# found" through the rest of module init.
foreach ($coreFn in @('Write-Log', 'Set-CacheObject', 'Get-CacheObject', 'Get-SettingValue')) {
    if (-not (Get-Command $coreFn -ErrorAction SilentlyContinue)) {
        throw "Core preload did not define '$coreFn'. Internal/Core.ps1 must be preloaded (and define its primitives) before module initialization continues."
    }
}

Write-Log "#####################################################################################"
Write-Log "Application started"
Write-Log "#####################################################################################"

Write-Log "PowerShell version: $($PSVersionTable.PSVersion.ToString())"
if($PSVersionTable.BuildVersion) {
    Write-Log "PowerShell build: $($PSVersionTable.BuildVersion.ToString())"
}
if($PSVersionTable.CLRVersion) {
    Write-Log "PowerShell CLR: $($PSVersionTable.CLRVersion.ToString())"
}
Write-Log "PowerShell edition: $($PSVersionTable.PSEdition)"

# The detailed OS name/build live in the Windows registry; other platforms fall
# through to the OSVersion string.
if ($script:IsWindowsOS) {
    try {
        $osName = Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "ProductName" -ErrorAction Stop
        $patchLevel = Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "UBR" -ErrorAction Stop
        $ver = [Version]::new([Environment]::OSVersion.Version.Major, [Environment]::OSVersion.Version.Minor, [Environment]::OSVersion.Version.Build, $patchLevel)
        Write-Log "OS: $osName $ver"
    }
    catch {
        Write-Log "OS version: $([environment]::OSVersion.VersionString)"
    }
}
else {
    Write-Log "OS version: $([environment]::OSVersion.VersionString)"
}
$script:WPFNS = "xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'"

Set-CacheObject "ScriptRoot" $script:AppRootFolder -Persistent
# Registry | Json | Memory. Resolved at the top of Internal/Core.ps1 (it has to be
# known before the first Write-Log); this only publishes it to the cache for
# callers that report the active backend.
Set-CacheObject "SettingType" $script:SettingsStoreMode -Persistent
Set-CacheObject "MainAppStarted" $false
Set-CacheObject "MainUIStarted" $false
Set-CacheObject "ShowUI" $false -Persistent

$allClasses = @()

$allClasses += (Get-ChildItem -path "$script:AppRootFolder/Classes" -Filter "*.ps1")
# Shared UI base classes (e.g. UIProvider) live under UI\Classes and are
# loaded ahead of the per-provider subclasses so #ImportOrder can place the
# base before its subclass.
$allClasses += (Get-ChildItem -path "$script:AppRootFolder/UI/Classes" -Filter "*.ps1" -ErrorAction SilentlyContinue)
if ($script:AppUIFolderExists) {
    $allClasses += (Get-ChildItem -path "$script:AppUIRootFolder/Classes" -Filter "*.ps1" -ErrorAction SilentlyContinue)
}

foreach ($file in $allClasses) {
    $fileContent = Get-Content -LiteralPath $file.FullName -Encoding UTF8

    $tmpPriority = ($fileContent -like "$($ImportOrderStr)*")
    $ImportOrder = 0

    if ($tmpPriority) {
        try {
            $ImportOrder = [int]$tmpPriority.Substring($ImportOrderStr.Length).Trim()
        }
        catch {
            $ImportOrder = 0
        }
    }
    Remove-Variable "tmpPriority"

    if ($ImportOrder -eq 0) {
        if ($fileContent -like "Class * :*Attribute*") {
            $ImportOrder = 10
        }
        elseif ($fileContent -like "class *" -and $fileContent -notlike "class * : *") {
            $ImportOrder = 20
        }
        elseif ($fileContent -like "class *") {
            $ImportOrder = 50
        }
        else {
            $ImportOrder = 100
        }
    }
    
    $dependencyFiles += [PSCustomObject]@{
        FileInfo = $file
        Order    = $ImportOrder
    }
}

# Primary sort: #ImportOrder (explicit marker, else the derived default -
# 10 attribute class / 20 base class / 50 derived class / 100 non-class).
# Secondary sort: FullName, so files sharing the SAME Order load in a stable,
# deterministic order instead of relying on filesystem enumeration order.
foreach ($dependencyFile in ($dependencyFiles | Sort-Object -Property Order, @{ Expression = { $_.FileInfo.FullName } })) {
    Write-Host "Import dependency $($dependencyFile.FileInfo.Name) based on order $($dependencyFile.Order)"
    if ($script:IsWindowsOS) { Unblock-File -Path $dependencyFile.FileInfo.Fullname }
    . $dependencyFile.FileInfo.Fullname

    $loadedFiles += $dependencyFile.FileInfo.Fullname
}

$allPSFiles = @()
$allPSFiles += Get-ChildItem -path "$script:AppRootFolder/Internal" -Filter "*.ps1" -ErrorAction SilentlyContinue
if ($script:AppUIFolderExists) {
    # Extensions BEFORE ClassExtensions: ClassExtensions captures helper
    # functions (Add-AvaloniaExportPropertyCheckbox, Add-AvaloniaDetailsButton, ...)
    # from Extensions/IntuneManagerExtensionHooks.ps1 via `${function:X}` at
    # module-load time. If Extensions loads second, the captures get $null
    # and ScriptMethod invocations fail with "expression after & not valid"
    # (Avalonia Export form per-policy-type hooks fail this way).
    $allPSFiles += Get-ChildItem -path "$script:AppUIRootFolder/Extensions" -Filter "*.ps1" -ErrorAction SilentlyContinue
    $allPSFiles += Get-ChildItem -path "$script:AppUIRootFolder/ClassExtensions" -Filter "*.ps1" -ErrorAction SilentlyContinue
}

$allPSFiles += Get-ChildItem -path "$script:AppRootFolder/Public" -Filter "*.ps1" -ErrorAction SilentlyContinue

foreach ($file in $allPSFiles) {
    if ($loadedFiles -contains $file.FullName) { Write-Verbose "File $($file.Name) already loaded"; continue; }

    Write-Host "Import file $($file.Name)"
    if ($script:IsWindowsOS) { Unblock-File -Path $file.FullName }
    if ($useDotSoure) {
        . $file.FullName 
    }
    else {
        $ExecutionContext.InvokeCommand.InvokeScript($false, ([scriptblock]::Create([IO.File]::ReadAllText($File.FileName))), $null, $null)
    }
    $loadedFiles += $File.FullName
}

# GetFolderPath(LocalApplicationData) is cross-platform: %LOCALAPPDATA% on Windows,
# $XDG_DATA_HOME / ~/.local/share on Linux/macOS. The %LOCALAPPDATA% env var only
# exists on Windows, so ExpandEnvironmentVariables left it literal elsewhere.
$script:AppDataFolder = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) "IntuneManagement"

if ([IO.Directory]::Exists($script:AppDataFolder) -eq $false) {
    [IO.Directory]::CreateDirectory($script:AppDataFolder) | Out-Null
}

Remove-Variable "loadedFiles"
Remove-Variable "useDotSoure"

# Concrete UI provider for the active backend. Subclasses are loaded by the
# Classes scan above (UI\Classes\UIProvider.ps1 + provider folder's
# *UIProvider.ps1). $script:UIProvider stays $null when the UI folder isn't
# present (headless / CLI use) so non-UI callers can guard with `if ($script:UIProvider)`.
$script:UIProvider = $null
if ($script:AppUIFolderExists) {
    try {
        if ($script:UIBackend -ieq 'Avalonia') {
            $script:UIProvider = [AvaloniaUIProvider]::new()
        }
        else {
            $script:UIProvider = [WPFUIProvider]::new()
        }
    }
    catch {
        Write-LogError "Failed to instantiate UI provider for backend '$script:UIBackend'" $_.Exception
    }
}

# Module-scope alias of the provider.
#
# Every Show-* function opens with `$ui = $script:UIProvider` and then uses
# `$ui.X()` - including inside event handlers. On the Avalonia backend those
# handlers are re-bound to the module by ConvertTo-AvaloniaEventScriptBlock,
# which DROPS function-local captures, so `$ui` was $null at click time and
# the handler threw before doing anything (this silently killed Save buttons,
# Browse buttons, Close/Escape and whole dialogs).
#
# An unqualified `$ui` inside a module-bound scriptblock falls back to module
# scope, so defining it here makes that entrenched idiom safe everywhere. It
# is the SAME object as $script:UIProvider - never assign anything else to a
# local named $ui.
$script:ui = $script:UIProvider

# Load the settings backend before the AppInitialized event so event handlers
# read persisted values. On non-Windows this also ensures the JSON settings store
# is active before any Get-SettingStoreValue call, so the (Windows-only) registry branch is
# never reached. Registry mode is initialized lazily inside Start-CoreApp, and
# Memory mode needs no initialization - Core.ps1 created its settings object at
# dot-source time.
if ($script:SettingsStoreMode -eq "Json") { Initialize-JsonSettings }

Invoke-AppEvent "AppInitialized"

Start-CoreApp

Invoke-AppEvent "AppStarted"
