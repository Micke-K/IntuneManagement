# Theme resolution shared by both UI backends.
#
# The AppTheme setting has three values:
#   Default - follow the OS app theme, and keep following it
#   Light   - always light
#   Dark    - always dark
#
# "Default" is not a theme of its own: it resolves to Light or Dark at the point
# of use. Both backends call Resolve-AppTheme and then apply the concrete result
# their own way - WPF merges Themes/<name>.xaml, Avalonia sets
# Application.RequestedThemeVariant - so this file stays free of toolkit types
# (R12) and neither backend reimplements the lookup (R10).

# Platform flags, mirroring $script:IsWindowsOS in IntuneManagement.psm1.
# $IsMacOS / $IsLinux exist only on PowerShell Core; on 5.1 they read as $null,
# which is correct because 5.1 is Windows-only. Captured once at load so tests
# can override them in module scope to exercise a branch off-platform.
$script:IsMacOSPlatform = [bool]$IsMacOS
$script:IsLinuxPlatform = [bool]$IsLinux

# The two external theme readers sit behind thin wrappers on purpose: Pester
# cannot mock a native command that does not exist on the machine running the
# test, so `defaults` is unmockable on Linux/Windows CI and `gsettings` on
# Windows/macOS. Wrapping them makes both branches of Get-SystemAppTheme
# testable everywhere. They are the only lines here that touch the OS.
function Get-MacOSInterfaceStyle {
    [CmdletBinding()]
    param()
    & defaults read -g AppleInterfaceStyle 2>$null
}

function Get-LinuxDesktopSetting {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Key)
    & gsettings get org.gnome.desktop.interface $Key 2>$null
}

# The OS app theme, as 'Light' or 'Dark'.
#
# Every platform falls back to Light when its theme cannot be read. Light is the
# historical default for this app, so an unreadable setting changes nothing.
function Get-SystemAppTheme {
    [CmdletBinding()]
    param()

    if ($script:IsWindowsOS) {
        # Personalize\AppsUseLightTheme is the *app* theme; SystemUsesLightTheme
        # is the separate taskbar/Start setting and is deliberately not used - a
        # user who runs light apps on a dark taskbar wants light apps.
        try {
            $key = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
            $value = (Get-ItemProperty -Path $key -Name 'AppsUseLightTheme' -ErrorAction Stop).AppsUseLightTheme
            if ($null -ne $value -and [int]$value -eq 0) { return 'Dark' }
            return 'Light'
        }
        catch {
            Write-LogDebug "Windows app theme not readable, assuming Light: $($_.Exception.Message)"
            return 'Light'
        }
    }

    if ($script:IsMacOSPlatform) {
        # 'defaults read -g AppleInterfaceStyle' prints 'Dark' in dark mode. In
        # light mode the key does not exist at all, so the command exits non-zero
        # and prints "does not exist" to stderr. That failure IS the light answer,
        # not an error - hence the catch returning Light rather than logging loudly.
        # The try/catch is required: on PS 7.4+ a non-zero native exit throws when
        # $ErrorActionPreference is Stop.
        try {
            $style = Get-MacOSInterfaceStyle
            if ($style -and (($style -join '').Trim() -eq 'Dark')) { return 'Dark' }
        }
        catch {
            Write-LogDebug "macOS AppleInterfaceStyle not set, assuming Light"
        }
        return 'Light'
    }

    if ($script:IsLinuxPlatform) {
        # There is no cross-desktop standard. GNOME 42+ exposes color-scheme
        # ('prefer-dark' / 'prefer-light' / 'default'); 'default' means the user
        # expressed no preference, so it is Light here. Older GNOME and several
        # derivatives only carry gtk-theme, whose name ends in '-dark' by
        # convention (Adwaita-dark, Yaru-dark). Anything else stays Light.
        # gsettings values come back single-quoted, hence the Trim.
        try {
            $scheme = Get-LinuxDesktopSetting 'color-scheme'
            if ($scheme) {
                $scheme = ($scheme -join '').Trim().Trim("'")
                if ($scheme -eq 'prefer-dark') { return 'Dark' }
                if ($scheme -eq 'prefer-light') { return 'Light' }
            }

            $gtkTheme = Get-LinuxDesktopSetting 'gtk-theme'
            if ($gtkTheme) {
                $gtkTheme = ($gtkTheme -join '').Trim().Trim("'")
                if ($gtkTheme -match '-dark$') { return 'Dark' }
            }
        }
        catch {
            Write-LogDebug "Linux desktop theme not readable (no gsettings?), assuming Light"
        }
        return 'Light'
    }

    return 'Light'
}

# Turn an AppTheme setting value into the concrete theme to apply.
#
# Pass nothing to read the current setting. Unknown values resolve as Default
# rather than throwing, so a hand-edited settings file cannot leave the app
# unthemed.
function Resolve-AppTheme {
    [CmdletBinding()]
    param([string]$ThemeName)

    if (-not $PSBoundParameters.ContainsKey('ThemeName') -or [string]::IsNullOrWhiteSpace($ThemeName)) {
        $ThemeName = Get-SettingValue 'AppTheme'
    }

    switch ($ThemeName) {
        'Light' { return 'Light' }
        'Dark'  { return 'Dark' }
        default { return (Get-SystemAppTheme) }
    }
}

# True when the setting says "follow the OS" - the backends use this to decide
# whether an OS theme change should be reacted to.
function Test-AppThemeFollowsSystem {
    [CmdletBinding()]
    param([string]$ThemeName)

    if (-not $PSBoundParameters.ContainsKey('ThemeName') -or [string]::IsNullOrWhiteSpace($ThemeName)) {
        $ThemeName = Get-SettingValue 'AppTheme'
    }
    return ($ThemeName -notin @('Light', 'Dark'))
}
