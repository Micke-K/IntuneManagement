# Experimental-platform notice - the DECISION half.
#
# macOS and Linux run the same engine as Windows through the Avalonia backend, but
# the macOS path has never been executed on a Mac. Users get told that once, with a
# "do not show again" opt-out, rather than discovering it from a stack trace.
#
# Why here and not in UI/<backend>/: UI/Classes/ loads for BOTH backends, so the
# gate is testable headlessly on Windows (Tests/ExperimentalPlatformNotice.Tests.ps1)
# even though the dialog itself only ever renders under Avalonia. The rendering half
# lives in UI/Avalonia/Extensions/ExperimentalPlatformNoticeAvalonia.ps1.
#
# Why AppInitialized: UI/Classes/ dot-sources before Internal/, so the settings
# sections do not exist at load time. See the header of UICommonSettings.ps1.

$script:ExperimentalPlatformNoticeKey = "HideExperimentalPlatformNotice"

function Add-ExperimentalPlatformSettings
{
    # Idempotent: the notice calls this defensively if it runs before AppInitialized
    # has fired, and a second Add-SettingsObject would put a duplicate row in the
    # Settings dialog.
    foreach($section in (Get-SettingsSections)) {
        if($section.Values | Where-Object Key -eq $script:ExperimentalPlatformNoticeKey) { return }
    }

    Add-SettingsObject -Title "Hide experimental platform notice" -Key $script:ExperimentalPlatformNoticeKey -Type "Boolean" `
        -Description "Stop showing the startup notice that macOS and Linux support is experimental. Clear this to see it again." `
        -DefaultValue $false -Section "General"
}
Add-AppEventHandler "AppInitialized" "Add-ExperimentalPlatformSettings"

function Test-ShouldShowExperimentalPlatformNotice
{
    # The opt-out wins everywhere, including under the forced override - otherwise a
    # developer who ticks the box on Windows cannot get the dialog back without
    # editing the settings store by hand.
    if((Get-SettingValue $script:ExperimentalPlatformNoticeKey) -eq $true) { return $false }

    # IM_EXPERIMENTAL_NOTICE=1 forces the dialog on Windows so it can be reviewed on
    # the machine it is developed on; it never renders there otherwise.
    if($script:IsWindowsOS -and $env:IM_EXPERIMENTAL_NOTICE -ne '1') { return $false }

    return $true
}
