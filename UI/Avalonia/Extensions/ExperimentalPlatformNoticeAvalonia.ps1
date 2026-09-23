# Experimental-platform notice - the RENDERING half. The decision lives in
# UI/Classes/ExperimentalPlatformNotice.ps1 (loads under both backends, so the gate
# is unit-testable on Windows); this file only draws it.
#
# Avalonia only, by design: the notice never shows on Windows, and WPF is the only
# backend there - so there is no WPF twin to keep in sync.
#
# Closure law: click handlers run through ConvertTo-AvaloniaEventScriptBlock, which
# strips function-local captures but keeps module scope. The handler below therefore
# reads the checkbox off $S.Tag and calls module functions by name only - never a
# local, never $script:UIProvider.X(). See UI/Avalonia/CLAUDE.md.

function Get-ExperimentalPlatformName
{
    if($IsMacOS)   { return "macOS" }
    if($IsLinux)   { return "Linux" }
    return "this platform"
}

function Show-ExperimentalPlatformNotice
{
    if(-not (Test-ShouldShowExperimentalPlatformNotice)) { return }

    $ui = $script:UIProvider
    $panel = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/ExperimentalPlatformNotice.axaml'))
    if(-not $panel) { return }

    $platform = Get-ExperimentalPlatformName
    $body = "IntuneManagement runs on $platform through the Avalonia UI backend, " +
            "but this build has not been verified there. Policy export, import, " +
            "compare and documentation all use the same engine as the Windows " +
            "build, so they are expected to work - the parts most likely to " +
            "misbehave are the window itself, file dialogs and the sign-in " +
            "browser flow." + [Environment]::NewLine + [Environment]::NewLine +
            "Word documentation output and MSI property extraction need Windows " +
            "and are unavailable here." + [Environment]::NewLine + [Environment]::NewLine +
            "Please report anything that breaks."

    $ui.SetXamlProperty($panel, 'txtNoticeBody', 'Text', $body)

    # Carry the checkbox on the button's Tag: the handler cannot capture it as a
    # function local (see the closure note above).
    $hostType = Get-AvaloniaHost
    $okButton = $hostType::FindByName($panel, 'btnNoticeOk')
    $checkBox = $hostType::FindByName($panel, 'chkHideNoticeAgain')
    if($okButton) { $okButton.Tag = $checkBox }

    $ui.AddXamlEvent($panel, 'btnNoticeOk', 'Add_Click', ({
        param($S, $E)
        try {
            $chk = $S.Tag
            if($chk -and $chk.IsChecked -eq $true) {
                Save-SettingStoreValue "" $script:ExperimentalPlatformNoticeKey $true
            }
        }
        catch { Write-LogError 'Failed to save the experimental platform notice preference' $_.Exception }

        Close-TopModalObject
    }))

    Show-ModalForm -FormTitle "Experimental platform" -FormObject $panel -HideButtons
}
