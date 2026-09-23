<#
    Platform-specific main-window chrome.

    MainWindow.axaml is authored Windows-first: ExtendClientAreaToDecorationsHint
    plus PreferSystemChrome merge our title bar into the OS caption area, and the
    title-bar content grid reserves 140px on the right for the Win11
    minimize/maximize/close buttons that the OS paints over our content.

    That hint is only honoured by the Win32 and macOS-native backends. Avalonia's
    X11 backend does not remove the window-manager decoration, so on Linux the WM
    keeps drawing its own title bar and the app draws a second one underneath it,
    complete with 140px of dead space reserved for buttons that are not there.

    macOS does honour the hint, but puts its traffic lights on the LEFT, so the
    same reservation has to move to the other side or the app icon and File menu
    end up underneath them.

    Windows keeps the authored layout untouched.
#>

# Width reserved inside the extended client area for OS-painted caption buttons.
$script:IMWindowsCaptionWidth = 140
$script:IMMacTrafficLightWidth = 80

function Set-MainWindowChrome {
    <#
    .SYNOPSIS
    Adapts the main window chrome to the running platform.

    .DESCRIPTION
    Called from Show-MainWindow after MainWindow.axaml is loaded and before the
    window is shown. On Windows this is a no-op: the XAML is already correct.
    #>
    param($Window)

    if (-not $Window) { return }

    if ($script:IsWindowsOS) { return }

    if ($IsMacOS) {
        # The hint works here, so we keep the merged title bar and only move the
        # caption reservation from the right (Windows) to the left (traffic lights).
        Set-XamlProperty $Window 'grdTitleBarContent' 'Margin' `
            ([Avalonia.Thickness]::new($script:IMMacTrafficLightWidth, 0, 0, 0))
        return
    }

    # Linux/X11, and any other backend that ignores the hint: hand the title bar
    # back to the window manager so we do not end up with two of them.
    try {
        $Window.ExtendClientAreaToDecorationsHint = $false
    } catch {
        Write-LogDebug "Set-MainWindowChrome: could not disable extended client area: $($_.Exception.Message)"
    }

    # Nothing overlaps our content now, so drop the caption reservation and hide
    # the duplicate centred view title. The WM caption already shows Window.Title,
    # which Set-MainTitle keeps in sync with the same text.
    Set-XamlProperty $Window 'grdTitleBarContent' 'Margin' ([Avalonia.Thickness]::new(0))
    Set-XamlProperty $Window 'txtTitleViewName' 'IsVisible' $false
}
