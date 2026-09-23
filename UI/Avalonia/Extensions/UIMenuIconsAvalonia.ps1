# Avalonia counterpart of UI/WPF/Extensions/UIMenuIconsWPF.ps1.
#
# ClassExtensions/IntuneCommonUIExtensions.ps1 attaches LoadIconImage/GetImage
# to every policy-type and group singleton at import time (the closures must
# capture the module-bound loader then) but no longer parses the icons there.
# Get-IntuneViewItems calls this right before it projects the singletons into
# ViewMenuItem rows, so the visuals load once the splash is up and the menu
# needs them. Each item keeps its own instance: a visual can only have one parent.

function Initialize-MenuItemIcons
{
    param($Items)

    foreach($item in @($Items))
    {
        if($null -eq $item) { continue }
        if($null -ne $item.IconImage) { continue }
        if(-not $item.PSObject.Methods['LoadIconImage']) { continue }
        try { $item.LoadIconImage() } catch { }
    }
}
