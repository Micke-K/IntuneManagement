# Menu icons load when a menu is built, not when the module is imported.
#
# ClassExtensions/IntuneCommonUIExtensions.ps1 attaches LoadIconImage/GetImage
# to every policy-type and group singleton at import time (the closures must
# capture the module-bound XAML loader then). It used to call LoadIconImage
# right away too, which parsed 82 icon files through XamlReader before Start.ps1
# had even reached Show-MainWindow - the whole pass ran before the splash
# existed. Show-ViewMenu calls this instead, once the splash is up and the menu
# actually needs the visuals. Each item keeps its own instance: a WPF visual can
# only have one parent.

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
