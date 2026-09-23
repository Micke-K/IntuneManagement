#ImportOrder 7

# Verbatim port of UI/WPF/Classes/IntuneToolsViewObject.ps1. The class shape is
# unchanged; the difference between WPF and Avalonia lives in the matching
# Extensions/IntuneToolsUI.ps1 (panel + items + tool placeholders).

class IntuneToolsViewObject : ViewObjectBase
{
    IntuneToolsViewObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._ID                  = "IntuneTools"
        $this._Title               = "Intune Tools"
        $this._Description         = "Tools that operate on Intune objects, such as listing assignments across all object types."
        $this._AddToMenu           = $true
        $this._ExpandInViewsMenu   = $true
    }

    [Object]GetViewPanel()
    {
        if($null -eq $this._ViewPanel)
        {
            $this._ViewPanel = Get-IntuneToolsViewPanel
        }
        return $this._ViewPanel
    }

    [Object[]]GetViewItems()
    {
        return @(Get-IntuneToolsViewItems)
    }

    Authenticate()
    {
        $active = $null
        try { $active = Get-AuthProvider } catch { }
        if($active) { $active.RefreshAmbientSession() }
    }

    [Object[]]OnItemChanged($SelectedItem)
    {
        return (Invoke-IntuneToolsActivateItem $SelectedItem)
    }
}
