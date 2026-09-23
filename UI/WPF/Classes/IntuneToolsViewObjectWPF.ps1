#ImportOrder 160
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
        # Tools list is short and curated — fine to also surface in the top
        # Views menu as a submenu (one entry per tool, grouped by Category).
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

    [PSCustomObject[]]GetViewItems()
    {
        return @(Get-IntuneToolsViewItems)
    }

    Authenticate()
    {
        # Auto-login policy: silent-only. If the silent path fails the user signs in
        # explicitly via the Login button — we never pop an unbidden OAuth window
        # from a view-activation event.
        #
        # Also gated to MSAL provider so MgGraph mode isn't hijacked by a successful
        # MSAL silent auth here.
        $active = $null
        try { $active = Get-AuthProvider } catch { }
        if($active) { $active.RefreshAmbientSession() }
    }

    [Object[]]OnItemChanged($SelectedItem)
    {
        return (Invoke-IntuneToolsActivateItem $SelectedItem)
    }
}
