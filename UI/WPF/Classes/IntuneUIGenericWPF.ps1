#ImportOrder 150
class IntuneViewObject : ViewObjectBase
{
    Hidden [IntunePolicyBase[]]$_Items = @()

    IntuneViewObject() : Base()
    {
        $this.Init()
    }
    
    Hidden Init()
    {
        $this._ID = "IntuneManagement"
        $this._Title = "Intune Manager"        
        $this._Description = "Manages Intune environments. This view can be used for copying objects in an Intune environment. It can also be used for backing up an entire Intune environment and cloning the Intune environment into another tenant."
        $this._AddToMenu = $false
    }
    
    [Object]GetViewPanel()
    {
        if($null -eq $this._ViewPanel) {
            $this._ViewPanel = Get-IntuneManagementViewPanel
        }
        return $this._ViewPanel
    }

    [PSCustomObject[]]GetViewItems()
    {
        return @(Get-IntuneViewItems)
    }
    
    Authenticate()
    {
        # Auto-login policy: at startup / view activation we only attempt SILENT auth
        # against the last-logged-on account. If it fails we surface "signed out" and
        # wait for the user to click the Login button. We do NOT pop an interactive
        # OAuth prompt unbidden — that would steal focus from whatever the user is
        # doing and contradict the explicit-login design.
        #
        # Also gated to MSAL provider so MgGraph mode isn't hijacked (a successful
        # silent MSAL auth here would flip the active provider).
        $active = $null
        try { $active = Get-AuthProvider } catch { }
        if($active) { $active.RefreshAmbientSession() }
    }

    [Object[]]OnItemChanged($SelectedItem)
    {
        return (Invoke-IntuneActivateObject $SelectedItem)
    }

    OnDeactivating($NewActiveView)
    {
        Invoke-IntuneDeactivatingView $NewActiveView
    }

    OnActivating($PreviousActiveView)
    {
        Invoke-IntuneActivatingView $PreviousActiveView
    }
}
