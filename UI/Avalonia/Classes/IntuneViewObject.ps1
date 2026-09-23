#ImportOrder 150

# Avalonia port of UI/WPF/Classes/IntuneUIGeneric.ps1.
#
# The legacy "Intune Manager" view — main object editor / list / bulk host. In
# WPF this is the landing view (`Show-View "IntuneManagement"`); we mirror that
# here. AddToMenu=$false so it doesn't show up in the Views menu — it's added
# explicitly by Invoke-IntuneUIAppInitialized.
#
# Method bodies delegate to functions in Extensions/IntuneManagerUI.ps1
# (mirrors the WPF split). Most are still stubbed pending the data-load and
# DataGrid column ports — see CLAUDE.md "Pending work".

class IntuneViewObject : ViewObjectBase
{
    Hidden [Object[]]$_Items = @()

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
        if ($null -eq $this._ViewPanel) {
            $this._ViewPanel = Get-IntuneManagementViewPanel
        }
        return $this._ViewPanel
    }

    [Object[]]GetViewItems()
    {
        return @(Get-IntuneViewItems)
    }

    Authenticate()
    {
        # Match WPF: silent-only at activation. Interactive auth happens via the
        # Sign-in button. Gated to MSAL so MgGraph mode isn't hijacked.
        $active = $null
        try { $active = Get-AuthProvider } catch { }
        if ($active) { $active.RefreshAmbientSession() }
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
