#ImportOrder 5

# Avalonia counterparts of the three debug ViewObjects in
# UI/WPF/Classes/CoreUIBaseClassesWPF.ps1 (LogViewObject, CachedObjectsViewObject,
# GraphCallsViewObject). Get-SubClasses("ViewObjectBase") in Show-MainWindow
# auto-discovers these and Add-ViewObject registers each one — same flow as
# the existing PreviewViewObject sandbox view.
#
# All three set _HideMenu=$true: per Show-View, that hides the in-view item
# menu (these panels host their own DataGrid rather than a row of view items).
# AddToMenu stays at the default $true so they appear in the left nav under
# the Views section.

class LogViewObject : ViewObjectBase
{
    LogViewObject() : base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._ID          = "CoreLog"
        $this._Title       = "Log"
        $this._Description = "View log items"
        $this._HideMenu    = $true
    }

    [Object] GetViewPanel()
    {
        if ($null -eq $this._ViewPanel) {
            $this._ViewPanel = Get-LogViewPanel
        }
        return $this._ViewPanel
    }

    [Object[]] GetViewItems() { return $null }

    Authenticate() {}

    [Object[]] OnItemChanged($SelectedItem) { return $null }

    OnDeactivating($NewActiveView) {}

    OnActivating($PreviousActiveView) {}

    OnActivated()
    {
        # Re-bind log items each activation so entries logged while the view
        # was inactive show up. ItemsSource is reassigned (not appended) so
        # the DataGrid picks up the latest snapshot.
        if ($this._ViewPanel) {
            Update-LogView -ViewPanel $this._ViewPanel
        }
    }
}

class CachedObjectsViewObject : ViewObjectBase
{
    CachedObjectsViewObject() : base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._ID          = "CoreCachedObjects"
        $this._Title       = "Cached Objects"
        $this._Description = "View cached items"
        $this._HideMenu    = $true
    }

    [Object] GetViewPanel()
    {
        if ($null -eq $this._ViewPanel) {
            $this._ViewPanel = Get-CachedObjectsViewPanel
        }
        return $this._ViewPanel
    }

    [Object[]] GetViewItems() { return $null }

    Authenticate() {}

    [Object[]] OnItemChanged($SelectedItem) { return $null }

    OnDeactivating($NewActiveView) {}

    OnActivating($PreviousActiveView) {}

    OnActivated()
    {
        if ($this._ViewPanel) {
            Update-CachedObjectsView -ViewPanel $this._ViewPanel
        }
    }
}

class GraphCallsViewObject : ViewObjectBase
{
    GraphCallsViewObject() : base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._ID          = "GraphCalls"
        $this._Title       = "Graph Calls"
        $this._Description = "View all calls to Microsoft Graph"
        $this._HideMenu    = $true
    }

    [Object] GetViewPanel()
    {
        if ($null -eq $this._ViewPanel) {
            $this._ViewPanel = Get-GraphCallsViewPanel
        }
        return $this._ViewPanel
    }

    [Object[]] GetViewItems() { return $null }

    Authenticate() {}

    [Object[]] OnItemChanged($SelectedItem) { return $null }

    OnDeactivating($NewActiveView) {}

    OnActivating($PreviousActiveView) {}

    OnActivated()
    {
        if ($this._ViewPanel) {
            Update-GraphCallsView -ViewPanel $this._ViewPanel
        }
    }
}
