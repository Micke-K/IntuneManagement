#ImportOrder 5

class LogViewObject : ViewObjectBase
{
    LogViewObject() : base()
    {
        $this.Init()
    }
    
    Hidden Init()
    {
        $this._ID = "CoreLog"
        $this._Title = "Log"
        $this._Description = "View log items"
        $this._HideMenu = $true
    }
    
    [Object]GetViewPanel()
    {
        # ToDo: Remove recreate panel each time 
        #if($null -eq $this._ViewPanel) {
            $this._ViewPanel = Get-LogViewPanel
        #}
        return $this._ViewPanel
    }

    [Object[]]GetViewItems()
    {
        return $null
    }

    Authenticate()
    {

    }    

    [Object[]]OnItemChanged($SelectedItem)
    {
        return $null
    }

    OnDeactivating($NewActiveView)
    {

    }

    OnActivating($PreviousActiveView)
    {

    }

    OnActivated()
    {

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
        $this._ID = "CoreCachedObjects"
        $this._Title = "Cached Objects"
        $this._Description = "View cached items"
        $this._HideMenu = $true
    }
    
    [Object]GetViewPanel()
    {
        # ToDo: Remove recreate panel each time 
        #if($null -eq $this._ViewPanel) {
            $this._ViewPanel = Get-CachedObjectsViewPanel
        #}
        return $this._ViewPanel
    }

    [Object[]]GetViewItems()
    {
        return $null
    }

    Authenticate()
    {

    }    

    [Object[]]OnItemChanged($SelectedItem)
    {
        return $null
    }

    OnDeactivating($NewActiveView)
    {

    }

    OnActivating($PreviousActiveView)
    {

    }

    OnActivated()
    {
        if($this._ViewPanel) {
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
        $this._ID = "GraphCalls"
        $this._Title = "Graph Calls"
        $this._Description = "View all calls to Microsoft Graph"
        $this._HideMenu = $true
    }
    
    [Object]GetViewPanel()
    {
        # ToDo: Remove recreate panel each time 
        #if($null -eq $this._ViewPanel) {
            $this._ViewPanel = Get-GraphCallsViewPanel
        #}
        return $this._ViewPanel
    }

    [Object[]]GetViewItems()
    {
        return $null
    }

    Authenticate()
    {

    }    

    [Object[]]OnItemChanged($SelectedItem)
    {
        return $null
    }

    OnDeactivating($NewActiveView)
    {

    }

    OnActivating($PreviousActiveView)
    {

    }

    OnActivated()
    {

    }    
}