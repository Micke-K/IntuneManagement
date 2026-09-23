#ImportOrder 1
class ViewObjectBase
{
    Hidden [String]$_ID = $null
    Hidden [String]$_Title = $null
    Hidden [String]$_Description = $null
    Hidden [String]$_AuthenticaionObject = $null
    Hidden [Boolean]$_HideMenu = $false
    Hidden [Object]$_ViewPanel = $null
    Hidden [Guid]$_SessionId = [Guid]::Empty
    Hidden [Boolean]$_AddToMenu = $true
    # When true, the top-bar Views menu surfaces this view's items as a
    # submenu beneath the view entry — clicking a child both activates the
    # parent view and selects that item in the left nav. Default off because
    # most views (IntuneManagement with ~50 policy types) would balloon the
    # Views menu past usability.
    Hidden [Boolean]$_ExpandInViewsMenu = $false
    
    ViewObjectBase()
    {
        if($this.GetType().Name -eq "ViewObjectBase") {
            throw "Abstract class. Object cannot be created"
        }
        elseif($script:SingletonObjects.ContainsKey($this.GetType().Name) -eq $true) {
            throw "Only one $($this.GetType().Name) object can be created"
        }

        Add-SingletonObject $this.GetType().Name $this

        ([ViewObjectBase]$this).Init()        
    }
    
    # Hidden Functions
    Hidden Init()
    {
        $this._SessionId = $script:SessionID
        Add-ObjectProperty $this "ID" { $this._ID }
        Add-ObjectProperty $this "Title" { $this._Title }
        Add-ObjectProperty $this "Description" { $this._Description }
        Add-ObjectProperty $this "Authentication" { $this._AuthenticaionObject }
        Add-ObjectProperty $this "HideMenu" { $this._HideMenu }
        Add-ObjectProperty $this "ViewPanel" { $this.GetViewPanel()  }
        Add-ObjectProperty $this "AddToMenu" { $this._AddToMenu }
        Add-ObjectProperty $this "ExpandInViewsMenu" { $this._ExpandInViewsMenu }
    }

    [Object[]]GetViewItems()
    {
        return $null
    }

    [Object]GetViewPanel()
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

class ViewItemBase
{
    [String]$Id = ""
    [String]$Name = ""
    [String]$Icon = $null

    ViewItemBase()
    {
    }
}
