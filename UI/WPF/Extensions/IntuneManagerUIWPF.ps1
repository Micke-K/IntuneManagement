Add-AppEvent "IntuneManagerUISelectedMenuItemChanged"
Add-AppEvent "IntuneManagerUISelectedItemChanged"

Add-AppEventHandler "AppInitialized" "Invoke-IntuneUIAppInitialized"
Add-AppEventHandler "SettingsUpdated" "Invoke-IntuneUIEventSettingsUpdated"

function Get-IntuneViewItems
{
    [OutputType([PSCustomObject[]])]
    param()

    # ToDo: Fix support for PolicyGroup OR PolicyType

    $viewType = Get-SettingValue "ObjectViewType" # "Type"
    if($viewType -eq "Type") {
        $viewItems = $script:IntuneTypes | Sort-Object -Property Title
    }
    else {
        $viewItems = $script:IntuneGroups | Sort-Object -Property Title
    }

    # Read-only "Intune Info" group (Baseline Templates, Tenant Settings, etc.) is
    # opt-in: hidden unless enabled via the gear menu's "Show read-only Info". Drop
    # the group (Group view) or its member types (API view).
    if((Get-SettingStoreValue "IntuneManager" "ShowIntuneInfo" "false") -ne "true") {
        if($viewType -eq "Type") {
            $viewItems = @($viewItems | Where-Object { -not ($_.PolicyGroup -and $_.PolicyGroup.Id -eq "IntuneInfo") })
        }
        else {
            $viewItems = @($viewItems | Where-Object { $_.Id -ne "IntuneInfo" })
        }
    }

    # Stamp AccessType / AccessInfo from the token's granted scopes so the
    # ListBox triggers in MainWindow.xaml can colour restricted rows (orange =
    # read-only, red = no access). Works for both view modes: groups aggregate
    # their member types. No-op when there is no token to diff against.
    Update-IntuneAccessLevels

    if((Get-SettingValue "HideNoAccess"))
    {
        # Drop rows the token cannot use at all. Groups aggregate to None only
        # when every member type is unusable, so a partially usable group is
        # never hidden - it stays visible and orange.
        $viewItems = @($viewItems | Where-Object { $_.AccessType -ne [APIAccess]::None })
    }

    # MenuLabel is the bound display string for each menu row (template uses
    # `Text="{Binding MenuLabel}"`). Default to Title; Update-IntuneViewItemCounts
    # later overrides with "Title (N)" when the gear's "Show item counts"
    # toggle is on. Always reset to Title here so a view-mode switch can't
    # leave a stale "(N)" attached when counts are off.
    $showCounts = (Get-SettingStoreValue "IntuneManager" "ShowMenuItemCounts" "false") -eq "true"
    foreach($item in $viewItems) {
        $item | Add-Member -NotePropertyName "MenuLabel" -NotePropertyValue $item.Title -Force
    }

    # Defer to Show-ViewMenu's post-bind hook to populate counts (it has access
    # to $script:lstMenuItems after the bind happens). $showCounts only here
    # so the variable is referenced — actual count population happens later.
    $script:_intuneShowItemCounts = $showCounts

    return ($viewItems)
}

# Batched type-count fetcher. Runs once per Refresh / first count-enable;
# cached results power both API view (1 count per row) and Group view (sum
# of child-type counts per group row). Skips silently if there's no signed-
# in token — counts will populate the next time the menu rebuilds.
function Update-IntuneViewItemCounts
{
    param([switch]$Force)

    if(-not $script:lstMenuItems -or -not $script:lstMenuItems.ItemsSource) { return }

    $tokenId = $null
    try { $tokenId = Get-DefaultTokenId } catch { }
    if($null -eq $tokenId) {
        Write-LogDebug "Update-IntuneViewItemCounts: no token id available - skipping"
        return
    }

    # Fetch ALL type counts (not just the visible rows) so Group view can roll
    # up without a second pass. The fetch + session cache live in
    # Internal/IntuneManager.ps1 (Get-IntuneTypeCounts) and are shared with the
    # Avalonia backend; -Force invalidates.
    [void](Get-IntuneTypeCounts -Force:$Force -TokenId $tokenId)

    # Project counts onto the currently-bound MenuLabels. Type rows: direct
    # lookup. Group rows: sum of every child type's count (skips types with
    # no count so a single 400 doesn't void the whole group).
    foreach($it in $script:lstMenuItems.ItemsSource) {
        $count = $null
        if($it -is [IntunePolicyTypeBase]) {
            if($script:_intuneTypeCounts.ContainsKey($it.Id)) { $count = $script:_intuneTypeCounts[$it.Id] }
        }
        elseif($it -is [IntunePolicyGroupBase]) {
            $sum = 0; $hasAny = $false
            foreach($child in @($it.PolicyTypes)) {
                if($child -and $script:_intuneTypeCounts.ContainsKey($child.Id)) {
                    $sum += $script:_intuneTypeCounts[$child.Id]; $hasAny = $true
                }
            }
            if($hasAny) { $count = $sum }
        }
        $label = if($null -ne $count) { "$($it.Title) ($count)" } else { $it.Title }
        $it | Add-Member -NotePropertyName "MenuLabel" -NotePropertyValue $label -Force
    }
    try { $script:lstMenuItems.Items.Refresh() } catch { }
}

function Get-IntuneManagementViewPanel
{
    $panel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\IntuneManagerPanel.xaml"))

    $script:dgIntuneManagerObjects = $panel.FindName("dgIntuneManagerObjects")
    $script:spIntuneManagerButtons = $panel.FindName("spSubMenu")
    $script:IntuneManagementFilterTextBox = $panel.FindName("txtFilter")
    $script:IntuneManagementObjectsCount = $panel.FindName("txtObjectsCount")

    $script:UIProvider.AddXamlEvent($panel, "btnCopy", "Add_Click", {
        Copy-IntuneManagerPolicy
    })

    $script:UIProvider.AddXamlEvent($panel, "btnView", "Add_Click", {
        Show-IntuneManagerDetailedView
    })

    $script:UIProvider.AddXamlEvent($panel, "btnCompare", "Add_Click", {
        $selectedItems = @($script:dgIntuneManagerObjects.ItemsSource | Where-Object IsSelected -eq $true)
        if($selectedItems.Count -eq 0)
        {
            $sel = $script:dgIntuneManagerObjects.SelectedItem
            if($sel) { $selectedItems = @($sel) }
        }
        if($selectedItems.Count -gt 0) { Show-GraphCompareForm $selectedItems }
    })

    $script:UIProvider.AddXamlEvent($panel, "btnExport", "Add_Click", {
        Show-IntuneManagerExportForm
    })

    $script:UIProvider.AddXamlEvent($panel, "btnImport", "Add_Click", {
        Show-IntuneManagerImportForm
    })

    $script:UIProvider.AddXamlEvent($panel, "btnDelete", "Add_Click", {
        Remove-GraphObjectsUI
    })

    $script:UIProvider.AddXamlEvent($panel, "btnDocument", "Add_Click", {
        # Prefer IsSelected-checked rows; fall back to highlighted row; fall back
        # to all currently-visible rows (matches the OLD project's btnDocument
        # behaviour where clicking with nothing selected documented the view).
        $items = @($script:dgIntuneManagerObjects.ItemsSource | Where-Object IsSelected -eq $true)
        if ($items.Count -eq 0 -and $null -ne $script:dgIntuneManagerObjects.SelectedItem) {
            $items = @($script:dgIntuneManagerObjects.SelectedItem)
        }
        if ($items.Count -eq 0) {
            $items = @($script:dgIntuneManagerObjects.ItemsSource)
        }
        if ($items.Count -eq 0) { return }
        Show-IntuneManagerDocumentForm -PolicyObject $items
    })

    $script:UIProvider.AddXamlEvent($panel, "txtFilter", "Add_TextChanged", {
        Invoke-FilterBoxChanged $this $script:dgIntuneManagerObjects -ObjectsCountTextBox $script:IntuneManagementObjectsCount
    })

    $dpd = [System.ComponentModel.DependencyPropertyDescriptor]::FromProperty([System.Windows.Controls.ItemsControl]::ItemsSourceProperty, [System.Windows.Controls.DataGrid])
    if($dpd)
    {
        $dpd.AddValueChanged($script:dgIntuneManagerObjects, {
            $script:UIProvider.SetXamlProperty($this.Parent, "txtFilter", "Text", "")
            # Clearing txtFilter only fires TextChanged when the text actually
            # changes; for an already-empty filter that no-ops, leaving the
            # objects-count label stale (or blank on first load). Refresh
            # explicitly with -ForceUpdate so the count tracks every ItemsSource
            # change — initial load, view switch, refresh, all paths.
            Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox $script:dgIntuneManagerObjects -ForceUpdate -ObjectsCountTextBox $script:IntuneManagementObjectsCount
        })
    }

    $script:dgIntuneManagerObjects.add_selectionChanged({
        
        $hasSelectedItems = ($script:dgIntuneManagerObjects.ItemsSource | Measure-Object).Count -gt 0 -and (($script:dgIntuneManagerObjects.ItemsSource | Where-Object IsSelected -eq $true) -or ($null -ne $script:dgIntuneManagerObjects.SelectedItem))

        $script:UIProvider.SetXamlProperty($script:spIntuneManagerButtons, "btnView", "IsEnabled", $hasSelectedItems)
        $script:UIProvider.SetXamlProperty($script:spIntuneManagerButtons, "btnCompare", "IsEnabled", $hasSelectedItems)
        $script:UIProvider.SetXamlProperty($script:spIntuneManagerButtons, "btnCopy", "IsEnabled", $hasSelectedItems)
        $script:UIProvider.SetXamlProperty($script:spIntuneManagerButtons, "btnDelete", "IsEnabled", $hasSelectedItems)
        $script:UIProvider.SetXamlProperty($this.Parent, "btnExport", "IsEnabled", $hasSelectedItems)
        # btnDocument stays enabled whenever any items are loaded — clicking
        # with nothing selected documents the entire view (matches OLD project).
        $script:UIProvider.SetXamlProperty($script:spIntuneManagerButtons, "btnDocument", "IsEnabled", (($script:dgIntuneManagerObjects.ItemsSource | Measure-Object).Count -gt 0))

        Invoke-AppEvent "IntuneManagerUISelectedItemChanged" $hasSelectedItems
    })

    $btnRefresh = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\RefreshButton.xaml"))
    if($btnRefresh)
    {
        $grdTitle = $panel.FindName("grdTitle")
        $btnRefresh.SetValue([System.Windows.Controls.Grid]::ColumnProperty, $grdTitle.ColumnDefinitions.Count - 1)
        $btnRefresh.Margin = "0,0,5,3"
        $btnRefresh.Cursor = "Hand"
        $btnRefresh.Name = "btnRefresh"
        $btnRefresh.Focusable = $false
        $grdTitle.Children.Add($btnRefresh) | Out-Null

        $tooltip = [System.Windows.Controls.ToolTip]::new()
        $tooltip.Content = "Refresh all objects"
        [System.Windows.Controls.ToolTipService]::SetToolTip($btnRefresh, $tooltip)

        $panel.RegisterName($btnRefresh.Name, $btnRefresh)

        $tooltip = [System.Windows.Controls.ToolTip]::new()
        $tooltip.Content = "Refresh objects"

        [System.Windows.Controls.ToolTipService]::SetToolTip($btnRefresh, $tooltip)

        $btnRefresh.Add_Click({
            Invoke-RefreshObjects
        })
    }

    $script:UIProvider.AddXamlEvent($panel, "btnLoadAllPages", "add_click", {
        Add-GraphPoliciesFromPaging "AllRemainingPages"
    })

    $script:UIProvider.AddXamlEvent($panel, "btnLoadNextPage", "add_click", {
        Add-GraphPoliciesFromPaging "NextPage"
    })

    return $panel
}

function Add-GraphPoliciesFromPaging
{
    param(
        [ValidateSet("NextPage", "AllRemainingPages")]
        [string]    
        $PagingType
    )

    Write-Status "Loading $($script:IntuneManagerSelectedObject.Title) objects"

    Get-GraphPolicies -Paging $PagingType | Get-GraphPolicyForUIList | ForEach-Object { 
        $script:intuneManagerPolicyCollection.Add([PSObject]$_)
    }
    $script:dgIntuneManagerObjects.ItemsSource.CommitNew()

    Set-GraphPagesButtonStatus
    Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox $script:dgIntuneManagerObjects -ForceUpdate -ObjectsCountTextBox $script:IntuneManagementObjectsCount
    Write-Status ""
}

function Set-GraphPagesButtonStatus
{
    $IntuneManagerViewObject = Get-SingletonObject "IntuneViewObject"

    $IntuneManagerPanael = $IntuneManagerViewObject.ViewPanel

    $script:UIProvider.SetXamlProperty($IntuneManagerPanael, "btnLoadAllPages", "Visibility", (?: ($script:GraphPagingCache) "Visible" "Collapsed"))
    $script:UIProvider.SetXamlProperty($IntuneManagerPanael, "btnLoadNextPage", "Visibility", (?: ($script:GraphPagingCache) "Visible" "Collapsed"))
}

function Clear-GraphObjects
{   
    $IntuneManagerViewObject = Get-SingletonObject "IntuneViewObject"

    $IntuneManagerPanael = $IntuneManagerViewObject.ViewPanel

    $script:UIProvider.SetXamlProperty($IntuneManagerPanael, "txtFormTitle", "Text", "")
    $script:UIProvider.SetXamlProperty($IntuneManagerPanael, "txtObjectsCount", "Text", "")
    # Keep grdTitle visible — the empty colored strip is the desired idle state so
    # the right pane isn't an unframed expanse of background.

    if($script:dgIntuneManagerObjects -and $script:dgIntuneManagerObjects.Children) {
        $script:dgIntuneManagerObjects.Children.Clear()
        $script:dgIntuneManagerObjects.ItemsSource = $null
    }
    
    [System.Windows.Forms.Application]::DoEvents()
}

function Invoke-RefreshObjects
{
    if(-not $script:IntuneManagerSelectedObject) { return }
    $txtFilterText = $null
    $txtFilter = $this.Parent.FindName("txtFilter")
    if($txtFilter) { $txtFilterText = $txtFilter.Text } 

    Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject

    if($txtFilterText -and $txtFilter)
    {
        $txtFilter.Text = $txtFilterText
        Invoke-FilterBoxChanged $txtFilter $script:dgIntuneManagerObjects -ObjectsCountTextBox $script:IntuneManagementObjectsCount
    }

    Write-Status ""
}

function Invoke-IntuneActivatingView
{
    param($PreviousActiveView)

    Add-GraphBulkMenu

    if($null -eq $PreviousActiveView) {
        Set-SplashWindowText "Authenticating"
        $obj = Get-SingletonObject "IntuneViewObject"
        if($obj) { $obj.Authenticate() }
    }
}

function Invoke-IntuneDeactivatingView
{
    param($NewActiveView)

    $tmp = $script:mnuMain.Items | Where-Object Name -eq "IntuneBulk"
    if($tmp) { $script:mnuMain.Items.Remove($tmp) }
}

function Add-GraphBulkMenu
{
    # Item order matches the original project: Export, Import, Delete, Compare,
    # Copy, Documentation. Features that are new in this version (Scope Tags,
    # Assignments) are appended after a separator.
    $menuItem = [System.Windows.Controls.MenuItem]::new()
    $menuItem.Header = "_Bulk"
    $menuItem.Name = "IntuneBulk"

    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "_Export"
    $subItem.Add_Click({Show-GraphBulkExportForm})
    $menuItem.AddChild($subItem) | Out-Null

    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "_Import"
    $subItem.Add_Click({Show-GraphBulkImportForm})
    $menuItem.AddChild($subItem) | Out-Null

    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "_Delete"
    $subItem.Name = "mnuBulkDelete"
    $allowBulkDelete = Get-SettingValue "AllowBulkDelete"
    # Add it hidden even if not enabled, the save settings will enable it
    $subItem.Visibility = (?: ($allowBulkDelete -eq $true) "Visible" "Collapsed")
    $subItem.Add_Click({Show-GraphBulkDeleteForm})
    $menuItem.AddChild($subItem) | Out-Null

    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "C_ompare"
    $subItem.Add_Click({ Show-GraphBulkCompareForm })
    $menuItem.AddChild($subItem) | Out-Null

    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "Cop_y"
    $subItem.Add_Click({Show-GraphBulkCopyForm})
    $menuItem.AddChild($subItem) | Out-Null

    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "Doc_umentation"
    $subItem.Add_Click({Show-GraphBulkDocumentationForm})
    $menuItem.AddChild($subItem) | Out-Null

    $menuItem.AddChild(([System.Windows.Controls.Separator]::new())) | Out-Null

    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "_Scope Tags"
    $subItem.Add_Click({Show-GraphBulkScopeTagForm})
    $menuItem.AddChild($subItem) | Out-Null

    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "_Assignments"
    $subItem.Add_Click({Show-GraphBulkAssignmentsForm})
    $menuItem.AddChild($subItem) | Out-Null

    Add-MenuItem $menuItem 1
}

function Invoke-IntuneActivateObject
{
    param($SelectedObject)

    if($null -eq $SelectedObject) { 
        $script:dgIntuneManagerObjects.ItemsSource = $null 
        return
    }

    Write-Status "Activate $($SelectedObject.Title)"

    $script:IntuneManagerSelectedObject = $SelectedObject

    Clear-GraphObjects

    Show-SelectedGraphPolicies -SinglePage -Restart
}    

function Show-SelectedGraphPolicies
{
    param(
        [switch]
        $AllPages,
        [switch]
        $SinglePage,
        [switch]
        $Restart        
    )

    if($Restart -eq $true) {
        $script:IntuneManagerCurrentPage = $null 
    }

    if($null -eq $script:IntuneManagerSelectedObject) {
        return
    }

    $params = @{}
    if($script:IntuneManagerSelectedObject -is [IntunePolicyGroupBase]) {
        $params.Add("PolicyGroup", $script:IntuneManagerSelectedObject.Id)
    }
    elseif($script:IntuneManagerSelectedObject -is [IntunePolicyTypeBase]) {
        $params.Add("PolicyType", $script:IntuneManagerSelectedObject.Id)
    }

    if($SinglePage -eq $true -and (Get-SettingValue "GetAllPages") -ne $true) {
        $params.Add("SinglePage", $true)
    }

    Write-Status "Loading $($script:IntuneManagerSelectedObject.Title) objects"

    if($script:IntuneManagerSelectedObject.ShowForm -ne $false)
    {
        $IntuneManagerViewObject = Get-SingletonObject "IntuneViewObject"
        $IntuneManagerPanael = $IntuneManagerViewObject.ViewPanel
        $script:UIProvider.SetXamlProperty($IntuneManagerPanael, "txtFormTitle", "Text", $script:IntuneManagerSelectedObject.Title)
        try {
            $script:UIProvider.SetXamlProperty($IntuneManagerPanael, "ccIcon", "Content", $script:IntuneManagerSelectedObject.GetImage())
        }
        catch {
        }
        $script:UIProvider.SetXamlProperty($IntuneManagerPanael, "grdTitle", "Visibility", "Visible")
    }    
    
    $graphObjects = Get-GraphPolicies @params | Get-GraphPolicyForUIList

    $script:dgIntuneManagerObjects.AutoGenerateColumns = $false
    $script:dgIntuneManagerObjects.Columns.Clear()

    if($graphObjects)
    {
        $column = Get-GridCheckboxColumn "IsSelected"
        $script:dgIntuneManagerObjects.Columns.Add($column)

        $column.Header.add_Click({
            foreach($Item in $script:dgIntuneManagerObjects.ItemsSource)
            { 
                $Item.IsSelected = $this.IsChecked
            }
            $script:dgIntuneManagerObjects.Items.Refresh()
        })

        $additionalColumns = @()
        $selectedType = Get-SelectedObjectTypeString
        if(-not $selectedType) { return }
        $additionalColsStr = Get-SettingStoreValue "IntuneManager\ObjectColumns\$selectedType" "$($script:IntuneManagerSelectedObject.Id)"
    
        if($additionalColsStr)
        {
            $additionalColumns += $additionalColsStr.Split(',')
        }

        $columns = @()
        
        if($additionalColumns.Count -eq 0 -or $additionalColumns[0] -ne "0") 
        {
            $columns += ?? $script:IntuneManagerSelectedObject.ViewProperties @("Name","PolicyName","ID")
        }
        
        # Add custom columns
        foreach($additionalCol in $additionalColumns)
        {
            if($additionalCol -eq "0" -or $additionalCol -eq "1") { continue }
            $columns += $additionalCol
        }

        # Add columns
        # One line per row: a multi-line value (store app descriptions, script
        # bodies) otherwise makes the row as tall as its text and wrecks scrolling.
        # The full text stays on the cell tooltip. Setting: ObjectListFirstLineOnly.
        $firstLineOnly = ((Get-SettingValue "ObjectListFirstLineOnly") -eq $true)
        # Character cap on top of the first line: one very long description
        # otherwise widens its column until the others leave the screen.
        $maxCellLength = 50
        try { $maxCellLength = [int](Get-SettingValue "ObjectListMaxCellLength") } catch { }
        foreach($columnInfo in $columns)
        {
            $bindingProp,$colHeader = $columnInfo.Split('=')
            if(-not $colHeader) { $colHeader = $bindingProp.Split(".")[-1] }

            $dgIntuneManagerObjects.Columns.Add((New-GridTextColumn -Path $bindingProp -Header $colHeader -FirstLineOnly:$firstLineOnly -MaxLength $maxCellLength))
        }
    }
    else {
        $graphObjects = @()
    }

    $script:intuneManagerPolicyCollection = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new($graphObjects)
    $script:dgIntuneManagerObjects.ItemsSource = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:intuneManagerPolicyCollection)

    $policyTypes = Get-IntuneManagerSelectedPolicyTypes

    $allowDelete = Get-SettingValue "AllowDelete"

    Set-IntuneManagerUIButtonStatus @("btnDelete") $policyTypes -ForceHide:($allowDelete -eq $false)

    Set-IntuneManagerUIButtonStatus @("btnImport","btnView","btnExport","btnCompare","btnCopy","btnDocument") $policyTypes

    $script:UIProvider.SetXamlProperty($script:spIntuneManagerButtons, "btnImport", "IsEnabled", $true) # Always allow Import if ObjectType allows it

    Set-GraphPagesButtonStatus

    Invoke-AppEvent "IntuneManagerUISelectedMenuItemChanged" $policyTypes

    Write-Status ""
}

function Set-IntuneManagerUIButtonStatus
{
    param($Buttons, $PolicyTypes, [switch]$ForceHide)

    foreach($btn in $Buttons) {
        $visibility = "Collapsed"
        if($ForceHide -ne $true) {
            foreach($policyType in $PolicyTypes) {
                if(-not $policyType.ShowButtons -or ($policyType.ShowButtons | Where-Object { $btn -like "*$($_)" } )) {
                    $visibility = "Visible"
                    break
                }
            }
        }

        $script:UIProvider.SetXamlProperty($script:spIntuneManagerButtons, $btn, "Visibility", $visibility)
    }
}

function Get-GraphPolicyForUIList
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyBase[]]
        $InputObject
    )

    Begin { $graphObjects = @() }

    Process {
        foreach($policy in $InputObject) {
            $tmp = [PSCustomObject]$policy
            if((-not $tmp.PSObject.Properties['JsonObject']) -or ($null -eq $tmp.JsonObject)) {
                $tmp | Add-Member -NotePropertyName "JsonObject" -NotePropertyValue $policy.Object -Force
            }
            if((-not $tmp.PSObject.Properties['_TokenId']) -or ($null -eq $tmp._TokenId)) {
                $tmp | Add-Member -NotePropertyName "_TokenId" -NotePropertyValue $policy.TokenId -Force
            }
            $tmp | Add-Member -NotePropertyName "IsSelected" -NotePropertyValue $false
            $graphObjects += $tmp
        }
    }

    End { $graphObjects | Sort-Object -Property Name }
}

function Get-SelectedObjectTypeString {

    if($script:IntuneManagerSelectedObject -is [IntunePolicyGroupBase]) {
        "PolicyGroup"
    }
    elseif($script:IntuneManagerSelectedObject -is [IntunePolicyTypeBase]) {
        "PolicyType"
    }
    else {
        ""
    }
}


function Get-IntuneManagerSelectedPolicyTypes
{
    [CmdletBinding()]
    param()

    $policyTypes = @()
    
    if($script:IntuneManagerSelectedObject -is [IntunePolicyGroupBase]) {
        $policyTypes += $script:IntuneManagerSelectedObject.PolicyTypes
    }
    else {
        $policyTypes += $script:IntuneManagerSelectedObject
    }
    return $policyTypes
}

function Invoke-IntuneUIAppInitialized
{    
    # Settings registrations moved out of this handler (2026-08-15):
    # GraphPageSize -> Internal/IntuneManager.ps1 (engine-consumed);
    # GetAllPages / ObjectViewType -> UI/Classes/UICommonSettings.ps1 (shared by
    # both backends, registered once).

    Add-ViewObject (Get-SingletonObject "IntuneViewObject")

    Add-AppEventHandler "AuthenticatedNewToken" "Invoke-IntuneUIEventNewAuthentication"
    Add-AppEventHandler "AuthenticationUserDisconnected" "Invoke-IntuneUIEventUserDisconnected"
    Add-AppEventHandler "AuthenticationFailed" "Invoke-IntuneUIEventAuthenticationFailed"
}

function Invoke-IntuneUIEventSettingsUpdated
{
    if((Get-SettingValue "ObjectViewType") -ne (Get-CacheObject "ObjectViewType-Cached")) {
        Set-CacheObject "ObjectViewType-Cached" (Get-SettingValue "ObjectViewType")
        Show-ViewMenu
    }

    Update-IntuneDeleteVisibility
}

# Called by Show-ViewMenu (CoreUI) whenever the IntuneView becomes the
# active view OR its menu is rebuilt. Makes the title-bar config gear
# visible, syncs all its checkable states to current settings, and
# one-shot wires the click handlers. If "show item counts" is on, also
# kicks the count fetcher to paint counts into the freshly-bound rows.
function Update-MenuTitleConfigForIntuneView
{
    if(-not $script:btnMenuTitleConfig) { return }

    $script:btnMenuTitleConfig.Visibility = "Visible"
    $script:btnMenuTitleConfig.ToolTip    = "Menu options: switch view, refresh"

    $current = Get-SettingValue "ObjectViewType"
    # "Group" is the default (matches the ObjectViewType DefaultValue registered in
    # UI/Classes/UICommonSettings.ps1);
    # anything not "Type" reads as Group so a missing/blank setting still
    # ticks the Group option rather than leaving both unchecked.
    if($script:mnuMenuTitleConfigGroup) { $script:mnuMenuTitleConfigGroup.IsChecked = ($current -ne "Type") }
    if($script:mnuMenuTitleConfigType)  { $script:mnuMenuTitleConfigType.IsChecked  = ($current -eq "Type") }

    # Read-only "Intune Info" group is opt-in (hidden by default). Sync the toggle.
    if($script:mnuMenuTitleConfigShowInfo) { $script:mnuMenuTitleConfigShowInfo.IsChecked = ((Get-SettingStoreValue "IntuneManager" "ShowIntuneInfo" "false") -eq "true") }

    # "Show item counts" is HIDDEN per user decision (2026-06-12) — the menu
    # item is collapsed in MainWindow.xaml and the populate paths below are
    # forced off so a previously-saved true setting can't resurrect it. The
    # counts machinery (Get-IntuneTypeCounts + label projection) is kept for
    # potential re-enable.
    $showCounts = $false
    if($script:mnuMenuTitleConfigShowCounts) { $script:mnuMenuTitleConfigShowCounts.IsChecked = $showCounts }

    # Wire-once. The Click handlers persist for the window lifetime so a
    # subsequent view switch + return doesn't re-register them.
    if(-not $script:_menuTitleConfigWired) {
        $script:_menuTitleConfigWired = $true

        $script:mnuMenuTitleConfigGroup.Add_Click({ Set-IntuneMenuObjectViewType "Group" })
        $script:mnuMenuTitleConfigType.Add_Click({  Set-IntuneMenuObjectViewType "Type" })

        if($script:mnuMenuTitleConfigShowInfo) {
            $script:mnuMenuTitleConfigShowInfo.Add_Click({
                # IsCheckable items flip IsChecked before the handler runs. Persist +
                # rebuild the nav so the read-only Info group appears/disappears.
                $enabled = [bool]$script:mnuMenuTitleConfigShowInfo.IsChecked
                Save-SettingStoreValue -SubPath "IntuneManager" -Key "ShowIntuneInfo" -Value ($enabled.ToString().ToLower())
                Show-ViewMenu
            })
        }

        if($script:mnuMenuTitleConfigRefresh) {
            $script:mnuMenuTitleConfigRefresh.Add_Click({
                # Re-bind the menu list (cheap — no Graph). Count re-fetch
                # removed with the hidden "Show item counts" feature.
                Show-ViewMenu
            })
        }

        if($script:mnuMenuTitleConfigShowCounts) {
            $script:mnuMenuTitleConfigShowCounts.Add_Click({
                # IsCheckable items flip IsChecked before the handler runs.
                $enabled = [bool]$script:mnuMenuTitleConfigShowCounts.IsChecked
                Save-SettingStoreValue -SubPath "IntuneManager" -Key "ShowMenuItemCounts" -Value ($enabled.ToString().ToLower())
                if($enabled) {
                    Update-IntuneViewItemCounts
                }
                else {
                    # Strip counts back to bare Title without a fresh batch.
                    if($script:lstMenuItems -and $script:lstMenuItems.ItemsSource) {
                        foreach($it in $script:lstMenuItems.ItemsSource) {
                            $it | Add-Member -NotePropertyName "MenuLabel" -NotePropertyValue $it.Title -Force
                        }
                        try { $script:lstMenuItems.Items.Refresh() } catch { }
                    }
                }
            })
        }

        if($script:mnuMenuTitleConfigFilterPlatforms -and $script:btnMenuTitleConfig.ContextMenu) {
            # Populate on the parent ContextMenu's Opened (not the submenu's
            # SubmenuOpened). A MenuItem with zero child items doesn't render
            # the submenu arrow and never fires SubmenuOpened — so the items
            # have to exist BEFORE the user hovers into the submenu. Cheap:
            # in-memory scan of the already-bound right-pane rows, no Graph.
            $script:btnMenuTitleConfig.ContextMenu.Add_Opened({ Update-IntuneFilterPlatformsSubmenu })
        }
    }

    # First paint after rebind: if counts are on, populate them. Uses the
    # cached $script:_intuneTypeCounts so a view-switch back doesn't trigger
    # a fresh batch — only first activation or explicit Refresh does.
    if($showCounts) {
        Update-IntuneViewItemCounts
    }
}

# Rebuilds the Filter Platforms submenu from the currently-bound right-pane
# rows. Each entry is an IsCheckable MenuItem keyed on the platform string
# (with "" sentinel for null/empty Platform). Currently-selected platforms
# are always present even if absent from the data, so the user can untick
# a filter they previously enabled but that no longer matches anything.
function Update-IntuneFilterPlatformsSubmenu
{
    if(-not $script:mnuMenuTitleConfigFilterPlatforms) { return }

    # 1. Collect platforms in the current grid
    $available = [System.Collections.Generic.HashSet[string]]::new()
    $hasNull   = $false
    if($script:dgIntuneManagerObjects -and $script:dgIntuneManagerObjects.ItemsSource) {
        $src = $script:dgIntuneManagerObjects.ItemsSource
        if($src -is [System.Windows.Data.ListCollectionView]) { $src = $src.SourceCollection }
        if($src) {
            foreach($it in $src) {
                $plat = $it.Platform
                if([string]::IsNullOrEmpty($plat)) { $hasNull = $true }
                else { [void]$available.Add([string]$plat) }
            }
        }
    }

    # 2. Include currently-selected filters so the user can untick them even
    #    if the new view no longer has any matching rows.
    if($script:_intunePlatformFilter) {
        foreach($s in $script:_intunePlatformFilter) {
            if($s -eq "") { $hasNull = $true }
            else { [void]$available.Add($s) }
        }
    }

    $script:mnuMenuTitleConfigFilterPlatforms.Items.Clear()

    $sorted = @($available | Sort-Object)
    if($sorted.Count -eq 0 -and -not $hasNull) {
        $empty = [System.Windows.Controls.MenuItem]::new()
        $empty.Header    = "(no policies loaded)"
        $empty.IsEnabled = $false
        [void]$script:mnuMenuTitleConfigFilterPlatforms.Items.Add($empty)
        return
    }

    # 3. One MenuItem per platform; "(no platform)" pinned at the bottom so
    #    it visually separates from named platforms.
    foreach($plat in $sorted) {
        $mi = [System.Windows.Controls.MenuItem]::new()
        $mi.Header      = $plat
        $mi.IsCheckable = $true
        $mi.IsChecked   = ($script:_intunePlatformFilter -and $script:_intunePlatformFilter.Contains($plat))
        # Tag stores the canonical key (== Header for named platforms; "" for
        # the no-platform row). Read in the click handler so the same handler
        # works for both.
        $mi.Tag         = $plat
        $mi.Add_Click({ Switch-IntunePlatformFilter ([string]$this.Tag) ([bool]$this.IsChecked) })
        [void]$script:mnuMenuTitleConfigFilterPlatforms.Items.Add($mi)
    }
    if($hasNull) {
        $mi = [System.Windows.Controls.MenuItem]::new()
        $mi.Header      = "(no platform)"
        $mi.IsCheckable = $true
        $mi.IsChecked   = ($script:_intunePlatformFilter -and $script:_intunePlatformFilter.Contains(""))
        $mi.Tag         = ""
        $mi.Add_Click({ Switch-IntunePlatformFilter ([string]$this.Tag) ([bool]$this.IsChecked) })
        [void]$script:mnuMenuTitleConfigFilterPlatforms.Items.Add($mi)
    }

    # 4. "Clear filter" footer when any filter is active.
    if($script:_intunePlatformFilter -and $script:_intunePlatformFilter.Count -gt 0) {
        [void]$script:mnuMenuTitleConfigFilterPlatforms.Items.Add([System.Windows.Controls.Separator]::new())
        $clear = [System.Windows.Controls.MenuItem]::new()
        $clear.Header = "Clear filter"
        $clear.Add_Click({
            if($script:_intunePlatformFilter) { $script:_intunePlatformFilter.Clear() }
            if($script:dgIntuneManagerObjects) {
                Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox $script:dgIntuneManagerObjects -ForceUpdate -ObjectsCountTextBox $script:IntuneManagementObjectsCount
            }
        })
        [void]$script:mnuMenuTitleConfigFilterPlatforms.Items.Add($clear)
    }
}

# Toggle a single platform on/off in $script:_intunePlatformFilter and
# re-evaluate the grid filter. Centralised so the submenu click handlers
# stay one-liners.
function Switch-IntunePlatformFilter
{
    param([string]$PlatformKey, [bool]$Enable)

    if($null -eq $script:_intunePlatformFilter) {
        $script:_intunePlatformFilter = [System.Collections.Generic.HashSet[string]]::new()
    }
    if($Enable) { [void]$script:_intunePlatformFilter.Add($PlatformKey) }
    else        { [void]$script:_intunePlatformFilter.Remove($PlatformKey) }

    if($script:dgIntuneManagerObjects) {
        Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox $script:dgIntuneManagerObjects -ForceUpdate -ObjectsCountTextBox $script:IntuneManagementObjectsCount
    }
}

# Single source of truth for switching the menu mode from the gear. Persists
# the setting so the choice survives a restart, keeps the SettingsUpdated
# cache key in sync so the Settings dialog won't fire a redundant rebuild,
# and rebuilds the menu list. Idempotent — selecting the same option twice
# is a no-op.
function Set-IntuneMenuObjectViewType
{
    param([ValidateSet("Group","Type")][string]$ViewType)

    $current = Get-SettingValue "ObjectViewType"
    if($current -eq $ViewType) {
        # Sync checkboxes anyway in case the user toggled one off — IsCheckable
        # menu items uncheck on click before this handler runs.
        if($script:mnuMenuTitleConfigGroup) { $script:mnuMenuTitleConfigGroup.IsChecked = ($ViewType -eq "Group") }
        if($script:mnuMenuTitleConfigType)  { $script:mnuMenuTitleConfigType.IsChecked  = ($ViewType -eq "Type") }
        return
    }

    Save-SettingStoreValue -SubPath "IntuneManager" -Key "ObjectViewType" -Value $ViewType
    Set-CacheObject "ObjectViewType-Cached" $ViewType

    if($script:mnuMenuTitleConfigGroup) { $script:mnuMenuTitleConfigGroup.IsChecked = ($ViewType -eq "Group") }
    if($script:mnuMenuTitleConfigType)  { $script:mnuMenuTitleConfigType.IsChecked  = ($ViewType -eq "Type") }

    Show-ViewMenu
}

function Invoke-IntuneUIEventShowMainWindow
{
    Update-IntuneDeleteVisibility
}

function Invoke-IntuneUISelectedItemChanged
{

}

function Invoke-IntuneUIEventNewAuthentication
{
    param($TokenInfo)

    Update-IntuneDeleteVisibility

    # Access marking (Internal/AccessLevel.ps1) needs the signed-in token, which
    # only exists once login completes. The nav was built white before the token
    # arrived, so rebuild it now to repaint the AccessType colours - otherwise the
    # menu stays uncoloured until the next view switch. AccessType is a plain
    # property with no change notification, so only a full ItemsSource rebuild
    # re-runs the colour triggers; refreshing the object grid alone does nothing
    # for the menu. Preserve and restore the selected menu item across the
    # rebuild: setting ItemsSource nulls the selection, which fires
    # OnItemChanged($null) and would blank a loaded object list. Restoring it
    # reloads that list (the same effect the old Invoke-RefreshObjects had).
    if(($script:ActiveView -and $script:ActiveView.ID -eq "IntuneManagement") -and
       (Get-Command Show-ViewMenu -ErrorAction SilentlyContinue)) {
        $selected = if($script:lstMenuItems) { $script:lstMenuItems.SelectedItem } else { $null }
        Show-ViewMenu
        if($selected -and $script:lstMenuItems) { $script:lstMenuItems.SelectedItem = $selected }
    }
    else {
        Invoke-RefreshObjects
    }
}

function Invoke-IntuneUIEventUserDisconnected
{
    param($TokenInfo)

    Update-IntuneDeleteVisibility

    # Sign-out is the mirror of Invoke-IntuneUIEventNewAuthentication: with no
    # token, Update-IntuneAccessLevels (run inside Show-ViewMenu -> GetViewItems)
    # resets every type/group back to Full, clearing the orange/red marking. But
    # AccessType is a plain property with no change notification, so the colours
    # only clear when the menu's ItemsSource is rebuilt - rebuild it here.
    # Unlike the login handler, do NOT restore the selection: letting it fall to
    # null fires OnItemChanged($null), which clears the object list that the
    # signed-out session can no longer load.
    if(($script:ActiveView -and $script:ActiveView.ID -eq "IntuneManagement") -and
       (Get-Command Show-ViewMenu -ErrorAction SilentlyContinue)) {
        Show-ViewMenu
    }
}

function Invoke-IntuneUIEventAuthenticationFailed
{
    param($TokenInfo)

    # AuthenticationFailed also fires on transient failures - this tenant's WAM
    # broker fails every silent refresh while the token is still perfectly valid
    # (see the SuppressFailedEvent handling in the auth core) - so rebuilding on
    # every failure would wrongly wipe the marking mid-session. Only treat it as a
    # sign-out when the default token is genuinely past expiry; Test-DefaultTokenExpired
    # returns $false for still-valid and SDK-managed tokens. When it is expired, the
    # reset is identical to an explicit disconnect, so reuse that handler (no
    # selection restore -> the object grid clears too).
    if(-not (Test-DefaultTokenExpired)) { return }
    Invoke-IntuneUIEventUserDisconnected $TokenInfo
}

# Refresh the visibility of every place that exposes a delete action so that toggling
# 'AllowDelete' / 'AllowBulkDelete' in Settings or signing in to a new tenant takes effect
# without having to restart the UI or re-enter a view. Safe to call before the relevant
# controls exist -- each path checks for the control first.
function Update-IntuneDeleteVisibility
{
    Update-IntuneManagerDeleteButton
    Update-IntuneBulkDeleteMenu
}

function Update-IntuneManagerDeleteButton
{
    if(-not $script:spIntuneManagerButtons) { return }
    $btn = $script:spIntuneManagerButtons.FindName("btnDelete")
    if(-not $btn) { return }

    $allowDelete = (Get-SettingValue "AllowDelete") -eq $true
    # Reuse the regular button-status path so the current policy type's ShowButtons rules
    # still apply -- don't expose Delete when no compatible type is selected.
    $policyTypes = @()
    if(Get-Command Get-IntuneManagerSelectedPolicyTypes -ErrorAction SilentlyContinue)
    {
        $policyTypes = @(Get-IntuneManagerSelectedPolicyTypes)
    }
    Set-IntuneManagerUIButtonStatus @("btnDelete") $policyTypes -ForceHide:(-not $allowDelete)
}

function Update-IntuneBulkDeleteMenu
{
    if(-not $script:mnuMain) { return }
    $bulkMenu = $script:mnuMain.Items | Where-Object { $_.Name -eq "IntuneBulk" } | Select-Object -First 1
    if(-not $bulkMenu) { return }
    $deleteItem = $bulkMenu.Items | Where-Object { $_.Name -eq "mnuBulkDelete" } | Select-Object -First 1
    if(-not $deleteItem) { return }

    $allowBulkDelete = (Get-SettingValue "AllowBulkDelete") -eq $true
    $deleteItem.Visibility = if($allowBulkDelete) { "Visible" } else { "Collapsed" }
}

#endregion

function Invoke-FilterBoxChanged
{
    param($TxtBox, $DataSource, [switch]$ForceUpdate, $ObjectsCountTextBox)

    if($DataSource.ItemsSource -is [System.Windows.Data.ListCollectionView])
    {
        # Capture the active text + platform filter into the predicate closure.
        # The platform set lives at $script: so a separate code path (the gear's
        # Filter Platforms submenu) can mutate it and then call this with
        # -ForceUpdate to re-evaluate the bound rows.
        $textValue   = if($TxtBox) { [string]$TxtBox.Text } else { "" }
        $platformSet = $script:_intunePlatformFilter
        $hasText     = -not [string]::IsNullOrEmpty($textValue)
        $hasPlatform = ($null -ne $platformSet -and $platformSet.Count -gt 0)

        if($hasText -or $hasPlatform) {
            # Text match runs ONLY against the fields actually shown in this grid,
            # derived from the grid's own bound columns (Name, Description, and
            # whatever else the view/ObjectColumns configure). Deriving from
            # $DataSource.Columns keeps this generic for every grid that calls this
            # function AND avoids matching heavy off-grid note-properties - the row
            # objects carry the full policy under 'JsonObject' plus '_TokenId', and
            # regex-matching those on every row/keystroke was the source of the lag.
            # Computed once here (not per row) and captured in the closure.
            $columnPaths = @()
            foreach($col in $DataSource.Columns) {
                if($col -is [System.Windows.Controls.DataGridBoundColumn] -and $col.Binding -and $col.Binding.Path) {
                    $p = [string]$col.Binding.Path.Path
                    if($p -and $p -ne "IsSelected") { $columnPaths += $p }
                }
            }
            # -like with the wildcard metacharacters escaped == a plain, fast,
            # case-insensitive substring match (no regex engine per field).
            $likePattern = if($hasText) { "*" + [System.Management.Automation.WildcardPattern]::Escape($textValue) + "*" } else { $null }

            $DataSource.ItemsSource.Filter = {
                param($Item)

                # Platform filter — empty Platform maps to the "" sentinel so a
                # "(no platform)" selection can target rows whose .Platform is
                # null/empty without ambiguity.
                if($hasPlatform) {
                    $plat = $Item.Platform
                    $key = if([string]::IsNullOrEmpty($plat)) { "" } else { [string]$plat }
                    if(-not $platformSet.Contains($key)) { return $false }
                }

                if($hasText) {
                    if($columnPaths.Count -gt 0) {
                        foreach($path in $columnPaths) {
                            # Resolve the (possibly dotted, e.g. "PolicyType.Title") bound path.
                            $val = $Item
                            foreach($seg in $path.Split('.')) {
                                if($null -eq $val) { break }
                                $val = $val.$seg
                            }
                            if($null -ne $val -and ([string]$val) -like $likePattern) { return $true }
                        }
                        return $false
                    }
                    # Defensive fallback (grid exposed no bound columns): scan
                    # properties but skip the internal/heavy ones so it can never
                    # regress to matching the whole policy object.
                    return ($null -ne ($Item.PSObject.Properties | Where-Object {
                        $_.Name -notin @("IsSelected","Object","ObjectType","JsonObject","_TokenId") -and
                        ([string]$_.Value) -like $likePattern
                    }))
                }
                return $true
            }.GetNewClosure()
        }
        else {
            $DataSource.ItemsSource.Filter = $null
        }

        if($ForceUpdate)
        {
            $DataSource.ItemsSource.Refresh()
        }
    }

    if($ObjectsCountTextBox)
    {
        # loadedCount = rows pulled from Graph and cached locally (across all loaded pages)
        # visibleCount = rows currently rendered after the txtFilter predicate
        # If $script:GraphPagingCache is non-empty, more pages are available on the server
        # (the Load More / Load All buttons are visible). We mark loadedCount with a "+"
        # in that case so the user can tell the count isn't the full set.
        $loadedCount = 0
        if($DataSource.ItemsSource.SourceCollection) {
            $loadedCount = $DataSource.ItemsSource.SourceCollection.Count
        }
        $visibleCount = ($DataSource.ItemsSource | Measure-Object).Count
        $morePages    = [bool]$script:GraphPagingCache
        $loadedLabel  = if($morePages) { "$loadedCount+" } else { "$loadedCount" }

        if($loadedCount -le 0) {
            $ObjectsCountTextBox.Text = ""
        }
        elseif($visibleCount -lt $loadedCount) {
            # Filter is active and hides some loaded rows.
            $ObjectsCountTextBox.Text = "Showing $visibleCount of $loadedLabel"
        }
        else {
            # No filter (or filter matches everything loaded).
            $ObjectsCountTextBox.Text = if($morePages) {
                "Objects: $loadedLabel (more available - click Load All)"
            } else {
                "Objects: $loadedCount"
            }
        }
    }
}

#region Bulk Forms
