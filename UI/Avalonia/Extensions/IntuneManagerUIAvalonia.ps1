# Avalonia port of UI/WPF/Extensions/IntuneManagerUI.ps1.
#
# Slice 2 + Slice 3a + Slice 3b1+3b2 + Slice 3c + Slice 3d1+3d2+3d3 +
# Slice 3e + Slice 3f + Slice 4a + Slice 4b + Slice 4c + Slice 4d +
# Slice 4e + Slice 4f: shell + activation + real data load + dynamic
# DataGrid columns + filter/refresh + paging + Delete + Copy (incl.
# scope-tag dual-list) + Export (without per-policy-type UI extensions) +
# Compare (Intune-vs-Intune AND file compare, with Match-based row
# coloring + Save/Copy output) + Import + View (Object Details — Json /
# Settings / Columns tabs, without per-policy-type AddUIDetailsExtension
# hooks) + Bulk Export + Bulk Import + Bulk Delete + Bulk Scope Tags +
# Bulk Assignments + Bulk Compare. All bulk forms ported.
# Copy, Export, Compare, Import,
# View, Bulk Export, Bulk Import, Bulk Delete, Bulk Scope Tags, Bulk
# Assignments, and Bulk Compare live in their own files
# (Extensions/IntuneManagerCopyUI.ps1, IntuneManagerExportUI.ps1,
# IntuneManagerCompareUI.ps1, IntuneManagerImportUI.ps1,
# IntuneManagerDetailsUI.ps1, IntuneManagerBulkExportUI.ps1,
# IntuneManagerBulkImportUI.ps1, IntuneManagerBulkDeleteUI.ps1,
# IntuneManagerBulkScopeTagUI.ps1, IntuneManagerBulkAssignmentsUI.ps1,
# IntuneManagerBulkCompareUI.ps1) per architecture R9.

Add-AppEvent "IntuneManagerUISelectedMenuItemChanged"
Add-AppEvent "IntuneManagerUISelectedItemChanged"

Add-AppEventHandler "AppInitialized"  "Invoke-IntuneUIAppInitialized"
Add-AppEventHandler "SettingsUpdated" "Invoke-IntuneUIEventSettingsUpdated"

function Get-IntuneViewItems
{
    [OutputType([Object[]])]
    param()

    $viewType = Get-SettingValue "ObjectViewType"
    if ($viewType -eq "Type") {
        $viewItems = $script:IntuneTypes  | Sort-Object -Property Title
    } else {
        $viewItems = $script:IntuneGroups | Sort-Object -Property Title
    }

    # Read-only "Intune Info" group (Baseline Templates, Tenant Settings, etc.) is
    # opt-in: hidden unless enabled via the gear menu's "Show read-only Info". Drop
    # the group (Group view) or its member types (API view).
    if ((Get-SettingStoreValue "IntuneManager" "ShowIntuneInfo" "false") -ne "true") {
        if ($viewType -eq "Type") {
            $viewItems = @($viewItems | Where-Object { -not ($_.PolicyGroup -and $_.PolicyGroup.Id -eq "IntuneInfo") })
        } else {
            $viewItems = @($viewItems | Where-Object { $_.Id -ne "IntuneInfo" })
        }
    }

    # Stamp AccessType / AccessInfo from the token's granted scopes so restricted
    # rows can be coloured (orange = read-only, red = no access). Groups
    # aggregate their member types. No-op when there is no token to diff against.
    Update-IntuneAccessLevels

    if ((Get-SettingValue 'HideNoAccess')) {
        # Drop rows the token cannot use at all. Groups aggregate to None only
        # when every member type is unusable, so a partially usable group is
        # never hidden - it stays visible and orange.
        $viewItems = @($viewItems | Where-Object { $_.AccessType -ne [APIAccess]::None })
    }

    # Icons are parsed here, after the splash is up, not at module import.
    Initialize-MenuItemIcons -Items $viewItems

    # Avalonia's binder needs real CLR properties for the {Binding MenuLabel}
    # in the lstMenuItems DataTemplate — Add-Member NoteProperty silently fails.
    # Wrap each underlying type in a ViewMenuItem and stash the original on Tag
    # so OnItemChanged can recover it.
    $wrapped = foreach ($item in $viewItems) {
        if ($null -eq $item) { continue }
        $access = [string]$item.AccessType
        [ViewMenuItem]@{
            Id              = [string]$item.ID
            Title           = [string]$item.Title
            MenuLabel       = [string]$item.Title
            IconImage       = $item.IconImage
            AccessType      = $access
            # $null rather than '' so rows with nothing to say get no tooltip.
            AccessInfo      = if ($item.AccessInfo) { [string]$item.AccessInfo } else { $null }
            IsAccessLimited = ($access -eq 'Limited')
            IsAccessNone    = ($access -eq 'None')
            Tag             = $item
        }
    }

    return @($wrapped)
}

function Get-IntuneManagementViewPanel
{
    $ui = $script:UIProvider
    $panel = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/IntuneManagerPanel.axaml'))
    if (-not $panel) { return $null }

    $hostType = Get-AvaloniaHost
    $script:dgIntuneManagerObjects        = $hostType::FindByName($panel, 'dgIntuneManagerObjects')
    $script:spIntuneManagerButtons        = $hostType::FindByName($panel, 'spSubMenu')
    $script:IntuneManagementFilterTextBox = $hostType::FindByName($panel, 'txtFilter')
    $script:IntuneManagementObjectsCount  = $hostType::FindByName($panel, 'txtObjectsCount')
    $script:grdIntuneManagerNotLoggedIn   = $hostType::FindByName($panel, 'grdNotLoggedIn')

    $ui.AddXamlEvent($panel, 'btnView', 'Add_Click', ({ Show-IntuneManagerDetailedView }))
    $ui.AddXamlEvent($panel, 'btnCompare', 'Add_Click', ({ Invoke-IntuneManagerCompare }))
    $ui.AddXamlEvent($panel, 'btnCopy', 'Add_Click', ({ Copy-IntuneManagerPolicy           }))
    $ui.AddXamlEvent($panel, 'btnDelete', 'Add_Click', ({ Remove-GraphObjectsUI            }))
    $ui.AddXamlEvent($panel, 'btnImport', 'Add_Click', ({ Show-IntuneManagerImportForm     }))
    $ui.AddXamlEvent($panel, 'btnExport', 'Add_Click', ({ Show-IntuneManagerExportForm     }))
    $ui.AddXamlEvent($panel, 'btnDocument', 'Add_Click', ({
        # Prefer IsSelected-checked rows; fall back to highlighted; fall back
        # to all visible. Scans run over the grid's CURRENT (filtered)
        # ItemsSource, never the unfiltered backing list - a checked row that
        # the filter hides must not be acted on.
        $visibleRows = @($script:dgIntuneManagerObjects.ItemsSource)
        $items = @($visibleRows | Where-Object { $_.IsSelected } | ForEach-Object { $_.Source } | Where-Object { $_ })
        if ($items.Count -eq 0 -and $script:dgIntuneManagerObjects.SelectedItem) {
            $row = $script:dgIntuneManagerObjects.SelectedItem
            $candidate = if ($row.PSObject.Properties['Source']) { $row.Source } else { $row }
            if ($candidate) { $items = @($candidate) }
        }
        if ($items.Count -eq 0) {
            $items = @($visibleRows | ForEach-Object { $_.Source } | Where-Object { $_ })
        }
        if ($items.Count -eq 0) { return }
        Show-IntuneManagerDocumentForm -PolicyObject $items
    }))
    $ui.AddXamlEvent($panel, 'btnRefresh', 'Add_Click', ({ Invoke-RefreshObjects }))

    $ui.AddXamlEvent($panel, 'btnLoadAllPages', 'Add_Click', ({ Add-GraphPoliciesFromPaging 'AllRemainingPages' }))
    $ui.AddXamlEvent($panel, 'btnLoadNextPage', 'Add_Click', ({ Add-GraphPoliciesFromPaging 'NextPage'          }))

    $ui.AddXamlEvent($panel, 'txtFilter', 'Add_TextChanged', ({
        Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox `
            $script:dgIntuneManagerObjects `
            -ObjectsCountTextBox $script:IntuneManagementObjectsCount
    }))

    if ($script:dgIntuneManagerObjects) {
        $script:dgIntuneManagerObjects.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
            Update-IntuneManagerActionButtons
        }))
    }

    return $panel
}

function Invoke-IntuneManagerNotPorted
{
    param([string]$Action)
    $ui = $script:UIProvider
    $ui.ShowMessageBox("$Action is not yet ported to the Avalonia backend.", "Not yet ported")
}

function Update-IntuneManagerActionButtons
{
    if (-not $script:dgIntuneManagerObjects -or -not $script:spIntuneManagerButtons) { return }

    $items = $script:dgIntuneManagerObjects.ItemsSource
    $count = if ($items) { @($items).Count } else { 0 }
    # WPF enables the action buttons when rows are CHECKED as well as when a
    # row is highlighted. Without the checked clause, ticking rows (including
    # select-all) left every button disabled.
    #
    # This runs on every SelectionChanged, so only ask whether ANY row is
    # checked - a Where-Object pass over the whole ItemsSource stutters keyboard
    # navigation once Load All has pulled in thousands of rows. A highlighted
    # row short-circuits the scan entirely (the common click case).
    $hasSelectedItems = $false
    if ($count -gt 0) {
        if ($null -ne $script:dgIntuneManagerObjects.SelectedItem) {
            $hasSelectedItems = $true
        }
        else {
            foreach ($row in $items) {
                if ($row.IsSelected) { $hasSelectedItems = $true; break }
            }
        }
    }

    foreach ($name in 'btnView','btnCompare','btnCopy','btnDelete') {
        $btn = (Get-AvaloniaHost)::FindByName($script:spIntuneManagerButtons, $name)
        if ($btn) { $btn.IsEnabled = $hasSelectedItems }
    }
    $btnExport = (Get-AvaloniaHost)::FindByName($script:spIntuneManagerButtons, 'btnExport')
    if ($btnExport) { $btnExport.IsEnabled = $hasSelectedItems }
    # btnDocument stays enabled whenever any rows are loaded — clicking with
    # nothing selected documents the whole view (matches OLD project's UX).
    $btnDocument = (Get-AvaloniaHost)::FindByName($script:spIntuneManagerButtons, 'btnDocument')
    if ($btnDocument) { $btnDocument.IsEnabled = ($count -gt 0) }

    Invoke-AppEvent "IntuneManagerUISelectedItemChanged" $hasSelectedItems
}


function Invoke-IntuneActivatingView
{
    param($PreviousActiveView)

    Add-GraphBulkMenu

    if ($null -eq $PreviousActiveView) {
        # WPF parity routes this through the splash window — no splash in
        # Avalonia, so write to the status overlay and clear it once
        # silent-auth finishes (success surfaces via AuthenticatedNewToken;
        # failure / no-op leaves us on the not-logged-in overlay).
        Write-Status "Authenticating"
        try {
            $obj = Get-SingletonObject "IntuneViewObject"
            if ($obj) { $obj.Authenticate() }
        } finally {
            Write-Status ""
        }
    }
}

function Invoke-IntuneDeactivatingView
{
    param($NewActiveView)

    if ($script:mnuMain) {
        $tmp = $script:mnuMain.Items | Where-Object { $_.Name -eq 'IntuneBulk' }
        if ($tmp) { $script:mnuMain.Items.Remove($tmp) | Out-Null }
    }
}

function Add-GraphBulkMenu
{
    if (-not $script:mnuMain) { return }

    # Skip if already added (re-activations).
    $existing = $script:mnuMain.Items | Where-Object { $_.Name -eq 'IntuneBulk' }
    if ($existing) { return }

    $menuItem      = New-Object Avalonia.Controls.MenuItem
    $menuItem.Header = '_Bulk'
    $menuItem.Name   = 'IntuneBulk'

    # Item order matches the original project: Export, Import, Delete, Compare,
    # Copy, Documentation. Features that are new in this version (Scope Tags,
    # Assignments) are appended after a separator.
    $bulkActions = @(
        @{ Header = '_Export';        Action = { Show-GraphBulkExportForm } }
        @{ Header = '_Import';        Action = { Show-GraphBulkImportForm } }
        @{ Header = '_Delete';        Action = { Show-GraphBulkDeleteForm }; Name = 'mnuBulkDelete' }
        @{ Header = 'C_ompare';       Action = { Show-GraphBulkCompareForm } }
        @{ Header = 'Cop_y';          Action = { Show-GraphBulkCopyForm } }
        @{ Header = 'Doc_umentation'; Action = { Show-GraphBulkDocumentationForm } }
        @{ Separator = $true }
        @{ Header = '_Scope Tags';    Action = { Show-GraphBulkScopeTagForm } }
        @{ Header = '_Assignments';   Action = { Show-GraphBulkAssignmentsForm } }
    )

    foreach ($entry in $bulkActions) {
        if ($entry.Separator) {
            $menuItem.Items.Add((New-Object Avalonia.Controls.Separator)) | Out-Null
            continue
        }
        $sub        = New-Object Avalonia.Controls.MenuItem
        $sub.Header = $entry.Header
        if ($entry.Name) { $sub.Name = $entry.Name }
        $sub.add_Click((ConvertTo-AvaloniaEventScriptBlock $entry.Action))
        if ($entry.Name -eq 'mnuBulkDelete') {
            $sub.IsVisible = ((Get-SettingValue 'AllowBulkDelete') -eq $true)
        }
        $menuItem.Items.Add($sub) | Out-Null
    }

    Add-MenuItem $menuItem 1
}

function Invoke-IntuneActivateObject
{
    param($SelectedItem)

    $payload = $null
    if ($SelectedItem -is [ViewMenuItem]) {
        $payload = $SelectedItem.Tag
    } else {
        $payload = $SelectedItem
    }

    if ($null -eq $payload) {
        if ($script:dgIntuneManagerObjects) {
            $script:dgIntuneManagerObjects.ItemsSource = $null
        }
        return $null
    }

    Write-Status "Activate $($payload.Title)"
    $script:IntuneManagerSelectedObject = $payload

    Clear-GraphObjects

    Show-SelectedGraphPolicies -SinglePage -Restart
    return $null
}

function Get-SelectedObjectTypeString
{
    if ($script:IntuneManagerSelectedObject -is [IntunePolicyGroupBase]) { return "PolicyGroup" }
    if ($script:IntuneManagerSelectedObject -is [IntunePolicyTypeBase])  { return "PolicyType"  }
    return ""
}

function Get-IntuneManagerSelectedPolicyTypes
{
    [CmdletBinding()]
    param()

    $policyTypes = @()
    if ($script:IntuneManagerSelectedObject -is [IntunePolicyGroupBase]) {
        $policyTypes += $script:IntuneManagerSelectedObject.PolicyTypes
    } else {
        $policyTypes += $script:IntuneManagerSelectedObject
    }
    return $policyTypes
}

function Get-GraphPolicyForUIList
{
    [CmdletBinding()]
    [OutputType([IntuneObjectRowItem[]])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyBase[]]
        $InputObject
    )

    Begin {
        $rows = [System.Collections.Generic.List[IntuneObjectRowItem]]::new()
    }

    Process {
        foreach ($policy in $InputObject) {
            if ($null -eq $policy) { continue }

            $row = [IntuneObjectRowItem]::new()
            $row.IsSelected   = $false
            $row.Name         = [string]$policy.Name
            $row.PolicyName   = [string]$policy.PolicyName
            $row.Platform     = [string]$policy.Platform
            $row.Description  = if ($policy.Description) { [string]$policy.Description } else { '' }
            $row.LastModified = if ($policy.LastModified) { [string]$policy.LastModified } else { '' }
            $row.Created      = if ($policy.Created)      { [string]$policy.Created      } else { '' }
            $row.ID           = [string]$policy.ID
            $row.JsonObject   = $policy.Object
            $row.Source       = $policy
            $row.TokenId      = if ($null -ne $policy.TokenId) { [int]$policy.TokenId } else { 0 }
            try { $row.ScopeTags = [string]$policy.ScopeTags } catch { }

            [void]$rows.Add($row)
        }
    }

    End {
        # In-place typed sort. Routing through `Sort-Object` would push each
        # IntuneObjectRowItem through the PSObject pipeline which stamps an ETS
        # wrapper on every item; Avalonia's DataGrid text-column binder reflects
        # on the runtime type and renders blank rows against the wrapper
        # (same failure mode as the Where-Object pipeline in the
        # Update-IntuneAssignmentsFilter fix — see [[avalonia-binding-needs-clr-types]]).
        $rows.Sort([Comparison[IntuneObjectRowItem]]{
            param($a, $b)
            [string]::Compare($a.Name, $b.Name, [StringComparison]::OrdinalIgnoreCase)
        })

        # Emit each row individually as a bare IntuneObjectRowItem. A
        # `,$rows` emit would treat the List as a single pipeline item (List<T>
        # is not auto-unrolled like Object[]) and the caller's `@()` would
        # collapse to a one-element array holding the whole List, which the
        # DataGrid then renders as one blank row.
        foreach ($r in $rows) { $r }
    }
}

function Show-SelectedGraphPolicies
{
    param([switch]$AllPages, [switch]$SinglePage, [switch]$Restart)

    $ui = $script:UIProvider
    if ($Restart) { $script:IntuneManagerCurrentPage = $null }
    if (-not $script:IntuneManagerSelectedObject) { return }

    $params = @{}
    if     ($script:IntuneManagerSelectedObject -is [IntunePolicyGroupBase]) { $params.Add('PolicyGroup', $script:IntuneManagerSelectedObject.Id) }
    elseif ($script:IntuneManagerSelectedObject -is [IntunePolicyTypeBase])  { $params.Add('PolicyType',  $script:IntuneManagerSelectedObject.Id) }

    if ($SinglePage -and (Get-SettingValue 'GetAllPages') -ne $true) {
        $params.Add('SinglePage', $true)
    }

    Write-Status "Loading $($script:IntuneManagerSelectedObject.Title) objects"

    $viewObject = Get-SingletonObject "IntuneViewObject"
    $panel = if ($viewObject) { $viewObject.ViewPanel } else { $null }
    if ($panel -and $script:IntuneManagerSelectedObject.ShowForm -ne $false) {
        $ui.SetXamlProperty($panel, 'txtFormTitle', 'Text', $script:IntuneManagerSelectedObject.Title)
        try {
            $img = $script:IntuneManagerSelectedObject.GetImage()
            if ($img) { $ui.SetXamlProperty($panel, 'ccIcon', 'Content', $img) }
        } catch { }
    }

    $rows = @(Get-GraphPolicies @params | Get-GraphPolicyForUIList)


    if ($script:dgIntuneManagerObjects) {
        $script:dgIntuneManagerObjects.AutoGenerateColumns = $false
        $script:dgIntuneManagerObjects.Columns.Clear()

        if ($rows.Count -gt 0) {
            Add-IntuneManagerDataGridColumns
        }


        # Avalonia DataGrid binds straight against an enumerable; the WPF
        # CollectionViewSource path doesn't apply. Filter is a subset rebuild
        # (see Invoke-FilterBoxChanged) — stash the unfiltered list so the
        # filter can restore it.
        $script:IntuneManagerAllRows = [System.Collections.Generic.List[object]]::new()
        foreach ($r in $rows) { [void]$script:IntuneManagerAllRows.Add($r) }


        $script:dgIntuneManagerObjects.ItemsSource = $script:IntuneManagerAllRows

    }

    Set-GraphPagesButtonStatus

    $policyTypes = Get-IntuneManagerSelectedPolicyTypes
    $allowDelete = Get-SettingValue 'AllowDelete'
    Set-IntuneManagerUIButtonStatus @('btnDelete') $policyTypes -ForceHide:($allowDelete -eq $false)
    Set-IntuneManagerUIButtonStatus @('btnImport','btnView','btnExport','btnCompare','btnCopy','btnDocument') $policyTypes
    if ($script:spIntuneManagerButtons) {
        $btnImport = (Get-AvaloniaHost)::FindByName($script:spIntuneManagerButtons, 'btnImport')
        if ($btnImport) { $btnImport.IsEnabled = $true }
    }


    Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox `
        $script:dgIntuneManagerObjects -ForceUpdate `
        -ObjectsCountTextBox $script:IntuneManagementObjectsCount


    Invoke-AppEvent "IntuneManagerUISelectedMenuItemChanged" $policyTypes


    Write-Status ""
}

function Add-IntuneManagerDataGridColumns
{
    if (-not $script:dgIntuneManagerObjects) { return }

    # WPF parity for the ObjectListFirstLineOnly setting: the 'firstLine' class
    # drives a MaxLines=1 cell style in Themes/Styles.axaml, so a multi-line
    # value (store app descriptions) no longer makes the row as tall as its text.
    if ((Get-SettingValue "ObjectListFirstLineOnly") -eq $true) {
        if (-not $script:dgIntuneManagerObjects.Classes.Contains('firstLine')) { $script:dgIntuneManagerObjects.Classes.Add('firstLine') }
    } else {
        [void]$script:dgIntuneManagerObjects.Classes.Remove('firstLine')
    }

    # Checkbox column for IsSelected. WPF parity: a CheckBox in the column
    # header toggles every row's IsSelected. Avalonia's DataGridCheckBoxColumn
    # has no bindable header CheckBox, so build a DataGridTemplateColumn:
    # header = a CheckBox control, cell template = a CheckBox bound TwoWay
    # to IsSelected. Initialize-AvaloniaGridSelectAllHeader walks the columns
    # post-add, finds the header CheckBox, and wires its IsCheckedChanged to
    # push IsSelected onto every row in ItemsSource.
    # Header and cell checkboxes are both LEFT-anchored: the Fluent column
    # header does not stretch its content, so a "centered" header checkbox is
    # actually pinned at the header's content-left edge - centering the cell
    # checkboxes can therefore never line up with it. Left-anchoring both
    # puts them on the same edge at any column width.
    $cellTemplateXaml = @'
<DataTemplate xmlns="https://github.com/avaloniaui" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay}" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="12,0,0,0"/>
</DataTemplate>
'@
    $headerCheckBox = New-Object Avalonia.Controls.CheckBox
    $headerCheckBox.IsChecked            = $false
    $headerCheckBox.HorizontalAlignment  = [Avalonia.Layout.HorizontalAlignment]::Left
    $headerCheckBox.VerticalAlignment    = [Avalonia.Layout.VerticalAlignment]::Center
    try { $headerCheckBox.SetValue([Avalonia.Controls.ToolTip]::TipProperty, 'Select/deselect all listed items') } catch { }

    $checkCol = New-Object Avalonia.Controls.DataGridTemplateColumn
    $checkCol.Header       = $headerCheckBox
    $checkCol.CellTemplate = (Get-AvaloniaHost)::LoadXaml($cellTemplateXaml)
    $checkCol.CanUserSort  = $false
    $script:dgIntuneManagerObjects.Columns.Add($checkCol) | Out-Null

    $selectAllHeaderCb = Initialize-AvaloniaGridSelectAllHeader -Grid $script:dgIntuneManagerObjects -BindingProperty 'IsSelected' -InitiallyChecked $false
    # Re-evaluate the action buttons after a select-all toggle: the generic
    # helper only flips row IsSelected, so without this the buttons kept the
    # state they had before every row got (un)checked. Module function by name
    # - closure-safe.
    if ($selectAllHeaderCb) {
        $selectAllHeaderCb.add_IsCheckedChanged((ConvertTo-AvaloniaEventScriptBlock {
            Update-IntuneManagerActionButtons
        }))
    }

    # Resolve the column list. The "ObjectColumns" per-type setting can name
    # paths beyond the canonical IntuneObjectRowItem fields (e.g.
    # `lastModifiedDateTime`, `Object.priority`); those bind through
    # IntunePolicyPathConverter against the row's Source so PSObject
    # NoteProperties resolve correctly (see [[avalonia-binding-needs-clr-types]]).
    $additionalColumns = @()
    $selectedType = Get-SelectedObjectTypeString
    if ($selectedType) {
        $additionalColsStr = Get-SettingStoreValue "IntuneManager\ObjectColumns\$selectedType" "$($script:IntuneManagerSelectedObject.Id)"
        if ($additionalColsStr) {
            $additionalColumns += $additionalColsStr.Split(',')
        }
    }

    $columns = @()
    if ($additionalColumns.Count -eq 0 -or $additionalColumns[0] -ne '0') {
        $columns += ?? $script:IntuneManagerSelectedObject.ViewProperties @('Name','PolicyName','ID')
    }
    foreach ($extra in $additionalColumns) {
        if ($extra -eq '0' -or $extra -eq '1') { continue }
        $columns += $extra
    }

    # Map lower-cased name -> the ACTUAL CLR property name. The lookup has to be
    # case-insensitive (view properties are declared as both "ID" and "Id"
    # across the policy classes) but the binding must use the real casing:
    # Avalonia's binder resolves paths by case-SENSITIVE reflection, so binding
    # "Id" against the property declared as "ID" silently rendered an empty
    # column - which is why Applications, Conditional Access and 11 other types
    # showed a blank Id in Avalonia but not in WPF.
    $rowClrProps = @{}
    foreach ($p in [IntuneObjectRowItem].GetProperties()) {
        $rowClrProps[$p.Name.ToLowerInvariant()] = $p.Name
    }
    # Construct via the host factory rather than [FullName]::new() — see the
    # comment on Host.CreatePathConverter for why the name-based resolution
    # blows up after Avalonia's XAML loader populates dynamic assemblies.
    $pathConverter = (Get-AvaloniaHost)::CreatePathConverter()

    foreach ($columnInfo in $columns) {
        if (-not $columnInfo) { continue }
        $bindingProp, $colHeader = ($columnInfo -split '=', 2)
        if (-not $colHeader) { $colHeader = $bindingProp.Split('.')[-1] }

        $col = New-Object Avalonia.Controls.DataGridTextColumn
        $col.Header = $colHeader
        $col.IsReadOnly = $true

        # Direct bind for the canonical row fields (Name, PolicyName, ID, ...);
        # everything else routes through the converter on Source so PSObject
        # NoteProperties get walked.
        #
        # CRITICAL: both bindings must be OneWay. Avalonia's
        # DataGridTextColumn.Binding defaults to TwoWay, and Avalonia's
        # binding engine fires writeback on sort / scroll / focus changes
        # even for IsReadOnly=$true columns. For the converter binding that
        # writeback would call IntunePolicyPathConverter.ConvertBack (which
        # returns null) and **stamp null into the row's Source property** —
        # silently breaking View / Compare / Copy because their handlers
        # read $selectedRow.Source. Verified by the 22-of-78 null-Source
        # diag the user produced 2026-06-07.
        $clrName = $rowClrProps[("$bindingProp").ToLowerInvariant()]
        if ($clrName) {
            # Bind the property's real name, not the caller's spelling.
            $directBinding = New-Object Avalonia.Data.Binding $clrName
            $directBinding.Mode = [Avalonia.Data.BindingMode]::OneWay
            $col.Binding = $directBinding
        } else {
            $b = New-Object Avalonia.Data.Binding 'Source'
            $b.Mode               = [Avalonia.Data.BindingMode]::OneWay
            $b.Converter          = $pathConverter
            $b.ConverterParameter = $bindingProp
            $col.Binding = $b
        }

        $script:dgIntuneManagerObjects.Columns.Add($col) | Out-Null
    }
}

function Set-IntuneManagerUIButtonStatus
{
    param($Buttons, $PolicyTypes, [switch]$ForceHide)

    if (-not $script:spIntuneManagerButtons) { return }

    foreach ($btnName in $Buttons) {
        $visible = $false
        if (-not $ForceHide) {
            foreach ($policyType in $PolicyTypes) {
                if (-not $policyType.ShowButtons -or
                    ($policyType.ShowButtons | Where-Object { $btnName -like "*$_" })) {
                    $visible = $true
                    break
                }
            }
        }
        $btn = (Get-AvaloniaHost)::FindByName($script:spIntuneManagerButtons, $btnName)
        if ($btn) { $btn.IsVisible = $visible }
    }
}

function Add-GraphPoliciesFromPaging
{
    param([ValidateSet('NextPage','AllRemainingPages')][string]$PagingType)

    if (-not $script:IntuneManagerSelectedObject -or -not $script:dgIntuneManagerObjects) { return }

    Write-Status "Loading $($script:IntuneManagerSelectedObject.Title) objects"

    $newRows = @(Get-GraphPolicies -Paging $PagingType | Get-GraphPolicyForUIList)
    if ($newRows.Count -gt 0) {
        # Rebuild into a NEW list instance rather than appending in place: the
        # no-filter path assigns this same collection back to ItemsSource, and
        # Avalonia ignores a reassignment of the identical instance - appended
        # rows never rendered.
        $merged = [System.Collections.Generic.List[object]]::new()
        foreach ($r in @($script:IntuneManagerAllRows)) { if ($r) { [void]$merged.Add($r) } }
        foreach ($r in $newRows) { [void]$merged.Add($r) }
        $script:IntuneManagerAllRows = $merged
    }

    Set-GraphPagesButtonStatus
    Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox `
        $script:dgIntuneManagerObjects -ForceUpdate `
        -ObjectsCountTextBox $script:IntuneManagementObjectsCount
    Write-Status ""
}

function Set-GraphPagesButtonStatus
{
    $ui = $script:UIProvider
    $viewObject = Get-SingletonObject "IntuneViewObject"
    if (-not $viewObject) { return }
    $panel = $viewObject.ViewPanel
    $visible = [bool]$script:GraphPagingCache
    $ui.SetXamlProperty($panel, 'btnLoadAllPages', 'IsVisible', $visible)
    $ui.SetXamlProperty($panel, 'btnLoadNextPage', 'IsVisible', $visible)
}

function Clear-GraphObjects
{
    $ui = $script:UIProvider
    $viewObject = Get-SingletonObject "IntuneViewObject"
    if (-not $viewObject) { return }
    $panel = $viewObject.ViewPanel

    $ui.SetXamlProperty($panel, 'txtFormTitle', 'Text', '')
    $ui.SetXamlProperty($panel, 'txtObjectsCount', 'Text', '')

    if ($script:dgIntuneManagerObjects) {
        $script:dgIntuneManagerObjects.Columns.Clear()
        $script:dgIntuneManagerObjects.ItemsSource = $null
    }
    $script:IntuneManagerAllRows = $null
    $script:GraphPagingCache = $null

    # Clear the filter box AFTER nulling AllRows so the TextChanged it raises
    # hits Invoke-FilterBoxChanged's early return. WPF clears the filter when
    # the grid's ItemsSource changes; without this, stale filter text silently
    # filtered the next policy type the user clicked.
    if ($script:IntuneManagementFilterTextBox) {
        $script:IntuneManagementFilterTextBox.Text = ''
    }
}

function Invoke-RefreshObjects
{
    if (-not $script:IntuneManagerSelectedObject) { return }

    $filterText = $null
    if ($script:IntuneManagementFilterTextBox) {
        $filterText = $script:IntuneManagementFilterTextBox.Text
    }

    Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject

    if ($filterText -and $script:IntuneManagementFilterTextBox) {
        $script:IntuneManagementFilterTextBox.Text = $filterText
        Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox `
            $script:dgIntuneManagerObjects -ForceUpdate `
            -ObjectsCountTextBox $script:IntuneManagementObjectsCount
    }

    Write-Status ""
}

function Invoke-FilterBoxChanged
{
    param($TxtBox, $DataSource, [switch]$ForceUpdate, $ObjectsCountTextBox)

    if (-not $DataSource) { return }
    if ($null -eq $script:IntuneManagerAllRows) {
        if ($ObjectsCountTextBox) { $ObjectsCountTextBox.Text = '' }
        return
    }

    $textValue   = if ($TxtBox) { [string]$TxtBox.Text } else { '' }
    $hasText     = -not [string]::IsNullOrEmpty($textValue)
    # Platform set lives at $script: so the gear's Filter Platforms submenu can
    # mutate it and re-call this with -ForceUpdate (mirrors the WPF filter).
    $platformSet = $script:_intunePlatformFilter
    $hasPlatform = ($null -ne $platformSet -and $platformSet.Count -gt 0)

    if ($hasText -or $hasPlatform) {
        $escaped = if ($hasText) { [regex]::Escape($textValue) } else { $null }
        $filtered = [System.Collections.Generic.List[object]]::new()
        foreach ($row in $script:IntuneManagerAllRows) {
            # Platform filter - empty Platform maps to the "" sentinel so a
            # "(no platform)" selection can target rows whose .Platform is
            # null/empty without ambiguity.
            if ($hasPlatform) {
                $plat = $row.Platform
                $key = if ([string]::IsNullOrEmpty($plat)) { "" } else { [string]$plat }
                if (-not $platformSet.Contains($key)) { continue }
            }
            if ($hasText) {
                $match = $false
                foreach ($prop in 'Name','PolicyName','Platform','LastModified','Created','ID','ScopeTags') {
                    $val = $row.$prop
                    if ($null -ne $val -and ($val -match $escaped)) { $match = $true; break }
                }
                if (-not $match) { continue }
            }
            [void]$filtered.Add($row)
        }
        $DataSource.ItemsSource = $filtered
    }
    else {
        $DataSource.ItemsSource = $script:IntuneManagerAllRows
    }

    if ($ObjectsCountTextBox) {
        $loadedCount  = $script:IntuneManagerAllRows.Count
        $visibleCount = if ($DataSource.ItemsSource) { @($DataSource.ItemsSource).Count } else { 0 }
        $morePages    = [bool]$script:GraphPagingCache
        $loadedLabel  = if ($morePages) { "$loadedCount+" } else { "$loadedCount" }

        if ($loadedCount -le 0) {
            $ObjectsCountTextBox.Text = ''
        }
        elseif ($visibleCount -lt $loadedCount) {
            $ObjectsCountTextBox.Text = "Showing $visibleCount of $loadedLabel"
        }
        else {
            $ObjectsCountTextBox.Text = if ($morePages) {
                "Objects: $loadedLabel (more available - click Load All)"
            } else {
                "Objects: $loadedCount"
            }
        }
    }
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
    if ((Get-SettingValue "ObjectViewType") -ne (Get-CacheObject "ObjectViewType-Cached")) {
        Set-CacheObject "ObjectViewType-Cached" (Get-SettingValue "ObjectViewType")
        Show-ViewMenu
    }
    # AllowDelete / AllowBulkDelete were only honoured while building the view
    # and the Bulk menu, so toggling them in Settings did nothing until the
    # next restart (WPF applies them immediately).
    Update-IntuneDeleteVisibility
}

function Invoke-IntuneUIEventNewAuthentication
{
    if ($script:ActiveView -and $script:ActiveView.Id -eq "IntuneManagement") {
        Update-IntuneDeleteVisibility

        # Access marking (Internal/AccessLevel.ps1) needs the signed-in token,
        # which only exists once login completes. The nav was built white before
        # the token arrived, so rebuild it now to repaint the AccessType colours -
        # otherwise the menu stays uncoloured until the next view switch. A plain
        # Show-ViewMenu blanked the loaded view because setting ItemsSource fires
        # the selection-changed event with a null selection; wrap it in the same
        # suppress+restore guard the count-label refresh uses so the rebuild
        # repaints without dropping the loaded list, then refresh objects.
        if (Get-Command Show-ViewMenu -ErrorAction SilentlyContinue) {
            $script:_suppressMenuSelectionEvents = $true
            try {
                # Restore by identity, not by reference. Show-ViewMenu binds FRESH
                # rows - Get-IntuneViewItems wraps every type in a new ViewMenuItem
                # on each call - so the old SelectedItem is never in the new list,
                # and assigning it back left SelectedIndex at -1: the nav lost its
                # selection at every login even though the same group was still
                # there. (The count-label refresh can restore by reference because
                # it re-binds the very same row objects.)
                $sel = if ($script:lstMenuItems) { $script:lstMenuItems.SelectedItem } else { $null }
                Show-ViewMenu
                if ($sel -and $script:lstMenuItems) {
                    $match = $null
                    foreach ($row in @($script:lstMenuItems.ItemsSource)) {
                        if ($row -and -not $row.IsHeader -and $row.Id -eq $sel.Id -and $row.Title -eq $sel.Title) { $match = $row; break }
                    }
                    if ($match) { $script:lstMenuItems.SelectedItem = $match }
                }
            }
            finally {
                $script:_suppressMenuSelectionEvents = $false
            }
        }

        Invoke-RefreshObjects
    }
}

function Invoke-IntuneUIEventUserDisconnected
{
    if ($script:ActiveView -and $script:ActiveView.Id -eq "IntuneManagement") {
        Update-IntuneDeleteVisibility

        # Mirror of the login handler: signed out there is no token, so
        # Update-IntuneAccessLevels (inside Show-ViewMenu) resets every type/group
        # to Full and clears the orange/red marking. AccessType has no change
        # notification, so rebuild the nav to repaint the default colours. Do NOT
        # use the suppress+restore guard here: letting the selection fall to null
        # fires OnItemChanged($null), which clears the object list the signed-out
        # session can no longer load.
        if (Get-Command Show-ViewMenu -ErrorAction SilentlyContinue) {
            Show-ViewMenu
        }
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
    if (-not (Test-DefaultTokenExpired)) { return }
    Invoke-IntuneUIEventUserDisconnected $TokenInfo
}

# Live-apply the delete-permission settings (WPF parity: the same trio lives
# in UI/WPF/Extensions/IntuneManagerUIWPF.ps1).
function Update-IntuneDeleteVisibility
{
    Update-IntuneManagerDeleteButton
    Update-IntuneBulkDeleteMenu
}

function Update-IntuneManagerDeleteButton
{
    if (-not $script:spIntuneManagerButtons) { return }

    $allowDelete = (Get-SettingValue 'AllowDelete') -eq $true
    # Route through the normal status path so the selected policy type's
    # ShowButtons rules still apply on top of the permission setting.
    $policyTypes = @(Get-IntuneManagerSelectedPolicyTypes)
    Set-IntuneManagerUIButtonStatus @('btnDelete') $policyTypes -ForceHide:(-not $allowDelete)
}

function Update-IntuneBulkDeleteMenu
{
    if (-not $script:mnuMain) { return }
    $bulkMenu = $script:mnuMain.Items | Where-Object { $_.Name -eq 'IntuneBulk' } | Select-Object -First 1
    if (-not $bulkMenu) { return }
    $deleteItem = $bulkMenu.Items | Where-Object { $_.Name -eq 'mnuBulkDelete' } | Select-Object -First 1
    if (-not $deleteItem) { return }

    # Avalonia uses a bool IsVisible, not WPF's Visibility enum.
    $deleteItem.IsVisible = ((Get-SettingValue 'AllowBulkDelete') -eq $true)
}

# Avalonia port of UI/WPF/Extensions/IntuneManagerUIWPF.ps1
# Update-MenuTitleConfigForIntuneView. Show-ViewMenu (CoreUIAvalonia) calls
# this whenever the Intune view becomes active or its menu is rebuilt.
# Reveals the gear button next to the menu title and syncs / wires the
# Group / API / Show item counts / Filter platforms / Refresh items in its
# MenuFlyout (full WPF parity).
function Update-AvaloniaMenuTitleConfigForIntuneView
{
    if (-not $script:btnMenuTitleConfig) { return }
    if (-not $script:Window) { return }

    $hostType = Get-AvaloniaHost
    $script:btnMenuTitleConfig.IsVisible = $true
    try { $script:btnMenuTitleConfig.SetValue([Avalonia.Controls.ToolTip]::TipProperty, 'Menu options: switch view, refresh') } catch { }

    $script:mnuMenuTitleConfigGroup           = $hostType::FindByName($script:Window, 'mnuMenuTitleConfigGroup')
    $script:mnuMenuTitleConfigType            = $hostType::FindByName($script:Window, 'mnuMenuTitleConfigType')
    $script:mnuMenuTitleConfigShowCounts      = $hostType::FindByName($script:Window, 'mnuMenuTitleConfigShowCounts')
    $script:mnuMenuTitleConfigFilterPlatforms = $hostType::FindByName($script:Window, 'mnuMenuTitleConfigFilterPlatforms')
    $script:mnuMenuTitleConfigRefresh         = $hostType::FindByName($script:Window, 'mnuMenuTitleConfigRefresh')
    $script:mnuMenuTitleConfigShowInfo        = $hostType::FindByName($script:Window, 'mnuMenuTitleConfigShowInfo')

    $current = Get-SettingValue 'ObjectViewType'
    # "Group" is the default; anything not explicitly "Type" reads as Group
    # so a missing/blank setting still ticks Group rather than leaving both
    # unchecked.
    if ($script:mnuMenuTitleConfigGroup) { $script:mnuMenuTitleConfigGroup.IsChecked = ($current -ne 'Type') }
    if ($script:mnuMenuTitleConfigType)  { $script:mnuMenuTitleConfigType.IsChecked  = ($current -eq 'Type') }

    # Read-only "Intune Info" group is opt-in (hidden by default). Sync the toggle.
    if ($script:mnuMenuTitleConfigShowInfo) { $script:mnuMenuTitleConfigShowInfo.IsChecked = ((Get-SettingStoreValue "IntuneManager" "ShowIntuneInfo" "false") -eq "true") }

    # "Show item counts" is HIDDEN per user decision (2026-06-12) — the menu
    # item is IsVisible=False in MainWindow.axaml and the populate paths below
    # are forced off so a previously-saved true setting can't resurrect it.
    # The counts machinery is kept for potential re-enable.
    $showCounts = $false
    if ($script:mnuMenuTitleConfigShowCounts) { $script:mnuMenuTitleConfigShowCounts.IsChecked = $showCounts }

    if (-not $script:_menuTitleConfigWired) {
        $script:_menuTitleConfigWired = $true

        if ($script:mnuMenuTitleConfigGroup) {
            $script:mnuMenuTitleConfigGroup.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                Set-AvaloniaIntuneMenuObjectViewType 'Group'
            }))
        }
        if ($script:mnuMenuTitleConfigType) {
            $script:mnuMenuTitleConfigType.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                Set-AvaloniaIntuneMenuObjectViewType 'Type'
            }))
        }
        if ($script:mnuMenuTitleConfigShowInfo) {
            $script:mnuMenuTitleConfigShowInfo.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                # ToggleType=CheckBox flips IsChecked before the handler runs. Persist +
                # rebuild the nav so the read-only Info group appears/disappears.
                $enabled = [bool]$script:mnuMenuTitleConfigShowInfo.IsChecked
                Save-SettingStoreValue -SubPath "IntuneManager" -Key "ShowIntuneInfo" -Value ($enabled.ToString().ToLower())
                Show-ViewMenu
            }))
        }
        if ($script:mnuMenuTitleConfigRefresh) {
            $script:mnuMenuTitleConfigRefresh.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                # Re-bind the menu list (cheap - no Graph). Count re-fetch
                # removed with the hidden "Show item counts" feature.
                Show-ViewMenu
            }))
        }
        if ($script:mnuMenuTitleConfigShowCounts) {
            $script:mnuMenuTitleConfigShowCounts.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                # ToggleType=CheckBox items flip IsChecked before the handler runs.
                $enabled = [bool]$script:mnuMenuTitleConfigShowCounts.IsChecked
                Save-SettingStoreValue -SubPath "IntuneManager" -Key "ShowMenuItemCounts" -Value ($enabled.ToString().ToLower())
                if ($enabled) {
                    Update-AvaloniaIntuneViewItemCounts
                }
                else {
                    # Strip counts back to bare Title without a fresh batch.
                    Update-AvaloniaIntuneMenuLabels -Reset
                }
            }))
        }
        if ($script:mnuMenuTitleConfigFilterPlatforms -and $script:btnMenuTitleConfig.Flyout) {
            # Populate on the parent flyout's Opened (the submenu items must
            # exist BEFORE the user hovers into the submenu). Cheap: in-memory
            # scan of the already-bound right-pane rows, no Graph.
            $script:btnMenuTitleConfig.Flyout.add_Opened((ConvertTo-AvaloniaEventScriptBlock {
                Update-AvaloniaIntuneFilterPlatformsSubmenu
            }))
        }
    }

    # First paint after rebind: if counts are on, populate them. Uses the
    # session count cache so a view-switch back doesn't trigger a fresh
    # batch - only first activation or explicit Refresh does.
    if ($showCounts) {
        Update-AvaloniaIntuneViewItemCounts
    }
}

# Project the shared per-type counts (Get-IntuneTypeCounts) onto the bound
# [ViewMenuItem] rows. Type rows: direct lookup. Group rows: sum of every
# child type's count. ViewMenuItem has no INPC so the labels only repaint
# after an ItemsSource rebind; selection is preserved with the suppression
# flag so the rebind can't retrigger a Graph reload of the current type.
function Update-AvaloniaIntuneViewItemCounts
{
    param([switch]$Force)

    if (-not $script:lstMenuItems -or -not $script:lstMenuItems.ItemsSource) { return }

    $tokenId = $null
    try { $tokenId = Get-DefaultTokenId } catch { }
    if ($null -eq $tokenId) {
        Write-LogDebug "Update-AvaloniaIntuneViewItemCounts: no token id available - skipping"
        return
    }

    $counts = Get-IntuneTypeCounts -Force:$Force -TokenId $tokenId
    Update-AvaloniaIntuneMenuLabels -Counts $counts
}

# Rewrites MenuLabel on every bound row ("Title (N)" or bare Title with
# -Reset) and rebinds the ListBox so the non-INPC rows repaint.
function Update-AvaloniaIntuneMenuLabels
{
    param([hashtable]$Counts, [switch]$Reset)

    if (-not $script:lstMenuItems -or -not $script:lstMenuItems.ItemsSource) { return }
    $rows = @($script:lstMenuItems.ItemsSource)

    foreach ($row in $rows) {
        if (-not $row -or $row.IsHeader) { continue }
        if ($Reset -or -not $Counts) {
            $row.MenuLabel = [string]$row.Title
            continue
        }
        $payload = $row.Tag
        $count = $null
        if ($payload -is [IntunePolicyTypeBase]) {
            if ($Counts.ContainsKey($payload.Id)) { $count = $Counts[$payload.Id] }
        }
        elseif ($payload -is [IntunePolicyGroupBase]) {
            $sum = 0; $hasAny = $false
            foreach ($child in @($payload.PolicyTypes)) {
                if ($child -and $Counts.ContainsKey($child.Id)) {
                    $sum += $Counts[$child.Id]; $hasAny = $true
                }
            }
            if ($hasAny) { $count = $sum }
        }
        $row.MenuLabel = if ($null -ne $count) { "$($row.Title) ($count)" } else { [string]$row.Title }
    }

    $script:_suppressMenuSelectionEvents = $true
    try {
        $sel = $script:lstMenuItems.SelectedItem
        $script:lstMenuItems.ItemsSource = $null
        $script:lstMenuItems.ItemsSource = $rows
        if ($sel) { $script:lstMenuItems.SelectedItem = $sel }
    }
    finally {
        $script:_suppressMenuSelectionEvents = $false
    }
}

# Rebuilds the Filter Platforms submenu from the currently-bound right-pane
# rows. Mirrors the WPF implementation: one checkable item per platform,
# "(no platform)" pinned at the bottom, Clear filter footer when active.
function Update-AvaloniaIntuneFilterPlatformsSubmenu
{
    if (-not $script:mnuMenuTitleConfigFilterPlatforms) { return }

    # 1. Collect platforms in the current grid (the unfiltered backing list).
    $available = [System.Collections.Generic.HashSet[string]]::new()
    $hasNull   = $false
    if ($script:IntuneManagerAllRows) {
        foreach ($it in $script:IntuneManagerAllRows) {
            $plat = $it.Platform
            if ([string]::IsNullOrEmpty($plat)) { $hasNull = $true }
            else { [void]$available.Add([string]$plat) }
        }
    }

    # 2. Include currently-selected filters so the user can untick them even
    #    if the new view no longer has any matching rows.
    if ($script:_intunePlatformFilter) {
        foreach ($s in $script:_intunePlatformFilter) {
            if ($s -eq "") { $hasNull = $true }
            else { [void]$available.Add($s) }
        }
    }

    $script:mnuMenuTitleConfigFilterPlatforms.Items.Clear()

    $sorted = @($available | Sort-Object)
    if ($sorted.Count -eq 0 -and -not $hasNull) {
        $empty = New-Object Avalonia.Controls.MenuItem
        $empty.Header    = "(no policies loaded)"
        $empty.IsEnabled = $false
        [void]$script:mnuMenuTitleConfigFilterPlatforms.Items.Add($empty)
        return
    }

    # 3. One MenuItem per platform; "(no platform)" pinned at the bottom.
    $platformClick = (ConvertTo-AvaloniaEventScriptBlock {
        param($s, $e)
        Switch-AvaloniaIntunePlatformFilter ([string]$s.Tag) ([bool]$s.IsChecked)
    })
    foreach ($plat in ($sorted + $(if ($hasNull) { @($null) } else { @() }))) {
        $mi = New-Object Avalonia.Controls.MenuItem
        $mi.Header     = if ($null -eq $plat) { "(no platform)" } else { $plat }
        $mi.ToggleType = [Avalonia.Controls.MenuItemToggleType]::CheckBox
        # Tag stores the canonical key ("" sentinel for the no-platform row).
        $mi.Tag        = if ($null -eq $plat) { "" } else { $plat }
        $mi.IsChecked  = ($script:_intunePlatformFilter -and $script:_intunePlatformFilter.Contains([string]$mi.Tag))
        $mi.add_Click($platformClick)
        [void]$script:mnuMenuTitleConfigFilterPlatforms.Items.Add($mi)
    }

    # 4. "Clear filter" footer when any filter is active.
    if ($script:_intunePlatformFilter -and $script:_intunePlatformFilter.Count -gt 0) {
        [void]$script:mnuMenuTitleConfigFilterPlatforms.Items.Add((New-Object Avalonia.Controls.Separator))
        $clear = New-Object Avalonia.Controls.MenuItem
        $clear.Header = "Clear filter"
        $clear.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            if ($script:_intunePlatformFilter) { $script:_intunePlatformFilter.Clear() }
            Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox `
                $script:dgIntuneManagerObjects -ForceUpdate `
                -ObjectsCountTextBox $script:IntuneManagementObjectsCount
        }))
        [void]$script:mnuMenuTitleConfigFilterPlatforms.Items.Add($clear)
    }
}

# Toggle a single platform on/off in $script:_intunePlatformFilter and
# re-evaluate the grid filter.
function Switch-AvaloniaIntunePlatformFilter
{
    param([string]$PlatformKey, [bool]$Enable)

    if ($null -eq $script:_intunePlatformFilter) {
        $script:_intunePlatformFilter = [System.Collections.Generic.HashSet[string]]::new()
    }
    if ($Enable) { [void]$script:_intunePlatformFilter.Add($PlatformKey) }
    else         { [void]$script:_intunePlatformFilter.Remove($PlatformKey) }

    Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox `
        $script:dgIntuneManagerObjects -ForceUpdate `
        -ObjectsCountTextBox $script:IntuneManagementObjectsCount
}

function Set-AvaloniaIntuneMenuObjectViewType
{
    param([ValidateSet('Group','Type')][string]$ViewType)

    $current = Get-SettingValue 'ObjectViewType'
    if ($current -eq $ViewType) {
        # ToggleType=CheckBox items flip IsChecked before the handler fires.
        # Re-sync so the user can't toggle one off without picking the other.
        if ($script:mnuMenuTitleConfigGroup) { $script:mnuMenuTitleConfigGroup.IsChecked = ($ViewType -eq 'Group') }
        if ($script:mnuMenuTitleConfigType)  { $script:mnuMenuTitleConfigType.IsChecked  = ($ViewType -eq 'Type') }
        return
    }

    Save-SettingStoreValue -SubPath 'IntuneManager' -Key 'ObjectViewType' -Value $ViewType
    Set-CacheObject 'ObjectViewType-Cached' $ViewType
    if ($script:mnuMenuTitleConfigGroup) { $script:mnuMenuTitleConfigGroup.IsChecked = ($ViewType -eq 'Group') }
    if ($script:mnuMenuTitleConfigType)  { $script:mnuMenuTitleConfigType.IsChecked  = ($ViewType -eq 'Type') }

    Show-ViewMenu
}
