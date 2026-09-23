# Object picker sub-modal (Avalonia).
#
# Split out of CoreUIAvalonia.ps1 so the shell file stays inside its R9 line
# budget. Avalonia port of the WPF Show-PickerDialog / Show-ObjectPickerDialog
# pair: a Window-shaped dialog shown through Host::ShowDialog for synchronous
# semantics, so callers can write
#     $selected = Show-ObjectPickerDialog ...; if ($selected) { ... }
# A grdModal overlay would have to push a DispatcherFrame from PowerShell to
# block, which the C# side already wraps.

# Show a one-line message in the picker's status area (hidden when empty).
function Set-PickerStatusText
{
    param([string]$Text)

    if (-not $script:txtPickerStatus -or -not $script:grpPickerStatus) { return }

    $script:txtPickerStatus.Text = $Text
    $script:grpPickerStatus.IsVisible = (-not [String]::IsNullOrWhiteSpace($Text))
}

# The scope object behind the current "Search in" selection, or $null. The
# ComboBox holds [SettingsListItem] projections (Avalonia cannot bind
# NoteProperties), so the real scope comes back through $script:pickerScopeMap.
function Get-PickerSelectedScope
{
    if (-not $script:cbPickerScope -or -not $script:pickerScopeMap) { return $null }

    $selected = $script:cbPickerScope.SelectedItem
    if (-not $selected) { return $null }

    return $script:pickerScopeMap[[string]$selected.Value]
}

# The tenant entry behind the current "Tenant" selection, or $null when the
# dialog is single-tenant. Same projection-and-map trick as the scope combo.
function Get-PickerSelectedTenant
{
    if (-not $script:cbPickerTenant -or -not $script:pickerTenantMap) { return $null }

    $selected = $script:cbPickerTenant.SelectedItem
    if (-not $selected) { return $null }

    return $script:pickerTenantMap[[string]$selected.Value]
}

# Drop whatever is in the results grid. Called when the tenant changes - those
# rows belong to the previous tenant and picking one would compare the wrong
# object.
function Clear-PickerResults
{
    $empty = [System.Collections.Generic.List[PickerObjectRow]]::new()
    if ($script:_pickerState) { $script:_pickerState.Rows = $empty }
    if ($script:dgPickerObjects) { $script:dgPickerObjects.ItemsSource = $empty }
    if ($script:btnPickerOK) { $script:btnPickerOK.IsEnabled = $false }
}

# Wired to the tenant combo's SelectionChanged. A module function rather than an
# inline handler so it can be reached by name from the rebound scriptblock.
function Invoke-PickerTenantChanged
{
    Clear-PickerResults

    $tenant = Get-PickerSelectedTenant
    if ($tenant) { Set-PickerStatusText "Search $($tenant.Title) for an object." }
}

function Show-PickerDialog
{
    param(
        $Title,
        $InitHandler,
        $ClickHandler,
        $OkHandler,
        [switch] $MultiSelect,
        $ClickHandlerArgs = $null,
        $LoadHandler = $null,
        [string] $LoadLabel = "Load all from Intune",
        # Entries from Get-PolicySearchScopes. When supplied, the dialog shows
        # a "Search in" dropdown above the search box.
        $Scopes = $null,
        [string] $SelectedScopeKey = $null,
        # Entries from Get-PolicySearchTenants. When supplied, the dialog shows
        # a "Tenant" dropdown above the scope dropdown.
        $Tenants = $null,
        [string] $SelectedTenantKey = $null
    )

    $ui = $script:UIProvider
    $dialogXaml = Join-Path $script:AppUIRootFolder 'XAML/PickerDialog.axaml'
    $dialog = $ui.GetXamlObject($dialogXaml)
    if (-not $dialog) { return }

    $dialog.Title = $Title

    $hostType            = (Get-AvaloniaHost)
    $script:pickerDialog       = $dialog
    $script:dgPickerObjects    = $hostType::FindByName($dialog, 'dgPickerObjects')
    $script:btnPickerOK        = $hostType::FindByName($dialog, 'btnPickerOK')
    $script:grpPickerStatus    = $hostType::FindByName($dialog, 'grpPickerStatus')
    $script:txtPickerStatus    = $hostType::FindByName($dialog, 'txtPickerStatus')
    $script:txtPickerSearch    = $hostType::FindByName($dialog, 'txtPickerSearch')
    $btnPickerSearch           = $hostType::FindByName($dialog, 'btnPickerSearch')

    if ($MultiSelect) {
        $script:dgPickerObjects.SelectionMode = [Avalonia.Controls.DataGridSelectionMode]::Extended
    }

    # Tenant selector - only shown when more than one tenant is signed in.
    $script:cbPickerTenant  = $hostType::FindByName($dialog, 'cbPickerTenant')
    $script:pickerTenantMap = $null
    if ($Tenants -and $script:cbPickerTenant) {
        $script:pickerTenantMap = @{}
        $tenantItems = [System.Collections.Generic.List[SettingsListItem]]::new()
        foreach ($tenant in @($Tenants)) {
            $script:pickerTenantMap[[string]$tenant.Key] = $tenant
            [void]$tenantItems.Add([SettingsListItem]@{ Name = $tenant.Title; Value = $tenant.Key })
        }
        $script:cbPickerTenant.ItemsSource = $tenantItems

        $tenantIndex = 0
        if ($SelectedTenantKey) {
            for ($i = 0; $i -lt $tenantItems.Count; $i++) {
                if ($tenantItems[$i].Value -eq $SelectedTenantKey) { $tenantIndex = $i; break }
            }
        }
        if ($tenantItems.Count -gt 0) { $script:cbPickerTenant.SelectedIndex = $tenantIndex }

        $grdPickerTenant = $hostType::FindByName($dialog, 'grdPickerTenant')
        if ($grdPickerTenant) { $grdPickerTenant.IsVisible = $true }

        # Switching tenant clears the grid: the rows in it came from the
        # previous tenant and the search has to be re-run against the new one.
        $script:cbPickerTenant.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
            Invoke-PickerTenantChanged
        }))
    }

    $script:cbPickerScope  = $hostType::FindByName($dialog, 'cbPickerScope')
    $script:pickerScopeMap = $null
    if ($Scopes -and $script:cbPickerScope) {
        $script:pickerScopeMap = @{}
        $scopeItems = [System.Collections.Generic.List[SettingsListItem]]::new()
        foreach ($scope in @($Scopes)) {
            $script:pickerScopeMap[[string]$scope.Key] = $scope
            [void]$scopeItems.Add([SettingsListItem]@{ Name = $scope.Title; Value = $scope.Key })
        }
        $script:cbPickerScope.ItemsSource = $scopeItems

        $selectedIndex = 0
        if ($SelectedScopeKey) {
            for ($i = 0; $i -lt $scopeItems.Count; $i++) {
                if ($scopeItems[$i].Value -eq $SelectedScopeKey) { $selectedIndex = $i; break }
            }
        }
        if ($scopeItems.Count -gt 0) { $script:cbPickerScope.SelectedIndex = $selectedIndex }

        $grdPickerScope = $hostType::FindByName($dialog, 'grdPickerScope')
        if ($grdPickerScope) { $grdPickerScope.IsVisible = $true }
    }

    $ui.AddXamlEvent($dialog, 'btnPickerCancel', 'add_Click', ((ConvertTo-AvaloniaEventScriptBlock {
        $script:pickerDialog.Close()
    }.GetNewClosure())))

    $script:pickerOKHandler = $OkHandler
    $ui.AddXamlEvent($dialog, 'btnPickerOK', 'add_Click', $OkHandler)

    $script:pickerSearchHandler = $ClickHandler
    if ($ClickHandler) {
        $ui.AddXamlEvent($dialog, 'btnPickerSearch', 'add_Click', $ClickHandler)
    }

    if ($ClickHandlerArgs -and $btnPickerSearch) {
        $btnPickerSearch.Tag = $ClickHandlerArgs
    }

    $btnPickerLoadAll = $hostType::FindByName($dialog, 'btnPickerLoadAll')
    if ($btnPickerLoadAll -and $LoadHandler -is [scriptblock]) {
        $btnPickerLoadAll.Content   = $LoadLabel
        $btnPickerLoadAll.IsVisible = $true
        $script:pickerLoadHandler   = $LoadHandler
        $btnPickerLoadAll.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            try {
                # Disabled only for the duration - the scope can change, so a
                # second load is legitimate.
                $S.IsEnabled = $false
                & $script:pickerLoadHandler
            } catch {
                Write-LogError "Picker LoadHandler failed" $_.Exception
            } finally {
                $S.IsEnabled = $true
            }
        }.GetNewClosure()))
    }

    $script:dgPickerObjects.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
        $script:btnPickerOK.IsEnabled = ($null -ne $script:dgPickerObjects.SelectedItem)
    }))

    $script:dgPickerObjects.add_DoubleTapped((ConvertTo-AvaloniaEventScriptBlock {
        if ($script:btnPickerOK.IsEnabled -and $script:pickerOKHandler) {
            & $script:pickerOKHandler
        }
    }))

    # Enter in the search box runs the search (matches WPF). Handled is set so
    # it does not fall through to btnPickerOK, which is IsDefault.
    $script:txtPickerSearch.add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
        param($S, $E)
        if ($E.Key -eq [Avalonia.Input.Key]::Enter -and $script:pickerSearchHandler) {
            $E.Handled = $true
            & $script:pickerSearchHandler
        }
    }))

    if ($InitHandler -is [scriptblock]) {
        & $InitHandler
    }

    $hostType::ShowDialog($script:pickerDialog, $script:Window)
}

# Project a policy object into the CLR row shape the picker DataGrid binds
# against (Avalonia ignores PSObject NoteProperties - see
# [[avalonia-binding-needs-clr-types]]). Module functions rather than
# scriptblocks because the handlers below cannot capture one - rebinding a
# scriptblock to the module strips its closure.
function ConvertTo-PickerObjectRow
{
    param($Source)

    if (-not $Source) { return $null }

    return [PickerObjectRow]@{
        Name            = [string]$Source.Name
        PolicyTypeTitle = if ($Source.PolicyType) { [string]$Source.PolicyType.Title } else { $null }
        Description     = [string]$Source.Description
        Source          = $Source
    }
}

function ConvertTo-PickerObjectRowList
{
    param($Items)

    $rows = [System.Collections.Generic.List[PickerObjectRow]]::new()
    foreach ($item in @($Items)) {
        $row = ConvertTo-PickerObjectRow $item
        if ($row) { [void]$rows.Add($row) }
    }
    return ,$rows
}

function Show-ObjectPickerDialog
{
    param(
        [string]      $Title,
        [object[]]    $Items,
        [object[]]    $DisplayColumns,
        [scriptblock] $LoadHandler,
        [string]      $LoadLabel = "Load all from Intune",
        # Entries from Get-PolicySearchScopes. When supplied, the dialog shows
        # a "Search in" dropdown above the search box.
        $Scopes = $null,
        [string]      $SelectedScopeKey = $null,
        # Entries from Get-PolicySearchTenants. When supplied, the dialog shows
        # a "Tenant" dropdown above the scope dropdown.
        $Tenants = $null,
        [string]      $SelectedTenantKey = $null,
        # Called as & $SearchHandler $searchText $scope and expected to return
        # the objects to display. When absent the Search button filters the
        # supplied -Items client-side (the original behaviour).
        [scriptblock] $SearchHandler = $null
    )

    # Everything the handlers need lives on $script:_pickerState, NOT in
    # captured locals: ConvertTo-AvaloniaEventScriptBlock rebinds each block to
    # the module, which strips the closure (proven 2026-08-27 with a real
    # function local raised through a real Avalonia Click). Module-scope state,
    # module functions by name, $S/$E and sender.Tag are what survive.
    #
    # DisplayColumns Binding paths are mapped to PickerObjectRow property names;
    # only the paths used by current callers (Name, PolicyType.Title,
    # Description) are recognised - extend PickerObjectRow and BindingMap
    # together if a new column is needed.
    $script:_pickerState = @{
        Columns    = $DisplayColumns
        BindingMap = @{
            'Name'             = 'Name'
            'PolicyType.Title' = 'PolicyTypeTitle'
            'Description'      = 'Description'
        }
        Rows          = (ConvertTo-PickerObjectRowList $Items)
        Result        = @{ Value = $null }
        SearchHandler = $SearchHandler
        LoadHandler   = $LoadHandler
    }

    $initHandler = (ConvertTo-AvaloniaEventScriptBlock {
        $ps = $script:_pickerState
        foreach ($colDef in $ps.Columns) {
            $col            = [Avalonia.Controls.DataGridTextColumn]::new()
            $col.Header     = $colDef.Header
            $col.IsReadOnly = $true

            $bindingPath = $ps.BindingMap[[string]$colDef.Binding]
            if (-not $bindingPath) {
                Write-LogError "Show-ObjectPickerDialog: unsupported column binding '$($colDef.Binding)' - extend PickerObjectRow + BindingMap together"
                continue
            }
            $col.Binding = [Avalonia.Data.Binding]::new($bindingPath)
            $col.Width   = [Avalonia.Controls.DataGridLength]::new(1, [Avalonia.Controls.DataGridLengthUnitType]::Star)

            $script:dgPickerObjects.Columns.Add($col) | Out-Null
        }

        $script:dgPickerObjects.ItemsSource = $ps.Rows
    })

    $clickHandler = (ConvertTo-AvaloniaEventScriptBlock {
        $ps = $script:_pickerState
        $filter = [string]$script:txtPickerSearch.Text

        if ($ps.SearchHandler) {
            # Server-side search: ask the caller for the objects matching the
            # text in the selected scope and show exactly what comes back. The
            # status line belongs to the search handler - it knows whether an
            # empty result means "no matches" or "keep typing".
            $scope = Get-PickerSelectedScope
            try {
                $found = @(& $ps.SearchHandler $filter $scope)
            }
            catch {
                Write-LogError "Picker SearchHandler failed" $_.Exception
                $found = @()
            }

            $ps.Rows = (ConvertTo-PickerObjectRowList $found)
            $script:dgPickerObjects.ItemsSource = $ps.Rows
            return
        }

        if ([string]::IsNullOrEmpty($filter)) {
            $script:dgPickerObjects.ItemsSource = $ps.Rows
        } else {
            # Where-Object stamps PSObject ETS wrappers; Avalonia DataGrid's
            # text-column binder reflects on the runtime type and renders blank
            # cells against the wrapper (see [[avalonia-binding-needs-clr-types]]).
            $filtered = [System.Collections.Generic.List[PickerObjectRow]]::new()
            foreach ($r in $ps.Rows) {
                if ($r.Name -ilike "*$filter*") { [void]$filtered.Add($r) }
            }
            $script:dgPickerObjects.ItemsSource = $filtered
        }
    })

    $okHandler = (ConvertTo-AvaloniaEventScriptBlock {
        $row = $script:dgPickerObjects.SelectedItem
        if ($row) {
            $script:_pickerState.Result.Value = $row.Source
        }
        $script:pickerDialog.Close()
    })

    $wrappedLoadHandler = $null
    if ($LoadHandler) {
        $wrappedLoadHandler = (ConvertTo-AvaloniaEventScriptBlock {
            $ps = $script:_pickerState
            $newItems = & $ps.LoadHandler

            if ($ps.SearchHandler) {
                # Scoped loads replace rather than merge - the returned set IS
                # everything in the newly selected scope, and merging would
                # leave stale rows from a previous scope in the grid.
                $ps.Rows = (ConvertTo-PickerObjectRowList $newItems)
                $script:dgPickerObjects.ItemsSource = $ps.Rows
                return
            }

            if ($newItems) {
                # Dedupe by Id; new entries appended so cached items stay first.
                $existingIds = @{}
                foreach ($r in $ps.Rows) {
                    if ($r.Source -and $r.Source.Id) { $existingIds[$r.Source.Id] = $true }
                }
                $merged = [System.Collections.Generic.List[PickerObjectRow]]::new()
                foreach ($r in $ps.Rows) { [void]$merged.Add($r) }
                foreach ($n in $newItems) {
                    if (-not $n) { continue }
                    if ($n.Id -and $existingIds.ContainsKey($n.Id)) { continue }
                    $row = ConvertTo-PickerObjectRow $n
                    if ($row) { [void]$merged.Add($row) }
                }
                $ps.Rows = $merged
            }
            Invoke-PickerSearchHandler
        })
    }

    try {
        Show-PickerDialog -Title $Title -InitHandler $initHandler -ClickHandler $clickHandler -OkHandler $okHandler -LoadHandler $wrappedLoadHandler -LoadLabel $LoadLabel -Scopes $Scopes -SelectedScopeKey $SelectedScopeKey -Tenants $Tenants -SelectedTenantKey $SelectedTenantKey
        return $script:_pickerState.Result.Value
    }
    finally {
        # Do not keep the handlers / rows alive between dialog invocations.
        $script:_pickerState = $null
    }
}

# Re-run the picker's current search/filter. A module function so the load
# handler can trigger it without capturing the click handler.
function Invoke-PickerSearchHandler
{
    if ($script:pickerSearchHandler) { & $script:pickerSearchHandler }
}
