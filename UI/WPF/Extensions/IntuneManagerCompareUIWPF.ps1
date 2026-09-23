
Add-AppEventHandler "CompareColumnVisibilityChanged" "Set-CompareGridColumnVisibility"

function Set-CompareGridColumnVisibility
{
    param($ShowCategory = $false, $ShowSubCategory = $false)

    if(-not $script:dgCompareInfo) { return }

    foreach($col in $script:dgCompareInfo.Columns)
    {
        $Path = $col.Binding.Path.Path
        if($Path -eq "Category")    { $col.Visibility = if($ShowCategory)    { "Visible" } else { "Collapsed" } }
        if($Path -eq "SubCategory") { $col.Visibility = if($ShowSubCategory) { "Visible" } else { "Collapsed" } }
    }
}

function Set-CompareRuntimeOptionsFromUI
{
    $params = @{}

    if($script:cbCompareType)
    {
        $params.CompareType = $script:cbCompareType.SelectedValue
        $params.CompareDefinition = $script:cbCompareType.SelectedItem
    }

    if($script:chkIgnoreCoreProperties)
    {
        $params.IgnoreCoreProperties = [bool]$script:chkIgnoreCoreProperties.IsChecked
    }

    if($script:cbCompareSave)
    {
        $params.SaveType = $script:cbCompareSave.SelectedValue
    }

    if($script:cbCompareOutputFormat)
    {
        $params.OutputProvider = $script:cbCompareOutputFormat.SelectedItem
    }

    if($script:cbCompareCSVDelimiter)
    {
        $params.CsvDelimiter = $script:cbCompareCSVDelimiter.Text
    }

    if($script:cbCompareMultiValueDelimiter -and $null -ne $script:cbCompareMultiValueDelimiter.SelectedValue)
    {
        $params.ObjectSeparator = [string]$script:cbCompareMultiValueDelimiter.SelectedValue
    }

    if($script:chkSkipCompareAssignments)
    {
        $params.SkipAssignments = [bool]$script:chkSkipCompareAssignments.IsChecked
    }

    Set-CompareRuntimeOptions @params
}

function Get-CompareMultiValueDelimiterItems
{
    @(
        [PSCustomObject]@{ Name = "New line"; Value = [System.Environment]::NewLine },
        [PSCustomObject]@{ Name = ";";        Value = ";" },
        [PSCustomObject]@{ Name = "|";        Value = "|" }
    )
}

function Initialize-CompareSharedOptionControls
{
    # Controls shared by the single and bulk compare forms.
    if($script:cbCompareMultiValueDelimiter)
    {
        $script:cbCompareMultiValueDelimiter.ItemsSource    = @(Get-CompareMultiValueDelimiterItems)
        $script:cbCompareMultiValueDelimiter.SelectedValue = (Get-SettingStoreValue "Compare" "ObjectSeparator" ([System.Environment]::NewLine))
        if($null -eq $script:cbCompareMultiValueDelimiter.SelectedItem)
        {
            $script:cbCompareMultiValueDelimiter.SelectedIndex = 0
        }
    }
    if($script:chkSkipCompareAssignments)
    {
        $script:chkSkipCompareAssignments.IsChecked = ((Get-SettingStoreValue "Compare" "SkipCompareAssignments" "false") -eq "true")
    }
}

function Save-CompareSharedOptionSettings
{
    if($script:cbCompareMultiValueDelimiter -and $null -ne $script:cbCompareMultiValueDelimiter.SelectedValue)
    {
        Save-SettingStoreValue "Compare" "ObjectSeparator" $script:cbCompareMultiValueDelimiter.SelectedValue
    }
    if($script:chkSkipCompareAssignments)
    {
        Save-SettingStoreValue "Compare" "SkipCompareAssignments" $(if($script:chkSkipCompareAssignments.IsChecked) { "true" } else { "false" })
    }
}

# ─── Object Picker ────────────────────────────────────────────────────────────

function Show-ObjectPickerDialog
{
    param(
        [string]   $Title,
        [object[]] $Items,
        [object[]] $DisplayColumns,
        [scriptblock] $LoadHandler,
        [string]   $LoadLabel = "Load all from Intune",
        # Entries from Get-PolicySearchScopes. When supplied, the dialog shows
        # a "Search in" dropdown above the search box.
        $Scopes = $null,
        [string]  $SelectedScopeKey = $null,
        # Entries from Get-PolicySearchTenants. When supplied, the dialog shows
        # a "Tenant" dropdown so the user can search a different tenant.
        $Tenants = $null,
        [string]  $SelectedTenantKey = $null,
        # Called as & $SearchHandler $searchText $scope and expected to return
        # the rows to display. When absent the Search button filters the
        # supplied -Items client-side (the original behaviour).
        [scriptblock] $SearchHandler = $null
    )

    # Per-call state on $script:_pickerState. Closures below dynamically read
    # this slot at INVOCATION time instead of capturing locals via
    # .GetNewClosure() — that captures every variable visible at create time,
    # including module-scope ones like $script:dgPickerObjects which are
    # populated AFTER the closures are built (Show-PickerDialog assigns them
    # when it materialises the dialog). Captured-as-null is what produced the
    # "You cannot call a method on a null-valued expression" errors at
    # $script:dgPickerObjects.Columns.Add and the 'cannot be found' error at
    # $script:dgPickerObjects.ItemsSource. Using a $script: bag keeps lookup
    # late-bound; assignments by Show-PickerDialog land in the same scope the
    # closures read from.
    $script:_pickerState = @{
        Columns = $DisplayColumns
        # ItemsBox lets the LoadHandler replace the list in-place; the apply
        # path reassigns ItemsBox.Items and every closure picks up the new
        # array on the next read.
        ItemsBox = @{ Items = Get-UIItemsArray $Items }
        Result   = @{ Value = $null }
        SearchHandler = $SearchHandler
    }

    $initHandler = {
        $ps = $script:_pickerState
        foreach($colDef in $ps.Columns)
        {
            $col            = [System.Windows.Controls.DataGridTextColumn]::new()
            $col.Header     = $colDef.Header
            $col.IsReadOnly = $true
            $col.Binding    = [System.Windows.Data.Binding]::new($colDef.Binding)
            $script:dgPickerObjects.Columns.Add($col)
        }
        $script:dgPickerObjects.ItemsSource = $ps.ItemsBox.Items
    }

    $clickHandler = {
        $ps = $script:_pickerState
        $filter = $script:txtPickerSearch.Text

        if($ps.SearchHandler)
        {
            # Server-side search: ask the caller for rows matching the text in
            # the selected scope, and show exactly what comes back.
            $scope = Get-PickerSelectedScope
            try {
                $script:pickerDialog.Cursor = [System.Windows.Input.Cursors]::Wait
                $found = Get-UIItemsArray (& $ps.SearchHandler $filter $scope)
            }
            catch {
                Write-LogError "Picker SearchHandler failed" $_.Exception
                $found = @()
            }
            finally {
                $script:pickerDialog.Cursor = $null
            }

            # The status line belongs to the search handler - it knows whether
            # an empty result means "no matches" or "keep typing".
            $ps.ItemsBox.Items = $found
            $script:dgPickerObjects.ItemsSource = $found
            return
        }

        if([string]::IsNullOrEmpty($filter))
        {
            $script:dgPickerObjects.ItemsSource = $ps.ItemsBox.Items
        }
        else
        {
            $script:dgPickerObjects.ItemsSource = @($ps.ItemsBox.Items | Where-Object { $_.Name -ilike "*$filter*" })
        }
    }

    $okHandler = {
        $script:_pickerState.Result.Value = $script:dgPickerObjects.SelectedItem
        $script:pickerDialog.Close()
    }

    $wrappedLoadHandler = $null
    if($LoadHandler) {
        # Stash the user-supplied loader so the wrapper script block can
        # reach it via $script: at runtime (same late-bind reasoning as
        # above — no closure capture).
        $script:_pickerState.UserLoadHandler = $LoadHandler
        $script:_pickerState.ClickHandler    = $clickHandler
        $wrappedLoadHandler = {
            $ps = $script:_pickerState
            $newItems = & $ps.UserLoadHandler

            if($ps.SearchHandler) {
                # Scoped loads replace rather than merge - the returned set IS
                # everything in the newly selected scope, and merging would
                # leave stale rows from a previous scope in the grid.
                $ps.ItemsBox.Items = Get-UIItemsArray $newItems
                $script:dgPickerObjects.ItemsSource = $ps.ItemsBox.Items
                return
            }

            if($newItems) {
                # Dedupe by Id; new entries appended so cached items stay first.
                $existingIds = @{}
                foreach($existing in $ps.ItemsBox.Items) {
                    if($existing -and $existing.Id) { $existingIds[$existing.Id] = $true }
                }
                $merged = [System.Collections.Generic.List[object]]::new()
                foreach($existing in $ps.ItemsBox.Items) { [void]$merged.Add($existing) }
                foreach($n in $newItems) {
                    if(-not $n) { continue }
                    if($n.Id -and $existingIds.ContainsKey($n.Id)) { continue }
                    [void]$merged.Add($n)
                }
                $ps.ItemsBox.Items = $merged.ToArray()
            }
            & $ps.ClickHandler
        }
    }

    try {
        Show-PickerDialog -title $Title -initHandler $initHandler -clickHandler $clickHandler -okHandler $okHandler -LoadHandler $wrappedLoadHandler -LoadLabel $LoadLabel -Scopes $Scopes -SelectedScopeKey $SelectedScopeKey -Tenants $Tenants -SelectedTenantKey $SelectedTenantKey
        return $script:_pickerState.Result.Value
    }
    finally {
        # Avoid keeping the closed-over Items / handlers alive between dialog
        # invocations; next Show-ObjectPickerDialog call replaces this anyway.
        $script:_pickerState = $null
    }
}

# ─── Bulk Compare ─────────────────────────────────────────────────────────────

function Show-GraphBulkCompareForm
{
    $script:bulkCompareForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\BulkCompare.xaml"), $true)
    if(-not $script:bulkCompareForm) { return }

    $script:cbCompareProvider.ItemsSource    = @($script:compareProviders)
    $script:cbCompareProvider.SelectedValue = (Get-SettingStoreValue "Compare" "Provider" "export")
    # A saved provider id that no longer registers (a provider was removed) must
    # not leave the combo blank - Start would then be a silent no-op. Mirrors the
    # Avalonia form's first-item fallback.
    if($null -eq $script:cbCompareProvider.SelectedItem -and $script:cbCompareProvider.Items.Count -gt 0)
    {
        $script:cbCompareProvider.SelectedIndex = 0
    }

    $script:cbCompareSave.ItemsSource    = @($script:compareOutputTypes)
    $script:cbCompareSave.SelectedValue = (Get-SettingStoreValue "Compare" "SaveType" "objectType")

    $script:cbCompareOutputFormat.ItemsSource    = @($script:compareOutputProviders)
    $script:cbCompareOutputFormat.SelectedValue = (Get-SettingStoreValue "Compare" "OutputFormat" "csv")

    $script:cbCompareType.ItemsSource    = $script:comparisonTypes
    $script:cbCompareType.SelectedValue = (Get-SettingStoreValue "Compare" "Type" "property")

    $script:cbCompareCSVDelimiter.ItemsSource    = @("", ",", ";", "-", "|")
    $script:cbCompareCSVDelimiter.SelectedValue = (Get-SettingStoreValue "Compare" "Delimiter" ";")

    Initialize-CompareSharedOptionControls
    if($script:chkIgnoreCoreProperties)
    {
        $script:chkIgnoreCoreProperties.IsChecked = ((Get-SettingStoreValue "Compare" "IgnoreCoreProperties" "true") -eq "true")
    }
    if($script:chkSkipMissingSourcePolicies)
    {
        $script:chkSkipMissingSourcePolicies.IsChecked = ((Get-SettingStoreValue "Compare" "SkipMissingSourcePolicies" "false") -eq "true")
    }
    if($script:chkSkipMissingDestinationPolicies)
    {
        $script:chkSkipMissingDestinationPolicies.IsChecked = ((Get-SettingStoreValue "Compare" "SkipMissingDestinationPolicies" "false") -eq "true")
    }

    $script:compareObjects = @()
    foreach($intuneGroup in $script:IntuneGroups)
    {
        if(-not $intuneGroup.Title) { continue }
        $script:compareObjects += [PSCustomObject]@{
            Title       = $intuneGroup.Title
            Selected    = $true
            ObjectGroup = $intuneGroup
        }
    }

    $column = Get-GridCheckboxColumn "Selected"
    $script:dgObjectsToCompare.Columns.Add($column)
    $column.Header.IsChecked = $true
    $column.Header.Add_Click({
        foreach($item in $script:dgObjectsToCompare.ItemsSource) { $item.Selected = $this.IsChecked }
        $script:dgObjectsToCompare.Items.Refresh()
    })

    $col = [System.Windows.Controls.DataGridTextColumn]::new()
    $col.Header    = "Object type"
    $col.IsReadOnly = $true
    $col.Binding   = [System.Windows.Data.Binding]::new("Title")
    $script:dgObjectsToCompare.Columns.Add($col)

    $script:dgObjectsToCompare.ItemsSource = $script:compareObjects

    $script:cbCompareOutputFormat.Add_SelectionChanged({
        $isCSV = ($script:cbCompareOutputFormat.SelectedItem -is [CompareCSVOutputProvider])
        $script:cbCompareCSVDelimiter.IsEnabled = $isCSV
    })

    $script:UIProvider.AddXamlEvent($script:bulkCompareForm, "btnClose", "add_click", {
        $script:bulkCompareForm = $null
        $script:UIProvider.ShowModalObject()
    })

    $script:UIProvider.AddXamlEvent($script:bulkCompareForm, "btnStartCompare", "add_click", {
        Write-Status "Compare objects"
        Save-SettingStoreValue "Compare" "Provider"     $script:cbCompareProvider.SelectedValue
        Save-SettingStoreValue "Compare" "Type"         $script:cbCompareType.SelectedValue
        Save-SettingStoreValue "Compare" "Delimiter"    $script:cbCompareCSVDelimiter.SelectedValue
        Save-SettingStoreValue "Compare" "OutputFormat" $script:cbCompareOutputFormat.SelectedValue
        Save-SettingStoreValue "Compare" "SaveType"     $script:cbCompareSave.SelectedValue
        Save-CompareSharedOptionSettings
        if($script:chkIgnoreCoreProperties)
        {
            Save-SettingStoreValue "Compare" "IgnoreCoreProperties" $(if($script:chkIgnoreCoreProperties.IsChecked) { "true" } else { "false" })
        }

        try
        {
            Set-CompareRuntimeOptionsFromUI
            $Provider = $script:cbCompareProvider.SelectedItem
            if($Provider)
            {
                if($script:chkSkipMissingSourcePolicies)
                {
                    $Provider.SkipMissingSourcePolicies = [bool]$script:chkSkipMissingSourcePolicies.IsChecked
                    Save-SettingStoreValue "Compare" "SkipMissingSourcePolicies" $(if($Provider.SkipMissingSourcePolicies) { "true" } else { "false" })
                }
                if($script:chkSkipMissingDestinationPolicies)
                {
                    $Provider.SkipMissingDestinationPolicies = [bool]$script:chkSkipMissingDestinationPolicies.IsChecked
                    Save-SettingStoreValue "Compare" "SkipMissingDestinationPolicies" $(if($Provider.SkipMissingDestinationPolicies) { "true" } else { "false" })
                }
            }
            $selectedGroups = @()
            if($script:dgObjectsToCompare) {
                $selectedGroups = @($script:dgObjectsToCompare.ItemsSource | Where-Object Selected -eq $true | ForEach-Object { $_.ObjectGroup })
            }
            if($Provider) { Start-BulkCompare $Provider -SelectedGroups $selectedGroups }
        }
        catch
        {
            $script:UIProvider.ShowMessageBox($_.Exception.Message, "Compare", "OK", "Error")
        }
        finally
        {
            Write-Status ""
        }
    })

    $script:cbCompareProvider.Add_SelectionChanged({ Set-CompareProviderOptions $this })

    Set-CompareProviderOptions $script:cbCompareProvider

    $script:UIProvider.ShowModalForm("Bulk Compare Objects", $script:bulkCompareForm, $true)
}

function Update-BulkCompareFormForProvider
{
    param($Provider)

    if(-not $script:grdObjectsToCompareSection) { return }
    $script:grdObjectsToCompareSection.Visibility = if($Provider -and $Provider.IgnoreGroups) { "Collapsed" } else { "Visible" }
}

# ─── Single Compare ──────────────────────────────────────────────────────────

function Show-GraphCompareForm
{
    param([object[]]$Policies)

    if(-not $Policies -or $Policies.Count -eq 0) { return }

    $script:compareForm            = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\CompareForm.xaml"), $true)
    if(-not $script:compareForm) { return }

    $script:compareSource          = $Policies[0]
    $script:compareTargetPolicy    = if($Policies.Count -ge 2) { $Policies[1] } else { $null }
    $script:compareInputPanel      = $null
    $script:compareInputPanelCache = @{}

    $script:cbCompareType.ItemsSource    = $script:comparisonTypes
    $script:cbCompareType.SelectedValue = (Get-SettingStoreValue "Compare" "Type" "property")

    $script:cbSingleCompareOutputFormat.ItemsSource    = @($script:compareOutputProviders)
    $script:cbSingleCompareOutputFormat.SelectedValue = (Get-SettingStoreValue "Compare" "SingleOutputFormat" "csv")

    if($script:chkIgnoreCoreProperties)
    {
        $script:chkIgnoreCoreProperties.IsChecked = ((Get-SettingStoreValue "Compare" "IgnoreCoreProperties" "true") -eq "true")
    }

    Initialize-CompareSharedOptionControls

    if($script:cbCompareFilter)
    {
        # 3-way result filter (All / Mismatch / Match).
        $savedFilter = Get-SettingStoreValue "Compare" "ResultFilter" "All"
        $script:cbCompareFilter.SelectedItem = @($script:cbCompareFilter.Items) | Where-Object { "$($_.Tag)" -eq "$savedFilter" } | Select-Object -First 1
        if(-not $script:cbCompareFilter.SelectedItem) { $script:cbCompareFilter.SelectedIndex = 0 }
        $script:cbCompareFilter.Add_SelectionChanged({ Update-CompareGridFilter })
    }

    $script:txtIntuneObject.Text = $Policies[0].Name

    $compareInputTypes = @(
        [PSCustomObject]@{ Name = "File";         Value = "file";   OptionsXaml = "CompareFileOptions"         }
        [PSCustomObject]@{ Name = "Intune Object"; Value = "intune"; OptionsXaml = "CompareIntuneObjectOptions" }
    )
    $script:cbCompareInput.ItemsSource = $compareInputTypes

    $script:cbCompareInput.Add_SelectionChanged({ Set-CompareInputOptions $this })
    $script:cbCompareInput.SelectedValue = if($script:compareTargetPolicy) { "intune" } else { (Get-SettingStoreValue "Compare" "InputType" "file") }

    $script:UIProvider.AddXamlEvent($script:compareForm, "btnClose", "add_click", {
        $script:compareForm              = $null
        $script:compareSource            = $null
        $script:compareTargetPolicy      = $null
        $script:compareInputPanel        = $null
        $script:compareInputPanelCache   = $null
        $script:chkIgnoreCoreProperties  = $null
        $script:cbCompareFilter          = $null
        $script:chkSkipCompareAssignments = $null
        $script:cbCompareMultiValueDelimiter = $null
        $script:txtCompareSummary        = $null
        $script:dgCompareInfo            = $null
        $script:UIProvider.ShowModalObject()
    })

    $script:UIProvider.AddXamlEvent($script:compareForm, "btnStartCompare", "add_click", {
        Write-Status "Compare objects"
        Save-SettingStoreValue "Compare" "Type"                 $script:cbCompareType.SelectedValue
        Save-SettingStoreValue "Compare" "InputType"            $script:cbCompareInput.SelectedValue
        Save-SettingStoreValue "Compare" "SingleOutputFormat"   $script:cbSingleCompareOutputFormat.SelectedValue
        if($script:chkIgnoreCoreProperties)
        {
            Save-SettingStoreValue "Compare" "IgnoreCoreProperties" $(if($script:chkIgnoreCoreProperties.IsChecked) { "true" } else { "false" })
        }
        if($script:cbCompareFilter -and $script:cbCompareFilter.SelectedItem)
        {
            Save-SettingStoreValue "Compare" "ResultFilter" ([string]$script:cbCompareFilter.SelectedItem.Tag)
        }
        Save-CompareSharedOptionSettings
        Set-CompareRuntimeOptionsFromUI
        Start-CompareObjects
        Write-Status ""
    })

    $script:UIProvider.AddXamlEvent($script:compareForm, "btnSwap", "add_click", {
        if(-not $script:compareSource -or -not $script:compareTargetPolicy) { return }
        $tmp                          = $script:compareSource
        $script:compareSource         = $script:compareTargetPolicy
        $script:compareTargetPolicy   = $tmp
        $script:txtIntuneObject.Text  = $script:compareSource.Name
        if($script:compareInputPanel)
        {
            $txt = $script:compareInputPanel.FindName("txtCompareIntuneObject")
            if($txt) { $txt.Text = $script:compareTargetPolicy.Name }
        }
    })

    $script:UIProvider.AddXamlEvent($script:compareForm, "btnCompareSave", "add_click", {
        if(($script:dgCompareInfo.ItemsSource | Measure-Object).Count -eq 0) { return }

        $outProv = if($script:cbSingleCompareOutputFormat -and $script:cbSingleCompareOutputFormat.SelectedItem) {
            $script:cbSingleCompareOutputFormat.SelectedItem
        } else {
            [CompareCSVOutputProvider]::new()
        }

        $sf = [System.Windows.Forms.SaveFileDialog]::new()
        $sf.FileName    = $script:compareSource.Name
        $sf.DefaultExt  = $outProv.Extension
        $sf.FilterIndex = if($outProv.Extension -eq "json") { 2 } else { 1 }
        $sf.Filter      = "CSV files (*.csv)|*.csv|JSON files (*.json)|*.json|All files (*.*)|*.*"
        $initialDir = Get-SettingStoreValue "Compare" "LastSaveDirectory"
        if(-not $initialDir) { $initialDir = Get-SettingValue "RootFolder" }
        if($initialDir) { $sf.InitialDirectory = $initialDir }

        if($sf.ShowDialog() -eq "OK")
        {
            Save-SettingStoreValue "Compare" "LastSaveDirectory" ([IO.FileInfo]$sf.FileName).DirectoryName
            $ext     = [IO.Path]::GetExtension($sf.FileName).TrimStart(".")
            $outProv = Get-CompareOutputProviderByExtension $ext
            $props   = Get-CompareOutputProps
            $outProv.FormatRows($script:dgCompareInfo.ItemsSource, $props) | Out-File -LiteralPath $sf.FileName -Force -Encoding UTF8
        }
    })

    $script:UIProvider.AddXamlEvent($script:compareForm, "btnCompareCopy", "add_click", {
        Set-CompareRuntimeOptionsFromUI
        (Get-CompareCsvInfo $script:dgCompareInfo.ItemsSource $script:compareSource) | Set-Clipboard
    })

    $script:UIProvider.ShowModalForm("Compare Intune Objects", $script:compareForm, $true)
}

function Set-CompareInputOptions
{
    param($Control)

    $InputType = $Control.SelectedItem
    if(-not $InputType) { return }

    $Panel = $null
    if($script:compareInputPanelCache -is [Hashtable] -and $script:compareInputPanelCache.ContainsKey($InputType.Value))
    {
        $Panel = $script:compareInputPanelCache[$InputType.Value]
    }
    elseif($InputType.OptionsXaml)
    {
        $Panel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\$($InputType.OptionsXaml).xaml"))
        if($Panel)
        {
            $script:compareInputPanelCache[$InputType.Value] = $Panel
            Register-CompareInputBrowseEvents $Panel $InputType
        }
    }

    $script:compareInputPanel              = $Panel
    $script:ccCompareInputOptions.Content  = $Panel

    if($Panel -and $InputType.Value -eq "intune" -and $script:compareTargetPolicy)
    {
        $txt = $Panel.FindName("txtCompareIntuneObject")
        if($txt) { $txt.Text = $script:compareTargetPolicy.Name }
    }
}

function Register-CompareInputBrowseEvents
{
    param($Panel, $InputType)

    $browseFile = $Panel.FindName("browseCompareObject")
    if($browseFile)
    {
        $browseFile.Tag = $Panel
        $browseFile.Add_Click({
            $p = $this.Tag
            $path = Get-SettingStoreValue "" "LastUsedFullPath"
            if($path) { $path = [IO.Directory]::GetParent($path).FullName }

            if($script:compareSource -and $script:compareSource.PolicyType)
            {
                # No ?? operator here - this file must parse on PS5.1 too.
                $typePath = [IO.Path]::Combine($(if($path) { $path } else { "" }), $script:compareSource.PolicyType.Folder)
                if([IO.Directory]::Exists($typePath)) { $path = $typePath }
            }

            $lastFile = Get-SettingStoreValue "Compare" "LastFile"
            if(-not [String]::IsNullOrEmpty($lastFile)) { $path = ([IO.FileInfo]$lastFile).DirectoryName }

            $of = [System.Windows.Forms.OpenFileDialog]::new()
            $of.Multiselect = $false
            $of.Filter      = "Json files (*.json)|*.json"
            if($path) { $of.InitialDirectory = $path }

            if($of.ShowDialog() -eq "OK")
            {
                $txt = $p.FindName("txtCompareFile")
                if($txt) { $txt.Text = $of.FileName }
                Save-SettingStoreValue "Compare" "LastFile" $of.FileName
            }
        })
    }

    $browseIntune = $Panel.FindName("browseIntuneObject")
    if($browseIntune)
    {
        $browseIntune.Tag = $Panel
        $browseIntune.Add_Click({
            $p = $this.Tag

            $sourceType = if($script:compareSource) { $script:compareSource.PolicyType } else { $null }
            if(-not $sourceType) {
                $script:UIProvider.ShowMessageBox("Select a source policy first.", "Compare", "OK", "Warning") | Out-Null
                return
            }
            $sourceId = if($script:compareSource) { $script:compareSource.Id } else { $null }

            # Seed with same-type policies already cached in the main view so the
            # dialog opens with something in it and no Graph call. Searching or
            # "Load all in scope" goes to Graph from there.
            $initial = @()
            if($script:intuneManagerPolicyCollection) {
                $initial = @($script:intuneManagerPolicyCollection | Where-Object {
                    $_.PolicyType -and $_.PolicyType.ID -eq $sourceType.ID -and
                    (-not $sourceId -or $_.Id -ne $sourceId)
                })
            }

            $exclude = @()
            if($sourceId) { $exclude = @($sourceId) }

            $selected = Show-PolicySearchDialog `
                -Title "Select Intune Object to Compare ($($sourceType.Title))" `
                -DefaultPolicyTypeId $sourceType.ID `
                -ExcludeIds $exclude `
                -InitialItems $initial
            if($selected)
            {
                $script:compareTargetPolicy = $selected
                $txt = $p.FindName("txtCompareIntuneObject")
                if($txt) { $txt.Text = $selected.Name }
            }
        })
    }
}

function Start-CompareObjects
{
    if(-not $script:compareSource) { return }

    $compareMode = if($script:cbCompareInput) { $script:cbCompareInput.SelectedValue } `
                   elseif($script:compareTargetPolicy) { "intune" } `
                   else { "file" }

    if($compareMode -eq "intune")
    {
        if(-not $script:compareTargetPolicy)
        {
            $script:UIProvider.ShowMessageBox("No Intune object selected for comparison", "Compare", "OK", "Error")
            return
        }

        Write-Status "Compare Intune objects"
        Resolve-FullPolicyForCompare $script:compareSource
        Resolve-FullPolicyForCompare $script:compareTargetPolicy

        $compareResult = Compare-PolicyObjects @($script:compareSource, $script:compareTargetPolicy)
        $script:dgCompareInfo.ItemsSource = $compareResult
        Update-CompareGridFilter
        Update-CompareSummary

        $tabCtrl = $script:compareForm.FindName("tabCompare")
        if($tabCtrl) { $tabCtrl.SelectedIndex = 1 }
        Write-Status ""
        return
    }

    $compareFile = ""
    if($script:compareInputPanel)
    {
        $txtFile = $script:compareInputPanel.FindName("txtCompareFile")
        if($txtFile) { $compareFile = $txtFile.Text }
    }
    if(-not $compareFile)
    {
        $script:UIProvider.ShowMessageBox("No file selected for comparison", "Compare", "OK", "Error")
        return
    }
    if(-not [IO.File]::Exists($compareFile))
    {
        $script:UIProvider.ShowMessageBox("File '$compareFile' not found", "Compare", "OK", "Error")
        return
    }

    try
    {
        $filePolicy = $script:compareSource.PolicyType.GetObject([IO.FileInfo]$compareFile)
        if(-not $filePolicy)
        {
            $script:UIProvider.ShowMessageBox("Failed to load file '$compareFile'. Object type may not match.", "Compare", "OK", "Error")
            return
        }
    }
    catch
    {
        $script:UIProvider.ShowMessageBox("Failed to load file '$compareFile': $($_.Exception.Message)", "Compare", "OK", "Error")
        return
    }

    Resolve-FullPolicyForCompare $script:compareSource

    if($script:compareSource.Object.'@OData.Type' -ne $filePolicy.Object.'@OData.Type')
    {
        if(($script:UIProvider.ShowMessageBox("The object types do not match.`n`nDo you want to compare the objects anyway?", "Compare", "YesNo", "Warning")) -ne "Yes")
        {
            return
        }
    }

    $compareResult = Compare-PolicyObjects @($script:compareSource, $filePolicy)
    $script:dgCompareInfo.ItemsSource = $compareResult
    Update-CompareGridFilter
    Update-CompareSummary

    $tabCtrl = $script:compareForm.FindName("tabCompare")
    if($tabCtrl) { $tabCtrl.SelectedIndex = 1 }
    Write-Status ""
}

function Update-CompareSummary
{
    if(-not $script:txtCompareSummary -or -not $script:dgCompareInfo) { return }
    $items = @($script:dgCompareInfo.ItemsSource)
    $total = $items.Count
    $mismatches = @($items | Where-Object { $_.Match -eq $false }).Count
    $skipped    = @($items | Where-Object { $null -eq $_.Match }).Count
    $matched    = $total - $mismatches - $skipped
    $script:txtCompareSummary.Text = "$total properties - $matched matched, $mismatches mismatched$(if($skipped -gt 0) { ", $skipped ignored" } else { '' })"
}

function Update-CompareGridFilter
{
    if(-not $script:dgCompareInfo -or -not $script:dgCompareInfo.Items) { return }
    $mode = if($script:cbCompareFilter -and $script:cbCompareFilter.SelectedItem) { [string]$script:cbCompareFilter.SelectedItem.Tag } else { "All" }
    switch($mode)
    {
        "Mismatch" { $script:dgCompareInfo.Items.Filter = [Predicate[object]]{ param($item) $item.Match -eq $false } }
        "Match"    { $script:dgCompareInfo.Items.Filter = [Predicate[object]]{ param($item) $item.Match -eq $true } }
        default    { $script:dgCompareInfo.Items.Filter = $null }
    }
}

# ─── Provider Options ─────────────────────────────────────────────────────────

function Set-CompareProviderOptions
{
    param($Control)

    $Provider      = $Control.SelectedItem
    $providerPanel = $null

    if($Provider -and $Provider.OptionsXaml)
    {
        # The cached panels are keyed by tenant so a tenant switch rebuilds them.
        # No current provider caches tenant data on its panel; the invalidation
        # stays as the only guard should a future one do so.
        if($script:CompareProviderOptionsCache -isnot [Hashtable] -or
           $script:CompareProviderOptionsCacheTenant -ne $script:OrganizationId)
        {
            $script:CompareProviderOptionsCache = @{}
            $script:CompareProviderOptionsCacheTenant = $script:OrganizationId
        }

        if($script:CompareProviderOptionsCache.ContainsKey($Provider.Value))
        {
            $providerPanel = $script:CompareProviderOptionsCache[$Provider.Value]
        }
        else
        {
            $providerPanel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\$($Provider.OptionsXaml).xaml"))
            if($providerPanel)
            {
                $script:CompareProviderOptionsCache[$Provider.Value] = $providerPanel
                Register-CompareProviderBrowseEvents $providerPanel $Provider
            }
            else
            {
                Write-Log "Failed to load provider options for '$($Provider.Name)'" 3
            }
        }

        if($providerPanel)
        {
            $providerPanel.DataContext = $Provider
        }

        $script:ccContentProviderOptions.Content = $providerPanel
    }
    else
    {
        $script:ccContentProviderOptions.Content = $null
    }

    $script:ccContentProviderOptions.Visibility = if($null -eq $script:ccContentProviderOptions.Content) { "Collapsed" } else { "Visible" }

    Update-BulkCompareFormForProvider $Provider
}

function Register-CompareProviderBrowseEvents
{
    param($Panel, $Provider)

    $browseExportPath = $Panel.FindName("browseExportPath")
    if($browseExportPath)
    {
        $browseExportPath.Tag = @{ Provider = $Provider; Panel = $Panel }
        $browseExportPath.Add_Click({
            $tag = $this.Tag
            $folder = $script:UIProvider.ShowFolderPicker($tag.Provider.ExportPath, "Select root folder for compare")
            if($folder)
            {
                $tag.Provider.ExportPath = $folder
                $dc = $tag.Panel.DataContext; $tag.Panel.DataContext = $null; $tag.Panel.DataContext = $dc
            }
        })
    }

    $browseSavePath = $Panel.FindName("browseSavePath")
    if($browseSavePath)
    {
        $browseSavePath.Tag = @{ Provider = $Provider; Panel = $Panel }
        $browseSavePath.Add_Click({
            $tag = $this.Tag
            $folder = $script:UIProvider.ShowFolderPicker($tag.Provider.SavePath, "Select save folder")
            if($folder)
            {
                $tag.Provider.SavePath = $folder
                $dc = $tag.Panel.DataContext; $tag.Panel.DataContext = $null; $tag.Panel.DataContext = $dc
            }
        })
    }

    $browseSource = $Panel.FindName("browseExportPathSource")
    if($browseSource)
    {
        $browseSource.Tag = @{ Provider = $Provider; Panel = $Panel }
        $browseSource.Add_Click({
            $tag = $this.Tag
            $folder = $script:UIProvider.ShowFolderPicker($tag.Provider.SourcePath, "Select source root folder")
            if($folder)
            {
                $tag.Provider.SourcePath = $folder
                $dc = $tag.Panel.DataContext; $tag.Panel.DataContext = $null; $tag.Panel.DataContext = $dc
            }
        })
    }

    $browseCompare = $Panel.FindName("browseExportPathCompare")
    if($browseCompare)
    {
        $browseCompare.Tag = @{ Provider = $Provider; Panel = $Panel }
        $browseCompare.Add_Click({
            $tag = $this.Tag
            $folder = $script:UIProvider.ShowFolderPicker($tag.Provider.ComparePath, "Select compare root folder")
            if($folder)
            {
                $tag.Provider.ComparePath = $folder
                $dc = $tag.Panel.DataContext; $tag.Panel.DataContext = $null; $tag.Panel.DataContext = $dc
            }
        })
    }
}
