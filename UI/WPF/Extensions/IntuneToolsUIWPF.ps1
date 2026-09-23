Add-AppEventHandler "AppInitialized" "Initialize-IntuneAssignmentsModule"

# ─── View ─────────────────────────────────────────────────────────────────────

function Get-IntuneToolsViewItems
{
    # MenuLabel is bound by the shared lstMenuItems template (MainWindow.xaml)
    # — IntuneView populates it dynamically (with optional "(N)" counts); these
    # tool rows are static so MenuLabel == Title is fine. Without it the tool
    # rows render blank because the template binds MenuLabel, not Title.
    #
    # Category drives the grouped-expander rendering in Show-ViewMenu — when
    # any item declares one, the menu list becomes a tree of expandable groups
    # (one per distinct Category) with the tools nested inside.
    $items = @()

    $items += [PSCustomObject]@{
        Title       = "Intune Assignments"
        MenuLabel   = "Intune Assignments"
        Id          = "IntuneAssignments"
        Category    = "Assignments"
        Description = "List assignments for Intune objects, from an exported folder or directly from Intune."
        IconImage   = Get-IntuneToolsItemIcon "DeviceConfiguration"
    }

    $items += [PSCustomObject]@{
        Title       = "Intune Filter Usage"
        MenuLabel   = "Intune Filter Usage"
        Id          = "IntuneFilterUsage"
        Category    = "Assignments"
        Description = "Show every assignment filter and the policies / apps that reference it."
        IconImage   = Get-IntuneToolsItemIcon "DeviceConfiguration"
    }

    $items += [PSCustomObject]@{
        Title       = "ADMX Import"
        MenuLabel   = "ADMX Import"
        Id          = "ADMXImport"
        Category    = "ADMX"
        Description = "Load an .admx (+ .adml) file, edit policy settings, and ingest as a Custom OMA-URI device-configuration profile."
        IconImage   = Get-IntuneToolsItemIcon "DeviceConfiguration"
    }

    $items += [PSCustomObject]@{
        Title       = "Reg Values"
        MenuLabel   = "Reg Values"
        Id          = "ADMXRegValues"
        Category    = "ADMX"
        Description = "Build a Custom OMA-URI policy from manually-added registry keys/values (HKLM + HKCU)."
        IconImage   = Get-IntuneToolsItemIcon "DeviceConfiguration"
    }

    return $items
}

function Get-IntuneToolsItemIcon
{
    param([string]$IconName)

    if(-not $IconName) { return $null }
    $iconFile = [IO.Path]::Combine($script:AppUIRootFolder, "Xaml", "Icons", "$IconName.xaml")
    if(-not [IO.File]::Exists($iconFile)) { return $null }
    return ($script:UIProvider.GetXamlObject($iconFile))
}

function Get-IntuneToolsViewPanel
{
    $panel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\IntuneTools.xaml"))
    if(-not $panel) { return $null }

    $script:grdToolsMain = $panel.FindName("grdToolsMain")
    return $panel
}

function Invoke-IntuneToolsActivateItem
{
    param($SelectedItem)

    if(-not $SelectedItem -or -not $script:grdToolsMain) { return $null }

    if($SelectedItem -is [System.Windows.Data.CollectionViewGroup]) {
        return $null
    }

    if(-not $SelectedItem.PSObject.Properties['Id']) {
        $label = $null
        foreach($prop in @('Title','MenuLabel','Name')) {
            if($SelectedItem.PSObject.Properties[$prop] -and $SelectedItem.$prop) {
                $label = [string]$SelectedItem.$prop
                break
            }
        }
        if($label) {
            $SelectedItem = @(Get-IntuneToolsViewItems) | Where-Object {
                $_.Title -eq $label -or $_.MenuLabel -eq $label -or $_.Id -eq $label
            } | Select-Object -First 1
        }
    }

    if(-not $SelectedItem -or -not $SelectedItem.PSObject.Properties['Id']) {
        Write-Log "Unable to activate Intune tool. Selected item did not include a tool Id." 2
        return $null
    }

    Write-Log "Activate tool: $($SelectedItem.Title)"

    switch($SelectedItem.Id)
    {
        "IntuneAssignments"  { Show-IntuneAssignmentsTool }
        "IntuneFilterUsage"  { Show-IntuneFilterUsageTool }
        "ADMXImport"         { Show-ADMXImportTool }
        "ADMXRegValues"      { Show-ADMXRegValuesTool }
        default
        {
            Write-Log "Unknown tool selected: $($SelectedItem.Id)" 2
            $script:grdToolsMain.Children.Clear()
        }
    }

    return $null
}

function Get-IntuneToolsDataGridVisibleItems
{
    param($DataGrid)

    if(-not $DataGrid -or -not $DataGrid.Items) { return @() }

    $items = @()
    foreach($item in $DataGrid.Items) {
        if($null -ne $item) { $items += $item }
    }
    return $items
}

# ─── Intune Assignments tool ──────────────────────────────────────────────────

function Show-IntuneAssignmentsTool
{
    if(-not $script:grdToolsMain) { return }

    $script:grdToolsMain.Children.Clear()

    $panel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\IntuneToolsAssignments.xaml"), $true)
    if(-not $panel)
    {
        Write-Log "Failed to load IntuneToolsAssignments.xaml" 3
        return
    }

    $script:intuneAssignmentsPanel        = $panel
    $script:intuneAssignmentsProviderPanelCache = @{}
    # Keep last-loaded rows across panel rebuilds so the user doesn't have to
    # re-run the load every time they navigate away from and back to the tool.
    if(-not $script:intuneAssignmentsRows) { $script:intuneAssignmentsRows = @() }

    $script:cbIntuneAssignmentsProvider.ItemsSource    = @($script:intuneAssignmentsProviders)
    $script:cbIntuneAssignmentsProvider.SelectedValue  = (Get-SettingStoreValue "IntuneAssignments" "Provider" "folder")

    $script:cbIntuneAssignmentsProvider.Add_SelectionChanged({ Set-IntuneAssignmentsProviderOptions $this })

    Set-IntuneAssignmentsProviderOptions $script:cbIntuneAssignmentsProvider

    $script:UIProvider.AddXamlEvent($panel, "btnGetIntuneAssignments", "Add_Click", {
        Start-IntuneAssignmentsLoad
    })

    $script:UIProvider.AddXamlEvent($panel, "btnIntuneAssignmentsCopy", "Add_Click", {
        $items = @(Get-IntuneToolsDataGridVisibleItems $script:dgIntuneAssignments)
        if($items.Count -eq 0) { return }
        ($items | Select-Object Name, Type, AssignmentCount, HasFilters, IncludedString, ExcludedString, IncludedFilterString, ExcludedFilterString | ConvertTo-Csv -NoTypeInformation) | Set-Clipboard
    })

    $script:UIProvider.AddXamlEvent($panel, "btnIntuneAssignmentsSave", "Add_Click", {
        $items = @(Get-IntuneToolsDataGridVisibleItems $script:dgIntuneAssignments)
        if($items.Count -eq 0) { return }

        $sf            = [System.Windows.Forms.SaveFileDialog]::new()
        $sf.FileName   = "IntuneAssignments_$((Get-Date).ToString('yyyyMMdd-HHmm')).csv"
        $sf.DefaultExt = "csv"
        $sf.Filter     = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $initialDir    = Get-SettingStoreValue "IntuneAssignments" "LastSaveDirectory"
        if(-not $initialDir) { $initialDir = Get-SettingValue "RootFolder" }
        if($initialDir) { $sf.InitialDirectory = $initialDir }

        if($sf.ShowDialog() -eq "OK")
        {
            Save-SettingStoreValue "IntuneAssignments" "LastSaveDirectory" ([IO.FileInfo]$sf.FileName).DirectoryName
            ($items | Select-Object Name, Type, AssignmentCount, HasFilters, IncludedString, ExcludedString, IncludedFilterString, ExcludedFilterString | ConvertTo-Csv -NoTypeInformation) | Out-File -LiteralPath $sf.FileName -Force -Encoding UTF8
        }
    })

    $script:UIProvider.AddXamlEvent($panel, "txtIntuneAssignmentsFilter", "Add_TextChanged", {
        Update-IntuneAssignmentsFilter
    })

    if($script:cbIntuneAssignmentsView)
    {
        $script:cbIntuneAssignmentsView.ItemsSource = @(Get-IntuneAssignmentViewOptions)
        $script:cbIntuneAssignmentsView.SelectedIndex = 0
        $script:cbIntuneAssignmentsView.Add_SelectionChanged({ Update-IntuneAssignmentsFilter })
    }

    if(@($script:intuneAssignmentsRows).Count -gt 0 -and $script:dgIntuneAssignments) {
        $script:dgIntuneAssignments.ItemsSource = $script:intuneAssignmentsRows
        Update-IntuneAssignmentsFilter
    }

    $script:grdToolsMain.Children.Add($panel) | Out-Null
}

function Set-IntuneAssignmentsProviderOptions
{
    param($Control)

    $Provider = $Control.SelectedItem
    if(-not $Provider) { return }

    Save-SettingStoreValue "IntuneAssignments" "Provider" $Provider.Value

    $providerPanel = $null
    if($Provider.OptionsXaml)
    {
        if($script:intuneAssignmentsProviderPanelCache.ContainsKey($Provider.Value))
        {
            $providerPanel = $script:intuneAssignmentsProviderPanelCache[$Provider.Value]
        }
        else
        {
            $providerPanel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\$($Provider.OptionsXaml).xaml"))
            if($providerPanel)
            {
                $script:intuneAssignmentsProviderPanelCache[$Provider.Value] = $providerPanel
                Register-IntuneAssignmentsProviderEvents $providerPanel $Provider
            }
            else
            {
                Write-Log "Failed to load provider options '$($Provider.OptionsXaml)'" 3
            }
        }

        if($providerPanel)
        {
            $providerPanel.DataContext = $Provider
        }
    }

    if($script:ccIntuneAssignmentsProviderOptions)
    {
        $script:ccIntuneAssignmentsProviderOptions.Content    = $providerPanel
        $script:ccIntuneAssignmentsProviderOptions.Visibility = if($null -eq $providerPanel) { "Collapsed" } else { "Visible" }
    }
}

function Register-IntuneAssignmentsProviderEvents
{
    param($Panel, $Provider)

    if($Provider -is [IntuneAssignmentsFolderProvider])
    {
        $Provider.ExportPath = Get-SettingStoreValue "IntuneAssignments" "ExportPath"

        $browse = $Panel.FindName("browseIntuneAssignmentsExportPath")
        if($browse)
        {
            $browse.Tag = @{ Provider = $Provider; Panel = $Panel }
            $browse.Add_Click({
                $tag    = $this.Tag
                $folder = $script:UIProvider.ShowFolderPicker($tag.Provider.ExportPath, "Select root folder for exported objects")
                if($folder)
                {
                    $tag.Provider.ExportPath = $folder
                    $dc = $tag.Panel.DataContext
                    $tag.Panel.DataContext = $null
                    $tag.Panel.DataContext = $dc
                }
            })
        }
    }
}

function Start-IntuneAssignmentsLoad
{
    if(-not $script:cbIntuneAssignmentsProvider) { return }

    $provider = $script:cbIntuneAssignmentsProvider.SelectedItem
    if(-not $provider)
    {
        $script:UIProvider.ShowMessageBox("Select an input source first", "Intune Assignments", "OK", "Error")
        return
    }

    try
    {
        if(-not $provider.Validate()) { return }

        $provider.SaveSettings()

        Write-Status "Get Intune assignments"

        $script:dgIntuneAssignments.ItemsSource = $null
        $rows = @($provider.GetAssignments())
        $script:intuneAssignmentsRows = $rows
        $script:dgIntuneAssignments.ItemsSource = $rows

        Update-IntuneAssignmentsFilter
        Update-IntuneAssignmentsSummary
    }
    catch
    {
        $script:UIProvider.ShowMessageBox($_.Exception.Message, "Intune Assignments", "OK", "Error")
    }
    finally
    {
        Write-Status ""
    }
}

function Get-IntuneAssignmentsSelectedView
{
    # Value of the View combo, defaulting to All before the panel is built.
    if(-not $script:cbIntuneAssignmentsView) { return "All" }
    $value = "$($script:cbIntuneAssignmentsView.SelectedValue)"
    if([string]::IsNullOrEmpty($value)) { return "All" }
    return $value
}

function Update-IntuneAssignmentsFilter
{
    if(-not $script:dgIntuneAssignments -or -not $script:dgIntuneAssignments.Items) { return }
    if(-not $script:txtIntuneAssignmentsFilter) { return }

    $text = "$($script:txtIntuneAssignmentsFilter.Text)".Trim()
    $view = Get-IntuneAssignmentsSelectedView

    if([string]::IsNullOrEmpty($text) -and $view -eq "All")
    {
        $script:dgIntuneAssignments.Items.Filter = $null
    }
    else
    {
        # Read text dynamically inside the predicate so it picks up the current filter on each refresh.
        $script:dgIntuneAssignments.Items.Filter = [Predicate[object]]{
            param($item)
            if(-not (Test-IntuneAssignmentViewMatch $item (Get-IntuneAssignmentsSelectedView))) { return $false }
            $current = "$($script:txtIntuneAssignmentsFilter.Text)".Trim()
            if([string]::IsNullOrEmpty($current)) { return $true }
            $pattern = [regex]::Escape($current)
            return (
                ($item.Name           -and $item.Name           -match $pattern) -or
                ($item.Type           -and $item.Type           -match $pattern) -or
                ($item.IncludedString -and $item.IncludedString -match $pattern) -or
                ($item.ExcludedString -and $item.ExcludedString -match $pattern) -or
                ($item.IncludedFilterString -and $item.IncludedFilterString -match $pattern) -or
                ($item.ExcludedFilterString -and $item.ExcludedFilterString -match $pattern) -or
                ($item.HasFilters -and "Has filters" -match $pattern)
            )
        }
    }

    Update-IntuneAssignmentsSummary
}

function Update-IntuneAssignmentsSummary
{
    if(-not $script:txtIntuneAssignmentsCount -or -not $script:dgIntuneAssignments) { return }
    $total    = @($script:intuneAssignmentsRows).Count
    $visible  = @($script:dgIntuneAssignments.Items).Count
    if($total -eq $visible)
    {
        $script:txtIntuneAssignmentsCount.Text = "$total objects"
    }
    else
    {
        $script:txtIntuneAssignmentsCount.Text = "$visible of $total objects"
    }
}

# ─── Intune Filter Usage tool ─────────────────────────────────────────────────
#
# Surfaces a where-used view for assignment filters. Each Intune assignment
# filter exposes a /payloads collection — one entry per (policy, group,
# include|exclude) combination that references the filter. We fan that out
# into a flat table so the user can answer "if I change this filter, what
# policies are affected?" without clicking through every policy in the portal.
#
# Migration notes from the old Extensions/IntuneFilterUsage.psm1:
#   - Old used the undefined Invoke-GraphRequest cmdlet; replaced with
#     Invoke-MSGraphAPI everywhere.
#   - Old used the legacy positional Invoke-GraphBatchRequest signature; the
#     current helper takes [List[PSCustomObject]] via -BatchObjects.
#   - Old XAML hardcoded foreground/background colors and had a placeholder
#     "Filter" text manipulated via Foreground/FontStyle/Tag; replaced with
#     the txt*.Items.Filter predicate pattern from the Assignments tool.
#   - "All Devices" / "All Users" virtual groups don't resolve via /groups;
#     they're handled with a pre-seeded lookup table (matches old behavior).

function Show-IntuneFilterUsageTool
{
    if(-not $script:grdToolsMain) { return }

    $script:grdToolsMain.Children.Clear()

    $panel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\IntuneToolsFilterUsage.xaml"), $true)
    if(-not $panel)
    {
        Write-Log "Failed to load IntuneToolsFilterUsage.xaml" 3
        return
    }

    $script:intuneFilterUsagePanel = $panel
    if(-not $script:intuneFilterUsageRows) { $script:intuneFilterUsageRows = @() }

    # Cache controls by FindName once so all handlers and helpers share the
    # same references — avoids repeat FindName probes on every keystroke.
    $script:dgIntuneFilterUsage        = $panel.FindName("dgIntuneFilterUsage")
    $script:txtIntuneFilterUsageFilter = $panel.FindName("txtIntuneFilterUsageFilter")
    $script:txtIntuneFilterUsageCount  = $panel.FindName("txtIntuneFilterUsageCount")
    $script:btnGetIntuneFilterUsage    = $panel.FindName("btnGetIntuneFilterUsage")

    if(-not $script:dgIntuneFilterUsage -or -not $script:btnGetIntuneFilterUsage) {
        Write-Log "Intune Filter Usage XAML loaded, but required controls were not found" 3
        return
    }

    $script:UIProvider.AddXamlEvent($panel, "btnGetIntuneFilterUsage", "Add_Click", {
        Start-IntuneFilterUsageLoad
    })

    $script:UIProvider.AddXamlEvent($panel, "btnIntuneFilterUsageCopy", "Add_Click", {
        $items = @(Get-IntuneToolsDataGridVisibleItems $script:dgIntuneFilterUsage)
        if($items.Count -eq 0) { return }
        ($items | Select-Object FilterName, PolicyName, PayloadType, Mode, GroupName | ConvertTo-Csv -NoTypeInformation) | Set-Clipboard
    })

    $script:UIProvider.AddXamlEvent($panel, "btnIntuneFilterUsageSave", "Add_Click", {
        $items = @(Get-IntuneToolsDataGridVisibleItems $script:dgIntuneFilterUsage)
        if($items.Count -eq 0) { return }

        $sf            = [System.Windows.Forms.SaveFileDialog]::new()
        $sf.FileName   = "IntuneFilterUsage_$((Get-Date).ToString('yyyyMMdd-HHmm')).csv"
        $sf.DefaultExt = "csv"
        $sf.Filter     = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $initialDir    = Get-SettingStoreValue "IntuneFilterUsage" "LastSaveDirectory"
        if(-not $initialDir) { $initialDir = Get-SettingValue "RootFolder" }
        if($initialDir) { $sf.InitialDirectory = $initialDir }

        if($sf.ShowDialog() -eq "OK")
        {
            Save-SettingStoreValue "IntuneFilterUsage" "LastSaveDirectory" ([IO.FileInfo]$sf.FileName).DirectoryName
            ($items | Select-Object FilterName, PolicyName, PayloadType, Mode, GroupName | ConvertTo-Csv -NoTypeInformation) | Out-File -LiteralPath $sf.FileName -Force -Encoding UTF8
        }
    })

    $script:UIProvider.AddXamlEvent($panel, "txtIntuneFilterUsageFilter", "Add_TextChanged", {
        Update-IntuneFilterUsageFilter
    })

    if(@($script:intuneFilterUsageRows).Count -gt 0 -and $script:dgIntuneFilterUsage) {
        $script:dgIntuneFilterUsage.ItemsSource = $script:intuneFilterUsageRows
        Update-IntuneFilterUsageFilter
    }

    $panel.HorizontalAlignment = "Stretch"
    $panel.VerticalAlignment   = "Stretch"
    $script:grdToolsMain.Children.Add($panel) | Out-Null
}

function Start-IntuneFilterUsageLoad
{
    if(-not $script:dgIntuneFilterUsage) { return }

    # Disable the Get button during the load so double-clicks can't fire a
    # parallel pass. A concurrent run would reset $script:_intuneEnrollmentConfigCache
    # mid-flight and risk a null-deref in the first pass. try/finally guarantees
    # re-enable even on exception.
    if($script:btnGetIntuneFilterUsage) { $script:btnGetIntuneFilterUsage.IsEnabled = $false }

    try {
        # Reset per-load lazy caches so a re-run picks up fresh server state.
        $script:_intuneEnrollmentConfigCache  = $null
        $script:_intuneManagedAppPolicyCache  = $null

        Write-Status "Get Intune filter usage"

        $script:dgIntuneFilterUsage.ItemsSource = $null
        try {
            $rows = @(Get-IntuneFilterUsageData)
        }
        catch {
            Write-LogError "Intune Filter Usage: load failed" $_.Exception
            Write-Status ""
            $script:UIProvider.ShowMessageBox("Failed to load filter usage:`n`n$($_.Exception.Message)", "Intune Filter Usage", "OK", "Error")
            return
        }

        $script:intuneFilterUsageRows         = $rows
        $script:dgIntuneFilterUsage.ItemsSource = $rows

        Update-IntuneFilterUsageFilter
        Update-IntuneFilterUsageSummary

        Write-Status ""
    }
    finally {
        if($script:btnGetIntuneFilterUsage) { $script:btnGetIntuneFilterUsage.IsEnabled = $true }
    }
}

function Update-IntuneFilterUsageFilter
{
    if(-not $script:dgIntuneFilterUsage -or -not $script:dgIntuneFilterUsage.Items) { return }
    if(-not $script:txtIntuneFilterUsageFilter) { return }

    $text = "$($script:txtIntuneFilterUsageFilter.Text)".Trim()

    if([string]::IsNullOrEmpty($text))
    {
        $script:dgIntuneFilterUsage.Items.Filter = $null
    }
    else
    {
        # Same dynamic-read pattern as the Assignments tool — the predicate
        # re-reads the textbox on every refresh so it tracks edits.
        $script:dgIntuneFilterUsage.Items.Filter = [Predicate[object]]{
            param($item)
            $current = "$($script:txtIntuneFilterUsageFilter.Text)".Trim()
            if([string]::IsNullOrEmpty($current)) { return $true }
            $pattern = [regex]::Escape($current)
            return (
                ($item.FilterName  -and $item.FilterName  -match $pattern) -or
                ($item.PolicyName  -and $item.PolicyName  -match $pattern) -or
                ($item.PayloadType -and $item.PayloadType -match $pattern) -or
                ($item.Mode        -and $item.Mode        -match $pattern) -or
                ($item.GroupName   -and $item.GroupName   -match $pattern)
            )
        }
    }

    Update-IntuneFilterUsageSummary
}

function Update-IntuneFilterUsageSummary
{
    if(-not $script:txtIntuneFilterUsageCount -or -not $script:dgIntuneFilterUsage) { return }
    $total   = @($script:intuneFilterUsageRows).Count
    $visible = @($script:dgIntuneFilterUsage.Items).Count
    if($total -eq $visible) {
        $script:txtIntuneFilterUsageCount.Text = "$total rows"
    }
    else {
        $script:txtIntuneFilterUsageCount.Text = "$visible of $total rows"
    }
}

# ─── ADMX Import tool (Phase 1: scaffold) ─────────────────────────────────────
#
# Migrated FROM Extensions/IntuneTools.psm1 in the original IntuneManagement
# project. Phase 1 wires the panel into the Intune Tools view: the XAML loads,
# the Load ADMX / Load ADML file dialogs work and save the last-used path, and
# the policy-name "Add Random" button is fully functional. The actual ADMX/ADML
# parsing, category tree building, setting properties dialog, and Intune
# ingestion are stubbed with a warning — Phase 2 ports them.
#
# Migration rules applied here (and to be applied to Phase 2 / 3):
#   - $global:<ctrl> from -AddVariables -> $script:<tool><Ctrl> set via FindName.
#   - Invoke-GraphRequest (undefined in current project) -> Invoke-MSGraphAPI
#     with explicit -TokenId.
#   - Hardcoded foreground/background colors -> theme via DynamicResource.

function Show-ADMXImportTool
{
    if(-not $script:grdToolsMain) { return }

    $script:grdToolsMain.Children.Clear()

    $panel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\IntuneToolsADMX.xaml"))
    if(-not $panel)
    {
        Write-Log "Failed to load IntuneToolsADMX.xaml" 3
        return
    }

    # Cache panel + controls used by event handlers. FindName once at panel-init
    # so click/select handlers don't re-probe the visual tree.
    $script:admxImportPanel              = $panel
    $script:btnADMXLoadADMX              = $panel.FindName("btnADMXLoadADMX")
    $script:btnADMXLoadADML              = $panel.FindName("btnADMXLoadADML")
    $script:btnADMXImport                = $panel.FindName("btnADMXImport")
    $script:btnADMXPolicyNameRandom      = $panel.FindName("btnADMXPolicyNameRandom")
    $script:tvADMXCategories             = $panel.FindName("tvADMXCategories")
    $script:dgADMXCategoryPolicies       = $panel.FindName("dgADMXCategoryPolicies")
    $script:txtADMXProfileName           = $panel.FindName("txtADMXProfileName")
    $script:txtADMXProfileDescription    = $panel.FindName("txtADMXProfileDescription")
    $script:txtADMXPolicyFileName        = $panel.FindName("txtADMXPolicyFileName")
    $script:txtADMXPolicyAppName         = $panel.FindName("txtADMXPolicyAppName")
    $script:txtADMXPolicyIngestName      = $panel.FindName("txtADMXPolicyIngestName")
    $script:chkADMXPolicyIngest          = $panel.FindName("chkADMXPolicyIngest")

    # Per-tool state set during Load ADMX. Initialised here so the file
    # dialogs can write to it without a null-check inside the click handler.
    $script:currentADMXFile              = $null
    $script:_admxImportEventsRegistered = $false

    # Load ADMX — fully functional in Phase 1: opens a file dialog, persists
    # the last-used path, and delegates to the (stubbed) parser. Phase 2 fills
    # in Start-ADMXLoadFile to actually parse the file and populate the tree.
    $script:UIProvider.AddXamlEvent($panel, "btnADMXLoadADMX", "Add_Click", {
        $of = [System.Windows.Forms.OpenFileDialog]::new()
        $of.Multiselect = $false
        $of.Filter      = "ADMX Files (*.admx)|*.admx"
        $of.FileName    = Get-SettingStoreValue "Tools" "ADMXLastADMXFile"
        if($of.ShowDialog() -eq "OK")
        {
            $script:currentADMXFile = [IO.FileInfo]$of.FileName
            Save-SettingStoreValue "Tools" "ADMXLastADMXFile" $of.FileName
            Write-Status "Loading policy settings from $($script:currentADMXFile.Name)"
            Start-ADMXLoadFile $of.FileName
            Write-Status ""
        }
    })

    $script:UIProvider.AddXamlEvent($panel, "btnADMXLoadADML", "Add_Click", {
        $of = [System.Windows.Forms.OpenFileDialog]::new()
        $of.Multiselect = $false
        $of.Filter      = "ADML Files (*.adml)|*.adml"
        $of.FileName    = Get-SettingStoreValue "Tools" "ADMXLastADMLFile"
        if($of.ShowDialog() -eq "OK")
        {
            Save-SettingStoreValue "Tools" "ADMXLastADMLFile" $of.FileName
            Write-Status "Loading ADML policy $($of.FileName)"
            Invoke-ADMXLoadSettings $of.FileName
            Write-Status ""
        }
    })

    $script:UIProvider.AddXamlEvent($panel, "btnADMXImport", "Add_Click", {
        Write-Status "Import ADMX policy"
        Import-ADMXPolicyToIntune
        Write-Status ""
    })

    # Append-or-replace GUID suffix on the policy file name. The trailing
    # segment after "_" is the unique-id portion; if it's already a 32-char
    # GUID we leave the existing prefix alone and replace only the suffix.
    $script:UIProvider.AddXamlEvent($panel, "btnADMXPolicyNameRandom", "Add_Click", {
        $guid = [Guid]::NewGuid().Guid
        if($script:txtADMXPolicyFileName.Text)
        {
            $parts = $script:txtADMXPolicyFileName.Text.Split('_')
            if($parts[-1].Length -eq $guid.Length) {
                # Already has a guid suffix — swap it.
                $script:txtADMXPolicyFileName.Text = (($parts[0..($parts.Count-2)] -join '_') + "_" + $guid)
            } else {
                $script:txtADMXPolicyFileName.Text = ($script:txtADMXPolicyFileName.Text + "_" + $guid)
            }
        }
        else { $script:txtADMXPolicyFileName.Text = $guid }
    })

    $script:grdToolsMain.Children.Add($panel) | Out-Null
}

# ─── ADMX Import — helpers (data layer, pure functions) ───────────────────────

function Add-ADMXCategoriesRecursive
{
    # Build a PSCustomObject tree of categories (matches old Add-ADMXCategories).
    # Adds branches into $Parent.Children, deduping on internal category name.
    param($Categories, $Parent, $SettingClass)

    foreach($cat in $Categories.ref)
    {
        $catPath = Get-ADMXCategoryIdPath $cat
        $tvObj = $Parent
        foreach($catName in $catPath.Split('/'))
        {
            $curParent = $tvObj.Children | Where-Object { $_.CategoryNode.name -eq $catName }
            if(-not $curParent)
            {
                $catNode = $script:_admxXML.policyDefinitions.categories.SelectSingleNode("$($script:_admxNSPrefix)category[@name='$catName']", $script:_admxNS)
                $curParent = [PSCustomObject]@{
                    Name         = Get-ADMXADMLString $catNode
                    CategoryName = $catName
                    CategoryNode = $catNode
                    SettingClass = $SettingClass
                    Children     = @()
                }
                $tvObj.Children += $curParent
            }
            $tvObj = $curParent
        }
    }
}

function Add-ADMXCategoryTreeNodeRecursive
{
    # Render a PSCustomObject category tree into a WPF TreeView.
    param($Obj, $Parent)

    $tvItem = [System.Windows.Controls.TreeViewItem]::new()
    $tvItem.Header = $Obj.Name
    $tvItem.Tag    = $Obj
    [void]$Parent.Items.Add($tvItem)

    $Obj.Children | Sort-Object -Property Name | ForEach-Object {
        Add-ADMXCategoryTreeNodeRecursive $_ $tvItem
    }
}

# ─── ADMX Import — loaders (replace Phase 1 stubs) ────────────────────────────

function Start-ADMXLoadFile
{
    param($FileName)

    # Reset all per-load state. Treated as a hard reset on every Load ADMX —
    # the old behavior, and the only safe contract because partial state
    # from a previous file would cross-contaminate the tree-building below.
    $script:_admxNS               = $null
    $script:_admxNSPrefix         = $null
    $script:_admxXML              = $null
    $script:_admxPolicies         = @()
    $script:_admxPoliciesHT       = @{}
    $script:_admxSupportedOn      = @()
    $script:_admxCategoryPaths    = @{}
    $script:_admxStringTable      = @{}
    $script:_admxlngADML          = $null

    if($script:txtADMXProfileName)        { $script:txtADMXProfileName.Text        = "" }
    if($script:txtADMXProfileDescription) { $script:txtADMXProfileDescription.Text = "" }
    if($script:txtADMXPolicyFileName)     { $script:txtADMXPolicyFileName.Text     = "" }
    if($script:txtADMXPolicyIngestName)   { $script:txtADMXPolicyIngestName.Text   = "" }
    if($script:txtADMXPolicyAppName)      { $script:txtADMXPolicyAppName.Text      = "" }
    if($script:btnADMXLoadADML)           { $script:btnADMXLoadADML.IsEnabled      = $false }

    $admxFI = [IO.FileInfo]$FileName
    if(-not $admxFI.Exists) {
        $script:UIProvider.ShowMessageBox("ADMX file not found: $FileName", "ADMX Import", "OK", "Error")
        return
    }

    # Convention: paired .adml lives in en-US\ next to the .admx, or sometimes
    # in the same folder. Either path is acceptable. Missing ADML is a soft
    # warning — Get-ADMXADMLString falls back to raw attributes if so.
    $admlFile = [IO.Path]::Combine($admxFI.DirectoryName, "en-US\$($admxFI.BaseName).adml")
    if(-not [IO.File]::Exists($admlFile)) {
        $admlFile = [IO.Path]::Combine($admxFI.DirectoryName, "$($admxFI.BaseName).adml")
    }
    if(-not [IO.File]::Exists($admlFile)) {
        Write-Log "Could not find an ADML file alongside $FileName" 2
        $admlFile = $null
    }

    try {
        Write-Log "Load ADMX file $FileName"
        [xml]$script:_admxXML = [IO.File]::ReadAllText($FileName)

        $namespace = $script:_admxXML.DocumentElement.NamespaceURI
        if($namespace) {
            $script:_admxNS = New-Object System.Xml.XmlNamespaceManager($script:_admxXML.NameTable)
            $script:_admxNS.AddNamespace("ns", $namespace)
            $script:_admxNSPrefix = "ns:"
        }
        else {
            $script:_admxNS       = $null
            $script:_admxNSPrefix = ""
        }

        # ADMX policy id is derived from the policyNamespaces/target prefix.
        # Used as the default value for the App Id text box.
        $prefix = $script:_admxXML.policyDefinitions.policyNamespaces.SelectSingleNode("$($script:_admxNSPrefix)target[@prefix]", $script:_admxNS)
        $policyId = $null
        if($prefix) {
            if($prefix.namespace) { $policyId = $prefix.namespace.Split('.')[-1] }
            else                  { $policyId = $prefix.prefix }
        }
        if(-not $policyId) {
            $policyId = $admxFI.BaseName -replace " ", ""
            Write-Log "Failed to get policy id from ADMX namespace; using file base name" 2
        }
        if($script:txtADMXPolicyAppName)  { $script:txtADMXPolicyAppName.Text  = $policyId }
        if($script:txtADMXPolicyFileName) { $script:txtADMXPolicyFileName.Text = $policyId }
    }
    catch {
        Write-LogError "Failed to load ADMX file" $_.Exception
        return
    }

    if($script:btnADMXLoadADML) { $script:btnADMXLoadADML.IsEnabled = $true }
    Invoke-ADMXLoadSettings $admlFile
}

function Invoke-ADMXLoadSettings
{
    # Phase 2: read the ADML (string + presentation tables), enumerate every
    # <policy> in the ADMX, and project to PSCustomObjects the UI binds to.
    # Then build the category tree as Device Configuration / User Configuration
    # roots with category sub-trees underneath each.
    param($AdmlFile)

    if($AdmlFile -and [IO.File]::Exists($AdmlFile)) {
        try {
            Write-Log "Load ADML file $AdmlFile"
            [xml]$script:_admxlngADML = [IO.File]::ReadAllText($AdmlFile)
        }
        catch {
            Write-LogError "Failed to load ADML file $AdmlFile" $_.Exception
        }
    }

    $script:_admxStringTable = @{}
    foreach($strNode in $script:_admxlngADML.policyDefinitionResources.resources.stringTable.string) {
        if(-not $script:_admxStringTable.ContainsKey($strNode.id)) {
            $script:_admxStringTable.Add($strNode.id, $strNode.'#text')
        }
    }

    # supportedOn table — used in the setting properties dialog to display the
    # "Supported on" line. Augmented with Windows.adml's SUPPORTED_* entries
    # so policies that reference common Windows constants still resolve.
    $script:_admxSupportedOn = @()
    foreach($polObj in $script:_admxXML.policyDefinitions.supportedOn.definitions.definition) {
        $script:_admxSupportedOn += [PSCustomObject]@{
            Id          = $polObj.Name
            DisplayName = (Get-ADMXADMLString $polObj)
        }
    }
    if($script:_admxWindowsADML) {
        foreach($winString in ($script:_admxWindowsADML.policyDefinitionResources.resources.stringTable.string | Where-Object Id -like "SUPPORTED_*")) {
            $script:_admxSupportedOn += [PSCustomObject]@{
                Id          = $winString.id
                DisplayName = $winString.'#text'
            }
        }
    }

    $devicePolicies = @()
    $userPolicies   = @()

    foreach($polObj in $script:_admxXML.policyDefinitions.policies.policy)
    {
        $category    = if($polObj.parentCategory.ref) { Get-ADMXCategoryNamePath $polObj.parentCategory.ref "/" } else { $null }
        $displayName = Get-ADMXADMLString $polObj
        $description = Get-ADMXADMLString $polObj "explainText"

        # Class can be Both / Machine / User. "Both" expands into two virtual
        # entries so it shows up in both Device + User configuration trees.
        $classArr = switch ($polObj.Class) {
            "Both"    { @("Device","User") }
            "Machine" { @("Device") }
            default   { @($polObj.Class) }
        }

        foreach($class in $classArr)
        {
            $key = "$($polObj.Name)_$class"
            if($script:_admxPoliciesHT.ContainsKey($key)) { continue }   # already seen (re-load ADML on existing tree)

            $newSetting = [PSCustomObject]@{
                Name              = $displayName
                Description       = $description
                OMAURIName        = $null
                OMAURIDescription = $null
                Category          = $category
                CategoryId        = $polObj.parentCategory.ref
                Id                = $polObj.Name
                Definition        = $polObj
                SettingStatus     = $null
                SettingStatusText = $null
                PolicySettings    = $null
                PolicyDefinition  = $null
                ElementsPanel     = $null
                ManualConfig      = $false
                SettingClass      = $class
                SupportedOn       = $null
            }
            $script:_admxPoliciesHT.Add($key, $newSetting)
            $script:_admxPolicies   += $newSetting
            if($class -eq "User") { $userPolicies   += $newSetting }
            else                  { $devicePolicies += $newSetting }
        }
    }

    $script:_admxPolicies | ForEach-Object { Set-ADMXSettingStatusText $_ }
    $script:_admxPolicies = $script:_admxPolicies | Sort-Object -Property Name

    if($script:tvADMXCategories) { $script:tvADMXCategories.Items.Clear() }

    $treeItems = @()

    # Device Configuration root
    $tvItem = [PSCustomObject]@{ Name = "Computer Configuration"; Children = @() }
    if($script:_admxXML.policyDefinitions.policies)
    {
        $policies = $script:_admxXML.policyDefinitions.policies.SelectNodes("$($script:_admxNSPrefix)policy[@class = 'Both' or @class = 'Machine']", $script:_admxNS)
        if($policies) {
            $categories = $policies.parentCategory | Select-Object ref -Unique
            Add-ADMXCategoriesRecursive $categories $tvItem "Device"
        }
    }
    $treeItems += $tvItem

    # User Configuration root
    $tvItem = [PSCustomObject]@{ Name = "User Configuration"; Children = @() }
    if($script:_admxXML.policyDefinitions.policies)
    {
        $policies = $script:_admxXML.policyDefinitions.policies.SelectNodes("$($script:_admxNSPrefix)policy[@class = 'Both' or @class = 'User']", $script:_admxNS)
        if($policies) {
            $categories = $policies.parentCategory | Select-Object ref -Unique
            Add-ADMXCategoriesRecursive $categories $tvItem "User"
        }
    }
    $treeItems += $tvItem

    $treeItems | ForEach-Object { Add-ADMXCategoryTreeNodeRecursive $_ $script:tvADMXCategories }

    # "All Policies" leaf under each root — quick way to see everything.
    if($devicePolicies.Count -gt 0 -and $script:tvADMXCategories.Items.Count -ge 1) {
        $tv = [System.Windows.Controls.TreeViewItem]::new()
        $tv.Header = "All Policies"
        $tv.Tag    = $devicePolicies
        $tv | Add-Member -MemberType NoteProperty -Name "AllPolicies" -Value $true
        [void]$script:tvADMXCategories.Items[0].Items.Add($tv)
    }
    if($userPolicies.Count -gt 0 -and $script:tvADMXCategories.Items.Count -ge 2) {
        $tv = [System.Windows.Controls.TreeViewItem]::new()
        $tv.Header = "All Policies"
        $tv.Tag    = $userPolicies
        $tv | Add-Member -MemberType NoteProperty -Name "AllPolicies" -Value $true
        try { [void]$script:tvADMXCategories.Items[1].Items.Add($tv) } catch { }
    }

    # Hook the tree + grid selection / double-click + context menu handlers.
    # Done after the tree is populated so we don't trigger a spurious selection
    # change before the data is in place.
    Register-ADMXImportTreeEvents
}

function Register-ADMXImportTreeEvents
{
    # Idempotent — only registers once per panel lifetime. ScriptBlocks read
    # $script: state, so they work the same across reloads of ADMX files.
    if($script:_admxImportEventsRegistered) { return }
    $script:_admxImportEventsRegistered = $true

    if($script:tvADMXCategories) {
        $script:tvADMXCategories.Add_SelectedItemChanged({
            $sel = $script:tvADMXCategories.SelectedItem
            if(-not $sel) { return }
            if($sel.AllPolicies -eq $true) {
                $script:dgADMXCategoryPolicies.ColumnWidth          = [System.Windows.Controls.DataGridLength]::Auto
                $script:dgADMXCategoryPolicies.Columns[0].Width     = [System.Windows.Controls.DataGridLength]::Auto
                $script:dgADMXCategoryPolicies.Columns[1].Width     = [System.Windows.Controls.DataGridLength]::Auto
                $script:dgADMXCategoryPolicies.Columns[2].Width     = [System.Windows.Controls.DataGridLength]::Auto
                $script:dgADMXCategoryPolicies.Columns[2].Visibility = "Visible"
                $list = $sel.Tag
            }
            else {
                $script:dgADMXCategoryPolicies.ColumnWidth          = [System.Windows.Controls.DataGridLength]"*"
                $script:dgADMXCategoryPolicies.Columns[0].Width     = [System.Windows.Controls.DataGridLength]"10*"
                $script:dgADMXCategoryPolicies.Columns[1].Width     = [System.Windows.Controls.DataGridLength]::Auto
                $script:dgADMXCategoryPolicies.Columns[2].Visibility = "Collapsed"
                $list = @($script:_admxPolicies | Where-Object {
                    $_.CategoryId   -eq $sel.Tag.CategoryName -and
                    $_.SettingClass -eq $sel.Tag.SettingClass
                })
            }
            $script:dgADMXCategoryPolicies.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[object]]::new(@($list))
        })
    }

    if($script:dgADMXCategoryPolicies) {
        $script:dgADMXCategoryPolicies.Add_MouseDoubleClick({
            if(-not $script:dgADMXCategoryPolicies.SelectedItem) { return }
            Show-ADMXSettingPropertiesDialog $script:dgADMXCategoryPolicies.SelectedItem
        })

        $mnu = $script:dgADMXCategoryPolicies.ContextMenu
        if($mnu) {
            $mnu.Add_Opened({
                $edit = $script:dgADMXCategoryPolicies.ContextMenu.Items | Where-Object Name -eq "mnuADMXSettingEdit" | Select-Object -First 1
                if($edit) { $edit.IsEnabled = $null -ne $script:dgADMXCategoryPolicies.SelectedItem }
            })
            $edit = $mnu.Items | Where-Object Name -eq "mnuADMXSettingEdit" | Select-Object -First 1
            if($edit) {
                $edit.Add_Click({
                    if($script:dgADMXCategoryPolicies.SelectedItem) {
                        Show-ADMXSettingPropertiesDialog $script:dgADMXCategoryPolicies.SelectedItem
                    }
                })
            }
        }
    }

    # Pre-load Windows.adml for SUPPORTED_* string resolution (best-effort).
    if(-not $script:_admxWindowsADML) {
        $winADML = "$($env:WinDir)\PolicyDefinitions\en-US\Windows.adml"
        if([IO.File]::Exists($winADML)) {
            try { [xml]$script:_admxWindowsADML = [IO.File]::ReadAllText($winADML) }
            catch { Write-LogDebug "Failed to load Windows.adml: $($_.Exception.Message)" }
        }
    }
}

# ─── ADMX Import — setting properties dialog ──────────────────────────────────

function Show-ADMXSettingPropertiesDialog
{
    param($SettingObj)

    $form = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\IntuneToolsADMXSettingProperties.xaml"))
    if(-not $form) { return }

    $script:_admxSettingsForm                 = $form
    $script:_admxGrdElements                  = $form.FindName("grdADMXElements")
    $script:_admxRbEnabled                    = $form.FindName("rbADMXSettingEnabled")
    $script:_admxRbDisabled                   = $form.FindName("rbADMXSettingDisabled")
    $script:_admxRbNotConfigured              = $form.FindName("rbADMXSettingNotConfigured")
    $script:_admxTxtSettings                  = $form.FindName("txtADMXSettings")
    $script:_admxChkManualConfig              = $form.FindName("chkADMXManualConfig")
    $script:_admxBtnPrev                      = $form.FindName("btnADMXPreviousSetting")
    $script:_admxBtnNext                      = $form.FindName("btnADMXNextSetting")
    $script:_admxTcPolicyConfig               = $form.FindName("tcADMXPolicyConfig")
    $script:_admxTabSettings                  = $form.FindName("tabADMXSettings")

    Set-ADMXSettingProperties $SettingObj

    $script:UIProvider.AddXamlEvent($form, "btnADMXSettingsOK", "Add_Click", {
        Save-ADMXSettings
        $script:UIProvider.ShowModalObject()
    })

    $script:UIProvider.AddXamlEvent($form, "btnADMXPreviousSetting", "Add_Click", {
        if($script:dgADMXCategoryPolicies.SelectedIndex -le 0) { return }
        Save-ADMXSettings
        $script:dgADMXCategoryPolicies.SelectedIndex = $script:dgADMXCategoryPolicies.SelectedIndex - 1
        Set-ADMXSettingButtonsStatus
        Set-ADMXSettingProperties $script:dgADMXCategoryPolicies.SelectedItem
    })

    $script:UIProvider.AddXamlEvent($form, "btnADMXNextSetting", "Add_Click", {
        if($script:dgADMXCategoryPolicies.SelectedIndex -ge ($script:dgADMXCategoryPolicies.ItemsSource.Count - 1)) { return }
        Save-ADMXSettings
        $script:dgADMXCategoryPolicies.SelectedIndex = $script:dgADMXCategoryPolicies.SelectedIndex + 1
        Set-ADMXSettingButtonsStatus
        Set-ADMXSettingProperties $script:dgADMXCategoryPolicies.SelectedItem
    })

    $script:UIProvider.AddXamlEvent($form, "btnADMXSettingsCancel", "Add_Click", {
        if($script:_admxGrdElements) { $script:_admxGrdElements.Children.Clear() }
        $script:_admxSettingsForm = $null
        $script:UIProvider.ShowModalObject()
    })

    $script:_admxRbEnabled.Add_Checked({
        $script:_admxCurItemStatus = 1
        Set-ADMXControlStatus
    })
    $script:_admxRbDisabled.Add_Checked({
        $script:_admxCurItemStatus = 0
        Set-ADMXControlStatus
    })
    $script:_admxRbNotConfigured.Add_Checked({
        $script:_admxCurItemStatus = $null
        Set-ADMXControlStatus
    })

    $script:_admxChkManualConfig.Add_Click({ Set-ADMXControlStatus })

    $script:_admxTcPolicyConfig.Add_SelectionChanged({
        param($src, $e)
        if($e.AddedItems[0] -eq $script:_admxTabSettings) {
            # When switching to the OMA-URI Settings tab, rebuild the string
            # from the live element controls unless the user picked manual mode.
            if($script:dgADMXCategoryPolicies.SelectedItem.ManualConfig -ne 1) {
                $script:_admxTxtSettings.Text = Get-ADMXSettingsString $script:_admxSettingsForm.DataContext
            }
        }
    })

    $script:UIProvider.ShowModalForm($SettingObj.Name, $form, $true)
}

function Set-ADMXSettingButtonsStatus
{
    $script:_admxBtnPrev.IsEnabled = $script:dgADMXCategoryPolicies.SelectedIndex -gt 0
    $script:_admxBtnNext.IsEnabled = $script:dgADMXCategoryPolicies.SelectedIndex -lt ($script:dgADMXCategoryPolicies.ItemsSource.Count - 1)
}

function Set-ADMXSettingProperties
{
    param($SettingObj)

    $script:_admxGrdElements.Children.Clear()
    $script:_admxCurItemStatus = $SettingObj.SettingStatus

    switch ($SettingObj.SettingStatus)
    {
        0       { $script:_admxRbDisabled.IsChecked      = $true }
        1       { $script:_admxRbEnabled.IsChecked       = $true }
        default { $script:_admxRbNotConfigured.IsChecked = $true }
    }

    $script:_admxTxtSettings.Text         = $SettingObj.PolicySettings
    $script:_admxChkManualConfig.IsChecked = $SettingObj.ManualConfig

    if(-not $SettingObj.PolicyDefinition) {
        $SettingObj.PolicyDefinition = Format-XML $SettingObj.Definition.OuterXml
    }

    if(-not $SettingObj.SupportedOn -and $SettingObj.Definition.supportedOn.ref) {
        $supRef = $SettingObj.Definition.supportedOn.ref.Split(':')[-1]
        $supObj = $script:_admxSupportedOn | Where-Object Id -eq $supRef
        if($supObj) { $SettingObj.SupportedOn = $supObj.DisplayName }
    }

    Set-ADMXElementsPanel $SettingObj

    $script:_admxSettingsForm.DataContext = $SettingObj
    Set-ADMXControlStatus
}

function Set-ADMXControlStatus
{
    # The element grid is enabled only when the policy is Enabled AND not in
    # manual-config mode. The text box accepts edits only when Enabled +
    # manual mode (mirror of the old toggle).
    $script:_admxGrdElements.IsEnabled  = ($script:_admxCurItemStatus -eq 1 -and $script:_admxChkManualConfig.IsChecked -eq $false)
    $script:_admxTxtSettings.IsReadOnly = ($script:_admxCurItemStatus -ne 1 -or $script:_admxChkManualConfig.IsChecked -eq $false)
}

function Save-ADMXSettings
{
    if($script:_admxChkManualConfig.IsChecked) {
        $script:_admxSettingsForm.DataContext.PolicySettings = $script:_admxTxtSettings.Text
    }
    else {
        $script:_admxSettingsForm.DataContext.PolicySettings = Get-ADMXSettingsString $script:_admxSettingsForm.DataContext
    }
    $script:_admxSettingsForm.DataContext.ManualConfig  = if($script:_admxChkManualConfig.IsChecked) { 1 } else { 0 }
    $script:_admxSettingsForm.DataContext.SettingStatus = $script:_admxCurItemStatus
    Set-ADMXSettingStatusText $script:_admxSettingsForm.DataContext
    [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:dgADMXCategoryPolicies.ItemsSource).Refresh()
    $script:_admxGrdElements.Children.Clear()
}

function Set-ADMXElementsPanel
{
    # Build the per-element editor UI for a policy. The element layout (and
    # which controls each element type uses) follows the ADMX/ADML schema:
    # https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-gpreg/
    #
    # Each element gets cached on $Item.ElementsPanel so reopening the dialog
    # for the same setting shows the user's previous edits unchanged.
    param($Item)

    if(-not $Item.Definition.elements) { return }

    # ListValue type used for list-box element rows. Defined inline so the
    # generic List<ListValue> binding below works without a separate file.
    if(-not ('ListValue' -as [type])) {
        Invoke-Expression @'
class ListValue { [string]$Key; [string]$Value }
'@
    }

    if(-not $Item.ElementsPanel)
    {
        $grd = [System.Windows.Controls.Grid]::new()
        $presentation = Get-ADMXPresentationNode $Item

        if($presentation)
        {
            foreach($presentationNode in $presentation.ChildNodes)
            {
                $ctrl = $null
                $elementNode = $Item.Definition.elements.ChildNodes | Where-Object id -eq $presentationNode.refId

                if($presentationNode.Label.'#text')    { $stringLabel = $presentationNode.Label.'#text' }
                elseif($presentationNode.Label)        { $stringLabel = $presentationNode.Label }
                else                                   { $stringLabel = $presentationNode.'#text' }

                if($stringLabel -and $presentationNode.LocalName -ne "CheckBox") {
                    $ctrl = [System.Windows.Controls.TextBlock]::new()
                    $ctrl.Text = $stringLabel
                    Add-GridObject $grd $ctrl
                }

                if($presentationNode.LocalName -eq "text") { continue }

                if($ctrl) { $ctrl.Margin = "0,5,0,0" }

                switch ($presentationNode.LocalName)
                {
                    "textbox" {
                        $ctrl = [System.Windows.Controls.TextBox]::new()
                        if($null -ne $presentationNode.defaultValue) { $ctrl.Text = $presentationNode.defaultValue }
                    }
                    { $_ -in @("DecimalTextBox","LongDecimalTextBox") } {
                        $ctrl = Get-NumericUpDownControl $presentationNode.refId (?? $elementNode.minValue 0) (?? $elementNode.maxValue 9999) (?? $elementNode.SpinStep 1)
                        if(-not $ctrl) { continue }
                        if($null -ne $presentationNode.defaultValue) { $ctrl.Children[0].Text = $presentationNode.defaultValue }
                    }
                    "multiText" {
                        $ctrl = [System.Windows.Controls.TextBox]::new()
                        $ctrl.Height        = 100
                        $ctrl.AcceptsReturn = $true
                    }
                    "CheckBox" {
                        $ctrl = [System.Windows.Controls.CheckBox]::new()
                        $ctrl.Content = $stringLabel
                        if($presentationNode.defaultChecked -eq $true) { $ctrl.IsChecked = $true }
                    }
                    { $_ -in @("ComboBox","DropdownList") } {
                        $ctrl = [System.Windows.Controls.ComboBox]::new()
                        $ctrl.DisplayMemberPath = "Name"
                        $ctrl.SelectedValuePath = "Value"

                        $valItems = @()
                        foreach($valItem in $elementNode.ChildNodes) {
                            $displayName = Get-ADMXADMLString $valItem
                            $value = $null
                            if($valItem.value.decimal.value)       { $value = $valItem.value.decimal.value }
                            elseif($valItem.value.longDecimal)     { $value = (?? $valItem.value.longDecimal.'#text' $valItem.value.longDecimal) }
                            elseif($valItem.value.string)          { $value = (?? $valItem.value.string.'#text'      $valItem.value.string) }
                            else {
                                Write-Log "Unsupported value type for $($elementNode.Id): $($valItem.value.InnerXml)" 2
                                $value = "<SET MANUALLY!!!>"
                            }
                            $valItems += [PSCustomObject]@{ Name = $displayName; Value = $value }
                        }
                        if($null -ne $presentationNode.defaultItem) {
                            try { $ctrl.SelectedIndex = $presentationNode.defaultItem } catch { }
                        }
                        if($presentationNode.NoSort -ne "true") { $valItems = $valItems | Sort-Object -Property Name }
                        $ctrl.ItemsSource = $valItems
                    }
                    "listBox" {
                        $ctrl = [System.Windows.Controls.DataGrid]::new()
                        $ctrl.CanUserAddRows      = $true
                        $ctrl.CanUserDeleteRows   = $true
                        $ctrl.CanUserSortColumns  = $false
                        $ctrl.CanUserResizeRows   = $false
                        $ctrl.AutoGenerateColumns = $false
                        $ctrl.ColumnWidth         = [System.Windows.Controls.DataGridLength]"*"

                        $col = [System.Windows.Controls.DataGridTextColumn]::new()
                        $col.Header   = "Value Name"
                        $col.Width    = [System.Windows.Controls.DataGridLength]"1*"
                        $col.Binding  = [System.Windows.Data.Binding]::new("Key")
                        if($elementNode.explicitValue -ne "true") { $col.Visibility = "Collapsed" }
                        $ctrl.Columns.Add($col)

                        $col = [System.Windows.Controls.DataGridTextColumn]::new()
                        $col.Header   = "Value"
                        $col.Width    = [System.Windows.Controls.DataGridLength]"1*"
                        $col.Binding  = [System.Windows.Data.Binding]::new("Value")
                        $ctrl.Columns.Add($col)

                        $ctrl.ItemsSource = [System.Collections.Generic.List[ListValue]]::new()
                    }
                    default {
                        Write-Log "Unsupported presentation control: $($presentationNode.LocalName) (refId: $($presentationNode.refId))" 2
                        continue
                    }
                }

                Add-GridObject $grd $ctrl
                if($presentationNode.refId) {
                    $ctrl.Tag  = $elementNode
                    $ctrl.Name = $presentationNode.refId
                }
            }
        }
        else
        {
            # Fallback: no presentation node defined. Build best-effort controls
            # directly from <elements>. Less polished but works for ADMX files
            # that ship without ADML presentation data.
            Write-Log "No presentation found for $($Item.Definition.Name); building fallback editor" 2
            $i = 0
            foreach($elementNode in $Item.Definition.elements.ChildNodes) {
                try {
                    $ctrl = $null
                    switch ($elementNode.LocalName) {
                        "text"      { $ctrl = [System.Windows.Controls.TextBox]::new() }
                        "multiText" {
                            $ctrl = [System.Windows.Controls.TextBox]::new()
                            $ctrl.Height        = 100
                            $ctrl.AcceptsReturn = $true
                        }
                        "enum" {
                            $ctrl = [System.Windows.Controls.ComboBox]::new()
                            $ctrl.DisplayMemberPath = "Name"
                            $ctrl.SelectedValuePath = "Value"
                            $valItems = @()
                            foreach($valItem in $elementNode.ChildNodes) {
                                $value = $null
                                if($valItem.value.decimal.value) { $value = $valItem.value.decimal.value }
                                elseif($valItem.value.string)    { $value = (?? $valItem.value.string.'#text' $valItem.value.string) }
                                else                              { $value = "<SET MANUALLY!!!>" }
                                $valItems += [PSCustomObject]@{ Name = (Get-ADMXADMLString $valItem); Value = $value }
                            }
                            $ctrl.ItemsSource = $valItems
                        }
                        "list" {
                            $ctrl = [System.Windows.Controls.DataGrid]::new()
                            $ctrl.CanUserAddRows      = $true
                            $ctrl.CanUserDeleteRows   = $true
                            $ctrl.CanUserSortColumns  = $false
                            $ctrl.CanUserResizeRows   = $false
                            $ctrl.AutoGenerateColumns = $false
                            $ctrl.ColumnWidth         = [System.Windows.Controls.DataGridLength]"*"
                            $col = [System.Windows.Controls.DataGridTextColumn]::new()
                            $col.Header  = "Value Name"
                            $col.Width   = [System.Windows.Controls.DataGridLength]"1*"
                            $col.Binding = [System.Windows.Data.Binding]::new("Key")
                            if($elementNode.explicitValue -ne "true") { $col.Visibility = "Collapsed" }
                            $ctrl.Columns.Add($col)
                            $col = [System.Windows.Controls.DataGridTextColumn]::new()
                            $col.Header  = "Value"
                            $col.Width   = [System.Windows.Controls.DataGridLength]"1*"
                            $col.Binding = [System.Windows.Data.Binding]::new("Value")
                            $ctrl.Columns.Add($col)
                            $ctrl.ItemsSource = [System.Collections.Generic.List[ListValue]]::new()
                        }
                        "decimal"   { $ctrl = [System.Windows.Controls.TextBox]::new() }
                        "boolean"   { $ctrl = [System.Windows.Controls.ComboBox]::new() }
                        default     {
                            Write-Log "Element type not supported in fallback: $($elementNode.LocalName)" 2
                            continue
                        }
                    }

                    $displayName = Get-ADMXADMLPresentationString $presentation $elementNode
                    if($displayName) {
                        $rd = [System.Windows.Controls.RowDefinition]::new()
                        $rd.Height = [double]::NaN
                        $grd.RowDefinitions.Add($rd)
                        $tb = [System.Windows.Controls.TextBlock]::new()
                        if($i -gt 0) { $tb.Margin = "0,5,0,0" }
                        $tb.Text = $displayName
                        $tb.SetValue([System.Windows.Controls.Grid]::RowProperty, $i)
                        [void]$grd.Children.Add($tb)
                        $i++
                    }
                    $rd = [System.Windows.Controls.RowDefinition]::new()
                    $rd.Height = [double]::NaN
                    $grd.RowDefinitions.Add($rd)

                    $ctrl.SetValue([System.Windows.Controls.Grid]::RowProperty, $i)
                    $ctrl.Tag  = $elementNode
                    $ctrl.Name = $elementNode.Id
                    [void]$grd.Children.Add($ctrl)
                    $i++
                }
                catch {
                    Write-LogError "Failed to add ADMX element $($elementNode.LocalName) with id $($elementNode.id)" $_.Exception
                }
            }
        }

        # Trailing row so the grid stretches; harmless if empty.
        [void]$grd.RowDefinitions.Add([System.Windows.Controls.RowDefinition]::new())
        $Item.ElementsPanel = $grd
    }

    if($Item.ElementsPanel) {
        [void]$script:_admxGrdElements.Children.Add($Item.ElementsPanel)
    }
}

function Get-ADMXSettingsString
{
    # Build the OMA-URI value <data id=... value=.../> XML from the live element
    # editor controls. Empty values are skipped (with a log line for required
    # ones). Returns the joined string, or $null if the policy isn't Enabled.
    param($Item)

    if(-not $Item -or -not $Item.Definition.elements) { return }
    if($script:_admxCurItemStatus -ne 1) {
        $Item.PolicySettings = $null
        return
    }

    $policySettings = @()
    foreach($elementNode in $Item.Definition.elements.ChildNodes) {
        $ctrl = [System.Windows.LogicalTreeHelper]::FindLogicalNode($Item.ElementsPanel, $elementNode.Id)
        if(-not $ctrl) {
            Write-Log "Could not find a control with id $($elementNode.Id)" 3
            continue
        }

        $ctrlValue = $null
        switch ($elementNode.LocalName) {
            "text"      { $ctrlValue = $ctrl.Text }
            "multiText" { $ctrlValue = $ctrl.Text -replace [Environment]::NewLine, "&#xF000;" }
            "enum"      { $ctrlValue = $ctrl.SelectedValue }
            "list"      {
                $i = 1
                $arr = @()
                foreach($kv in $ctrl.ItemsSource) {
                    if(-not $kv.Value) { continue }
                    $arr += "$((?? $kv.Key $i))&#xF000;$($kv.Value)"
                    $i++
                }
                $ctrlValue = $arr -join "&#xF000;"
            }
            default {
                # decimal / longDecimal land in the NumericUpDown wrapper grid
                # whose first child is the TextBox holding the value.
                if($elementNode.LocalName -eq "decimal" -or $ctrl.Tag.LocalName -eq "longDecimal") {
                    $ctrlValue = $ctrl.Children[0].Text
                }
                elseif($elementNode.LocalName -eq "boolean") {
                    if($ctrl -is [System.Windows.Controls.CheckBox]) {
                        $ctrlValue = if($ctrl.IsChecked) { "1" } else { "0" }
                    }
                    elseif($ctrl -is [System.Windows.Controls.ComboBox]) {
                        $ctrlValue = $ctrl.SelectedValue
                    }
                    else {
                        Write-Log "Boolean element type not supported: $($elementNode.LocalName)" 2
                        continue
                    }
                }
                else {
                    Write-Log "Element type not supported: $($elementNode.LocalName)" 2
                    continue
                }
            }
        }

        if(-not $ctrlValue) {
            if($elementNode.required -eq $true) {
                Write-Log "Required value is missing for $($elementNode.Id)" 3
            }
            else {
                Write-Log "Value not set for $($elementNode.Id) - value will not be added"
            }
            continue
        }

        $policySettings += "<data id=`"$($elementNode.Id)`" value=`"$ctrlValue`"/>"
    }
    return ($policySettings -join [Environment]::NewLine)
}

# ─── ADMX Import — Intune ingestion ───────────────────────────────────────────

function Import-ADMXPolicyToIntune
{
    if(-not $script:txtADMXProfileName.Text.Trim()) {
        $script:UIProvider.ShowMessageBox("Profile Name must be specified", "ADMX Import", "OK", "Error")
        return
    }
    if(-not $script:txtADMXPolicyFileName.Text.Trim()) {
        $script:UIProvider.ShowMessageBox("ADMX Policy Name must be specified", "ADMX Import", "OK", "Error")
        return
    }

    $configured = @()
    foreach($admxPolicy in @($script:_admxPolicies | Where-Object { $_.SettingStatus -in @(0,1) }))
    {
        if($admxPolicy.SettingStatus -eq 0) {
            $policyValue = "<disabled/>"
        }
        else {
            $policyValue = "<enabled/>"
            if($admxPolicy.PolicySettings) { $policyValue += "`n$($admxPolicy.PolicySettings)" }
        }

        if((Get-SettingValue "FormatOMAURI") -eq $true) {
            $policyValue = Update-XmlFormatting $policyValue
        }

        $catPath   = Get-ADMXCategoryOMAURIPath $admxPolicy.Definition.parentCategory.ref
        $omaUri    = "./$($admxPolicy.SettingClass)/Vendor/MSFT/Policy/Config/$($script:txtADMXPolicyAppName.Text)~Policy~$catPath/$($admxPolicy.Definition.name)"

        $desc =
            if($admxPolicy.OMAURIDescription)                                { $admxPolicy.OMAURIDescription }
            elseif($admxPolicy.Description -and $admxPolicy.Description.Length -gt 1000) { $admxPolicy.Description.Substring(0, 1000) }
            else                                                              { $admxPolicy.Description }

        $configured += [PSCustomObject]@{
            "@odata.type" = "#microsoft.graph.omaSettingString"
            displayName   = (?? $admxPolicy.OMAURIName $admxPolicy.Name)
            description   = $desc
            omaUri        = $omaUri
            value         = $policyValue
        }
    }

    $intuneObj = [PSCustomObject]@{
        "@odata.type"   = "#microsoft.graph.windows10CustomConfiguration"
        displayName     = $script:txtADMXProfileName.Text
        omaSettings     = @()
        roleScopeTagIds = @()
        assignments     = @()
    }
    if($script:txtADMXProfileDescription.Text) {
        $intuneObj | Add-Member -MemberType NoteProperty -Name "description" -Value $script:txtADMXProfileDescription.Text
    }

    if($script:chkADMXPolicyIngest.IsChecked) {
        $xmlString = Format-XML $script:_admxXML
        if((Get-SettingValue "FormatOMAURI") -eq $true) {
            $xmlString = Update-XmlFormatting $xmlString
        }

        $ingestName = if($script:txtADMXPolicyIngestName.Text) { $script:txtADMXPolicyIngestName.Text }
                      else { "$($script:currentADMXFile.Name) Ingestion" }

        $intuneObj.omaSettings += [PSCustomObject]@{
            "@odata.type" = "#microsoft.graph.omaSettingString"
            displayName   = $ingestName
            description   = $null
            omaUri        = "./Device/Vendor/MSFT/Policy/ConfigOperations/ADMXInstall/$($script:txtADMXPolicyAppName.Text)/Policy/$($script:txtADMXPolicyFileName.Text)"
            value         = $xmlString
        }
    }

    $intuneObj.omaSettings += $configured
    $json = $intuneObj | ConvertTo-Json -Depth 20

    $tokenId = Get-DefaultTokenId
    $resp = Invoke-MSGraphAPI -Url "deviceManagement/deviceConfigurations" -Content $json -HttpMethod "POST" -TokenId $tokenId -FullResponseObject

    if($resp -and $resp.Success) {
        Write-Log "Device configuration profile '$($intuneObj.displayName)' created with id $($resp.Content.Id)"
        $script:UIProvider.ShowMessageBox("Profile '$($intuneObj.displayName)' created successfully.", "ADMX Import", "OK", "Information")
    }
    else {
        $errInfo = if($resp) { "$($resp.StatusCode) $($resp.StatusDescription)" } else { "no response" }
        Write-Log "Failed to create device configuration profile '$($intuneObj.displayName)' - $errInfo" 3
        $script:UIProvider.ShowMessageBox("Failed to create device configuration profile '$($intuneObj.displayName)'.`n`n$errInfo`n`nCheck the log for details.", "ADMX Import", "OK", "Error")
    }
}

# ─── Reg Values tool (Phase 1: scaffold) ──────────────────────────────────────
#
# Same migration story as ADMX Import. Phase 1 loads the XAML and wires the
# basic buttons; Phase 3 ports the actual data model (ADMXRegProfile,
# ADMXRegPolicy, ADMXRegPolicyElement classes) + the registry-value editor
# dialog + the Intune ingestion path.

function Show-ADMXRegValuesTool
{
    if(-not $script:grdToolsMain) { return }

    $script:grdToolsMain.Children.Clear()

    # The ADMXReg* C# classes back the bound DataContext below. Lazy-loaded so
    # users who never open this tool don't pay the Add-Type cost.
    Add-ADMXRegClasses

    $panel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\IntuneToolsADMXRegValues.xaml"))
    if(-not $panel)
    {
        Write-Log "Failed to load IntuneToolsADMXRegValues.xaml" 3
        return
    }

    $script:_admxRegValuesPanel      = $panel
    $script:cbADMXRegPolicyType      = $panel.FindName("cbADMXRegPolicyType")
    $script:dgADMXRegAddedPolicies   = $panel.FindName("dgADMXRegAddedPolicies")
    $script:mnuADMXRegPoliciesCtx    = $panel.FindName("mnuADMXRegPoliciesContextMenu")
    $script:mnuADMXRegPolicyEdit     = $panel.FindName("mnuADMXRegPolicyEdit")

    # Profile / hive / status combos are bound by static lists. Safe to wire
    # up-front — no backend dependency.
    $script:cbADMXRegPolicyType.ItemsSource = @(
        [PSCustomObject]@{ Name = "Policy";     Value = "Policy" }
        [PSCustomObject]@{ Name = "Preference"; Value = "Preference" }
    )

    # DataContext is the live ADMXRegProfile — the XAML binds every editable
    # field through it (ProfileName / ProfileDescription / PolicyType /
    # ADMXPolicies). Replacing the DataContext is how "Clear" resets state.
    $panel.DataContext = [ADMXRegProfile]::new()

    $script:UIProvider.AddXamlEvent($panel, "btnADMXRegClear", "Add_Click", {
        $script:_admxRegValuesPanel.DataContext = [ADMXRegProfile]::new()
    })

    $script:UIProvider.AddXamlEvent($panel, "btnADMXAddRegValue", "Add_Click", {
        Show-ADMXRegSettingsDialog
    })

    $script:UIProvider.AddXamlEvent($panel, "btnADMXRegImport", "Add_Click", {
        Write-Status "Import Reg Settings policy"
        Import-ADMXRegProfileToIntune
        Write-Status ""
    })

    # Double-click an added policy to re-open the editor for that policy.
    $script:dgADMXRegAddedPolicies.Add_MouseDoubleClick({
        if(-not $script:dgADMXRegAddedPolicies.SelectedItem) { return }
        Show-ADMXRegSettingsDialog $script:dgADMXRegAddedPolicies.SelectedItem
    })

    if($script:mnuADMXRegPoliciesCtx) {
        $script:mnuADMXRegPoliciesCtx.Add_Opened({
            $script:mnuADMXRegPolicyEdit.IsEnabled = $null -ne $script:dgADMXRegAddedPolicies.SelectedItem
        })
    }
    if($script:mnuADMXRegPolicyEdit) {
        $script:mnuADMXRegPolicyEdit.Add_Click({
            if(-not $script:dgADMXRegAddedPolicies.SelectedItem) { return }
            Show-ADMXRegSettingsDialog $script:dgADMXRegAddedPolicies.SelectedItem
        })
    }

    $script:grdToolsMain.Children.Add($panel) | Out-Null
}

# ─── Reg Values — C# classes (INotifyPropertyChanged for live DataGrid binding)─

# ─── Reg Values — registry-path safety list ───────────────────────────────────
#
# The MDM ADMX-ingest path can't write to certain Microsoft-owned roots. The
# original lists came from:
#   https://docs.microsoft.com/en-us/windows/client-management/mdm/win32-and-centennial-app-policy-configuration
# Carried over verbatim so behaviour matches the old tool.
$script:_admxRegUnsupportedLocations = @(
    'System'
    'Software\Microsoft'
    'Software\Policies\Microsoft'
)
$script:_admxRegUnsupportedOverride  = @(
    'Software\Policies\Microsoft\Office'
    'Software\Microsoft\Office'
    'Software\Microsoft\Windows\CurrentVersion\Explorer'
    'Software\Microsoft\Internet Explorer'
    'software\policies\microsoft\shared tools\proofing tools'
    'software\policies\microsoft\imejp'
    'software\policies\microsoft\ime\shared'
    'software\policies\microsoft\shared tools\graphics filters'
    'software\policies\microsoft\windows\currentversion\explorer'
    'software\policies\microsoft\softwareprotectionplatform'
    'software\policies\microsoft\officesoftwareprotectionplatform'
    'software\policies\microsoft\windows\windows search\preferences'
    'software\policies\microsoft\exchange'
    'software\microsoft\shared tools\proofing tools'
    'software\microsoft\shared tools\graphics filters'
    'software\microsoft\windows\windows search\preferences'
    'software\microsoft\exchange'
    'software\policies\microsoft\vba\security'
    'software\microsoft\onedrive'
    'software\Microsoft\Edge'
    'Software\Microsoft\EdgeUpdate'
)

# Template used to synthesize the ADMX XML payload for the ingest profile. One
# child <policy> is cloned per UI-defined ADMXRegPolicy; the seed entry is then
# removed before the XML is serialised.
$script:_admxRegTemplate = @"
<policyDefinitions revision="1.0" schemaVersion="1.0">
    <categories>
        <category name="RegImport" />
    </categories>
    <policies>
        <policy name="" class="" displayName="" explainText="" presentation="" key="" valueName="">
            <parentCategory ref="RegImport" />
            <supportedOn ref="windows:SUPPORTED_Windows7" />
            <enabledValue>
                <decimal value="1" />
            </enabledValue>
            <disabledValue>
                <decimal value="0" />
            </disabledValue>
            <elements>
            </elements>
        </policy>
    </policies>
</policyDefinitions>
"@

function Get-ADMXRegIsKeySupported
{
    param($RegKey)

    if(-not $RegKey) { return $true }

    $tmpPath = $RegKey.Trim('\') + "\"

    $blockedSource = ""
    foreach($blockedPath in $script:_admxRegUnsupportedLocations) {
        if($tmpPath -like "$blockedPath\*") { $blockedSource = $blockedPath; break }
    }
    if($blockedSource) {
        foreach($exemptPath in $script:_admxRegUnsupportedOverride) {
            if($tmpPath -like "$exemptPath\*") { $blockedSource = ""; break }
        }
    }

    if($blockedSource) {
        $script:UIProvider.ShowMessageBox("The registry key '$RegKey' is not supported.`n`nBlocked by root key: $blockedSource", "Unsupported reg key", "OK", "Error")
        return $false
    }
    return $true
}

# ─── Reg Values — settings dialog ─────────────────────────────────────────────

function Show-ADMXRegSettingsDialog
{
    param($RegPolicy)

    $form = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\IntuneToolsADMXAddRegPolicy.xaml"))
    if(-not $form) { return }

    $script:_admxRegPoliciesForm = $form

    # New vs edit: when invoked from the "Add" button we get $null; double-click
    # on a row passes the existing ADMXRegPolicy.
    $newPolicy = -not $RegPolicy
    if($newPolicy) { $RegPolicy = [ADMXRegPolicy]::new() }
    $script:_admxRegIsNewPolicy = $newPolicy

    $newElement = [ADMXRegPolicyElement]::new()
    $form.DataContext = [PSCustomObject]@{
        RegPolicy     = $RegPolicy
        PolicyElement = $newElement
    }

    # Cache child controls used by event handlers and the show/hide helper.
    $script:_admxRegCbHive            = $form.FindName("cbADMXRegHive")
    $script:_admxRegCbPolicyStatus    = $form.FindName("cbADMXRegPolicyStatus")
    $script:_admxRegCbDataType        = $form.FindName("cbADMXRegElementDataType")
    $script:_admxRegTxtElementKey     = $form.FindName("txtADMXRegElementKey")
    $script:_admxRegTxtElementValName = $form.FindName("txtADMXRegElementValueName")
    $script:_admxRegTxtSeparator      = $form.FindName("txtADMXRegAttributeValueSeparator")
    $script:_admxRegSpSeparator       = $form.FindName("spADMXRegAttributeValueSeparator")
    $script:_admxRegSpValuePrefix     = $form.FindName("spADMXRegAttributeValuePrefix")
    $script:_admxRegTxtValuePrefix    = $form.FindName("txtADMXRegAttributeValuePrefix")
    $script:_admxRegSpExpandable      = $form.FindName("spADMXRegAttributeExpandable")
    $script:_admxRegChkExpandable     = $form.FindName("chkADMXRegAttributeExpandable")
    $script:_admxRegSpAdditive        = $form.FindName("spADMXRegAttributeAdditive")
    $script:_admxRegChkAdditive       = $form.FindName("chkADMXRegAttributeAdditive")
    $script:_admxRegBtnElementAdd     = $form.FindName("btnADMXRegElementAdd")
    $script:_admxRegBtnElementUpdate  = $form.FindName("btnADMXRegElementNew")
    $script:_admxRegTxtKey            = $form.FindName("txtADMXRegKey")
    $script:_admxRegDgElements        = $form.FindName("dgADMXRegAddedElements")

    $script:_admxRegCbHive.ItemsSource = @(
        [PSCustomObject]@{ Name = "HKEY_LOCAL_MACHINE"; Value = "HKLM" }
        [PSCustomObject]@{ Name = "HKEY_CURRENT_USER";  Value = "HKCU" }
    )
    $script:_admxRegCbPolicyStatus.ItemsSource = @(
        [PSCustomObject]@{ Name = "Enabled";  Value = "Enabled"  }
        [PSCustomObject]@{ Name = "Disabled"; Value = "Disabled" }
    )
    # longDecimal omitted intentionally — Intune's ADMX-ingest path doesn't
    # currently accept QWORD. Note carried over from the original tool.
    $script:_admxRegCbDataType.ItemsSource = @(
        [PSCustomObject]@{ Name = "String";          Value = "text"      }
        [PSCustomObject]@{ Name = "Multi-string";    Value = "multiText" }
        [PSCustomObject]@{ Name = "List";            Value = "list"      }
        [PSCustomObject]@{ Name = "DWORD (32-bit)";  Value = "decimal"   }
    )

    # When the user picks a different data type, attribute control visibility
    # follows. Centralised in Set-ADMXRegAttributeControls so the "edit existing
    # element" code path doesn't have to repeat the logic.
    $script:_admxRegCbDataType.Add_SelectionChanged({ Set-ADMXRegAttributeControls })

    # Double-click on an element row swaps the current PolicyElement for the
    # selected one, switching the form from add-mode to update-mode.
    $script:_admxRegDgElements.Add_MouseDoubleClick({
        if(-not $script:_admxRegDgElements.SelectedItem) { return }
        $script:_admxRegBtnElementAdd.Visibility    = "Collapsed"
        $script:_admxRegBtnElementUpdate.Visibility = "Visible"

        $selected = $script:_admxRegDgElements.SelectedItem
        $tmp = $script:_admxRegPoliciesForm.DataContext
        $script:_admxRegPoliciesForm.DataContext = $null
        $tmp.PolicyElement = $selected
        $script:_admxRegPoliciesForm.DataContext = $tmp

        Set-ADMXRegAttributeControls
    })

    $script:UIProvider.AddXamlEvent($form, "btnADMXRegElementAdd", "Add_Click", {
        if($script:_admxRegTxtElementKey.Text -and (Get-ADMXRegIsKeySupported $script:_admxRegTxtElementKey.Text) -eq $false) { return }
        if(-not $script:_admxRegPoliciesForm.DataContext.RegPolicy.Key) {
            $script:UIProvider.ShowMessageBox("The Key value must be specified for the policy", "Reg Values", "OK", "Error")
            return
        }
        $tmp = $script:_admxRegPoliciesForm.DataContext
        $script:_admxRegPoliciesForm.DataContext = $null
        if($script:_admxRegTxtSeparator) {
            $tmp.PolicyElement.AttributeSeparator = $script:_admxRegTxtSeparator.Text
        }
        $tmp.RegPolicy.PolicyElements.Add($tmp.PolicyElement)
        $tmp.PolicyElement = [ADMXRegPolicyElement]::new()
        $script:_admxRegPoliciesForm.DataContext = $tmp
    })

    $script:UIProvider.AddXamlEvent($form, "btnADMXRegElementNew", "Add_Click", {
        $tmp = $script:_admxRegPoliciesForm.DataContext
        $script:_admxRegPoliciesForm.DataContext = $null
        $tmp.PolicyElement = [ADMXRegPolicyElement]::new()
        $script:_admxRegPoliciesForm.DataContext = $tmp
        $script:_admxRegBtnElementAdd.Visibility    = "Visible"
        $script:_admxRegBtnElementUpdate.Visibility = "Collapsed"
    })

    $script:UIProvider.AddXamlEvent($form, "btnADMXRegAddNew", "Add_Click", {
        if((Get-ADMXRegIsKeySupported $script:_admxRegTxtKey.Text) -eq $false) { return }
        if($script:_admxRegIsNewPolicy) {
            $script:_admxRegValuesPanel.DataContext.ADMXPolicies.Add(
                $script:_admxRegPoliciesForm.DataContext.RegPolicy)
        }
        $script:UIProvider.ShowModalObject()
    })

    $script:UIProvider.AddXamlEvent($form, "btnADMXRegCancel", "Add_Click", {
        $script:_admxRegPoliciesForm = $null
        $script:UIProvider.ShowModalObject()
    })

    $script:UIProvider.ShowModalForm("Add new reg policy", $form, $true)
}

function Set-ADMXRegAttributeControls
{
    # Show or hide the per-data-type attribute controls so users only see the
    # toggles that are relevant for the picked DataType. Old behaviour preserved.
    $dt = $script:_admxRegCbDataType.SelectedValue

    $showSep = ($dt -eq "list" -or $dt -eq "multiText")
    $script:_admxRegSpSeparator.Visibility   = if($showSep) { "Visible" } else { "Collapsed" }
    $script:_admxRegTxtSeparator.Visibility  = if($showSep) { "Visible" } else { "Collapsed" }
    if($showSep -and [string]::IsNullOrEmpty($script:_admxRegTxtSeparator.Text)) {
        $script:_admxRegTxtSeparator.Text = ";"
    }

    $showPrefix = ($dt -eq "list")
    $script:_admxRegSpValuePrefix.Visibility  = if($showPrefix) { "Visible" } else { "Collapsed" }
    $script:_admxRegTxtValuePrefix.Visibility = if($showPrefix) { "Visible" } else { "Collapsed" }

    $script:_admxRegTxtElementValName.IsReadOnly = ($dt -eq "list")

    $showExpandable = ($dt -eq "text" -or $dt -eq "list")
    $script:_admxRegSpExpandable.Visibility   = if($showExpandable) { "Visible" } else { "Collapsed" }
    $script:_admxRegChkExpandable.Visibility  = if($showExpandable) { "Visible" } else { "Collapsed" }

    $showAdditive = ($dt -eq "list")
    $script:_admxRegSpAdditive.Visibility     = if($showAdditive) { "Visible" } else { "Collapsed" }
    $script:_admxRegChkAdditive.Visibility    = if($showAdditive) { "Visible" } else { "Collapsed" }
}

# ─── Reg Values — Intune ingestion ────────────────────────────────────────────

function Import-ADMXRegProfileToIntune
{
    $regProfile = $script:_admxRegValuesPanel.DataContext
    if(-not $regProfile -or -not $regProfile.ProfileName) {
        $script:UIProvider.ShowMessageBox("Custom Profile Name must be specified", "Reg Values", "OK", "Error")
        return
    }
    if($regProfile.ADMXPolicies.Count -eq 0) {
        $script:UIProvider.ShowMessageBox("Add at least one registry policy before importing", "Reg Values", "OK", "Information")
        return
    }

    [xml]$xml = $script:_admxRegTemplate

    # The synthesized category needs a unique name so re-imports under the same
    # ingest namespace don't collide.
    $guidId = [Guid]::NewGuid().Guid
    $regPolicyFileName = "RegPolicy_$guidId"
    $xml.policyDefinitions.categories.category.name = ($xml.policyDefinitions.categories.category.name + "_" + $guidId)

    $intuneObj = [PSCustomObject]@{
        "@odata.type"   = "#microsoft.graph.windows10CustomConfiguration"
        displayName     = $regProfile.ProfileName
        omaSettings     = @()
        roleScopeTagIds = @()
        assignments     = @()
    }
    if($regProfile.ProfileDescription) {
        $intuneObj | Add-Member -MemberType NoteProperty -Name description -Value $regProfile.ProfileDescription
    }

    $admxRegSettings = @()

    foreach($regPolicy in $regProfile.ADMXPolicies)
    {
        $omaUriString = if($regPolicy.PolicyStatus -eq "Enabled") { "<enabled/>" } else { "<disabled/>" }

        if($regPolicy.PolicyName) {
            $policyName = $regPolicy.PolicyName -replace " ", "_"
        }
        else {
            $policyName = [Guid]::NewGuid().Guid
        }

        # Clone the seed <policy> for each user-defined policy. We mutate the
        # clone, then append; the original is removed once the loop completes.
        $newNode = $xml.policyDefinitions.policies.ChildNodes[0].CloneNode($true)
        $newNode.name         = $policyName
        $newNode.class        = if($regPolicy.Hive -eq "HKLM") { "Machine" } else { "User" }
        $newNode.displayName  = "`$(string.$policyName)"
        $newNode.presentation = "`$(presentation.$policyName)"
        $newNode.key          = $regPolicy.Key.Trim('\')
        $newNode.parentCategory.ref = $xml.policyDefinitions.categories.category.name

        if(-not $regPolicy.StatusValueName) {
            [void]$newNode.RemoveChild($newNode.SelectSingleNode("enabledValue"))
            [void]$newNode.RemoveChild($newNode.SelectSingleNode("disabledValue"))
        }
        else {
            $newNode.valueName = $regPolicy.StatusValueName
        }

        $omaUriItems = @()
        if($null -eq $regPolicy.PolicyElements -or $regPolicy.PolicyElements.Count -eq 0) {
            [void]$newNode.RemoveChild($newNode.SelectSingleNode("elements"))
        }
        else
        {
            foreach($element in $regPolicy.PolicyElements)
            {
                $child = $xml.CreateElement($element.DataType)
                $elementId = $null

                if($element.DataType -in @("multiText","list")) {
                    # User-set separator is stored per element; fall back to
                    # ';' for legacy in-memory objects.
                    $splitter = ?? $element.AttributeSeparator ";"
                    $value = $element.Value -replace $splitter, "&#xF000;"
                }
                else {
                    $value = $element.Value
                }

                # valueName goes on every element type except "list" (which is
                # keyed by an Id derived from the policy/element key path).
                if($element.DataType -ne "list") {
                    Add-ADMXRegXmlAttribute $child "valueName" $element.ValueName
                }
                else {
                    if($element.AttributePrefix) {
                        Add-ADMXRegXmlAttribute $child "valuePrefix" $element.AttributePrefix
                    }
                }

                if(($element.DataType -eq "text" -or $element.DataType -eq "list") -and $element.AttributeExpandable) {
                    Add-ADMXRegXmlAttribute $child "expandable" "true"
                }
                if($element.DataType -eq "list" -and $element.AttributeAdditive) {
                    Add-ADMXRegXmlAttribute $child "additive" "true"
                }
                if($element.AttributeSoft) {
                    Add-ADMXRegXmlAttribute $child "soft" "true"
                }

                if($element.DataType -eq "list") {
                    $keyStr = ?? $element.Key $regPolicy.Key
                    if($keyStr) { $idStr = $keyStr.Trim('\').Split('\')[-1] }
                    else        { $idStr = [Guid]::NewGuid().Guid }
                    $elementId = $idStr + "_Id"
                    Add-ADMXRegXmlAttribute $child "id" $elementId
                }
                else {
                    $elementId = $element.ValueName + "_Id"
                    Add-ADMXRegXmlAttribute $child "id" $elementId
                }

                if($element.Key) {
                    Add-ADMXRegXmlAttribute $child "key" $element.Key.Trim('\')
                }

                $escapedValue = [System.Security.SecurityElement]::Escape($value) -replace '&amp;#xF000;', '&#xF000;'
                $omaUriItems += "<data id=`"$elementId`" value=`"$escapedValue`"/>"

                [void]$newNode.SelectSingleNode("elements").AppendChild($child)
            }
        }

        if($omaUriItems.Count -gt 0) {
            $omaUriString = $omaUriString + [Environment]::NewLine + [Environment]::NewLine + ($omaUriItems -join [Environment]::NewLine)
        }
        if((Get-SettingValue "FormatOMAURI") -eq $true) {
            $omaUriString = Update-XmlFormatting $omaUriString
        }

        [void]$xml.policyDefinitions.SelectSingleNode("policies").AppendChild($newNode)

        $direction = if($regPolicy.Hive -eq "HKLM") { "Device" } else { "User" }
        $admxRegSettings += [PSCustomObject]@{
            "@odata.type" = "#microsoft.graph.omaSettingString"
            displayName   = "Set $policyName"
            omaUri        = "./$direction/Vendor/MSFT/Policy/Config/IntuneManagementReg~$($regProfile.PolicyType)~$($newNode.parentCategory.ref)/$policyName"
            value         = $omaUriString
        }
    }

    # Remove the seed/template policy before serialisation.
    [void]$xml.policyDefinitions.SelectSingleNode("policies").RemoveChild(
        $xml.policyDefinitions.policies.SelectSingleNode("policy"))

    $xmlString = Format-XML $xml
    if((Get-SettingValue "FormatOMAURI") -eq $true) {
        $xmlString = Update-XmlFormatting $xmlString
    }

    $intuneObj.omaSettings += [PSCustomObject]@{
        "@odata.type" = "#microsoft.graph.omaSettingString"
        displayName   = "Reg ADMX Ingestion"
        description   = "This XML is generated by Intune Management tool"
        omaUri        = "./Device/Vendor/MSFT/Policy/ConfigOperations/ADMXInstall/IntuneManagementReg/$($regProfile.PolicyType)/$regPolicyFileName"
        value         = $xmlString
    }
    $intuneObj.omaSettings += $admxRegSettings

    $json = $intuneObj | ConvertTo-Json -Depth 20

    $tokenId = Get-DefaultTokenId
    $resp = Invoke-MSGraphAPI -Url "deviceManagement/deviceConfigurations" -Content $json -HttpMethod "POST" -TokenId $tokenId -FullResponseObject

    if($resp -and $resp.Success) {
        Write-Log "Custom profile '$($intuneObj.displayName)' created with id $($resp.Content.Id)"
        $script:UIProvider.ShowMessageBox("Profile '$($intuneObj.displayName)' created successfully.", "Reg Values", "OK", "Information")
    }
    else {
        $errInfo = if($resp) { "$($resp.StatusCode) $($resp.StatusDescription)" } else { "no response" }
        Write-Log "Failed to create device configuration profile '$($intuneObj.displayName)' - $errInfo" 3
        $script:UIProvider.ShowMessageBox("Failed to create device configuration profile '$($intuneObj.displayName)'.`n`n$errInfo`n`nCheck the log for details.", "Reg Values", "OK", "Error")
    }
}
