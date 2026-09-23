# Bulk Export form.
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Show-GraphBulkExportForm
{
    $script:bulkExportForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\BulkExportForm.xaml"), $true)
    if(-not $script:bulkExportForm) { return }

    $script:dgObjectsToExport = $script:bulkExportForm.FindName("dgObjectsToExport")

    $exportSettings = [IntuneManagerExportSettings]::new()

    $script:bulkExportForm.DataContext = $exportSettings

    $script:UIProvider.AddXamlEvent($script:bulkExportForm, "browseExportPath", "add_click", {
        $folder = Get-Folder ($script:UIProvider.GetXamlProperty($script:bulkExportForm, "txtExportPath", "Text")) "Select root folder for export"
        if($folder)
        {
            $script:bulkExportForm.DataContext.ExportFolder = $folder
            $tmp = $script:bulkExportForm.DataContext
            $script:bulkExportForm.DataContext = $null
            $script:bulkExportForm.DataContext = $tmp
        }
    })

    # Register dynamic export-properties / UI extensions for every exportable
    # policy type up front. Done once (not per mode switch) so toggling
    # Group/API doesn't re-register and we never miss a type that only appears
    # in the other view.
    $script:bulkExportExportableGroups = @($script:IntuneGroups | Where-Object {
        $_.Title -and (-not ($_.ShowButtons -is [Object[]]) -or $_.ShowButtons -contains "Export")
    })
    foreach($intuneGroup in $script:bulkExportExportableGroups)
    {
        $intuneGroup.PolicyTypes | Add-IntuneManagerExportProperties -ExportSettings $exportSettings
        $intuneGroup.PolicyTypes | Add-IntuneManagerExportUIExtensions -Form $script:bulkExportForm
    }

    $column = Get-GridCheckboxColumn "Selected"
    $script:dgObjectsToExport.Columns.Add($column)

    $column.Header.IsChecked = $true # All items are checked by default
    $column.Header.add_Click({
            foreach($Item in $script:dgObjectsToExport.ItemsSource)
            {
                $Item.Selected = $this.IsChecked
            }
            $script:dgObjectsToExport.Items.Refresh()
        }
    )

    # Add Object type column
    $binding = [System.Windows.Data.Binding]::new("Title")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Object type"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:dgObjectsToExport.Columns.Add($column)

    # Default view is "Group" (matches the IsChecked=True on the XAML radio).
    $script:bulkExportMode = "Group"
    Update-BulkExportObjectList

    # Repopulate the grid when the user flips between Group and API views.
    $script:UIProvider.AddXamlEvent($script:bulkExportForm, "rbBulkExportViewGroup", "add_Checked", {
        $script:bulkExportMode = "Group"
        Update-BulkExportObjectList
    })
    $script:UIProvider.AddXamlEvent($script:bulkExportForm, "rbBulkExportViewType", "add_Checked", {
        $script:bulkExportMode = "Type"
        Update-BulkExportObjectList
    })

    $script:UIProvider.AddXamlEvent($script:bulkExportForm, "btnClose", "add_click", {
        $script:bulkExportForm = $null
        Show-ModalObject
    })

    # NOTE: the "Save settings for batch job" button was removed 2026-06-12 per
    # user decision — scheduled/CI runs use the public drivers directly
    # (Start-GraphBulkExport -SettingsFile files can still be produced with
    # Save-GraphBulkExportSettings from a script).

    # btnExport: collect the selected rows from the DataGrid and call the public
    # driver. The UI does NOT contain export logic — it's a thin caller of
    # Start-GraphBulkExport so the same operation is scriptable headlessly.
    $script:UIProvider.AddXamlEvent($script:bulkExportForm, "btnExport", "add_click", {
        $selection = Get-BulkExportSelectedIds
        $unit = if ($script:bulkExportMode -eq "Type") { "policy type" } else { "object group" }

        if ($selection.Count -eq 0) {
            $script:UIProvider.ShowMessageBox("Select at least one $unit to export.", "Bulk Export", "OK", "Warning") | Out-Null
            return
        }

        # Read the path straight from the TextBox rather than relying on the binding —
        # the field is bound TwoWay/PropertyChanged so they should agree, but pulling
        # from the control directly is bullet-proof. Trim and unquote (users sometimes
        # paste paths with surrounding quotes).
        $rawPath = ($script:UIProvider.GetXamlProperty($script:bulkExportForm, "txtExportPath", "Text"))
        if ($rawPath) { $rawPath = ([string]$rawPath).Trim().Trim('"').Trim("'") }
        if (-not $rawPath) {
            $script:UIProvider.ShowMessageBox("Select an export folder.", "Bulk Export", "OK", "Warning") | Out-Null
            return
        }
        # Resolve to a full path so the success message shows the canonical location,
        # not whatever relative-ish text the user typed.
        try { $resolvedPath = [IO.Path]::GetFullPath($rawPath) }
        catch {
            $script:UIProvider.ShowMessageBox("Invalid export folder path: $($_.Exception.Message)", "Bulk Export", "OK", "Error") | Out-Null
            return
        }
        $script:bulkExportForm.DataContext.ExportFolder = $resolvedPath
        Write-Log "Bulk export starting. Folder: $resolvedPath"

        $startParams = @{
            ExportSettings = $script:bulkExportForm.DataContext
        }
        if ($script:bulkExportMode -eq "Type") {
            $startParams.PolicyType = $selection
        }
        else {
            $startParams.PolicyGroup = $selection
        }

        try {
            $summary = Start-GraphBulkExport @startParams
            Write-Status $null
            $msg = ("Exported {0} policies across {1} types ({2} failed) in {3:hh\:mm\:ss}.`n`nFolder:`n{4}" -f `
                $summary.Policies, $summary.Types, $summary.Failed, $summary.Duration, $resolvedPath)
            $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
            if($unknownNote) { $msg += "`n`n$unknownNote" }
            $script:UIProvider.ShowMessageBox($msg, "Bulk Export", "OK", "Information") | Out-Null
        }
        catch {
            Write-LogError "Bulk export failed" $_.Exception
            $script:UIProvider.ShowMessageBox("Bulk export failed: $($_.Exception.Message)", "Bulk Export", "OK", "Error") | Out-Null
        }

        $script:bulkExportForm = $null
        Show-ModalObject
    })

    $script:UIProvider.ShowModalForm("Bulk Export", $script:bulkExportForm, $true)
}

# Rebuild the bulk-export DataGrid for the current $script:bulkExportMode.
# Group mode lists policy groups (Configuration, Compliance, ...); Type mode
# lists individual policy types (APIs). Both sort by Title so the order is
# stable instead of class-load order.
function Update-BulkExportObjectList
{
    if (-not $script:dgObjectsToExport) { return }

    $script:exportObjects = @()

    if ($script:bulkExportMode -eq "Type") {
        $exportableGroupIds = @($script:bulkExportExportableGroups | ForEach-Object { $_.Id })
        $sortedTypes = $script:IntuneTypes |
            Where-Object { $_.PolicyGroup -and ($exportableGroupIds -contains $_.PolicyGroup.Id) } |
            Sort-Object Title
        foreach ($intuneType in $sortedTypes) {
            $script:exportObjects += New-Object PSObject -Property @{
                Title       = $intuneType.Title
                Selected    = $true
                ObjectGroup = $null
                ObjectType  = $intuneType
            }
        }
    }
    else {
        $sortedGroups = $script:bulkExportExportableGroups | Sort-Object Title
        foreach ($intuneGroup in $sortedGroups) {
            $script:exportObjects += New-Object PSObject -Property @{
                Title       = $intuneGroup.Title
                Selected    = (?? $intuneGroup.BulkExport $true)
                ObjectGroup = $intuneGroup
                ObjectType  = $null
            }
        }
    }

    $script:dgObjectsToExport.ItemsSource = $script:exportObjects
}

# Pull the Id list out of the currently-checked rows. In Group mode that's
# PolicyGroup IDs (for -PolicyGroup); in Type mode it's PolicyType IDs (for
# -PolicyType). Returns @() when nothing is selected.
function Get-BulkExportSelectedIds
{
    if ($script:bulkExportMode -eq "Type") {
        return @($script:exportObjects |
            Where-Object { $_.Selected -eq $true -and $_.ObjectType -and $_.ObjectType.Id } |
            ForEach-Object { $_.ObjectType.Id })
    }
    return @($script:exportObjects |
        Where-Object { $_.Selected -eq $true -and $_.ObjectGroup -and $_.ObjectGroup.Id } |
        ForEach-Object { $_.ObjectGroup.Id })
}

# ─── Bulk Scope Tags ──────────────────────────────────────────────────────────
# Thin UI wrapper around Set-GraphBulkScopeTags. The form mirrors Bulk Export's
# Group / API selector for picking object types, adds an action selector
# (Add / Replace / Remove), an orphan-cleanup checkbox, and a dual-list scope
# tag picker. All business logic lives in the public command — the UI only
# collects inputs, calls it, and shows the summary.
# Get-BulkScopeTagCatalog moved to Internal/IntuneScopeTags.ps1 so the
# Avalonia tree can share it.

