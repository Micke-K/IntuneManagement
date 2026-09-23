# Bulk Scope Tag form.
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Show-GraphBulkScopeTagForm
{
    $script:bulkScopeTagForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\BulkScopeTagForm.xaml"))
    if(-not $script:bulkScopeTagForm) { return }

    $script:dgBulkScopeTagObjects   = $script:bulkScopeTagForm.FindName("dgBulkScopeTagObjects")
    $script:lstBulkScopeTagAvail    = $script:bulkScopeTagForm.FindName("lstBulkScopeTagAvailable")
    $script:lstBulkScopeTagSelected = $script:bulkScopeTagForm.FindName("lstBulkScopeTagSelected")
    $script:txtBulkScopeTagStatus   = $script:bulkScopeTagForm.FindName("txtBulkScopeTagStatus")

    $scopeTagSettings = [IntuneManagerScopeTagSettings]::new()
    $script:bulkScopeTagForm.DataContext = $scopeTagSettings

    # ── Object list (mirrors Bulk Export DataGrid setup) ──────────────────────
    $script:bulkScopeTagEligibleTypes  = @($script:IntuneTypes  | Where-Object { $_.ScopeTagProperty })
    $script:bulkScopeTagEligibleGroups = @($script:IntuneGroups | Where-Object {
        $_.Title -and ($_.PolicyTypes | Where-Object { $_.ScopeTagProperty })
    })

    $column = Get-GridCheckboxColumn "Selected"
    $script:dgBulkScopeTagObjects.Columns.Add($column)
    $column.Header.IsChecked = $true
    $column.Header.add_Click({
        foreach($item in $script:dgBulkScopeTagObjects.ItemsSource) {
            $item.Selected = $this.IsChecked
        }
        $script:dgBulkScopeTagObjects.Items.Refresh()
    })

    $col = [System.Windows.Controls.DataGridTextColumn]::new()
    $col.Header     = "Object type"
    $col.IsReadOnly = $true
    $col.Binding    = [System.Windows.Data.Binding]::new("Title")
    $script:dgBulkScopeTagObjects.Columns.Add($col)

    $script:bulkScopeTagMode = "Group"
    Update-BulkScopeTagObjectList

    $script:UIProvider.AddXamlEvent($script:bulkScopeTagForm, "rbBulkScopeTagViewGroup", "add_Checked", {
        $script:bulkScopeTagMode = "Group"; Update-BulkScopeTagObjectList
    })
    $script:UIProvider.AddXamlEvent($script:bulkScopeTagForm, "rbBulkScopeTagViewType", "add_Checked", {
        $script:bulkScopeTagMode = "Type"; Update-BulkScopeTagObjectList
    })

    # ── Action radio → settings.Action sync ───────────────────────────────────
    # XAML radio buttons aren't directly bound to a string property; wire change
    # handlers so the [IntuneManagerScopeTagSettings] always reflects the UI.
    $script:UIProvider.AddXamlEvent($script:bulkScopeTagForm, "rbBulkScopeTagAdd",     "add_Checked", { $script:bulkScopeTagForm.DataContext.Action = "Add" })
    $script:UIProvider.AddXamlEvent($script:bulkScopeTagForm, "rbBulkScopeTagReplace", "add_Checked", { $script:bulkScopeTagForm.DataContext.Action = "Replace" })
    $script:UIProvider.AddXamlEvent($script:bulkScopeTagForm, "rbBulkScopeTagRemove",  "add_Checked", { $script:bulkScopeTagForm.DataContext.Action = "Remove" })

    # ── Scope tag picker (dual-list with click-to-move) ───────────────────────
    # Keep the modal open path cheap. The tenant dependency cache is used first;
    # if it is cold, the live Graph fetch runs after the dialog has rendered.
    $script:_bulkScopeTagAll = @()
    $script:_bulkScopeTagSelectedIds = [System.Collections.Generic.HashSet[string]]::new()

    $script:colBulkScopeTagAvail    = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    $script:colBulkScopeTagSelected = [System.Collections.ObjectModel.ObservableCollection[object]]::new()

    $script:_syncBulkScopeTagLists = {
        $script:colBulkScopeTagAvail.Clear()
        $script:colBulkScopeTagSelected.Clear()
        foreach($t in $script:_bulkScopeTagAll) {
            if($script:_bulkScopeTagSelectedIds.Contains($t.Id)) {
                [void]$script:colBulkScopeTagSelected.Add($t)
            } else {
                [void]$script:colBulkScopeTagAvail.Add($t)
            }
        }
    }
    & $script:_syncBulkScopeTagLists

    $script:lstBulkScopeTagAvail.ItemsSource    = $script:colBulkScopeTagAvail
    $script:lstBulkScopeTagSelected.ItemsSource = $script:colBulkScopeTagSelected
    $script:txtBulkScopeTagStatus.Text = "Loading scope tags..."

    # Double-click toggles list membership in either direction; the Add/Remove
    # buttons do the same for the highlighted rows (double-click alone was not
    # discoverable).
    $script:_bulkScopeTagAddSelected = {
        $ids = @($script:lstBulkScopeTagAvail.SelectedItems | ForEach-Object { [string]$_.Id })
        foreach($id in $ids) { [void]$script:_bulkScopeTagSelectedIds.Add($id) }
        & $script:_syncBulkScopeTagLists
    }
    $script:_bulkScopeTagRemoveSelected = {
        $ids = @($script:lstBulkScopeTagSelected.SelectedItems | ForEach-Object { [string]$_.Id })
        foreach($id in $ids) { [void]$script:_bulkScopeTagSelectedIds.Remove($id) }
        & $script:_syncBulkScopeTagLists
    }

    $script:lstBulkScopeTagAvail.Add_MouseDoubleClick({ & $script:_bulkScopeTagAddSelected })
    $script:lstBulkScopeTagSelected.Add_MouseDoubleClick({ & $script:_bulkScopeTagRemoveSelected })
    $script:UIProvider.AddXamlEvent($script:bulkScopeTagForm, "btnBulkScopeTagAddTag",    "add_click", { & $script:_bulkScopeTagAddSelected })
    $script:UIProvider.AddXamlEvent($script:bulkScopeTagForm, "btnBulkScopeTagRemoveTag", "add_click", { & $script:_bulkScopeTagRemoveSelected })

    # ── Apply ─────────────────────────────────────────────────────────────────
    $script:UIProvider.AddXamlEvent($script:bulkScopeTagForm, "btnBulkScopeTagApply", "add_click", {
        # Concurrent-click guard. Apply triggers a multi-step Graph run; without
        # this, a double-click (or impatient second press) fires two parallel
        # bulk PATCH passes against the same selection.
        $btnApply = $script:bulkScopeTagForm.FindName("btnBulkScopeTagApply")
        $btnClose = $script:bulkScopeTagForm.FindName("btnBulkScopeTagClose")

        $selectionIds = Get-BulkScopeTagSelectedObjectIds
        $unit = if($script:bulkScopeTagMode -eq "Type") { "policy type" } else { "object group" }

        if($selectionIds.Count -eq 0) {
            $script:UIProvider.ShowMessageBox("Select at least one $unit to update.", "Bulk Scope Tags", "OK", "Warning") | Out-Null
            return
        }

        $settings = $script:bulkScopeTagForm.DataContext
        $settings.ScopeTagIds = @($script:_bulkScopeTagSelectedIds)

        # Preflight guards
        if($settings.ScopeTagIds.Count -eq 0 -and -not $settings.CleanupOrphans) {
            $script:UIProvider.ShowMessageBox("Pick at least one scope tag, or enable 'Cleanup orphans' for a pure cleanup pass.", "Bulk Scope Tags", "OK", "Warning") | Out-Null
            return
        }

        # Mass-wipe guard: Replace with no selected tags will strip every tag
        # (including Default) from every matched policy. Defending here rather
        # than in the public command keeps the cmdlet itself non-prompting.
        if($settings.Action -eq "Replace" -and $settings.ScopeTagIds.Count -eq 0) {
            $proceed = $script:UIProvider.ShowMessageBox(
                "Replace with no scope tags selected will REMOVE every tag (including Default) from every policy that matches the filter.`n`nMost Intune policies require the Default tag and the API will reject empty roleScopeTagIds - you may end up with a large failure count.`n`nContinue?",
                "Confirm wipe-all", "YesNo", "Warning")
            if($proceed -ne "Yes") { return }
        }

        # Default-missing soft guard — mirrors the Details-view "Default scope
        # tag missing" warning. Triggers for Add (tags selected but Default
        # not included) and Replace (tags selected but Default not included).
        # Remove is unaffected because we only remove what the user picked.
        if($settings.Action -in @("Add","Replace") -and
           $settings.ScopeTagIds.Count -gt 0 -and
           $settings.ScopeTagIds -notcontains "0") {
            $proceed = $script:UIProvider.ShowMessageBox(
                "The 'Default' scope tag (Id 0) is not in the selected list.`n`nIntune usually requires Default to be present; PATCH calls may fail for policies that end up without it.`n`nContinue anyway?",
                "Default scope tag missing", "YesNo", "Warning")
            if($proceed -ne "Yes") { return }
        }

        $tagNames = @($script:_bulkScopeTagAll |
            Where-Object { $script:_bulkScopeTagSelectedIds.Contains($_.Id) } |
            ForEach-Object { $_.Name })
        $tagList = if($tagNames.Count -gt 0) { $tagNames -join ", " } else { "(none - cleanup only)" }
        $cleanup = if($settings.CleanupOrphans) { "yes" } else { "no" }
        $filterText = if($settings.Filter) { $settings.Filter } else { "(none)" }

        $confirm = $script:UIProvider.ShowMessageBox(@"
About to update scope tags on every policy that matches:

  Action          : $($settings.Action)
  Tags            : $tagList
  Cleanup orphans : $cleanup
  Name filter     : $filterText
  $unit count     : $($selectionIds.Count)

Continue?
"@, "Confirm Bulk Scope Tag update", "YesNo", "Warning")
        if($confirm -ne "Yes") { return }

        # Disable both buttons for the duration of the run so a stray click
        # doesn't restart it (Apply) or yank the form mid-flight (Close).
        if($btnApply) { $btnApply.IsEnabled = $false }
        if($btnClose) { $btnClose.IsEnabled = $false }
        # No pre-call Write-Status: Set-GraphBulkScopeTags now sets its own two-line
        # status ("Bulk scope tags — <Action>" + sub-step) as soon as it starts.
        $startParams = @{ ScopeTagSettings = $settings }
        if($script:bulkScopeTagMode -eq "Type") { $startParams.PolicyType = $selectionIds }
        else                                    { $startParams.PolicyGroup = $selectionIds }

        try {
            $summary = Set-GraphBulkScopeTags @startParams
            Write-Status $null

            $msg = ("Scanned: {0}`nMatched filter: {1}`nUpdated: {2}`nNo change: {3}`nFailed: {4}`nDuration: {5:hh\:mm\:ss}" -f `
                $summary.PoliciesScanned, $summary.PoliciesMatched, $summary.PoliciesUpdated, $summary.PoliciesSkipped, $summary.PoliciesFailed, $summary.Duration)
            $script:txtBulkScopeTagStatus.Text = ("Last run: updated {0}, no change {1}, failed {2}" -f $summary.PoliciesUpdated, $summary.PoliciesSkipped, $summary.PoliciesFailed)
            $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
            if($unknownNote) { $msg += "`n`n$unknownNote" }
            $script:UIProvider.ShowMessageBox($msg, "Bulk Scope Tags", "OK", "Information") | Out-Null

            # Reflect the new tag values in the main DataGrid (Set-GraphBulkScopeTags
            # already mirrored Object.<prop> + cleared _ScopeTags on updated rows).
            try { $script:dgIntuneManagerObjects.Items.Refresh() }
            catch { Write-LogDebug "Main grid refresh failed after bulk scope tag run: $($_.Exception.Message)" }
        }
        catch {
            Write-LogError "Bulk Scope Tags run failed" $_.Exception
            $script:UIProvider.ShowMessageBox("Bulk Scope Tags run failed: $($_.Exception.Message)", "Bulk Scope Tags", "OK", "Error") | Out-Null
        }
        finally {
            if($btnApply) { $btnApply.IsEnabled = $true }
            if($btnClose) { $btnClose.IsEnabled = $true }
        }
    })

    $script:UIProvider.AddXamlEvent($script:bulkScopeTagForm, "btnBulkScopeTagClose", "add_click", {
        $script:bulkScopeTagForm = $null
        Show-ModalObject
    })

    $script:UIProvider.ShowModalForm("Bulk Scope Tags", $script:bulkScopeTagForm, $true)

    try {
        $scopeTagLoadError = $null
        $script:_bulkScopeTagAll = @(Get-BulkScopeTagCatalog -TokenId (Get-DefaultTokenId) -LoadError ([ref]$scopeTagLoadError))
        & $script:_syncBulkScopeTagLists
        $script:txtBulkScopeTagStatus.Text = "$($script:_bulkScopeTagAll.Count) scope tag(s) loaded"
        if($scopeTagLoadError) {
            $script:UIProvider.ShowMessageBox("Could not load scope tags from the tenant.`n`n$scopeTagLoadError`n`nThe picker will only show 'Default'.", "Bulk Scope Tags", "OK", "Warning") | Out-Null
        }
    }
    catch {
        Write-LogError "Bulk Scope Tags: failed to initialize scope tag picker" $_.Exception
        $script:txtBulkScopeTagStatus.Text = "Scope tag loading failed"
    }
}

# Rebuild the bulk-scope-tag object DataGrid for the current view mode.
function Update-BulkScopeTagObjectList
{
    if(-not $script:dgBulkScopeTagObjects) { return }

    $script:bulkScopeTagObjects = @()

    if($script:bulkScopeTagMode -eq "Type") {
        $sortedTypes = $script:bulkScopeTagEligibleTypes | Sort-Object Title
        foreach($pt in $sortedTypes) {
            $script:bulkScopeTagObjects += New-Object PSObject -Property @{
                Title       = $pt.Title
                Selected    = $true
                ObjectGroup = $null
                ObjectType  = $pt
            }
        }
    }
    else {
        $sortedGroups = $script:bulkScopeTagEligibleGroups | Sort-Object Title
        foreach($grp in $sortedGroups) {
            $script:bulkScopeTagObjects += New-Object PSObject -Property @{
                Title       = $grp.Title
                Selected    = $true
                ObjectGroup = $grp
                ObjectType  = $null
            }
        }
    }

    $script:dgBulkScopeTagObjects.ItemsSource = $script:bulkScopeTagObjects
}

function Get-BulkScopeTagSelectedObjectIds
{
    if($script:bulkScopeTagMode -eq "Type") {
        return @($script:bulkScopeTagObjects |
            Where-Object { $_.Selected -eq $true -and $_.ObjectType -and $_.ObjectType.Id } |
            ForEach-Object { $_.ObjectType.Id })
    }
    return @($script:bulkScopeTagObjects |
        Where-Object { $_.Selected -eq $true -and $_.ObjectGroup -and $_.ObjectGroup.Id } |
        ForEach-Object { $_.ObjectGroup.Id })
}

# ─── Bulk Assignments ─────────────────────────────────────────────────────────
# Thin UI wrapper around Set-GraphBulkAssignments. Mirrors the Bulk Scope Tag
# form's structure: Action radio, name filter, eligible-type selector
# (Group/API), plus an assignments list with Add buttons. Phase 1 supports
# group / exclusion-group / all-devices / all-users targets with optional
# assignment filters — app intent + per-platform settings come in later phases.
