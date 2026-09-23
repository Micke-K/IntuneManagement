# Avalonia port of the Bulk Scope Tag dialog from
# UI/WPF/Extensions/IntuneManagerUI.ps1 (Show-GraphBulkScopeTagForm + helpers).
# New file because Bulk Scope Tags is a self-contained subsystem and the WPF
# original sits in the giant IntuneManagerUI file we're trying not to grow
# further (architecture rule R9).
#
# Slice 4d: Bulk Scope Tags. Largest of the bulk forms — combines the
# Group/API DataGrid (Slice 4a-style) with the dual-list scope-tag picker
# from the Copy dialog (Slice 3d3 — see [[CopyDialogScopeTagItem]]).
# Driver Set-GraphBulkScopeTags lives in Public/, so the click handler just
# preflight-validates, prompts, and forwards to the cmdlet.
#
# Get-BulkScopeTagCatalog lives in Internal/IntuneScopeTags.ps1 and is
# shared with the WPF tree.

function Update-BulkScopeTagObjectList
{
    if (-not $script:dgBulkScopeTagObjects) { return }

    $rows = [System.Collections.Generic.List[BulkScopeTagRowItem]]::new()

    if ($script:bulkScopeTagMode -eq "Type") {
        $sortedTypes = $script:bulkScopeTagEligibleTypes | Sort-Object Title
        foreach ($pt in $sortedTypes) {
            $rows.Add([BulkScopeTagRowItem]@{
                Title       = [string]$pt.Title
                Selected    = $true
                ObjectGroup = $null
                ObjectType  = $pt
            })
        }
    } else {
        $sortedGroups = $script:bulkScopeTagEligibleGroups | Sort-Object Title
        foreach ($grp in $sortedGroups) {
            $rows.Add([BulkScopeTagRowItem]@{
                Title       = [string]$grp.Title
                Selected    = $true
                ObjectGroup = $grp
                ObjectType  = $null
            })
        }
    }

    $script:bulkScopeTagRows = @($rows)
    $script:dgBulkScopeTagObjects.ItemsSource = $script:bulkScopeTagRows
}

function Get-BulkScopeTagSelectedObjectIds
{
    if ($null -eq $script:bulkScopeTagRows) { return @() }
    if ($script:bulkScopeTagMode -eq "Type") {
        return @($script:bulkScopeTagRows |
            Where-Object { $_.Selected -and $_.ObjectType -and $_.ObjectType.Id } |
            ForEach-Object { $_.ObjectType.Id })
    }
    return @($script:bulkScopeTagRows |
        Where-Object { $_.Selected -and $_.ObjectGroup -and $_.ObjectGroup.Id } |
        ForEach-Object { $_.ObjectGroup.Id })
}

# Rebuild both dual-list collections from the catalog + the AssignedIds set.
# Reads module-scope state so it is safe to call from event handlers (no
# captured locals / GetNewClosure — see [[avalonia-closure-dynamic-module]]).
function Sync-BulkScopeTagLists
{
    $st = $script:_bulkScopeTagState
    if (-not $st) { return }
    $state = $st.State
    $state.AvailableCollection.Clear()
    $state.AssignedCollection.Clear()
    foreach ($tag in $state.AllScopeTags) {
        if ($state.AssignedIds.Contains([string]$tag.Id)) { [void]$state.AssignedCollection.Add($tag) }
        else { [void]$state.AvailableCollection.Add($tag) }
    }
}

function Move-BulkScopeTagSelection
{
    param([string]$Action, [string[]]$Ids)
    $st = $script:_bulkScopeTagState
    if (-not $st -or $Ids.Count -eq 0) { return }
    $state = $st.State
    if ($Action -eq 'assign') { foreach ($id in $Ids) { [void]$state.AssignedIds.Add([string]$id) } }
    else                      { foreach ($id in $Ids) { [void]$state.AssignedIds.Remove([string]$id) } }
    Sync-BulkScopeTagLists
}

function Show-GraphBulkScopeTagForm
{
    $ui = $script:UIProvider
    $form = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/BulkScopeTagForm.axaml'))
    if (-not $form) { return }

    $hostType = Get-AvaloniaHost
    $script:dgBulkScopeTagObjects   = $hostType::FindByName($form, 'dgBulkScopeTagObjects')
    $lstAvail                       = $hostType::FindByName($form, 'lstBulkScopeTagAvailable')
    $lstSelected                    = $hostType::FindByName($form, 'lstBulkScopeTagSelected')
    $txtStatus                      = $hostType::FindByName($form, 'txtBulkScopeTagStatus')

    $rbAdd      = $hostType::FindByName($form, 'rbBulkScopeTagAdd')
    $rbReplace  = $hostType::FindByName($form, 'rbBulkScopeTagReplace')
    $rbRemove   = $hostType::FindByName($form, 'rbBulkScopeTagRemove')
    $chkClean   = $hostType::FindByName($form, 'chkBulkScopeTagCleanupOrphans')
    $txtFilter  = $hostType::FindByName($form, 'txtBulkScopeTagNameFilter')
    $rbGroup    = $hostType::FindByName($form, 'rbBulkScopeTagViewGroup')
    $rbType     = $hostType::FindByName($form, 'rbBulkScopeTagViewType')
    $btnApply   = $hostType::FindByName($form, 'btnBulkScopeTagApply')
    $btnClose   = $hostType::FindByName($form, 'btnBulkScopeTagClose')

    $scopeTagSettings = [IntuneManagerScopeTagSettings]::new()

    # Eligible types/groups: same gate as the WPF original (must have a
    # ScopeTagProperty defined on the IntuneType).
    $script:bulkScopeTagEligibleTypes  = @($script:IntuneTypes  | Where-Object { $_.ScopeTagProperty })
    $script:bulkScopeTagEligibleGroups = @($script:IntuneGroups | Where-Object {
        $_.Title -and ($_.PolicyTypes | Where-Object { $_.ScopeTagProperty })
    })

    $script:bulkScopeTagMode = "Group"
    Update-BulkScopeTagObjectList

    $headerCb = Initialize-AvaloniaGridSelectAllHeader -Grid $script:dgBulkScopeTagObjects -BindingProperty 'Selected' -InitiallyChecked $true

    # ── Scope tag picker (dual-list) state ──────────────────────────────────
    # ObservableCollections bound to the two ListBoxes, a HashSet of selected
    # ids as source of truth, rebuilt by Sync-BulkScopeTagLists /
    # Move-BulkScopeTagSelection.
    $state = [PSCustomObject]@{
        AllScopeTags         = @()
        AssignedIds          = [System.Collections.Generic.HashSet[string]]::new()
        AvailableCollection  = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
        AssignedCollection   = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    }

    # Module-scope state so event handlers resolve it at click time (no captured
    # locals / GetNewClosure / $script:UIProvider — see [[avalonia-closure-dynamic-module]]).
    $script:_bulkScopeTagState = @{
        State       = $state
        HeaderCb    = $headerCb
        RbReplace   = $rbReplace
        RbRemove    = $rbRemove
        ChkClean    = $chkClean
        TxtFilter   = $txtFilter
        TxtStatus   = $txtStatus
        LstAvail    = $lstAvail
        LstSelected = $lstSelected
        BtnApply    = $btnApply
        BtnClose    = $btnClose
        Settings    = $scopeTagSettings
    }

    if ($rbGroup) {
        $rbGroup.add_IsCheckedChanged({
            param($s, $e)
            if (-not $s.IsChecked) { return }
            $script:bulkScopeTagMode = "Group"
            Update-BulkScopeTagObjectList
            $st = $script:_bulkScopeTagState
            if ($st -and $st.HeaderCb) { $st.HeaderCb.IsChecked = $true }
        })
    }
    if ($rbType) {
        $rbType.add_IsCheckedChanged({
            param($s, $e)
            if (-not $s.IsChecked) { return }
            $script:bulkScopeTagMode = "Type"
            Update-BulkScopeTagObjectList
            $st = $script:_bulkScopeTagState
            if ($st -and $st.HeaderCb) { $st.HeaderCb.IsChecked = $true }
        })
    }

    if ($lstAvail)    { $lstAvail.ItemsSource    = $state.AvailableCollection }
    if ($lstSelected) { $lstSelected.ItemsSource = $state.AssignedCollection }
    if ($txtStatus)   { $txtStatus.Text = "Loading scope tags..." }

    if ($lstAvail) {
        $lstAvail.add_DoubleTapped({
            try {
                $st = $script:_bulkScopeTagState
                if (-not $st -or $st.LstAvail.SelectedItems.Count -eq 0) { return }
                $ids = @($st.LstAvail.SelectedItems | ForEach-Object { [string]$_.Id })
                Move-BulkScopeTagSelection 'assign' $ids
            } catch { Write-LogError "Bulk Scope Tag double-tap (available) failed" $_.Exception }
        })
    }
    if ($lstSelected) {
        $lstSelected.add_DoubleTapped({
            try {
                $st = $script:_bulkScopeTagState
                if (-not $st -or $st.LstSelected.SelectedItems.Count -eq 0) { return }
                $ids = @($st.LstSelected.SelectedItems | ForEach-Object { [string]$_.Id })
                Move-BulkScopeTagSelection 'unassign' $ids
            } catch { Write-LogError "Bulk Scope Tag double-tap (selected) failed" $_.Exception }
        })
    }

    # Add / Remove buttons - same moves as the double-taps, but discoverable.
    $btnAddTag    = $hostType::FindByName($form, 'btnBulkScopeTagAddTag')
    $btnRemoveTag = $hostType::FindByName($form, 'btnBulkScopeTagRemoveTag')
    if ($btnAddTag) {
        $btnAddTag.add_Click({
            try {
                $st = $script:_bulkScopeTagState
                if (-not $st -or $st.LstAvail.SelectedItems.Count -eq 0) { return }
                $ids = @($st.LstAvail.SelectedItems | ForEach-Object { [string]$_.Id })
                Move-BulkScopeTagSelection 'assign' $ids
            } catch { Write-LogError "Bulk Scope Tag Add button failed" $_.Exception }
        })
    }
    if ($btnRemoveTag) {
        $btnRemoveTag.add_Click({
            try {
                $st = $script:_bulkScopeTagState
                if (-not $st -or $st.LstSelected.SelectedItems.Count -eq 0) { return }
                $ids = @($st.LstSelected.SelectedItems | ForEach-Object { [string]$_.Id })
                Move-BulkScopeTagSelection 'unassign' $ids
            } catch { Write-LogError "Bulk Scope Tag Remove button failed" $_.Exception }
        })
    }

    # Apply — preflight + confirm + Set-GraphBulkScopeTags. All work runs
    # inside the click handler because Show-ModalForm is non-blocking.
    if ($btnApply) {
        $btnApply.add_Click({
            $st = $script:_bulkScopeTagState
            if (-not $st) { return }
            try {
                $state            = $st.State
                $scopeTagSettings = $st.Settings

                # Action radio -> settings.Action
                if ($st.RbReplace -and $st.RbReplace.IsChecked)     { $scopeTagSettings.Action = "Replace" }
                elseif ($st.RbRemove -and $st.RbRemove.IsChecked)   { $scopeTagSettings.Action = "Remove" }
                else                                                { $scopeTagSettings.Action = "Add" }

                $scopeTagSettings.CleanupOrphans = if ($st.ChkClean) { [bool]$st.ChkClean.IsChecked } else { $false }
                $scopeTagSettings.Filter         = if ($st.TxtFilter) { [string]$st.TxtFilter.Text } else { '' }
                $scopeTagSettings.ScopeTagIds    = @($state.AssignedIds)

                $selectionIds = Get-BulkScopeTagSelectedObjectIds
                $unit = if ($script:bulkScopeTagMode -eq "Type") { "policy type" } else { "object group" }

                if ($selectionIds.Count -eq 0) {
                    Show-MessageBox "Select at least one $unit to update." "Bulk Scope Tags" "OK" "Warning" | Out-Null
                    return
                }

                if ($scopeTagSettings.ScopeTagIds.Count -eq 0 -and -not $scopeTagSettings.CleanupOrphans) {
                    Show-MessageBox "Pick at least one scope tag, or enable 'Cleanup orphans' for a pure cleanup pass." "Bulk Scope Tags" "OK" "Warning" | Out-Null
                    return
                }

                # Mass-wipe guard: Replace with no selected tags strips every
                # tag (including Default) from every matched policy.
                if ($scopeTagSettings.Action -eq "Replace" -and $scopeTagSettings.ScopeTagIds.Count -eq 0) {
                    $proceed = Show-MessageBox "Replace with no scope tags selected will REMOVE every tag (including Default) from every policy that matches the filter.`n`nMost Intune policies require the Default tag and the API will reject empty roleScopeTagIds - you may end up with a large failure count.`n`nContinue?" "Confirm wipe-all" "YesNo" "Warning"
                    if ($proceed -ne "Yes") { return }
                }

                # Default-missing soft guard. Add/Replace with tags selected
                # but Default not included is usually a mistake.
                if ($scopeTagSettings.Action -in @("Add","Replace") -and
                    $scopeTagSettings.ScopeTagIds.Count -gt 0 -and
                    $scopeTagSettings.ScopeTagIds -notcontains "0") {
                    $proceed = Show-MessageBox "The 'Default' scope tag (Id 0) is not in the selected list.`n`nIntune usually requires Default to be present; PATCH calls may fail for policies that end up without it.`n`nContinue anyway?" "Default scope tag missing" "YesNo" "Warning"
                    if ($proceed -ne "Yes") { return }
                }

                $tagNames = @($state.AllScopeTags |
                    Where-Object { $state.AssignedIds.Contains([string]$_.Id) } |
                    ForEach-Object { $_.Name })
                $tagList    = if ($tagNames.Count -gt 0) { $tagNames -join ", " } else { "(none - cleanup only)" }
                $cleanup    = if ($scopeTagSettings.CleanupOrphans) { "yes" } else { "no" }
                $filterText = if ($scopeTagSettings.Filter) { $scopeTagSettings.Filter } else { "(none)" }

                $confirmMsg = @"
About to update scope tags on every policy that matches:

  Action          : $($scopeTagSettings.Action)
  Tags            : $tagList
  Cleanup orphans : $cleanup
  Name filter     : $filterText
  $unit count     : $($selectionIds.Count)

Continue?
"@
                $confirm = Show-MessageBox $confirmMsg "Confirm Bulk Scope Tag update" "YesNo" "Warning"
                if ($confirm -ne "Yes") { return }

                # Disable both buttons while the run executes — a stray click
                # would otherwise fire a parallel pass against the same set.
                if ($st.BtnApply) { $st.BtnApply.IsEnabled = $false }
                if ($st.BtnClose) { $st.BtnClose.IsEnabled = $false }

                try {
                    $startParams = @{ ScopeTagSettings = $scopeTagSettings }
                    if ($script:bulkScopeTagMode -eq "Type") { $startParams.PolicyType = $selectionIds }
                    else                                     { $startParams.PolicyGroup = $selectionIds }

                    $summary = Set-GraphBulkScopeTags @startParams
                    Write-Status ""

                    $msg = ("Scanned: {0}`nMatched filter: {1}`nUpdated: {2}`nNo change: {3}`nFailed: {4}`nDuration: {5:hh\:mm\:ss}" -f `
                        $summary.PoliciesScanned, $summary.PoliciesMatched, $summary.PoliciesUpdated, $summary.PoliciesSkipped, $summary.PoliciesFailed, $summary.Duration)
                    if ($st.TxtStatus) {
                        $st.TxtStatus.Text = ("Last run: updated {0}, no change {1}, failed {2}" -f $summary.PoliciesUpdated, $summary.PoliciesSkipped, $summary.PoliciesFailed)
                    }
                    $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
                    if ($unknownNote) { $msg += "`n`n$unknownNote" }
                    Show-MessageBox $msg "Bulk Scope Tags" "OK" "Information" | Out-Null

                    if ($script:IntuneManagerSelectedObject) {
                        Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject | Out-Null
                    }
                } catch {
                    Write-LogError "Bulk Scope Tags run failed" $_.Exception
                    Show-MessageBox "Bulk Scope Tags run failed: $($_.Exception.Message)" "Bulk Scope Tags" "OK" "Error" | Out-Null
                } finally {
                    if ($st.BtnApply) { $st.BtnApply.IsEnabled = $true }
                    if ($st.BtnClose) { $st.BtnClose.IsEnabled = $true }
                }
            } catch {
                Write-LogError "Bulk Scope Tags Apply handler failed" $_.Exception
            }
        })
    }

    if ($btnClose) {
        $btnClose.add_Click({
            $script:dgBulkScopeTagObjects = $null
            $script:bulkScopeTagRows = $null
            $script:_bulkScopeTagState = $null
            Show-ModalObject
        })
    }

    $form.add_KeyDown({
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            $script:dgBulkScopeTagObjects = $null
            $script:bulkScopeTagRows = $null
            $script:_bulkScopeTagState = $null
            Show-ModalObject
            $e.Handled = $true
        }
    })

    $ui.ShowModalForm("Bulk Scope Tags", $form, $true)

    # Load the catalog after the form has rendered. Cache lookup is cheap;
    # the live fetch falls back inside Get-BulkScopeTagCatalog when cold.
    try {
        $scopeTagLoadError = $null
        $catalog = @(Get-BulkScopeTagCatalog -TokenId (Get-DefaultTokenId) -LoadError ([ref]$scopeTagLoadError))
        $state.AllScopeTags = @($catalog | ForEach-Object {
            [CopyDialogScopeTagItem]@{ Id = [string]$_.Id; Name = [string]$_.Name }
        })
        Sync-BulkScopeTagLists
        if ($txtStatus) { $txtStatus.Text = "$($state.AllScopeTags.Count) scope tag(s) loaded" }
        if ($scopeTagLoadError) {
            $ui.ShowMessageBox("Could not load scope tags from the tenant.`n`n$scopeTagLoadError`n`nThe picker will only show 'Default'.", "Bulk Scope Tags", "OK", "Warning") | Out-Null
        }
    } catch {
        Write-LogError "Bulk Scope Tags: failed to initialize scope tag picker" $_.Exception
        if ($txtStatus) { $txtStatus.Text = "Scope tag loading failed" }
    }
}
