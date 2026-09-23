# Avalonia port of UI/WPF/Extensions/IntuneManagerUI.ps1::Show-IntuneManagerDetailedView
# (lines 666-1135 in WPF). Lives in its own file per architecture R9 — Details
# is a self-contained subsystem and the WPF original sits in the giant
# IntuneManagerUI file we're trying not to grow further.
#
# Slice 3f: ObjectDetails.axaml + Show-IntuneManagerDetailedView.
#   - JSON tab: txtValue + Load full + Copy
#   - Settings tab: name + description + scope-tag dual-list + Save (PATCH)
#   - Columns tab: Basic/Object property pickers + ObjectColumnInfo editor
#                  + Up/Down/Delete/Clear/Override + Reset/Save
#
# Per-policy-type AddUIDetailsExtension is invoked at the bottom — Avalonia
# variants live in UI/Avalonia/ClassExtensions/Intune*ClassUIExtension*.ps1
# and route through Add-AvaloniaDetailsButton (see
# Extensions/IntuneManagerExtensionHooks.ps1) instead of building WPF
# controls directly.
#
# Dirty-tracking signatures + Confirm-DetailsViewClose work the same as WPF,
# but read from the captured $state hashtable rather than $script: vars so
# multiple modal opens don't trample each other.

function Show-IntuneManagerDetailedView
{
    param(
        $FormTitle = "",
        [switch]$NoLoadFull
    )

    $ui = $script:UIProvider
    if (-not $script:dgIntuneManagerObjects) { return }

    $selectedRow = $script:dgIntuneManagerObjects.SelectedItem
    if (-not $selectedRow -and $script:dgIntuneManagerObjects.ItemsSource) {
        $selectedRow = @($script:dgIntuneManagerObjects.ItemsSource |
            Where-Object { $_.PSObject.Properties['IsSelected'] -and $_.IsSelected -eq $true } |
            Select-Object -First 1)
        if ($selectedRow.Count -gt 0) { $selectedRow = $selectedRow[0] }
    }
    if (-not $selectedRow) { return }

    # IntuneObjectRowItem.Source holds the underlying IntunePolicyBase, which
    # is what JsonObject / Object / Get() etc. live on. WPF read directly from
    # the row because rows there were the IntunePolicyBase itself.
    $selectedItem = $selectedRow.Source
    if (-not $selectedItem -or -not $selectedItem.JsonObject) {
        Write-Log "View: selected row carries no policy object" 3
        return
    }

    $detailsForm = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/ObjectDetails.axaml'))
    if (-not $detailsForm) { return }

    $hostType = Get-AvaloniaHost

    # State captured in closures rather than $script: vars so reopening the
    # dialog doesn't trample a still-running instance and so the close
    # confirmation sees the right form.
    $state = @{
        Form                = $detailsForm
        SelectedItem        = $selectedItem
        ScopeTagPropName    = $null
        AllScopeTags        = @()
        AssignedIds         = [System.Collections.Generic.HashSet[string]]::new()
        AvailableCollection = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
        AssignedCollection  = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
        ColObjectProperties = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
        InitialSettingsSig  = ""
        InitialColumnsSig   = ""
        LstAvail            = $null
        LstAssigned         = $null
        TxtValue            = $null
        LstBasic            = $null
        LstObjProps         = $null
        LstObjCols          = $null
        GrdObjCols          = $null
        ChkOverride         = $null
        HostType            = $null
    }
    # Event handlers must reach this through $script: - closure captures do
    # not survive ConvertTo-AvaloniaEventScriptBlock (NewBoundScriptBlock
    # rebinds to the module and drops the .GetNewClosure() state).
    $script:_detailsTagState = $state

    if (-not $FormTitle) { $FormTitle = $selectedItem.PolicyName }
    if ($selectedItem.Name) { $FormTitle = "$FormTitle - $($selectedItem.Name)" }

    # ---- JSON tab ----------------------------------------------------------
    $txtValue = $hostType::FindByName($detailsForm, 'txtValue')
    $state.TxtValue = $txtValue
    $state.HostType = $hostType
    $btnFull  = $hostType::FindByName($detailsForm, 'btnFull')
    $btnCopy  = $hostType::FindByName($detailsForm, 'btnCopy')

    if ($txtValue) { $txtValue.Text = [string]$selectedItem.JsonString }

    if ($btnFull -and $selectedItem.PolicyType.AllowFullDetails -eq $false) {
        $btnFull.IsVisible = $false
    }

    if ($btnCopy) {
        $btnCopy.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            try {
                $st = $script:_detailsTagState
                if ($st -and $st.TxtValue -and $st.TxtValue.Text) { $st.TxtValue.Text | Set-Clipboard }
            } catch { Write-LogError "Details Copy failed" $_.Exception }
        }))
    }

    if ($btnFull) {
        $btnFull.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            try {
                $st = $script:_detailsTagState
                if (-not $st -or -not $st.SelectedItem) { return }
                Write-Status "Get full object $($st.SelectedItem.Name)"
                [void]$st.SelectedItem.Get()
                if ($st.SelectedItem.IsFullObject -and $st.TxtValue) {
                    $st.TxtValue.Text = [string]$st.SelectedItem.JsonString
                }
                Write-Status ""
            } catch { Write-LogError "Load full failed" $_.Exception; Write-Status "" }
        }))
    }

    # ---- Settings tab ------------------------------------------------------
    $txtName = $hostType::FindByName($detailsForm, 'txtObjectName')
    $txtDesc = $hostType::FindByName($detailsForm, 'txtObjectDescription')
    if ($txtName) { $txtName.Text = [string]$selectedItem.Name }
    if ($txtDesc) { $txtDesc.Text = [string]$selectedItem.Description }

    # Snapshots for the save-confirmation dialog: it names the object and
    # says exactly what changed, computed against these open-time values via
    # $script: state (reliable regardless of handler capture behavior).
    $state.TxtName      = $txtName
    $state.TxtDesc      = $txtDesc
    $state.OriginalName = if ($txtName) { [string]$txtName.Text } else { [string]$selectedItem.Name }
    $state.OriginalDesc = if ($txtDesc) { [string]$txtDesc.Text } else { '' }

    $pnlScopeTagsLabel = $hostType::FindByName($detailsForm, 'pnlScopeTagsLabel')
    $grdScopeTags      = $hostType::FindByName($detailsForm, 'grdScopeTags')
    $lstAvail          = $hostType::FindByName($detailsForm, 'lstAvailableScopeTags')
    $lstAssigned       = $hostType::FindByName($detailsForm, 'lstAssignedScopeTags')
    $btnTagAssign      = $hostType::FindByName($detailsForm, 'btnScopeTagAssign')
    $btnTagUnassign    = $hostType::FindByName($detailsForm, 'btnScopeTagUnassign')
    $btnSettingsSave   = $hostType::FindByName($detailsForm, 'btnObjectSettingsSave')

    if ($selectedItem.PolicyType.ScopeTagProperty) {
        $state.ScopeTagPropName = [string]$selectedItem.PolicyType.ScopeTagProperty
    }

    if (-not $state.ScopeTagPropName) {
        if ($pnlScopeTagsLabel) { $pnlScopeTagsLabel.IsVisible = $false }
        if ($grdScopeTags)      { $grdScopeTags.IsVisible      = $false }
    } else {
        $allScopeTags = @()
        try {
            $depObjects = Get-GraphDependencySourceObjects $selectedItem -DefaultPoliciesOnly
            if ($depObjects -and $depObjects.ContainsKey('ScopeTags')) {
                $allScopeTags = @($depObjects['ScopeTags'])
            }
        } catch { Write-LogDebug "Scope tag dependency cache lookup failed: $($_.Exception.Message)" }

        if ($allScopeTags.Count -eq 0) {
            try {
                $allScopeTags = @(Get-GraphPolicies -PolicyType 'ScopeTags' -TokenId $selectedItem._TokenId)
            } catch { Write-LogError "Failed to load scope tags" $_.Exception }
        }

        $currentIds = $null
        if ($selectedItem.Object -and $selectedItem.Object.PSObject.Properties[$state.ScopeTagPropName]) {
            $currentIds = $selectedItem.Object.($state.ScopeTagPropName)
        }
        foreach ($id in @($currentIds)) {
            if ($null -ne $id) { [void]$state.AssignedIds.Add([string]$id) }
        }
        $state.OriginalAssignedIds = [System.Collections.Generic.HashSet[string]]::new($state.AssignedIds)

        # Project to CopyDialogScopeTagItem (already a CLR class with Id/Name) —
        # same shape works here. Dedupe by Id (cache may carry both the
        # synthetic Default and a real Default tag).
        $dedup  = @()
        $seenIds = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($tag in $allScopeTags) {
            $idStr = [string]$tag.Id
            if (-not $seenIds.Add($idStr)) { continue }
            $dedup += [CopyDialogScopeTagItem]@{ Id = $idStr; Name = [string]$tag.Name }
        }
        $state.AllScopeTags = @($dedup | Sort-Object Name)


        if ($lstAvail) {
            $lstAvail.ItemsSource  = $state.AvailableCollection
        }
        if ($lstAssigned) {
            $lstAssigned.ItemsSource  = $state.AssignedCollection
        }
        $state.LstAvail    = $lstAvail
        $state.LstAssigned = $lstAssigned

        # Helpers live in $script: and read $script: state only - local
        # captures would be lost the moment an event handler runs them.
        $script:_detailsSyncTagLists = {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            $st.AvailableCollection.Clear()
            $st.AssignedCollection.Clear()
            foreach ($tag in $st.AllScopeTags) {
                if ($st.AssignedIds.Contains($tag.Id)) {
                    [void]$st.AssignedCollection.Add($tag)
                } else {
                    [void]$st.AvailableCollection.Add($tag)
                }
            }
        }
        & $script:_detailsSyncTagLists

        $script:_detailsMoveTags = {
            param([string]$Action, [string[]]$Ids)
            $st = $script:_detailsTagState
            if (-not $st -or $Ids.Count -eq 0) { return }
            if ($Action -eq 'assign') {
                foreach ($id in $Ids) { [void]$st.AssignedIds.Add([string]$id) }
            } else {
                foreach ($id in $Ids) { [void]$st.AssignedIds.Remove([string]$id) }
            }
            & $script:_detailsSyncTagLists
        }

        if ($btnTagAssign) {
            $btnTagAssign.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                $st = $script:_detailsTagState
                if (-not $st) { return }
                try {
                    $st = $script:_detailsTagState
                    if (-not $st -or -not $st.LstAvail -or $st.LstAvail.SelectedItems.Count -eq 0) {
                        Write-Status "Select one or more scope tag(s) in the Available list first"
                        return
                    }
                    $ids = @($st.LstAvail.SelectedItems | ForEach-Object { [string]$_.Id })
                    & $script:_detailsMoveTags 'assign' $ids
                } catch { Write-LogError "Scope-tag assign handler failed" $_.Exception }
            }))
        }
        if ($btnTagUnassign) {
            $btnTagUnassign.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                $st = $script:_detailsTagState
                if (-not $st) { return }
                try {
                    $st = $script:_detailsTagState
                    if (-not $st -or -not $st.LstAssigned -or $st.LstAssigned.SelectedItems.Count -eq 0) {
                        Write-Status "Select one or more scope tag(s) in the Assigned list first"
                        return
                    }
                    $ids = @($st.LstAssigned.SelectedItems | ForEach-Object { [string]$_.Id })
                    & $script:_detailsMoveTags 'unassign' $ids
                } catch { Write-LogError "Scope-tag unassign handler failed" $_.Exception }
            }))
        }

        # Avalonia's MouseDoubleClick equivalent on ListBox is DoubleTapped.
        if ($lstAvail) {
            $lstAvail.add_DoubleTapped((ConvertTo-AvaloniaEventScriptBlock {
                $st = $script:_detailsTagState
                if (-not $st) { return }
                try {
                    $st = $script:_detailsTagState
                    if (-not $st -or -not $st.LstAvail -or $st.LstAvail.SelectedItems.Count -eq 0) { return }
                    $ids = @($st.LstAvail.SelectedItems | ForEach-Object { [string]$_.Id })
                    & $script:_detailsMoveTags 'assign' $ids
                } catch { Write-LogError "Scope-tag double-tap (available) failed" $_.Exception }
            }))
        }
        if ($lstAssigned) {
            $lstAssigned.add_DoubleTapped((ConvertTo-AvaloniaEventScriptBlock {
                $st = $script:_detailsTagState
                if (-not $st) { return }
                try {
                    $st = $script:_detailsTagState
                    if (-not $st -or -not $st.LstAssigned -or $st.LstAssigned.SelectedItems.Count -eq 0) { return }
                    $ids = @($st.LstAssigned.SelectedItems | ForEach-Object { [string]$_.Id })
                    & $script:_detailsMoveTags 'unassign' $ids
                } catch { Write-LogError "Scope-tag double-tap (assigned) failed" $_.Exception }
            }))
        }
    }

    if ($btnSettingsSave) {
        $btnSettingsSave.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            try {
                # Name the object and say what changed - all read from the
                # $script: state snapshots taken at form open.
                $st = $script:_detailsTagState
                $originalName = if ($st -and $st.OriginalName) { [string]$st.OriginalName } else { [string]$st.SelectedItem.Name }
                $changes = @()
                if ($st -and $st.TxtName) {
                    $newName = [string]$st.TxtName.Text
                    if ($newName -and $newName -cne $originalName) { $changes += "New name: '$newName'" }
                }
                if ($st -and $st.TxtDesc -and ([string]$st.TxtDesc.Text) -cne [string]$st.OriginalDesc) {
                    $changes += "Description updated"
                }
                if ($st -and $null -ne $st.OriginalAssignedIds -and $null -ne $st.AssignedIds) {
                    $added   = @($st.AssignedIds         | Where-Object { -not $st.OriginalAssignedIds.Contains($_) }).Count
                    $removed = @($st.OriginalAssignedIds | Where-Object { -not $st.AssignedIds.Contains($_) }).Count
                    if ($added -or $removed) { $changes += "Scope tags: $added added, $removed removed" }
                }
                $confirmMsg = "Are you sure you want to update '$originalName'?"
                if ($changes.Count) { $confirmMsg += "`n`n" + ($changes -join "`n") }

                if (($ui.ShowMessageBox($confirmMsg, "Update object information?", "YesNo", "Warning")) -ne "Yes") {
                    return
                }

                Write-Status "Update object information for $($st.SelectedItem.Name)"
                $nameValue = if ($st.TxtName) { [string]$st.TxtName.Text } else { '' }
                if (-not $nameValue) {
                    $ui.ShowMessageBox("Name property must not be empty!", "Error", "OK", "Error")
                    Write-Status ""
                    return
                }

                $nameProp = (?? $st.SelectedItem.PolicyType.NameProperty "displayName")
                $idProp   = (?? $st.SelectedItem.PolicyType.IDProperty "id")

                $updateHT = @{}
                $updateHT.Add($idProp, $st.SelectedItem.ID)
                $updateHT.Add($nameProp, $nameValue)

                if ($st.SelectedItem.Object."@odata.type") {
                    $updateHT.Add("@odata.type", $st.SelectedItem.JsonObject."@odata.type")
                }

                if ($st.TxtDesc -and $st.TxtDesc.IsEnabled -eq $true) {
                    $descValue = [string]$st.TxtDesc.Text
                    if (-not $descValue -and $st.SelectedItem.Description) {
                        $updateHT["description"] = ""
                    } elseif ($descValue) {
                        $updateHT["description"] = $descValue
                    }
                }

                if ($st.ScopeTagPropName -and $null -ne $st.AssignedIds) {
                    $assignedIds = @($st.AssignedIds)

                    # Soft guard: Intune treats roleScopeTagIds as containing "0"
                    # (Default) by default. A non-empty list omitting Default is
                    # usually rejected by Graph; warn before letting the PATCH
                    # 400 silently. Empty list isn't flagged — that's a deliberate
                    # reset.
                    if ($assignedIds.Count -gt 0 -and $assignedIds -notcontains "0") {
                        $proceed = $ui.ShowMessageBox("The 'Default' scope tag (Id 0) is not in the assigned list.`n`nIntune usually requires Default to be present and the API call may fail.`n`nContinue anyway?", "Default scope tag missing", "YesNo", "Warning")
                        if ($proceed -ne "Yes") { Write-Status ""; return }
                    }
                    $updateHT.Add($st.ScopeTagPropName, $assignedIds)
                }

                $updateObj = [PSCustomObject]$updateHT
                $api = "$($st.SelectedItem.PolicyType.API)/$($st.SelectedItem.ID)"
                $json = $updateObj | ConvertTo-Json -Depth 20

                $ret = Invoke-MSGraphAPI $api -HttpMethod "PATCH" -Content $json -FullResponseObject -TokenId $st.SelectedItem._TokenId
                Write-Status ""

                if ($ret.Success -eq $false) {
                    Write-Log "Failed to update object information for $($st.SelectedItem.Name). Error: $($ret.StatusCode) - $($ret.StatusDescription)"
                    $ui.ShowMessageBox("Object information could not be verified!`n`nCheck the log file", "Update warning", "OK", "Warning")
                } else {
                    Write-Log "Object information updated for $($st.SelectedItem.Name)"

                    # Refresh in-memory object so the parent grid's ScopeTags
                    # column reflects the new value. Clearing _ScopeTags forces
                    # the lazy getter to re-resolve. Avalonia DataGrid has no
                    # Items.Refresh — rebind ItemsSource to force re-pull (the
                    # row item is also a CLR class without INPC).
                    if ($st.ScopeTagPropName -and $null -ne $st.AssignedIds) {
                        $assignedIds = @($st.AssignedIds)
                        try { $st.SelectedItem.Object.($st.ScopeTagPropName) = $assignedIds } catch { }
                        $st.SelectedItem._ScopeTags       = $null
                        $st.SelectedItem._ScopeTagsString = $null
                    }
                    try {
                        # IntuneObjectRowItem carries SNAPSHOT strings copied at
                        # build time, so re-binding alone still showed the old
                        # name / scope tags. Update the row that wraps this
                        # policy before the rebind.
                        $current = $script:dgIntuneManagerObjects.ItemsSource
                        foreach ($row in @($current)) {
                            if (-not $row -or -not [object]::ReferenceEquals($row.Source, $st.SelectedItem)) { continue }
                            $row.Name = [string]$st.SelectedItem.Name
                            try { $row.ScopeTags = [string]$st.SelectedItem.ScopeTags } catch { }
                        }
                        $script:dgIntuneManagerObjects.ItemsSource = $null
                        $script:dgIntuneManagerObjects.ItemsSource = $current
                    } catch { Write-LogDebug "DataGrid refresh failed: $($_.Exception.Message)" }

                    $st.InitialSettingsSig = Get-DetailsSettingsSignature -State $st
                }
            } catch { Write-LogError "Settings Save failed" $_.Exception; Write-Status "" }
        }))
    }

    # ---- Columns tab -------------------------------------------------------
    $lstBasic       = $hostType::FindByName($detailsForm, 'lstBasicProperties')
    $lstObjProps    = $hostType::FindByName($detailsForm, 'lstObjectProperties')
    $lstObjCols     = $hostType::FindByName($detailsForm, 'lstObjectColumns')
    $btnBasicAdd    = $hostType::FindByName($detailsForm, 'btnBasicColumnsAdd')
    $btnObjAdd      = $hostType::FindByName($detailsForm, 'btnObjectColumnsAdd')
    $btnColUp       = $hostType::FindByName($detailsForm, 'btnObjectColumnsMoveUp')
    $btnColDown     = $hostType::FindByName($detailsForm, 'btnObjectColumnsMoveDown')
    $btnColDelete   = $hostType::FindByName($detailsForm, 'btnObjectColumnsDelete')
    $btnColClear    = $hostType::FindByName($detailsForm, 'btnObjectColumnsClear')
    $btnColReset    = $hostType::FindByName($detailsForm, 'btnObjectColumnsReset')
    $btnColSave     = $hostType::FindByName($detailsForm, 'btnObjectColumnsSave')
    $grdObjCols     = $hostType::FindByName($detailsForm, 'grdObjectColumns')
    $chkOverride    = $hostType::FindByName($detailsForm, 'chkObjectColumnOverride')
    $lblColConfig   = $hostType::FindByName($detailsForm, 'lblObjectColumnsConfig')

    $state.LstBasic    = $lstBasic
    $state.LstObjProps = $lstObjProps
    $state.LstObjCols  = $lstObjCols
    $state.GrdObjCols  = $grdObjCols
    $state.ChkOverride = $chkOverride

    $skipBasic = @("JsonObject", "Object", "JsonString", "TenantId", "TokenId", "IsSelected")
    $arrBasic = foreach ($prop in ($selectedItem.PSObject.Properties | Where-Object { $skipBasic -notcontains $_.Name })) {
        [DetailsPropertyItem]@{ Name = $prop.Name; Value = $prop; Source = "Basic" }
    }
    # Typed cast, not a bare pipeline: Sort-Object re-wraps each item in a
    # PSObject, which Avalonia's binder cannot see through, so the list would
    # render blank rows against DisplayMemberBinding.
    $arrBasic = [DetailsPropertyItem[]]@($arrBasic | Sort-Object -Property Name)

    $skipObj = @("@odata.editLink")
    $arrObjProps = foreach ($prop in ($selectedItem.Object.PSObject.Properties | Where-Object { $skipObj -notcontains $_.Name -and $_.Name -notlike "*?@odata.*" })) {
        [DetailsPropertyItem]@{ Name = $prop.Name; Value = $prop; Source = "Object" }
    }
    $arrObjProps = [DetailsPropertyItem[]]@($arrObjProps | Sort-Object -Property Name)

    # Display fields come from DisplayMemberBinding in ObjectDetails.axaml.
    if ($lstBasic) {
        $lstBasic.ItemsSource  = $arrBasic
    }
    if ($lstObjProps) {
        $lstObjProps.ItemsSource  = $arrObjProps
    }
    if ($lstObjCols) {
        $lstObjCols.ItemsSource  = $state.ColObjectProperties
        $lstObjCols.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            if ($st.GrdObjCols) { $st.GrdObjCols.DataContext = $st.LstObjCols.SelectedItem }
        }))
    }

    if ($btnBasicAdd) {
        $btnBasicAdd.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            $sel = if ($st.LstBasic) { $st.LstBasic.SelectedItem } else { $null }
            if ($sel) { $st.ColObjectProperties.Add(([ObjectColumnInfo]::new($sel.Name, ""))) }
        }))
    }
    if ($btnObjAdd) {
        $btnObjAdd.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            $sel = if ($st.LstObjProps) { $st.LstObjProps.SelectedItem } else { $null }
            if ($sel) { $st.ColObjectProperties.Add(([ObjectColumnInfo]::new("Object." + $sel.Name, ""))) }
        }))
    }
    if ($btnColUp) {
        $btnColUp.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            if (-not $st.LstObjCols) { return }
            $idx = $st.LstObjCols.SelectedIndex
            if ($idx -gt 0) {
                $tmp = $st.ColObjectProperties[$idx]
                $st.ColObjectProperties.RemoveAt($idx)
                $st.ColObjectProperties.Insert(($idx - 1), $tmp)
                $st.LstObjCols.SelectedIndex = $idx - 1
            }
        }))
    }
    if ($btnColDown) {
        $btnColDown.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            if (-not $st.LstObjCols) { return }
            $idx = $st.LstObjCols.SelectedIndex
            if ($idx -ge 0 -and $idx -lt ($st.ColObjectProperties.Count - 1)) {
                $tmp = $st.ColObjectProperties[$idx]
                $st.ColObjectProperties.RemoveAt($idx)
                $st.ColObjectProperties.Insert(($idx + 1), $tmp)
                $st.LstObjCols.SelectedIndex = $idx + 1
            }
        }))
    }
    if ($btnColDelete) {
        $btnColDelete.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            if (-not $st.LstObjCols) { return }
            $idx = $st.LstObjCols.SelectedIndex
            if ($idx -ge 0) {
                if (($ui.ShowMessageBox("Are you sure you want to remove selected column?", "Remove Columns?", "YesNo", "Warning")) -ne "Yes") {
                    return
                }
                $st.ColObjectProperties.RemoveAt($idx)
            }
        }))
    }
    if ($btnColClear) {
        $btnColClear.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            if (($ui.ShowMessageBox("Are you sure you want to clear custom column settings?", "Clear Custom Columns?", "YesNo", "Warning")) -ne "Yes") {
                return
            }
            $st.ColObjectProperties.Clear()
        }))
    }
    if ($btnColReset) {
        $btnColReset.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            $st.ColObjectProperties.Clear()
            Show-DetailsObjectDefaultColumns -State $st -Reset
        }))
    }
    if ($btnColSave) {
        $btnColSave.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_detailsTagState
            if (-not $st) { return }
            try {
                $selectedType = Get-DetailsSelectedObjectTypeString
                if (-not $selectedType) { return }

                if (($ui.ShowMessageBox("Are you sure you want to save custom column settings?", "Save Custom Columns?", "YesNo", "Warning")) -ne "Yes") {
                    return
                }

                if ($st.ColObjectProperties.Count -gt 0) {
                    $arrCols = @()
                    if ($st.ChkOverride -and $st.ChkOverride.IsChecked -eq $true) { $arrCols += "0" }
                    foreach ($col in $st.ColObjectProperties) {
                        $tmp = $col.Property
                        if ($col.Header -and $col.Header -cne $col.Property) {
                            $tmp = "$($tmp)=$($col.Header)"
                        }
                        $arrCols += $tmp
                    }
                    Save-SettingStoreValue "IntuneManager\ObjectColumns\$selectedType" "$($script:IntuneManagerSelectedObject.Id)" ($arrCols -join ",")
                } else {
                    Remove-SettingStoreValue "IntuneManager\ObjectColumns\$selectedType" "$($script:IntuneManagerSelectedObject.Id)"
                }

                Show-DetailsObjectDefaultColumns -State $st
                $st.InitialColumnsSig = Get-DetailsColumnsSignature -State $st
            } catch { Write-LogError "Columns Save failed" $_.Exception }
        }))
    }

    Show-DetailsObjectDefaultColumns -State $state
    $state.InitialSettingsSig = Get-DetailsSettingsSignature -State $state
    $state.InitialColumnsSig  = Get-DetailsColumnsSignature  -State $state

    # Hook ConfirmClose so the Show-ModalForm close button checks for unsaved
    # changes. Mirrors the WPF impl which attached this as a ScriptMethod.
    #
    # PowerShell ScriptMethod scriptblocks invoked through Add-Member execute
    # outside the defining module's session state — module-scoped functions
    # like Test-DetailsViewHasUnsavedChanges / Show-MessageBox aren't visible
    # by name. Capture them as module-bound scriptblocks first (`${function:X}`
    # preserves the binding) and invoke via `&` from inside the closure.
    #
    # NOTE: ConvertTo-AvaloniaEventScriptBlock would STRIP these closure
    # captures (NewBoundScriptBlock loses GetNewClosure state), so the SB
    # would fail with "expression after & not valid" when invoked. The
    # captured ${function:X} SBs already carry their own module binding,
    # so we don't need ConvertTo here — plain GetNewClosure is sufficient.
    $sbTestDirty   = ${function:Test-DetailsViewHasUnsavedChanges}
    $sbShowMsgBox  = ${function:Show-MessageBox}
    $sbClearState  = ${function:Clear-DetailsTagState}
    $detailsForm | Add-Member -MemberType ScriptMethod -Name ConfirmClose -Value ({
        # Release the state on every path that actually closes - it pins the
        # selected policy including its full JSON after "Load full".
        if (-not (& $sbTestDirty -State $state)) { & $sbClearState; return $true }
        $result = & $sbShowMsgBox `
            -Text "You have unsaved changes in Settings or Columns.`n`nClose without saving?" `
            -Caption "Unsaved changes" -Button "YesNo" -Icon "Warning"
        if ($result -eq "Yes") { & $sbClearState; return $true }
        return $false
    }.GetNewClosure()) -Force

    # Per-policy-type Details buttons (Download/Edit for Script/PolicyFile,
    # Upload/Download for Application). Routes through Add-AvaloniaDetailsButton
    # to insert into pnlButtons. Failures here must not block the modal.
    if ($selectedItem.PolicyType.AddUIDetailsExtension) {
        try { $selectedItem.PolicyType.AddUIDetailsExtension($detailsForm) }
        catch { Write-LogError "AddUIDetailsExtension failed for $($selectedItem.PolicyType.Name)" $_.Exception }
    }

    $ui.ShowModalForm($FormTitle, $detailsForm)
}

# ---------------------------------------------------------------------------
# Helpers (private to this file). State is passed in explicitly rather than
# pulled from $script: vars so a future "open two details dialogs at once"
# flow doesn't require restructuring.
# ---------------------------------------------------------------------------

function Get-DetailsSelectedObjectTypeString
{
    if ($script:IntuneManagerSelectedObject -is [IntunePolicyGroupBase]) { return "PolicyGroup" }
    if ($script:IntuneManagerSelectedObject -is [IntunePolicyTypeBase])  { return "PolicyType" }
    return ""
}

function Show-DetailsObjectDefaultColumns
{
    param(
        [Parameter(Mandatory)] $State,
        [switch]$Reset
    )

    $hostType = Get-AvaloniaHost
    $lblColConfig = $hostType::FindByName($State.Form, 'lblObjectColumnsConfig')
    $chkOverride  = $hostType::FindByName($State.Form, 'chkObjectColumnOverride')

    if ($Reset -ne $true) {
        $selectedType = Get-DetailsSelectedObjectTypeString
        if (-not $selectedType) { return }
        $strColSettings = Get-SettingStoreValue "IntuneManager\ObjectColumns\$selectedType" "$($script:IntuneManagerSelectedObject.Id)"
    } else {
        $strColSettings = ""
    }

    $State.ColObjectProperties.Clear()
    $defaultColumns = $script:IntuneManagerSelectedObject.ViewProperties

    if ($strColSettings) {
        $arrColSettings = $strColSettings -split ",|;"
        if ($chkOverride) {
            $chkOverride.IsChecked = ($arrColSettings.Count -gt 0 -and $arrColSettings[0] -eq "0")
        }

        $start = 0
        if ($arrColSettings.Count -gt 0 -and ($arrColSettings[0] -eq "0" -or $arrColSettings[0] -eq "1")) {
            $start = 1
        }

        $colArr = @()
        for ($i = $start; $i -lt $arrColSettings.Count; $i++) {
            $colProp, $colHeader = $arrColSettings[$i].Split("=")
            if (-not $colHeader) { $colHeader = $colProp }
            $State.ColObjectProperties.Add([ObjectColumnInfo]::new($colProp, $colHeader))
            $colArr += $colProp
        }

        if ($arrColSettings.Count -eq 0 -or $arrColSettings[0] -ne "0") {
            $colArr = @($defaultColumns) + $colArr
        }

        if ($lblColConfig) { $lblColConfig.Text = ($colArr -join ',') }
    } else {
        if ($lblColConfig) { $lblColConfig.Text = "$($defaultColumns -join ',') (Default)" }
    }
}

function Get-DetailsSettingsSignature
{
    param([Parameter(Mandatory)] $State)

    if (-not $State.Form) { return "" }
    $hostType = Get-AvaloniaHost
    $txtName  = $hostType::FindByName($State.Form, 'txtObjectName')
    $txtDesc  = $hostType::FindByName($State.Form, 'txtObjectDescription')

    $scopeTagIds = @()
    if ($State.ScopeTagPropName -and $null -ne $State.AssignedIds) {
        $scopeTagIds = @($State.AssignedIds | Sort-Object)
    }

    $descriptionEnabled = $false
    if ($txtDesc) { $descriptionEnabled = [bool]$txtDesc.IsEnabled }

    $signature = [PSCustomObject]@{
        Name               = if ($txtName) { [string]$txtName.Text } else { "" }
        DescriptionEnabled = $descriptionEnabled
        Description        = if ($descriptionEnabled -and $txtDesc) { [string]$txtDesc.Text } else { $null }
        ScopeTagProperty   = [string]$State.ScopeTagPropName
        ScopeTagIds        = $scopeTagIds
    }
    return ($signature | ConvertTo-Json -Depth 10 -Compress)
}

function Get-DetailsColumnsSignature
{
    param([Parameter(Mandatory)] $State)

    if (-not $State.Form) { return "" }
    $hostType = Get-AvaloniaHost
    $chkOverride = $hostType::FindByName($State.Form, 'chkObjectColumnOverride')

    $columns = @()
    foreach ($col in @($State.ColObjectProperties)) {
        if (-not $col) { continue }
        $columns += [PSCustomObject]@{ Property = [string]$col.Property; Header = [string]$col.Header }
    }

    $signature = [PSCustomObject]@{
        OverrideDefaultColumns = if ($chkOverride) { [bool]$chkOverride.IsChecked } else { $false }
        Columns                = $columns
    }
    return ($signature | ConvertTo-Json -Depth 20 -Compress)
}

function Test-DetailsViewHasUnsavedChanges
{
    param([Parameter(Mandatory)] $State)

    if (-not $State.Form) { return $false }
    $settingsDirty = $false
    $columnsDirty  = $false

    try {
        if ($State.InitialSettingsSig) {
            $settingsDirty = ((Get-DetailsSettingsSignature -State $State) -ne $State.InitialSettingsSig)
        }
    } catch { Write-LogDebug "Details settings dirty check failed: $($_.Exception.Message)" }

    try {
        if ($State.InitialColumnsSig) {
            $columnsDirty = ((Get-DetailsColumnsSignature -State $State) -ne $State.InitialColumnsSig)
        }
    } catch { Write-LogDebug "Details columns dirty check failed: $($_.Exception.Message)" }

    return ($settingsDirty -or $columnsDirty)
}

function Clear-DetailsTagState
{
    # Called from the ConfirmClose ScriptMethod, which runs outside the module's
    # session state - a bare $script: assignment there would target the caller's
    # scope, not the module's. Capture this as ${function:Clear-DetailsTagState}
    # and invoke it with & so the write lands in module scope.
    $script:_detailsTagState = $null
}
