# Avalonia port of the Copy dialog from UI/WPF/Extensions/IntuneManagerUI.ps1
# (Copy-IntuneManagerPolicy). New file because Copy is a self-contained
# subsystem and the WPF original sits in the giant CoreUI/IntuneManagerUI
# files we're trying not to grow further (architecture rule R9).
#
# Slice 3b1: core copy without scope tags.
# Slice 3b2: scope-tag dual-list. Panels in CopyDialog.axaml shown only
# when the source policy type defines a ScopeTagProperty; mirrors the
# WPF impl with the canonical assigned-ids HashSet + ObservableCollections
# and re-seeds by NAME on tenant change so cross-tenant copies pre-pick
# the same-named tags in the destination.

function Copy-IntuneManagerPolicy
{
    $ui = $script:UIProvider
    if (-not $script:dgIntuneManagerObjects -or -not $script:dgIntuneManagerObjects.SelectedItem) {
        $ui.ShowMessageBox("No object selected`n`nSelect the $($script:IntuneManagerSelectedObject.Title) item you want to copy", "Error", "OK", "Error")
        return
    }

    $sourceRow = $script:dgIntuneManagerObjects.SelectedItem
    $sourceItem = $sourceRow.Source
    if (-not $sourceItem) {
        Write-Log "Copy: selected row carries no policy object" 3
        $ui.ShowMessageBox("Cannot copy: the selected item has no underlying policy object.", "Copy", "OK", "Error")
        return
    }

    $copyForm = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/CopyDialog.axaml'))
    if (-not $copyForm) { return }

    $hostType = Get-AvaloniaHost
    $cbDestTenant = $hostType::FindByName($copyForm, 'cbDestinationTenant')
    $txtName      = $hostType::FindByName($copyForm, 'txtObjectName')
    $txtDesc      = $hostType::FindByName($copyForm, 'txtObjectDescription')
    $btnOk        = $hostType::FindByName($copyForm, 'btnOk')
    $btnCancel    = $hostType::FindByName($copyForm, 'btnCancel')
    $pnlScopeTagsLabel = $hostType::FindByName($copyForm, 'pnlCopyScopeTagsLabel')
    $grdScopeTags      = $hostType::FindByName($copyForm, 'grdCopyScopeTags')
    $lstAvailable      = $hostType::FindByName($copyForm, 'lstCopyAvailableScopeTags')
    $lstAssigned       = $hostType::FindByName($copyForm, 'lstCopyAssignedScopeTags')
    $btnAssign         = $hostType::FindByName($copyForm, 'btnCopyScopeTagAssign')
    $btnUnassign       = $hostType::FindByName($copyForm, 'btnCopyScopeTagUnassign')

    # Default new-name pattern. PolicyType.CopyDefaultName allows %Name%-style
    # placeholders against properties of the source row; mirror the WPF
    # substitution loop. Falls back to "Original - Copy".
    $newName = "$($sourceRow.Name) - Copy"
    $copyDefault = $null
    try {
        if ($sourceItem.PolicyType -and $sourceItem.PolicyType.CopyDefaultName) {
            $copyDefault = $sourceItem.PolicyType.CopyDefaultName
        }
    } catch { }
    if ($copyDefault) {
        $newName = $copyDefault
        foreach ($p in $sourceRow.PSObject.Properties) {
            $val = [string]$sourceRow.$($p.Name)
            $newName = $newName -replace "%$($p.Name)%", $val
        }
    }
    if ($txtName) { $txtName.Text = $newName }

    # Description gating: HasDescription=$false on the type means the policy
    # has no description field at all; greyed out so the user knows it won't
    # round-trip.
    $hasDescription = $true
    try {
        if ($sourceItem.HasDescription -eq $false) { $hasDescription = $false }
    } catch { }
    if ($txtDesc) {
        if ($hasDescription) {
            try { $txtDesc.Text = [string]$sourceItem.Description } catch { }
        } else {
            $txtDesc.IsEnabled = $false
        }
    }

    # Tenant combo: project Get-TokenInfo entries into CopyDialogTenantItem so
    # the ItemTemplate can bind to TenantNameEx (Avalonia binding doesn't
    # walk PSObject NoteProperty — see [[avalonia-binding-needs-clr-types]]).
    $tenantItems = @()
    foreach ($t in (Get-TokenInfo)) {
        if (-not $t) { continue }
        $name = [string]$t.TenantName
        if ($t.IsDefault) { $name = "$name (Current)" }
        $tenantItems += [CopyDialogTenantItem]@{
            TenantNameEx = $name
            Id           = if ($null -ne $t.Id) { [int]$t.Id } else { 0 }
            IsDefault    = [bool]$t.IsDefault
            Source       = $t
        }
    }
    if ($cbDestTenant) {
        $cbDestTenant.ItemsSource = $tenantItems
        $defaultItem = $tenantItems | Where-Object { $_.IsDefault } | Select-Object -First 1
        if (-not $defaultItem -and $tenantItems.Count -gt 0) { $defaultItem = $tenantItems[0] }
        if ($defaultItem) { $cbDestTenant.SelectedItem = $defaultItem }
    }

    # ---- Scope tags (dual-list) ----------------------------------------------
    # Mirrors WPF UI/WPF/Extensions/IntuneManagerUI.ps1 ~lines 1947-2114. Single
    # source of truth = the AssignedIds HashSet on $state; ObservableCollections
    # are re-derived on every change. Cross-tenant assignment is matched by
    # NAME so picking a different tenant still pre-checks same-named tags.
    $state = @{
        ScopeTagPropName       = $null
        SourceScopeTagNames    = @()
        AllScopeTags           = @()
        AssignedIds            = [System.Collections.Generic.HashSet[string]]::new()
        AvailableCollection    = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
        AssignedCollection     = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
        LstAvail               = $lstAvailable
        LstAssigned            = $lstAssigned
    }
    # Event handlers must reach this through $script: - closure captures do
    # not survive ConvertTo-AvaloniaEventScriptBlock (NewBoundScriptBlock
    # rebinds to the module and drops the .GetNewClosure() state).
    $script:_copyDialogTagState = $state

    try {
        if ($sourceItem.PolicyType -and $sourceItem.PolicyType.ScopeTagProperty) {
            $state.ScopeTagPropName = [string]$sourceItem.PolicyType.ScopeTagProperty
        }
    } catch { }

    if ($state.ScopeTagPropName) {
        # Capture source tag NAMES so cross-tenant re-seeds work.
        $sourceTagIds = @()
        try {
            if ($sourceItem.Object -and $sourceItem.Object.PSObject.Properties[$state.ScopeTagPropName]) {
                $sourceTagIds = @($sourceItem.Object.($state.ScopeTagPropName) | ForEach-Object { [string]$_ })
            }
        } catch { }
        try {
            $srcDeps = Get-GraphDependencySourceObjects $sourceItem -DefaultPoliciesOnly
            if ($srcDeps -and $srcDeps.ContainsKey('ScopeTags')) {
                foreach ($id in $sourceTagIds) {
                    $tag = @($srcDeps['ScopeTags']) | Where-Object { [string]$_.Id -eq $id } | Select-Object -First 1
                    if ($tag) { $state.SourceScopeTagNames += [string]$tag.Name }
                }
            }
        } catch { Write-LogDebug "Source scope-tag name capture failed: $($_.Exception.Message)" }

        if ($pnlScopeTagsLabel) { $pnlScopeTagsLabel.IsVisible = $true }
        if ($grdScopeTags)      { $grdScopeTags.IsVisible      = $true }


        if ($lstAvailable) {
            $lstAvailable.ItemsSource   = $state.AvailableCollection
        }
        if ($lstAssigned) {
            $lstAssigned.ItemsSource    = $state.AssignedCollection
        }

        # All three helpers live in $script: and read $script: state only -
        # local captures would be lost the moment an event handler runs them.
        $script:_copyDialogSyncTagLists = {
            $st = $script:_copyDialogTagState
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

        # Reload destination tenant's scope tags. Cache lookup first; live
        # fetch on cold cache. Dedupes by Id (cache may carry both the
        # synthetic Default and a real Default tag).
        $script:_copyDialogReloadTags = {
            param([int]$DestTokenId)

            $st = $script:_copyDialogTagState
            if (-not $st) { return }

            $destTags = @()
            try {
                $destTokenInfo = Get-TokenInfo $DestTokenId
                if ($destTokenInfo) {
                    $cacheId = "DependencyObjects_$($destTokenInfo.TenantId)"
                    $deps = Get-CacheObject $cacheId
                    if ($deps -and $deps.ContainsKey('ScopeTags')) {
                        $destTags = @($deps['ScopeTags'])
                    }
                }
            } catch { Write-LogDebug "Copy scope tags: cache lookup failed: $($_.Exception.Message)" }

            if ($destTags.Count -eq 0) {
                try {
                    $destTags = @(Get-GraphPolicies -PolicyType 'ScopeTags' -TokenId $DestTokenId)
                    $destTags = @([PSCustomObject]@{ ID = 0; Name = 'Default' }) + $destTags
                } catch { Write-LogError "Failed to load scope tags for destination tenant" $_.Exception }
            }

            $dedup = @()
            $seenIds = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($tag in $destTags) {
                $idStr = [string]$tag.Id
                if (-not $seenIds.Add($idStr)) { continue }
                $dedup += [CopyDialogScopeTagItem]@{ Id = $idStr; Name = [string]$tag.Name }
            }
            $st.AllScopeTags = @($dedup | Sort-Object Name)

            $st.AssignedIds.Clear()
            $seedNames = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($n in $st.SourceScopeTagNames) { [void]$seedNames.Add($n) }
            foreach ($tag in $st.AllScopeTags) {
                if ($seedNames.Contains($tag.Name)) { [void]$st.AssignedIds.Add($tag.Id) }
            }

            & $script:_copyDialogSyncTagLists
        }

        if ($cbDestTenant -and $cbDestTenant.SelectedItem) {
            & $script:_copyDialogReloadTags ([int]$cbDestTenant.SelectedItem.Id)
        }

        if ($cbDestTenant) {
            $cbDestTenant.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
                param($s, $e)
                if ($s.SelectedItem) { & $script:_copyDialogReloadTags ([int]$s.SelectedItem.Id) }
            }))
        }

        $script:_copyDialogMoveTags = {
            param([string]$Action, [string[]]$Ids)
            $st = $script:_copyDialogTagState
            if (-not $st -or $Ids.Count -eq 0) { return }
            if ($Action -eq 'assign') {
                foreach ($id in $Ids) { [void]$st.AssignedIds.Add([string]$id) }
            } else {
                foreach ($id in $Ids) { [void]$st.AssignedIds.Remove([string]$id) }
            }
            & $script:_copyDialogSyncTagLists
        }

        if ($btnAssign) {
            $btnAssign.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                try {
                    $st = $script:_copyDialogTagState
                    if (-not $st -or -not $st.LstAvail -or $st.LstAvail.SelectedItems.Count -eq 0) { return }
                    $ids = @($st.LstAvail.SelectedItems | ForEach-Object { [string]$_.Id })
                    & $script:_copyDialogMoveTags 'assign' $ids
                } catch { Write-LogError "Copy scope-tag assign handler failed" $_.Exception }
            }))
        }
        if ($btnUnassign) {
            $btnUnassign.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                try {
                    $st = $script:_copyDialogTagState
                    if (-not $st -or -not $st.LstAssigned -or $st.LstAssigned.SelectedItems.Count -eq 0) { return }
                    $ids = @($st.LstAssigned.SelectedItems | ForEach-Object { [string]$_.Id })
                    & $script:_copyDialogMoveTags 'unassign' $ids
                } catch { Write-LogError "Copy scope-tag unassign handler failed" $_.Exception }
            }))
        }

        # Avalonia's MouseDoubleClick equivalent on ListBox is DoubleTapped.
        if ($lstAvailable) {
            $lstAvailable.add_DoubleTapped((ConvertTo-AvaloniaEventScriptBlock {
                try {
                    $st = $script:_copyDialogTagState
                    if (-not $st -or -not $st.LstAvail -or $st.LstAvail.SelectedItems.Count -eq 0) { return }
                    $ids = @($st.LstAvail.SelectedItems | ForEach-Object { [string]$_.Id })
                    & $script:_copyDialogMoveTags 'assign' $ids
                } catch { Write-LogError "Copy scope-tag double-tap (available) failed" $_.Exception }
            }))
        }
        if ($lstAssigned) {
            $lstAssigned.add_DoubleTapped((ConvertTo-AvaloniaEventScriptBlock {
                try {
                    $st = $script:_copyDialogTagState
                    if (-not $st -or -not $st.LstAssigned -or $st.LstAssigned.SelectedItems.Count -eq 0) { return }
                    $ids = @($st.LstAssigned.SelectedItems | ForEach-Object { [string]$_.Id })
                    & $script:_copyDialogMoveTags 'unassign' $ids
                } catch { Write-LogError "Copy scope-tag double-tap (assigned) failed" $_.Exception }
            }))
        }
    }

    # Result box captured by the button closures — they set Confirmed before
    # tearing down the modal.
    # ShowModalForm is non-blocking, so the copy runs from the OK/Enter handlers
    # (Invoke-IntuneManagerCopyExecute) while state is live - stashed in script
    # scope so the module-scope helper can read the controls after dispatch.
    $script:_imCopyFormState = @{
        SourceRow    = $sourceRow
        SourceItem   = $sourceItem
        TxtName      = $txtName
        TxtDesc      = $txtDesc
        CbDestTenant = $cbDestTenant
        BtnOk        = $btnOk
        State        = $state
    }

    if ($btnOk) {
        $btnOk.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            try { Invoke-IntuneManagerCopyExecute } catch { Write-LogError "Copy execution failed" $_.Exception }
            Show-ModalObject
            $script:_imCopyFormState = $null
            $script:_copyDialogTagState = $null
            if ($script:dgIntuneManagerObjects) { $script:dgIntuneManagerObjects.Focus() }
        }.GetNewClosure()))
    }
    if ($btnCancel) {
        $btnCancel.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            Show-ModalObject
            $script:_imCopyFormState = $null
            $script:_copyDialogTagState = $null
            if ($script:dgIntuneManagerObjects) { $script:dgIntuneManagerObjects.Focus() }
        }.GetNewClosure()))
    }

    # Avalonia has no IsDefault/IsCancel on Button — handle Enter/Escape.
    $copyForm.add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
        param($s, $e)
        $stCopy = $script:_imCopyFormState
        if ($e.Key -eq [Avalonia.Input.Key]::Enter -and $stCopy -and $stCopy.BtnOk -and $stCopy.BtnOk.IsEnabled) {
            try { Invoke-IntuneManagerCopyExecute } catch { Write-LogError "Copy execution failed" $_.Exception }
            Show-ModalObject
            $script:_imCopyFormState = $null
            $script:_copyDialogTagState = $null
            $e.Handled = $true
        } elseif ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            Show-ModalObject
            $script:_imCopyFormState = $null
            $script:_copyDialogTagState = $null
            $e.Handled = $true
        }
    }.GetNewClosure()))

    $ui.ShowModalForm("Copy object", $copyForm, $true)

    # Focus + select the name so the user can type a new value immediately.
    if ($txtName) {
        try { $txtName.Focus() } catch { }
        try { $txtName.SelectAll() } catch { }
    }

}

function Invoke-IntuneManagerCopyExecute
{
    # Runs the copy for the live $script:_imCopyFormState. Called from the OK/Enter
    # handlers (ShowModalForm does not block, so it can't run after it returns).
    $st = $script:_imCopyFormState
    if (-not $st) { return }
    $ui           = $script:UIProvider
    $sourceRow    = $st.SourceRow
    $sourceItem   = $st.SourceItem
    $txtName      = $st.TxtName
    $txtDesc      = $st.TxtDesc
    $cbDestTenant = $st.CbDestTenant
    $state        = $st.State

    $finalName = if ($txtName) { [string]$txtName.Text } else { '' }
    if ([string]::IsNullOrWhiteSpace($finalName)) {
        Write-Log "New name cannot be empty. Copy object skipped" 2
        Write-Status ""
        return
    }

    $finalDesc = $null
    if ($txtDesc -and $txtDesc.IsEnabled) {
        $finalDesc = [string]$txtDesc.Text
    }

    $tokenId = $null
    if ($cbDestTenant -and $cbDestTenant.SelectedItem) {
        $tokenId = $cbDestTenant.SelectedItem.Id
    }

    Write-Status "Copy $($sourceRow.Name)"

    $copyArgs = @{
        InputObject = $sourceItem
        Name        = $finalName
        Description = $finalDesc
    }
    if ($null -ne $tokenId) { $copyArgs['TokenId'] = $tokenId }

    # Forward scope-tag picks (when section was shown). Pass even an empty
    # array so the new copy explicitly reflects user intent — Copy-GraphPolicy
    # only overrides the cloned JSON when this is non-null. Read from the
    # canonical AssignedIds set rather than the UI-derived collection.
    if ($state.ScopeTagPropName -and $null -ne $state.AssignedIds) {
        $copyArgs['ScopeTagIds'] = @($state.AssignedIds)
    }

    $newPolicies = $null
    try {
        $newPolicies = Copy-GraphPolicy @copyArgs
    } catch {
        Write-LogError "Copy-GraphPolicy failed" $_.Exception
        $ui.ShowMessageBox("Copy failed: $($_.Exception.Message)", "Error", "OK", "Error")
        Write-Status ""
        return
    }

    if ($newPolicies -and (Get-SettingValue "RefreshObjectsAfterCopy") -eq $true) {
        Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject
    }

    Write-Status ""
    if ($script:dgIntuneManagerObjects) { $script:dgIntuneManagerObjects.Focus() }
}
