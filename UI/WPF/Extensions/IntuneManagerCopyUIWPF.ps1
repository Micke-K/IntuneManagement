# Per-view Copy dialog (Copy-IntuneManagerPolicy).
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Copy-IntuneManagerPolicy
{
    if(-not $script:dgIntuneManagerObjects.SelectedItem)
    {
        $script:UIProvider.ShowMessageBox("No object selected`n`nSelect the $($script:IntuneManagerSelectedObject.Title) item you want to copy", "Error", "OK", "Error")
        return
    }

    $script:copyForm = Initialize-Window ($script:AppUIRootFolder + "\Xaml\CopyDialog.xaml")

    if(-not $script:copyForm) { return }

    $newName = "$($script:dgIntuneManagerObjects.SelectedItem.Name) - Copy"
    if($script:dgIntuneManagerObjects.SelectedItem.PolicyType.CopyDefaultName)
    {
        $newName = $script:dgIntuneManagerObjects.SelectedItem.PolicyType.CopyDefaultName
        $script:dgIntuneManagerObjects.SelectedItem.PSObject.Properties | ForEach-Object { $newName =  $newName -replace "%$($_.Name)%", $script:dgIntuneManagerObjects.SelectedItem."$($_.Name)" }
    }

    $script:UIProvider.SetXamlProperty($script:copyForm, "txtObjectName", "Text", $newName)
    if($script:dgIntuneManagerObjects.SelectedItem.HasDescription -ne $false)
    {
        $script:UIProvider.SetXamlProperty($script:copyForm, "txtObjectDescription", "Text", $script:dgIntuneManagerObjects.SelectedItem.Description)
    }
    else
    {
        $script:UIProvider.SetXamlProperty($script:copyForm, "txtObjectDescription", "IsEnabled", $false)
    }

    $tokenList = @()
    $tokenList += Get-TokenInfo | ForEach-Object { 
        $name = $_.TenantName
        if($_.IsDefault) {
            $name = $name + " (Current)"
        }
        $_ | Add-Member -NotePropertyName "TenantNameEx" -NotePropertyValue $name | Out-Null
        $_
    }
    $script:UIProvider.SetXamlProperty($script:copyForm, "cbDestinationTenant", "ItemsSource", $tokenList)
    $script:UIProvider.SetXamlProperty($script:copyForm, "cbDestinationTenant", "SelectedItem", ($tokenList | Where-Object IsDefault -eq $true))

    # ---- Scope tags (dual-list) ----------------------------------------------
    # Same pattern as the Details view, with two cross-cutting concerns:
    # (1) destination tenant determines which tag list to show — scope tag IDs
    #     are per-tenant, so when the user picks a different tenant the list
    #     reloads from THAT tenant's dependency cache.
    # (2) Assigned-list seed crosses tenants by NAME (source tag names that
    #     exist in the destination get pre-checked; the rest drop with a log).
    $script:_copyScopeTagPropName = $null
    $sourceItem = $script:dgIntuneManagerObjects.SelectedItem
    if($sourceItem.PolicyType.ScopeTagProperty) {
        $script:_copyScopeTagPropName = [string]$sourceItem.PolicyType.ScopeTagProperty
    }

    if(-not $script:_copyScopeTagPropName) {
        $script:UIProvider.SetXamlProperty($script:copyForm, "pnlCopyScopeTagsLabel", "Visibility", "Collapsed")
        $script:UIProvider.SetXamlProperty($script:copyForm, "grdCopyScopeTags",      "Visibility", "Collapsed")
    }
    else {
        # Capture the source tags by NAME so we can re-seed the Assigned list
        # against the destination tenant whenever the tenant combo changes.
        $script:_copySourceScopeTagNames = @()
        $sourceTagIds = @()
        if($sourceItem.Object -and $sourceItem.Object.PSObject.Properties[$script:_copyScopeTagPropName]) {
            $sourceTagIds = @($sourceItem.Object.$script:_copyScopeTagPropName | ForEach-Object { [string]$_ })
        }
        try {
            $srcDeps = Get-GraphDependencySourceObjects $sourceItem -DefaultPoliciesOnly
            if($srcDeps -and $srcDeps.ContainsKey("ScopeTags")) {
                foreach($id in $sourceTagIds) {
                    $tag = @($srcDeps["ScopeTags"]) | Where-Object { [string]$_.Id -eq $id } | Select-Object -First 1
                    if($tag) { $script:_copySourceScopeTagNames += [string]$tag.Name }
                }
            }
        }
        catch { Write-LogDebug "Source scope-tag name capture failed: $($_.Exception.Message)" }

        $script:colCopyAvailableScopeTags = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
        $script:colCopyAssignedScopeTags  = [System.Collections.ObjectModel.ObservableCollection[object]]::new()

        $script:UIProvider.SetXamlProperty($script:copyForm, "lstCopyAvailableScopeTags", "ItemsSource", $script:colCopyAvailableScopeTags)
        $script:UIProvider.SetXamlProperty($script:copyForm, "lstCopyAssignedScopeTags",  "ItemsSource", $script:colCopyAssignedScopeTags)

        # Single source of truth (same pattern as Details view): full list +
        # assigned-ids set; ObservableCollections derived from both. Mutual
        # exclusion is structural — no chance of a tag drifting between lists.
        $script:_copyAllScopeTags  = @()
        $script:_copyAssignedIds   = [System.Collections.Generic.HashSet[string]]::new()

        $script:_syncCopyScopeTagLists = {
            $script:colCopyAvailableScopeTags.Clear()
            $script:colCopyAssignedScopeTags.Clear()
            foreach($tag in $script:_copyAllScopeTags) {
                if($script:_copyAssignedIds.Contains($tag.Id)) {
                    [void]$script:colCopyAssignedScopeTags.Add($tag)
                } else {
                    [void]$script:colCopyAvailableScopeTags.Add($tag)
                }
            }
        }

        # Reload both lists for a given destination tenant. Cache lookup first;
        # live fetch if cold. Re-seeds the assigned-ids set by name lookup so
        # cross-tenant copies pre-pick the same-named tags in the destination.
        # Dedupes by Id (the cache can hold both the synthetic Default and the
        # real Default scope tag).
        $script:_reloadCopyScopeTags = {
            param([int]$DestTokenId)

            $destTags = @()
            try {
                $destTokenInfo = Get-TokenInfo $DestTokenId
                if($destTokenInfo) {
                    $cacheId = "DependencyObjects_$($destTokenInfo.TenantId)"
                    $deps = Get-CacheObject $cacheId
                    if($deps -and $deps.ContainsKey("ScopeTags")) {
                        $destTags = @($deps["ScopeTags"])
                    }
                }
            }
            catch { Write-LogDebug "Copy scope tags: cache lookup failed: $($_.Exception.Message)" }

            if($destTags.Count -eq 0) {
                # Cold cache — fetch live for this tenant. ScopeTags is small; cheap.
                try {
                    $destTags = @(Get-GraphPolicies -PolicyType "ScopeTags" -TokenId $DestTokenId)
                    # Mirror what Initialize-TenantDependencyCache does: prepend the
                    # synthetic Default so it's available for assignment.
                    $destTags = @([PSCustomObject]@{ ID = 0; Name = "Default" }) + $destTags
                }
                catch { Write-LogError "Failed to load scope tags for destination tenant" $_.Exception }
            }

            # Project + dedupe by Id.
            $dedup = @()
            $seenIds = [System.Collections.Generic.HashSet[string]]::new()
            foreach($tag in $destTags) {
                $idStr = [string]$tag.Id
                if(-not $seenIds.Add($idStr)) { continue }
                $dedup += [PSCustomObject]@{ Id = $idStr; Name = [string]$tag.Name }
            }
            $script:_copyAllScopeTags = @($dedup | Sort-Object Name)

            # Re-seed assigned-ids from the captured source tag NAMES (cross-tenant
            # safe). Names that don't exist in the destination simply don't get
            # selected — the user can pick something else.
            $script:_copyAssignedIds.Clear()
            $seedNames = [System.Collections.Generic.HashSet[string]]::new()
            foreach($n in $script:_copySourceScopeTagNames) { [void]$seedNames.Add($n) }
            foreach($tag in $script:_copyAllScopeTags) {
                if($seedNames.Contains($tag.Name)) { [void]$script:_copyAssignedIds.Add($tag.Id) }
            }

            & $script:_syncCopyScopeTagLists
        }

        # Initial load against the default-selected destination tenant.
        $initialDest = $script:UIProvider.GetXamlProperty($script:copyForm, "cbDestinationTenant", "SelectedItem")
        if($initialDest) { & $script:_reloadCopyScopeTags $initialDest.Id }

        $script:UIProvider.AddXamlEvent($script:copyForm, "cbDestinationTenant", "Add_SelectionChanged", {
            $newDest = $script:UIProvider.GetXamlProperty($script:copyForm, "cbDestinationTenant", "SelectedItem")
            if($newDest) { & $script:_reloadCopyScopeTags $newDest.Id }
        })

        # Move = mutate the canonical assigned-ids set, then re-derive both lists.
        $script:_moveCopyScopeTags = {
            param([string]$Action, [string[]]$Ids)
            if($Ids.Count -eq 0) { return }
            if($Action -eq "assign") {
                foreach($id in $Ids) { [void]$script:_copyAssignedIds.Add([string]$id) }
            } else {
                foreach($id in $Ids) { [void]$script:_copyAssignedIds.Remove([string]$id) }
            }
            & $script:_syncCopyScopeTagLists
        }

        $script:UIProvider.AddXamlEvent($script:copyForm, "btnCopyScopeTagAssign", "Add_Click", {
            try {
                $lb = $script:copyForm.FindName("lstCopyAvailableScopeTags")
                if(-not $lb -or $lb.SelectedItems.Count -eq 0) { return }
                $ids = @($lb.SelectedItems | ForEach-Object { [string]$_.Id })
                & $script:_moveCopyScopeTags "assign" $ids
            }
            catch { Write-LogError "Copy scope-tag assign handler failed" $_.Exception }
        })

        $script:UIProvider.AddXamlEvent($script:copyForm, "btnCopyScopeTagUnassign", "Add_Click", {
            try {
                $lb = $script:copyForm.FindName("lstCopyAssignedScopeTags")
                if(-not $lb -or $lb.SelectedItems.Count -eq 0) { return }
                $ids = @($lb.SelectedItems | ForEach-Object { [string]$_.Id })
                & $script:_moveCopyScopeTags "unassign" $ids
            }
            catch { Write-LogError "Copy scope-tag unassign handler failed" $_.Exception }
        })

        $script:UIProvider.AddXamlEvent($script:copyForm, "lstCopyAvailableScopeTags", "Add_MouseDoubleClick", {
            try {
                $lb = $script:copyForm.FindName("lstCopyAvailableScopeTags")
                if(-not $lb -or $lb.SelectedItems.Count -eq 0) { return }
                $ids = @($lb.SelectedItems | ForEach-Object { [string]$_.Id })
                & $script:_moveCopyScopeTags "assign" $ids
            }
            catch { Write-LogError "Copy scope-tag double-click (available) failed" $_.Exception }
        })

        $script:UIProvider.AddXamlEvent($script:copyForm, "lstCopyAssignedScopeTags", "Add_MouseDoubleClick", {
            try {
                $lb = $script:copyForm.FindName("lstCopyAssignedScopeTags")
                if(-not $lb -or $lb.SelectedItems.Count -eq 0) { return }
                $ids = @($lb.SelectedItems | ForEach-Object { [string]$_.Id })
                & $script:_moveCopyScopeTags "unassign" $ids
            }
            catch { Write-LogError "Copy scope-tag double-click (assigned) failed" $_.Exception }
        })
    }

    $script:copyForm.Add_ContentRendered({
        $txtName = $script:copyForm.FindName("txtObjectName")
        if($txtName)
        {
            $txtName.SelectAll()
        }
    })

    $script:UIProvider.AddXamlEvent($script:copyForm, "btnOk", "Add_Click", {
        $script:copyForm.DialogResult = $true
    })

    $script:copyForm.Owner = $script:window
    $script:copyForm.Icon = $script:Window.Icon     
    $ret = $script:copyForm.ShowDialog()

    if($ret)
    {
        $txtDescrption = $null
        $newName = $script:UIProvider.GetXamlProperty($script:copyForm, "txtObjectName", "Text")
        if(-not $newName)
        {
            Write-Log "New name cannot be empty. Copy object skipped" 2
            Write-Status ""
            return
        }

        if(($script:UIProvider.GetXamlProperty($script:copyForm, "txtObjectDescription", "IsEnabled")) -eq $true)
        {
            $txtDescrption = $script:UIProvider.GetXamlProperty($script:copyForm, "txtObjectDescription", "Text")
        }

        # Export profile
        Write-Status "Copy $($script:dgIntuneManagerObjects.SelectedItem.Name)"

        #$tokenId = $script:dgIntuneManagerObjects.SelectedItem.TokenId
        $selectedTenant = $script:UIProvider.GetXamlProperty($script:copyForm, "cbDestinationTenant", "SelectedItem")
        $tokenId = $selectedTenant.Id

        $copyArgs = @{
            InputObject = $script:dgIntuneManagerObjects.SelectedItem
            Name        = $newName
            Description = $txtDescrption
            TokenId     = $tokenId
        }
        # Forward the user's scope-tag picks (when the section was shown). Pass
        # even an empty array so the new copy explicitly reflects user intent —
        # CopyObject will override the cloned JSON only when this is non-null.
        # Read from the canonical $script:_copyAssignedIds set (single source of
        # truth) rather than the UI-derived ObservableCollection.
        if($script:_copyScopeTagPropName -and $null -ne $script:_copyAssignedIds) {
            $copyArgs['ScopeTagIds'] = @($script:_copyAssignedIds)
        }

        $newPolicies = Copy-GraphPolicy @copyArgs

        if($newPolicies -and (Get-SettingValue "RefreshObjectsAfterCopy") -eq $true) {
            Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject
        }

        #if($script:dgIntuneManagerObjects.SelectedItem.CopyObject($newName, $txtDescrption, 0)) {
        #    if((Get-SettingValue "RefreshObjectsAfterCopy") -eq $true)
        #    {
        #        Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject
        #    }
        #}

        Write-Status ""    
    }
    $script:dgIntuneManagerObjects.Focus()
}

#region Events

