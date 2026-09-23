# Internal helpers for Start-GraphBulkExport (Public/Start-GraphBulkExport.ps1).
# Extracted here per architecture rule R11 (keep public driver files thin); these
# are module-internal and never exported. All dot-source into the same module
# scope, so the public driver calls them by name.

# NOTE: the five legacy Sync-BulkExport* policy-hydration helpers (reusable
# settings, branding images, app-config target apps, role-assignment details,
# Terms-of-Use files) were retired. Each now lives on its owning policy class via
# the GetSubResourceBatchRequests / ApplySubResourceBatchResult / FinalizeSubResources
# contract, driven by Internal/PolicySubresourceFetch.ps1 (single + batch + parallel
# in one place). Internal/PolicyHydrateExtras.ps1 was deleted.

# Flag every policy whose export file would land on the same path as another so the
# per-policy GetFileName appends `_<id>` and the duplicate doesn't silently
# overwrite its twin on disk.
#
# Filename uniqueness was previously left to the user-facing AddIDToExportFile
# setting (off by default). Real tenants — one lab tenant has 8 CA policies all named
# "Temp Policy 1" and ~30 SettingsCatalog clones — relied on that setting being on
# to avoid data loss on bulk export. The collision detector kicks in automatically
# so users don't need to discover the setting after losing files.
#
# Two things decide whether two policies collide, and BOTH have to be part of the
# bucket key:
#
#   * The FILE NAME, not the displayName. The two differ whenever sanitization
#     strips a character, and it is the file name that collides on disk. "Wi-Fi:
#     Corp" and "Wi-Fi Corp" are distinct displayNames that both sanitize to
#     "Wi-Fi Corp.json". This always mattered on Windows (':' has always been an
#     invalid file name char) and PortableFileNames extends the same class of
#     collision to Linux and macOS.
#
#   * The FOLDER, which is why this runs once across every type rather than once
#     per type. With AddObjectType off, every type writes straight into the export
#     root, so a per-type pass compared each type only against itself and never
#     noticed a Compliance policy and a Configuration policy of the same name
#     landing on one path. Any two types sharing a .Folder collide the same way
#     even with AddObjectType on.
#
# The name comes from the policy's OWN GetFileName rather than a re-derivation of
# the base transform: subclasses override it (DeviceEnrollmentObject names default
# policies from their id suffix), and a detector that re-implements the base rule
# buckets those by a name they never write under.
function Set-BulkExportFilenameCollisionFlag {
    param(
        [Parameter(Mandatory)]$Policies,
        # Mirrors ExportSettings.AddObjectType: on = a per-type subfolder, off =
        # everything in one directory. Decides whether the type's folder is part of
        # the destination path at all.
        [switch]$AddObjectType
    )

    if (-not $Policies -or @($Policies).Count -lt 2) { return }

    # Every name is read before any flag is set. GetFileName's output depends on
    # _NeedsIdInFilename, so mutating as we go would make later policies in the same
    # bucket hash to their already-disambiguated name and escape detection.
    $byPath = @{}
    foreach ($p in $Policies) {
        if (-not $p) { continue }

        $name = $null
        try { $name = $p.GetFileName($null) }
        catch { Write-LogDebug "Bulk export: could not resolve a file name for '$($p.Name)' - excluded from collision detection: $($_.Exception.Message)" }
        if (-not $name) { continue }

        $folder = if ($AddObjectType) { [string]$p.PolicyType.Folder } else { '' }

        # ToLowerInvariant because the collision is on the file system, and Windows
        # and macOS are case-insensitive: "Baseline.json" and "baseline.json" are one
        # file there. Over-flagging on Linux is harmless - the id suffix is only ever
        # added, never dropped.
        $key = "$folder/$name".ToLowerInvariant()

        if (-not $byPath.ContainsKey($key)) {
            $byPath[$key] = [System.Collections.Generic.List[object]]::new()
        }
        [void]$byPath[$key].Add($p)
    }

    $collided = 0
    foreach ($key in $byPath.Keys) {
        $bucket = $byPath[$key]
        if ($bucket.Count -lt 2) { continue }
        foreach ($p in $bucket) {
            if ($p.PSObject.Properties['_NeedsIdInFilename']) {
                $p._NeedsIdInFilename = $true
                $collided++
            }
        }
    }

    if ($collided -gt 0) {
        Write-Log "Bulk export: $collided policy file(s) resolved to duplicate file names - appending '_<id>' to their filenames to avoid silent overwrite"
    }
}

# ExportFullMembershipPrefixes is gone: all group paths are batched now, so its
# performance rationale went with them, and its only observable output was the
# #DirectMembers / #DirectMemberCount stamps on group sidecars, which nothing in
# the module ever read. An automation run downstream could have read them though,
# so a leftover value must not be dropped in silence - the export would "succeed"
# with the member lists simply absent from the sidecars.
#
# Two places can still hold one, and they need different treatment:
#
#   * A settings file. Read-only from here - rewriting the caller's automation
#     input is not ours to do, so this warns on every run until they edit it.
#
#   * The settings store, where every export used to persist the key alongside
#     AddCompanyName and ExportAssignments. Nothing reads it now, so it is
#     cleared after being reported: otherwise the warning would be permanent for
#     a value the user has no UI left to clear.
#
# The store holds more than one candidate. The old read went through
# Get-SettingValue, which resolves "<tenantId>\IntuneManager" BEFORE the global
# "IntuneManager" path, so a tenant override was the value that actually applied
# and cannot be skipped here. Every applicable path is walked in that same
# precedence order.
#
# A stored empty string reads back the same as "not stored" (Get-SettingStoreValue
# treats "" as unset), so only a non-empty leftover is reportable - which is also
# the only one that ever changed any output. The removal is unconditional
# regardless, so the dead key does not sit in every user's store for ever.
function Clear-BulkExportLegacyMembershipSetting
{
    param(
        $FileSettings,
        [string] $SettingsFile,
        [int] $TokenId = 0
    )

    $removedKey = 'ExportFullMembershipPrefixes'
    $sidecarNote = 'Group sidecars no longer carry the #DirectMembers / #DirectMemberCount member lists.'

    if ($FileSettings -and $FileSettings.PSObject.Properties[$removedKey]) {
        $fileValue = ([string]$FileSettings.$removedKey).Trim()
        if ($fileValue) {
            Write-Log "Bulk export: settings file '$SettingsFile' sets $removedKey ('$fileValue'), which has been removed - the value is ignored. $sidecarNote Remove the property from the settings file to silence this warning." 2
        }
        else {
            # Every file written by an older build carries the key. One that was
            # never set changes nothing, so it is not worth a warning.
            Write-LogDebug "Bulk export: settings file '$SettingsFile' still carries the removed $removedKey property with an empty value - ignored, no change in output."
        }
    }

    # Two tenant ids can be the applicable one: the organization the session is
    # connected to (which is what Get-SettingValue itself keyed on) and, on a
    # cross-tenant automation run, the tenant behind this export's own token.
    $subPaths = [System.Collections.Generic.List[string]]::new()
    $tenantIds = [System.Collections.Generic.List[string]]::new()
    if ($script:OrganizationId) { [void]$tenantIds.Add([string]$script:OrganizationId) }
    $tokenInfo = Get-OperationTokenInfo $TokenId
    if ($tokenInfo -and $tokenInfo.TenantId) { [void]$tenantIds.Add([string]$tokenInfo.TenantId) }

    foreach ($tenantId in $tenantIds) {
        # A path stored flat by an older build ("<tenantId>\IntuneManager" as one
        # property name) resolves from this same string - Get-SettingsTreeNode
        # checks the flat property before splitting the path.
        $tenantPath = "$tenantId\IntuneManager"
        if (-not $subPaths.Contains($tenantPath)) { [void]$subPaths.Add($tenantPath) }
    }
    [void]$subPaths.Add("IntuneManager")   # global last, as it was read last

    foreach ($subPath in $subPaths) {
        $storedValue = Get-SettingStoreValue $subPath $removedKey
        if ($storedValue) {
            Write-Log "Bulk export: the saved setting $subPath\$removedKey ('$storedValue') has been removed - it is ignored and is being cleared from the settings store. $sidecarNote" 2
        }
        # Not inside the if: an empty leftover is worth no warning but is still a
        # dead key. Removal touches the store only when the key is actually there.
        Remove-SettingStoreValue $subPath $removedKey
    }
}

# Pre-warm the AAD-object cache that Add-GraphMigrationObject reads on every
# group reference. Walks three sources visible in the already-fetched policy
# bodies: assignment targets, CA conditions.users.include/excludeGroups, and
# compliance scheduledActionConfigurations.notificationMessageCCList. Without
# this each unique group is fetched one-by-one (sequential Invoke-MSGraphAPI
# calls per policy) during Phase 3, which dominates the wall time for CA-heavy
# or assignment-heavy tenants. Batched here in parallel chunks; subsequent
# per-policy lookups are all cache hits.
function Sync-BulkExportMigrationGroups
{
    param(
        [Parameter(Mandatory)] $Policies,
        [int] $TokenId = 0
    )

    if (-not $Policies -or $Policies.Count -eq 0) { return }

    # Collect unique group ids referenced by any group-style assignment target.
    $groupIds = [System.Collections.Generic.HashSet[string]]::new()
    $groupTargetTypes = @(
        '#microsoft.graph.groupAssignmentTarget'
        '#microsoft.graph.exclusionGroupAssignmentTarget'
        '#microsoft.graph.cloudPcManagementGroupAssignmentTarget'
    )

    # CA condition group ids are special-cased strings, not real group ids.
    $caSkipIds = @('All','None','GuestsOrExternalUsers')

    foreach ($p in $Policies) {
        $body = $p.JsonObject
        if (-not $body) { continue }

        # Assignments live on JsonObject.Assignments after Get-GraphPolicies populates them.
        foreach ($a in @($body.Assignments)) {
            foreach ($target in @($a.target)) {
                if (-not $target -or -not $target.'@odata.type') { continue }
                if ($groupTargetTypes -contains $target.'@odata.type' -and $target.groupId) {
                    [void]$groupIds.Add($target.groupId)
                }
            }
        }

        # CA conditions reference groups by id directly in the policy body. These
        # are resolved by IntuneConditionalAccessClasses.ps1 PostExportCommand —
        # one Graph GET per id during the write phase if not pre-cached here.
        $caUsers = $body.conditions.users
        if ($caUsers) {
            foreach ($id in @($caUsers.includeGroups) + @($caUsers.excludeGroups)) {
                if (-not $id) { continue }
                if ($caSkipIds -contains $id) { continue }
                [void]$groupIds.Add($id)
            }
        }

        # Compliance notification CC groups — resolved by IntuneComplianceClasses
        # PostExportCommand. The expand on scheduledActionConfigurations is on by
        # default for compliance policies, so the ids are already in the body.
        foreach ($rule in @($body.scheduledActionsForRule)) {
            foreach ($cfg in @($rule.scheduledActionConfigurations)) {
                foreach ($id in @($cfg.notificationMessageCCList)) {
                    if ($id) { [void]$groupIds.Add($id) }
                }
            }
        }
    }
    if ($groupIds.Count -eq 0) { return }

    # One token, never the whole list - see the prefetch above.
    $tokenInfo = Get-OperationTokenInfo $TokenId
    if (-not $tokenInfo -or -not $tokenInfo.TenantId) {
        Write-Log "Bulk export: skipping migration-group prefetch - no token info for TokenId $TokenId" 2
        return
    }
    $cacheKey  = "AADObjectCache_$($tokenInfo.TenantId)"
    $cacheFile = "TenantCache_$($tokenInfo.TenantId)"
    $cache     = Get-CacheObject $cacheKey @{}
    if ($cache -isnot [Hashtable]) { $cache = @{} }

    # Skip ids already cached (positive OR negative — null entries mean "known missing").
    $needFetch = [System.Collections.Generic.List[string]]::new()
    foreach ($gid in $groupIds) {
        if (-not $cache.ContainsKey($gid)) { [void]$needFetch.Add($gid) }
    }

    if ($needFetch.Count -gt 0) {
        Write-Log "Bulk export: pre-fetching $($needFetch.Count) AAD group(s) referenced by policies (assignments, CA conditions, compliance notifications)"
        Write-Status -Detail ("Pre-fetching {0} group(s)" -f $needFetch.Count) -SkipLog -Force

        # Full bodies, a thousand ids per POST, through /directoryObjects/getByIds -
        # the same resolver the documentation preload uses. This used to be one
        # GET groups/<id> sub-request per group inside $batch (fifty POSTs for a
        # thousand groups); now it is one or two POSTs. Verified live 2026-09-07:
        # the body is identical to GET groups/<id> in every property and value
        # except an extra @odata.type, which is stripped below so the Groups/
        # sidecars stay byte-identical to before. A group the directory does not
        # return is simply absent from the response: deleted, so it is
        # negative-cached exactly as a 404 was.
        $resolved = 0
        $missing  = 0
        for ($i = 0; $i -lt $needFetch.Count; $i += 1000) {
            $end   = [Math]::Min($i + 999, $needFetch.Count - 1)
            $chunk = @($needFetch[$i..$end])

            $resp = $null
            try {
                $body = (@{ ids = $chunk; types = @('group') } | ConvertTo-Json -Compress)
                $resp = Invoke-MSGraphAPI -Url 'directoryObjects/getByIds' -Content $body -HttpMethod POST -TokenId $TokenId -ODataMetadata 'none'
            }
            catch {
                Write-LogError "Bulk export: getByIds group prefetch failed for a chunk of $($chunk.Count)" $_.Exception
            }
            if ($null -eq $resp) {
                # Transport or auth failure: leave the chunk UNcached so the per-policy
                # path can still resolve those ids later, rather than negative-caching
                # a thousand live groups on one bad call.
                continue
            }

            $seen = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($g in @($resp.value)) {
                if (-not $g -or -not $g.id) { continue }
                if ($g.PSObject.Properties['@odata.type']) { [void]$g.PSObject.Properties.Remove('@odata.type') }
                $cache[[string]$g.id] = $g
                [void]$seen.Add([string]$g.id)
                $resolved++
            }
            foreach ($gid in $chunk) {
                if (-not $seen.Contains($gid)) {
                    $cache[$gid] = $null
                    $missing++
                }
            }
        }
        Set-CacheObject $cacheKey $cache $cacheFile
        Write-Log "Bulk export: migration-group prefetch complete - $resolved resolved, $missing missing/deleted"
    }
}

# Pre-walk the nested-group hierarchy in batches so the per-policy recursion
# inside Add-GraphMigrationObject doesn't fire one un-batched GET per parent
# group plus one per child body. BFS, level-by-level: at each level we batch
# /members/microsoft.graph.group for every group at that level, attach the
# resulting children list to the parent body as `#NestedChildren`, then batch
# /groups/{id} for any newly-discovered children that aren't yet in the
# AADObjectCache. Stops when there are no more groups to walk or when we hit
# MaxDepth-1 (depth 1 means "just the directly-assigned group", so no walk).
#
# The dedup inside Add-GraphMigrationObject's recursion still handles cycles
# correctly — this function is purely a Graph-call optimizer; it doesn't
# change which groups end up in the export.
function Sync-BulkExportNestedGroupHierarchy
{
    param(
        [Parameter(Mandatory)] [int] $MaxDepth,
        [int] $TokenId = 0
    )

    if ($MaxDepth -le 1) { return }

    # One token, never the whole list - see the prefetch above.
    $tokenInfo = Get-OperationTokenInfo $TokenId
    if (-not $tokenInfo -or -not $tokenInfo.TenantId) {
        Write-Log "Bulk export: skipping nested-group hierarchy prefetch - no token info for TokenId $TokenId" 2
        return
    }

    $cacheKey  = "AADObjectCache_$($tokenInfo.TenantId)"
    $cacheFile = "TenantCache_$($tokenInfo.TenantId)"
    $cache     = Get-CacheObject $cacheKey @{}
    if ($cache -isnot [Hashtable]) { $cache = @{} }

    # Seed the walk with every group already in cache (i.e. every group the
    # earlier Sync-BulkExportMigrationGroups resolved). These are the level-1
    # starting points for the hierarchy.
    $currentLevelIds = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $cache.GetEnumerator()) {
        if ($null -eq $entry.Value) { continue }                                          # negative-cached
        if ($entry.Value.PSObject.Properties.Name -contains '#NestedChildren') { continue } # already walked this session
        [void]$currentLevelIds.Add([string]$entry.Key)
    }

    if ($currentLevelIds.Count -eq 0) { return }

    $totalParents  = 0
    $totalChildren = 0

    # MaxDepth=2 means "walk one level": stamp #NestedChildren on every level-1
    # group. MaxDepth=3 → walk two levels. Loop bound is MaxDepth-1.
    for ($depth = 1; $depth -lt $MaxDepth -and $currentLevelIds.Count -gt 0; $depth++) {

        Write-Log "Bulk export: nested-group hierarchy - fetching children for $($currentLevelIds.Count) group(s) at depth $depth"
        Write-Status -Detail ("Pre-fetching nested groups (depth {0}, {1} parent(s))" -f $depth, $currentLevelIds.Count) -SkipLog -Force

        # Batch-fetch /members/microsoft.graph.group for every parent at this level.
        $memBatch  = [System.Collections.Generic.List[PSCustomObject]]::new()
        $idToParent = @{}
        $i = 0
        foreach ($gid in $currentLevelIds) {
            $i++
            $reqId = "nest_${depth}_$i"
            [void]$memBatch.Add([PSCustomObject]@{
                id      = $reqId
                method  = 'GET'
                url     = "groups/$gid/members/microsoft.graph.group?`$select=id,displayName"
                headers = @{ Accept = 'application/json;odata.metadata=minimal' }
            })
            $idToParent[$reqId] = $gid
        }

        $memResults = Invoke-GraphBatchRequest -BatchObjects $memBatch -BatchType "Nested group members (depth $depth)" -TokenId $TokenId -AllPages -SkipWarnings -IncludedFailed

        # Walk results: stamp #NestedChildren on parents, collect new child ids.
        $newChildIds = [System.Collections.Generic.HashSet[string]]::new()
        $childListsByParent = @{}
        foreach ($r in $memResults) {
            $parentId = $idToParent["$($r.Id)"]
            if (-not $parentId) { continue }

            $children = [System.Collections.Generic.List[PSCustomObject]]::new()
            if ($r.body -and $r.Status -ge 200 -and $r.Status -lt 300) {
                foreach ($c in @($r.body.value)) {
                    if (-not $c -or -not $c.id) { continue }
                    [void]$children.Add([PSCustomObject]@{ id = $c.id; displayName = $c.displayName })
                    if (-not $cache.ContainsKey($c.id)) { [void]$newChildIds.Add($c.id) }
                }
            }
            else {
                Write-LogDebug "Bulk export: nested-group members fetch for $parentId returned HTTP $($r.Status)"
            }

            $parentBody = $cache[$parentId]
            if ($null -ne $parentBody) {
                Add-Member -InputObject $parentBody -MemberType NoteProperty -Name '#NestedChildren' -Value $children.ToArray() -Force
            }
            $childListsByParent[$parentId] = $children
            $totalParents++
        }

        # Batch-fetch bodies for newly-discovered children so the recursion
        # inside Add-GraphMigrationObject hits cache for them too.
        if ($newChildIds.Count -gt 0) {
            Write-Status -Detail ("Pre-fetching {0} nested-group bod(ies) (depth {1})" -f $newChildIds.Count, $depth) -SkipLog -Force

            $bodyBatch  = [System.Collections.Generic.List[PSCustomObject]]::new()
            $idToChild  = @{}
            $j = 0
            foreach ($gid in $newChildIds) {
                $j++
                $reqId = "nestbody_${depth}_$j"
                [void]$bodyBatch.Add([PSCustomObject]@{
                    id      = $reqId
                    method  = 'GET'
                    url     = "groups/$gid"
                    headers = @{ Accept = 'application/json;odata.metadata=none' }
                })
                $idToChild[$reqId] = $gid
            }

            $bodyResults = Invoke-GraphBatchRequest -BatchObjects $bodyBatch -BatchType "Nested group bodies (depth $depth)" -TokenId $TokenId -SkipWarnings -IncludedFailed

            foreach ($r in $bodyResults) {
                $gid = $idToChild["$($r.Id)"]
                if (-not $gid) { continue }
                if ($r.body -and $r.Status -ge 200 -and $r.Status -lt 300) {
                    $cache[$gid] = $r.body
                }
                else {
                    $cache[$gid] = $null  # negative-cache so we don't re-query during the write phase
                }
            }
            $totalChildren += $newChildIds.Count
        }

        # Next iteration walks the children we just found that weren't already in cache
        # and don't yet have their own #NestedChildren stamp.
        $nextLevel = [System.Collections.Generic.List[string]]::new()
        foreach ($pair in $childListsByParent.GetEnumerator()) {
            foreach ($child in $pair.Value) {
                $body = $cache[$child.id]
                if ($null -eq $body) { continue }
                if ($body.PSObject.Properties.Name -contains '#NestedChildren') { continue }
                [void]$nextLevel.Add($child.id)
            }
        }
        $currentLevelIds = $nextLevel
    }

    Set-CacheObject $cacheKey $cache $cacheFile
    Write-Log "Bulk export: nested-group hierarchy prefetch complete - $totalParents parent(s) walked, $totalChildren new child group bod(ies) cached"
}
