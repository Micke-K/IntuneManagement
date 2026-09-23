# Intune cross-tenant migration-table feature (export/import of Groups, ScopeTags
# etc. between environments). Moved out of Internal/MSGraph.ps1 (architecture R4 -
# MSGraph.ps1 should hold only generic Graph helpers, not Intune feature logic).
# All module-internal; shares the $script:_migFile* caches set up by Start-GraphBulkExport.

#region Migration Table functions
function Get-MigrationTableInfo
{
    param($Path, $TenantId)

    $FileName = Get-GraphMigrationTableFromPath $Path

    $str = $null
    $sameTenant = $false
    if($FileName -and [IO.File]::Exists($FileName))
    {
        $migFileObj = ConvertFrom-Json ([IO.File]::ReadAllText($FileName))
        if($migFileObj.TenantId -and $migFileObj.TenantId -eq $TenantId) 
        { 
            $sameTenant = $true
            $str = "Current tenant. Migration table will not be used"
        }
        elseif($migFileObj.Organization)
        {
            $str = "Objects exported from $($migFileObj.Organization) ($($migFileObj.TenantId))"
        }
    }

    if(-not $str)
    {
        # Hide controls?
        $str = "No migration table found"
    }
    $str, $sameTenant
}

function Add-GraphMigrationInfo
{
    # MigrationRoot (optional): the export root directory where MigrationTable.json
    # and Groups/ should live. When supplied, takes precedence over the legacy
    # ".Parent.FullName" guess in Add-GraphMigrationObject — needed because that
    # guess assumed $Folder is always a per-type subfolder one level under the
    # export root, which is wrong when AddObjectType=false (policies written
    # straight to the export root). Without this, MigrationTable.json + Groups/
    # ended up one level too high and the "Parent name of the folder is not
    # organization name" warning fired falsely.
    #
    # MaxGroupDepth (optional): depth of group-membership recursion. 1 (default)
    # exports only the directly-assigned group; >1 walks into the group's member
    # groups up to MaxGroupDepth levels deep. Cycle-safe via the per-folder dedup
    # index in Add-GraphMigrationObject.
    param($PolicyObject, $Folder, [string]$MigrationRoot, [int]$MaxGroupDepth = 1)

    if(-not $PolicyObject) { return }

    foreach($assignment in $PolicyObject.JsonObject.Assignments)
    {
        foreach($assignmentTarget in $assignment.target)
        {
            if(-not $assignmentTarget."@odata.type") { continue }

            $assignmentTargetType = $assignmentTarget."@odata.type"

            if($assignmentTargetType -eq "#microsoft.graph.groupAssignmentTarget" -or
                $assignmentTargetType -eq "#microsoft.graph.exclusionGroupAssignmentTarget" -or
                $assignmentTargetType -eq "#microsoft.graph.cloudPcManagementGroupAssignmentTarget")
            {
                Add-GraphMigrationObject $assignmentTarget.groupid "groups" "Group" $Folder $PolicyObject._TokenId -MigrationRoot $MigrationRoot -MaxGroupDepth $MaxGroupDepth
            }
            elseif($assignmentTargetType -eq "#microsoft.graph.allLicensedUsersAssignmentTarget" -or
                $assignmentTargetType -eq "#microsoft.graph.allDevicesAssignmentTarget")
            {
                # No need to migrate All Users or All Devices
            }
            else
            {
                Write-Log "Unsupported migration object: $assignmentTargetType" 3
            }

            # Assignment filters ride on EVERY target type as a sibling property (incl.
            # All Users / All Devices). Capture the filter to the migration table + an
            # AssignmentFilters sidecar so cross-tenant import can resolve/create it.
            $filterId = [string]$assignmentTarget.deviceAndAppManagementAssignmentFilterId
            if($filterId -and $filterId -ne "00000000-0000-0000-0000-000000000000")
            {
                Add-GraphMigrationObject $filterId "deviceManagement/assignmentFilters" "AssignmentFilter" $Folder $PolicyObject._TokenId -MigrationRoot $MigrationRoot
            }
        }
    }
}

function Add-GraphMigrationObject
{
    # MigrationRoot (optional): explicit export root for MigrationTable.json and
    # Groups/. Use it when known (Export-GraphPolicy passes $exportFolderRoot).
    # Otherwise the fallback is the legacy "$Folder.Parent.FullName" heuristic,
    # which is correct when $Folder is a per-type subfolder under the export root
    # but wrong when $Folder IS the export root (AddObjectType=false case).
    #
    # MaxGroupDepth / CurrentGroupDepth: drive nested-group export. When saving a
    # Group sidecar, if CurrentGroupDepth < MaxGroupDepth we fetch the group's
    # members and recurse for any that are themselves groups (CurrentGroupDepth+1).
    # The per-folder dedup index naturally breaks cycles (a group already added
    # to MigrationTable.json won't be re-added or re-recursed).
    param($ObjectId, $GraphAPI, $ObjectTypeName, $Folder, $TokenId = 0, [string]$MigrationRoot,
          [int]$MaxGroupDepth = 1, [int]$CurrentGroupDepth = 1)

    if(-not $ObjectId) { return }

    # Get-OperationTokenInfo, not Get-TokenInfo: a caller with no token of its own
    # passes 0, the codebase's spelling for "the default token", and the list
    # accessor answers 0 with EVERY registered token. With two tenants signed in
    # $tokeInfo.TenantId was an ARRAY, so the cache key below became
    # "AADObjectCache_tenant-a tenant-b" - one partition holding two directories,
    # which then leaked one tenant's objects into the other's migration table. A
    # token with no TenantId is the same defect spelled differently (every tenant
    # sharing "AADObjectCache_"), so it stops here too.
    $tokeInfo = Get-OperationTokenInfo ([int]$TokenId)
    if(-not $tokeInfo -or -not $tokeInfo.TenantId) { return }

    # Pin this operation to the resolved token. 0 is a moving target: it is stored in
    # the pending-fetch queue, and Resolve-GraphMigrationObjectsPending drains that
    # queue much later - after a Login or a tenant switch has possibly made a
    # DIFFERENT token the default. The queued id, the nested /members GET and the
    # recursion into child groups all have to name the tenant this export is of, not
    # whichever tenant is default when the flush happens.
    if([int]$tokeInfo.Id -gt 0) { $TokenId = [int]$tokeInfo.Id }
    $cacheTenantId = [string]$tokeInfo.TenantId

    # ----- Per-process MigrationTable.json cache + dedup index -----
    # This function used to read MigrationTable.json from disk, parse it, append
    # via `$Objects += @{...}` (O(n²) array rebuild), then write the entire file
    # back to disk — EVERY single call. For a bulk export of ~400 policies with a
    # few group assignments each that's ~1200 round-trips through disk + JSON
    # serialize. With the cache below, each call is an in-memory hashtable lookup +
    # List<object>.Add; Save-GraphMigrationFilesPending writes each touched file to disk
    # once at end of bulk export.
    if(-not $script:_migFileCache)         { $script:_migFileCache         = @{} }
    if(-not $script:_migFileObjectsIndex)  { $script:_migFileObjectsIndex  = @{} }
    if(-not $script:_migFileDirty)         { $script:_migFileDirty         = @{} }
    # Per-$Folder cache of the resolved MigrationTable.json path. Without this,
    # Get-GraphMigrationTableFromPath fires on every call (~1200× during a bulk
    # export), each call does Expand-FileName + 2 [IO.File]::Exists + when the
    # file doesn't exist yet it Write-Log's "Could not find migration table" —
    # which through Write-Log's ObservableCollection.Add triggers UI marshalling.
    # Resolve once per folder; on cache hit later we skip the lookup entirely.
    if(-not $script:_migFilePathCache)     { $script:_migFilePathCache     = @{} }

    # Key the cache on $Folder (stable across all calls in this bulk export) rather
    # than on the resolved migration-file path (which is $null on first call when the
    # file doesn't exist yet, then changes — leading to a cache miss on every call).
    $cacheKey = $Folder

    if($script:_migFileCache.ContainsKey($cacheKey)) {
        # Fast path: cache already populated for this folder, skip disk probe + log spam.
        $migFileObj    = $script:_migFileCache[$cacheKey]
        $migrationFile = $script:_migFilePathCache[$cacheKey]   # may be $null until first add resolves it
    }
    else {
        # First call for this folder — resolve the path once and seed the caches.
        $migrationFile = Get-GraphMigrationTableFromPath $Folder
        $script:_migFilePathCache[$cacheKey] = $migrationFile

        if($migrationFile -and [IO.File]::Exists($migrationFile)) {
            $migFileObj = ConvertFrom-Json ([IO.File]::ReadAllText($migrationFile))
            # Normalise Objects to a List so adds are O(1) instead of O(n) array rebuilds.
            $list = [System.Collections.Generic.List[object]]::new()
            if($migFileObj.Objects) { foreach($o in $migFileObj.Objects) { [void]$list.Add($o) } }
            $migFileObj.Objects = $list
            $script:_migFileCache[$cacheKey] = $migFileObj

            # Build lookup index for O(1) duplicate-check instead of O(n) Where-Object.
            $idx = @{}
            foreach($o in $list) { $idx["$($o.Id)|$($o.Type)"] = $true }
            $script:_migFileObjectsIndex[$cacheKey] = $idx
        }
        else {
            # No existing file — start with an empty migration object. Path will be
            # resolved on first object add below.
            $migrationFile = $null
            $migFileObj = ([PSCustomObject]@{
                TenantId     = $script:organizationId
                Organization = $script:organizationName
                Objects      = [System.Collections.Generic.List[object]]::new()
            })
            $script:_migFileCache[$cacheKey]        = $migFileObj
            $script:_migFileObjectsIndex[$cacheKey] = @{}
        }
    }

    # Check if object is already processed. The cache key is normalised on the resolved
    # token's tenant id (see $cacheTenantId above) on BOTH read and write - earlier code
    # accidentally wrote to $script:Organization.Id which could diverge from the token's
    # tenant id and produce a permanent cache miss.
    $graphObj = Get-GraphMigrationObject $ObjectId $cacheTenantId

    $AADObjectCache = Get-CacheObject "AADObjectCache_$cacheTenantId" (@{})
    if(-not $graphObj -and $AADObjectCache.ContainsKey($ObjectId) -eq $false)
    {
        # Not in the cache, positively or negatively. This function never calls
        # Graph itself any more: it used to issue one direct GET per unknown id
        # right here, which on an export with no prefetch (single-policy export,
        # a depth bumped at the call site) was a serial round-trip per assignment
        # group at ~1.6 s each. The miss is queued instead and
        # Resolve-GraphMigrationObjectsPending fetches every queued id at once -
        # getByIds for directory objects, $batch for the rest - then replays this
        # call as a cache hit. Save-GraphMigrationFilesPending drains the queue,
        # and every export path ends with that.
        if(-not $script:_migPendingFetch) { $script:_migPendingFetch = [System.Collections.Generic.List[object]]::new() }
        [void]$script:_migPendingFetch.Add([PSCustomObject]@{
            ObjectId          = $ObjectId
            GraphAPI          = $GraphAPI
            ObjectTypeName    = $ObjectTypeName
            Folder            = $Folder
            TokenId           = $TokenId
            MigrationRoot     = $MigrationRoot
            MaxGroupDepth     = $MaxGroupDepth
            CurrentGroupDepth = $CurrentGroupDepth
        })
        return
    }

    if($graphObj)
    {
        $objectAdded = $false
        # Add object to cache (positive)
        if($AADObjectCache -is [Hashtable] -and $AADObjectCache.ContainsKey($ObjectId) -eq $false) { $AADObjectCache.Add($ObjectId, $graphObj) }

        $indexKey = "$ObjectId|$ObjectTypeName"
        $index    = $script:_migFileObjectsIndex[$cacheKey]
        if(-not $index.ContainsKey($indexKey)) {

            # FIX: was $GraphObject (undefined → wrote {Id=$null, DisplayName=$null} into
            # every MigrationTable.json entry). The correct variable is $graphObj.
            [void]$migFileObj.Objects.Add([PSCustomObject]@{
                Id          = $graphObj.Id
                DisplayName = $graphObj.displayName
                Type        = $ObjectTypeName
            })
            $index[$indexKey] = $true
            $objectAdded      = $true

            # First time we add anything against a still-unresolved migration file path:
            # resolve the canonical location now so the flush writes to the right path.
            # Cache is keyed on $Folder which is stable across calls in this bulk export,
            # so no key migration is needed. Save the resolved path back into the path
            # cache so cache-hit calls later in this run see the right file.
            if(-not $migrationFile) {
                if($MigrationRoot) {
                    # Caller passed the export root explicitly — most reliable.
                    $migrationFile = [IO.Path]::Combine($MigrationRoot, "MigrationTable.json")
                }
                else {
                    # Legacy heuristic: assume $Folder is a per-type subfolder and the
                    # export root is one level up. Works for the typical AddObjectType=true
                    # path but lands the file one level too high when $Folder is already
                    # the export root.
                    $folderInfo    = [IO.DirectoryInfo]$Folder
                    $migrationFile = [IO.Path]::Combine($folderInfo.Parent.FullName, "MigrationTable.json")
                    if($folderInfo.Parent.Name -ne $tokeInfo.TenantName) {
                        Write-Log "Parent name of the folder is not organization name: $($folderInfo.Parent.Name)" 2
                    }
                }
                $script:_migFilePathCache[$cacheKey] = $migrationFile
                Write-Log "Create new Migration file: $migrationFile"
            }

            # Defer the actual disk write until Save-GraphMigrationFilesPending at end of
            # bulk export — eliminates ~1200 read+write cycles for a typical run.
            # Dirty map: cacheKey ($Folder) -> resolved file path.
            $script:_migFileDirty[$cacheKey] = $migrationFile
        }

        # Sidecar json — write immediately, only once per new entry. Groups go to
        # Groups\, assignment filters to AssignmentFilters\ (same name as the type's
        # normal export folder, so the Assignments viewer finds them either way).
        if($objectAdded -and $ObjectTypeName -in @("Group","AssignmentFilter") -and $migrationFile)
        {
            $sidecarFolder = if($ObjectTypeName -eq "AssignmentFilter") { "AssignmentFilters" } else { "Groups" }
            $grouspPath = Join-Path ([IO.Path]::GetDirectoryName($migrationFile)) $sidecarFolder
            # New-Item, not mkdir: on Linux/macOS `mkdir` resolves to the native
            # binary, which rejects -Path/-Force and cannot be silenced by
            # -ErrorAction, so the folder was never created and every sidecar
            # write failed.
            if(-not (Test-Path $grouspPath)) { New-Item -ItemType Directory -Path $grouspPath -Force -ErrorAction SilentlyContinue | Out-Null }
            $FileName = [IO.Path]::Combine($grouspPath, "$((Remove-InvalidFileNameChars $graphObj.displayName)).json")

            # Defensive strip before serialising the group sidecar: any literal
            # `members` array or members nav links should never end up in a
            # sidecar that could later be re-POSTed to Graph (Graph treats
            # `members@odata.bind` on POST /groups as "create with these
            # members"). A group body fetched with $expand or a future Graph
            # projection change could carry one in without this.
            foreach ($m in @('members','members@odata.context','members@odata.associationLink','members@odata.navigationLink','members@odata.bind')) {
                if ($graphObj.PSObject.Properties.Name -contains $m) {
                    $graphObj.PSObject.Properties.Remove($m) | Out-Null
                }
            }

            Save-GraphObjectToFile $graphObj $FileName

            # Nested-group recursion. Only walks deeper when the caller opted in
            # (MaxGroupDepth > 1) and there's still budget. The dedup index above
            # ($script:_migFileObjectsIndex) prevents cycles and duplicate work.
            if($MaxGroupDepth -gt 1 -and $CurrentGroupDepth -lt $MaxGroupDepth)
            {
                try {
                    # Bulk export pre-walks the hierarchy in batches and stamps each
                    # parent body with `#NestedChildren = [{id, displayName}, ...]`.
                    # When that's present we skip the inline /members GET entirely —
                    # otherwise the recursion fires one un-batched call per parent
                    # which dominates wall time for tenants with many group
                    # assignments and depth > 1.
                    $members = $null
                    if($graphObj.PSObject.Properties.Name -contains '#NestedChildren') {
                        $members = @($graphObj.'#NestedChildren')
                    }
                    else {
                        # Fallback for callers that didn't pre-walk the hierarchy
                        # (single-policy export, depth bumped at call site, etc.).
                        # Use Graph's type-cast syntax to filter to nested groups
                        # server-side — skips users / service principals / devices
                        # without round-tripping them, and doesn't depend on
                        # @odata.type being present in the response. Smaller payload.
                        $membersResp = Invoke-MSGraphAPI -Url "groups/$ObjectId/members/microsoft.graph.group?`$select=id,displayName" -TokenId $TokenId -AllPages -ODataMetadata "Minimal"
                        $members = @()
                        if($membersResp -and $membersResp.value) { $members = @($membersResp.value) }
                        elseif($membersResp -is [Array])        { $members = @($membersResp) }
                    }

                    foreach($m in $members)
                    {
                        if(-not $m -or -not $m.id) { continue }
                        Add-GraphMigrationObject $m.id "groups" "Group" $Folder $TokenId `
                            -MigrationRoot      $MigrationRoot `
                            -MaxGroupDepth      $MaxGroupDepth `
                            -CurrentGroupDepth ($CurrentGroupDepth + 1)
                    }
                }
                catch {
                    Write-LogDebug "Nested-group recursion for $ObjectId failed at depth $CurrentGroupDepth/$MaxGroupDepth : $($_.Exception.Message)"
                }
            }
        }
    }
    else
    {
        # Negative-result cache: remember that this object id doesn't resolve so the next
        # caller doesn't hit Graph again for the same missing reference.
        if($AADObjectCache -is [Hashtable] -and $AADObjectCache.ContainsKey($ObjectId) -eq $false) { $AADObjectCache.Add($ObjectId, $null) }
        Set-CacheObject "AADObjectCache_$cacheTenantId" $AADObjectCache "TenantCache_$cacheTenantId"
        Write-Log "No $ObjectTypeName found with ID $($ObjectId). It might be deleted." 2
    }
}

# Resolve every migration object Add-GraphMigrationObject queued as a cache miss,
# in as few Graph calls as possible, then replay the queued calls as cache hits
# so they write their MigrationTable entry and sidecar exactly as before.
#
# Directory objects (groups, users, devices, service principals) go through
# /directoryObjects/getByIds, a thousand ids per POST. Anything else (assignment
# filters, apps) goes through Invoke-GraphBatchRequest, which decides batch versus
# direct from the UseBatchAPI setting. Ids are grouped by token first so a
# cross-tenant run never resolves one tenant's ids against another's directory.
#
# A replayed call can queue more work (nested-group recursion finds children that
# are not cached), so this loops; ten rounds is far beyond any real nesting depth.
# Ids whose fetch failed outright are dropped with a warning rather than retried
# every round - and are NOT negative-cached, because that cache persists and a
# transient failure must not mark a thousand live groups as deleted.
function Resolve-GraphMigrationObjectsPending
{
    if(-not $script:_migPendingFetch -or $script:_migPendingFetch.Count -eq 0) { return }

    $directoryTypes = @{ groups = 'group'; users = 'user'; devices = 'device'; servicePrincipals = 'servicePrincipal' }
    $round = 0
    while($script:_migPendingFetch.Count -gt 0 -and $round -lt 10)
    {
        $round++
        $pending = @($script:_migPendingFetch.ToArray())
        $script:_migPendingFetch.Clear()
        $unresolvable = [System.Collections.Generic.HashSet[string]]::new()

        $byToken = @{}
        foreach($p in $pending) {
            $k = [int]$p.TokenId
            if(-not $byToken.ContainsKey($k)) { $byToken[$k] = [System.Collections.Generic.List[object]]::new() }
            [void]$byToken[$k].Add($p)
        }

        # Raw queued token id -> the positive id it resolved to. The replay loop below
        # runs over every token at once, outside this loop, so it needs the mapping to
        # avoid replaying the raw 0 - which would be re-resolved against whatever is
        # default by then, i.e. possibly another tenant than the one just fetched from.
        $resolvedToken = @{}

        foreach($rawTokenId in @($byToken.Keys))
        {
            # Defensive: Add-GraphMigrationObject pins the id before it queues anything,
            # so a 0 should no longer reach here. If one does - an entry queued by
            # something else, or a future caller - resolve it ONCE here and use the
            # resolved id for both the fetch and the replay. Fetching one tenant's
            # directory and replaying against "whatever is default now" is the defect.
            $tokenInfo = Get-OperationTokenInfo $rawTokenId
            if(-not $tokenInfo -or -not $tokenInfo.TenantId) { continue }
            $tokenId   = if([int]$tokenInfo.Id -gt 0) { [int]$tokenInfo.Id } else { [int]$rawTokenId }
            $resolvedToken[[int]$rawTokenId] = $tokenId

            $cacheKey  = "AADObjectCache_$($tokenInfo.TenantId)"
            $cacheFile = "TenantCache_$($tokenInfo.TenantId)"
            $cache     = Get-CacheObject $cacheKey @{}
            if($cache -isnot [Hashtable]) { $cache = @{} }

            $idsByApi = @{}
            foreach($p in $byToken[$rawTokenId]) {
                $id = [string]$p.ObjectId
                if(-not $id -or $cache.ContainsKey($id)) { continue }
                $api = "$($p.GraphAPI)".Trim('/')
                if(-not $idsByApi.ContainsKey($api)) { $idsByApi[$api] = [System.Collections.Generic.HashSet[string]]::new() }
                [void]$idsByApi[$api].Add($id)
            }

            foreach($api in @($idsByApi.Keys))
            {
                $ids = @($idsByApi[$api])
                if($ids.Count -eq 0) { continue }
                $leaf = ($api -split '/')[-1]
                Write-Log "Migration objects: resolving $($ids.Count) queued $leaf id(s) for tenant $($tokenInfo.TenantId)"
                Write-Status -Detail ("Resolving {0} {1}" -f $ids.Count, $leaf) -SkipLog -Force

                if($directoryTypes.ContainsKey($leaf))
                {
                    for($i = 0; $i -lt $ids.Count; $i += 1000) {
                        $end   = [Math]::Min($i + 999, $ids.Count - 1)
                        $chunk = @($ids[$i..$end])
                        $resp  = $null
                        try {
                            $body = (@{ ids = $chunk; types = @($directoryTypes[$leaf]) } | ConvertTo-Json -Compress)
                            $resp = Invoke-MSGraphAPI -Url 'directoryObjects/getByIds' -Content $body -HttpMethod POST -TokenId $tokenId -ODataMetadata 'none'
                        }
                        catch { Write-LogError "Migration objects: getByIds failed for $($chunk.Count) $leaf id(s)" $_.Exception }
                        if($null -eq $resp) {
                            foreach($id in $chunk) { [void]$unresolvable.Add($id) }
                            continue
                        }
                        $seen = [System.Collections.Generic.HashSet[string]]::new()
                        foreach($o in @($resp.value)) {
                            if(-not $o -or -not $o.id) { continue }
                            if($o.PSObject.Properties['@odata.type']) { [void]$o.PSObject.Properties.Remove('@odata.type') }
                            $cache[[string]$o.id] = $o
                            [void]$seen.Add([string]$o.id)
                        }
                        foreach($id in $chunk) { if(-not $seen.Contains($id)) { $cache[$id] = $null } }
                    }
                }
                else
                {
                    $batch = [System.Collections.Generic.List[PSCustomObject]]::new()
                    $idByReq = @{}
                    $n = 0
                    foreach($id in $ids) {
                        $n++
                        $reqId = "migobj_$n"
                        [void]$batch.Add([PSCustomObject]@{ id = $reqId; method = 'GET'; url = "$api/$id"; headers = @{ Accept = 'application/json;odata.metadata=none' } })
                        $idByReq[$reqId] = $id
                    }
                    $results = @(Invoke-GraphBatchRequest -BatchObjects $batch -BatchType "Migration objects ($leaf)" -TokenId $tokenId -SkipWarnings -IncludedFailed)
                    $answered = [System.Collections.Generic.HashSet[string]]::new()
                    foreach($r in $results) {
                        $id = $idByReq["$($r.Id)"]
                        if(-not $id) { continue }
                        [void]$answered.Add($id)
                        if($r.body -and $r.Status -ge 200 -and $r.Status -lt 300) { $cache[$id] = $r.body }
                        else { $cache[$id] = $null }
                    }
                    foreach($id in $ids) { if(-not $answered.Contains($id)) { [void]$unresolvable.Add($id) } }
                }
            }

            Set-CacheObject $cacheKey $cache $cacheFile
        }

        if($unresolvable.Count -gt 0) {
            Write-Log "Migration objects: $($unresolvable.Count) id(s) could not be fetched and are left out of the migration table" 2
        }

        # Every queued id is now a cache hit, positive or negative: replay.
        foreach($p in $pending) {
            if($unresolvable.Contains([string]$p.ObjectId)) { continue }

            # The token the ids were actually fetched with, not the one that was queued.
            $replayToken = if($resolvedToken.ContainsKey([int]$p.TokenId)) { $resolvedToken[[int]$p.TokenId] } else { [int]$p.TokenId }

            Add-GraphMigrationObject $p.ObjectId $p.GraphAPI $p.ObjectTypeName $p.Folder $replayToken `
                -MigrationRoot $p.MigrationRoot -MaxGroupDepth $p.MaxGroupDepth -CurrentGroupDepth $p.CurrentGroupDepth
        }
        Write-Status -Detail "" -SkipLog -Force
    }

    if($script:_migPendingFetch.Count -gt 0) {
        Write-Log "Migration objects: $($script:_migPendingFetch.Count) id(s) still queued after $round round(s) - giving up" 2
        $script:_migPendingFetch.Clear()
    }
}

# Persist every MigrationTable.json that was mutated during this bulk export. Called
# once at the end of Start-GraphBulkExport so the in-memory cache batches all the
# Add-GraphMigrationObject mutations into a single write per file. Drains the
# queued cache misses first - on a fresh export nothing is dirty until they are.
function Save-GraphMigrationFilesPending
{
    Resolve-GraphMigrationObjectsPending

    if(-not $script:_migFileDirty -or $script:_migFileDirty.Count -eq 0) { return }

    $objectsByFile = @{}
    foreach($entry in @($script:_migFileDirty.GetEnumerator())) {
        $cacheKey = $entry.Key
        $filePath = $entry.Value
        if(-not $filePath) { continue }
        $obj = $script:_migFileCache[$cacheKey]
        if(-not $obj) { continue }

        $fileKey = [IO.Path]::GetFullPath($filePath).ToLowerInvariant()
        if(-not $objectsByFile.ContainsKey($fileKey)) {
            $objectsByFile[$fileKey] = [PSCustomObject]@{
                FilePath     = $filePath
                TenantId     = $obj.TenantId
                Organization = $obj.Organization
                Objects      = [System.Collections.Generic.List[object]]::new()
                Index        = @{}
            }
        }

        $target = $objectsByFile[$fileKey]
        foreach($migrationObject in @($obj.Objects)) {
            if(-not $migrationObject -or -not $migrationObject.Id -or -not $migrationObject.Type) { continue }
            $indexKey = "$($migrationObject.Id)|$($migrationObject.Type)"
            if($target.Index.ContainsKey($indexKey)) { continue }
            [void]$target.Objects.Add($migrationObject)
            $target.Index[$indexKey] = $true
        }
    }

    foreach($entry in @($objectsByFile.GetEnumerator())) {
        $filePath = $entry.Value.FilePath
        $obj = [PSCustomObject]@{
            TenantId     = $entry.Value.TenantId
            Organization = $entry.Value.Organization
            Objects      = $entry.Value.Objects.ToArray()
        }
        try {
            # Same encoding as the exported policy files - a MigrationTable in a
            # different encoding to the objects beside it is the same import
            # hazard.
            Save-GraphTextToFile (ConvertTo-GraphExportJson $obj -Depth 50) $filePath
            Write-LogDebug "Migration file flushed: $filePath"
        }
        catch {
            Write-LogError "Failed to flush Migration File $filePath" $_.Exception
        }
    }
    # Clear dirty set so a subsequent bulk export starts clean (cache itself is kept
    # so re-runs in the same process skip the disk-load cost).
    $script:_migFileDirty = @{}
}

function Get-GraphMigrationTableFromPath
{    
    param($Path)
    
    # Migration table must be located in the root of the import path
    $path = Expand-FileName $Path
    
    for($i = 0;$i -lt 2;$i++)
    {
        if($i -gt 0)
        {
            # Get parent directory
            $path = [io.path]::GetDirectoryName($path)
        }

        $migFileName = Join-Path $path "MigrationTable.json"
        try
        {
            if([IO.File]::Exists($migFileName))
            {                
                return $migFileName
            }
        }
        catch {}
    }

    # Downgraded from Write-Log (visible / file-logged / UI-marshalled per call) to
    # Write-LogDebug — for a fresh bulk export this happens once per per-type folder
    # and is informational, not a problem. The original Write-Log fired 1000+ times
    # before path caching landed, contributing ~2s of UI marshalling per run.
    Write-LogDebug "Could not find migration table for path '$Path'"
}

function Get-GraphMigrationObject
{
    param($ObjectId, $TenantId)

    $AADObjectCache = Get-CacheObject "AADObjectCache_$($TenantId)" @{}

    if($AADObjectCache.ContainsKey($ObjectId)) { return $AADObjectCache[$ObjectId] }
} 
