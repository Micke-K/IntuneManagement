# Per-policy Entra group + assignment-filter migration on IMPORT.
#
# When a policy exported from another tenant is imported, its json still carries the
# SOURCE tenant's group ids (assignments, CA include/exclude groups, notification CC
# lists, role members, ...), assignment-filter ids (deviceAndAppManagementAssignment-
# FilterId on every target type) and - for some payloads - group SIDs. This module
# resolves ONLY the references the imported policy actually carries (unlike the 3.x
# behavior of bulk-creating every group in the migration table up front):
#
#   1. Update-JsonForEnvironment scans the policy json for every GUID and SID
#      (Get-DependencyIDs) and hands them to Resolve-GraphMigrationGroups.
#   2. A GUID/SID is only considered when it matches an EXPORTED group - the
#      Groups\*.json sidecars plus the MigrationTable.json Group entries under the
#      import folder's export root. Anything else is ignored (no Graph calls).
#   3. Each matched group is resolved in the target tenant: by id first (fast path,
#      also covers already-migrated ids), then by displayName.
#   4. Missing groups are CREATED - gated on the CreateGroupOnImport setting - from
#      the sidecar stripped to cloud-creatable properties (dynamic groups keep their
#      membershipRule/groupTypes). A group whose sidecar shows onPremisesSyncEnabled
#      is an AD-synced group: ConvertSyncedGroupOnImport=true recreates it as a
#      cloud Entra group, false skips it with a warning.
#   5. The returned maps translate source id -> target id and source SID -> target
#      SID in the policy json.
#
# Results are cached per (export root + target tenant) so a group shared by many
# policies in one bulk import is resolved/created exactly once.

# --- caches -------------------------------------------------------------------
# $script:_groupSidecarIndex : exportRoot -> @{ ById = @{ id -> entry }; BySid = @{ sid -> entry } }
#   entry = @{ Id; Name; Sid; Synced; Sidecar (full object or $null) }
# $script:_groupMigrationMap : "exportRoot|targetTenant" -> @{ sourceId -> resolved @{ Id; Sid } or $null (unresolvable) }

function Reset-GraphGroupMigrationCache
{
    $script:_groupSidecarIndex = @{}
    $script:_groupMigrationMap = @{}
    $script:_filterListCache   = @{}
}

# Resolve the export root for a policy loaded from file: the folder that holds
# MigrationTable.json / Groups\. Probes the file's own folder first (AddObjectType
# = false exports) then its parent (the normal per-type subfolder layout) - same
# 2-level heuristic as Get-GraphMigrationTableFromPath.
function Get-GraphGroupMigrationRoot
{
    param($PolicyObject)

    $fileInfo = ?? $PolicyObject._ClonedFromObject.FileInfo $PolicyObject.FileInfo
    if(-not $fileInfo) { return $null }

    $folder = [IO.Path]::GetDirectoryName($fileInfo.FullName)
    for($i = 0; $i -lt 2 -and $folder; $i++)
    {
        if([IO.File]::Exists([IO.Path]::Combine($folder, "MigrationTable.json")) -or
           [IO.Directory]::Exists([IO.Path]::Combine($folder, "Groups")) -or
           [IO.Directory]::Exists([IO.Path]::Combine($folder, "AssignmentFilters")))
        {
            return $folder
        }
        $folder = [IO.Path]::GetDirectoryName($folder)
    }
    return $null
}

# Build (and cache) the exported-group index for an export root: every group the
# export knows about, keyed by source id and by SID. Sidecars win over bare
# migration-table entries (they carry the full object for re-creation).
function Get-GraphGroupSidecarIndex
{
    param([string]$ExportRoot)

    if(-not $script:_groupSidecarIndex) { $script:_groupSidecarIndex = @{} }
    if($script:_groupSidecarIndex.ContainsKey($ExportRoot)) { return $script:_groupSidecarIndex[$ExportRoot] }

    $index = @{ ById = @{}; BySid = @{} }

    $groupsPath = [IO.Path]::Combine($ExportRoot, "Groups")
    if([IO.Directory]::Exists($groupsPath))
    {
        foreach($file in [IO.Directory]::EnumerateFiles($groupsPath, "*.json"))
        {
            try
            {
                $g = ConvertFrom-Json ([IO.File]::ReadAllText($file))
                if(-not $g.id) { continue }
                $entry = @{
                    Kind    = "Group"
                    Id      = [string]$g.id
                    Name    = [string]$g.displayName
                    Sid     = [string]$g.securityIdentifier
                    Synced  = ($g.onPremisesSyncEnabled -eq $true)
                    Sidecar = $g
                }
                $index.ById[$entry.Id] = $entry
                if($entry.Sid) { $index.BySid[$entry.Sid] = $entry }
            }
            catch
            {
                Write-Log "Failed to parse group sidecar $file" 2
            }
        }
    }

    # Assignment-filter sidecars (written by Add-GraphMigrationObject; the folder name
    # matches the AssignmentFilters type's normal export folder, so a full bulk export
    # that included the type feeds the same index).
    $filtersPath = [IO.Path]::Combine($ExportRoot, "AssignmentFilters")
    if([IO.Directory]::Exists($filtersPath))
    {
        foreach($file in [IO.Directory]::EnumerateFiles($filtersPath, "*.json"))
        {
            try
            {
                $f = ConvertFrom-Json ([IO.File]::ReadAllText($file))
                if(-not $f.id) { continue }
                $index.ById[[string]$f.id] = @{
                    Kind    = "Filter"
                    Id      = [string]$f.id
                    Name    = [string]$f.displayName
                    Sid     = $null
                    Synced  = $false
                    Sidecar = $f
                }
            }
            catch
            {
                Write-Log "Failed to parse assignment-filter sidecar $file" 2
            }
        }
    }

    # Migration-table entries without a sidecar (older exports): name-only.
    $migFile = [IO.Path]::Combine($ExportRoot, "MigrationTable.json")
    if([IO.File]::Exists($migFile))
    {
        try
        {
            $migObj = ConvertFrom-Json ([IO.File]::ReadAllText($migFile))
            foreach($m in @($migObj.Objects))
            {
                if($m.Type -notin @("Group","AssignmentFilter") -or -not $m.Id) { continue }
                if($index.ById.ContainsKey([string]$m.Id)) { continue }
                $index.ById[[string]$m.Id] = @{
                    Kind    = if($m.Type -eq "AssignmentFilter") { "Filter" } else { "Group" }
                    Id      = [string]$m.Id
                    Name    = [string]$m.DisplayName
                    Sid     = $null
                    Synced  = $false
                    Sidecar = $null
                }
            }
        }
        catch
        {
            Write-Log "Failed to parse migration table $migFile" 2
        }
    }

    $script:_groupSidecarIndex[$ExportRoot] = $index
    return $index
}

# Create a group in the target tenant from an exported entry. Sidecar-based when
# available (stripped to cloud-creatable properties; dynamic groups keep their
# rule), else a default cloud security group. Returns the created group or $null.
function New-GraphMigrationGroup
{
    param($Entry, [int]$TokenId)

    $keepProps = @("displayName","description","mailEnabled","mailNickname","securityEnabled",
                   "membershipRule","groupTypes","membershipRuleProcessingState")

    $body = [ordered]@{}
    if($Entry.Sidecar)
    {
        foreach($prop in $Entry.Sidecar.PSObject.Properties)
        {
            if($prop.Name -notin $keepProps) { continue }
            if($null -eq $prop.Value) { continue }
            $body[$prop.Name] = $prop.Value
        }
    }

    if(-not $body["displayName"]) { $body["displayName"] = $Entry.Name }
    if(-not $body["displayName"]) { return $null }
    $body["displayName"] = ([string]$body["displayName"]).Trim()
    if(-not $body.Contains("mailEnabled"))     { $body["mailEnabled"] = $false }
    if(-not $body.Contains("securityEnabled")) { $body["securityEnabled"] = $true }
    # mailNickname is mandatory on POST /groups and often null on synced/security groups.
    if(-not $body["mailNickname"]) { $body["mailNickname"] = (New-Guid).Guid.SubString(0, 10) }

    Write-Log "Creating Entra group '$($body["displayName"])' in target tenant (referenced by imported policy)"
    return Invoke-MSGraphAPI -Url "/groups" -HttpMethod "POST" -Content (ConvertTo-Json $body -Depth 10) -TokenId $TokenId
}

# Create an assignment filter in the target tenant from an exported sidecar. Filters
# without a sidecar (table-only entries) cannot be created - platform + rule are
# mandatory and unknowable from the name alone.
function New-GraphMigrationFilter
{
    param($Entry, [int]$TokenId)

    if(-not $Entry.Sidecar) {
        Write-Log "Assignment filter '$($Entry.Name)' has no exported sidecar (platform/rule unknown) - cannot create it in the target tenant" 2
        return $null
    }

    $keepProps = @("displayName","description","platform","rule","assignmentFilterManagementType")
    $body = [ordered]@{}
    foreach($prop in $Entry.Sidecar.PSObject.Properties)
    {
        if($prop.Name -notin $keepProps) { continue }
        if($null -eq $prop.Value) { continue }
        $body[$prop.Name] = $prop.Value
    }
    if(-not $body["displayName"] -or -not $body["rule"] -or -not $body["platform"]) {
        Write-Log "Assignment filter sidecar for '$($Entry.Name)' is missing displayName/platform/rule - cannot create" 2
        return $null
    }

    Write-Log "Creating assignment filter '$($body["displayName"])' in target tenant (referenced by imported policy)"
    return Invoke-MSGraphAPI -Url "deviceManagement/assignmentFilters" -HttpMethod "POST" -Content (ConvertTo-Json $body -Depth 10) -TokenId $TokenId
}

# Target-tenant assignment-filter list, cached per map key. Intune endpoints do not
# reliably honor server-side displayName $filter, so name matching is client-side.
function Get-GraphMigrationTargetFilters
{
    param([string]$MapKey, [int]$TokenId)

    if(-not $script:_filterListCache) { $script:_filterListCache = @{} }
    if($script:_filterListCache.ContainsKey($MapKey)) { return $script:_filterListCache[$MapKey] }

    $list = @()
    $resp = Invoke-MSGraphAPI -Url "deviceManagement/assignmentFilters?`$select=id,displayName" -ODataMetadata "none" -NoError -TokenId $TokenId
    if($resp -and $resp.Value) { $list = @($resp.Value) }
    $script:_filterListCache[$MapKey] = $list
    return $list
}

# Main entry point - called from Update-JsonForEnvironment with every GUID and SID
# found in the imported policy json. Covers BOTH Entra groups and assignment filters
# (the sidecar index tags each entry with its Kind). Returns
# @{ IdMap = @{src->target}; SidMap = @{src->target} } containing ONLY references
# that need rewriting.
function Resolve-GraphMigrationGroups
{
    param($Guids, $Sids, $PolicyObject, [int]$TokenId)

    $result = @{ IdMap = @{}; SidMap = @{} }

    # Same-tenant import: source ids are valid as-is, nothing to translate.
    #
    # Get-OperationTokenInfo, not Get-TokenInfo: the import path hands its own TokenId
    # down, and 0 means "every token" to the list accessor. With two tenants signed in,
    # [string]$tokenInfo.TenantId was "tenant-a tenant-b", so the same-tenant early
    # return below never fired - a same-tenant import ran the whole cross-tenant group
    # translation, and cached its results under a map key naming both tenants.
    $tokenInfo = Get-OperationTokenInfo $TokenId
    $targetTenant = if($tokenInfo -and $tokenInfo.TenantId) { [string]$tokenInfo.TenantId } else { [string](Get-CurrentTenantId) }
    if(-not $targetTenant) { return $result }
    if($PolicyObject.TenantID -and ([string]$PolicyObject.TenantID) -eq $targetTenant) { return $result }

    $exportRoot = Get-GraphGroupMigrationRoot $PolicyObject
    if(-not $exportRoot) { return $result }

    $index = Get-GraphGroupSidecarIndex $exportRoot
    if($index.ById.Count -eq 0) { return $result }

    if(-not $script:_groupMigrationMap) { $script:_groupMigrationMap = @{} }
    $mapKey = "$exportRoot|$targetTenant"
    if(-not $script:_groupMigrationMap.ContainsKey($mapKey)) { $script:_groupMigrationMap[$mapKey] = @{} }
    $map = $script:_groupMigrationMap[$mapKey]

    $createEnabled  = (Get-SettingValue "CreateGroupOnImport") -ne $false
    $convertSynced  = (Get-SettingValue "ConvertSyncedGroupOnImport") -ne $false

    # Collect the entries this policy actually references (by id or by SID).
    $wanted = @{}
    foreach($guid in @($Guids)) {
        if($guid -and $index.ById.ContainsKey([string]$guid)) { $wanted[[string]$guid] = $index.ById[[string]$guid] }
    }
    foreach($sid in @($Sids)) {
        if($sid -and $index.BySid.ContainsKey([string]$sid)) {
            $entry = $index.BySid[[string]$sid]
            $wanted[$entry.Id] = $entry
        }
    }
    if($wanted.Count -eq 0) { return $result }

    foreach($sourceId in $wanted.Keys)
    {
        $entry = $wanted[$sourceId]

        # Cached resolution (positive or negative) from an earlier policy in this run.
        if($map.ContainsKey($sourceId))
        {
            $resolved = $map[$sourceId]
        }
        else
        {
            $resolved = $null
            $target = $null

            if($entry.Kind -eq "Filter")
            {
                # 1. By id - covers same-guid edge cases and pre-migrated environments.
                $target = Invoke-MSGraphAPI -Url "deviceManagement/assignmentFilters/$sourceId" -ODataMetadata "none" -NoError -TokenId $TokenId

                # 2. By display name - client-side match against the cached target list
                # (Intune endpoints do not reliably honor server-side displayName filters).
                if(-not $target -and $entry.Name)
                {
                    $wantName = $entry.Name.Trim()
                    $target = Get-GraphMigrationTargetFilters -MapKey $mapKey -TokenId $TokenId |
                        Where-Object { $_.displayName -and $_.displayName.Trim() -eq $wantName } | Select-Object -First 1
                }

                # 3. Create (same gate as groups - one 'create referenced objects' toggle).
                if(-not $target)
                {
                    if(-not $createEnabled)
                    {
                        Write-Log "Assignment filter '$($entry.Name)' ($sourceId) does not exist in the target tenant and CreateGroupOnImport is disabled - reference not translated" 2
                    }
                    else
                    {
                        $target = New-GraphMigrationFilter -Entry $entry -TokenId $TokenId
                        if(-not $target -and $entry.Name)
                        {
                            # The create can fail on a duplicate name when an
                            # earlier resolution pass already created the filter
                            # but a transient list failure cached an empty
                            # target list - re-list fresh and match once more.
                            if($script:_filterListCache) { $script:_filterListCache.Remove($mapKey) }
                            $wantName = $entry.Name.Trim()
                            $target = Get-GraphMigrationTargetFilters -MapKey $mapKey -TokenId $TokenId |
                                Where-Object { $_.displayName -and $_.displayName.Trim() -eq $wantName } | Select-Object -First 1
                        }
                        if($target -and $script:_filterListCache -and $script:_filterListCache.ContainsKey($mapKey)) {
                            # Keep the cached target list current for later policies.
                            $script:_filterListCache[$mapKey] = @($script:_filterListCache[$mapKey]) + @($target)
                        }
                    }
                }
            }
            else
            {
                # 1. By id - covers same-guid edge cases and pre-migrated environments.
                $target = Invoke-MSGraphAPI -Url "/groups/$sourceId" -ODataMetadata "none" -NoError -TokenId $TokenId

                # 2. By display name.
                if(-not $target -and $entry.Name)
                {
                    # Percent-encoded literal: a raw & or # in the name used to split the
                    # query, Graph 400'd, -NoError hid it, and step 3 created a duplicate.
                    $literal = ConvertTo-ODataStringLiteral $entry.Name.Trim()
                    $resp = Invoke-MSGraphAPI -Url "/groups?`$filter=displayName eq $literal&`$select=id,displayName,securityIdentifier" -ODataMetadata "none" -NoError -TokenId $TokenId
                    if($resp -and $resp.Value) { $target = $resp.Value | Select-Object -First 1 }
                }

                # 3. Create.
                if(-not $target)
                {
                    if(-not $createEnabled)
                    {
                        Write-Log "Group '$($entry.Name)' ($sourceId) does not exist in the target tenant and CreateGroupOnImport is disabled - reference not translated" 2
                    }
                    elseif($entry.Synced -and -not $convertSynced)
                    {
                        Write-Log "Group '$($entry.Name)' ($sourceId) is an AD-synced group and ConvertSyncedGroupOnImport is disabled - reference not translated" 2
                    }
                    else
                    {
                        if($entry.Synced)
                        {
                            Write-Log "Group '$($entry.Name)' is AD-synced in the source tenant - creating it as a cloud Entra group (ConvertSyncedGroupOnImport)" 2
                        }
                        $target = New-GraphMigrationGroup -Entry $entry -TokenId $TokenId
                        if(-not $target) { Write-Log "Failed to create group '$($entry.Name)' in target tenant" 3 }
                    }
                }
            }

            $resolved = if($target) { @{ Id = [string]$target.id; Sid = [string]$target.securityIdentifier } } else { $null }
            $map[$sourceId] = $resolved
        }

        if(-not $resolved) { continue }

        if($resolved.Id -and $resolved.Id -ne $sourceId) { $result.IdMap[$sourceId] = $resolved.Id }
        if($entry.Sid -and $resolved.Sid -and $resolved.Sid -ne $entry.Sid) { $result.SidMap[$entry.Sid] = $resolved.Sid }
    }

    return $result
}
