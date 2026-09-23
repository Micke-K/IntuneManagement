# Data layer for the Intune Tools view (Filter Usage report + enrollment-config
# lookups). Moved out of the per-backend UI files (UI/WPF + UI/Avalonia) where it
# was duplicated, into one module-internal home. Pure Graph/data - no XAML, no
# UIProvider, no backend row classes; the UI Start-*Load functions project the
# returned PSCustomObjects into their own CLR row types. (architecture R11)

# Map a payload's payloadType to the Graph endpoint that owns that resource +
# the display label we'll show in the Type column. Returns $null when the
# payloadType isn't one we recognize — caller falls back to a multi-endpoint
# probe (legacy behavior). Listed in the same order the old module checked.
function Get-IntuneFilterPayloadDispatch
{
    param([string]$PayloadType, [string]$PayloadId)

    if(-not $PayloadId) { return $null }

    switch ($PayloadType)
    {
        "application" {
            return [PSCustomObject]@{
                Url       = "deviceAppManagement/mobileApps/$PayloadId/?`$select=displayName"
                TypeLabel = "Application"
            }
        }
        "win32app" {
            # Old comment used "Proactive Remediations" — the endpoint is
            # /deviceHealthScripts which is exactly that, despite the
            # payloadType name being "win32app".
            return [PSCustomObject]@{
                Url       = "deviceManagement/deviceHealthScripts/$PayloadId/?`$select=displayName,isGlobalScript"
                TypeLabel = "Proactive Remediation"
            }
        }
        "deviceManagmentConfigurationAndCompliancePolicy" {
            # Yes, the typo "Managment" is upstream from Microsoft — kept verbatim.
            return [PSCustomObject]@{
                Url       = "deviceManagement/configurationPolicies/$PayloadId/?`$select=name,platforms,technologies,templateReference"
                TypeLabel = "Settings Catalog"
            }
        }
        "groupPolicyConfiguration" {
            return [PSCustomObject]@{
                Url       = "deviceManagement/groupPolicyConfigurations/$PayloadId/?`$select=displayName"
                TypeLabel = "Administrative Templates"
            }
        }
        default { return $null }
    }
}

# Which endpoints to probe when the payloadType has no dispatch entry.
#
# The generic set is the legacy three. Two payload types are known to be
# CONFIGURATION-OR-COMPLIANCE and never an app configuration, so they skip that
# third probe (measured live 2026-09-06: both deviceCompliancePolicies and
# deviceConfigurations were observed behind deviceConfigurationAndCompliance, so
# neither can become a single-URL dispatch entry - two probes is the floor).
function Get-IntuneFilterProbeEndpoints
{
    param([string]$PayloadType, [string]$PayloadId)

    $probes = @(
        @{ Suffix = "_dcp"; Url = "deviceManagement/deviceCompliancePolicies/$PayloadId/?`$select=displayName"; Type = "Compliance Policy" }
        @{ Suffix = "_dc";  Url = "deviceManagement/deviceConfigurations/$PayloadId/?`$select=displayName";   Type = "Device Configuration" }
    )

    if($PayloadType -notin @("deviceConfigurationAndCompliance", "androidEnterpriseConfiguration")) {
        $probes += @{ Suffix = "_mac"; Url = "deviceAppManagement/mobileAppConfigurations/$PayloadId/?`$select=displayName"; Type = "App Configuration" }
    }

    return $probes
}

# App Protection / managed-app policies, keyed by the payloadId shape.
#
# THE TRAP: /assignmentFilters/<id>/payloads reports a managed-app policy with
# payloadType "unknown" AND a BARE guid, while the policy's real id carries a
# type prefix (T_ targeted app protection, A_ app configuration, I_ Windows app
# protection, M_ information protection). Verified live: GET managedAppPolicies/
# <bare guid> is a 404, GET managedAppPolicies/T_<same guid> works. So these can
# never be resolved by id from the batch - list them once and match on the
# prefix-stripped id instead, the same way enrollment configurations are handled.
function Get-IntuneManagedAppPolicyLookup
{
    param([int]$TokenId)

    if($script:_intuneManagedAppPolicyCache) { return $script:_intuneManagedAppPolicyCache }

    $lookup = @{}
    try {
        $resp = Invoke-MSGraphAPI -Url "deviceAppManagement/managedAppPolicies?`$select=id,displayName" -TokenId $TokenId -AllPages
        foreach($p in @($resp.value)) {
            if(-not $p.id) { continue }
            $bare = "$($p.id)" -replace '^[A-Za-z]+_', ''
            if($bare) { $lookup[$bare] = $p }
        }
    }
    catch {
        Write-LogDebug "Failed to preload managed app policies: $($_.Exception.Message)"
    }

    $script:_intuneManagedAppPolicyCache = $lookup
    return $lookup
}

# Label for a managed-app policy row. App protection and app configuration both
# live under managedAppPolicies; the @odata.type is what separates them.
function Get-IntuneManagedAppPolicyTypeLabel
{
    param($Policy)

    $odata = "$($Policy.'@odata.type')"
    if($odata -match 'ManagedAppConfiguration') { return "App Configuration" }
    if($odata -match 'InformationProtection')   { return "Information Protection" }
    return "App Protection"
}

# Lazily load the deviceEnrollmentConfigurations list — used for payloads with
# payloadType=enrollmentConfiguration. There's no GET-by-id pattern that fits
# the batch model cleanly here (the configType discriminator decides what kind
# it is), so we fetch the list once and look up by Id in memory.
function Get-IntuneEnrollmentConfigurationLookup
{
    param([int]$TokenId)

    if($script:_intuneEnrollmentConfigCache) { return $script:_intuneEnrollmentConfigCache }

    $configs = @()
    try {
        $base = (Invoke-MSGraphAPI -Url "deviceManagement/deviceEnrollmentConfigurations?`$select=displayName,id,deviceEnrollmentConfigurationType" -TokenId $TokenId -AllPages)
        if($base -and $base.value) { $configs += @($base.value) }
        # The portal also separately enumerates enrollmentNotificationsConfiguration
        # (it's filtered out of the default list response) — preserve that fetch.
        $notif = (Invoke-MSGraphAPI -Url "deviceManagement/deviceEnrollmentConfigurations?`$filter=deviceEnrollmentConfigurationType eq 'EnrollmentNotificationsConfiguration'&`$select=displayName,id,deviceEnrollmentConfigurationType" -TokenId $TokenId -AllPages)
        if($notif -and $notif.value) { $configs += @($notif.value) }
    }
    catch {
        Write-LogDebug "Failed to preload enrollment configurations: $($_.Exception.Message)"
    }

    $lookup = @{}
    foreach($c in $configs) {
        if($c.id) { $lookup[$c.id] = $c }
    }
    $script:_intuneEnrollmentConfigCache = $lookup
    return $lookup
}

function Get-IntuneEnrollmentConfigurationTypeLabel
{
    # Friendly label for the Type column based on the discriminator on each
    # deviceEnrollmentConfiguration subtype.
    param([string]$ConfigType)

    switch -Regex ($ConfigType)
    {
        '(?i)^enrollmentNotificationsConfiguration$'              { return "Enrollment notifications" }
        '(?i)^windows10EnrollmentCompletionPageConfiguration$'    { return "Enrollment Status Page" }
        '(?i)^limit$'                                             { return "Enrollment Limit" }
        '(?i)^singlePlatformRestriction$'                         { return "Enrollment Restriction" }
        '(?i)^platformRestrictions$'                              { return "Enrollment Restrictions (default)" }
        '(?i)^windowsHelloForBusiness$'                           { return "Windows Hello for Business" }
        '(?i)^deviceComanagementAuthorityConfiguration$'          { return "Co-management Authority" }
        '(?i)^windowsRestore$'                                    { return "Windows Restore" }
        default                                                   { return "Enrollment Configuration" }
    }
}

function Get-IntuneFilterUsageData
{
    # Pull every assignment filter, then for each one fetch /payloads. Each
    # payload references one policy by (payloadId, payloadType) — we resolve
    # the policy display name in a second batch, then resolve group display
    # names in a third batch. Three round-trip-batched phases instead of
    # one-per-payload + one-per-group.

    $tokenId = Get-DefaultTokenId

    # Phase 1: list all filters.
    Write-Status "Loading assignment filters..."
    $filterResp = Invoke-MSGraphAPI -Url "deviceManagement/assignmentFilters" -TokenId $tokenId -AllPages
    if(-not $filterResp) { return @() }
    $filters = @()
    if($filterResp.value) { $filters = @($filterResp.value) }
    elseif($filterResp -is [Array]) { $filters = @($filterResp) }
    if($filters.Count -eq 0) { return @() }

    # Phase 2: /payloads per filter in one batch.
    Write-Status "Fetching payloads for $($filters.Count) filter(s)..."
    $payloadBatch = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach($f in $filters) {
        [void]$payloadBatch.Add([PSCustomObject]@{
            id      = [string]$f.id
            method  = "GET"
            url     = "deviceManagement/assignmentFilters/$($f.id)/payloads"
            headers = @{ "Accept" = "application/json" }
        })
    }
    $payloadResults = Invoke-GraphBatchRequest -BatchObjects $payloadBatch -BatchType "FilterPayloads" -TokenId $tokenId

    # Build a (filter, payload) work list from the batch responses.
    $filterById = @{}
    foreach($f in $filters) { $filterById[[string]$f.id] = $f }

    $filtersWithPayload = [System.Collections.Generic.HashSet[string]]::new()
    $work = [System.Collections.Generic.List[object]]::new()
    foreach($r in $payloadResults) {
        if(-not $r.body) { continue }
        $values = @()
        if($r.body.value) { $values = @($r.body.value) }
        $filter = $filterById["$($r.Id)"]
        if(-not $filter) { continue }
        foreach($p in $values) {
            [void]$filtersWithPayload.Add([string]$filter.id)
            [void]$work.Add([PSCustomObject]@{ Filter = $filter; Payload = $p })
        }
    }

    if($work.Count -eq 0) {
        Write-Log "No filter payloads found across $($filters.Count) filter(s). Showing filters as not used."
        return @($filters | ForEach-Object {
            [PSCustomObject]@{
                FilterName  = $_.displayName
                Platform    = [string]$_.platform
                FilterType  = [string]$_.assignmentFilterManagementType
                PolicyName  = "<not used>"
                PayloadType = "No payloads"
                Mode        = ""
                GroupId     = ""
                GroupName   = ""
            }
        })
    }

    # Phase 3: resolve policy display name per payload. We batch by sub-request
    # id, then post-process — each work item gets a unique GUID prefix so we
    # can match the response back to its (filter, payload) pair. Three-way
    # fallback (deviceCompliancePolicies / deviceConfigurations /
    # mobileAppConfigurations) preserved for unrecognized payloadType values.
    $policyBatch       = [System.Collections.Generic.List[PSCustomObject]]::new()
    $workByGuid        = @{}
    $manualEnrolment   = [System.Collections.Generic.List[object]]::new()
    $manualManagedApp  = [System.Collections.Generic.List[object]]::new()
    $unknownTypeCounts = @{}   # diagnostic — surface unrecognized payloadTypes for future dispatch tuning

    foreach($w in $work) {
        $payload = $w.Payload
        $guid    = [Guid]::NewGuid().Guid
        $workByGuid[$guid] = $w
        $w | Add-Member -NotePropertyName "_BatchGuid" -NotePropertyValue $guid -Force

        if($payload.payloadType -eq "enrollmentConfiguration") {
            # Resolved in-memory from the pre-loaded list — skip the batch.
            [void]$manualEnrolment.Add($w)
            continue
        }

        # Managed-app policies report payloadType "unknown" with a prefix-less id
        # (see Get-IntuneManagedAppPolicyLookup). Resolve those in memory; anything
        # else calling itself "unknown" still falls through to the probe below.
        if($payload.payloadType -eq "unknown") {
            $mamLookup = Get-IntuneManagedAppPolicyLookup -TokenId $tokenId
            if($mamLookup -and $mamLookup.ContainsKey([string]$payload.payloadId)) {
                [void]$manualManagedApp.Add($w)
                continue
            }
        }

        $dispatch = Get-IntuneFilterPayloadDispatch -PayloadType $payload.payloadType -PayloadId $payload.payloadId
        if($dispatch) {
            [void]$policyBatch.Add([PSCustomObject]@{
                id      = $guid
                method  = "GET"
                url     = $dispatch.Url
                headers = @{ "Accept" = "application/json" }
            })
            $w | Add-Member -NotePropertyName "_TypeLabel" -NotePropertyValue $dispatch.TypeLabel -Force
        }
        else {
            # Unknown payloadType — try the three common endpoints in
            # parallel and pick the one that returns 200. Track the unrecognized
            # type so we can extend Get-IntuneFilterPayloadDispatch later (cuts
            # the batch sub-request count from 3 to 1 for known types).
            $ptKey = if([string]::IsNullOrEmpty($payload.payloadType)) { "<null>" } else { [string]$payload.payloadType }
            if(-not $unknownTypeCounts.ContainsKey($ptKey)) { $unknownTypeCounts[$ptKey] = 0 }
            $unknownTypeCounts[$ptKey]++

            foreach($probe in (Get-IntuneFilterProbeEndpoints -PayloadType $payload.payloadType -PayloadId $payload.payloadId)) {
                $subId = "$guid$($probe.Suffix)"
                [void]$policyBatch.Add([PSCustomObject]@{
                    id      = $subId
                    method  = "GET"
                    url     = $probe.Url
                    headers = @{ "Accept" = "application/json" }
                })
            }
        }
    }

    if($unknownTypeCounts.Count -gt 0) {
        $summary = ($unknownTypeCounts.GetEnumerator() | Sort-Object Key | ForEach-Object { "$($_.Key):$($_.Value)" }) -join ", "
        # Informational, not a warning: nothing is wrong and the end user can do
        # nothing about it. It is a note for whoever maintains the dispatch table.
        Write-Log "Intune Filter Usage: payloadType(s) resolved by endpoint probe rather than a direct lookup (extend Get-IntuneFilterPayloadDispatch to optimize): $summary"
    }

    $policyResults = @()
    if($policyBatch.Count -gt 0) {
        Write-Status "Resolving $($policyBatch.Count) policy reference(s)..."
        $policyResults = @(Invoke-GraphBatchRequest -BatchObjects $policyBatch -BatchType "FilterPayloadNames" -TokenId $tokenId -SkipWarnings -IncludedFailed)
    }

    # Bucket policy results by base GUID; pick the first success per work item.
    $resolvedByGuid = @{}
    foreach($r in $policyResults) {
        $baseGuid = "$($r.Id)"
        # Strip any trailing "_dcp" / "_dc" / "_mac" suffix to get the work-item key.
        $baseGuid = $baseGuid -replace '_(dcp|dc|mac)$', ''
        if(-not $workByGuid.ContainsKey($baseGuid)) { continue }
        if($r.Status -ge 300 -or -not $r.body) { continue }
        if(-not $resolvedByGuid.ContainsKey($baseGuid)) {
            # For unknown-payloadType probes, derive the type label from which
            # endpoint actually answered (suffix tells us).
            $derivedType = $null
            if("$($r.Id)" -match '_(dcp|dc|mac)$') {
                $derivedType = switch ($matches[1]) {
                    'dcp' { "Compliance Policy" }
                    'dc'  { "Device Configuration" }
                    'mac' { "App Configuration" }
                }
            }
            $resolvedByGuid[$baseGuid] = [PSCustomObject]@{ Body = $r.body; DerivedType = $derivedType }
        }
    }

    # Resolve enrollment-configuration payloads from the in-memory list.
    $enrollmentLookup = $null
    if($manualEnrolment.Count -gt 0) {
        $enrollmentLookup = Get-IntuneEnrollmentConfigurationLookup -TokenId $tokenId
    }

    # Same for managed-app policies (already loaded above if any matched).
    $managedAppLookup = $null
    if($manualManagedApp.Count -gt 0) {
        $managedAppLookup = Get-IntuneManagedAppPolicyLookup -TokenId $tokenId
    }

    # Build the row list. Defer group-name resolution to Phase 4.
    $rows         = [System.Collections.Generic.List[object]]::new()
    $allGroupIds  = [System.Collections.Generic.HashSet[string]]::new()

    foreach($w in $work) {
        $filter  = $w.Filter
        $payload = $w.Payload

        $policyName = $null
        $typeLabel  = $w._TypeLabel

        if($payload.payloadType -eq "enrollmentConfiguration" -and $enrollmentLookup) {
            $cfg = $enrollmentLookup[$payload.payloadId]
            if($cfg) {
                $policyName = $cfg.displayName
                $typeLabel  = Get-IntuneEnrollmentConfigurationTypeLabel $cfg.deviceEnrollmentConfigurationType
            }
        }
        elseif($managedAppLookup -and $managedAppLookup.ContainsKey([string]$payload.payloadId)) {
            $mam = $managedAppLookup[[string]$payload.payloadId]
            $policyName = $mam.displayName
            $typeLabel  = Get-IntuneManagedAppPolicyTypeLabel $mam
        }
        else {
            $resolved = $resolvedByGuid[$w._BatchGuid]
            if($resolved) {
                $body = $resolved.Body
                $policyName = if($body.name) { $body.name } else { $body.displayName }
                if($resolved.DerivedType) { $typeLabel = $resolved.DerivedType }
                # Settings Catalog templateReference can carry a richer label.
                if($payload.payloadType -eq "deviceManagmentConfigurationAndCompliancePolicy" -and $body.templateReference -and $body.templateReference.templateDisplayName) {
                    $typeLabel = "Settings Catalog ($($body.templateReference.templateDisplayName))"
                }
            }
        }

        if(-not $policyName) {
            # Couldn't resolve. Keep the row so the filter itself is still visible
            # and the user can see the stale/unsupported payload reference.
            Write-Log "Filter '$($filter.displayName)': failed to resolve payload $($payload.payloadId) (type: $($payload.payloadType))" 2
            $policyName = "<unresolved: $($payload.payloadId)>"
            if(-not $typeLabel) {
                $typeLabel = if($payload.payloadType) { $payload.payloadType } else { "Unknown payload" }
            }
        }

        $mode = if($payload.assignmentFilterType -eq "Include") { "Include" } else { "Exclude" }

        if($payload.groupId) { [void]$allGroupIds.Add([string]$payload.groupId) }

        [void]$rows.Add([PSCustomObject]@{
            FilterName  = $filter.displayName
            Platform    = [string]$filter.platform
            FilterType  = [string]$filter.assignmentFilterManagementType
            PolicyName  = $policyName
            PayloadType = if($typeLabel) { $typeLabel } else { $payload.payloadType }
            Mode        = $mode
            GroupId     = $payload.groupId
            GroupName   = $payload.groupId   # placeholder; resolved below
        })
    }

    foreach($filter in $filters) {
        if($filtersWithPayload.Contains([string]$filter.id)) { continue }
        [void]$rows.Add([PSCustomObject]@{
            FilterName  = $filter.displayName
            Platform    = [string]$filter.platform
            FilterType  = [string]$filter.assignmentFilterManagementType
            PolicyName  = "<not used>"
            PayloadType = "No payloads"
            Mode        = ""
            GroupId     = ""
            GroupName   = ""
        })
    }

    # Phase 4: resolve group names in one batch. Pre-seed the well-known
    # virtual groups that don't resolve via /groups (these are baked into the
    # Intune assignment model). Cache otherwise.
    $groupNames = @{
        "adadadad-808e-44e2-905a-0b7873a8a531" = "All Devices"
        "acacacac-9df4-4c7d-9d50-4ef0226f57a9" = "All Users"
    }
    $toLookup = @($allGroupIds | Where-Object { -not $groupNames.ContainsKey($_) })

    if($toLookup.Count -gt 0) {
        Write-Status "Resolving $($toLookup.Count) group(s)..."
        $groupBatch = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach($gid in $toLookup) {
            [void]$groupBatch.Add([PSCustomObject]@{
                id      = [string]$gid
                method  = "GET"
                url     = "groups/$gid/?`$select=displayName,id"
                headers = @{ "Accept" = "application/json" }
            })
        }
        $groupResults = @(Invoke-GraphBatchRequest -BatchObjects $groupBatch -BatchType "FilterGroupNames" -TokenId $tokenId -SkipWarnings -IncludedFailed)
        foreach($r in $groupResults) {
            if($r.Status -ge 300 -or -not $r.body) { continue }
            $name = if($r.body.displayName) { $r.body.displayName } else { "$($r.Id)" }
            $groupNames["$($r.Id)"] = $name
        }
    }

    foreach($row in $rows) {
        if(-not $row.GroupId) {
            $row.GroupName = ""
            continue
        }
        $key = [string]$row.GroupId
        if($groupNames.ContainsKey($key)) {
            $row.GroupName = $groupNames[$key]
        }
        else {
            # Group didn't resolve via /groups (deleted, throttled, or denied) —
            # mark explicitly so the user doesn't mistake the GUID for a name.
            # Short-form GUID keeps the column narrow.
            $short = if($key.Length -ge 8) { $key.Substring(0, 8) } else { $key }
            $row.GroupName = "<unresolved: $short...>"
        }
    }

    return $rows
}
