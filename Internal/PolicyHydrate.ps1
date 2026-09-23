# Unified policy hydration orchestrator.
#
# Every caller that needs a fully populated IntunePolicyBase — bulk export,
# single-policy export, UI list-row open, Compare, Copy — funnels through
# Invoke-PolicyHydrate. The orchestrator owns:
#   1. Body fetch     — N=1 direct GET, N>1 parallel $batch (chunks of 20).
#   2. Nav properties — Add-GraphNavigationProperties per row that needs it.
#   3. Sub-resources  — Invoke-PolicySubresourceFetch drives the per-class
#                       GetSubResourceBatchRequests / ApplySubResourceBatchResult
#                       / FinalizeSubResources contract (branding images, target
#                       apps, role assignments, reusable settings, ToU files, …).
#                       This replaced the old Invoke-PolicyExtraData /
#                       Sync-BulkExport* fan-out helpers (deleted with
#                       Internal/PolicyHydrateExtras.ps1).

function Invoke-PolicyHydrate
{
    param(
        [Parameter(Mandatory)] $Policies,
        [int]$TokenId = 0
    )

    $list = [System.Collections.Generic.List[object]]::new()
    foreach($p in @($Policies)) {
        if(-not $p) { continue }
        if(-not $p.PolicyType) { continue }
        if($p._IsFullObject) { continue }
        # Single-object endpoints return the object itself from the list call, so
        # the body is already complete and there is no per-id URL to fetch -
        # appending .id would duplicate the last path segment
        # (androidManagedStoreAccountEnterpriseSettings/androidManagedStore...)
        # and Graph 400s. Mark it full so GetFullObject()/Get() do not retry
        # forever, but keep it in the list so the nav-property and sub-resource
        # stages still run. This has to sit BEFORE the Id guard: Tenant Settings
        # (deviceManagement/settings) returns an EMPTY id, so an Id check first
        # would drop it here and leave it permanently un-hydrated.
        if($p.PolicyType.SingleObject -eq $true) {
            $p._IsFullObject = $true
            [void]$list.Add($p)
            continue
        }
        if(-not $p.Id) { continue }
        [void]$list.Add($p)
    }
    if($list.Count -eq 0) { return }

    # Cross-tenant Compare / Copy can hand us policies from different tokens in
    # one call. Body fetch + sub-resource fetch + Phase-A wrapper all reach
    # Graph with a single TokenId per chunk, so a mixed-token batch would
    # auth half the requests under the wrong tenant. Group by effective
    # _TokenId (fall back to the caller-supplied $TokenId) and dispatch one
    # per-token hydrate pass; the common single-token case takes the fast
    # path with no extra dictionary work.
    $byToken = $null
    $firstEff = $null
    foreach($p in $list) {
        $eff = if($null -ne $p._TokenId) { [int]$p._TokenId } else { [int]$TokenId }
        if($null -eq $firstEff) { $firstEff = $eff; continue }
        if($eff -ne $firstEff) {
            # Plain hashtable, not [ordered]: keys are integer token ids and an ordered
            # dictionary treats an integer indexer as a positional index, so $byToken[7]
            # would throw instead of addressing the token-7 bucket. Dispatch order does
            # not matter here.
            if($null -eq $byToken) { $byToken = @{} }
        }
    }
    if($null -ne $byToken) {
        foreach($p in $list) {
            $eff = if($null -ne $p._TokenId) { [int]$p._TokenId } else { [int]$TokenId }
            if(-not $byToken.Contains($eff)) {
                $byToken[$eff] = [System.Collections.Generic.List[object]]::new()
            }
            [void]$byToken[$eff].Add($p)
        }
        foreach($key in $byToken.Keys) {
            Invoke-PolicyHydrate -Policies $byToken[$key].ToArray() -TokenId $key
        }
        return
    }

    # Every policy left here shares one effective token (the multi-token case was split
    # and recursed above). Adopt it so the batch body fetch and the sub-resource fetch
    # authenticate under the policies' own token even when the caller passed the default
    # (0) -TokenId. Invoke-PolicyHydrateBodySingle already re-derives per policy, but the
    # batch and sub-resource helpers use $TokenId verbatim.
    if($null -ne $firstEff) { $TokenId = [int]$firstEff }

    if($list.Count -eq 1) {
        Invoke-PolicyHydrateBodySingle -Policy $list[0] -TokenId $TokenId
    }
    else {
        Invoke-PolicyHydrateBodyBatch -Policies $list -TokenId $TokenId
    }

    $hydrated = [System.Collections.Generic.List[object]]::new()
    foreach($p in $list) {
        if(-not $p._IsFullObject) { continue }
        if($p.PolicyType.NavigationProperties -ne $true) {
            Add-GraphNavigationProperties $p
        }
        [void]$hydrated.Add($p)
    }
    if($hydrated.Count -eq 0) { return }

    Invoke-PolicySubresourceFetch -Policies $hydrated -TokenId $TokenId
}

function Invoke-PolicyHydrateBodySingle
{
    param(
        [Parameter(Mandatory)]$Policy,
        [int]$TokenId = 0
    )

    $url = Get-PolicyHydrateBodyUrl -Policy $Policy
    if(-not $url) { return }

    $effTokenId = if($null -ne $Policy._TokenId) { $Policy._TokenId } else { $TokenId }

    $body = Invoke-MSGraphAPI -Url $url -TokenId $effTokenId
    if($body) {
        Merge-PolicyHydrateBody -Policy $Policy -Body $body
        $Policy._IsFullObject = $true
    }
    else {
        Write-Warning "Failed to get full object for $($Policy.Name)"
    }
}

# Merge the per-id body response into the existing JsonObject instead of
# replacing it wholesale. Properties present in $Body win; properties only on
# the existing JsonObject (typically `assignments` placed there by the list-
# stage $expand or by Add-GraphPolicyAssignments) are preserved.
#
# Without this merge, a body URL that doesn't include `$expand=assignments`
# silently drops the assignment rows the caller already populated — see the
# regression spotted on Manged App Test.json during the 2026-06-06 bulk-
# export diff session.
function Merge-PolicyHydrateBody
{
    param(
        [Parameter(Mandatory)]$Policy,
        [Parameter(Mandatory)]$Body
    )

    if($null -eq $Policy.JsonObject) {
        $Policy.JsonObject = $Body
        return
    }

    foreach($prop in $Body.PSObject.Properties) {
        if($Policy.JsonObject.PSObject.Properties[$prop.Name]) {
            $Policy.JsonObject.($prop.Name) = $prop.Value
        }
        else {
            Add-Member -InputObject $Policy.JsonObject -MemberType NoteProperty -Name $prop.Name -Value $prop.Value -Force
        }
    }
}

function Invoke-PolicyHydrateBodyBatch
{
    param(
        [Parameter(Mandatory)] $Policies,
        [int]$TokenId = 0
    )

    $batchObjects = [System.Collections.Generic.List[PSCustomObject]]::new()
    $policyById   = @{}
    $idx          = 0
    foreach($p in $Policies) {
        $url = Get-PolicyHydrateBodyUrl -Policy $p
        if(-not $url) { continue }
        $idx++
        $reqId = "hyd_$idx"
        [void]$batchObjects.Add([PSCustomObject]@{
            id      = $reqId
            method  = "GET"
            url     = $url.TrimStart('/')
            headers = @{ "Accept" = "application/json;odata.metadata=$($p.PolicyType.ODataMetadata)" }
        })
        $policyById[$reqId] = $p
    }
    if($batchObjects.Count -eq 0) { return }

    $throttle = 4
    try {
        $cfg = Get-SettingValue "ParallelBatchThrottle"
        if($cfg) { $throttle = [int]$cfg }
    } catch {}
    if($throttle -lt 1) { $throttle = 1 }
    $chunkSize = $throttle * 20

    $total     = $batchObjects.Count
    $processed = 0
    $batchNum  = 0
    $totalBatches = [Math]::Ceiling($total / [double]$chunkSize)

    Write-Log "Policy hydrate: fetching $total full policy object(s) via parallel batch (chunk=$chunkSize)"

    while($processed -lt $total) {
        $take  = [Math]::Min($chunkSize, $total - $processed)
        $chunk = [System.Collections.Generic.List[PSCustomObject]]::new($batchObjects.GetRange($processed, $take))
        $batchNum++

        Write-Status -Detail ("Fetching item {0} of {1} · batch {2} of {3}" -f ($processed + $take), $total, $batchNum, $totalBatches) -SkipLog -Force

        $results = Invoke-GraphBatchRequest -BatchObjects $chunk -BatchType "Hydrate:Body" -TokenId $TokenId

        foreach($r in $results) {
            $policy = $policyById["$($r.Id)"]
            if(-not $policy -or -not $r.body) { continue }
            Merge-PolicyHydrateBody -Policy $policy -Body $r.body
            $policy._IsFullObject = $true
        }

        $processed += $take
    }
}

# Resolve the Graph URL for hydrating a policy's main body. Most types use
# IntunePolicyBase.GetObjectURL() unchanged. AppConfigurationManagedAppObject
# / AppProtectionPolicyObject route through deviceAppManagement/<_objectClass>
# instead of their type's _API.
#
# The expand list for these wrappers is hardcoded per _objectClass because the
# list-stage JsonObject for these endpoints does NOT include the
# `apps@odata.navigationLink` / `settings@odata.navigationLink` / similar
# hints that IntunePolicyBase.GetObjectURL relies on. Without the hints, the
# generic expand-builder produces no $expand clause and the per-id GET
# returns a body missing the embedded collections — silently dropping
# $.apps and $.settings on every targetedManagedAppConfiguration export.
function Get-PolicyHydrateBodyUrl
{
    param([Parameter(Mandatory)]$Policy)

    if(-not $Policy.Id) { return $null }

    # A single-object endpoint IS the object. Its `id` can equal the last path
    # segment (androidManagedStoreAccountEnterpriseSettings), so the generic
    # "$API/$id" build would request …/x/x and fail. Nothing to fetch here.
    if($Policy.PolicyType.SingleObject -eq $true) { return $null }

    $typeName = $Policy.GetType().Name
    if($typeName -in @('AppConfigurationManagedAppObject', 'AppProtectionPolicyObject')) {
        if($Policy._objectClass) {
            $expand = $null
            switch($Policy._objectClass) {
                'windowsInformationProtectionPolicies' {
                    $expand = "?`$expand=protectedAppLockerFiles,exemptAppLockerFiles"
                }
                'targetedManagedAppConfigurations' {
                    # apps = the policy's app references; settings = the configured
                    # key/value pairs; assignments = the target groups. All three are
                    # missing from a plain per-id GET (list-stage nav-link hints absent).
                    # assignments MUST be expanded here: this type sets
                    # _ExpandAssignmentsList=$false, so the list stage never fetched
                    # assignments and Merge-PolicyHydrateBody has nothing to preserve —
                    # without this expand the documented policy shows no assignments.
                    # $expand=assignments on the body URL is the supported path for these
                    # managed-app types (their /assignments sub-resource returns 400).
                    $expand = "?`$expand=apps,settings,assignments"
                }
                default {
                    $url = ([IntunePolicyBase]$Policy).GetObjectURL()
                    $parts = $url.Split('?')
                    if($parts.Length -gt 1) { $expand = '?' + $parts[1] }
                }
            }
            return "deviceAppManagement/$($Policy._objectClass)/$($Policy.Id)$expand"
        }
    }

    return ([IntunePolicyBase]$Policy).GetObjectURL()
}
