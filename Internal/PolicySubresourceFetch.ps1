#region Policy sub-resource orchestrator
#
# Drives the per-class GetSubResourceBatchRequests / ApplySubResourceBatchResult
# contract defined on IntunePolicyBase. The contract lets a policy class declare
# the extra Graph GETs it needs to be fully hydrated beyond the main object body
# (relationships, script content, etc.) so the hydrator pipeline can fan them
# out via $batch in parallel instead of paying one round-trip per policy.
#
# Phase loop:
#   1. Phase 1 requests come from policy.GetSubResourceBatchRequests(1).
#   2. Identical URLs across policies are coalesced into ONE batch sub-request
#      and the single response is fanned out to every (policy, key) that asked
#      for it (e.g. two AppConfig policies referencing the same mobileApp fetch
#      it once). Method + Headers must also match for coalescing.
#   3. Each phase batch is dispatched via Invoke-GraphBatchRequest.
#   4. Each response is routed back to its owning policy(ies) via
#      ApplySubResourceBatchResult($phase, $key, $body), which returns any
#      follow-up requests for phase+1.
#   5. Loop continues until no policy returns more requests, or until $maxPhases.
#   6. After all phases, FinalizeSubResources() runs once per opted-in policy so
#      classes can assemble derived properties from the applied responses.
#
# Opted-in policies are flagged via $_HasSubResourceBatch on the class. Default
# is $false, so unchanged classes stay on the inline path until they migrate.

function Invoke-PolicySubresourceFetch
{
    param(
        [Parameter(Mandatory)] $Policies,
        [int] $TokenId = 0
    )

    if(-not $Policies) { return }
    $list = @($Policies)
    if($list.Count -eq 0) { return }

    $optIn = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach($p in $list) {
        if($null -eq $p) { continue }
        if($p._HasSubResourceBatch -ne $true) { continue }
        if(-not $p.Id) { continue }
        # Clear any per-phase state from a previous hydrate of the same wrapper
        # (e.g. UI list-row re-open). Without this, partial state from a failed
        # phase-2 in a prior run would silently leak into the next hydrate's
        # ApplySubResourceBatchResult lookups.
        if($p.PSObject.Properties['_SubResourceState']) {
            $p._SubResourceState = $null
        }
        [void]$optIn.Add($p)
    }
    if($optIn.Count -eq 0) { return }

    # Phase 1: ask every opted-in policy what it needs.
    $pending = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach($policy in $optIn) {
        try {
            $reqs = $policy.GetSubResourceBatchRequests(1)
        }
        catch {
            Write-LogError "Policy sub-resource: GetSubResourceBatchRequests on $($policy.GetType().Name) ($($policy.Name)) threw" $_.Exception
            continue
        }
        foreach($r in @($reqs)) {
            if($r -and $r.Url -and $r.Key) {
                [void]$pending.Add([PSCustomObject]@{ Policy = $policy; Request = $r })
            }
        }
    }

    if($pending.Count -eq 0) {
        foreach($p in $optIn) { $p._IsFullObject = $true }
        return
    }

    Write-Log "Policy sub-resource: $($pending.Count) phase-1 request(s) across $($optIn.Count) policy(ies)"

    $maxPhases = 5
    $globalIdx = 0

    for($phase = 1; $phase -le $maxPhases -and $pending.Count -gt 0; $phase++) {

        $batchObjects = [System.Collections.Generic.List[PSCustomObject]]::new()
        # reqId -> list of { Policy; Key } so one coalesced response fans out to
        # every requester. urlKey -> reqId so identical requests reuse one slot.
        $routing      = @{}
        $reqIdByUrl   = @{}

        foreach($entry in $pending) {
            $req     = $entry.Request
            $headers = if($req.Headers) { $req.Headers } else { @{ Accept = 'application/json;odata.metadata=minimal' } }
            $method  = if($req.Method)  { $req.Method }  else { 'GET' }
            $url     = ([string]$req.Url).TrimStart('/')

            # Coalesce identical (method + url + headers) requests so a shared
            # sub-resource is fetched once. Headers are serialized into the key
            # so requests differing only by Accept aren't wrongly merged.
            $hdrKey  = ($headers.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ';'
            $urlKey  = "$method`n$url`n$hdrKey"

            if($reqIdByUrl.ContainsKey($urlKey)) {
                $reqId = $reqIdByUrl[$urlKey]
            }
            else {
                $globalIdx++
                $reqId = "psrf_${phase}_${globalIdx}"
                $reqIdByUrl[$urlKey] = $reqId
                [void]$batchObjects.Add([PSCustomObject]@{
                    id      = $reqId
                    method  = $method
                    url     = $url
                    headers = $headers
                })
                $routing[$reqId] = [System.Collections.Generic.List[PSCustomObject]]::new()
            }
            [void]$routing[$reqId].Add([PSCustomObject]@{ Policy = $entry.Policy; Key = [string]$req.Key })
        }

        $batchType = "Hydrate:Subresource:Phase$phase"
        $results = @(Invoke-GraphBatchRequest -BatchObjects $batchObjects -BatchType $batchType -TokenId $TokenId -SkipWarnings -IncludedFailed)

        $next = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach($result in $results) {
            $routes = $routing["$($result.Id)"]
            if(-not $routes) { continue }

            $body = $null
            if($result.Status -ge 200 -and $result.Status -lt 300) {
                $body = $result.body
            }

            foreach($route in $routes) {
                if($null -eq $body) {
                    Write-Log "Policy sub-resource: $($route.Policy.GetType().Name) ($($route.Policy.Name)) phase $phase key '$($route.Key)' returned HTTP $($result.Status)" 2
                }
                try {
                    $followups = $route.Policy.ApplySubResourceBatchResult($phase, $route.Key, $body)
                }
                catch {
                    Write-LogError "Policy sub-resource: ApplySubResourceBatchResult on $($route.Policy.GetType().Name) (phase=$phase, key=$($route.Key)) threw" $_.Exception
                    continue
                }

                foreach($f in @($followups)) {
                    if($f -and $f.Url -and $f.Key) {
                        [void]$next.Add([PSCustomObject]@{ Policy = $route.Policy; Request = $f })
                    }
                }
            }
        }

        $pending = $next
    }

    if($pending.Count -gt 0) {
        # Hitting the phase cap means a class's ApplySubResourceBatchResult keeps
        # returning follow-ups indefinitely. Always a bug in the class — emit at
        # error level so it surfaces in the log filter, not just verbose noise.
        Write-Log "Policy sub-resource: phase cap ($maxPhases) reached with $($pending.Count) request(s) still pending - runaway in a class implementation; data is incomplete" 3
    }

    # Finalize pass: let each opted-in policy assemble derived properties from
    # the responses already applied (e.g. AppConfiguration's #CustomRefTargetedApps).
    # Every IntunePolicyBase inherits FinalizeSubResources; the method-existence
    # guard keeps the orchestrator tolerant of stand-in objects that only
    # implement the request/apply half of the contract.
    foreach($p in $optIn) {
        if(-not $p.PSObject.Methods['FinalizeSubResources']) { continue }
        try {
            $p.FinalizeSubResources()
        }
        catch {
            Write-LogError "Policy sub-resource: FinalizeSubResources on $($p.GetType().Name) ($($p.Name)) threw" $_.Exception
        }
    }

    foreach($p in $optIn) {
        $p._IsFullObject = $true
    }
}

#endregion
