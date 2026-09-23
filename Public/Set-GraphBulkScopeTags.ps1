function Set-GraphBulkScopeTags
{
    <#
    .SYNOPSIS
        Bulk-update Intune policy scope tag assignments.

    .DESCRIPTION
        Public, UI-independent driver for the Bulk Scope Tags tool. Lists policies
        for the requested type(s) / group(s), applies the chosen action (Add /
        Replace / Remove), optionally strips orphan tag ids (ids that no longer
        resolve to a real scope tag in the tenant), and PATCHes the new
        roleScopeTagIds value back to each policy.

        PATCH calls are batched 20 at a time via Invoke-GraphBatchRequest, so a
        full-tenant update is one round-trip per 20 policies rather than one per
        policy.

    .PARAMETER ScopeTagSettings
        An [IntuneManagerScopeTagSettings] instance. The UI binds to one; callers
        can construct one directly.

    .PARAMETER Action
        Overrides ScopeTagSettings.Action when supplied.

    .PARAMETER ScopeTagIds
        Overrides ScopeTagSettings.ScopeTagIds when supplied.

    .PARAMETER Filter
        Name filter (literal substring, case-insensitive) matched against each
        policy's Name. Empty = no filter. Overrides ScopeTagSettings.Filter when
        supplied.

    .PARAMETER CleanupOrphans
        Strip tag ids that don't resolve to a real scope tag. Overrides
        ScopeTagSettings.CleanupOrphans when supplied.

    .PARAMETER PolicyType
        Restrict the run to these PolicyType IDs.

    .PARAMETER PolicyGroup
        Restrict the run to these PolicyGroup IDs (expanded to their member types).

    .PARAMETER TokenId
        Authentication token id. Defaults to the current default token.

    .EXAMPLE
        $s = [IntuneManagerScopeTagSettings]::new()
        $s.Action      = "Add"
        $s.ScopeTagIds = @("3","4")
        Set-GraphBulkScopeTags -ScopeTagSettings $s -PolicyGroup DeviceConfiguration

    .EXAMPLE
        # Pure orphan cleanup across every type that supports scope tags.
        Set-GraphBulkScopeTags -CleanupOrphans

    .OUTPUTS
        PSCustomObject with summary statistics
        (Types, PoliciesScanned, PoliciesMatched, PoliciesUpdated, PoliciesSkipped, PoliciesFailed, Duration).
    #>
    [CmdletBinding()]
    param(
        [IntuneManagerScopeTagSettings]
        $ScopeTagSettings,

        [ValidateSet("Add","Replace","Remove")]
        [string]
        $Action,

        [string[]]
        $ScopeTagIds,

        [string]
        $Filter,

        [Nullable[bool]]
        $CleanupOrphans,

        [string[]]
        $PolicyType,

        [string[]]
        $PolicyGroup,

        [Int]
        $TokenId = (Get-DefaultTokenId)
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    if(-not $ScopeTagSettings) { $ScopeTagSettings = [IntuneManagerScopeTagSettings]::new() }

    if($PSBoundParameters.ContainsKey('Action'))         { $ScopeTagSettings.Action         = $Action }
    if($PSBoundParameters.ContainsKey('ScopeTagIds'))    { $ScopeTagSettings.ScopeTagIds    = @($ScopeTagIds) }
    if($PSBoundParameters.ContainsKey('Filter'))         { $ScopeTagSettings.Filter         = $Filter }
    if($PSBoundParameters.ContainsKey('CleanupOrphans')) { $ScopeTagSettings.CleanupOrphans = [bool]$CleanupOrphans }

    if($ScopeTagSettings.Action -notin @("Add","Replace","Remove")) {
        throw "Invalid Action '$($ScopeTagSettings.Action)'. Expected Add, Replace, or Remove."
    }

    # ---- 1. Resolve target policy types ----
    # Only types that actually advertise a ScopeTagProperty are eligible — the rest
    # don't support roleScopeTagIds and the PATCH would 400.
    $selection = Resolve-IntuneTargetSelectors -PolicyType $PolicyType -PolicyGroup $PolicyGroup -Caller 'Set-GraphBulkScopeTags'
    $unknownSelectors = $selection.Unknown   # ids from -PolicyType/-PolicyGroup that matched nothing
    $targetTypes = [System.Collections.Generic.List[object]]::new()
    $targetTypes.AddRange([object[]]$selection.Types)

    if(-not $PolicyType -and -not $PolicyGroup) {
        foreach($grp in $script:IntuneGroups) {
            if(-not $grp.Title) { continue }
            foreach($pt in $grp.PolicyTypes) { [void]$targetTypes.Add($pt) }
        }
    }

    $targetTypes = @($targetTypes |
        Where-Object { $_.ScopeTagProperty } |
        Sort-Object -Property Id -Unique)

    if($targetTypes.Count -eq 0) {
        Write-Log "Set-GraphBulkScopeTags: no policy types with ScopeTagProperty selected" 2
        return [PSCustomObject]@{
            Types            = 0
            PoliciesScanned  = 0
            PoliciesMatched  = 0
            PoliciesUpdated  = 0
            PoliciesSkipped  = 0
            PoliciesFailed   = 0
            UnknownSelectors = $unknownSelectors
            Duration         = $stopwatch.Elapsed
        }
    }

    # ---- 2. Pre-compile name filter ----
    # A literal substring, like 3.x and every other bulk driver. This command
    # writes scope tags, so '[Test]' must mean the text [Test] and never a
    # character class that matches most of the tenant.
    $filterRegex = $null
    if($ScopeTagSettings.Filter) {
        $filterRegex = [Regex]::new([Regex]::Escape($ScopeTagSettings.Filter), 'IgnoreCase')
    }

    # ---- 3. Load current scope tag catalogue (for orphan detection) ----
    # Always loaded — even when CleanupOrphans is off — because the summary log
    # and downstream UI may report orphans encountered.
    # Local effective copy so a tag-load failure doesn't mutate the caller's
    # settings instance (the UI re-uses the same object across runs).
    $effectiveCleanup = [bool]$ScopeTagSettings.CleanupOrphans
    $validTagIds = [System.Collections.Generic.HashSet[string]]::new()
    try {
        $tags = @(Get-GraphPolicies -PolicyType "ScopeTags" -TokenId $TokenId -ErrorAction Stop)
        foreach($t in $tags) {
            if($null -ne $t.Id) { [void]$validTagIds.Add([string]$t.Id) }
        }
        # The synthetic "Default" tag (Id 0) is always treated as valid even when
        # not surfaced by the live list — Intune requires it on most policies.
        [void]$validTagIds.Add("0")
    }
    catch {
        Write-Log "Set-GraphBulkScopeTags: failed to load scope tag catalogue - orphan cleanup disabled for this run. $($_.Exception.Message)" 2
        $effectiveCleanup = $false
    }

    # Selected tag ids → string set for fast membership tests in Add/Remove logic.
    $selectedIds = [System.Collections.Generic.HashSet[string]]::new()
    foreach($id in @($ScopeTagSettings.ScopeTagIds)) {
        if($null -ne $id -and "$id".Length -gt 0) { [void]$selectedIds.Add([string]$id) }
    }

    if($selectedIds.Count -eq 0) {
        if(-not ($ScopeTagSettings.Action -eq "Add" -and $effectiveCleanup)) {
            throw "No scope tag IDs were selected. Only Add with CleanupOrphans can run without selected tags."
        }
    }

    Write-Log "Set-GraphBulkScopeTags: action=$($ScopeTagSettings.Action), tags=$($selectedIds.Count), cleanup=$effectiveCleanup, types=$($targetTypes.Count)"

    # ---- 4. List policies for every selected type ----
    # Two-line status: action context on the primary line stays pinned across
    # the listing and the per-policy PATCH phases; sub-step lives on the detail line.
    Write-Status `
        -Text   ("Bulk scope tags - {0}" -f $ScopeTagSettings.Action) `
        -Detail ("Listing policies for {0} policy type(s)" -f $targetTypes.Count)
    $allPolicies = @()
    try {
        $allPolicies = @(Get-GraphPolicies -PolicyType @($targetTypes | ForEach-Object { $_.Id }) -TokenId $TokenId -ErrorAction Stop)
    }
    catch {
        Write-LogError "Set-GraphBulkScopeTags: failed to list policies" $_.Exception
    }

    # ---- 5. Build the PATCH batch ----
    $scanned   = 0
    $matched   = 0
    $skipped   = 0
    $batch     = [System.Collections.Generic.List[PSCustomObject]]::new()
    $byBatchId = @{}   # batch id -> PSCustomObject @{Policy=...; NewIds=...}

    foreach($p in $allPolicies)
    {
        $scanned++
        if(-not $p -or -not $p.PolicyType -or -not $p.PolicyType.ScopeTagProperty) { continue }

        if($filterRegex -and -not $filterRegex.IsMatch([string]$p.Name)) { continue }

        $matched++

        $prop = [string]$p.PolicyType.ScopeTagProperty
        if(-not $p.Object -or -not $p.Object.PSObject.Properties[$prop]) {
            $skipped++
            Write-LogDebug "Set-GraphBulkScopeTags: '$($p.Name)' [$($p.PolicyType.Title)] does not expose '$prop' in Graph response - skipped"
            continue
        }

        # Current ids — defensive: property may not exist on the JSON yet, or
        # may be empty. Stringify everything for consistent set ops.
        $currentIds = [System.Collections.Generic.HashSet[string]]::new()
        foreach($id in @($p.Object.$prop)) {
            if($null -ne $id) { [void]$currentIds.Add([string]$id) }
        }

        $newIds = [System.Collections.Generic.HashSet[string]]::new($currentIds)
        switch ($ScopeTagSettings.Action) {
            "Add"     { foreach($id in $selectedIds) { [void]$newIds.Add($id) } }
            "Remove"  { foreach($id in $selectedIds) { [void]$newIds.Remove($id) } }
            "Replace" { $newIds = [System.Collections.Generic.HashSet[string]]::new($selectedIds) }
        }

        if($effectiveCleanup) {
            $stale = @($newIds | Where-Object { -not $validTagIds.Contains($_) })
            foreach($id in $stale) { [void]$newIds.Remove($id) }
        }

        # Stable comparison (order-independent): same membership → no PATCH.
        if($newIds.Count -eq $currentIds.Count -and
           (@($newIds | Where-Object { -not $currentIds.Contains($_) }).Count -eq 0)) {
            $skipped++
            continue
        }

        $ht = [ordered]@{}
        $ht[$prop] = @($newIds)
        # Some types require @odata.type on PATCH (subtype-discriminated entities).
        # Mirror the Details-view behaviour: include it when the source object has it.
        if($p.Object -and $p.Object.'@odata.type') {
            $ht['@odata.type'] = $p.Object.'@odata.type'
        }

        $bodyObj = [PSCustomObject]$ht
        $bodyJson = $bodyObj | ConvertTo-Json -Depth 20 -Compress

        $batchId = [string]$batch.Count

        # Graph $batch sub-requests require the body to be an OBJECT (not a string).
        # Invoke-GraphBatchRequest currently passes BatchObjects.body through as-is,
        # so we attach the parsed object directly.
        $batchObj = [PSCustomObject]@{
            id      = $batchId
            method  = "PATCH"
            url     = Get-GraphBulkScopeTagPatchUrl -Policy $p
            body    = ($bodyJson | ConvertFrom-Json)
            headers = @{
                "Content-Type" = "application/json"
                "If-Match"     = "*"
            }
        }
        [void]$batch.Add($batchObj)
        $byBatchId[$batchId] = [PSCustomObject]@{ Policy = $p; NewIds = @($newIds) }
    }

    # ---- 6. Dispatch the batch ----
    $updated = 0
    $failed  = 0

    if($batch.Count -gt 0) {
        Write-Status -Detail ("Updating scope tags on {0} policy(ies)" -f $batch.Count) -SkipLog
        $results = @(Invoke-GraphBatchRequest -BatchObjects $batch -BatchType "BulkScopeTags" -TokenId $TokenId -IncludedFailed -SkipWarnings)
        if($results.Count -eq 0) {
            $failed += $batch.Count
            Write-Log "Set-GraphBulkScopeTags: batch returned no responses; treating all $($batch.Count) PATCH request(s) as failed" 3
        }
        else {
            foreach($r in $results) {
                $entry = $byBatchId["$($r.Id)"]
                if(-not $entry) { continue }
                $statusCode = 0
                try { $statusCode = [int]$r.status } catch { }

                if($statusCode -ge 200 -and $statusCode -lt 300) {
                    $updated++
                    # Mirror to in-memory PSCustomObject so a subsequent UI refresh
                    # shows the new tag list without an extra GET.
                    try {
                        $prop = [string]$entry.Policy.PolicyType.ScopeTagProperty
                        if($entry.Policy.Object.PSObject.Properties[$prop]) {
                            $entry.Policy.Object.$prop = $entry.NewIds
                        }
                        else {
                            $entry.Policy.Object | Add-Member -MemberType NoteProperty -Name $prop -Value $entry.NewIds -Force
                        }
                        $entry.Policy._ScopeTags       = $null
                        $entry.Policy._ScopeTagsString = $null
                    }
                    catch { Write-LogDebug "Set-GraphBulkScopeTags: in-memory mirror failed for $($entry.Policy.Name): $($_.Exception.Message)" }
                }
                else {
                    $failed++
                    $msg = ""
                    if($r.body -and $r.body.error -and $r.body.error.message) { $msg = $r.body.error.message }
                    Write-Log "Set-GraphBulkScopeTags: PATCH failed for '$($entry.Policy.Name)' [$($entry.Policy.PolicyType.Title)] - $statusCode $msg" 3
                }
            }

            $seenIds = [System.Collections.Generic.HashSet[string]]::new()
            foreach($r in $results) { [void]$seenIds.Add([string]$r.Id) }
            foreach($pending in $byBatchId.Keys) {
                if(-not $seenIds.Contains([string]$pending)) {
                    $failed++
                    $entry = $byBatchId[$pending]
                    Write-Log "Set-GraphBulkScopeTags: no batch response for '$($entry.Policy.Name)' [$($entry.Policy.PolicyType.Title)]" 3
                }
            }
        }
    }

    Write-Status ""

    [PSCustomObject]@{
        Types            = $targetTypes.Count
        PoliciesScanned  = $scanned
        PoliciesMatched  = $matched
        PoliciesUpdated  = $updated
        PoliciesSkipped  = $skipped
        PoliciesFailed   = $failed
        UnknownSelectors = $unknownSelectors
        Duration         = $stopwatch.Elapsed
    }
}
