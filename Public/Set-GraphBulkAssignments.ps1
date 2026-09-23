function Set-GraphBulkAssignments
{
    <#
    .SYNOPSIS
        Bulk-update Intune policy assignments.

    .DESCRIPTION
        Public, UI-independent driver for the Bulk Assignments tool. Lists
        policies for the requested type(s) / group(s), applies the chosen
        Action (Add / Replace / Remove), and POSTs the new assignments list
        to the policy's /assign endpoint (the canonical "replace all
        assignments" operation Intune uses internally).

        Supported shapes:
          * Simple — `{target}` only. Default for most types
            (DeviceConfiguration, Compliance, EndpointSecurity, Scripts,
            EnrollmentConfiguration, Autopilot, PolicySet, etc.).
          * App — `{target, intent}`. Used by types whose AssignmentsType
            is `mobileAppAssignments`. Per-platform `settings` is NOT
            populated yet (Phase 3); Graph defaults are accepted.

        Targets supported in this phase:
          * Group, exclusion-group, all-devices, all-users
          * Optional assignment filter (include/exclude) on group targets

        Health-script assignments (`runSchedule` + `runRemediationScript`)
        are filtered out — they need a per-type schedule editor that lands
        in a later phase.

        The /assign endpoint REPLACES all assignments for the policy, so
        Add and Remove modes GET each policy's current assignments first
        (via Get-GraphPolicies -IncludeAssignments), compute the new list,
        and POST the merged result. Replace mode skips the merge step.

    .PARAMETER AssignmentSettings
        An [IntuneManagerAssignmentSettings] instance. The UI binds to one;
        callers can construct one directly.

    .PARAMETER Action
        Overrides AssignmentSettings.Action when supplied.

    .PARAMETER Assignments
        Overrides AssignmentSettings.Assignments when supplied.

    .PARAMETER Filter
        Name filter (literal substring, case-insensitive) matched against each
        policy's Name. Empty = no filter. Overrides AssignmentSettings.Filter
        when supplied.

    .PARAMETER PolicyType
        Restrict the run to these PolicyType IDs.

    .PARAMETER PolicyGroup
        Restrict the run to these PolicyGroup IDs (expanded to their member types).

    .PARAMETER TokenId
        Authentication token id. Defaults to the current default token.

    .EXAMPLE
        $s = [IntuneManagerAssignmentSettings]::new()
        $s.Action      = 'Add'
        $s.Assignments = @([PSCustomObject]@{
            TargetType = 'groupAssignmentTarget'
            GroupId    = '<aad-group-id>'
            GroupName  = 'All Helpdesk Devices'
        })
        Set-GraphBulkAssignments -AssignmentSettings $s -PolicyGroup DeviceConfiguration

    .OUTPUTS
        PSCustomObject @{
            Types, PoliciesScanned, PoliciesMatched, PoliciesUpdated,
            PoliciesSkipped, PoliciesFailed, PoliciesUnsupported, Duration
        }
    #>
    [CmdletBinding()]
    param(
        [IntuneManagerAssignmentSettings]
        $AssignmentSettings,

        [ValidateSet("Add","Replace","Remove")]
        [string]
        $Action,

        [PSCustomObject[]]
        $Assignments,

        [string]
        $Filter,

        [string[]]
        $PolicyType,

        [string[]]
        $PolicyGroup,

        [Int]
        $TokenId = (Get-DefaultTokenId)
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    if(-not $AssignmentSettings) { $AssignmentSettings = [IntuneManagerAssignmentSettings]::new() }

    if($PSBoundParameters.ContainsKey('Action'))      { $AssignmentSettings.Action      = $Action }
    if($PSBoundParameters.ContainsKey('Assignments')) { $AssignmentSettings.Assignments = @($Assignments) }
    if($PSBoundParameters.ContainsKey('Filter'))      { $AssignmentSettings.Filter      = $Filter }

    if($AssignmentSettings.Action -notin @("Add","Replace","Remove")) {
        throw "Invalid Action '$($AssignmentSettings.Action)'. Expected Add, Replace, or Remove."
    }

    $chosen = @($AssignmentSettings.Assignments | Where-Object { $_ })
    if($chosen.Count -eq 0) {
        throw "No assignments selected. Specify at least one target in AssignmentSettings.Assignments."
    }

    # Phase 1 supports only the simple target-only shape. Anything else (apps
    # with intent + settings, health scripts with runSchedule) would silently
    # drop user fields during the Graph round-trip, so reject up-front.
    $simpleTargetTypes = @($chosen | Where-Object {
        $_.TargetType -in @(
            "groupAssignmentTarget",
            "exclusionGroupAssignmentTarget",
            "allDevicesAssignmentTarget",
            "allLicensedUsersAssignmentTarget")
    })
    if($simpleTargetTypes.Count -ne $chosen.Count) {
        $unknown = @($chosen | Where-Object { $_ -notin $simpleTargetTypes } | ForEach-Object { $_.TargetType }) -join ', '
        throw "Unsupported assignment target type(s): $unknown. Phase 1 supports groupAssignmentTarget, exclusionGroupAssignmentTarget, allDevicesAssignmentTarget, allLicensedUsersAssignmentTarget."
    }

    foreach($a in $chosen) {
        if($a.TargetType -in @("groupAssignmentTarget","exclusionGroupAssignmentTarget") -and -not $a.GroupId) {
            throw "Assignment of type '$($a.TargetType)' is missing GroupId."
        }
    }

    # installIntent enum values accepted by Graph for mobileAppAssignment.
    # Non-app shapes ignore Intent — see Build-AssignmentBody below.
    $validIntents = @("available","notAvailable","required","uninstall","availableWithoutEnrollment")
    foreach($a in $chosen) {
        if($a.Intent -and $a.Intent -notin $validIntents) {
            throw "Invalid Intent '$($a.Intent)' for target '$($a.GroupName)'. Expected one of: $($validIntents -join ', ')."
        }
    }

    # ---- 1. Resolve target policy types ----
    # Filter to types that support the bulk assignment action shape. Some
    # Intune objects expose an `assignments` navigation property for export or
    # documentation, but do not support the replace-all /assign action used by
    # this tool.
    $selection = Resolve-IntuneTargetSelectors -PolicyType $PolicyType -PolicyGroup $PolicyGroup -Caller 'Set-GraphBulkAssignments'
    $unknownSelectors = $selection.Unknown   # ids from -PolicyType/-PolicyGroup that matched nothing
    $targetTypes = [System.Collections.Generic.List[object]]::new()
    $targetTypes.AddRange([object[]]$selection.Types)
    $unsupportedSelected = [System.Collections.Generic.List[object]]::new()

    if(-not $PolicyType -and -not $PolicyGroup) {
        foreach($grp in $script:IntuneGroups) {
            if(-not $grp.Title) { continue }
            foreach($pt in $grp.PolicyTypes) { [void]$targetTypes.Add($pt) }
        }
    }

    $targetTypes = @($targetTypes | Sort-Object -Property Id -Unique)

    $eligibleTypes = [System.Collections.Generic.List[object]]::new()
    foreach($pt in $targetTypes) {
        if(-not $pt.SupportsAssignments) { continue }
        if(-not $pt.AssignmentsType)     { continue }
        if(-not (Test-BulkAssignmentSupported $pt)) {
            [void]$unsupportedSelected.Add($pt)
            continue
        }
        [void]$eligibleTypes.Add($pt)
    }

    if($eligibleTypes.Count -eq 0) {
        Write-Log "Set-GraphBulkAssignments: no eligible policy types selected (need SupportsAssignments=true and default target-only shape)" 2
        return [PSCustomObject]@{
            Types               = 0
            PoliciesScanned     = 0
            PoliciesMatched     = 0
            PoliciesUpdated     = 0
            PoliciesSkipped     = 0
            PoliciesFailed      = 0
            PoliciesUnsupported = 0
            UnsupportedTypes    = @($unsupportedSelected | ForEach-Object { $_.Title })
            UnknownSelectors    = $unknownSelectors
            Duration            = $stopwatch.Elapsed
        }
    }

    # ---- 2. Pre-compile name filter ----
    # A literal substring, like 3.x and every other bulk driver. This command
    # writes assignments, so '[Test]' must mean the text [Test] and never a
    # character class that matches most of the tenant.
    $filterRegex = $null
    if($AssignmentSettings.Filter) {
        $filterRegex = [Regex]::new([Regex]::Escape($AssignmentSettings.Filter), 'IgnoreCase')
    }

    Write-Log "Set-GraphBulkAssignments: action=$($AssignmentSettings.Action), assignments=$($chosen.Count), eligible=$($eligibleTypes.Count), unsupported=$($unsupportedSelected.Count)"

    # Surface PolicyTypes whose assignment @odata.type the bulk tool can't
    # resolve. These will still PATCH (Graph sometimes accepts an entry
    # without an explicit @odata.type, sometimes 400s) but the user should
    # know — once per type per run, not once per policy. The fix is to set
    # _AssignmentObjectType on the PolicyType class (see IntunePolicyBase).
    foreach($pt in $eligibleTypes) {
        if(-not (Get-BulkAssignmentObjectType $pt)) {
            Write-Log "Set-GraphBulkAssignments: no assignment @odata.type resolved for PolicyType '$($pt.Id)' ($($pt.Title)). Set _AssignmentObjectType on the class or extend Get-BulkAssignmentObjectType." 2
        }
    }

    # ---- 3. List policies (with assignments) for every eligible type ----
    # Two-line status: action context pinned on the primary line for the whole
    # run; the detail line rotates between listing and the PATCH phase.
    Write-Status `
        -Text   ("Bulk assignments - {0}" -f $AssignmentSettings.Action) `
        -Detail ("Listing policies for {0} policy type(s)" -f $eligibleTypes.Count)
    $allPolicies = @()
    try {
        $allPolicies = @(Get-GraphPolicies -PolicyType @($eligibleTypes | ForEach-Object { $_.Id }) -TokenId $TokenId -IncludeAssignments -ErrorAction Stop)
    }
    catch {
        Write-LogError "Set-GraphBulkAssignments: failed to list policies" $_.Exception
    }

    # The cached policy list can carry STALE assignments from an earlier pass
    # in the same session (Add-GraphPolicyAssignments skips policies whose
    # assignments are already populated). The Add/Remove merge must see the
    # LIVE state - with stale data a Remove computes "nothing changed" and
    # silently no-ops. Force a refresh for every policy this run will touch.
    $refresh = @($allPolicies | Where-Object {
        $_ -and $_.PolicyType -and $_.Object -and
        (Test-BulkAssignmentSupported $_.PolicyType) -and
        (-not $filterRegex -or $filterRegex.IsMatch([string]$_.Name))
    })
    foreach($p in $refresh) { Remove-Property $p.Object 'assignments' }
    if($refresh.Count -gt 0) {
        Add-GraphPolicyAssignments -Policies $refresh -TokenId $TokenId
    }

    # ---- 4. Build the per-policy POST batch ----
    # Tuples carry both Target (final POST shape) and Intent (only set for
    # app-shape types). Tuple → assignment-object conversion happens once
    # right before serialisation so the shape branch lives in one place.
    $scanned = 0
    $matched = 0
    $skipped = 0
    $batch     = [System.Collections.Generic.List[PSCustomObject]]::new()
    $byBatchId = @{}

    foreach($p in $allPolicies)
    {
        $scanned++
        if(-not $p -or -not $p.PolicyType) { continue }
        if(-not (Test-BulkAssignmentSupported $p.PolicyType)) { continue }

        if($filterRegex -and -not $filterRegex.IsMatch([string]$p.Name)) { continue }
        $matched++

        $shape = Get-BulkAssignmentShape $p.PolicyType

        # For app shape: figure out which per-platform settings entry the
        # user's chosen list applies to THIS app (source) and which Graph
        # @odata.type to stamp on the resulting `settings` object (target).
        # Inheritance shortcut: win32CatalogApp reuses Win32 LOB settings
        # since the inheriting type has no extra fields of its own.
        $sourceSettingsType = $null
        $targetSettingsType = $null
        if($shape -eq "app" -and $p.Object -and $p.Object.'@odata.type') {
            $targetSettingsType = Get-AppSettingsTypeForPolicy $p.Object.'@odata.type'
            $sourceSettingsType = if($targetSettingsType -eq 'win32CatalogAppAssignmentSettings') {
                'win32LobAppAssignmentSettings'
            } else {
                $targetSettingsType
            }
        }

        # Project current assignments to (Target, Intent, Settings) tuples.
        # Settings on current assignments are already in Graph form (have
        # @odata.type), so pass through verbatim — they're only re-emitted
        # when the assignment survives Add/Remove.
        $current = @()
        if($p.Object -and $p.Object.PSObject.Properties['assignments']) {
            $current = @($p.Object.assignments | Where-Object { $_ -and $_.target })
        }
        $currentTuples = @()
        $currentKeys   = [System.Collections.Generic.HashSet[string]]::new()
        foreach($a in $current) {
            # Script shape: preserve runSchedule (PSCustomObject from Graph
            # with @odata.type already set) and runRemediationScript
            # ([Nullable[bool]]) so Add/Remove keeps them on surviving rows.
            $runSchedule        = $null
            $runRemediation     = $null
            if($shape -eq "script") {
                if($a.PSObject.Properties['runSchedule'] -and $a.runSchedule)         { $runSchedule = $a.runSchedule }
                if($a.PSObject.Properties['runRemediationScript'])                    { $runRemediation = [Nullable[bool]]$a.runRemediationScript }
            }
            $tuple = [PSCustomObject]@{
                Target         = ConvertTo-AssignmentTarget $a.target
                Intent         = if($shape -eq "app") { [string]$a.intent } else { $null }
                Settings       = if($shape -eq "app" -and $a.settings) { $a.settings } else { $null }
                RunSchedule    = $runSchedule
                RunRemediation = $runRemediation
            }
            $currentTuples += $tuple
            [void]$currentKeys.Add((Get-AssignmentSignature $tuple.Target $tuple.Intent))
        }

        # Build chosen tuples for THIS shape. Intent only on apps; Settings
        # only when (a) shape is app AND (b) the chosen descriptor has a
        # per-platform hashtable matching this app's settings type. Same
        # chosen list flows across mixed-type runs unchanged.
        $chosenTuples = @()
        $chosenKeys   = [System.Collections.Generic.HashSet[string]]::new()
        foreach($c in $chosen) {
            $settings       = $null
            $runSchedule    = $null
            $runRemediation = $null
            if($shape -eq "app" -and $sourceSettingsType -and $c.Settings -is [Hashtable] -and $c.Settings.ContainsKey($sourceSettingsType)) {
                $h = $c.Settings[$sourceSettingsType]
                if($h -is [Hashtable] -and $h.Count -gt 0) {
                    $settings = ConvertTo-AppSettingsObject -Hash $h -GraphType $targetSettingsType
                }
            }
            if($shape -eq "script" -and $c.Settings -is [Hashtable] -and $c.Settings.ContainsKey('deviceHealthScriptAssignment')) {
                $h = $c.Settings['deviceHealthScriptAssignment']
                if($h -is [Hashtable]) {
                    if($h.ContainsKey('runRemediationScript')) { $runRemediation = [Nullable[bool]]$h['runRemediationScript'] }
                    if($h.ContainsKey('scheduleType'))         { $runSchedule    = Build-HealthScriptSchedule -Spec $h }
                }
            }
            # Some app types cannot take an assignment filter - Graph answers
            # "The Assignment Filters are not supported for this app type" and the
            # whole policy failed. The portal simply does not offer a filter for
            # them, so assign the group as asked and drop the filter, with a log
            # line naming the policy. Live-verified for webApp (2026-09-13).
            $descriptorForTarget = $c
            if($shape -eq "app" -and $c.FilterId -and $p.Object -and
               ([string]$p.Object.'@odata.type') -in $script:BulkAssignmentAppTypesWithoutFilters) {
                Write-Log "Set-GraphBulkAssignments: '$($p.Name)' is a $($p.Object.'@odata.type' -replace '^#microsoft\.graph\.','') - assignment filters are not supported for this app type, assigning without the filter" 2
                $descriptorForTarget = [PSCustomObject]@{
                    TargetType = $c.TargetType
                    GroupId    = $c.GroupId
                    GroupName  = $c.GroupName
                }
            }
            $tuple = [PSCustomObject]@{
                Target         = ConvertTo-AssignmentTarget (Build-AssignmentTarget $descriptorForTarget)
                Intent         = if($shape -eq "app") { if($c.Intent) { [string]$c.Intent } else { "required" } } else { $null }
                Settings       = $settings
                RunSchedule    = $runSchedule
                RunRemediation = $runRemediation
            }
            $chosenTuples += $tuple
            [void]$chosenKeys.Add((Get-AssignmentSignature $tuple.Target $tuple.Intent))
        }

        # Compute new tuple set per Action.
        $newTuples = @()
        switch ($AssignmentSettings.Action) {
            "Replace" {
                $newTuples = $chosenTuples
            }
            "Add" {
                $seen = [System.Collections.Generic.HashSet[string]]::new()
                foreach($t in $currentTuples) {
                    if($seen.Add((Get-AssignmentSignature $t.Target $t.Intent))) { $newTuples += $t }
                }
                foreach($t in $chosenTuples) {
                    if($seen.Add((Get-AssignmentSignature $t.Target $t.Intent))) { $newTuples += $t }
                }
            }
            "Remove" {
                foreach($t in $currentTuples) {
                    $sig = Get-AssignmentSignature $t.Target $t.Intent
                    if($chosenKeys.Contains($sig)) { continue }
                    $newTuples += $t
                }
            }
        }

        # No-op if the resulting assignments are identical to current
        # (order-independent). Use the full tuple signature here, not just
        # target+intent, so Replace can update app settings and health-script
        # schedule/remediation fields for an existing target.
        $currentFullKeys = [System.Collections.Generic.HashSet[string]]::new()
        foreach($t in $currentTuples) { [void]$currentFullKeys.Add((Get-AssignmentFullSignature $t)) }
        $newFullKeys = [System.Collections.Generic.HashSet[string]]::new()
        foreach($t in $newTuples) { [void]$newFullKeys.Add((Get-AssignmentFullSignature $t)) }
        if($newFullKeys.Count -eq $currentFullKeys.Count -and
           (@($newFullKeys | Where-Object { -not $currentFullKeys.Contains($_) }).Count -eq 0)) {
            $skipped++
            # A Remove that finds nothing to remove is the case worth explaining:
            # either the target was never there, or the refresh read came back
            # empty and the assignment is about to be left behind.
            if($AssignmentSettings.Action -eq 'Remove') {
                Write-Log "Set-GraphBulkAssignments: '$($p.Name)' [$($p.PolicyType.Id)] - nothing to remove (current=$($currentTuples.Count), chosen=$($chosenTuples.Count))"
            }
            continue
        }

        # Tuples → POST objects (shape-aware).
        #   app    → {target, intent, settings?}
        #   script → {target, runRemediationScript?, runSchedule?}
        #   simple → {target}
        # All optional fields are omitted entirely when null/absent so Graph
        # keeps its defaults rather than seeing $null / empty objects.
        $assignmentObjectType = Get-BulkAssignmentObjectType $p.PolicyType
        $newAssignments = foreach($t in $newTuples) {
            if($shape -eq "app") {
                $obj = [ordered]@{}
                if($assignmentObjectType) { $obj['@odata.type'] = $assignmentObjectType }
                $obj.target = $t.Target
                $obj.intent = (?? $t.Intent "required")
                if($t.Settings) { $obj.settings = $t.Settings }
                [PSCustomObject]$obj
            }
            elseif($shape -eq "script") {
                $obj = [ordered]@{}
                if($assignmentObjectType) { $obj['@odata.type'] = $assignmentObjectType }
                $obj.target = $t.Target
                if($null -ne $t.RunRemediation) { $obj.runRemediationScript = [bool]$t.RunRemediation }
                if($t.RunSchedule)              { $obj.runSchedule          = $t.RunSchedule }
                [PSCustomObject]$obj
            }
            else {
                $obj = [ordered]@{}
                if($assignmentObjectType) { $obj['@odata.type'] = $assignmentObjectType }
                $obj.target = $t.Target
                [PSCustomObject]$obj
            }
        }

        $body = [PSCustomObject]@{
            $p.PolicyType.AssignmentsType = @($newAssignments)
        }
        $bodyJson = $body | ConvertTo-Json -Depth 20 -Compress

        $batchId = [string]$batch.Count
        [void]$batch.Add([PSCustomObject]@{
            id      = $batchId
            method  = "POST"
            # AssignAction is "assign" for almost every type; policySets have
            # no /assign segment and take the replacement list via /update.
            url     = "$($p.PolicyType.API)/$($p.Id)/$($p.PolicyType.AssignAction)"
            body    = ($bodyJson | ConvertFrom-Json)
            headers = @{ "Content-Type" = "application/json" }
        })
        $byBatchId[$batchId] = [PSCustomObject]@{ Policy = $p; NewAssignments = @($newAssignments) }
    }

    # ---- 6. Dispatch the batch ----
    $updated = 0
    $failed  = 0
    if($batch.Count -gt 0) {
        Write-Status -Detail ("Updating assignments on {0} policy(ies)" -f $batch.Count) -SkipLog
        $results = @(Invoke-GraphBatchRequest -BatchObjects $batch -BatchType "BulkAssignments" -TokenId $TokenId -IncludedFailed -SkipWarnings)

        if($results.Count -eq 0) {
            $failed += $batch.Count
            Write-Log "Set-GraphBulkAssignments: batch returned no responses; treating all $($batch.Count) request(s) as failed" 3
        }
        else {
            foreach($r in $results) {
                $entry = $byBatchId["$($r.Id)"]
                if(-not $entry) { continue }
                $statusCode = 0
                try { $statusCode = [int]$r.status } catch { }

                if($statusCode -ge 200 -and $statusCode -lt 300) {
                    $updated++
                    # Mirror so a subsequent UI refresh sees the new list
                    # without a fresh round-trip.
                    try {
                        if($entry.Policy.Object.PSObject.Properties['assignments']) {
                            $entry.Policy.Object.assignments = $entry.NewAssignments
                        }
                        else {
                            $entry.Policy.Object | Add-Member -MemberType NoteProperty -Name 'assignments' -Value $entry.NewAssignments -Force
                        }
                    }
                    catch { Write-LogDebug "Set-GraphBulkAssignments: in-memory mirror failed for $($entry.Policy.Name): $($_.Exception.Message)" }
                }
                else {
                    $failed++
                    $msg = ""
                    if($r.body -and $r.body.error -and $r.body.error.message) { $msg = $r.body.error.message }
                    Write-Log "Set-GraphBulkAssignments: /assign failed for '$($entry.Policy.Name)' [$($entry.Policy.PolicyType.Title)] - $statusCode $msg" 3
                }
            }

            # Pending entries with no response → treat as failed (matches
            # the codex fix pattern from Set-GraphBulkScopeTags).
            $seenIds = [System.Collections.Generic.HashSet[string]]::new()
            foreach($r in $results) { [void]$seenIds.Add([string]$r.Id) }
            foreach($pending in $byBatchId.Keys) {
                if(-not $seenIds.Contains([string]$pending)) {
                    $failed++
                    $entry = $byBatchId[$pending]
                    Write-Log "Set-GraphBulkAssignments: no batch response for '$($entry.Policy.Name)' [$($entry.Policy.PolicyType.Title)]" 3
                }
            }
        }
    }

    Write-Status ""

    [PSCustomObject]@{
        Types               = $eligibleTypes.Count
        PoliciesScanned     = $scanned
        PoliciesMatched     = $matched
        PoliciesUpdated     = $updated
        PoliciesSkipped     = $skipped
        PoliciesFailed      = $failed
        PoliciesUnsupported = 0   # placeholder — unsupported TYPES are reported separately below
        UnsupportedTypes    = @($unsupportedSelected | ForEach-Object { $_.Title })
        UnknownSelectors    = $unknownSelectors
        Duration            = $stopwatch.Elapsed
    }
}
