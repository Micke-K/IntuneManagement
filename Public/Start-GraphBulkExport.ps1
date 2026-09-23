function Start-GraphBulkExport {
    <#
    .SYNOPSIS
        Bulk-export Intune policy objects to disk.

    .DESCRIPTION
        Public, UI-independent driver for bulk export. The WPF UI is a thin caller of
        this function; the same function can be invoked from a scheduled task or
        CI/CD pipeline without loading any XAML.

        Parameter precedence: SettingsFile < ExportSettings instance < explicit params.

    .PARAMETER ExportSettings
        An [IntuneManagerExportSettings] instance. If omitted, a fresh instance is
        created (which loads defaults from user settings).

    .PARAMETER SettingsFile
        Path to a JSON file produced by the "Save settings for batch job" button.
        Loaded first; any other explicit parameters override its values.

    .PARAMETER ExportFolder
        Root folder for the export. Overrides ExportSettings.ExportFolder.

    .PARAMETER Filter
        Name filter (literal substring, case-insensitive) matched against each
        policy's Name. Empty = no filter.

    .PARAMETER ExportAssignments
        Include assignments in each exported file.

    .PARAMETER AddCompanyName
        Append the current tenant's organization display name as an additional folder
        level under ExportFolder.

    .PARAMETER PolicyType
        Restrict export to these PolicyType IDs. If both PolicyType and PolicyGroup
        are omitted, every group that allows export is processed.

    .PARAMETER PolicyGroup
        Restrict export to these PolicyGroup IDs (expanded to their member types).

    .PARAMETER TokenId
        Authentication token id. Defaults to the current default token.

    .EXAMPLE
        Start-GraphBulkExport -ExportFolder C:\IntuneExport -PolicyGroup DeviceConfiguration

    .EXAMPLE
        Start-GraphBulkExport -SettingsFile .\nightly-export.json

    .OUTPUTS
        PSCustomObject with summary statistics (Types, Policies, Failed, Duration).
    #>
    [CmdletBinding()]
    param(
        [IntuneManagerExportSettings]
        $ExportSettings,

        [string]
        $SettingsFile,

        [string]
        $ExportFolder,

        [string]
        $Filter,

        [Nullable[bool]]
        $ExportAssignments,

        [Nullable[bool]]
        $AddCompanyName,

        [string[]]
        $PolicyType,

        [string[]]
        $PolicyGroup,

        [Int]
        $TokenId = (Get-DefaultTokenId)
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # ---- 1. Resolve settings ----
    $fileSettings = $null
    if ($SettingsFile) {
        if (-not (Test-Path -LiteralPath $SettingsFile)) {
            throw "Settings file '$SettingsFile' not found"
        }
        try {
            $fileSettings = [IO.File]::ReadAllText($SettingsFile) | ConvertFrom-Json
        }
        catch {
            throw "Failed to parse settings file '$SettingsFile': $($_.Exception.Message)"
        }
    }

    if (-not $ExportSettings) {
        $ExportSettings = [IntuneManagerExportSettings]::new()
    }

    # ExportFullMembershipPrefixes was removed. A value can survive that in a
    # settings file OR in the settings store (global and per-tenant), and none of
    # them can be honoured any more - report them before the export starts rather
    # than silently dropping membership data from the output. Runs whether or not
    # a settings file was given, because a stored value needs no file to exist.
    Clear-BulkExportLegacyMembershipSetting -FileSettings $fileSettings -SettingsFile $SettingsFile -TokenId $TokenId

    # File overrides defaults; explicit params override file.
    if ($fileSettings) {
        if ($null -ne $fileSettings.ExportFolder)      { $ExportSettings.ExportFolder      = $fileSettings.ExportFolder }
        if ($null -ne $fileSettings.Filter)            { $ExportSettings.Filter            = $fileSettings.Filter }
        if ($null -ne $fileSettings.ExportAssignments) { $ExportSettings.ExportAssignments = [bool]$fileSettings.ExportAssignments }
        if ($null -ne $fileSettings.AddCompanyName)    { $ExportSettings.AddCompanyName    = [bool]$fileSettings.AddCompanyName }
        if ($null -ne $fileSettings.AddObjectType)     { $ExportSettings.AddObjectType     = [bool]$fileSettings.AddObjectType }
        if ($null -ne $fileSettings.ExportNestedGroupLevels) {
            $n = 0
            if([int]::TryParse([string]$fileSettings.ExportNestedGroupLevels, [ref]$n) -and $n -ge 1) {
                $ExportSettings.ExportNestedGroupLevels = $n
            }
        }
        if (-not $PolicyType   -and $fileSettings.PolicyType)   { $PolicyType   = @($fileSettings.PolicyType) }
        if (-not $PolicyGroup  -and $fileSettings.PolicyGroup)  { $PolicyGroup  = @($fileSettings.PolicyGroup) }
    }

    if ($PSBoundParameters.ContainsKey('ExportFolder'))      { $ExportSettings.ExportFolder      = $ExportFolder }
    if ($PSBoundParameters.ContainsKey('Filter'))            { $ExportSettings.Filter            = $Filter }
    if ($PSBoundParameters.ContainsKey('ExportAssignments')) { $ExportSettings.ExportAssignments = [bool]$ExportAssignments }
    if ($PSBoundParameters.ContainsKey('AddCompanyName'))    { $ExportSettings.AddCompanyName    = [bool]$AddCompanyName }

    if (-not $ExportSettings.ExportFolder) {
        throw "ExportFolder is required (set on ExportSettings, -ExportFolder, or SettingsFile)"
    }

    # ---- 2. Determine target policy types ----
    $selection = Resolve-IntuneTargetSelectors -PolicyType $PolicyType -PolicyGroup $PolicyGroup -Caller 'Start-GraphBulkExport'
    $unknownSelectors = $selection.Unknown   # ids from -PolicyType/-PolicyGroup that matched nothing
    $targetTypes = [System.Collections.Generic.List[object]]::new()
    $targetTypes.AddRange([object[]]$selection.Types)

    if (-not $PolicyType -and -not $PolicyGroup) {
        foreach ($grp in $script:IntuneGroups) {
            if (-not $grp.Title) { continue }
            if ($grp.ShowButtons -is [Object[]] -and $grp.ShowButtons -notcontains "Export") { continue }
            foreach ($pt in $grp.PolicyTypes) { [void]$targetTypes.Add($pt) }
        }
    }

    # De-duplicate (a type can belong to >1 group)
    $targetTypes = @($targetTypes | Sort-Object -Property Id -Unique)

    if ($targetTypes.Count -eq 0) {
        Write-Log "Bulk export: no policy types selected" 2
        return [PSCustomObject]@{ Types = 0; Policies = 0; Failed = 0; UnknownSelectors = $unknownSelectors; Duration = $stopwatch.Elapsed }
    }

    # Pre-compile the name filter once. A literal substring, like 3.x and every
    # other bulk driver: '[Test]' means the text [Test], not a character class.
    $filterRegex = $null
    if ($ExportSettings.Filter) {
        $filterRegex = [Regex]::new([Regex]::Escape($ExportSettings.Filter), 'IgnoreCase')
    }

    Write-Log "Bulk export: $($targetTypes.Count) policy type(s) → $($ExportSettings.ExportFolder)"

    # Honour the ClearCacheBeforeExportImport setting (bit 1 = bulk export, on
    # by default) so re-exports always fetch fresh dependency objects.
    Invoke-GraphCacheClearBeforeOperation -Operation BulkExport -TokenId $TokenId

    # Reset migration-table and AppConfig target-app caches at the start of every
    # bulk export. The caches are designed to dedupe work WITHIN a single export
    # (1 disk write per file at end, 1 Graph call per unique target app), so a
    # stale carry-over from a previous export to a different folder/tenant would
    # produce wrong contents in MigrationTable.json.
    $script:_migFileCache              = @{}
    $script:_migFileObjectsIndex       = @{}
    $script:_migFileDirty              = @{}
    $script:_migFilePathCache          = @{}
    $script:_appConfigTargetAppCache   = @{}
    $script:_migPendingFetch           = [System.Collections.Generic.List[object]]::new()

    # Tell Export-GraphPolicy.End NOT to flush migration files per type — we batch
    # the writes across the entire bulk export and flush once at the very end.
    # Wrapped in try/finally so the flag always clears even if Phase 3 throws,
    # otherwise the next standalone Export-GraphPolicy call would also skip its flush.
    $script:_bulkExportDeferMigFlush = $true
    $script:_skipDirectGet = $true
    try {

    # ---- 3. Export ----
    $totalPolicies = 0
    $failedTypes   = 0

    # One pipeline, whatever the concurrency setting:
    #   1.  List every type's objects in ONE Get-GraphPolicies call.
    #   2.  Batch-GET the full body for every listed object across all types.
    #   2.5 Pre-cache every Entra group the export will reference.
    #   2.6 Pre-walk the nested-group hierarchy when nested export is on.
    #   3.  Write each policy to disk type-by-type.
    # Whether the $batch POSTs behind steps 1, 2, 2.5 and 2.6 go out one at a
    # time or concurrently is decided inside Invoke-GraphBatchRequest from the
    # UseParallelBatchAPI setting - it is not this function's concern. The old
    # per-type interleaved loop that ran when that setting was off is gone: it
    # skipped the group prefetch entirely, which cost a measured 22 minutes on a
    # 963-object tenant (one direct GET per assignment group at ~1.6 s each,
    # versus ~2 s for twenty of them in one $batch). See Docs/GraphBatching.md.
    # Phase 1 — list ALL types in a single Get-GraphPolicies call so the listing
    # round-trips ride the parallel-batch dispatcher (chunks of 20, throttle-many
    # POSTs concurrently). Calling Get-GraphPolicies once per type instead would
    # produce N single-item batches dispatched sequentially — for ~50 types that's
    # ~50 round-trips at ~1s each. Combined it's ceil(N/20) rounds run in parallel.
    Write-Status ("Bulk export - listing {0} policy type(s)" -f $targetTypes.Count) -SkipLog -Force

    $perType    = [ordered]@{}
    foreach ($pt in $targetTypes) { $perType[$pt.Id] = [System.Collections.Generic.List[object]]::new() }
    $allTypeIds = @($targetTypes | ForEach-Object { $_.Id })

    try {
        $allPolicies = @(Get-GraphPolicies -PolicyType $allTypeIds -TokenId $TokenId -IncludeAssignments:$ExportSettings.ExportAssignments -ErrorAction Stop)
    }
    catch {
        Write-LogError "Bulk export: failed to list policies" $_.Exception
        $failedTypes = $targetTypes.Count
        $allPolicies = @()
    }

    # Bucket each returned policy back into its owning type so Phase 3 can iterate
    # type-by-type (which keeps per-type folder routing + status messages intact).
    # Dedup by policy Id within each bucket — some Graph endpoints have overlapping
    # paging cursors so Get-GraphPolicies' AllPages merge can include the same item
    # twice. Without this we'd write the same file (and run all subclass Get()
    # extras) once per duplicate, multiplying the wall time for nothing.
    $seenByType = @{}
    foreach ($p in $allPolicies) {
        $tid = $null
        if ($p.PolicyType) { $tid = $p.PolicyType.Id }
        if (-not ($tid -and $perType.Contains($tid))) { continue }
        if (-not $seenByType.ContainsKey($tid)) {
            $seenByType[$tid] = [System.Collections.Generic.HashSet[string]]::new()
        }
        $key = "$($p.Id)"
        if (-not $seenByType[$tid].Add($key)) {
            # Already seen this id under this type — skip duplicate.
            continue
        }
        [void]$perType[$tid].Add($p)
    }

    if ($filterRegex) {
        foreach ($tid in @($perType.Keys)) {
            $filtered = [System.Collections.Generic.List[object]]::new()
            foreach ($p in $perType[$tid]) {
                if ($filterRegex.IsMatch([string]$p.Name)) { [void]$filtered.Add($p) }
            }
            $perType[$tid] = $filtered
        }
    }

    # Phase 2 — batch-fetch full objects across every type at once
    $needFull = [System.Collections.Generic.List[object]]::new()
    $allPoliciesFlat = [System.Collections.Generic.List[object]]::new()
    foreach ($pt in $targetTypes) {
        $policies = $perType[$pt.Id]
        if (-not $policies) { continue }
        foreach ($p in $policies) {
            [void]$allPoliciesFlat.Add($p)
            if ($p.IsFullObject -eq $false -and $p.Id) { [void]$needFull.Add($p) }
        }
    }
    if ($allPoliciesFlat.Count -gt 0) {
        Write-Status "Bulk export - hydrating policies" -SkipLog -Force
        Invoke-PolicyHydrate -Policies $allPoliciesFlat -TokenId $TokenId
    }

    # Phase 2.5 — pre-cache every AAD group the export will touch in one parallel
    # batch. Sources: assignment targets (when ExportAssignments is on), plus CA
    # conditions.users.includeGroups/excludeGroups and compliance notificationMessageCCList
    # (always exported, so always worth prefetching). Without this, the per-policy
    # PostExportCommand path does a sequential Graph GET per unknown groupId during
    # Phase 3 — hundreds of single-item round-trips for a CA-heavy tenant.
    if ($allPoliciesFlat.Count -gt 0) {
        Sync-BulkExportMigrationGroups -Policies $allPoliciesFlat -TokenId $TokenId
    }

    # Phase 2.6 — when nested-group export is on, walk the hierarchy in batches
    # and cache each parent's child-group list (plus any new child bodies) so the
    # per-policy recursion inside Add-GraphMigrationObject is pure cache hits.
    # Skipped when ExportNestedGroupLevels = 1 (no recursion happens then).
    if ($ExportSettings.ExportNestedGroupLevels -gt 1) {
        Sync-BulkExportNestedGroupHierarchy -MaxDepth $ExportSettings.ExportNestedGroupLevels -TokenId $TokenId
    }

    # Phase 3 — write to disk (sequential I/O; Get() inside Export-GraphPolicy is a no-op
    # for non-override types because Phase 2 already pre-populated them). Status updates
    # per type so the bar doesn't freeze for the duration of the write phase — without
    # these the user sees the last Phase 2 status message for minutes on end.
    $totalToWrite = 0
    foreach ($pt in $targetTypes) { $totalToWrite += $perType[$pt.Id].Count }
    $writeProgress = 0
    $typeIndex     = 0

    # Once, across EVERY type - not per type inside the loop below. Two policies
    # collide when they resolve to the same destination path, and the folder is part
    # of that path: with AddObjectType off every type writes into the export root, so
    # a per-type pass never compared a Compliance policy against a Configuration
    # policy of the same name and one silently overwrote the other.
    Set-BulkExportFilenameCollisionFlag -Policies $allPoliciesFlat `
        -AddObjectType:($ExportSettings.AddObjectType -eq $true)

    foreach ($pt in $targetTypes) {
        $typeIndex++
        $policies = $perType[$pt.Id]
        if (-not $policies -or $policies.Count -eq 0) {
            Write-Log "Bulk export: no $($pt.Title) objects matched"
            continue
        }
        Write-Status `
            -Text   ("Bulk export - {0} ({1} of {2})" -f $pt.Title, $typeIndex, $targetTypes.Count) `
            -Detail ("Writing {0} object(s) · {1} of {2} total" -f $policies.Count, $writeProgress, $totalToWrite) `
            -SkipLog -Force
        try {
            $policies | Export-GraphPolicy -ExportSettings $ExportSettings
            $totalPolicies += $policies.Count
            $writeProgress += $policies.Count
        }
        catch {
            $failedTypes++
            Write-LogError "Bulk export: failed while exporting $($pt.Title)" $_.Exception
        }
    }

    # Flush every MigrationTable.json that Add-GraphMigrationObject mutated. The
    # in-memory cache batched all the per-policy mutations during Phase 3; this one
    # call writes each touched file to disk exactly once. Wrapped in try so any
    # disk error doesn't mask the export-finished status.
    try { Save-GraphMigrationFilesPending } catch {
        Write-LogError "Bulk export: failed to flush migration table(s)" $_.Exception
    }

    }
    finally {
        # Clear the per-type-flush guard so future standalone Export-GraphPolicy
        # calls flush normally in their End block. Runs even on Phase 3 exceptions.
        $script:_bulkExportDeferMigFlush = $false
        $script:_skipDirectGet = $false
    }

    Write-Status $null

    # ---- 4. Persist user-settings (round-trip the form's "remember last used") ----
    try { $ExportSettings.Save() } catch { }
    # Persist to the SAME SubPath these keys are registered under ("IntuneManager") so
    # Get-SettingValue reads them back - a root-path write here was a dead location
    # (same bug the import side fixed; see IntuneBaseClasses.ps1 ImportSettings.Save).
    Save-SettingStoreValue "IntuneManager" "AddCompanyName"          $ExportSettings.AddCompanyName
    Save-SettingStoreValue "IntuneManager" "ExportAssignments"        $ExportSettings.ExportAssignments
    Save-SettingStoreValue "IntuneManager" "ExportNestedGroupLevels"  $ExportSettings.ExportNestedGroupLevels

    $stopwatch.Stop()
    $summary = [PSCustomObject]@{
        Types            = $targetTypes.Count
        Policies         = $totalPolicies
        Failed           = $failedTypes
        UnknownSelectors = $unknownSelectors
        Duration         = $stopwatch.Elapsed
    }
    Write-Log ("Bulk export finished. Types={0} Policies={1} Failed={2} Duration={3}" -f `
        $summary.Types, $summary.Policies, $summary.Failed, $summary.Duration)
    return $summary
}
