#region Batch functions

# The ONLY places the two batching settings are read. Every other call path asks
# these, so what each setting means is defined here and nowhere else:
#   UseBatchAPI          combine requests into POST /$batch - on EVERY path that
#                        can batch (listing, hydrate, sub-resources, assignments,
#                        import, delete), not just the four that used to check.
#   UseParallelBatchAPI  send those $batch POSTs concurrently (PowerShell 7+ only).
#                        Concurrency and nothing else - it no longer selects a
#                        pipeline anywhere.
# Static.Tests.ps1 fails the build if either key is consulted outside this file.
# Map of every call path: Docs/GraphBatching.md.
function Test-GraphBatchEnabled
{
    return ((Get-SettingValue "UseBatchAPI") -eq $true)
}

function Test-GraphParallelEnabled
{
    return ((Get-SettingValue "UseParallelBatchAPI") -eq $true -and $PSVersionTable.PSVersion.Major -ge 7)
}

# A parallel dispatch needs more than this many queued sub-requests to be worth
# the runspace fan-out; a smaller queue goes out as one serial POST even with
# UseParallelBatchAPI on. This used to be a bare 20 in the dispatch condition,
# which cost real debugging time when "parallel is on but nothing is parallel".
$script:GraphParallelBatchMinQueue = 20

function Invoke-GraphBatchRequest
{
    param(

        [System.Collections.Generic.List[PSCustomObject]]$BatchObjects,
        # Description of the batch objects eg. Policies, Delete etc.
        [string]$BatchType,
        [switch]$SkipWarnings,
        [switch]$IncludedFailed,
        [switch]$AllPages,
        [int]$TokenId = 0)

    if($BatchObjects.Count -eq 0) { return }

    # Resolve %OrganizationId% in every sub-request URL up-front. Invoke-MSGraphAPI
    # does this for direct calls (see its body), but a batch ships sub-request URLs
    # inside the POST payload — Graph parses each verbatim and fails 400 if the
    # placeholder leaks through. Endpoints that need this today: organization/
    # %OrganizationId%/branding... Resolution mirrors Invoke-MSGraphAPI: ask the
    # active provider for the call's tenant id, fall back to the script default.
    $needSubst = $false
    foreach($o in $BatchObjects) { if($o.url -match '%OrganizationId%') { $needSubst = $true; break } }
    if($needSubst) {
        $callTenantOrgId = $null
        try {
            $provider = Get-AuthProvider
            if($provider) {
                $userInfo = $provider.GetUserInfo($TokenId)
                if($userInfo -and $userInfo.TenantId) { $callTenantOrgId = $userInfo.TenantId }
            }
        }
        catch { }
        if(-not $callTenantOrgId) { $callTenantOrgId = (Get-CurrentTenantId) }

        if($callTenantOrgId) {
            foreach($o in $BatchObjects) {
                if($o.url -match '%OrganizationId%') {
                    $o.url = $o.url -replace '%OrganizationId%', $callTenantOrgId
                }
            }
        }
        else {
            Write-Log "Invoke-GraphBatchRequest: could not resolve %OrganizationId% for TokenId $TokenId - sub-requests using that placeholder will likely 400" 2
        }
    }

    # O(1) lookup: batch id -> original request object (for status reporting + URL updates on paging)
    $byId = @{}
    foreach($o in $BatchObjects) { $byId["$($o.id)"] = $o }

    # O(1) lookup: batch id -> accumulated result object (for AllPages stitching)
    $resultsById = @{}

    # O(1) lookup: batch id -> cumulative page count across re-dispatches caused by per-item
    # @odata.nextLink stitching. Each sub-item starts at 1 (first dispatch counts as page 1);
    # Add-GraphBatchResult increments when an item is kept alive for another page. Surfaced
    # back into $script:AllGraphCalls so the Graph Calls log shows per-item totals.
    $pagesById = @{}
    foreach($o in $BatchObjects) { $pagesById["$($o.id)"] = 1 }

    $batchResults        = [System.Collections.Generic.List[PSCustomObject]]::new()
    $requestObjects      = [System.Collections.Generic.List[PSCustomObject]]::new($BatchObjects)

    # One-per-second endpoints (Conditional Access, named locations, identity
    # protection) get their own queue: never in parallel dispatch, one sub-request
    # per $batch, each gated on the tenant's clock. See Internal/GraphRateLimits.ps1.
    $pacedObjects        = Split-GraphPacedBatchRequests -RequestObjects $requestObjects -BatchType $BatchType

    $directResults = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ---- Paced queue: always direct ----
    # A paced sub-request used to go out as a $batch envelope around ONE request.
    # The envelope is a whole extra round-trip, and Graph counts the sub-request
    # against the one-per-second limit exactly as it counts a direct call - on
    # one measured export 95 of 135 $batch POSTs were such envelopes. So the
    # paced queue goes direct, one call at a time, and Invoke-MSGraphAPI gates
    # every paced URL on the tenant clock itself. Still never parallel, still
    # one per second, whatever the two settings say.
    if($pacedObjects.Count -gt 0)
    {
        foreach($r in @(Invoke-GraphBatchRequestDirect -RequestObjects $pacedObjects -BatchType $BatchType -TokenId $TokenId `
            -AllPages:$AllPages -SkipWarnings:$SkipWarnings -IncludedFailed:$IncludedFailed -Reason 'one per second')) { [void]$directResults.Add($r) }
        $pacedObjects.Clear()
    }

    # ---- Direct dispatch: batching off, or a single request ----
    # With UseBatchAPI off every queued request goes out as its own direct call
    # and comes back in the same {id, status, headers, body} shape, so no caller
    # changes and every call path honours the setting. A lone request takes the
    # same path even with batching on: a $batch envelope around one GET is a
    # whole extra round-trip for nothing (fifty-three of them on one measured
    # export).
    if($requestObjects.Count -gt 0 -and (-not (Test-GraphBatchEnabled) -or $requestObjects.Count -eq 1))
    {
        $reason = if(Test-GraphBatchEnabled) { 'single request' } else { 'batching off' }
        foreach($r in @(Invoke-GraphBatchRequestDirect -RequestObjects $requestObjects -BatchType $BatchType -TokenId $TokenId `
            -AllPages:$AllPages -SkipWarnings:$SkipWarnings -IncludedFailed:$IncludedFailed -Reason $reason)) { [void]$directResults.Add($r) }
        return $directResults.ToArray()
    }
    if($requestObjects.Count -eq 0) { return $directResults.ToArray() }

    $curBatch            = 1
    $expectedReturnCount = 0
    $batchResultCount    = 0
    $maxRetryCount       = 10
    # Server errors get a much smaller budget than throttling. A 429 clears when
    # the rate window rolls over, so ten rounds are worth waiting out. A 5xx that
    # repeats is usually a permanent data error - Graph returning 500 for an
    # object it cannot serialize (a Terms of Use agreement with a missing policy
    # file, say) will still return 500 on the tenth try, after 100 s of back-off.
    $maxServerErrorRetry = 3
    $retryObjects        = @{}

    # Decide whether to dispatch the chunks in parallel. Pre-acquire the token once so each
    # runspace can do a raw Invoke-WebRequest with a static Authorization header -- no module
    # functions inside the runspaces (they wouldn't have access anyway), no token refresh
    # races. If anything goes wrong setting this up we silently fall back to sequential.
    $useParallel  = $false
    $parallelCtx  = $null
    if(Test-GraphParallelEnabled)
    {
        $parallelCtx = Initialize-ParallelBatchContext -TokenId $TokenId
        if($parallelCtx) { $useParallel = $true }
    }
    $parallelSkipLogged = $false

    # Per-chunk progress. This function dispatches 20 sub-requests per POST; for large
    # inputs (600+ sub-requests across list loads, assignments, tools, bulk ops) the
    # caller's single Write-Status would otherwise sit frozen for the whole run.
    # -Detail keeps the caller's status text and adds a live progress line under it
    # (same pattern as Invoke-PolicyHydrateBodyBatch). Only shown when more than one
    # chunk is needed so single-batch calls don't flicker. -AllPages re-dispatches can
    # grow the workload beyond the initial count - the denominator tracks that.
    $totalRequests = $requestObjects.Count
    $showProgress  = ($totalRequests -gt 20)

    while($requestObjects.Count -gt 0)
    {
        $queueObjects = $requestObjects
        $chunkSize    = 20
        $pending      = $requestObjects.Count

        # ------- Parallel path -------
        if($useParallel -and $queueObjects.Count -le $script:GraphParallelBatchMinQueue -and -not $parallelSkipLogged)
        {
            # Say so once per call. Without this line a user with the setting on
            # and a small queue had no way to tell why nothing ran in parallel.
            Write-Log "Batch $BatchType`: parallel requested but the queue is $($queueObjects.Count) (threshold $($script:GraphParallelBatchMinQueue)) - dispatching serially"
            $parallelSkipLogged = $true
        }
        if($useParallel -and $queueObjects.Count -gt $script:GraphParallelBatchMinQueue)
        {
            if($showProgress) {
                Write-Status -Detail ("{0}: dispatching {1} request(s) in parallel batches" -f $BatchType, $queueObjects.Count) -SkipLog -Force
            }
            $allResponses = Invoke-ParallelGraphBatchPosts -RequestObjects $queueObjects -Context $parallelCtx -BatchType $BatchType
            $expectedReturnCount += $queueObjects.Count

            $parallelRetryRef     = [ref]0
            $parallelStatusCounts = @{}
            foreach($batchResult in $allResponses)
            {
                Add-GraphBatchResult `
                    -BatchResult $batchResult `
                    -ById $byId `
                    -ResultsById $resultsById `
                    -PagesById $pagesById `
                    -RequestObjects $queueObjects `
                    -RetryObjects $retryObjects `
                    -MaxRetryCount $maxRetryCount `
                    -BatchResults $batchResults `
                    -AllPages:$AllPages `
                    -SkipWarnings:$SkipWarnings `
                    -IncludedFailed:$IncludedFailed `
                    -BatchResultCountRef ([ref]$batchResultCount) `
                    -RetryAfterRef $parallelRetryRef `
                    -MaxServerErrorRetryCount $maxServerErrorRetry `
                    -RetryStatusCounts $parallelStatusCounts
            }

            Update-BatchPageCounts -PagesById $pagesById

            # If any batch came back retryable, back off before re-dispatching.
            if($queueObjects.Count -gt 0 -and $parallelRetryRef.Value -gt 0)
            {
                $sleep = [Math]::Max(5, $parallelRetryRef.Value)
                Write-Log "Parallel batch $BatchType returned $(Format-GraphRetryStatus $parallelStatusCounts). Waiting $sleep seconds before retry..." 2
                # Sliced, pumped wait: a plain Start-Sleep here freezes the window for
                # the whole back-off and, on Avalonia, never even paints the message.
                Wait-GraphThrottle -Seconds $sleep -BatchType $BatchType -Queued $queueObjects.Count
            }

            # Loop continues if -AllPages re-queued anything via $queueObjects.
            continue
        }

        # ------- Sequential path (default) -------
        $take     = [Math]::Min($chunkSize, $queueObjects.Count)
        $batchArr = $queueObjects.GetRange(0, $take).ToArray()

        $expectedReturnCount += $batchArr.Count

        $batchObj = [PSCustomObject]@{ requests = $batchArr }
        $json     = $batchObj | ConvertTo-Json -Depth 20

        Write-Log "Invoke batch $curBatch $BatchType ($($batchArr.Count) requests)"

        # $batchResultCount = completed sub-requests so far; denominator grows when
        # -AllPages keeps items alive for more pages.
        $progressTotal = [Math]::Max($totalRequests, $batchResultCount + $pending)
        if($showProgress) {
            Write-Status -Detail ("{0}: {1} of {2} requests - batch {3}" -f $BatchType, $batchResultCount, $progressTotal, $curBatch) -SkipLog -Force
        }

        $retryAfter = 0
        $tmpResults = Invoke-MSGraphAPI -Url "`$batch" -Body $json -Method "POST" -TokenId $TokenId

        # If Invoke-MSGraphAPI returned $null (auth failure, network error, etc.),
        # the batch never gets processed. Without bailing, the outer while() loop
        # spins forever because $queueObjects is only reduced by Add-GraphBatchResult.
        if($null -eq $tmpResults)
        {
            Write-Log "Batch call returned no result (auth or transport failure). Aborting batch dispatch with $pending request(s) unprocessed." 3
            break
        }

        $retryAfterRef     = [ref]0
        $retryStatusCounts = @{}
        foreach($batchResult in $tmpResults.responses)
        {
            Add-GraphBatchResult `
                -BatchResult $batchResult `
                -ById $byId `
                -ResultsById $resultsById `
                -PagesById $pagesById `
                -RequestObjects $queueObjects `
                -RetryObjects $retryObjects `
                -MaxRetryCount $maxRetryCount `
                -BatchResults $batchResults `
                -AllPages:$AllPages `
                -SkipWarnings:$SkipWarnings `
                -IncludedFailed:$IncludedFailed `
                -BatchResultCountRef ([ref]$batchResultCount) `
                -RetryAfterRef $retryAfterRef `
                -MaxServerErrorRetryCount $maxServerErrorRetry `
                -RetryStatusCounts $retryStatusCounts
        }
        $retryAfter = $retryAfterRef.Value

        Update-BatchPageCounts -PagesById $pagesById

        if($queueObjects.Count -gt 0 -and $retryAfter -ne 0)
        {
            if($retryAfter -lt 5) { $retryAfter = 5 }
            # Report what Graph actually returned. This used to say "429 - Too many
            # requests" for every retryable status, so a permanent 500 was reported
            # to the user as throttling.
            Write-Log "Batch $BatchType returned $(Format-GraphRetryStatus $retryStatusCounts). Waiting $retryAfter seconds before retry." 2
            # Sliced, pumped wait: a plain Start-Sleep here freezes the window for
            # the whole back-off and, on Avalonia, never even paints the message.
            Wait-GraphThrottle -Seconds $retryAfter -BatchType $BatchType -Queued $queueObjects.Count
        }
        $curBatch++
    }

    # Clear the progress detail line so it doesn't linger under the caller's next status.
    if($showProgress) { Write-Status -Detail "" -SkipLog -Force }

    if($batchResultCount -ne $expectedReturnCount -and -not $SkipWarnings)
    {
        Write-Log "Not all batch objects returned. Expected $expectedReturnCount but only got $batchResultCount" 2
    }

    $pagedItems = @($pagesById.GetEnumerator() | Where-Object { $_.Value -gt 1 })
    if($pagedItems.Count -gt 0)
    {
        $totalPages = ($pagedItems | Measure-Object -Property Value -Sum).Sum
        Write-Log "Batch $BatchType paged: $($pagedItems.Count) sub-item(s) required $totalPages page(s) total"
    }

    foreach($r in $batchResults) { [void]$directResults.Add($r) }
    return $directResults.ToArray()
}

# Dispatch a batch queue as individual Invoke-MSGraphAPI calls and return the
# results in the batch response shape ({id, status, headers, body}). Used when
# UseBatchAPI is off and for single-request queues. Each sub-request's Accept
# header becomes -ODataMetadata (the wrapper builds Accept itself and would
# refuse a duplicate), other headers ride on -AdditionalHeaders, and the status
# is read back from the telemetry row the wrapper records for every call -
# success and failure alike - so a 404 or 429 that the wrapper has already
# retried out reports as exactly that rather than as a silent $null.
function Invoke-GraphBatchRequestDirect
{
    param(
        [System.Collections.Generic.List[PSCustomObject]]$RequestObjects,
        [string]$BatchType,
        [switch]$SkipWarnings,
        [switch]$IncludedFailed,
        [switch]$AllPages,
        [int]$TokenId = 0,
        # Shown in the log and status: 'batching off', 'single request', 'one per second'.
        [string]$Reason = 'batching off'
    )

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    if(-not $RequestObjects -or $RequestObjects.Count -eq 0) { return $results.ToArray() }

    $total = $RequestObjects.Count
    if($total -gt 1) { Write-Log "Direct dispatch $BatchType ($total requests, $Reason)" }

    $n = 0
    foreach($o in $RequestObjects)
    {
        $n++
        if($total -gt 1) {
            Write-Status -Detail ("{0}: {1} of {2} requests - direct calls ({3})" -f $BatchType, $n, $total, $Reason) -SkipLog -Force
        }

        $method = if($o.method) { "$($o.method)".ToUpperInvariant() } else { "GET" }
        # Parity with $batch: a sub-request with no Accept, or an Accept whose
        # metadata value is empty (the hydrate header for a type that declares
        # none reads "odata.metadata="), gets Graph's own default, minimal.
        # Defaulting to full here put @odata annotations and navigation links
        # into every Applications export that the batch path did not have.
        $odata  = "minimal"
        $extra  = @{}
        if($o.headers) {
            foreach($h in @($o.headers.Keys)) {
                $v = "$($o.headers[$h])"
                if($h -ieq 'Accept') {
                    if($v -match 'odata\.metadata=(full|minimal|none)') { $odata = $Matches[1] }
                }
                elseif($h -ieq 'Content-Type') { }
                else { $extra[$h] = $o.headers[$h] }
            }
        }

        $body = $null
        if($o.PSObject.Properties['body'] -and $null -ne $o.body) {
            $body = if($o.body -is [string]) { $o.body } else { ($o.body | ConvertTo-Json -Depth 20) }
        }

        $callArgs = @{ Url = $o.url; HttpMethod = $method; TokenId = $TokenId; ODataMetadata = $odata }
        if($body)                        { $callArgs.Content           = $body }
        if($extra.Count -gt 0)           { $callArgs.AdditionalHeaders = $extra }
        if($AllPages -and $method -eq 'GET') { $callArgs.AllPages      = $true }
        # SkipWarnings controls only this dispatcher's warning below. Passing it
        # through as Invoke-MSGraphAPI -NoError would return on the first error
        # and bypass that function's 429/5xx retry handling.
        $previousLastCall = if($script:AllGraphCalls -and $script:AllGraphCalls.Count -gt 0) {
            $script:AllGraphCalls[$script:AllGraphCalls.Count - 1]
        } else { $null }
        $content = $null
        try { $content = Invoke-MSGraphAPI @callArgs } catch { }

        $status = 0
        # Invoke-MSGraphAPI keeps a capped 2,000-row telemetry log. At the cap it
        # removes the oldest row before adding this call, so Count does not grow.
        # The newly appended row is a distinct object even when the size is
        # unchanged (and the latest retry row carries the final status).
        if($script:AllGraphCalls -and $script:AllGraphCalls.Count -gt 0) {
            $row = $script:AllGraphCalls[$script:AllGraphCalls.Count - 1]
            if(-not [object]::ReferenceEquals($row, $previousLastCall) -and $row.StatusCode) {
                $status = [int]$row.StatusCode
            }
        }
        if($status -eq 0 -and $null -ne $content) { $status = 200 }

        $ok = ($status -ge 200 -and $status -lt 300)
        if(-not $ok -and -not $IncludedFailed) {
            if(-not $SkipWarnings) { Write-Log "Direct request $($o.id) ($method $($o.url)) returned $status. Skipping..." 2 }
            continue
        }
        [void]$results.Add([PSCustomObject]@{ id = $o.id; status = $status; headers = @{}; body = $content })
    }

    if($total -gt 1) { Write-Status -Detail "" -SkipLog -Force }
    return $results.ToArray()
}

# Walk back through $script:AllGraphCalls and overwrite each batch sub-item's
# Response.PageCount with the cumulative page count from $PagesById. We can't assume the
# update only touches the last row: the parallel dispatch path adds one telemetry row per
# in-flight batch POST, so a single Invoke-GraphBatchRequest cycle may have appended many
# rows. Walks until every id in $PagesById has been seen at least once, or we exhaust the
# log (safety cap: stop at the first non-batch row we hit going backward to avoid scanning
# the entire history). Idempotent — re-running with the same data is a no-op.
function Update-BatchPageCounts
{
    param([hashtable]$PagesById)

    if(-not $script:AllGraphCalls -or $script:AllGraphCalls.Count -eq 0) { return }
    if(-not $PagesById -or $PagesById.Count -eq 0) { return }

    $remaining = [System.Collections.Generic.HashSet[string]]::new()
    foreach($k in $PagesById.Keys) { [void]$remaining.Add("$k") }

    for($i = $script:AllGraphCalls.Count - 1; $i -ge 0 -and $remaining.Count -gt 0; $i--)
    {
        $row = $script:AllGraphCalls[$i]
        if(-not $row.IsBatch) { break }
        if(-not $row.BatchRequests) { continue }

        foreach($item in $row.BatchRequests)
        {
            $key = "$($item.Id)"
            if($remaining.Contains($key) -and $item.Response)
            {
                $item.Response.PageCount = $PagesById[$key]
                [void]$remaining.Remove($key)
            }
        }
    }
}

# Process one batch result and update the bookkeeping state. Extracted so the sequential
# and parallel paths share identical processing semantics (retry-on-429, AllPages stitching,
# failure handling, etc.).
function Add-GraphBatchResult
{
    param(
        $BatchResult,
        [hashtable]$ById,
        [hashtable]$ResultsById,
        [hashtable]$PagesById,
        [System.Collections.Generic.List[PSCustomObject]]$RequestObjects,
        [hashtable]$RetryObjects,
        [int]$MaxRetryCount,
        [int]$MaxServerErrorRetryCount = 3,
        # Optional tally of the retryable statuses seen this round, so the caller
        # can report what actually came back instead of assuming throttling.
        [hashtable]$RetryStatusCounts,
        [System.Collections.Generic.List[PSCustomObject]]$BatchResults,
        [switch]$AllPages,
        [switch]$SkipWarnings,
        [switch]$IncludedFailed,
        [ref]$BatchResultCountRef,
        [ref]$RetryAfterRef
    )

    $keepToGetNextPage = $false
    $failed = $false
    $requestObject = $ById["$($BatchResult.Id)"]

    if($requestObject.Method -eq "DELETE" -and $BatchResult.Status -eq 200)
    {
        [void]$BatchResults.Add($BatchResult)
        $BatchResultCountRef.Value++
        [void]$RequestObjects.Remove($requestObject)
        return
    }

    if($BatchResult.Status -ge 300 -or -not $BatchResult.body)
    {
        if(($BatchResult.Status -eq 429 -or $BatchResult.Status -in @(500, 502, 503, 504)) -and $requestObject)
        {
            # 429 = throttling; 500/502/503/504 = transient server / gateway errors.
            # Both are safe to re-issue - requeue this sub-request (honoring Retry-After
            # when present), but on SEPARATE budgets: throttling is worth waiting out,
            # a repeating server error is not (see $maxServerErrorRetry above).
            # Honor Retry-After when present, else default to 10s so an absent header still
            # backs off (rather than re-dispatching this sub-request with zero delay).
            $wait = Get-GraphRetryAfterSeconds $BatchResult.headers.'Retry-After'
            if($RetryAfterRef -and $wait -gt $RetryAfterRef.Value) { $RetryAfterRef.Value = $wait }

            # Counted per sub-request AND per kind, so a throttled request that also
            # meets one server blip doesn't lose its throttle budget (or vice versa).
            $attempt = Register-GraphRetryAttempt -Id $BatchResult.Id -Status $BatchResult.Status `
                -RetryObjects $RetryObjects -MaxThrottleRetry $MaxRetryCount `
                -MaxServerErrorRetry $MaxServerErrorRetryCount -StatusCounts $RetryStatusCounts

            if($attempt.Exhausted)
            {
                Write-Log "Giving up on batch object $($BatchResult.Id) after $($attempt.Attempts) attempt(s) returning $($BatchResult.Status). Removing..." 3
                [void]$RequestObjects.Remove($requestObject)
            }
            return
        }
        $failed = $true
        if(-not $SkipWarnings)
        {
            Write-Log "Batch result $($BatchResult.Status) for URL $($requestObject.URL). Skipping..." 2
        }
    }
    elseif($AllPages -and $BatchResult.body.'@odata.nextLink')
    {
        $keepToGetNextPage = $true
        $uri    = [URI]$BatchResult.body.'@odata.nextLink'
        $newUrl = $uri.PathAndQuery.Substring($uri.PathAndQuery.IndexOf("/",1)).TrimStart("/")
        if($newUrl) { $requestObject.URL = "/" + $newUrl }
        if($PagesById) { $PagesById["$($BatchResult.Id)"]++ }
    }

    if(-not $keepToGetNextPage)
    {
        [void]$RequestObjects.Remove($requestObject)
    }

    if(-not $failed -or $IncludedFailed)
    {
        Write-LogDebug "Value count for $($BatchResult.Id): $(($BatchResult.body.value | Measure-Object).Count)"

        $existing = $ResultsById["$($BatchResult.Id)"]
        if($existing -and $existing.body.value -is [Array] -and $BatchResult.body.value -is [Array])
        {
            $existing.body.value += $BatchResult.body.value
        }
        else
        {
            [void]$BatchResults.Add($BatchResult)
            $ResultsById["$($BatchResult.Id)"] = $BatchResult
        }
        $BatchResultCountRef.Value++
    }
}

# Pre-acquire a fresh access token + endpoint info so each parallel runspace can do raw
# HTTP without invoking module functions or touching MSAL state. Returns $null if anything
# is wrong (caller falls back to sequential).
function Initialize-ParallelBatchContext
{
    param([int]$TokenId = 0)

    try
    {
        # Get the access token through the provider that OWNS this token id, exactly
        # as Invoke-MSGraphAPI does - routing by the merely-active provider instead
        # broke mixed-provider sessions: a provider that does not own the id returns
        # no token (parallelism silently degrades to serial), and one that ignores the
        # id entirely (MgGraph) hands back ITS bearer for another tenant's batch.
        # Falling back to the active provider keeps id 0 / unregistered ids working.
        # An earlier version called Connect-EntraEnvironment + Get-FullToken
        # unconditionally, which forced an MSAL silent auth even when MgGraph was
        # active and triggered the auto-switch back to MSAL.
        $authProvider = Resolve-AuthTokenProvider $TokenId
        if(-not $authProvider) { $authProvider = Get-AuthProvider }
        if(-not $authProvider) { return $null }

        # The runspaces below call Invoke-WebRequest themselves, so a provider
        # that answers every request in-process (the offline mock) would be
        # bypassed and its bearer sent to the real Graph. No context means the
        # caller dispatches serially through Invoke-MSGraphAPI, which routes to
        # the provider.
        if($authProvider.RoutesAllRequests)
        {
            Write-LogDebug "Parallel batch: provider '$($authProvider.Id)' routes every request itself - dispatching serially"
            return $null
        }

        $graphDomain = Get-GraphDomain $TokenId
        $accessToken = $authProvider.GetAccessToken($TokenId, "https://$graphDomain")
        if(-not $accessToken) { return $null }

        $graphVersion = if($script:defaultVersion) { $script:defaultVersion }
                        elseif((Get-SettingValue "UseGraphV1") -eq $true) { "v1.0" }
                        else { "beta" }
        $script:defaultVersion = $graphVersion

        $throttle = 4
        try
        {
            $cfg = Get-SettingValue "ParallelBatchThrottle"
            if($cfg) { $throttle = [int]$cfg }
        } catch {}
        if($throttle -lt 1) { $throttle = 1 }
        if($throttle -gt 20) { $throttle = 20 }

        $proxy = Get-ProxyURI

        return [PSCustomObject]@{
            BatchUrl     = "https://$graphDomain/$graphVersion/`$batch"
            AccessToken  = $accessToken
            ProxyUri     = $proxy
            ThrottleLimit = $throttle
            IsPSv7       = ($PSVersionTable.PSVersion.Major -ge 7)
            # ProviderId is propagated into each parallel runspace so the Graph Calls
            # log can show which provider minted these batches. Without it the
            # Provider column would be blank for parallel-batch rows.
            ProviderId   = $authProvider.Id
        }
    }
    catch
    {
        Write-Log "Failed to initialize parallel batch context: $($_.Exception.Message)" 2
        return $null
    }
}

# Build all $batch payloads from the current $RequestObjects queue and dispatch them
# concurrently via ForEach-Object -Parallel. The token + URL + proxy are captured via
# $using: so the runspaces don't need module state. Returns a flat list of batch responses
# in the same shape the sequential path produces (so the caller can use the same
# Add-GraphBatchResult processing).
function Invoke-ParallelGraphBatchPosts
{
    param(
        [System.Collections.Generic.List[PSCustomObject]]$RequestObjects,
        [PSCustomObject]$Context,
        [string]$BatchType
    )

    # Build batch payloads from the queue without removing the items.
    # Add-GraphBatchResult will handle queue mutations as it processes responses
    # (removing successfully-processed items, leaving paged items in place with updated URLs).
    $payloads = [System.Collections.Generic.List[object]]::new()
    $index = 0
    while($index -lt $RequestObjects.Count)
    {
        $take     = [Math]::Min(20, $RequestObjects.Count - $index)
        $batchArr = $RequestObjects.GetRange($index, $take).ToArray()
        [void]$payloads.Add([PSCustomObject]@{
            Json = ([PSCustomObject]@{ requests = $batchArr } | ConvertTo-Json -Depth 20)
        })
        $index += $take
    }

    Write-Log "Dispatching $($payloads.Count) batch(es) in parallel (throttle=$($Context.ThrottleLimit)) for $BatchType"

    $url        = $Context.BatchUrl
    $token      = $Context.AccessToken
    $proxy      = $Context.ProxyUri
    $isPSv7     = $Context.IsPSv7
    $providerId = $Context.ProviderId

    $rawResults = $payloads | ForEach-Object -ThrottleLimit $Context.ThrottleLimit -Parallel {
        $localUrl        = $using:url
        $localToken      = $using:token
        $localProxy      = $using:proxy
        $localPSv7       = $using:isPSv7
        $localProviderId = $using:providerId

        $requestId = [Guid]::NewGuid().Guid
        $headers = @{
            'Content-Type'           = 'application/json; charset=utf-8'
            'Authorization'          = "Bearer $localToken"
            'x-ms-client-request-id' = $requestId
        }

        $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($_.Json)
        $params = @{
            Uri             = $localUrl
            Method          = 'POST'
            Headers         = $headers
            Body            = $bodyBytes
            UseBasicParsing = $true
            ErrorAction     = 'Stop'
        }
        if($localProxy)  { $params.Proxy = $localProxy }
        if($localPSv7)   { $params.ProgressAction = 'SilentlyContinue' }

        $sentAt = Get-Date
        try
        {
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $response = Invoke-WebRequest @params
            $stopwatch.Stop()

            $bytes = 0
            try {
                if($response.RawContentLength -gt 0) { $bytes = [long]$response.RawContentLength }
                elseif($response.Content) { $bytes = [long]$response.Content.Length }
            } catch {}

            $parsed = $response.Content | ConvertFrom-Json -Depth 20

            # Build per-item records for the Graph log with the same columns as the top grid
            # (StatusCode, KB, ObjectCount).
            $perItemById = @{}
            $totalObjectCount = 0
            foreach($r in $parsed.responses) {
                $rid = "$($r.id)"
                $itemBytes = 0
                $itemCount = 0
                if($null -ne $r.body) {
                    try { $itemBytes = ($r.body | ConvertTo-Json -Depth 20 -Compress).Length } catch {}
                    if($r.body.value -is [Array]) {
                        $itemCount = $r.body.value.Count
                    }
                    elseif($r.status -ge 200 -and $r.status -lt 300) {
                        $itemCount = 1
                    }
                }
                $totalObjectCount += $itemCount
                $perItemById[$rid] = [PSCustomObject]@{
                    StatusCode  = $r.status
                    KB          = [Math]::Round($itemBytes / 1024.0, 1)
                    ObjectCount = $itemCount
                    PageCount   = 1
                }
            }

            $envelope = $_.Json | ConvertFrom-Json -Depth 20
            $batchItems = @()
            foreach($req in $envelope.requests) {
                $rid  = "$($req.id)"
                $resp = $perItemById[$rid]
                if(-not $resp) { $resp = [PSCustomObject]@{ StatusCode = $null; KB = 0.0; ObjectCount = 0; PageCount = 1 } }
                $batchItems += [PSCustomObject]@{
                    Id       = $req.id
                    Method   = $req.method
                    URL      = $req.url
                    Response = $resp
                }
            }

            [PSCustomObject]@{
                Success     = $true
                Responses   = $parsed.responses
                Error       = $null
                Telemetry   = [PSCustomObject]@{
                    ID            = $requestId
                    URL           = $localUrl
                    Method        = 'POST'
                    IsBatch       = $true
                    BatchRequests = $batchItems
                    StatusCode    = [int]$response.StatusCode
                    Time          = $sentAt
                    Duration      = $stopwatch.Elapsed.TotalMilliseconds
                    KB            = [Math]::Round($bytes / 1024.0, 1)
                    ObjectCount   = $totalObjectCount
                    PageCount     = 1
                    ErrorMessage  = $null
                    Provider      = $localProviderId
                }
            }
        }
        catch
        {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
            [PSCustomObject]@{
                Success     = $false
                Responses   = $null
                Error       = $_.Exception.Message
                StatusCode  = $statusCode
                Telemetry   = [PSCustomObject]@{
                    ID            = $requestId
                    URL           = $localUrl
                    Method        = 'POST'
                    IsBatch       = $true
                    BatchRequests = @()
                    StatusCode    = $statusCode
                    Time          = $sentAt
                    Duration      = $null
                    KB            = 0.0
                    ObjectCount   = 0
                    PageCount     = 1
                    ErrorMessage  = $_.Exception.Message
                    Provider      = $localProviderId
                }
            }
        }
    }

    # Surface the per-batch telemetry into $script:AllGraphCalls so the Graph log UI sees it.
    if($null -eq $script:AllGraphCalls) {
        $script:AllGraphCalls = [System.Collections.Generic.List[PSCustomObject]]::new()
    }
    $allResponses = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach($r in $rawResults)
    {
        if($r.Telemetry) {
            if($script:AllGraphCalls.Count -ge 2000) { $script:AllGraphCalls.RemoveAt(0) }
            [void]$script:AllGraphCalls.Add($r.Telemetry)
        }
        if(-not $r.Success)
        {
            Write-Log "Parallel batch POST failed: $($r.Error)" 3
            continue
        }
        Write-LogDebug "Parallel batch took $($r.Telemetry.Duration) ms, $($r.Telemetry.KB) KB, $($r.Telemetry.ObjectCount) objects ($($r.Responses.Count) responses)"
        foreach($resp in $r.Responses) { [void]$allResponses.Add($resp) }
    }

    return $allResponses
}

#endregion

#region Graph Metadata
$script:GraphMetaDataNeedsRefresh = $false
function Get-GraphMetaData
{
    # -NoDownload: use the cached file, never fetch. AppInitialized reaches here during
    # Import-Module, so once the cache aged out EVERY import paid for a 7-8 MB Graph
    # beta $metadata download - and an offline/proxied machine stalled on import for
    # something nothing had asked for. Real consumers omit it and fetch on first use.
    param([switch]$NoDownload)

    # A load-time -NoDownload call may populate the cache from an expired file so
    # validation can run offline. The first real consumer must still perform the
    # refresh promised by the normal path. Clear only that provisional stale value;
    # reset the marker before downloading so a failed refresh falls back once rather
    # than retrying the network on every metadata lookup in this process.
    $refreshProvisionalCache = (-not $NoDownload -and $script:GraphMetaDataNeedsRefresh -eq $true)
    if($refreshProvisionalCache)
    {
        $script:GraphMetaDataXML = $null
        $script:GraphMetaDataNeedsRefresh = $false
    }

    if(-not $script:GraphMetaDataXML)
    {
        # Graph metadata does not support Content-Length in response so size can not be used to check if it is updated
        # There also no other version information in response headers. Use file date to update every week
        Write-Log "Load Graph MetaData file"
        $url = "https://$(Get-GraphDomain)/beta/`$metadata"
        $fileFullPath = Join-Path $script:AppDataFolder "GraphMetaData.xml"
        $fi = [IO.FileInfo]$fileFullPath
        $maxAge = (Get-Date).AddDays(-14)
        $cacheIsFresh = ($fi.Exists -and ($fi.LastWriteTime -gt $maxAge -or $fi.CreationTime -gt $maxAge))
        # -NoDownload ignores the age limit: there is no refresh to fall back to, and
        # stale schema is fine for the load-time sanity check that passes it.
        if($fi.Exists -and ($NoDownload -or $cacheIsFresh))
        {
            try
            {
                # -Raw reads the whole file as one string; without it Get-Content
                # splits the 7-8 MB metadata into a per-line string array before
                # the [xml] cast, which is ~6x slower (~1.1s vs ~0.2s at startup).
                [xml]$script:GraphMetaDataXML = [IO.File]::ReadAllText($fi.FullName)
                # Mark only an expired cache loaded for the non-downloading import
                # path. A subsequent normal call consumes this marker and refreshes.
                $script:GraphMetaDataNeedsRefresh = ($NoDownload -and -not $cacheIsFresh)
            }
            catch { }
        }

        if(-not $script:GraphMetaDataXML -and -not $NoDownload)
        {
            Start-DownloadFile $url $fi.FullName
            $fi.Refresh()
            if($fi.Exists) {
                try {
                    [xml]$script:GraphMetaDataXML = [IO.File]::ReadAllText($fi.FullName)
                    $script:GraphMetaDataXML.Save($fi.FullName)
                    $script:GraphMetaDataNeedsRefresh = $false
                }
                catch {
                    Write-LogError "Failed to get $($fi.Name)" $_.Exception
                }
            }
        }

        if(-not $script:GraphMetaDataXML -and $fi.Exists)
        {
            Write-Log "Using old version of Graph MetaData file" 2
            try
            {
                [xml]$script:GraphMetaDataXML = [IO.File]::ReadAllText($fi.FullName)
                # If this is the import-time offline fallback, preserve the promise
                # to refresh once a real consumer asks. After a failed normal refresh
                # the marker stays false, avoiding repeated network attempts.
                if($NoDownload) { $script:GraphMetaDataNeedsRefresh = $true }
            }
            catch { }
        }
    }
}

function Get-GraphObjectClassName
{
    param($Type)

    Get-GraphMetaData

    if(-not $script:GraphMetaDataXML) { return }

    $objectClassName = $null
    
    $nodes = $script:GraphMetaDataXML.SelectNodes("//*[@Type='Collection(graph.$($Type))']")
    if($null -ne $nodes -and $nodes.Count -gt 0)
    {
        foreach($node in $nodes)
        {
            if($node.ParentNode.Name -eq "deviceAppManagement")
            {
                $objectClassName = $node.Name
                break
            }
        }
    }

    $objectClassName
}

function Get-GraphAllEntityTypes
{
    param($EntityType, $Xml, $HashTable)

    if(-not $HashTable.ContainsKey($EntityType))
    {
        $HashTable.Add($EntityType, $Xml.SelectSingleNode("//*[name()='EntityType' and @Name='$EntityType']"))
    }

    $nodes = $Xml.SelectNodes("//*[@BaseType='graph.$EntityType']")

    foreach($node in $nodes)
    {
        if($node.Abstract -ne "true")
        {
            $HashTable.Add($node.Name, $node)
        }
        Get-GraphAllEntityTypes $node.Name $Xml $HashTable
    }    
}

function Get-GraphEntityTypeObject
{
    param($EntityType, $Xml, $SkipProperties = @())

    $props = Get-GraphEntityTypeProperties $EntityType $Xml

    if(-not $props) { return }

    $obj = [PSCustomObject]@{
        
    }

    foreach($prop in $props)
    {
        if($prop.Name -in $SkipProperties) { continue }
        $obj | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $null
    }
    $obj
}

function Get-GraphEntityTypeProperties
{
    param($EntityType, $Xml)

    Get-GraphMetaData

    if(-not $Xml) { $Xml = $script:GraphMetaDataXML }
    if(-not $Xml) { return }

    $tmpEntity = $Xml.SelectSingleNode("//*[name()='EntityType' and @Name='$EntityType']")
    if(-not $tmpEntity) { return }

    $entities = @()
    $entities += $tmpEntity
    
    while($tmpEntity.BaseType)
    {        
        $baseType = $tmpEntity.BaseType.Split('.')[-1]
        $tmpEntity = $Xml.SelectSingleNode("//*[name()='EntityType' and @Name='$baseType']")
        if($tmpEntity) 
        {
            $entities += $tmpEntity
        }
    }
    $properties = @()
    [array]::Reverse($entities)
    foreach($enitiy in $entities)
    {
        $properties += $enitiy.SelectNodes("*[name()='Property' or name()='NavigationProperty']")
    }
    
    $properties 
}

#endregion

#endregion 

#region Navigation Properties
<# # !!! ToDo: Delete
function Set-GraphNavigationPropertiesFromFile
{
    param($NavPropObject)

    if(-not $NavPropObject.File -or -not $NavPropObject.ImportedObject)
    {
        return
    }

    # Reload data from file. Some object properties was removed before import...
    $objFileInfo = Get-GraphObjectFromFile $NavPropObject.File.FileInfo.FullName

    if(-not ($objFileInfo.PSObject.Properties | Where-Object { $_.Name -like "#CustomRef_*" })) { return }

    Set-GraphNavigationProperties $NavPropObject.ImportedObject $objFileInfo $NavPropObject.File.ObjectType
}
#>

function Set-GraphNavigationProperties
{
    param($ImportedPolicy, $SourcePolicy)

    if(-not $ImportedPolicy -or -not $SourcePolicy)
    {
        return
    }

    if((Get-SettingValue "ResolveReferenceInfo") -ne $true) { return }

    $entityName = $SourcePolicy.JsonObject.'@odata.type'.Split('.')[-1]

    $graphObjectProperties = Get-GraphEntityTypeProperties $entityName
    
    foreach($navigationProperty in ($graphObjectProperties | Where-Object LocalName -eq "NavigationProperty" ))
    {
        # Is this the correct way of filter out Assignments, summaries etc.?
        if($navigationProperty.ContainsTarget -eq $true) { continue }

        if(-not ($SourcePolicy.JsonObject."$($navigationProperty.Name)@odata.associationLink")) { continue }

        $associationLink = $SourcePolicy.JsonObject."$($navigationProperty.Name)@odata.associationLink" -replace $SourcePolicy.Id,$ImportedPolicy.Id
        # Assumption that the reference object is of the same object type
        $navigationPolicyType = $ImportedPolicy.PolicyType
        $nameProp = ?? $navigationPolicyType.NameProperty "displayName"

        $refBodyObjs = @()
        $refObjName = $null

        if($navigationProperty.Type -like "Collection(*")
        {
            $multiNavProperty = $true
            $method = "POST"
        }
        else
        {
            $multiNavProperty = $false
            $method = "PUT"
        }

        if($ImportedPolicy.TenantID -and $ImportedPolicy.TenantID -eq $SourcePolicy.TenantID)
        {
            $navigationAPI = Get-GraphAPIUrl $SourcePolicy.JsonObject."$($navigationProperty.Name)@odata.navigationLink" $SourcePolicy.PolicyType.API

            $navObject = Invoke-MSGraphAPI -URL $navigationAPI -NoError -TokenId $ImportedPolicy.TokenId -ODataMetadata "minimal" 

            if($multiNavProperty)
            {
                $navProperties = $navObject.Value
            }
            else
            {
                $navProperties = $navObject
            }

            if(-not $navProperties) { 
                Write-Log "No navigation object returned based on link $($SourcePolicy.JsonObject."$($navigationProperty.Name)@odata.navigationLink")" 2
                continue
            }

            foreach($navProp in $navProperties)
            {
                $refBodyObjs += [PSCustomObject]@{                    
                    RefObjName = $navProp."$nameProp"
                    RefObjId = $navProp.Id
                    RefBody = ([PSCustomObject]@{
                        "@odata.id" = ("https://$(Get-GraphDomain)/$($ImportedPolicy.PolicyType.APIVersion)/$($ImportedPolicy.PolicyType.API)('$($navProp.Id)')")
                    })
                }
            }
        }
        else
        {
            if(-not ($SourcePolicy.JsonObject."#CustomRef_$($navigationProperty.Name)")) { continue } # Not included in the export file

            $refObjNames, $fromODataType, $policyTypeId = $SourcePolicy.JsonObject."#CustomRef_$($navigationProperty.Name)" -split "[|][:][|]" 

            foreach($refObjName in $refObjNames.Split(","))
            {            
                $refObjects = Invoke-MSGraphAPI -URL "$($ImportedPolicy.PolicyType.API)?`$filter=$($nameProp) eq '$($refObjName)'" -NoError -TokenId $ImportedPolicy.TokenId

                $objectsFound = ($refObjects.value | Where-Object '@odata.type' -eq $fromODataType | Measure-Object).Count

                if($objectsFound -eq 1)
                {
                    # Are there any references that allows multiple ref objects?
                    foreach($refObj in $refObjects.value)
                    {
                        $refBodyObjs += [PSCustomObject]@{
                            RefObjName = ?? $refObjName
                            RefObjId = $refObj.Id
                            RefBody = ([PSCustomObject]@{
                                "@odata.id" = ("https://$(Get-GraphDomain)/$($ImportedPolicy.PolicyType.APIVersion)/$($ImportedPolicy.PolicyType.API)('$($refObj.Id)')")
                            })
                        }
                    }
                }
                elseif($objectsFound -gt 1)
                {
                    Write-Log "Multiple objects ($objectsFound) found with $($ImportedPolicy.PolicyType.NameProperty) $refObjName. Skipping reference." 2
                    continue
                }
                else
                {
                    Write-Log "No object found with $($ImportedPolicy.PolicyType.NameProperty) $refObjName" 2
                    continue
                }
            }
        }

        foreach($refObject in $refBodyObjs)
        {
            Write-Log "Add $($refObject.RefObjName) ($($refObject.RefObjId)) to navigation property $($navigationProperty.Name)"
            $body = $refObject.RefBody | ConvertTo-Json -Depth 50
            $response = Invoke-MSGraphAPI -URL $associationLink -HttpMethod $method -Content $body -TokenId $ImportedPolicy.TokenId -FullResponseObject
            if($response.Success) {
                Write-LogDebug "Reference updated successfully"
            }
            else {
                Write-LogDebug "Failed to update reference" 2
            }
        }
    }    
}

# One escaping rule for every OData string literal this module puts in a URL.
#
# OData doubles a quote inside a literal. The literal is then percent-encoded so
# & # + and friends cannot split or truncate the query string - the migration
# by-name lookup once sent "displayName eq 'R&D Devices'" raw, Graph answered 400,
# -NoError swallowed it and the import created a duplicate group.
#
# [Uri]::EscapeDataString is NOT the same on both hosts: .NET Framework (PS5.1)
# leaves ' ( ) * ! unescaped, .NET (PS7) encodes them. The replaces below bring 5.1
# up to the .NET form so the URL is byte-identical on both. Graph decodes either,
# so this is about determinism (and anything keyed on the URL), not correctness.
#
# Returns the literal INCLUDING its quotes: displayName eq $(ConvertTo-ODataStringLiteral $name)
function ConvertTo-ODataStringLiteral
{
    param([string]$Value)

    $escaped = [Uri]::EscapeDataString(([string]$Value).Replace("'", "''"))
    foreach($pair in @(@("'", '%27'), @('(', '%28'), @(')', '%29'), @('*', '%2A'), @('!', '%21')))
    {
        $escaped = $escaped.Replace($pair[0], $pair[1])
    }
    return "'$escaped'"
}

function Get-GraphAPIUrl
{
    param([String]$FullUrl, [String]$API)

    $link = $FullUrl
    $x = $link.IndexOf($API, [System.StringComparison]::CurrentCultureIgnoreCase)
    if($x -gt 0) {
        return  $link.SubString($x)
    }
    else {
        Write-log "Failed to find API string '$($API)' in $FullUrl. Using full URL" 2
    }
    return $FullUrl
}

function Add-GraphNavigationProperties
{
    param($PolicyObject)
    
    if($PolicyObject.PolicyType.NavigationProperties -ne $true) { return }

    if(-not $PolicyObject.JsonObject.'@odata.type') { return }

    if((Get-SettingValue "ResolveReferenceInfo") -ne $true) { return }

    $entityName = $PolicyObject.JsonObject.'@odata.type'.Split('.')[-1]

    $props = Get-GraphEntityTypeProperties $entityName

    foreach($prop in ($props | Where-Object LocalName -eq "NavigationProperty" ))
    {
        # Is this the correct way of filter out Assignments, summaries etc.?
        if($prop.ContainsTarget -eq $true) { continue }

        if(-not ($PolicyObject.JsonObject."$($prop.Name)@odata.navigationLink")) { continue }
        $navigationPolicyType = Get-PolicyTypeFromURL $PolicyObject.JsonObject."$($prop.Name)@odata.navigationLink".Split("(")[0]
        if(-not $navigationPolicyType) { continue }
        if($navigationPolicyType -is [Array]) { 
            if(($navigationPolicyType | Where-Object Id -eq $PolicyObject.PolicyType.Id)) {
                $navigationPolicyType = $PolicyObject.PolicyType
            }
            else {
                $navigationPolicyType = $navigationPolicyType[0]
            }
            Write-LogDebug "Array returned for $($prop.Name). Using '$($navigationPolicyType.Title)'." 2 # Expected for policy types using same API
        }
        $navigationAPI = Get-GraphAPIUrl $PolicyObject.JsonObject."$($prop.Name)@odata.navigationLink" $navigationPolicyType.API

        $navProp = Invoke-MSGraphAPI -URL $navigationAPI -ODataMetadata "minimal" -NoError -TokenId $PolicyObject.TokenId
        if($navProp)
        {
            $value = $null
            $refType = ""
            if($navProp.value -is [Object[]])
            {
                if($navProp.value.Count -gt 0 -and $navProp.value[0].'@odata.type') { $refType = $navProp.value[0].'@odata.type' }
                $refValues = @()
                $navProp.value | ForEach-Object { 
                    $tmpObject = $navigationPolicyType.GetObject($_)
                    if(-not $refType) {
                        # @odata.type not returned so get full object
                        if($tmpObject.Get()) {
                            $refType = $tmpObject.JsonObject.'@odata.type'
                        }
                    }

                    $refValues += $tmpObject.Name
                }

                if($refValues.Count -gt 0)
                {
                    if(($refValues -join "") -like "*,*")
                    {
                        Write-Log "One or mor referenced objects has the comma (,) character in the name. Cannot add navigation property $($prop.Name)" 3
                    }
                    $value = ($refValues -join ",") 
                }
            }
            else
            {
                if($navProp.'@odata.type') { $refType = $navProp.'@odata.type' }
                $value = $navigationPolicyType.GetObject($navProp).Name
            }
            if($refType -and $value)
            {
                $value = ($value + "|:|" + $refType) # + "|:|" + $navigationPolicyType.Id)
                $PolicyObject.JsonObject | Add-Member -NotePropertyName "#CustomRef_$($prop.Name)" -NotePropertyValue $value
            }            
        }
    }
}

function Get-DependencyIDs
{
    param($PolicyText)

    $regExpGuid = "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
    $regExpSID = "S-1-\d{1,2}-\d{1,2}-\d{1,10}\-\d{1,10}\-\d{1,10}(?:-\d{1,10}){0,14}"
    
    $uniqueGuids = New-Object System.Collections.Generic.HashSet[String]
    $uniqueSIDs = New-Object System.Collections.Generic.HashSet[String]
    
    # Use regular expressions to extract the GUIDs
    [regex]::Matches($PolicyText, $regExpGuid) | ForEach-Object { $uniqueGuids.Add($_.Value) | Out-Null }
        
    # Use regular expressions to extract the SIDs
    [regex]::Matches($PolicyText, $regExpSID) | ForEach-Object { $uniqueSIDs.Add($_.Value) | Out-Null }

    $uniqueGuids, $uniqueSIDs
}

#endregion

#region Save Graph objects
function Invoke-SortJsonGraphObject
{
    param($Object)

    if ($null -eq $Object) { return $null }
    
    if ($Object -is [System.Collections.Hashtable]) {
        $sorted = [ordered]@{}
        
        # First, add @odata properties
        $Object.Keys | Where-Object { $_ -match '^@odata' } | Sort-Object | ForEach-Object { $sorted[$_] = $Object[$_] }
        
        # Then @ and # prefixed properties grouped by base name
        $Object.Keys | Where-Object { $_ -match '^[@#]' -and $_ -notmatch '^@odata' } | Sort-Object | ForEach-Object { $sorted[$_] = $Object[$_] }
        
        # Then regular properties
        $Object.Keys | Where-Object { $_ -notmatch '^[@#]' } | Sort-Object | ForEach-Object { 
            $value = $Object[$_]
            if ($value -is [System.Collections.Hashtable]) {
                $sorted[$_] = Invoke-SortJsonGraphObject $value
            } elseif ($value -is [array]) {
                $sorted[$_] = @($value | ForEach-Object { 
                    if ($_ -is [System.Collections.Hashtable]) { Invoke-SortJsonGraphObject $_ } else { $_ }
                })
            } else {
                $sorted[$_] = $value
            }
        }
        
        return $sorted
    } elseif ($Object -is [array]) {
        return @($Object | ForEach-Object {
            if ($_ -is [System.Collections.Hashtable]) { Invoke-SortJsonGraphObject $_ } else { $_ }
        })
    } else {
        return $Object
    }
}

function ConvertTo-JsonSortedGraph
{
    param($Object, [int]$Depth = 50, [switch]$Compress)

    # Convert to hashtable first
    if ($Object -is [PSCustomObject]) {
        $ht = @{}
        $Object.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
        $Object = $ht
    }

    # Sort the object
    $sorted = Invoke-SortJsonGraphObject $Object

    # Convert back to JSON with sorted structure
    $sorted | ConvertTo-Json -Depth $Depth -Compress:$Compress
}

# THE serializer for everything written by export - policy JSON, the
# MigrationTable and the assignment sidecars - so one setting governs all of
# them. Honours:
#
#   SortJsonProperties   alphabetical property order
#   ExportJsonFormat     indented (host-native) or compact (single line)
#
# Compact is the only format that is byte-identical on PowerShell 5.1 and 7.
# 5.1 writes two spaces after the colon and aligns nested values under the key;
# 7 uses plain two-space indentation. Anyone diffing, hashing or version-
# controlling exports across both hosts wants compact.
function ConvertTo-GraphExportJson
{
    param($Object, [int]$Depth = 50)

    $compress = ((Get-SettingValue "ExportJsonFormat" "indented") -eq "compact")

    if ((Get-SettingValue "SortJsonProperties") -eq $true) {
        return (ConvertTo-JsonSortedGraph $Object -Depth $Depth -Compress:$compress)
    }

    return ($Object | ConvertTo-Json -Depth $Depth -Compress:$compress)
}

# Encoding for files written by export, from the "Export file encoding" setting.
#
# These writes used to be a bare Out-File, whose default is UTF-16LE on
# PowerShell 5.1 and UTF-8 on 7 - and this app runs on both. The same policy
# exported from two hosts was byte-incompatible, which broke re-import and made
# every export look changed in git.
function Get-ExportFileEncoding
{
    switch ((Get-SettingValue "ExportFileEncoding" "utf8"))
    {
        "utf8bom" { return (New-Object System.Text.UTF8Encoding($true)) }
        # $false = little endian, $true = emit the BOM (what Out-File -Encoding
        # Unicode produced, so an old pipeline keeps working byte for byte).
        "unicode" { return (New-Object System.Text.UnicodeEncoding($false, $true)) }
        default   { return (New-Object System.Text.UTF8Encoding($false)) }
    }
}

# Single write path for exported text, so encoding cannot drift between the
# policy files and the migration table.
function Save-GraphTextToFile
{
    param([string]$Text, [string]$FileName)

    [IO.File]::WriteAllText($FileName, $Text, (Get-ExportFileEncoding))
}

function Save-GraphObjectToFile
{
    param($GraphObject, $FileName)

    $json = ConvertTo-GraphExportJson $GraphObject -Depth 50

    # Mask the object's OWN tenant id in the serialized JSON, not the script-global one.
    # On cross-tenant exports (object fetched from tenant A while default is tenant B),
    # the previous code masked B's id in A's data, leaving A's id in the file and producing
    # a mixed-tenant export. Internal/ExportTokens.ps1 owns that resolution so the
    # sidecars and the policy files agree on which tenant a file belongs to.
    $maskTenantId = (Get-GraphObjectOrganizationInfo $GraphObject).OrganizationId

    # Tenant id only. Sidecars carry group and filter display names, and migration
    # resolves those in the target tenant by EXACT display name (see
    # Internal/IntuneGroupMigration.ps1), so tokenizing the organization-name prefix
    # common in them ("Contoso-All-Users") would make that lookup miss the group it
    # meant to find, or match another. Internal/ExportTokens.ps1 owns the replace.
    if ($maskTenantId) {
        $json = Convert-GraphOrganizationValueToToken $json -OrganizationId $maskTenantId -Tokens "OrganizationId"
    }

    try
    {
        Save-GraphTextToFile $json $FileName
    }
    catch
    {
        Write-LogError "Failed to save file $FileName" $_.Exception
    }
}

#endregion

function Update-JsonForEnvironment
{
    param($Json, $PolicyObject, [int]$TokenId)

    if(-not $PolicyObject.TenantID) { 
        Write-Log "No Tenant ID for object $($PolicyObject.Name). Json not updated" 2
        return $Json
    }

    # One scan serves both translation passes below: every GUID and SID in the body.
    $GUIDs, $SIDs = Get-DependencyIDs $Json

    if((Get-SettingValue "ResolveReferenceInfo") -eq $true)
    {
        $dependencySourceObjects = Get-GraphDependencySourceObjects $PolicyObject
        $dependencyDestinationObjects = $null

        foreach($GUID in $GUIDs)
        {
            if(($PolicyObject.ID -and $GUID -eq $PolicyObject.ID) -or ($PolicyObject._ClonedFromObject -and $PolicyObject._ClonedFromObject.Id -eq $GUID)) { continue }

            if($dependencySourceObjects.ContainsKey($GUID) -eq $true) {
                if($null -eq $dependencyDestinationObjects) {
                    $dependencyDestinationObjects = Get-GraphDependencyIntuneObjects $PolicyObject $TokenId
                }

                $dependecyPolicy = $dependencyDestinationObjects.Values | Where-Object { $_.Name -eq $dependencySourceObjects[$GUID].Name -and $_.PolicyName -eq $dependencySourceObjects[$GUID].PolicyName }
                if(($dependecyPolicy | Measure-Object).Count -eq 1) {
                    Write-LogDebug "Replace $GUID with $($dependecyPolicy.Id) - $($dependecyPolicy.Name)"
                    $Json = $Json -replace $GUID,$dependecyPolicy.Id
                }
                elseif(($dependecyPolicy | Measure-Object).Count -gt 1) {
                    Write-Log "Multiple dependency policies fround with name '$($dependencySourceObjects[$GUID].Name)'. Cannot replace id $GUID" 3
                }
                else {
                    Write-Log "No dependency policy found with name '$($dependencySourceObjects[$GUID].Name)'. Cannot replace id $GUID" 3
                }
            }
        }
    }

    # Per-policy Entra group migration (Internal/IntuneGroupMigration.ps1): only the
    # GUIDs/SIDs that match groups the EXPORT knows about (Groups\ sidecars +
    # MigrationTable.json) are resolved in the target tenant and created when missing
    # (CreateGroupOnImport / ConvertSyncedGroupOnImport gates). Runs outside the
    # ResolveReferenceInfo gate: group translation is required for a cross-tenant
    # import to produce valid assignments regardless of the dependency-id setting.
    try
    {
        $groupMaps = Resolve-GraphMigrationGroups -Guids $GUIDs -Sids $SIDs -PolicyObject $PolicyObject -TokenId $TokenId
        foreach($sourceId in $groupMaps.IdMap.Keys)
        {
            Write-LogDebug "Replace group id $sourceId with $($groupMaps.IdMap[$sourceId])"
            $Json = $Json -replace $sourceId, $groupMaps.IdMap[$sourceId]
        }
        foreach($sourceSid in $groupMaps.SidMap.Keys)
        {
            Write-LogDebug "Replace group SID $sourceSid with $($groupMaps.SidMap[$sourceSid])"
            $Json = $Json -replace [regex]::Escape($sourceSid), $groupMaps.SidMap[$sourceSid]
        }
    }
    catch
    {
        Write-LogError "Group migration failed for $($PolicyObject.Name)" $_.Exception
    }

    #return updated json
    $json
}

#region Dependency Functions

# Key used inside the dependency cache hashtable to track which policy type IDs have been
# fully loaded (all objects fetched).  Stored alongside the policy entries so the flag
# survives cache hits and prevents redundant full-type fetches when specific-ID loads have
# already populated some objects of that type.
$script:DepFullyLoadedKey = "__FullyLoaded__"

function Get-GraphDependencySourceObjects
{
    param($PolicyObject, [switch]$DefaultPoliciesOnly)

    $scopeTagsId = "ScopeTags"

    $fileInfo = ?? $PolicyObject._ClonedFromObject.FileInfo $PolicyObject.FileInfo

    $dependencyObjects = $null

    if($fileInfo) {
        $exportPath = [IO.Path]::GetDirectoryName($fileInfo.DirectoryName)
        $cacheId = "DependencyObjects_$exportPath"
        $dependencyObjects = Get-CacheObject $cacheId
        if($dependencyObjects -is [HashTable]) { return $dependencyObjects }
        $dependencyObjects = @{}

        # Determine which type IDs this object depends on.
        # Prefer the object's own GetDependencyReferences() for precision; fall back to the
        # type-level _Dependencies list.
        $depTypeIds = @()
        $depRefs = $PolicyObject.GetDependencyReferences()
        if($depRefs.Count -gt 0) {
            $depTypeIds = @($depRefs | Select-Object -ExpandProperty TypeId -Unique)
        }
        elseif($PolicyObject.PolicyType.Dependencies) {
            $depTypeIds = @($PolicyObject.PolicyType.Dependencies)
        }

        # Scope tags are always required (integer IDs, handled via the special $scopeTagsId key)
        if($depTypeIds -notcontains $scopeTagsId) { $depTypeIds += $scopeTagsId }

        # Map type IDs to export folder names and load only those subfolders
        $depFolders = @($depTypeIds | ForEach-Object {
            $t = $script:IntuneTypes | Where-Object Id -eq $_
            if($t) { $t.Folder }
        } | Where-Object { $_ })

        if($depFolders.Count -gt 0) {
            Write-Log "Loading dependency objects from subfolders: $($depFolders -join ', ')"
            # -PolicyTypes is what makes Get-PoliciesFromFolder narrow candidates by
            # the file's parent folder. Without it resolution is global and rests on
            # each type's CheckPolicy alone - Applications only self-identifies from
            # @odata.id, which an exported file is not guaranteed to carry, so an
            # app dependency silently failed to load.
            $policies = Get-PoliciesFromFolder -Path $exportPath -SubFolders $depFolders -PolicyTypes @($script:IntuneTypes)
        }
        else {
            Write-Log "No dependency folders found for $($PolicyObject.PolicyType.Id) - skipping file dependency load"
            $policies = @()
        }

        foreach($policy in $policies) {
            if($null -eq $policy.Id) { continue }

            if($policy.Id.ToString().Length -ge 36) {
                if($dependencyObjects.ContainsKey($policy.Id) -eq $false) {
                    $dependencyObjects.Add($policy.Id, $policy)
                }
            }
            elseif($policy.PolicyType.Id -eq $scopeTagsId) {
                if($dependencyObjects.ContainsKey($scopeTagsId) -eq $false) {
                    $dependencyObjects.Add($scopeTagsId, @())
                }
                $dependencyObjects[$scopeTagsId] += $policy
            }
        }

        # The file-dependency cache is purely folder-scoped — the resolved objects come from
        # disk and have no tenant identity. Tag with FolderCache_<exportPath> so a folder-wide
        # invalidation (e.g. user re-imports the same folder) can drop them; do NOT tag with
        # TenantCache_*, which would wipe this entry on tenant disconnect for no reason.
        Set-CacheObject $cacheId $dependencyObjects "FolderCache_$exportPath"
    }
    else {
        # Map the policy's tenant to a registered token via the central registry
        # (provider-agnostic - works for MSAL / OAuth / MgGraph). Prefer the tenant
        # match; fall back to the object's own _TokenID, then the default. Missing
        # token is non-fatal: pass 0 and let Invoke-MSGraphAPI's provider routing
        # handle auth without a specific id.
        $effectiveTokenId = 0
        $tokenInfo = $null
        if($PolicyObject.TenantId) { $tokenInfo = Get-TokenInfoForTenant $PolicyObject.TenantId }
        if($tokenInfo) {
            $effectiveTokenId = $tokenInfo.Id
        }
        elseif($PolicyObject._TokenID) {
            $effectiveTokenId = [int]$PolicyObject._TokenID
        }

        $params = @{}
        if($DefaultPoliciesOnly -eq $true) {
            $params.Add("DefaultPoliciesOnly", $true)
        }
        $depRefs = $PolicyObject.GetDependencyReferences()
        if($depRefs.Count -gt 0) {
            $params.Add("DependencyRefs", $depRefs)
        }
        $dependencyObjects = Get-GraphDependencyIntuneObjects $PolicyObject $effectiveTokenId @params
    }

    $dependencyObjects
}

function Get-GraphDependencyIntuneObjects
{
    param($PolicyObject, $TokenId, [switch]$DefaultPoliciesOnly, [PSCustomObject[]]$DependencyRefs)

    # Get-OperationTokenInfo, not Get-TokenInfo: both callers can reach here with 0, whose
    # every-token answer would name both tenants in the DependencyObjects_<tenant> key.
    $tokeInfo = Get-OperationTokenInfo $TokenId
    if(-not $tokeInfo) { return }

    # Per-tenant re-entry guard. The old script:GraphGetDependencyPolicies flag was
    # process-global which silently swallowed concurrent calls for different tenants.
    if($null -eq $script:GraphGetDependencyPoliciesByTenant) {
        $script:GraphGetDependencyPoliciesByTenant = @{}
    }
    if($script:GraphGetDependencyPoliciesByTenant[$tokeInfo.TenantId] -eq $true) { return }

    try {
        $script:GraphGetDependencyPoliciesByTenant[$tokeInfo.TenantId] = $true

        $scopeTagsId = "ScopeTags"

        $filterPolicyTypes = @()
        if($DefaultPoliciesOnly -eq $true) { $filterPolicyTypes += $scopeTagsId }

        $cacheId = "DependencyObjects_$($tokeInfo.TenantId)"
        $dependencyObjects = Get-CacheObject $cacheId @{}

        # Retrieve or create the set of fully-loaded type IDs stored inside the cache entry.
        # "Fully loaded" means every object of that type was fetched; a partial load (specific
        # IDs only) does NOT mark the type as fully loaded, so later callers can still load
        # additional specific objects of the same type.
        if(-not $dependencyObjects.ContainsKey($script:DepFullyLoadedKey)) {
            $dependencyObjects[$script:DepFullyLoadedKey] = [System.Collections.Generic.HashSet[string]]::new()
        }
        $fullyLoadedTypes = $dependencyObjects[$script:DepFullyLoadedKey]

        # Determine which type IDs are required for this object
        $requiredTypeIds = @()
        if($DependencyRefs.Count -gt 0) {
            $requiredTypeIds = @($DependencyRefs | Select-Object -ExpandProperty TypeId -Unique)
        }
        elseif($PolicyObject.PolicyType.Dependencies) {
            $requiredTypeIds = @($PolicyObject.PolicyType.Dependencies)
        }

        # Scope tags are always required
        if($requiredTypeIds -notcontains $scopeTagsId) { $requiredTypeIds += $scopeTagsId }

        # Separate type IDs into: those with specific object IDs (targeted fetch) and those
        # that need a full type load.  Scope tags always use full load (integer IDs, not GUIDs).
        $fullLoadTypes = @()

        foreach($depTypeId in $requiredTypeIds) {
            if(-not ($script:IntuneTypes | Where-Object Id -eq $depTypeId)) { continue }
            if($filterPolicyTypes.Count -gt 0 -and $filterPolicyTypes -notcontains $depTypeId) { continue }
            if($fullyLoadedTypes.Contains($depTypeId)) { continue }

            $specificRefs = @($DependencyRefs | Where-Object { $_.TypeId -eq $depTypeId -and $_.Id })

            if($specificRefs.Count -gt 0 -and $depTypeId -ne $scopeTagsId) {
                # Fetch only the specific objects whose IDs are not yet in cache
                $depType = $script:IntuneTypes | Where-Object Id -eq $depTypeId
                foreach($ref in $specificRefs) {
                    if($dependencyObjects.ContainsKey($ref.Id)) { continue }
                    Write-Log "Fetching specific dependency: $depTypeId / $($ref.Id)"
                    $rawObj = Invoke-MSGraphAPI -Url "$($depType.API)/$($ref.Id)" -TokenId $TokenId
                    if($rawObj -and $rawObj.id) {
                        $depObj = $depType.GetObject($rawObj)
                        if($depObj) {
                            $depObj._TokenID = $tokeInfo.Id
                            $depObj.TenantId = $tokeInfo.TenantId
                            $dependencyObjects.Add($depObj.Id, $depObj)
                        }
                    }
                }
            }
            else {
                $fullLoadTypes += $depTypeId
            }
        }

        # Full-type loads (no specific IDs, or scope tags which use integer IDs)
        foreach($dependencyType in $fullLoadTypes) {
            if($dependencyObjects.ContainsKey($dependencyType)) { continue }

            $dependencyPolicyObjects = Get-GraphPolicies -PolicyType $dependencyType -TokenId $TokenId

            foreach($dependencyPolicy in $dependencyPolicyObjects) {
                if($dependencyPolicy.PolicyType.Id -eq $scopeTagsId) {
                    if($dependencyObjects.ContainsKey($scopeTagsId) -eq $false) {
                        $dependencyObjects.Add($scopeTagsId, @())
                        $dependencyObjects[$scopeTagsId] += [PSCustomObject]@{ ID = 0; Name = "Default" }
                    }
                    $dependencyObjects[$scopeTagsId] += $dependencyPolicy
                }
                else {
                    if(-not $dependencyObjects.ContainsKey($dependencyPolicy.Id)) {
                        $dependencyObjects.Add($dependencyPolicy.Id, $dependencyPolicy)
                    }
                }
            }
            $fullyLoadedTypes.Add($dependencyType) | Out-Null
        }

        # Persistent: this cache holds ScopeTags/Filters that are preloaded on auth and
        # should NOT be timed out mid-flow. Explicit invalidation happens on tenant
        # disconnect via Clear-TenantCache.
        Set-CacheObject $cacheId $dependencyObjects "TenantCache_$($tokeInfo.TenantId)" -Persistent
    }
    finally {
        $script:GraphGetDependencyPoliciesByTenant[$tokeInfo.TenantId] = $false
    }
    $dependencyObjects
}

# Force a per-tenant preload of types that are always required by Compare/Import flows.
# Today: ScopeTags and Filters (assignmentFilters). Called from
# Invoke-MSALUIEventNewAuthentication so the dependency cache is warm before any
# user-driven Compare or Import work.
function Initialize-TenantDependencyCache
{
    param($TokenId)

    $tokeInfo = Get-TokenInfo $TokenId
    if(-not $tokeInfo) { return }

    $cacheId = "DependencyObjects_$($tokeInfo.TenantId)"
    $dependencyObjects = Get-CacheObject $cacheId @{}
    if(-not $dependencyObjects.ContainsKey($script:DepFullyLoadedKey)) {
        $dependencyObjects[$script:DepFullyLoadedKey] = [System.Collections.Generic.HashSet[string]]::new()
    }
    $fullyLoadedTypes = $dependencyObjects[$script:DepFullyLoadedKey]

    $scopeTagsId = "ScopeTags"
    $filtersId   = "AssignmentFilters"

    # Only preload the types not already cached, and that are actually registered.
    $toLoad = @(@($scopeTagsId, $filtersId) | Where-Object {
        -not $fullyLoadedTypes.Contains($_) -and ($script:IntuneTypes | Where-Object Id -eq $_)
    })
    if($toLoad.Count -eq 0) { return }

    Write-Log "Preloading dependency type(s) '$($toLoad -join ", ")' for tenant $($tokeInfo.TenantName)"

    # One Get-GraphPolicies call with multiple types coalesces into a single $batch
    # request (one round-trip) instead of one list call per type. It returns every
    # type's policies in one array, each wrapped with its own PolicyType, so the loop
    # below still routes ScopeTags to their bucket and everything else by Id.
    $loaded = Get-GraphPolicies -PolicyType $toLoad -TokenId $TokenId

    foreach($depPolicy in $loaded) {
        if($depPolicy.PolicyType.Id -eq $scopeTagsId) {
            if($dependencyObjects.ContainsKey($scopeTagsId) -eq $false) {
                $dependencyObjects.Add($scopeTagsId, @())
                $dependencyObjects[$scopeTagsId] += [PSCustomObject]@{ ID = 0; Name = "Default" }
            }
            $dependencyObjects[$scopeTagsId] += $depPolicy
        }
        elseif($depPolicy.Id) {
            if(-not $dependencyObjects.ContainsKey($depPolicy.Id)) {
                $dependencyObjects.Add($depPolicy.Id, $depPolicy)
            }
        }
    }
    foreach($t in $toLoad) { $fullyLoadedTypes.Add($t) | Out-Null }

    Set-CacheObject $cacheId $dependencyObjects "TenantCache_$($tokeInfo.TenantId)" -Persistent
}

# Wipe every cache entry tagged for a given tenant. Called on disconnect / tenant switch
# so stale dependency, AAD, baseline-template, and ADMX objects don't leak across
# tenants.
function Clear-TenantCache
{
    param($TenantId)

    if(-not $TenantId) { return }
    Write-Log "Clearing tenant cache for $TenantId"
    Clear-CacheObject -Tags "TenantCache_$TenantId" -Force
    if($null -ne $script:GraphGetDependencyPoliciesByTenant) {
        $script:GraphGetDependencyPoliciesByTenant.Remove($TenantId) | Out-Null
    }
}

function Get-GraphTranslatedDependencyObject
{
    param($ObjectId, $SourcePolicy, $Policy, $PolicyType)

    $sourceDependecnyObjects = Get-GraphDependencySourceObjects $SourcePolicy
    $environmentDependecnyObjects = Get-GraphDependencySourceObjects $Policy

    if($PolicyType -eq "ScopeTags") {
        $sourceScopeTags = $sourceDependecnyObjects["ScopeTags"]
        $environmentScopeTags = $environmentDependecnyObjects["ScopeTags"]

        $sourceObject = $sourceScopeTags | Where-Object Id -eq $ObjectId
        if(-not $sourceObject) {
            Write-Log "No Scope Tag with Id $ObjectId found in the source environment" 2
            return
        }
        if($sourceScopeTags -and $environmentScopeTags) {
            $environmentObject = $environmentScopeTags | Where-Object Name -eq $sourceObject.Name
            if(-not $environmentObject) {
                Write-Log "No Scope Tag with Id $ObjectId found in the source environment" 2
                return
            }
            Write-LogDebug "Found $($environmentObject.Name) with Id $($environmentObject.Id) based on $($sourceObject.Id) in Source environment"
            return $environmentObject
        }
        else {
            Write-Log "Could not get Scope tags from Source or Destination environment" 3
            return
        }
    }
    else {
        
    }
}

#endregion

#Region Remove Graph object properties
function Remove-GraphPropertiesForImport
{
    param(
        [PSCustomObject]
        $PolicyObject, 
        [PSCustomObject]
        $PropertyObject,
        [String[]]
        $KeepProperties)

    if($PolicyObject.PolicyType.SkipRemovingProperties -eq $true) { return }

    $removeProperties = @()

    if($PolicyObject.PolicyType.PropertiesToRemove)
    {
        $removeProperties += $PolicyObject.PolicyType.PropertiesToRemove
    }

    if($removeProperties.Count -eq 0 -or $PolicyObject.PolicyType.SkipRemoveDefaultProperties -ne $true)
    {
        # Default properties to delete
        $removeProperties += @('lastModifiedDateTime','createdDateTime','supportsScopeTags','id','modifiedDateTime')
    }

    # Remove OData properties
    foreach($odataProp in ($PropertyObject.PSObject.Properties | Where-Object { $_.Name -like "*@Odata*Link" -or $_.Name -like "*@odata.context" -or $_.Name -like "*@odata.id" -or ($_.Name -like "*@odata.type" -and $_.Name -ne "@odata.type")})) # -or $_.Name -like "#CustomRef*"
    {        
        $removeProperties += $odataProp.Name
    }

    foreach($propertyToRemove in $removeProperties)
    {
        # Allow override deleting default propeties e.g. some object types requires the Id property
        if(($propertyToRemove -in $PolicyObject.PolicyType.SkipRemoveProperties) -or ($propertyToRemove -in $KeepProperties)) { continue }
        Remove-Property $PropertyObject $propertyToRemove
    }

    if($PolicyObject.PolicyObject.SkipRemovingChildProperties -ne $true)
    {
        foreach($prop in ($PropertyObject.PSObject.Properties))
        {
            if($PropertyObject."$($prop.Name)"."@odata.type")
            {
                foreach($childObj in ($PropertyObject."$($prop.Name)"))
                {
                    Remove-GraphPropertiesForImport  $PolicyObject $childObj $KeepProperties
                }
            }
        }
    }
}    

#endregion

# Application type/name/platform helpers (Get-GraphApplicationName / -Type /
# -Platform / -TypeGroup) moved to Internal/IntuneAppManagement.ps1 on 2026-06-21:
# they are application-specific, and MSGraph.ps1 is generic-only (R4).

function Get-GraphScopeTags
{
    param($PolicyObject)

    $scopeTagsId = "ScopeTags"

    $scopeTags = @()

    $dependencyObjects = Get-GraphDependencySourceObjects $PolicyObject -DefaultPoliciesOnly
    if($dependencyObjects -and $dependencyObjects.ContainsKey($scopeTagsId)) {
        
        foreach($scopeTagId in ($PolicyObject.Object."$($PolicyObject.PolicyType.ScopeTagProperty)")) {            
            $scoppeTagPolicy = $dependencyObjects[$scopeTagsId] | Where-Object Id -eq $scopeTagId
            if($scoppeTagPolicy) {
                $scopeTags += $scoppeTagPolicy.Name
            }
            else {
                Write-Warning "No Scope Tag object found with id $scopeTagId"
            }
        }
    }
    else {

    }
    $scopeTags
}

function Add-GraphAssignmentsToObject
{
    param($PolicyObject, $SourceObject)

    # AutoPilot and TaC are using assignments and not assign like other object types
    $api = "$($PolicyObject.PolicyType.API)/$($PolicyObject.Id)/assignments"

    $assignments = $SourceObject.JsonObject.assignments

    # These profiles don't support importing of multiple assignments with { "assignment" [...]}
    # Each assignment must be imported separately 

    foreach($assignment in $assignments)
    {
        if($assignment.Source -and $assignment.Source -ne "direct") { continue }

        foreach($prop in $assignment.PSObject.Properties)
        {
            if($prop.Name -in @("Target")) { continue }
            Remove-Property $assignment $prop.Name
        }

        foreach($prop in $assignment.target.PSObject.Properties)
        {
            if($prop.Name -in @("@odata.type","groupId")) { continue }
            Remove-Property $assignment.target $prop.Name
        }

        $json = Update-JsonForEnvironment ($assignment | ConvertTo-Json -Depth 20) $PolicyObject $PolicyObject.TokenId
        $response = Invoke-MSGraphAPI -Url $api -Body $json -Method "POST" -TokenId $PolicyObject.TokenId -FullResponseObject
        if($response.Success) {
            Write-LogDebug "Assignment added successfully"
        }
        else {
            Write-LogDebug "Failed to add assignments" 2
        }
    }

    @{"Import"=$false}
}

#endregion
