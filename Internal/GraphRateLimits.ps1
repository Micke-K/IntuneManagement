# Pacing for the Graph endpoints that allow one request per second.
#
# Microsoft Graph limits the Conditional Access and identity-protection
# resources (conditionalAccessPolicy, namedLocation, riskDetection, riskyUser
# and the Conditional Access siblings) to ONE request per second per tenant,
# counted across every application that talks to the tenant, and it sends no
# Retry-After header when it throttles them. Every $batch sub-request counts on
# its own. A list of forty Conditional Access policies fetched twenty per
# $batch, or in parallel, therefore throttles at once and then backs off ten
# seconds at a time - which the user sees as a hang.
#
# This file keeps a "last sent" clock per (rate class, tenant) and lets a
# caller wait out the remainder of the interval before it sends. The single
# request path gates in Invoke-MSGraphAPI; the batch layer keeps paced requests
# out of parallel dispatch and posts them one per $batch. The
# GraphPaceIdentityEndpoints setting (default on) switches the whole thing off.

$script:GraphRateClasses = @(
    [PSCustomObject]@{
        Class           = 'ConditionalAccess'
        Label           = 'Conditional Access'
        IntervalSeconds = 1.0
        # Resource paths below the API version, compared case-insensitively.
        Prefixes        = @(
            'identity/conditionalAccess/policies',
            'identity/conditionalAccess/namedLocations',
            'identity/conditionalAccess/authenticationStrengths',
            'identity/conditionalAccess/authenticationContextClassReferences',
            'identityProtection/riskDetections',
            'identityProtection/riskyUsers'
        )
    }
)

# (class|tenant) -> [DateTime] UTC of the last send
$script:GraphPaceLastSend = @{}

function Test-GraphPacingEnabled
{
    return -not ((Get-SettingValue "GraphPaceIdentityEndpoints") -eq $false)
}

# The pure classifier: the rate class for a URL, or $null when the URL is not
# paced. Accepts every shape the engine produces - absolute (nextLinks),
# relative with or without a leading slash, with or without the API version,
# with query strings and OData key segments.
function Resolve-GraphRateClass
{
    param([string]$Url)

    if([string]::IsNullOrWhiteSpace($Url)) { return $null }

    $path = $Url.Trim()
    if($path -match '^https?://[^/]+(/.*)?$') { $path = "$($Matches[1])" }
    $path = ($path -split '[?#]', 2)[0]
    $path = $path.TrimStart('/')
    $path = $path -replace '^(beta|v1\.0)/', ''

    foreach($rc in $script:GraphRateClasses)
    {
        foreach($prefix in $rc.Prefixes)
        {
            if(-not $path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
            # Segment boundary: 'policies' must not match 'policiesTemplates'.
            if($path.Length -eq $prefix.Length -or $path[$prefix.Length] -eq '/' -or $path[$prefix.Length] -eq '(')
            {
                return $rc.Class
            }
        }
    }
    return $null
}

# The classifier callers use: honours the setting.
function Get-GraphRateClass
{
    param([string]$Url)

    if(-not (Test-GraphPacingEnabled)) { return $null }
    return (Resolve-GraphRateClass -Url $Url)
}

function Get-GraphRateClassInfo
{
    param([string]$Class)

    foreach($rc in $script:GraphRateClasses) { if($rc.Class -eq $Class) { return $rc } }
    return $null
}

# Mockable clock.
function Get-GraphPaceNow
{
    return [DateTime]::UtcNow
}

# The limit is per tenant across all applications, so two tokens for the same
# tenant must share one clock. Resolution mirrors the %OrganizationId%
# substitution in Invoke-MSGraphAPI; the token id is the last resort.
function Get-GraphPaceTenantKey
{
    param($TokenId = 0)

    try
    {
        $provider = Get-AuthProvider
        if($provider)
        {
            $userInfo = $provider.GetUserInfo($TokenId)
            if($userInfo -and $userInfo.TenantId) { return "$($userInfo.TenantId)" }
        }
    }
    catch { }

    try
    {
        $tenantId = Get-CurrentTenantId
        if($tenantId) { return "$tenantId" }
    }
    catch { }

    return "token:$TokenId"
}

# Pull the paced requests out of a batch queue (in place) and return them as a
# queue of their own, in their original order. Graph counts every $batch
# sub-request on its own and sends no Retry-After for these, so batching them
# twenty at a time only throttles at once. Returns an empty queue when the
# setting is off.
function Split-GraphPacedBatchRequests
{
    param(
        [System.Collections.Generic.List[PSCustomObject]]$RequestObjects,
        [string]$BatchType = 'Graph'
    )

    $paced = [System.Collections.Generic.List[PSCustomObject]]::new()
    if(-not (Test-GraphPacingEnabled)) { return ,$paced }

    $total = $RequestObjects.Count
    for($i = $RequestObjects.Count - 1; $i -ge 0; $i--)
    {
        if(Resolve-GraphRateClass -Url $RequestObjects[$i].url)
        {
            $paced.Insert(0, $RequestObjects[$i])
            $RequestObjects.RemoveAt($i)
        }
    }

    if($paced.Count -gt 0)
    {
        Write-Log "Batch $($BatchType): $($paced.Count) of $total request(s) target one-per-second endpoints - sending them one at a time after the rest"
    }
    # The comma keeps an empty or single-item list from unrolling on return.
    return ,$paced
}

# Wait until the tenant's interval for the class has passed since the last
# send, then record this send. Call it immediately before the request goes out.
function Wait-GraphRatePace
{
    param(
        [Parameter(Mandatory = $true)][string]$Class,
        $TokenId = 0
    )

    $rc = Get-GraphRateClassInfo -Class $Class
    if(-not $rc) { return }

    $key = "{0}|{1}" -f $Class, (Get-GraphPaceTenantKey -TokenId $TokenId)
    $now = Get-GraphPaceNow

    if($script:GraphPaceLastSend.ContainsKey($key))
    {
        $last      = $script:GraphPaceLastSend[$key]
        $remaining = $rc.IntervalSeconds - ($now - $last).TotalSeconds
        if($remaining -gt 0)
        {
            Write-LogDebug ("{0}: pacing - waiting {1:0.00}s before the next request to the tenant" -f $rc.Label, $remaining)
            # Sliced and pumped so the window stays alive; no status text of its
            # own - the callers show their progress.
            Wait-UIAware -Seconds $remaining -SliceMilliseconds 100
            # The send happens right after the wait: never record it earlier
            # than the interval boundary, even if the clock did not move.
            $after = Get-GraphPaceNow
            $boundary = $last.AddSeconds($rc.IntervalSeconds)
            $now = if($after -gt $boundary) { $after } else { $boundary }
        }
    }

    $script:GraphPaceLastSend[$key] = $now
}

# Record one retryable batch response against its sub-request budget and say
# whether that budget is now spent. Throttling and server errors are counted
# separately on purpose: a 429 clears when the rate window rolls over, while a
# 5xx that repeats is usually a permanent data error and re-sending it ten times
# only freezes the app. $RetryObjects is keyed by batch sub-request id.
function Register-GraphRetryAttempt
{
    param(
        [Parameter(Mandatory = $true)]$Id,
        [int]$Status,
        [Parameter(Mandatory = $true)][hashtable]$RetryObjects,
        [int]$MaxThrottleRetry    = 10,
        [int]$MaxServerErrorRetry = 3,
        # Optional tally so the caller can report the statuses it actually saw.
        [hashtable]$StatusCounts
    )

    if($null -ne $StatusCounts)
    {
        $key = "$Status"
        if(-not $StatusCounts.ContainsKey($key)) { $StatusCounts[$key] = 0 }
        $StatusCounts[$key]++
    }

    $isThrottled = ($Status -eq 429)
    if(-not $RetryObjects.ContainsKey($Id)) { $RetryObjects[$Id] = @{ Throttle = 0; ServerError = 0 } }
    $counter = $RetryObjects[$Id]
    if($isThrottled) { $counter.Throttle++ } else { $counter.ServerError++ }

    $used  = if($isThrottled) { $counter.Throttle }  else { $counter.ServerError }
    $limit = if($isThrottled) { $MaxThrottleRetry }  else { $MaxServerErrorRetry }

    return [PSCustomObject]@{ Attempts = $used; Exhausted = ($used -ge $limit) }
}

# Render a tally of retryable statuses ({"429" = 2; "500" = 1}) as "429 x2, 500".
# Used by the batch dispatcher so its back-off message names the status Graph
# actually returned.
function Format-GraphRetryStatus
{
    param([hashtable]$StatusCounts)

    if(-not $StatusCounts -or $StatusCounts.Count -eq 0) { return "a retryable status" }

    $parts = $StatusCounts.GetEnumerator() | Sort-Object { [int]$_.Key } | ForEach-Object {
        if($_.Value -gt 1) { "$($_.Key) x$($_.Value)" } else { "$($_.Key)" }
    }
    $text = ($parts -join ', ')
    if($StatusCounts.ContainsKey("429") -and $StatusCounts.Count -eq 1) { return "$text - 'Too many requests'" }
    return "status $text"
}

# ---- Retry-After ----
# Moved here from Internal/MSGraph.ps1 with the pacing work: parsing the back-off
# header Graph sends belongs beside the clock for the endpoints that never send one.
# Normalize an HTTP Retry-After header value into a wait, in seconds. Retry-After can be
# delta-seconds (e.g. "120") or an HTTP-date (e.g. "Fri, 31 Jan 2026 23:59:59 GMT"); both
# forms are honored. When the header is absent or unparseable we return $DefaultSeconds so
# the caller always backs off instead of busy-retrying a throttled/erroring backend. Shared
# by the single-request retry (Invoke-MSGraphAPI) and the per-batch-item retry in MSGraph.ps1 so both
# treat 429/5xx back-off identically.
function Get-GraphRetryAfterSeconds
{
    param(
        $RetryAfterValue,
        [int]$DefaultSeconds = 10
    )

    if($null -eq $RetryAfterValue) { return $DefaultSeconds }

    $raw = ([string]$RetryAfterValue).Trim()
    if(-not $raw) { return $DefaultSeconds }

    # delta-seconds (the form Graph almost always uses)
    $seconds = 0
    if([int]::TryParse($raw, [ref]$seconds))
    {
        if($seconds -gt 0) { return $seconds }
        return $DefaultSeconds
    }

    # HTTP-date form - wait until that instant (invariant culture; header is always GMT)
    $when = [DateTimeOffset]::MinValue
    if([DateTimeOffset]::TryParse($raw, [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$when))
    {
        $delta = [int][Math]::Ceiling(($when - [DateTimeOffset]::UtcNow).TotalSeconds)
        if($delta -gt 0) { return $delta }
        return $DefaultSeconds
    }

    return $DefaultSeconds
}
