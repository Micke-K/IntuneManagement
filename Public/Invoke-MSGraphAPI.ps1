function Invoke-MSGraphAPI {
    param (
        [Parameter(Mandatory)]
        [String]
        $Url,

        [Alias("Body")]
        [String]
        $Content,

        [HashTable]
        $Headers,

        [ValidateSet("GET", "POST", "OPTIONS", "DELETE", "PATCH", "PUT")]
        [Alias("Method")]
        [String]
        $HttpMethod = "GET",

        [HashTable]
        $AdditionalHeaders,

        [string]
        $Outfile = "",

        [Switch]
        $SkipAuthentication,

        [ValidateSet("full", "minimal", "none", "skip")]
        [String]
        $ODataMetadata = "full",

        [ValidateSet("beta", "v1.0")]
        [String]
        $GraphVersion = "",

        [switch]
        $AllPages,

        [int]
        $PageSize = -1,

        [switch]
        $Batch,

        [switch]
        $NoError,

        [Int]
        $TokenId = 0,

        [Switch]
        $FullResponseObject
    )

    if ($null -eq $tokenId -or $tokenId -eq 0) {
        $TokenId = Get-DefaultTokenId
    }

    # Token acquisition goes through the token's OWNING provider, resolved from the
    # central registry by id. This is what lets several environments be live at once
    # across different providers (tenant A on MSAL, tenant B on OAuth) and have each
    # call reach the right one - routing by the active provider alone could not. For
    # id 0 / an unregistered id we fall back to the active provider (headless
    # -SkipAuthentication internal calls and pre-registry states rely on this).
    # GetAccessToken handles its own session lookup, expiry pre-flight, and silent
    # refresh; each provider has its own session model.
    $authProvider = Resolve-AuthTokenProvider $TokenId
    if (-not $authProvider) {
        $authProvider = Get-AuthProvider
    }
    if (-not $authProvider) {
        Write-Log "No authentication provider is active. Cannot invoke Graph API." 3
        return
    }

    $graphDomain = Get-GraphDomain $TokenId
    $graphResource = "https://$graphDomain"

    # Always ask the provider for an access token so the Authorization header is set.
    # The historical $SkipAuthentication flag means "I'm already inside an auth flow,
    # do not trigger another active re-auth" — NOT "do not send a token." When that
    # flag is set we still need the cached token; we just don't bail if it's missing.
    $accessToken = $authProvider.GetAccessToken($TokenId, $graphResource)

    # An expired token the provider could not silently refresh is as useless as no
    # token - sending it just produces a 401 (and, worse, a doomed /ME during the
    # AuthenticationFailed handler). Unless this is an internal auth-flow call
    # (-SkipAuthentication), treat "expired" the same as "missing": null it out so the
    # guard below aborts cleanly instead of firing the request. The provider already
    # fires AuthenticationFailed on the refresh miss, so the UI reverts to Sign-in on
    # its own. GetAccessTokenExpiry returns MaxValue for SDK-managed providers
    # (MgGraph) and unknown expiries, so this never blocks those.
    if ($accessToken -and $SkipAuthentication -ne $true) {
        try {
            $tokenExpiry = $authProvider.GetAccessTokenExpiry($TokenId, $graphResource)
            if ($tokenExpiry -ne [datetime]::MaxValue -and $tokenExpiry -le (Get-Date)) {
                Write-Log "Access token for TokenId $TokenId is expired and could not be refreshed - skipping Graph call ($Url)" 2
                $accessToken = $null
            }
        }
        catch { }
    }

    if (-not $accessToken -and $SkipAuthentication -ne $true) {
        Write-Log "Could not obtain a valid access token from provider '$($authProvider.Id)' for TokenId $TokenId" 3
        return
    }

    if (-not $GraphVersion) {
        if (-not $script:defaultVersion) {
            if ((Get-SettingValue "UseGraphV1") -eq $true) {
                $script:defaultVersion = "v1.0"
            }
            else {
                $script:defaultVersion = "beta"
            }
        }
        $GraphVersion = $script:defaultVersion 
    }

    $Params = @{}

    $requestId = [Guid]::NewGuid().guid

    if (-not $Headers) {
        $Headers = @{
            'Content-Type'           = 'application/json; charset=utf-8'
            'x-ms-client-request-id' = $requestId
        }
        if ($accessToken) {
            $Headers['Authorization'] = "Bearer $accessToken"
        }
    }

    if ($HttpMethod -eq "GET" -and $ODataMetadata -ne "Skip") {
        # Note: odata.metadata=full in Accept 
        # @odata.type is not always included with default (minimum). 
        # That is required to identify the object type in some functions
        # It does include a lot of info we don't need... 
        $Headers.Add("Accept", "application/json;odata.metadata=$ODataMetadata")
    }
    #elseif($Content)
    #{
    #    # Upload content as UTF8 to support international and extended characters
    #    $Content = [System.Text.Encoding]::UTF8.GetBytes($Content)
    #}

    if ($AdditionalHeaders -is [HashTable]) {
        foreach ($key in $AdditionalHeaders.Keys) {
            if ($Headers.ContainsKey($key)) { continue }

            $Headers.Add($key, $AdditionalHeaders[$key])
        }
    }

    # Multi Admin Approval: tenants with an access policy hold app-auth writes for a
    # second admin. Graph wants a base64 justification header on the first attempt;
    # a caller resubmitting an already-approved request passes 'x-msft-approval-code'
    # through -AdditionalHeaders instead. Never send both - the approval code wins.
    # GET is never gated, so this only touches write verbs.
    if ($HttpMethod -in @("POST", "PATCH", "PUT", "DELETE") -and
        -not $Headers.ContainsKey($script:MSGraphApprovalCodeHeader) -and
        -not $Headers.ContainsKey($script:MSGraphApprovalJustifyHeader)) {
        $maaJustification = ConvertTo-MSGraphApprovalJustification (Get-SettingValue "MultiAdminApprovalJustification")
        if ($maaJustification) {
            $Headers[$script:MSGraphApprovalJustifyHeader] = $maaJustification
        }
    }

    if ($Content) { $Params.Add("Body", [System.Text.Encoding]::UTF8.GetBytes($Content)) }
    if ($Headers) { $Params.Add("Headers", $Headers) }
    if ($Outfile) {
        $dirName = [IO.Path]::GetDirectoryName($Outfile)
        try {
            [IO.Directory]::CreateDirectory($dirName) | Out-Null
        }
        catch {
            
        }
        if ([IO.Directory]::Exists($dirName)) {
            $Params.Add("OutFile", $OutFile)
            $Params.Add("PassThru", $true)
        }
        else {
            Write-Log "Failed to create directory for OutFile $Outfile" 3
            return
        }
    }

    if (($Url -notmatch "^http://|^https://")) {
        $Url = "https://$graphDomain/$GraphVersion/" + $Url.TrimStart('/')
    }

    # Resolve %OrganizationId% from the active provider's view of this token's tenant.
    # Phase 3: was reading MSAL globals directly; now goes through GetUserInfo so the
    # MgGraph provider works correctly too. Falls back to the script global (default
    # tenant snapshot) only if the provider returns nothing.
    if ($Url -match "%OrganizationId%") {
        $callTenantOrgId = $null
        try {
            $userInfo = $authProvider.GetUserInfo($TokenId)
            if ($userInfo -and $userInfo.TenantId) { $callTenantOrgId = $userInfo.TenantId }
        }
        catch { }
        if (-not $callTenantOrgId) { $callTenantOrgId = (Get-CurrentTenantId) }
        $Url = $Url -replace "%OrganizationId%", $callTenantOrgId
    }

    $uri = [uri]$Url

    if ($PageSize -gt 0 -and $uri.Query.IndexOf("`$top=") -eq -1 -and $uri.Segments[-1] -eq '$batch') {
        if (($url.IndexOf('?')) -eq -1) {
            $url = "$($url.Trim())?"
        }
        else {
            $url = "$($url.Trim())&"
        }
        $url = "$($url.Trim())`$top=$($PageSize)"
        Write-LogDebug "Use page size $PageSize"
    }

    $proxyURI = Get-ProxyURI
    if ($proxyURI) {
        $Params.Add("proxy", $proxyURI)
        #$Params.Add("UseBasicParsing", $true)
    }

    # Each counter is incremented ONLY by its own status class, in the retry branch
    # below. An earlier version also bumped $retryCount for every caught exception,
    # which made each 429 cost two of the ten and silently halved the budget.
    # Compared with -lt, so the budget is exactly $retryMax retries - the same
    # arithmetic Register-GraphRetryAttempt uses for the batch path.
    $retryCount = 0
    $retryMax = 10
    # Server errors are capped far below throttling - see the same split in
    # Invoke-GraphBatchRequest. Ten rounds of 10 s back-off on a 500 that will
    # never clear is 100 s of frozen app for a request that cannot succeed.
    $serverErrorRetryMax   = 3
    $serverErrorRetryCount = 0
    # CAE claims-challenge retry is one-shot — if the new token still gets a 401 we
    # let the error surface instead of spinning. Per-call flag, not per-loop.
    $claimsRetryDone = $false

    $returnValue = [PSCustomObject]@{
        Content           = $null
        StatusCode        = $null
        StatusDescription = $null
        Success           = $false
        ErrorCode         = $null
        ErrorMessage      = $null
        ErrorRequestId    = $null
        ErrorClientRequestId = $null
        ErrorDate         = $null
        ErrorContent      = $null
        ErrorRawContent   = $null
        # Multi Admin Approval outcome. ApprovalPending means the write was accepted
        # and is queued for a second admin - callers should report "pending", not
        # "failed". ApprovalCode is the id to resubmit with once approved.
        ApprovalPending   = $false
        ApprovalCode      = $null
        ApprovalAdvice    = $null
    }

    do {
        $retryRequest = $false

        if(-not $script:AllGraphCalls) {
            $script:AllGraphCalls = [System.Collections.Generic.List[PSCustomObject]]::new()
        }

        $webRequestInfo = [PSCustomObject]@{
            ID = $requestId
            URL = $Url
            Method = $HttpMethod
            IsBatch = ($uri.Segments[-1] -eq '$batch')
            BatchRequests = @()
            StatusCode = $null
            Time = Get-Date
            Duration = $null
            KB = 0.0
            ObjectCount = 0
            PageCount = 0
            ErrorMessage = $null
            ErrorCode = $null
            ErrorRequestId = $null
            ErrorClientRequestId = $null
            # Phase 3: track which auth provider minted the token for this call so the
            # Graph Calls log can show MSAL vs MgGraph at a glance.
            Provider = $authProvider.Id
        }
        $totalBytes = [long]0

        # Cap the in-memory call log to avoid unbounded growth (was O(n²) with array `+=`).
        if($script:AllGraphCalls.Count -ge 2000) { $script:AllGraphCalls.RemoveAt(0) }
        [void]$script:AllGraphCalls.Add($webRequestInfo)

        if($null -eq $script:IsPSv7) { $script:IsPSv7 = $PSVersionTable.PSVersion.Major -ge 7 }

        try {
            Write-LogDebug "Invoke graph API: $Url (Request ID: $requestId)"
            $allValues = [System.Collections.Generic.List[object]]::new()
            $stopwatch = [System.Diagnostics.Stopwatch]::new()
            do {
                if($script:IsPSv7 -and -not $Params.ContainsKey("ProgressAction")) {
                    $Params.Add("ProgressAction", "SilentlyContinue")
                }

                $Params["Uri"]    = $Url
                $Params["Method"] = $HttpMethod
                # One-per-second endpoints (Conditional Access, named locations, identity
                # protection): wait out the tenant's interval before this request goes out.
                $paceClass = Get-GraphRateClass -Url $Url
                if($paceClass) { Wait-GraphRatePace -Class $paceClass -TokenId $TokenId }
                # Bound every request so a hung/stalled call can't freeze the UI thread
                # indefinitely (the whole app is single-threaded-with-pump). Applies to
                # the -OutFile photo download too, since that shares $Params.
                if(-not $Params.ContainsKey("TimeoutSec")) {
                    $reqTimeout = Get-SettingValue "MSGraphRequestTimeoutSec"
                    if(-not $reqTimeout -or [int]$reqTimeout -le 0) { $reqTimeout = 100 }
                    $Params["TimeoutSec"] = [int]$reqTimeout
                }
                $response = $null

                $stopwatch.Restart()
                # Provider-routed request: when the owning provider couldn't yield a raw
                # bearer (e.g. an SDK-backed provider whose in-memory token cache is
                # opaque), let it run the request itself and return a response shaped like
                # Invoke-WebRequest's. $null means "not routed - use the raw token". This
                # replaces a hardcoded provider-Id check with the InvokeWebRequest capability.
                if ($authProvider -and (-not $accessToken -or $authProvider.RoutesAllRequests)) {
                    $response = $authProvider.InvokeWebRequest($Url, $HttpMethod, $Content, $Headers)
                }
                if ($null -eq $response) {
                    $response = Invoke-WebRequest @Params -UseBasicParsing -ErrorAction Stop
                }
                $stopwatch.Stop()
                $webRequestInfo.Duration = $stopwatch.Elapsed.TotalMilliseconds

                # Track bytes received. Invoke-WebRequest exposes RawContentLength as a long.
                # Fall back to Content.Length if RawContentLength isn't set (rare).
                try {
                    if($response.RawContentLength -gt 0) { $totalBytes += [long]$response.RawContentLength }
                    elseif($response.Content) { $totalBytes += [long]$response.Content.Length }
                } catch {}

                $contentObject = $response.Content | ConvertFrom-Json -ErrorAction Stop
                # Count successful pages only — incrementing before the request would
                # double-count when 429/CAE retries re-enter the outer retry loop and
                # re-run the inner pagination loop. Increment lives after the parse so a
                # parse failure also doesn't inflate the count.
                $webRequestInfo.PageCount++

                $returnValue.Content = $contentObject
                $returnValue.StatusDescription = $response.StatusDescription
                $returnValue.StatusCode = $response.StatusCode
                $returnValue.Success = $true
                $webRequestInfo.StatusCode = $response.StatusCode

                # Count returned objects from this page.
                #  - List endpoints return { "value": [...] } -> count the array.
                #  - Single-entity endpoints like /me or /users/{id} return the object directly -> count as 1.
                #  - Batch envelopes are handled separately below (don't double-count here).
                if($contentObject.value -is [Array]) {
                    $webRequestInfo.ObjectCount += $contentObject.value.Count
                }
                elseif(-not $webRequestInfo.IsBatch -and $null -ne $contentObject) {
                    $webRequestInfo.ObjectCount += 1
                }

                # For batch calls, populate BatchRequests with each request item + its response
                # status code, KB, and ObjectCount so the Graph log UI can show the per-item breakdown.
                if($webRequestInfo.IsBatch -and $Content) {
                    try {
                        $requestEnvelope = $Content | ConvertFrom-Json -Depth 20
                        if($requestEnvelope.requests) {
                            $perItemById = @{}
                            $batchObjectTotal = 0
                            foreach($r in $contentObject.responses) {
                                $rid = "$($r.id)"
                                $itemBytes = 0
                                $itemCount = 0
                                if($null -ne $r.body) {
                                    try {
                                        # Approximate per-item byte size by serializing the body
                                        # (response is JSON; raw bytes-over-wire aren't broken out
                                        # per item by the $batch envelope).
                                        $itemBytes = ($r.body | ConvertTo-Json -Depth 20 -Compress).Length
                                    } catch {}
                                    # Count: value array -> array length; single-object success -> 1; error body -> 0.
                                    if($r.body.value -is [Array]) {
                                        $itemCount = $r.body.value.Count
                                    }
                                    elseif($r.status -ge 200 -and $r.status -lt 300) {
                                        $itemCount = 1
                                    }
                                }
                                $batchObjectTotal += $itemCount
                                $perItemById[$rid] = [PSCustomObject]@{
                                    StatusCode  = $r.status
                                    KB          = [Math]::Round($itemBytes / 1024.0, 1)
                                    ObjectCount = $itemCount
                                    PageCount   = 1
                                }
                            }
                            $webRequestInfo.ObjectCount += $batchObjectTotal

                            $items = [System.Collections.Generic.List[PSCustomObject]]::new()
                            foreach($req in $requestEnvelope.requests) {
                                $rid = "$($req.id)"
                                $resp = $perItemById[$rid]
                                if(-not $resp) { $resp = [PSCustomObject]@{ StatusCode = $null; KB = 0.0; ObjectCount = 0 } }
                                [void]$items.Add([PSCustomObject]@{
                                    Id       = $req.id
                                    Method   = $req.method
                                    URL      = $req.url
                                    Response = $resp
                                })
                            }
                            $webRequestInfo.BatchRequests = $items.ToArray()
                        }
                    }
                    catch {
                        # Telemetry-only; never let it break the API call.
                    }
                }

                $webRequestInfo.KB = [Math]::Round($totalBytes / 1024.0, 1)
                Write-LogDebug "Invoke-WebRequest took $($webRequestInfo.Duration) ms, $($webRequestInfo.KB) KB, $($webRequestInfo.ObjectCount) objects, page $($webRequestInfo.PageCount) ($Url)"

                if ($AllPages -eq $true -and $HttpMethod -eq "GET" -and $contentObject.value -is [Array]) {
                    foreach($v in $contentObject.value) { [void]$allValues.Add($v) }
                    if ($contentObject.'@odata.nextLink') {
                        $Url = $contentObject.'@odata.nextLink'
                    }
                    else { break }
                }
                else {
                    break
                }

            } while ($contentObject.'@odata.nextLink')

            # Assign the accumulated pages back to the returned object.
            # The previous code checked `$returnValue.Content -is [Array]`, which is never true
            # because Content is the parsed JSON object {value, @odata.nextLink}, not an array.
            if ($allValues.Count -gt 0 -and $null -ne $returnValue.Content -and $null -ne $returnValue.Content.PSObject.Properties['value']) {
                $returnValue.Content.value = $allValues.ToArray()
            }
        }
        catch {
            $webRequestInfo.Duration = ((Get-Date) - $webRequestInfo.Time).TotalMilliseconds
            try {
                $webRequestInfo.StatusCode = [int]$_.Exception.Response.StatusCode
            }
            catch{ $webRequestInfo.StatusCode = $_.Exception.Response.StatusCode }

            if ($NoError -eq $true) { return }
            # CAE claims challenge: Graph returns 401 with WWW-Authenticate containing
            # claims="..." when the token must be re-acquired to satisfy a tenant policy
            # change. Re-mint the token with that challenge and retry once. Only providers
            # that manage claims challenges themselves (SupportsClaimsChallenge) need this
            # manual round-trip; SDK-managed providers handle CAE internally and report
            # $false, so the 401 surfaces unchanged.
            $claims = $null
            if (-not $claimsRetryDone -and $authProvider.SupportsClaimsChallenge -and
                [int]$_.Exception.Response.StatusCode -eq 401) {
                try {
                    $wwwAuth = $_.Exception.Response.Headers["WWW-Authenticate"]
                    if ($wwwAuth) {
                        # Header is one or more Bearer challenges separated by commas. We
                        # only care about a claims="..." parameter; capture between the
                        # first set of quotes after claims=. CIAM/ESTS may also emit it
                        # unquoted, so accept both.
                        $m = [regex]::Match($wwwAuth, 'claims="([^"]+)"')
                        if (-not $m.Success) {
                            $m = [regex]::Match($wwwAuth, 'claims=([^,\s]+)')
                        }
                        if ($m.Success) { $claims = $m.Groups[1].Value }
                    }
                }
                catch { }
            }
            if ($claims) {
                # Delegate the re-acquire to the owning provider so this caller stays
                # provider-agnostic. $allowInteractive is a caller-context decision: a
                # provider MAY prompt only when the main app window is up and this isn't a
                # nested auth-flow call; headless / automation stays silent-only and the
                # 401 surfaces if the challenge can't be satisfied silently.
                Write-Log "401 with CAE claims challenge. Re-acquiring token once." 2
                $claimsRetryDone = $true
                try {
                    $allowInteractive = [bool]($script:MainAppStarted -and $SkipAuthentication -ne $true)
                    $accessToken = $authProvider.GetClaimsToken($TokenId, $graphResource, $claims, $allowInteractive)
                    if ($accessToken) {
                        $Headers['Authorization'] = "Bearer $accessToken"
                        if ($Params.ContainsKey('Headers')) { $Params['Headers'] = $Headers }
                        $retryRequest = $true
                    }
                    else {
                        Write-Log "Claims re-acquire did not produce a token; surfacing 401 to caller (user must re-login)." 2
                    }
                }
                catch {
                    Write-LogError "Failed to re-acquire token with CAE claims" $_.Exception
                }
            }
            elseif ((([int]$_.Exception.Response.StatusCode -eq 429) -and $retryCount -lt $retryMax) -or
                    (([int]$_.Exception.Response.StatusCode -in @(500, 502, 503, 504)) -and $serverErrorRetryCount -lt $serverErrorRetryMax)) {
                # 429 = throttling; 500/502/503/504 = transient server / gateway errors
                # (backend hiccup or upstream timeout). Both are safe to re-issue, but on
                # separate budgets: throttling clears when the window rolls over, while a
                # 5xx that repeats is usually permanent. Honor the Retry-After header when
                # Graph sends one, else default to 10s so we back off instead of hammering
                # a struggling backend, then retry the same request.
                $transientCode = [int]$_.Exception.Response.StatusCode
                # Retry-After header shape differs by PS host: PS7 surfaces a typed
                # HttpResponseHeaders.RetryAfter (RetryConditionHeaderValue); PS5.1's
                # WebException exposes a string-indexable WebHeaderCollection. Try both.
                $retryAfterRaw = $null
                try { $retryAfterRaw = $_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds } catch { }
                if($null -eq $retryAfterRaw) { try { $retryAfterRaw = $_.Exception.Response.Headers['Retry-After'] } catch { } }
                $wait = Get-GraphRetryAfterSeconds $retryAfterRaw
                if($transientCode -eq 429) { $retryCount++ } else { $serverErrorRetryCount++ }
                $retryRequest = $true
                Write-Log "$transientCode - transient error (throttling or gateway). Wait $wait s before retry" 2
                # Sliced, pumped wait so the window keeps painting (and shows the
                # countdown) instead of freezing for the whole back-off.
                Wait-UIAware -Seconds $wait -DetailFormat ("Graph {0} - retrying in {{0}}s" -f $transientCode)
            }
            else {
                $graphError = ConvertFrom-MSGraphErrorResponse $_
                $extMessage = if($graphError.Message) { ". Response message: $($graphError.Message)" } else { $null }

                $webRequestInfo.ErrorMessage = $graphError.Message
                $webRequestInfo.ErrorCode = $graphError.Code
                $webRequestInfo.ErrorRequestId = $graphError.RequestId
                $webRequestInfo.ErrorClientRequestId = $graphError.ClientRequestId

                # Classify 400/403/412 before logging. Multi Admin Approval reports a
                # queued write as 412 with outer code "BadRequest", which reads like a
                # hard failure but is the documented success path - log it as a warning
                # and hand the approval code back to the caller.
                $approvalInfo = Get-MSGraphApprovalInfo $_ $graphError $HttpMethod $Url
                $returnValue.ApprovalPending = $approvalInfo.IsApprovalPending
                $returnValue.ApprovalCode    = $approvalInfo.ApprovalCode
                $returnValue.ApprovalAdvice  = $approvalInfo.Advice

                if ($approvalInfo.IsApprovalPending) {
                    Write-Log "$Url - $($approvalInfo.Advice)" 2
                }
                else {
                    if ($approvalInfo.Advice) { Write-Log $approvalInfo.Advice 2 }
                    Write-LogError "Failed to invoke MS Graph with URL $Url (Request ID: $requestId). Status code: $($_.Exception.Response.StatusCode)$extMessage" $_.Exception
                }
                $returnValue.StatusCode = $_.Exception.Response.StatusCode
                $returnValue.StatusDescription = $extMessage
                $returnValue.ErrorCode = $graphError.Code
                $returnValue.ErrorMessage = $graphError.Message
                $returnValue.ErrorRequestId = $graphError.RequestId
                $returnValue.ErrorClientRequestId = $graphError.ClientRequestId
                $returnValue.ErrorDate = $graphError.Date
                $returnValue.ErrorContent = $graphError.Content
                $returnValue.ErrorRawContent = $graphError.RawContent
            }            
        }
    } while ($retryRequest -eq $true)
    
    #Write-Debug "$(($ret | Select-Object *))"
    
    if ($FullResponseObject -eq $true) {
        $returnValue
    }
    else {
        $returnValue.Content
    }
}
