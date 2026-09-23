#ImportOrder 26

# Microsoft.Graph PowerShell SDK implementation of AuthenticationProvider.
#
# Wraps the Microsoft.Graph.Authentication module (Connect-MgGraph, Disconnect-MgGraph,
# Get-MgContext). The SDK itself uses MSAL.NET underneath but with its own session
# state and its own on-disk token cache, separate from our MSAL provider — by design,
# per the user's decision to keep caches separate.
#
# Status:
#   * The class always registers (even without the SDK installed) so it is selectable
#     in Settings; the SDK modules are resolved / prompted-for on first Connect.
#   * `Connect-IntuneManagement -Provider MgGraph -...` routes here.
#   * Invoke-MSGraphAPI routes requests for MgGraph-owned tokens through this provider:
#     when the SDK's opaque cache can't yield a raw bearer, the request runs via the
#     provider's own pipeline (see InvokeWebRequest / Invoke-MgGraphRequestAsWebResponse).
#
# Token-extraction note: the SDK does not expose the raw access token via a public
# cmdlet. We reach into [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance
# which is the documented (in source) but undocumented (in MS Learn) accessor. In SDK
# v2 the AccessToken is a SecureString; we unprotect at the last moment.
#
# Minimum SDK version for CAE: Microsoft.Graph.Authentication 2.37.0+ is recommended.
# Earlier versions had a token-cache bug (fixed by PR #3573, May 2026) where the
# `caeEnabled: true` capability was not included when caching tokens — so a CAE
# claim-challenge round-trip could re-prompt instead of resolving silently. Older
# SDK versions still work for non-CAE flows.
class AuthenticationMgGraph : AuthenticationProvider {

    AuthenticationMgGraph() {
        $this.Id           = "MgGraph"
        $this.DisplayName  = "Microsoft Graph PowerShell SDK"

        # Required capabilities (see AuthenticationProvider contract).
        $this.SupportsInteractive       = $true
        $this.SupportsClientSecret      = $true
        $this.SupportsCertificate       = $true

        # Optional: SDK has -Identity flag for managed identity.
        $this.SupportsIdentityProvider  = $true

        # Optional: SDK accepts -AccessToken.
        $this.SupportsBYOToken          = $true

        # The SDK keeps cached accounts in a private InMemoryTokenCache byte[] (the
        # serialized MSAL-v3 cache). GetCachedAccounts() reaches in via reflection
        # and rehydrates an MSAL public-client app to enumerate the accounts —
        # source pattern: github.com/microsoftgraph/msgraph-sdk-powershell.
        # NOTE: this cache is in-memory only (NOT persisted to disk) — accounts only
        # show up within the current PowerShell session, and clicking one cannot
        # "switch to" that account because Connect-MgGraph has no -LoginHint
        # parameter. The list is informational; the user must re-Connect-MgGraph
        # to change accounts.
        $this.SupportsCachedUsers       = $true

        # SDK has no per-call tenant switching — you Disconnect and Connect with a
        # different -TenantId. The active session is single-tenant.
        $this.SupportsMultiTenant       = $false

        # The SDK refreshes internally, but a manual "Refresh" action is meaningful
        # for users — we re-trigger Connect-MgGraph (silent if the cache has a
        # valid refresh token, interactive otherwise).
        $this.SupportsRefresh           = $true

        # No way to evict a single cached account from the SDK's token cache via
        # public cmdlets. Disconnect-MgGraph clears the active session only.
        $this.SupportsForget            = $false
    }

    [void] Initialize() {
        # Module presence check happens at registration time in
        # Internal/AuthenticationMgGraphHelpers.ps1 — by the time we get here, the SDK is
        # known to be installed. We do NOT eagerly Import-Module (load cost is
        # ~hundreds of ms); the first Connect() call imports lazily.
    }

    # The SDK v2 in-memory token cache is opaque, so we can't hand Invoke-MSGraphAPI a
    # raw bearer. Route the request through the SDK's own pipeline (which auths it) and
    # wrap the result so it quacks like Invoke-WebRequest's response.
    [object] InvokeWebRequest([string]$Url, [string]$Method, [object]$Body, [hashtable]$Headers) {
        return (Invoke-MgGraphRequestAsWebResponse -Url $Url -Method $Method -Body $Body -Headers $Headers)
    }

    # Native session inspector rows for the profile "Session Info" dialog: Get-MgContext
    # properties (the closest MgGraph equivalent of MSAL's AuthenticationResult).
    [PSCustomObject[]] GetSessionInfoRows() {
        $rows = @()
        try {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if($ctx) {
                foreach($prop in ($ctx | Get-Member -MemberType Properties)) {
                    $value = $ctx."$($prop.Name)"
                    if($prop.Name -eq "Scopes" -and $value) { $value = ($value -join "`n") }
                    if($value -is [SecureString]) { $value = "<SecureString>" }
                    $rows += [PSCustomObject]@{ Name = $prop.Name; Value = $value }
                }
            }
        } catch { }
        return [PSCustomObject[]]$rows
    }

    # Silent cross-session resume. Microsoft.Graph SDK v2 persists credentials by
    # default (ContextScope.CurrentUser): the MSAL cache lives at
    # %LOCALAPPDATA%\.IdentityService\mg.msal.cache and the AuthenticationRecord
    # anchor at %USERPROFILE%\.mg\mg.authrecord.json. Both must exist; if so, a
    # plain Connect-MgGraph -NoWelcome silently rehydrates the session via
    # Azure.Identity's MsalCacheHelper. No browser, no prompt.
    [bool] TryResumeSession() {
        try {
            if(-not (Resolve-MgGraphModule)) { return $false }

            $anchor = Join-Path $env:USERPROFILE ".mg\mg.authrecord.json"
            if(-not (Test-Path $anchor)) {
                Write-LogDebug "MgGraph: no auth record at $anchor - skipping silent resume"
                return $false
            }

            try {
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop | Out-Null
            }
            catch {
                Write-LogError "MgGraph TryResumeSession: failed to import Microsoft.Graph.Authentication" $_.Exception
                return $false
            }

            # If a context is already active (e.g. another caller already connected
            # during this session), respect it.
            $existing = $null
            try { $existing = Get-MgContext -ErrorAction SilentlyContinue } catch { }
            if($existing) {
                Write-Log "MgGraph: already signed in as $($existing.Account) (tenant $($existing.TenantId)); silent resume not needed"
                return $true
            }

            Write-Log "MgGraph: attempting silent resume from persisted Azure.Identity cache..."

            # NoWelcome suppresses banner; no Scopes parameter means Azure.Identity
            # uses whatever scopes were in the AuthenticationRecord. If the cache or
            # record is stale, Connect-MgGraph will throw / require interaction —
            # we treat any failure as "resume not possible, user must click Login".
            Connect-MgGraph -NoWelcome -ErrorAction Stop | Out-Null

            $ctx = $null
            try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch { }
            if($ctx) {
                Write-Log "MgGraph: silent resume succeeded - signed in as $($ctx.Account) (tenant $($ctx.TenantId))"
                return $true
            }
            Write-Log "MgGraph: Connect-MgGraph completed but Get-MgContext returned nothing - silent resume failed" 2
            return $false
        }
        catch {
            Write-LogDebug "MgGraph: silent resume failed: $($_.Exception.Message)"
            return $false
        }
    }

    [PSCustomObject] Connect([hashtable]$Arguments) {
        Write-Log "AuthenticationMgGraph.Connect starting"

        # Step 1: ensure the required SDK module is available. If not, offer to
        # install it. If the user declines or install fails, return $null so
        # Connect-IntuneManagement can fall back to MSAL.
        if(-not (Resolve-MgGraphModule)) {
            Write-Log "Microsoft.Graph.Authentication not available - MgGraph provider cannot connect" 2
            return $null
        }

        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop | Out-Null
        }
        catch {
            Write-LogError "Failed to import Microsoft.Graph.Authentication" $_.Exception
            return $null
        }

        $mgArgs = @{ NoWelcome = $true }

        if($Arguments.TenantId) { $mgArgs['TenantId'] = $Arguments.TenantId }
        if($Arguments.AppId)    { $mgArgs['ClientId'] = $Arguments.AppId }

        # Cloud / sovereign environment selection. Prefer the flat -Cloud arg
        # (Phase 1, 2026-05-22). Fall back to translating the legacy GraphEnvironment+GCCType
        # pair so direct provider callers passing the old shape still work during the
        # deprecation window. Connect-MgGraph -Environment accepts: Global / USGov / USGovDOD / China.
        $cloudValue = if($Arguments.Cloud) { [string]$Arguments.Cloud }
                      elseif($Arguments.GraphEnvironment -or $Arguments.GCCType) {
                          Convert-LegacyToCloud -GraphEnvironment ([string]$Arguments.GraphEnvironment) -GCCType ([string]$Arguments.GCCType)
                      }
                      else { "Public" }

        $cloudEntry = Get-CloudByValue $cloudValue
        $mgEnv = $cloudEntry.MgEnvironment
        if($mgEnv -and $mgEnv -ne "Global") {
            $mgArgs['Environment'] = $mgEnv
            Write-LogDebug "MgGraph: using -Environment $mgEnv (Cloud=$cloudValue)"
        }

        # Dispatch on auth method. Connect-MgGraph parameter sets are mutually
        # exclusive, so we pick exactly one.
        if($Arguments.ContainsKey('Secret') -and $Arguments.Secret) {
            if(-not $Arguments.AppId)    { Write-Log "MgGraph: -AppId required with -Secret" 3; return $null }
            if(-not $Arguments.TenantId) { Write-Log "MgGraph: -TenantId required with -Secret" 3; return $null }
            $secStr = if($Arguments.Secret -is [SecureString]) { $Arguments.Secret }
                      else { ConvertTo-SecureString ([string]$Arguments.Secret) -AsPlainText -Force }
            $mgArgs['ClientSecretCredential'] = [PSCredential]::new($Arguments.AppId, $secStr)
            # ClientId + TenantId are conveyed via the credential here; remove the
            # standalone entries so we don't conflict with the credential parameter set.
            $mgArgs.Remove('ClientId') | Out-Null
        }
        elseif($Arguments.ContainsKey('Certificate') -and $Arguments.Certificate) {
            if($Arguments.Certificate -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) {
                $mgArgs['Certificate'] = $Arguments.Certificate
            }
            else {
                # Treat as thumbprint string
                $mgArgs['CertificateThumbprint'] = [string]$Arguments.Certificate
            }
        }
        elseif($Arguments.ContainsKey('CertificatePath') -and $Arguments.CertificatePath) {
            $cert = Get-PfxCertificate -FilePath $Arguments.CertificatePath -Password $Arguments.CertificatePassword -ErrorAction Stop
            $mgArgs['Certificate'] = $cert
        }
        elseif($Arguments.ContainsKey('Token') -and $Arguments.Token) {
            $tokStr = if($Arguments.Token -is [SecureString]) { $Arguments.Token }
                      else { ConvertTo-SecureString ([string]$Arguments.Token) -AsPlainText -Force }
            $mgArgs['AccessToken'] = $tokStr
        }
        elseif($Arguments.ContainsKey('ManagedIdentity') -and $Arguments.ManagedIdentity) {
            $mgArgs['Identity'] = $true
            # For user-assigned managed identity, the user can pass a specific client id.
            if($Arguments.ManagedIdentityClientId) { $mgArgs['ClientId'] = $Arguments.ManagedIdentityClientId }
        }
        elseif($Arguments.ContainsKey('DeviceCode') -and $Arguments.DeviceCode) {
            # Microsoft.Graph.Authentication 2.x exposes device code as the
            # -UseDeviceCode switch on Connect-MgGraph. Emits the code and
            # verification URL to the console and blocks until the user
            # completes auth in a browser on any device.
            $mgArgs['UseDeviceCode'] = $true
            if($Arguments.Scopes) { $mgArgs['Scopes'] = $Arguments.Scopes }
        }
        else {
            # Interactive. Default scopes match what the rest of the app uses; callers
            # can override via $Arguments.Scopes.
            if($Arguments.Scopes) { $mgArgs['Scopes'] = $Arguments.Scopes }
        }

        $authModeLog = if($mgArgs.ContainsKey('ClientSecretCredential')) { 'client secret' }
                       elseif($mgArgs.ContainsKey('Certificate'))        { 'certificate' }
                       elseif($mgArgs.ContainsKey('CertificateThumbprint')) { 'certificate (thumbprint)' }
                       elseif($mgArgs.ContainsKey('AccessToken'))        { 'BYO token' }
                       elseif($mgArgs.ContainsKey('Identity'))           { 'managed identity' }
                       else                                              { 'interactive (browser)' }
        Write-Log "Calling Connect-MgGraph (mode: $authModeLog)..."

        try {
            # Wipe any stale cached token from a prior session before the new auth.
            [AuthenticationMgGraph]::ClearTokenCache()

            Connect-MgGraph @mgArgs -ErrorAction Stop | Out-Null
            Write-Log "Connect-MgGraph completed successfully"
        }
        catch {
            Write-LogError "Connect-MgGraph failed" $_.Exception
            return $null
        }

        # Verify Get-MgContext returns a session. If it does, we're signed in —
        # even if we can't pull a raw bearer token out of the SDK. Invoke-MSGraphAPI
        # has an SDK-routed fallback for that case (uses Invoke-MgGraphRequest, which
        # the SDK auths internally with its own in-memory cache).
        $ctx = $null
        try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch { }
        if(-not $ctx) {
            Write-Log "Connect-MgGraph completed but Get-MgContext returned nothing. Treating as failed auth." 3
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
            [AuthenticationMgGraph]::ClearTokenCache()
            return $null
        }
        $verifyToken = $this.GetAccessToken(0, "https://$($cloudEntry.GraphHost)")
        if($verifyToken) {
            Write-LogDebug "MgGraph session verified - token extracted ($($verifyToken.Length) chars)"
        }
        else {
            # SDK v2 keeps tokens opaque by design. Invoke-MSGraphAPI has a routed
            # fallback that uses Invoke-MgGraphRequest (the SDK auths internally),
            # so this is normal — debug-level only.
            Write-LogDebug "MgGraph: session valid (Get-MgContext: tenant=$($ctx.TenantId)) but raw bearer not extractable. Graph calls route via Invoke-MgGraphRequest."
        }

        # Realign the active auth provider so subsequent Invoke-MSGraphAPI calls route
        # here. Symmetric to the same logic in MSAL's Add-MSALTokenInfo.
        if(Get-Command Set-ActiveAuthProvider -ErrorAction SilentlyContinue) {
            $cur = Get-AuthProvider
            if($cur -and $cur.Id -ne $this.Id) {
                Write-Log "Active auth provider auto-switched from '$($cur.Id)' to '$($this.Id)' because MgGraph authentication succeeded"
                Set-ActiveAuthProvider -Id $this.Id
            }
        }

        # Phase 3: persist per-tenant cloud memory. Prefer the Cloud arg the caller
        # asked for; fall back to mapping Get-MgContext.Environment back to a Cloud
        # value (the SDK's -Environment values match ours 1:1 via Clouds[].MgEnvironment).
        # Declared before the try so it's always in scope for Register-AuthToken below.
        $persistCloud = if($cloudValue) { $cloudValue } else { $null }
        try {
            if(-not $persistCloud -and $ctx -and $ctx.Environment) {
                $match = $script:Clouds | Where-Object MgEnvironment -eq $ctx.Environment | Select-Object -First 1
                if($match) { $persistCloud = $match.Value }
            }
            if(-not $persistCloud) { $persistCloud = Get-DefaultCloud }
            if($persistCloud -and $ctx.TenantId) {
                Save-TenantCloud -TenantId $ctx.TenantId -Cloud $persistCloud
                Save-SettingStoreValue "" "LastLoggedOnCloud" $persistCloud
            }
        }
        catch {
            Write-LogDebug "Phase 3 MgGraph cloud memory write failed: $($_.Exception.Message)"
        }

        # Register with the central token registry (single-session invariant: drop
        # any prior entry first so repeated Connect never accumulates entries).
        # The registry fires AuthenticatedNewToken with the canonical [IMAuthToken]
        # - MgGraph now participates in the auth events for the first time.
        if($this.CurrentTokenId -gt 0) {
            Unregister-AuthToken -TokenId $this.CurrentTokenId
        }
        $this.CurrentTokenId = Get-NextAuthTokenId
        return (Register-AuthToken -Provider $this -TokenId $this.CurrentTokenId -Cloud $persistCloud)
    }

    [bool] Disconnect([int]$TokenId) {
        try {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            [AuthenticationMgGraph]::ClearTokenCache()
            if($this.CurrentTokenId -gt 0) {
                Unregister-AuthToken -TokenId $this.CurrentTokenId
                $this.CurrentTokenId = 0
            }
            return $true
        }
        catch {
            Write-LogError "Disconnect-MgGraph failed" $_.Exception
            return $false
        }
    }

    # Re-trigger Connect-MgGraph against the current session's tenant. With a valid
    # refresh token in the cache the SDK does this silently; otherwise it prompts.
    [bool] Refresh([int]$TokenId) {
        try {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            $mgArgs = @{ NoWelcome = $true }
            if($ctx -and $ctx.TenantId) { $mgArgs['TenantId'] = $ctx.TenantId }
            if($ctx -and $ctx.ClientId) { $mgArgs['ClientId'] = $ctx.ClientId }
            if($ctx -and $ctx.Scopes)   { $mgArgs['Scopes']   = @($ctx.Scopes) }
            if($ctx -and $ctx.Environment -and $ctx.Environment -ne "Global") { $mgArgs['Environment'] = $ctx.Environment }
            Write-Log "MgGraph Refresh: re-running Connect-MgGraph (tenant: $($ctx.TenantId))"
            # Invalidate the cached bearer so the next GetAccessToken call re-extracts.
            [AuthenticationMgGraph]::ClearTokenCache()
            Connect-MgGraph @mgArgs -ErrorAction Stop | Out-Null
            return $true
        }
        catch {
            Write-LogError "MgGraph Refresh failed" $_.Exception
            return $false
        }
    }

    static [void] ClearTokenCache() {
        [AuthenticationMgGraph]::CachedToken = $null
        [AuthenticationMgGraph]::CachedTokenExpiry = [DateTimeOffset]::MinValue
        [AuthenticationMgGraph]::CachedTokenTenantId = $null
    }

    # Cached token + expiry to avoid hitting the SDK on every request.
    static [string]$CachedToken
    static [DateTimeOffset]$CachedTokenExpiry = [DateTimeOffset]::MinValue
    static [string]$CachedTokenTenantId

    # The single global token id for this provider's one live session. MgGraph is
    # single-session by SDK design (one ambient Get-MgContext), so it holds exactly
    # one registry entry at a time. 0 = not registered.
    [int]$CurrentTokenId = 0

    # Robust token extraction. Microsoft.Graph SDK v2 doesn't expose an access token
    # accessor — AuthContext.AccessToken stays null in the delegated flow. We use the
    # SDK's own HttpClient (which has its auth DelegatingHandler attached) to make a
    # cheap HEAD-style request; the handler mutates the request to add
    # "Authorization: Bearer <token>" before sending. We then read the header off the
    # request. Same mechanism the SDK itself uses internally on every Mg* cmdlet —
    # no reflection, no private APIs.
    [string] GetAccessToken([int]$TokenId, [string]$Resource) {
        try {
            # Return cached token if still valid (>5 min until expiry). The SDK refreshes
            # internally on every call, so caching at our layer avoids per-request
            # network round-trips just to "pull" a token.
            $ctxTenantId = $null
            try { $ctxTenantId = (Get-MgContext -ErrorAction SilentlyContinue).TenantId } catch { }

            if([AuthenticationMgGraph]::CachedToken -and
               [AuthenticationMgGraph]::CachedTokenExpiry -gt [DateTimeOffset]::UtcNow.AddMinutes(5) -and
               [AuthenticationMgGraph]::CachedTokenTenantId -eq $ctxTenantId) {
                return [AuthenticationMgGraph]::CachedToken
            }

            $sessionType = "Microsoft.Graph.PowerShell.Authentication.GraphSession" -as [type]
            if(-not $sessionType) {
                Write-LogDebug "MgGraph: GraphSession type not available; SDK not loaded?"
                return $null
            }
            $session = $sessionType::Instance
            if(-not $session) {
                Write-LogDebug "MgGraph: GraphSession.Instance is null"
                return $null
            }

            # The SDK exposes GraphHttpClient: an HttpClient wired with auth handlers.
            $httpClient = $session.GraphHttpClient
            if(-not $httpClient) {
                Write-LogDebug "MgGraph: GraphHttpClient is null (Connect-MgGraph not run?)"
                return $null
            }

            # Strategy A — direct AuthContext read (works on some SDK builds, fast-path).
            if($session.AuthContext -and $session.AuthContext.AccessToken) {
                $tok = [AuthenticationMgGraph]::UnprotectString($session.AuthContext.AccessToken)
                if($tok) {
                    [AuthenticationMgGraph]::SaveTokenCache($tok, $ctxTenantId)
                    Write-LogDebug "MgGraph: token from AuthContext.AccessToken (fast-path)"
                    return $tok
                }
            }

            # Strategy B — HttpClient sniff: make a trivial GET via the SDK's HttpClient,
            # then read the Authorization header that the DelegatingHandler attached.
            # This is the supported public contract of the SDK's auth pipeline.
            $graphResource = if($Resource) { $Resource.TrimEnd('/') } else { "https://$(Get-GraphDomain)" }
            $req = [System.Net.Http.HttpRequestMessage]::new(
                [System.Net.Http.HttpMethod]::Get,
                "$graphResource/v1.0/`$metadata")
            try {
                $task = $httpClient.SendAsync(
                    $req,
                    [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead)
                # 30s timeout — interactive auth could be required if cache is cold.
                if(-not $task.Wait(30000)) {
                    Write-LogDebug "MgGraph: HttpClient sniff timed out"
                    return $null
                }
                # Drop the response (we only care about the request headers the handler
                # populated). Dispose to free the socket.
                try { $task.Result.Dispose() } catch { }
            }
            catch {
                Write-LogDebug "MgGraph: HttpClient sniff failed: $($_.Exception.Message)"
            }

            if($req.Headers.Authorization -and
               $req.Headers.Authorization.Scheme -eq 'Bearer' -and
               $req.Headers.Authorization.Parameter) {
                $token = $req.Headers.Authorization.Parameter
                [AuthenticationMgGraph]::SaveTokenCache($token, $ctxTenantId)
                Write-LogDebug "MgGraph: token via HttpClient auth-handler sniff"
                return $token
            }

            # Last resort — reflection probe.
            $token = [AuthenticationMgGraph]::FindTokenViaReflection($session)
            if($token) {
                [AuthenticationMgGraph]::SaveTokenCache($token, $ctxTenantId)
                Write-LogDebug "MgGraph: token via reflection"
                return $token
            }

            $ctxJson = "<null>"
            if($session.AuthContext) {
                try {
                    $ctxJson = ($session.AuthContext | Select-Object TenantId, ClientId, AppName, AuthType, @{n='AccessTokenPresent';e={[bool]$_.AccessToken}}, @{n='Scopes';e={($_.Scopes -join ', ')}} | ConvertTo-Json -Compress)
                }
                catch { $ctxJson = "<unserializable>" }
            }
            Write-Log "MgGraph: could not extract access token. AuthContext=$ctxJson" 2
            return $null
        }
        catch {
            Write-LogError "Failed to extract MgGraph access token" $_.Exception
            return $null
        }
    }

    # Persist the freshly-acquired token + expiry. Expiry parsed from the JWT exp claim;
    # fall back to "now + 50 minutes" if parsing fails (Entra tokens default to 60 min).
    static [void] SaveTokenCache([string]$Token, [string]$TenantId) {
        [AuthenticationMgGraph]::CachedToken = $Token
        [AuthenticationMgGraph]::CachedTokenTenantId = $TenantId
        $expiry = [DateTimeOffset]::UtcNow.AddMinutes(50)
        try {
            # Decode JWT exp claim
            $jwt = Get-JWTtoken $Token
            if($jwt -and $jwt.Payload -and $jwt.Payload.exp) {
                $expiry = [DateTimeOffset]::FromUnixTimeSeconds([long]$jwt.Payload.exp)
            }
        }
        catch { }
        [AuthenticationMgGraph]::CachedTokenExpiry = $expiry
    }

    # Unprotect a SecureString or pass a plain string through.
    static [string] UnprotectString($Value) {
        if($null -eq $Value) { return $null }
        if($Value -is [SecureString]) {
            return [System.Net.NetworkCredential]::new("", $Value).Password
        }
        return [string]$Value
    }

    # Walks the session object graph looking for a property/field whose name suggests it
    # holds an access token. Limited to 2 levels deep to avoid infinite recursion.
    static [string] FindTokenViaReflection($obj) {
        if($null -eq $obj) { return $null }
        try {
            foreach($prop in $obj.PSObject.Properties) {
                if($prop.Name -match 'AccessToken|BearerToken|JWT|^Token$') {
                    $val = $prop.Value
                    if($val) {
                        $unwrapped = [AuthenticationMgGraph]::UnprotectString($val)
                        if($unwrapped -and $unwrapped.Length -gt 20) { return $unwrapped }
                    }
                }
            }
            foreach($prop in $obj.PSObject.Properties) {
                if($prop.Name -in 'AuthContext','InMemoryTokenCache','GraphOption','RequestContext') {
                    $child = $prop.Value
                    if($child) {
                        foreach($childProp in $child.PSObject.Properties) {
                            if($childProp.Name -match 'AccessToken|BearerToken|JWT|^Token$') {
                                $val = $childProp.Value
                                if($val) {
                                    $unwrapped = [AuthenticationMgGraph]::UnprotectString($val)
                                    if($unwrapped -and $unwrapped.Length -gt 20) { return $unwrapped }
                                }
                            }
                        }
                    }
                }
            }
        }
        catch { }
        return $null
    }

    [datetime] GetAccessTokenExpiry([int]$TokenId, [string]$Resource) {
        # SDK does not expose the expiry to consumers. We return MaxValue and rely on
        # MgGraph's internal silent refresh (it owns the cache).
        return [datetime]::MaxValue
    }

    # Cached tenant display name to avoid hitting /organization on every refresh.
    static [hashtable]$TenantNameCache = @{}

    [PSCustomObject] GetUserInfo([int]$TokenId) {
        try {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if(-not $ctx) { return $null }

            # SDK reports AuthType as Delegated / AppOnly. Map to our taxonomy.
            $authType = switch ($ctx.AuthType) {
                "AppOnly"   { "ClientCredential" }
                "Delegated" { "Interactive" }
                default     { "$($ctx.AuthType)" }
            }

            # Get-MgContext doesn't surface tenant display name. Fetch it once per
            # tenant via Invoke-MgGraphRequest /organization (the SDK handles auth);
            # cache for the rest of the session.
            $tenantName = $null
            if($ctx.TenantId) {
                if([AuthenticationMgGraph]::TenantNameCache.ContainsKey($ctx.TenantId)) {
                    $tenantName = [AuthenticationMgGraph]::TenantNameCache[$ctx.TenantId]
                }
                else {
                    try {
                        $org = Invoke-MgGraphRequest -Method GET -Uri "https://$(Get-GraphDomain)/v1.0/organization" -OutputType PSObject -ErrorAction Stop
                        if($org -and $org.value -and $org.value.Count -gt 0 -and $org.value[0].displayName) {
                            $tenantName = $org.value[0].displayName
                            [AuthenticationMgGraph]::TenantNameCache[$ctx.TenantId] = $tenantName
                        }
                    }
                    catch {
                        Write-LogDebug "MgGraph GetUserInfo: /organization fetch failed ($($_.Exception.Message)); tenant display name will be unknown"
                    }
                }
            }

            # Expiry comes from the JWT exp claim we parsed into CachedTokenExpiry.
            # If no token has been minted yet this session, trigger one — cheap when
            # the SDK's in-memory cache is warm.
            $expiresOn = $null
            if([AuthenticationMgGraph]::CachedTokenExpiry -gt [DateTimeOffset]::MinValue) {
                $expiresOn = [AuthenticationMgGraph]::CachedTokenExpiry.LocalDateTime
            }
            else {
                try {
                    [void]$this.GetAccessToken(0, "https://$(Get-GraphDomain)")
                    if([AuthenticationMgGraph]::CachedTokenExpiry -gt [DateTimeOffset]::MinValue) {
                        $expiresOn = [AuthenticationMgGraph]::CachedTokenExpiry.LocalDateTime
                    }
                }
                catch { }
            }

            return [PSCustomObject]@{
                Provider    = $this.Id
                DisplayName = $ctx.Account
                UPN         = $ctx.Account
                UserId      = $null   # Not surfaced by Get-MgContext
                TenantId    = $ctx.TenantId
                TenantName  = $tenantName
                AppId       = $ctx.ClientId
                AppName     = $ctx.AppName
                AuthType    = $authType
                ExpiresOn   = $expiresOn
            }
        }
        catch {
            return $null
        }
    }

    [PSCustomObject[]] GetCachedAccounts() {
        # Delegate to a module function — PS class method bodies are parsed strictly
        # (type references like [Microsoft.Identity.Client.PublicClientApplicationBuilder]
        # have to resolve at PARSE time, before MSAL DLLs are loaded). Module
        # functions are late-bound and tolerate this.
        $rows = @()
        try {
            $rows = @(Get-MgGraphCachedMsalAccounts -ProviderId $this.Id)
        }
        catch {
            Write-LogDebug "MgGraph GetCachedAccounts failed: $($_.Exception.Message)"
        }
        return [PSCustomObject[]]$rows
    }

    [PSCustomObject[]] GetAvailableTenants([int]$TokenId) {
        # SDK does not surface available tenants. Tenant switching means
        # Disconnect + Connect with -TenantId, which the UI can offer via a manual
        # tenant id entry (Phase 3 UI concern).
        return [PSCustomObject[]]@()
    }
}
