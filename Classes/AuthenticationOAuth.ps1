#ImportOrder 27

# Pure-PowerShell implementation of AuthenticationProvider.
#
# No MSAL.NET DLL, no Microsoft.Graph.Authentication SDK. The class talks to
# https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token directly via
# Invoke-RestMethod. Designed for automation: scheduled tasks, CI runners,
# AKS workload identity, Azure VMs, App Service / Functions.
#
# Supported flows (dispatched by Connect() based on which arg is set, in
# this priority order):
#
#   Token                                BYO bearer (no flow)
#   DeviceCode                           device code grant (RFC 8628) — v2
#   ManagedIdentity (no Federated*)      IMDS / App Service / Functions
#   FederatedTokenFile / FederatedToken  Workload Identity Federation
#   Certificate / CertificatePath        client_credentials w/ private_key_jwt
#   Secret                               client_credentials w/ shared secret
#   Credential (PSCredential)            ROPC (grant_type=password)
#
# All endpoint heavy lifting (HTTP, JWT signing, IMDS detection, error
# surfacing) lives in Internal/AuthenticationOAuthHelpers.ps1 — module
# functions are late-bound and tolerate types that don't resolve at parse
# time, which keeps this class file portable.
#
# Per-tenant cloud memory is stamped via Save-TenantCloud at the end of every
# successful Connect, matching what AuthenticationMSAL and AuthenticationMgGraph
# do. Same for AppEvents (AuthenticatedNewToken / AuthenticationTokenRefresh /
# AuthenticationUserDisconnected / AuthenticationFailed).

class AuthenticationOAuth : AuthenticationProvider {

    # Token cache. Static so every reference to the provider sees the same
    # store within a session. Keys are integer TokenIds (consistent with the
    # other providers).
    static [hashtable]$Tokens = @{}
    static [int]$NextTokenId  = 1

    # Cached tenant display name per TenantId so GetUserInfo doesn't hit
    # /organization on every refresh event. Failed lookups (e.g. app-only token
    # without Organization.Read.All) cache $null so they aren't retried either.
    static [hashtable]$TenantNameCache = @{}

    AuthenticationOAuth() {
        $this.Id          = "OAuth"
        $this.DisplayName = "Direct OAuth (no SDK)"

        # Required contract.
        $this.SupportsInteractive       = $true    # device code flow (RFC 8628)
        $this.SupportsClientSecret      = $true
        $this.SupportsCertificate       = $true

        # Recommended.
        $this.SupportsIdentityProvider  = $true    # IMDS + workload federation
        $this.SupportsBYOToken          = $true
        $this.SupportsClaimsChallenge   = $true    # re-mints via the /token body (see GetClaimsToken)

        $this.SupportsMultiTenant       = $true
        $this.SupportsRefresh           = $true
        $this.SupportsForget            = $true
        $this.SupportsCachedUsers       = $false   # No persistent on-disk cache
    }

    [void] Initialize() {
        # Register the OAuth settings section (browser-login config). Guard on the
        # settings API being loaded (module-load order); Initialize() runs from
        # Register-AuthProvider, before the Settings dialog renders.
        if(-not (Get-Command -Name Add-SettingsSection -ErrorAction SilentlyContinue)) { return }
        if(-not (Get-Command -Name Add-SettingsObject  -ErrorAction SilentlyContinue)) { return }
        try {
            # Order 9 = directly under the MSAL section (8). App (client) id and tenant id
            # are NOT registered here - MSAL and OAuth are two ways of authenticating with
            # the SAME app, so both resolve it from the common Entra settings in the
            # Authentication section (Get-EntraApp: dropdown -> custom app id -> default
            # Microsoft Graph PowerShell public client).
            Add-SettingsSection -Title "OAuth" -Id "OAuth" -Order 9

            Add-SettingsObject -Title "OAuth browser prompt" -Key "OAuthPrompt" -Type "List" -DefaultValue "select_account" `
                -ItemsSource @(
                    [PSCustomObject]@{ Name = "Select account"; Value = "select_account" },
                    [PSCustomObject]@{ Name = "Force login";    Value = "login" },
                    [PSCustomObject]@{ Name = "Consent";        Value = "consent" },
                    [PSCustomObject]@{ Name = "None (silent)";  Value = "none" }
                ) `
                -Description "OAuth /authorize prompt behaviour. 'Force login' re-authenticates even with an active browser session (equivalent to force-interactive); 'None' fails if interaction would be required." `
                -Section "OAuth"

            Add-SettingsObject -Title "OAuth login hint (UPN)" -Key "OAuthLoginHint" -Type "String" -DefaultValue "" `
                -Description "Optional UPN to pre-fill on the sign-in page (login_hint)." `
                -Section "OAuth"

            Add-SettingsObject -Title "OAuth redirect port" -Key "OAuthRedirectPort" -Type "Int" -DefaultValue 0 `
                -Description "Fixed loopback port for the browser redirect (http://localhost:<port>). 0 = pick a free port automatically. Set a fixed port only if your app registration requires a specific http://localhost:<port> redirect." `
                -Section "OAuth"

            Add-SettingsObject -Title "Remember login (cache token)" -Key "OAuthCacheToken" -Type "Boolean" -DefaultValue $false `
                -Description "Persist the OAuth refresh token (DPAPI-encrypted, current user) so the app silently resumes the browser session after a restart. When off, you sign in again after each restart (usually a quick browser redirect via existing SSO)." `
                -Section "OAuth"
        }
        catch {
            Write-LogError "Failed to register OAuth settings section" $_.Exception
        }
    }

    # Silent cross-restart resume for the browser flow. When "Remember login" is on and
    # a DPAPI-cached refresh token exists, mint a fresh token via the refresh_token grant
    # (no browser) and register it as the default. Returns $true on success. Base is a
    # no-op; MSAL/other providers have their own resume paths.
    [bool] TryResumeSession() {
        try {
            if((Get-SettingValue "OAuthCacheToken") -ne $true) { return $false }
            $cache = Read-OAuthTokenCache
            if(-not $cache -or -not $cache.RefreshToken) { return $false }

            $cloudValue = if($cache.Cloud) { [string]$cache.Cloud } else { Get-DefaultCloud }
            $authority  = if($cache.Authority) { [string]$cache.Authority } else { (Get-CloudByValue $cloudValue).AADAuthority }
            $resource   = if($cache.Resource) { [string]$cache.Resource } else { "https://$((Get-CloudByValue $cloudValue).GraphHost)" }
            $tenant     = if($cache.TenantId) { [string]$cache.TenantId } else { 'organizations' }
            $scope      = "$resource/.default offline_access"

            $tokenResp = Invoke-OAuthTokenRequest -Authority $authority -TenantId $tenant -Body @{
                grant_type    = 'refresh_token'
                client_id     = $cache.ClientId
                refresh_token = $cache.RefreshToken
                scope         = $scope
            }
            if(-not $tokenResp -or -not $tokenResp.access_token) { return $false }

            $expiresAt = [DateTime]::UtcNow.AddMinutes(50)
            if($tokenResp.expires_in) { $expiresAt = [DateTime]::UtcNow.AddSeconds([int]$tokenResp.expires_in) }

            # Recover the real tenant from the token when we resumed under 'organizations'.
            $tenantId = [string]$cache.TenantId
            try {
                $jwt = Get-JWTtoken $tokenResp.access_token
                if($jwt -and $jwt.Payload -and $jwt.Payload.tid) { $tenantId = [string]$jwt.Payload.tid }
            } catch { }

            $cred = [ordered]@{
                AuthMethod = 'AuthCode'
                Cloud      = $cloudValue
                Authority  = $authority
                Resource   = $resource
                TenantId   = $tenantId
                ClientId   = $cache.ClientId
            }
            $tokenId = Get-NextAuthTokenId
            $entry = [ordered]@{
                Id              = $tokenId
                AccessToken     = [string]$tokenResp.access_token
                RefreshToken    = if($tokenResp.refresh_token) { [string]$tokenResp.refresh_token } else { [string]$cache.RefreshToken }
                ExpiresAt       = $expiresAt
                TenantId        = $tenantId
                ClientId        = $cache.ClientId
                AuthMethod      = 'AuthCode'
                Cloud           = $cloudValue
                Resource        = $resource
                CredentialState = $cred
                AcquiredAt      = [DateTime]::UtcNow
            }
            [AuthenticationOAuth]::Tokens[$tokenId] = $entry

            # A rotated refresh token must overwrite the cached one.
            if($tokenResp.refresh_token) {
                [void](Save-OAuthTokenCache -Data @{
                    RefreshToken = [string]$tokenResp.refresh_token
                    ClientId     = $cache.ClientId
                    TenantId     = $tenantId
                    Cloud        = $cloudValue
                    Authority    = $authority
                    Resource     = $resource
                    AuthMethod   = 'AuthCode'
                })
            }

            [void](Register-AuthToken -Provider $this -TokenId $tokenId -Cloud $cloudValue -Default)
            Write-Log "OAuth provider: resumed cached browser session (TokenId=$tokenId, tenant=$tenantId)"
            return $true
        }
        catch {
            Write-LogError "OAuth provider: TryResumeSession failed" $_.Exception
            return $false
        }
    }

    Hidden [string] ResolvePublicClientId([string]$AppId) {
        if(-not [String]::IsNullOrWhiteSpace($AppId)) { return $AppId }

        # Match MSAL's app selection: Settings -> Entra app dropdown, then custom
        # app id, then the default Microsoft Graph PowerShell public client.
        if(Get-Command Get-EntraApp -ErrorAction SilentlyContinue) {
            try {
                $entraApp = Get-EntraApp
                if($entraApp -and -not [String]::IsNullOrWhiteSpace([string]$entraApp.ClientId)) {
                    return [string]$entraApp.ClientId
                }
            } catch { }
        }

        return "14d82eec-204b-4c2f-b7e8-296a70dab67e"
    }

    Hidden [string] ResolveTenantId([string]$TenantId) {
        if(-not [String]::IsNullOrWhiteSpace($TenantId)) { return $TenantId }

        # Shared identity config: the common Entra settings supply the tenant the same
        # way they supply the app id. Get-EntraApp surfaces EntraCustomTenantId on
        # custom-app rows; the built-in app rows have no TenantId property, which
        # safely yields $null here.
        if(Get-Command Get-EntraApp -ErrorAction SilentlyContinue) {
            try {
                $entraApp = Get-EntraApp
                if($entraApp -and -not [String]::IsNullOrWhiteSpace([string]$entraApp.TenantId)) {
                    return [string]$entraApp.TenantId
                }
            } catch { }
        }

        # 'organizations' = any work/school account; the real tenant id is recovered
        # from the token's tid claim after sign-in.
        return 'organizations'
    }

    # === Auth lifecycle ===

    [PSCustomObject] Connect([hashtable]$Arguments) {
        Write-Log "AuthenticationOAuth.Connect starting"

        $cloudValue = if($Arguments.Cloud) { [string]$Arguments.Cloud } else { Get-DefaultCloud }
        $cloudEntry = Get-CloudByValue $cloudValue
        $authority  = $cloudEntry.AADAuthority
        $graphHost  = $cloudEntry.GraphHost
        $defaultScope = "https://$graphHost/.default"
        $resource     = "https://$graphHost"

        $tenantId = [string]$Arguments.TenantId
        $clientId = [string]$Arguments.AppId

        # Determine flow + run it. State stored in $cred is what Refresh() and
        # GetAccessToken() use to re-acquire when the access token expires.
        $cred = $null
        $tokenResp = $null
        $authMethod = $null

        try {
            if($Arguments.ContainsKey('Token') -and $Arguments.Token) {
                $authMethod = 'BYO'
                $cred = [ordered]@{
                    AuthMethod  = $authMethod
                    Cloud       = $cloudValue
                    Authority   = $authority
                    Resource    = $resource
                    TenantId    = $tenantId
                    ClientId    = $clientId
                }
                $tokenResp = [PSCustomObject]@{
                    access_token = [string]$Arguments.Token
                    expires_in   = $null    # unknown; parsed from JWT exp below
                    token_type   = 'Bearer'
                }
            }
            elseif( ($Arguments.ContainsKey('Browser') -and $Arguments.Browser) -or
                    ($Arguments.ContainsKey('Interactive') -and $Arguments.Interactive -and
                     -not ($Arguments.ContainsKey('DeviceCode') -and $Arguments.DeviceCode) -and
                     ((Get-CacheObject "ShowUI") -eq $true)) ) {
                # Browser (Authorization Code + PKCE, loopback redirect). Chosen for an
                # explicit -Browser, or for -Interactive when a GUI is present (ShowUI);
                # headless -Interactive keeps going to device code below.
                $authMethod = 'AuthCode'
                # Client id / tenant: explicit args win, else the common Entra settings
                # (same app selection as MSAL - dropdown, custom app id, then the
                # well-known Microsoft Graph PowerShell public client, which already
                # has http://localhost registered).
                $clientId = $this.ResolvePublicClientId($clientId)
                $acTenant = $this.ResolveTenantId($tenantId)
                $prompt    = if($Arguments.Prompt) { [string]$Arguments.Prompt } else { [string](Get-SettingValue "OAuthPrompt") }
                $loginHint = if($Arguments.LoginHint) { [string]$Arguments.LoginHint } else { [string](Get-SettingValue "OAuthLoginHint") }
                $redirectPort = if($Arguments.RedirectPort) { [int]$Arguments.RedirectPort } else { [int](Get-SettingValue "OAuthRedirectPort") }
                $timeoutSec = 600
                try { $t = [int](Get-SettingValue "MSGraphInteractiveTimeoutSec"); if($t -gt 0) { $timeoutSec = $t } } catch { }

                $cred = [ordered]@{
                    AuthMethod  = $authMethod
                    Cloud       = $cloudValue
                    Authority   = $authority
                    Resource    = $resource
                    TenantId    = $tenantId
                    ClientId    = $clientId
                }
                try {
                    $tokenResp = Invoke-OAuthAuthCodeFlow -Authority $authority -TenantId $acTenant -ClientId $clientId `
                        -Scope "$defaultScope offline_access" -RedirectPort $redirectPort -Prompt $prompt -LoginHint $loginHint -TimeoutSec $timeoutSec
                }
                catch [System.OperationCanceledException] {
                    # The user clicked Cancel on the sign-in overlay. That is a decision,
                    # not a broken browser, so do not start a second (device code) login
                    # behind their back - let it out and leave the session signed out.
                    Write-Log "OAuth provider: browser sign-in cancelled by the user" 2
                    throw
                }
                catch {
                    # Browser unavailable (headless -Browser, no default browser, or the
                    # loopback port is blocked). Fall back to device code so sign-in can
                    # still complete. Refresh works identically for both (refresh_token).
                    Write-Log "OAuth provider: browser sign-in failed ($($_.Exception.Message)); falling back to device code" 2
                    $tokenResp = Invoke-OAuthDeviceCodeFlow -Authority $authority -TenantId $acTenant -ClientId $clientId -Scope "$defaultScope offline_access"
                }
            }
            elseif(($Arguments.ContainsKey('DeviceCode') -and $Arguments.DeviceCode) -or
                   ($Arguments.ContainsKey('Interactive') -and $Arguments.Interactive)) {
                if($Arguments.Interactive -and -not $Arguments.DeviceCode) {
                    # -Interactive reached here means no GUI was available for the
                    # browser flow above (headless), so use RFC 8628 device code -
                    # keeping the Connect-IntuneManagement -Interactive story working
                    # on every provider.
                    Write-Log "OAuth provider: -Interactive routed to device code flow (no GUI for browser login)"
                }
                $authMethod = 'DeviceCode'
                $clientId = $this.ResolvePublicClientId($clientId)
                if(-not $clientId) { throw "AppId is required for device code auth" }
                # Tenant from the common Entra settings when not explicit (same
                # resolution as the browser flow above).
                $dcTenant = $this.ResolveTenantId($tenantId)
                $cred = [ordered]@{
                    AuthMethod  = $authMethod
                    Cloud       = $cloudValue
                    Authority   = $authority
                    Resource    = $resource
                    TenantId    = $tenantId
                    ClientId    = $clientId
                }
                # offline_access so the response includes a refresh_token —
                # that's what AcquireFromState uses for silent renewal.
                $tokenResp = Invoke-OAuthDeviceCodeFlow -Authority $authority -TenantId $dcTenant -ClientId $clientId -Scope "$defaultScope offline_access"
            }
            elseif($Arguments.ContainsKey('ManagedIdentity') -and $Arguments.ManagedIdentity -and
                  -not ($Arguments.FederatedTokenFile -or $Arguments.FederatedToken)) {
                $authMethod = 'IMDS'
                if(-not $clientId -and $Arguments.ManagedIdentityClientId) {
                    $clientId = [string]$Arguments.ManagedIdentityClientId
                }
                $cred = [ordered]@{
                    AuthMethod  = $authMethod
                    Cloud       = $cloudValue
                    Authority   = $authority
                    Resource    = $resource
                    TenantId    = $tenantId
                    ClientId    = $clientId
                }
                $tokenResp = Invoke-OAuthIMDS -Resource $resource -ClientId $clientId
                # IMDS responds with token_type / access_token / expires_in (string seconds);
                # tenant_id is also included. Use the IMDS-reported tenant if our caller
                # didn't supply one.
                if(-not $tenantId -and $tokenResp.tenant_id) {
                    $tenantId = $tokenResp.tenant_id
                    $cred.TenantId = $tenantId
                }
            }
            elseif($Arguments.FederatedTokenFile -or $Arguments.FederatedToken) {
                $authMethod = 'Federated'
                if(-not $tenantId) { throw "TenantId is required for federated credential auth" }
                if(-not $clientId) { throw "AppId is required for federated credential auth" }
                $cred = [ordered]@{
                    AuthMethod         = $authMethod
                    Cloud              = $cloudValue
                    Authority          = $authority
                    Resource           = $resource
                    TenantId           = $tenantId
                    ClientId           = $clientId
                    FederatedTokenFile = [string]$Arguments.FederatedTokenFile
                    FederatedToken     = [string]$Arguments.FederatedToken
                }
                $assertion = Get-OAuthFederatedAssertion -FederatedToken $cred.FederatedToken -FederatedTokenFile $cred.FederatedTokenFile
                $tokenResp = Invoke-OAuthTokenRequest -Authority $authority -TenantId $tenantId -Body @{
                    grant_type            = 'client_credentials'
                    client_id             = $clientId
                    scope                 = $defaultScope
                    client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
                    client_assertion      = $assertion
                }
            }
            elseif($Arguments.ContainsKey('Certificate') -and $Arguments.Certificate) {
                $authMethod = 'Certificate'
                if(-not $tenantId) { throw "TenantId is required for certificate auth" }
                if(-not $clientId) { throw "AppId is required for certificate auth" }
                $cert = $this.ResolveCertificate($Arguments.Certificate, $null, $null)
                $cred = [ordered]@{
                    AuthMethod  = $authMethod
                    Cloud       = $cloudValue
                    Authority   = $authority
                    Resource    = $resource
                    TenantId    = $tenantId
                    ClientId    = $clientId
                    Certificate = $cert
                }
                $assertion = Get-OAuthClientAssertion -Certificate $cert -ClientId $clientId -Authority $authority -TenantId $tenantId
                $tokenResp = Invoke-OAuthTokenRequest -Authority $authority -TenantId $tenantId -Body @{
                    grant_type            = 'client_credentials'
                    client_id             = $clientId
                    scope                 = $defaultScope
                    client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
                    client_assertion      = $assertion
                }
            }
            elseif($Arguments.ContainsKey('CertificatePath') -and $Arguments.CertificatePath) {
                $authMethod = 'Certificate'
                if(-not $tenantId) { throw "TenantId is required for certificate auth" }
                if(-not $clientId) { throw "AppId is required for certificate auth" }
                $cert = $this.ResolveCertificate($null, [string]$Arguments.CertificatePath, $Arguments.CertificatePassword)
                $cred = [ordered]@{
                    AuthMethod          = $authMethod
                    Cloud               = $cloudValue
                    Authority           = $authority
                    Resource            = $resource
                    TenantId            = $tenantId
                    ClientId            = $clientId
                    Certificate         = $cert
                    CertificatePath     = [string]$Arguments.CertificatePath
                    CertificatePassword = $Arguments.CertificatePassword
                }
                $assertion = Get-OAuthClientAssertion -Certificate $cert -ClientId $clientId -Authority $authority -TenantId $tenantId
                $tokenResp = Invoke-OAuthTokenRequest -Authority $authority -TenantId $tenantId -Body @{
                    grant_type            = 'client_credentials'
                    client_id             = $clientId
                    scope                 = $defaultScope
                    client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
                    client_assertion      = $assertion
                }
            }
            elseif($Arguments.ContainsKey('Secret') -and $Arguments.Secret) {
                $authMethod = 'ClientSecret'
                if(-not $tenantId) { throw "TenantId is required for client-secret auth" }
                if(-not $clientId) { throw "AppId is required for client-secret auth" }
                $secretPlain = if($Arguments.Secret -is [SecureString]) {
                    [System.Net.NetworkCredential]::new("", $Arguments.Secret).Password
                } else {
                    [string]$Arguments.Secret
                }
                $cred = [ordered]@{
                    AuthMethod = $authMethod
                    Cloud      = $cloudValue
                    Authority  = $authority
                    Resource   = $resource
                    TenantId   = $tenantId
                    ClientId   = $clientId
                    Secret     = if($Arguments.Secret -is [SecureString]) { $Arguments.Secret } else { ConvertTo-SecureString $secretPlain -AsPlainText -Force }
                }
                $tokenResp = Invoke-OAuthTokenRequest -Authority $authority -TenantId $tenantId -Body @{
                    grant_type    = 'client_credentials'
                    client_id     = $clientId
                    client_secret = $secretPlain
                    scope         = $defaultScope
                }
            }
            elseif($Arguments.ContainsKey('Credential') -and $Arguments.Credential) {
                $authMethod = 'Password'
                if(-not $tenantId) { throw "TenantId is required for password auth" }
                if(-not $clientId) { throw "AppId is required for password auth" }
                $pscred = [PSCredential]$Arguments.Credential
                $passwordPlain = $pscred.GetNetworkCredential().Password
                $cred = [ordered]@{
                    AuthMethod = $authMethod
                    Cloud      = $cloudValue
                    Authority  = $authority
                    Resource   = $resource
                    TenantId   = $tenantId
                    ClientId   = $clientId
                    Username   = $pscred.UserName
                    Password   = $pscred.Password
                }
                $tokenResp = Invoke-OAuthTokenRequest -Authority $authority -TenantId $tenantId -Body @{
                    grant_type = 'password'
                    client_id  = $clientId
                    username   = $pscred.UserName
                    password   = $passwordPlain
                    scope      = "$defaultScope offline_access"
                }
            }
            else {
                throw "AuthenticationOAuth.Connect: no supported credential argument was provided. Pass -Secret, -Certificate, -CertificatePath, -ManagedIdentity, -FederatedTokenFile/-FederatedToken, -Credential, -DeviceCode, or -Token."
            }
        }
        catch [System.OperationCanceledException] {
            # An interactive user explicitly abandoned the flow. Do not publish the
            # canonical AuthenticationFailed event: consumers use that event for real
            # authentication faults (and may tear down/redraw an existing session).
            Write-Log "OAuth provider Connect cancelled by the user (method: $authMethod)" 2
            return $null
        }
        catch {
            Write-LogError "OAuth provider Connect failed (method: $authMethod)" $_.Exception
            Invoke-AuthTokenFailed -Provider $this.Id -TenantId $tenantId -Message "OAuth Connect failed (method: $authMethod)" -Exception $_.Exception
            return $null
        }

        if(-not $tokenResp -or -not $tokenResp.access_token) {
            Write-Log "OAuth provider: token request returned no access_token (method: $authMethod)" 3
            Invoke-AuthTokenFailed -Provider $this.Id -TenantId $tenantId -Message "OAuth token request returned no access_token (method: $authMethod)"
            return $null
        }

        # Compute expiry. Prefer expires_in (relative seconds; reliable across all
        # flows). Fall back to JWT exp claim if expires_in is absent (BYO token).
        $expiresAt = [DateTime]::UtcNow.AddMinutes(50)
        if($tokenResp.expires_in) {
            $secs = [int]$tokenResp.expires_in
            $expiresAt = [DateTime]::UtcNow.AddSeconds($secs)
        }
        else {
            try {
                $jwt = Get-JWTtoken $tokenResp.access_token
                if($jwt -and $jwt.Payload -and $jwt.Payload.exp) {
                    $expiresAt = [DateTimeOffset]::FromUnixTimeSeconds([long]$jwt.Payload.exp).UtcDateTime
                }
            } catch { }
        }

        # If TenantId wasn't supplied (BYO token, IMDS without tenant_id),
        # extract it from the access_token's `tid` claim so multi-tenant
        # routing (Get-GraphDomain, Save-TenantCloud) still works.
        if(-not $tenantId) {
            try {
                $jwt = Get-JWTtoken $tokenResp.access_token
                if($jwt -and $jwt.Payload -and $jwt.Payload.tid) {
                    $tenantId = [string]$jwt.Payload.tid
                    $cred.TenantId = $tenantId
                }
            } catch { }
        }

        # Allocate a TokenId from the central registry so ids are globally unique
        # across providers (MSAL/OAuth/MgGraph), not just within this provider.
        $tokenId = Get-NextAuthTokenId
        $entry = [ordered]@{
            Id              = $tokenId
            AccessToken     = [string]$tokenResp.access_token
            RefreshToken    = if($tokenResp.refresh_token) { [string]$tokenResp.refresh_token } else { $null }
            ExpiresAt       = $expiresAt
            TenantId        = $tenantId
            ClientId        = $clientId
            AuthMethod      = $authMethod
            Cloud           = $cloudValue
            Resource        = $resource
            CredentialState = $cred
            AcquiredAt      = [DateTime]::UtcNow
        }
        [AuthenticationOAuth]::Tokens[$tokenId] = $entry

        # Opt-in cross-restart persistence for the browser flow: DPAPI-encrypt the
        # refresh token when "Remember login" (OAuthCacheToken) is on. Only for the
        # interactive browser method - service/BYO flows re-acquire from their own
        # credentials and shouldn't leave a refresh token on disk.
        if($authMethod -eq 'AuthCode' -and $entry.RefreshToken -and ((Get-SettingValue "OAuthCacheToken") -eq $true)) {
            [void](Save-OAuthTokenCache -Data @{
                RefreshToken = $entry.RefreshToken
                ClientId     = $clientId
                TenantId     = $tenantId
                Cloud        = $cloudValue
                Authority    = $authority
                Resource     = $resource
                AuthMethod   = $authMethod
            })
        }

        Write-Log "OAuth provider: acquired token (TokenId=$tokenId, method=$authMethod, tenant=$tenantId, expires=$($expiresAt.ToLocalTime().ToString('s')))"

        # Persist per-tenant cloud memory so subsequent calls land on the same authority.
        try {
            if($tenantId) {
                Save-TenantCloud -TenantId $tenantId -Cloud $cloudValue
                Save-SettingStoreValue "" "LastLoggedOnCloud" $cloudValue
            }
        } catch {
            Write-LogDebug "OAuth provider: Save-TenantCloud failed: $($_.Exception.Message)"
        }

        # Realign the active provider so subsequent Invoke-MSGraphAPI calls route here.
        try {
            $cur = Get-AuthProvider
            if($cur -and $cur.Id -ne $this.Id) {
                Write-Log "Active auth provider auto-switched from '$($cur.Id)' to '$($this.Id)' because OAuth authentication succeeded"
                Set-ActiveAuthProvider -Id $this.Id
            }
        } catch { }

        # Register with the central token registry, which fires AuthenticatedNewToken
        # with the canonical [IMAuthToken] payload (no longer fired directly here).
        $token = Register-AuthToken -Provider $this -TokenId $tokenId -Cloud $cloudValue
        return $token
    }

    [bool] Disconnect([int]$TokenId) {
        if(-not [AuthenticationOAuth]::Tokens.ContainsKey($TokenId)) { return $false }
        $entry = [AuthenticationOAuth]::Tokens[$TokenId]
        Write-Log "OAuth provider: disconnected TokenId=$TokenId (tenant=$($entry.TenantId))"
        # Unregister BEFORE dropping the backend entry so the disconnect snapshot can
        # still hydrate via GetUserInfo. The registry fires AuthenticationUserDisconnected
        # (and promotes a survivor default if needed).
        Unregister-AuthToken -TokenId $TokenId
        [AuthenticationOAuth]::Tokens.Remove($TokenId) | Out-Null
        # Sign-out of a browser session also drops the persisted refresh token so a
        # later launch doesn't silently resume the account the user just signed out of.
        if($entry.AuthMethod -eq 'AuthCode') { Clear-OAuthTokenCache }
        return $true
    }

    [bool] Refresh([int]$TokenId) {
        if(-not [AuthenticationOAuth]::Tokens.ContainsKey($TokenId)) { return $false }
        try {
            $newToken = $this.AcquireFromState($TokenId)
            return [bool]$newToken
        }
        catch {
            Write-LogError "OAuth provider: refresh failed for TokenId=$TokenId" $_.Exception
            return $false
        }
    }

    # Re-run whatever flow originally minted this entry, replacing the access
    # token (and refresh_token, where applicable) in place. Fires
    # AuthenticationTokenRefresh on success. Used both by the Refresh() public
    # method and by GetAccessToken's preflight expiry check.
    Hidden [string] AcquireFromState([int]$TokenId) {
        return $this.AcquireFromState($TokenId, $null)
    }

    # $Claims carries a CAE claims challenge (from a Graph 401) into the /token
    # request so the re-minted token satisfies the tenant's policy change.
    Hidden [string] AcquireFromState([int]$TokenId, [string]$Claims) {
        if(-not [AuthenticationOAuth]::Tokens.ContainsKey($TokenId)) { return $null }
        $entry = [AuthenticationOAuth]::Tokens[$TokenId]
        $cred  = $entry.CredentialState
        $defaultScope = "$($cred.Resource)/.default"

        $tokenResp = $null
        switch($cred.AuthMethod) {
            'BYO' {
                # BYO bearer can't be silently refreshed — caller must Connect again.
                Write-Log "OAuth provider: BYO token (TokenId=$TokenId) cannot be refreshed automatically" 2
                return $null
            }
            'IMDS' {
                if($Claims) {
                    # IMDS / App Service token endpoints don't accept a claims
                    # parameter — a CAE challenge can't be satisfied here.
                    Write-Log "OAuth provider: managed identity tokens cannot satisfy a CAE claims challenge (TokenId=$TokenId)" 2
                    return $null
                }
                $tokenResp = Invoke-OAuthIMDS -Resource $cred.Resource -ClientId $cred.ClientId
            }
            'Federated' {
                $assertion = Get-OAuthFederatedAssertion -FederatedToken $cred.FederatedToken -FederatedTokenFile $cred.FederatedTokenFile
                $tokenResp = Invoke-OAuthTokenRequest -Authority $cred.Authority -TenantId $cred.TenantId -Claims $Claims -Body @{
                    grant_type            = 'client_credentials'
                    client_id             = $cred.ClientId
                    scope                 = $defaultScope
                    client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
                    client_assertion      = $assertion
                }
            }
            'Certificate' {
                $assertion = Get-OAuthClientAssertion -Certificate $cred.Certificate -ClientId $cred.ClientId -Authority $cred.Authority -TenantId $cred.TenantId
                $tokenResp = Invoke-OAuthTokenRequest -Authority $cred.Authority -TenantId $cred.TenantId -Claims $Claims -Body @{
                    grant_type            = 'client_credentials'
                    client_id             = $cred.ClientId
                    scope                 = $defaultScope
                    client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
                    client_assertion      = $assertion
                }
            }
            'ClientSecret' {
                $secretPlain = [System.Net.NetworkCredential]::new("", $cred.Secret).Password
                $tokenResp = Invoke-OAuthTokenRequest -Authority $cred.Authority -TenantId $cred.TenantId -Claims $Claims -Body @{
                    grant_type    = 'client_credentials'
                    client_id     = $cred.ClientId
                    client_secret = $secretPlain
                    scope         = $defaultScope
                }
            }
            'DeviceCode' {
                # Silent-only here: never re-prompt with a new device code from a
                # refresh path. No refresh token → the 401/expiry surfaces and the
                # user re-runs Connect with -DeviceCode explicitly.
                if(-not $entry.RefreshToken) {
                    Write-Log "OAuth provider: device-code token (TokenId=$TokenId) has no refresh token - run Connect-IntuneManagement -DeviceCode again" 2
                    return $null
                }
                $dcTenant = if($cred.TenantId) { $cred.TenantId } else { 'organizations' }
                $tokenResp = Invoke-OAuthTokenRequest -Authority $cred.Authority -TenantId $dcTenant -Claims $Claims -Body @{
                    grant_type    = 'refresh_token'
                    client_id     = $cred.ClientId
                    refresh_token = $entry.RefreshToken
                    scope         = "$defaultScope offline_access"
                }
            }
            'AuthCode' {
                # Browser (auth-code) refresh is silent - the refresh_token grant,
                # identical to DeviceCode. No browser/redirect needed. Missing refresh
                # token means the user must sign in again.
                if(-not $entry.RefreshToken) {
                    Write-Log "OAuth provider: browser token (TokenId=$TokenId) has no refresh token - sign in again" 2
                    return $null
                }
                $acTenant = if($cred.TenantId) { $cred.TenantId } else { 'organizations' }
                $tokenResp = Invoke-OAuthTokenRequest -Authority $cred.Authority -TenantId $acTenant -Claims $Claims -Body @{
                    grant_type    = 'refresh_token'
                    client_id     = $cred.ClientId
                    refresh_token = $entry.RefreshToken
                    scope         = "$defaultScope offline_access"
                }
            }
            'Password' {
                # Prefer refresh_token grant if we got one; falls back to ROPC re-prompt.
                if($entry.RefreshToken) {
                    $tokenResp = Invoke-OAuthTokenRequest -Authority $cred.Authority -TenantId $cred.TenantId -Claims $Claims -Body @{
                        grant_type    = 'refresh_token'
                        client_id     = $cred.ClientId
                        refresh_token = $entry.RefreshToken
                        scope         = "$defaultScope offline_access"
                    }
                }
                else {
                    $passwordPlain = [System.Net.NetworkCredential]::new("", $cred.Password).Password
                    $tokenResp = Invoke-OAuthTokenRequest -Authority $cred.Authority -TenantId $cred.TenantId -Claims $Claims -Body @{
                        grant_type = 'password'
                        client_id  = $cred.ClientId
                        username   = $cred.Username
                        password   = $passwordPlain
                        scope      = "$defaultScope offline_access"
                    }
                }
            }
            default {
                throw "OAuth provider: unknown AuthMethod '$($cred.AuthMethod)' on TokenId=$TokenId"
            }
        }

        if(-not $tokenResp -or -not $tokenResp.access_token) {
            Write-Log "OAuth provider: refresh request returned no access_token for TokenId=$TokenId" 3
            return $null
        }

        $expiresAt = [DateTime]::UtcNow.AddMinutes(50)
        if($tokenResp.expires_in) {
            $expiresAt = [DateTime]::UtcNow.AddSeconds([int]$tokenResp.expires_in)
        }
        $entry.AccessToken = [string]$tokenResp.access_token
        if($tokenResp.refresh_token) { $entry.RefreshToken = [string]$tokenResp.refresh_token }
        $entry.ExpiresAt   = $expiresAt
        $entry.AcquiredAt  = [DateTime]::UtcNow
        [AuthenticationOAuth]::Tokens[$TokenId] = $entry

        Write-LogDebug "OAuth provider: refreshed TokenId=$TokenId (method=$($cred.AuthMethod), expires=$($expiresAt.ToLocalTime().ToString('s')))"

        # Re-persist the rotated refresh token for the browser flow when caching is on.
        if($cred.AuthMethod -eq 'AuthCode' -and $entry.RefreshToken -and ((Get-SettingValue "OAuthCacheToken") -eq $true)) {
            [void](Save-OAuthTokenCache -Data @{
                RefreshToken = $entry.RefreshToken
                ClientId     = $cred.ClientId
                TenantId     = $entry.TenantId
                Cloud        = $entry.Cloud
                Authority    = $cred.Authority
                Resource     = $cred.Resource
                AuthMethod   = 'AuthCode'
            })
        }

        Update-AuthToken -TokenId $TokenId
        return $entry.AccessToken
    }

    # CAE entry point — Invoke-MSGraphAPI calls this when Graph answers 401 with
    # a WWW-Authenticate claims challenge. Re-mints the token with the challenge
    # attached; returns the new access token, or $null when the flow can't
    # satisfy claims (BYO, managed identity, no refresh token).
    [string] AcquireWithClaims([int]$TokenId, [string]$ClaimsChallenge) {
        return $this.AcquireFromState($TokenId, $ClaimsChallenge)
    }

    # Satisfy a CAE claims challenge by re-running the credential flow with the claims
    # attached to the /token body. Returns $null when the flow can't satisfy the claims
    # (e.g. BYO / managed identity). OAuth has no interactive browser escalation, so
    # $AllowInteractive is not used.
    [string] GetClaimsToken([int]$TokenId, [string]$Resource, [string]$ClaimsChallenge, [bool]$AllowInteractive) {
        return $this.AcquireWithClaims($TokenId, $ClaimsChallenge)
    }

    # === Token retrieval ===

    [string] GetAccessToken([int]$TokenId, [string]$Resource) {
        if($TokenId -le 0) {
            # Caller didn't pin a TokenId; pick the most recently-acquired entry.
            if([AuthenticationOAuth]::Tokens.Count -eq 0) { return $null }
            $latest = [AuthenticationOAuth]::Tokens.Values | Sort-Object AcquiredAt -Descending | Select-Object -First 1
            $TokenId = $latest.Id
        }
        if(-not [AuthenticationOAuth]::Tokens.ContainsKey($TokenId)) { return $null }
        $entry = [AuthenticationOAuth]::Tokens[$TokenId]

        # 5-minute slack so we don't hand out a token that expires mid-request.
        if($entry.ExpiresAt -le [DateTime]::UtcNow.AddMinutes(5)) {
            $refreshed = $this.AcquireFromState($TokenId)
            if($refreshed) { return $refreshed }
            # Refresh failed; surface the stale token rather than $null and let
            # Invoke-MSGraphAPI handle the inevitable 401 — same behavior as MSAL.
        }
        return [string]$entry.AccessToken
    }

    [datetime] GetAccessTokenExpiry([int]$TokenId, [string]$Resource) {
        if(-not [AuthenticationOAuth]::Tokens.ContainsKey($TokenId)) {
            return [datetime]::MaxValue
        }
        # ExpiresAt is stored UTC; return LOCAL to match the provider contract (MSAL
        # returns ExpiresOn.LocalDateTime) and the local-time comparisons in
        # Invoke-MSGraphAPI / Test-DefaultTokenExpired. Returning raw UTC made callers
        # in ahead-of-UTC timezones see a just-issued token as already expired.
        return [AuthenticationOAuth]::Tokens[$TokenId].ExpiresAt.ToLocalTime()
    }

    # === Display / picker data ===

    [PSCustomObject] GetUserInfo([int]$TokenId) {
        if($TokenId -le 0 -and [AuthenticationOAuth]::Tokens.Count -gt 0) {
            $latest = [AuthenticationOAuth]::Tokens.Values | Sort-Object AcquiredAt -Descending | Select-Object -First 1
            $TokenId = $latest.Id
        }
        if(-not [AuthenticationOAuth]::Tokens.ContainsKey($TokenId)) { return $null }
        $entry = [AuthenticationOAuth]::Tokens[$TokenId]

        $upn = $null; $userId = $null; $appName = $null
        try {
            $jwt = Get-JWTtoken $entry.AccessToken
            if($jwt -and $jwt.Payload) {
                $upn    = if($jwt.Payload.upn) { [string]$jwt.Payload.upn }
                          elseif($jwt.Payload.preferred_username) { [string]$jwt.Payload.preferred_username }
                          elseif($jwt.Payload.unique_name) { [string]$jwt.Payload.unique_name }
                          else { $null }
                $userId = if($jwt.Payload.oid) { [string]$jwt.Payload.oid } else { $null }
                $appName = if($jwt.Payload.app_displayname) { [string]$jwt.Payload.app_displayname } else { $null }
            }
        } catch { }

        # Map AuthMethod to the cross-provider taxonomy (Interactive / ClientCredential / BYO / ...).
        $authType = switch ($entry.AuthMethod) {
            'BYO'          { 'BYO' }
            'DeviceCode'   { 'Interactive' }
            'AuthCode'     { 'Interactive' }
            'IMDS'         { 'ManagedIdentity' }
            'Federated'    { 'WorkloadFederation' }
            'Certificate'  { 'ClientCredential' }
            'ClientSecret' { 'ClientCredential' }
            'Password'     { 'Password' }
            default        { [string]$entry.AuthMethod }
        }

        # Caller-friendly default for headless flows: when there's no UPN
        # (app-only token), display the app id.
        $displayName = if($upn) { $upn } else { $entry.ClientId }

        # Tenant display name — one /organization GET per tenant per session,
        # mirroring AuthenticationMgGraph.TenantNameCache. App-only tokens
        # without Organization.Read.All fail the lookup; the $null result is
        # cached too so it isn't retried on every refresh event.
        $tenantName = $null
        if($entry.TenantId) {
            if([AuthenticationOAuth]::TenantNameCache.ContainsKey($entry.TenantId)) {
                $tenantName = [AuthenticationOAuth]::TenantNameCache[$entry.TenantId]
            }
            else {
                $tenantName = Get-OAuthTenantOrganizationName -Resource $entry.Resource -AccessToken $entry.AccessToken
                [AuthenticationOAuth]::TenantNameCache[$entry.TenantId] = $tenantName
            }
        }

        return [PSCustomObject]@{
            Provider    = $this.Id
            DisplayName = $displayName
            UPN         = $upn
            UserId      = $userId
            TenantId    = $entry.TenantId
            TenantName  = $tenantName
            AppId       = $entry.ClientId
            AppName     = $appName
            AuthType    = $authType
            ExpiresOn   = $entry.ExpiresAt.ToLocalTime()
        }
    }

    [PSCustomObject[]] GetCachedAccounts() {
        # Pure-headless: surface the token cache as the account list. Each entry
        # represents one Connect() call within the session.
        $rows = @()
        foreach($entry in [AuthenticationOAuth]::Tokens.Values) {
            $upn = $null
            try {
                $jwt = Get-JWTtoken $entry.AccessToken
                if($jwt -and $jwt.Payload) {
                    $upn = if($jwt.Payload.upn) { [string]$jwt.Payload.upn }
                           elseif($jwt.Payload.preferred_username) { [string]$jwt.Payload.preferred_username }
                           else { $null }
                }
            } catch { }
            $rows += [PSCustomObject]@{
                Provider = $this.Id
                Username = if($upn) { $upn } else { $entry.ClientId }
                UserId   = $null
                TenantId = $entry.TenantId
                Native   = $entry
            }
        }
        return [PSCustomObject[]]$rows
    }

    [PSCustomObject[]] GetAvailableTenants([int]$TokenId) {
        # Tenant enumeration costs an extra Graph call (/tenants or /me/memberOf
        # /transitive). Headless callers usually know which tenant they want;
        # leave this as a v2 enhancement.
        return [PSCustomObject[]]@()
    }

    # Resolve a -Certificate argument (thumbprint string OR X509Certificate2 OR
    # path-to-pfx) into a usable X509Certificate2. Reuses the MSAL provider's
    # resolver where available so behavior stays identical.
    Hidden [System.Security.Cryptography.X509Certificates.X509Certificate2] ResolveCertificate($CertObject, [string]$Path, $Password) {
        if($CertObject -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) {
            return $CertObject
        }
        if($CertObject) {
            # Treat as thumbprint string — search Cert: stores the same way MSAL does.
            if(Get-Command Resolve-MSALCertificate -ErrorAction SilentlyContinue) {
                $cert = Resolve-MSALCertificate $CertObject
                if($cert) { return $cert }
            }
            # Fallback: walk Cert:\CurrentUser\My then Cert:\LocalMachine\My.
            $thumb = ([string]$CertObject).Replace(' ', '').ToUpperInvariant()
            foreach($store in @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')) {
                $hit = Get-ChildItem $store -ErrorAction SilentlyContinue | Where-Object Thumbprint -EQ $thumb | Select-Object -First 1
                if($hit) { return $hit }
            }
            throw "Could not find certificate with thumbprint '$CertObject' in CurrentUser\My or LocalMachine\My"
        }
        if($Path) {
            if(-not (Test-Path $Path)) { throw "Certificate file not found: $Path" }
            return Get-PfxCertificate -FilePath $Path -Password $Password -ErrorAction Stop
        }
        throw "ResolveCertificate: neither object nor path was supplied"
    }
}
