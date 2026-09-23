# Module-level helpers for AuthenticationOAuth.
#
# PowerShell class methods bind type names at parse time, so [System.Net.Http.*]
# / [System.Security.Cryptography.*] usage that needs late binding (or any
# helper that calls module functions like Write-Log) lives here, not on the
# class. AuthenticationMgGraph follows the same split.
#
# Functions:
#   Invoke-OAuthTokenRequest        — POSTs to /oauth2/v2.0/token, surfaces Entra error fields
#   Invoke-OAuthDeviceCodeFlow      — device code grant: request code, poll /token until signed in
#   Invoke-OAuthAuthCodeFlow        — browser auth-code + PKCE over a loopback HttpListener redirect
#   Get-OAuthFreePort               — OS-assigned free loopback port for the redirect listener
#   Save/Read/Clear-OAuthTokenCache — opt-in DPAPI refresh-token cache (Remember login)
#   Get-OAuthClientAssertion        — builds + signs the private_key_jwt for cert auth
#   Invoke-OAuthIMDS                — Managed identity: IMDS / App Service / Functions
#   Get-OAuthFederatedAssertion     — workload identity federation token loader
#   Get-OAuthErrorBody              — parse the JSON error body off a failed web request (PS5.1 + PS7)
#   Get-OAuthAADSTSHint             — map common AADSTS error codes to actionable messages
#   ConvertTo-OAuthBase64Url        — bytes → base64url string
#   ConvertFrom-OAuthBase64Url      — base64url string → bytes

function ConvertTo-OAuthBase64Url
{
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][byte[]]$Bytes)
    $b64 = [Convert]::ToBase64String($Bytes)
    return $b64.Replace('+', '-').Replace('/', '_').TrimEnd('=')
}

function ConvertFrom-OAuthBase64Url
{
    [CmdletBinding()]
    [OutputType([byte[]])]
    param([Parameter(Mandatory)][string]$String)
    $s = $String.Replace('-', '+').Replace('_', '/')
    while($s.Length % 4) { $s += '=' }
    return [Convert]::FromBase64String($s)
}

# Build and sign a client_assertion JWT for the private_key_jwt flow. Entra
# rejects with AADSTS700027 if x5t isn't the SHA-1 thumbprint base64url-encoded
# (NOT hex), if nbf/exp aren't unix seconds, or if aud isn't the tenant-specific
# token endpoint. Encoding follows RFC 7523 §3.
function Get-OAuthClientAssertion
{
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$Authority,
        [Parameter(Mandatory)][string]$TenantId
    )

    if(-not $Certificate.HasPrivateKey) {
        throw "Certificate '$($Certificate.Subject)' has no associated private key - cannot sign client assertion"
    }

    # SHA-1 thumbprint of the DER cert, base64url encoded. The x5t header tells
    # Entra which key in the app's keyCredentials list signed this assertion.
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $thumbBytes = $sha1.ComputeHash($Certificate.RawData)
    }
    finally {
        $sha1.Dispose()
    }
    $x5t = ConvertTo-OAuthBase64Url -Bytes $thumbBytes

    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $exp = $now + 600   # 10-minute lifetime is plenty for one /token round-trip

    $header = [ordered]@{
        alg = 'RS256'
        typ = 'JWT'
        x5t = $x5t
    }
    $payload = [ordered]@{
        aud = "https://$Authority/$TenantId/oauth2/v2.0/token"
        iss = $ClientId
        sub = $ClientId
        jti = [Guid]::NewGuid().ToString()
        nbf = $now
        exp = $exp
        iat = $now
    }

    $headerJson  = $header  | ConvertTo-Json -Compress
    $payloadJson = $payload | ConvertTo-Json -Compress
    $headerB64   = ConvertTo-OAuthBase64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes($headerJson))
    $payloadB64  = ConvertTo-OAuthBase64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes($payloadJson))
    $signing     = "$headerB64.$payloadB64"

    # Pull the RSA private key. Newer PFX imports expose it via GetRSAPrivateKey;
    # legacy PrivateKey property still works on older runtimes.
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if(-not $rsa) {
        throw "Could not access RSA private key on certificate '$($Certificate.Subject)' - is the key exportable?"
    }
    try {
        $sigBytes = $rsa.SignData(
            [System.Text.Encoding]::UTF8.GetBytes($signing),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    }
    finally {
        # GetRSAPrivateKey returns a fresh disposable handle (per CryptoAPI docs)
        # — disposing here releases the CSP/CNG resource without touching the
        # certificate itself.
        try { $rsa.Dispose() } catch { }
    }

    $sigB64 = ConvertTo-OAuthBase64Url -Bytes $sigBytes
    return "$signing.$sigB64"
}

# Parse the JSON error body off a failed Invoke-RestMethod call. PS7 surfaces
# the body in ErrorDetails.Message (HttpResponseException has no response
# stream); PS5.1 needs the classic GetResponseStream read. Returns $null when
# there's no parsable JSON body.
function Get-OAuthErrorBody
{
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param([Parameter(Mandatory)]$ErrorRecord)

    try {
        if($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
            return ($ErrorRecord.ErrorDetails.Message | ConvertFrom-Json)
        }
    } catch { }
    try {
        $stream = $ErrorRecord.Exception.Response.GetResponseStream()
        if($stream) {
            $reader = New-Object IO.StreamReader($stream)
            return ($reader.ReadToEnd() | ConvertFrom-Json)
        }
    } catch { }
    return $null
}

# Actionable messages for the AADSTS codes automation users hit most. Keyed by
# the numeric code Entra puts in error_codes (and in the AADSTSnnnnn prefix of
# error_description).
$script:OAuthAADSTSHints = @{
    700016  = "The app registration (client id) was not found in this tenant. Check -AppId and -TenantId."
    700027  = "Client assertion signature validation failed - the certificate does not match any key on the app registration. Upload the certificate's public key to the app, or check -Certificate / -CertificatePath."
    7000215 = "Invalid client secret. The secret is wrong or expired - create a new client secret on the app registration."
    50126   = "Wrong username or password (ROPC). Check the -Credential values."
    50076   = "Multi-factor authentication is required for this account, which ROPC cannot satisfy. Use -DeviceCode instead."
    50034   = "The user account was not found in this tenant. Check the username's domain and -TenantId."
    65001   = "Admin consent is missing for a delegated scope. Grant consent to the application in Entra."
    50105   = "The signed-in user is not assigned to the application. Assign the user (or a group) to the app in Entra."
    70011   = "Invalid scope value. The scope must look like 'https://graph.microsoft.com/.default'."
}

# Return the human hint for the first recognized AADSTS code on a parsed Entra
# error body (error_codes array preferred, AADSTSnnnnn prefix in
# error_description as fallback). $null when nothing matches.
function Get-OAuthAADSTSHint
{
    [CmdletBinding()]
    [OutputType([string])]
    param($ErrorDetail)

    if(-not $ErrorDetail) { return $null }

    $codes = @()
    if($ErrorDetail.error_codes) { $codes += @($ErrorDetail.error_codes | ForEach-Object { [int]$_ }) }
    if($ErrorDetail.error_description) {
        $m = [regex]::Match([string]$ErrorDetail.error_description, 'AADSTS(\d+)')
        if($m.Success) { $codes += [int]$m.Groups[1].Value }
    }

    foreach($code in $codes) {
        if($script:OAuthAADSTSHints.ContainsKey($code)) {
            return $script:OAuthAADSTSHints[$code]
        }
    }
    return $null
}

# POST to /oauth2/v2.0/token. Entra returns 4xx with a JSON body that includes
# `error`, `error_description`, `error_codes`, `correlation_id` and `timestamp`
# — surface those as a single error so the caller's log shows what Entra
# actually said instead of just "Bad Request". Known AADSTS codes get an
# actionable hint appended (the original Entra wording is always kept).
#
# -Claims carries a CAE claims challenge from a Graph 401 WWW-Authenticate
# header. The header value is base64-encoded JSON; /token wants the decoded
# JSON in the `claims` form field. Raw JSON is passed through unchanged.
function Invoke-OAuthTokenRequest
{
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)][string]$Authority,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][hashtable]$Body,
        [string]$Claims
    )

    if($Claims) {
        $claimsJson = $Claims
        if(-not $Claims.TrimStart().StartsWith('{')) {
            try {
                $decoded = [System.Text.Encoding]::UTF8.GetString((ConvertFrom-OAuthBase64Url -String $Claims))
                if($decoded.TrimStart().StartsWith('{')) { $claimsJson = $decoded }
            } catch { }
        }
        $Body = $Body.Clone()
        $Body['claims'] = $claimsJson
    }

    $url = "https://$Authority/$TenantId/oauth2/v2.0/token"
    Write-LogDebug "OAuth /token POST -> $url (grant: $($Body.grant_type)$(if($Claims) { ', with claims challenge' }))"

    try {
        $resp = Invoke-RestMethod -Method POST -Uri $url -Body $Body `
            -ContentType 'application/x-www-form-urlencoded' `
            -ErrorAction Stop
        return $resp
    }
    catch {
        $detail = Get-OAuthErrorBody -ErrorRecord $_

        if($detail -and $detail.error) {
            $msg = "Entra rejected token request: $($detail.error)"
            if($detail.error_description) { $msg += " - $($detail.error_description -replace '\r?\n',' ')" }
            $hint = Get-OAuthAADSTSHint -ErrorDetail $detail
            if($hint)                     { $msg += " Hint: $hint" }
            if($detail.correlation_id)    { $msg += " (correlation: $($detail.correlation_id))" }
            throw $msg
        }
        throw   # rethrow original — no JSON body to parse
    }
}

# Device code grant (RFC 8628). POST /devicecode to get a user_code +
# verification URI, show them to the user, then poll /token with the
# device_code until the user finishes signing in on their other device.
# Returns the /token response (access_token + refresh_token when the scope
# includes offline_access). Throws on decline, timeout, or a hard error.
function Invoke-OAuthDeviceCodeFlow
{
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)][string]$Authority,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$Scope
    )

    $dcUrl = "https://$Authority/$TenantId/oauth2/v2.0/devicecode"
    Write-LogDebug "OAuth /devicecode POST -> $dcUrl"
    try {
        $dc = Invoke-RestMethod -Method POST -Uri $dcUrl `
            -Body @{ client_id = $ClientId; scope = $Scope } `
            -ContentType 'application/x-www-form-urlencoded' `
            -ErrorAction Stop
    }
    catch {
        $detail = Get-OAuthErrorBody -ErrorRecord $_
        if($detail -and $detail.error) {
            $msg = "Entra rejected device code request: $($detail.error)"
            if($detail.error_description) { $msg += " - $($detail.error_description -replace '\r?\n',' ')" }
            $hint = Get-OAuthAADSTSHint -ErrorDetail $detail
            if($hint) { $msg += " Hint: $hint" }
            throw $msg
        }
        throw
    }

    if(-not $dc.device_code) { throw "Device code request returned no device_code" }

    # Entra's message field is the canonical user instruction ("To sign in, use
    # a web browser to open ... and enter the code ..."). Write it both to the
    # console (the user has to act on it NOW) and the log.
    $instruction = if($dc.message) { [string]$dc.message }
                   else { "To sign in, open $($dc.verification_uri) in a browser and enter the code $($dc.user_code)" }
    Write-Log "OAuth device code: $instruction"
    Write-Host $instruction -ForegroundColor Yellow

    $interval = if($dc.interval) { [int]$dc.interval } else { 5 }
    $lifetime = if($dc.expires_in) { [int]$dc.expires_in } else { 900 }
    $deadline = [DateTime]::UtcNow.AddSeconds($lifetime)
    $tokenUrl = "https://$Authority/$TenantId/oauth2/v2.0/token"

    # This loop has nothing cancellable to stop - it is a synchronous poll - so the
    # action is a no-op and cancelling just breaks the loop. Passing -OnCancel still
    # matters: it resets a cancel request left over from an earlier operation.
    # This wait can run for the full code lifetime (15 min by default).
    Write-Status "Waiting for sign-in" $instruction -CancelText "Cancel" -OnCancel { }
    try {
    while([DateTime]::UtcNow -lt $deadline) {
        # Pumped, cancellable wait - a bare Start-Sleep here would freeze the UI for
        # the whole interval and the Cancel button would never get its click.
        if(Wait-StatusCancel -Seconds $interval) {
            # OperationCanceledException, like the browser flow: user cancellation is
            # not a failure and callers must be able to tell the two apart.
            throw [System.OperationCanceledException]::new("Device code sign-in was cancelled")
        }
        try {
            return Invoke-RestMethod -Method POST -Uri $tokenUrl `
                -Body @{
                    grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                    client_id   = $ClientId
                    device_code = $dc.device_code
                } `
                -ContentType 'application/x-www-form-urlencoded' `
                -ErrorAction Stop
        }
        catch {
            $detail = Get-OAuthErrorBody -ErrorRecord $_
            $errCode = if($detail) { [string]$detail.error } else { $null }
            switch($errCode) {
                'authorization_pending' { }                          # user hasn't finished yet — keep polling
                'slow_down'             { $interval += 5 }           # RFC 8628 §3.5
                'authorization_declined' { throw "Device code sign-in was declined by the user" }
                'expired_token'          { throw "Device code expired before sign-in completed - run Connect-IntuneManagement -DeviceCode again" }
                default {
                    if($detail -and $detail.error) {
                        $msg = "Device code polling failed: $($detail.error)"
                        if($detail.error_description) { $msg += " - $($detail.error_description -replace '\r?\n',' ')" }
                        throw $msg
                    }
                    throw
                }
            }
        }
    }
    throw "Device code expired before sign-in completed - run Connect-IntuneManagement -DeviceCode again"
    }
    finally {
        # Clear the overlay and disarm on every exit - success, expiry, decline or
        # cancel - so the Cancel button never outlives this wait.
        Write-Status ""
    }
}

# Pick an OS-assigned free TCP port on the loopback interface for the redirect
# listener (bind port 0, read what the OS handed out, release it). There is a tiny
# race between release and HttpListener.Start(), but on loopback it is negligible.
function Get-OAuthFreePort
{
    [CmdletBinding()]
    [OutputType([int])]
    param()
    $l = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    try { $l.Start(); return [int]$l.LocalEndpoint.Port } finally { $l.Stop() }
}

# Authorization Code + PKCE browser login (native-app pattern, .NET BCL only - no
# MSAL). Opens the system default browser to /authorize, captures the redirect on a
# loopback HttpListener, then exchanges the code at /token. Returns the SAME
# /token response shape as Invoke-OAuthDeviceCodeFlow ({access_token, refresh_token,
# expires_in, ...}) so AuthenticationOAuth.Connect handles both identically.
#
# The redirect wait polls GetContextAsync and pumps the UI (Invoke-UIPump) so the
# window stays responsive, with a wall-clock timeout - same model as the MSAL
# interactive poll (Get-MsalAuthenticationToken). Throws on timeout / browser-launch
# failure / OAuth error so the caller can fall back to device code when headless.
function Invoke-OAuthAuthCodeFlow
{
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)][string]$Authority,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$Scope,
        [int]$RedirectPort = 0,
        [string]$Prompt,
        [string]$LoginHint,
        [int]$TimeoutSec = 180
    )

    # --- PKCE (RFC 7636, S256) ---
    $verifierBytes = New-Object byte[] 32
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($verifierBytes) } finally { $rng.Dispose() }
    $verifier   = ConvertTo-OAuthBase64Url -Bytes $verifierBytes
    $sha        = [System.Security.Cryptography.SHA256]::Create()
    try { $challengeBytes = $sha.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($verifier)) } finally { $sha.Dispose() }
    $challenge  = ConvertTo-OAuthBase64Url -Bytes $challengeBytes

    # --- state (CSRF) ---
    $stateBytes = New-Object byte[] 16
    $rng2 = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng2.GetBytes($stateBytes) } finally { $rng2.Dispose() }
    $state = ConvertTo-OAuthBase64Url -Bytes $stateBytes

    $port = if($RedirectPort -gt 0) { $RedirectPort } else { Get-OAuthFreePort }
    $redirectUri = "http://localhost:$port/"

    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add($redirectUri)
    try {
        $listener.Start()
    }
    catch {
        throw "Could not start the loopback redirect listener on $redirectUri : $($_.Exception.Message)"
    }

    try {
        # --- build /authorize URL ---
        $q = [System.Collections.Generic.List[string]]::new()
        $q.Add("client_id=" + [System.Uri]::EscapeDataString($ClientId))
        $q.Add("response_type=code")
        $q.Add("redirect_uri=" + [System.Uri]::EscapeDataString($redirectUri))
        $q.Add("response_mode=query")
        $q.Add("scope=" + [System.Uri]::EscapeDataString($Scope))
        $q.Add("code_challenge=" + [System.Uri]::EscapeDataString($challenge))
        $q.Add("code_challenge_method=S256")
        $q.Add("state=" + [System.Uri]::EscapeDataString($state))
        if($Prompt)    { $q.Add("prompt=" + [System.Uri]::EscapeDataString($Prompt)) }
        if($LoginHint) { $q.Add("login_hint=" + [System.Uri]::EscapeDataString($LoginHint)) }
        $authorizeUrl = "https://$Authority/$TenantId/oauth2/v2.0/authorize?" + ($q -join '&')

        Write-LogDebug "OAuth /authorize -> redirect_uri=$redirectUri prompt=$Prompt hint=$LoginHint"

        # --- launch the default browser ---
        try {
            Start-Process $authorizeUrl -ErrorAction Stop | Out-Null
        }
        catch {
            throw "Could not open the default browser for sign-in: $($_.Exception.Message)"
        }

        # HttpListener.GetContextAsync() takes no CancellationToken, so cancelling
        # means stopping the listener out from under it - the task then faults and
        # the loop below exits on the request flag rather than on IsCompleted.
        # The status click runs later, outside this statement's local scope.
        # Preserve the listener explicitly so cancellation always stops the
        # GetContextAsync wait rather than merely setting the polled flag.
        $cancelListener = $listener
        Set-StatusCancelAction ({ try { $cancelListener.Stop() } catch { } }.GetNewClosure())
        Write-Status "Waiting for browser sign-in..." -CancelText "Cancel"

        # --- wait for the redirect (responsive poll + timeout) ---
        $ctxTask  = $listener.GetContextAsync()
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSec)
        while(-not $ctxTask.IsCompleted) {
            if(Test-StatusCancelRequested) {
                throw [System.OperationCanceledException]::new("Browser sign-in was cancelled")
            }
            if([DateTime]::UtcNow -gt $deadline) {
                throw "Browser sign-in timed out after $TimeoutSec s"
            }
            Invoke-UIPump
            Start-Sleep -Milliseconds 100
        }
        # Cancel also completes the task - stopping the listener faults it - so the
        # loop above can exit without ever seeing the flag when the click lands
        # during a sleep slice. Re-check before touching the result, or a cancelled
        # sign-in would surface as an HttpListener error instead, and the caller
        # would treat it as "browser broken" and fall back to device code.
        if(Test-StatusCancelRequested) {
            throw [System.OperationCanceledException]::new("Browser sign-in was cancelled")
        }
        $context = $ctxTask.GetAwaiter().GetResult()
        $req     = $context.Request

        # Respond so the browser tab shows a friendly message, then release.
        $html = "<html><head><title>Sign-in complete</title></head><body style='font-family:Segoe UI,sans-serif;padding:2em'><h3>Sign-in complete</h3><p>You can close this tab and return to IntuneManagement.</p></body></html>"
        try {
            $buf = [System.Text.Encoding]::UTF8.GetBytes($html)
            $context.Response.ContentType = "text/html; charset=utf-8"
            $context.Response.ContentLength64 = $buf.Length
            $context.Response.OutputStream.Write($buf, 0, $buf.Length)
            $context.Response.OutputStream.Close()
        }
        catch { }

        # HttpListenerRequest.QueryString is a populated NameValueCollection (no
        # System.Web dependency needed).
        $qs         = $req.QueryString
        $returnState = [string]$qs['state']
        $errCode     = [string]$qs['error']
        $code        = [string]$qs['code']

        if($errCode) {
            $errDesc = [string]$qs['error_description']
            throw "Browser sign-in failed: $errCode$(if($errDesc) { " - $($errDesc -replace '\r?\n',' ')" })"
        }
        if($returnState -ne $state) {
            throw "Browser sign-in state mismatch - possible CSRF; aborting."
        }
        if(-not $code) {
            throw "Browser sign-in returned no authorization code."
        }

        # --- exchange the code for tokens ---
        return Invoke-OAuthTokenRequest -Authority $Authority -TenantId $TenantId -Body @{
            grant_type    = 'authorization_code'
            client_id     = $ClientId
            code          = $code
            redirect_uri  = $redirectUri
            code_verifier = $verifier
            scope         = $Scope
        }
    }
    finally {
        # Disarm first: a click arriving after this point must not reach a listener
        # that is about to be closed.
        Clear-StatusCancelAction
        try { $listener.Stop(); $listener.Close() } catch { }
    }
}

# Managed identity acquisition. Detects the host environment in this order:
#
#   1. App Service / Functions / Container Apps
#      — IDENTITY_ENDPOINT + IDENTITY_HEADER env vars set; api-version=2019-08-01
#   2. Azure Arc-enabled servers
#      — IDENTITY_ENDPOINT + IMDS_ENDPOINT env vars (challenge-response, not yet
#        supported here; treat as App Service variant if header is set)
#   3. Azure VM / VMSS / AKS host node
#      — IMDS at 169.254.169.254 with Metadata: true; api-version=2018-02-01
#
# AKS pods that opt into Workload Identity Federation should NOT use this path —
# they should pass the projected token via Get-OAuthFederatedAssertion instead.
function Invoke-OAuthIMDS
{
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)][string]$Resource,
        [string]$ClientId
    )

    $appSvcEndpoint = $env:IDENTITY_ENDPOINT
    $appSvcHeader   = $env:IDENTITY_HEADER

    if($appSvcEndpoint -and $appSvcHeader) {
        # App Service / Functions / Container Apps flow.
        $params = @{
            'api-version' = '2019-08-01'
            'resource'    = $Resource
        }
        if($ClientId) { $params['client_id'] = $ClientId }
        $query = ($params.GetEnumerator() | ForEach-Object { "$([uri]::EscapeDataString($_.Key))=$([uri]::EscapeDataString($_.Value))" }) -join '&'
        $url = "$appSvcEndpoint`?$query"
        Write-LogDebug "OAuth IMDS (App Service) -> $url"
        return Invoke-RestMethod -Method GET -Uri $url `
            -Headers @{ 'X-IDENTITY-HEADER' = $appSvcHeader } `
            -ErrorAction Stop
    }

    # VM / VMSS / AKS host-node flow.
    $params = @{
        'api-version' = '2018-02-01'
        'resource'    = $Resource
    }
    if($ClientId) { $params['client_id'] = $ClientId }
    $query = ($params.GetEnumerator() | ForEach-Object { "$([uri]::EscapeDataString($_.Key))=$([uri]::EscapeDataString($_.Value))" }) -join '&'
    $url = "http://169.254.169.254/metadata/identity/oauth2/token`?$query"
    Write-LogDebug "OAuth IMDS (VM) -> $url"
    return Invoke-RestMethod -Method GET -Uri $url `
        -Headers @{ Metadata = 'true' } `
        -TimeoutSec 5 `
        -ErrorAction Stop
}

# Workload identity federation — the assertion is already a signed JWT minted
# by the workload's identity provider (AKS service-account token, GitHub
# Actions OIDC token, Azure DevOps OIDC token, etc.). We don't sign or
# validate it; we just hand it to Entra as the client_assertion.
function Get-OAuthFederatedAssertion
{
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [string]$FederatedToken,
        [string]$FederatedTokenFile
    )

    if($FederatedToken) { return $FederatedToken.Trim() }
    if(-not $FederatedTokenFile) {
        throw "Federated assertion required but neither -FederatedToken nor -FederatedTokenFile was supplied"
    }
    if(-not (Test-Path $FederatedTokenFile)) {
        throw "Federated token file not found: $FederatedTokenFile"
    }
    $tok = ([IO.File]::ReadAllText($FederatedTokenFile)).Trim()
    if(-not $tok) {
        throw "Federated token file '$FederatedTokenFile' is empty"
    }
    if(-not $tok.StartsWith('eyJ')) {
        # Not strictly required to be a JWT but every real federated token IS one;
        # warn loudly if it isn't, so a misconfigured file fails fast rather than
        # surfacing as AADSTS50027 from the wire.
        Write-Log "Federated token from '$FederatedTokenFile' does not look like a JWT (missing 'eyJ' prefix) - passing through anyway" 2
    }
    return $tok
}

# Resolve the tenant's display name via /organization with a raw bearer token.
# Returns $null on any failure (e.g. app-only token without
# Organization.Read.All) — the caller caches the result either way so a failed
# lookup isn't retried.
function Get-OAuthTenantOrganizationName
{
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$Resource,
        [Parameter(Mandatory)][string]$AccessToken
    )

    try {
        $org = Invoke-RestMethod -Method GET -Uri "$Resource/v1.0/organization?`$select=displayName" `
            -Headers @{ Authorization = "Bearer $AccessToken" } -ErrorAction Stop
        if($org -and $org.value -and @($org.value).Count -gt 0 -and $org.value[0].displayName) {
            return [string]$org.value[0].displayName
        }
    }
    catch {
        Write-LogDebug "OAuth: /organization lookup failed ($($_.Exception.Message)); tenant display name will be unknown"
    }
    return $null
}

# ===================== OAuth token cache (DPAPI, opt-in) =====================
#
# Optional cross-restart persistence for the browser/auth-code refresh token,
# gated by the OAuthCacheToken ("Remember login") setting. DPAPI (ProtectedData,
# CurrentUser scope) - same protection the MSAL cache uses - written to
# %LOCALAPPDATA%\IntuneManagement\oauthcache.bin. Only the refresh token + minimal
# metadata is stored; never an access token. Windows-only; degrades to no-op
# (in-memory session only) elsewhere or if the DPAPI type can't be loaded.

# App-specific entropy so the blob isn't decryptable by other apps' CurrentUser DPAPI.
$script:_oauthCacheEntropy = [System.Text.Encoding]::UTF8.GetBytes("IntuneManagement.OAuth.TokenCache.v1")

function Get-OAuthTokenCachePath {
    $folder = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) "IntuneManagement"
    return (Join-Path $folder "oauthcache.bin")
}

# Resolve System.Security.Cryptography.ProtectedData across PS 5.1 (System.Security)
# and PS7 (separate assembly, may need loading from $PSHOME). Returns the type or
# $null (persistence then silently disabled).
function Get-OAuthProtectedDataType {
    $t = 'System.Security.Cryptography.ProtectedData' -as [type]
    if($t) { return $t }
    if(-not $script:IsWindowsOS) { return $null }
    foreach($asm in 'System.Security','System.Security.Cryptography.ProtectedData') {
        try { Add-Type -AssemblyName $asm -ErrorAction Stop } catch { }
        $t = 'System.Security.Cryptography.ProtectedData' -as [type]
        if($t) { return $t }
    }
    try {
        $dll = Join-Path $PSHOME 'System.Security.Cryptography.ProtectedData.dll'
        if(Test-Path -LiteralPath $dll) { Add-Type -Path $dll -ErrorAction Stop }
    } catch { }
    return ('System.Security.Cryptography.ProtectedData' -as [type])
}

# Persist the minimal refresh-token state. $Data is a hashtable
# (RefreshToken/ClientId/TenantId/Cloud/Authority/Resource/AuthMethod).
function Save-OAuthTokenCache {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Data)
    $pd = Get-OAuthProtectedDataType
    if(-not $pd) { Write-LogDebug "OAuth cache: DPAPI unavailable - not persisting"; return $false }
    try {
        $path   = Get-OAuthTokenCachePath
        $folder = Split-Path -Parent $path
        if(-not (Test-Path -LiteralPath $folder)) { [void][IO.Directory]::CreateDirectory($folder) }
        $json   = ($Data | ConvertTo-Json -Depth 5 -Compress)
        $plain  = [System.Text.Encoding]::UTF8.GetBytes($json)
        $scope  = [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        $prot   = $pd::Protect($plain, $script:_oauthCacheEntropy, $scope)
        [IO.File]::WriteAllBytes($path, $prot)
        Write-LogDebug "OAuth cache: saved refresh-token state to $path"
        return $true
    }
    catch { Write-LogError "OAuth cache: failed to save token cache" $_.Exception; return $false }
}

# Returns the persisted hashtable, or $null when absent/unreadable.
function Read-OAuthTokenCache {
    [CmdletBinding()]
    param()
    $path = Get-OAuthTokenCachePath
    if(-not (Test-Path -LiteralPath $path)) { return $null }
    $pd = Get-OAuthProtectedDataType
    if(-not $pd) { return $null }
    try {
        $prot  = [IO.File]::ReadAllBytes($path)
        $scope = [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        $plain = $pd::Unprotect($prot, $script:_oauthCacheEntropy, $scope)
        $json  = [System.Text.Encoding]::UTF8.GetString($plain)
        $obj   = $json | ConvertFrom-Json
        $ht = @{}
        foreach($p in $obj.PSObject.Properties) { $ht[$p.Name] = $p.Value }
        return $ht
    }
    catch { Write-LogError "OAuth cache: failed to read token cache (removing it)" $_.Exception; Clear-OAuthTokenCache; return $null }
}

function Clear-OAuthTokenCache {
    [CmdletBinding()]
    param()
    try {
        $path = Get-OAuthTokenCachePath
        if(Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction Stop; Write-LogDebug "OAuth cache: cleared $path" }
    }
    catch { Write-LogError "OAuth cache: failed to clear token cache" $_.Exception }
}

# OAuth provider has zero dependencies (no DLL, no SDK module) so it always
# registers — unlike MgGraph which has to defer until the SDK module is present.
function Invoke-OAuthProviderInitialize {
    [CmdletBinding()]
    param()

    if(-not (Get-Command -Name Register-AuthProvider -ErrorAction SilentlyContinue)) {
        Write-LogDebug "OAuth provider: AuthenticationCore not available, skipping registration"
        return
    }

    try {
        Register-AuthProvider -Provider ([AuthenticationOAuth]::new())
    }
    catch {
        Write-LogError "Failed to register AuthenticationOAuth provider" $_.Exception
    }
}

Invoke-OAuthProviderInitialize
