function Connect-IntuneManagement {
    <#
    .SYNOPSIS
        Authenticate to Microsoft Graph for use with IntuneManagement.

    .DESCRIPTION
        Supports four non-interactive authentication methods:
          - App registration with a client secret
          - App registration with a certificate (thumbprint, X509Certificate2 object, or .pfx file)
          - Bring-your-own token (pass a raw Bearer token string)
          - Managed Identity (Azure VM / Azure Function workload identity)

        The auth backend is selected via -Provider:
          - MSAL    (default; uses the bundled MSAL.NET DLLs)
          - MgGraph (uses the Microsoft.Graph.Authentication PowerShell module; the
                    module must be installed: Install-Module Microsoft.Graph.Authentication)

        If -Provider is omitted, the value of the "ActiveAuthProvider" setting is
        used (default MSAL).

    .EXAMPLE
        # Client secret with the default provider (MSAL)
        Connect-IntuneManagement -TenantId "contoso.onmicrosoft.com" -AppId "00000000-..." -Secret "abc123"

    .EXAMPLE
        # Client secret via the Microsoft.Graph SDK provider
        Connect-IntuneManagement -Provider MgGraph -TenantId "contoso.onmicrosoft.com" -AppId "00000000-..." -Secret "abc123"

    .EXAMPLE
        # Certificate thumbprint (looked up in Cert:\CurrentUser\My then Cert:\LocalMachine\My)
        Connect-IntuneManagement -TenantId "contoso.onmicrosoft.com" -AppId "00000000-..." -Certificate "A1B2C3..."

    .EXAMPLE
        # Certificate from .pfx file
        Connect-IntuneManagement -TenantId "contoso.onmicrosoft.com" -AppId "00000000-..." `
            -CertificatePath "C:\certs\app.pfx" -CertificatePassword (ConvertTo-SecureString "pass" -AsPlainText -Force)

    .EXAMPLE
        # Bring your own Graph Bearer token
        Connect-IntuneManagement -Token $myToken

    .EXAMPLE
        # System-assigned managed identity (Azure VM / Azure Function)
        Connect-IntuneManagement -Provider MgGraph -ManagedIdentity

    .EXAMPLE
        # User-assigned managed identity
        Connect-IntuneManagement -Provider MgGraph -ManagedIdentity -AppId "00000000-..."

    .EXAMPLE
        # Direct-OAuth provider (no SDK, no DLL): client secret
        Connect-IntuneManagement -Provider OAuth -TenantId "contoso.onmicrosoft.com" -AppId "00000000-..." -Secret "abc"

    .EXAMPLE
        # Direct-OAuth provider: workload identity federation (AKS / GitHub Actions OIDC)
        Connect-IntuneManagement -Provider OAuth -TenantId "..." -AppId "..." `
            -FederatedTokenFile $env:AZURE_FEDERATED_TOKEN_FILE

    .EXAMPLE
        # Direct-OAuth provider: PSCredential (ROPC; non-MFA accounts only)
        Connect-IntuneManagement -Provider OAuth -TenantId "..." -AppId "..." -Credential (Get-Credential)

    .EXAMPLE
        # Device code sign-in on the default provider (MSAL). MFA / FIDO2 /
        # YubiKey capable; the browser auth happens on any other device.
        Connect-IntuneManagement -DeviceCode

    .EXAMPLE
        # Device code sign-in on the Direct-OAuth provider. Uses the app id
        # selected in Settings -> Entra, unless -AppId is supplied.
        Connect-IntuneManagement -Provider OAuth -DeviceCode

    .EXAMPLE
        # Device code sign-in on the Microsoft.Graph SDK provider
        Connect-IntuneManagement -Provider MgGraph -DeviceCode

    .EXAMPLE
        # Interactive sign-in (browser popup / WAM broker on the default MSAL
        # provider). Silent-from-cache first, falls back to browser prompt.
        Connect-IntuneManagement -Interactive

    .EXAMPLE
        # Interactive against a specific tenant, force a fresh browser prompt
        Connect-IntuneManagement -Interactive -TenantId "contoso.onmicrosoft.com" -ForceInteractive

    .EXAMPLE
        # Interactive against a sovereign cloud
        Connect-IntuneManagement -Interactive -Cloud USGov

    .EXAMPLE
        # Interactive against the Direct-OAuth provider — no browser popup path
        # exists in that provider, so this transparently routes to device code
        # (headless-friendly, MFA/FIDO2 capable). Uses the app id selected in
        # Settings -> Entra, unless -AppId is supplied.
        Connect-IntuneManagement -Provider OAuth -Interactive
    #>
    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true,  ParameterSetName = 'Secret')]
        [Parameter(Mandatory = $true,  ParameterSetName = 'Certificate')]
        [Parameter(Mandatory = $true,  ParameterSetName = 'CertificatePath')]
        [Parameter(Mandatory = $true,  ParameterSetName = 'OAuthFederated')]
        [Parameter(Mandatory = $true,  ParameterSetName = 'OAuthCredential')]
        [Parameter(Mandatory = $false, ParameterSetName = 'DeviceCode')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Token')]
        [Parameter(Mandatory = $false, ParameterSetName = 'ManagedIdentity')]
        [string]$TenantId,

        [Parameter(Mandatory = $true,  ParameterSetName = 'Secret')]
        [Parameter(Mandatory = $true,  ParameterSetName = 'Certificate')]
        [Parameter(Mandatory = $true,  ParameterSetName = 'CertificatePath')]
        [Parameter(Mandatory = $true,  ParameterSetName = 'OAuthFederated')]
        [Parameter(Mandatory = $true,  ParameterSetName = 'OAuthCredential')]
        [Parameter(Mandatory = $false, ParameterSetName = 'DeviceCode')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [Parameter(Mandatory = $false, ParameterSetName = 'ManagedIdentity')]
        [string]$AppId,

        [Parameter(Mandatory = $true, ParameterSetName = 'Secret')]
        [string]$Secret,

        # Thumbprint string or X509Certificate2 object
        [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
        $Certificate,

        [Parameter(Mandatory = $true, ParameterSetName = 'CertificatePath')]
        [string]$CertificatePath,

        [Parameter(Mandatory = $false, ParameterSetName = 'CertificatePath')]
        [SecureString]$CertificatePassword,

        [Parameter(Mandatory = $true, ParameterSetName = 'Token')]
        [string]$Token,

        # System-assigned (no -AppId) or user-assigned (with -AppId) managed identity.
        # MgGraph and OAuth providers support this; MSAL rejects with a clear error.
        [Parameter(Mandatory = $true, ParameterSetName = 'ManagedIdentity')]
        [switch]$ManagedIdentity,

        # OAuth provider only — workload identity federation. Path to a file
        # containing an OIDC JWT to exchange at the /token endpoint as
        # client_assertion (jwt-bearer). Used by AKS Workload Identity, GitHub
        # Actions OIDC, Azure DevOps OIDC, etc.
        [Parameter(Mandatory = $true,  ParameterSetName = 'OAuthFederated')]
        [string]$FederatedTokenFile,

        # OAuth provider only — pass the federated assertion inline (alternative
        # to -FederatedTokenFile when the caller already has the JWT in memory).
        [Parameter(Mandatory = $false, ParameterSetName = 'OAuthFederated')]
        [string]$FederatedToken,

        # OAuth provider only — username/password (ROPC) via PSCredential. Limited
        # to non-MFA accounts; intended for legacy automation. Use a managed
        # identity / federated credential / cert / secret in preference.
        [Parameter(Mandatory = $true, ParameterSetName = 'OAuthCredential')]
        [PSCredential]$Credential,

        # Device code sign-in (RFC 8628). Prints a code + verification URL;
        # the user completes auth (MFA / FIDO2 / YubiKey all work) in a browser
        # on any device while this call polls for the token. Supported by every
        # provider: MSAL uses MSAL.NET's AcquireTokenWithDeviceCode (token lands
        # in the MSAL cache and refreshes silently); MgGraph uses Connect-MgGraph
        # -UseDeviceCode; OAuth speaks the RFC 8628 flow directly. TenantId is
        # optional on MSAL/OAuth ('organizations' / 'common' as appropriate).
        [Parameter(Mandatory = $true, ParameterSetName = 'DeviceCode')]
        [switch]$DeviceCode,

        # Interactive sign-in — browser popup / WAM broker (MSAL) or device code
        # (OAuth). Silent-from-cache first, falls back to interactive prompt
        # when the cache is cold. Default parameter set — you can call
        # `Connect-IntuneManagement` with no args and get an interactive prompt.
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [switch]$Interactive,

        # Interactive-only: pin the account to sign in with (MSAL only). Same
        # semantics as passing $global:MSALLoginHint.
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [string]$User,

        # Interactive-only: bypass the token cache and force a fresh browser
        # prompt even when a cached token exists (MSAL only).
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [switch]$ForceInteractive,

        # Interactive-only: use the Windows broker (WAM) instead of a browser
        # popup. Requires PS7+ on Windows (MSAL only).
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [switch]$AuthenticationBroker,

        # Interactive-only: force the OAuth provider's browser (Authorization Code +
        # PKCE, loopback redirect) flow explicitly, even headless. Without this,
        # -Interactive uses the browser only when a GUI is present, else device code.
        # (OAuth only; ignored by MSAL/MgGraph.)
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [switch]$Browser,

        # New flat Cloud taxonomy (Phase 1 of the cloud redesign, 2026-05-22). Replaces
        # -GraphEnvironment + -GCCType. If omitted, falls back to the "DefaultCloud"
        # setting (defaults to Public). The legacy pair is still accepted for one
        # release with a deprecation warning — see resolution block below.
        [Parameter(Mandatory = $false)]
        [ValidateSet("Public", "USGov", "USGovDOD", "China")]
        [string]$Cloud,

        [Parameter(Mandatory = $false)]
        [ValidateSet("public", "usGov", "china")]
        [string]$GraphEnvironment = "public",

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [ValidateSet("", "gcc", "gccHigh", "gccDoD")]
        [string]$GCCType,

        [switch]$DefaultToken,

        # Pick the auth backend. If omitted, uses the value of the "ActiveAuthProvider"
        # setting (default MSAL). The named provider must be registered (MgGraph requires
        # Microsoft.Graph.Authentication installed). Valid values are enumerated
        # dynamically from every registered AuthenticationProvider — adding a new
        # provider via Register-AuthProvider is sufficient, no edit here needed.
        [Parameter(Mandatory = $false)]
        # Completion (not [ValidateSet([AuthProviderValues])]) so the module still
        # imports on Windows PowerShell 5.1 - the generator implements the PS7-only
        # IValidateSetValuesGenerator. Unknown providers are handled at runtime below.
        [ArgumentCompleter({ param($commandName, $parameterName, $wordToComplete) @(& (Get-Module IntuneManagement) { Get-AuthProviderValues }) | Where-Object { $_ -like "$wordToComplete*" } })]
        [string]$Provider
    )

    # Resolve the provider. If -Provider is explicitly set, use it; otherwise use the
    # currently active provider from AuthenticationCore. Parameter sets that ONLY make
    # sense for the OAuth provider (workload-identity federation, ROPC PSCredential)
    # auto-route to OAuth — saves the caller from having to also pass -Provider OAuth.
    # DeviceCode used to be OAuth-only ('OAuthDeviceCode'); every provider now
    # supports it, so the auto-route no longer needs to force -Provider OAuth
    # for that case. Federated / ROPC (Credential) remain OAuth-exclusive.
    $oauthOnlySets = @('OAuthFederated','OAuthCredential')
    if(-not $Provider -and $PSCmdlet.ParameterSetName -in $oauthOnlySets) {
        $Provider = "OAuth"
        Write-LogDebug "Connect-IntuneManagement auto-selected -Provider OAuth (parameter set: $($PSCmdlet.ParameterSetName))"
    }

    if($Provider) {
        $authProvider = Get-AuthProvider -Id $Provider
        if(-not $authProvider) {
            Write-Log "Provider '$Provider' is not registered. For MgGraph, install Microsoft.Graph.Authentication." 3
            return
        }
    }
    else {
        $authProvider = Get-AuthProvider
        if(-not $authProvider) {
            Write-Log "No authentication provider is active. Cannot continue." 3
            return
        }
    }

    Write-LogDebug "Connect-IntuneManagement routing to provider '$($authProvider.Id)' (parameter set: $($PSCmdlet.ParameterSetName))"

    # Resolve the target cloud once, here, so every downstream branch sees a consistent
    # answer. Precedence: -Cloud wins; explicit -GraphEnvironment/-GCCType is the legacy
    # path (deprecated, warns); otherwise read the DefaultCloud setting. Also derive the
    # legacy GraphEnvironment/GCCType pair from the resolved Cloud so MSAL functions that
    # still take the old form (Phase 4 will migrate them) keep working.
    if($PSBoundParameters.ContainsKey('Cloud')) {
        $resolvedCloud = $Cloud
    }
    elseif($PSBoundParameters.ContainsKey('GraphEnvironment') -or $PSBoundParameters.ContainsKey('GCCType')) {
        Write-Log "-GraphEnvironment and -GCCType are deprecated; use -Cloud (Public/USGov/USGovDOD/China) instead. They will be removed in a future release." 2
        $resolvedCloud = Convert-LegacyToCloud -GraphEnvironment $GraphEnvironment -GCCType $GCCType
    }
    else {
        $resolvedCloud = Get-DefaultCloud
    }
    $cloudEntry = Get-CloudByValue $resolvedCloud
    $GraphEnvironment = $cloudEntry.LegacyEnv
    $GCCType          = if([string]::IsNullOrWhiteSpace($cloudEntry.LegacyGCC)) { $null } else { $cloudEntry.LegacyGCC }
    Write-LogDebug "Connect-IntuneManagement resolved Cloud=$resolvedCloud (legacy GraphEnvironment='$GraphEnvironment', GCCType='$GCCType')"

    # Path A - providers wired into the built-in Connect-* entry points
    # (UsesBuiltInConnectPath, i.e. MSAL) are driven through them directly. Their
    # Connect() ALSO forwards to these functions, but skipping the extra hop keeps
    # stack traces clean and behaviour identical to pre-refactor.
    if($authProvider.UsesBuiltInConnectPath) {
        $sharedParams = @{
            GraphEnvironment = $GraphEnvironment
            GCCType          = $GCCType
            DefaultToken     = $DefaultToken
        }

        switch ($PSCmdlet.ParameterSetName) {
            'Secret' {
                return (Connect-WithClientCredentials -TenantId $TenantId -AppId $AppId -Secret $Secret @sharedParams)
            }
            'Certificate' {
                $cert = Resolve-MSALCertificate $Certificate
                if (-not $cert) {
                    Write-Log "Cannot resolve certificate '$Certificate'. Provide a valid thumbprint or X509Certificate2 object." 3
                    return
                }
                return (Connect-WithClientCredentials -TenantId $TenantId -AppId $AppId -Certificate $cert @sharedParams)
            }
            'CertificatePath' {
                $cert = Resolve-MSALCertificate -CertificatePath $CertificatePath -Password $CertificatePassword
                if (-not $cert) {
                    Write-Log "Cannot load certificate from '$CertificatePath'." 3
                    return
                }
                return (Connect-WithClientCredentials -TenantId $TenantId -AppId $AppId -Certificate $cert @sharedParams)
            }
            'Token' {
                return (Add-BYOTokenInfo -Token $Token -TenantId $TenantId @sharedParams)
            }
            'ManagedIdentity' {
                Write-Log "MSAL provider does not support -ManagedIdentity. Use -Provider MgGraph." 3
                return
            }
            'DeviceCode' {
                # Full MSAL device-code flow via Connect-EntraEnvironment's
                # -DeviceCode switch (uses MSAL.NET's AcquireTokenWithDeviceCode
                # under the hood). Token lands in the MSAL cache so subsequent
                # silent refreshes work identically to interactive sign-in.
                # Splat matches Connect-EntraEnvironment's parameter surface —
                # GraphEnvironment/GCCType are NOT its parameters (used by
                # the client-credentials helpers), so -Cloud carries the cloud.
                $dcArgs = @{
                    DefaultToken = $DefaultToken
                    DeviceCode   = $true
                    Cloud        = $resolvedCloud
                }
                if($TenantId) { $dcArgs['TenantId'] = $TenantId }
                if($AppId)    { $dcArgs['AppId']    = $AppId }
                return (Connect-EntraEnvironment @dcArgs)
            }
            'Interactive' {
                # Interactive MSAL flow — browser popup, or WAM broker when
                # -AuthenticationBroker is set. Delegates to Connect-EntraEnvironment
                # which owns the MSAL public-client PCA plumbing. Splat only the
                # keys Connect-EntraEnvironment declares; -Cloud drives the
                # sovereign-cloud selection there (the legacy Environment param
                # is derived from Cloud downstream). GCCType is not a
                # Connect-EntraEnvironment parameter (it's used by the
                # client-credentials helpers only).
                $iArgs = @{
                    DefaultToken         = $DefaultToken
                    ForceInteractive     = $ForceInteractive
                    AuthenticationBroker = $AuthenticationBroker
                    Cloud                = $resolvedCloud
                }
                if($TenantId) { $iArgs['TenantId'] = $TenantId }
                if($AppId)    { $iArgs['AppId']    = $AppId }
                if($User)     { $iArgs['User']     = $User }
                return (Connect-EntraEnvironment @iArgs)
            }
        }
    }

    # Path B — provider-aware path. Pack the parameters into a hashtable and let the
    # provider class translate. This is the route for MgGraph, OAuth, and any future
    # provider.
    $providerArgs = @{}
    if($TenantId)            { $providerArgs['TenantId']            = $TenantId }
    if($AppId)               { $providerArgs['AppId']               = $AppId }
    if($Secret)              { $providerArgs['Secret']              = $Secret }
    if($Certificate)         { $providerArgs['Certificate']         = $Certificate }
    if($CertificatePath)     { $providerArgs['CertificatePath']     = $CertificatePath }
    if($CertificatePassword) { $providerArgs['CertificatePassword'] = $CertificatePassword }
    if($Token)               { $providerArgs['Token']               = $Token }
    if($ManagedIdentity)     { $providerArgs['ManagedIdentity']     = $true }
    if($FederatedTokenFile)  { $providerArgs['FederatedTokenFile']  = $FederatedTokenFile }
    if($FederatedToken)      { $providerArgs['FederatedToken']      = $FederatedToken }
    if($Credential)          { $providerArgs['Credential']          = $Credential }
    if($DeviceCode)          { $providerArgs['DeviceCode']          = $true }
    if($Interactive -or $PSCmdlet.ParameterSetName -eq 'Interactive') {
        $providerArgs['Interactive'] = $true
    }
    if($User)                { $providerArgs['User']                = $User }
    if($ForceInteractive)    { $providerArgs['ForceInteractive']    = $true }
    if($AuthenticationBroker){ $providerArgs['AuthenticationBroker'] = $true }
    if($Browser)             { $providerArgs['Browser']             = $true }
    if($DeviceCode)          { $providerArgs['DeviceCode']          = $true }
    $providerArgs['Cloud']            = $resolvedCloud
    $providerArgs['GraphEnvironment'] = $GraphEnvironment
    $providerArgs['GCCType']          = $GCCType
    $providerArgs['DefaultToken']     = $DefaultToken.IsPresent

    $result = $authProvider.Connect($providerArgs)

    # MgGraph fallback: if the non-default provider failed (user declined the SDK
    # install, or install failed), fall back to MSAL so the user is still signed in.
    # ManagedIdentity / OAuth-only parameter sets have no MSAL equivalent —
    # don't fall back for them.
    $noFallbackSets = @('ManagedIdentity','OAuthFederated','OAuthCredential','DeviceCode','Interactive')
    if(-not $result -and $authProvider.Id -ne "MSAL" -and $PSCmdlet.ParameterSetName -notin $noFallbackSets) {
        $msal = Get-AuthProvider -Id "MSAL"
        if($msal) {
            Write-Log "Provider '$($authProvider.Id)' failed to connect - falling back to MSAL"
            # Strip provider-specific fields before retrying.
            $providerArgs.Remove('ManagedIdentity')     | Out-Null
            $providerArgs.Remove('FederatedTokenFile')  | Out-Null
            $providerArgs.Remove('FederatedToken')      | Out-Null
            $providerArgs.Remove('Credential')          | Out-Null
            $result = $msal.Connect($providerArgs)
            if($result) {
                # Make MSAL the active provider for the rest of the session — otherwise
                # subsequent Invoke-MSGraphAPI calls would still target the failed
                # provider.
                Set-ActiveAuthProvider -Id "MSAL"
            }
        }
    }

    return $result
}
