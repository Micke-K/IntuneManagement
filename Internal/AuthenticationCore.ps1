# Provider-agnostic authentication core.
#
# This file owns the provider registry, the central token registry, and the
# provider-agnostic facade functions. Each provider self-registers at load: MSAL in
# Invoke-MSALInitialize, OAuth in Invoke-OAuthProviderInitialize, MgGraph in
# Invoke-MgGraphProviderInitialize. The user picks the active provider via the
# ActiveAuthProvider setting.
#
# Consumers are provider-agnostic: Invoke-MSGraphAPI resolves each call's owning
# provider by TokenId (Resolve-AuthTokenProvider) and gets the bearer from it (or lets
# the provider run the request), so MSAL, OAuth and MgGraph tokens can be live at once
# across different tenants.

# Provider registry. Keyed by provider Id ("MSAL", "MgGraph", "Az", ...).
$script:AuthProviders = @{}
$script:ActiveAuthProviderId = $null

# ---- Central token registry (provider-agnostic) ----
# Single source of truth for live tokens across ALL providers. Solves three
# problems that per-provider stores can't: (1) globally-unique token ids (MSAL
# and OAuth otherwise both allocate from 1 and collide); (2) id -> owning
# provider routing so a Graph call reaches the token's own provider, not just
# whichever provider is "active"; (3) one place that fires the auth events with
# one payload shape ([IMAuthToken]).
#
# Records hold only the stable routing essentials; the full IMAuthToken is
# hydrated on read from the owning provider's GetUserInfo(TokenId) so ExpiresOn
# / TenantName never go stale and token refresh needs no registry write.
$script:AuthTokens         = @{}   # [int]TokenId -> @{ TokenId; ProviderId; Cloud; IsDefault }
$script:AuthTokenNextId    = 1     # central monotonic id allocator (feeds every provider)
$script:DefaultAuthTokenId = 0     # 0 = no default

# Add the events first so providers can fire them during their Initialize() call
# without races. Idempotent — Add-AppEvent is safe to call repeatedly.
Add-AppEvent "AuthProviderRegistered"
Add-AppEvent "ActiveAuthProviderChanged"

# ---- Common Entra app + cloud data (provider-neutral) ----
# Shared by every auth provider: MSAL and OAuth resolve their public client from
# Get-EntraApp; MgGraph and the UI cloud pickers read $script:Clouds; everyone
# calls Get-DefaultCloud. NOTE: these arrays MUST stay above the Authentication
# settings registrations below - the EntraApp / DefaultCloud List settings capture
# them as ItemsSource at dot-source time.

$script:EntraApps = @(
    (New-Object PSObject -Property @{Name = ""; ClientId = ""; RedirectUri = ""; Authority = "" }),
    (New-Object PSObject -Property @{Name = "*** Do NOT use *** Microsoft Intune PowerShell"; ClientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"; RedirectUri = "urn:ietf:wg:oauth:2.0:oob"; }),
    (New-Object PSObject -Property @{Name = "Microsoft Graph PowerShell"; ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"; RedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"; })
)

$script:DefaultEntraAppId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

$script:EntraEnvironments = @(
    [PSCustomObject]@{
        Name  = "Entra Public"
        Value = "public"
        URL   = "login.microsoftonline.com"
        GraphURL = "graph.microsoft.com"
    },
    [PSCustomObject]@{
        Name  = "Entra US Government"
        Value = "usGov"
        URL   = "login.microsoftonline.us"
    },
    [PSCustomObject]@{
        Name     = "Entra China"
        Value    = "china"
        URL      = "login.partner.microsoftonline.cn"
        GraphURL = "microsoftgraph.chinacloudapi.cn"
    }
)

# Phase 1 of the cloud redesign (2026-05-22): single Cloud enum replaces the awkward
# GraphEnvironment+GCCType pair. Each entry carries everything a provider needs:
#   AADAuthority  — login host used to build MSAL .WithAuthority("https://<host>/<tenant>")
#   GraphHost     — Graph hostname for absolute @odata.bind / @odata.id URLs
#   ArmHost       — Azure Resource Manager hostname the tenant list is READ from
#   ArmAudiences  — the OAuth resource identifier(s) an ARM token is REQUESTED on, most
#                   likely first. Not the same thing as ArmHost: they happen to be the
#                   same string in the public cloud, but a sovereign tenant may only
#                   know the older management.core.* identifier, so both are listed and
#                   Internal/EntraTenantList.ps1 falls back in order.
#   MgEnvironment — Connect-MgGraph -Environment value (Global / USGov / USGovDOD / China)
#   LegacyEnv     — old GraphEnvironment value (public / usGov / china) — used by the
#   LegacyGCC     — old GCCType value (gccHigh / gccDoD / $null) — deprecation translator
# "GCC" (commercial Graph from a Gov tenant) collapses into Public: same endpoints,
# no separate cloud entry needed.
$script:Clouds = @(
    [PSCustomObject]@{
        Value         = "Public"
        Name          = "Public (Global)"
        AADAuthority  = "login.microsoftonline.com"
        GraphHost     = "graph.microsoft.com"
        ArmHost       = "management.azure.com"
        ArmAudiences  = @("https://management.azure.com", "https://management.core.windows.net")
        MgEnvironment = "Global"
        LegacyEnv     = "public"
        LegacyGCC     = $null
    },
    [PSCustomObject]@{
        Value         = "USGov"
        Name          = "US Government (GCC High)"
        AADAuthority  = "login.microsoftonline.us"
        GraphHost     = "graph.microsoft.us"
        ArmHost       = "management.usgovcloudapi.net"
        ArmAudiences  = @("https://management.usgovcloudapi.net", "https://management.core.usgovcloudapi.net")
        MgEnvironment = "USGov"
        LegacyEnv     = "usGov"
        LegacyGCC     = "gccHigh"
    },
    [PSCustomObject]@{
        Value         = "USGovDOD"
        Name          = "US Government (DoD)"
        AADAuthority  = "login.microsoftonline.us"
        GraphHost     = "dod-graph.microsoft.us"
        ArmHost       = "management.usgovcloudapi.net"
        ArmAudiences  = @("https://management.usgovcloudapi.net", "https://management.core.usgovcloudapi.net")
        MgEnvironment = "USGovDOD"
        LegacyEnv     = "usGov"
        LegacyGCC     = "gccDoD"
    },
    [PSCustomObject]@{
        Value         = "China"
        Name          = "China (Vianet)"
        AADAuthority  = "login.partner.microsoftonline.cn"
        GraphHost     = "microsoftgraph.chinacloudapi.cn"
        ArmHost       = "management.chinacloudapi.cn"
        ArmAudiences  = @("https://management.chinacloudapi.cn", "https://management.core.chinacloudapi.cn")
        MgEnvironment = "China"
        LegacyEnv     = "china"
        LegacyGCC     = $null
    }
)

function Get-CloudByValue
{
    param([string]$Value)
    if(-not $Value) { $Value = "Public" }
    $entry = $script:Clouds | Where-Object Value -eq $Value | Select-Object -First 1
    if(-not $entry) {
        Write-Log "Unknown cloud '$Value' - falling back to Public" 2
        $entry = $script:Clouds | Where-Object Value -eq "Public" | Select-Object -First 1
    }
    return $entry
}

function Convert-LegacyToCloud
{
    # Translate the historical (GraphEnvironment, GCCType) pair to a new Cloud value.
    # Used during the deprecation window so callers passing -GraphEnvironment/-GCCType
    # still land on the correct sovereign cloud after we route through the flat enum.
    param(
        [string]$GraphEnvironment,
        [string]$GCCType
    )
    if(-not $GraphEnvironment) { return "Public" }
    foreach($c in $script:Clouds) {
        $cGcc = if($c.LegacyGCC) { $c.LegacyGCC } else { "" }
        $inGcc = if($GCCType)    { $GCCType }    else { "" }
        if($c.LegacyEnv -eq $GraphEnvironment -and $cGcc -eq $inGcc) {
            return $c.Value
        }
    }
    # usGov without gccHigh/gccDoD historically meant commercial Graph from a Gov
    # tenant (the legacy "gcc" value). Collapse to Public — same endpoints.
    if($GraphEnvironment -eq "usGov") { return "Public" }
    return "Public"
}

function Get-DefaultCloud
{
    # Persisted default cloud. Falls back to Public if the setting is missing
    # or empty (corrupted config).
    $val = Get-SettingValue "DefaultCloud" "Public"
    if(-not $val) { return "Public" }
    return $val
}

function Set-DefaultCloud
{
    # Single write path for the DefaultCloud setting. UI code must use this instead
    # of Save-SettingStoreValue with a hardcoded path, so reader (Get-DefaultCloud via the
    # setting's SubPath) and writers always agree on Authentication\DefaultCloud.
    param([Parameter(Mandatory)][string]$Value)
    Save-SettingStoreValue "Authentication" "DefaultCloud" $Value
}

# --- Per-tenant cloud memory ---

function Resolve-CloudFromIss
{
    # Best-effort Cloud detection from a JWT 'iss' (issuer) claim. We can't tell
    # USGov (GCC High) from USGovDOD purely from iss — both issue from
    # login.microsoftonline.us / sts.windows.us. When the caller already knows which
    # USGov variant was requested (via -FallbackCloud), that wins; otherwise we
    # default to USGov (the more common GCC variant). The same logic applies to any
    # other ambiguous family.
    param(
        [string]$Iss,
        [string]$FallbackCloud
    )
    if(-not $Iss) { return $FallbackCloud }
    $low = $Iss.ToLowerInvariant()
    if($low -match "microsoftonline\.us|sts\.windows\.us") {
        if($FallbackCloud -eq "USGovDOD") { return "USGovDOD" }
        return "USGov"
    }
    if($low -match "partner\.microsoftonline\.cn|chinacloudapi\.cn") {
        return "China"
    }
    if($low -match "microsoftonline\.com|sts\.windows\.net|login\.microsoft\.com") {
        return "Public"
    }
    return $FallbackCloud
}

function Save-TenantCloud
{
    # Persist the cloud a given tenant was authenticated against. Used at silent
    # refresh + next-launch resume so we hit the right authority without re-asking.
    param([string]$TenantId, [string]$Cloud)
    if(-not $TenantId -or -not $Cloud) { return }
    Save-SettingStoreValue "" "TenantCloud_$TenantId" $Cloud
}

function Get-TenantCloud
{
    param([string]$TenantId)
    if(-not $TenantId) { return $null }
    $val = Get-SettingStoreValue "" "TenantCloud_$TenantId" ""
    if(-not $val) { return $null }
    return [string]$val
}

function Get-StartupCloudHint
{
    # Resolution chain for a sign-in that doesn't carry an explicit cloud:
    #   1. TenantCloud_<TenantId> if the caller knows the target tenant
    #   2. LastLoggedOnCloud (the cloud of the most recent default sign-in)
    #   3. DefaultCloud setting (user-configured default)
    # Returns a Cloud value (Public/USGov/USGovDOD/China) — never $null.
    param([string]$TenantId)
    if($TenantId) {
        $t = Get-TenantCloud $TenantId
        if($t) { return $t }
    }
    $last = Get-SettingStoreValue "" "LastLoggedOnCloud" ""
    if($last) { return [string]$last }
    return (Get-DefaultCloud)
}

function New-EntraApp {
    param($ClientId, $TenantId, $RedirectUri, $Authority)

    return New-Object PSObject -Property @{
        ClientId = $ClientId
        TenantId = $TenantId
        RedirectUri = $RedirectUri
        Authority = $Authority
    }
}

function Get-EntraApp {
    [CmdletBinding()]
    param($AppId,
            $RedirectURI,
            $Authority)

    # Resolution order, highest priority first:
    #   1. Caller-supplied $AppId (explicit override).
    #   2. "EntraApp" setting (the Settings → Authentication → Entra application
    #      dropdown — picks one of the built-in apps in $script:EntraApps).
    #   3. "EntraCustomAppId" setting (the Settings → Authentication → Application Id
    #      field — a custom Entra App registration the user controls).
    #   4. $script:DefaultEntraAppId (Microsoft Graph PowerShell first-party app).
    # The previous version defaulted $entraAppId to the Microsoft Graph PowerShell
    # GUID when the dropdown was empty, which then matched a built-in entry and
    # never fell through to EntraCustomAppId — so a user who filled in only the
    # custom Application Id field still got logged in via the first-party app.
    $entraAppObj = $null

    if(-not [String]::IsNullOrWhiteSpace($AppId)) {
        $entraAppObj = $script:EntraApps | Where-Object ClientId -eq $AppId
        if(-not $entraAppObj) {
            $entraAppObj = New-EntraApp -ClientId $AppId `
                                        -TenantId (Get-SettingValue "EntraCustomTenantId") `
                                        -RedirectUri (Get-SettingValue "EntraCustomAppRedirect" $RedirectURI) `
                                        -Authority (Get-SettingValue "EntraCustomAuthority" $Authority)
        }
    }

    if(-not $entraAppObj) {
        $dropdownAppId = Get-SettingValue "EntraApp"
        if(-not [String]::IsNullOrWhiteSpace($dropdownAppId)) {
            $entraAppObj = $script:EntraApps | Where-Object ClientId -eq $dropdownAppId
        }
    }

    if(-not $entraAppObj) {
        $customAppId = Get-SettingValue "EntraCustomAppId"
        if(-not [String]::IsNullOrWhiteSpace($customAppId)) {
            $entraAppObj = New-EntraApp -ClientId $customAppId `
                                        -TenantId (Get-SettingValue "EntraCustomTenantId") `
                                        -RedirectUri (Get-SettingValue "EntraCustomAppRedirect" $RedirectURI) `
                                        -Authority (Get-SettingValue "EntraCustomAuthority" $Authority)
        }
    }

    if((-not $entraAppObj -or -not $entraAppObj.ClientId) -and $script:DefaultEntraAppId) {
        $entraAppObj = $script:EntraApps | Where-Object ClientId -eq $script:DefaultEntraAppId
    }

    $entraAppObj
}

function Get-EntraEnvironment
{
    [CmdletBinding()]
    param($Environment = $null)

    if([String]::IsNullOrEmpty($Environment)) { $Environment = "public"}

    $msalEnvironment = $script:EntraEnvironments | Where-Object Value -eq $Environment
    if(-not $msalEnvironment) {
        Write-Log "Entra environment $Environment not found. Using pulic" 2
    }
    Write-LogDebug "Use Entra environment $($msalEnvironment.Value)"
    $msalEnvironment
}

# ---- end common Entra app + cloud data ----

# Setting that picks which registered provider is active. Read once at AppInitialized
# time (after every provider has had a chance to register). If the setting points at
# a provider that isn't registered (e.g. one this build doesn't ship) we keep
# whatever was set during registration — usually MSAL since it always registers.
#
# ItemsSource starts empty: this file loads before any provider registers, so the
# real list is built from the registry in Invoke-AuthCoreOnAppInitialized. Both
# settings dialogs read ItemsSource at render time, long after AppInitialized.
Add-SettingsSection -Title "Authentication" -Id "Authentication" -Order 7
Add-SettingsObject -Title "Active authentication provider" -Key "ActiveAuthProvider" -Type "List" -DefaultValue "MSAL" `
    -ItemsSource @() `
    -Description "Which authentication backend to use. MSAL is the built-in default. MgGraph requires the Microsoft.Graph.Authentication PowerShell module to be installed. OAuth is a pure-PowerShell provider for automation (CI / scheduled tasks / managed identity / workload identity federation) - no SDK required." `
    -Section "Authentication"

# Common app-identity + cloud settings, shared by every provider (moved here from
# the former Entra section). MSAL and OAuth resolve their public client from these
# via Get-EntraApp - they are two ways of authenticating with the SAME app.
Add-SettingsObject -Title "Entra application" -Key "EntraApp" -Type "List" -SelectedValuePath "ClientId" -ItemsSource $script:EntraApps `
    -Description "Built-in Entra application used for interactive sign-in by both the MSAL and OAuth providers. Leave empty to use the custom Application Id below, or the default Microsoft Graph PowerShell app." `
    -Section "Authentication"

Add-SettingsObject -Title "Application Id" -Key "EntraCustomAppId" -Type "String" `
    -Description "Custom Entra application (client) id used by MSAL and OAuth sign-in when no built-in application is selected above. The app registration needs the http://localhost redirect URI for browser-based login." `
    -Section "Authentication"

Add-SettingsObject -Title "Redirect URL" -Key "EntraCustomAppRedirect" -Type "String" `
    -Section "Authentication"

Add-SettingsObject -Title "Tenant Id" -Key "EntraCustomTenantId" -Type "String" `
    -Section "Authentication"

Add-SettingsObject -Title "Authority" -Key "EntraCustomAuthority" -Type "String" `
    -Section "Authentication"

Add-SettingsObject -Title "Default cloud" -Key "DefaultCloud" -Type "List" -DefaultValue "Public" `
    -SelectedValuePath "Value" -ItemsSource $script:Clouds `
    -Description "Microsoft cloud this tool signs in to by default. Public covers commercial + GCC commercial; USGov is GCC High; USGovDOD is GCC DoD; China is the Vianet cloud." `
    -Section "Authentication"

Add-SettingsObject -Title "Interactive login timeout (seconds)" -Key "MSGraphInteractiveTimeoutSec" -Type "Int" -DefaultValue 600 `
    -Description "Maximum time to wait for an interactive login to complete (MSAL embedded/broker and OAuth browser flows). Only applies to flows a user is waiting in front of; silent and app-secret token requests have their own much shorter cap. The default was 180 s, which killed legitimate sign-ins that needed an account picker plus an MFA approval - and the sign-in status has a Cancel button, so an abandoned window does not depend on this timeout." `
    -Section "Authentication"

function Register-AuthProvider {
    <#
    .SYNOPSIS
        Register an authentication provider in the multi-provider auth core.
    .DESCRIPTION
        The provider must inherit from AuthenticationProvider and set its Id. The
        first provider registered becomes the active one automatically; pass
        -SetActive to force-switch.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AuthenticationProvider]$Provider,

        [switch]$SetActive
    )

    if(-not $Provider.Id) {
        Write-Log "Refusing to register an AuthenticationProvider without an Id" 3
        return
    }

    if($script:AuthProviders.ContainsKey($Provider.Id)) {
        Write-Log "AuthenticationProvider '$($Provider.Id)' is already registered; replacing" 2
    }

    $script:AuthProviders[$Provider.Id] = $Provider
    Write-Log "Registered authentication provider: $($Provider.Id) - $($Provider.DisplayName)"

    try { $Provider.Initialize() }
    catch { Write-LogError "Provider $($Provider.Id) Initialize() threw" $_.Exception }

    Invoke-AppEvent "AuthProviderRegistered" $Provider | Out-Null

    if($SetActive -or -not $script:ActiveAuthProviderId) {
        Set-ActiveAuthProvider -Id $Provider.Id
    }
}

function Get-AuthProvider {
    <#
    .SYNOPSIS
        Get a registered authentication provider by Id, or the active one if -Id is
        omitted.
    #>
    [CmdletBinding()]
    [OutputType([AuthenticationProvider])]
    param([string]$Id)

    if(-not $Id) { $Id = $script:ActiveAuthProviderId }
    if(-not $Id) { return $null }
    if(-not $script:AuthProviders.ContainsKey($Id)) { return $null }
    return $script:AuthProviders[$Id]
}

function Get-RegisteredAuthProviders {
    <#
    .SYNOPSIS
        List every registered authentication provider.
    #>
    return @($script:AuthProviders.Values)
}

# Accessor backing the -Provider ArgumentCompleter on Connect-IntuneManagement /
# Get-AuthToken (invoked via & (Get-Module IntuneManagement) { Get-AuthProviderValues }).
# A [ValidateSet([IValidateSetValuesGenerator])] was deliberately avoided so the
# module still imports on Windows PowerShell 5.1 (that interface is PS7-only).
# Returns the sorted list of registered provider Ids (MSAL, MgGraph, OAuth, ...).
function Get-AuthProviderValues {
    if (-not $script:AuthProviders -or $script:AuthProviders.Count -eq 0) { return @() }
    return @($script:AuthProviders.Keys | Sort-Object)
}

# ===================== Central token registry =====================
# See the $script:AuthTokens comment block at the top of this file.

# The single global token-id allocator. Every provider draws its token id from
# here instead of its own counter, so ids are unique across MSAL / OAuth / MgGraph.
function Get-NextAuthTokenId {
    $id = $script:AuthTokenNextId
    $script:AuthTokenNextId = $id + 1
    return $id
}

# Hydrate an [IMAuthToken] for a registry record: ask the owning provider for the
# live identity, then stamp the registry-authoritative fields (TokenId, Provider,
# Cloud, IsDefault). Tolerates a provider that returns $null (e.g. torn-down
# session) by falling back to the stored record fields — never triggers auth.
function ConvertTo-IMAuthToken {
    param([Parameter(Mandatory)][hashtable]$Record)

    $provider = $null
    if ($Record.ProviderId -and $script:AuthProviders.ContainsKey($Record.ProviderId)) {
        $provider = $script:AuthProviders[$Record.ProviderId]
    }

    $info = $null
    if ($provider) {
        try { $info = $provider.GetUserInfo($Record.TokenId) }
        catch { Write-LogDebug "ConvertTo-IMAuthToken: GetUserInfo failed on '$($Record.ProviderId)' for token $($Record.TokenId): $($_.Exception.Message)" }
    }

    $token = if ($info -is [IMAuthToken]) { $info } else { [IMAuthToken]::new() }
    if ($info -and $info -isnot [IMAuthToken]) {
        # Provider still returns a loose PSCustomObject (pre-migration) - copy the
        # known fields across so callers always get a typed IMAuthToken.
        foreach ($p in 'TenantId','TenantName','Account','UPN','UserId','AppId','AppName','AuthType','ExpiresOn') {
            if ($info.PSObject.Properties[$p]) { $token.$p = $info.$p }
        }
        if (-not $token.Account -and $info.PSObject.Properties['DisplayName']) { $token.Account = $info.DisplayName }
    }

    # Registry-authoritative overlay.
    $token.TokenId   = [int]$Record.TokenId
    $token.Provider  = [string]$Record.ProviderId
    if ($Record.Cloud) { $token.Cloud = [string]$Record.Cloud }
    $token.IsDefault = ($script:DefaultAuthTokenId -eq $Record.TokenId)
    return $token
}

# Return the [IMAuthToken] for one id, or $null when unknown.
function Get-AuthTokenById {
    param([int]$TokenId)
    if (-not $script:AuthTokens.ContainsKey($TokenId)) { return $null }
    return ConvertTo-IMAuthToken -Record $script:AuthTokens[$TokenId]
}

# Return every live token as [IMAuthToken[]], ordered by id.
function Get-AuthTokenList {
    if (-not $script:AuthTokens -or $script:AuthTokens.Count -eq 0) { return [IMAuthToken[]]@() }
    $rows = foreach ($id in ($script:AuthTokens.Keys | Sort-Object)) {
        ConvertTo-IMAuthToken -Record $script:AuthTokens[$id]
    }
    return [IMAuthToken[]]@($rows)
}

function Get-DefaultAuthTokenId {
    return $script:DefaultAuthTokenId
}

# Make a token the default. IsDefault everywhere derives from this single value,
# so there is no per-record flag to keep in sync.
function Set-DefaultAuthToken {
    param([Parameter(Mandatory)][int]$TokenId)
    if (-not $script:AuthTokens.ContainsKey($TokenId)) {
        Write-Log "Set-DefaultAuthToken: token $TokenId is not registered" 2
        return
    }
    $script:DefaultAuthTokenId = $TokenId
}

# Resolve a token id to its owning [AuthenticationProvider]; $null when unknown
# (callers fall back to the active provider for id 0 / unregistered ids).
function Resolve-AuthTokenProvider {
    param([int]$TokenId)
    if ($TokenId -le 0) { return $null }
    if (-not $script:AuthTokens.ContainsKey($TokenId)) { return $null }
    $providerId = $script:AuthTokens[$TokenId].ProviderId
    if ($providerId -and $script:AuthProviders.ContainsKey($providerId)) {
        return $script:AuthProviders[$providerId]
    }
    return $null
}

# Register a freshly-acquired token. Providers call this after storing the token
# in their own backend, passing the id they drew from Get-NextAuthTokenId. Sets
# the first token (or an explicitly-defaulted one) as default and fires
# AuthenticatedNewToken with the hydrated [IMAuthToken]. Returns that token.
function Register-AuthToken {
    # -Provider is intentionally untyped: only $Provider.Id is used here (hydration
    # later resolves the live provider from $script:AuthProviders by that id). Keeping
    # it untyped matches how the codebase injects fake providers in tests.
    param(
        [Parameter(Mandatory)]$Provider,
        [Parameter(Mandatory)][int]$TokenId,
        [string]$Cloud,
        [switch]$Default
    )

    $script:AuthTokens[$TokenId] = @{
        TokenId    = $TokenId
        ProviderId = $Provider.Id
        Cloud      = $Cloud
    }

    if ($Default -or $script:DefaultAuthTokenId -le 0) {
        $script:DefaultAuthTokenId = $TokenId
    }

    $token = ConvertTo-IMAuthToken -Record $script:AuthTokens[$TokenId]
    try { Invoke-AppEvent "AuthenticatedNewToken" $token | Out-Null } catch { }
    return $token
}

# Remove a token from the registry. Fires AuthenticationUserDisconnected with a
# final snapshot; if the removed token was the default and others remain, promotes
# the highest-id survivor and fires AuthenticatedNewToken for it.
function Unregister-AuthToken {
    param([Parameter(Mandatory)][int]$TokenId)
    if (-not $script:AuthTokens.ContainsKey($TokenId)) { return }

    $snapshot = ConvertTo-IMAuthToken -Record $script:AuthTokens[$TokenId]
    $wasDefault = ($script:DefaultAuthTokenId -eq $TokenId)
    $script:AuthTokens.Remove($TokenId) | Out-Null

    $promoted = $null
    if ($wasDefault) {
        $script:DefaultAuthTokenId = 0
        if ($script:AuthTokens.Count -gt 0) {
            $nextId = ($script:AuthTokens.Keys | Sort-Object -Descending | Select-Object -First 1)
            $script:DefaultAuthTokenId = $nextId
            $promoted = ConvertTo-IMAuthToken -Record $script:AuthTokens[$nextId]
        }
    }

    try { Invoke-AppEvent "AuthenticationUserDisconnected" $snapshot | Out-Null } catch { }
    if ($promoted) {
        try { Invoke-AppEvent "AuthenticatedNewToken" $promoted | Out-Null } catch { }
    }
}

# Fire AuthenticationTokenRefresh for a token after a silent renewal. No registry
# write needed (IMAuthToken is hydrated on read); this just notifies consumers.
function Update-AuthToken {
    param([Parameter(Mandatory)][int]$TokenId)
    if (-not $script:AuthTokens.ContainsKey($TokenId)) { return }
    $token = ConvertTo-IMAuthToken -Record $script:AuthTokens[$TokenId]
    try { Invoke-AppEvent "AuthenticationTokenRefresh" $token | Out-Null } catch { }
}

# Fire AuthenticationFailed with a canonical shape, so every provider reports
# failures the same way (replaces ad-hoc Exception / $null / raw-object payloads).
function Invoke-AuthTokenFailed {
    param(
        [string]$Provider,
        [string]$TenantId,
        [string]$ErrorCode,
        [string]$Message,
        $Exception
    )
    $payload = [PSCustomObject]@{
        Provider  = $Provider
        TenantId  = $TenantId
        ErrorCode = $ErrorCode
        Message   = if ($Message) { $Message } elseif ($Exception) { $Exception.Message } else { $null }
        Exception = $Exception
    }
    try { Invoke-AppEvent "AuthenticationFailed" $payload | Out-Null } catch { }
}

# True when the current default (signed-in) token has passed its expiry with no
# valid replacement. Provider-agnostic: uses the owning provider's
# GetAccessTokenExpiry(), which returns [datetime]::MaxValue when expiry is
# unknown or the SDK manages refresh (MgGraph) - those are treated as NOT expired
# so we never falsely flip a still-managed session to a signed-out UI. Returns
# $false when nothing is signed in. Used by the UI to decide when the title-bar
# profile control should fall back to the Sign-in icon.
function Test-DefaultTokenExpired {
    try {
        if (-not (Get-Command Get-AuthProvider -ErrorAction SilentlyContinue)) { return $false }
        $provider = Get-AuthProvider
        $tokenId  = Get-DefaultTokenId
        if (-not $provider -or -not $tokenId) { return $false }
        $exp = $provider.GetAccessTokenExpiry($tokenId, "https://graph.microsoft.com")
        return ($exp -ne [datetime]::MaxValue -and $exp -le (Get-Date))
    }
    catch { return $false }
}

# ===================== end token registry =====================

function Set-ActiveAuthProvider {
    <#
    .SYNOPSIS
        Make a registered provider the active one. Subsequent calls to Get-AuthProvider
        (without -Id) return this provider.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Id)

    if(-not $script:AuthProviders.ContainsKey($Id)) {
        Write-Log "Cannot set active auth provider '$Id' - not registered" 3
        return
    }
    $previous = $script:ActiveAuthProviderId
    $script:ActiveAuthProviderId = $Id
    Write-Log "Active authentication provider: $Id"

    if($previous -ne $Id) {
        Invoke-AppEvent "ActiveAuthProviderChanged" $script:AuthProviders[$Id] $previous | Out-Null
    }
}

# === Facades that route to the ACTIVE provider ===
# Convenience wrappers for callers that just want "the current provider". The hot path
# (Invoke-MSGraphAPI) does NOT use these: it resolves each call's OWNING provider by
# TokenId (Resolve-AuthTokenProvider) so a call reaches the token's own provider, not
# merely whichever provider is active.

function Get-AuthProviderAccessToken {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [int]$TokenId,
        [string]$Resource = "https://graph.microsoft.com",
        [string]$ProviderId
    )
    $provider = Get-AuthProvider $ProviderId
    if(-not $provider) { return $null }
    return $provider.GetAccessToken($TokenId, $Resource)
}

function Get-AuthProviderUserInfo {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param([int]$TokenId, [string]$ProviderId)

    $provider = Get-AuthProvider $ProviderId
    if(-not $provider) { return $null }
    return $provider.GetUserInfo($TokenId)
}

function Get-AuthProviderCachedAccounts {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param([string]$ProviderId)

    $provider = Get-AuthProvider $ProviderId
    if(-not $provider) { return [PSCustomObject[]]@() }
    return $provider.GetCachedAccounts()
}

# --- Active-tenant / current-user accessors (architecture R7) ---
# The org/tenant identifiers and the signed-in user are auth-module state
# (set by the authentication providers on every token change). Non-auth code
# must read them through these accessors instead of reaching into the module's
# internal $script: variables directly. The values are module-scope, so these
# are thin getters; the point is encapsulation - if the backing field is ever
# renamed or sourced differently, only these change.

function Get-CurrentTenantId {
    [OutputType([string])]
    param()
    return $script:OrganizationId
}

function Get-CurrentOrganizationName {
    [OutputType([string])]
    param()
    return $script:OrganizationName
}

function Get-CurrentUser {
    # The signed-in user object (Graph /me shape when available; may be a
    # display-name string for app/secret logins, or $null when signed out).
    param()
    return $script:CurrentUser
}

function Get-AuthProviderAvailableTenants {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param([int]$TokenId, [string]$ProviderId)

    $provider = Get-AuthProvider $ProviderId
    if(-not $provider) { return [PSCustomObject[]]@() }
    return $provider.GetAvailableTenants($TokenId)
}

# Honour the user's "ActiveAuthProvider" setting after every provider has had a
# chance to register. AppInitialized fires near the end of module load — by then
# both MSAL and (if installed) MgGraph have registered themselves.
function Invoke-AuthCoreOnAppInitialized {
    # Populate the ActiveAuthProvider setting's dropdown from the registry now that
    # every provider has registered. MSAL first (it's the default), rest by Id.
    $settingObj = (Get-SettingsSection "Authentication").Values | Where-Object Key -eq "ActiveAuthProvider"
    if($settingObj) {
        $settingObj.ItemsSource = @(Get-RegisteredAuthProviders |
            Sort-Object @{ Expression = { $_.Id -ne "MSAL" } }, Id |
            ForEach-Object { [PSCustomObject]@{ Name = $_.DisplayName; Value = $_.Id } })
    }

    # IM_AUTH_PROVIDER picks the provider for this session only (launch configs,
    # "start with OAuth" vs "start with MSAL"); it outranks the persisted setting
    # but never writes it, so the next plain launch is back on the saved choice.
    $desired = $env:IM_AUTH_PROVIDER
    $source = "environment (IM_AUTH_PROVIDER)"
    if(-not $desired) {
        $desired = Get-SettingValue "ActiveAuthProvider"
        $source = "setting"
    }
    if($desired) {
        if($script:AuthProviders.ContainsKey($desired)) {
            if($script:ActiveAuthProviderId -ne $desired) {
                Write-Log "Active auth provider switched to '$desired' per $source"
                Set-ActiveAuthProvider -Id $desired
            }
        }
        else {
            Write-Log "$source requested auth provider '$desired' but it isn't registered (active = '$script:ActiveAuthProviderId')" 2
        }
    }

    # Give the active provider a chance to silently resume a persisted session at
    # startup (no interactive prompt). MSAL overrides TryResumeSession to do a silent
    # Connect-EntraEnvironment from its on-disk cache; MgGraph resumes its
    # Azure.Identity disk cache. Providers with nothing to resume inherit the base
    # no-op ($false).
    $active = Get-AuthProvider
    if($active) {
        try {
            if($active.TryResumeSession()) {
                Write-LogDebug "Provider '$($active.Id)' resumed a persisted session at startup"
            }
        }
        catch {
            Write-LogError "Provider '$($active.Id)' TryResumeSession threw" $_.Exception
        }
    }
}

Add-AppEventHandler "AppInitialized" "Invoke-AuthCoreOnAppInitialized"

# Provider-neutral post-auth state sync. Reads the active provider's normalized
# GetUserInfo() shape (defined on the AuthenticationProvider base class) and
# writes the module-scope globals ($script:OrganizationId, $script:OrganizationName,
# $script:CurrentUser) that downstream code -- Get-CurrentTenantId,
# Test-DocumentationGraphAvailable, env-badge helpers, etc. -- reads. Replaces
# the MSAL-only path in Get-MSALUserInfo (now Update-MSALUserProfile) for these
# generic fields, so headless OAuth / MgGraph sessions (no UI event handlers)
# still populate the tenant id.
function Sync-AuthContextFromProvider {
    [CmdletBinding()]
    param([int]$TokenId = 0)

    $provider = Get-AuthProvider
    if(-not $provider) {
        $script:OrganizationId   = $null
        $script:OrganizationName = $null
        $script:CurrentUser      = $null
        return
    }

    $u = $null
    try { $u = $provider.GetUserInfo($TokenId) }
    catch { Write-LogDebug "Sync-AuthContextFromProvider: GetUserInfo failed on '$($provider.Id)': $($_.Exception.Message)" }
    if(-not $u) { return }

    if($u.TenantId)   { $script:OrganizationId   = [string]$u.TenantId }
    if($u.TenantId)   { $script:OrganizationName = if($u.TenantName) { [string]$u.TenantName } else { [string]$u.TenantId } }

    # Minimal, provider-independent user shape. UI code that wants MSAL-specific
    # enrichment (photo, JWT claim inspection) layers it on separately via
    # Update-MSALUserProfile.
    $script:CurrentUser = [PSCustomObject]@{
        displayName       = $u.DisplayName
        userPrincipalName = $u.UPN
        Id                = $u.UserId
    }
}

# Module-scope AuthenticatedNewToken handler. Fires regardless of whether the UI
# is loaded, so headless sessions get the same essential state updates the UI
# handlers used to do exclusively (organization id/name, scope-tag + filter
# preload into the persistent dependency cache).
function Invoke-AuthCoreOnNewToken {
    [CmdletBinding()]
    param($TokenInfo)
    if(-not $TokenInfo) { return }

    # Payload is an [IMAuthToken] (.TokenId) once a provider registers via the
    # token registry; legacy MSAL fire sites still pass a Get-TokenInfo projection
    # (.Id) until they migrate. Accept either so the handler is correct throughout
    # the rollout.
    $tokenId = 0
    if($TokenInfo.PSObject.Properties['TokenId'] -and $TokenInfo.TokenId) { $tokenId = [int]$TokenInfo.TokenId }
    elseif($TokenInfo.PSObject.Properties['Id'] -and $TokenInfo.Id) { $tokenId = [int]$TokenInfo.Id }

    try { Sync-AuthContextFromProvider -TokenId $tokenId }
    catch { Write-LogError 'Sync-AuthContextFromProvider failed on AuthenticatedNewToken' $_.Exception }

    # Preload ScopeTags + AssignmentFilters into DependencyObjects_<TenantId> so
    # downstream flows (Documentation NameFilter 'scope:...', Copy / Import
    # scope-tag pickers) resolve without a mid-run Graph call. Was previously
    # only invoked from the UI event handlers -- headless flows never got it.
    if($tokenId -gt 0) {
        try { Initialize-TenantDependencyCache -TokenId $tokenId }
        catch { Write-LogError 'Initialize-TenantDependencyCache failed on AuthenticatedNewToken' $_.Exception }
    }
}

# Declare the event name here in addition to the declaration in
# AuthenticationMSALHelpers.ps1. Internal/ files are dot-sourced alphabetically, so
# AuthenticationCore.ps1 loads BEFORE AuthenticationMSALHelpers.ps1 — without this
# call, Add-AppEventHandler below silently fails because the event doesn't yet
# exist in $script:AppEventTriggers. Add-AppEvent is idempotent (checks
# ContainsKey), so re-declaring in Invoke-MSALInitialize is harmless.
# Add Authentication Events
Add-AppEvent "AuthenticatedNewToken"
Add-AppEvent "AuthenticationTokenRefresh"
Add-AppEvent "AuthenticationUserDisconnected"
Add-AppEvent "AuthenticationFailed"

Add-AppEventHandler "AuthenticatedNewToken" "Invoke-AuthCoreOnNewToken"
