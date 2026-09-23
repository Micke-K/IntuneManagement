#ImportOrder 25

# Base class for authentication providers.
#
# Naming convention: concrete providers are named Authentication<Backend>
# (AuthenticationMSAL, AuthenticationMgGraph, AuthenticationAz). The base class
# stays AuthenticationProvider — there's only one of those.
#
# Three concrete providers ship: AuthenticationMSAL (MSAL.NET), AuthenticationMgGraph
# (Microsoft.Graph SDK) and AuthenticationOAuth (pure PowerShell for automation).
#
# Concrete providers override what they support. Methods that MUST be overridden
# throw NotImplemented when called on the base; optional methods return safe defaults
# so a provider that doesn't support a feature simply reports back $false / $null.
#
# Return shapes for the *Info / *Accounts / *Tenants methods are uniform across all
# providers — that's the whole point of the abstraction. Consumers (Invoke-MSGraphAPI,
# profile dialog, account picker) read these uniform shapes directly. Invoke-MSGraphAPI
# resolves a call's OWNING provider by TokenId and gets the bearer from it (or lets the
# provider run the request), so MSAL / MgGraph / OAuth tokens can all be live at once.
#
# === Required capabilities (contract) ===
# Every concrete provider MUST support and prove out these auth modes:
#   * Interactive login            (SupportsInteractive   = $true)
#   * App with client secret       (SupportsClientSecret  = $true)
#   * App with certificate         (SupportsCertificate   = $true)
# Setting any of these to $false on a concrete provider is a contract violation —
# the abstraction assumes consumers can pick any of the three.
#
# === Recommended capabilities ===
# Providers SHOULD also support, where the backend allows it:
#   * Identity Provider login      (SupportsIdentityProvider = $true)
#     Covers managed identity, workload identity federation, and federated
#     credential assertions. Not required because Connect-IntuneManagement
#     doesn't yet expose those parameter sets, but expected for full coverage.
class AuthenticationProvider {
    # Identity
    [string]$Id
    [string]$DisplayName

    # Capability flags. The profile UI reads these to enable / disable buttons and the
    # connect facade routes only to providers that support the requested mode.
    [bool]$SupportsMultiTenant      = $true
    [bool]$SupportsRefresh          = $true
    [bool]$SupportsForget           = $true
    [bool]$SupportsCachedUsers      = $true

    # Required by contract. See block comment above.
    [bool]$SupportsInteractive      = $true
    [bool]$SupportsClientSecret     = $true
    [bool]$SupportsCertificate      = $true

    # Recommended. Managed identity / workload identity federation / federated
    # credential. Off by default — concrete providers opt in when implemented.
    [bool]$SupportsIdentityProvider = $false

    # Optional. Provider accepts an externally-issued bearer token (no auth flow).
    [bool]$SupportsBYOToken         = $false

    # Optional. Provider can satisfy a CAE (Continuous Access Evaluation) claims
    # challenge - a Graph 401 carrying a WWW-Authenticate claims="..." parameter - by
    # re-minting the token via GetClaimsToken(). Providers whose SDK handles CAE
    # internally, or that can't re-mint (pure BYO), leave this $false and the caller
    # surfaces the 401 unchanged.
    [bool]$SupportsClaimsChallenge  = $false

    # Optional. The provider is wired directly into the module's built-in
    # Connect-EntraEnvironment / Connect-IntuneManagement entry points (the original
    # pre-abstraction MSAL path). Callers that want that rich handling (sovereign-cloud
    # -Cloud selection, ForceRefresh by TokenId, ...) drive such a provider through those
    # functions instead of the generic Connect()/Refresh() hop, keeping behaviour and
    # stack traces identical. Providers whose logic lives entirely in their own Connect()
    # / Refresh() leave this $false.
    [bool]$UsesBuiltInConnectPath   = $false

    # Optional. Provider can launch an interactive admin/user consent prompt for the
    # app's delegated scopes (MSAL: Start-MSALConsentPrompt). The profile UI shows a
    # "Request consent" action only when this is $true.
    [bool]$SupportsConsentPrompt    = $false

    # The provider answers every Graph request itself through InvokeWebRequest,
    # even though GetAccessToken returned a bearer. Off for the real providers -
    # a bearer means "send it over HTTPS". On for a provider whose backend is not
    # Graph at all (the offline mock tenant serves canned data from disk).
    [bool]$RoutesAllRequests        = $false

    # Always Graph for now; -Resource is plumbed through so future phases can mint
    # tokens for other audiences (Key Vault, Storage, etc.) without API changes.
    [string]$DefaultResource = "https://graph.microsoft.com"

    # Called once at module init for provider-specific setup (DLL loads, settings).
    # Default is a no-op.
    [void] Initialize() { }

    # === Auth lifecycle ===
    # Must override:
    [PSCustomObject] Connect([hashtable]$Arguments) {
        throw "Provider '$($this.Id)' does not implement Connect"
    }

    # Optional (return $false if not supported):
    [bool] Disconnect([int]$TokenId)                                    { return $false }
    [bool] ForgetAccount([string]$AccountIdentifier)                    { return $false }
    [bool] Refresh([int]$TokenId)                                       { return $false }

    # List the tenants the signed-in account can reach, for a tenant picker or a
    # pre-flight check before Connect -TenantId. Microsoft Graph has no API for
    # this: findTenantInformationByTenantId resolves ONE tenant you already know,
    # and the multi-tenant-organization APIs only cover a configured MTO. The only
    # general source is Azure Resource Manager's /tenants, which needs a token for
    # a different audience - so a provider can implement this only if it can mint
    # one for the same account.
    #
    # Return $null when the provider cannot enumerate (the default). Otherwise
    # return @{ Tenants = <rows>; ConsentMissing = <bool>; Message = <string> } so
    # the caller can tell "no tenants" from "the app lacks the permission".
    [PSCustomObject] GetAccessibleTenants([int]$TokenId)                { return $null }
    [PSCustomObject] SwitchTenant([int]$TokenId, [string]$NewTenantId) { return $null }

    # Called once at app startup (AppInitialized event) for the active provider only.
    # Implementations should attempt a SILENT (non-interactive) sign-in if their backend
    # has persisted credentials on disk. Return $true if the user is now signed in,
    # $false otherwise. Default is no-op — MSAL has its own cross-session refresh path
    # via $script:MSALDefaultToken on startup, so it doesn't need this hook.
    [bool] TryResumeSession() { return $false }

    # Optional. Cheap SILENT ambient-session refresh, invoked on view activation to keep
    # the default token fresh without ever prompting. Providers that have no such
    # mechanism - or where a silent auth here could hijack the active session - leave the
    # base no-op. (MSAL refreshes via Connect-EntraEnvironment -ForceSilent.)
    [void] RefreshAmbientSession() { }

    # === Token retrieval ===
    # Must override. -Resource is informational for providers that mint per-audience
    # tokens; today only Graph is required.
    [string] GetAccessToken([int]$TokenId, [string]$Resource) {
        throw "Provider '$($this.Id)' does not implement GetAccessToken"
    }

    # Optional. Returns DateTime.MaxValue when expiry is unknown (callers should treat
    # MaxValue as "no preflight refresh needed").
    [datetime] GetAccessTokenExpiry([int]$TokenId, [string]$Resource) {
        return [datetime]::MaxValue
    }

    # Optional. Re-mint an access token to satisfy a CAE claims challenge returned by
    # the resource (Graph 401 -> WWW-Authenticate claims="..."). Implemented only by
    # providers with SupportsClaimsChallenge = $true; the base returns $null so a
    # provider that can't satisfy the challenge simply lets the 401 surface.
    #   $Resource         - target audience (informational; Graph today)
    #   $ClaimsChallenge  - the claims value parsed from the 401 challenge
    #   $AllowInteractive - caller context: $true when a UI is present and this is NOT a
    #                       nested auth-flow call, so the provider MAY prompt if it can't
    #                       satisfy the challenge silently; $false forces silent-only.
    # Returns the new access token, or $null if the challenge couldn't be satisfied.
    [string] GetClaimsToken([int]$TokenId, [string]$Resource, [string]$ClaimsChallenge, [bool]$AllowInteractive) {
        return $null
    }

    # Optional. Providers whose backend SDK owns the HTTP pipeline and won't hand out a
    # raw bearer token override this to run the request themselves and return a response
    # object shaped like Invoke-WebRequest's (StatusCode / Headers / Content /
    # RawContentLength). Return $null to signal "I don't route requests - use the raw
    # access token from GetAccessToken via Invoke-WebRequest". Base default: $null.
    [object] InvokeWebRequest([string]$Url, [string]$Method, [object]$Body, [hashtable]$Headers) {
        return $null
    }

    # === Display / picker data (uniform shapes for UI consumption) ===

    # Returns a PSCustomObject with these properties (or $null if unknown):
    #   Provider     - this.Id
    #   DisplayName  - user's display name
    #   UPN          - user principal name (login id)
    #   UserId       - tenant-scoped object id
    #   TenantId     - GUID
    #   TenantName   - organization display name
    #   AppId        - Entra app client id
    #   AppName      - Entra app display name
    #   AuthType     - "Interactive" | "ClientCredential" | "BYO" | "DeviceCode" | ...
    #   ExpiresOn    - DateTime (local) or $null
    [PSCustomObject] GetUserInfo([int]$TokenId) { return $null }

    # Returns PSCustomObject[] with these per-row properties:
    #   Provider, Username, UserId, TenantId, Native (provider-opaque)
    [PSCustomObject[]] GetCachedAccounts() { return [PSCustomObject[]]@() }

    # Returns PSCustomObject[] with these per-row properties:
    #   Provider, TenantId, TenantName, Native
    [PSCustomObject[]] GetAvailableTenants([int]$TokenId) { return [PSCustomObject[]]@() }

    # Returns Name/Value rows describing the provider's native session/context for the
    # profile "Session Info" inspector (MSAL: AuthenticationResult fields; MgGraph:
    # Get-MgContext properties). Base default: no rows.
    [PSCustomObject[]] GetSessionInfoRows() { return [PSCustomObject[]]@() }

    # Decoded id-token JWT ({Header, Payload}) for the profile popup's "Id Token"
    # inspector, or $null when the provider doesn't surface an id token. The UI shows
    # the Id Token button only when this returns non-null, so non-MSAL providers hide
    # it automatically. Keeps raw-JWT knowledge inside the owning provider.
    [object] GetIdTokenJwt([int]$TokenId) { return $null }

    # === Provider-specific UI hook ===
    # MSAL adds "MSAL Token / Access Token / ID Token" inspector buttons here, for
    # example. Returns a WPF element to splice into the profile popup, or $null.
    [object] BuildLoginMenu([object]$Window) { return $null }
}
