#ImportOrder 24

# Canonical, provider-agnostic authentication token descriptor.
#
# One typed shape used everywhere a token/identity is surfaced: the auth-event
# payloads (AuthenticatedNewToken / AuthenticationTokenRefresh /
# AuthenticationUserDisconnected), every provider's GetUserInfo() return value,
# the central token registry in Internal/AuthenticationCore.ps1, and the public
# Get-IMAuthToken function.
#
# Providers populate everything they can see from the token itself (Provider,
# TenantId, TenantName, Cloud, Account/UPN/UserId, AppId/AppName, AuthType,
# ExpiresOn). The registry overlays the two fields a provider cannot know on its
# own: the GLOBAL TokenId (registry-allocated, unique across providers) and
# IsDefault (derived from the registry's single default-id).
#
# ImportOrder 24: must parse before AuthenticationProvider.ps1 (#ImportOrder 25),
# whose GetUserInfo signature returns [IMAuthToken], and the concrete providers
# (26/27). Core primitives load at 1/10, so 24 is free and safely after them.

class IMAuthToken {
    [int]                $TokenId       # GLOBAL id (registry-allocated); 0 = unassigned
    [string]             $Provider      # owning provider Id: "MSAL" | "OAuth" | "MgGraph"
    [string]             $TenantId
    [string]             $TenantName
    [string]             $Cloud         # "Public" | "USGov" | "USGovDOD" | "China"
    [string]             $Account       # display name (UPN, or app name for app-only)
    [string]             $UPN
    [string]             $UserId
    [string]             $AppId
    [string]             $AppName
    [string]             $AuthType      # Interactive | ClientCredential | BYO | ManagedIdentity | WorkloadFederation | Password
    [bool]               $IsDefault
    [Nullable[datetime]] $ExpiresOn

    [string] ToString() {
        $tenant = if ($this.TenantName) { $this.TenantName } elseif ($this.TenantId) { $this.TenantId } else { '?' }
        return "$($this.Provider)/$tenant ($($this.TokenId))"
    }
}
