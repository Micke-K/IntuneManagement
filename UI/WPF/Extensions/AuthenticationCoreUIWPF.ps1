# Provider-agnostic authentication UI helpers.
#
# Phase 1 placeholder. The profile popup and account picker still live in
# MSGraphAuthenticationUI.ps1 and read MSAL state directly. Phase 2 will migrate
# them here once the MgGraph provider lands and we have a second consumer to
# validate the abstraction.
#
# When that happens, expect these helpers to live here:
#   Show-AuthProviderProfilePopup    - build the popup from $provider.GetUserInfo
#   Add-AuthProviderCachedAccountRow - render a row for $provider.GetCachedAccounts
#   Add-AuthProviderTenantRow        - tenant switcher from $provider.GetAvailableTenants
#   Update-AuthProviderUI            - refresh the title-bar profile picture
#
# All of these will route through Get-AuthProvider so they work for any registered
# provider with no per-provider UI code.

# Currently empty by design. Do not delete — the file is part of the auth
# abstraction's stable file layout and Phase 2 will populate it.
