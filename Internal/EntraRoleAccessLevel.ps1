# Layer 2 for Entra ID (directory) objects: Conditional Access and its siblings.
#
# Internal/EffectivePermissions.ps1 asks Intune "what can this user do" via
# deviceManagement/getEffectivePermissions. That call is Intune-only - it knows
# nothing about Conditional Access, Named Locations, Authentication Strengths or
# Authentication Context, which are Entra ID directory objects
# (identity/conditionalAccess/*) governed by Entra RBAC, not Intune RBAC.
#
# Entra ID has no equivalent per-object "getEffectivePermissions" surfaced as a
# simple Graph call, so this layer reads the directory roles already carried in
# the token's `wids` claim (no extra Graph traffic) and marks from the well-known
# roles that govern Conditional Access:
#
#   * Global Administrator                 -> Full (governs everything).
#   * Conditional Access Administrator     -> write (and therefore read).
#   * Security Administrator               -> write (and therefore read).
#   * Global Reader / Security Reader      -> read only -> a writable type is
#                                             marked read-only (orange).
#
# Conservative on purpose - it only ever DOWNGRADES (Get-WorstAccessLevel), and
# only when the token proves a governing role is present:
#   * No governing role in wids -> Unknown ($null). A custom directory role can
#     grant Conditional Access access without being one of the ids above, so a
#     missing role is NOT reported as no-access; Layer 1 (token scopes) stands.
#     This layer therefore never produces None, only Full / read-only / Unknown.
#   * A user who holds Global Reader AND a custom role that grants write would be
#     shown read-only. That combination is unusual and read-only is the safe
#     reading; scope-limited administrative units are likewise not modelled, the
#     same caveat Internal/EffectivePermissions.ps1 carries for Intune scope tags.
#
# Terms of Use (identityGovernance/termsOfUse) is deliberately NOT covered: it is
# a different API family and permission model (Agreement.ReadWrite.All), so
# Layer 1 alone marks it.

# Directory role template ids are fixed across tenants (wids carries the template
# id, not the tenant's role object id).
$script:EntraDirectoryRoleTemplateIds = @{
    GlobalAdministrator            = "62e90394-69f5-4237-9190-012177145e10"
    GlobalReader                   = "f2ef992c-3afb-46b9-b7cf-a126ee74c451"
    SecurityAdministrator          = "194ae4cb-b126-40b2-bd5b-6091b380977d"
    SecurityReader                 = "5d6b6bb7-de71-4623-b4af-96380a352509"
    ConditionalAccessAdministrator = "b1be1c3e-b65d-4f19-8427-f6fa0d97feb9"
}

# _API prefix -> the Entra roles that grant write / read for that area. Longest
# matching prefix wins (same rule as the Intune map). Global Administrator implies
# both everywhere and is handled separately, so it is not repeated here.
$script:EntraRoleApiMap = @{
    "identity/conditionalAccess" = @{
        Category   = "Conditional Access"
        WriteRoles = @("SecurityAdministrator", "ConditionalAccessAdministrator")
        ReadRoles  = @("GlobalReader", "SecurityReader")
    }
}

# The Entra-role entry for an API, or $null when this layer does not govern it.
function Get-EntraRoleCategoryForApi {
    [CmdletBinding()]
    param([string]$Api)

    if([string]::IsNullOrWhiteSpace($Api)) { return $null }
    $api = $Api.Trim().TrimStart('/')

    $best = $null
    foreach($key in $script:EntraRoleApiMap.Keys) {
        if($api -eq $key -or $api.StartsWith("$key/", [System.StringComparison]::OrdinalIgnoreCase)) {
            if(-not $best -or $key.Length -gt $best.Length) { $best = $key }
        }
    }
    if(-not $best) { return $null }
    return $script:EntraRoleApiMap[$best]
}

# The set of directory role template ids in the token, case-insensitive.
function Get-EntraDirectoryRoleIds {
    [CmdletBinding()]
    param($Claims)

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if(-not $Claims -or -not $Claims.PSObject.Properties['wids']) { return ,$set }
    foreach($w in @($Claims.wids)) {
        if($w) { [void]$set.Add(([string]$w).Trim()) }
    }
    return ,$set
}

# Layer 2 verdict for one Entra-governed policy type against the token claims.
# Returns @{ Level = [APIAccess]; Info; Missing = @() } or $null (Unknown) when
# this layer does not apply or cannot tell. Never returns None (see the header).
function Get-PolicyTypeEntraRoleAccess {
    [CmdletBinding()]
    param($PolicyType, $Claims)

    if(-not $PolicyType -or -not $Claims) { return $null }
    $entry = Get-EntraRoleCategoryForApi ([string]$PolicyType._API)
    if(-not $entry) { return $null }

    $wids = Get-EntraDirectoryRoleIds $Claims
    if($wids.Count -eq 0) { return $null }   # no directory roles to judge by

    if($wids.Contains($script:EntraDirectoryRoleTemplateIds.GlobalAdministrator)) {
        return @{ Level = [APIAccess]::Full; Info = ""; Missing = @() }
    }

    $hasWrite = $false
    foreach($role in @($entry.WriteRoles)) {
        $id = $script:EntraDirectoryRoleTemplateIds[$role]
        if($id -and $wids.Contains($id)) { $hasWrite = $true; break }
    }
    if($hasWrite) { return @{ Level = [APIAccess]::Full; Info = ""; Missing = @() } }

    $hasRead = $false
    foreach($role in @($entry.ReadRoles)) {
        $id = $script:EntraDirectoryRoleTemplateIds[$role]
        if($id -and $wids.Contains($id)) { $hasRead = $true; break }
    }
    if(-not $hasRead) { return $null }   # a role we do not model may still grant access

    # Read confirmed. Read-only-by-design types are Full; writable types are
    # marked read-only because a reader role grants no write.
    if(-not (Test-PolicyTypeDeclaresWrite $PolicyType)) { return @{ Level = [APIAccess]::Full; Info = ""; Missing = @() } }
    return @{
        Level   = [APIAccess]::Limited
        Info    = "Entra role: read-only for $($entry.Category) (directory role grants read, not write)"
        Missing = @()
    }
}

# Tenant-wide read-only floor from the Global Reader directory role, for the types
# neither the Intune RBAC layer (Internal/EffectivePermissions.ps1) nor the
# Conditional Access layer above resolve: Device Categories, Windows 365, Entra
# Branding, Terms of Use - anything whose API maps to no Intune RBAC category and
# is not Conditional Access, so it would otherwise sit at the token level (Full).
#
# Global Reader is the read-only twin of Global Administrator: it can read
# essentially every admin surface in the tenant but write none, so a writable
# type is read-only for it. This is the catch-all applied after the two
# API-specific layers return Unknown. Only ever DOWNGRADES:
#   * Global Administrator -> $null. It can write; Full from Layer 1 stands.
#   * Global Reader        -> a writable type => Limited (read-only); a
#                             read-only-by-design type => $null (Full stands).
#   * Neither role in wids -> $null. Layer 1 stands; a custom role may grant more.
# Never returns None: if the token also lacks the read scope, Layer 1 already
# marked the type None and a downgrade-only verdict cannot lift it. Scope-limited
# administrative units are not modelled (the same caveat as the layers above).
function Get-PolicyTypeDirectoryRoleReadFloor {
    [CmdletBinding()]
    param($PolicyType, $Claims)

    if(-not $PolicyType -or -not $Claims) { return $null }

    $wids = Get-EntraDirectoryRoleIds $Claims
    if($wids.Count -eq 0) { return $null }

    # A global writer keeps Full; only a pure global reader forces read-only.
    if($wids.Contains($script:EntraDirectoryRoleTemplateIds.GlobalAdministrator)) { return $null }
    if(-not $wids.Contains($script:EntraDirectoryRoleTemplateIds.GlobalReader))    { return $null }

    if(-not (Test-PolicyTypeDeclaresWrite $PolicyType)) { return $null }   # read-only feature: Full stands
    return @{
        Level   = [APIAccess]::Limited
        Info    = "Entra role: read-only (Global Reader grants read across the tenant, not write)"
        Missing = @()
    }
}
