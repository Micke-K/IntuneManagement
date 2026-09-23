# Permission-based access marking for the left-hand navigation.
#
# Every registered policy type declares the Graph permissions it needs in
# $this._Permissions. The signed-in token carries what was actually granted:
# 'scp' (space-separated string) on delegated logins, 'roles' (array) on
# app-only logins. Diffing the two yields a per-type access level that the nav
# renders as a colour, via the APIAccess enum in Classes/IntuneBaseClasses.ps1:
#
#   Full    - every declared permission is granted                (no colour)
#   Limited - the type declares a ReadWrite permission but only   (orange)
#             the matching Read variant is granted: readable,
#             not writable
#   None    - at least one declared permission is granted in      (red)
#             neither form, so the type cannot be used at all
#
# The two aggregations are deliberately different in kind:
#
#   * Within a TYPE the declared permissions are conjunctive - a type needs all
#     of them - so the worst permission decides the type.
#   * Across a GROUP the member types are independent, so None is reserved for
#     "nothing in this group is usable". Anything short of uniformly-Full but
#     not wholly-None is Limited. Without that asymmetry a group with one
#     inaccessible child out of five would render red and imply total denial.
#
# A type that declares only a Read permission (several do) is Full when that
# Read permission is granted - it is a read-only feature by design, not a
# degraded one.
#
# When there is no token, or the token carries no scp/roles claims at all,
# everything is left at the Full default so the nav shows no colour. A sea of
# red on a token we simply could not inspect is worse than no signal.

# Turn a ReadWrite permission into its Read counterpart, or $null when the
# permission is not ReadWrite-shaped. Covers both layouts in use:
#   DeviceManagementConfiguration.ReadWrite.All -> DeviceManagementConfiguration.Read.All
#   Policy.ReadWrite.ConditionalAccess          -> Policy.Read.ConditionalAccess
function Get-PermissionReadVariant {
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Permission)

    if([string]::IsNullOrWhiteSpace($Permission)) { return $null }
    if($Permission -notmatch '\.ReadWrite\.') { return $null }
    return ($Permission -replace '\.ReadWrite\.', '.Read.')
}

# Turn a Read permission into its ReadWrite counterpart, or $null when the
# permission is not Read-shaped. The mirror of Get-PermissionReadVariant, used to
# satisfy a required Read scope from the granted ReadWrite superset:
#   DeviceManagementConfiguration.Read.All -> DeviceManagementConfiguration.ReadWrite.All
#   Policy.Read.ConditionalAccess          -> Policy.ReadWrite.ConditionalAccess
# A type that declares only a bare Read permission (the read-only-by-design types
# in Classes/IntuneInfoClasses.ps1 etc.) is fully usable when the token carries
# the ReadWrite scope for the same resource - ReadWrite is a strict superset of
# Read in Graph - even though the separate Read scope was never consented. Without
# this the type was falsely marked None while the read call itself succeeded.
function Get-PermissionWriteVariant {
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Permission)

    if([string]::IsNullOrWhiteSpace($Permission)) { return $null }
    if($Permission -notmatch '\.Read\.') { return $null }
    return ($Permission -replace '\.Read\.', '.ReadWrite.')
}

# Normalise any permission collection into a case-insensitive HashSet.
#
# This exists because PowerShell enumerates collections on output, so a bare
# `return $hashSet` hands the caller a plain string or object[] instead. On
# those, .Contains() resolves to String.Contains / IList.Contains - a
# case-SENSITIVE, and for a string even SUBSTRING, match. That silently
# produces wrong access levels, so every entry point normalises first and
# returns with the unary comma to stop the unrolling.
function ConvertTo-PermissionSet {
    [CmdletBinding()]
    param($Permissions)

    if($Permissions -is [System.Collections.Generic.HashSet[string]]) { return ,$Permissions }

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach($p in @($Permissions)) {
        if($p) { [void]$set.Add(([string]$p).Trim()) }
    }
    return ,$set
}

# Decoded claims of the access token currently in play, or $null.
#
# Deliberately provider-agnostic. Reading $script:MSALDefaultToken directly
# would only ever work for the MSAL provider - an OAuth or MgGraph session has
# no such global, so the permission marking silently did nothing there. The
# active provider is asked for its token through the AuthenticationCore facade
# instead, exactly as Invoke-MSGraphAPI does.
function Get-AccessTokenClaims {
    [CmdletBinding()]
    param($TokenInfo, [int]$TokenId = 0)

    # 1. An explicit TokenInfo that already carries a decoded JWT (MSAL paths,
    #    and how the unit tests inject a payload).
    if($TokenInfo) {
        $payload = $null
        try { $payload = $TokenInfo.JWTAccessToken.Payload } catch { }
        if($payload) { return $payload }
    }

    # 2. The token's OWN provider (the explicit -TokenId, else the default token),
    #    via its access token. Get-GraphDomain keeps the resource correct in
    #    sovereign clouds.
    #
    #    Routing by owner and not by "whichever provider is active" matters as soon
    #    as a second login is live: these claims are combined with the user's Intune
    #    RBAC, which is read through Invoke-MSGraphAPI and therefore routed to the
    #    token's owner. Asking the active provider for someone else's token id would
    #    pair one account's scopes with another account's role. Id 0 / an
    #    unregistered id still falls back to the active provider (Get-AuthProvider
    #    treats an empty ProviderId as "the active one"), as before.
    $tokenId = if($TokenId -gt 0) { $TokenId } else { Get-DefaultAuthTokenId }
    $owner   = $null
    try { $owner = Resolve-AuthTokenProvider $tokenId } catch { }

    try {
        $domain = $null
        try { $domain = Get-GraphDomain $tokenId } catch { }
        if(-not $domain) { $domain = "graph.microsoft.com" }

        $providerId = if($owner) { $owner.Id } else { $null }

        $accessToken = Get-AuthProviderAccessToken -TokenId $tokenId -Resource "https://$domain" -ProviderId $providerId
        if($accessToken) {
            $jwt = Get-JWTtoken $accessToken
            if($jwt -and $jwt.Payload) { return $jwt.Payload }
        }
    }
    catch {
        Write-LogDebug "Get-AccessTokenClaims: could not read the token's provider: $($_.Exception.Message)"
    }

    # 3. Legacy MSAL global, for callers that run before the registry is set up -
    #    and ONLY for those. Once a provider owns the token, its failure to produce a
    #    bearer means "cannot tell", not "use the MSAL global": that global belongs to
    #    a different sign-in, so returning its claims would pair one account's scopes
    #    with another account's Intune role, which is the mixing step 2 exists to
    #    prevent. No owner (id 0, or an id the registry does not know) is the only
    #    case where there is no identity to contradict.
    if($owner) {
        Write-LogDebug "Get-AccessTokenClaims: no token from provider $($owner.Id) for token id $tokenId - not falling back to the MSAL global"
        return $null
    }

    $payload = $null
    try { $payload = $script:MSALDefaultToken.JWTAccessToken.Payload } catch { }
    return $payload
}

# Collect the permissions granted by a token into a case-insensitive set.
# Returns $null (not an empty set) when the token carries no permission claims,
# so callers can tell "nothing granted" from "cannot tell".
function Get-GrantedGraphPermissions {
    [CmdletBinding()]
    param($TokenInfo)

    $payload = Get-AccessTokenClaims $TokenInfo
    if(-not $payload) { return $null }

    $granted = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if($payload.scp) {
        foreach($s in ([string]$payload.scp).Split(' ')) {
            if($s) { [void]$granted.Add($s.Trim()) }
        }
    }
    if($payload.roles) {
        foreach($r in @($payload.roles)) {
            if($r) { [void]$granted.Add(([string]$r).Trim()) }
        }
    }

    if($granted.Count -eq 0) { return $null }
    # Unary comma: without it PowerShell unrolls the set. See ConvertTo-PermissionSet.
    return ,$granted
}

# Access level for a single policy type. $Granted is the set from
# Get-GrantedGraphPermissions. A type declaring no permissions is Full - we
# have nothing to judge it on and must not invent a warning.
function Get-PolicyTypeAccessLevel {
    [CmdletBinding()]
    [OutputType([APIAccess])]
    param($PolicyType, $Granted)

    if(-not $PolicyType -or -not $Granted) { return [APIAccess]::Full }

    $required = @($PolicyType._Permissions | Where-Object { $_ })
    if($required.Count -eq 0) { return [APIAccess]::Full }

    $grantedSet = ConvertTo-PermissionSet $Granted

    $worst = [APIAccess]::Full
    foreach($perm in $required) {
        if($grantedSet.Contains($perm)) { continue }

        # A required Read scope is fully satisfied by holding the ReadWrite
        # superset (ReadWrite implies Read in Graph), so the type is usable even
        # if the bare Read scope was never consented.
        $writeVariant = Get-PermissionWriteVariant $perm
        if($writeVariant -and $grantedSet.Contains($writeVariant)) { continue }

        $readVariant = Get-PermissionReadVariant $perm
        if($readVariant -and $grantedSet.Contains($readVariant)) {
            # Readable but not writable. Keep looking - a later permission may
            # be missing outright, which outranks this.
            if($worst -eq [APIAccess]::Full) { $worst = [APIAccess]::Limited }
            continue
        }

        # Granted in neither form: the type cannot function.
        return [APIAccess]::None
    }
    return $worst
}

# Human-readable detail for a type's tooltip. Empty when the type is Full.
function Get-PolicyTypeAccessInfo {
    [CmdletBinding()]
    [OutputType([string])]
    param($PolicyType, $Granted, [APIAccess]$Level)

    if($Level -eq [APIAccess]::Full) { return "" }
    if(-not $PolicyType -or -not $Granted) { return "" }

    $grantedSet = ConvertTo-PermissionSet $Granted

    $readOnly = @()
    $missing  = @()
    foreach($perm in @($PolicyType._Permissions | Where-Object { $_ })) {
        if($grantedSet.Contains($perm)) { continue }
        # Read satisfied by the ReadWrite superset: not missing at all.
        $writeVariant = Get-PermissionWriteVariant $perm
        if($writeVariant -and $grantedSet.Contains($writeVariant)) { continue }
        $readVariant = Get-PermissionReadVariant $perm
        if($readVariant -and $grantedSet.Contains($readVariant)) { $readOnly += $perm }
        else { $missing += $perm }
    }

    $parts = @()
    if($readOnly.Count -gt 0) { $parts += "Read-only: missing $($readOnly -join ', ')" }
    if($missing.Count  -gt 0) { $parts += "No access: missing $($missing -join ', ')" }
    return ($parts -join "`n")
}

# The access a type NEEDS to be fully usable: "ReadWrite" when it declares a
# write scope, "Read" when it is read-only by design, "" when it declares no
# permission at all (nothing to judge). Drives the Permissions popup's Required
# column so "Full" is never shown where it would read as "I have write".
function Get-PolicyTypeRequiredAccess {
    [CmdletBinding()]
    [OutputType([string])]
    param($PolicyType)

    $perms = @($PolicyType._Permissions | Where-Object { $_ })
    if($perms.Count -eq 0) { return "" }
    foreach($perm in $perms) { if(Get-PermissionReadVariant $perm) { return "ReadWrite" } }
    return "Read"
}

# Map an APIAccess level to the concrete capability it represents for a type:
# None -> "None"; Limited -> "Read" (readable, not writable - only writable types
# ever reach Limited); Full -> the type's Required level (ReadWrite for a writable
# type, Read for a read-only one). This is what the popup shows instead of the
# bare enum, so a read-only feature reads as "Read", not "Full".
#
# Limited has two shapes, and the enum cannot tell them apart: a role that allows
# no write action at all, and one that allows some (Create and Update but not
# Delete or Assign is a common custom role). -PartialWrite names the second, so
# the label says "Partial write" rather than "Read" - which would contradict the
# tooltip that already says the user can write.
function Get-AccessCapabilityLabel {
    [CmdletBinding()]
    [OutputType([string])]
    param([APIAccess]$Level, [string]$RequiredAccess, [switch]$PartialWrite)

    switch($Level) {
        ([APIAccess]::None)    { return "None" }
        ([APIAccess]::Limited) { if($PartialWrite) { return "Partial write" } else { return "Read" } }
        default                { if($RequiredAccess) { return $RequiredAccess } else { return "Read" } }
    }
}

# The bottom-line verdict for the Effective column: Full -> "Match" (you have what
# the type needs), Limited -> "Read-only" (you can read a type that needs write)
# or "Partial write" when -PartialWrite says some writes are allowed,
# None -> "No access".
function Get-AccessResultLabel {
    [CmdletBinding()]
    [OutputType([string])]
    param([APIAccess]$EffectiveLevel, [switch]$PartialWrite)

    switch($EffectiveLevel) {
        ([APIAccess]::None)    { return "No access" }
        ([APIAccess]::Limited) { if($PartialWrite) { return "Partial write" } else { return "Read-only" } }
        default                { return "Match" }
    }
}

# Aggregate member-type levels into a group level. See the header for why None
# requires ALL children to be None.
function Get-PolicyGroupAccessLevel {
    [CmdletBinding()]
    [OutputType([APIAccess])]
    param([APIAccess[]]$ChildLevels)

    $levels = @($ChildLevels)
    if($levels.Count -eq 0) { return [APIAccess]::Full }

    $noneCount = @($levels | Where-Object { $_ -eq [APIAccess]::None }).Count
    if($noneCount -eq $levels.Count) { return [APIAccess]::None }

    $fullCount = @($levels | Where-Object { $_ -eq [APIAccess]::Full }).Count
    if($fullCount -eq $levels.Count) { return [APIAccess]::Full }

    return [APIAccess]::Limited
}

# Tooltip breakdown for a group, e.g. "2 of 5 read-only, 1 of 5 no access".
function Get-PolicyGroupAccessInfo {
    [CmdletBinding()]
    [OutputType([string])]
    param([APIAccess[]]$ChildLevels)

    $levels = @($ChildLevels)
    if($levels.Count -eq 0) { return "" }

    $limited = @($levels | Where-Object { $_ -eq [APIAccess]::Limited }).Count
    $none    = @($levels | Where-Object { $_ -eq [APIAccess]::None }).Count
    if(($limited + $none) -eq 0) { return "" }

    $parts = @()
    if($limited -gt 0) { $parts += "$limited of $($levels.Count) read-only" }
    if($none    -gt 0) { $parts += "$none of $($levels.Count) no access" }
    return ($parts -join ', ')
}

# Stamp AccessType / AccessInfo onto every registered policy type and group.
# Called from each backend's Get-IntuneViewItems, so it re-runs on every menu
# rebuild (login, view-mode switch, settings change) without needing its own
# event subscription. Safe to call with no token: everything resets to Full.
function Update-IntuneAccessLevels {
    [CmdletBinding()]
    param($TokenInfo)

    # Resolve the claims separately from the permission set so we can tell
    # "signed out" (say nothing) from "signed in but the token carries no
    # scp/roles" (worth a warning - it means marking cannot work at all).
    $claims  = Get-AccessTokenClaims $TokenInfo

    # An expired default token is effectively signed out. A JWT still DECODES
    # after expiry, so Get-AccessTokenClaims can hand back stale claims and leave
    # the nav marked (coloured) for a session that can no longer call Graph. When
    # we are marking from the default token (no explicit TokenInfo), treat a
    # confirmed-expired token as no token so the marking resets to Full - the same
    # end state as a sign-out - until a silent refresh or new login re-marks it.
    # Test-DefaultTokenExpired is provider-agnostic and returns $false for still
    # valid tokens and SDK-managed (MgGraph) sessions, so this never trips
    # mid-session on a transient failure where the token is still good.
    if(-not $TokenInfo -and $claims -and (Test-DefaultTokenExpired)) {
        Write-LogDebug "Update-IntuneAccessLevels: default token has expired; resetting access marking to Full"
        $claims = $null
    }

    $granted = if($claims) { Get-GrantedGraphPermissions ([PSCustomObject]@{
                   JWTAccessToken = [PSCustomObject]@{ Payload = $claims } }) } else { $null }

    # Layer 2: the signed-in user's Intune RBAC (Internal/EffectivePermissions.ps1).
    # $null when it does not apply (setting off, app-only token, lookup failed).
    # It can only make a type worse, never better.
    $rbac = $null
    if($claims) { try { $rbac = Get-IntuneRbacContext -Claims $claims } catch { Write-LogDebug "Get-IntuneRbacContext failed: $($_.Exception.Message)" } }
    $rbacLimited = 0; $rbacNone = 0

    # Layer 2 for Entra ID objects (Conditional Access etc.): driven by the
    # directory roles in the token, no Graph call (Internal/EntraRoleAccessLevel.ps1).
    # Same setting gate as the Intune RBAC layer; also only downgrades.
    $entraClaims = if($claims -and (Test-RbacAccessMarkingEnabled)) { $claims } else { $null }

    $types = @($script:IntuneTypes | Where-Object { $_ })
    foreach($type in $types) {
        $level = Get-PolicyTypeAccessLevel $type $granted
        $info  = Get-PolicyTypeAccessInfo $type $granted $level
        # A type is either Intune-governed (RBAC verdict) or Entra-governed
        # (directory-role verdict) - never both, so try RBAC first and fall back.
        $verdict = $null
        if($rbac) { $verdict = Get-PolicyTypeRbacAccess $type $rbac }
        if(-not $verdict -and $entraClaims) { $verdict = Get-PolicyTypeEntraRoleAccess $type $entraClaims }
        # Catch-all: a Global Reader reads the whole tenant but writes nothing, so
        # any writable type the two layers above did not resolve is read-only.
        if(-not $verdict -and $entraClaims) { $verdict = Get-PolicyTypeDirectoryRoleReadFloor $type $entraClaims }
        if($verdict -and $verdict.Level -ne [APIAccess]::Full) {
            $merged = Get-WorstAccessLevel $level $verdict.Level
            if($merged -ne $level) { if($merged -eq [APIAccess]::None) { $rbacNone++ } else { $rbacLimited++ } }
            $level = $merged
            $info  = (@($info, $verdict.Info) | Where-Object { $_ }) -join "`n"
        }
        $type.AccessType = $level
        $type.AccessInfo = $info
    }

    foreach($group in @($script:IntuneGroups | Where-Object { $_ })) {
        # Prefer the group's own member list; fall back to scanning the type
        # registry for types wired to this group.
        $members = @($group._PolicyTypes | Where-Object { $_ })
        if($members.Count -eq 0) {
            $members = @($types | Where-Object { $_.PolicyGroup -and $_.PolicyGroup.Id -eq $group.Id })
        }

        $childLevels = @($members | ForEach-Object { $_.AccessType })
        $group.AccessType = Get-PolicyGroupAccessLevel $childLevels
        $group.AccessInfo = Get-PolicyGroupAccessInfo $childLevels
    }

    # Report at warning level, not debug. The first cut logged this through
    # Write-LogDebug, which the Debug setting suppresses by default - so an
    # unmarked nav produced no explanation anywhere and looked like the feature
    # had simply not shipped.
    if(-not $claims) {
        Write-LogDebug "Update-IntuneAccessLevels: no access token available yet; access marking left at Full"
        return
    }
    if(-not $granted) {
        Write-Log ("Access marking unavailable: the access token carries no 'scp' or 'roles' claim, so " +
                   "per-type permissions cannot be determined. The navigation is left unmarked.") 2
        return
    }

    $limited = @($types | Where-Object { $_.AccessType -eq [APIAccess]::Limited }).Count
    $none    = @($types | Where-Object { $_.AccessType -eq [APIAccess]::None }).Count
    if(($limited + $none) -eq 0) {
        Write-Log "Access marking: all $($types.Count) policy types are fully accessible with the current token"
    }
    else {
        Write-Log ("Access marking: of $($types.Count) policy types, $limited are read-only " +
                   "(orange) and $none have no access (red) with the current token.") 2
    }
    if(($rbacLimited + $rbacNone) -gt 0) {
        Write-Log ("Access marking: the signed-in user's role (Intune RBAC or Entra directory role) lowered " +
                   "$rbacLimited type(s) to read-only and $rbacNone to no access beyond what the token scopes " +
                   "allow. Refresh the token from the Profile popup after a role change.") 2
    }
}
