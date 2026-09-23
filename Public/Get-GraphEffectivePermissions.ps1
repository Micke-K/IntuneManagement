function Get-GraphEffectivePermissions {
    <#
    .SYNOPSIS
        What the signed-in identity can actually do per policy type: the app's
        token scopes combined with the user's Intune role permissions.

    .DESCRIPTION
        For a delegated login the effective access to Intune is the intersection of
        two things: the scopes the APP was consented (the token's scp claim) and
        the Intune RBAC / Entra directory roles of the USER. The token only shows
        the first. This command evaluates both for every registered policy type
        and reports the combined level, so a script can find out before a bulk
        import that the user is read-only for Device configurations instead of
        collecting 403s halfway through.

        Each row carries two views of the same answer. The concrete capability
        labels are what the Permissions popup shows: Required (Read or ReadWrite -
        what the type needs), TokenAccess / RoleAccess / EffectiveAccess (None,
        Read or ReadWrite - what each layer grants in those terms; RoleAccess is
        $null when no role layer applies) and Result (Match / Read-only / No
        access). The *Level fields below are the raw enum kept for programmatic
        callers.

        Levels: Full (usable), Limited (readable, not writable), None (unusable).
        RbacLevel carries the user-role verdict: the Intune RBAC level for
        Intune-governed types, or - for Entra ID objects such as Conditional
        Access and Named Locations - the level implied by the directory roles in
        the token (Conditional Access / Security Administrator = write, Global /
        Security Reader = read-only). It is $null when no role governs the type:
        an app-only token (application permissions bypass Intune RBAC), a type
        neither layer governs (Terms of Use, branding), no governing directory
        role in the token, or a failed lookup. EffectiveLevel is then the TokenLevel.

        The Intune answer is read once per token and refreshed when the token is
        re-issued (Force refresh in the Profile popup, or a new sign-in). Scope
        tags are not modelled: a user scoped to some tags can still be refused on
        individual objects.

        Exported as Get-IMGraphEffectivePermissions (the module applies the 'IM'
        command prefix).

    .PARAMETER TokenId
        Evaluate the token with this id (see Get-IMAuthToken) instead of the
        default token.

    .PARAMETER PolicyType
        Only report these policy type ids (e.g. DeviceConfiguration, SettingsCatalog).

    .PARAMETER Raw
        Return the Intune RBAC context itself - source (DirectoryRole shortcut or
        Graph), the resource actions the user is Allowed, the Catalog of actions
        Intune defines at all (from deviceManagement/resourceOperations; $null when
        that call failed) and the raw getEffectivePermissions response - instead of
        the per-type table. $null when RBAC does not apply.

    .EXAMPLE
        # Types the signed-in user cannot change, with the reason
        Get-IMGraphEffectivePermissions | Where-Object EffectiveLevel -ne Full |
            Format-Table Id, TokenLevel, RbacLevel, EffectiveLevel, Reason

    .EXAMPLE
        # Abort a bulk import if Settings Catalog is not writable
        $sc = Get-IMGraphEffectivePermissions -PolicyType SettingsCatalog
        if($sc.EffectiveLevel -ne 'Full') { throw "Settings Catalog: $($sc.Reason)" }

    .EXAMPLE
        # Every Intune resource action the user is allowed
        (Get-IMGraphEffectivePermissions -Raw).Allowed | Sort-Object
    #>
    [CmdletBinding()]
    param(
        [int]$TokenId = 0,

        [ArgumentCompleter({ & (Get-Module IntuneManagement) { Get-IntunePolicyTypeValues } })]
        [string[]]$PolicyType,

        [switch]$Raw
    )

    $claims = Get-AccessTokenClaims -TokenId $TokenId
    if(-not $claims) {
        Write-Log "Get-GraphEffectivePermissions: no access token available. Connect first (Connect-IMIntuneManagement)." 2
        return
    }

    $granted = Get-GrantedGraphPermissions ([PSCustomObject]@{ JWTAccessToken = [PSCustomObject]@{ Payload = $claims } })
    $rbac    = Get-IntuneRbacContext -Claims $claims -TokenId $TokenId -IgnoreSetting

    if($Raw) { return $rbac }

    $types = @($script:IntuneTypes | Where-Object { $_ })
    if($PolicyType) { $types = @($types | Where-Object { $_.Id -in $PolicyType }) }

    foreach($type in $types) {
        $tokenLevel = Get-PolicyTypeAccessLevel $type $granted
        $tokenInfo  = Get-PolicyTypeAccessInfo $type $granted $tokenLevel

        $grantedSet = if($granted) { ConvertTo-PermissionSet $granted } else { $null }
        $missingScopes = @()
        foreach($perm in @($type._Permissions | Where-Object { $_ })) {
            if(-not $grantedSet) { continue }
            if($grantedSet.Contains($perm)) { continue }
            # A required Read scope is covered by the granted ReadWrite superset.
            $writeVariant = Get-PermissionWriteVariant $perm
            if($writeVariant -and $grantedSet.Contains($writeVariant)) { continue }
            $missingScopes += $perm
        }

        # Intune-governed types get the RBAC verdict; Entra-governed types
        # (Conditional Access etc.) fall back to the directory-role verdict.
        $verdict = if($rbac) { Get-PolicyTypeRbacAccess $type $rbac } else { $null }
        if(-not $verdict) { $verdict = Get-PolicyTypeEntraRoleAccess $type $claims }
        # Catch-all: Global Reader reads the whole tenant but writes nothing.
        if(-not $verdict) { $verdict = Get-PolicyTypeDirectoryRoleReadFloor $type $claims }
        $rbacLevel = if($verdict) { $verdict.Level } else { $null }
        $effective = if($verdict) { Get-WorstAccessLevel $tokenLevel $verdict.Level } else { $tokenLevel }
        $category  = Get-PolicyTypeRbacCategory $type

        # Concrete capability labels for the popup: what the type needs (Required),
        # what each layer grants in those terms (None/Read/ReadWrite), and the
        # bottom line (Match/Read-only/No access). The *Level fields above stay for
        # programmatic callers that compare against 'Full'/'Limited'/'None'.
        $required = Get-PolicyTypeRequiredAccess $type

        # A role that allows some writes but not all is Limited, but not read-only.
        # The role column says so whenever the verdict does; the effective and
        # result columns only when the token can write too - a read-only token
        # over a partial-write role really is read-only.
        $rolePartial      = [bool]($verdict -and $verdict.Partial)
        $effectivePartial = $rolePartial -and $tokenLevel -eq [APIAccess]::Full

        [PSCustomObject]@{
            Id               = $type.Id
            Title            = $type.Title
            Group            = if($type.PolicyGroup) { $type.PolicyGroup.Id } else { $null }
            ResourceCategory = if($category) { $category.Category } else { $null }
            Required         = $required
            TokenAccess      = Get-AccessCapabilityLabel $tokenLevel $required
            RoleAccess       = if($verdict) { Get-AccessCapabilityLabel $verdict.Level $required -PartialWrite:$rolePartial } else { $null }
            EffectiveAccess  = Get-AccessCapabilityLabel $effective $required -PartialWrite:$effectivePartial
            Result           = Get-AccessResultLabel $effective -PartialWrite:$effectivePartial
            TokenLevel       = [string]$tokenLevel
            RbacLevel        = if($null -ne $rbacLevel) { [string]$rbacLevel } else { $null }
            EffectiveLevel   = [string]$effective
            MissingScopes    = $missingScopes
            MissingActions   = if($verdict) { @($verdict.Missing) } else { @() }
            Reason           = (@($tokenInfo, $(if($verdict) { $verdict.Info })) | Where-Object { $_ }) -join "; "
            RbacSource       = if($rbac) { $rbac.Source } else { $null }
        }
    }
}
