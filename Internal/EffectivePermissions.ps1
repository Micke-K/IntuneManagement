# Layer 2 of the access marking: the signed-in USER's Intune RBAC, on top of the
# APP's consented scopes that Internal/AccessLevel.ps1 (Layer 1) diffs.
#
# For a delegated login the effective access is the intersection of both. The
# token carries scp (app consent) and wids (Entra directory roles) but nothing
# about Intune role assignments - a user holding only the built-in Read Only
# Operator role signs in with a token that says ReadWrite.All, Layer 1 marks
# everything Full, and every PATCH gets a 403. This file asks Graph what the
# user can actually do and lets Update-IntuneAccessLevels stamp the worse of
# the two answers. See Docs/EffectivePermissions-Plan.md.
#
# Rules, in priority order:
#   * Setting off, no claims, or an app-only token -> no Layer 2 at all.
#     Application permissions bypass Intune RBAC, so 'roles' IS the answer.
#   * Intune Administrator / Global Administrator in wids -> Full for every
#     Intune category, no Graph call. Both grant complete Intune RBAC.
#   * Global Reader in wids -> read guaranteed for every Intune category, never
#     write. Used as a FLOOR, not a short-circuit: the getEffectivePermissions
#     call still runs (an additional Intune role may add write on top), but a
#     category the response omits is still treated as readable, and when the
#     call comes back empty the type is marked read-only instead of no-access.
#     A user who holds Global Reader plus an Intune role that grants write keeps
#     that write (it comes from the response); the floor only ever adds read.
#   * Otherwise one GET of deviceManagement/getEffectivePermissions per token.
#     The answer is cached under the token FINGERPRINT (tid|oid|iat): any newly
#     minted token - routine renewal, explicit login, Force refresh in the
#     Profile popup - misses the cache and re-asks. There is no TTL on purpose:
#     Force refresh is the one refresh path, exactly as for the token itself.
#   * Plus one GET of deviceManagement/resourceOperations: the CATALOGUE of
#     resource actions that exist. getEffectivePermissions answers with the
#     allowed actions only - notAllowedResourceActions comes back empty - so on
#     its own it cannot tell "the role denies this action" from "no such action
#     for this category", and both a denied write and a category that has no
#     Assign action look identical. Without the catalogue the write check was a
#     contradiction and every mapped type came out Full, which is the whole
#     point of Layer 2 missing. The catalogue is tenant-level and
#     token-independent, so it is cached per tenant and survives a refresh.
#   * Layer 2 can only DOWNGRADE. A missing scope is fatal regardless of RBAC.
#   * Unknown never colours. A failed call, a type with no category mapping,
#     an API that Intune RBAC does not govern, or a category the response never
#     mentions all return $null and Layer 1 stands. A sea of red on data we
#     could not read is worse than no signal.
#
# Scope tags are NOT modelled. getEffectivePermissions is the global answer; a
# user scoped to a subset of tags can still 403 on specific objects.

# Directory role template ids that imply full Intune RBAC. Template ids are
# fixed across tenants (wids carries the template id, not the tenant's role
# object id).
$script:RbacFullAccessRoleTemplateIds = @(
    "3a2c62db-5318-420d-8d74-23affee5d9d5",   # Intune Administrator
    "62e90394-69f5-4237-9190-012177145e10"    # Global Administrator
)

# Directory role template ids that imply tenant-wide READ but no write. Global
# Reader reads every Intune resource; it is a floor on read, not a replacement
# for the per-action lookup (see the header). Kept separate from the full-access
# list so it downgrades a writable type to read-only rather than granting Full.
$script:RbacReadOnlyRoleTemplateIds = @(
    "f2ef992c-3afb-46b9-b7cf-a126ee74c451"    # Global Reader
)

# The Graph function and the scope it is asked for. Both are what the Intune
# portal uses. Kept as variables so a live finding is a one-line change.
$script:RbacEffectivePermissionsUrl = "/deviceManagement/getEffectivePermissions(scope='*')"
$script:RbacResourceActionPrefix     = "Microsoft.Intune_"

# The action catalogue: every resource action Intune defines, one entry per
# action, with .id exactly '<prefix><Category>_<Action>'. Tenant-level and the
# same for every caller, so it is cached per tenant rather than per token.
$script:RbacResourceOperationsUrl = "/deviceManagement/resourceOperations"

# Actions that make a type writable. Only the ones that EXIST for a category
# count - Roles has no Assign, ManagedGooglePlay uses Modify - so a category
# with fewer write actions is not penalised for actions it lacks. Existence
# comes from the catalogue, which is why Layer 2 fetches it.
$script:RbacWriteActions = @("Create", "Update", "Delete", "Assign", "Modify")

# _API path -> Intune RBAC resource category (the middle segment of
# Microsoft.Intune_<Category>_<Action>). Longest prefix wins so a specific
# sub-path can override its parent. An entry may also override the action
# suffixes when a category does not use the plain Read/Create/... names.
#
# Every Category name below exists in the live resourceOperations catalogue, and
# a test asserts that against the catalogue fixture - a typo like the former
# 'Filters' (the portal's label; the resource is 'AssignmentFilter') now fails
# the suite instead of silently marking the type Unknown. UNVERIFIED marks the
# other half, the API -> category association, which the catalogue cannot
# confirm: it is safe because an action name the catalogue does not know yields
# Unknown, not a colour. Types can override any of this with _ResourceCategory.
$script:RbacApiCategoryMap = @{
    # Device configuration family
    "deviceManagement/deviceConfigurations"                = "DeviceConfigurations"
    "deviceManagement/configurationPolicies"               = "DeviceConfigurations"
    "deviceManagement/configurationPolicyTemplates"        = "DeviceConfigurations"
    "deviceManagement/groupPolicyConfigurations"           = "DeviceConfigurations"
    "deviceManagement/groupPolicyUploadedDefinitionFiles"  = "DeviceConfigurations"
    "deviceManagement/reusablePolicySettings"              = "DeviceConfigurations"
    "deviceManagement/hardwareConfigurations"              = "DeviceConfigurations"
    "deviceManagement/deviceManagementScripts"             = "DeviceConfigurations"
    "deviceManagement/deviceShellScripts"                  = "DeviceConfigurations"
    "deviceManagement/deviceCustomAttributeShellScripts"   = "DeviceConfigurations"
    "deviceManagement/deviceHealthScripts"                 = "DeviceConfigurations"   # UNVERIFIED (remediations)
    "deviceManagement/windowsFeatureUpdateProfiles"        = "DeviceConfigurations"   # UNVERIFIED
    "deviceManagement/windowsQualityUpdateProfiles"        = "DeviceConfigurations"   # UNVERIFIED
    "deviceManagement/windowsQualityUpdatePolicies"        = "DeviceConfigurations"   # UNVERIFIED
    "deviceManagement/windowsDriverUpdateProfiles"         = "DeviceConfigurations"   # UNVERIFIED
    "deviceManagement/inventoryPolicies"                   = "DeviceConfigurations"   # UNVERIFIED
    "deviceManagement/deviceEnrollmentConfigurations"      = "DeviceConfigurations"   # UNVERIFIED (ESP / restrictions)
    # Security baselines are intents; the template catalogue sits beside them.
    "deviceManagement/intents"                             = "SecurityBaselines"
    "deviceManagement/templates"                           = "SecurityBaselines"
    # Compliance family (Intune spells the category 'Polices')
    "deviceManagement/deviceCompliancePolicies"            = "DeviceCompliancePolices"
    "deviceManagement/compliancePolicies"                  = "DeviceCompliancePolices"
    "deviceManagement/deviceComplianceScripts"             = "DeviceCompliancePolices"  # UNVERIFIED
    "deviceManagement/notificationMessageTemplates"        = "DeviceCompliancePolices"  # UNVERIFIED
    # Tenant administration
    "deviceManagement/roleDefinitions"                     = "Roles"
    "deviceManagement/roleScopeTags"                       = "Roles"
    "deviceManagement/termsAndConditions"                  = "TermsAndConditions"
    "deviceManagement/assignmentFilters"                   = "AssignmentFilter"      # 'Filters' is the portal label, not the resource
    "deviceManagement/operationApprovalPolicies"           = @{
        Category = "MultiAdminApproval"
        Actions  = @{ Read = "ReadAccessPolicy"; Create = "CreateAccessPolicy"; Update = "UpdateAccessPolicy"; Delete = "DeleteAccessPolicy" }
    }
    "deviceManagement/intuneBrandingProfiles"              = "Customization"
    "deviceManagement/settings"                            = "Organization"           # UNVERIFIED
    # Enrollment
    "deviceManagement/appleUserInitiatedEnrollmentProfiles" = "AppleEnrollmentProfiles"  # UNVERIFIED
    "deviceManagement/depOnboardingSettings"               = "AppleEnrollmentProfiles"   # UNVERIFIED
    # Android Enterprise ('AndroidSync'). There is no Create/Delete action for a
    # profile: the one write action is UpdateEnrollmentProfiles, so it has to be
    # named here or the type would look writable to anyone who can read.
    "deviceManagement/androidForWorkEnrollmentProfiles"    = @{
        Category = "AndroidSync"
        Actions  = @{ Update = "UpdateEnrollmentProfiles" }
    }
    "deviceManagement/androidDeviceOwnerEnrollmentProfiles" = @{
        Category = "AndroidSync"
        Actions  = @{ Update = "UpdateEnrollmentProfiles" }
    }
    "deviceManagement/androidManagedStoreAccountEnterpriseSettings" = "ManagedGooglePlay" # UNVERIFIED
    # Autopilot profiles are governed by the Enrollment programs permission,
    # whose *profile* actions are the AppleEnrollmentProfiles resource - the
    # action names there are generic ("Read profile", "Assign profile") and the
    # catalogue has no Windows- or Autopilot-specific resource at all. The
    # alternative reading, EnrollmentProgramToken (the *token* actions), is the
    # same portal permission group and every built-in role grants the two sets
    # together, so the verdict is identical either way; only a custom role that
    # ticks tokens without profiles could tell them apart.
    "deviceManagement/windowsAutopilotDeploymentProfiles"  = "AppleEnrollmentProfiles"
    # Apps
    "deviceAppManagement/mobileApps"                       = "MobileApps"
    "deviceAppManagement/mobileAppConfigurations"          = "MobileApps"              # UNVERIFIED
    "deviceAppManagement/iosLobAppProvisioningConfigurations" = "MobileApps"           # UNVERIFIED
    "deviceAppManagement/managedAppPolicies"               = "ManagedApps"
    "deviceAppManagement/targetedManagedAppConfigurations" = "ManagedApps"
    "deviceAppManagement/policySets"                       = "PolicySets"
}

# Intune-governed APIs that deliberately have NO category. The completeness test
# (Tests/EffectivePermissions.Tests.ps1) fails on any deviceManagement/ or
# deviceAppManagement/ type that is in neither table, so a new type cannot fall
# through silently. Value = the reason.
$script:RbacUnmappedApis = @{
    "deviceManagement/deviceCategories"          = "no matching resource in the action catalogue"
    "deviceManagement/virtualEndpoint"           = "Windows 365 uses its own RBAC namespace"
    "deviceAppManagement/vppTokens"              = "no VPP resource in the action catalogue (MicrosoftStoreForBusiness is a different store)"
}

#region Category resolution

# The category entry for a policy type: @{ Category; Actions } or $null when the
# type is not Intune-governed / unmapped. _ResourceCategory on the type wins.
function Get-PolicyTypeRbacCategory {
    [CmdletBinding()]
    param($PolicyType)

    if(-not $PolicyType) { return $null }

    $override = $null
    try { $override = $PolicyType._ResourceCategory } catch { }
    if($override) { return @{ Category = [string]$override; Actions = @{} } }

    $api = $null
    try { $api = [string]$PolicyType._API } catch { }
    return (Get-RbacCategoryForApi $api)
}

function Get-RbacCategoryForApi {
    [CmdletBinding()]
    param([string]$Api)

    if([string]::IsNullOrWhiteSpace($Api)) { return $null }
    $api = $Api.Trim().TrimStart('/')

    # Longest matching prefix wins; a prefix must end at a path boundary.
    # "$($key)?" not "$key?": on PS7 '?' is a legal variable-name character, so
    # "$key?" reads the (empty) variable 'key?' and StartsWith("") matches all.
    $best = $null
    foreach($key in $script:RbacApiCategoryMap.Keys) {
        if($api -eq $key -or $api.StartsWith("$key/", [System.StringComparison]::OrdinalIgnoreCase) -or
           $api.StartsWith("$($key)?", [System.StringComparison]::OrdinalIgnoreCase)) {
            if(-not $best -or $key.Length -gt $best.Length) { $best = $key }
        }
    }
    if(-not $best) { return $null }

    $entry = $script:RbacApiCategoryMap[$best]
    if($entry -is [string]) { return @{ Category = $entry; Actions = @{} } }
    $actions = if($entry.Actions) { $entry.Actions } else { @{} }
    return @{ Category = [string]$entry.Category; Actions = $actions }
}

# Full resource action name for a category + logical action, honouring the
# entry's suffix overrides.
function Get-RbacActionName {
    [CmdletBinding()]
    param($Entry, [string]$Action)

    $suffix = $Action
    if($Entry.Actions -and $Entry.Actions.ContainsKey($Action)) { $suffix = $Entry.Actions[$Action] }
    return "$($script:RbacResourceActionPrefix)$($Entry.Category)_$suffix"
}

#endregion

#region Response parsing

# Turn a getEffectivePermissions payload into two sets: the actions the role
# allows, and the ones it explicitly denies. The shape is walked defensively -
# value[].resourceActions[].allowed/notAllowed - and a bare object with
# resourceActions is accepted too.
#
# NotAllowed is empty in practice: Graph answers with the allowed list only. It
# is still read, because when it is populated it is a better source for "this
# action exists" than the catalogue (Get-RbacActionCatalog falls back to it).
function ConvertTo-RbacActionSets {
    [CmdletBinding()]
    param($Response)

    $allowed    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $notAllowed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if(-not $Response) { return @{ Allowed = $allowed; NotAllowed = $notAllowed } }

    $items = @()
    if($Response.PSObject.Properties['value']) { $items = @($Response.value) } else { $items = @($Response) }

    foreach($item in $items) {
        if(-not $item) { continue }
        $resourceActions = @()
        if($item.PSObject.Properties['resourceActions']) { $resourceActions = @($item.resourceActions) }
        elseif($item.PSObject.Properties['allowedResourceActions']) { $resourceActions = @($item) }
        foreach($ra in $resourceActions) {
            if(-not $ra) { continue }
            foreach($a in @($ra.allowedResourceActions)) {
                if($a) { [void]$allowed.Add(([string]$a).Trim()) }
            }
            foreach($a in @($ra.notAllowedResourceActions)) {
                if($a) { [void]$notAllowed.Add(([string]$a).Trim()) }
            }
        }
    }
    return @{ Allowed = $allowed; NotAllowed = $notAllowed }
}

#endregion

#region Action catalogue

function Get-RbacCatalogCacheName {
    param([string]$TenantId)
    return "RbacResourceOperations_$TenantId"
}

# The set of resource action names that EXIST, from deviceManagement/
# resourceOperations. $null when the call fails or returns nothing usable -
# callers must treat that as "existence unknown", not as "nothing exists".
#
# Cached per tenant with no timeout and no token fingerprint: the catalogue is a
# property of the Intune service, not of the signed-in user, so a token refresh
# re-asks getEffectivePermissions but reuses this.
function Get-RbacActionCatalog {
    [CmdletBinding()]
    param([string]$TenantId, [int]$TokenId = 0)

    # The cached value is a hashtable wrapping the set, and every return of the
    # set has a leading comma. PowerShell enumerates a HashSet when it is written
    # to the pipeline - both here and inside Get-CacheObject - so without those
    # two guards the caller gets an object[] of names instead of the set, and
    # object[].Contains() is case-SENSITIVE. A hashtable is not enumerated.
    $cacheName = Get-RbacCatalogCacheName $TenantId
    $cached = Get-CacheObject $cacheName
    if($cached -and $cached.Actions.Count -gt 0) { return ,$cached.Actions }

    $response = $null
    try {
        $params = @{ Url = $script:RbacResourceOperationsUrl; GraphVersion = "beta"; ODataMetadata = "minimal"; NoError = $true; AllPages = $true }
        if($TokenId -gt 0) { $params.TokenId = $TokenId }
        $response = Invoke-MSGraphAPI @params
    }
    catch { }

    $actions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach($op in @($response.value)) {
        if($op -and $op.id) { [void]$actions.Add(([string]$op.id).Trim()) }
    }
    if($actions.Count -eq 0) {
        Write-Log ("Intune RBAC marking: the resource action catalogue ($($script:RbacResourceOperationsUrl)) could not be read. " +
                   "A type is marked read-only only when the role allows no write action at all for its category.") 2
        return $null
    }

    Set-CacheObject $cacheName @{ Actions = $actions } "TenantCache_$TenantId" -Persistent
    return ,$actions
}

#endregion

#region Context (fetch + cache)

function Get-RbacTokenFingerprint {
    [CmdletBinding()]
    param($Claims)
    if(-not $Claims) { return $null }
    return "$($Claims.tid)|$($Claims.oid)|$($Claims.iat)"
}

function Test-RbacAppOnlyClaims {
    [CmdletBinding()]
    param($Claims)
    if(-not $Claims) { return $true }
    if($Claims.PSObject.Properties['idtyp'] -and ([string]$Claims.idtyp) -eq 'app') { return $true }
    # No scp at all means no delegated consent - app-only or unreadable.
    return (-not $Claims.PSObject.Properties['scp'] -or -not $Claims.scp)
}

function Test-RbacFullAccessRole {
    [CmdletBinding()]
    param($Claims)
    if(-not $Claims -or -not $Claims.PSObject.Properties['wids']) { return $false }
    foreach($w in @($Claims.wids)) {
        if($w -and $script:RbacFullAccessRoleTemplateIds -contains ([string]$w).Trim()) { return $true }
    }
    return $false
}

function Test-RbacReadOnlyRole {
    [CmdletBinding()]
    param($Claims)
    if(-not $Claims -or -not $Claims.PSObject.Properties['wids']) { return $false }
    foreach($w in @($Claims.wids)) {
        if($w -and $script:RbacReadOnlyRoleTemplateIds -contains ([string]$w).Trim()) { return $true }
    }
    return $false
}

function Test-RbacAccessMarkingEnabled {
    [CmdletBinding()]
    param()
    $v = $null
    try { $v = Get-SettingValue "UseRbacAccessMarking" } catch { }
    # Compare as text: `$false -eq ""` is TRUE in PowerShell (the right side is
    # coerced to bool), which would read an explicit off as "unset".
    $text = [string]$v
    if($text -eq "") { return $true }
    return ($text -eq "true")
}

function Get-RbacCacheName {
    param([string]$TenantId)
    return "EffectivePermissions_$TenantId"
}

# The Layer 2 context for the current (or given) token, or $null when Layer 2
# does not apply. Never prompts and never throws: every failure path returns
# $null after one log line, and a failed fetch is remembered for the token so a
# nav rebuild does not re-ask until the token changes.
#
# Context shape:
#   Fingerprint, TenantId, AsOf (token iat), Source ('DirectoryRole'|'Graph'),
#   AllAllowed (bool), Allowed / NotAllowed / Catalog (HashSet[string], Catalog
#   $null when the catalogue could not be read), Raw (response)
function Get-IntuneRbacContext {
    [CmdletBinding()]
    param($Claims, [int]$TokenId = 0, [switch]$IgnoreSetting)

    if(-not $IgnoreSetting -and -not (Test-RbacAccessMarkingEnabled)) { return $null }

    if(-not $Claims) { $Claims = Get-AccessTokenClaims }
    if(-not $Claims) { return $null }
    if(Test-RbacAppOnlyClaims $Claims) {
        Write-LogDebug "Intune RBAC marking: app-only token, application permissions are the effective set"
        return $null
    }

    $fingerprint = Get-RbacTokenFingerprint $Claims
    $tenantId    = [string]$Claims.tid
    $asOf        = $null
    try { if($Claims.iat) { $asOf = [datetime]::new(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc).AddSeconds([double]$Claims.iat).ToLocalTime() } } catch { }

    $cacheName = Get-RbacCacheName $tenantId
    $cached = Get-CacheObject $cacheName
    if($cached -and $cached.Fingerprint -eq $fingerprint) {
        if($cached.Failed) { return $null }
        return $cached
    }

    $ctx = [PSCustomObject]@{
        Fingerprint = $fingerprint
        TenantId    = $tenantId
        AsOf        = $asOf
        Source      = $null
        AllAllowed  = $false
        AllRead     = $false
        Allowed     = $null
        NotAllowed  = $null
        Catalog     = $null
        Raw         = $null
        Failed      = $false
        Error       = $null
    }

    if(Test-RbacFullAccessRole $Claims) {
        $ctx.Source     = "DirectoryRole"
        $ctx.AllAllowed = $true
        Write-Log "Intune RBAC marking: Intune Administrator / Global Administrator role in token; full Intune access assumed, no permission lookup"
    }
    else {
        # Global Reader (or equivalent) guarantees read everywhere. Recorded now
        # so it applies whether or not the per-action lookup below succeeds.
        $ctx.AllRead = (Test-RbacReadOnlyRole $Claims)
        $response = $null
        try {
            $params = @{ Url = $script:RbacEffectivePermissionsUrl; GraphVersion = "beta"; ODataMetadata = "minimal"; NoError = $true }
            if($TokenId -gt 0) { $params.TokenId = $TokenId }
            $response = Invoke-MSGraphAPI @params
        }
        catch {
            $ctx.Error = $_.Exception.Message
        }

        $sets = ConvertTo-RbacActionSets $response
        if(-not $response -or ($sets.Allowed.Count -eq 0 -and $sets.NotAllowed.Count -eq 0)) {
            if($ctx.AllRead) {
                # A pure Global Reader with no Intune RBAC assignment: the lookup
                # says nothing, but the directory role still guarantees read. Mark
                # from the role instead of failing, so the type shows read-only.
                $ctx.Source = "DirectoryRole"
                Write-Log "Intune RBAC marking: Global Reader directory role in token; read-only Intune access assumed (getEffectivePermissions returned nothing)"
            }
            else {
                $ctx.Failed = $true
                if(-not $ctx.Error) { $ctx.Error = "empty or unreadable response" }
                Write-Log ("Intune RBAC marking unavailable: getEffectivePermissions returned nothing usable ($($ctx.Error)). " +
                           "The navigation reflects token scopes only until the next token refresh.") 2
            }
        }
        else {
            $ctx.Source     = "Graph"
            $ctx.Allowed    = $sets.Allowed
            $ctx.NotAllowed = $sets.NotAllowed
            $ctx.Raw        = $response
            # A populated notAllowed list would be the authoritative statement of
            # what exists; the catalogue is what makes the verdict possible when
            # it is empty, which is every response seen so far.
            $ctx.Catalog = Get-RbacActionCatalog -TenantId $tenantId -TokenId $TokenId
            if(-not $ctx.Catalog -and $sets.NotAllowed.Count -gt 0) {
                $ctx.Catalog = [System.Collections.Generic.HashSet[string]]::new($sets.Allowed, [System.StringComparer]::OrdinalIgnoreCase)
                foreach($a in $sets.NotAllowed) { [void]$ctx.Catalog.Add($a) }
            }
            $of = if($ctx.Catalog) { " of $($ctx.Catalog.Count)" } else { "" }
            Write-Log "Intune RBAC marking: $($sets.Allowed.Count)$of Intune resource actions allowed for the signed-in user"
        }
    }

    # Persistent + tenant tag: no timeout (the fingerprint is the invalidation)
    # and Clear-TenantCache sweeps it with the other per-tenant entries.
    Set-CacheObject $cacheName $ctx "TenantCache_$tenantId" -Persistent
    if($ctx.Failed) { return $null }
    return $ctx
}

function Clear-IntuneRbacContext {
    [CmdletBinding()]
    param([string]$TenantId)

    if($TenantId) {
        Clear-CacheObject -Name (Get-RbacCacheName $TenantId)
        Clear-CacheObject -Name (Get-RbacCatalogCacheName $TenantId)
        return
    }
    foreach($name in @($script:cacheObjects.Keys | Where-Object { $_ -like "EffectivePermissions_*" -or $_ -like "RbacResourceOperations_*" })) {
        Clear-CacheObject -Name $name
    }
}

function Invoke-RbacEventUserDisconnected {
    param($Snapshot)
    $tenantId = $null
    try { $tenantId = [string]$Snapshot.TenantId } catch { }
    Clear-IntuneRbacContext -TenantId $tenantId
}

#endregion

#region Per-type evaluation

# True when the type declares at least one ReadWrite permission, i.e. it is a
# writable feature rather than read-only by design. Shared by the read-guarantee
# and write-check paths.
function Test-PolicyTypeDeclaresWrite {
    [CmdletBinding()]
    param($PolicyType)
    foreach($perm in @($PolicyType._Permissions | Where-Object { $_ })) {
        if(Get-PermissionReadVariant $perm) { return $true }
    }
    return $false
}

# Layer 2 verdict for one policy type against a context. Returns
# @{ Level = [APIAccess]; Info = <tooltip text>; Missing = @(actions) } or $null
# for Unknown (no category, category never mentioned, no context).
function Get-PolicyTypeRbacAccess {
    [CmdletBinding()]
    param($PolicyType, $Context)

    if(-not $PolicyType -or -not $Context) { return $null }
    $entry = Get-PolicyTypeRbacCategory $PolicyType
    if(-not $entry) { return $null }

    if($Context.AllAllowed) { return @{ Level = [APIAccess]::Full; Info = ""; Missing = @() } }

    # A directory read-only role (Global Reader) with no per-action data: read is
    # guaranteed, write is not. Read-only-by-design types are Full; the rest are
    # read-only. Handled before the Allowed check so a pure Global Reader whose
    # getEffectivePermissions came back empty is still marked.
    if($Context.AllRead -and -not $Context.Allowed) {
        if(-not (Test-PolicyTypeDeclaresWrite $PolicyType)) { return @{ Level = [APIAccess]::Full; Info = ""; Missing = @() } }
        return @{
            Level   = [APIAccess]::Limited
            Info    = "Intune role: read-only for $($entry.Category) (Global Reader directory role grants read, not write)"
            Missing = @()
        }
    }
    if(-not $Context.Allowed) { return $null }

    # The catalogue says which actions exist. Without it (the call failed) every
    # existence question is unanswerable, and the verdict falls back to the one
    # thing the allowed list alone can prove - see the write check below.
    $catalog    = $Context.Catalog
    $readAction = Get-RbacActionName $entry "Read"

    if($catalog -and -not $catalog.Contains($readAction)) {
        # Intune has no such action: this row of the category table is wrong, not
        # the user's role. Say nothing rather than something false.
        return $null
    }

    # AllRead (Global Reader) guarantees read even when the response omits this
    # category's read action.
    if(-not $Context.AllRead -and -not $Context.Allowed.Contains($readAction)) {
        if(-not $catalog) {
            # Could be a denied read or a wrong mapping. Unknown.
            return $null
        }
        return @{
            Level   = [APIAccess]::None
            Info    = "Intune role: no read access to $($entry.Category) (missing $readAction)"
            Missing = @($readAction)
        }
    }

    # Read-only by design (declares only Read scopes) needs nothing more.
    if(-not (Test-PolicyTypeDeclaresWrite $PolicyType)) { return @{ Level = [APIAccess]::Full; Info = ""; Missing = @() } }

    $missing = @()
    $exists  = 0
    foreach($action in $script:RbacWriteActions) {
        $name = Get-RbacActionName $entry $action
        if($catalog -and -not $catalog.Contains($name)) { continue }   # no such action for this category
        $exists++
        if(-not $Context.Allowed.Contains($name)) { $missing += $name }
    }
    if(-not $catalog) {
        # Naming individual missing actions would be inventing them, so the only
        # honest degraded verdict is the coarse one: a role that allows no write
        # action at all for the category cannot write it.
        if($missing.Count -lt $exists) { return @{ Level = [APIAccess]::Full; Info = ""; Missing = @() } }
        return @{
            Level   = [APIAccess]::Limited
            Info    = "Intune role: read-only for $($entry.Category) (no write action allowed)"
            Missing = @()
        }
    }
    if($missing.Count -eq 0) { return @{ Level = [APIAccess]::Full; Info = ""; Missing = @() } }

    # "read-only" is only true when the role allows no write action at all. A role
    # that can Create and Update but not Delete or Assign - a common custom role - is
    # still Limited, but calling that read-only misdescribes what the user can do.
    $short = @($missing | ForEach-Object { $_.Substring($_.LastIndexOf('_') + 1) })
    $info = if($missing.Count -ge $exists) {
        "Intune role: read-only for $($entry.Category) (missing $($short -join ', '))"
    }
    else {
        "Intune role: partial write access to $($entry.Category) (missing $($short -join ', '))"
    }
    return @{
        Level   = [APIAccess]::Limited
        Info    = $info
        Missing = $missing
        # Carried as data, not only as tooltip text, so the Permissions popup's
        # Role/Effective/Result columns can say "Partial write" instead of
        # contradicting the tooltip with "Read" and "Read-only".
        Partial = ($missing.Count -lt $exists)
    }
}

# The worse of two levels. None outranks Limited outranks Full.
function Get-WorstAccessLevel {
    [CmdletBinding()]
    [OutputType([APIAccess])]
    param([APIAccess]$A, [APIAccess]$B)
    if($A -eq [APIAccess]::None -or $B -eq [APIAccess]::None) { return [APIAccess]::None }
    if($A -eq [APIAccess]::Limited -or $B -eq [APIAccess]::Limited) { return [APIAccess]::Limited }
    return [APIAccess]::Full
}

#endregion

function Invoke-RbacEventAppInitialized {
    Add-SettingsObject -Title "Use Intune role permissions for access marking" -Key "UseRbacAccessMarking" -Type "Boolean" `
        -Description "Also ask Intune which resource actions the signed-in user's role allows, and mark menu items the user cannot change (orange) or read (red). Off = mark from the app's token scopes only. Refresh the token from the Profile popup after a role change." `
        -DefaultValue $true -Section "General"
}

Add-AppEventHandler "AppInitialized" "Invoke-RbacEventAppInitialized"
Add-AppEventHandler "AuthenticationUserDisconnected" "Invoke-RbacEventUserDisconnected"
