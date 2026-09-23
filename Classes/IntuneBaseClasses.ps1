enum APIAccess
{
    Full
    Limited
    None
}

enum OMetadata
{
    Full
    Minimal
    None
    Skip
}

# A group of APIs eg Device Configurations has Settings Catalog, Templates etc.
class IntunePolicyGroupBase 
{
    Hidden [string]$_ID = ""
    Hidden [string]$_Name = ""
    Hidden [Object]$_Icon = $null
    Hidden [IntunePolicyTypeBase[]]$_PolicyTypes = @()
    # Permission state for the left-nav colour, stamped by
    # Update-IntuneAccessLevels (Internal/AccessLevel.ps1). Full renders
    # normally, Limited orange, None red. Groups aggregate their member types:
    # None only when nothing in the group is usable. AccessInfo is the tooltip
    # breakdown. Both stay at their defaults when there is no token to diff
    # against, so an unreadable token shows no colour rather than all red.
    [APIAccess]$AccessType = [APIAccess]::Full
    [string]$AccessInfo = ""
    Hidden [Object]$IconImage = $null
    # Column list for the group view. $null = derive from the member types
    # (Get-IntuneDefaultColumns, Internal/IntuneViewColumns.ps1); set only to
    # hard-override the derivation.
    Hidden [string[]]$_ViewProperties = $null
    # Columns the whole group can fill, appended after Policy type - e.g.
    # "ScriptType=Script type". Keep to values every member type supplies.
    Hidden [string[]]$_ExtraColumns = @()
    # $false hides the Policy type column, $true forces it; $null = show it when
    # the group has more than one member type.
    Hidden [object]$_ShowPolicyTypeColumn = $null
    # $false hides Platform, $true forces it even when most members cannot fill it
    # (Tenant administration: assignment filters are worth the blanks); $null = the
    # half-the-members rule.
    Hidden [object]$_ShowPlatformColumn = $null
    # $null = all buttons / all bulk operations. Set to e.g. @("Export","View")
    # for read-only groups (UI button gating + bulk-form type lists both check it).
    Hidden [object]$_ShowButtons = $null

    IntunePolicyGroupBase()
    {
        if($this.GetType().Name -eq "IntunePolicyGroupBase") {
            throw "Abstract class. Object cannot be created"
        }
        elseif($script:SingletonObjects.ContainsKey($this.GetType().Name) -eq $true) {
            throw "Only one $($this.GetType().Name) object can be created"
        }

        Add-SingletonObject $this.GetType().Name $this

        ([IntunePolicyGroupBase]$this).Init()
    }

    # Hidden Functions
    Hidden Init()
    {
        Add-ObjectProperty $this "ID" { $this._ID }
        Add-ObjectProperty $this "Title" { $this._Name }
        Add-ObjectProperty $this "Icon" { Get-StringOrDefault $this._Icon $this._ID }
        #Add-ObjectProperty $this "IconImage" { return $this._IconImage }
        Add-ObjectProperty $this "PolicyTypes" { $this._PolicyTypes }
        Add-ObjectProperty $this "ViewProperties" { ?? $this._ViewProperties (Get-IntuneDefaultColumns $this) }
        Add-ObjectProperty $this "ShowButtons" { $this._ShowButtons }
    }

    # Public Functions
    [void]AddPolicyType([IntunePolicyTypeBase]$PolicyType)
    {
        if(-not $PolicyType) { return }

        if($null -eq ($this._PolicyTypes | Where-Object ID -eq $PolicyType.ID)) {
            $this._PolicyTypes += $PolicyType
        }
        else {
            Write-Log "Cannot add PolicyType to PolicyGroup. $($PolicyType.ID) ($($PolicyType.Title)) alredy added" 2
        }
    }
}

# Represents an API
class IntunePolicyTypeBase
{
    Hidden [IntunePolicyGroupBase]$_PolicyGroup = $null
    Hidden [String]$_Name = $null

    # See IntunePolicyGroupBase.AccessType. For a type this is derived straight
    # from _Permissions vs the token's granted scopes/roles.
    [APIAccess]$AccessType = [APIAccess]::Full
    [string]$AccessInfo = ""

    Hidden [String]$_ID = $null
    Hidden [String]$_APITitle = $null
    Hidden [String]$_APIVersion = $null
    Hidden [String]$_API = $null
    Hidden [String]$_APIPOST = $null
    Hidden [String]$_APIPUT = $null
    Hidden [String]$_APIDELETE = $null
    Hidden [String]$_PolicyName = $null
    Hidden [String]$_PolicyBaseName = $null
    Hidden [String[]]$_PropertiesToRemove = $null
    Hidden [String[]]$_PropertiesToRemoveForUpdate = $null
    # Properties the PROPERTY compare should not match on for this type -
    # server-maintained state that legitimately differs between two
    # identically-configured objects (rendered greyed-out, like the global
    # isAssigned skip). Membership-style props skipped here should still be
    # covered semantically by the type's documentation compare.
    Hidden [String[]]$_ComparePropertiesToSkip = @()
    Hidden [String[]]$_SkipRemoveProperties = @()
    Hidden [String[]]$_Permissions = $null
    # Intune RBAC resource category (the middle of Microsoft.Intune_<Cat>_<Action>)
    # used by the access marking's Layer 2 (Internal/EffectivePermissions.ps1).
    # $null = derive from _API; set only where the derivation is wrong.
    Hidden [String]$_ResourceCategory = $null
    # Column list for this type's view. $null = derive (Get-IntuneDefaultColumns,
    # Internal/IntuneViewColumns.ps1) from the flags below; set only to
    # hard-override the derivation.
    Hidden [string[]]$_ViewProperties = $null
    # Rows of this type resolve a Platform, per row or via _PlatformName.
    Hidden [bool]$_HasPlatform = $true
    # The endpoint returns lastModifiedDateTime / modifiedDateTime.
    Hidden [bool]$_HasModified = $true
    # Column that tells rows of this ONE type apart - "PolicyName=Type" where the
    # object class sets a per-row name, "ApplicationType=Type" for apps. It takes
    # the place of the generic Policy type column, which would repeat itself.
    Hidden [string]$_SubTypeColumn = $null
    # Type-specific additions, e.g. "ApplicationTypeGroup=App type".
    Hidden [string[]]$_ExtraColumns = @()
    Hidden [String]$_Expand = $null
    Hidden [String]$_Icon = $null
    Hidden [Object]$IconImage = $null
    Hidden [String[]]$_Dependencies = $null

    Hidden [String]$_QueryList = $null
    Hidden [Boolean]$_QuerySearch = $false
    Hidden [Boolean]$_NavigationProperties = $false

    Hidden [Boolean]$_ExpandAssignments = $true
    Hidden [Boolean]$_ExpandAssignmentsList = $true
    Hidden [Boolean]$_SupportsAssignments = $true
    Hidden [String]$_AssignmentsType = "assignments"
    # Action segment used when writing assignments ({API}/{id}/<action>).
    # Most types use the canonical "assign" action; policySets have no
    # /assign segment - their update action takes a full `assignments`
    # replacement list with the same semantics.
    Hidden [String]$_AssignAction = "assign"
    # When $true, per-policy assignments are fetched via
    # {API}/{id}?$expand=assignments instead of the {API}/{id}/assignments
    # navigation URL (policySets: the navigation GET returns 400, only the
    # expand variant works - same family as _ExpandAssignmentsList).
    Hidden [Boolean]$_AssignmentsViaExpand = $false
    # Optional override consulted by the Bulk Assignments tool. When the
    # default Get-BulkAssignmentObjectType heuristic (AssignmentsType / Id
    # lookup) can't derive a Graph @odata.type for this PolicyType's
    # assignment entries, set this field to the explicit
    # "#microsoft.graph.XAssignment" string. The lookup logs a warning when
    # both this override and the heuristic return null — that's the signal
    # a new type needs to set it.
    Hidden [String]$_AssignmentObjectType = $null
    Hidden [String[]]$_AssignmentPropertiesToKeep = $null
    Hidden [String[]]$_AssignmentTargetPropertiesToKeep = $null

    Hidden [Uint32]$_ImportOrder = 1000
    Hidden [OMetadata]$_ODataMetadata = [OMetadata]::Full
    Hidden [Boolean]$_SkipAddIDOnFileName = $false
    Hidden [Boolean]$_ScopeTagsReturnedInList = $true
    Hidden [String]$_ScopeTagProperty = "roleScopeTagIds"
    Hidden [String]$_ObjectClass = $null

    Hidden [String]$_CopyDefaultName = $null

    Hidden [Boolean]$_HasDescription = $true
    Hidden [Boolean]$_RequiresFullObject = $true

    Hidden [Boolean]$_VerifyObject = $false

    # True when this type's CheckPolicy authoritatively matches/rejects a *file* object
    # by @odata.type (a complete matcher), rather than relying on the base @odata.id
    # matcher which is blind to files. When true, a CheckPolicy=$false is a real
    # rejection, so the single-narrowed-candidate folder-trust fallback in
    # Get-PoliciesTypeFromObject is skipped. Leave $false for types whose CheckPolicy
    # cannot reliably identify their own file objects (the base @odata.id matcher, or
    # ApplicationType's assignments-context heuristic).
    Hidden [Boolean]$_StrictODataTypeCheck = $false

    Hidden [String]$_NameProperty = "displayName"
    # Whether this type's endpoint honours a server-side name filter
    # (contains/tolower on _NameProperty). Verified live for most Intune
    # endpoints; set to $false on the ones that answer HTTP 400/500 so the
    # search filters client-side instead. See GetNameFilterClause.
    Hidden [Boolean]$_SupportsNameFilter = $true
    Hidden [String]$_IdProperty = "id"

    Hidden [String]$_Folder = $null

    # Optional default platform name for this policy type. When set, the per-row
    # IntunePolicyBase.Platform getter falls back to this value if the row didn't
    # set its own _PlatformName. Used by DeviceEnrollment subtypes (ESP, WHfB,
    # Co-Mgmt, Windows Restore) so the platform doesn't have to be derived from
    # @odata.type on every Object instance.
    Hidden [String]$_PlatformName = $null

    Hidden [String[]]$_PolicyFileAttributes = @()

    Hidden [Uint32]$_PolicyTypeOrder = 100

    # $null = all buttons. Set to e.g. @("Export","View") for read-only types
    # (the Intune Info group). Gates the main-view buttons and bulk-form lists.
    Hidden [object]$_ShowButtons = $null

    # Set when the API returns ONE object instead of a value collection (e.g.
    # androidManagedStoreAccountEnterpriseSettings, deviceManagement/settings).
    # Get-GraphPolicies wraps the body as a single row. Combine with
    # _HasPageSizeSupport = $false so no $top is appended to the list URL.
    Hidden [Boolean]$_SingleObject = $false
    Hidden [Uint32]$_TopItems = 1000

    Hidden [Boolean]$_HasPageSizeSupport = $true

    IntunePolicyTypeBase()
    {
        if($this.GetType().Name -eq "IntunePolicyTypeBase") {
            throw "Abstract class. Object cannot be created"
        }
        elseif($script:SingletonObjects.ContainsKey($this.GetType().Name) -eq $true) {
            throw "Only one $($this.GetType().Name) object can be created"
        }

        Add-SingletonObject $this.GetType().Name $this

        ([IntunePolicyTypeBase]$this).Init()
    }

    Static [PSCustomObject]Get([IntunePolicyBase]$GraphObject)
    {
        return $GraphObject.Object
    }

    [String]GetListURL()
    {
        return ([IntunePolicyTypeBase]$this).GetListURL($null)
    }

    [String]GetListURL([string]$NameFilter)
    {
        $params = @()
        $stringURI = $this._API
        if($this._QueryList) {
            $stringURI += $this._QueryList
        }

        # Server-side name search. _QueryList already carries a $filter on some
        # types (the Settings Catalog siblings filter by templateFamily), and a
        # URL may only have ONE $filter - so merge into the existing clause with
        # 'and' rather than appending a second parameter. The existing clause is
        # parenthesised because those are 'or' chains and 'and' binds tighter.
        $nameClause = ([IntunePolicyTypeBase]$this).GetNameFilterClause($NameFilter)
        if($nameClause) {
            $existingMatch = if($this._QueryList) { [Regex]::Match($this._QueryList, '\$filter=([^&]+)') } else { $null }
            if($existingMatch -and $existingMatch.Success) {
                # Splice by index - no regex replacement, so quotes and
                # parentheses in either clause are safe.
                $merged = "`$filter=($($existingMatch.Groups[1].Value)) and $nameClause"
                $stringURI = $this._API +
                             $this._QueryList.Substring(0, $existingMatch.Index) +
                             $merged +
                             $this._QueryList.Substring($existingMatch.Index + $existingMatch.Length)
            }
            else {
                $params += "`$filter=$nameClause"
            }
        }
        if((Get-SettingValue "ExpandAssignments") -eq $true -and $this.ExpandAssignmentsList -ne $false)
        {            # Expand assignments so they can be used in custom columns
            if(-not $this._QueryList -or ($this._QueryList.IndexOf('$expand=',[System.StringComparison]::InvariantCultureIgnoreCase)) -eq -1)
            {
                $params += "`$expand=assignments"
            }
        }

        # ToDo: Add PageSize Setting        
        $pageSizeStr = Get-SettingValue "GraphPageSize" $this._TopItems
        try {
            $pageSize = [uint32]$pageSizeStr
        }
        catch {
            $pageSize = 20
        }

        if($this.HasPageSizeSupport -eq $false)
        {
            # Do nothing...
        }
        elseif($pageSize -gt 0 -and (-not $this._QueryList -or ($this._QueryList.IndexOf('$top=',[System.StringComparison]::InvariantCultureIgnoreCase) -eq -1))) {
            $params += "`$top=$pageSize"
        }

        $url = $stringURI
        if($params.Count -gt 0) {
            $url += (?: ([String]::IsNullOrEmpty($this._QueryList)) "?" "&")
            $url += ($params -join "&")
        }

        # NOTE: %OrganizationId% is intentionally NOT substituted here. It used to be
        # replaced with $script:OrganizationId (the *default* tenant), which broke
        # cross-tenant lookups when the caller supplied -TokenId pointing at a different
        # token. Substitution now happens in Invoke-MSGraphAPI using the call's TokenId.

        return $url
    }

    # Base URL used when fetching a single policy's assignments ({base}/{id}/assignments
    # or {base}/{id}?$expand=assignments). Defaults to the list API. Types whose list
    # collection is polymorphic must override this: $expand and the assignments
    # navigation both resolve against the collection's *declared* type, so a base type
    # that doesn't declare `assignments` returns 400 for either form (App protection -
    # managedAppPolicies is Collection(managedAppPolicy), which has no navigation
    # properties at all). Return $null to skip the fetch for an object that cannot
    # carry assignments.
    [String]GetAssignmentsBaseURL([PSCustomObject]$PolicyObject)
    {
        return $this._API
    }

    [PSCustomObject]GetListBatchObject()
    {
        return ([IntunePolicyTypeBase]$this).GetListBatchObject($null)
    }

    [PSCustomObject]GetListBatchObject([string]$NameFilter)
    {
        return [PSCustomObject]@{
            id = $this.Id
            method = "GET"
            url = ([IntunePolicyTypeBase]$this).GetListURL($NameFilter).TrimStart('/')
            headers = @{"Accept"="application/json;odata.metadata=$($this._ODataMetadata)"}
        }
    }

    # Server-side name search. Returns the OData filter clause for $NameFilter,
    # or an empty string when this type cannot be searched that way.
    #
    # contains(tolower(prop),'term') rather than startswith(prop,'term'):
    # verified live 2026-08-27 that Graph's startswith is case-SENSITIVE on
    # nearly every one of these endpoints (a search for "test" found nothing
    # named "[Testing] ..."), while the tolower/contains form is accepted and
    # returns exactly what a client-side substring match would. It also keeps
    # the substring semantics the picker had before it searched server-side.
    #
    # Types whose endpoint rejects any $filter set _SupportsNameFilter = $false
    # and are filtered client-side by Get-GraphPolicies instead.
    [String]GetNameFilterClause([string]$NameFilter)
    {
        if([String]::IsNullOrWhiteSpace($NameFilter)) { return "" }
        if($this._SupportsNameFilter -ne $true) { return "" }
        # Single-object endpoints return the object itself, not a collection -
        # there is nothing to filter and Graph rejects the attempt.
        if($this._SingleObject -eq $true) { return "" }

        $prop = Get-StringOrDefault $this._NameProperty "displayName"
        # Escaping (quote doubling, percent-encoding, PS5.1/PS7 normalisation) is
        # ConvertTo-ODataStringLiteral's job - one rule for every literal.
        $literal = ConvertTo-ODataStringLiteral $NameFilter.ToLowerInvariant()
        return "contains(tolower($prop),$literal)"
    }

    # Hidden Functions
    Hidden Init()
    {    
        Add-ObjectProperty $this "ID" { $this._ID }
        Add-ObjectProperty $this "PolicyGroup" { $this._PolicyGroup }
        Add-ObjectProperty $this "APIVersion" { Get-StringOrDefault $this._APIVersion "Beta" }
        Add-ObjectProperty $this "API" { $this._API }
        Add-ObjectProperty $this "APIPOST" { ?? $this._APIPOST $this._API }
        Add-ObjectProperty $this "APIPUT" { ?? $this._APIPUT $this._API }
        Add-ObjectProperty $this "APIDELETE" { ?? $this._APIDELETE $this._API }
        Add-ObjectProperty $this "PolicyName" { $this._PolicyName }
        Add-ObjectProperty $this "PolicyBaseName" { ?? $this._PolicyBaseName $this._PolicyName }
        Add-ObjectProperty $this "PropertiesToRemove" { $this._PropertiesToRemove }
        Add-ObjectProperty $this "PropertiesToRemoveForUpdate" { $this._PropertiesToRemoveForUpdate }
        Add-ObjectProperty $this "ComparePropertiesToSkip" { $this._ComparePropertiesToSkip }
        Add-ObjectProperty $this "SkipRemoveProperties" { $this._SkipRemoveProperties }
        Add-ObjectProperty $this "Permissions" { $this._Permissions }
        Add-ObjectProperty $this "ViewProperties" { ?? $this._ViewProperties (Get-IntuneDefaultColumns $this) }
        Add-ObjectProperty $this "Expand" { $this._Expand }
        Add-ObjectProperty $this "Dependencies" { $this._Dependencies }
        Add-ObjectProperty $this "Icon" { Get-StringOrDefault $this._Icon $this._ID }
        #Add-ObjectProperty $this "IconImage" { $this._IconImage }
        Add-ObjectProperty $this "QueryList" { $this._QueryList }
        Add-ObjectProperty $this "QuerySearch" { Get-BoolOrDefault $this._QuerySearch $false }
        Add-ObjectProperty $this "ExpandAssignments" { $this._ExpandAssignments }
        Add-ObjectProperty $this "ExpandAssignmentsList" { $this._ExpandAssignmentsList }
        Add-ObjectProperty $this "SupportsAssignments" { $this._SupportsAssignments }
        Add-ObjectProperty $this "ScopeTagsReturnedInList" { $this._ScopeTagsReturnedInList }
        Add-ObjectProperty $this "ImportOrder" { $this._ImportOrder }
        Add-ObjectProperty $this "ODataMetadata" { $this._ODataMetadata }
        Add-ObjectProperty $this "AssignmentsType" { $this._AssignmentsType }
        Add-ObjectProperty $this "AssignAction" { ?? $this._AssignAction "assign" }
        Add-ObjectProperty $this "AssignmentsViaExpand" { Get-BoolOrDefault $this._AssignmentsViaExpand $false }
        Add-ObjectProperty $this "AssignmentObjectType" { $this._AssignmentObjectType }
        Add-ObjectProperty $this "AssignmentPropertiesToKeep" { $this._AssignmentPropertiesToKeep }
        Add-ObjectProperty $this "AssignmentTargetPropertiesToKeep" { $this._AssignmentTargetPropertiesToKeep }
        Add-ObjectProperty $this "SkipAddIDOnFileName" { Get-BoolOrDefault $this._SkipAddIDOnFileName $false }
        Add-ObjectProperty $this "ScopeTagProperty" { Get-StringOrDefault $this._ScopeTagProperty }
        Add-ObjectProperty $this "NavigationProperties" { $this._NavigationProperties }
        Add-ObjectProperty $this "HasDescription" { $this._HasDescription }
        Add-ObjectProperty $this "CopyDefaultName" { $this._CopyDefaultName }
        Add-ObjectProperty $this "NameProperty" { $this._NameProperty }
        Add-ObjectProperty $this "SupportsNameFilter" { $this._SupportsNameFilter }
        Add-ObjectProperty $this "IDProperty" { $this._IDProperty }

        Add-ObjectProperty $this "VerifyObject" { $this._VerifyObject }
        Add-ObjectProperty $this "StrictODataTypeCheck" { $this._StrictODataTypeCheck }

        Add-ObjectProperty $this "Folder" { ?? $this._Folder $this._ID }

        Add-ObjectProperty $this "Title" { ?? $this._APITitle $this._PolicyName }

        Add-ObjectProperty $this "HasPageSizeSupport" { $this._HasPageSizeSupport }
        Add-ObjectProperty $this "ShowButtons" { $this._ShowButtons }
        Add-ObjectProperty $this "SingleObject" { Get-BoolOrDefault $this._SingleObject $false }
    }

    [IntunePolicyBase]GetObject([PSCustomObject]$JsonObj)
    {
        if($this._ObjectClass) {
            $policyObject = (New-Object -TypeName $this._ObjectClass -ArgumentList $JsonObj)
            if($this.VerifyObject -and $this.CheckPolicy($JsonObj) -eq $false) {
                $policyObject = $null
            }
            # When several PolicyType siblings share the same _ObjectClass (e.g. all
            # deviceEnrollmentConfigurations subtypes wrap rows in DeviceEnrollmentObject),
            # the object's Init() picks a single hard-coded singleton — usually the base
            # class. Override with the *calling* subtype so the policy's PolicyType.Folder
            # and other type-keyed lookups resolve to the bucket that actually claimed it.
            if($null -ne $policyObject) {
                $policyObject._PolicyType = $this
                $this.LogResolvedPolicy($JsonObj)
            }
            return $policyObject
        }
        else {
            Write-Log "Object class is missing for $($this.Title) $($JsonObj."$($this.NameProperty)")"
        }
        return $null
    }

    [IntunePolicyBase]GetObject([String]$Json)
    {
        $jsonObj = (ConvertFrom-Json $Json)
        return ($this.GetObject($jsonObj))
    }

    [IntunePolicyBase]GetObject([IO.FileInfo]$File)
    {
        try {
            $tmpObject = (New-Object -TypeName $this._ObjectClass -ArgumentList $File)
            if($this.VerifyObject -and $this.CheckPolicy($tmpObject.JsonObject) -eq $false) {
                $tmpObject = $null
            }
            if($null -ne $tmpObject) {
                $tmpObject._PolicyType = $this
                $this.LogResolvedPolicy($tmpObject.JsonObject)
            }
            return $tmpObject
        }
        catch {}
        return $null
    }

    AddExportProperties([IntuneManagerExportSettings]$ExportSettings)
    {

    }

    AddImportProperties([IntuneManagerImportSettings]$ImportSettings)
    {

    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.'@odata.id') {
            if($PolicyObject.'@odata.id'.StartsWith($this._API.TrimStart('/'))) {
                return $true
            }
        }

        return $false
    }

    # Uniform "found policy" breadcrumb, logged once per object at the point a
    # type claims it - both the list-load path (Get-GraphPolicies) and the
    # file-load path (Get-GraphPolicyFromFile) route every resolved object
    # through GetObject, so this is the single chokepoint. It replaces the
    # ad-hoc Write-Log lines that a handful of CheckPolicy overrides used to
    # carry: those fired only for the subset of types that both overrode
    # CheckPolicy AND happened to include the line, so the log was verbose for
    # some types and silent for others. Reporting from here makes it consistent
    # for every type. The discriminators (@odata.type / @odata.id, plus the two
    # template/setting keys some types match on) are appended only when present.
    LogResolvedPolicy([PSCustomObject]$PolicyObject)
    {
        $name = $PolicyObject."$($this.NameProperty)"

        $parts = @()
        if($PolicyObject.'@odata.type') { $parts += "@odata.type: $($PolicyObject.'@odata.type')" }
        if($PolicyObject.'@odata.id')   { $parts += "@odata.id: $($PolicyObject.'@odata.id')" }
        if($PolicyObject.templateReference.templateFamily) { $parts += "templateFamily: $($PolicyObject.templateReference.templateFamily)" }
        if($PolicyObject.settingDefinitionId) { $parts += "settingDefinitionId: $($PolicyObject.settingDefinitionId)" }

        $basis = if($parts.Count -gt 0) { " Found based on $($parts -join ', ')" } else { "" }
        # Verbose channel: this fires once per resolved object (every type, every
        # list/file load), so it is diagnostic noise in the normal log. Write-LogDebug
        # only emits when the Debug setting is on, keeping the default log clean while
        # the breadcrumb is still available when troubleshooting type resolution.
        Write-LogDebug "Found policy $name ($($this.Title)).$basis"
    }

    # Single-policy entry point preserved for back-compat with Compare and a
    # handful of subtype overrides (e.g. windows10CustomConfiguration OMA
    # decrypt). Routes through Invoke-PolicyHydrate so there is exactly one
    # hydration path; the body fetch / nav-property expansion / sub-resource
    # fan-out all live in the unified orchestrator.
    [Boolean]GetFullObject($PolicyObject)
    {
        if($PolicyObject._IsFullObject) { return $true }
        if($null -eq $PolicyObject._TokenId) { return $false }

        Invoke-PolicyHydrate -Policies @($PolicyObject) -TokenId $PolicyObject._TokenId

        if(-not $PolicyObject._IsFullObject) {
            Write-Warning "Failed to get full object for $($PolicyObject.Name)"
            return $false
        }
        return $true
    }

    # Returns a hashtable describing how to compare settings of this policy type, or $null for default property comparison.
    # Expected keys when non-null:
    #   Prop        : property name on .Object holding the settings array (required)
    #   GetKey      : scriptblock { param($s) ... } returning a stable string key per setting (required)
    #   GetValue    : scriptblock { param($s) ... } returning the display value per setting (required)
    #   GetCategory : scriptblock { param($s) ... } returning the category label (optional)
    #   Hydrate     : scriptblock { param($policy) ... } loading any data missing after GetFullObject (optional)
    [Hashtable]GetCompareConfig()
    {
        return $null
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {

    }

    Hidden PreExportCommand([IntunePolicyBase]$PolicyObject, [PSCustomObject]$ExportObject)
    {

    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        return $null
    }

    [Hashtable]PreImportAssignmentsCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        return $null
    }

    PostImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {

    }

    PostImportAssignmentsCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject, [PSCustomObject]$ImportedAssignments)
    {

    }

    PostCopyCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {

    }

    PostBulkImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {

    }

    [Hashtable]PreDeleteCommand([IntunePolicyBase]$PolicyObject)
    {
        return $null
    }

    [Hashtable]PreUpdateCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        return $null
    }

    [Hashtable]PreReplaceCommand([IntunePolicyBase]$PolicyObject)
    {
        return $null
    }

    PostReplaceCommand([IntunePolicyBase]$PolicyObject, [IntunePolicyBase]$SourceObject)
    {

    }

    [IntunePolicyBase[]]PreImportPolicies([IntunePolicyBase[]]$PolicyObjects)
    {
        return @($PolicyObjects)
    }
}

class IntunePolicyBase
{
    # Hidden Properties
    Hidden [IO.FileInfo]$FileInfoObject = $null
    Hidden [PSCustomObject]$JsonObject = $null
    Hidden [IntunePolicyTypeBase]$_PolicyType = $null
    Hidden [String]$_PolicyName = $null
    Hidden [boolean]$_IsFullObject = $false
    # Set to $true on subclasses that implement GetSubResourceBatchRequests /
    # ApplySubResourceBatchResult so the unified hydrator
    # (Invoke-PolicySubresourceFetch) can fan out their extra GETs in $batch
    # form during Invoke-PolicyHydrate.
    Hidden [boolean]$_HasSubResourceBatch = $false
    # Set by Start-GraphBulkExport when another policy in the same type bucket
    # shares this one's displayName, so GetFileName appends `_<id>` and the
    # collision doesn't silently overwrite a sibling on disk.
    Hidden [boolean]$_NeedsIdInFilename = $false
    Hidden [string]$_PlatformName = $null
    Hidden [PSCustomObject]$_ClonedFromObject = $null
    Hidden [int]$_TokenID = $null
    Hidden [String[]]$_ScopeTags = $()
    Hidden [String]$_ScopeTagsString = $null

    [string]$TenantId = $null
    

    IntunePolicyBase() {
        ([IntunePolicyBase]$this).Init()
    }
    
    IntunePolicyBase([String]$Json) {
        ([IntunePolicyBase]$this).Init($Json)
    }

    IntunePolicyBase([PSCustomObject]$JsonObj) {
        ([IntunePolicyBase]$this).Init($JsonObj)
    }    

    IntunePolicyBase([IO.FileInfo]$FileInfo) {
        ([IntunePolicyBase]$this).Init($FileInfo)
    }

    # Hidden Functions
    Hidden Init()
    {
        Add-ObjectProperty $this "ID" { return $this.GetId() }
        Add-ObjectProperty $this "Name" { return $this.GetName() } { $this.SetName($args[0]) }
        Add-ObjectProperty $this "PolicyType" { $this._PolicyType }
        Add-ObjectProperty $this "PolicyName" { $this.GetPolicyName() }
        Add-ObjectProperty $this "PolicyBaseName" { $this.GetPolicyBaseName() }
        # Per-row _PlatformName wins; if absent, fall back to the type's default.
        # Lets multiple subtypes share one Object class (e.g. DeviceEnrollment
        # buckets all use DeviceEnrollmentObject) and still report the right
        # Platform per bucket via type-level _PlatformName.
        Add-ObjectProperty $this "Platform" {
            if($this._PlatformName) { return $this._PlatformName }
            if($this._PolicyType -and $this._PolicyType._PlatformName) { return $this._PolicyType._PlatformName }
            return $null
        }
        Add-ObjectProperty $this "Object" { $this.JsonObject }
        # Bound by the grid (and therefore searchable by the filter) on every view.
        Add-ObjectProperty $this "Description" { if($this.JsonObject) { $this.JsonObject.description } }
        Add-ObjectProperty $this "IsFromFile" { $this.GetFromFile() }
        Add-ObjectProperty $this "FileInfo" { ([PSCustomObject]$this.FileInfoObject) }
        Add-ObjectProperty $this "IsFullObject" { $this._IsFullObject }
        Add-ObjectProperty $this "JsonString" { if($this.JsonObject) { ConvertTo-Json $this.JsonObject -Depth 50} }
        Add-ObjectProperty $this "TokenId" { $this._TokenId }
        Add-ObjectProperty $this "Created" { 
            if($this.JsonObject.createdDateTime) {
                Get-Date $this.JsonObject.createdDateTime
            } 
            else { "" }
        }

        Add-ObjectProperty $this "LastModified" { 
            if($this.JsonObject.lastModifiedDateTime) {
                Get-Date $this.JsonObject.lastModifiedDateTime
            } 
            elseif($this.JsonObject.modifiedDateTime) {
                Get-Date $this.JsonObject.modifiedDateTime
            }
            else { "" }
        }

        Add-ObjectProperty $this "ScopeTags" {
            if($null -eq $this._ScopeTags) {
                $this._ScopeTags = Get-GraphScopeTags $this
                $this._ScopeTagsString = $this._ScopeTags -join ","
            }
            return $this._ScopeTagsString
        }

        # Scope-tag NAMES as an array (same lazy resolution as ScopeTags, which
        # returns the comma-joined string). Lets callers use array operators —
        # e.g. a documentation filter `$_.ScopeTagNames -contains 'Production'`.
        Add-ObjectProperty $this "ScopeTagNames" {
            if($null -eq $this._ScopeTags) {
                $this._ScopeTags = Get-GraphScopeTags $this
                $this._ScopeTagsString = $this._ScopeTags -join ","
            }
            return $this._ScopeTags
        }

        # Both fields can carry a comma-separated flags value, so translation
        # goes through Get-PlatformDisplayName rather than a direct key lookup.
        if($this.JsonObject.platforms) {
            $this._PlatformName = Get-PlatformDisplayName $this.JsonObject.platforms
        }
        elseif($this.JsonObject.platformType) {
            $this._PlatformName = Get-PlatformDisplayName $this.JsonObject.platformType
        }
        
        if(-not $this._PlatformName -and $this.JsonObject.'@OData.Type'){
            $this._PlatformName = Get-PolicyPlatformName $this.JsonObject.'@OData.Type'
        }
    }

    Hidden Init([string]$Json)
    {
        if($Json) { $this.LoadJson($Json) }
        ([IntunePolicyBase]$this).Init()
    }

    Hidden Init([PSCustomObject]$JsonObject) {
        $this.JsonObject = $JsonObject
        ([IntunePolicyBase]$this).Init()
    }    

    Hidden Init([IO.FileInfo]$FileInfo)
    {
        if($FileInfo -and $FileInfo.Exists) {
            Write-LogDebug "Loading object from file $($FileInfo.FullName)"            
            # ReadAllText, not Get-Content: this is the load path for every policy
            # read from disk (import, compare, file-based documentation). Bare
            # Get-Content decodes as the ANSI codepage on PS5.1, which mangles any
            # non-ASCII policy name or description on the way in.
            $Json = [IO.File]::ReadAllText($FileInfo.FullName)
            $this.LoadJson($Json)
            $this.FileInfoObject = $FileInfo
        }
        else 
        {
            $errorString = "Failed to load object from file. File $($FileInfo.FullName) not found"
            Write-Log $errorString 3
            throw $errorString
        }
        ([IntunePolicyBase]$this).Init()
    }

    Hidden LoadJson([String]$Json)
    {
        $this.JsonObject = $Json | ConvertFrom-Json 
    }

    Hidden [String] GetName()
    {
        return $this.JsonObject."$($this.PolicyType._NameProperty)"
    }

    Hidden SetName($Name)
    {
        $this.JsonObject."$($this.PolicyType._NameProperty)" = $Name
    }

    Hidden [String] GetId()
    {
        return $this.JsonObject.id
    }

    Hidden [Boolean]GetFromFile()
    {
        return ($null -ne $this.FileInfoObject)
    }

    Hidden [String]GetPolicyName() 
    {
        return ?? $this._PolicyName $this._PolicyType.PolicyName
    }

    Hidden [String]GetPolicyBaseName()
    {
        return ?? $this._PolicyBaseName $this._PolicyType.PolicyBaseName
    }    

    # Public functions
    [String]GetObjectURL()
    {
        $params = @()
        $uri = [uri]"$($this.PolicyType.API)/$($this.id)"

        $expand = @()
        # NOTE: the property check below reads `.ExpandAssignments` not
        # `.ExpandAssignmentsList`. This LOOKS like a typo — and an earlier
        # session "fixed" it to ExpandAssignmentsList — but the wrong-name
        # behaviour is load-bearing here:
        # * `_ExpandAssignmentsList=$false` is meant to skip $expand=assignments
        #   on the LIST URL because some Graph endpoints reject it there.
        # * For the per-id BODY URL, `$expand=assignments` is the only working
        #   path for types whose `/assignments` sub-resource returns 400
        #   (managedAppPolicies, notificationMessageTemplates).
        # The typo'd check resolves to `$null -ne $false` (true), so the body
        # URL keeps the assignments expand regardless of the list-URL opt-out.
        # Correcting the name regressed AppProtection's actual assignment
        # rows to null on bulk export.
        if($this.JsonObject.'assignments@odata.navigationLink' -and $this.PolicyType.ExpandAssignments -ne $false)
        {
            $expand += "assignments"
        }
    
        if($this.JsonObject.'apps@odata.navigationLink')
        {
            $expand += "apps"
        }
    
        if($this.JsonObject.'settings@odata.navigationLink')
        {
            $expand += "settings"
        }
    
        if($this.JsonObject.'roleAssignments@odata.navigationLink')
        {
            $expand += "roleAssignments"
        }
        
        if($this.JsonObject.'privacyAccessControls@odata.associationLink')
        {
            $expand += "microsoft.graph.windows10GeneralConfiguration/privacyAccessControls"
        }    
        
        if($this.PolicyType.Expand)
        {
            foreach($objExpand in $this.PolicyType.Expand.Split(","))
            {
                if($objExpand -notin $expand) { $expand += $objExpand}
            }
        }
        
        if($expand.Count -gt 0) {
            $params += ('$expand=' + ($expand -join ",")) # ToDo: Check if expand is set
        }

        $url = $uri.OriginalString
        if($params.Count -gt 0) {
            $url += (?: ([String]::IsNullOrEmpty($uri.Query)) "?" "&")
            $url += ($params -join "&")
        }

        return $url
    }

    [Boolean]Get()
    {
        if($this._IsFullObject) { return $true }
        if($null -eq $this._TokenId) { return $false }
        Invoke-PolicyHydrate -Policies @($this) -TokenId $this._TokenId
        return $this._IsFullObject
    }

    # Bulk-export sub-resource contract. Override in subclasses that need to
    # fetch extra resources beyond the main object body (e.g. relationships,
    # script content, target apps). Phase 1 is the entry point; phases 2+ come
    # from follow-ups returned by ApplySubResourceBatchResult.
    # Each request: [PSCustomObject]@{ Key = '<unique-within-this-policy>'; Url = '<graph-relative>'; Method = 'GET' (optional); Headers = @{} (optional) }
    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        return @()
    }

    # Routes a batched response back into the policy. Returns follow-up requests
    # for the next phase (empty array = done). $Body is $null for non-2xx
    # responses; the orchestrator already logged the HTTP status.
    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        return @()
    }

    # Called once per opted-in policy by Invoke-PolicySubresourceFetch AFTER all
    # phases complete. Override to build a derived property from the responses
    # already applied via ApplySubResourceBatchResult (e.g. AppConfiguration's
    # #CustomRefTargetedApps, which needs every targeted-app body resolved before
    # it can be assembled). Default no-op. Runs even when GetSubResourceBatchRequests
    # returned nothing, so a class can finalize from cache/state alone.
    [void] FinalizeSubResources()
    {
    }

    GetAssignments()
    {
        if(-not $this.Object.Assignments -and $null -ne $this._TokenId)
        {
            $url = "$($this._PolicyType.API)/$($this.id)/assignments"
            $assignments = (Invoke-MSGraphAPI -Url $url -TokenId $this._TokenId).Value
            if($assignments)
            {
                $this.Object.Assignments = $assignments
            }
        }
    }

    # Returns the specific objects this policy depends on.
    # Each entry: [PSCustomObject]@{ TypeId = "<policy type ID>"; Id = "<object GUID>" }
    # Override in derived classes to enable targeted dependency loading.
    # The base implementation returns an empty array (falls back to type-level _Dependencies).
    [PSCustomObject[]] GetDependencyReferences()
    {
        return @()
    }

    [String]GetFileName()
    {
        # `$this.` is required: a bare `GetFileName($null)` is parsed as a COMMAND
        # invocation, not a method call, so this overload threw "The term
        # 'GetFileName' is not recognized" for every caller that omitted the path.
        return ($this.GetFileName($null))
    }

    [String]GetFileName([String]$Path)
    {
        $fileName = $this.Name.Trim('.')
        # Append `_<id>` to the filename when EITHER the user-facing setting
        # AddIDToExportFile is on (matches the OLD project's behaviour) OR the
        # bulk-export collision detector flagged this specific policy because
        # another policy in the same type bucket shares its displayName.
        # Without the collision case, exporting N same-named policies would
        # silently overwrite the same file N times and only the last one would
        # survive on disk (seen in a lab tenant: 81 CA policies -> 39 unique
        # displayName -> 42 lost without this guard).
        $forceId = (Get-SettingValue "AddIDToExportFile") -eq $true -or $this._NeedsIdInFilename -eq $true
        if($forceId -and $this.Id -and $this.PolicyType.SkipAddIDOnFileName -ne $true) {
            $fileName = ($fileName + "_" + $this.Id)
        }
        $fileName = "$((Remove-InvalidFileNameChars $fileName)).json"

        if($Path) {
            $fileName = [IO.Path]::Combine($Path, $fileName)
        }

        return $fileName
    }

    [string]ExportToFile([String]$ParentFolder)
    {
        return ($this.ExportToFile($ParentFolder, $null))
    }

    [string]ExportToFile([String]$ParentFolder, [String]$FileName)
    {
        $returnValue = $null

        if(-not $FileName) { $FileName = $this.GetFileName($ParentFolder) }

        # Clone JsonObject via JSON round-trip so PreExportCommand can mutate
        # without affecting the in-memory policy. Same pattern as Clone().
        $exportObject = ConvertFrom-Json (ConvertTo-Json $this.JsonObject -Depth 50)

        $this.PolicyType.PreExportCommand($this, $exportObject)

        # Central serializer so the policy file honours the same
        # SortJsonProperties / ExportJsonFormat settings as the MigrationTable
        # and the sidecars. It previously called ConvertTo-Json directly, which
        # meant SortJsonProperties was silently ignored for policy exports.
        $json = ConvertTo-GraphExportJson $exportObject -Depth 50

        # Organization values -> placeholders. Internal/ExportTokens.ps1 owns the
        # replace, the "Replace organization values in export files" setting and
        # the hidden ExportReplaceTokens key that narrows which values are
        # replaced.
        #
        # Masked against THIS policy's own organization, not the default token's:
        # exporting a policy listed from tenant A while tenant B is the default
        # otherwise masked B's id in A's data and left A's real tenant id in the
        # file. Same reasoning as Save-GraphObjectToFile does for the sidecars.
        $orgInfo = Get-GraphObjectOrganizationInfo $this
        $json = Convert-GraphOrganizationValueToToken $json `
                    -OrganizationId $orgInfo.OrganizationId `
                    -OrganizationName $orgInfo.OrganizationName


        try
        {
            # [IO.File]::WriteAllText is ~10x faster than Out-File for this hot path
            # (called once per exported policy). Out-File goes through PowerShell's
            # pipeline + .NET StreamWriter wrapper with Format-Table conversion; the
            # direct .NET call skips all that.
            #
            # Encoding comes from the "Export file encoding" setting, which every
            # exported file shares - policy JSON, the MigrationTable and the
            # assignment sidecars. This used to hard-code UTF-8 WITH BOM, and the
            # sidecars used a bare Out-File (UTF-16LE on PS5.1). Consumers that read
            # the export as plain UTF-8 - git, the Intune portal, other tooling -
            # choke on a BOM, so the default is now UTF-8 without one.
            [System.IO.File]::WriteAllText($fileName, $json, (Get-ExportFileEncoding))
            $returnValue = $fileName
        }
        catch
        {
            Write-LogError "Failed to save file $fileName" $_.Exception
        }

        return $returnValue
    }

    Hidden SetDescription($Description)
    {
        if($this.HasDescription -ne $false) {
            $this.JsonObject.description = $Description
        }
    }    
    
    Hidden [IntunePolicyBase]Clone()
    {
        $cloned = $this._PolicyType.GetObject((ConvertTo-Json $this.JsonObject -Depth 50))

        $cloned._ClonedFromObject = ?? $this._ClonedFromObject $this

        $cloned.TenantId = $this.TenantId

        # Import runs against the clone, and PreImportCommand hooks read FileInfo to
        # find files that sit next to the exported json - the ADMX/ADML pair, the
        # Terms of Use PDFs. Rebuilding the clone from JSON alone dropped it, so
        # those lookups saw $null and silently fell back to the app packages folder.
        $cloned.FileInfoObject = $this.FileInfoObject

        return $cloned
    }

    Hidden [IntunePolicyBase]CopyObject([String]$Name, [String]$Description, [int]$ToTokenId)
    {
        # Backward-compat overload — inherits scope tags from the source.
        return $this.CopyObject($Name, $Description, $ToTokenId, $null)
    }

    Hidden [IntunePolicyBase]CopyObject([String]$Name, [String]$Description, [int]$ToTokenId, [string[]]$ScopeTagIds)
    {
        Write-Log "Copy $($this.PolicyType.Title) object $($this.Name)"

        if($this.IsFromFile -eq $false -and $this.IsFullObject -eq $false -and $this._RequiresFullObject -ne $false) {
            $this.Get() | Out-Null
        }

        $clonedObject = $this.Clone()

        if($clonedObject)
        {
            if($clonedObject._RequiresFullObject -eq $true) {
                [void]$clonedObject.Get()
            }

            $clonedObject.SetName($Name)
            if($Description) {
                $clonedObject.SetDescription($Description)
            }

            # Override scope tags on the clone when the caller passed an explicit
            # list. $null = leave as-is (inherit). Empty array = explicit reset
            # (Intune will re-apply Default automatically in that case).
            if($null -ne $ScopeTagIds -and $clonedObject.PolicyType.ScopeTagProperty) {
                $stProp = [string]$clonedObject.PolicyType.ScopeTagProperty
                if($clonedObject.JsonObject.PSObject.Properties[$stProp]) {
                    $clonedObject.JsonObject.$stProp = @($ScopeTagIds)
                } else {
                    $clonedObject.JsonObject | Add-Member -MemberType NoteProperty -Name $stProp -Value @($ScopeTagIds) -Force
                }
            }

            Add-GraphNavigationProperties $this

            $newObj = $clonedObject.ImportObject($ToTokenId, $null)
            if(-not $newObj)
            {
                return $null
            }

            if($newObj.PolicyType.NavigationProperties -eq $true) {
                Set-GraphNavigationProperties $newObj $this
            }

            $newObj.PolicyType.PostCopyCommand($newObj, $this)

            return $newObj
        }

        return $null
    }

    Hidden [PSCustomObject]ImportObject([int]$TokenId, [HashTable]$BulkImport)
    {
        # Get-OperationTokenInfo, not Get-TokenInfo: the latter treats 0 as "list
        # every token", and 0 is also how a caller spells "the default token". With
        # two tenants signed in that made $tokenInfo an ARRAY, so .TenantId below
        # restored %OrganizationId% as two space-joined guids.
        $tokenInfo = Get-OperationTokenInfo $TokenId

        Write-Log "Import $($this.PolicyType.Title) object '$($this.Name)' to $($tokenInfo.TenantName)"

        $clonedObject = $this.Clone()
        
        Remove-GraphPropertiesForImport $clonedObject $clonedObject.JsonObject

        $params = @{}
        $strAPI = (?? $clonedObject.PolicyType.APIPOST $clonedObject.PolicyType.API)
        $method = "POST"

        $ret = $clonedObject.PolicyType.PreImportCommand($clonedObject)

        if($ret -is [HashTable])
        {
            if($ret.ContainsKey("Import") -and $ret["Import"] -eq $false)
            {
                # Import handled manually
                return $false
            }

            if($ret.ContainsKey("API"))
            {
                $strAPI = $ret["API"]
            }
            
            if($ret.ContainsKey("Method"))
            {
                $method = $ret["Method"]
            }

            if($ret.ContainsKey("AdditionalHeaders") -and $ret["AdditionalHeaders"] -is [HashTable])
            {
                $params.Add("AdditionalHeaders",$ret["AdditionalHeaders"])
            }
        }

        # Resolve scope tags on the cloned body before serialization (ImportScopeTags
        # setting + same/cross-tenant remap). Keeps source IDs same-tenant, remaps
        # cross-tenant, or resets to Default when the setting is off.
        Set-GraphImportScopeTags $clonedObject $TokenId

        $json = ConvertTo-Json $clonedObject.JsonObject -Depth 50
        if($tokenInfo -and $clonedObject._ClonedFromObject.TenantId -ne $tokenInfo.TenantId)
        {
            # Call Update-JsonForEnvironment before importing the object
            # E.g. PolicySets contains references, AppConfiguration policies reference apps etc.
            $json = Update-JsonForEnvironment $json $clonedObject $TokenId
        }
    
        # Placeholders -> values. Always runs: an export made when the replace
        # setting was on carries placeholders whatever it is set to now.
        # Internal/ExportTokens.ps1 decides which placeholders are restored.
        #
        # Restored to the TARGET tenant of this import, not to whichever tenant
        # happens to hold the default token: importing into tenant B on a
        # non-default token used to write tenant A's GUID into B's policy.
        $json = Convert-GraphOrganizationTokenToValue $json `
                    -OrganizationId $tokenInfo.TenantId `
                    -OrganizationName $tokenInfo.TenantName

        if($BulkImport) {
            Write-Status "Add $($this.Name) ($($this.PolicyName)) with id $($this.Id) to Bulk Import list"            
            
            $bulkImportObject = [PSCustomObject]@{
                id = ($this.Name + "_" + (New-Guid).Guid.SubString(0,8))
                method = $method
                url = $strAPI.TrimStart('/')
                headers = @{"Content-Type"="application/json;odata.metadata=none"}
                body = $json | ConvertFrom-Json #[System.Text.Encoding]::UTF8.GetBytes($json)
            }
            $BulkImport.Add($bulkImportObject, [PSCustomObject]@{
                ImportObject = $clonedObject
                FromObject   = $this
            })
            return $null
        }
    
        $newObj = $null

        $responseObject = $clonedObject.ImportObject($strAPI, $json, $method, $TokenId, $params)

        $newObj = $clonedObject.ProcessImportResponse($TokenId, $responseObject, $method)
        
        return $newObj
    }        

    Hidden [PSCustomObject]ImportObject([string]$API,[string]$Json, [string]$Method, [int]$TokenID, [Hashtable]$Params)
    {
        $responseObject = (Invoke-MSGraphAPI -Url $API -Content $Json -HttpMethod $Method -TokenId $TokenID -FullResponseObject @Params)

        return $responseObject
    }

    Hidden [IntunePolicyBase]UpdateObject([IntunePolicyBase]$ExistingObject, [int]$TokenId)
    {
        if($null -eq $ExistingObject) {
            Write-Log "Cannot update $($this.PolicyType.Title) object '$($this.Name)'. Existing object was not supplied." 3
            return $null
        }

        # One token, never the whole list - see ImportObject above.
        $tokenInfo = Get-OperationTokenInfo $TokenId
        Write-Log "Update $($this.PolicyType.Title) object '$($ExistingObject.Name)' from imported object '$($this.Name)'"

        $updateObject = $this.Clone()
        $targetId = $ExistingObject.Id
        if(-not $targetId) {
            Write-Log "Cannot update $($this.PolicyType.Title) object '$($this.Name)'. Existing object has no id." 3
            return $null
        }

        if($updateObject.JsonObject.PSObject.Properties['id']) {
            $updateObject.JsonObject.id = $targetId
        }
        else {
            $updateObject.JsonObject | Add-Member -MemberType NoteProperty -Name 'id' -Value $targetId -Force
        }
        $updateObject._TokenID = $TokenId
        if($tokenInfo) { $updateObject.TenantId = $tokenInfo.TenantId }

        $params = @{}
        $strAPI = (?? $updateObject.PolicyType.APIPUT $updateObject.PolicyType.API) + "/$targetId"
        $method = "PATCH"

        $ret = $updateObject.PolicyType.PreUpdateCommand($updateObject, $ExistingObject)
        if($ret -is [HashTable])
        {
            if($ret.ContainsKey("Update") -and $ret["Update"] -eq $false)
            {
                Write-Log "Update skipped by $($updateObject.PolicyType.Title) pre-update command for '$($ExistingObject.Name)'" 2
                return $null
            }

            if($ret.ContainsKey("API"))
            {
                $strAPI = $ret["API"]
            }
            
            if($ret.ContainsKey("Method"))
            {
                $method = $ret["Method"]
            }

            if($ret.ContainsKey("AdditionalHeaders") -and $ret["AdditionalHeaders"] -is [HashTable])
            {
                $params.Add("AdditionalHeaders",$ret["AdditionalHeaders"])
            }
        }

        Remove-GraphPropertiesForImport $updateObject $updateObject.JsonObject
        # isAssigned is a computed status flag many endpoints refuse to accept
        # back on PATCH - the original project stripped it on every update.
        Remove-Property $updateObject.JsonObject "isAssigned"
        foreach($propertyToRemove in @($updateObject.PolicyType.PropertiesToRemoveForUpdate)) {
            Remove-Property $updateObject.JsonObject $propertyToRemove
        }

        $json = ConvertTo-Json $updateObject.JsonObject -Depth 50
        if($tokenInfo -and $updateObject._ClonedFromObject.TenantId -ne $tokenInfo.TenantId)
        {
            $json = Update-JsonForEnvironment $json $updateObject $TokenId
        }
    
        # Placeholders -> values. Always runs: an export made when the replace
        # setting was on carries placeholders whatever it is set to now.
        # Internal/ExportTokens.ps1 decides which placeholders are restored.
        #
        # Restored to the TARGET tenant of this import, not to whichever tenant
        # happens to hold the default token: importing into tenant B on a
        # non-default token used to write tenant A's GUID into B's policy.
        $json = Convert-GraphOrganizationTokenToValue $json `
                    -OrganizationId $tokenInfo.TenantId `
                    -OrganizationName $tokenInfo.TenantName

        $responseObject = Invoke-MSGraphAPI -Url $strAPI -Content $json -HttpMethod $method -TokenId $TokenId -FullResponseObject @params
        if($responseObject.Success) {
            $ExistingObject._TokenID = $TokenId
            if($tokenInfo) { $ExistingObject.TenantId = $tokenInfo.TenantId }
            [void]$ExistingObject.Get()

            Write-Log "$($ExistingObject.PolicyType.Title) object updated successfully: $($ExistingObject.Name) ($($ExistingObject.Id))"

            Import-GraphObjectAssignment $ExistingObject $this

            return $ExistingObject
        }

        # Include the Graph error body - StatusDescription alone ("Bad
        # Request") says nothing about WHY the update was rejected.
        $errorDetail = ""
        try {
            $errorContent = if($responseObject.Content -is [string]) { $responseObject.Content | ConvertFrom-Json } else { $responseObject.Content }
            if($errorContent.error.message) { $errorDetail = " - $($errorContent.error.message)" }
        } catch { }
        Write-Log "Failed to update $($ExistingObject.PolicyType.Title) object '$($ExistingObject.Name)' ($($ExistingObject.Id)). $($responseObject.StatusCode) $($responseObject.StatusDescription)$errorDetail" 3
        return $null
    }

    Hidden [PSCustomObject]ProcessImportResponse([int]$TokenId, [PSCustomObject]$Response, [String]$Method)
    {
        $newObj = $null
        # One token, never the whole list - see ImportObject above.
        $tokenInfo = Get-OperationTokenInfo $TokenId
        if($Response.Success)
        {
            if($method -eq "POST" -and $Response.Content) {
                $newObj = $this.PolicyType.GetObject($Response.Content)
                if(-not $newObj -and $this.PolicyType._ObjectClass) {
                    # GetObject runs CheckPolicy for _VerifyObject types and can
                    # reject a minimal POST response body (e.g. Settings Catalog
                    # create responses may omit templateReference). We KNOW the
                    # object belongs to this type - we just created it - so
                    # construct it directly without verification.
                    $newObj = New-Object -TypeName $this.PolicyType._ObjectClass -ArgumentList $Response.Content
                }
            }
            else {
                [void]$this.Get()
                $newObj = $this 
            }
            if($tokenInfo) {
                $newObj._TokenID = $tokenInfo.Id
                $newObj.TenantId = $tokenInfo.TenantId
            }
            else {
                $newObj._TokenID = $TokenId
                try {
                    $userInfo = (Get-AuthProvider).GetUserInfo($TokenId)
                    if($userInfo -and $userInfo.TenantId) { $newObj.TenantId = $userInfo.TenantId }
                } catch { }
            }

            Write-Log "$($newObj.PolicyType.Title) object imported successfully: $($newObj.Name) ($($newObj.Id))"

            # PostImportCommand post-processes using SOURCE data that was stripped
            # from this clone before the create POST (AdminTemplate definitionValues,
            # app script/content file info, ...). $this here is the stripped clone,
            # so hand it the ORIGINAL via _ClonedFromObject; fall back to $this for a
            # direct (non-clone) call.
            $sourceObj = $this._ClonedFromObject
            if($null -eq $sourceObj) { $sourceObj = $this }
            $newObj.PolicyType.PostImportCommand($newObj, $sourceObj)

            Import-GraphObjectAssignment $newObj $this
        }
        return $newObj
    }

    Hidden [bool]Set([string]$Json)
    {
        return ($this.Set($null, $Json, $null, $null))
    }

    Hidden [bool]Set([string]$API, [string]$Json)
    {
        return ($this.Set($API, $Json, $null, $null))
    }

    Hidden [bool]Set([string]$API, [string]$Json, [string]$Method, [Hashtable]$Params)
    {
        if($null -eq $Json) {
            Write-Log "Cannot update object without Json data"
            return $false
        }
        
        if(-not $API) {
            $API = (?? $this.PolicyType.APIPOST $this.PolicyType.API) + "/$($this.Id)"
        }

        if($null -ne $this._TokenID) {
            $tokenId = $this._TokenID
        }
        else {
            Write-Log "Updating object $($this.Name) ($($this.Id)) with default token. TokenId not set on object" 2
            $tokenId = 0
        }

        if(-not $Method) {
            $Method = "PATCH"
        }

        $responseObject = (Invoke-MSGraphAPI -Url $API -Content $Json -HttpMethod $Method -TokenId $tokenId -FullResponseObject @Params)
        if($responseObject.Success) {
            Write-Log "$($this.Name) ($($this.Id)) updated successfully"
            [void]$this.Get()
        }

        return $responseObject.Success
    }

    Hidden [bool]Delete([HashTable]$BulkDelete)
    {
        $deleteAPI = $null
        $ret = $this.PolicyType.PreDeleteCommand($this)
        if($ret -is [HashTable])
        {
            if($ret.ContainsKey("Delete") -and $ret["Delete"] -eq $false)
            {
                # Delete handled manually or aborted
                return $false
            }

            if($ret.ContainsKey("API"))
            {
                $deleteAPI = $ret["API"]
            }
        }

        if(-not $deleteAPI) {
            $deleteAPI = $this.PolicyType.APIDELETE
        }

        if($BulkDelete -is [HashTable])
        {
            Write-Status "Add $($this.Name) ($($this.PolicyName)) with id $($this.Id) to Bulk Delete list"            
            $url = ($deleteAPI + "/$($this.Id)")
            $bulkDeleteObject = [PSCustomObject]@{
                id = $this.Id 
                method = "DELETE"
                url = $url.TrimStart('/')
                headers = @{"Accept"="application/json;odata.metadata=none"}
            }
            $BulkDelete.Add($bulkDeleteObject, $this)            
        }
        else {
            Write-Status "Delete $($this.Name) ($($this.PolicyName)) with id $($this.Id)"
            $deleteAPI = ($deleteAPI + "/$($this.Id)")
            $response = Invoke-MSGraphAPI -Url $deleteAPI -HttpMethod "DELETE" -ODataMetadata "none" -FullResponseObject
            if($response.Success) {
                Write-LogDebug "Policy deleted successfully"
            }
            return $response.Success
        }            
        return $true
    }
}

<#
class IntuneImportExportPolicyBase : PSCustomObject, System.ComponentModel.INotifyPropertyChanged
{
    [Boolean]$IsSelected = $true
    [Boolean]$_AddObjectType = $null
    [Boolean]$_AddCompanyName = $null
    [System.Collections.ArrayList]$PropertyChanged = @()

    IntuneImportExportPolicyBase() {
        $this.PSObject._AddObjectType = (Get-SettingValue "AddObjectType")
        $this.PSObject._AddCompanyName = (Get-SettingValue "AddCompanyName")

        Add-ObjectProperty $this "AddObjectType" { return $this.PSObject._AddObjectType } { 
            param([Boolean]$Value)
            $this.PSObject._AddObjectType = $Value;
            $this.PSObject.psobject.NotifyPropertyChanged('AddObjectType')
        }

        Add-ObjectProperty $this "AddCompanyName" { return $this.PSObject._AddObjectType } { 
            param([Boolean]$Value)
            $this.psobject._AddCompanyName = $Value;
            $this.psobject.NotifyPropertyChanged('AddCompanyName')
        }        
    }

    add_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]$handler)
    {
        $this.psobject.PropertyChanged.Add($handler)
    }
    remove_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]$handler)
    {
        $this.psobject.PropertyChanged.Remove($handler)
    }

    NotifyPropertyChanged([string]$propname)
    {
        if ($this.psobject.PropertyChanged)
        {
            $evargs = [System.ComponentModel.PropertyChangedEventArgs]::new($propname)
            $this.psobject.PropertyChanged.Invoke($this, $evargs) # invokes every member
        }
    }
}

class IntuneExportPolicy : IntuneImportExportPolicyBase
{
    [string]$_ExportFolder = $null
    [Boolean]$_AddExportAssignments = $false

    IntuneExportPolicy() : Base()
    {
        $this.PSObject._ExportFolder = (?? (Get-SettingStoreValue "" "LastUsedRoot") (Get-SettingValue "RootFolder"))
        $this.PSObject._AddExportAssignments = (Get-SettingValue "ExportAssignments")

        Add-ObjectProperty $this "ExportFolder" { return $this.PSObject._ExportFolder } ({ 
            param([String]$Value)
            $this.psobject._ExportFolder = $Value;
            $this.psobject.NotifyPropertyChanged('ExportFolder')
        })
            
        Add-ObjectProperty $this "AddExportAssignments" { return $this.PSObject._AddExportAssignments } { 
            param([Boolean]$Value)
            $this.psobject._AddExportAssignments = $Value;
            $this.psobject.NotifyPropertyChanged('AddExportAssignments')
        }
    }
}
#>

class IntuneManagerExportSettings
{
    [Boolean]$AddObjectType = $true
    [Boolean]$AddCompanyName = $true
    [string]$Filter = $null
    [Boolean]$ExportAssignments = $false
    # Depth of group-membership recursion. 1 (default) = only the group directly
    # assigned to the policy. 2 = also exports groups that are direct members
    # of the assigned group. 3 = one more level. Cycle-safe via the existing
    # per-folder MigrationTable dedup index.
    [Int]$ExportNestedGroupLevels = 1

    Hidden [string]$ExportFolder = $null

    IntuneManagerExportSettings()
    {
        $this.AddObjectType = (Get-SettingValue "AddObjectType")
        $this.AddCompanyName = (Get-SettingValue "AddCompanyName")
        $this.ExportAssignments = (Get-SettingValue "ExportAssignments")
        $this.ExportFolder = (?? (Get-SettingStoreValue "" "LastUsedRoot") (Get-SettingValue "RootFolder"))

        # Setting is registered as String so the form can use a free-form
        # editor; coerce to Int with a floor of 1 here so consumers can treat
        # it as a plain integer.
        $rawDepth = Get-SettingValue "ExportNestedGroupLevels" "1"
        $parsed = 0
        if([int]::TryParse([string]$rawDepth, [ref]$parsed) -and $parsed -ge 1) {
            $this.ExportNestedGroupLevels = $parsed
        } else {
            $this.ExportNestedGroupLevels = 1
        }
    }

    Save()
    {
        Save-SettingStoreValue "" "LastUsedRoot" $this.ExportFolder
    }
}

class IntuneManagerImportSettings 
{
    [string]$ImportFolder = $null
    [String]$MigrationTableInfo = $false
    [Boolean]$ImportAssignments = $false
    [Boolean]$ImportScopeTags = $false
    [Boolean]$ReplaceDependencyIDs = $false
    [Boolean]$SameTenant = $false
    [String]$ImportType = $null

    IntuneManagerImportSettings() 
    {
        $path = Get-SettingStoreValue "" "LastUsedFullPath"
        $this.ImportFolder =(?? $path (Get-SettingValue "RootFolder"))
        $this.ImportAssignments = (Get-SettingValue "ImportAssignments")
        $this.ImportScopeTags = (Get-SettingValue "ImportScopeTags")
        $this.ImportType = (Get-SettingValue "ImportType" "alwaysImport")
        $this.ReplaceDependencyIDs = (Get-SettingValue "ResolveReferenceInfo") -ne $false
    }

    Save()
    {
        # Persist the per-import toggles to the SAME SubPath they are registered
        # under ("IntuneManager") so the import pipeline reads them back via
        # Get-SettingValue. (Was a broken copy of the export Save() that wrote a
        # non-existent $this.ExportFolder.)
        Save-SettingStoreValue "IntuneManager" "ImportAssignments"    $this.ImportAssignments
        Save-SettingStoreValue "IntuneManager" "ImportScopeTags"      $this.ImportScopeTags
        Save-SettingStoreValue "IntuneManager" "ResolveReferenceInfo" $this.ReplaceDependencyIDs
        if($this.ImportType) { Save-SettingStoreValue "IntuneManager" "ImportType" $this.ImportType }
    }
}

# State holder for Bulk Scope Tags. UI binds to it; Set-GraphBulkScopeTags
# reads it. Kept separate from IntuneManagerImportSettings because the two
# operations have unrelated knobs and a shared class would force every UI to
# carry the union.
class IntuneManagerScopeTagSettings
{
    # Action applied to every selected policy:
    #   Add     — union of current tag list + ScopeTagIds (no-op if all already present)
    #   Replace — overwrite current list with ScopeTagIds
    #   Remove  — remove ScopeTagIds from the current list
    [string]$Action = "Add"

    # IDs of the scope tags the user chose in the picker. May be empty when
    # CleanupOrphans is the only thing the user wants done.
    [String[]]$ScopeTagIds = @()

    # Name filter — matched as regex (falls back to literal substring if not
    # a valid regex). Empty = touch every listed policy.
    [string]$Filter = $null

    # When true, any tag id on a policy that doesn't resolve to a real tag in
    # the tenant's current catalogue is stripped from the resulting list. Lets
    # users clean up references to deleted tags as a side-effect of any action,
    # or run a pure cleanup pass (Action=Add with no ScopeTagIds).
    [Boolean]$CleanupOrphans = $false
}

# State holder for Bulk Assignments (Phase 1: group / exclusion-group / all-
# devices / all-users targets with optional assignment filters; no app intent
# or per-platform settings yet — those land in later phases).
class IntuneManagerAssignmentSettings
{
    # Action applied to every selected policy:
    #   Add     — union of current assignments + chosen Assignments (no-op if
    #             all already present, compared by target type + groupId +
    #             filterId + filterType)
    #   Replace — overwrite current assignments with chosen Assignments
    #   Remove  — remove the chosen Assignments from the current list
    [string]$Action = "Add"

    # User-chosen assignments to apply. Each entry is a PSCustomObject:
    #   @{
    #     TargetType  = 'groupAssignmentTarget' | 'exclusionGroupAssignmentTarget'
    #                   | 'allDevicesAssignmentTarget' | 'allLicensedUsersAssignmentTarget'
    #     GroupId     = <string> (group/exclusion targets only)
    #     GroupName   = <string> (display only)
    #     FilterId    = <string> (optional, group/exclusion only)
    #     FilterName  = <string> (display only)
    #     FilterType  = 'include' | 'exclude' (when FilterId set)
    #   }
    # PSCustomObject[] (not a typed class) so the UI can mutate in place and
    # the public command can ingest without an extra projection step.
    [PSCustomObject[]]$Assignments = @()

    # Name filter — matched as regex (falls back to literal substring if not
    # a valid regex). Empty = touch every listed policy.
    [string]$Filter = $null
}
