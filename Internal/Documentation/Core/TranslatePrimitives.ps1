# Shared translate primitives used by phase-3 input providers.
#
# Ported from old Documentation.psm1 (Invoke-TranslateBoolean / Option / MultiOption /
# Table / Duration + Add-PropertyInfo + Add-NotConfiguredProperty + Get-PropertyInfo
# + ObjectInfo walker customization callbacks).
#
# Key design call: keep the OLD function signatures so input providers port
# mechanically. State that used to live in $script:objectSettingsData,
# $script:currentObject, $script:CurrentSubCategory, $script:propLevel etc. now
# lives on $script:_currentDocContext, a module-local [DocumentationContext]
# pointer the engine sets via Set-CurrentDocumentationContext before invoking
# any translate primitive. This mirrors the old $script:* pattern but with a
# single typed indirection — input providers and handlers don't need to know
# the context exists.

# ---- Module-local current-context indirection ----

$script:_currentDocContext = $null

function Set-CurrentDocumentationContext {
    param([DocumentationContext]$Context)
    $script:_currentDocContext = $Context
    # Drive Get-LanguageString off the requested documentation language so local
    # (non-Graph) strings localize too, not just the Accept-Language schema fetches.
    # Get-LanguageString reads $script:CurrentLanguage; without this it always used
    # 'en'. The public entry points save/restore this around the run.
    $script:CurrentLanguage = if ($Context -and $Context.Language) { $Context.Language } else { 'en' }
}

function Get-CurrentDocumentationContext {
    if (-not $script:_currentDocContext) {
        throw "No current documentation context. Call Set-CurrentDocumentationContext before invoking translate primitives."
    }
    return $script:_currentDocContext
}

# Accept-Language headers for the active documentation language (empty hashtable for
# en / none). Pass to Invoke-MSGraphAPI -AdditionalHeaders on any fetch that returns
# Microsoft-localized schema (ADMX definitions, categories, setting definitions) so
# the content matches the requested documentation language. Falls back to the current
# module context when no $Context is supplied.
function Get-DocAcceptLanguageHeaders {
    param([DocumentationContext]$Context)
    if (-not $Context) { $Context = $script:_currentDocContext }
    $lang = if ($Context) { $Context.Language } else { $null }
    if ($lang -and $lang -ne 'en') { return @{ 'Accept-Language' = $lang } }
    return @{}
}

# ---- Property accumulator helpers (replaces old Add-PropertyInfo et al.) ----

# Add a basic-info row (name/value pair shown in the policy's header table).
# EntityKey is the source field name on the raw object (e.g. 'displayName',
# 'state', 'createdDateTime'); used by compare and other downstream tools to
# key rows independently of the localized Name. Optional; callers that don't
# care about compare can omit it.
function Add-BasicPropertyValue {
    param(
        [string]$Name,
        [object]$Value,
        [string]$EntityKey
    )
    $ctx = Get-CurrentDocumentationContext
    $ctx.AddBasic($Name, $Value, $EntityKey)
}

# Build the rich PSCustomObject the old engine put into $script:objectSettingsData.
# Schema kept compatible with what output providers and handlers expect.
function Get-PropertyInfo {
    param($Prop, $Value, $OriginalValue, $JsonValue, $TableValue)

    # Get-LanguageString (Internal/LanguageString.ps1) can throw when a
    # nameResourceKey path resolves to a nested object instead of a leaf
    # string (e.g. a key like "WslCompliance" where WslCompliance is a
    # section, not a string). Guarding both calls preserves the row but
    # logs the bad key for follow-up.
    if ($Prop.nameResource) {
        $name = $Prop.nameResource
    }
    elseif ($Prop.nameResourceKey) {
        $key = if ($Prop.nameResourceKey.Contains('.')) { $Prop.nameResourceKey } else { "SettingDetails.$($Prop.nameResourceKey)" }
        try { $name = Get-LanguageString $key }
        catch { Write-Log "Get-LanguageString '$key' failed: $($_.Exception.Message)" 2; $name = $Prop.nameResourceKey }
    }
    else {
        $name = $Prop.entityKey
    }

    $description = ""
    if ($Prop.descriptionResource) {
        $description = $Prop.descriptionResource
    }
    elseif ($Prop.descriptionResourceKey) {
        # Via the shared resolver so renamed keys are redirected and numeric
        # metadata artifacts are skipped. See Get-ObjectInfoResourceString.
        $description = Get-ObjectInfoResourceString $Prop.descriptionResourceKey
        if ($null -eq $description) { $description = "" }
    }

    $categoryStr = $null
    if ($Prop.category) {
        # New project's category helper (old code used Get-Category — renamed)
        try { $categoryStr = Get-PolicyObjectCategoryString $Prop.category }
        catch { Write-Log "Get-PolicyObjectCategoryString '$($Prop.category)' failed: $($_.Exception.Message)" 2 }
    }

    if (-not $JsonValue -and $null -ne $OriginalValue -and "$OriginalValue" -ne "") {
        $JsonValue = $OriginalValue | ConvertTo-Json -Depth 50 -Compress
    }

    $defValue = $null
    if ($Prop.emptyValueResourceKey) {
        try { $defValue = Get-LanguageString $Prop.emptyValueResourceKey }
        catch { Write-Log "Get-LanguageString '$($Prop.emptyValueResourceKey)' failed: $($_.Exception.Message)" 2 }
    }
    else {
        $defValue = $Prop.defaultValue
    }

    $ctx = Get-CurrentDocumentationContext
    return [PSCustomObject]@{
        Name              = $name
        Description       = $description
        Value             = $Value
        Category          = $categoryStr
        SubCategory       = $ctx.CurrentSubCategory
        Property          = $Prop.entityKey
        DataType          = $Prop.dataType
        RawValue          = $OriginalValue
        RawJsonValue      = $JsonValue
        DefaultValue      = $defValue
        FullValueTable    = $TableValue
        UnconfiguredValue = $Prop.unconfiguredValue
        AlwaysAddValue    = $Prop.alwaysAddValue -eq $true
        Enabled           = $Prop.Enabled
        EntityKey         = $Prop.EntityKey
        # PropLevel uses -1 internally as a "reset on next recursion" sentinel
        # (set by dataType 8 sub-headers / dataType 5 groups). That sentinel must
        # never reach a row: a negative level renders as padding-left:0px, pushing
        # the setting name left of its category header. Clamp so top-level settings
        # align with the header (default cell padding) and only real child
        # settings (positive levels) indent.
        Level             = [Math]::Max(0, [int]$ctx.PropLevel)
    }
}

# Add a translated property to the context. Routes to BasicInfo when the
# prop is in the synthetic category 1000 (header info); otherwise to FilteredSettings.
function Add-PropertyInfo {
    param($Prop, $Value, $OriginalValue, $JsonValue, $TableValue)

    if ($Prop.Category -eq "1000") {
        $name = if ($Prop.nameResource) {
            $Prop.nameResource
        }
        elseif ($Prop.nameResourceKey) {
            $key = if ($Prop.nameResourceKey.Contains('.')) { $Prop.nameResourceKey } else { "SettingDetails.$($Prop.nameResourceKey)" }
            try { Get-LanguageString $key } catch { Write-Log "Get-LanguageString '$key' failed: $($_.Exception.Message)" 2; $Prop.nameResourceKey }
        }
        else { $Prop.entityKey }
        Add-BasicPropertyValue $name $Value
        return
    }

    $info = Get-PropertyInfo $Prop $Value $OriginalValue $JsonValue $TableValue
    $ctx  = Get-CurrentDocumentationContext
    $ctx.AddSetting($info)

    Invoke-CustomPostAddValue $Prop
}

# Pre-built PSCustomObject path (rarer — used when a custom handler has already
# constructed the row in the result shape).
function Add-PropertyInfoObject {
    param($PropInfo)
    if (-not $PropInfo) { return }
    $ctx = Get-CurrentDocumentationContext
    $ctx.AddSetting($PropInfo)
}

# Custom handlers (Conditional Access, PolicySet, etc.) build their own row
# PSCustomObjects directly (no $prop metadata indirection) and add them via this
# helper. Old code: Add-CustomSettingObject in Documentation.psm1:4041.
function Add-CustomSettingObject {
    param($SettingsObj)
    if (-not $SettingsObj) { return }
    $ctx = Get-CurrentDocumentationContext
    $ctx.AddSetting($SettingsObj)
}

# Adds the standard "Created" / "Modified" BasicInfo rows. Reads dates through
# the wrapper's `.Created` / `.LastModified` script properties (which already
# handle the lastModifiedDateTime/modifiedDateTime fallback). Old code:
# Add-BasicAdditionalValues in Documentation.psm1:697.
function Add-BasicAdditionalValues {
    param($PolicyObject)
    # Per-handler private helpers (e.g. Invoke-CADocBasicInfo) don't always
    # have $PolicyObject in scope when they fan out to Add-BasicAdditionalValues;
    # the engine has already pinned the wrapper to the context so we just
    # read it back. Explicit param still wins for callers that have it.
    if (-not $PolicyObject) { $PolicyObject = (Get-CurrentDocumentationContext).PolicyObject }
    if (-not $PolicyObject) { return }
    $obj = $PolicyObject.JsonObject

    if ($PolicyObject.Created -is [datetime]) {
        Add-BasicPropertyValue (Get-LanguageString 'Inputs.createdDateTime') (Format-BasicDateValue $PolicyObject.Created) 'createdDateTime'
    }
    if ($PolicyObject.LastModified -is [datetime]) {
        # Note: old engine uses TableHeaders.lastModified (= "Last modified") here,
        # NOT Inputs.lastModifiedDateTime (which resolves to "Last updated Time").
        # Entity-key reflects which of the two underlying fields was used so
        # downstream tools that key by entityKey still see the right name.
        $modKey = if ($obj.lastModifiedDateTime) { 'lastModifiedDateTime' } else { 'modifiedDateTime' }
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.lastModified') (Format-BasicDateValue $PolicyObject.LastModified) $modKey
    }

    # Version row (when the object has a numeric version, e.g. enrollment
    # restriction configurations). Truthy check skips 0 (system-created defaults).
    if ($obj.version) {
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.eDPPolicyAppsListVersionName') $obj.version 'version'
    }
}

# Formats a [datetime] the way the old engine did: long date + long time, in
# local time. Kept separate so any downstream change to the date format only
# touches one place.
function Format-BasicDateValue {
    param([datetime]$Value)
    $tmp = if ($Value.Kind -eq 'Utc') { $Value.ToLocalTime() } else { $Value }
    return "$($tmp.ToLongDateString()) $($tmp.ToLongTimeString())"
}

# Emits the conventional BasicInfo header rows that almost every handler / input
# provider starts with: Name, Description, plus Platform-supported / Profile-type
# from the ObjectCategories.json lookup. Old code: Add-BasicDefaultValues at
# Documentation.psm1:607.
function Add-BasicDefaultValues {
    param($PolicyObject, [string]$ProfileTypeName, [string[]]$SkipProperties = @())
    # Context fallback (see Add-BasicAdditionalValues for rationale).
    if (-not $PolicyObject) { $PolicyObject = (Get-CurrentDocumentationContext).PolicyObject }
    if (-not $PolicyObject) { return }
    $obj = $PolicyObject.JsonObject

    # Name: routed through the wrapper's GetName(), which resolves the per-type
    # _NameProperty (e.g. `fileName` for ADMX, `displayName` for most). Avoids
    # the displayName-vs-name guessing the helper used to do.
    $nameProp = if ($PolicyObject.PolicyType -and $PolicyObject.PolicyType._NameProperty) {
        $PolicyObject.PolicyType._NameProperty
    } else { 'displayName' }
    Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.nameName') $PolicyObject.Name $nameProp

    # Always emit a Description row — old engine includes one with value=""
    # even for objects (e.g. Named Locations) that don't have a description
    # property in their raw Graph payload at all.
    $descValue = if ($obj.description) { $obj.description } else { '' }
    Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.descriptionName') $descValue 'description'

    # Platform-supported + Profile-type rows from Config/ObjectCategories.json
    # lookup. If the type isn't catalogued (e.g. assignment filters, named
    # locations), these stay silent and the handler emits its own Profile-type
    # row without duplication.
    $odata = $obj.'@odata.type'
    if ($odata -and (Get-Command Get-PolicyObjectCategoryInfo -ErrorAction SilentlyContinue)) {
        try {
            $objInfo = Get-PolicyObjectCategoryInfo $odata
            if ($objInfo) {
                if ($objInfo.PlatformLanguageId -and -not $SkipProperties.Contains('platformSupported')) {
                    $platformType = Get-LanguageString "Platform.$($objInfo.PlatformLanguageId)"
                    if ($platformType) {
                        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.platformSupported') $platformType 'platformSupported'
                    }
                }
                $profileType = if ($ProfileTypeName) { $ProfileTypeName }
                               elseif ($objInfo.PolicyType) { Get-LanguageString "ConfigurationTypes.$($objInfo.PolicyType)" }
                               else { $null }
                if ($profileType -and -not $SkipProperties.Contains('@odata.type')) {
                    Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') $profileType '@odata.type'
                }
            }
        } catch { }
    }
}

# Resolves roleScopeTagIds / roleScopeTags arrays to display names and appends
# them as a single comma-joined BasicInfo row. Cache lives on the context
# ($ctx.ScopeTags); populated lazily on first call. Old code: Add-ScopeTagStrings
# at Documentation.psm1:797, with caching at Documentation.psm1:214.
function Add-ScopeTagStrings {
    param($Obj)
    if (-not $Obj) { return }

    $prop = $null
    if ($Obj.PSObject.Properties['roleScopeTagIds']) { $prop = 'roleScopeTagIds' }
    elseif ($Obj.PSObject.Properties['roleScopeTags']) { $prop = 'roleScopeTags' }
    if (-not $prop) { return }

    $ids = @($Obj.$prop)
    if ($ids.Count -eq 0) { return }

    $ctx = Get-CurrentDocumentationContext

    # Lazy cache population — source-tenant-specific scope tags; skipped when the
    # source tenant is unavailable (another tenant's tags would be wrong).
    if ((-not $ctx.ScopeTags -or $ctx.ScopeTags.Count -eq 0) -and -not $ctx.SourceTenantUnavailable) {
        if (Test-DocumentationGraphAvailable) {
            try {
                $resp = Invoke-MSGraphAPI -Url '/deviceManagement/roleScopeTags'
                if ($resp.Value) { $ctx.ScopeTags = $resp.Value }
            }
            catch {
                Write-LogError 'Failed to load scope tags for documentation' $_.Exception
            }
        }
    }

    $names = foreach ($id in $ids) {
        if ($id -eq '0') {
            Get-LanguageString 'SettingDetails.default'
        }
        elseif ($ctx.ScopeTags) {
            $tag = $ctx.ScopeTags | Where-Object Id -eq $id | Select-Object -First 1
            if ($tag -and $tag.displayName) { $tag.displayName } else { $id }
        }
        else {
            $id
        }
    }

    Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.scopeTags') (($names) -join $ctx.ObjectSeparator) 'roleScopeTagIds'
}

# Translate $obj.assignments into structured Assignment rows on the context.
# Port of old Invoke-TranslateAssignments (Documentation.psm1:3610, ~370 LOC).
# Trimmed for v1:
#   - Generic assignments: full parity (group/include/exclude/filter rows)
#   - App (mobileApp) assignments: Group + Intent + Filter; per-intent extra
#     settings (restart, install-time, notifications, deliveryOptimization, ...)
#     deferred — those add ~150 LOC of per-prop translation and need their own
#     fixture coverage before they're worth porting. Apps still document
#     correctly, just without the extra-settings block.
#
# Honors $ctx.Options.ExcludeAssignments (default false).
# Offline-safe: group/filter resolution falls back to ID when no Graph access.
# Resolve Entra group IDs -> displayName via /directoryObjects/getByIds, chunked
# at 1000 (the Graph batch-getByIds limit). Populates the run cache
# Context.GroupNamesById, storing $null for IDs the directory did not return so
# unresolved IDs are not retried on later policies (negative cache, matching the
# assignment-filter pattern). Only IDs not already cached are queried, so the
# opt-in run preload and the per-object lazy path share one cache without
# double-fetching. Returns a hashtable of the requested IDs that resolved to a
# non-empty name.
function Resolve-DocumentationGroupNames {
    param(
        [string[]]$GroupIds,
        [DocumentationContext]$Context
    )
    $out = @{}
    if (-not $Context -or -not $GroupIds) { return $out }

    $toQuery = @($GroupIds | Where-Object { $_ -and -not $Context.GroupNamesById.ContainsKey($_) } | Select-Object -Unique)
    for ($i = 0; $i -lt $toQuery.Count; $i += 1000) {
        $end   = [Math]::Min($i + 999, $toQuery.Count - 1)
        $chunk = @($toQuery[$i..$end])
        $returned = @{}
        try {
            $body = (@{ ids = $chunk; types = @('group') } | ConvertTo-Json -Compress)
            $resp = Invoke-MSGraphAPI -Url '/directoryObjects/getByIds?$select=displayName,id' -Content $body -HttpMethod 'POST'
            foreach ($g in @($resp.value)) {
                if ($g.id) { $returned[[string]$g.id] = [string]$g.displayName }
            }
        }
        catch {
            Write-LogError 'Failed to resolve assignment groups via /directoryObjects/getByIds' $_.Exception
            # Leave this chunk uncached so a transient failure can be retried
            # later (do NOT negative-cache on error).
            continue
        }
        foreach ($id in $chunk) {
            if ($returned.ContainsKey($id) -and $returned[$id]) {
                $Context.GroupNamesById[$id] = $returned[$id]
            }
            else {
                $Context.GroupNamesById[$id] = $null   # looked up, not found - don't retry
            }
        }
    }

    foreach ($id in $GroupIds) {
        if ($Context.GroupNamesById.ContainsKey($id) -and $Context.GroupNamesById[$id]) {
            $out[$id] = $Context.GroupNamesById[$id]
        }
    }
    return $out
}

function Add-AssignmentsForObject {
    param($Obj)
    if (-not $Obj) { return }

    $ctx = Get-CurrentDocumentationContext
    if ($ctx.Options.ExcludeAssignments) { return }

    $assignments = @($Obj.assignments)
    if ($assignments.Count -eq 0) { return }

    # Collect unique IDs so we resolve in batch (one Graph call, not N).
    # The all-zeros GUID is Intune's "no filter" sentinel; never a real
    # filter id. Anything that doesn't shape like a GUID would produce a
    # malformed URL on the per-id resolution path — drop those too so the
    # batch lookup below is always safe to interpolate.
    $groupIds  = @($assignments.target.groupId | Where-Object { $_ } | Select-Object -Unique)
    $filterIds = @($assignments.target.deviceAndAppManagementAssignmentFilterId |
                   Where-Object {
                       (Test-AssignmentFilterDefined $_) -and
                       ($_ -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
                   } |
                   Select-Object -Unique)

    # ---- Resolve groups ----
    # Cache hits first: the opt-in run preload (Sync-DocumentationGroupPreload)
    # and earlier policies this run populate $ctx.GroupNamesById. Only IDs still
    # missing hit Graph, and Resolve-DocumentationGroupNames caches the result so
    # later policies referencing the same group are free. When the preload ran,
    # every ID is already a hit and no per-object Graph call happens here.
    $groupNamesById = @{}
    if ($groupIds.Count -gt 0) {
        $missing = @($groupIds | Where-Object { -not $ctx.GroupNamesById.ContainsKey($_) })
        foreach ($gid in $groupIds) {
            if ($ctx.GroupNamesById.ContainsKey($gid) -and $ctx.GroupNamesById[$gid]) {
                $groupNamesById[$gid] = $ctx.GroupNamesById[$gid]
            }
        }
        if ($missing.Count -gt 0 -and -not $ctx.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
            $resolved = Resolve-DocumentationGroupNames -GroupIds $missing -Context $ctx
            foreach ($gid in $resolved.Keys) { $groupNamesById[$gid] = $resolved[$gid] }
        }
        # Offline fallback: leave IDs unresolved — output providers display the
        # raw GUID, matching old engine behavior when the migration table is
        # unavailable.
    }

    # ---- Resolve filters ----
    # Normally the run prefetch (Initialize-DocumentationRunPrefetch) has
    # already seeded $ctx.FilterNamesById — from the login-time tenant
    # dependency cache or its batched tenant-wide GET — and set FiltersLoaded.
    # This lazy fallback covers the single-policy Get-GraphDocumentation path.
    # One tenant-wide list GET, not per-id: the Intune beta endpoint doesn't
    # honour `$filter=id in (...)` (the StatelessPayloadLinkingService proxy
    # 400s the OData syntax). Same shape as Get-CDAllTenantApps below.
    if ($filterIds.Count -gt 0 -and -not $ctx.SourceTenantUnavailable -and -not $ctx.FiltersLoaded -and (Test-DocumentationGraphAvailable)) {
        try {
            $resp = Invoke-MSGraphAPI -Url '/deviceManagement/assignmentFilters?$select=id,displayName&$top=999' -ODataMetadata 'minimal'
            foreach ($f in @($resp.value)) {
                if ($f.id) { $ctx.FilterNamesById[$f.id] = $f.displayName }
            }
        }
        catch {
            Write-LogError 'Failed to resolve assignment filters' $_.Exception
        }
        # Stamp regardless of success so a tenant-wide 5xx doesn't trigger a
        # retry per policy (cache acts as both lookup AND negative-cache).
        $ctx.FiltersLoaded = $true
    }
    $filterNamesById = @{}
    foreach ($id in $filterIds) {
        if ($ctx.FilterNamesById.ContainsKey($id) -and $ctx.FilterNamesById[$id]) {
            $filterNamesById[$id] = $ctx.FilterNamesById[$id]
        }
    }

    # ---- Translate each assignment ----
    $includeLabel = Get-LanguageString 'TableHeaders.includedGroups'
    $excludeLabel = Get-LanguageString 'TableHeaders.excludedGroups'
    $noFilter     = Get-LanguageString 'AssignmentFilters.noFilters'
    $filterInc    = Get-LanguageString 'SettingDetails.include'
    $filterExc    = Get-LanguageString 'SettingDetails.exclude'

    $included = @()
    $excluded = @()
    $appRows  = @()

    foreach ($assignment in $assignments) {
        if (-not $assignment -or -not $assignment.target) { continue }

        $isAppAssignment = ($assignment.PSObject.Properties['intent'] -and
                            $assignment.PSObject.Properties['settings'])

        $groupMode = if ($assignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
            'exclude'
        } else { 'include' }

        $groupName = switch ($assignment.target.'@odata.type') {
            '#microsoft.graph.allDevicesAssignmentTarget'        { Get-LanguageString 'SettingDetails.allDevices' }
            '#microsoft.graph.allLicensedUsersAssignmentTarget'  { Get-LanguageString 'SettingDetails.allUsers' }
            default {
                $gid = [string]$assignment.target.groupId
                if ($gid -and $groupNamesById.ContainsKey($gid)) { $groupNamesById[$gid] }
                elseif ($gid) { $gid }
                else { 'unknown' }
            }
        }

        $filterName = $null; $filterMode = $null
        if ($assignment.target.PSObject.Properties['deviceAndAppManagementAssignmentFilterId']) {
            $filterName = $noFilter; $filterMode = $noFilter
            $fid = [string]$assignment.target.deviceAndAppManagementAssignmentFilterId
            # Test-AssignmentFilterDefined, not a truthy check: the all-zeros sentinel
            # means "no filter" and keeps the $noFilter labels. Reporting it verbatim
            # printed a filter named 00000000-0000-0000-0000-000000000000 in Exclude
            # mode on assignments that cannot carry a filter at all.
            if (Test-AssignmentFilterDefined $fid) {
                $filterName = if ($filterNamesById.ContainsKey($fid)) { $filterNamesById[$fid] } else { $fid }
                $filterMode = if ($assignment.target.deviceAndAppManagementAssignmentFilterType -eq 'include') { $filterInc } else { $filterExc }
            }
        }

        if ($isAppAssignment) {
            $row = @{
                Group        = $groupName
                GroupMode    = Get-LanguageString "AssignmentAction.$groupMode"
                Category     = Get-LanguageString "InstallIntent.$($assignment.intent)"
                RawIntent    = $assignment.intent
                Type         = 'AppAssignment'
                RawJsonValue = ($assignment | ConvertTo-Json -Depth 50 -Compress)
            }
            if ($groupMode -eq 'include') {
                $row['Filter']     = $filterName
                $row['FilterMode'] = $filterMode

                # Per-intent extra settings (port of old Documentation.psm1:3776-3895).
                # Only Included app assignments carry these. Output providers turn each
                # key of the ordered 'Settings' dictionary into a column.
                if ($null -ne $assignment.settings) {
                    $settingsProps = [ordered]@{}
                    # Ordered to match the Intune portal's assignment columns (iOS VPP):
                    # VPN, License type, Prevent automatic app updates, Uninstall on device
                    # removal, Install as removable, Prevent iCloud app backup. Non-portal /
                    # other-app-type keys follow.
                    foreach ($settingProp in @('useDeviceContext','vpnConfigurationId','useDeviceLicensing','preventAutoAppUpdate',
                                               'uninstallOnDeviceRemoval','isRemovable','preventManagedAppBackup',
                                               'androidManagedStoreAppTrackIds','deliveryOptimizationPriority',
                                               'installTimeSettings','notifications','restartSettings')) {
                        if (-not ($assignment.settings.PSObject.Properties | Where-Object Name -EQ $settingProp)) { continue }
                        $sVal = $assignment.settings.$settingProp

                        if ($settingProp -eq 'useDeviceLicensing') {
                            $value = if ($sVal -eq $true) { Get-LanguageString 'SettingDetails.licenseTypeDevice' }
                                     else { Get-LanguageString 'SettingDetails.licenseTypeUser' }
                        }
                        elseif ($settingProp -eq 'restartSettings') {
                            if ($null -eq $sVal) { $value = Get-LanguageString 'SettingDetails.disabledOption' }
                            else {
                                $arr = @()
                                $arr += "$(Get-LanguageString 'Assignment.RestartGracePeriod.durationInMinutes')=$($sVal.gracePeriodInMinutes)"
                                $arr += "$(Get-LanguageString 'Assignment.RestartGracePeriod.countdownDialog')=$($sVal.countdownDisplayBeforeRestartInMinutes)"
                                if ($null -eq $sVal.restartNotificationSnoozeDurationInMinutes) {
                                    $arr += "$(Get-LanguageString 'Assignment.RestartGracePeriod.allowSnooze')=$(Get-LanguageString 'SettingDetails.no')"
                                } else {
                                    $arr += "$(Get-LanguageString 'Assignment.RestartGracePeriod.allowSnooze')=$(Get-LanguageString 'SettingDetails.yes')"
                                    $arr += "$(Get-LanguageString 'Assignment.RestartGracePeriod.snoozeDurationInMinutes')=$($sVal.restartNotificationSnoozeDurationInMinutes)"
                                }
                                $value = $arr -join $ctx.ObjectSeparator
                            }
                        }
                        elseif ($settingProp -eq 'notifications') {
                            $value = Get-LanguageString "AppResources.AssignmentToast.$sVal"
                            if (-not $value) { $value = $sVal }
                        }
                        elseif ($settingProp -eq 'installTimeSettings') {
                            $asap = Get-LanguageString 'Assignment.SoftwareInstallationTime.defaultTime'
                            $startValue = $asap; $value = $asap
                            if ($sVal) {
                                if ($sVal.startDateTime) {
                                    $instTime = Get-Date $sVal.startDateTime
                                    if ($sVal.useLocalTime -eq $false) { $instTime = $instTime.AddHours(($instTime.ToUniversalTime() - $instTime).Hours) }
                                    $startValue = "$($instTime.ToShortDateString()) $($instTime.ToShortTimeString())"
                                }
                                if ($sVal.deadlineDateTime) {
                                    $endTime = Get-Date $sVal.deadlineDateTime
                                    if ($sVal.useLocalTime -eq $false) { $endTime = $endTime.AddHours(($endTime.ToUniversalTime() - $endTime).Hours) }
                                    $value = "$($endTime.ToShortDateString()) $($endTime.ToShortTimeString())"
                                }
                            }
                            $settingsProps['startTimeColumnLabel'] = $startValue
                            if ($assignment.intent -eq 'available') { continue }   # no deadline column on available
                        }
                        elseif ($settingProp -eq 'deliveryOptimizationPriority') {
                            $tmpStr  = Get-LanguageString 'AppResources.DeliveryOptimizationPriority.displayText'
                            $tmpType = if ($sVal -ne 'foreground') { Get-LanguageString 'AppResources.DeliveryOptimizationPriority.backgroundNormal' }
                                       else { Get-LanguageString 'AppResources.DeliveryOptimizationPriority.foreground' }
                            $value = $tmpStr -f $tmpType
                        }
                        elseif ("$sVal" -eq 'notConfigured') { $value = Get-LanguageString 'BooleanActions.notConfigured' }
                        else { $value = $sVal }

                        $settingsProps[$settingProp] = $value
                    }
                    if ($settingsProps.Count -gt 0) { $row['Settings'] = $settingsProps }
                }
            }
            $appRows += [PSCustomObject]$row
        }
        else {
            $row = [PSCustomObject]@{
                GroupMode = if ($groupMode -eq 'include') { $includeLabel } else { $excludeLabel }
                Group     = $groupName
                Type      = 'GenericAssignment'
                Category  = if ($groupMode -eq 'include') { $includeLabel } else { $excludeLabel }
            }
            if ($groupMode -eq 'include' -and $null -ne $filterMode) {
                $row | Add-Member -MemberType NoteProperty -Name 'Filter'     -Value $filterName -Force
                $row | Add-Member -MemberType NoteProperty -Name 'FilterMode' -Value $filterMode -Force
            }
            if ($groupMode -eq 'include') { $included += $row } else { $excluded += $row }
        }
    }

    foreach ($r in $included) { $ctx.AddAssignment($r) }
    foreach ($r in $excluded) { $ctx.AddAssignment($r) }

    # Sort app rows by intent order: required -> available -> availableWithoutEnrollment -> uninstall
    foreach ($intent in @('required','available','availableWithoutEnrollment','uninstall')) {
        foreach ($r in ($appRows | Where-Object RawIntent -EQ $intent)) {
            $ctx.AddAssignment($r)
        }
    }
}

# Track properties that couldn't be translated (booleanActions miss, unknown option,
# notConfigured fallback). Used downstream for diagnostic output.
function Add-NotConfiguredProperty {
    param($Prop)
    $ctx = Get-CurrentDocumentationContext
    $ctx.UnconfiguredProperties += $Prop
}

# ---- ObjectInfo walker customization callbacks ----
#
# Schema-driven Manifest/Profile translation uses a separate customizer registry.
# Whole-object handlers cannot provide these callbacks because registering one
# bypasses the ObjectInfo input providers entirely.

# ---- AppConfig / CustomOMAUri helpers ----

# Infers a Platform.* language id from the object's @odata.type. Used by
# handlers that need a "Platform: iOS" / "Mac" / "Windows10" basic info row.
# Faithful port of old Documentation.psm1:879. Note: the old code's
# Contains("androidForWork") branch is unreachable (it's a case-sensitive
# match against an already-lowercased string); preserving for parity.
function Get-ObjectPlatformFromType {
    param($Obj)
    if (-not $Obj.'@OData.Type' -and -not $Obj.'@odata.type') { return $null }
    $t = (($Obj.'@OData.Type'),($Obj.'@odata.type') | Where-Object { $_ })[0].ToLower()

    if     ($t.Contains('ios'))         { 'iOS' }
    elseif ($t.Contains('mac'))         { 'Mac' }
    elseif ($t.Contains('windowsphone')){ 'WindowsPhone' }
    elseif ($t.Contains('windows') -or $t.Contains('win32') -or $t.Contains('mirosoftstore')) { 'Windows10' }
    elseif ($t.Contains('androidforwork')) { 'androidForWork' }
    elseif ($t.Contains('android'))     { 'Android' }
    else { $null }
}

# Returns the in-progress per-object settings buffer. Handlers use this to
# avoid double-adding settings the ObjectInfo walker has already emitted
# (matches old Documentation.psm1:386 Get-DocumentedSettings).
function Get-DocumentedSettings {
    $ctx = Get-CurrentDocumentationContext
    return @($ctx.SettingsData)
}

# App-catalog lookup for AppConfig handlers. Live: /deviceAppManagement/mobileApps
# (paged). Offline: returns empty list so handlers fall through to raw app IDs.
# Cached on $ctx via a hashtable extension on the singleton.
function Get-CDAllTenantApps {
    $ctx = Get-CurrentDocumentationContext
    if (-not $ctx.PSObject.Properties['_AllTenantApps']) {
        $ctx | Add-Member -MemberType NoteProperty -Name '_AllTenantApps' -Value $null -Force
    }
    if ($ctx._AllTenantApps) { return $ctx._AllTenantApps }
    if ($ctx.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) {
        $ctx._AllTenantApps = @()
        return $ctx._AllTenantApps
    }
    try {
        $resp = Invoke-MSGraphAPI -Url '/deviceAppManagement/mobileApps?$select=displayName,id&$top=999'
        $ctx._AllTenantApps = @($resp.Value)
    } catch {
        Write-LogError 'Failed to load /deviceAppManagement/mobileApps' $_.Exception
        $ctx._AllTenantApps = @()
    }
    return $ctx._AllTenantApps
}

# Notification message templates, for enrollment notifications. An enrollment
# notification policy stores only a template id per channel; the subject and body
# a reader actually wants are the template's localized messages. $expand pulls
# every template with its messages in ONE request, cached for the whole run, so
# documenting N notification policies costs one call rather than 2N.
# Offline: empty list, and the caller renders the channel without its text.
#
# Two things about the cache. It is keyed by the policy's OWN token, because with
# two tenants signed in the active provider may belong to the other one, and its
# templates would document this policy's messages as "Not configured". And it
# lives for one run only: Reset-DocumentationTenantLookups drops it at the start
# of each bulk run and whenever a run flips to offline, so a message edited
# between two runs is read fresh and an offline run never answers from what a
# live one cached. The offline guard runs before the cache is consulted for the
# same reason.
function Get-CDNotificationMessageTemplates {
    $ctx = Get-CurrentDocumentationContext
    if ($ctx.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) { return @() }

    if (-not $ctx.PSObject.Properties['_NotificationMessageTemplates']) {
        $ctx | Add-Member -MemberType NoteProperty -Name '_NotificationMessageTemplates' -Value $null -Force
    }
    if ($ctx._NotificationMessageTemplates -isnot [hashtable]) { $ctx._NotificationMessageTemplates = @{} }

    # A policy loaded from a live tenant carries the token it came from; one
    # loaded from a file carries 0 and a test fixture nothing, and both mean
    # "the default token". The default is whatever is signed in NOW, and that
    # changes when the user switches tenants or accounts between two
    # single-object calls, which keep this cache. So the key is the CONNECTION
    # the token resolves to - tenant, account and app, plus the resolved id -
    # never the bare token id: a "0" entry would go on answering for whatever
    # happened to be default when it was filled. Nor the tenant alone: two
    # accounts in one tenant see different templates once scope tags apply, and
    # an app-only token and a user token need not be authorized alike.
    $tokenId = 0
    if ($ctx.PolicyObject -and $null -ne $ctx.PolicyObject._TokenId) { $tokenId = [int]$ctx.PolicyObject._TokenId }
    $tokenInfo = $null
    try { $tokenInfo = Get-OperationTokenInfo $tokenId } catch { }
    if ($tokenInfo -and [int]$tokenInfo.Id -gt 0) { $tokenId = [int]$tokenInfo.Id }
    $cacheKey = if ($tokenInfo) { "$tokenId|$($tokenInfo.TenantID)|$($tokenInfo.User)|$($tokenInfo.ClientID)" } else { "token:$tokenId" }
    if ($ctx._NotificationMessageTemplates.ContainsKey($cacheKey)) { return $ctx._NotificationMessageTemplates[$cacheKey] }

    try {
        $params = @{ Url = '/deviceManagement/notificationMessageTemplates?$expand=localizedNotificationMessages' }
        if ($tokenId -gt 0) { $params.TokenId = $tokenId }
        $resp = Invoke-MSGraphAPI @params
        $ctx._NotificationMessageTemplates[$cacheKey] = @($resp.Value)
    }
    catch {
        Write-LogError 'Failed to load /deviceManagement/notificationMessageTemplates' $_.Exception
        $ctx._NotificationMessageTemplates[$cacheKey] = @()
    }
    return $ctx._NotificationMessageTemplates[$cacheKey]
}

# Managed-app catalog lookup. The managedAppStatuses('managedAppList') endpoint
# is the same source used by the old documentation provider. Offline callers can
# pre-seed _AllManagedApps on the context from exported fixture data.
function Get-CDAllManagedApps {
    $ctx = Get-CurrentDocumentationContext
    if (-not $ctx.PSObject.Properties['_AllManagedApps']) {
        $ctx | Add-Member -MemberType NoteProperty -Name '_AllManagedApps' -Value $null -Force
    }
    if ($ctx._AllManagedApps) { return $ctx._AllManagedApps }
    if ($ctx.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) {
        $ctx._AllManagedApps = @()
        return $ctx._AllManagedApps
    }
    try {
        $resp = Invoke-MSGraphAPI -Url "/deviceAppManagement/managedAppStatuses('managedAppList')"
        $ctx._AllManagedApps = @($resp.content.appList)
    }
    catch {
        Write-LogError "Failed to load managed app catalog" $_.Exception
        $ctx._AllManagedApps = @()
    }
    return $ctx._AllManagedApps
}

# Given a policy's $obj.Apps array, partitions display names into
# (customApps, publishedApps) based on Microsoft-first-party flag.
# Old code: DocumentationCustom.psm1:207. Offline returns ([],[]).
function Get-CDMobileApps {
    param($Apps)
    if (-not $Apps) { return @(@(), @()) }
    $managed = Get-CDAllManagedApps
    $customApps = @()
    $publishedApps = @()
    foreach ($tmpApp in $Apps) {
        $appObj = $managed | Where-Object {
            $pkgMatch  = $tmpApp.mobileAppIdentifier.packageId    -and $_.appIdentifier.packageId    -eq $tmpApp.mobileAppIdentifier.packageId
            $bundMatch = $tmpApp.mobileAppIdentifier.bundleId     -and $_.appIdentifier.bundleId     -eq $tmpApp.mobileAppIdentifier.bundleId
            $winMatch  = $tmpApp.mobileAppIdentifier.windowsAppId -and $_.appIdentifier.windowsAppId -eq $tmpApp.mobileAppIdentifier.windowsAppId
            $typeMatch = $_.appIdentifier.'@odata.type' -eq $tmpApp.mobileAppIdentifier.'@odata.type'
            ($pkgMatch -or $bundMatch -or $winMatch) -and $typeMatch
        } | Select-Object -First 1
        if ($appObj -and $appObj.isFirstParty) { $publishedApps += $appObj.displayName }
        elseif ($appObj)                       { $customApps    += $appObj.displayName }
        else {
            # Not in the managed-app catalog - that IS what a custom app is (a
            # hand-entered bundle/package id). Previously these were dropped, so the
            # "Custom apps" row always rendered empty.
            $rawId = @($tmpApp.mobileAppIdentifier.packageId, $tmpApp.mobileAppIdentifier.bundleId, $tmpApp.mobileAppIdentifier.windowsAppId) | Where-Object { $_ } | Select-Object -First 1
            if ($rawId) { $customApps += $rawId }
        }
    }
    return @(,$customApps + ,$publishedApps)
}

# Append a row collection to $ctx.CustomTables (e.g. AdditionalSettings or
# Permissions tables emitted by Android app-config handlers). Maps to old
# Documentation.psm1 Add-CustomTable.
function Add-CustomTable {
    param(
        [string]$TableId,
        [string[]]$Columns = @('Name','Value'),
        $Values,
        [int]$Order = 100,
        [string]$LanguageId = ''
    )
    $ctx = Get-CurrentDocumentationContext
    $ctx.AddCustomTable([PSCustomObject]@{
        Id         = $TableId
        Columns    = $Columns
        Values     = $Values
        LanguageId = $LanguageId
        Order      = $Order
    })
}

function Get-CustomChildObject {
    param($Obj, $Prop)
    return Invoke-DocumentationObjectInfoGetChildObject $Obj $Prop
}

function Get-CustomPropertyObject {
    param($Obj, $Prop)
    return Invoke-DocumentationObjectInfoGetPropertyObject $Obj $Prop
}

function Get-CustomProfileValue {
    param($Obj, $Prop)
    return Invoke-DocumentationObjectInfoGetProfileValue $Obj $Prop
}

function Invoke-CustomPostAddValue {
    param($Prop)
    Invoke-DocumentationObjectInfoPostAddValue $Prop
}

function Invoke-ChildSections {
    param($Obj, $SectionObject)
    # Recurse into both .Children and .ChildSettings via the walker
    # (Invoke-TranslateSection in ObjectInfoWalker.ps1). Old code at
    # Documentation.psm1:2859. The walker preserves propLevel via
    # $script:_currentSectionParent so nesting depth tracks correctly.
    $ctx = Get-CurrentDocumentationContext
    $childObj = Get-CustomChildObject $Obj $SectionObject

    $savedLevel = $ctx.PropLevel
    if (Get-Command Invoke-TranslateSection -ErrorAction SilentlyContinue) {
        if ($SectionObject.Children) {
            Invoke-TranslateSection $childObj $SectionObject.Children $null -Parent $SectionObject
        }
        $ctx.PropLevel = $savedLevel
        if ($SectionObject.ChildSettings) {
            Invoke-TranslateSection $childObj $SectionObject.ChildSettings $null -Parent $SectionObject
        }
    }
}

# ---- Translate primitives ----

function Invoke-TranslateBoolean {
    param($Obj, $Prop)

    $propValue = $Obj."$($Prop.entityKey)"
    if ($null -eq $propValue) { $propValue = $Prop.unconfiguredValue }

    if ($propValue -is [string] -and ($propValue -eq 'true' -or $propValue -eq 'false')) {
        $propValue = [bool]::Parse($propValue)
    }

    if ("$propValue" -eq 'notConfigured') {
        return (Get-LanguageString 'BooleanActions.notConfigured')
    }

    # Booleans where $true maps to a single language id
    $singleTrue = @{
        0 = 'BooleanActions.allow';      1 = 'BooleanActions.require'
        2 = 'BooleanActions.enable';     3 = 'BooleanActions.block'
        4 = 'BooleanActions.configured'; 5 = 'BooleanActions.disable'
        6 = 'BooleanActions.limit';      7 = 'BooleanActions.show'
        8 = 'BooleanActions.hide';       9 = 'BooleanActions.yes'
    }
    if ($singleTrue.ContainsKey([int]$Prop.booleanActions) -and $propValue -eq $true) {
        return (Get-LanguageString $singleTrue[[int]$Prop.booleanActions])
    }

    # Booleans where both $true and $false map to language ids (true-id / false-id)
    $pairs = @{
        100 = @('BooleanActions.block',     'BooleanActions.allow')
        101 = @('BooleanActions.require',   'SettingDetails.notRequired')
        102 = @('BooleanActions.enable',    'BooleanActions.disable')
        107 = @('BooleanActions.show',      'BooleanActions.hide')
        108 = @('BooleanActions.hide',      'BooleanActions.show')
        109 = @('BooleanActions.yes',       'SettingDetails.no')
        110 = @('SettingDetails.no',        'BooleanActions.yes')
        120 = @('SettingDetails.onOption',  'SettingDetails.offOption')
        200 = @('BooleanActions.allow',     'BooleanActions.block')
        201 = @('SettingDetails.notRequired','BooleanActions.require')
        220 = @('SettingDetails.offOption', 'SettingDetails.onOption')
    }
    if ($pairs.ContainsKey([int]$Prop.booleanActions)) {
        $ids = $pairs[[int]$Prop.booleanActions]
        $id  = if ($propValue) { $ids[0] } else { $ids[1] }
        return (Get-LanguageString $id)
    }

    Add-NotConfiguredProperty $Prop
    return (Get-LanguageString 'BooleanActions.notConfigured')
}

function Invoke-TranslateOption {
    param($Obj, $Prop, [switch]$SkipOptionChildren, $PropValue = $null)

    if ($null -eq $PropValue) {
        $PropValue = $Obj."$($Prop.entityKey)"
    }

    # Quirk preserved from old code: defenderSecurityCenterDisableRansomwareUI as $true
    # gets coerced to "blockOption" so the option lookup below can match it.
    if ($Obj.defenderSecurityCenterDisableRansomwareUI -eq $true) {
        $Obj.defenderSecurityCenterDisableRansomwareUI = 'blockOption'
    }

    foreach ($option in $Prop.options) {
        if ("$PropValue" -ne "$($option.Value)") { continue }

        $optionValue = $null
        if ($option.nameResource) {
            # $option, not $Prop. The old engine read $prop.nameResource here
            # (Documentation.psm1:3214), so a literal option label rendered the
            # PROPERTY's name as its value - "Token type : Token type". The two
            # sibling translators (MultiOption, MultiOptionBoolean) always read
            # $option.nameResource, which is what makes the typo visible as a bug
            # rather than a convention. No shipped manifest set nameResource on a
            # single-option property, so nothing rendered before this changes.
            $optionValue = $option.nameResource
        }
        elseif ($option.displayText) {
            $optionValue = $option.displayText
        }
        elseif ($option.nameResourceKey) {
            if ($option.nameResourceKey -eq 'notConfigured') {
                Add-NotConfiguredProperty $Prop
            }
            $key = if ($option.nameResourceKey.Contains('.')) { $option.nameResourceKey } else { "SettingDetails.$($option.nameResourceKey)" }
            $optionValue = Get-LanguageString $key
        }
        else {
            $optionValue = $option.Value
        }

        # Return shape preserved from old code (some callers consume the pair)
        @{ Option = $option; Value = $optionValue }

        Add-PropertyInfo $Prop $optionValue $PropValue

        if (-not $SkipOptionChildren) {
            Invoke-ChildSections (Get-CustomChildObject $Obj $Prop) $option
        }
        break
    }

    if ($PropValue -is [bool] -and $Prop.ChildSettings.Count -gt 0) {
        Write-Log "Child properties for boolean $($Prop.EntityKey) value=$PropValue added. Disabled items might be included." 2
    }

    Invoke-ChildSections $Obj $Prop
}

function Invoke-TranslateMultiOption {
    param($Obj, $Prop)

    $propValues = $null
    if ($Obj.PSObject.Properties.Name -contains $Prop.entityKey) {
        $propValues = $Obj."$($Prop.entityKey)"
        if ($propValues -is [string]) { $propValues = $propValues.Split(',') }
    }
    elseif ($Prop.entityKey -like '*List') {
        $tmpProp = $Prop.entityKey.Substring(0, $Prop.entityKey.Length - 4)
        $propValues = $Obj.$tmpProp
    }

    $ctx = Get-CurrentDocumentationContext
    $selectedValues = @()
    foreach ($propValue in $propValues) {
        $option = $Prop.Options | Where-Object Value -EQ $propValue
        if (-not $option) { continue }

        if ($option.nameResource) {
            $selectedValues += $option.nameResource
        }
        else {
            $key = if ($option.nameResourceKey.Contains('.')) { $option.nameResourceKey } else { "SettingDetails.$($option.nameResourceKey)" }
            $selectedValues += (Get-LanguageString $key)
        }
    }

    if ($selectedValues.Count -gt 0) {
        return ($selectedValues -join $ctx.PropertySeparator)
    }

    Add-NotConfiguredProperty $Prop
    return (Get-LanguageString 'BooleanActions.notConfigured')
}

function Invoke-TranslateMultiOptionBoolean {
    param($Obj, $Prop, $SelectedValue = $true)

    $propObj = $Obj."$($Prop.entityKey)"
    if (-not $propObj) { return (Get-LanguageString 'BooleanActions.notConfigured') }

    $ctx = Get-CurrentDocumentationContext
    $selectedValues = @()
    foreach ($propValue in $propObj.PSObject.Properties.Name) {
        if ($propObj.$propValue -isnot [bool])  { continue }
        $option = $Prop.options | Where-Object value -EQ $propValue
        if (-not $option) { continue }

        if ($propObj.$propValue -ne $SelectedValue) { continue }

        if ($option.nameResource) {
            $selectedValues += $option.nameResource
        }
        else {
            $key = if ($option.nameResourceKey.Contains('.')) { $option.nameResourceKey } else { "SettingDetails.$($option.nameResourceKey)" }
            $selectedValues += (Get-LanguageString $key)
        }
    }

    if ($selectedValues.Count -gt 0) {
        return ($selectedValues -join $ctx.PropertySeparator)
    }
    return (Get-LanguageString 'BooleanActions.notConfigured')
}

function Invoke-TranslateTable {
    param($Obj, $Prop)

    $propValue = if ($Prop.entityKey -eq '.') { $Obj } else { $Obj."$($Prop.entityKey)" }
    $ctx = Get-CurrentDocumentationContext

    $items         = @()
    $itemFullValue = @()
    foreach ($item in $propValue) {
        $itemValues   = @()
        $htFullPropInfo = [ordered]@{}

        foreach ($column in $Prop.Columns) {
            if ($column.metadata.entityKey -eq 'unusedForSingleItems') {
                $itemValues += $item
            }
            elseif ($column.metadata.entityKey -eq $Prop.entityKey -and ($Prop.Columns | Measure-Object).Count -eq 1) {
                # Self-referencing single-column tables
                $itemValues += $item
            }
            elseif (($Prop.Columns | Measure-Object).Count -eq 1 -and `
                    $null -eq $item."$($column.metadata.entityKey)" -and `
                    $null -eq $Obj."$($column.metadata.entityKey)" -and `
                    $item -is [string]) {
                # String-list with declared (but empty) entity key
                $itemValues += $item
            }
            else {
                $itemTmpVal = $null
                if ($item.PSObject.Properties | Where-Object Name -Like $column.metadata.entityKey) {
                    $itemTmpVal = $item."$($column.metadata.entityKey)"
                }
                else {
                    $itemTmpVal = $Obj."$($column.metadata.entityKey)"
                }
                $itemValues += $itemTmpVal

                if ($Prop.Columns.Count -gt 1) {
                    if ($column.metadata.nameResourceKey) {
                        $key = if ($column.metadata.nameResourceKey.Contains('.')) { $column.metadata.nameResourceKey } else { "SettingDetails.$($column.metadata.nameResourceKey)" }
                        $colName = Get-LanguageString $key
                        if (-not $colName) { $colName = $column.metadata.entityKey }
                        $htFullPropInfo.Add($colName, ($itemTmpVal -join $ctx.PropertySeparator))
                    }
                    else {
                        $nameForLog = if ($Prop.nameResourceKey) { $Prop.nameResourceKey } else { $Prop.entityKey }
                        Write-Log "Property $nameForLog does not have nameResourceKey on one of the columns" 2
                    }
                }
            }
        }

        if ($htFullPropInfo.Count -gt 0) {
            $itemFullValue += [PSCustomObject]$htFullPropInfo
        }

        $sep = if ($Prop.separator) { $Prop.separator } else { $ctx.PropertySeparator }
        $items += ($itemValues -join $sep)
    }

    if ($items.Count -gt 0) {
        $params = @{}
        if ($itemFullValue.Count -gt 0) { $params['TableValue'] = $itemFullValue }

        if ((-not $Prop.nameResourceKey -or $Prop.nameResourceKey -eq 'Empty') -and $Prop.columns[0].metadata.nameResourceKey) {
            Add-PropertyInfo $Prop.columns[0].metadata ($items -join $ctx.ObjectSeparator) $propValue @params
        }
        else {
            Add-PropertyInfo $Prop ($items -join $ctx.ObjectSeparator) $propValue @params
        }
    }
    else {
        if ((-not $Prop.nameResourceKey -or $Prop.nameResourceKey -eq 'Empty') -and $Prop.Columns[0].metadata.nameResourceKey) {
            Add-PropertyInfo $Prop.Columns[0].metadata $null
        }
        else {
            Add-PropertyInfo $Prop $null
        }
    }

    Invoke-ChildSections $Obj $Prop
}

function Invoke-TranslateDuration {
    param($Obj, $Prop)
    $raw = $Obj."$($Prop.entityKey)"
    # Optional manifest hint ("durationUnit": "minutes"|"days"|"hours"|"seconds"):
    # render the ISO8601 duration in the unit the row label promises. Graph stores
    # normalized durations (30240 minutes round-trips as P21D), so the legacy
    # first-component fallback below would show "21" under a "(minutes)" label.
    if ($Prop.durationUnit -and $raw) {
        $ts = Get-DurationValue $raw -ReturnTimeSpan
        if ($ts -is [timespan]) {
            $value = switch ([string]$Prop.durationUnit) {
                'days'    { $ts.TotalDays }
                'hours'   { $ts.TotalHours }
                'seconds' { $ts.TotalSeconds }
                default   { $ts.TotalMinutes }
            }
            return [string][math]::Round($value)
        }
    }
    Get-DurationValue $raw
}

function Get-DurationValue {
    param($DurationValue, [switch]$ReturnTimeSpan)

    if (-not $DurationValue -or -not $DurationValue.StartsWith('P')) { return "0" }

    # Ported from old Documentation.psm1:3461. Without -ReturnTimeSpan the function
    # returns the bare digits of the FIRST non-empty component (e.g. "P70D" -> "70",
    # "P1Y" -> "1", "PT15M" -> "15"). With -ReturnTimeSpan it accumulates
    # years+days+hours+minutes+seconds into a real TimeSpan.
    #
    # NB: the old code's $DurationValue.Split($arr) is a no-op - a [string[]] passed
    # to String.Split() without StringSplitOptions binds to no char[] overload and
    # returns the whole string, so the function always returned the raw ISO8601
    # value ("P70D" rendered verbatim under a "(days)" label). Split explicitly on
    # [string[]] with StringSplitOptions::None so empty inter-delimiter segments are
    # kept and the $values indices stay aligned with the delimiter loop below.
    $arr = @('P','T','Y','D','H','M','S')
    $values = $DurationValue.Split([string[]]$arr, [System.StringSplitOptions]::None)

    $years = 0; $days = 0; $hours = 0; $minutes = 0; $seconds = 0
    $i = 0
    foreach ($tmp in $arr) {
        if ($DurationValue.Contains($tmp)) {
            if ($ReturnTimeSpan) {
                switch ($tmp) {
                    'Y' { $years   = [int]$values[$i] }
                    'D' { $days    = [int]$values[$i] }
                    'H' { $hours   = [int]$values[$i] }
                    'M' { $minutes = [int]$values[$i] }
                    'S' { $seconds = [int]$values[$i] }
                }
            }
            elseif (-not [string]::IsNullOrEmpty($values[$i])) {
                return $values[$i]
            }
            $i++
        }
    }

    if ($ReturnTimeSpan) {
        $days += ($years * 365)  # approximate; matches old code
        return [timespan]::new($days, $hours, $minutes, $seconds)
    }
    return "0"
}
