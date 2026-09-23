# Internal helpers for Set-GraphBulkAssignments (Public/Set-GraphBulkAssignments.ps1).
# Extracted here per architecture rule R11 (keep public driver files thin). These
# are module-internal (not exported) but several are also called by the bulk-
# assignment UI (Test-BulkAssignmentSupported, Get-BulkAssignmentObjectType, ...)
# and the online tests - all run in module scope, so name resolution is unchanged.

# App @odata.types Graph refuses an assignment filter for. Set-GraphBulkAssignments
# assigns these without the filter rather than failing the policy. Extend as
# further types are verified live; do not guess.
$script:BulkAssignmentAppTypesWithoutFilters = @(
    '#microsoft.graph.webApp'
)

# Stable string key for an assignment — used for set-membership (dedupe,
# no-op detection, removal matching). Encodes the target discriminator +
# the fields we expose in the UI; Intent is appended for app-shape
# assignments so "Required to Group X" and "Available to Group X" are
# distinct entries.
function Get-AssignmentSignature
{
    param($Target, [string]$Intent)

    if(-not $Target) { return "" }

    $type = [string]$Target.'@odata.type'
    # Strip the leading '#microsoft.graph.' if present so signatures compare
    # equal regardless of whether Graph returned the prefix.
    $type = $type -replace '^#?microsoft\.graph\.', ''

    $groupId    = [string]$Target.groupId
    $filterId   = [string]$Target.deviceAndAppManagementAssignmentFilterId
    $filterType = [string]$Target.deviceAndAppManagementAssignmentFilterType
    # "none" and "" are equivalent — Intune sometimes writes one, sometimes
    # the other. Normalise so signatures don't drift between Add runs. The all-zeros
    # sentinel is the third spelling of the same thing (Test-AssignmentFilterDefined):
    # without it, one assignment carrying the sentinel and an identical one carrying no
    # filter property hash differently, so an already-assigned target looks new.
    if($filterType -eq "none" -or -not (Test-AssignmentFilterDefined $filterId)) {
        $filterType = ""
        $filterId   = ""
    }

    $intentPart = if([string]::IsNullOrEmpty($Intent)) { "" } else { [string]$Intent }
    return "$type|$groupId|$filterId|$filterType|$intentPart".ToLowerInvariant()
}

function Get-AssignmentFullSignature
{
    param($Tuple)

    if(-not $Tuple) { return "" }

    $runRemediationPart = ""
    if($null -ne $Tuple.RunRemediation) { $runRemediationPart = [string][bool]$Tuple.RunRemediation }

    $parts = @(
        (Get-AssignmentSignature $Tuple.Target $Tuple.Intent),
        (ConvertTo-StableAssignmentJson $Tuple.Settings),
        (ConvertTo-StableAssignmentJson $Tuple.RunSchedule),
        $runRemediationPart
    )
    return ($parts -join "|").ToLowerInvariant()
}

function ConvertTo-StableAssignmentJson
{
    param($Value)

    if($null -eq $Value) { return "" }
    return (ConvertTo-StableAssignmentObject $Value | ConvertTo-Json -Depth 50 -Compress)
}

function ConvertTo-StableAssignmentObject
{
    param($Value)

    if($null -eq $Value) { return $null }
    if($Value -is [string]) { return $Value }
    if($Value -is [System.Collections.IDictionary]) {
        $ordered = [ordered]@{}
        foreach($key in @($Value.Keys | Sort-Object)) {
            $ordered[[string]$key] = ConvertTo-StableAssignmentObject $Value[$key]
        }
        return [PSCustomObject]$ordered
    }
    if($Value -is [System.Collections.IEnumerable] -and $Value -isnot [PSCustomObject]) {
        $items = [System.Collections.Generic.List[object]]::new()
        foreach($item in $Value) {
            [void]$items.Add((ConvertTo-StableAssignmentObject $item))
        }
        return $items.ToArray()
    }
    if($Value -is [PSCustomObject]) {
        $ordered = [ordered]@{}
        foreach($prop in @($Value.PSObject.Properties | Sort-Object Name)) {
            $ordered[$prop.Name] = ConvertTo-StableAssignmentObject $prop.Value
        }
        return [PSCustomObject]$ordered
    }
    return $Value
}

# Project a raw assignment target (from Graph response) down to just the
# fields the bulk tool supports. Keeps the result POST-safe.
function ConvertTo-AssignmentTarget
{
    param($Target)

    $out = [ordered]@{ "@odata.type" = [string]$Target.'@odata.type' }
    if($Target.groupId) { $out.groupId = [string]$Target.groupId }
    if($Target.deviceAndAppManagementAssignmentFilterId) {
        $out.deviceAndAppManagementAssignmentFilterId   = [string]$Target.deviceAndAppManagementAssignmentFilterId
        $out.deviceAndAppManagementAssignmentFilterType = [string]$Target.deviceAndAppManagementAssignmentFilterType
    }
    return [PSCustomObject]$out
}

# Build a fresh Graph target object from a UI-supplied assignment descriptor
# (the PSCustomObject the UI puts into AssignmentSettings.Assignments).
function Build-AssignmentTarget
{
    param($Descriptor)

    $obj = [ordered]@{
        "@odata.type" = "#microsoft.graph.$($Descriptor.TargetType)"
    }
    if($Descriptor.GroupId) { $obj.groupId = [string]$Descriptor.GroupId }
    if($Descriptor.FilterId) {
        $obj.deviceAndAppManagementAssignmentFilterId   = [string]$Descriptor.FilterId
        $obj.deviceAndAppManagementAssignmentFilterType = if($Descriptor.FilterType) { [string]$Descriptor.FilterType } else { "include" }
    }
    return [PSCustomObject]$obj
}

# Bulk tool supports three assignment shapes:
#   "simple" — `{target}` only. Most types.
#   "app"    — `{target, intent, settings?}`. mobileAppAssignments.
#   "script" — `{target, runSchedule?, runRemediationScript?}`. Health
#              scripts (deviceHealthScriptAssignments).
function Get-BulkAssignmentShape
{
    param($PolicyType)

    if(-not $PolicyType) { return $null }
    switch ([string]$PolicyType.AssignmentsType) {
        "mobileAppAssignments"          { return "app" }
        "deviceHealthScriptAssignments" { return "script" }
    }
    if(-not $PolicyType.AssignmentPropertiesToKeep) { return "simple" }
    return $null
}

function Test-BulkAssignmentSupported
{
    param($PolicyType)

    if(-not $PolicyType) { return $false }
    if(-not $PolicyType.SupportsAssignments) { return $false }
    if(-not $PolicyType.AssignmentsType)     { return $false }

    # These types either do not expose a policy assignment action or use a
    # custom assignment flow that is not the replace-all /assign contract.
    # AppProtection uses the polymorphic /managedAppPolicies endpoint for
    # list/read, but /assign is only bound to the concrete platform subtypes
    # (iosManagedAppProtections, androidManagedAppProtections, windowsManagedAppProtections,
    # mdmWindowsInformationProtectionPolicies). Calling /managedAppPolicies/{id}/assign
    # is unsupported.
    # AppleEnrollmentTypes has an `assignments` navigation property but no
    # `/assign` action — assignments are managed by POST/DELETE on the nested
    # /assignments collection, not the replace-all contract.
    if($PolicyType.Id -in @(
        "AppleEnrollmentTypes",
        "AppProtection",
        "Autopilot",
        "IntuneBranding",
        "Notifications",
        "ScopeTags",
        "TermsAndConditions"
    )) { return $false }

    return ($null -ne (Get-BulkAssignmentShape $PolicyType))
}

function Get-BulkAssignmentObjectType
{
    param($PolicyType)

    if(-not $PolicyType) { return $null }

    # Per-type override on the PolicyType class takes precedence over the
    # built-in heuristic below. Lets new PolicyTypes declare their assignment
    # @odata.type next to the type definition instead of editing this central
    # switch — see _AssignmentObjectType on IntunePolicyTypeBase.
    $override = [string]$PolicyType.AssignmentObjectType
    if($override) { return $override }

    switch ([string]$PolicyType.AssignmentsType) {
        "deviceHealthScriptAssignments"     { return "#microsoft.graph.deviceHealthScriptAssignment" }
        "deviceManagementScriptAssignments" { return "#microsoft.graph.deviceManagementScriptAssignment" }
        "enrollmentConfigurationAssignments"{ return "#microsoft.graph.enrollmentConfigurationAssignment" }
        "mobileAppAssignments"              { return "#microsoft.graph.mobileAppAssignment" }
        "hardwareConfigurationAssignments"  { return "#microsoft.graph.hardwareConfigurationAssignment" }
    }

    switch ([string]$PolicyType.Id) {
        "EndpointSecurity"                 { return "#microsoft.graph.deviceManagementIntentAssignment" }
        "EnrollmentSettingsCatalog"        { return "#microsoft.graph.deviceManagementConfigurationPolicyAssignment" }
        "DeviceConfiguration"              { return "#microsoft.graph.deviceConfigurationAssignment" }
        "CompliancePolicies"               { return "#microsoft.graph.deviceCompliancePolicyAssignment" }
        "CompliancePoliciesV2"             { return "#microsoft.graph.deviceManagementConfigurationPolicyAssignment" }
        "SettingsCatalog"                  { return "#microsoft.graph.deviceManagementConfigurationPolicyAssignment" }
        "EndpointSecuritySettingsCatalog"  { return "#microsoft.graph.deviceManagementConfigurationPolicyAssignment" }
        "DeviceConfigurationScripts"       { return "#microsoft.graph.deviceManagementConfigurationPolicyAssignment" }
        "InventoryPolicies"                { return "#microsoft.graph.deviceManagementConfigurationPolicyAssignment" }
        "AppConfigurationManagedDevice"    { return "#microsoft.graph.managedDeviceMobileAppConfigurationAssignment" }
        "AppConfigurationManagedApp"       { return "#microsoft.graph.targetedManagedAppPolicyAssignment" }
        "AndroidOEMConfig"                 { return "#microsoft.graph.managedDeviceMobileAppConfigurationAssignment" }
        "Applications"                     { return "#microsoft.graph.mobileAppAssignment" }
        "IosLobAppProvisioningConfigurations" { return "#microsoft.graph.iosLobAppProvisioningConfigurationAssignment" }
        "PolicySets"                       { return "#microsoft.graph.policySetAssignment" }
        "UpdatePolicies"                   { return "#microsoft.graph.deviceConfigurationAssignment" }
        "iOSiPadOSUpdatePolicies"          { return "#microsoft.graph.deviceConfigurationAssignment" }
        "macOSUpdatePolicies"              { return "#microsoft.graph.deviceConfigurationAssignment" }
        "FeatureUpdates"                   { return "#microsoft.graph.windowsFeatureUpdateProfileAssignment" }
        "QualityUpdates"                   { return "#microsoft.graph.windowsQualityUpdateProfileAssignment" }
        "QualityUpdatePolicies"            { return "#microsoft.graph.windowsQualityUpdatePolicyAssignment" }
        "DriverUpdateProfiles"             { return "#microsoft.graph.windowsDriverUpdateProfileAssignment" }
        "AdministrativeTemplates"          { return "#microsoft.graph.groupPolicyConfigurationAssignment" }
        "W365ProvisioningPolicies"         { return "#microsoft.graph.cloudPcProvisioningPolicyAssignment" }
        "W365UserSettings"                 { return "#microsoft.graph.cloudPcUserSettingAssignment" }
    }

    return $null
}

# Map a policy's @odata.type (e.g. "#microsoft.graph.win32LobApp") to the
# matching mobileAppAssignmentSettings derived type name (e.g.
# "win32LobAppAssignmentSettings"). The Graph naming convention is
# `<appType>AssignmentSettings`, so a strip-and-append is enough. Returns
# $null when the input doesn't look like a known app type.
function Get-AppSettingsTypeForPolicy
{
    param([string]$PolicyOdataType)

    if([string]::IsNullOrEmpty($PolicyOdataType)) { return $null }
    $name = $PolicyOdataType -replace '^#?microsoft\.graph\.', ''
    if([string]::IsNullOrEmpty($name)) { return $null }
    return "${name}AssignmentSettings"
}

# Build a deviceHealthScriptRunSchedule PSCustomObject from a UI spec
# hashtable. Picks the right @odata.type based on scheduleType
# (Hourly/Daily/Once) and emits only the fields each schedule supports.
# Returns $null when the spec doesn't have enough information.
function Build-HealthScriptSchedule
{
    param([Hashtable]$Spec)

    if(-not $Spec -or -not $Spec.ContainsKey('scheduleType')) { return $null }

    $type = [string]$Spec['scheduleType']
    $interval = 0
    if($Spec.ContainsKey('interval')) {
        try { $interval = [int]$Spec['interval'] } catch { $interval = 0 }
    }

    switch ($type) {
        "Hourly" {
            return [PSCustomObject]@{
                "@odata.type" = "#microsoft.graph.deviceHealthScriptHourlySchedule"
                interval      = $interval
            }
        }
        "Daily" {
            $obj = [ordered]@{
                "@odata.type" = "#microsoft.graph.deviceHealthScriptDailySchedule"
                interval      = $interval
                useUtc        = [bool]$Spec['useUtc']
            }
            if($Spec.ContainsKey('time') -and $Spec['time']) {
                $obj.time = [string]$Spec['time']
            }
            return [PSCustomObject]$obj
        }
        "Once" {
            $obj = [ordered]@{
                "@odata.type" = "#microsoft.graph.deviceHealthScriptRunOnceSchedule"
                interval      = $interval
                useUtc        = [bool]$Spec['useUtc']
            }
            if($Spec.ContainsKey('date') -and $Spec['date']) { $obj.date = [string]$Spec['date'] }
            if($Spec.ContainsKey('time') -and $Spec['time']) { $obj.time = [string]$Spec['time'] }
            return [PSCustomObject]$obj
        }
    }
    return $null
}

# Convert a user-built settings hashtable (from the App settings... dialog)
# into a Graph-shaped PSCustomObject with the requested @odata.type stamped
# at the top. Nested complex types (win32LobAppRestartSettings etc.) are
# expected to already carry their own @odata.type in the hashtable; this
# function only wraps the outer object.
function ConvertTo-AppSettingsObject
{
    param([Hashtable]$Hash, [string]$GraphType)

    if(-not $Hash) { return $null }
    $obj = [ordered]@{ "@odata.type" = "#microsoft.graph.$GraphType" }
    foreach($k in $Hash.Keys) {
        $obj[$k] = ConvertTo-AppSettingsValue $Hash[$k]
    }
    return [PSCustomObject]$obj
}

# Recursive companion to ConvertTo-AppSettingsObject. Walks arbitrary
# nesting depth, converting any [Hashtable] (including nested ones inside
# arrays) into [PSCustomObject] so ConvertTo-Json emits JSON objects rather
# than the @{key=value} debug format. Arrays are preserved and their
# elements recursed individually. Strings (which are IEnumerable) and
# PSCustomObject (already correct shape) pass through unchanged.
function ConvertTo-AppSettingsValue
{
    param($Value)

    if($null -eq $Value) { return $null }
    if($Value -is [Hashtable]) {
        $nested = [ordered]@{}
        foreach($k in $Value.Keys) {
            $nested[$k] = ConvertTo-AppSettingsValue $Value[$k]
        }
        return [PSCustomObject]$nested
    }
    if($Value -is [string]) { return $Value }
    if($Value -is [System.Collections.IList]) {
        $items = [System.Collections.Generic.List[object]]::new()
        foreach($item in $Value) {
            [void]$items.Add((ConvertTo-AppSettingsValue $item))
        }
        return $items.ToArray()
    }
    return $Value
}
