#ImportOrder 220

#########################################################################################
#
# Script Group
#
#########################################################################################

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class DeviceEnrollmentGroup : IntunePolicyGroupBase
{
    DeviceEnrollmentGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "DeviceEnrollments"
        $this._Name = "Device enrollment"
        $this._Icon = "WindowsEnrollments"
    }
}

#########################################################################################
#
# Autopilot
#
#########################################################################################

# region Autopilot
class AutopilotType : IntunePolicyTypeBase
{
    AutopilotType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceEnrollmentGroup")
        $this._PolicyName = "Autopilot"
        $this._ID = "Autopilot"
        # Platform default: the endpoint serves exactly one platform and the
        # objects carry no platforms/platformType field, so the column would
        # otherwise be blank (windowsAutopilotDeploymentProfiles is
        # Windows-only).
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        $this._API = "deviceManagement/windowsAutopilotDeploymentProfiles"
        $this._CopyDefaultName = "%Name% Copy"
        $this._Permissions = @("DeviceManagementServiceConfig.ReadWrite.All")
        $this._ObjectClass = "AutopilotObject"
        $this._PropertiesToRemoveForUpdate = @('managementServiceAppId')        

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreDeleteCommand([IntunePolicyBase]$PolicyObject)
    {
        Write-Log "Delete AutoPilot profile assignments"

        foreach($assignment in $PolicyObject.Assignments)
        {
            if($assignment.Source -ne "direct") { continue }

            $api = "$($PolicyObject.PolicyType.API)/$($PolicyObject.Id)/assignments/$($assignment.Id)"

            $repsone = Invoke-MSGraphAPI -Url $api -HttpMethod "DELETE" -TokenId $PolicyObject._TokenId -FullResponseObject
            if($repsone.Success)
            {
                Write-LogDebug "Assignemnt with Id $($assignment.Id) deleted successfully"
            }
        }
        return $null
    }

    [Hashtable]PreImportAssignmentsCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        return (Add-GraphAssignmentsToObject $PolicyObject $SourceObject)
    }

}

Class AutopilotObject : IntunePolicyBase
{
    AutopilotObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AutopilotObject() : Base()
    {
        $this.Init()

        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
    }

    Hidden Init()
    {

        $this._PolicyType = (Get-SingletonObject "AutopilotType")
    }
}

#########################################################################################
#
# Device Enrollment
#
#########################################################################################

# region Device Enrollment

# DeviceEnrollmentType is the SHARED BASE for everything served by
# /deviceManagement/deviceEnrollmentConfigurations. The Graph endpoint returns
# eight subtypes of an abstract base type (per the schema):
#   deviceEnrollmentLimitConfiguration                  ('limit' / 'defaultLimit')
#   deviceEnrollmentPlatformRestrictionConfiguration    ('singlePlatformRestriction')
#   deviceEnrollmentPlatformRestrictionsConfiguration   ('platformRestrictions' / 'defaultPlatformRestrictions')
#   deviceEnrollmentWindowsHelloForBusinessConfiguration('windowsHelloForBusiness' / 'defaultWindowsHelloForBusiness')
#   windows10EnrollmentCompletionPageConfiguration      ('windows10EnrollmentCompletionPageConfiguration' / 'defaultWindows10EnrollmentCompletionPageConfiguration')
#   deviceComanagementAuthorityConfiguration            ('deviceComanagementAuthorityConfiguration')
#   deviceEnrollmentNotificationConfiguration           ('enrollmentNotificationsConfiguration')
#   windowsRestoreDeviceEnrollmentConfiguration         ('windowsRestore')
#
# Each logical bucket gets its own thin subtype below. The base only carries
# import/export/replace behavior — it is NOT registered as a policy type itself
# (no _PolicyGroup), so $script:IntuneTypes only contains the concrete buckets.
#
# Server-side filtering by `deviceEnrollmentConfigurationType eq '...'` is value-exact
# and excludes the 'default*' variants, so subtypes filter client-side via CheckPolicy
# on @odata.type. With identical _API/_QueryList across siblings, Get-GraphPolicies
# coalesces them into ONE batch sub-request and fans rows out to the matching subtype.
class DeviceEnrollmentType : IntunePolicyTypeBase
{
    # Abstract: only concrete subtypes (EnrollmentStatusPageType, EnrollmentLimitType, ...)
    # may be instantiated. Auto-discovery loops in Invoke-IntuneEventAppInitialized and
    # the UI extensions consult Test-ClassIsAbstract on each candidate and skip those
    # declaring this static marker, so $script:IntuneTypes only contains concrete buckets.
    static [bool] $IsAbstract = $true

    DeviceEnrollmentType() : Base()
    {
        ([DeviceEnrollmentType]$this).Init()
    }

    Init()
    {
        # Everything in this Init is shared across every concrete bucket. Subtypes only
        # override _ID, _APITitle, _PolicyName, _Folder, _QueryList, _PlatformName (when
        # platform-specific), and CheckPolicy. Subtypes do NOT need to call this Init
        # explicitly — the constructor chain already runs it via : Base() → DeviceEnrollmentType()
        # body's ([DeviceEnrollmentType]$this).Init().
        $this._API                         = "deviceManagement/deviceEnrollmentConfigurations"
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._Permissions                 = @("DeviceManagementServiceConfig.ReadWrite.All")
        $this._SkipRemoveProperties        = @('Id')
        $this._PropertiesToRemoveForUpdate = @('priority')
        $this._Dependencies                = @('Applications')
        $this._AssignmentsType             = "enrollmentConfigurationAssignments"
        $this._Icon                        = "EnrollmentStatusPage"
        $this._PolicyGroup                 = (Get-SingletonObject "DeviceEnrollmentGroup")
        $this._ObjectClass                 = "DeviceEnrollmentObject"
        $this._VerifyObject                = $true
        # _ExpandAssignmentsList stays at its default ($true) so the list URL appends
        # &$expand=assignments. The Intune portal does this on every enrollment-config
        # subtype, and Graph accepts it here, so we get assignments inline and skip the
        # follow-up /assignments round-trip in Add-GraphPolicyAssignments.
        # No _QueryList here — each subtype owns its filter. Combining filters across
        # subtypes via OR was unreliable (Graph's batch endpoint dropped most matches).
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.Object.Priority -eq 0)
        {
            $ret = @{}
            $ret.Add("API","$($PolicyObject.PolicyType.API)/$($PolicyObject.Id)")
            $ret.Add("Method","PATCH") # Default profile always exists so update them
            $ret
        }
        else
        {
            Remove-Property $PolicyObject.Object "Id"
        }
        return $null
    }

    [Hashtable]PreImportAssignmentsCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        if($SourceObject.Object.Priority -eq 0) { return  @{ "Import" = $false } }
        return $null
    }

    [Hashtable]PreDeleteCommand([IntunePolicyBase]$PolicyObject)
    {
        if($PolicyObject.Object.Priority -eq 0) { return  @{ "Delete" = $false } }
        return $null
    }

    [Hashtable]PreReplaceCommand([IntunePolicyBase]$PolicyObject)
    {
        if($PolicyObject.Object.Priority -eq 0) { return  @{ "Delete" = $false } }
        return $null
    }

    PostReplaceCommand([IntunePolicyBase]$PolicyObject, [IntunePolicyBase]$SourceObject)
    {
        Set-EnrollmentRestrictionsPriority $PolicyObject $SourceObject
    }
}

# Shared Object class for every DeviceEnrollment subtype. PolicyName, Folder and
# Platform routing all flow from the subtype's TYPE class (_PolicyName / _Folder /
# _PlatformName), not per-instance switches on @odata.type:
#   * IntunePolicyTypeBase.GetObject sets policyObject._PolicyType = $this (the calling
#     subtype) after construction, so $obj.PolicyType.Folder / .PolicyName resolve to
#     the right subtype's values.
#   * IntunePolicyBase.Platform getter falls back to _PolicyType._PlatformName when the
#     instance doesn't set its own — that's how Windows-only buckets (ESP, WHfB,
#     CoMgmt, WindowsRestore) report Platform=Windows.
# The constructor is intentionally empty: IntunePolicyBase's : Base($JsonObj) chain
# already runs the Add-ObjectProperty setup. Anything we'd add to a subclass Init here
# is per-subtype concern and lives on the subtype TYPE.
Class DeviceEnrollmentObject : IntunePolicyBase
{
    DeviceEnrollmentObject([PSCustomObject]$JsonObj) : Base($JsonObj) { }
    DeviceEnrollmentObject()                         : Base()         { }

    [String]GetFileName([String]$Path)
    {
        # Default policies (priority=0) have a localized boilerplate displayName
        # ("All users and all devices") that collides across subtypes — both
        # DefaultLimit and DefaultPlatformRestrictions land on the same name and
        # one overwrites the other. Their id-suffix is unique per subtype
        # (DefaultLimit / DefaultPlatformRestrictions / DefaultWindowsHelloForBusiness /
        # DefaultWindows10EnrollmentCompletionPageConfiguration), so we use that.
        # Non-default policies have id format <randomGuid>_<configTypeSuffix> where
        # the suffix is identical for every policy of a given subtype, so we use
        # displayName instead.
        # We use TWO signals — priority OR a TenantId-prefixed id — because each
        # has a failure mode on its own:
        #   * priority alone: depends on Graph exposing 'priority' on every payload
        #     and on it staying =0 for default policies (mostly true but not
        #     guaranteed across endpoints / future schema changes).
        #   * TenantId prefix alone: requires $this.TenantId to be populated by the
        #     time GetFileName runs; it isn't always (file-load paths, certain
        #     bulk-export code paths skip the TenantId-set step).
        # Combining them is robust: a row that's a default in EITHER sense uses the
        # id-suffix, everything else uses displayName.

        $isDefault = ($this.Object.priority -eq 0) -or `
                     ($this.TenantId -and $this.Id -and $this.Id.StartsWith($this.TenantId + "_"))

        if($isDefault) {
            $parts = $this.Id -split '_', 2
            $name = if($parts.Count -ge 2) { $parts[1] } else { $null }
        }
        else {
            $name = $this.Object.displayName
        }
        if(-not $name) { $name = $this.Id }

        # Same id-suffix rule as IntunePolicyBase.GetFileName, which this override
        # replaced wholesale and therefore silently dropped: both the user-facing
        # AddIDToExportFile setting AND the bulk-export collision flag were ignored
        # here, so two non-default enrollment policies of the same subtype sharing a
        # displayName still wrote the same file and one overwrote the other - the
        # collision was DETECTED and then not acted on.
        #
        # A default policy takes its name from the id suffix already, which is unique
        # per subtype, so the flag will not normally fire for one; the rule is applied
        # unconditionally anyway rather than only in the else-branch, so an id-suffix
        # name that does somehow collide is still disambiguated.
        $forceId = (Get-SettingValue "AddIDToExportFile") -eq $true -or $this._NeedsIdInFilename -eq $true
        if($forceId -and $this.Id -and $this.PolicyType.SkipAddIDOnFileName -ne $true -and $name -ne $this.Id) {
            $name = ($name + "_" + $this.Id)
        }

        $fileName = "$((Remove-InvalidFileNameChars $name)).json"
        if($Path) { $fileName = [IO.Path]::Combine($Path, $fileName) }
        return $fileName
    }
}

# Concrete bucket conventions:
#   * Constructor delegates to : Base() (= DeviceEnrollmentType) which runs the shared
#     Init via the constructor chain. Subtype Init does NOT call the base Init
#     explicitly — that would run the base Init twice.
#   * Subtype Init only sets bucket-specific fields: _ID, _APITitle, _PolicyName,
#     _Folder, _QueryList, optionally _PlatformName, then registers via AddPolicyType.
#   * CheckPolicy stays as a defensive client-side filter; @odata.type is the primary
#     discriminator with deviceEnrollmentConfigurationType as a fallback.

class EnrollmentStatusPageType : DeviceEnrollmentType
{
    EnrollmentStatusPageType() : Base() { ([EnrollmentStatusPageType]$this).Init() }
    Init()
    {
        $this._ID           = "EnrollmentStatusPage"
        $this._APITitle     = "Enrollment Status Page"
        $this._PolicyName   = "Enrollment Status Page"
        $this._Folder       = "EnrollmentStatusPage"
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        # Single eq is enough: empirically Graph returns both default and non-default
        # ESPs for this filter. (Adding 'defaultWindows10…' as a second clause causes
        # a 400 — that enum value is in the schema but rejected by the live filter parser.)
        $this._QueryList    = "?`$filter=deviceEnrollmentConfigurationType eq 'windows10EnrollmentCompletionPageConfiguration'"
        if($null -ne $this._PolicyGroup) { $this._PolicyGroup.AddPolicyType($this) }
    }
    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if ($PolicyObject.'@odata.type' -eq '#microsoft.graph.windows10EnrollmentCompletionPageConfiguration') { return $true }
        return ($PolicyObject.deviceEnrollmentConfigurationType -eq 'windows10EnrollmentCompletionPageConfiguration')
    }
}

class EnrollmentRestrictionsPageType : DeviceEnrollmentType
{
    EnrollmentRestrictionsPageType() : Base() { ([EnrollmentRestrictionsPageType]$this).Init() }
    Init()
    {
        $this._ID         = "EnrollmentRestrictions"
        $this._APITitle   = "Enrollment Restrictions"
        $this._PolicyName = "Device platform restrictions"
        $this._Folder     = "EnrollmentRestrictions"
        # Cross-platform (covers iOS / Android / Windows etc.) — no _PlatformName.
        # Single eq clause empirically returns BOTH per-platform Block Android-style
        # configs (@odata.type singular) AND the default combined platform restrictions
        # (@odata.type plural, id-suffix _DefaultPlatformRestrictions). 'limit' lives
        # on EnrollmentLimitType because stacking a second eq clause on the same
        # property in batch mode caused Graph's batch endpoint to drop most matches.
        $this._QueryList  = "?`$filter=deviceEnrollmentConfigurationType eq 'singlePlatformRestriction'"
        if($null -ne $this._PolicyGroup) { $this._PolicyGroup.AddPolicyType($this) }
    }
    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if ($PolicyObject.'@odata.type' -in @(
            '#microsoft.graph.deviceEnrollmentPlatformRestrictionConfiguration',
            '#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration')) { return $true }
        return ($PolicyObject.deviceEnrollmentConfigurationType -in @(
            'singlePlatformRestriction',
            'platformRestrictions',
            'defaultPlatformRestrictions'))
    }
}

class EnrollmentLimitType : DeviceEnrollmentType
{
    EnrollmentLimitType() : Base() { ([EnrollmentLimitType]$this).Init() }
    Init()
    {
        $this._ID         = "EnrollmentLimit"
        $this._HasPlatform = $false
        $this._APITitle   = "Enrollment Limit"
        $this._PolicyName = "Device limit restrictions"
        # Same folder as platform restrictions — matches the OLD baseline export where
        # EnrollmentRestrictions/ bundled limit + platform restriction policies together.
        $this._Folder     = "EnrollmentRestrictions"
        $this._QueryList  = "?`$filter=deviceEnrollmentConfigurationType eq 'limit'"
        if($null -ne $this._PolicyGroup) { $this._PolicyGroup.AddPolicyType($this) }
    }
    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if ($PolicyObject.'@odata.type' -eq '#microsoft.graph.deviceEnrollmentLimitConfiguration') { return $true }
        return ($PolicyObject.deviceEnrollmentConfigurationType -in @('limit', 'defaultLimit'))
    }
}

class WindowsHelloForBusinessType : DeviceEnrollmentType
{
    WindowsHelloForBusinessType() : Base() { ([WindowsHelloForBusinessType]$this).Init() }
    Init()
    {
        $this._ID           = "WindowsHelloForBusiness"
        $this._APITitle     = "Windows Hello for Business"
        $this._PolicyName   = "Windows Hello for Business"
        $this._Folder       = "WindowsHelloForBusiness"
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        $this._QueryList    = "?`$filter=deviceEnrollmentConfigurationType eq 'windowsHelloForBusiness'"
        if($null -ne $this._PolicyGroup) { $this._PolicyGroup.AddPolicyType($this) }
    }
    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if ($PolicyObject.'@odata.type' -eq '#microsoft.graph.deviceEnrollmentWindowsHelloForBusinessConfiguration') { return $true }
        return ($PolicyObject.deviceEnrollmentConfigurationType -eq 'windowsHelloForBusiness')
    }
}

class CoManagementSettingsType : DeviceEnrollmentType
{
    CoManagementSettingsType() : Base() { ([CoManagementSettingsType]$this).Init() }
    Init()
    {
        $this._ID           = "CoManagementSettings"
        $this._APITitle     = "Co-Management Settings"
        $this._PolicyName   = "Co-Management Settings"
        $this._Folder       = "CoManagementSettings"
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        $this._QueryList    = "?`$filter=deviceEnrollmentConfigurationType eq 'deviceComanagementAuthorityConfiguration'"
        if($null -ne $this._PolicyGroup) { $this._PolicyGroup.AddPolicyType($this) }
    }
    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if ($PolicyObject.'@odata.type' -eq '#microsoft.graph.deviceComanagementAuthorityConfiguration') { return $true }
        return ($PolicyObject.deviceEnrollmentConfigurationType -eq 'deviceComanagementAuthorityConfiguration')
    }
}

class WindowsRestoreType : DeviceEnrollmentType
{
    WindowsRestoreType() : Base() { ([WindowsRestoreType]$this).Init() }
    Init()
    {
        $this._ID           = "WindowsRestore"
        $this._APITitle     = "Windows Restore"
        $this._PolicyName   = "Windows Restore"
        $this._Folder       = "WindowsRestore"
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        $this._QueryList    = "?`$filter=deviceEnrollmentConfigurationType eq 'windowsRestore'"
        if($null -ne $this._PolicyGroup) { $this._PolicyGroup.AddPolicyType($this) }
    }
    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if ($PolicyObject.'@odata.type' -eq '#microsoft.graph.windowsRestoreDeviceEnrollmentConfiguration') { return $true }
        return ($PolicyObject.deviceEnrollmentConfigurationType -eq 'windowsRestore')
    }
}

#########################################################################################
#
# Enrollment Notification
#
#########################################################################################

# region Enrollment Notification
# Note: Email and Push notifications are defined in the Notifications class in the Compliance class file.
# This bucket follows the same shape as the other DeviceEnrollment subtypes:
# no _QueryList (so the URL coalesces with siblings in Get-GraphPolicies), client-side
# filter via CheckPolicy on @odata.type, shared DeviceEnrollmentObject.
class EnrollmentNotificationType : DeviceEnrollmentType
{
    EnrollmentNotificationType() : Base() { ([EnrollmentNotificationType]$this).Init() }
    Init()
    {
        $this._ID         = "EnrollmentNotification"
        $this._HasPlatform = $false
        $this._APITitle   = "Enrollment notifications"
        $this._PolicyName = "Enrollment notification"
        $this._Folder     = "EnrollmentNotifications"
        # Cross-platform (email + push notifications target any enrolled device).
        $this._QueryList  = "?`$filter=deviceEnrollmentConfigurationType eq 'enrollmentNotificationsConfiguration'"
        if($null -ne $this._PolicyGroup) { $this._PolicyGroup.AddPolicyType($this) }
    }
    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if ($PolicyObject.'@odata.type' -eq '#microsoft.graph.deviceEnrollmentNotificationConfiguration') { return $true }
        return ($PolicyObject.deviceEnrollmentConfigurationType -eq 'enrollmentNotificationsConfiguration')
    }

    # ToDo: Add support for importing, exporting, copying Notifications between environment eg
    # notificationTemplates property has a string list of actual notification template policies Email_<GUID of Notification Template>
}

#########################################################################################
#
# Generic functions
#
#########################################################################################
function Set-EnrollmentRestrictionsPriority
{
    param($PolicyObject, $SourceObj)

    if($PolicyObject.Object.Priority -eq 0) { return }

    $api = "$($PolicyObject.PolicyType.API)/$($PolicyObject.Id)/setpriority"

    $priority = [PSCustomObject]@{
        priority = $SourceObj.Object.Priority
    }
    $json = $priority | ConvertTo-Json -Depth 20

    Write-Log "Update priority for $($PolicyObject.Name) to $($PolicyObject.Object.Priority)"
    Invoke-MSGraphAPI -Url $api -HttpMethod "POST" -Content $json -TokenId $PolicyObject._TokenId
}

#########################################################################################
#
# Android Device Owner Enrollment Profiles
#
#########################################################################################
#
# Profile used to enrol corporate-owned Android devices (dedicated devices,
# fully-managed, AOSP, Teams devices) via QR code or token. Listed flat at
# /deviceManagement/androidDeviceOwnerEnrollmentProfiles — not part of the
# deviceEnrollmentConfigurations multi-subtype tree.

# region AndroidDeviceOwnerEnrollmentProfilesType
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class AndroidDeviceOwnerEnrollmentProfilesType : IntunePolicyTypeBase
{
    AndroidDeviceOwnerEnrollmentProfilesType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceEnrollmentGroup")
        $this._PolicyName  = "Android Enterprise — corporate"
        $this._ID          = "AndroidDeviceOwnerEnrollmentProfiles"
        $this._API         = "deviceManagement/androidDeviceOwnerEnrollmentProfiles"
        # AndroidCOWP icon (Corporate-Owned With Profile) is the closest
        # existing match for the Device Owner enrolment surface.
        $this._Icon        = "AndroidCOWP"
        $this._Permissions = @("DeviceManagementServiceConfig.ReadWrite.All")
        # Server-generated bits that come back on GET but Graph rejects on
        # POST/PATCH. enrolledDeviceCount + token{*} + qrCode* are derived;
        # accountId is set from the calling tenant.
        $this._PropertiesToRemove          = @('accountId','enrolledDeviceCount','enrollmentTokenUsageCount','qrCodeContent','qrCodeImage','tokenCreationDateTime','tokenExpirationDateTime','tokenValue')
        $this._PropertiesToRemoveForUpdate = @('accountId','enrolledDeviceCount','enrollmentTokenUsageCount','qrCodeContent','qrCodeImage','tokenCreationDateTime','tokenExpirationDateTime','tokenValue','enrollmentMode','enrollmentTokenType')
        # Profile, not policy — no group assignments.
        $this._SupportsAssignments = $false
        # Graph returns HTTP 400 on `androidDeviceOwnerEnrollmentProfiles?$expand=assignments`,
        # even though SupportsAssignments=$false; the list-URL builder still
        # appends the expand unless this is explicitly suppressed.
        $this._ExpandAssignmentsList = $false
        $this._ObjectClass = "AndroidDeviceOwnerEnrollmentProfileObject"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class AndroidDeviceOwnerEnrollmentProfileObject : IntunePolicyBase
{
    AndroidDeviceOwnerEnrollmentProfileObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AndroidDeviceOwnerEnrollmentProfileObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AndroidDeviceOwnerEnrollmentProfilesType")
        # Language pack uses Platform.androidForWork for AOSP/Android Enterprise
        # surfaces; fall back to a literal if the key is missing in en-US.
        $this._PlatformName = Get-LanguageString "Platform.androidForWork" -IgnoreMissing
        if(-not $this._PlatformName) { $this._PlatformName = "Android Enterprise" }
    }
}

#########################################################################################
#
# Android For Work Enrollment Profiles
#
#########################################################################################
#
# Profile used to enrol personal Android devices into a managed Work Profile
# (BYOD). Endpoint at /deviceManagement/androidForWorkEnrollmentProfiles. No
# scope tag support, no assignments — purely token + QR for the end user.

# region AndroidForWorkEnrollmentProfilesType
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class AndroidForWorkEnrollmentProfilesType : IntunePolicyTypeBase
{
    AndroidForWorkEnrollmentProfilesType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceEnrollmentGroup")
        $this._PolicyName  = "Android Enterprise — work profile"
        $this._ID          = "AndroidForWorkEnrollmentProfiles"
        $this._API         = "deviceManagement/androidForWorkEnrollmentProfiles"
        # AndroidGooglePlay icon — work-profile enrolment is the personal-device
        # / Play-store-managed surface, so the GP icon reads better than the
        # corporate AndroidCOWP one used for the Device Owner type.
        $this._Icon        = "AndroidGooglePlay"
        $this._Permissions = @("DeviceManagementServiceConfig.ReadWrite.All")
        # Schema has no roleScopeTagIds, so disable the scope-tag column /
        # detail-view widget for this type. Leaving the default ("roleScopeTagIds")
        # would surface an empty UI and a 400 on save.
        $this._ScopeTagProperty = $null
        # Server-generated fields Graph rejects on POST/PATCH.
        $this._PropertiesToRemove          = @('accountId','enrolledDeviceCount','qrCodeContent','qrCodeImage','tokenValue','tokenExpirationDateTime')
        $this._PropertiesToRemoveForUpdate = @('accountId','enrolledDeviceCount','qrCodeContent','qrCodeImage','tokenValue','tokenExpirationDateTime')
        $this._SupportsAssignments = $false
        # Graph returns HTTP 400 on `androidForWorkEnrollmentProfiles?$expand=assignments`.
        $this._ExpandAssignmentsList = $false
        $this._ObjectClass = "AndroidForWorkEnrollmentProfileObject"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class AndroidForWorkEnrollmentProfileObject : IntunePolicyBase
{
    AndroidForWorkEnrollmentProfileObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AndroidForWorkEnrollmentProfileObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AndroidForWorkEnrollmentProfilesType")
        $this._PlatformName = Get-LanguageString "Platform.androidForWork" -IgnoreMissing
        if(-not $this._PlatformName) { $this._PlatformName = "Android Enterprise" }
    }
}

#########################################################################################
#
# Settings Catalog
#
#########################################################################################

# region Settings Catalog

class EnrollmentSettingsCatalogType : SettingsCatalogTypeBase
{
    EnrollmentSettingsCatalogType() : Base()
    {
        ([EnrollmentSettingsCatalogType]$this).Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceEnrollmentGroup")
        $this._APITitle = "Enrollment Policies (Settings Catalog)"
        $this._ID = "EnrollmentSettingsCatalog"
        # Startup default only - Invoke-IntuneSettingsCatalogAuthenticated rebuilds
        # this from _FamilyTypes once Graph reports the live template list.
        # windowsOsRecoveryPolicies is NOT listed here: it falls to SettingsCatalog's
        # catch-all spec, so claiming it would only fetch rows CheckPolicy rejects.
        $this._QueryList = "?`$filter=templateReference/templateFamily eq 'enrollmentConfiguration'"
        $this._Icon = "EnrollmentStatusPage"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyTypeOrder = 160
    }
}

<#
class AutopilotDevicePreparationSettingsCatalogType : SettingsCatalogTypeBase
{
    AutopilotDevicePreparationSettingsCatalogType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceEnrollmentGroup")
        $this._ID = "AutopilotDevicePreparationSettingsCatalog"
        $this._APITitle = "Autopilot Device Preparation (Settings Catalog)"
        #$this._QueryList = "?`$filter=(technologies has 'enrollment') and (platforms eq 'windows10') and (TemplateReference/templateId eq '80d33118-b7b4-40d8-b15f-81be745e053f_1') and (Templatereference/templateFamily eq 'enrollmentConfiguration')"
        $this._QueryList = "?`$filter=(technologies has 'enrollment') and (platforms eq 'windows10') and (Templatereference/templateFamily eq 'enrollmentConfiguration')"
        $this._Folder = "SettingsCatalog"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyTypeOrder = 100
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.'@odata.type' -ne "#microsoft.graph.deviceManagementConfigurationPolicy") { return $false }

        if($PolicyObject.templateReference.templateFamily -and $PolicyObject.templateReference.templateFamily -eq 'enrollmentConfiguration') {
            return $true
        }
        
        return $false
    }
}
#>
#endregion