#ImportOrder 220

#########################################################################################
#
# Windows Update Group
#
#########################################################################################

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class WindowsUpdateGroup : IntunePolicyGroupBase
{
    WindowsUpdateGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "WindowsUpdates"
        $this._Name = "Windows 10 and later updates"
        $this._Icon = "UpdatePolicies"
    }
}

#########################################################################################
#
# Update Rings
#
#########################################################################################

# region Update Rings
class WindowsUpdatePolicyType : IntunePolicyTypeBase
{
    WindowsUpdatePolicyType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "WindowsUpdateGroup")
        $this._PolicyName = "Update rings"
        $this._ID = "UpdatePolicies"
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing   # every member of the group is Windows
        $this._API = "deviceManagement/deviceConfigurations"
        $this._QueryList = "?`$filter=isof(%27microsoft.graph.windowsUpdateForBusinessConfiguration%27)"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._ObjectClass = "WindowsUpdatePolicyObject"
        $this._PropertiesToRemove = @('version','qualityUpdatesPauseStartDate','featureUpdatesPauseStartDate','qualityUpdatesWillBeRolledBack','featureUpdatesWillBeRolledBack')
        $this._Icon = "UpdatePolicies"
        $this._ExpandAssignmentsList = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyTypeOrder = 90
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.'@odata.type' -ne "#microsoft.graph.windowsUpdateForBusinessConfiguration") { return $false }

        return $true
    }
}

Class WindowsUpdatePolicyObject : IntunePolicyBase
{
    WindowsUpdatePolicyObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    WindowsUpdatePolicyObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        #$this._PlatformName = Get-LanguageString "Platform.windows"
        #Add-ObjectProperty $this "ScriptType" { "PowerShell script" }

        $this._PolicyType = (Get-SingletonObject "WindowsUpdatePolicyType")
    }
}

#########################################################################################
#
# Feature Updates
#
#########################################################################################

# region Feature Updates
class FeatureUpdateType : IntunePolicyTypeBase
{
    FeatureUpdateType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "WindowsUpdateGroup")
        $this._PolicyName = "Feature updates"
        $this._ID = "FeatureUpdates"
        # Platform default: the endpoint serves exactly one platform and the
        # objects carry no platforms/platformType field, so the column would
        # otherwise be blank (windowsFeatureUpdateProfiles is Windows-only).
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceManagement/windowsFeatureUpdateProfiles"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._ObjectClass = "FeatureUpdateObject"
        $this._Dependencies = @('Applications')
        # Read-only/derived properties the update PATCH must not send back.
        $this._PropertiesToRemoveForUpdate = @('deployableContentDisplayName','endOfSupportDate')
        $this._TopItems = 0
        # Graph returns HTTP 400 on `windowsFeatureUpdateProfiles?$expand=assignments`;
        # assignments are loaded via Add-GraphPolicyAssignments after listing.
        $this._ExpandAssignmentsList = $false
        # The endpoint caps `$top` at 200; the bulk-export default of 1000 produces
        # `400: The limit of '200' for Top query has been exceeded`. Disabling page-
        # size for this type drops $top entirely — Graph applies its own default.
        $this._HasPageSizeSupport = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportAssignmentsCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        return (@{ "API" = "$($this.API)/$($PolicyObject.Id)/microsoft.graph.managedDeviceMobileAppConfiguration/assign" })
    }
}

Class FeatureUpdateObject : IntunePolicyBase
{
    FeatureUpdateObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    FeatureUpdateObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "FeatureUpdateType")
    }
}

#########################################################################################
#
# Quality Updates
#
#########################################################################################

# region Quality Update Profiles
class QualityUpdateProfileType : IntunePolicyTypeBase
{
    QualityUpdateProfileType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "WindowsUpdateGroup")
        $this._PolicyName = "Quality updates (Profile)"
        $this._ID = "QualityUpdates"
        # Platform default: the endpoint serves exactly one platform and the
        # objects carry no platforms/platformType field, so the column would
        # otherwise be blank (windowsQualityUpdateProfiles is Windows-only).
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceManagement/windowsQualityUpdateProfiles"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._ObjectClass = "QualityUpdateProfileObject"
        $this._PropertiesToRemoveForUpdate = @('releaseDateDisplayName','deployableContentDisplayName')
        $this._TopItems = 0
        $this._Icon = "UpdatePolicies"
        # Graph returns HTTP 400 on `windowsQualityUpdateProfiles?$expand=assignments`.
        $this._ExpandAssignmentsList = $false
        # Endpoint caps `$top` at 200; bulk-export default of 1000 returns 400.
        $this._HasPageSizeSupport = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class QualityUpdateProfileObject : IntunePolicyBase
{
    QualityUpdateProfileObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    QualityUpdateProfileObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "QualityUpdateProfileType")
    }
}

# region Quality Update Policies
class QualityUpdatePolicyType : IntunePolicyTypeBase
{
    QualityUpdatePolicyType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "WindowsUpdateGroup")
        $this._PolicyName = "Quality updates (Policy)"
        $this._ID = "QualityUpdatePolicies"
        # Platform default: the endpoint serves exactly one platform and the
        # objects carry no platforms/platformType field, so the column would
        # otherwise be blank (windowsQualityUpdatePolicies is Windows-only).
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceManagement/windowsQualityUpdatePolicies"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._ObjectClass = "QualityUpdatePolicyObject"
        $this._TopItems = 0
        $this._Icon = "UpdatePolicies"
        # Graph returns HTTP 400 on `windowsQualityUpdatePolicies?$expand=assignments`.
        $this._ExpandAssignmentsList = $false
        # Endpoint caps `$top` at 200; bulk-export default of 1000 returns 400.
        $this._HasPageSizeSupport = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class QualityUpdatePolicyObject : IntunePolicyBase
{
    QualityUpdatePolicyObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    QualityUpdatePolicyObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "QualityUpdatePolicyType")
    }
}

#########################################################################################
#
# Driver Updates
#
#########################################################################################

# region Driver Updates
class DriverUpdateType : IntunePolicyTypeBase
{
    DriverUpdateType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "WindowsUpdateGroup")
        $this._PolicyName = "Driver updates"
        $this._ID = "DriverUpdateProfiles"
        # Platform default: the endpoint serves exactly one platform and the
        # objects carry no platforms/platformType field, so the column would
        # otherwise be blank (windowsDriverUpdateProfiles is Windows-only).
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceManagement/windowsDriverUpdateProfiles"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._ObjectClass = "DriverUpdateObject"
        $this._PropertiesToRemoveForUpdate = @('releaseDateDisplayName','deployableContentDisplayName')
        $this._TopItems = 0
        $this._Icon = "UpdatePolicies"
        # Graph returns HTTP 400 on `windowsDriverUpdateProfiles?$expand=assignments`.
        $this._ExpandAssignmentsList = $false
        # Endpoint caps `$top` at 200; bulk-export default of 1000 returns 400.
        $this._HasPageSizeSupport = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class DriverUpdateObject : IntunePolicyBase
{
    DriverUpdateObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    DriverUpdateObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "DriverUpdateType")
    }
}

#region Maintenance windows (Settings Catalog)

# Windows Update maintenance windows are a Settings Catalog template family
# (maintenanceWindows, backed by the Update/MaintenanceWindow* CSP, Windows 11
# 24H2 + KB5077181). Without this class they fall into SettingsCatalogType's
# catch-all and list under Configuration; every other family is routed to its
# domain group (enrollmentConfiguration -> Device enrollment, endpointSecurity*
# -> Endpoint Security, deviceConfigurationScripts -> Scripts), so this one
# belongs here.
#
# _QueryList is the pre-login default. Invoke-IntuneSettingsCatalogAuthenticated
# rebuilds it from the tenant's live template list and fills _FamilyTypes, which
# is what CheckPolicy matches on; if the tenant has no such template the default
# filter simply returns nothing. SettingsCatalogObject resolves its PolicyType by
# family, so no dedicated object class is needed.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class WindowsUpdateSettingsCatalogType : SettingsCatalogTypeBase
{
    WindowsUpdateSettingsCatalogType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "WindowsUpdateGroup")
        $this._PolicyName = "Maintenance window"
        $this._APITitle = "Maintenance windows"
        $this._ID = "WindowsUpdateSettingsCatalog"
        $this._QueryList = "?`$filter=templateReference/templateFamily eq 'maintenanceWindows'"
        $this._Icon = "UpdatePolicies"
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyTypeOrder = 60
    }
}

#endregion
