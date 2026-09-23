#ImportOrder 220

#########################################################################################
#
# Windows 365 Group
#
#########################################################################################

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class Windows365Group : IntunePolicyGroupBase
{
    Windows365Group() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "Windows365"
        $this._Name = "Windows 365"
        $this._Icon = "Devices"

    }
}

#########################################################################################
#
# W365 Provisioning Policy
#
#########################################################################################

# region W365 Provisioning Policy
class Win365ProvisioningType : IntunePolicyTypeBase
{
    Win365ProvisioningType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "Windows365Group")
        $this._APITitle = "W365 Provisioning Policies"
        $this._PolicyName = "W365 Provisioning"
        $this._ID = "W365ProvisioningPolicies"
        # Platform default: the endpoint serves exactly one platform and the
        # objects carry no platforms/platformType field, so the column would
        # otherwise be blank (Cloud PC provisioning is Windows-only).
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceManagement/virtualEndpoint/provisioningPolicies"
        $this._Permissions = @("CloudPC.ReadWrite.All")
        $this._Icon = "Devices"
        $this._ExpandAssignmentsList = $false
        $this._ObjectClass = "Win365ProvisioningObject"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class Win365ProvisioningObject : IntunePolicyBase
{
    Win365ProvisioningObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }
    Win365ProvisioningObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "Win365ProvisioningType")
    }
}

#endregion

#########################################################################################
#
# W365 User Settings
#
#########################################################################################

# region W365 User Settings
class Win365UserSettingsType : IntunePolicyTypeBase
{
    Win365UserSettingsType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "Windows365Group")
        $this._APITitle = "W365 User Settings"
        $this._PolicyName = "W365 User Setting"
        $this._ID = "W365UserSettings"
        # Platform default: the endpoint serves exactly one platform and the
        # objects carry no platforms/platformType field, so the column would
        # otherwise be blank (Cloud PC user settings are Windows-only).
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceManagement/virtualEndpoint/userSettings"
        $this._Permissions = @("CloudPC.ReadWrite.All")
        $this._Icon = "Devices"
        $this._ExpandAssignmentsList = $false
        $this._ObjectClass = "Win365UserSettingObject"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class Win365UserSettingObject : IntunePolicyBase
{
    Win365UserSettingObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }
    Win365UserSettingObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "Win365UserSettingsType")
    }
}

#endregion