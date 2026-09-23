#ImportOrder 220

#########################################################################################
#
# Apple Update Group
#
#########################################################################################

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class AppleUpdateGroup : IntunePolicyGroupBase
{
    AppleUpdateGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "AppleUpdates"
        $this._Name = "Apple updates"
        $this._Icon = "AppleUpdates"
    }
}


#########################################################################################
#
# iOS/iPadOS Updates
#
#########################################################################################

# region iOS/iPadOS Updates
class iOSiPadOSPolicyType : IntunePolicyTypeBase
{
    iOSiPadOSPolicyType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "AppleUpdateGroup")
        $this._PolicyName = "iOS/iPadOS update policies"
        $this._ID = "iOSiPadOSUpdatePolicies"
        $this._API = "deviceManagement/deviceConfigurations"
        $this._QueryList = "?`$filter=isof(%27microsoft.graph.iosUpdateConfiguration%27)"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._ObjectClass = "iOSiPadOSPolicyObject"
        $this._Icon = "iOSUpdates"
        $this._ExpandAssignmentsList = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyTypeOrder = 90
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.'@odata.type' -ne "#microsoft.graph.iosUpdateConfiguration") { return $false }

        return $true
    }
}

Class iOSiPadOSPolicyObject : IntunePolicyBase
{
    iOSiPadOSPolicyObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    iOSiPadOSPolicyObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "iOSiPadOSPolicyType")
    }
}


#########################################################################################
#
# MacOS Updates
#
#########################################################################################

# region MacOS Updates
class macOSPolicyType : IntunePolicyTypeBase
{
    macOSPolicyType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "AppleUpdateGroup")
        $this._PolicyName = "macOS update policies"
        $this._ID = "macOSUpdatePolicies"
        $this._API = "deviceManagement/deviceConfigurations"
        $this._QueryList = "?`$filter=isof(%27microsoft.graph.macOSSoftwareUpdateConfiguration%27)"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._ObjectClass = "macOSPolicyObject"
        $this._Icon = "MacOSUpdates"
        $this._ExpandAssignmentsList = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyTypeOrder = 90
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.'@odata.type' -ne "#microsoft.graph.macOSSoftwareUpdateConfiguration") { return $false }

        return $true
    }
}

Class macOSPolicyObject : IntunePolicyBase
{
    macOSPolicyObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    macOSPolicyObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "macOSPolicyType")
    }
}
