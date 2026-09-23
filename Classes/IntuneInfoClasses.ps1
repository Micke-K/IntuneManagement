#ImportOrder 220

#########################################################################################
#
# Intune Info Group
#
# Read-only objects ported from the original project's "Intune Info" view
# (Extensions/EndpointManagerInfo.psm1). Every type here is Export/View only:
# no import, delete, copy or assignment support. Android Google Play status
# and Tenant Settings are single-object endpoints (_SingleObject; the body IS
# the row) with a synthesized displayName because the API has none.
#
#########################################################################################

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class IntuneInfoGroup : IntunePolicyGroupBase
{
    IntuneInfoGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "IntuneInfo"
        # Templates are this group's substance: Version fills for both template
        # kinds, State for templates, VPP tokens and the Google Play binding.
        $this._ExtraColumns = @("Version", "State")
        $this._Name = "Intune Info"
        $this._Icon = "Report"
        $this._ShowButtons = @("Export","View")
    }
}

#########################################################################################
#
# Baseline Templates - Intent
#
#########################################################################################

class BaselineTemplatesType : IntunePolicyTypeBase
{
    BaselineTemplatesType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "IntuneInfoGroup")
        $this._PolicyName = "Baseline Templates - Intent"
        $this._ID = "BaselineTemplates"
        $this._ExtraColumns = @("Version", "State")
        $this._API = "deviceManagement/templates"
        $this._Icon = "SecurityBaselines"
        $this._Permissions = @("DeviceManagementConfiguration.Read.All")
        $this._ObjectClass = "BaselineTemplatesObject"
        $this._ShowButtons = @("Export","View")
        $this._SupportsAssignments = $false
        $this._ExpandAssignmentsList = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class BaselineTemplatesObject : IntunePolicyBase
{
    BaselineTemplatesObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    BaselineTemplatesObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "BaselineTemplatesType")
        # Same vocabulary as the Settings Catalog templates, so one Version / State
        # column reads evenly across the Intune Info group.
        Add-ObjectProperty $this "Version" { if($this.JsonObject) { $this.JsonObject.versionInfo } }
        Add-ObjectProperty $this "State"   { if($null -eq $this.JsonObject) { $null } elseif($this.JsonObject.isDeprecated -eq $true) { "Deprecated" } else { "Active" } }
    }
}

#########################################################################################
#
# Baseline Templates - Settings Catalog
#
#########################################################################################

class BaselineTemplatesSettingsCatalogType : IntunePolicyTypeBase
{
    BaselineTemplatesSettingsCatalogType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "IntuneInfoGroup")
        $this._PolicyName = "Templates - Settings Catalog"
        $this._ID = "BaselineTemplatesSettingsCatalog"
        $this._ExtraColumns = @("Version", "State")
        $this._API = "deviceManagement/configurationPolicyTemplates"
        $this._Icon = "SecurityBaselines"
        $this._Permissions = @("DeviceManagementConfiguration.Read.All")
        $this._ObjectClass = "BaselineTemplatesSettingsCatalogObject"
        $this._ShowButtons = @("Export","View")
        $this._SupportsAssignments = $false
        $this._ExpandAssignmentsList = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class BaselineTemplatesSettingsCatalogObject : IntunePolicyBase
{
    BaselineTemplatesSettingsCatalogObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    BaselineTemplatesSettingsCatalogObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "BaselineTemplatesSettingsCatalogType")
        Add-ObjectProperty $this "Version" { if($this.JsonObject) { $this.JsonObject.displayVersion } }
        Add-ObjectProperty $this "State"   { if($this.JsonObject) { ConvertTo-DisplayWords $this.JsonObject.lifecycleState } }
    }
}

#########################################################################################
#
# Apple VPP Tokens
#
#########################################################################################

class AppleVPPTokensType : IntunePolicyTypeBase
{
    AppleVPPTokensType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "IntuneInfoGroup")
        $this._PolicyName = "Apple VPP Tokens"
        $this._ID = "AppleVPPTokens"
        $this._ExtraColumns = @("State")
        $this._HasPlatform = $false
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceAppManagement/vppTokens"
        $this._Icon = "AppleVPPTokens"
        $this._Permissions = @("DeviceManagementApps.Read.All")
        $this._ObjectClass = "AppleVPPTokensObject"
        $this._NameProperty = "appleId"
        $this._ShowButtons = @("Export","View")
        $this._SupportsAssignments = $false
        $this._ExpandAssignmentsList = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class AppleVPPTokensObject : IntunePolicyBase
{
    AppleVPPTokensObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AppleVPPTokensObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AppleVPPTokensType")
        Add-ObjectProperty $this "State" { if($this.JsonObject) { ConvertTo-DisplayWords $this.JsonObject.state } }
    }
}

#########################################################################################
#
# Apple Enrollment Tokens (DEP / Apple Business Manager)
#
#########################################################################################

class AppleEnrollmentTokensType : IntunePolicyTypeBase
{
    AppleEnrollmentTokensType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "IntuneInfoGroup")
        $this._PolicyName = "Apple Enrollment Tokens"
        $this._ID = "AppleEnrollmentTokens"
        $this._ExtraColumns = @("State")
        $this._HasPlatform = $false
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceManagement/depOnboardingSettings"
        $this._QueryList = "?`$top=100"
        $this._Icon = "AppleEnrollmentTokens"
        $this._Permissions = @("DeviceManagementServiceConfig.Read.All")
        $this._ObjectClass = "AppleEnrollmentTokensObject"
        $this._NameProperty = "tokenName"
        $this._ShowButtons = @("Export","View")
        $this._SupportsAssignments = $false
        $this._ExpandAssignmentsList = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class AppleEnrollmentTokensObject : IntunePolicyBase
{
    AppleEnrollmentTokensObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AppleEnrollmentTokensObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AppleEnrollmentTokensType")
        # A DEP token has no state field; whether it has expired is the state that
        # matters, and it lines up with the VPP token's Valid / Expired.
        Add-ObjectProperty $this "State" {
            if($null -eq $this.JsonObject -or -not $this.JsonObject.tokenExpirationDateTime) { return $null }
            $exp = [DateTime]::MinValue
            if([DateTime]::TryParse([string]$this.JsonObject.tokenExpirationDateTime, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$exp)) {
                if($exp -lt (Get-Date)) { "Expired" } else { "Valid" }
            }
        }
    }
}

#########################################################################################
#
# Android Google Play (managed store account status - single object)
#
#########################################################################################

class AndroidGooglePlayType : IntunePolicyTypeBase
{
    AndroidGooglePlayType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "IntuneInfoGroup")
        $this._PolicyName = "Android Google Play"
        $this._ID = "AndroidGooglePlay"
        $this._ExtraColumns = @("State")
        # Platform default: the endpoint serves exactly one platform and the
        # objects carry no platforms/platformType field, so the column would
        # otherwise be blank (Managed Google Play is the Android Enterprise
        # connection).
        $this._PlatformName = Get-LanguageString "Platform.androidEnterprise" -IgnoreMissing
        $this._API = "deviceManagement/androidManagedStoreAccountEnterpriseSettings"
        $this._Icon = "AndroidGooglePlay"
        $this._Permissions = @("DeviceManagementConfiguration.Read.All")
        $this._ObjectClass = "AndroidGooglePlayObject"
        $this._ShowButtons = @("Export","View")
        $this._SupportsAssignments = $false
        $this._ExpandAssignmentsList = $false
        $this._SingleObject = $true
        $this._HasPageSizeSupport = $false
        $this._SkipAddIDOnFileName = $true

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [IntunePolicyBase]GetObject([PSCustomObject]$JsonObj)
    {
        # The settings object has no displayName - synthesize one so the grid
        # row and the export file name aren't blank.
        if($JsonObj -and -not $JsonObj.PSObject.Properties['displayName']) {
            $JsonObj | Add-Member -MemberType NoteProperty -Name 'displayName' -Value 'Managed Google Play' -Force
        }
        return ([IntunePolicyTypeBase]$this).GetObject($JsonObj)
    }
}

Class AndroidGooglePlayObject : IntunePolicyBase
{
    AndroidGooglePlayObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AndroidGooglePlayObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AndroidGooglePlayType")
        Add-ObjectProperty $this "State" { if($this.JsonObject) { ConvertTo-DisplayWords $this.JsonObject.bindStatus } }
    }
}

#########################################################################################
#
# Tenant Settings (Intune service settings - single object)
#
#########################################################################################

class TenantSettingsType : IntunePolicyTypeBase
{
    TenantSettingsType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "IntuneInfoGroup")
        $this._PolicyName = "Tenant Settings"
        $this._ID = "TenantSettings"
        $this._HasPlatform = $false
        $this._API = "deviceManagement/settings"
        $this._Icon = "TenantSettings"
        $this._Permissions = @("DeviceManagementConfiguration.Read.All")
        $this._ObjectClass = "TenantSettingsObject"
        $this._ShowButtons = @("Export","View")
        $this._SupportsAssignments = $false
        $this._ExpandAssignmentsList = $false
        $this._SingleObject = $true
        $this._HasPageSizeSupport = $false
        $this._SkipAddIDOnFileName = $true

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [IntunePolicyBase]GetObject([PSCustomObject]$JsonObj)
    {
        # deviceManagement/settings has neither id nor displayName - synthesize
        # the name so the grid row and the export file name aren't blank.
        if($JsonObj -and -not $JsonObj.PSObject.Properties['displayName']) {
            $JsonObj | Add-Member -MemberType NoteProperty -Name 'displayName' -Value 'Intune Tenant Settings' -Force
        }
        return ([IntunePolicyTypeBase]$this).GetObject($JsonObj)
    }
}

Class TenantSettingsObject : IntunePolicyBase
{
    TenantSettingsObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    TenantSettingsObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "TenantSettingsType")
    }
}
