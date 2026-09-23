#ImportOrder 205

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class EndpointSecurityGroup : IntunePolicyGroupBase
{
    EndpointSecurityGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "EndpointSecurity"
        $this._Name = "Endpoint Security"
        $this._Icon = "EndpointSecurity"
        $this._ExtraColumns = @("TemplateFamily=Parent type")   # two of three members supply it

    }
}

#########################################################################################
#
# Intents
#
#########################################################################################

# region Intents

class EndpointSecurityType : IntunePolicyTypeBase
{
    EndpointSecurityType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "EndpointSecurityGroup")
        $this._PolicyName = "Endpoint Security"
        $this._APITitle = "Endpoint Security (Intents)"

        $this._ID = "EndpointSecurity"
        $this._API = "deviceManagement/intents"
        $this._PolicyBaseName = "Intents"
        $this._PropertiesToRemove = @('Settings','@OData.Type')
        # Graph: "Properties not patchable specified: IsAssigned,
        # IsMigratingToConfigurationPolicy, TemplateId". Settings are updated
        # via the /updateSettings action, not the intent PATCH.
        $this._PropertiesToRemoveForUpdate = @('isAssigned','isMigratingToConfigurationPolicy','templateId','settings')
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._SubTypeColumn = "PolicyName=Type"   # per-row template name (Antivirus, Firewall, ...)
        $this._ExtraColumns  = @("TemplateFamily=Parent type")
        $this._Expand = "Settings"
        $this._Icon = "EndpointSecurity"
        $this._Dependencies = @("ReusableSettings")
        $this._ObjectClass = "IntentObject"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]GetCompareConfig()
    {
        return @{
            Prop        = "settings"
            GetKey      = { param($s) "$($s.definitionId)" }
            GetValue    = { param($s) Get-IntentSettingValue $s }
            GetCategory = { param($s) Get-IntentSettingCategory $s }
        }
    }

    Hidden [PSCustomObject]GetTemplate($TemplateId)
    {
        # Tag the baseline-template cache with TenantCache_<tenantId> so it gets wiped by
        # Clear-TenantCache on disconnect — previous code stored it untagged and the
        # previous tenant's templates would leak across tenant switches.
        $tenantId = $script:OrganizationId
        $cacheId = "BaseLineTemplates_$tenantId"
        $baseLineTemplates = Get-CacheObject $cacheId
        if(-not $baseLineTemplates)
        {
            $baseLineTemplates = (Invoke-MSGraphAPI -Url "/deviceManagement/templates").Value
            if($baseLineTemplates) {
                Set-CacheObject $cacheId $baseLineTemplates "TenantCache_$tenantId"
            }
        }

        if(-not $baseLineTemplates) { return $null}

        return ($baseLineTemplates | Where-Object Id -eq $TemplateId)
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        return (@{
            "API"="deviceManagement/templates/$($PolicyObject.JsonObject.templateId)/createInstance"
        })
    }

    PostImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {      
        $this.UpdateSettings($PolicyObject, $SourceObject.JsonObject.Settings)
    }

    Hidden UpdateSettings($PolicyObject, $Settings)
    {
        if(($Settings | Measure-Object).Count -eq 0) { return } 

        $clonedSettings = @()
        $Settings | ConvertTo-Json -Depth 50 | ConvertFrom-Json | ForEach-Object { $clonedSettings += $_ }
        $newSettings = ([HashTable]@{
            "settings" = $clonedSettings
        })            
        Remove-GraphPropertiesForImport $PolicyObject $newSettings.Settings -KeepProperties "@odata.type"
        
        $response = Invoke-MSGraphAPI -Url "$($this.API)/$($PolicyObject.id)/updateSettings" -Body ($newSettings | ConvertTo-Json -Depth 50) -Method "POST" -FullResponseObject -TokenId $PolicyObject._TokenID
        if($response.Success) {
            Write-Log "Settings updated successfully"
        }
    }
}

class IntentObject : IntunePolicyBase
{
    Hidden [PSCustomObject]$_BaselineTemplate = $null

    IntentObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    IntentObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "EndpointSecurityType")        

        if($this._PolicyType -and $this.JsonObject.templateId) {
            $this._BaselineTemplate = $this._PolicyType.GetTemplate($this.JsonObject.templateId)

            if($this._BaselineTemplate) {
                Add-ObjectProperty $this "BaselineTemplate" { $this._BaselineTemplate }
                Add-ObjectProperty $this "TemplateFamily" { (Get-EndpointSecurityCategoryName (?: ($this.BaselineTemplate.templateSubtype -eq "none") $this.BaselineTemplate.templateType $this.BaselineTemplate.templateSubtype)) } # ToDo: Get actual language string
                Add-ObjectProperty $this "Category" { $this.TemplateFamily }
                Add-ObjectProperty $this "TemplateVersion" { $this.BaselineTemplate.versionInfo }
                $this._PlatformName = Get-LanguageString "Platform.$($this._BaselineTemplate.platformType)" -IgnoreMissing
                
            }
        }

        $this._PolicyName = ?? $this.BaselineTemplate.displayName $this._PolicyType.PolicyBaseType
    }
}

function Get-EndpointSecurityCategoryName
{
    param($TemplateType)

    if(-not $TemplateType)
    {
        Write-Log "Get-EndpointSecurityCategoryName called with empty Category" 2
        return
    }

    $returnString = $null

    if($TemplateType.StartsWith("endpointSecurity"))
    {
        $TemplateType = $TemplateType.Substring(16)
    }

    if($TemplateType -eq "none")
    {
        return ""
    }
    elseif($TemplateType -eq "accountProtection")
    {
        $returnString = Get-LanguageString "SecurityTemplate.accountProtection"
    }
    elseif($TemplateType -eq "antivirus")
    {
        $returnString = Get-LanguageString "SecurityTemplate.antivirus"
    }
    elseif($TemplateType -eq "diskEncryption")
    {
        $returnString = Get-LanguageString "SecurityTemplate.diskEncryption"
    }
    elseif($TemplateType -eq "endpointDetectionReponse" -or $TemplateType -eq "EndpointDetectionAndResponse")
    {
        $returnString = Get-LanguageString "SecurityTemplate.eDR"
    }    
    elseif($TemplateType -eq "attackSurfaceReduction")
    {
        $returnString = Get-LanguageString "SecurityTemplate.aSR"
    }
    elseif($TemplateType -eq "attackSurfaceReduction")
    {
        $returnString = Get-LanguageString "SecurityTemplate.aSR"
    }
    elseif($TemplateType -eq "firewall")
    {
        $returnString = Get-LanguageString "SecurityTemplate.firewall"
    }
    elseif($TemplateType -eq "applicationControl")
    {
        $returnString = Get-LanguageString "PolicyType.applicationControl"
    }
    elseif($TemplateType -eq "securityBaseline" -or 
        $TemplateType -eq "advancedThreatProtectionSecurityBaseline" -or
        $TemplateType -eq "microsoftEdgeSecurityBaseline" -or
        $TemplateType -eq "baseline")
    {
        $returnString = Get-LanguageString "Titles.securityBaselines"
    }
    elseif($TemplateType -eq "enrollmentConfiguration")
    {
        $returnString = Get-LanguageString "SettingDetails.enrollment"
    }

    if([String]::IsNullOrEmpty($returnString))
    {
        Write-Log "Could not translate templateSubtype $TemplateType" 2
        return $TemplateType
    }

    return $returnString
}

#endregion

#########################################################################################
#
# Settings Catalog
#
#########################################################################################

# region Settings catalog
class EndpointSecuritySettingsCatalogType : SettingsCatalogTypeBase
{
    EndpointSecuritySettingsCatalogType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "EndpointSecurityGroup")
        $this._ID = "EndpointSecuritySettingsCatalog"
        # The template version a policy was created from - how you spot an
        # antivirus or baseline policy still on a superseded template.
        $this._ExtraColumns = @("Object.templateReference.templateDisplayVersion=Template version")
        $this._APITitle = "Endpoint Security (Settings Catalog)"
        $this._QueryList = "?`$filter=templateReference/templateFamily eq 'baseline' or templateReference/templateFamily eq 'endpointSecurityAccountProtection' or templateReference/templateFamily eq 'endpointSecurityAntivirus' or templateReference/templateFamily eq 'endpointSecurityDiskEncryption' or templateReference/templateFamily eq 'endpointSecurityEndpointDetectionAndResponse' or templateReference/templateFamily eq 'endpointSecurityAttackSurfaceReduction' or templateReference/templateFamily eq 'endpointSecurityFirewall' or templateReference/templateFamily eq 'endpointSecurityApplicationControl'"
        $this._Icon = "EndpointSecurity"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyTypeOrder = 100
    }
}

#########################################################################################
#
# Reusable Settings
#
#########################################################################################

# region Reusable Settings
class ReusableSettingsEndpointSecurityType : ReusableSettingsTypeBase
{
    ReusableSettingsEndpointSecurityType() : Base()
    {
        ([ReusableSettingsEndpointSecurityType]$this).Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "EndpointSecurityGroup")
        # !!! For now...Should check actual supported settingDefinitionId for Endpoint Security.
        $this._QueryList = "?`$filter=settingDefinitionId ne 'linux_customcompliance_discoveryscript_reusablesetting'"
        $this._ID = "ReusableSettingsEndpointSecurity"
        $this._HasPlatform = $false
        $this._Folder = "ReusableSettings"
        $this._ObjectClass = "ReusableSettingEndpointSecurityObject"
        $this._ImportOrder = 75
        $this._PolicyTypeOrder = 145
        $this._Icon = "EndpointSecurity"
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.'@odata.type' -ne "#microsoft.graph.deviceManagementReusablePolicySetting") { return $false }

        # !!! For now...Should check actual supported settingDefinitionId for Endpoint Security. deviceConfigurationScripts
        if($PolicyObject.settingDefinitionId -ne 'linux_customcompliance_discoveryscript_reusablesetting') {
            return $true
        }

        return $false
    }
}

Class ReusableSettingEndpointSecurityObject : ReusableSettingsObjectBase
{
    ReusableSettingEndpointSecurityObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    ReusableSettingEndpointSecurityObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "ReusableSettingsEndpointSecurityType")
    }
}

#endregion
