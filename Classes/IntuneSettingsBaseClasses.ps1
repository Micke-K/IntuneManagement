#ImportOrder 110

#########################################################################################
#
# Settings Catalog
#
#########################################################################################

# region Settings Catalog

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class SettingsCatalogTypeBase : IntunePolicyTypeBase
{
    static [bool] $IsAbstract = $true
    Hidden [string[]]$_FamilyTypes = @()

    SettingsCatalogTypeBase() : Base()
    {
        ([SettingsCatalogTypeBase]$this).Init()
    }

    Init()
    {
        $this._NameProperty = "name"
        $this._PolicyName = "Settings Catalog"
        $this._ID = "SettingsCatalogBase"
        $this._API = "deviceManagement/configurationPolicies"
        $this._PolicyBaseName = "Settings Catalog"
        $this._PropertiesToRemove = @('settingCount')
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._SubTypeColumn = "TemplateFamily=Family"
        $this._Expand = "Settings"
        $this._Icon = "DeviceConfiguration"
        $this._Dependencies = @("ReusableSettings")
        $this._ObjectClass = "SettingsCatalogObject"
        $this._VerifyObject = $true

        $this._PolicyTypeOrder = 200
    }

    [Hashtable]PreUpdateCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        # Settings Catalog updates must PUT the full body (including settings);
        # a PATCH carrying the settings payload is rejected by Graph. Same
        # contract as the portal and the original project
        # (Start-PreUpdateSettingsCatalog).
        return @{ "Method" = "PUT" }
    }

    [Hashtable]GetCompareConfig()
    {
        return @{
            Prop        = "settings"
            GetKey      = { param($s) Get-SettingsCatalogSettingKey $s }
            GetValue    = { param($s) Get-SettingsCatalogSettingValue $s }
            GetCategory = { param($s) Get-SettingsCatalogSettingCategory $s }
        }
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.'@odata.type' -ne "#microsoft.graph.deviceManagementConfigurationPolicy") { return $false }

        if($PolicyObject.templateReference.templateFamily -notin $this._FamilyTypes) {
            return $false
        }

        return $true
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.JsonObject.templateReference.templateId) {
            # I do not like this at all and it is a lazy but simple implementation...
            # It turns out that settingInstanceTemplateId and settingValueTemplateId are case sensitive
            # and there is ONE setting with a different casing in the Windows Baseline template.
            # The export saves it with lowercase which causes the import to fail.

            Write-Log "Get template $($PolicyObject.JsonObject.templateReference.templateId)"
            $templateObj = Invoke-MSGraphAPI -Url "/deviceManagement/configurationPolicyTemplates('$($PolicyObject.JsonObject.templateReference.templateId)')"
            if($templateObj.lifecycleState -and $templateObj.lifecycleState -ne "active") {
                Write-Log "Template '$($templateObj.displayName)' '$($templateObj.displayVersion)' is in '$($templateObj.lifecycleState)' state. Current state: $($templateObj.lifecycleState). Import might fail." 2
            }
            #Todo: Should probably check for the latest active version and use that instead of the one in the templateReference

            if(-not $script:baseLineTemplate) {
                $script:baseLineTemplate = @{}
            }
            if($script:baseLineTemplate.ContainsKey($PolicyObject.JsonObject.templateReference.templateId)) {
                $templateReference = $script:baseLineTemplate[$PolicyObject.JsonObject.templateReference.templateId]
            }
            else {
                Write-Log "Get template settings for '$($templateObj.displayName)' '$($templateObj.displayVersion)' ($($PolicyObject.JsonObject.templateReference.templateId))"
                $templateReference = Invoke-MSGraphAPI -Url "/deviceManagement/configurationPolicyTemplates('$($PolicyObject.JsonObject.templateReference.templateId)')/settingTemplates?`$expand=settingDefinitions&top=1000"
                $script:baseLineTemplate.Add($PolicyObject.JsonObject.templateReference.templateId, $templateReference)
            }

            if($templateReference) {
                $settingsJson = $PolicyObject.JsonObject.Settings | ConvertTo-Json -Depth 50
                $templateIDs, $dummy = Get-DependencyIDs ($templateReference | ConvertTo-Json -Depth 50)
                $objectIDs, $dummy = Get-DependencyIDs $settingsJson
                $diff = Compare-Object @($templateIDs) @($objectIDs) -CaseSensitive
                $updated = $false
                foreach($diffItem in ($diff | Where-Object SideIndicator -eq "=>")) {
                    $templateID = $templateIDs | Where-Object { $_ -eq $diffItem.InputObject }
                    if($templateID) {
                        # Found but with different casing
                        $settingsJson = $settingsJson -replace $diffItem.InputObject, $templateID
                        $updated = $true
                    }
                }
                if($updated) {
                    $PolicyObject.JsonObject.Settings = @($settingsJson | ConvertFrom-Json -Depth 50)
                }
            }
        }
        return $null
    }
}

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class SettingsCatalogObject : IntunePolicyBase
{
    Hidden [String]$_TemplateFamilyName = $null

    SettingsCatalogObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    SettingsCatalogObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $settingCatalogClasses = Get-IntuneSettingsCatalogClasses
        $policyType = $settingCatalogClasses | Where-Object { $_._FamilyTypes -contains $this.JsonObject.templateReference.templateFamily }
        if($policyType) {
            $this._PolicyType = $policyType
        }

        $this._PolicyName = ?? $this.JsonObject.templateReference.templateDisplayName $this._PolicyType.PolicyBaseType

        # Opt into the sub-resource batching contract so Invoke-PolicyHydrate
        # fans out the assignments fetch via $batch.
        $this._HasSubResourceBatch = $true

        Add-ObjectProperty $this "TemplateVersion" { $this.JsonObject.templateReference.templateDisplayVersion  }
        if($this.JsonObject.templateReference.templateFamily) {
            $this._TemplateFamilyName = (?? (Get-EndpointSecurityCategoryName $this.JsonObject.templateReference.templateFamily) $this.JsonObject.templateReference.templateFamily)
        }
        else {
            $this._TemplateFamilyName = $null
        }
        Add-ObjectProperty $this "TemplateFamily" { $this._TemplateFamilyName } 
        Add-ObjectProperty $this "Category" { $this._TemplateFamilyName  }
    }

    #Hidden [String] GetName()
    #{
    #    return $this.JsonObject.Name
    #}

    # Sub-resource contract. The per-id body GET ($expand=assignments,settings)
    # returns an empty `assignments` array — known Graph quirk on
    # configurationPolicies. We hit the dedicated /assignments endpoint via
    # $batch instead so Invoke-PolicyHydrate fans them out in parallel.
    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        return @([PSCustomObject]@{
            Key = 'assignments'
            Url = "$($this._PolicyType.API)/$($this.Id)/assignments"
        })
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -eq 1 -and $Key -eq 'assignments') {
            # Comma-prefix forces an array reference through the if-expression —
            # without it, PowerShell unwraps a single-element @() on assignment
            # and ConvertTo-Json then emits a bare object instead of [{...}].
            # Lowercase `assignments` matches the Graph wire shape and the
            # `assignments@odata.*` metadata properties the body fetch leaves
            # alongside. PSObject is case-insensitive on read, so existing
            # callers reading `Assignments` keep working.
            $assignments = if($Body -and $Body.value) { ,@($Body.value) } else { ,@() }
            Add-Member -InputObject $this.JsonObject -MemberType NoteProperty -Name 'assignments' -Value $assignments -Force
        }
        return @()
    }
}

#endregion

#########################################################################################
#
# Reusable Settings
#
#########################################################################################

# region Reusable Settings

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class ReusableSettingsTypeBase : IntunePolicyTypeBase
{
    static [bool] $IsAbstract = $true

    ReusableSettingsTypeBase() : Base()
    {
        ([ReusableSettingsTypeBase]$this).Init()
    }

    Init()
    {
        $this._PolicyName = "Reusable Settings"
        $this._ID = "ReusableSettingsBase"
        $this._API = "deviceManagement/reusablePolicySettings"
        $this._PolicyBaseName = "Reusable Settings"
        $this._PropertiesToRemove = @('Settings','@OData.Type')
        $this._Permissions=@("DeviceManagementConfiguration.ReadWrite.All")
        $this._ImportOrder = 70
        $this._ExpandAssignmentsList = $false
        $this._SkipRemoveProperties = @("@OData.Type")
        $this._ObjectClass = "ReusableSettingObject"
        $this._SupportsAssignments = $false

        $this._PolicyTypeOrder = 210
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        return $false
    }
}

class ReusableSettingsObjectBase : IntunePolicyBase
{
    ReusableSettingsObjectBase([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.InitReusable() }

    ReusableSettingsObjectBase() : Base() { $this.InitReusable() }

    Hidden InitReusable()
    {
        # The list endpoint omits settingInstance; hydration fetches it via the
        # sub-resource contract (single GET coalesced into the hydrate $batch).
        # Owns the API in one place — was previously duplicated in
        # Sync-BulkExportReusableSettings (Internal/PolicyHydrateExtras.ps1).
        $this._HasSubResourceBatch = $true
    }

    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        if($this.JsonObject.settingInstance) { return @() }   # already present
        return @([PSCustomObject]@{
            Key     = 'reusableSettingInstance'
            Url     = "$($this._PolicyType.API)/$($this.Id)?`$select=settinginstance,displayname,description"
            Headers = @{ Accept = 'application/json;odata.metadata=none' }
        })
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -eq 1 -and $Key -eq 'reusableSettingInstance' -and $Body -and $Body.settingInstance) {
            Add-Member -InputObject $this.JsonObject -MemberType NoteProperty -Name 'settingInstance' -Value $Body.settingInstance -Force
        }
        return @()
    }
}



#endregion
