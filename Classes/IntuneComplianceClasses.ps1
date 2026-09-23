#ImportOrder 220

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class ComplianceGroup : IntunePolicyGroupBase
{
    ComplianceGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "Compliance"
        $this._Name = "Compliance"
        $this._Icon = "CompliancePolicies"
    }
}

#########################################################################################
#
# Device Compliance
#
#########################################################################################

# region Device Compliance
class DeviceComplianceType : IntunePolicyTypeBase
{
    DeviceComplianceType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ComplianceGroup")
        $this._PolicyName = "Compliance Policy"
        $this._ID = "CompliancePolicies"
        $this._API = "deviceManagement/deviceCompliancePolicies"
        # This endpoint answers HTTP 400 to startswith() on displayName
        # (verified 2026-08-27; 'displayName eq' works, prefix search does not),
        # so name searches filter client-side instead.
        $this._SupportsNameFilter = $false
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._Expand = "scheduledActionsForRule(`$expand=scheduledActionConfigurations)"
        # "Locations" (v3's deprecated Intune managementConditions type) is not a
        # PolicyType here, so listing it resolved to nothing on import.
        $this._Dependencies = @("Notifications","ComplianceScripts")
        $this._ObjectClass = "DeviceComplianceObject"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        foreach($scheduledActionsForRule in $PolicyObject.JsonObject.scheduledActionsForRule)
        {
            foreach($scheduledActionConfiguration in $scheduledActionsForRule.scheduledActionConfigurations)
            {
                foreach($notificationMessageCCGroup in $scheduledActionConfiguration.notificationMessageCCList)
                {
                    Add-GraphMigrationObject $notificationMessageCCGroup "groups" "Group" ([IO.Path]::GetDirectoryName($PathToFile)) $PolicyObject._TokenId
                }
            }
        }
    }

    [Hashtable]PreUpdateCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        $api = "deviceManagement/deviceCompliancePolicies/$($PolicyObject.Id)/scheduleActionsForRules"

        $tmpObj = [PSCustomObject]@{
            deviceComplianceScheduledActionForRules = $PolicyObject.JsonObject.scheduledActionsForRule
        }

        $json = ConvertTo-Json $tmpObj -Depth 20
        Invoke-MSGraphAPI -Url $api -Content $json -HttpMethod "POST" -TokenId $PolicyObject._TokenId | Out-Null
    
        Remove-Property $PolicyObject.JsonObject "scheduledActionsForRule"

        return (@{})
    }
}

Class DeviceComplianceObject : IntunePolicyBase
{
    DeviceComplianceObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    DeviceComplianceObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "DeviceComplianceType")
    }
}

#########################################################################################
#
# Device Compliance V2
#
#########################################################################################

Class DeviceComplianceV2Type : IntunePolicyTypeBase
{
    DeviceComplianceV2Type() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ComplianceGroup")
        # Its own name so a mixed Compliance list tells V1 from V2 without a
        # separate "Policy base" column.
        $this._PolicyName = "Compliance Policy (Settings Catalog)"
        $this._PolicyBaseName = "Compliance Policy V2"
        $this._APITitle = "Compliance Policy (Linux)"
        $this._ID = "CompliancePoliciesV2"
        $this._API = "deviceManagement/compliancePolicies"
        $this._PropertiesToRemove = @('settingCount')
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._NameProperty = "Name"
        $this._Expand = "settings"        
        $this._ObjectClass = "DeviceComplianceV2Object"
        $this._Icon = "CompliancePolicies"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreUpdateCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        # V2 compliance runs on the settings-catalog engine: updates must PUT
        # the full body (including settings); PATCH with settings is rejected.
        return @{ "Method" = "PUT" }
    }
}

Class DeviceComplianceV2Object : IntunePolicyBase
{
    DeviceComplianceV2Object([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    DeviceComplianceV2Object() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "DeviceComplianceV2Type")
    }
}

#########################################################################################
#
# Compliance Scripts
#
#########################################################################################

Class ComplianceScriptsType : IntunePolicyTypeBase
{
    ComplianceScriptsType() : Base()
    {
        ([ComplianceScriptsType]$this).Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ComplianceGroup")
        $this._APITitle = "Compliance Scripts"
        $this._PolicyName = "Compliance Script"
        $this._ID = "ComplianceScripts"
        $this._API = "deviceManagement/deviceComplianceScripts"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._ObjectClass = "ComplianceScriptObject"
        $this._Icon = "Scripts"
        # Graph rejects PATCHing the read-only version back ("Invalid property
        # name: Version").
        $this._PropertiesToRemoveForUpdate = @('version')
        # Custom compliance scripts are REFERENCED by compliance policies, not
        # assigned to devices - the portal has no Assignments blade and Graph's
        # /assign action rejects the request (400 "Action parameters do not
        # contain parameter 'deviceHealthScriptAssignments'").
        $this._SupportsAssignments = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class ComplianceScriptObject : IntunePolicyBase
{
    ComplianceScriptObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    ComplianceScriptObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PlatformName = Get-LanguageString "Platform.windows10AndLater"
        $this._PolicyType = (Get-SingletonObject "ComplianceScriptsType")
    }
}

#########################################################################################
#
# Compliance Scripts - Linux
#
#########################################################################################

Class ComplianceScriptsLinuxType : ReusableSettingsTypeBase
{
    ComplianceScriptsLinuxType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ComplianceGroup")
        $this._APITitle = "Compliance Scripts (Linux)"
        $this._PolicyName = "Compliance Script"
        $this._ID = "ComplianceScriptsScriptsLinux"
        $this._QueryList = "?`$filter=settingDefinitionId eq 'linux_customcompliance_discoveryscript_reusablesetting'"
        # Reusable settings carry no roleScopeTagIds in the Graph schema.
        $this._ScopeTagProperty = ""
        $this._ObjectClass = "ComplianceScriptLinuxObject"
        $this._Icon = "Scripts"
        $this._Folder = "ReusableSettings"
        $this._PolicyTypeOrder = 140
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.'@odata.type' -ne "#microsoft.graph.deviceManagementReusablePolicySetting") { return $false }

        if($PolicyObject.settingDefinitionId -eq 'linux_customcompliance_discoveryscript_reusablesetting') {
            return $true
        }
        
        return $false
    }
}

Class ComplianceScriptLinuxObject : ReusableSettingsObjectBase
{
    ComplianceScriptLinuxObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    ComplianceScriptLinuxObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "ComplianceScriptsLinuxType")
        $this._PlatformName = Get-LanguageString "Platform.linux" -IgnoreMissing
    }
}

#########################################################################################
#
# Compliance Notifications
#
#########################################################################################

Class NotificationsType : IntunePolicyTypeBase
{
    NotificationsType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ComplianceGroup")
        $this._PolicyName = "Notifications"
        $this._ID = "Notifications"
        $this._HasPlatform = $false
        $this._API = "deviceManagement/notificationMessageTemplates"
        #$this._QueryList = "?`$filter=displayName ne 'EnrollmentNotificationInternalMEO'"
        $this._Permissions = @("DeviceManagementServiceConfig.ReadWrite.All")
        $this._ObjectClass = "NotificationObject"
        $this._ImportOrder = 40
        $this._Expand = "localizedNotificationMessages"
        $this._ExpandAssignmentsList = $false
        # notificationMessageTemplate has no `assignments` navigation property at
        # all (its only nav is localizedNotificationMessages), so both the nav GET
        # and ?$expand=assignments return 400. Notification templates are targeted
        # from compliance policies, not assigned - don't ask for assignments.
        $this._SupportsAssignments = $false
        # notificationMessageTemplate has no description property in Graph.
        $this._HasDescription = $false
        # localizedNotificationMessages is a navigation property - PATCHing it
        # inline is rejected; messages are managed on their own sub-endpoint.
        $this._PropertiesToRemoveForUpdate = @('localizedNotificationMessages','defaultLocale')

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        Remove-Property $PolicyObject.JsonObject "defaultLocale"
        Remove-Property $PolicyObject.JsonObject "localizedNotificationMessages"
        Remove-Property $PolicyObject.JsonObject "localizedNotificationMessages@odata.context"

        return $null
    }

    PostImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        $this.UpdateNotificationMessages($PolicyObject, $SourceObject.JsonObject.localizedNotificationMessages)
    }

    Hidden UpdateNotificationMessages($PolicyObject, $localizedNotificationMessages)
    {
        if(-not $localizedNotificationMessages) {
            return
        }

        $updated = $false
        foreach($localizedNotificationMessage in $localizedNotificationMessages)
        {            
            Remove-GraphPropertiesForImport $this $localizedNotificationMessage
            $response = Invoke-MSGraphAPI -Url "$($this.API)/$($PolicyObject.Id)/localizedNotificationMessages" -Body ($localizedNotificationMessage | ConvertTo-Json -Depth 20) -Method "POST" -FullResponseObject
            if($response.Success) {
                Write-log "Notification message '$($localizedNotificationMessage.subject)' ($($localizedNotificationMessage.locale)) added successfully"
                $updated = $true
            }
            else {
                Write-log "Failed to add notification message: '$($localizedNotificationMessage.subject)' ($($localizedNotificationMessage.locale))" 3
            }
        }
        if($updated) {
            [void]$PolicyObject.Get()
        }
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.'displayName' -eq "EnrollmentNotificationInternalMEO") { return $false } # Skip built in

        return (([IntunePolicyTypeBase]$this).CheckPolicy($PolicyObject))
    }
}

Class NotificationObject : IntunePolicyBase
{
    NotificationObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    NotificationObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "NotificationsType")
    }
}

