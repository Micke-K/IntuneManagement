#ImportOrder 220

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class AppProtectionGroup : IntunePolicyGroupBase
{
    AppProtectionGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "AppProtection"
        $this._Name = "App protection policies"
        $this._Icon = "AppProtection"
    }
}

#########################################################################################
#
# App Protection
#
#########################################################################################

# region App Protection
class AppProtectionType : IntunePolicyTypeBase
{
    AppProtectionType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "AppProtectionGroup")
        $this._PolicyName = "App protection policy"
        $this._ID = "AppProtection"
        $this._SubTypeColumn = "ManagementType=Management type"
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceAppManagement/managedAppPolicies"
        $this._Permissions = @("DeviceManagementApps.ReadWrite.All")
        $this._Dependencies = @("Applications")
        $this._ObjectClass = "AppProtectionPolicyObject"
        $this._ExpandAssignmentsList = $false
        # Assignments are fetched per platform collection - see
        # GetAssignmentsBaseURL below. _AssignmentsViaExpand stays $false:
        # once the URL targets a concrete subtype the plain navigation GET
        # works, and it returns just the assignments instead of the whole policy.
        $this._PropertiesToRemove = @('exemptAppLockerFiles')
        $this._PropertiesToRemoveForUpdate = @("protectedAppLockerFiles","version") # ToDo: !!! Add support for protectedAppLockerFiles?
        $this._VerifyObject = $true
        # CheckPolicy is a complete @odata.type matcher (the managedAppPolicies allowlist),
        # so a rejection is authoritative - do not let the folder-trust fallback rescue a
        # foreign object misplaced in this type's export folder.
        $this._StrictODataTypeCheck = $true
        $this._Icon = "AppConfiguration"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    # managedAppPolicies is Collection(managedAppPolicy), and managedAppPolicy
    # declares no navigation properties - so both {API}/{id}/assignments and
    # {API}/{id}?$expand=assignments return 400 ("Could not find a property named
    # 'assignments' on type 'microsoft.graph.managedAppPolicy'"). Every concrete
    # subtype inherits `assignments`, so route through the per-platform collection
    # (_objectClass, set in the object's constructor via Get-AppConfigurationClass).
    # defaultManagedAppProtection is the exception - it has no assignments at all.
    [String]GetAssignmentsBaseURL([PSCustomObject]$PolicyObject)
    {
        if(-not $PolicyObject._objectClass) { return $this._API }
        if($PolicyObject._objectClass -eq "defaultManagedAppProtections") { return $null }

        return "deviceAppManagement/$($PolicyObject._objectClass)"
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        # apps is a navigation property set later via targetApps (PostImportCommand),
        # not an inline body property - strip it from the POST body.
        Remove-Property $PolicyObject.JsonObject "apps"
        Remove-Property $PolicyObject.JsonObject "apps@odata.context"

        # The polymorphic managedAppPolicies collection (used for listing) rejects
        # POST, so import must target the per-platform collection. _objectClass is
        # the metadata-derived endpoint segment (e.g. iosManagedAppProtections),
        # set on the object at construction via Get-AppConfigurationClass.
        if($PolicyObject._objectClass)
        {
            return @{"API"="deviceAppManagement/$($PolicyObject._objectClass)"}
        }

        return (@{})
    }

    PostImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        if($SourceObject.Object.Apps) {
            # No "@odata.type" on the created object so reload new object
            #$newObject = (Invoke-MSGraphAPI "$($objectType.API)?`$filter=id eq '$($obj.Id)'").Value
            $newObject = Invoke-MSGraphAPI -Url "$($this.API)/$($PolicyObject.Id)" -TokenId $PolicyObject._TokenID
            if($newObject)
            {
                try
                {
                    $apps = [PSCustomObject]@{ 
                        appGroupType = $PolicyObject.Object.appGroupType
                        apps = @($SourceObject.Object.Apps)                 
                    } 
                    $json = $apps | ConvertTo-Json -Depth 20

                    # Created object carries no @odata.type; use the source object's
                    # metadata-derived endpoint segment (_objectClass).
                    if($SourceObject._objectClass)
                    {
                        Invoke-MSGraphAPI -Url "deviceAppManagement/$($SourceObject._objectClass)/$($PolicyObject.Id)/targetApps" -Content $json -HttpMethod POST -TokenId $PolicyObject._TokenID | Out-Null
                    }
                }
                catch {
                    # The policy was created; only the targetApps association failed. Keep
                    # going but make the partial import visible instead of swallowing it.
                    Write-LogError "Failed to assign target apps to imported App protection policy '$($PolicyObject.displayName)' ($($PolicyObject.Id)) via deviceAppManagement/$($SourceObject._objectClass)/$($PolicyObject.Id)/targetApps" $_.Exception
                }
            }
        }
    }

    [Hashtable]PreImportAssignmentsCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        # /assign is bound to the concrete platform subtype; the polymorphic
        # managedAppPolicies collection returns 400 for the assign action.
        if($SourceObject._objectClass)
        {
            return @{"API"="deviceAppManagement/$($SourceObject._objectClass)/$($PolicyObject.Id)/assign"}
        }
        return $null
    }

    [Hashtable]PreUpdateCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$ExistingObject)
    {
        # managedAppPolicies rejects PATCH, and apps is a navigation property set
        # via targetApps rather than inline. Re-post apps to the per-platform
        # endpoint, strip them from the PATCH body, then PATCH that endpoint.
        #
        # Routing note: an imported object's POST response carries no
        # @odata.type, so ITS _objectClass can be null - the update object came
        # from a file that always has the type, so prefer that one.
        if(-not $ExistingObject._objectClass -and $PolicyObject._objectClass)
        {
            $ExistingObject._objectClass = $PolicyObject._objectClass
        }

        if($PolicyObject.JsonObject.apps -and $ExistingObject._objectClass)
        {
            try
            {
                $apps = [PSCustomObject]@{
                    appGroupType = $PolicyObject.JsonObject.appGroupType
                    apps         = @($PolicyObject.JsonObject.apps)
                }
                $json = $apps | ConvertTo-Json -Depth 20
                Invoke-MSGraphAPI -Url "deviceAppManagement/$($ExistingObject._objectClass)/$($ExistingObject.Id)/targetApps" -Content $json -HttpMethod POST -TokenId $ExistingObject._TokenID | Out-Null
            }
            catch {
                # The policy PATCH still proceeds; surface the failed targetApps update.
                Write-LogError "Failed to update target apps for App protection policy '$($ExistingObject.displayName)' ($($ExistingObject.Id)) via deviceAppManagement/$($ExistingObject._objectClass)/$($ExistingObject.Id)/targetApps" $_.Exception
            }
        }
        Remove-Property $PolicyObject.JsonObject "apps"
        Remove-Property $PolicyObject.JsonObject "apps@odata.context"
        # assignments is a navigation property (managed via the /assign action),
        # not inline-PATCHable - PATCHing it 400s on the platform entity type.
        Remove-Property $PolicyObject.JsonObject "assignments"
        Remove-Property $PolicyObject.JsonObject "assignments@odata.context"

        if($ExistingObject._objectClass)
        {
            return @{"API"="deviceAppManagement/$($ExistingObject._objectClass)/$($ExistingObject.Id)"}
        }
        return $null
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        # App Protection owns the polymorphic deviceAppManagement/managedAppPolicies
        # collection: the per-platform *ManagedAppProtection variants (ios / android /
        # windows / default) plus the two Windows Information Protection policy types.
        # An explicit allowlist is required: file objects carry only @odata.type (no
        # top-level @odata.id), so this runs as the file->type discriminator. A previous
        # "accept everything except targetedManagedAppConfiguration" greedily claimed
        # unrelated policy types (Compliance, CA, etc.) when resolving from an export
        # folder. targetedManagedAppConfiguration is App Config, not App Protection
        # (see AppConfigurationManagedAppType.CheckPolicy).
        $odata = [string]$PolicyObject.'@odata.type'
        if($odata -match 'ManagedAppProtection$' -or
           $odata -eq '#microsoft.graph.mdmWindowsInformationProtectionPolicy' -or
           $odata -eq '#microsoft.graph.windowsInformationProtectionPolicy')
        {
            return $true
        }

        return $false
    }
}

Class AppProtectionPolicyObject : IntunePolicyBase
{
    Hidden [string]$_objectClass = $null
    Hidden [string]$_managemntType = $null

    AppProtectionPolicyObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AppProtectionPolicyObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AppProtectionType")

        if($this.Object."@odata.type" -eq "#microsoft.graph.mdmWindowsInformationProtectionPolicy") {
            $this._managemntType = "With enrollment"
        }
        elseif($this.Object."@odata.type" -eq "#microsoft.graph.windowsInformationProtectionPolicy") {
            $this._managemntType = "Without enrollment"
        }
        else {
            $this._managemntType = "All app types"
        }

        Add-ObjectProperty $this "ManagementType" { $this._managemntType }

        Get-AppConfigurationClass $this

        if($this.JsonObject."@odata.type" -eq "#microsoft.graph.iosManagedAppProtection") {
            $platformName = Get-LanguageString "AppProtection.iOSPlatformLabel"
            if($platformName) {
                #$this._PlatformName = $platformName
            }
        }
    }
}
