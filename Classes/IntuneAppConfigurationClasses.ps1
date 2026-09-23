#ImportOrder 220

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class AppConfigurationGroup : IntunePolicyGroupBase
{
    AppConfigurationGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "AppConfiguration"
        $this._Name = "App configuration policies"
        $this._Icon = "AppConfiguration"
        $this._ExtraColumns = @("EnrolmentType=Enrolment type")
    }
}

#########################################################################################
#
# App Configuration (App)
#
#########################################################################################

# region App Configuration (App)
class AppConfigurationManagedAppType : IntunePolicyTypeBase
{
    AppConfigurationManagedAppType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "AppConfigurationGroup")
        $this._PolicyName = "App configuration (App)"
        $this._ID = "AppConfigurationManagedApp"
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceAppManagement/targetedManagedAppConfigurations"
        $this._Permissions = @("DeviceManagementApps.ReadWrite.All")
        $this._Dependencies = @("Applications")
        $this._ObjectClass = "AppConfigurationManagedAppObject"
        $this._ExpandAssignmentsList = $false
        # CheckPolicy authoritatively matches this type by @odata.type (plus the base
        # @odata.id fallback), so a rejection is real - skip the folder-trust fallback.
        $this._StrictODataTypeCheck = $true
        $this._Icon = "AppConfiguration"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        # targetedManagedAppConfiguration is surfaced by the polymorphic
        # managedAppPolicies collection alongside App Protection variants, so match it
        # explicitly by @odata.type. This also lets a file object (which carries only
        # @odata.type, no top-level @odata.id) resolve to this type; the base
        # @odata.id-based CheckPolicy would reject it.
        if($PolicyObject.'@odata.type' -eq '#microsoft.graph.targetedManagedAppConfiguration')
        {
            return $true
        }

        # Fall back to the base @odata.id matcher (deviceAppManagement/targetedManagedAppConfigurations)
        # for objects that carry an id but no top-level @odata.type.
        return ([IntunePolicyTypeBase]$this).CheckPolicy($PolicyObject)
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        # apps is a navigation property set after create via targetApps
        # (PostImportCommand); strip it, then POST to the type API
        # (deviceAppManagement/targetedManagedAppConfigurations).
        Remove-Property $PolicyObject.JsonObject "apps"
        Remove-Property $PolicyObject.JsonObject "apps@odata.context"

        return $null
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

                    Invoke-MSGraphAPI -Url "$($this.API)/$($PolicyObject.Id)/targetApps" -Content $json -HttpMethod POST -TokenId $PolicyObject._TokenID | Out-Null
                }
                catch {
                    # The policy was created; only the targetApps association failed. Keep
                    # going but make the partial import visible instead of swallowing it.
                    Write-LogError "Failed to assign target apps to imported App configuration policy '$($PolicyObject.displayName)' ($($PolicyObject.Id)) via $($this.API)/$($PolicyObject.Id)/targetApps" $_.Exception
                }
            }
        }
    }

    [Hashtable]PreUpdateCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$ExistingObject)
    {
        # apps is a navigation property set via targetApps, not inline. Re-post
        # apps to the type API, strip them from the PATCH body, then PATCH the
        # type API (deviceAppManagement/targetedManagedAppConfigurations).
        if($PolicyObject.JsonObject.apps)
        {
            try
            {
                $apps = [PSCustomObject]@{
                    appGroupType = $PolicyObject.JsonObject.appGroupType
                    apps         = @($PolicyObject.JsonObject.apps)
                }
                $json = $apps | ConvertTo-Json -Depth 20
                Invoke-MSGraphAPI -Url "$($this.API)/$($ExistingObject.Id)/targetApps" -Content $json -HttpMethod POST -TokenId $ExistingObject._TokenID | Out-Null
            }
            catch {
                # The policy PATCH still proceeds; surface the failed targetApps update.
                Write-LogError "Failed to update target apps for App configuration policy '$($ExistingObject.displayName)' ($($ExistingObject.Id)) via $($this.API)/$($ExistingObject.Id)/targetApps" $_.Exception
            }
        }
        Remove-Property $PolicyObject.JsonObject "apps"
        Remove-Property $PolicyObject.JsonObject "apps@odata.context"
        # assignments is a navigation property (managed via the /assign action),
        # not inline-PATCHable.
        Remove-Property $PolicyObject.JsonObject "assignments"
        Remove-Property $PolicyObject.JsonObject "assignments@odata.context"

        return $null
    }
}

Class AppConfigurationManagedAppObject : IntunePolicyBase
{
    Hidden [string]$_objectClass = $null

    AppConfigurationManagedAppObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AppConfigurationManagedAppObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AppConfigurationManagedAppType")

        Add-ObjectProperty $this "EnrolmentType" { "Managed apps" }

        Get-AppConfigurationClass $this

        # Targeted-app resolution runs through the sub-resource contract; see the
        # Get-/Add-/Build-AppConfigTargetApp* helpers below.
        $this._HasSubResourceBatch = $true
    }

    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        return (Get-AppConfigTargetAppRequests $this)
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -eq 1) { Add-AppConfigTargetAppResult $Key $Body }
        return @()
    }

    [void] FinalizeSubResources()
    {
        Build-AppConfigTargetAppRefs $this
    }
}

#########################################################################################
#
# App Configuration (Device)
#
#########################################################################################

# region App Configuration (Device)
class AppConfigurationManagedDeviceType : IntunePolicyTypeBase
{
    AppConfigurationManagedDeviceType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "AppConfigurationGroup")
        $this._PolicyName = "App configuration (Device)"
        $this._ID = "AppConfigurationManagedDevice"
        $this._API = "deviceAppManagement/mobileAppConfigurations"
        $this._QueryList = "?`$filter=microsoft.graph.androidManagedStoreAppConfiguration/appSupportsOemConfig%20eq%20false%20or%20isof(%27microsoft.graph.androidManagedStoreAppConfiguration%27)%20eq%20false"
        $this._Permissions = @("DeviceManagementApps.ReadWrite.All")
        $this._Dependencies = @("Applications")
        $this._ObjectClass = "AppConfigurationManagedDeviceObject"
        $this._ExpandAssignmentsList = $false
        $this._Icon = "AppConfiguration"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportAssignmentsCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        return (@{"API"="deviceAppManagement/mobileAppConfigurations/$($PolicyObject.Id)/microsoft.graph.managedDeviceMobileAppConfiguration/assign"})
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        Import-AppConfigurationTargetedApps $PolicyObject
        return $null
    }
}

Class AppConfigurationManagedDeviceObject : IntunePolicyBase
{
    AppConfigurationManagedDeviceObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AppConfigurationManagedDeviceObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AppConfigurationManagedDeviceType")

        Add-ObjectProperty $this "EnrolmentType" { "Managed devices" }

        # Targeted-app resolution runs through the sub-resource contract; see the
        # Get-/Add-/Build-AppConfigTargetApp* helpers.
        $this._HasSubResourceBatch = $true
    }

    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        return (Get-AppConfigTargetAppRequests $this)
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -eq 1) { Add-AppConfigTargetAppResult $Key $Body }
        return @()
    }

    [void] FinalizeSubResources()
    {
        Build-AppConfigTargetAppRefs $this
    }
}

#########################################################################################
#
# Generic Functions
#
#########################################################################################

function Get-AppConfigurationClass
{
    param($Policy)

    try {
        $tmp = $Policy.Object."@odata.type".Split('.')[-1]
        $Policy._objectClass = Get-GraphObjectClassName $tmp
    }
    catch { }

    if($null -eq $Policy._objectClass) {
        Write-Log "Could not get class name for $($Policy.Name) ($($Policy.Object."@odata.type"))" 3
    }    
}

function Get-AppConfigurationFullObject
{
    param($Policy)

    if(-not $Policy.Object."@odata.type" -or -not $Policy._objectClass) { return $false }

    $expand = $null
    if($Policy._objectClass -eq "windowsInformationProtectionPolicies")
    {
        $expand = "?`$expand=protectedAppLockerFiles,exemptAppLockerFiles"
    }
    else {
        $url = $Policy.GetObjectURL()

        $tmpArr = $url.Split("?")
        if($tmpArr.Length -gt 1) {
            $expand = "?" + $tmpArr[1]
        }
    }

    $fullObject = (Invoke-MSGraphAPI -Url "deviceAppManagement/$($Policy._objectClass)/$($Policy.Id)$expand" -TokenId $Policy._TokenId)
    if($fullObject)
    {
        $Policy.JsonObject = $fullObject
        $Policy._IsFullObject = $true

        return $true
    }
    return $false
}

# Targeted-app resolution for AppConfiguration policies.
#
# The policy body lists app IDs in targetedMobileApps; cross-tenant export needs
# each app's displayName + @odata.type to re-map them on import into another
# tenant (#CustomRefTargetedApps). The mobileApps/<id> lookup is owned ONLY by
# Get-AppConfigTargetAppRequests below — the sub-resource contract on the
# AppConfiguration*Object classes drives the fetch (coalesced across policies by
# Invoke-PolicySubresourceFetch's URL de-dup), caches results, then builds the
# ref string. Previously this was duplicated in Sync-BulkExportAppConfigurationTargetApps
# (Internal/PolicyHydrateExtras.ps1).

# Process-wide cache: appId -> app body (or $null for a 404/miss). Shared across
# every AppConfig policy in a hydrate run so the same app is fetched once.
function Get-AppConfigTargetAppCache
{
    if($null -eq $script:_appConfigTargetAppCache) { $script:_appConfigTargetAppCache = @{} }
    return $script:_appConfigTargetAppCache
}

# Only these polymorphic AppConfig @odata.types carry targetedMobileApps that
# need tenant-specific remapping. (androidManagedAppProtection is listed for
# parity with the legacy filter; App Protection policies use `apps`, not
# `targetedMobileApps`, so they never actually match.)
function Test-AppConfigHasTargetApps
{
    param($Policy)
    return ($Policy.JsonObject.'@OData.Type' -in @(
            '#microsoft.graph.androidManagedAppProtection',
            '#microsoft.graph.androidForWorkMobileAppConfiguration',
            '#microsoft.graph.androidManagedStoreAppConfiguration',
            '#microsoft.graph.iosMobileAppConfiguration'
        ) -and @($Policy.JsonObject.targetedMobileApps).Count -gt 0)
}

# Sub-resource requests for every targeted app not already cached. The
# deviceAppManagement/mobileApps/<id> URL lives ONLY here.
function Get-AppConfigTargetAppRequests
{
    param($Policy)
    if(-not (Test-AppConfigHasTargetApps $Policy)) { return @() }
    $cache = Get-AppConfigTargetAppCache
    $reqs = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach($appId in @($Policy.JsonObject.targetedMobileApps)) {
        if(-not $appId -or $cache.ContainsKey($appId)) { continue }
        [void]$reqs.Add([PSCustomObject]@{
            Key     = "targetApp_$appId"
            Url     = "deviceAppManagement/mobileApps/$appId"
            Headers = @{ Accept = 'application/json;odata.metadata=minimal' }
        })
    }
    return $reqs.ToArray()
}

# Cache one targeted-app sub-resource response. Stores $null for misses so 404s
# aren't re-requested by a later policy in the same run.
function Add-AppConfigTargetAppResult
{
    param([string]$Key, $Body)
    if($Key -notlike 'targetApp_*') { return }
    $appId = $Key.Substring('targetApp_'.Length)
    (Get-AppConfigTargetAppCache)[$appId] = $Body
}

# Build #CustomRefTargetedApps from the cached app bodies (finalize step).
function Build-AppConfigTargetAppRefs
{
    param($Policy)
    if(-not (Test-AppConfigHasTargetApps $Policy)) { return }
    $cache = Get-AppConfigTargetAppCache
    $targetedApps = @()
    foreach($appId in @($Policy.JsonObject.targetedMobileApps)) {
        $appObj = $cache[$appId]
        if($appObj) {
            Write-Log "Add target app info $($appObj.displayName) ($($appObj.Id)) of type $($appObj.'@OData.Type')"
            $targetedApps += $appObj.displayName + '|!|' + $appObj.Id + '|!|' + $appObj.'@OData.Type'
        }
        else {
            Write-Log "No app found with id $appId" 2
        }
    }
    if($targetedApps.Count -gt 0) {
        Add-Member -InputObject $Policy.JsonObject -MemberType NoteProperty -Name '#CustomRefTargetedApps' -Value ($targetedApps -join '|*|') -Force
    }
}

function Import-AppConfigurationTargetedApps
{
    param($Policy)

    if($Policy.JsonObject."#CustomRefTargetedApps" -and $Policy.JsonObject.targetedMobileApps)
    {
        Write-Log "Adding app targets for $($Policy.JsonObject.displayName)"

        $targetedAppsInfo = $Policy.JsonObject."#CustomRefTargetedApps"

        $translatedTargetedApps = @()

        if($targetedAppsInfo)
        {
            foreach($targetedApp in ($targetedAppsInfo -split "[|][*][|]"))
            {
                $appName, $appId, $appType = $targetedApp -split "[|][!][|]"
                if(-not $appName -or -not $appId)
                {
                    Write-Log "App Name and Id is missing in string: $appApp" 2
                    continue
                }
                $tmpApps = (Invoke-MSGraphAPI -Url "/deviceAppManagement/mobileApps?`$filter=displayName eq '$appName'" -TokenId $Policy._TokenId).value
                if(-not $tmpApps)
                {
                    Write-Log "No application found with name $appName. $appId will not be translated and added to target list" 2
                    continue
                }
                $tmpApp = $tmpApps | Where-Object '@OData.Type' -eq $appType
                if(-not $tmpApp)
                {
                    Write-Log "No $appName application found of type $appType. $appId will not be translated and added to target list" 2
                }
                elseif(($tmpApp | Measure-Object).Count -gt 1) {
                    Write-Log "$(($tmpApp | Measure-Object).Count) applications found with name '$appName' of type $appType. $appId will not be translated and added to target list" 2
                }
                else {
                    Write-Log "Found '$appName' with id $($tmpApp.Id) ($appType)"
                    $translatedTargetedApps += $tmpApp.Id
                }
            }

            if($translatedTargetedApps.Count -gt 0) {
                Write-Log "Updating translated targeted apps"
                $Policy.JsonObject.targetedMobileApps = $translatedTargetedApps
            }
            else {
                Write-Log "Could not find targeted apps in the evnironment. Verify that they are added. Policy import might fail" 3
            }
        }
    }    
}
