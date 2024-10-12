<#

A module that handles custom documentation tasks

This will add properties at runtime that is required for the documentation

This module will also document some objects based on PowerShell functions

#>

function Get-ModuleVersion
{
    '1.6.6'
}

function Invoke-InitializeModule
{
    Add-DocumentationProvicer ([PSCustomObject]@{
        Name="Custom"
        Priority = 1000 # The priority of the Provider. Lower number has higher priority.
        InitializeDocumentation = { Initialize-CDDocumentation @args } 
        DocumentObject = { Invoke-CDDocumentObject @args }
        GetCustomProfileValue = { Add-CDDocumentCustomProfileValue @args }
        GetCustomChildObject = { Get-CDDocumentCustomChildObject  @args }
        GetCustomPropertyObject = { Get-CDDocumentCustomPropertyObject  @args }
        AddCustomProfileProperty = { Add-CDDocumentCustomProfileProperty @args }
        PostAddValue = { Invoke-CDDocumentCustomPostAdd @args }
        ObjectDocumented = { Invoke-CDDocumentCustomObjectDocumented @args }
        TranslateSectionFile = { Invoke-CDDocumentTranslateSectionFile @args }
        PostSettingsCatalog = { Invoke-CDDocumentPostSettingsCatalog @args }
    })
}

function Initialize-CDDocumentation
{
    $script:allTenantApps = $null
    $script:allTermsOfUse = $null 
    $script:allAuthenticationStrength = $null
    $script:allAuthenticationContextClasses = $null 
    $script:allCustomCompliancePolicies = $null 
}

function Invoke-CDDocumentObject
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType
    $type = $obj.'@OData.Type'

    if($type -eq '#microsoft.graph.conditionalAccessPolicy')
    {
        Invoke-CDDocumentConditionalAccess $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory") #,"RawValue","Description"
        }
    }
    elseif($type -eq '#microsoft.graph.agreement')
    {
        Invoke-CDDocumentTermsOfUse $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value") #,"RawValue","Description"
        }
    }    
    elseif($type -eq '#microsoft.graph.countryNamedLocation')
    {
        Invoke-CDDocumentCountryNamedLocation $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value") 
        }
    }
    elseif($type -eq '#microsoft.graph.ipNamedLocation')
    {
        Invoke-CDDocumentIPNamedLocation $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value") 
        }
    }
    elseif($type -eq '#microsoft.graph.androidForWorkMobileAppConfiguration' -or
            $type -eq '#microsoft.graph.androidManagedStoreAppConfiguration') {

        Invoke-CDDocumentAndroidManagedStoreAppConfiguration $documentationObj
        
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }
    elseif($type -eq '#microsoft.graph.iosMobileAppConfiguration')
    {
        Invoke-CDDocumentMobileAppConfiguration $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }
    elseif($type -eq '#microsoft.graph.targetedManagedAppConfiguration')
    {
        Invoke-CDDocumentManagedAppConfig $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }
    elseif($type -eq '#microsoft.graph.policySet')
    {
        Invoke-CDDocumentPolicySet $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }
    elseif($type -eq '#microsoft.graph.windows10CustomConfiguration' -or 
        $type -eq '#microsoft.graph.androidForWorkCustomConfiguration' -or
        $type -eq '#microsoft.graph.androidWorkProfileCustomConfiguration' -or
        $type -eq '#microsoft.graph.androidCustomConfiguration')
    {
        Invoke-CDDocumentCustomOMAUri $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }
    elseif($type -eq '#microsoft.graph.notificationMessageTemplate')
    {
        Invoke-CDDocumentNotification $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }
    elseif($type -eq '#microsoft.graph.deviceAndAppManagementAssignmentFilter')
    {
        Invoke-CDDocumentAssignmentFilter $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category")
        }
    }    
    elseif($type -eq '#microsoft.graph.deviceComanagementAuthorityConfiguration')
    {
        Invoke-CDDocumentCoManagementSettings $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }
    elseif($type -eq '#microsoft.graph.windowsKioskConfiguration')
    {
        Invoke-CDDocumentWindowsKioskConfiguration $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }    
    elseif($type -eq '#microsoft.graph.deviceEnrollmentPlatformRestrictionConfiguration' -or
           $type -eq '#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration')
    {
        Invoke-CDDocumentDeviceEnrollmentPlatformRestrictionConfiguration $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }
    elseif($type -eq '#microsoft.graph.deviceAndAppManagementRoleDefinition')
    {
        Invoke-CDDocumentDeviceAndAppManagementRoleDefinition $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }
    elseif($type -eq '#microsoft.graph.deviceComplianceScript')
    {
        Invoke-CDDocumentDeviceComplianceScript $documentationObj
        return [PSCustomObject]@{
            Properties = @("Name","Value","Category","SubCategory")
        }
    }
    elseif($type -eq '#microsoft.graph.roleScopeTag')
    {
        Invoke-CDDocumentScopeTag $documentationObj
        return $true
    }
}

function Get-CDAllManagedApps
{
    if(-not $script:allManagedApps)
    {
        $script:allManagedApps = (Invoke-GraphRequest -Url "/deviceAppManagement/managedAppStatuses('managedAppList')").content.appList
    }
    $script:allManagedApps
}

function Get-CDAllCloudApps
{
    if(-not $script:allCloudApps)
    {
        $script:allCloudApps = (Invoke-GraphRequest -url "/servicePrincipals?`$select=displayName,appId&top=999" -ODataMetadata "minimal" -AllPages).value
    }
    $script:allCloudApps
}

function Get-CDAllTenantApps
{
    if(-not $script:allTenantApps)
    {
        $script:allTenantApps = Get-DocOfflineObjects "Applications"
        if(-not $script:allTenantApps)
        {
            $script:allTenantApps =(Invoke-GraphRequest -url "/deviceAppManagement/mobileApps?`$select=displayName,id&top=999" -ODataMetadata "minimal" -AllPages).value
        }
    }
    $script:allTenantApps
}

function Get-CDMobileApps
{
    param($apps)

    $managedApps = Get-CDAllManagedApps
    $publishedApps = @()
    $customApps = @()
    foreach($tmpApp in $apps)
    {
        $appObj = $managedApps | Where { (($tmpApp.mobileAppIdentifier.packageId -and $_.appIdentifier.packageId -eq $tmpApp.mobileAppIdentifier.packageId) -or ($tmpApp.mobileAppIdentifier.bundleId -and $_.appIdentifier.bundleId -eq $tmpApp.mobileAppIdentifier.bundleId)) -and $_.appIdentifier."@odata.type" -eq $tmpApp.mobileAppIdentifier."@odata.type" }
        if($appObj -and $appObj.isFirstParty)
        {
            $publishedApps += $appObj.displayName
        }
        elseif($appObj)
        {
            $customApps += $appObj.displayName
        }
    }

    @($customApps,$publishedApps)
}

<#
.SYNOPSIS
Custom documentation for a value 

.DESCRIPTION
Ignore or create a custom value for a property
Return false to skip further processing of the property

.PARAMETER obj
The object to check. This could be a property of the profile object

.PARAMETER prop
Current property

.PARAMETER topObj
The profile object 

.PARAMETER propSeparator
Property separator character

.PARAMETER objSeparator
Object separator character
#>

function Invoke-CDDocumentCustomPostAdd
{
    param($obj, $prop, $propSeparator, $objSeparator)

    if($obj.'@OData.Type' -eq "#microsoft.graph.windowsUpdateForBusinessConfiguration")
    {
        if($prop.EntityKey -eq "featureUpdatesDeferralPeriodInDays")
        {
            # Inject Windows 11 update setting. Not included in the file
            $tmpProp = [PSCustomObject]@{
                nameResourceKey = "allowWindows11UpgradeName"
                descriptionResourceKey = "allowWindows11UpgradeDescription"
                entityKey = "allowWindows11Upgrade"
                dataType = 0
                booleanActions = 109
                category = $prop.Category
            }
            $propValue = Invoke-TranslateBoolean $obj $tmpProp

            $script:UpdateCategory = $prop.Category

            Add-PropertyInfo $tmpProp $propValue -originalValue $obj.allowWindows11Upgrade
        }

        if($prop.EntityKey -eq "featureUpdatesRollbackWindowInDays")
        {
            if($obj.businessReadyUpdatesOnly -eq "businessReadyOnly" -or $obj.businessReadyUpdatesOnly -eq "all" -or $obj.businessReadyUpdatesOnly -eq "userDefined")
            {
                $propValue = Get-LanguageString "BooleanActions.notConfigured"
            }
            else
            {
                $propValue = Get-LanguageString "BooleanActions.enable"
            }

            # Inject Pre-release setting. Not included in the file
            $tmpProp = [PSCustomObject]@{
                nameResourceKey = "preReleaseBuilds"
                descriptionResourceKey = "preReleaseBuildsDescription"
                entityKey = "preReleaseEnabled" # Not a class property!
                dataType = 0
                booleanActions = 2
                category = $prop.Category
            }

            Add-PropertyInfo $tmpProp $propValue -originalValue $obj.businessReadyUpdatesOnly

            if($obj.businessReadyUpdatesOnly -ne "businessReadyOnly" -and $obj.businessReadyUpdatesOnly -ne "all" -and $obj.businessReadyUpdatesOnly -ne "userDefined")
            {   
                # Pre-release channel selected. Inject info 
                $propValue = Get-LanguageString "SettingDetails.$($obj.businessReadyUpdatesOnly)Option"

                $tmpProp = [PSCustomObject]@{
                    nameResourceKey = "preReleaseChannel"
                    descriptionResourceKey = "preReleaseBuildsDescription"
                    entityKey = "businessReadyUpdatesOnly"
                    dataType = 0
                    booleanActions = 2
                    category = $prop.Category
                }

                Add-PropertyInfo $tmpProp $propValue -originalValue $obj.businessReadyUpdatesOnly
            }
        }
    }
}

function Add-CDDocumentCustomProfileValue
{
    param($obj, $prop, $topObj, $propSeparator, $objSeparator)
    
    if($obj.'@OData.Type' -eq "#microsoft.graph.windowsDeliveryOptimizationConfiguration" -and
        $prop.entityKey -eq "groupIdSourceSelector")
    {
        Invoke-TranslateOption $obj $prop -SkipOptionChildren | Out-Null
        return $false
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.androidManagedAppProtection" -or 
        $obj.'@OData.Type' -eq "#microsoft.graph.iosManagedAppProtection")
    {
        if($prop.entityKey -eq "apps")
        {
            $customApps,$publishedApps = Get-CDMobileApps $obj.Apps

            Add-PropertyInfo $prop ($publishedApps -join $objSeparator) -originalValue ($publishedApps -join $propSeparator)
            $propInfo = Get-PropertyInfo $prop ($customApps -join $objSeparator) -originalValue ($customApps -join $propSeparator)
            $propInfo.Name = Get-LanguageString "SettingDetails.customApps"
            $propInfo.Description = ""
            Add-PropertyInfoObject $propInfo
            return $false
        }        
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.windowsInformationProtectionPolicy" -or 
        $obj.'@OData.Type' -eq "#microsoft.graph.mdmWindowsInformationProtectionPolicy")
    {
        if($prop.entityKey -eq "enterpriseIPRanges")
        {
            $IPRanges = @()

            foreach($ipRange in $obj.enterpriseIPRanges)
            {
                $ranges = @()
                
                foreach($range in $ipRange.ranges)
                {
                    $ranges += ($range.lowerAddress + '-' + $range.upperAddress)
                }

                if($ranges.Count -gt 0)
                {
                    $IPRanges += ($ipRange.displayName + $propSeparator + ($ranges -join $propSeparator))
                }
            }

            $tmpArr = ($IPRanges | Where {$_.Contains('.')})
            if(($tmpArr | measure).Count -gt 0)
            {
                foreach($ipV4 in $tmpArr)
                {
                    Add-PropertyInfo $prop $ipV4 -originalValue $ipV4
                }
            }
            else
            {
                Add-PropertyInfo $prop $null
            }

            $tmpArr = ($IPRanges | Where {$_.Contains(':')})            
            
            if(($tmpArr | measure).Count -gt 0)
            {
                foreach($ipV6 in $tmpArr)
                {
                    $propInfo = Get-PropertyInfo $prop $ipV6 -originalValue $ipV6
                    $propInfo.Name = Get-LanguageString "WipPolicySettings.iPv6Ranges"
                    Add-PropertyInfoObject $propInfo
                }
            }
            else
            {
                $propInfo = Get-PropertyInfo $prop $null
                $propInfo.Name = Get-LanguageString "WipPolicySettings.iPv6Ranges"
                Add-PropertyInfoObject $propInfo
            }
            
            return $false
        }
        elseif($prop.entityKey -eq "enterpriseProxiedDomains")
        {
            foreach($tmpObj in $obj.enterpriseProxiedDomains)
            {
                $propValue = ($tmpObj.displayName + $propSeparator + ($tmpObj.proxiedDomains.ipAddressOrFQDN -join $propSeparator))
                Add-PropertyInfo $prop $propValue -originalValue $propValue
            }
            return $false
        }
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.windows*SCEPCertificateProfile")
    {
        if($prop.entityKey -eq "subjectNameFormat" -or $prop.entityKey -eq "subjectAlternativeNameType")
        {
            return $false # Skip these properties
        }        
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.windows10GeneralConfiguration")
    {        
        if($prop.EntityKey -eq "startMenuAppListVisibility")
        {
            $value = $obj.startMenuAppListVisibility
            if($value.IndexOf(", ") -eq -1)
            {
                $value = $value -replace ",",", " # Option values in json file has space afte , but value in object don't
            }
            Invoke-TranslateOption $obj $prop -PropValue $value
            return $false
        }

        $privacyAccessControls = $obj.privacyAccessControls | Where { $_.dataCategory -eq $prop.EntityKey -and $_.appDisplayName -eq $null }
        if($privacyAccessControls)
        {
            Invoke-TranslateOption $privacyAccessControls $prop -PropValue ($privacyAccessControls.accessLevel)
            return $false
        }
    }
    elseif($topObj.'@OData.Type' -like "#microsoft.graph.windows10EndpointProtectionConfiguration")
    {
        if($prop.EntityKey -eq "applicationGuardEnabled") { return $false }
        elseif($prop.EntityKey -eq "bitLockerRecoveryPasswordRotation") 
        { 
            Invoke-TranslateOption  $topObj $prop 
            return $false
        }
    }
    elseif($topObj.'@OData.Type' -like "#microsoft.graph.windowsHealthMonitoringConfiguration")
    {
        if($prop.EntityKey -eq "configDeviceHealthMonitoringScope") 
        { 
            if(($prop.options | Where value -eq "healthMonitoring"))
            {
                # Duplicate sections for health monitoring. Remove the old one
                return $false
            }
        }
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.windows10VpnConfiguration")
    {
        if($prop.EntityKey -eq "enableSplitTunneling" -and $prop.enabled -eq $false) 
        { 
            # SplitTunneling settings are moved to another file
            return $false
        }
        elseif($prop.EntityKey -eq "eapXml" -and $obj.eapXml)
        {
            $propValue = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.eapXml)) 
            Add-PropertyInfo $prop $propValue -originalValue $propValue
            return $false
        }                
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.windowsUpdateForBusinessConfiguration")
    {
        if($prop.EntityKey -eq "businessReadyUpdatesOnly" -or 
            $prop.EntityKey -eq "autoRestartNotificationDismissal" -or 
            $prop.EntityKey -eq "scheduleRestartWarningInHours" -or 
            $prop.EntityKey -eq "scheduleImminentRestartWarningInMinutes" -or
            $prop.EntityKey -eq "deliveryOptimizationMode") 
        { 
            # Not used anymore
            return $false
        }             
    }
}

<#
.SYNOPSIS
Change property source object before getting the property 

.DESCRIPTION
By default the object itself is always used when checking property values. 
This function changes the source object BEFORE a property is documented

.PARAMETER obj
The object to check

.PARAMETER prop
Current property

#>
function Get-CDDocumentCustomPropertyObject
{
    param($obj, $prop)

    if($obj.'@OData.Type' -like "#microsoft.graph.windows10EndpointProtectionConfiguration")
    {
        if($prop.EntityKey -eq "startupAuthenticationRequired")
        {
            return $obj.bitLockerSystemDrivePolicy
        }
        elseif($prop.EntityKey -eq "bitLockerSyntheticFixedDrivePolicyrequireEncryptionForWriteAccess")
        {
            return $obj.bitLockerFixedDrivePolicy
        }
        elseif($prop.EntityKey -eq "bitLockerSyntheticRemovableDrivePolicyrequireEncryptionForWriteAccess")
        {
            return $obj.bitLockerRemovableDrivePolicy
        }        
    }

    <#
    if($obj.'@OData.Type' -like "#microsoft.graph.windowsKioskConfiguration")
    {
        if($prop.nameResourceKey -eq "kioskSelectionName")
        {
            return $obj.kioskProfiles[0].appConfiguration
        }
    }
    #>
}

<#
.SYNOPSIS
Changes the source object to use for child properties

.DESCRIPTION
By default the object itself is always used when getting property values. 
This function changes the source property AFTER the property is processed but BEFORE child properties are documented

.PARAMETER obj
The object to check

.PARAMETER prop
Current property

#>
function Get-CDDocumentCustomChildObject
{
    param($obj, $prop)

    if($obj.'@OData.Type' -like "#microsoft.graph.windows10GeneralConfiguration")
    {
        if($prop.EntityKey -eq "syntheticDefenderDetectedMalwareActionsEnabled")
        {
            return $obj.defenderDetectedMalwareActions
        }
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.iosDeviceFeaturesConfiguration")
    {
        if($prop.EntityKey -eq "kerberosPrincipalName")
        {
            return $obj.singleSignOnSettings
        }
        elseif($prop.EntityKey -eq "singleSignOnExtensionType")
        {
            return $obj.iosSingleSignOnExtension
        }
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.macOSDeviceFeaturesConfiguration")
    {
        if($prop.EntityKey -eq "singleSignOnExtensionType")
        {
            return $obj.macOSSingleSignOnExtension
        }
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.windows10EndpointProtectionConfiguration")
    {
        if($prop.EntityKey -eq "applicationGuardPrintSettings")
        {
            return $obj.applicationGuardPrintSettings
        }
        if($prop.EntityKey -eq "firewallSyntheticIPsecExemptions")
        {
            return $obj.firewallSyntheticIPsecExemptions
        }
    }    
}

<#
.SYNOPSIS
Add cutom properties to the object

.DESCRIPTION
Many of the properties in profile translation files are based on calculated values. This function will add these extra properties to the object

.PARAMETER obj
The object to check

.PARAMETER propSeparator
Property separator character

.PARAMETER objSeparator
Object separator character

#>
function Add-CDDocumentCustomProfileProperty
{
    param($obj, $propSeparator, $objSeparator)

    $retValue = $false

    if($obj.'@OData.Type' -eq "#microsoft.graph.androidWorkProfileGeneralDeviceConfiguration" -or
            $obj.'@OData.Type' -eq "#microsoft.graph.androidDeviceOwnerGeneralDeviceConfiguration")
    {
        #Build vpnAlwaysOnPackageIdentifierSelector property
        $packageId = $null
        if(![String]::IsNullOrEmpty($obj.vpnAlwaysOnPackageIdentifier))
        {
            if(-not $obj.vpnAlwaysOnPackageIdentifier -or $obj.vpnAlwaysOnPackageIdentifier -notin @("com.cisco.anyconnect.vpn.android.avf","com.f5.edge.client_ics","com.paloaltonetworks.globalprotect","net.pulsesecure.pulsesecure"))
            {
                $packageId = "custom"
            }
            else
            {
                $packageId = $obj.vpnAlwaysOnPackageIdentifier
            }
        }
        $obj | Add-Member Noteproperty -Name "vpnAlwaysOnPackageIdentifierSelector" -Value $packageId -Force        
        $obj | Add-Member Noteproperty -Name "vpnAlwaysOnEnabled" -Value (![String]::IsNullOrEmpty($obj.vpnAlwaysOnPackageIdentifier)) -Force

        if(($obj.PSObject.Properties | Where Name -eq "globalProxy"))
        {
            $obj | Add-Member Noteproperty -Name "globalProxyEnabled" -Value ($obj.globalProxy -ne $null) -Force
            if($obj.globalProxy.proxyAutoConfigURL)
            {
                $globalProxyTypeSelector = "proxyAutoConfig"
                $obj | Add-Member Noteproperty -Name "globalProxyProxyAutoConfigURL" -Value $obj.globalProxy.proxyAutoConfigURL -Force
            }
            if($obj.globalProxy.host)
            {
                $globalProxyTypeSelector = "direct"
                $obj | Add-Member Noteproperty -Name "globalProxyHost" -Value $obj.globalProxy.host -Force
                $obj | Add-Member Noteproperty -Name "globalProxyPort" -Value $obj.globalProxy.port -Force
                $obj | Add-Member Noteproperty -Name "globalProxyExcludedHosts" -Value $obj.globalProxy.excludedHosts -Force
            }
            $obj | Add-Member Noteproperty -Name "globalProxyTypeSelector" -Value $globalProxyTypeSelector  -Force
        }

        if(($obj.PSObject.Properties | Where Name -eq "factoryResetDeviceAdministratorEmails"))
        {
            $factoryResetProtections = "factoryResetProtectionDisabled"
            if(($obj.factoryResetDeviceAdministratorEmails | measure).Count -gt 0)
            {
                $factoryResetProtections = "factoryResetProtectionEnabled"
            }
            $obj | Add-Member Noteproperty -Name "factoryResetProtections" -Value $factoryResetProtections -Force
            $obj | Add-Member Noteproperty -Name "googleAccountEmailAddressesList" -Value ($obj.factoryResetDeviceAdministratorEmails -join $objSeparator) -Force
        }
        
        if(($obj.PSObject.Properties | Where Name -eq "passwordBlockKeyguardFeatures"))
        {
            $obj | Add-Member Noteproperty -Name "passwordBlockKeyguardFeaturesList" -Value $obj.passwordBlockKeyguardFeatures -Force
        }
        
        if(($obj.PSObject.Properties | Where Name -eq "stayOnModes"))
        {
            $obj | Add-Member Noteproperty -Name "stayOnModesList" -Value $obj.stayOnModes -Force
        }

        if(($obj.PSObject.Properties | Where Name -eq "playStoreMode"))
        {
            $obj | Add-Member Noteproperty -Name "publicPlayStoreEnabled" -Value ($obj.playStoreMode -eq "blockList") -Force
        }
        $retValue = $true
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.androidEasEmailProfileConfiguration")
    {
        if(!($obj.PSObject.Properties | Where Name -eq "domainNameSourceType"))
        {
            $obj | Add-Member Noteproperty -Name "domainNameSourceType" -Value (?: ($obj.customDomainName -ne $null) "CustomDomainName" "AAD") -Force
        }
        $retValue = $true
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.windowsDeliveryOptimizationConfiguration")
    {
        if(!($obj.PSObject.Properties | Where Name -eq "groupIdSourceSelector"))
        {
            $obj | Add-Member Noteproperty -Name "groupIdSourceSelector" -Value (?? $obj.groupIdSource.groupIdSourceOption "notConfigured") -Force
        }
        $retValue = $true
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.windows10GeneralConfiguration")
    {
        if(!($obj.PSObject.Properties | Where Name -eq "networkProxyUseScriptUrlName"))
        {
            $obj | Add-Member Noteproperty -Name "networkProxyUseScriptUrlName" -Value ([String]::IsNullOrEmpty($obj.networkProxyAutomaticConfigurationUrl) -ne $null) -Force
        }

        $obj | Add-Member Noteproperty -Name "syntheticDefenderDetectedMalwareActionsEnabled" -Value ($obj.defenderDetectedMalwareActions -ne $null) -Force
        
        if(!($obj.PSObject.Properties | Where Name -eq "networkProxyUseManualServerName"))
        {
            $obj | Add-Member Noteproperty -Name "networkProxyUseManualServerName" -Value ($obj.networkProxyServer.address -ne $null) -Force
            if($obj.networkProxyServer.address -ne $null)
            {
                $obj | Add-Member Noteproperty -Name "networkProxyServerName" -Value $obj.networkProxyServer.address.Split(':')[0] -Force
                $obj | Add-Member Noteproperty -Name "networkProxyServerPort" -Value $obj.networkProxyServer.address.Split(':')[1] -Force
            }
            else
            {
                $obj | Add-Member Noteproperty -Name "networkProxyServerName" -Value "" -Force
                $obj | Add-Member Noteproperty -Name "networkProxyServerPort" -Value "" -Force
            }
            $exceptions = $null
            if($obj.networkProxyServer.exceptions)
            {
                $exceptions = ($obj.networkProxyServer.exceptions -join $propSeparator)
            }
            $obj | Add-Member Noteproperty -Name "networkProxyExceptionsTextString" -Value $exceptions -Force
            $obj | Add-Member Noteproperty -Name "useForLocalAddresses" -Value ($obj.networkProxyServer.useForLocalAddresses -eq $true) -Force
        }

        $obj | Add-Member Noteproperty -Name "edgeDisplayHomeButton" -Value ($obj.networkProxyServer.useForLocalAddresses -eq $true) -Force

        $searchEngineValue = 0
        if($obj.edgeSearchEngine.edgeSearchEngineOpenSearchXmlUrl -eq "default")
        {
            $searchEngineValue = 1
        }
        elseif($obj.edgeSearchEngine.edgeSearchEngineOpenSearchXmlUrl -eq "bing")
        {
            $searchEngineValue = 2
        }
        elseif($obj.edgeSearchEngine.edgeSearchEngineOpenSearchXmlUrl -eq "https://go.microsoft.com/fwlink/?linkid=842596")
        {
            $searchEngineValue = 3
        }
        elseif($obj.edgeSearchEngine.edgeSearchEngineOpenSearchXmlUrl -eq "https://go.microsoft.com/fwlink/?linkid=842600")
        {
            $searchEngineValue = 4
        }
        elseif($obj.edgeSearchEngine.edgeSearchEngineOpenSearchXmlUrl)
        {
            $searchEngineValue = 5
        }

        $obj | Add-Member Noteproperty -Name "edgeSearchEngineDropDown" -Value $searchEngineValue -Force

        $privacyApps = $obj.privacyAccessControls | Where { $_.appDisplayName -ne $null }

        $curApp = $null

        $perAppPrivacy = @()
        foreach($appItem in $privacyApps)
        {
            if($curApp -ne $appItem.appDisplayName)
            {
                $perAppPrivacy += [PSCustomObject]@{
                    appPackageName = $appItem.appPackageFamilyName
                    appName = $appItem.appDisplayName
                    #exceptions = $obj.privacyAccessControls | Where { $_.appPackageFamilyName -ne $appItem.appPackageFamilyName }
                }
                #($appItem.appPackageFamilyName + $propSeparator + $appItem.appDisplayName)
                $curApp = $appItem.appDisplayName
            }
        }
        $obj | Add-Member Noteproperty -Name "perAppPrivacy" -Value $perAppPrivacy -Force
        
        $retValue = $true
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.androidManagedAppProtection")
    {
        $obj | Add-Member Noteproperty -Name "overrideFingerprint" -Value ($obj.pinRequiredInsteadOfBiometricTimeout -ne $null -and $obj.pinRequiredInsteadOfBiometricTimeout -ne "PT0S")
        $obj | Add-Member Noteproperty -Name "pinReset" -Value ($obj.periodBeforePinReset -ne $null -and $obj.periodBeforePinReset -ne "PT0S")
        $obj | Add-Member Noteproperty -Name "managedBrowserSelection" -Value (?: $obj.customBrowserPackageId  "unmanagedBrowser" $obj.managedBrowser)
        $obj | Add-Member Noteproperty -Name "encryptOrgData" -Value ($obj.appDataEncryptionType -ne "useDeviceSettings")
        
        $retValue = $true
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.iosManagedAppProtection")
    {
        $sendDataOption = $obj.allowedOutboundDataTransferDestinations 
        if($obj.allowedOutboundDataTransferDestinations -eq "managedApps")
        {
            if($obj.disableProtectionOfManagedOutboundOpenInData -eq $false -and 
                $obj.filterOpenInToOnlyManagedApps -eq $true)
                {
                    $sendDataOption = "managedAppsWithOpenInSharing"
                }
            elseif($obj.disableProtectionOfManagedOutboundOpenInData -eq $true -and 
                $obj.filterOpenInToOnlyManagedApps -eq $false)
                {
                    $sendDataOption = "managedAppsWithOSSharing"
                }
        }

        $obj | Add-Member Noteproperty -Name "sendDataSelector" -Value $sendDataOption

        $obj | Add-Member Noteproperty -Name "overrideFingerprint" -Value ($obj.pinRequiredInsteadOfBiometricTimeout -ne $null -and $obj.pinRequiredInsteadOfBiometricTimeout -ne "PT0S")
        $obj | Add-Member Noteproperty -Name "pinReset" -Value ($obj.periodBeforePinReset -ne $null -and $obj.periodBeforePinReset -ne "PT0S")
        $obj | Add-Member Noteproperty -Name "managedBrowserSelection" -Value (?: $obj.customBrowserPackageId  "unmanagedBrowser" $obj.managedBrowser)
        $obj | Add-Member Noteproperty -Name "encryptOrgData" -Value ($obj.appDataEncryptionType -ne "useDeviceSettings")
        
        $retValue = $true
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.windowsUpdateForBusinessConfiguration")
    {
        $obj | Add-Member Noteproperty -Name "useDeadLineSettings" -Value ($obj.deadlineForFeatureUpdatesInDays -ne $null -or
                                                                            $obj.deadlineForQualityUpdatesInDays -ne $null -or
                                                                            $obj.deadlineGracePeriodInDays -ne $null -or
                                                                            $obj.postponeRebootUntilAfterDeadline -ne $null)
    
        $retValue = $true
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.azureADWindowsAutopilotDeploymentProfile" -or
           $obj.'@OData.Type' -eq "#microsoft.graph.activeDirectoryWindowsAutopilotDeploymentProfile")
    {
        $obj | Add-Member Noteproperty -Name "applyDeviceNameTemplate" -Value (?: ([String]::IsNullOrEmpty($obj.deviceNameTemplate)) $false  $true)

        if($obj.'@OData.Type' -eq "#microsoft.graph.azureADWindowsAutopilotDeploymentProfile")
        {
            $joinType = "azureAD"
        }
        else
        {
            $joinType = "hybrid"
        }

        $obj.outOfBoxExperienceSettings | Add-Member Noteproperty -Name "azureADJoinType" -Value $joinType

        $obj.outOfBoxExperienceSettings | Add-Member Noteproperty -Name "isLanguageSet" -Value (?: ([String]::IsNullOrEmpty($obj.language)) $false $true)
        
        if([String]::IsNullOrEmpty($obj.language))
        {
            $obj.language = "user-select"
        }
        
        $retValue = $true
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.officeSuiteApp")
    {
        $obj | Add-Member Noteproperty -Name "VersionToInstall" -Value (?: ([String]::IsNullOrEmpty($obj.targetVersion)) (Get-LanguageString "SettingDetails.latest") $obj.targetVersion)

        $obj | Add-Member Noteproperty -Name "useMicrosoftSearchAsDefault" -Value ($obj.excludedApps.bing -eq $false)

        if($obj.officeConfigurationXml)
        {
            $xmlConfig = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.officeConfigurationXml)) 
            $obj | Add-Member Noteproperty -Name "MSAppsConfigXml" -Value $xmlConfig
        }
        $retValue = $true
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.windowsWifiEnterpriseEAPConfiguration")
    {
        if($obj.authenticationMethod -ne "derivedCredential")
        {
            if($obj."#CustomRef_identityCertificateForClientAuthentication" -and $obj.'@ObjectFromFile' -eq $true)
            {
                $idCert = $obj."#CustomRef_identityCertificateForClientAuthentication"
                $idx = $idCert.IndexOf("|:|")
                if($idx -gt -1)
                {
                    $idCertType =  $idCert.SubString($idx + 3)
                }                                             
            }
            else
            {
                $idCert = Invoke-GraphRequest -URL $obj."identityCertificateForClientAuthentication@odata.navigationLink" -ODataMetadata "minimal" -NoError
                $idCertType = $idCert.'@OData.Type'
            }

            if($idCertType -like "*Pkcs*")
            {
                $clientCertType = "PKCS certificate"
            }
            elseif($idCertType -like "*SCEP*")
            {
                $clientCertType = "SCEP certificate"
            }

            $obj.authenticationMethod = $clientCertType

            $retValue = $true
        }
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.windows10VpnConfiguration")
    {
        if($obj.windowsInformationProtectionDomain)
        {
            $syntheticWipOrApps = 1
        }
        elseif($obj.onlyAssociatedAppsCanUseConnection)
        {
            $syntheticWipOrApps = 2
        }
        else
        {
            $syntheticWipOrApps = 0
        }
        $obj | Add-Member Noteproperty -Name "syntheticWipOrApps" -Value $syntheticWipOrApps -Force
        
        if($null -eq $obj.profileTarget)
        {
            $obj.profileTarget = "user"
        }

        $retValue = $true
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.iosDeviceFeaturesConfiguration")
    {
        #singleSignOnSettings
        $obj | Add-Member Noteproperty -Name "kerberosPrincipalName" -Value (?? $obj.singleSignOnSettings.kerberosPrincipalName "notConfigured") -Force

        #iosSingleSignOnExtension
        $obj | Add-Member Noteproperty -Name "singleSignOnExtensionType" -Value (?? $obj.iosSingleSignOnExtension."@OData.Type" "notConfigured") -Force

        $retValue = $true
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.macOSDeviceFeaturesConfiguration")
    {
        #macOSSingleSignOnExtension
        $obj | Add-Member Noteproperty -Name "singleSignOnExtensionType" -Value (?? $obj.macOSSingleSignOnExtension."@OData.Type" "notConfigured") -Force

        $retValue = $true
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.windows10EndpointProtectionConfiguration")
    {
        $allowPrintProps = $obj.PSObject.Properties | Where { $_.Name -like "applicationGuardAllowPrint*" }
        $obj | Add-Member Noteproperty -Name "applicationGuardAllowPrinting" -Value (($allowPrintProps | Where Value -eq $true).Count -gt 0)-Force
        $obj | Add-Member Noteproperty -Name "applicationGuardPrintSettings" -Value @(($allowPrintProps | Where Value -eq $true).Name) -Force
        
        $fwProps = $obj.PSObject.Properties | Where { $_.Name -like "firewallIPSecExemptionsAllow*" }
        $obj | Add-Member Noteproperty -Name "firewallSyntheticPresharedKeyEncodingMethod" -Value (($fwProps | Where Value -eq $true).Count -gt 0)-Force
        $obj | Add-Member Noteproperty -Name "firewallSyntheticIPsecExemptions" -Value @(($fwProps | Where Value -eq $true).Name) -Force

        $obj | Add-Member Noteproperty -Name "firewallSyntheticProfileDomainfirewallEnabled" -Value @($obj.firewallProfileDomain -ne $null) -Force
        $obj | Add-Member Noteproperty -Name "firewallSyntheticProfilePrivatefirewallEnabled" -Value @($obj.firewallProfilePrivate -ne $null) -Force
        $obj | Add-Member Noteproperty -Name "firewallSyntheticProfilePublicfirewallEnabled" -Value @($obj.firewallProfilePublic -ne $null) -Force

        Add-DefenderFirewallSettings $obj.firewallProfileDomain "Domain"
        Add-DefenderFirewallSettings $obj.firewallProfilePrivate "Private"
        Add-DefenderFirewallSettings $obj.firewallProfilePublic "Public"

        $obj | Add-Member Noteproperty -Name "bitLockerBaseConfigureEncryptionMethods" -Value (?: ($obj.bitLockerSystemDrivePolicy.encryptionMethod -ne $null) $true $null) -Force
        $obj | Add-Member Noteproperty -Name "bitLockerSystemDriveEncryptionMethod" -Value $obj.bitLockerSystemDrivePolicy.encryptionMethod -Force
        $obj | Add-Member Noteproperty -Name "bitLockerFixedDriveEncryptionMethod" -Value $obj.bitLockerFixedDrivePolicy.encryptionMethod -Force
        $obj | Add-Member Noteproperty -Name "bitLockerRemovableDriveEncryptionMethod" -Value $obj.bitLockerRemovableDrivePolicy.encryptionMethod -Force

        $obj.bitLockerSystemDrivePolicy | Add-Member Noteproperty -Name "bitLockerMinimumPinLength" -Value (?: ($obj.bitLockerSystemDrivePolicy.minimumPinLength -ne $null) $true $null) -Force
        $obj.bitLockerSystemDrivePolicy | Add-Member Noteproperty -Name "bitLockerSyntheticSystemDrivePolicybitLockerDriveRecovery" -Value (?: ($obj.bitLockerSystemDrivePolicy.recoveryOptions -ne $null) $true $null)  -Force
        
        if($obj.bitLockerSystemDrivePolicy.prebootRecoveryUrl -eq $null -and $obj.bitLockerSystemDrivePolicy.prebootRecoveryEnableMessageAndUrl -eq $null)
        {
            $bitLockerPrebootRecoveryMsgURLOption = "default"
        }
        elseif($obj.bitLockerSystemDrivePolicy.prebootRecoveryUrl -eq "" -and $obj.bitLockerSystemDrivePolicy.prebootRecoveryEnableMessageAndUrl -eq "")
        {
            $bitLockerPrebootRecoveryMsgURLOption = "empty"
        }
        elseif($obj.bitLockerSystemDrivePolicy.prebootRecoveryUrl)
        {
            $bitLockerPrebootRecoveryMsgURLOption = "customURL"
        }
        elseif($obj.bitLockerSystemDrivePolicy.prebootRecoveryEnableMessageAndUrl)
        {
            $bitLockerPrebootRecoveryMsgURLOption = "customMessage"
        }

        $obj.bitLockerSystemDrivePolicy | Add-Member Noteproperty -Name "bitLockerPrebootRecoveryMsgURLOption" -Value $bitLockerPrebootRecoveryMsgURLOption -Force
        
        foreach($tmpProp in ($obj.bitLockerSystemDrivePolicy.recoveryOptions.PSObject.Properties).Name)
        {
            $obj.bitLockerSystemDrivePolicy | Add-Member Noteproperty -Name "bitLockerSyntheticSystemDrivePolicy$($tmpProp)" -Value $obj.bitLockerSystemDrivePolicy.recoveryOptions.$tmpProp -Force
        }

        $obj.bitLockerFixedDrivePolicy | Add-Member Noteproperty -Name "bitLockerSyntheticFixedDrivePolicybitLockerDriveRecovery" -Value (?: ($obj.bitLockerFixedDrivePolicy.recoveryOptions -ne $null) $true $null) -Force

        foreach($tmpProp in ($obj.bitLockerFixedDrivePolicy.recoveryOptions.PSObject.Properties).Name)
        {
            $obj.bitLockerFixedDrivePolicy | Add-Member Noteproperty -Name "bitLockerSyntheticFixedDrivePolicy$($tmpProp)" -Value $obj.bitLockerFixedDrivePolicy.recoveryOptions.$tmpProp -Force
        }        

        $obj.bitLockerFixedDrivePolicy | Add-Member Noteproperty -Name "bitLockerSyntheticFixedDrivePolicyrequireEncryptionForWriteAccess" -Value $obj.bitLockerFixedDrivePolicy.requireEncryptionForWriteAccess -Force
        $obj.bitLockerRemovableDrivePolicy | Add-Member Noteproperty -Name "bitLockerSyntheticRemovableDrivePolicyrequireEncryptionForWriteAccess" -Value $obj.bitLockerRemovableDrivePolicy.requireEncryptionForWriteAccess -Force
        
        $appLockerApplicationControlType = "notConfigured"
        if($obj.appLockerApplicationControl -eq "enforceComponentsStoreAppsAndSmartlocker")
        {
            $appLockerApplicationControlType = "allow"
        }
        if($obj.appLockerApplicationControl -eq "auditComponentsAndStoreApps")
        {
            $appLockerApplicationControlType = "audit"
        }
        $obj | Add-Member Noteproperty -Name "appLockerApplicationControlType" -Value $appLockerApplicationControlType -Force

        $retValue = $true
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.iosGeneralDeviceConfiguration")
    {
        if([String]::IsNullOrEmpty($obj.KioskModeAppTypeDropDown))
        {
            $kioskMode = $null
            if($obj.kioskModeAppStoreUrl)
            {
                $kioskMode = 0
            }
            elseif($obj.kioskModeManagedAppId)
            {
                $kioskMode = 1
            }
            elseif($obj.kioskModeBuiltInAppId)
            {
                $kioskMode = 2
            }
            if($kioskMode -ne $null)
            {
                $obj | Add-Member Noteproperty -Name "KioskModeAppTypeDropDown" -Value $kioskMode -Force 
            }
        }

        $MediaContentRatingRegionSelectorDropDown = "notConfigured"
        foreach($mediaRatingProp in ($obj.PSObject.Properties | Where { $_.Name -like "mediaContentRating*" -and $_.Name -notlike "*@odata.type" -and $_.Name -ne "mediaContentRatingApps"}).Name)
        {
            if($obj.$mediaRatingProp -ne $null)
            {
                $MediaContentRatingRegionSelectorDropDown = $mediaRatingProp
                break
            }
        }
        $obj | Add-Member Noteproperty -Name "MediaContentRatingRegionSelectorDropDown" -Value $MediaContentRatingRegionSelectorDropDown -Force

        $networkUsageRulesCellularDataBlockType = "none"
        $networkUsageRulesCellularRoamingDataBlockType = "none"

        $tmpRule = $obj.networkUsageRules | Where cellularDataBlocked -eq $true
        if($tmpRule)
        {
            $networkUsageRulesCellularDataBlockType = ?: ($tmpRule.managedApps) "choose" "all"
            $obj | Add-Member Noteproperty -Name "networkUsageRulesCellularDataList" -Value ($tmpRule.managedApps -join $objSeparator) -Force
        }
        $tmpRule = $obj.networkUsageRules | Where cellularDataBlockWhenRoaming -eq $true
        if($tmpRule)
        {
            $networkUsageRulesCellularRoamingDataBlockType = ?: ($tmpRule.managedApps) "choose" "all"

            $obj | Add-Member Noteproperty -Name "networkUsageRulesCellularRoamingDataList" -Value $tmpRule.managedApps -Force
        }
        $obj | Add-Member Noteproperty -Name "networkUsageRulesCellularDataBlockType" -Value $networkUsageRulesCellularDataBlockType -Force
        $obj | Add-Member Noteproperty -Name "networkUsageRulesCellularRoamingDataBlockType" -Value $networkUsageRulesCellularRoamingDataBlockType -Force

        $retValue = $true
    }    
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.macOSEndpointProtectionConfiguration")
    {
        $firewallAllowedApps = $obj.firewallApplications | Where allowsIncomingConnections -eq $true
        $firewallBlockedApps = $obj.firewallApplications | Where allowsIncomingConnections -eq $false

        $obj | Add-Member Noteproperty -Name "firewallAllowedApps" -Value $firewallAllowedApps
        $obj | Add-Member Noteproperty -Name "firewallBlockedApps" -Value $firewallBlockedApps

        $retValue = $true
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.windowsFeatureUpdateProfile")
    {
        if(-not $script:win10FeatureUpdates)
        {
            $script:win10FeatureUpdates = (Invoke-GraphRequest -URL "/deviceManagement/windowsUpdateCatalogItems/microsoft.graph.windowsFeatureUpdateCatalogItem").value
        }

        $verInfo = $script:win10FeatureUpdates | Where version -eq $obj.featureUpdateVersion

        if($verInfo)
        {
            $verInfoTxt = $verInfo.displayName
        }
        else
        {
            $verInfoTxt = "{0} ({1})" -f $obj.featureUpdateVersion,(Get-LanguageString "WindowsFeatureUpdate.EndOFSupportStatus.notSupported")
        }

        $obj | Add-Member Noteproperty -Name "featureUpdateDisplayName" -Value $verInfoTxt

        if($obj.rolloutSettings.offerStartDateTimeInUTC -and
            $obj.rolloutSettings.offerEndDateTimeInUTC)
        {
            $featureUpdateRolloutOption = "gradualRollout"
            $obj | Add-Member Noteproperty -Name "featureUpdateRolloutStartDate" -Value ((Get-Date $obj.rolloutSettings.offerStartDateTimeInUTC).ToLongDateString())
            $obj | Add-Member Noteproperty -Name "featureUpdateRolloutEndDate" -Value ((Get-Date $obj.rolloutSettings.offerEndDateTimeInUTC).ToLongDateString())
            if($null -ne $obj.rolloutSettings.offerIntervalInDays)
            {
                $obj | Add-Member Noteproperty -Name "featureUpdateRolloutInterval" -Value ($obj.rolloutSettings.offerIntervalInDays)
            }
        }
        elseif($obj.rolloutSettings.offerStartDateTimeInUTC)
        {
            $featureUpdateRolloutOption = "startDateOnly"
            $obj | Add-Member Noteproperty -Name "featureUpdateRolloutStartDate" -Value ((Get-Date $obj.rolloutSettings.offerStartDateTimeInUTC).ToLongDateString())
        }
        else
        {
            $featureUpdateRolloutOption = "immediateStart"            
        }

        $obj | Add-Member Noteproperty -Name "featureUpdateRolloutOption" -Value $featureUpdateRolloutOption

        $retValue = $true
    }    
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.iosUpdateConfiguration")
    {
        if(-not $script:iOSAvailableUpdateVersions)
        {
            $script:iOSAvailableUpdateVersions = (Invoke-GraphRequest -URL "/deviceManagement/deviceConfigurations/getIosAvailableUpdateVersions").value
            $script:iOSAvailableUpdateVersions = $script:iOSAvailableUpdateVersions | Sort -property productVersion -Descending
        }

        $verInfo = $script:iOSAvailableUpdateVersions | Where productVersion -eq $obj.desiredOsVersion

        $versionText = "{0} {1}" -f (Get-LanguageString "SoftwareUpdates.IosUpdatePolicy.Settings.IOSVersion.prefix"), $obj.desiredOsVersion
        if(-not $verInfo)
        {
            $versionText = "$versionText ($(Get-LanguageString "SoftwareUpdates.IosUpdatePolicy.Settings.IOSVersion.noLongerSupported"))"
        }
        elseif($verInfo[0].productVersion -eq $obj.desiredOsVersion)
        {
            $versionText = "$versionText ($(Get-LanguageString "SoftwareUpdates.IosUpdatePolicy.Settings.IOSVersion.latestUpdate"))"
        }
        $obj | Add-Member Noteproperty -Name "versionInfo" -Value $versionText

        $timeWidows = @()
        foreach($timeWindow in $obj.customUpdateTimeWindows)
        {
            $startDay = Get-LanguageString "SettingDetails.$($timeWindow.startDay)"
            $endDay = Get-LanguageString "SettingDetails.$($timeWindow.endDay)"
            for($i = 0;$i -lt 2;$i++)
            {
                if($i -eq 0)
                {
                    $hour=[int]$timeWindow.startTime.Split(":")[0]
                }
                else
                {
                    $hour=[int]$timeWindow.endTime.Split(":")[0]
                }

                if($hour -gt 12)
                {
                    $when = "PM"
                    $hour = $hour - 12
                }
                else
                {
                    $when = "AM"
                }
                if($hour -eq 0) { $hourStr = "twelve" }
                elseif($hour -eq 1) { $hourStr = "one" }
                elseif($hour -eq 2) { $hourStr = "two" }
                elseif($hour -eq 3) { $hourStr = "three" }
                elseif($hour -eq 4) { $hourStr = "four" }
                elseif($hour -eq 5) { $hourStr = "five" }
                elseif($hour -eq 6) { $hourStr = "six" }
                elseif($hour -eq 7) { $hourStr = "seven" }
                elseif($hour -eq 8) { $hourStr = "eight" }
                elseif($hour -eq 9) { $hourStr = "nine" }
                elseif($hour -eq 10) { $hourStr = "ten" }
                elseif($hour -eq 11) { $hourStr = "eleven" }

                if($i -eq 0)
                {
                    $startTime = Get-LanguageString "SettingDetails.$($hourStr)$($when)Option"
                }
                else
                {
                    $endTime = Get-LanguageString "SettingDetails.$($hourStr)$($when)Option"
                }                
            }
            $timeWidows += ($startDay + $propSeparator + $startTime + $propSeparator + $endDay + $propSeparator + $endTime)
        }
        $obj | Add-Member Noteproperty -Name "timeWidows" -Value ($timeWidows -join $objSeparator)
    } 
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.windows10EnrollmentCompletionPageConfiguration")
    {
        if($obj.selectedMobileAppIds.Count -eq 0)
        {
            $apps = Get-LanguageString "EnrollmentStatusScreen.Apps.useSelectedAppsAll"
        }
        else
        {
            $allApps = Get-CDAllTenantApps
            $appsArr = @()
            foreach($appId in $obj.selectedMobileAppIds)
            {
                $tmpApp = $allApps | Where Id -eq $appId
                if($tmpApp)
                {
                    $appsArr += $tmpApp.displayName
                }
                else
                {
                    Write-Log "No app found with id $appId" 3
                }
            }
            $apps = $appsArr -join $objSeparator
        }
        $obj | Add-Member Noteproperty -Name "showCustomErrorMessage" -Value (-not [string]::IsNullOrEmpty($obj.customErrorMessage))
        $obj | Add-Member Noteproperty -Name "waitForApps" -Value $apps
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.win32LobApp")
    {
        $requirementRulesSummary = @()
        $detectionRulesSummary = @()
        $returnCodes = @()
        $detectionRules = @()
        $requirementRules = @()
        foreach($rc in $obj.returnCodes)
        {
            $returnCodes += [PSCustomObject]@{
                returnCode = $rc.returnCode
                type = (Get-LanguageString "Win32ReturnCodes.CodeTypes.$($rc.type)")
            }
            #$returnCodes += ("{0} {1}" -f @($rc.returnCode,(Get-LanguageString "Win32ReturnCodes.CodeTypes.$($rc.type)")))
        }

        $dependencyApps = @()
        $supersededApps = @()
        if($obj.dependentAppCount -gt 0 -or $obj.supersededAppCount -gt 0)
        {
            # ToDo: Add support for Offline documentation
            $relationships = (Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$($obj.Id)/relationships?`$filter=targetType%20eq%20microsoft.graph.mobileAppRelationshipType%27child%27").value
            foreach($rel in $relationships)
            {
                if($rel."@odata.type" -eq "#microsoft.graph.mobileAppDependency")
                {
                    $dependencyApps += ("{0} {1}" -f @($rel.targetDisplayName,(Get-LanguageString "SettingDetails.$((?: ($rel.dependencyType -eq "autoInstall") "win32DependenciesAutoInstall" "win32DependenciesDetect"))")))
                }
                elseif($rel."@odata.type" -eq "#microsoft.graph.mobileAppSupersedence")
                {
                    $supersededApps += ("{0} {1}" -f @($rel.targetDisplayName,(Get-LanguageString "SettingDetails.$((?: ($rel.supersedenceType -eq "update") "win32SupersedenceUpdate" "win32SupersedenceReplace"))")))
                }
            }
        }

        foreach($rule in $obj.requirementRules)
        {
            if($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppFileSystemRequirement")
            {
                $lngId = "fileType"
                $textValue = $rule.path
            }
            elseif($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppRegistryRequirement")
            {
                $lngId = "registry"
                $textValue = $rule.keyPath
            }
            else #win32LobAppProductCodeDetection
            {
                $lngId = "script"
                $textValue = $rule.displayName
                Add-ObjectScript $rule.displayName ("{0} - {1}" -f @($obj.displayName, "Requirement script")) $rule.ScriptContent
            }
            $requirementRulesSummary += ("{0} {1}" -f @((Get-LanguageString "Win32Requirements.AdditionalRequirements.RequirementTypeOptions.$lngId"),$textValue))
        
            $requirementRules += Add-CDDocumentRequirementRule $rule
        }

        if(($obj.detectionRules | Where '@OData.Type' -eq "#microsoft.graph.win32LobAppPowerShellScriptDetection"))
        {
            $detectionRulesType = Get-LanguageString "DetectionRules.RuleConfigurationOptions.customScript"
            foreach($rule in $obj.detectionRules)
            {
                $header = (Get-LanguageString "ProactiveRemediations.Create.Settings.DetectionScriptMultiLineTextBox.label")
                Add-ObjectScript $header ("{0} - {1}" -f @($obj.displayName,$header)) $rule.ScriptContent
            }
        }
        else
        {
            $detectionRulesType = Get-LanguageString "DetectionRules.RuleConfigurationOptions.manual"

            foreach($rule in $obj.detectionRules)
            {
                if($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppFileSystemDetection")
                {
                    $lngId = "file"
                    $textValue = $rule.path
                }
                elseif($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppRegistryDetection")
                {
                    $lngId = "registry"
                    $textValue = $rule.keyPath
                }
                else #win32LobAppProductCodeDetection
                {
                    $lngId = "mSI"
                    $textValue = $rule.productCode
                }
                
                $detectionRulesSummary += ("{0} {1}" -f @((Get-LanguageString "DetectionRules.Manual.RuleTypeOptions.$lngId"),$textValue))

                $detectionRules += Add-CDDocumentDetectionRule $rule
            }
        }
        
        $obj | Add-Member Noteproperty -Name "requirementRulesSummary" -Value ($requirementRulesSummary -join $objSeparator) -Force 
        $obj | Add-Member Noteproperty -Name "detectionRulesSummary" -Value ($detectionRulesSummary -join $objSeparator) -Force 
        $obj | Add-Member Noteproperty -Name "dependencyApps" -Value ($dependencyApps -join $objSeparator) -Force 
        $obj | Add-Member Noteproperty -Name "supersededApps" -Value ($supersededApps -join $objSeparator) -Force 
        $obj | Add-Member Noteproperty -Name "detectionRulesType" -Value $detectionRulesType -Force 
        $obj | Add-Member Noteproperty -Name "requirementRulesTranslated" -Value $requirementRules -Force 
        $obj | Add-Member Noteproperty -Name "detectionRulesTranslated" -Value $detectionRules -Force 
        $obj | Add-Member Noteproperty -Name "returnCodes" -Value $returnCodes -Force 
        $obj | Add-Member Noteproperty -Name "win10Release" -Value (Get-LanguageString "MinimumOperatingSystem.Windows.V10Release.release$($obj.minimumSupportedWindowsRelease)") -Force 
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.deviceHealthScript")
    {
        $obj | Add-Member Noteproperty -Name "detectionScriptAdded" -Value (-not [String]::IsNullOrEmpty($obj.detectionScriptContent))
        $obj | Add-Member Noteproperty -Name "remediationScriptAdded" -Value (-not [String]::IsNullOrEmpty($obj.remediationScriptContent))
        $obj | Add-Member Noteproperty -Name "useLoggedOnCredentials" -Value ($obj.runAsAccount -ne "system")

        if($obj.detectionScriptContent)
        {
            $obj | Add-Member Noteproperty -Name "detectionScriptContentString" -Value ([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(($obj.detectionScriptContent))))
            $header = Get-LanguageString "ProactiveRemediations.Create.Settings.DetectionScriptMultiLineTextBox.label"
            Add-ObjectScript $header ("{1} - {0}" -f $obj.displayName,$header) $obj.detectionScriptContent
        }
        if($obj.remediationScriptContent)
        {
            $obj | Add-Member Noteproperty -Name "remediationScriptContentString" -Value ([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(($obj.remediationScriptContent))))
            $header = Get-LanguageString "ProactiveRemediations.Create.Settings.RemediationScriptMultiLineTextBox.label"
            Add-ObjectScript $header ("{1} - {0}" -f $obj.displayName,$header) $obj.remediationScriptContent
        }
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.deviceManagementScript")
    {
        if($obj.ScriptContent)
        {
            Add-ObjectScript $obj.FileName ("{1} - {0}" -f $obj.displayName,(Get-LanguageString "WindowsManagement.powerShellScriptObjectName")) $obj.ScriptContent
        }
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.deviceShellScript")
    {
        if($obj.ScriptContent)
        {
            Add-ObjectScript $obj.FileName ("{1} - {0}" -f $obj.displayName,(Get-LanguageString "WindowsManagement.shellScriptObjectName")) $obj.ScriptContent
        }
    }
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.deviceCustomAttributeShellScript")
    {
        if($obj.ScriptContent)
        {
            Add-ObjectScript $obj.FileName ("{1} - {0}" -f $obj.displayName,(Get-LanguageString "WindowsManagement.customAttributeObjectName")) $obj.ScriptContent
        }
    }    
    elseif($obj.'@OData.Type' -eq "#microsoft.graph.windows10TeamGeneralConfiguration")
    {
        $obj | Add-Member Noteproperty -Name "syntheticAzureOperationalInsightsEnabled" -Value ($obj.azureOperationalInsightsBlockTelemetry -eq $false)
        $obj | Add-Member Noteproperty -Name "syntheticMaintenanceWindowEnabled" -Value ($obj.maintenanceWindowBlocked -eq $false)
    }
    elseif($obj.'@OData.Type' -like "#microsoft.graph.windowsKioskConfiguration")
    {
        if($obj.kioskProfiles[0].appConfiguration."@odata.type" -eq "#microsoft.graph.windowsKioskSingleWin32App")
        {
            $uwpAppType = "win32App"
            $obj.kioskProfiles[0].appConfiguration."@odata.type" = "#microsoft.graph.windowsKioskSingleUWPApp"
        }
        elseif($obj.kioskProfiles[0].appConfiguration.uwpApp.appUserModelId -like "Microsoft.MicrosoftEdge*")
        {
            $uwpAppType = "edge"
        }
        elseif($obj.kioskProfiles[0].appConfiguration.uwpApp.appUserModelId -like "Microsoft.KioskBrowser*")
        {
            $uwpAppType = "kioskBrowser"
        }
        elseif($obj.kioskProfiles[0].appConfiguration.uwpApp.appUserModelId)
        {
            $uwpAppType = "managed"
        }

        $obj.kioskProfiles[0].appConfiguration | Add-Member Noteproperty -Name "uwpAppType" -Value $uwpAppType
        
        if($obj.windowsKioskForceUpdateSchedule)
        {
            $obj | Add-Member Noteproperty -Name "hasForceRestart" -Value $true
        }
    }    
    elseif($obj.'@OData.Type' -like "#microsoft.graph.windowsWifiConfiguration")
    {
        if($obj.wifiSecurityType -eq "wpa2Personal")
        {
            $obj.preSharedKey = "********"
        }
    }

    if(($obj.PSObject.Properties | where Name -eq "securityRequireSafetyNetAttestationBasicIntegrity") -and 
    ($obj.PSObject.Properties | where Name -eq "securityRequireSafetyNetAttestationCertifiedDevice"))
    {
        $androidSafetyNetAttestationOptions = "notConfigured"
        if($obj.securityRequireSafetyNetAttestationBasicIntegrity -eq $true -and 
        $obj.securityRequireSafetyNetAttestationCertifiedDevice -eq $true)
        {
            $androidSafetyNetAttestationOptions = 'basicIntegrityAndCertified'
        }
        elseif($obj.securityRequireSafetyNetAttestationBasicIntegrity -eq $true)
        {
            $androidSafetyNetAttestationOptions = 'basicIntegrity'
        }
        $obj | Add-Member Noteproperty -Name "androidSafetyNetAttestationOptions" -Value $androidSafetyNetAttestationOptions -Force

        $retValue = $true
    }
    
    
    if(($obj.PSObject.Properties | Where Name -eq "periodOfflineBeforeWipeIsEnforced"))
    {
        #Conditional Launch settings for AppProtection policies

        $conditionalLaunch = @()

        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "maxPinAttempts" "maximumPinRetries" (?: ($obj.appActionIfMaximumPinRetriesExceeded -eq "block") "resetPin" "wipeData"))
        
        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "offlineGracePeriod" "periodOfflineBeforeAccessCheck" "blockMinutes")
        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "offlineGracePeriod" "periodOfflineBeforeWipeIsEnforced" "wipeDays")

        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "minAppVersion" "minimumWipeAppVersion" "wipeData")
        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "minAppVersion" "minimumRequiredAppVersion" "blockAccess")
        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "minAppVersion" "minimumWarningAppVersion" "warn")
        
        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "minSdkVersion" "minimumRequiredSdkVersion" "blockAccess")
        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "minSdkVersion" "minimumWipeSdkVersion" "wipeData")

        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "onlineButUnableToCheckin" "appActionIfUnableToAuthenticateUser" (?: ($obj.appActionIfUnableToAuthenticateUser -eq "block") "blockAccess" "wipeData") -SkipValue) 

        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "jailbrokenRootedDevices" "appActionIfDeviceComplianceRequired" (?: ($obj.appActionIfDeviceComplianceRequired -eq "block") "blockAccess" "wipeData") -SkipValue) 

        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "minOSVersion" "minimumWipeOsVersion" "wipeData")
        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "minOSVersion" "minimumRequiredOsVersion" "blockAccess")
        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "minOSVersion" "minimumWarningOsVersion" "warn")

        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "maxOSVersion" "maximumWipeOsVersion" "wipeData")
        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "maxOSVersion" "maximumRequiredOsVersion" "blockAccess")
        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "maxOSVersion" "maximumWarningOsVersion" "warn")

        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "deviceModels" "allowedIosDeviceModels" (?: ($obj.appActionIfIosDeviceModelNotAllowed -eq "block") "allowSpecifiedBlock" "allowSpecifiedWipe")) 

        $conditionalLaunch += (Get-ConditionalLaunchSetting $obj "maximumAllowedDeviceThreatLevel" "maximumAllowedDeviceThreatLevel" (?: ($obj.appActionIfDeviceComplianceRequired -eq "block") "blockAccess" "wipeData")) 

        if($conditionalLaunch.Count -gt 0)
        {
            $obj | Add-Member Noteproperty -Name "ConditionalLaunchSettings" -Value @($conditionalLaunch)
        }

        $retValue = $true
    }

    return $retValue
}

function Add-CDDocumentRequirementRule
{
    param($rule)

    $strYes = Get-LanguageString "SettingDetails.yes"
    $strNo = Get-LanguageString "SettingDetails.no"

    $ruleInfo = @()

    if($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppFileSystemRequirement")
    {
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.requirementType")
            value = (Get-LanguageString "Win32Requirements.AdditionalRequirements.RequirementTypeOptions.fileType")
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.FileRule.path")
            value = $rule.path
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.FileRule.fileOrFolder")
            value = $rule.fileOrFolderName
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.File.property")
            value = switch($rule.detectionType)
            {
                "createdDate" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.dateCreated") }
                "modifiedDate" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.dateModified") }
                "doesNotExist" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.doesNotExist") }
                "exists" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.fileOrFolderExists") }
                "sizeInMB" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.sizeInMB") }
                "version" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.version") }
                Default { Get-LanguageString "BooleanActions.notConfigured" }
            }
        }        
        
        if($rule.detectionValue -and $rule.operator)
        {
            $ruleInfo += [PSCustomObject]@{
                property = (Get-LanguageString "DetectionRules.Manual.FileRule.operator")
                value = (Get-CDDocumentOperatorString $rule.operator)
            }

            $detectionValue = $rule.detectionValue 
            if($rule.detectionType -eq "createdDate" -or $rule.detectionType -eq "modifiedDate")
            {
                try { 
                    $tmpDate = Get-Date $rule.detectionValue
                    $detectionValue = $tmpDate.ToString()
                } catch {}
            }
    
            $ruleInfo += [PSCustomObject]@{
                property = (Get-LanguageString "DetectionRules.Manual.FileRule.value")
                value = $detectionValue
            }
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.RegistryRule.associatedWith32Bit")
            value = (?: ($rule.check32BitOn64System -eq $true) ($strYes) ($strNo))
        }
    }
    elseif($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppRegistryRequirement")
    {
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.requirementType")
            value = (Get-LanguageString "Win32Requirements.AdditionalRequirements.RequirementTypeOptions.registry")
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Registry.keyPath")
            value = $rule.keyPath
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Registry.valueName")
            value = $rule.valueName
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Registry.registryRequirement")
            value = switch($rule.detectionType)
            {
                "doesNotExist" 
                {
                    if($rule.valueName)
                    {
                        (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.valueDoesNotExist")
                    }
                    else
                    {
                        (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.keyDoesNotExist")
                    }
                }
                "exists" { 
                    if($rule.valueName)
                    {
                        (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.valueExists")
                    }
                    else
                    {
                        (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.keyExists")
                    }
                }
                "integer" { (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.integerComparison") }
                "string" { (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.stringComparison") }
                "version" { (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.versionComparison") }
                Default { Get-LanguageString "BooleanActions.notConfigured" }
            }
        }        
        
        if($rule.detectionValue -and $rule.operator)
        {
            $ruleInfo += [PSCustomObject]@{
                property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Registry.operator")
                value = (Get-CDDocumentOperatorString $rule.operator)
            }
    
            $ruleInfo += [PSCustomObject]@{
                property = (Get-LanguageString "DetectionRules.Manual.RegistryRule.value")
                value = $rule.detectionValue
            }
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.RegistryRule.associatedWith32Bit")
            value = (?: ($rule.check32BitOn64System -eq $true) ($strYes) ($strNo))
        }
    }
    elseif($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppPowerShellScriptRequirement")
    {
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.requirementType")
            value = (Get-LanguageString "Win32Requirements.AdditionalRequirements.RequirementTypeOptions.script")
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.scriptName")
            value = $rule.displayName
        }

        <#
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.scriptContent")
            $scriptContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($rule.scriptContent))
            value = $scriptContent
        }
        #>
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.CustomScript.runAs32Bit")
            value = (?: ($rule.runAs32Bit -eq $true) ($strYes) ($strNo))
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.loggedOnCredentials")
            value = (?: ($rule.runAsAccount -ne "system") ($strYes) ($strNo))
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.enforceSignatureCheck")
            value = (?: ($rule.enforceSignatureCheck -eq $true) ($strYes) ($strNo))
        }        

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.requirementMethod")
            value = switch($rule.detectionType)
            {
                "string" { (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.string") }
                "dateTime" { (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.dateTime") }
                "integer" { (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.integer") }
                "float" { (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.float") }
                "version" { (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.version") }
                "boolean" { (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.boolean") }
                Default { Get-LanguageString "BooleanActions.notConfigured" }
            }
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Registry.operator")
            value = (Get-CDDocumentOperatorString $rule.operator)
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "Win32Requirements.AdditionalRequirements.Script.value")
            value = $rule.detectionValue
        }        
    }
    return $ruleInfo
}

function Add-CDDocumentDetectionRule
{
    param($rule)

    $ruleInfo = @()
    
    if($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppFileSystemDetection")
    {
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.ruleType")
            value = (Get-LanguageString "DetectionRules.Manual.RuleTypeOptions.file")
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.FileRule.path")
            value = $rule.path
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.FileRule.fileOrFolder")
            value = $rule.fileOrFolderName
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.FileRule.detectionMethod")
            value = switch($rule.detectionType)
            {
                "createdDate" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.dateCreated") }
                "modifiedDate" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.dateModified") }
                "doesNotExist" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.doesNotExist") }
                "exists" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.fileOrFolderExists") }
                "sizeInMB" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.sizeInMB") }
                "version" { (Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.version") }
                Default { Get-LanguageString "BooleanActions.notConfigured" }
            }
        }        
        
        if($rule.detectionValue -and $rule.operator)
        {
            $ruleInfo += [PSCustomObject]@{
                property = (Get-LanguageString "DetectionRules.Manual.FileRule.operator")
                value = (Get-CDDocumentOperatorString $rule.operator)
            }

            $detectionValue = $rule.detectionValue 
            if($rule.detectionType -eq "createdDate" -or $rule.detectionType -eq "modifiedDate")
            {
                try { 
                    $tmpDate = Get-Date $rule.detectionValue
                    $detectionValue = $tmpDate.ToString()
                } catch {}
            }
    
            $ruleInfo += [PSCustomObject]@{
                property = (Get-LanguageString "DetectionRules.Manual.FileRule.value")
                value = $detectionValue
            }
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.RegistryRule.associatedWith32Bit")
            value = (?: ($rule.check32BitOn64System -eq $true) (Get-LanguageString "SettingDetails.yes") (Get-LanguageString "SettingDetails.no"))
        }
    }
    elseif($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppRegistryDetection")
    {
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.ruleType")
            value = (Get-LanguageString "DetectionRules.Manual.RuleTypeOptions.registry")
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.RegistryRule.keyPath")
            value = $rule.keyPath
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.RegistryRule.valueName")
            value = $rule.valueName
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.RegistryRule.detectionMethod")
            value = switch($rule.detectionType)
            {
                "doesNotExist" 
                {
                    if($rule.valueName)
                    {
                        (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.valueDoesNotExist")
                    }
                    else
                    {
                        (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.keyDoesNotExist")
                    }
                }
                "exists" { 
                    if($rule.valueName)
                    {
                        (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.valueExists")
                    }
                    else
                    {
                        (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.keyExists")
                    }
                }
                "integer" { (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.integerComparison") }
                "string" { (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.stringComparison") }
                "version" { (Get-LanguageString "DetectionRules.Manual.RegistryRule.DetectionMethodOptions.versionComparison") }
                Default { Get-LanguageString "BooleanActions.notConfigured" }
            }
        }        
        
        if($rule.detectionValue -and $rule.operator)
        {
            $ruleInfo += [PSCustomObject]@{
                property = (Get-LanguageString "DetectionRules.Manual.RegistryRule.operator")
                value = (Get-CDDocumentOperatorString $rule.operator)
            }
    
            $ruleInfo += [PSCustomObject]@{
                property = (Get-LanguageString "DetectionRules.Manual.RegistryRule.value")
                value = $rule.detectionValue
            }
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.RegistryRule.associatedWith32Bit")
            value = (?: ($rule.check32BitOn64System -eq $true) (Get-LanguageString "SettingDetails.yes") (Get-LanguageString "SettingDetails.no"))
        }
    }
    else #win32LobAppProductCodeDetection
    {
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.ruleType")
            value = (Get-LanguageString "DetectionRules.Manual.RuleTypeOptions.mSI")
        }

        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.MsiRule.productCode")
            value = $rule.productCode
        }
        
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString "DetectionRules.Manual.MsiRule.productVersionCheck")
            value = (?: ($null -ne $rule.productVersion) (Get-LanguageString "SettingDetails.yes") (Get-LanguageString "SettingDetails.no"))
        }

        if($null -ne $rule.productVersion)
        {
            $ruleInfo += [PSCustomObject]@{
                property = (Get-LanguageString "DetectionRules.Manual.MsiRule.operator")
                value = (Get-CDDocumentOperatorString $rule.productVersionOperator)
            }
        }

        if($null -ne $rule.productVersion)
        {
            $ruleInfo += [PSCustomObject]@{
                property = (Get-LanguageString "DetectionRules.Manual.MsiRule.productVersion")
                value = (Get-CDDocumentOperatorString $rule.productVersion)
            }
        }        
    }    

    return $ruleInfo   
}

function Get-CDDocumentOperatorString
{
    param($operator)

    $lngString = switch ($operator)
    {
        "notConfigured" { Get-LanguageString "BooleanActions.notConfigured" }
        "equal" { Get-LanguageString "DetectionRules.ComparisonOperators.equals" }
        "notEqual" { Get-LanguageString "DetectionRules.ComparisonOperators.notEqualTo" }
        "greaterThan" { Get-LanguageString "DetectionRules.ComparisonOperators.greaterThan" }
        "greaterThanOrEqual" { Get-LanguageString "DetectionRules.ComparisonOperators.greaterThanOrEqualTo" }
        "lessThan" { Get-LanguageString "DetectionRules.ComparisonOperators.lessThan" }
        "lessThanOrEqual" { Get-LanguageString "DetectionRules.ComparisonOperators.lessThanOrEqualTo" }
        "exists" { Get-LanguageString "DetectionRules.Manual.FileRule.DetectionMethodOptions.fileOrFolderExists" }
        Default { $operator }
    }

    $lngString
}

# App Config
function Invoke-CDDocumentAndroidManagedStoreAppConfiguration
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicAdditionalValues $obj $objectType
    #Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "SettingDetails.appConfiguration")
    #Add-BasicPropertyValue (Get-LanguageString "Inputs.enrollmentTypeLabel") (Get-LanguageString "EnrollmentType.devicesWithEnrollment")

    $allApps = Get-CDAllTenantApps
    $appsList = @()

    foreach($id in ($obj.targetedMobileApps))
    {
        $tmpApp = $allApps | Where Id -eq $id
        $appsList += ?? $tmpApp.displayName $id
    }

    Add-BasicPropertyValue (Get-LanguageString "SettingDetails.targetedAppLabel") ($appsList -join $objSeparator)
    
    $category = Get-LanguageString "TableHeaders.settings"

    if($obj.payloadJson)
    {
        $payloadData = $null
        try
        {
            $payloadData = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.payloadJson)) | ConvertFrom-Json
        }
        catch
        {
            Write-LogError "Failed to get Json payload" $_.Exception
            return
        }

        # Not the best way. BundleId should be used but then full app info is required
        if($obj.packageId -eq "com.microsoft.office.outlook")
        {
            if([IO.File]::Exists(($global:AppRootFolder + "\Documentation\ObjectInfo\#AppConfigOutlookDevice.json")))
            {
                $tmp = $payloadData.managedProperty | Where { $_.key -eq "com.microsoft.outlook.EmailProfile.AccountType" }
                if($tmp){ $configEmail=$true }else{ $configEmail=$false }
                $outlookSettings = [PSCustomObject]@{
                    configureEmail = $configEmail
                }

                foreach($managedProperty in $payloadData.managedProperty)
                {
                    $valueProperty = $managedProperty.PSObject.Properties | Where-Object Name -like "value*"
                    $outlookSettings | Add-Member Noteproperty -Name $managedProperty.key -Value $valueProperty.Value -Force
                }

                $jsonObj = Get-Content ($global:AppRootFolder + "\Documentation\ObjectInfo\#AppConfigOutlookDevice.json") | ConvertFrom-Json
                Invoke-TranslateSection $outlookSettings $jsonObj
            }
        }                
        
        $addedSettings = Get-DocumentedSettings

        $additionalSettings = @()

        foreach($managedProperty in $payloadData.managedProperty)
        {
            if(($addedSettings | Where EntityKey -eq $managedProperty.key)) { continue }

            $valueProperty = $managedProperty.PSObject.Properties | Where-Object Name -like "value*"

            $value = $valueProperty.value

            if($value -is [Array]) {
                $value = $value -join ","
            }

            $additionalSettings += ([PSCustomObject]@{
                Name = $managedProperty.key
                ValueType = $valueProperty.Name.SubString(5)
                Value = $value
                EntityKey = $managedProperty.key
                Category = Get-LanguageString "TACSettings.generalSettings"
                SubCategory = Get-LanguageString "SettingDetails.additionalConfiguration"
            })
        }

        if($additionalSettings.Count -gt 0) {
            Add-CustomTable "AdditionalSettings" @("Name","ValueType","Value") $additionalSettings -Order 110
        }

        $permissions = @()

        foreach($permission in $obj.permissionActions)
        {
            $permissionTemp = $permission.permission.Split('.')[-1]
            if($permissionTemp) {
                $permissionLngId = $permissionTemp -replace "_", ""

                $permissionStr = ?? (Get-LanguageString "AndroidForWorkAppPermissions.Permissions.$($permissionLngId)") $permissionTemp
            }
            else {
                $permissionStr = $permission.permission
            }

            $permissions += ([PSCustomObject]@{
                Permission = $permissionStr
                Action = ?? (Get-LanguageString "AndroidForWorkAppPermissions.Action.$($permission.action)") $permission.action
                EntityKey = $permission.permission
            })
        }

        if($permissions.Count -gt 0) {
            Add-CustomTable "Permissions" @("Permission","Action") $permissions -Order 115 -LanguageId "AndroidForWorkAppPermissions.permissionsTitle"
        }
    }
}

function Invoke-CDDocumentMobileAppConfiguration
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicAdditionalValues $obj $objectType
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "SettingDetails.appConfiguration")
    Add-BasicPropertyValue (Get-LanguageString "Inputs.enrollmentTypeLabel") (Get-LanguageString "EnrollmentType.devicesWithEnrollment")
    
    $platformId = Get-ObjectPlatformFromType $obj
    Add-BasicPropertyValue (Get-LanguageString "Inputs.platformLabel") (Get-LanguageString "Platform.$platformId")

    $allApps = Get-CDAllTenantApps
    $appsList = @()
    foreach($id in ($obj.targetedMobileApps))
    {
        $tmpApp = $allApps | Where Id -eq $id
        $appsList += ?? $tmpApp.displayName $id
    }

    Add-BasicPropertyValue (Get-LanguageString "SettingDetails.targetedAppLabel") ($appsList -join $objSeparator)
    
    Add-BasicAdditionalValues $obj $objectType
    
    $category = Get-LanguageString "TableHeaders.settings"

    if($obj.encodedSettingXml)
    {
        $xml = $null
        try
        {
            $xml = [xml]([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.encodedSettingXml)))
        }
        catch
        {
            Write-LogError "Failed to convert XML data to XML" $_.Exception
            return
        }

        for($i = 0;$i -lt $xml.dict.ChildNodes.Count;$i++)
        {
            $name = $xml.dict.ChildNodes[$i].'#text'
            $i++
            $value = $xml.dict.ChildNodes[$i].'#text'

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $name
                Value = $value
                EntityKey = $name
                Category = $category
            })             
        }     
    }
    else 
    {
        $isOutlook = $false

        foreach($targetedAppId in $obj.targetedMobileApps) {
            $app = $allApps | Where Id -eq $targetedAppId
            if($app.displayName -eq "Microsoft Outlook") {
                $isOutlook = $true
                break
            }
        }

        # Not the best way. BundleId should be used but then full app info is required
        if($isOutlook -or ($obj.packageId | Where { $_.appConfigKey -like "com.microsoft.outlook*" }))
        {
            if([IO.File]::Exists(($global:AppRootFolder + "\Documentation\ObjectInfo\#AppConfigOutlookDevice.json")))
            {
                $tmp = $obj.settings | Where { $_.appConfigKey -eq "com.microsoft.outlook.EmailProfile.AccountType" }
                if($tmp){ $configEmail=$true }else{ $configEmail=$false }
                $outlookSettings = [PSCustomObject]@{
                    configureEmail = $configEmail
                }
                foreach($setting in $obj.settings)
                {
                    if($setting.appConfigKeyType -eq "booleanType")
                    {
                        $value = $setting.appConfigKeyValue -eq "true"
                    }
                    else
                    {
                        $value = $setting.appConfigKeyValue
                    }
                    $outlookSettings | Add-Member Noteproperty -Name $setting.appConfigKey -Value $value -Force
                }

                $jsonObj = Get-Content ($global:AppRootFolder + "\Documentation\ObjectInfo\#AppConfigOutlookDevice.json") | ConvertFrom-Json
                Invoke-TranslateSection $outlookSettings $jsonObj
            }
        }                
        
        $addedSettings = Get-DocumentedSettings

        foreach($setting in $obj.settings)
        {
            if(($addedSettings | Where EntityKey -eq $setting.appConfigKey)) { continue }

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $setting.appConfigKey
                Value = $setting.appConfigKeyValue
                EntityKey = $setting.appConfigKey
                Category = Get-LanguageString "TACSettings.generalSettings"
                SubCategory = Get-LanguageString "SettingDetails.additionalConfiguration"
            })
        }
    }
}

function Invoke-CDDocumentManagedAppConfig
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "SettingDetails.appConfiguration")
    
    $customApps,$publishedApps = Get-CDMobileApps $obj.Apps

    Add-BasicPropertyValue (Get-LanguageString "Inputs.enrollmentTypeLabel") (Get-LanguageString "EnrollmentType.devicesWithoutEnrollment")
    Add-BasicPropertyValue (Get-LanguageString "SettingDetails.publicApps") ($publishedApps -join  $script:objectSeparator)
    Add-BasicPropertyValue (Get-LanguageString "SettingDetails.customApps") ($customApps -join  $script:objectSeparator)

    Add-BasicAdditionalValues $obj $objectType

    $addedSettings = @()

    $appSettings = [PSCustomObject]@{ }
    foreach($setting in $obj.customSettings)
    {
        $appSettings | Add-Member Noteproperty -Name $setting.name -Value $setting.value -Force
    }

    if(($obj.Apps | Where { $_.mobileAppIdentifier.packageId -eq "com.microsoft.office.outlook" }))
    {
        if([IO.File]::Exists(($global:AppRootFolder + "\Documentation\ObjectInfo\#AppConfigOutlookApp.json")))
        {
            $jsonObj = Get-Content ($global:AppRootFolder + "\Documentation\ObjectInfo\#AppConfigOutlookApp.json") | ConvertFrom-Json
            Invoke-TranslateSection $appSettings $jsonObj
        }
    }

    if(($obj.Apps | Where { $_.mobileAppIdentifier.bundleId -like "com.microsoft.msedge" }))
    {
        if($appSettings.'com.microsoft.intune.mam.managedbrowser.bookmarks')
        {
            $appSettings.'com.microsoft.intune.mam.managedbrowser.bookmarks' = $appSettings.'com.microsoft.intune.mam.managedbrowser.bookmarks'.Replace("||",$script:objectSeparator).Replace("|",$script:propertySeparator)
        }

        if($appSettings.'com.microsoft.intune.mam.managedbrowser.AllowListURLs')
        {
            $appSettings.'com.microsoft.intune.mam.managedbrowser.AllowListURLs' = $appSettings.'com.microsoft.intune.mam.managedbrowser.AllowListURLs'.Replace("|",$script:objectSeparator)
        }

        if($appSettings.'com.microsoft.intune.mam.managedbrowser.BlockListURLs')
        {
            $appSettings.'com.microsoft.intune.mam.managedbrowser.BlockListURLs' = $appSettings.'com.microsoft.intune.mam.managedbrowser.BlockListURLs'.Replace("|",$script:objectSeparator)
        }

        if([IO.File]::Exists(($global:AppRootFolder + "\Documentation\ObjectInfo\#AppConfigEdgeApp.json")))
        {
            $jsonObj = Get-Content ($global:AppRootFolder + "\Documentation\ObjectInfo\#AppConfigEdgeApp.json") | ConvertFrom-Json
            Invoke-TranslateSection $appSettings $jsonObj
        }
    }

    $addedSettings = Get-DocumentedSettings

    $category = Get-LanguageString "TACSettings.generalSettings" 

    foreach($setting in $obj.customSettings)
    {
        if(($addedSettings | Where EntityKey -eq $setting.name)) { continue }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $setting.name
            Value = $setting.value
            EntityKey = $setting.name
            Category = $category
        })
    }       
}

# Document Named locations
function Invoke-CDDocumentCountryNamedLocation
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "AzureCA.menuItemNamedNetworks")
    Add-BasicAdditionalValues $obj $objectType
    
    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "AzureCA.NamedLocation.Form.CountryLookup.ariaLabel"
        Value = Get-LanguageString "AzureCA.NamedLocation.Form.CountryLookup.$((?: ($obj.countryLookupMethod -eq "clientIpAddress") "ip" "gps"))"
        EntityKey = "countryLookupMethod"
    })

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "AzureCA.NamedLocation.Form.Include.label"
        Value = Get-LanguageString (?: ($obj.includeUnknownCountriesAndRegions -eq $true) "Inputs.enabled" "Inputs.disabled")
        EntityKey = "includeUnknownCountriesAndRegions"
    })        

    $countryList = @()
    foreach($country in $obj.countriesAndRegions)
    {
        $countryList += Get-LanguageString "CountryNames.countryName$($country.ToLower())"
    }

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "AzureCA.NamedLocation.Type.countries"
        Value = $countryList -join $script:objectSeparator
        EntityKey = "countriesAndRegions"
    })         
}

function Invoke-CDDocumentIPNamedLocation
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "AzureCA.menuItemNamedNetworks")
    Add-BasicAdditionalValues $obj $objectType

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "AzureCA.NamedLocation.Form.Trusted.label"
        Value = Get-LanguageString (?: ($obj.isTrusted -eq $true) "Inputs.enabled" "Inputs.disabled")
        EntityKey = "isTrusted"
    })        

    $ipList = @()
    foreach($ip in $obj.ipRanges)
    {
        $ipList += $ip.cidrAddress
    }

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "AzureCA.NamedLocation.Type.ipRanges"
        Value = $ipList -join $script:objectSeparator
        EntityKey = "ipRanges"
    })         
}

# Document Terms of Use
function Invoke-CDDocumentTermsOfUse
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    $offLabel = Get-LanguageString "SettingDetails.offOption"
    $onLabel = Get-LanguageString "SettingDetails.onOption"

    ###################################################
    # Basic info
    ###################################################

    Add-BasicPropertyValue (Get-LanguageString "SettingDetails.nameName") $obj.displayName
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "AzureCA.menuItemTermsOfUse") 
        
    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "TermsOfUse.Wizard.agreementIsViewingBeforeAcceptanceRequiredLabel"
        Value = ?: $obj.isViewingBeforeAcceptanceRequired $onLabel $offLabel
        Category = $null
        SubCategory = $null
        EntityKey = "isViewingBeforeAcceptanceRequired"
    })

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "TermsOfUse.Wizard.agreementIsPerDeviceAcceptanceRequiredLabel"
        Value = ?: $obj.isPerDeviceAcceptanceRequired $onLabel $offLabel
        Category = $null
        SubCategory = $null
        EntityKey = "isPerDeviceAcceptanceRequired"
    })

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "TermsOfUse.Wizard.isAcceptanceExpirationEnabledLabel"
        Value = ?: $obj.termsExpiration $onLabel $offLabel
        Category = $null
        SubCategory = $null
        EntityKey = "isAcceptanceExpirationEnabledLabel"
    })
    
    if($obj.termsExpiration.startDateTime)
    {
        try
        {
            if($obj.termsExpiration.startDateTime -is [DateTime])
            {
                $tmpDate = $obj.termsExpiration.startDateTime
            }
            else
            {
                $tmpDate = ([DateTime]::Parse($obj.termsExpiration.startDateTime))
            }
            $tmpDateStr = ($tmpDate).ToShortDateString()
        }
        catch
        {
            Write-Log "Failed to parse date from string $($obj.termsExpiration.startDateTime)" 2
            $tmpDateStr = $obj.termsExpiration.startDateTime
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "TermsOfUse.Wizard.acceptanceExpirationStartDateTimeLabel"
            Value = $tmpDateStr
            Category = $null
            SubCategory = $null
            EntityKey = "startDateTime"
        })

        if($obj.termsExpiration.frequency -eq "P365D")
        {
            $value = Get-LanguageString "TermsOfUse.AcceptanceExpirationFrequency.annually"
        }
        elseif($obj.termsExpiration.frequency -eq "P180D")
        {
            $value = Get-LanguageString "TermsOfUse.AcceptanceExpirationFrequency.biannually"
        }
        elseif($obj.termsExpiration.frequency -eq "P30D")
        {
            $value = Get-LanguageString "TermsOfUse.AcceptanceExpirationFrequency.monthly"
        }
        elseif($obj.termsExpiration.frequency -eq "P90D")
        {
            $value = Get-LanguageString "TermsOfUse.AcceptanceExpirationFrequency.quarterly"
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "TermsOfUse.Wizard.acceptanceExpirationFrequencyLabel"
            Value = $value
            Category = $null
            SubCategory = $null
            EntityKey = "frequency"
        })        
    }
    if($null -ne $obj.userReacceptRequiredFrequency)
    {
        $days = Get-DurationValue $obj.userReacceptRequiredFrequency
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "TermsOfUse.Wizard.acceptanceDurationLabel"
            Value = $days
            Category = $null
            SubCategory = $null
            EntityKey = "userReacceptRequiredFrequency"
        })
    } 
}

# Document Conditional Access policy
function Invoke-CDDocumentConditionalAccess
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    #Add-BasicDefaultValues $obj $objectType
    Add-BasicPropertyValue (Get-LanguageString "SettingDetails.nameName") $obj.displayName
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "AzureCA.conditionalAccessBladeTitle") 

    if($obj.state -eq "enabledForReportingButNotEnforced")
    {
        $state = Get-LanguageString "AzureCA.PolicyState.reportOnly"
    }
    elseif($obj.state -eq "disabled")
    {
        $state = Get-LanguageString "AzureCA.PolicyState.off"
    }
    else
    {
        $state = Get-LanguageString "AzureCA.PolicyState.on"
    }

    Add-BasicPropertyValue (Get-LanguageString "AzureCA.policyEnforceLabel") $state

    Add-BasicAdditionalValues $obj $objectType

    $includeLabel = Get-LanguageString "AzureCA.userSelectionBladeIncludeTabTitle"
    $excludeLabel = Get-LanguageString "AzureCA.userSelectionBladeExcludeTabTitle"

    if($obj.conditions.clientApplications.includeServicePrincipals -or $obj.conditions.clientApplications.excludeServicePrincipals)
    {
        ###################################################
        # Workload
        ###################################################

        $ids = @()
        foreach($id in ($obj.conditions.clientApplications.includeServicePrincipals + $obj.conditions.clientApplications.excludeServicePrincipals))
        {
            if($id -in $ids) { continue }
            elseif($id -eq "ServicePrincipalsInMyTenant") { continue }
            
            $ids += $id
        }

        $category = Get-LanguageString "AzureCA.workloadIdentities"

        $idInfo = $null

        if($ids.Count -gt 0)
        {
            $ht = @{}
            $ht.Add("ids", @($ids | Unique))

            $body = $ht | ConvertTo-Json

            # ToDo: Get from MigFile for Offline
            $idInfo = (Invoke-GraphRequest -Url "/directoryObjects/getByIds?`$select=displayName,id" -Content $body -Method "Post").Value
        }
        
        if((($obj.conditions.clientApplications.includeServicePrincipals | Where { $_ -eq "ServicePrincipalsInMyTenant"}) -ne $null))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.servicePrincipalRadioAll"
                Category = $category
                SubCategory = $includeLabel
                EntityKey = "includeServicePrincipals"
            })        
        }
        elseif((($obj.conditions.clientApplications.includeServicePrincipals | Where { $_ -eq "None"}) -ne $null))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.chooseApplicationsNone"
                Category = $category
                SubCategory = $includeLabel
                EntityKey = "includeServicePrincipals"
            })        
        }
        elseif($ids.Count -gt 0 -and $obj.conditions.clientApplications.includeServicePrincipals)
        {
            #$category = Get-LanguageString "AzureCA.selectedSP"
            $tmpObjs = @() 
            foreach($id in ($obj.conditions.clientApplications.includeServicePrincipals))
            {
                $idObj = $idInfo | Where Id -eq $id
                $tmpObjs += ?? $idObj.displayName $id
            }
            
            if($tmpObjs.count -gt 0)
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = $category
                    Value = $tmpObjs -join $script:objectSeparator
                    Category = $category
                    SubCategory = $includeLabel
                    EntityKey = "includeServicePrincipals"
                })
            }
        } 

        if($obj.conditions.clientApplications.servicePrincipalFilter)
        {
            if($obj.conditions.clientApplications.servicePrincipalFilter.mode -eq "include") 
            {
                $filterMode = "included"
            }
            else
            {
                $filterMode = "excluded"
            }
    
            #AzureCA.PolicyBlade.Conditions.DeviceAttributes.AssignmentFilter.Blade
            #AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.title
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.AppliesTo.$filterMode"
                Value = $obj.conditions.clientApplications.servicePrincipalFilter.rule
                Category = $category
                SubCategory = Get-LanguageString "AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.title"
                EntityKey = "excludeServicePrincipalDevices"
            })           
        }         

        if((($obj.conditions.clientApplications.excludeServicePrincipals | Where { $_ -eq "ServicePrincipalsInMyTenant"}) -ne $null))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.servicePrincipalRadioAll"
                Category = $category
                SubCategory = $excludeLabel
                EntityKey = "excludeServicePrincipals"
            })        
        }
        elseif($ids.Count -gt 0)
        {
            #$category = Get-LanguageString "AzureCA.selectedSP"
            $tmpObjs = @() 
            foreach($id in ($obj.conditions.clientApplications.excludeServicePrincipals))
            {
                $idObj = $idInfo | Where Id -eq $id
                $tmpObjs += ?? $idObj.displayName $id
            }

            if($tmpObjs.count -gt 0)
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = $category
                    Value = $tmpObjs -join $script:objectSeparator
                    Category = $category
                    SubCategory = $excludeLabel
                    EntityKey = "excludeServicePrincipals"
                })
            }
        }
    }
    else
    {
        ###################################################
        # User and groups
        ###################################################

        $ids = @()
        foreach($id in ($obj.conditions.users.includeUsers + $obj.conditions.users.includeGroups + $obj.conditions.users.excludeUsers + $obj.conditions.users.excludeGroups))
        {
            if($id -in $ids) { continue }
            elseif($id -eq "GuestsOrExternalUsers") { continue }
            elseif($id -eq "All") { continue }
            elseif($id -eq "None") { continue }
            
            $ids += $id
        }

        $roleIds = @()
        foreach($id in ($obj.conditions.users.includeRoles + $obj.conditions.users.excludeRoles))
        {
            if($id -in $ids) { continue }
            $roleIds += $id
        }
        
        $idInfo = $null

        if($ids.Count -gt 0)
        {
            $ht = @{}
            $ht.Add("ids", @($ids | Unique))

            $body = $ht | ConvertTo-Json

            # ToDo: Get from MigFile for Offline
            $idInfo = (Invoke-GraphRequest -Url "/directoryObjects/getByIds?`$select=displayName,id" -Content $body -Method "Post").Value
        }

        if($roleIds.Count -gt 0 -and -not $script:allAadRoles)
        {
            $script:allAadRoles =(Invoke-GraphRequest -url "/directoryRoleTemplates?`$select=Id,displayName" -ODataMetadata "minimal").value
        }

        $category = Get-LanguageString "AzureCA.usersGroupsLabel"

        if((($obj.conditions.users.includeUsers | Where { $_ -eq "All"}) -ne $null))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.allUsersString"
                Category = $category
                SubCategory = $includeLabel
                EntityKey = "includeUsers"
            })        
        }
        elseif((($obj.conditions.users.includeUsers | Where { $_ -eq "None"}) -ne $null))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.chooseApplicationsNone"
                Category = $category
                SubCategory = $includeLabel
                EntityKey = "includeUsers"
            })        
        }
        else
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.userSelectionBladeSelectedUsers"
                Category = $category
                SubCategory = $includeLabel
                EntityKey = "includeUsers"
            })  

            if((($obj.conditions.users.includeUsers | Where { $_ -eq "GuestsOrExternalUsers"}) -ne $null))
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "AzureCA.allGuestUserLabel"
                    Value = Get-LanguageString "Inputs.enabled" #$((?: (($obj.conditions.users.includeUsers | Where { $_ -eq "GuestsOrExternalUsers"}) -ne $null) "enabled" "disabled"))"
                    Category = $category
                    SubCategory = $includeLabel
                    EntityKey = "includeGuestsOrExternalUsers"
                })
            }

            if($obj.conditions.users.includeRoles.Count -gt 0)
            {
                $tmpObjs = @() 
                foreach($id in $obj.conditions.users.includeRoles)
                {
                    $idObj = $script:allAadRoles | Where Id -eq $id
                    $tmpObjs += ?? $idObj.displayName $id
                }

                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "AzureCA.directoryRolesLabel"
                    Value = $tmpObjs -join $script:objectSeparator
                    Category = $category
                    SubCategory = $includeLabel
                    EntityKey = "includeRoles"
                })
            }

            if(($obj.conditions.users.includeUsers + $obj.conditions.users.includeGroups).Count -gt 0)
            {
                $tmpObjs = @() 
                foreach($id in ($obj.conditions.users.includeUsers + $obj.conditions.users.includeGroups))
                {
                    if($id -eq "GuestsOrExternalUsers") { continue }
                    $idObj = $idInfo | Where Id -eq $id
                    $tmpObjs += ?? $idObj.displayName $id
                }
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = $category
                    Value = $tmpObjs -join $script:objectSeparator
                    Category = $category
                    SubCategory = $includeLabel
                    EntityKey = "includeUsersGroups"
                })
            }
        }
        
        if((($obj.conditions.users.excludeUsers | Where { $_ -eq "GuestsOrExternalUsers"}) -ne $null))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "AzureCA.allGuestUserLabel"
                Value = Get-LanguageString "Inputs.enabled" #$((?: (($obj.conditions.users.excludeUsers | Where { $_ -eq "GuestsOrExternalUsers"}) -ne $null) "enabled" "disabled"))"
                Category = $category
                SubCategory = $excludeLabel
                EntityKey = "excludeGuestsOrExternalUsers"
            })
        }

        if($obj.conditions.users.excludeRoles.Count -gt 0)
        {
            $tmpObjs = @() 
            foreach($id in $obj.conditions.users.excludeRoles)
            {
                $idObj = $script:allAadRoles | Where Id -eq $id
                $tmpObjs += ?? $idObj.displayName $id
            }

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "AzureCA.directoryRolesLabel"
                Value = $tmpObjs -join $script:objectSeparator
                Category = $category
                SubCategory = $excludeLabel
                EntityKey = "excludeRoles"
            })
        }

        if(($obj.conditions.users.excludeUsers + $obj.conditions.users.excludeGroups).Count -gt 0)
        {
            $tmpObjs = @() 
            foreach($id in ($obj.conditions.users.excludeUsers + $obj.conditions.users.excludeGroups))
            {
                if($id -eq "GuestsOrExternalUsers") { continue }
                $idObj = $idInfo | Where Id -eq $id
                $tmpObjs += ?? $idObj.displayName $id
            }
            
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $category
                Value = $tmpObjs -join $script:objectSeparator
                Category = $category
                SubCategory = $excludeLabel
                EntityKey = "excludeUsersGroups"
            })
        }
    }

    ###################################################
    # Cloud apps or actions
    ###################################################

    $category = Get-LanguageString "AzureCA.UserActions.appsOrActionsTitle"
    $cloudAppsLabel = Get-LanguageString "AzureCA.policyCloudAppsLabel"    
    
    $cloudApps = Get-CDAllCloudApps
    
    if((($obj.conditions.applications.includeApplications | Where { $_ -eq "All"}) -ne $null))
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value = Get-LanguageString "AzureCA.cloudappsSelectionBladeAllCloudapps" #Get-LanguageString "Inputs.enabled"
            Category = $category
            SubCategory = $cloudAppsLabel
            EntityKey = "includeApplications"
        })        
    }
    elseif((($obj.conditions.applications.excludeApplications | Where { $_ -eq "None"}) -ne $null))
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value = Get-LanguageString "AzureCA.chooseApplicationsNone" #Get-LanguageString "Inputs.enabled"
            Category = $category
            SubCategory = $cloudAppsLabel
            EntityKey = "includeApplications"
        })        
    }
    elseif($obj.conditions.applications.includeApplications.Count -gt 0)
    {
        $tmpObjs = @() 
        foreach($id in ($obj.conditions.applications.includeApplications))
        {
            $idObj = $cloudApps | Where AppId -eq $id
            $tmpObjs += ?? $idObj.displayName $id
        }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value = $tmpObjs -join $script:objectSeparator
            Category = $category
            SubCategory = $cloudAppsLabel 
            EntityKey = "includeApplications"
        })        
    }    

    if($obj.conditions.applications.excludeApplications.Count -gt 0)
    {
        $tmpObjs = @() 
        foreach($id in ($obj.conditions.applications.excludeApplications))
        {
            $idObj = $cloudApps | Where AppId -eq $id
            $tmpObjs += ?? $idObj.displayName $id
        }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $excludeLabel
            Value = $tmpObjs -join $script:objectSeparator
            Category = $category
            SubCategory = $cloudAppsLabel 
            EntityKey = "excludeApplications"
        })        
    }  

    if($obj.conditions.applications.includeUserActions.Count -gt 0)
    {
        $userActionsLabel = Get-LanguageString "AzureCA.UserActions.label"
        if(($obj.conditions.applications.includeUserActions | Where { $_ -eq "urn:user:registersecurityinfo" }))
        {
            $value =  Get-LanguageString "AzureCA.UserActions.registerSecurityInfo"
        }
        else
        {
            $value =  Get-LanguageString "AzureCA.UserActions.registerOrJoinDevices"
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.UserActions.selectionInfo"
            Value =  $value
            Category = $category
            SubCategory = $userActionsLabel
            EntityKey = "includeUserActions"
        })           
    }

    if($obj.conditions.applications.includeAuthenticationContextClassReferences.Count -gt 0)
    {
        $tmpObjs = @() 
        if(-not $script:allAuthenticationContextClasses)
        {
            $script:allAuthenticationContextClasses = (Invoke-GraphRequest -url "/identity/conditionalAccess/authenticationContextClassReferences" -ODataMetadata "minimal").value 
        }

        foreach($id in ($obj.conditions.applications.includeAuthenticationContextClassReferences))
        {
            $idObj = $script:allAuthenticationContextClasses | Where Id -eq $id
            $tmpObjs += ?? $idObj.displayName $id
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.AuthContext.checkBoxInfo"
            Value =  $tmpObjs -join $script:objectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.AuthContext.label"
            EntityKey = "includeAuthenticationContextClassReferences"
        })           
    }

    ###################################################
    # Conditions
    ###################################################

    $category = Get-LanguageString "AzureCA.helpConditionsTitle"

    #$category = Get-LanguageString "AzureCA.policyConditionUserRisk"

    if($obj.conditions.userRiskLevels.Count -gt 0)
    {
        $tmpObjs = @() 
        foreach($id in ($obj.conditions.userRiskLevels))
        {
            $tmpObjs += Get-LanguageString "AzureCA.$($id)Risk"
        }        

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $script:objectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.policyConditionUserRisk"
            EntityKey = "userRiskLevels"
        })           
    }

    if($obj.conditions.signInRiskLevels.Count -gt 0)
    {
        $tmpObjs = @() 
        foreach($id in ($obj.conditions.signInRiskLevels))
        {
            $tmpObjs += Get-LanguageString "AzureCA.$($id)Risk"
        }        

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $script:objectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.policyConditionSigninRisk"
            EntityKey = "signInRiskLevels"
        })           
    }
    
    if($obj.conditions.platforms.includePlatforms.Count -gt 0)
    {
        $tmpObjs = @() 
        foreach($id in ($obj.conditions.platforms.includePlatforms))
        {
            if($id -eq "all")
            {
                $tmpObjs += Get-LanguageString "AzureCA.allDevicePlatforms"
            }
            else
            {
                $tmpObjs += Get-LanguageString "AzureCA.$($id)DisplayName"
            }
        }        

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $script:objectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.devicePlatform"
            EntityKey = "includePlatforms"
        })           
    }
    
    if($obj.conditions.platforms.excludePlatforms.Count -gt 0)
    {
        $tmpObjs = @() 
        foreach($id in ($obj.conditions.platforms.excludePlatforms))
        {
            $tmpObjs += Get-LanguageString "AzureCA.$($id)DisplayName"
        }        

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $excludeLabel
            Value =  $tmpObjs -join $script:objectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.devicePlatform"
            EntityKey = "excludePlatforms"
        })           
    }
    
    if(-not $script:allNamedLocations -and ($obj.conditions.locations.includeLocations.Count -gt 0 -or $obj.conditions.locations.excludeLocations.Count))
    {
        $script:allNamedLocations = Get-DocOfflineObjects "NamedLocations"
        if(-not $script:allNamedLocations)
        {
            # Might be better to get them one by one
            $script:allNamedLocations = (Invoke-GraphRequest -url "/identity/conditionalAccess/namedLocations?`$select=displayName,Id&top=999" -ODataMetadata "minimal").value
        }
        if(-not $script:allNamedLocations) {  $script:allNamedLocations = @()}
        elseif($script:allNamedLocations -isnot [Object[]]) {  $script:allNamedLocations = @($script:allNamedLocations) }

        $script:allNamedLocations += [PSCustomObject]@{
            displayName = Get-LanguageString "AzureCA.chooseLocationTrustedIpsItem"
            id =  "00000000-0000-0000-0000-000000000000"
        }
    }

    if(-not $script:allTermsOfUse -and (($obj.grantControls.termsOfUse | measure).Count -gt 0))
    {
        $script:allTermsOfUse = Get-DocOfflineObjects "TermsOfUse"
        if(-not $script:allTermsOfUse)
        {
            $script:allTermsOfUse  = (Invoke-GraphRequest -url "/identityGovernance/termsOfUse/agreements?`$select=displayName,Id&top=999" -ODataMetadata "minimal").value
        }
        if(-not $script:allTermsOfUse ) {  $script:allTermsOfUse  = @()}
        elseif($script:allTermsOfUse  -isnot [Object[]]) {  $script:allTermsOfUse  = @($script:allTermsOfUse ) }
    }

    <#
    if(-not $script:allAuthenticationStrength -and (($obj.grantControls.authenticationStrength | measure).Count -gt 0))
    {
        $script:allAuthenticationStrength = Get-DocOfflineObjects "AuthenticationStrengths"
        if(-not $script:allAuthenticationStrength)
        {
            $script:allAuthenticationStrength  = (Invoke-GraphRequest -url "/identity/conditionalAccess/authenticationStrengths/policies?`$select=displayName,Id" -ODataMetadata "minimal").value
        }
        if(-not $script:allAuthenticationStrength ) {  $script:allAuthenticationStrength  = @()}
        elseif($script:allAuthenticationStrength  -isnot [Object[]]) {  $script:allAuthenticationStrength  = @($script:allAuthenticationStrength ) }
    } 
    #>   

    if($obj.conditions.locations.includeLocations.Count -gt 0)
    {
        $tmpObjs = @() 
        foreach($id in ($obj.conditions.locations.includeLocations))
        {
            if($id -eq "AllTrusted")
            {
                $tmpObjs += Get-LanguageString "AzureCA.allTrustedLocationLabel"
            }
            elseif($id -eq "All")
            {
                $tmpObjs += Get-LanguageString "AzureCA.locationsAllLocationsLabel"
            }
            else
            {
                $idObj = $script:allNamedLocations | Where Id -eq $id
                $tmpObjs += ?? $idObj.displayName $id
            }
        }        

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $script:objectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.policyConditionLocation"
            EntityKey = "includeLocations"
        })           
    }
    
    if($obj.conditions.locations.excludeLocations.Count -gt 0)
    {
        $tmpObjs = @() 
        foreach($id in ($obj.conditions.locations.excludeLocations))
        {
            if($id -eq "AllTrusted")
            {
                $tmpObjs += Get-LanguageString "AzureCA.allTrustedLocationLabel"
            }
            elseif($id -eq "All")
            {
                $tmpObjs += Get-LanguageString "AzureCA.locationsAllLocationsLabel"
            }
            else
            {
                $idObj = $script:allNamedLocations | Where Id -eq $id
                $tmpObjs += ?? $idObj.displayName $id
            }
        }        

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $excludeLabel
            Value =  $tmpObjs -join $script:objectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.policyConditionLocation"
            EntityKey = "excludeLocations"
        })           
    }
    
    if($obj.conditions.clientAppTypes.Count -gt 0)
    {
        $tmpObjs = @() 
        foreach($id in ($obj.conditions.clientAppTypes))
        {
            if($id -eq "browser") { $tmpObjs += Get-LanguageString "AzureCA.clientAppWebBrowser" }
            elseif($id -eq "mobileAppsAndDesktopClients") { $tmpObjs += Get-LanguageString "AzureCA.clientAppMobileDesktop" }
            elseif($id -eq "exchangeActiveSync") { $tmpObjs += Get-LanguageString "AzureCA.clientAppExchangeActiveSync" }
            elseif($id -eq "other") { $tmpObjs += Get-LanguageString "AzureCA.clientTypeOtherClients" }
            elseif($id -eq "all") { break } # Not configured
            else
            {
                $tmpObjs += $id
                Write-Log "Unsupported app type: $id" 3
            }
        }        

        if($tmpObjs.Count -gt 0)
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value =  $tmpObjs -join $script:objectSeparator
                Category = $category
                SubCategory = Get-LanguageString "AzureCA.policyConditioniClientApp"
                EntityKey = "clientAppTypes"
            })
        }           
    }

    if($obj.conditions.devices.includeDevices.Count -gt 0)
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  Get-LanguageString "AzureCA.deviceStateAll"
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.deviceStateConditionSelectorLabel"
            EntityKey = "includeDevices"
        })           
    }

    if($obj.conditions.devices.excludeDevices.Count -gt 0)
    {
        $tmpObjs = @() 
        foreach($id in ($obj.conditions.devices.excludeDevices))
        {
            $tmpObjs += Get-LanguageString "AzureCA.classicPolicyControlRequire$($id)Device"
        }        

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $excludeLabel
            Value =  $tmpObjs -join $script:objectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.deviceStateConditionSelectorLabel"
            EntityKey = "excludeDevices"
        })           
    }

    if($obj.conditions.devices.deviceFilter)
    {
        if($obj.conditions.devices.deviceFilter.mode -eq "include") 
        {
            $filterMode = "included"
        }
        else
        {
            $filterMode = "excluded"
        }

        #AzureCA.PolicyBlade.Conditions.DeviceAttributes.AssignmentFilter.Blade
        #AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.title
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.AppliesTo.$filterMode"
            Value = $obj.conditions.devices.deviceFilter.rule
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.title"
            EntityKey = "includeDevices"
        })           
    }    
    
    ###################################################
    # Grant
    ###################################################

    $category = Get-LanguageString "AzureCA.policyControlBladeTitle"

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "AzureCA.policyControlContentDescription"
        Value =  Get-LanguageString "AzureCA.$((?: (($obj.grantControls.builtInControls | Where { $_ -eq "block"}) -ne $null) "policyControlBlockAccessDisplayedName" "policyControlAllowAccessDisplayedName"))"
        Category = $category
        SubCategory = ""
        EntityKey = "policyControl"
    })

    if($null -eq (($obj.grantControls.builtInControls | Where { $_ -eq "block"}) ))
    {
        if(($obj.grantControls.builtInControls | measure).Count -gt 0)
        {
            if(($obj.grantControls.builtInControls | Where { $_ -eq "mfa"}))
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "AzureCA.policyControlMfaChallengeDisplayedName"
                    Value =   Get-LanguageString "Inputs.enabled"
                    Category = $category
                    SubCategory = ""
                    EntityKey = "mfa"
                })
            }

            if(($obj.grantControls.builtInControls | Where { $_ -eq "compliantDevice"}))
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "AzureCA.policyControlCompliantDeviceDisplayedName"
                    Value =   Get-LanguageString "Inputs.enabled"
                    Category = $category
                    SubCategory = ""
                    EntityKey = "compliantDevice"
                })
            }

            if(($obj.grantControls.builtInControls | Where { $_ -eq "domainJoinedDevice"}))
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "AzureCA.policyControlRequireDomainJoinedDisplayedName"
                    Value =   Get-LanguageString "Inputs.enabled"
                    Category = $category
                    SubCategory = ""
                    EntityKey = "domainJoinedDevice"
                })
            }
            
            if(($obj.grantControls.builtInControls | Where { $_ -eq "approvedApplication"}))
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "AzureCA.policyControlRequireMamDisplayedName"
                    Value =   Get-LanguageString "Inputs.enabled"
                    Category = $category
                    SubCategory = ""
                    EntityKey = "approvedApplication"
                })
            }
            
            if(($obj.grantControls.builtInControls | Where { $_ -eq "compliantApplication"}))
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "AzureCA.policyControlRequireCompliantAppDisplayedName"
                    Value =   Get-LanguageString "Inputs.enabled"
                    Category = $category
                    SubCategory = ""
                    EntityKey = "compliantApplication"
                })
            }

            if(($obj.grantControls.builtInControls | Where { $_ -eq "passwordChange"}))
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "AzureCA.policyControlRequiredPasswordChangeDisplayedName"
                    Value =   Get-LanguageString "Inputs.enabled"
                    Category = $category
                    SubCategory = ""
                    EntityKey = "passwordChange"
                })
            }
        }

        if(($obj.grantControls.termsOfUse | measure).Count -gt 0)
        {
            $termsOfUse = @()
            foreach($tmpId in $obj.grantControls.termsOfUse)
            {
                $touObj = $script:allTermsOfUse | Where Id -eq $tmpId
                $termsOfUse += ?? $touObj.displayName $tmpId
            }
    
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "AzureCA.menuItemTermsOfUse"
                Value =   $termsOfUse -join $script:objectSeparator
                Category = $category
                SubCategory = ""
                EntityKey = "termsOfUse"
            })            
        }

        if(($obj.grantControls.authenticationStrength | measure).Count -gt 0)
        {
            $authenticationStrngth = @()
            foreach($tmpId in $obj.grantControls.authenticationStrength)
            {
                $authenticationStrngth += ?? $obj.grantControls.authenticationStrength.displayName $tmpId
            }
    
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "AzureCA.WhatIfBlade.authenticationStrength"
                Value =   $authenticationStrngth -join $script:objectSeparator
                Category = $category
                SubCategory = ""
                EntityKey = "authenticationStrength"
            })            
        }
        
        if(($obj.grantControls.customAuthenticationFactors | measure).Count -gt 0)
        {    
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "AzureCA.menuItemClaimProviderControls"
                Value =   $obj.grantControls.customAuthenticationFactors -join $script:objectSeparator
                Category = $category
                SubCategory = ""
                EntityKey = "customAuthenticationFactors"
            })            
        }        
    
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.descriptionContentForControlsAndOr"
            Value =   Get-LanguageString "AzureCA.$((?: ($obj.grantControls.operator -eq "OR") "requireOneControlText" "requireAllControlsText"))" 
            Category = $category
            SubCategory = ""
            EntityKey = "grantOperator"
        }) 
    }       

    ###################################################
    # Session
    ###################################################

    $category = Get-LanguageString "AzureCA.sessionControlBladeTitle"

    if($obj.sessionControls.applicationEnforcedRestrictions.isEnabled -eq $true)
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.sessionControlsAppEnforcedLabel"
            Value = Get-LanguageString "Inputs.enabled"
            Category = $category
            SubCategory = ""
            EntityKey = "applicationEnforcedRestrictions"
        })
    }
    
    if($obj.sessionControls.cloudAppSecurity.isEnabled -eq $true)
    {
        if($obj.sessionControls.cloudAppSecurity.cloudAppSecurityType -eq "mcasConfigured") { $strId = "useCustomControls" }
        elseif($obj.sessionControls.cloudAppSecurity.cloudAppSecurityType -eq "monitorOnly") { $strId = "monitorOnly" }
        elseif($obj.sessionControls.cloudAppSecurity.cloudAppSecurityType -eq "blockDownloads") { $strId = "blockDownloads" }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.sessionControlsCasLabel"
            Value =  Get-LanguageString "AzureCA.CAS.BuiltinPolicy.Option.$strId"
            Category = $category
            SubCategory = ""
            EntityKey = "cloudAppSecurity"
        })
    }
    
    if($obj.sessionControls.signInFrequency.isEnabled -eq $true)
    {
        if($obj.sessionControls.signInFrequency.type -eq "hours")
        {
            if($obj.sessionControls.signInFrequency.value -gt 1)
            {
                $value = (Get-LanguageString "AzureCA.SessionLifetime.SignInFrequency.Option.Hour.plural") -f $obj.sessionControls.signInFrequency.value
            }
            else
            {
                $value = Get-LanguageString "AzureCA.SessionLifetime.SignInFrequency.Option.Hour.singular"
            }
        }
        elseif($obj.sessionControls.signInFrequency.type -eq "days")
        {
            if($obj.sessionControls.signInFrequency.value -gt 1)
            {
                $value = (Get-LanguageString "AzureCA.SessionLifetime.SignInFrequency.Option.Day.plural") -f $obj.sessionControls.signInFrequency.value
            }
            else
            {
                $value = Get-LanguageString "AzureCA.SessionLifetime.SignInFrequency.Option.Day.singular"
            }
        }
        else
        {
            $value = Get-LanguageString "AzureCA.SessionControls.SignInFrequency.everytime"
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.SessionLifetime.SignInFrequency.Option.label"
            Value =  $value
            Category = $category
            SubCategory = ""
            EntityKey = "SignInFrequency"
        })
    }

    if($null -ne $obj.sessionControls.continuousAccessEvaluation) 
    {
        if($obj.sessionControls.continuousAccessEvaluation.mode -eq "strictLocation")
        {
            $value = Get-LanguageString "AzureCA.SessionControls.Cae.strictLocation"
        }
        else
        {
            $value = Get-LanguageString "AzureCA.SessionControls.Cae.disable"
        }        

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.SessionControls.Cae.checkboxLabel"
            Value =  $value
            Category = $category
            SubCategory = ""
            EntityKey = "continuousAccessEvaluation"
        })        
    }    
    
    if($obj.sessionControls.persistentBrowser.isEnabled -eq $true)
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.SessionLifetime.PersistentBrowser.Option.label"
            Value =  Get-LanguageString "AzureCA.SessionLifetime.PersistentBrowser.Option.$($obj.sessionControls.persistentBrowser.mode)"
            Category = $category
            SubCategory = ""
            EntityKey = "persistentBrowser"
        })
    }    
}

#region Document Policy Sets
function Invoke-CDDocumentPolicySet
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "SettingDetails.appConfiguration")
    
    ###################################################
    # Settings
    ###################################################

    $addedSettings = @()

    $policySetSettings = (
        [PSCustomObject]@{
            Types = @(
                @('#microsoft.graph.mobileAppPolicySetItem','appTitle'),
                @('#microsoft.graph.targetedManagedAppConfigurationPolicySetItem','appConfigurationTitle'),
                @('#microsoft.graph.managedAppProtectionPolicySetItem','appProtectionTitle'),
                @('#microsoft.graph.iosLobAppProvisioningConfigurationPolicySetItem','iOSAppProvisioningTitle'))
            Category = (Get-LanguageString "PolicySet.appManagement")
        },
        [PSCustomObject]@{
            Types = @(
                @('#microsoft.graph.deviceConfigurationPolicySetItem','deviceConfigurationTitle'),
                @('#microsoft.graph.deviceCompliancePolicyPolicySetItem','deviceComplianceTitle'),
                @('#microsoft.graph.deviceManagementScriptPolicySetItem','powershellScriptTitle'))
            Category = (Get-LanguageString "PolicySet.deviceManagement")
        }, 
        [PSCustomObject]@{
            Types = @(
                @('#microsoft.graph.enrollmentRestrictionsConfigurationPolicySetItem','deviceTypeRestrictionTitle'),
                @('#microsoft.graph.windowsAutopilotDeploymentProfilePolicySetItem','windowsAutopilotDeploymentProfileTitle'),
                @('#microsoft.graph.windows10EnrollmentCompletionPageConfigurationPolicySetItem','enrollmentStatusSettingTitle'))
            Category = (Get-LanguageString "PolicySet.deviceEnrollment")
        }
    )

    foreach($policySettingType in $policySetSettings)
    {
        foreach($subType in $policySettingType.Types)
        {
            foreach($setting in ($obj.items | where '@OData.Type' -eq $subType[0]))
            {
                if($setting.status -eq "error")
                {
                    Write-Log "Skipping missing $($subType[0]) type with id $($setting.id). Error code: $($setting.errorCode)"
                    continue
                }

                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = $setting.displayName
                    Value = (Get-CDDocumentPolicySetValue $setting)
                    EntityKey = $setting.id
                    Category = $policySettingType.Category
                    SubCategory = (Get-LanguageString "PolicySet.$($subType[1])")
                })
            }
        }
    }
}

function Get-CDDocumentPolicySetValue
{
    param($policySetItem)

    if($policySetItem.'@OData.Type' -eq '#microsoft.graph.enrollmentRestrictionsConfigurationPolicySetItem' -or 
        $policySetItem.'@OData.Type' -eq '#microsoft.graph.windows10EnrollmentCompletionPageConfigurationPolicySetItem')
    {
        return $policySetItem.Priority
    }
    elseif($policySetItem.'@OData.Type' -eq '#microsoft.graph.windowsAutopilotDeploymentProfilePolicySetItem')
    {
        if($policySetItem.itemType -eq '#microsoft.graph.azureADWindowsAutopilotDeploymentProfile')
        {
            return (Get-LanguageString "Autopilot.DirectoryService.azureAD")
        }
        elseif($policySetItem.itemType -eq '#microsoft.graph.activeDirectoryWindowsAutopilotDeploymentProfile')
        {
            return (Get-LanguageString "Autopilot.DirectoryService.activeDirectoryAD")
        }
    }
    # ToDo: Add support for all PolicySet items 
}
#endregion

#region Custom Profile
function Invoke-CDDocumentCustomOMAUri
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    #Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "PolicyType.custom")

    $platformId = Get-ObjectPlatformFromType $obj
    Add-BasicPropertyValue (Get-LanguageString "Inputs.platformLabel") (Get-LanguageString "Platform.$platformId")

    ###################################################
    # Settings
    ###################################################

    $addedSettings = @()
    $category = Get-LanguageString "SettingDetails.customPolicyOMAURISettingsName"

    foreach($setting in $obj.omaSettings)
    {
        # Add the name of the OMA-URI setting
        Add-CustomSettingObject ([PSCustomObject]@{            
            Name = (Get-LanguageString "SettingDetails.nameName")
            Value =  $setting.displayName
            EntityKey = "displayName_$($setting.omaUri)"
            Category = $category
            SubCategory = $setting.displayName
        })

        # Add the description of the OMA-URI setting
        Add-CustomSettingObject ([PSCustomObject]@{            
            Name = (Get-LanguageString "TableHeaders.description")
            Value =  $setting.description
            EntityKey = "description_$($setting.omaUri)"
            Category = $category
            SubCategory = $setting.displayName
        })

        # Add the OMA-URI path of the OMA-URI setting
        Add-CustomSettingObject ([PSCustomObject]@{            
            Name = (Get-LanguageString "SettingDetails.oMAURIName")
            Value =  $setting.omaUri
            EntityKey = "omaUri_$($setting.omaUri)"
            Category = $category
            SubCategory = $setting.displayName
        })

        if($setting.'@OData.Type' -eq '#microsoft.graph.omaSettingString')
        {
            $value = (Get-LanguageString "SettingDetails.stringName")
        }
        elseif($setting.'@OData.Type' -eq '#microsoft.graph.omaSettingBase64')
        {
            $value = (Get-LanguageString "SettingDetails.base64Name")
        }
        elseif($setting.'@OData.Type' -eq '#microsoft.graph.omaSettingBoolean')
        {
            $value = (Get-LanguageString "SettingDetails.booleanName")
        }
        elseif($setting.'@OData.Type' -eq '#microsoft.graph.omaSettingDateTime')
        {
            $value = (Get-LanguageString "SettingDetails.dateTimeName")
        }
        elseif($setting.'@OData.Type' -eq '#microsoft.graph.omaSettingFloatingPoint')
        {
            $value = (Get-LanguageString "SettingDetails.floatingPointName")
        }
        elseif($setting.'@OData.Type' -eq '#microsoft.graph.omaSettingInteger')
        {
            $value = (Get-LanguageString "SettingDetails.integerName")
        }
        elseif($setting.'@OData.Type' -eq '#microsoft.graph.omaSettingStringXml')
        {
            $value = (Get-LanguageString "SettingDetails.stringXMLName")
        }
        else
        {
            $value = $null
        }

        if($value)
        {
            # Add the type of the OMA-URI setting
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString "SettingDetails.dataTypeName")
                Value =  $value
                EntityKey = "type_$($setting.omaUri)"
                Category = $category
                SubCategory = $setting.displayName
            })
        }

        $value = $setting.value
        # Add the type of the OMA-URI setting
        if($setting.isEncrypted -ne $true)
        {
            if($setting.'@OData.Type' -eq '#microsoft.graph.omaSettingStringXml')
            {
                $value = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($value))
            }

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString "SettingDetails.valueName")
                Value =  $value
                EntityKey = "value_$($setting.omaUri)"
                Category = $category
                SubCategory = $setting.displayName
            })
        }
        else # ToDo: Add check button
        {
            if($obj.'@ObjectFromFile' -ne $true)
            {
                $xmlValue = Invoke-GraphRequest -Url "/deviceManagement/deviceConfigurations/$($obj.Id)/getOmaSettingPlainTextValue(secretReferenceValueId='$($setting.secretReferenceValueId)')"
                $value = $xmlValue.Value
                if($value)
                {
                    Add-CustomSettingObject ([PSCustomObject]@{
                        Name = (Get-LanguageString "SettingDetails.valueName")
                        Value =  $value
                        EntityKey = "value_$($setting.omaUri)"
                        Category = $category
                        SubCategory = $setting.displayName
                    })
                }
            }
        }        
    }
}
#endregion

#region Notification
function Invoke-CDDocumentNotification
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "Titles.notifications")

    ###################################################
    # Settings
    ###################################################

    $category = Get-LanguageString "TableHeaders.settings"

    if($obj.brandingOptions)
    {
        $brandingOptions = $obj.brandingOptions.Split(',')
    }
    else
    {
        $brandingOptions = @()
    }

    foreach($brandingOption in @('includeCompanyLogo','includeCompanyName','includeContactInformation','includeCompanyPortalLink'))
    {
        if($brandingOption -eq 'includeCompanyLogo')
        {
            $label = (Get-LanguageString "NotificationMessage.companyLogo")
        }
        elseif($brandingOption -eq 'includeCompanyName')
        {
            $label = (Get-LanguageString "NotificationMessage.companyName")
        }
        elseif($brandingOption -eq 'includeContactInformation')
        {
            $label = (Get-LanguageString "NotificationMessage.companyContact")
        }
        elseif($brandingOption -eq 'includeCompanyPortalLink')
        {
            $label = (Get-LanguageString "NotificationMessage.iwLink")
        }

        if(($brandingOption -in $brandingOptions))
        {
            $value = Get-LanguageString "BooleanActions.enable"
        }
        else
        {
            $value = Get-LanguageString "BooleanActions.disable"
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $label
            Value =  $value
            EntityKey = $brandingOption
            Category = $category
            SubCategory = $null
        })
    }
    
    #$subCategory = Get-LanguageString "NotificationMessage.localeLabel"
    $subCategory = Get-LanguageString "NotificationMessage.listTitle"

    foreach($template in $obj.localizedNotificationMessages)
    {
        $first,$second = $template.locale.Split('-')
        $baseInfo = [cultureinfo]$first
        $lng = $baseInfo.EnglishName.ToLower()
        if($first -eq 'en')
        {
            if($second -eq "US")
            {
                $lng = ($lng + "US")
            }
            elseif($second -eq "GB")
            {
                $lng = ($lng + "UK")
            }
        }
        elseif($first -eq 'es')
        {
            if($second -eq "es")
            {
                $lng = ($lng + "Spain")
            }
            elseif($second -eq "mx")
            {
                $lng = ($lng + "Mexico")
            }
        }
        elseif($first -eq 'fr')
        {
            if($second -eq "ca")
            {
                $lng = ($lng + "Canada")
            }
            elseif($second -eq "fr")
            {
                $lng = ($lng + "France")
            }
        }
        elseif($first -eq 'pt')
        {
            if($second -eq "pt")
            {
                $lng = ($lng + "Portugal")
            }
            elseif($second -eq "br")
            {
                $lng = ($lng + "Brazil")
            }
        }
        elseif($first -eq 'zh')
        {
            if($second -eq "tw")
            {
                $lng = ($lng + "Traditional")
            }
            elseif($second -eq "cn")
            {
                $lng = ($lng + "Simplified")
            }
        }
        elseif($first -eq 'nb')
        {
            $lng = "norwegian"
        }        
       
        $label = Get-LanguageString "NotificationMessage.NotificationMessageTemplatesTab.$lng"

        if(-not $label) { continue }

        $value = $template.subject

        if($template.isDefault)
        {
            $value = ($value + $script:objectSeparator + (Get-LanguageString "NotificationMessage.isDefaultLocale") + ": " + (Get-LanguageString "SettingDetails.trueOption"))
        }

        $fullValue = ($value + $script:objectSeparator + $template.messageTemplate)

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $label
            Value =  $fullValue            
            EntityKey = $template.locale
            Category = $category
            SubCategory = $subCategory
        })        
    }
}
#endregion

#region
function Invoke-CDDocumentAssignmentFilter
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicAdditionalValues $obj $objectType

    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "Filters.filters")
    Add-BasicPropertyValue (Get-LanguageString "Inputs.platformLabel") (Get-LanguageString "Platform.$($obj.platform)")

    ###################################################
    # Settings
    ###################################################

    $label = Get-LanguageString "Filters.ruleSyntax"

    $category = Get-LanguageString "SettingDetails.rules"

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = $label
        Value =  $obj.rule
        EntityKey = "rule"
        Category = $category
        SubCategory = $null
    })
}
#endregion

#region Co-ManagementSettings
function Invoke-CDDocumentCoManagementSettings
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicAdditionalValues $obj $objectType
    
    # "Filters" is not in the translation file
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") ((Get-LanguageString "WindowsEnrollment.coManagementAuthorityTitle").Trim())
    Add-BasicPropertyValue (Get-LanguageString "Inputs.platformLabel") (Get-LanguageString "Platform.Windows10")

    ###################################################
    # Settings
    ###################################################

    $category = Get-LanguageString "TableHeaders.settings"
    $valueYes = Get-LanguageString "BooleanActions.yes"
    $valueNo = Get-LanguageString "SettingDetails.no"
    
    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "CoManagementAuthority.installAgent"
        Value = ?: ($obj.installConfigurationManagerAgent -eq $true) $valueYes $valueNo
        EntityKey = "managedDeviceAuthority"
        Category = $category
        SubCategory = $null
    })

    if(($obj.installConfigurationManagerAgent -eq $true))
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "CoManagementAuthority.commandLineArgs"
            Value = $obj.configurationManagerAgentCommandLineArgument
            EntityKey = "managedDeviceAuthority"
            Category = $category
            SubCategory = $null
        })
    }

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "CoManagementAuthority.managedDeviceOwnership"
        Value = ?: ($obj.managedDeviceAuthority -eq 1) $valueYes $valueNo
        EntityKey = "managedDeviceAuthority"
        Category = $category
        SubCategory = Get-LanguageString "CoManagementAuthority.advancedProperty"
    })


}
#endregion

#region Windows Kiosk
function Invoke-CDDocumentWindowsKioskConfiguration
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicAdditionalValues $obj $objectType
    # "Filters" is not in the translation file
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "Category.kioskConfigurationV2")
    Add-BasicPropertyValue (Get-LanguageString "Inputs.platformLabel") (Get-LanguageString "Platform.$($obj.platform)")

    ###################################################
    # Settings
    ################################################### 
    
    $category = Get-LanguageString "Category.kiosk"

    if($obj.kioskProfiles[0].appConfiguration."@odata.type" -eq "#microsoft.graph.windowsKioskSingleWin32App" -or
        $obj.kioskProfiles[0].appConfiguration."@odata.type" -eq "#microsoft.graph.windowsKioskSingleUWPApp")
    {
        $kisokModeType = "single"
        $kioskMode = Get-LanguageString "SettingDetails.kioskSelectionSingleMode"
    }
    else
    {
        $kisokModeType = "multi"
        $kioskMode = Get-LanguageString "SettingDetails.kioskSelectionMultiMode"
    }

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "SettingDetails.kioskSelectionName"
        Value = $kioskMode
        EntityKey = "kioskMode"
        Category = $category
        SubCategory = $null
    })
    
    <#
    if($kisokModeType -eq "multi")
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "SettingDetails.kioskTargetSModeName"
            Value = $kioskMode
            EntityKey = "kioskMode"
            Category = $category
            SubCategory = $null
        })        
    }
    #>

    $logonTypeLngId = switch($obj.kioskProfiles[0].userAccountsConfiguration."@odata.type")
    {
        "#microsoft.graph.windowsKioskAutologon" { "kioskUserLogonTypeAutologon" }
        "#microsoft.graph.windowsKioskAzureADUser" { "kioskAADUserAndGroup" }
        "#microsoft.graph.windowsKioskAzureADGroup" { "kioskAADUserAndGroup" }
        "#microsoft.graph.windowsKioskLocalUser" { "kioskAppTypeStore" }
        "#microsoft.graph.windowsKioskVisitor" { "kioskVisitor" }
    }

    if($logonTypeLngId)
    {
        $logonType = Get-LanguageString "SettingDetails.$($logonTypeLngId)"
    }
    else
    {
        Write-Log "Unknown kiosk user logon type. $($obj.kioskProfiles[0].userAccountsConfiguration."@odata.type")" 2
    }

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "SettingDetails.kioskSelectionUsers"
        Value = $logonType
        EntityKey = "userAccountsConfigurationType"
        Category = $category
        SubCategory = $null
    })
    
    if($logonTypeLngId -eq "kioskAADUserAndGroup")
    {
        $users = @()
        $obj.kioskProfiles[0].userAccountsConfiguration | ForEach-Object { 
            if($_."@odata.type" -eq "#microsoft.graph.windowsKioskAzureADUser")
            {
                $users += "$($_.userPrincipalName)$($script:propertySeparator )$((Get-LanguageString "SettingDetails.kioskAADUser"))"
            }
            else
            {
                $users += "$($_.displayName)$($script:propertySeparator )$((Get-LanguageString "SettingDetails.kioskAADGroup"))"
            }
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "SettingDetails.kioskUserAccountName"
            Value = $users -join $script:objectSeparator
            EntityKey = "userAccounts"
            Category = $category
            SubCategory = $null
        })
    }
    elseif($obj.kioskProfiles[0].userAccountsConfiguration."@odata.type" -eq "#microsoft.graph.windowsKioskLocalUser")
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "SettingDetails.kioskUserAccountName"
            Value = $obj.kioskProfiles[0].userAccountsConfiguration.userName
            EntityKey = "userName"
            Category = $category
            SubCategory = $null
        })
    }

    if($kisokModeType -eq "single")
    {
        if($obj.kioskProfiles[0].appConfiguration."@odata.type" -eq "#microsoft.graph.windowsKioskSingleWin32App")
        {
            $uwpAppType = "win32App" 
            $appType = Get-LanguageString "SettingDetails.selectWin32AppForEdge86"
        }
        elseif($obj.kioskProfiles[0].appConfiguration."@odata.type" = "#microsoft.graph.windowsKioskSingleUWPApp")
        {
            if($obj.kioskProfiles[0].appConfiguration.uwpApp.appUserModelId -like "Microsoft.MicrosoftEdge*")
            {
                $uwpAppType = "edge"
                $appType = Get-LanguageString "SettingDetails.selectMicrosoftEdgeApp"
            }
            elseif($obj.kioskProfiles[0].appConfiguration.uwpApp.appUserModelId -like "Microsoft.KioskBrowser*")
            {
                $uwpAppType = "kioskBrowser"
                $appType = Get-LanguageString "SettingDetails.selectKioskBrowserApp"        
            }
            else
            {
                $uwpAppType = "storeApp"
                $appType = Get-LanguageString "SettingDetails.selectStoreApp"        
            }
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "SettingDetails.kioskApplicationType"
            Value = $appType
            EntityKey = "kioskApplicationType"
            Category = $category
            SubCategory = $null
        })

        $edgeKioskModeType = (?: ($obj.kioskProfiles[0].appConfiguration.win32App.edgeKioskType -eq "publicBrowsing") (Get-LanguageString "SettingDetails.edgeKioskModeTypePublicBrowsingInPrivate") (Get-LanguageString "SettingDetails.edgeKioskModeTypeDigitalSignage"))
        if($uwpAppType -eq "win32App")
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.win32EdgeKioskUrl"
                Value = $obj.kioskProfiles[0].appConfiguration.win32App.edgeKiosk
                EntityKey = "edgeKiosk"
                Category = $category
                SubCategory = $null
            }) 

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.edgeKioskModeType"
                Value = $edgeKioskModeType
                EntityKey = "edgeKioskType"
                Category = $category
                SubCategory = $null
            }) 

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.edgeKioskResetAfterIdleTimeInMinutesName"
                Value = $obj.kioskProfiles[0].appConfiguration.win32App.edgeKioskIdleTimeoutMinutes
                EntityKey = "edgeKioskIdleTimeoutMinutes"
                Category = $category
                SubCategory = $null
            }) 
        }
        elseif($uwpAppType -eq "edge")
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.edgeKioskModeType"
                Value = $edgeKioskModeType
                EntityKey = "edgeKioskType"
                Category = $category
                SubCategory = $null
            }) 
        }
        elseif($uwpAppType -eq "kioskBrowser")
        {
            $show = Get-LanguageString "BooleanActions.show"
            $hide = Get-LanguageString "BooleanActions.hide"

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.win10KioskBrowserDefaultWebsiteName"
                Value = $obj.kioskBrowserDefaultUrl
                EntityKey = "kioskBrowserDefaultUrl"
                Category = $category
                SubCategory = $null
            })
            
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.win10KioskBrowserHomeButtonName"
                Value = (?: $obj.kioskBrowserEnableHomeButton $show $hide)
                EntityKey = "kioskBrowserEnableHomeButton"
                Category = $category
                SubCategory = $null
            })
            
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.win10KioskBrowserNavigationButtonName"
                Value = (?: $obj.kioskBrowserEnableNavigationButtons $show $hide)
                EntityKey = "kioskBrowserEnableNavigationButtons"
                Category = $category
                SubCategory = $null
            })

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.win10KioskBrowserEnableEndSessionButtonName"
                Value = (?: $obj.kioskBrowserEnableEndSessionButton $show $hide)
                EntityKey = "kioskBrowserEnableEndSessionButton"
                Category = $category
                SubCategory = $null
            })
            
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.edgeKioskResetAfterIdleTimeInMinutesName"
                Value = $obj.kioskBrowserRestartOnIdleTimeInMinutes
                EntityKey = "kioskBrowserRestartOnIdleTimeInMinutes"
                Category = $category
                SubCategory = $null
            })
            
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.win10AllowedWebsitesName"
                Value = $obj.kioskBrowserBlockedURLs -join $script:objectSeparator
                EntityKey = "kioskBrowserBlockedURLs"
                Category = $category
                SubCategory = $null
            })        
        }
        elseif($uwpAppType -eq "storeApp")
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.kioskModeAppStoreUrlOrManagedAppIdName"
                Value = $obj.kioskProfiles[0].appConfiguration.uwpApp.name
                EntityKey = "edgeKioskType"
                Category = $category
                SubCategory = $null
            }) 
        }    
    }
    
    if($kisokModeType -eq "multi")
    {
        $apps = @()
        foreach($app in $obj.kioskProfiles[0].appConfiguration.apps)
        {            
            $kioskTypeLngId = switch($app.appType)
            {
                "aumId" { "kioskAppTypeAUMID" }
                "desktop" { "kioskAppTypeDesktop" }
                "store" { "kioskAppTypeStore" }
                Default { "kioskAppTypeUnknown" }
            }

            $kioskTileLngId = switch($app.startLayoutTileSize)
            {
                "medium" { "kioskTileMedium" } 
                "small" { "kioskTileSmall" } 
                "wide" { "kioskTileWide" } 
                "large" { "kioskTileLarge" } 
            }            

            $apps += $app.Name + $script:propertySeparator + (Get-LanguageString "SettingDetails.$($kioskTypeLngId)") +
            $script:propertySeparator + (?: ($app.autoLaunch -eq $true) (Get-LanguageString "SettingDetails.yes") (Get-LanguageString "SettingDetails.no")) +
            $script:propertySeparator + (Get-LanguageString "SettingDetails.$($kioskTileLngId)")
        }

        if($apps.Count -gt 0)
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.kioskAppTableName"
                Value = ($apps -join $script:objectSeparator)
                EntityKey = "kioskApps"
                Category = $category
                SubCategory = $null
            })
        }        

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "SettingDetails.alternativeStartLayoutName"
            Value = (?: ($obj.kioskProfiles[0].appConfiguration.startMenuLayoutXml -ne $null) (Get-LanguageString "SettingDetails.yes") (Get-LanguageString "SettingDetails.no"))
            EntityKey = "alternativeStartLayout"
            Category = $category
            SubCategory = $null
        })

        if($obj.kioskProfiles[0].appConfiguration.startMenuLayoutXml -ne $null)
        {           
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.kioskStartMenuLayoutXmlName"
                Value = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.kioskProfiles[0].appConfiguration.startMenuLayoutXml))
                EntityKey = "startMenuLayoutXml"
                Category = $category
                SubCategory = $null
            })
        }          

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "SettingDetails.kioskShowTaskbarName"
            Value = (?: ($obj.kioskProfiles[0].appConfiguration.showTaskBar) (Get-LanguageString "BooleanActions.show") (Get-LanguageString "BooleanActions.hide"))
            EntityKey = "showTaskBar"
            Category = $category
            SubCategory = $null
        })

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "SettingDetails.win10KioskAccessDownloadsFolderName"
            Value = (?: ($obj.kioskProfiles[0].appConfiguration.allowAccessToDownloadsFolder) (Get-LanguageString "SettingDetails.yes") (Get-LanguageString "SettingDetails.no"))
            EntityKey = "allowAccessToDownloadsFolder"
            Category = $category
            SubCategory = $null
        })
    }

    if($obj.windowsKioskForceUpdateSchedule)
    {
        $forceUpdateSchedule = Get-LanguageString "BooleanActions.require"        
    }
    else
    {
        $forceUpdateSchedule = Get-LanguageString "BooleanActions.notConfigured"
    }

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "SettingDetails.kioskForceRestart"
        Value = $forceUpdateSchedule
        EntityKey = "windowsKioskForceUpdateSchedule"
        Category = $category
        SubCategory = $null
    })

    if($obj.windowsKioskForceUpdateSchedule)
    {        
        try
        {
            $startDateObj = Get-Date $obj.windowsKioskForceUpdateSchedule.startDateTime -ErrorAction Stop

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.kioskStartDateTime"
                Value = ($startDateObj.ToShortDateString() + $script:objectSeparator + $startDateObj.ToShortTimeString())
                EntityKey = "startDateTime"
                Category = $category
                SubCategory = $null
            })

            if($obj.windowsKioskForceUpdateSchedule.recurrence -eq "weekly")
            {
                $recurrenceType = "kioskWeekly"
            }
            elseif($obj.windowsKioskForceUpdateSchedule.recurrence -eq "monthly")
            {
                $recurrenceType = "kioskMonthly"
            }
            else
            {
                $recurrenceType = "kioskDaily"
            }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.kioskRecurrence"
                Value = Get-LanguageString "SettingDetails.$($recurrenceType)"
                EntityKey = "recurrence"
                Category = $category
                SubCategory = $null
            })
            
            if($obj.windowsKioskForceUpdateSchedule.recurrence -eq "weekly")
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "SettingDetails.dayOfWeek"
                    Value = Get-LanguageString "SettingDetails.$($obj.windowsKioskForceUpdateSchedule.dayofWeek)"
                    EntityKey = "dayofWeek"
                    Category = $category
                    SubCategory = $null
                })
            } 
            
            if($obj.windowsKioskForceUpdateSchedule.recurrence -eq "monthly")
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "SettingDetails.dayOfMonth"
                    Value = $obj.windowsKioskForceUpdateSchedule.dayofMonth
                    EntityKey = "dayofMonth"
                    Category = $category
                    SubCategory = $null
                })
            } 
        }
        catch
        {

        }
    }    

  

}
#endregion

#region
function Invoke-CDDocumentDeviceEnrollmentPlatformRestrictionConfiguration
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicAdditionalValues $obj $objectType
    # "Filters" is not in the translation file
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "Titles.deviceTypeEnrollmentRestrictions")
    
    if($obj.platformType -eq "androidForWork")
    {
        $lngId = "androidWorkProfile"
    }
    elseif($obj.platformType -eq "mac")
    {
        $lngId = "macOS"
    }
    elseif($obj.platformType -eq "ios")
    {
        $lngId = "iOS"
    }
    elseif($obj.platformType -eq "android")
    {
        $lngId = "android"
    }
    elseif($obj.platformType -eq "windows")
    {
        $lngId = "windows"
    }
    else
    {
        $lngId = $null
    }

    if($obj.'@OData.Type' -eq '#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration')
    {
        $platform = Get-LanguageString "AzureCA.classicPolicyAllPlatforms"
        $properties = @("androidForWorkRestriction","androidRestriction","iosRestriction","macRestriction","windowsRestriction")
        $policyType = "all"
    }
    else
    {
        $platform = Get-LanguageString "Platform.$($lngId)"
        $properties = @("platformRestriction")
        $policyType = "platform"
    }
    
    Add-BasicPropertyValue (Get-LanguageString "Inputs.platformLabel") $platForm

    $allowStr = Get-LanguageString  "BooleanActions.allow"
    $blockStr = Get-LanguageString  "BooleanActions.block" 
    $category = Get-LanguageString  "EnrollmentRestrictions.DeviceType.platformSettings" 
    $subCategory = $null
    $connotRestrictStr = Get-LanguageString "EnrollmentRestrictions.DeviceType.cannotRestrict"
    
    foreach($prop in $properties)
    {
        if($prop -eq "androidForWorkRestriction")
        {
            $typeId = "androidWorkProfile"
        }
        elseif($prop -eq "macRestriction")
        {
            $typeId = "macOS"
        }
        elseif($prop -eq "iosRestriction")
        {
            $typeId = "iOS"
        }
        elseif($prop -eq "androidRestriction")
        {
            $typeId = "android"
        }
        elseif($prop -eq "windowsRestriction")
        {
            $typeId = "windows"
        }
        else
        {
            $typeId = $lngId
        }

        $typeStr = Get-LanguageString "Platform.$($typeId)"

        if($typeId -eq "macOS")
        {
            $version = $connotRestrictStr
        }
        elseif($obj.$prop.osMinimumVersion -or $obj.$prop.osMaximumVersion)
        {
            $version = "{0}-{1}" -f $obj.$prop.osMinimumVersion,$obj.$prop.osMaximumVersion
        }
        else
        {
            $version = ""
        }

        #$blockedSkus = $obj.blockedSkus -join $script:propertySeparator

        if($policyType -eq "all")
        {
            $subCategory = $typeStr
        }

        if($typeId -eq "androidWorkProfile" -or $typeId -eq "andriod")
        {
            $blockedManufacturers = ($obj.$prop.blockedManufacturers -join $script:propertySeparator)
        }
        else
        {
            $blockedManufacturers = $connotRestrictStr
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "EnrollmentRestrictions.DeviceType.type"
            Value = $typeStr
            EntityKey = "platformType"
            Category = $category
            SubCategory = $subCategory
        })
        
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "EnrollmentRestrictions.DeviceType.platform"
            Value = (?: $obj.$prop.platformBlocked $blockStr $allowStr)
            EntityKey = "platformBlocked"
            Category = $category
            SubCategory = $subCategory
        })
        
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "EnrollmentRestrictions.DeviceType.versions"
            Value = $version
            EntityKey = "versions"
            Category = $category
            SubCategory = $subCategory
        })      
        
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "EnrollmentRestrictions.DeviceType.personal"
            Value = (?: $obj.$prop.personalDeviceEnrollmentBlocked $blockStr $allowStr)
            EntityKey = "platformBlocked"
            Category = $category
            SubCategory = $subCategory
        })        

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "EnrollmentRestrictions.DeviceType.deviceManufacturer"
            Value = $blockedManufacturers
            EntityKey = "platformBlocked"
            Category = $category
            SubCategory = $subCategory
        })        
    }

}
#endregion

#region 
function Invoke-CDDocumentDeviceAndAppManagementRoleDefinition
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    Add-BasicAdditionalValues $obj $objectType
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "RoleAssignment.rolesMenuTitle")

    $roleResources = (Invoke-GraphRequest -Url "/deviceManagement/resourceOperations").Value
    
    if(-not $roleResources)
    {
        Write-Log "Could not get resource information for Intune roles" 3
        return
    }

    $assignedActions = @()
    foreach($actionId in $obj.permissions[0].actions)
    {
        $actionResource = $roleResources | Where Id -eq $actionId

        if(-not $actionResource)
        {
            Write-Log "Could not find a permission resource with ID $actionId" 3
            continue 
        }
        $assignedActions += $actionResource
    }

    $category = Get-LanguageString "Titles.permissions"
    $subCategory = $null
    foreach($resourceName in (($assignedActions | Select resourceName -Unique | sort-object -property resourceName).resourceName)) #@{e={$_.rootproperties.rootname}}
    {
        $resourceActions = @()
        foreach($action in ($assignedActions | where resourceName -eq $resourceName))
        {
            $resourceId = $action.resource
            $resourceActions += $action.actionName
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $resourceName
            Value = ($resourceActions -join $script:objectSeparator)
            EntityKey = $resourceId
            Category = $category
            SubCategory = $subCategory
        })        
    }

    $category = Get-LanguageString TableHeaders.assignments
    foreach($roleAssignment in $obj.roleAssignments)
    {
        $assignmentInfo = (Invoke-GraphRequest -Url "/deviceManagement/roleAssignments('$($roleAssignment.id)')?`$expand=microsoft.graph.deviceAndAppManagementRoleAssignment/roleScopeTags" -ODataMetadata "Skip")
        if(-not $assignmentInfo)
        {
            Write-Log "Failed to get assignment info"
            continue
        }
        $ids = @()
        foreach($id in @($assignmentInfo.scopeMembers,$assignmentInfo.members))
        {
            if($ids -notcontains $id) { $ids += $id }
        }

        $content = @{"ids"=$ids } | ConvertTo-Json
        $idInfo = (Invoke-GraphRequest -Url "/directoryObjects/getByIds?`$select=displayName,id" -Content $content -Method POST).value

        $subCategory = $assignmentInfo.displayName
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "SettingDetails.nameName"
            Value = $assignmentInfo.displayName
            EntityKey = "displayName"
            Category = $category
            SubCategory = $subCategory
        })

        if($assignmentInfo.description)
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "SettingDetails.descriptionName"
                Value = $assignmentInfo.description
                EntityKey = "displayName"
                Category = $category
                SubCategory = $subCategory
            })
        }

        $admins = @()
        foreach($id in $assignmentInfo.members)
        {
            $objInfo = $idInfo | Where Id -eq $id
            $admins += (?: ($objInfo.displayName) ($objInfo.displayName) ($id))            
        }

        if($admins.Count -gt 0)
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "RoleAssignment.RoleAssignmentAdmin"
                Value = ($admins -join $script:objectSeparator)
                EntityKey = "members"
                Category = $category
                SubCategory = $subCategory
            })
        }

        $scopeMembers = @()
        foreach($id in $assignmentInfo.scopeMembers)
        {
            $objInfo = $idInfo | Where Id -eq $id
            $scopeMembers += (?: ($objInfo.displayName) ($objInfo.displayName) ($id))            
        }

        if($scopeMembers.Count -gt 0)
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "RoleAssignment.RoleAssignmentScope"
                Value = ($scopeMembers -join $script:objectSeparator)
                EntityKey = "scopeMembers"
                Category = $category
                SubCategory = $subCategory
            })
        }
        
        $scopeTags = @()
        foreach($scopeTag in $assignmentInfo.roleScopeTags)
        {
            $scopeTags += $scopeTag.displayName
        }

        if($scopeTags.Count -gt 0)
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "TableHeaders.scopeTags"
                Value = ($scopeTags -join $script:objectSeparator)
                EntityKey = "scopeTags"
                Category = $category
                SubCategory = $subCategory
            })
        }  
    }
}
#endregion

#region 
function Invoke-CDDocumentCustomObjectDocumented
{
    param($obj, $objType, $documentationInfo)
    
    if($obj.'@Odata.type' -eq '#microsoft.graph.windows10EndpointProtectionConfiguration')
    {
        # Skip adding Xbox Services and Windows Encryption if not configured
        # Not a very good way of doing this but they have values even if not configured 
        # so this will remove them from the documentation

        $customProperties = @()
        $customProperties += [PSCustomObject]@{
            CategoryLanguageID = "bitLocker"
            SkipProperties = @("startupAuthenticationTpm*")
        }

        $customProperties += [PSCustomObject]@{
            CategoryLanguageID = "xboxServices"
            SkipProperties = @()
        }

        foreach($customProp in $customProperties)
        {
            $categoryStr = Get-LanguageString "Category.$($customProp.CategoryLanguageID)"
            $categorySettings = $documentationInfo.Settings | Where Category -eq $categoryStr
            $custom = $false
            foreach($categorySetting in $categorySettings)
            {
                $skip = $false
                foreach($SkipProperty in $customProp.SkipProperties)
                {
                    if($categorySetting.EntityKey -like $SkipProperty)
                    {
                        $skip = $true
                        break
                    }
                }
                if($skip) { continue }
                if($null -ne $categorySetting.RawValue -and $categorySetting.RawValue -ne $categorySetting.DefaultValue)
                {   
                    $custom = $true
                    break
                }
            }
            #$categorySettings | ForEach-Object {if($_.RawValue -ne $null -and  
            #    $_.RawValue -ne $_.DefaultValue){$custom = $true}}
            if($custom -eq $false)
            {
                Write-Log "Remove category $categoryStr"
                $documentationInfo.Settings = $documentationInfo.Settings | Where Category -ne $categoryStr
            }
        }        
    }
}
#endregion

#region
function Invoke-CDDocumentTranslateSectionFile
{
    param($obj, $objectType, $fileInfo, $categoryObj)

    if($obj.'@OData.Type' -eq "#microsoft.graph.windows10CompliancePolicy" -and $fileInfo.BaseName -eq "customcompliance_compliancewindows10")
    {
        $category = Get-Category $categoryObj."$($fileInfo.BaseName)".category
               
        if($null -eq $obj.deviceCompliancePolicyScript)
        {
            $propValue = Get-LanguageString "BooleanActions.notConfigured"
        }
        else
        {
            $propValue = Get-LanguageString "BooleanActions.require"
        }        
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "SettingDetails.adminConfiguredComplianceSettingName"
            Value = $propValue 
            EntityKey = "deviceCompliancePolicyScript"
            Category = $category
            SubCategory = $null
        })

        if($obj.deviceCompliancePolicyScript)
        {
            if($null -eq $script:allCustomCompliancePolicies)
            {
                $script:allCustomCompliancePolicies = (Invoke-GraphRequest -url "/deviceManagement/deviceComplianceScripts?`$select=displayName,id" -ODataMetadata "minimal").value
            }

            $customScript = $script:allCustomCompliancePolicies | Where Id -eq $obj.deviceCompliancePolicyScript.deviceComplianceScriptId

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "CustomCompliance.FilePicker.scriptFileLabel"
                Value = $customScript.displayName
                EntityKey = "deviceComplianceScriptName"
                Category = $category
                SubCategory = $null
            })

            if($obj.deviceCompliancePolicyScript.rulesContent)
            {
                $propValue = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.deviceCompliancePolicyScript.rulesContent)) 
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = Get-LanguageString "CustomCompliance.UploadFile.jsonFileLabel"
                    Value = $propValue 
                    EntityKey = "jsonFileContent"
                    Category = $category
                    SubCategory = $null
                })
            }            
        }

        return $true
    }
    return $false
}
#endregion

#region
function Invoke-CDDocumentDeviceComplianceScript
{
    param($documentationObj)

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","
    
    ###################################################
    # Basic info
    ###################################################

    Add-BasicDefaultValues $obj $objectType
    if($obj.publisher)
    {
        Add-BasicPropertyValue (Get-LanguageString "SettingDetails.publisher") $obj.publisher
    }    
    Add-BasicAdditionalValues $obj $objectType
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "Titles.complianceScriptManagementPreview")

    $category = Get-LanguageString "TableHeaders.settings"

    $valueYes = Get-LanguageString "BooleanActions.yes"
    $valueNo = Get-LanguageString "SettingDetails.no"

    if($obj.detectionScriptContent)
    {
        $propValue = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.detectionScriptContent)) 
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "ProactiveRemediations.Create.Settings.DetectionScriptMultiLineTextBox.label"
            Value = $propValue 
            EntityKey = "detectionScriptContent"
            Category = $category
            SubCategory = $null
        })
    }

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "WindowsManagement.scriptContextLabel"
        Value = (?: ($obj.runAsAccount -eq "system")  $valueNo $valueYes)
        EntityKey = "runAsAccount"
        Category = $category
        SubCategory = $null
    })

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "WindowsManagement.enforceSignatureCheckLabel"
        Value = (?: ($obj.enforceSignatureCheck -eq $false)  $valueNo $valueYes)
        EntityKey = "enforceSignatureCheck"
        Category = $category
        SubCategory = $null
    })    

    Add-CustomSettingObject ([PSCustomObject]@{
        Name = Get-LanguageString "WindowsManagement.runAs64BitLabel"
        Value = (?: ($obj.runAs32Bit -eq $true)  $valueNo $valueYes)
        EntityKey = "runAs32Bit"
        Category = $category
        SubCategory = $null
    }) 
}
#endregion

#region Settings Catalog

function Invoke-CDDocumentPostSettingsCatalog
{
    param($obj, $objectType, $settingsData)

    if($obj.templateReference.TemplateId.StartsWith("19c8aa67-f286-4861-9aa0-f23541d31680"))
    {
        $reusableSettingsType = Get-GraphObjectType "ReusableSettings"
        if($reusableSettingsType)
        {
            foreach($setting in ($settingsData | Where SettingId -eq "vendor_msft_firewall_mdmstore_firewallrules_{firewallrulename}_remoteaddressdynamickeywords"))
            {
                $reusableSettings = Invoke-GraphRequest -Url "$($reusableSettingsType.API)/$($setting.RawValue)"
                if($reusableSettings.displayName)
                {
                    $setting.Value = $reusableSettings.displayName
                }
                else
                {
                    Write-Log "No Reusable Settings object found with ID $($setting.RawValue)" 2
                }
            }
        }
    }
}
#endregion

#region Scope Tags
function Invoke-CDDocumentScopeTag
{
    param($obj, $objectType)

    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","

    $groupIDs, $groupInfo, $filterIds,$filtersInfo = Get-ObjectAssignments $obj.Object

    $nameLabel = Get-LanguageString "Inputs.displayNameLabel"
    $descriptionLabel = Get-LanguageString "TableHeaders.description" 
    $assignmentsLabel = Get-LanguageString "TableHeaders.assignments"

    $scopeTagInfo = Get-TableObjects $obj.ObjectType
    
    if(-not $scopeTagInfo)
    {
        $scopeTagInfo = [PSCustomObject]@{
            TypeName = (Get-LanguageString "SettingDetails.scopeTags")
            ObjectType = $obj.ObjectType
            Properties = @($nameLabel, "id", $descriptionLable, $assignmentsLabel)
            Items = @()
        }
        Set-TableObjects $scopeTagInfo
    }

    $scopeTagInfo.Items += ([PSCustomObject]@{
        $nameLabel = $obj.displayName
        ID = $obj.Id
        $descriptionLabel = $obj.Description
        $assignmentsLabel = ($groupInfo.displayName -join $script:objectSeparator)
        Object = $documentationObj.Object
    })
}

#endregion