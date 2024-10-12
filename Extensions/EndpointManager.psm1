<#
.SYNOPSIS
Module for managing Intune objects

.DESCRIPTION
This module is for the Endpoint Manager/Intune View. It manages Export/Import/Copy of Intune objects

.NOTES
  Author:         Mikael Karlsson
#>
function Get-ModuleVersion
{
    '3.9.8'
}

function Invoke-InitializeModule
{
    #Add settings
    $global:appSettingSections += (New-Object PSObject -Property @{
        Title = "Endpoint Manager/Intune"
        Id = "EndpointManager"
        Values = @()
        Priority = 10
    })

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Application"
        Key = "EMAzureApp"
        Type = "List" 
        SelectedValuePath = "ClientId"
        ItemsSource = $global:MSGraphGlobalApps
        DefaultValue = ""
        SubPath = "EndpointManager"
    }) "EndpointManager"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Application Id"
        Key = "EMCustomAppId"
        Type = "String"
        DefaultValue = ""
        SubPath = "EndpointManager"
    }) "EndpointManager"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Redirect URL"
        Key = "EMCustomAppRedirect"
        Type = "String"
        DefaultValue = ""
        SubPath = "EndpointManager"
    }) "EndpointManager"
    
    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Tenant Id"
        Key = "EMCustomTenantId"
        Type = "String"
        DefaultValue = ""
        SubPath = "EndpointManager"
    }) "EndpointManager"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Authority"
        Key = "EMCustomAuthority"
        Type = "String"
        DefaultValue = ""
        SubPath = "EndpointManager"
    }) "EndpointManager"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "App packages folder"
        Key = "EMIntuneAppPackages"
        Type = "Folder"
        Description = "Root folder where intune app packages are located"
        SubPath = "EndpointManager"
    }) "EndpointManager"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Save Encryption File"
        Key = "EMSaveEncryptionFile"
        Type = "Boolean"
        Description = "Save encryption file when uploading an app. This can then be used to when downloading the app file."
        SubPath = "EndpointManager"
    }) "EndpointManager"    

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "App download folder"
        Key = "EMIntuneAppDownloadFolder"
        Type = "Folder"
        Description = "Folder where app packages will be downloaded and where encryption files will be saved"
        SubPath = "EndpointManager"
    }) "EndpointManager"

    Get-SettingValue "ProxyURI"

    if($global:FirstTimeRunning) {
        Save-Setting "EndpointManager" "EMAzureApp" $global:DefaultAzureApp
    }

    $currentAppID = Get-SettingValue "EMAzureApp"
    $customAppID = Get-SettingValue "EMCustomAppId"
    $global:informOldAzureApp = $false

    if(($global:OldAzureApps -is [Array] -and $currentAppID -in $global:OldAzureApps) -or (-not $currentAppID -and -not $customAppID))
    {
        $global:informOldAzureApp = $true
        Write-Log "Microsoft Intune PowerShell is being decomissioned. Please change to a supported app eg Microsoft Graph or a custom app!" 2            
    }

    $viewPanel = Get-XamlObject ($global:AppRootFolder + "\Xaml\EndpointManagerPanel.xaml") -AddVariables
    
    Set-EMViewPanel $viewPanel

    #Add menu group and items
    $global:EMViewObject = (New-Object PSObject -Property @{ 
        Title = "Intune Manager"
        Description = "Manages Intune environments. This view can be used for copying objects in an Intune environment. It can also be used for backing up an entire Intune environment and cloning the Intune environment into another tenant."
        ID="IntuneGraphAPI" 
        ViewPanel = $viewPanel
        AuthenticationID = "MSAL"
        ItemChanged = { Show-GraphObjects -ObjectTypeChanged; Invoke-ModuleFunction "Invoke-GraphObjectsChanged"; Write-Status ""}
        Deactivating = { Invoke-EMDeactivateView }
        Activating = { Invoke-EMActivatingView  }
        Authentication = (Get-MSALAuthenticationObject)
        Authenticate = { Invoke-EMAuthenticateToMSAL @args }
        AppInfo = (Get-GraphAppInfo "EMAzureApp" $global:DefaultAzureApp "EM")
        SaveSettings = { Invoke-EMSaveSettings }

        Permissions = @()
    })

    Add-ViewObject $global:EMViewObject

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Device Configuration"
        Id = "DeviceConfiguration"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/deviceConfigurations"
        QUERYLIST = "`$filter=not%20isof(%27microsoft.graph.windowsUpdateForBusinessConfiguration%27)%20and%20not%20isof(%27microsoft.graph.iosUpdateConfiguration%27)"
        #ExportFullObject = $false
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        PropertiesToRemove = @("privacyAccessControls")
        PostFileImportCommand = { Start-PostFileImportDeviceConfiguration @args }
        PostCopyCommand = { Start-PostCopyDeviceConfiguration @args }
        PostGetCommand = { Start-PostGetDeviceConfiguration @args }
        GroupId = "DeviceConfiguration"
        NavigationProperties=$true
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Conditional Access"
        Id = "ConditionalAccess"
        ViewID = "IntuneGraphAPI"
        API = "/identity/conditionalAccess/policies"
        Permissons=@("Policy.Read.All","Policy.ReadWrite.ConditionalAccess","Application.Read.All")
        Dependencies = @("NamedLocations","Applications","TermsOfUse","AuthenticationStrengths","AssignmentFilters")
        GroupId = "ConditionalAccess"
        ImportExtension = { Add-ConditionalAccessImportExtensions @args }
        PreImportCommand = { Start-PreImportConditionalAccess @args }
        PostExportCommand = { Start-PostExportConditionalAccess  @args }
        ExpandAssignmentsList = $false
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Terms of use"
        Id = "TermsOfUse"
        ViewID = "IntuneGraphAPI"
        ViewProperties = @("id", "displayName")
        Expand = "files"
        QUERYLIST = "`$expand=files"
        API = "/identityGovernance/termsOfUse/agreements"
        Permissons=@("Agreement.ReadWrite.All")
        PreImportCommand = { Start-PreImportTermsOfUse @args }
        PostExportCommand = { Start-PostExportTermsOfUse  @args }
        GroupId = "ConditionalAccess"        
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Named Locations"
        Id = "NamedLocations"
        ViewID = "IntuneGraphAPI"
        API = "/identity/conditionalAccess/namedLocations"
        Permissons=@("Policy.ReadWrite.ConditionalAccess")
        ImportOrder = 50
        GroupId = "ConditionalAccess"
        ExpandAssignmentsList = $false
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Endpoint Security"
        Id = "EndpointSecurity"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/intents"
        PropertiesToRemove = @('Settings','@OData.Type')        
        PreImportCommand = { Start-PreImportEndpointSecurity @args }
        PostListCommand = { Start-PostListEndpointSecurity @args }
        PostExportCommand = { Start-PostExportEndpointSecurity @args }
        PostFileImportCommand = { Start-PostFileImportEndpointSecurity @args }
        PostGetCommand = { Start-PostGetEndpointSecurity @args }
        #PreCopyCommand = { Start-PreCopyEndpointSecurity @args }
        PostCopyCommand = { Start-PostCopyEndpointSecurity @args }
        PreUpdateCommand = { Start-PreUpdateEndpointSecurity @args }
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Dependencies = @("ReusableSettings")
        GroupId = "EndpointSecurity"        
    })    

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Compliance Policies"
        Id = "CompliancePolicies"
        ViewID = "IntuneGraphAPI"
        Expand = "scheduledActionsForRule(`$expand=scheduledActionConfigurations)"
        API = "/deviceManagement/deviceCompliancePolicies"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Dependencies = @("Locations","Notifications","ComplianceScripts")
        PostExportCommand = { Start-PostExportCompliancePolicies @args }
        PreUpdateCommand = { Start-PreUpdateCompliancePolicies @args }
        GroupId = "CompliancePolicies"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Compliance Policies - V2"
        Id = "CompliancePoliciesV2"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/compliancePolicies"
        NameProperty = "Name"
        PropertiesToRemove = @('settingCount')
        ViewProperties = @("name","description","Id")
        Expand="settings"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        GroupId = "CompliancePolicies"
        Icon = "CompliancePolicies"
    })

    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Compliance Scripts"
        Id = "ComplianceScripts"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/deviceComplianceScripts"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        GroupId = "CompliancePolicies"
        Icon = "Scripts"
    })        

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Intune Branding"
        Id = "IntuneBranding"
        API = "/deviceManagement/intuneBrandingProfiles"
        ViewID = "IntuneGraphAPI"
        NameProperty = "profileName"
        ViewProperties = @("profileName", "displayName", "description", "id","isDefaultProfile")
        PreImportCommand = { Start-PreImportIntuneBranding @args }
        PostImportCommand = { Start-PostImportIntuneBranding @args }
        PostGetCommand = { Start-PostGetIntuneBranding @args }
        PostExportCommand = { Start-PostExportIntuneBranding  @args }
        PreDeleteCommand = { Start-PreDeleteIntuneBranding @args }
        PreUpdateCommand = { Start-PreUpdateIntuneBranding @args }
        Permissons=@("DeviceManagementApps.ReadWrite.All")
        Icon = "Branding"
        SkipRemoveProperties = @('Id') # Id is removed by PreImport. Required for default profile
        PropertiesToRemoveForUpdate = @('isDefaultProfile','disableClientTelemetry')
        GroupId = "TenantAdmin"
    })

    <#
    # BUG in Graph? Cannot create default branding. Can only create it when importing another object
    # Header required Accept-Language: sv-SE
    # Documentation says to use Content-Language but that doesn't work

    # Could work with https://main.iam.ad.ext.azure.com/api/LoginTenantBrandings
    
    #>

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Azure Branding"
        Id = "AzureBranding"
        API = "/organization/%OrganizationId%/branding/localizations"
        ViewID = "IntuneGraphAPI"
        ViewProperties = @("Id")
        PreImportCommand = { Start-PreImportAzureBranding  @args }
        PostListCommand = { Start-PostListAzureBranding @args }
        ShowButtons = @("Export","View")
        NameProperty = "Id"
        Permissons=@("Organization.ReadWrite.All")
        Icon = "Branding"
        SkipRemoveProperties = @('Id')
        GroupId = "Azure"
        SkipAddIDOnExport = $true
        ExpandAssignmentsList = $false
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Enrollment Status Page"
        Id = "EnrollmentStatusPage"
        API = "/deviceManagement/deviceEnrollmentConfigurations"
        ViewID = "IntuneGraphAPI"
        PreImportCommand = { Start-PreImportESP @args }
        PostExportCommand = { Start-PostExportESP @args }
        PreDeleteCommand = { Start-PreDeleteEnrollmentRestrictions @args } # Note: Uses same PreDelete as restrictions
        PreReplaceCommand = { Start-PreReplaceEnrollmentRestrictions @args } # Note: Uses same PreReplaceCommand as restrictions
        PostReplaceCommand = { Start-PostReplaceEnrollmentRestrictions @args } # Note: Uses same PostReplaceCommand as restrictions
        PreFilesImportCommand = { Start-PreFilesImportEnrollmentRestrictions @args } # Note: Uses same PreFilesImportCommand as restrictions
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsEnrollmentRestrictions @args } # Note: Uses same PreFilesImportCommand as restrictions
        PostListCommand = { Start-PostListESP @args }
        #PreUpdateCommand = { Start-PreUpdateEnrollmentRestrictions @args } # Note: Uses same PreUpdateCommand as restrictions
        #QUERYLIST = "`$filter=endsWith(id,'Windows10EnrollmentCompletionPageConfiguration')"
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        SkipRemoveProperties = @('Id')
        Dependencies = @("Applications")
        AssignmentsType = "enrollmentConfigurationAssignments"
        PropertiesToRemoveForUpdate = @('priority')
        GroupId = "WinEnrollment"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Enrollment Restrictions"
        Id = "EnrollmentRestrictions"
        API = "/deviceManagement/deviceEnrollmentConfigurations"
        ViewID = "IntuneGraphAPI"
        #QUERYLIST = "`$filter=not endsWith(id,'Windows10EnrollmentCompletionPageConfiguration')"
        PostExportCommand = { Start-PostExportEnrollmentRestrictions @args }
        PreImportCommand = { Start-PreImportEnrollmentRestrictions @args }
        PreDeleteCommand = { Start-PreDeleteEnrollmentRestrictions @args }
        PreReplaceCommand = { Start-PreReplaceEnrollmentRestrictions @args }
        PostReplaceCommand = { Start-PostReplaceEnrollmentRestrictions @args }
        PreFilesImportCommand = { Start-PreFilesImportEnrollmentRestrictions @args }
        PostListCommand = { Start-PostListEnrollmentRestrictions @args }
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsEnrollmentRestrictions @args }
        #PreUpdateCommand = { Start-PreUpdateEnrollmentRestrictions @args }
        PropertiesToRemoveForUpdate = @('priority')
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        SkipRemoveProperties = @('Id')
        AssignmentsType = "enrollmentConfigurationAssignments"
        GroupId = "EnrollmentRestrictions"
        ViewProperties = @("displayName","platformType","description","Id")
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Co-Management Settings"
        Id = "CoManagementSettings"
        ViewID = "IntuneGraphAPI"        
        API = "/deviceManagement/deviceEnrollmentConfigurations"
        PostReplaceCommand = { Start-PostReplaceEnrollmentRestrictions @args } # Note: Uses same PostReplaceCommand as restrictions
        PreFilesImportCommand = { Start-PreFilesImportEnrollmentRestrictions @args } # Note: Uses same PreFilesImportCommand as restrictions
        PostListCommand = { Start-PostListCoManagementSettings @args }
        PropertiesToRemoveForUpdate = @('priority')
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        SkipRemoveProperties = @('Id')        
        GroupId = "WinEnrollment"
        Icon = "EnrollmentStatusPage"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Administrative Templates"
        Id = "AdministrativeTemplates"
        API = "/deviceManagement/groupPolicyConfigurations"
        ViewID = "IntuneGraphAPI"
        PostGetCommand = { Start-PostGetAdministrativeTemplate @args }
        PostExportCommand = { Start-PostExportAdministrativeTemplate @args }
        PostCopyCommand = { Start-PostCopyAdministrativeTemplate @args }
        PostFileImportCommand = { Start-PostFileImportAdministrativeTemplate @args }
        PreImportCommand = { Start-PreImportAdministrativeTemplate @args }
        LoadObject = { Start-LoadAdministrativeTemplate @args }
        PropertiesToRemove = @("definitionValues","policyConfigurationIngestionType")
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Icon="DeviceConfiguration"
        GroupId = "DeviceConfiguration"
        CompareValue = "CombinedValueWithLabel"
        Dependencies = @("ADMXFiles")
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Scripts (PowerShell)"
        Id = "PowerShellScripts"
        API = "/deviceManagement/deviceManagementScripts"
        ViewID = "IntuneGraphAPI"
        DetailExtension = { Add-ScriptExtensions @args }
        ExportExtension = { Add-ScriptExportExtensions @args }
        PostExportCommand = { Start-PostExportScripts @args }
        Permissons=@("DeviceManagementManagedDevices.ReadWrite.All")
        AssignmentsType = "deviceManagementScriptAssignments"
        Icon="Scripts"
        GroupId = "Scripts"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Scripts (Shell)"
        Id = "MacScripts"
        API = "/deviceManagement/deviceShellScripts"
        ViewID = "IntuneGraphAPI"
        DetailExtension = { Add-ScriptExtensions @args }
        ExportExtension = { Add-ScriptExportExtensions @args }
        PostExportCommand = { Start-PostExportScripts @args }
        Permissons=@("DeviceManagementManagedDevices.ReadWrite.All")
        AssignmentsType = "deviceManagementScriptAssignments"
        Icon="Scripts"
        GroupId = "Scripts"
    })    

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Custom Attributes"
        Id = "MacCustomAttributes"
        API = "/deviceManagement/deviceCustomAttributeShellScripts"
        ViewID = "IntuneGraphAPI"
        Permissons=@("DeviceManagementManagedDevices.ReadWrite.All")
        AssignmentsType = "deviceManagementScriptAssignments"
        Icon="CustomAttributes"
        GroupId = "CustomAttributes" # MacOS Settings
        DetailExtension = { Add-ScriptExtensions @args }
        ExportExtension = { Add-ScriptExportExtensions @args }
        PostExportCommand = { Start-PostExportScripts @args }
        PropertiesToRemoveForUpdate = @('customAttributeName','customAttributeType','displayName')
        #PreUpdateCommand = { Start-PreUpdateMacCustomAttributes @args }
    })    

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Terms and Conditions"
        Id = "TermsAndConditions"
        API = "/deviceManagement/termsAndConditions"
        ViewID = "IntuneGraphAPI"
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        ExpandAssignments = $false # Not supported for this object type
        PostExportCommand = { Start-PostExportTermsAndConditions @args }
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsTermsAndConditions @args }
        GroupId = "TenantAdmin"
        ExpandAssignmentsList = $false
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "App Protection"
        Id = "AppProtection"
        API = "/deviceAppManagement/managedAppPolicies"
        ViewID = "IntuneGraphAPI"
        PreGetCommand = { Start-GetAppProtection @args }
        PostListCommand = { Start-PostListAppProtection @args }
        PreImportCommand = { Start-PreImportAppProtection @args }
        PostImportCommand = { Start-PostImportAppProtection  @args }
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsAppProtection @args }
        PreUpdateCommand = { Start-PreUpdateAppProtection  @args }
        ExportFullObject = $true
        PropertiesToRemove = @('exemptAppLockerFiles')
        PropertiesToRemoveForUpdate = @("protectedAppLockerFiles","version") # ToDo: !!! Add support for protectedAppLockerFiles?
        Permissons=@("DeviceManagementApps.ReadWrite.All")
        Dependencies = @("Applications")
        GroupId = "AppProtection"
        ExpandAssignmentsList = $false
    })

    # These are also included in the managedAppPolicies API
    # So all custom commands will be handled by the same functions as App Protection
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "App Configuration (App)"
        Id = "AppConfigurationManagedApp"
        API = "/deviceAppManagement/targetedManagedAppConfigurations"
        ViewID = "IntuneGraphAPI"
        PreGetCommand = { Start-GetAppProtection @args }
        PreImportCommand = { Start-PreImportAppProtection @args }
        PostImportCommand = { Start-PostImportAppProtection  @args }
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsAppProtection @args }
        PreUpdateCommand = { Start-PreUpdateAppConfigurationApp @args }
        Permissons=@("DeviceManagementApps.ReadWrite.All")
        Dependencies = @("Applications")
        Icon = "AppConfiguration"
        GroupId = "AppConfiguration"
        ExpandAssignmentsList = $false
    })    

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "App Configuration (Device)"
        Id = "AppConfigurationManagedDevice"
        API = "/deviceAppManagement/mobileAppConfigurations"
        QUERYLIST = "`$filter=microsoft.graph.androidManagedStoreAppConfiguration/appSupportsOemConfig%20eq%20false%20or%20isof(%27microsoft.graph.androidManagedStoreAppConfiguration%27)%20eq%20false"
        ViewID = "IntuneGraphAPI"
        Permissons=@("DeviceManagementApps.ReadWrite.All")
        Dependencies = @("Applications")
        PreFilesImportCommand = { Start-PreFilesImportAppConfiguration @args }
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsAppConfiguration @args }
        PostExportCommand = { Start-PostExportAppConfiguration @args }
        Icon = "AppConfiguration"
        GroupId = "AppConfiguration"
    })
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Applications"
        Id = "Applications"
        API = "/deviceAppManagement/mobileApps"
        ViewID = "IntuneGraphAPI"
        PropertiesToRemove = @('uploadState','publishingState','isAssigned','dependentAppCount','supersedingAppCount','supersededAppCount','committedContentVersion','isFeatured','size','categories') #,'minimumSupportedWindowsRelease'
        QUERYLIST = "`$filter=(microsoft.graph.managedApp/appAvailability%20eq%20null%20or%20microsoft.graph.managedApp/appAvailability%20eq%20%27lineOfBusiness%27%20or%20isAssigned%20eq%20true)&`$orderby=displayName"
        QuerySearch=$true
        Permissons=@("DeviceManagementApps.ReadWrite.All")
        AssignmentsType="mobileAppAssignments"
        AssignmentProperties = @("@odata.type","target","settings","intent")
        AssignmentTargetProperties = @("@odata.type","groupId","deviceAndAppManagementAssignmentFilterId","deviceAndAppManagementAssignmentFilterType")
        ImportOrder = 60
        Expand="categories,assignments" # ODataMetadata is set to minimal so assignments can't be autodetected
        ODataMetadata="minimal" # categories property not supported with ODataMetadata full
        PostFileImportCommand = { Start-PostFileImportApplications @args }
        PostCopyCommand = { Start-PostCopyApplications @args }
        PreUpdateCommand = { Start-PreUpdateApplication  @args }
        PreImportCommand = { Start-PreImportCommandApplication  @args }
        DetailExtension = { Add-DetailExtensionApplications @args }
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsApplications @args }
        PreDeleteCommand = { Start-PreDeleteApplications @args }
        PostExportCommand = { Start-PostExportApplications @args }
        PostListCommand = { Start-PostListApplications @args }
        ExportExtension = { Add-ScriptExportApplications @args }
        PostGetCommand  = { Start-PostGetApplications @args }
        PostImportCommand = { Start-PostImportApplications @args }
        PostFilesImportCommand = { Start-PostFilesImportApplications @args }
        GroupId = "Apps"
        ScopeTagsReturnedInList = $false
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Autopilot"
        Id = "AutoPilot"
        API = "/deviceManagement/windowsAutopilotDeploymentProfiles"
        ViewID = "IntuneGraphAPI"
        CopyDefaultName = "%displayName% Copy" # '-' is not allowed in the name
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsAutoPilot @args }
        PreDeleteCommand = { Start-PreDeleteAutoPilot @args }
        PropertiesToRemoveForUpdate = @('managementServiceAppId')
        GroupId = "WinEnrollment"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Policy Sets"
        Id = "PolicySets"
        API = "/deviceAppManagement/policySets"
        ViewID = "IntuneGraphAPI"
        Expand = "Items"
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsPolicySets @args }
        PreImportCommand = { Start-PreImportPolicySets @args }
        PreUpdateCommand = { Start-PreUpdatePolicySets @args }
        PostListCommand = { Start-PostListPolicySets @args }
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        ImportOrder = 2000 # Policy Sets reference other objects so make sure it is imported last
        Dependencies = @("Applications","AppConfiguration","AppProtection","AutoPilot","EnrollmentRestrictions","EnrollmentStatusPage","DeviceConfiguration","AdministrativeTemplates","SettingsCatalog","CompliancePolicies")
        GroupId = "PolicySets"
        ExpandAssignmentsList = $false # expand is not allowed, IsAssigned is set in PostListCommand
    })
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Update Policies"
        Id = "UpdatePolicies"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/deviceConfigurations"
        QUERYLIST = "`$filter=isof(%27microsoft.graph.windowsUpdateForBusinessConfiguration%27)%20or%20isof(%27microsoft.graph.iosUpdateConfiguration%27)"
        #ExportFullObject = $false
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        GroupId = "WinUpdatePolicies"
        PropertiesToRemoveForUpdate = @('version','qualityUpdatesPauseStartDate','featureUpdatesPauseStartDate','qualityUpdatesWillBeRolledBack','featureUpdatesWillBeRolledBack')
    })
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Feature Updates"
        Id = "FeatureUpdates"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/windowsFeatureUpdateProfiles"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        GroupId = "WinFeatureUpdates"
        PropertiesToRemoveForUpdate = @('deployableContentDisplayName','endOfSupportDate')
        #PreUpdateCommand = { Start-PreUpdateFeatureUpdates @args } 
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Quality Updates"
        Id = "QualityUpdates"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/windowsQualityUpdateProfiles"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Icon = "UpdatePolicies"
        GroupId = "WinQualityUpdates"
        PropertiesToRemoveForUpdate = @('releaseDateDisplayName','deployableContentDisplayName')
    })    

    # Locations are not FULLY supported 
    # They will be imported but Compliance Policies will not be updated with new Location object after import
    # ToDo: Add support Export/Import Location Settings
    # Location object - Only used by Android Device Admins Compliance Policies 
    # - These should probably be migrated to Android Enterprise anyway. That is the recommendation by Google
    # Property that needs to be updated on the Compliance Policy
    # deviceManagement/managementConditionStatements/$obj.conditionStatementId
    
    # Location objects support removed from Intune
    <#
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Locations"
        Id = "Locations"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/managementConditions"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        PreImportCommand = { Start-PreImportLocations @args }
        ImportOrder = 30
        GroupId = "CompliancePolicies"
    })
    #>

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Settings Catalog"
        Id = "SettingsCatalog"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/configurationPolicies"
        PropertiesToRemove = @('settingCount')
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        NameProperty = "Name"
        ViewProperties = @("name","description","Id")
        Expand="Settings"
        Icon="DeviceConfiguration"
        PostExportCommand = { Start-PostExportSettingsCatalog  @args }
        PreUpdateCommand = { Start-PreUpdateSettingsCatalog  @args }
        PostGetCommand = { Start-PostGetSettingsCatalog  @args }
        Dependencies = @("ReusableSettings")
        GroupId = "DeviceConfiguration"        
    })   
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Role Definitions"
        Id = "RoleDefinitions"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/roleDefinitions"
        QUERYLIST = "`$filter=isBuiltIn%20eq%20false"
        PostExportCommand = { Start-PostExportRoleDefinitions @args }
        PreImportCommand = { Start-PreImportRoleDefinitions @args }
        PostFileImportCommand = { Start-PostFileImportRoleDefinitions @args }
        Permissons=@("DeviceManagementRBAC.ReadWrite.All")
        ImportOrder = 20
        #expand=roleassignments
        PropertiesToRemoveForUpdate = @('isBuiltInRoleDefinition','isBuiltIn','roleAssignments') ### !!! ToDo: Add support for roleAssignments
        GroupId = "TenantAdmin"
        ExpandAssignments = $false
        ExpandAssignmentsList = $false
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Scope (Tags)"
        Id = "ScopeTags"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/roleScopeTags"
        QUERYLIST = "`$filter=isBuiltIn%20eq%20false"
        Permissons=@("DeviceManagementRBAC.ReadWrite.All")
        PostExportCommand = { Start-PostExportScopeTags @args }
        PostGetCommand = { Start-PostGetScopeTags @args }
        ImportOrder = 10
        DocumentAll = $true
        GroupId = "TenantAdmin"
        ExpandAssignmentsList = $false # Adds the assignmnets property but always empty
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Notifications"
        Id = "Notifications"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/notificationMessageTemplates"
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        ImportOrder = 40
        Expand = "localizedNotificationMessages"
        PreImportCommand = { Start-PreImportNotifications @args }
        PostFileImportCommand = { Start-PostFileImportNotifications @args }
        PostCopyCommand = { Start-PostCopyNotifications @args }
        PropertiesToRemoveForUpdate = @('defaultLocale','localizedNotificationMessages') ### !!! ToDo: Add support for localizedNotificationMessages
        GroupId = "CompliancePolicies"
        ExpandAssignmentsList = $false
    })    
    
    # This has some pre-reqs for working!
    # Import is tested and verified in a tenant with Googple Play connection configured
    # And the OEM app was dpwnloaded e.g. Knox Service Plugin
    # Import failed in a tenant where Google Play was NOT configured
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Android OEM Config"
        Id = "AndroidOEMConfig"
        ViewID = "IntuneGraphAPI"
        QUERYLIST = "`$filter=microsoft.graph.androidManagedStoreAppConfiguration/appSupportsOemConfig%20eq%20true"
        API = "/deviceAppManagement/mobileAppConfigurations"
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsAppConfiguration @args }
        PreFilesImportCommand = { Start-PreFilesImportAppConfiguration @args }
        PostExportCommand = { Start-PostExportAppConfiguration @args }
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Icon="DeviceConfiguration"
        Dependencies = @("Applications")
        GroupId = "DeviceConfiguration"
    })

    # Copy/Export/Import not verified!
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Apple Enrollment Types"
        Id = "AppleEnrollmentTypes"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/appleUserInitiatedEnrollmentProfiles"
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        PropertiesToRemoveForUpdate = @('platform')
        GroupId = "AppleEnrollment"
    })
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Filters"
        Id = "AssignmentFilters"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/assignmentFilters"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        ImportOrder = 15
        GroupId = "TenantAdmin"
        PropertiesToRemoveForUpdate = @('platform')
        ExpandAssignmentsList = $false
        PropertiesToRemove = @("payloads")
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Health Scripts"
        Id = "DeviceHealthScripts"
        ViewID = "IntuneGraphAPI"
        QUERYLIST = "`$filter=isGlobalScript%20eq%20false" # Looks like filters are not working for deviceHealthScripts
        API = "/deviceManagement/deviceHealthScripts"
        PreDeleteCommand = { Start-PreDeleteDeviceHealthScripts @args }
        PreImportCommand = { Start-PreImportDeviceHealthScripts @args }
        PreUpdateCommand = { Start-PreUpdateDeviceHealthScripts @args }
        PostExportCommand = { Start-PostExportDeviceHealthScripts  @args }
        ExportExtension = { Add-ScriptExportExtensions @args }
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        GroupId = "EndpointAnalytics"
        Icon = "Report"
        AssignmentsType = "deviceHealthScriptAssignments"
        AssignmentProperties = @("target","runSchedule","runRemediationScript")
        PropertiesToRemoveForUpdate = @('version','isGlobalScript','highestAvailableVersion')
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "ADMX Files"
        Id = "ADMXFiles"
        ViewID = "IntuneGraphAPI"
        NameProperty = "fileName"
        API = "/deviceManagement/groupPolicyUploadedDefinitionFiles"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        ImportOrder = 45
        GroupId = "DeviceConfiguration"
        Icon = "DeviceConfiguration"
        ExpandAssignmentsList = $false
        PreFilesImportCommand = { Start-PreFilesImportADMXFiles @args }
        PreImportCommand = { Start-PreImportADMXFiles @args }
        PostImportCommand = { Start-PostImportADMXFiles @args }
        PreDeleteCommand = { Start-PreDeleteADMXFiles @args }
        ViewProperties = @("fileName","status","Id")
        PropertiesToRemove = @("languageCodes","targetPrefix","targetNamespace","policyType","revision","status","uploadDateTime")
    })

    <#
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "iOS Enrollment Profile"
        Id = "iOSDepProfile"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/depIOSEnrollmentProfile"
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        GroupId = "DeviceConfiguration"
        Icon = "DeviceConfiguration"
        ExpandAssignmentsList = $false
        ViewProperties = @("fileName","status","Id")
    })
    #>

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Reusable Settings"
        Id = "ReusableSettings"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/reusablePolicySettings"
        PropertiesToRemove = @('Settings','@OData.Type')        
        PostGetCommand = { Start-PostGetReusableSettings @args }
        ImportOrder = 70
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")        
        ExpandAssignmentsList = $false
        SkipRemoveProperties = @("@OData.Type")
        Icon = "EndpointSecurity"
        GroupId = "EndpointSecurity"
    })
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Authentication Strengths"
        Id = "AuthenticationStrengths"
        ViewID = "IntuneGraphAPI"
        API = "/identity/conditionalAccess/authenticationStrengths/policies"
        PreImportCommand = { Start-PreImportCommandAuthenticationStrengths @args }
        PropertiesToRemove = @()
        ImportOrder = 45
        Permissons=@("Policy.ReadWrite.ConditionalAccess")        
        ExpandAssignmentsList = $false
        Icon = "ConditionalAccess"
        GroupId = "EndpointSecurity"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Authentication Context"
        Id = "AuthenticationContext"
        ViewID = "IntuneGraphAPI"
        API = "/identity/conditionalAccess/authenticationContextClassReferences"
        PropertiesToRemove = @("@odata.type")
        SkipRemoveProperties = @('Id') 
        ImportOrder = 46
        PreImportCommand = { Start-PreImportCommandAuthenticationContext @args }
        Permissons=@("Policy.ReadWrite.ConditionalAccess")
        ExpandAssignmentsList = $false
        Icon = "ConditionalAccess"
        GroupId = "EndpointSecurity"
    })
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "W365 Provisioning Policies"
        Id = "W365ProvisioningPolicies"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/virtualEndpoint/provisioningPolicies"
        Permissons=@("CloudPC.ReadWrite.All")
        Icon = "Devices"
        GroupId = "DeviceConfiguration"
    })
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "W365 User Settings"
        Id = "W365UserSettings"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/virtualEndpoint/userSettings"
        Permissons = @("CloudPC.ReadWrite.All")
        Icon = "Devices"
        GroupId = "DeviceConfiguration"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Driver Update Profiles"
        Id = "DriverUpdateProfiles"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/windowsDriverUpdateProfiles"
        Permissons = @("DeviceManagementConfiguration.ReadWrite.All")
        Icon = "UpdatePolicies"
        GroupId = "WinDriverUpdatePolicies"
    })    
}

function Invoke-EMAuthenticateToMSAL
{
    param($params = @{})

    $global:EMViewObject.AppInfo = Get-GraphAppInfo "EMAzureApp" $global:DefaultAzureApp "EM"
    Set-MSALCurrentApp $global:EMViewObject.AppInfo
    & $global:msalAuthenticator.Login -Account (?? $global:MSALToken.Account.UserName (Get-Setting "" "LastLoggedOnUser")) @params
}

function Invoke-EMDeactivateView
{    
    $tmp = $mnuMain.Items | Where Name -eq "EMBulk"
    if($tmp) { $mnuMain.Items.Remove($tmp) }
}

function Invoke-EMActivatingView
{
    Show-MSALError
    
    # Refresh values in case they have changed
    $global:EMViewObject.AppInfo = (Get-GraphAppInfo "EMAzureApp" $global:DefaultAzureApp "EM")
    if(-not $global:EMViewObject.Authentication)
    {
        $global:EMViewObject.Authentication = Get-MSALAuthenticationObject    
    }

    # Add View specific menus
    Add-GraphBulkMenu
}

function Invoke-EMSaveSettings
{
    $tmpApp = Get-GraphAppInfo "EMAzureApp" $global:DefaultAzureApp

    if($global:appObj.ClientID -ne $tmpApp.ClientId -and $global:MSALToken)
    {
        # The app has changed. Need to authenticate to the new app
        Write-Status "Logging in to $((?? $global:appObj.Name "selected application"))"
        $global:EMViewObject.AppInfo = $tmpApp
        Set-MSALCurrentApp $global:EMViewObject.AppInfo
        Clear-MSALCurentUserVaiables
        Connect-MSALUser -Account $global:MSALToken.Account.Username
        Write-Status ""
    }

    Set-EMUIStatus
}

function Invoke-GraphAuthenticationUpdated
{
    Set-EMUIStatus

    $script:CustomADMXDefinitions = $null
}

function Set-EMUIStatus
{
    # Hide/Show Delete button
    $allowDelete = Get-SettingValue "EMAllowDelete"    
    $global:btnDelete.Visibility = (?: ($allowDelete -eq $true) "Visible" "Collapsed")

    # Hide/Show Delete on Bulk menu
    $allowBulkDelete = Get-SettingValue "EMAllowBulkDelete"
    $mnuBulk = $mnuMain.Items | Where Name -eq "EMBulk"

    if($mnuBulk) 
    {
        $mnuBulkDelete = $mnuBulk.Items | Where Name -eq "mnuBulkDelete"
        if($mnuBulkDelete)
        {
            $mnuBulkDelete.Visibility = (?: ($allowBulkDelete -eq $true) "Visible" "Collapsed")
        }
    }    
}

function Set-EMViewPanel
{
    param($panel)
    
    # ToDo: Create View specific pannel and move this to graph
    Add-XamlEvent $panel "btnView" "Add_Click" -scriptBlock ([scriptblock]{ 
        Show-GraphObjectInfo
    })

    Add-XamlEvent $panel "btnDelete" "Add_Click" -scriptBlock ([scriptblock]{ 
        Remove-GraphObjects
    })
    
    Add-XamlEvent $panel "btnCopy" "Add_Click" -scriptBlock ([scriptblock]{ 
        Copy-GraphObject
    })

    Add-XamlEvent $panel "btnExport" "Add_Click" -scriptBlock ([scriptblock]{
        Show-GraphExportForm
    })

    Add-XamlEvent $panel "btnImport" "Add_Click" -scriptBlock ([scriptblock]{
        Show-GraphImportForm
    })
    
    Add-XamlEvent $panel "txtFilter" "Add_LostFocus" ({ #param($obj, $e)
        Invoke-FilterBoxChanged $this
        #$e.Handled = $true
    })
    
    Add-XamlEvent $panel "txtFilter" "Add_GotFocus" ({
        if($this.Tag -eq "1" -and $this.Text -eq "Filter") { $this.Text = "" }
        Invoke-FilterBoxChanged $this
    })
    
    Add-XamlEvent $panel "txtFilter" "Add_TextChanged" ({
        Invoke-FilterBoxChanged $this
    })

    Invoke-FilterBoxChanged ($panel.FindName("txtFilter"))

    $allowDelete = Get-SettingValue "EMAllowDelete"
    Set-XamlProperty $panel "btnDelete" "Visibility" (?: ($allowDelete -eq $true) "Visible" "Collapsed")    

    $global:dgObjects.add_selectionChanged({        
        Invoke-ModuleFunction "Invoke-EMSelectedItemsChanged"
    })

    # ToDo: Move this to the view object
    $dpd = [System.ComponentModel.DependencyPropertyDescriptor]::FromProperty([System.Windows.Controls.ItemsControl]::ItemsSourceProperty, [System.Windows.Controls.DataGrid])
    if($dpd)
    {
        $dpd.AddValueChanged($global:dgObjects, {
            Set-XamlProperty $global:dgObjects.Parent "txtFilter" "Text" ""
            $enabled = (?: ($null -eq $this.ItemsSource -or ($this.ItemsSource | measure).Count -eq 0) $false $true)
            Set-XamlProperty $global:dgObjects.Parent "btnImport" "IsEnabled" $true # Always all Import if ObjectType allows it
            Set-XamlProperty $global:dgObjects.Parent "btnExport" "IsEnabled" $enabled
        })
    }

    $btnRefresh = Get-XamlObject ($global:AppRootFolder + "\Xaml\RefreshButton.xaml")
    if($btnRefresh)
    {
        $btnRefresh.SetValue([System.Windows.Controls.Grid]::ColumnProperty,$grdTitle.ColumnDefinitions.Count - 1)
        $btnRefresh.Margin = "0,0,5,3"
        $btnRefresh.Cursor = "Hand"
        $btnRefresh.Name = "btnRefresh"
        $btnRefresh.Focusable = $false
        $grdTitle.Children.Add($btnRefresh) | Out-Null

        $tooltip = [System.Windows.Controls.ToolTip]::new()
        $tooltip.Content = "Refresh all objects"
        [System.Windows.Controls.ToolTipService]::SetToolTip($btnRefresh, $tooltip)           

        $panel.RegisterName($btnRefresh.Name, $btnRefresh)

        $tooltip = [System.Windows.Controls.ToolTip]::new()
        $tooltip.Content = "Refresh objects"

        [System.Windows.Controls.ToolTipService]::SetToolTip($btnRefresh, $tooltip)

        $btnRefresh.Add_Click({
            $txtFilterText = $null
            $txtFilter = $this.Parent.FindName("txtFilter")
            if($txtFilter) { $txtFilterText = $txtFilter.Text } #= "" }
            
            Show-GraphObjects $txtFilterText

            if($txtFilterText -and $txtFilter)
            {
                $txtFilter.Text = $txtFilterText
                Invoke-FilterBoxChanged $txtFilter
            }

            Write-Status ""
        })
    }

    $global:btnLoadAllPages.add_click({
        Write-Status "Loading $($global:curObjectType.Title) objects"
        [array]$graphObjects = Get-GraphObjects -property $global:curObjectType.ViewProperties -objectType $global:curObjectType -AllPages
        $graphObjects | ForEach-Object { $global:dgObjects.ItemsSource.AddNewItem($_) | Out-Null }
        $global:dgObjects.ItemsSource.CommitNew()
        Set-GraphPagesButtonStatus
        Invoke-FilterBoxChanged $global:txtFilter -ForceUpdate
        Write-Status ""
    })

    $global:btnLoadNextPage.add_click({
        Write-Status "Loading $($global:curObjectType.Title) objects"
        [array]$graphObjects = Get-GraphObjects -property $global:curObjectType.ViewProperties -objectType $global:curObjectType -SinglePage
        $graphObjects | ForEach-Object { $global:dgObjects.ItemsSource.AddNewItem($_) | Out-Null }
        $global:dgObjects.ItemsSource.CommitNew()
        Set-GraphPagesButtonStatus
        Invoke-FilterBoxChanged $global:txtFilter
        Write-Status ""
    })    
}

function Invoke-GraphObjectsChanged
{
    $btnRefresh = $global:EMViewObject.ViewPanel.FindName("btnRefresh")

    if($btnRefresh)
    {
        $tooltip = [System.Windows.Controls.ToolTipService]::GetToolTip($btnRefresh)
        if($global:lstMenuItems.SelectedItem.QuerySearch -eq $true)
        {
            $tooltip.Content = "Refresh objects based on filter. Note: Only filtered objects will be returned. Clear filter and press refresh to reload other objects"
        }
        else
        {
            $tooltip.Content = "Refresh all objects"
        }
    }
}

function Invoke-EMSelectedItemsChanged
{
    $hasSelectedItems = ($global:dgObjects.ItemsSource | Where IsSelected -eq $true) -or ($null -ne $global:dgObjects.SelectedItem)
    Set-XamlProperty $global:dgObjects.Parent "btnView" "IsEnabled" $hasSelectedItems #(?: ($null -eq ($global:dgObjects.SelectedItem)) $false $true)
    Set-XamlProperty $global:dgObjects.Parent "btnCopy" "IsEnabled" $hasSelectedItems #(?: ($null -eq $global:dgObjects.SelectedItem) $false $true)
    Set-XamlProperty $global:dgObjects.Parent "btnDelete" "IsEnabled" $hasSelectedItems #(?: ($null -eq $global:dgObjects.SelectedItem -and $global:curObjectType.AllowDelete -ne $false) $false $true)
}

function Invoke-FilterBoxChanged 
{ 
    param($txtBox,[switch]$ForceUpdate)

    $filter = $null
    
    if($txtBox.Text.Trim() -eq "" -and $txtBox.IsFocused -eq $false)
    {
        $txtBox.FontStyle = "Italic"
        $txtBox.Tag = 1
        $txtBox.Text = "Filter"
        $txtBox.Foreground="Lightgray"
    }
    elseif($ForceUpdate -eq $true)
    {
        $dgObjects.ItemsSource.Filter = $dgObjects.ItemsSource.Filter
    }
    elseif($txtBox.Tag -eq "1" -and $txtBox.Text -eq "Filter" -and $txtBox.IsFocused -eq $false)
    {
        
    }
    else
    {            
        $txtBox.FontStyle = "Normal"
        $txtBox.Tag = $null
        $txtBox.Foreground="Black"
        $txtBox.Background="White"

        if($txtBox.Text)
        {
            $filter = {
                param ($item)

                return ($null -ne ($item.PSObject.Properties | Where { $_.Name -notin @("IsSelected","Object", "ObjectType") -and $_.Value -match [regex]::Escape($txtBox.Text) }))

                foreach($prop in ($item.PSObject.Properties | Where { $_.Name -notin @("IsSelected","Object", "ObjectType")}))
                {
                    if($prop.Value -match [regex]::Escape($txtBox.Text)) { return $true }
                }
                $false
            }
        }         
    }

    if($dgObjects.ItemsSource -is [System.Windows.Data.ListCollectionView] -and $txtBox.IsFocused -eq $true)
    {
        $dgObjects.ItemsSource.Filter = $filter
    }

    $allObjectsCount = 0
    if($dgObjects.ItemsSource.SourceCollection)
    {
        $allObjectsCount = $dgObjects.ItemsSource.SourceCollection.Count
    }

    $objCount = ($dgObjects.ItemsSource | measure).Count
    if($objCount -gt 0)
    {
        $strAllObjectsInfo = ""
        if($allObjectsCount -gt $objCount)
        {
            $strAllObjectsInfo = " ($($allObjectsCount))"
        }
        $global:txtEMObjects.Text = "Objects: $objCount$strAllObjectsInfo"
    }
    else
    {
        $global:txtEMObjects.Text = ""
    }
}
#region Endpoint Security (Intents) functions

function Start-PreImportEndpointSecurity
{
    param($obj, $objectType)

    @{
        "API"="deviceManagement/templates/$($obj.templateId)/createInstance"
    }
}

function Start-PostListEndpointSecurity
{
    param($objList, $objectType)

    if(-not $script:baseLineTemplates)
    {
        $script:baseLineTemplates = (Invoke-GraphRequest -Url "/deviceManagement/templates").Value
    }
    if(-not $script:baseLineTemplates) { return }

    foreach($obj in $objList)
    {        
        if(-not $obj.Object.templateId) { continue }
        if($obj.Object.templateId -ne $baseLineTemplate.Id)
        {
            $baseLineTemplate = $script:baseLineTemplates | Where Id -eq $obj.Object.templateId
        }
        
        if($baseLineTemplate)
        {
            $obj | Add-Member -MemberType NoteProperty -Name "Type" -Value $baseLineTemplate.displayName
            $obj | Add-Member -MemberType NoteProperty -Name "Category" -Value (?: ($baseLineTemplate.templateSubtype -eq "none") $baseLineTemplate.templateType $baseLineTemplate.templateSubtype)
        }
        
    }
    $objList
}

function Start-PostExportEndpointSecurity
{
    param($obj, $objectType, $path)

    $fileName = (Get-GraphObjectName $obj $objectType).Trim('.')
    if((Get-SettingValue "AddIDToExportFile") -eq $true -and $obj.Id)
    {
        $fileName = ($fileName + "_" + $obj.Id)
    }

    $settings = Invoke-GraphRequest -Url "$($objectType.API)/$($obj.id)/settings"
    $settingsJson = "{ `"settings`": $((ConvertTo-Json  $settings.value -Depth 20 ))`n}"
    $fileName = "$path\$((Remove-InvalidFileNameChars $fileName))_Settings.json"
    Save-GraphObjectToFile $settingsJson $fileName
}

function Start-PostFileImportEndpointSecurity
{
    param($obj, $objectType, $file)

    $settings = Get-EMSettingsObject $obj $objectType $file
    if($settings)
    {
        Start-GraphPreImport $settings
        Invoke-GraphRequest -Url "$($objectType.API)/$($obj.id)/updateSettings" -Body ($settings | ConvertTo-Json -Depth 50) -Method "POST"
    }    
}

function Start-PreCopyEndpointSecurity
{
    param($obj, $objectType, $newName)

    $false

    # Intents has a createCopy method. Use "manual" copy to have one standard and making sure Copy works the same as Export/Import
    # These objects supports duplicate in the portal
    # Keep for reference
    #
    # $objData = "{`"displayName`":`"$($newName)`"}"
    #
    #Invoke-GraphRequest -Url "/deviceManagement/intents/$($obj.Id)/createCopy" -Content $objData -HttpMethod "POST" | Out-Null
    #$true
}

function Start-PostCopyEndpointSecurity
{
    param($objCopyFrom, $objNew, $objectType)

    $settings = Invoke-GraphRequest -Url "$($objectType.API)/$($objCopyFrom.id)/settings" -ODataMetadata "Skip"
    if($settings)
    {
        $settingsObj = New-object PSObject @{ "Settings" = $settings.Value }
        Invoke-GraphRequest -Url "$($objectType.API)/$($objNew.id)/updateSettings" -Body ($settingsObj | ConvertTo-Json -Depth 20) -Method "POST"
    }
}

function Start-PreUpdateEndpointSecurity
{
    param($obj, $objectType, $curObject, $fromObj)

    if(-not $fromObj.settings) { return }

    $strAPI = "/deviceManagement/intents/$($curObject.Object.id)/updateSettings"
    
    $curObject = Get-GraphObject $curObject.Object $objectType

    $curValues = @()
    foreach($val in $curObject.Object.settings)
    {
        if($fromObj.settings | Where { $_.definitionId -eq $val.definitionId}) { continue }

        # Set all existing values to null
        # Note: This will not remove them from the configured list just set them Not Configured
        $curValues += [PSCustomObject]@{
            '@odata.type' = $val.'@odata.type'
            definitionId = $val.definitionId
            id = $val.id
            valueJson = "null"
        }
    }

    $curValues += $fromObj.settings

    <#
    if($curValues.Count -gt 0)
    {
        $tmpObj = [PSCustomObject]@{
            settings = $curValues
        }
        $json = ConvertTo-Json $tmpObj -Depth 20

        # Set all existing values to null
        # Note: This will not remove them from the configured list just set them Not Configured
        Invoke-GraphRequest -Url $strAPI -Content $json -HttpMethod "POST" | Out-Null
    }
    #>

    $tmpObj = [PSCustomObject]@{
        settings = $curValues
    }
    Start-GraphPreImport $tmpObj.settings 

    $json = ConvertTo-Json $tmpObj -Depth 20
    Invoke-GraphRequest -Url $strAPI -Content $json -HttpMethod "POST" | Out-Null

    Remove-Property $obj "templateId"
}

function Start-PostGetEndpointSecurity
{
    param($obj, $objectType)
    
    Add-EndpointSecurityInfo $obj
}

function local:Add-EndpointSecurityInfo
{
    param($obj, $baseLineTemplate = $null)
    
}
#endregion

#region 

function Start-PostFileImportDeviceConfiguration
{
    param($obj, $objectType, $importFile)

    if($obj.'@OData.Type' -like "#microsoft.graph.windows10GeneralConfiguration")
    {
        $tmpObj = Get-GraphObjectFromFile $importFile

        if(($tmpObj.privacyAccessControls | measure).Count -gt 0)
        {
            $privacyObj = [PSCustomObject]@{
                windowsPrivacyAccessControls = $tmpObj.privacyAccessControls
            }
            $json =  $privacyObj | ConvertTo-Json -Depth 20
            $ret = Invoke-GraphRequest -Url "deviceManagement/deviceConfigurations('$($obj.Id)')/windowsPrivacyAccessControls" -Body $json -Method "POST"
        }
    }
}

function Start-PostCopyDeviceConfiguration
{
    param($objCopyFrom, $objNew, $objectType)

    if($objCopyFrom.'@OData.Type' -like "#microsoft.graph.windows10GeneralConfiguration")
    {
        if(($objCopyFrom.privacyAccessControls | measure).Count -gt 0)
        {
            $privacyObj = [PSCustomObject]@{
                windowsPrivacyAccessControls = $objCopyFrom.privacyAccessControls
            }
            $json =  $privacyObj | ConvertTo-Json -Depth 20
            Invoke-GraphRequest -Url "deviceManagement/deviceConfigurations('$($objNew.Id)')/windowsPrivacyAccessControls" -Body $json -Method "POST" | Out-null
        }
    }
}

function Start-PostGetDeviceConfiguration
{
    param($obj, $objectType)
    
    if(($obj.Object.omaSettings | measure).Count -gt 0)
    {
        foreach($omaSetting in ($obj.Object.omaSettings | Where isEncrypted -eq $true))
        {
            if($omaSetting.isEncrypted -eq $false) { continue }

            $xmlValue = Invoke-GraphRequest -Url "/deviceManagement/deviceConfigurations/$($obj.Object.Id)/getOmaSettingPlainTextValue(secretReferenceValueId='$($omaSetting.secretReferenceValueId)')"
            if($xmlValue.Value)
            {
                $omaSetting.isEncrypted = $false
                $omaSetting.secretReferenceValueId = $null
                
                if($omaSetting.'@odata.type' -eq "#microsoft.graph.omaSettingStringXml" -or 
                $omaSetting.'value@odata.type' -eq "#Binary")
                {
                    $Bytes = [System.Text.Encoding]::UTF8.GetBytes($xmlValue.Value)
                    $omaSetting.value = [Convert]::ToBase64String($bytes)
                }
                else
                {
                    $omaSetting.value = $xmlValue.Value
                }
            }
        }
    }  
}

#endregion

#region Compliance Policy
function Start-PostExportCompliancePolicies
{
    param($obj, $objectType, $exportPath)

    foreach($scheduledActionsForRule in $obj.scheduledActionsForRule)
    {
        foreach($scheduledActionConfiguration in $scheduledActionsForRule.scheduledActionConfigurations)
        {
            foreach($notificationMessageCCGroup in $scheduledActionConfiguration.notificationMessageCCList)
            {
                Add-GroupMigrationObject $notificationMessageCCGroup
            }
        }
    }
}

function Start-PreUpdateCompliancePolicies
{
    param($obj, $objectType, $curObject, $fromObj)

    $strAPI = "/deviceManagement/deviceCompliancePolicies/$($curObject.Object.id)/scheduleActionsForRules"

    $tmpObj = [PSCustomObject]@{
        deviceComplianceScheduledActionForRules = $obj.scheduledActionsForRule
    }

    $json = ConvertTo-Json $tmpObj -Depth 20
    Invoke-GraphRequest -Url $strAPI -Content $json -HttpMethod "POST" | Out-Null

    Remove-Property $obj "scheduledActionsForRule"
}

#endregion

#region Intune Branding functions
function Start-PreImportIntuneBranding
{
    param($obj, $objectType)

    $ret = @{}
    $global:brandingClone = $null

    if($obj.isDefaultProfile)
    {
        
        # Looks like the ID is the same for all tenants so skip this for now
        <#
        $defObj  = (Invoke-GraphRequest -Url "/deviceManagement/intuneBrandingProfiles?`$filter=isDefaultProfile eq true&`$select=id,displayName").Value[0]
        if($defObj)
        {  
            $obj.Id = $defObj.Id
        }
        #>        

        $ret.Add("API",($objectType.API + "/" + $obj.Id))
        $ret.Add("Method","PATCH") # Default profile always exists so update it

        foreach($prop in @("profileName","isDefaultProfile","disableClientTelemetry","profileDescription"))
        {
            Remove-Property $obj $prop
        }

        $ret
    }
    else
    {
        # Create new Branding profile does not support images data in the json 
        # Workaround: (as done by the portal)
        # Create a new profile with basic info
        # Patch the profile with all the info

        $global:brandingClone = $obj | ConvertTo-Json -Depth 20 | ConvertFrom-Json

        foreach($prop in ($obj.PSObject.Properties | Where {$_.Name -notin @("profileName","profileDescription","roleScopeTagIds")})) #"customPrivacyMessage"
        {
            Remove-Property $obj $prop.Name
        }
    }
    Remove-Property $obj "Id"
}

function Start-PostImportIntuneBranding
{
    param($obj, $objectType)

    if($obj.isDefaultProfile -or -not $global:brandingClone) { return }

    foreach($prop in @("Id","isDefaultProfile","customPrivacyMessage","disableClientTelemetry")) #"isDefaultProfile","disableClientTelemetry"
    {
        Remove-Property $global:brandingClone $prop
    }
    $json = ($global:brandingClone | ConvertTo-Json -Depth 20)
    Invoke-GraphRequest -Url "$($objectType.API)/$($obj.Id)" -Body $json -Method "PATCH" | Out-Null
}

function Start-PostGetIntuneBranding
{
    param($obj, $objectType)

    foreach($imgType in @("themeColorLogo","lightBackgroundLogo","landingPageCustomizedImage"))
    {
        Write-LogDebug "Get $imgType for $($obj.Object.profileName)"
        $imgJson = Invoke-GraphRequest -Url "$($objectType.API)/$($obj.Object.Id)/$imgType"
        if($imgJson.Value)
        {
            $obj.Object.$imgType = $imgJson
        }
    }
}

function Start-PostExportIntuneBranding
{
    param($obj, $objectType, $path)

    foreach($imgType in @("themeColorLogo","lightBackgroundLogo","landingPageCustomizedImage"))
    {
        if($obj.$imgType.Value)
        {
            $fileName = "$path\$((Get-GraphObjectName $obj $objectType))_$imgType.jpg" 
            [IO.File]::WriteAllBytes($fileName, [System.Convert]::FromBase64String($obj.$imgType.Value))
        }
    }
}

function Start-PreDeleteIntuneBranding
{
    param($obj, $objectType)

    if($obj.isDefaultProfile -eq $true)
    {
        @{ "Delete" = $false }
    }
}

function Start-PreUpdateIntuneBranding
{
    param($obj, $objectType, $curObject, $fromObj)

    if($curObject.Object.isDefaultProfile)
    {
        foreach($prop in @("profileName","isDefaultProfile","disableClientTelemetry","profileDescription"))
        {
            Remove-Property $obj $prop
        }
    }
}

#endregion

#region Azure Branding functions
function Start-PreImportAzureBranding
{
    param($obj, $objectType)

    Remove-Property $obj "@odata.Type"

    $ret = @{}
    if($obj.Id -eq "0")
    {
        #$ret.Add("Method","PATCH") # Default profile always exists so update it
        #$ret.Add("API",($objectType.API + "/0"))
    }

    $ret.Add("API",($objectType.API + "/$($global:Organization.Id)/branding/localizations"))

    # This is NOT wat the documentation says
    # Documentation says to use Content-Language
    # Any place the documentation states to use Accept-Language is for Get operation
    # https://docs.microsoft.com/en-us/graph/api/organizationalbrandingproperties-get?view=graph-rest-beta&tabs=http#request-headers
    $ret.Add("AdditionalHeaders", @{ "Accept-Language" = $obj.Id })

    $ret
}

function Start-PostListAzureBranding
{
    param($objList, $objectType)

    foreach($obj in $objList)
    {
        if(-not $obj.Object.id) { continue }
        try
        {
            if($obj.Object.id -eq "0")
            {
                $language = "Default"
            }
            else
            {
                $language = ([cultureinfo]::GetCultureInfo($obj.Object.id)).DisplayName
            }

            $obj | Add-Member -MemberType NoteProperty -Name "Language" -Value $language
        }
        catch{}
    }
    $objList
}

#endregion

#region Script functions
function Add-ScriptExtensions
{
    param($form, $buttonPanel, $index = 0)

    $btnDownload = New-Object System.Windows.Controls.Button    
    $btnDownload.Content = 'Download'
    $btnDownload.Name = 'btnDownload'
    $btnDownload.Margin = "0,0,5,0"  
    $btnDownload.Width = "100"
    
    $btnDownload.Add_Click({
        Invoke-DownloadScript
    })

    $tmp = $form.FindName($buttonPanel)
    if($tmp) 
    { 
        $tmp.Children.Insert($index, $btnDownload)
    }

    $btnDownload = New-Object System.Windows.Controls.Button    
    $btnDownload.Content = 'Edit'
    $btnDownload.Name = 'btnEdit'
    $btnDownload.Margin = "0,0,5,0"  
    $btnDownload.Width = "100"
    
    $btnDownload.Add_Click({
        Invoke-EditScript
    })

    $tmp = $form.FindName($buttonPanel)
    if($tmp) 
    { 
        $tmp.Children.Insert($index, $btnDownload)
    }    
}

function Add-ScriptExportExtensions
{
    param($form, $buttonPanel, $index = 0)

    $ctrl = $form.FindName("chkExportScript")
    if(-not $ctrl)
    {
        $xaml =  @"
<StackPanel $($global:wpfNS) Orientation="Horizontal" Margin="0,0,5,0">
<Label Content="Export script" />
<Rectangle Style="{DynamicResource InfoIcon}" ToolTip="Export the script associated with PowerShell or Shell profiles" />
</StackPanel>
"@
        $label = [Windows.Markup.XamlReader]::Parse($xaml)

        $global:chkExportScript = [System.Windows.Controls.CheckBox]::new()
        $global:chkExportScript.IsChecked = $true
        $global:chkExportScript.VerticalAlignment = "Center" 
        $global:chkExportScript.Name = "chkExportScript" 

        @($label, $global:chkExportScript)
    }
}

function Start-PostExportScripts
{
    param($obj, $objectType, $exportPath)

    if($obj.scriptContent -and $global:chkExportScript.IsChecked)
    {
        Write-Log "Export script $($obj.FileName)"
        $fileName = [IO.Path]::Combine($exportPath, $obj.FileName)
        [IO.File]::WriteAllBytes($fileName, ([System.Convert]::FromBase64String($obj.scriptContent)))
    }
}

function Invoke-DownloadScript
{
    if(-not $global:dgObjects.SelectedItem.Object.id) { return }

    $obj = (Get-GraphObject $global:dgObjects.SelectedItem $global:curObjectType).Object
    Write-Status ""

    if($obj.scriptContent)
    {            
        Write-Log "Download PowerShell script '$($obj.FileName)' from $($obj.displayName)"
        
        $dlgSave = New-Object -Typename System.Windows.Forms.SaveFileDialog
        $dlgSave.InitialDirectory = Get-SettingValue "IntuneRootFolder" $env:Temp
        $dlgSave.FileName = $obj.FileName    
        if($dlgSave.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $dlgSave.Filename)
        {
            # Changed to WriteAllBytes to get rid of BOM characters from Custom Attribute file 
            [IO.File]::WriteAllBytes($dlgSave.FileName, ([System.Convert]::FromBase64String($obj.scriptContent)))
        }
    }    
}

function Invoke-EditScript
{
    if(-not $global:dgObjects.SelectedItem.Object.id) { return }

    $obj = (Get-GraphObject $global:dgObjects.SelectedItem $global:curObjectType)
    Write-Status ""
    if(-not $obj.Object.scriptContent) { return }
    $script:currentScriptObject = $obj

    $script:editForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\EditScriptDialog.xaml")
    
    if(-not $script:editForm) { return }

    Set-XamlProperty $script:editForm "txtEditScriptTitle" "Text" "Edit: $($obj.Object.displayName)"
    
    $scriptText = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.Object.scriptContent))
    Set-XamlProperty $script:editForm "txtScriptText" "Text" $scriptText

    $script:currentModal = $null
    if($global:grdModal.Children.Count -gt 0)
    {
        $script:currentModal = $global:grdModal.Children[0]
    }

    Add-XamlEvent $script:editForm "btnSaveScriptEdit" "add_click" ({
        $scriptText = Get-XamlProperty $script:editForm "txtScriptText" "Text"
        $pre = [System.Text.Encoding]::UTF8.GetPreamble()
        $utfBOM = [System.Text.Encoding]::UTF8.GetString($pre)
        if($scriptText.startsWith($utfBOM))
        {
            # Remove UTF8 BOM bytes
            $scriptText = $scriptText.Remove(0, $utfBOM.Length)
        }
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($scriptText)
        $encodedText = [Convert]::ToBase64String($bytes)

        if($script:currentScriptObject.Object.scriptContent -ne $encodedText)
        {
            # Save script
            if(([System.Windows.MessageBox]::Show("Are you sure you want to update the script?`n`nObject:`n$($script:currentScriptObject.displayName)", "Update script?", "YesNo", "Warning")) -eq "Yes")
            {
                Write-Status "Update $($script:currentScriptObject.displayName)"
                $obj =  $script:currentScriptObject.Object | ConvertTo-Json -Depth 20 | ConvertFrom-Json
                $obj.scriptContent = $encodedText
                Start-GraphPreImport $obj $script:currentScriptObject.ObjectType
                foreach($prop in $script:currentScriptObject.ObjectType.PropertiesToRemoveForUpdate)
                {
                    Remove-Property $obj $prop
                }                
                Remove-Property $obj "Assignments"
                Remove-Property $obj "isAssigned"

                $json = ConvertTo-Json $obj -Depth 15

                $objectUpdated = (Invoke-GraphRequest -Url "$($script:currentScriptObject.ObjectType.API)/$($script:currentScriptObject.Object.Id)" -Content $json -HttpMethod "PATCH")
                if(-not $objectUpdated)
                {
                    Write-Log "Failed to update script" 3
                    [System.Windows.MessageBox]::Show("Failed to save the script object. See log for more information","Update failed!", "OK", "Error")
                }
                Write-Status ""
            }
        }
        
        $global:grdModal.Children.Clear()
        if($script:currentModal)
        {
            $global:grdModal.Children.Add($script:currentModal)
        }
        [System.Windows.Forms.Application]::DoEvents()
    })    
    
    Add-XamlEvent $script:editForm "btnCancelScriptEdit" "add_click" ({
        $global:grdModal.Children.Clear()
        if($script:currentModal)
        {
            $global:grdModal.Children.Add($script:currentModal)
        }
        [System.Windows.Forms.Application]::DoEvents()
    })
    
    $global:grdModal.Children.Clear()
    $script:editForm.SetValue([System.Windows.Controls.Grid]::RowProperty,1)
    $script:editForm.SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
    $global:grdModal.Children.Add($script:editForm) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
}

#endregion

#region Terms and Conditions
function Start-PostExportTermsAndConditions
{
    param($obj, $objectType, $path)

    Add-EMAssignmentsToExportFile $obj $objectType $path 
}

function Start-PreImportAssignmentsTermsAndConditions
{
    param($obj, $objectType, $file, $assignments)

    Add-EMAssignmentsToObject $obj $objectType $file $assignments
}
#endregion

#region App Protection functions

function Start-GetAppProtection
{
    param($obj, $objectType)

    if(-not $obj."@odata.type") { return }

    Get-GraphMetaData
    
    $objectClass = $null
    if($global:metaDataXML)
    {
        try 
        {
            $tmp = $obj."@odata.type".Split('.')[-1]
            $objectClass = Get-GraphObjectClassName $tmp
        }
        catch 
        {
            
        }
        $expand = $null
        if($objectClass -eq "windowsInformationProtectionPolicies")
        {
            $expand = "?`$expand=protectedAppLockerFiles,exemptAppLockerFiles"
        }

        if($objectClass)
        {
            @{"API"="/deviceAppManagement/$objectClass/$($obj.Id)$expand"}
        }
    }
}

function Start-PostListAppProtection
{
    param($objList, $objectType)

    # App Configurations for Managed Apps are included in App Protections e.g. the /deviceAppManagement/managedAppPolicies API
    # For some reason, the $filter option is not supported to filter out these objects
    # e.g. not isof(...) to excluded the type, not startsWith(id, 'A_') to exlude based on Id
    # These filters generates a request error so filter them out manually in this function instead
    # The portal is probably doing the same thing since these are included in the return but not in the UI
    $objList | Where { $_.Object.'@OData.Type' -ne '#microsoft.graph.targetedManagedAppConfiguration' }
}

function Start-PreImportAppProtection
{
    param($obj, $objectType)
    
    if(($obj.Apps | measure).Count -gt 0)    
    {        
        $global:ImportObjectInfo = @{ Apps=$obj.Apps }
    }
    else
    {        
        $global:ImportObjectInfo = $null
    }

    $global:ImportObjectClass = $null
    if($obj."@odata.type")
    {
        try
        {
            $global:ImportObjectClass = Get-GraphObjectClassName ($obj."@odata.type".Split('.')[-1])
        }
        catch {}
    }

    Remove-Property $obj "apps"
    Remove-Property $obj "apps@odata.context"

    try
    {
        $tmp = $obj."@odata.type".Split('.')[-1]
        $objectClass = Get-GraphObjectClassName $tmp
        if($objectClass)
        {
            @{"API"="/deviceAppManagement/$objectClass"}
        }
    }
    catch {}
}

function Start-PostImportAppProtection
{
    param($obj, $objectType, $file)
    
    if($global:ImportObjectInfo.Apps)
    {
        # No "@odata.type" on the created object so reload new object
        #$newObject = (Invoke-GraphRequest "$($objectType.API)?`$filter=id eq '$($obj.Id)'").Value
        $newObject = Invoke-GraphRequest "$($objectType.API)/$($obj.Id)"
        if($newObject)
        {
            try
            {
                $tmp = $newObject."@odata.type".Split('.')[-1]
                $objectClass = Get-GraphObjectClassName $tmp

                $apps = [PSCustomObject]@{ 
                    appGroupType = $obj.appGroupType
                    apps = @($global:ImportObjectInfo.Apps)                 
                } 
                $json = $apps | ConvertTo-Json -Depth 20
                
                Invoke-GraphRequest -Url "/deviceAppManagement/$objectClass/$($obj.Id)/targetApps" -Content $json -HttpMethod POST | Out-Null
            }
            catch {}
        }
    }
    $global:ImportObjectInfo = $null
}

function Start-PreImportAssignmentsAppProtection
{
    param($obj, $objectType, $file, $assignments)

    if($global:ImportObjectClass)
    {
        @{"API"="/deviceAppManagement/$($global:ImportObjectClass)/$($obj.Id)/assign"}
    }
}

function Start-PreUpdateAppConfigurationApp
{
    param($obj, $objectType, $curObject, $fromObj)
    
    if($obj.Apps)
    {
        try
        {
            Write-Log "Update App Configuruation Apps"

            $apps = [PSCustomObject]@{ 
                appGroupType = $obj.appGroupType
                apps = @($obj.Apps)                 
            } 
            $json = $apps | ConvertTo-Json -Depth 20
            $objectClass = 'targetedManagedAppConfigurations'

            Invoke-GraphRequest -Url "/deviceAppManagement/$objectClass/$($curObject.Object.Id)/targetApps" -Content $json -HttpMethod POST | Out-Null
        }
        catch {}
    }

    Remove-Property $obj "apps"
}

function Start-PreUpdateAppProtection
{
    param($obj, $objectType, $curObject, $fromObj)

    if($curObject.Object.'@OData.Type' -eq "#microsoft.graph.windowsInformationProtectionPolicy")
    {
        $api = "/deviceAppManagement/windowsInformationProtectionPolicies/$($curObject.Object.Id)"
    }
    elseif($curObject.Object.'@OData.Type' -eq "#microsoft.graph.mdmWindowsInformationProtectionPolicy")
    {
        $api = "/deviceAppManagement/mdmWindowsInformationProtectionPolicies/$($curObject.Object.Id)"
    }
    elseif($curObject.Object.'@OData.Type' -eq "#microsoft.graph.iosManagedAppProtection")
    {
        $api = "/deviceAppManagement/iosManagedAppProtections/$($curObject.Object.Id)"
    }
    elseif($curObject.Object.'@OData.Type' -eq "#microsoft.graph.androidManagedAppProtection")
    {
        $api = "/deviceAppManagement/androidManagedAppProtections/$($curObject.Object.Id)"        
    }
    else
    {
        return (Start-PreUpdateAppConfigurationApp $obj $objectType $curObject $fromObj)
    }
    
    if($obj.Apps)
    {
        try
        {
            Write-Log "Update App Protection Apps"
            
            $apps = [PSCustomObject]@{ 
                appGroupType = $obj.appGroupType
                apps = @($obj.Apps)                 
            } 
            $json = $apps | ConvertTo-Json -Depth 20

            Invoke-GraphRequest -Url "$api/targetApps" -Content $json -HttpMethod POST | Out-Null
        }
        catch {}
        
        Remove-Property $obj "apps"
    }

    @{ "API" = $api }

}
#endregion

#region App Configuration
function Start-PostExportAppConfiguration
{
    param($obj, $objectType, $path)

    #Add-EMAssignmentsToExportFile $obj $objectType $path

    Write-Log "Export app config for $($objectType.Id) with OData.Type: $($obj.'@OData.Type')"

    if($obj.'@OData.Type' -eq "#microsoft.graph.androidManagedAppProtection" -or
        $obj.'@OData.Type' -eq "#microsoft.graph.androidForWorkMobileAppConfiguration" -or
        $obj.'@OData.Type' -eq "#microsoft.graph.androidManagedStoreAppConfiguration" -or
        $obj.'@OData.Type' -eq "#microsoft.graph.iosMobileAppConfiguration")
    {
        $fileName = (Get-GraphObjectName $obj $objectType).Trim('.')
        if((Get-SettingValue "AddIDToExportFile") -eq $true -and $obj.Id)
        {
            $fileName = ($fileName + "_" + $obj.Id)
        }
        $tmpObj = $null
        $fileName = "$path\$((Remove-InvalidFileNameChars $fileName)).json"
        if([IO.File]::Exists($fileName))
        {
            $tmpObj = Get-GraphObjectFromFile $fileName
        }
        else
        {
            Write-Log "File not found: $fileName. Could not add App names." 3
        }

        if(($tmpObj.targetedMobileApps | measure).Count -gt 0)
        {        
            Write-Log "Add target apps info"
            $targetedApps = @()
            foreach($appId in $tmpObj.targetedMobileApps)
            {            
                $appObj = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$($appId)" #?`select=id,displayName" -ODataMetadata "Minimal"
                if($appObj) 
                {
                    Write-Log "Add target app info $($appObj.displayName) ($($appObj.Id)) of type $($appObj.'@OData.Type')"
                    $targetedApps += $appObj.displayName + '|!|' + $appObj.Id  + '|!|' + $appObj.'@OData.Type'
                }
            }

            if($targetedApps.Count -gt 0) 
            {
                Write-Log "Add CustomRefTargetedApps property"
                $tmpObj | Add-Member -MemberType NoteProperty -Name "#CustomRefTargetedApps" -Value ($targetedApps -join "|*|")
                Write-Log "Save file $fileName"
                Save-GraphObjectToFile $tmpObj $fileName
            }
        }
        else 
        {
            Write-Log "No target apps found" 2
        }
    }
}

function Start-PreFilesImportAppConfiguration
{
    param($objectType, $filesToImport)

    $targetedAppsObjects = $filesToImport | Where { $null -ne $_.Object."#CustomRefTargetedApps" }
    
    if(($targetedAppsObjects | measure).Count -gt 0)
    {
        Write-Log "Policies with Targeted Apps detected"
        foreach($fileObject in $targetedAppsObjects)
        {
            Add-AppConfigurationTargets $objectType $fileObject
        }    
    }
    $filesToImport
}

function local:Add-AppConfigurationTargets
{
    param($obj, $fileObj)

    if($fileObj.Object."#CustomRefTargetedApps" -and $fileObj.Object.targetedMobileApps)
    {
        Write-Log "Adding app target for $($fileObj.Object.displayName)"

        $targetedAppsInfo = $fileObj.Object."#CustomRefTargetedApps"

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
                $tmpApps = (Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps?`$filter=displayName eq '$appName'").value
                if(-not $tmpApps)
                {
                    Write-Log "No application found with name $appName. $appId will not be translated and added to target list" 2
                    continue
                }

                Write-Log "Found $(($tmpApps | measure).Count) applications" 2
                foreach ($tmpApp in $tmpApps) {
                    Write-Log "Found '$($tmpApp.displayName)' ($($tmpApp.id)) of type $($($tmpApp.'@OData.Type'))"
                }

                $tmpApp = $tmpApps | Where-Object '@OData.Type' -eq $appType
                if(-not $tmpApp)
                {
                    Write-Log "No $appName application found of type $appType. $appId will not be translated and added to target list" 2
                }
                elseif(($tmpApp | measure).Count -gt 1) {
                    Write-Log "$(($tmpApp | measure).Count) applications found with name '$appName' of type $appType. $appId will not be translated and added to target list" 2
                }
                else {
                    Write-Log "Found '$appName' with id $($tmpApp.Id) ($appType)"
                    $translatedTargetedApps += $tmpApp.Id
                }
            }

            if($translatedTargetedApps.Count -gt 0) {
                Write-Log "Updating translated targeted apps"
                $fileObj.Object.targetedMobileApps = $translatedTargetedApps
            }
            else {
                Write-Log "Could not find targeted apps in the evnironment. Verify that they are added. Policy import might fail" 3
            }
        }
    }    
}

function Start-PreImportAssignmentsAppConfiguration
{
    param($obj, $objectType, $file, $assignments)

    @{"API"="/deviceAppManagement/mobileAppConfigurations/$($obj.Id)/microsoft.graph.managedDeviceMobileAppConfiguration/assign"}
}
#endregon

#region Applications

function Start-PostCopyApplications
{
    param($objCopyFrom, $objNew, $objectType)

    Start-ImportApp $objNew
    Write-Status ""
}

function Start-PostFileImportApplications
{
    param($obj, $objectType, $file)
    
    $tmpObj = Get-GraphObjectFromFile $file

    if(-not ($obj.PSObject.Properties | Where Name -eq '@odata.type'))
    {
        # Add @odata.type property if it is missing. Required by app package import
        $obj | Add-Member -MemberType NoteProperty -Name '@odata.type' -Value $objectType.'@odata.type'
    }

    $fi = [IO.FileInfo]$file
    $tmpFilName = $fi.DirectoryName + "\" + $obj.FileName

    if([IO.File]::Exists($tmpFilName) -eq $false)
    {
        $tmpFilName = $null
    }
    
    Start-ImportApp $obj $tmpFilName
}

function local:Start-ImportApp
{
    param($obj, $packageFile = $null)
    
    if(-not $obj.'@odata.type') { return }

    if($null -eq $packageFile)
    {
        $pkgPath = Get-SettingValue "EMIntuneAppPackages"

        if(-not $pkgPath -or [IO.Directory]::Exists($pkgPath) -eq $false) 
        {
            Write-LogDebug "Package source directory is either missing or does not exist" 2
            return 
        }

        $packageFile = "$($pkgPath)\$($obj.fileName)"
    }
    $fi = [IO.FileInfo]$packageFile

    if($fi.Exists -eq $false) 
    {
        Write-LogDebug "Package source file $($fi.FullName) not found" 2
        return 
    }

    Write-Status "Import appliction package file $($fi.FullName)"
    Write-Log "Import application file '$($($fi.FullName))' for $($obj.displayName)"

    $appType = $obj.'@odata.type'.Trim('#')

    if($appType -eq "microsoft.graph.win32LobApp")
    {
        $fileEncryptionInfo = Copy-Win32LOBPackage $packageFile $obj
    }
    elseif($appType -eq "microsoft.graph.windowsMobileMSI")
    {
        $fileEncryptionInfo = Copy-MSILOB $packageFile $obj
    }
    elseif($appType -eq "microsoft.graph.windowsUniversalAppX")
    {
        $fileEncryptionInfo = Copy-MSIXLOB $packageFile $obj
    }    
    elseif($appType -eq "microsoft.graph.iosLOBApp")
    {
        $fileEncryptionInfo = Copy-iOSLOB $packageFile $obj
    }
    elseif($appType -eq "microsoft.graph.androidLOBApp")
    {
        $fileEncryptionInfo = Copy-AndroidLOB $packageFile $obj
    }
    else
    {
        Write-Log "Unsupported application type $appType. File will not be uploaded" 2    
    }

    if((Get-SettingValue "EMSaveEncryptionFile") -eq $true)
    {
        if($fileEncryptionInfo)
        {
            $jsonEncryptionInfo = $fileEncryptionInfo | ConvertTo-Json -Depth 10
            
            $pkgPath = Get-SettingValue "EMIntuneAppDownloadFolder" (Get-SettingValue "EMIntuneAppPackages")
            if($pkgPath -and [IO.Directory]::Exists($pkgPath))
            {
                $obj = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$($obj.id)" -ODataMetadata "Minimal"
                $fullPath = $pkgPath + "\$($obj.displayName)_$($obj.id)_$($obj.committedContentVersion).json"
                $jsonEncryptionInfo | Out-File -FilePath $fullPath -Force -Encoding utf8
            }
        }
    }
}

function Start-PreUpdateApplication
{
    param($obj, $objectType, $curObject, $fromObj)

    if($curObject.Object.'@OData.type' -eq "#microsoft.graph.windowsMobileMSI")
    {
        Remove-Property $obj "useDeviceContext"
    }
    elseif($curObject.Object.'@OData.type' -eq "#microsoft.graph.officeSuiteApp")
    {
        Remove-Property $obj "officeConfigurationXml"
        Remove-Property $obj "officePlatformArchitecture"
        Remove-Property $obj "developer"
        Remove-Property $obj "owner"
        Remove-Property $obj "publisher"
    }

    Remove-Property $obj "appStoreUrl"
}

function Start-PreImportCommandApplication
{
    param($obj, $objectType, $file, $assignments)

    if($obj.'@OData.Type' -in @('#microsoft.graph.microsoftStoreForBusinessApp','#microsoft.graph.androidStoreApp'))
    {
        Write-Log "App type '$($obj.'@OData.Type')' not supported for import" 2
        @{ "Import" = $false }
    }

    if($obj.'@OData.Type' -eq '#microsoft.graph.officeSuiteApp')
    {
        if($obj.officeSuiteAppDefaultFileFormat -eq "notConfigured")
        {
            $obj.officeSuiteAppDefaultFileFormat = "officeOpenXMLFormat"
        }
    }
} 

function Add-DetailExtensionApplications
{
    param($form, $buttonPanel, $index = 0)

    $btnUpload = New-Object System.Windows.Controls.Button    
    $btnUpload.Content = 'Upload'
    $btnUpload.Name = 'btnUploadAppfile'
    $btnUpload.Margin = "0,0,5,0"  
    $btnUpload.Width = "100"
    
    $btnUpload.Add_Click({
        if($global:dgObjects.SelectedItem.Object.publishingState -ne "notPublished")
        {
            # Only allow upload of not published apps
            # Use portal to replace app file...
            if(([System.Windows.MessageBox]::Show("Are you sure you want to upload a new file for the app?`n`nApplication:`n$($global:dgObjects.SelectedItem.Object.displayName)", "Update app file?", "YesNo", "Warning")) -ne "Yes")
            {
                return
            }
        }
    
        $pkgPath = Get-SettingValue "EMIntuneAppPackages"

        $of = [System.Windows.Forms.OpenFileDialog]::new()
        $of.FileName = $global:dgObjects.SelectedItem.Object.fileName
        $of.DefaultExt = "*.intunewin"
        $of.Filter = "Intune Win32 (*.intunewin)|*.*"
        $of.Multiselect = $false

        if($pkgPath -and [IO.Directory]::Exists($pkgPath))
        {
            $of.InitialDirectory = $pkgPath
        }
        
        if($of.ShowDialog() -eq "OK")
        {
            Write-Status "Import $($global:dgObjects.SelectedItem.Object.displayName) file"
            Start-ImportApp $global:dgObjects.SelectedItem.Object $of.FileName
            Write-Status ""
        }
    })

    $tmp = $form.FindName($buttonPanel)
    if($tmp) 
    { 
        $tmp.Children.Insert($index, $btnUpload)
    }

    $btnDownload = New-Object System.Windows.Controls.Button    
    $btnDownload.Content = 'Download'
    $btnDownload.Name = 'btnDownloadAppfile'
    $btnDownload.Margin = "0,0,5,0"  
    $btnDownload.Width = "100"
    
    $btnDownload.Add_Click({
        Write-Status "Download file"
        $obj = $global:dgObjects.SelectedItem.Object
        #$obj = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$($obj.id)"

        $pkgPath = Get-SettingValue "EMIntuneAppDownloadFolder" (Get-SettingValue "EMIntuneAppPackages")

        $dlgSave = [System.Windows.Forms.SaveFileDialog]::new()
        $dlgSave.InitialDirectory = $pkgPath
        $dlgSave.FileName = ($obj.FileName + ".encrypted")
        $dlgSave.DefaultExt = "*.encrypted"
        $dlgSave.Filter = "Encrypted intunewin (*.encrypted)|*.encrypted|All files (*.*)|*.*"

        if($dlgSave.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $dlgSave.Filename)
        {
            $contentFileObj = Start-DownloadAppContent $obj $dlgSave.FileName

            if([IO.File]::Exists($dlgSave.FileName))
            {
                $fullPath = Find-AppEncryptionFile $obj $contentFileObj $pkgPath
                if([IO.File]::Exists($fullPath) -eq $false)
                {
                    if(([System.Windows.MessageBox]::Show("Could not find decryption file for $($obj.displayName)`nApp Id: $($obj.id)`nContent version $($obj.committedContentVersion)`n`nDo you want to browse for the file?", "Encryption file not found", "YesNo", "Warning")) -eq "Yes")
                    {
                        $of = [System.Windows.Forms.OpenFileDialog]::new()
                        $of.InitialDirectory = $pkgPath
                        $of.DefaultExt = "*.json"
                        $of.Filter = "Json (*.json)|*.json"
                        $of.Multiselect = $false
                        
                        if($of.ShowDialog() -eq "OK")
                        {
                            $fullPath = $of.FileName
                        }                    
                    }
                }

                if([IO.File]::Exists($fullPath))
                {
                    Write-Status "Decrypting file"
                    $encryptionInfo = ConvertFrom-Json (Get-Content -Path $fullPath -Raw)
                    if($encryptionInfo.fileEncryptionInfo)
                    {
                        $encryptionInfo = $encryptionInfo.fileEncryptionInfo
                    }
                    $destination = $pkgPath + "\$($obj.FileName)"
                    Start-DecryptFile $dlgSave.Filename $destination $encryptionInfo.encryptionKey $encryptionInfo.initializationVector
                    try { [IO.File]::Delete($dlgSave.Filename) }
                    catch {
                        Write-LogError "Failed to delete exported encrypted file" $_.Exception
                    }                
                }
                else
                {
                    Write-Log "Decryption file for $($obj.displayName) not found. Skipping decryption" 2
                }
            }
        }

        Write-Status ""
    })

    $tmp = $form.FindName($buttonPanel)
    if($tmp) 
    { 
        $tmp.Children.Insert($index, $btnDownload)
    }    
}

function Find-AppEncryptionFile
{
    param($obj, $contentFileObj, $rootFolders)

    $search = @()
    $search += "$($obj.displayName)_$($obj.id)_$($obj.committedContentVersion)"
    $search += "$([IO.Path]::GetFileNameWithoutExtension($obj.fileName))_$($contentFileObj.size)"
    $search += "$($obj.displayName)_$($contentFileObj.size)"

    foreach($rootFolder in $rootFolders)
    {
        foreach($searchName in $search)
        {
            $fullName = ($rootFolder + "\$($searchName).json")
            if([IO.File]::Exists($fullName))
            {
                return $fullName
            }
        }
    }
}

function Start-PreImportAssignmentsApplications
{
    param($obj, $objectType, $file, $assignments)

    if($obj.'@odata.type' -eq "#microsoft.graph.windowsMicrosoftEdgeApp")
    {
        foreach($assignment in $assignments)
        {
            Remove-Property $assignment.target "deviceAndAppManagementAssignmentFilterId"
            Remove-Property $assignment.target "deviceAndAppManagementAssignmentFilterType"
        }
        @{"Assignments"=$assignments}
    }
    elseif($obj.'@odata.type' -eq "#microsoft.graph.winGetApp")
    {
        Write-LogDebug "Wait for app to be published"
        $i = 2
        Start-Sleep -s ($i)
        $x = 0
        while($x -lt 10)
        {
            ###!!!
            $appInfo = Invoke-GraphRequest -Url "$($objectType.API)/$($obj.id)" -ODataMetadata "skip"
            if($appInfo.publishingState -eq "Published")
            {
                Write-LogDebug "Application $($obj.displayName) is published"
                return
            }
            Start-Sleep -s ($i)
            $x++
            if($x -ge 5) { $i++ }
        }

        Write-Log "Application '$($obj.displayName)' is not published. Skipping assignment" 2
        @{"Import"=$false}
    }
}

function Start-PreDeleteApplications
{
    if($obj.'@odata.type' -eq "#microsoft.graph.microsoftStoreForBusinessApp")
    {
        # Don't delete Microsoft Store for Business Apps
        @{ "Delete" = $false }
    }
}

function Start-PostExportApplications
{
    param($obj, $objectType, $path)
    
    if($global:chkExportScript.IsChecked)
    {
        $fileName = Get-GraphObjectFile $obj $objectType
        $fi = [IO.FileInfo]"$path\$fileName"

        try
        {
            foreach($rule in ($obj.detectionRules | Where '@OData.Type' -eq "#microsoft.graph.win32LobAppPowerShellScriptDetection"))
            {
                if($rule.ScriptContent)
                {
                    [IO.File]::WriteAllBytes(("$path\$($fi.BaseName)_DetectionScript.ps1"), ([System.Convert]::FromBase64String($rule.ScriptContent)))
        
                }
            }

            foreach($rule in $obj.requirementRules)
            {
                if($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppPowerShellScriptRequirement")
                {
                    if($rule.ScriptContent)
                    {
                        [IO.File]::WriteAllBytes(("$path\$($fi.BaseName)_RequirementScript.ps1"), ([System.Convert]::FromBase64String($rule.ScriptContent)))
                    }
                }
            }
        }
        catch
        {
            Write-LogError "Failed to export scripts" $_.Exception
        }
    }

    Save-Setting "Intune" "ExportAppFile" $global:chkExportApplicationFile.IsChecked
    if($global:chkExportApplicationFile.IsChecked)
    {
        $encryptionSource = Get-SettingValue "EMIntuneAppDownloadFolder" (Get-SettingValue "EMIntuneAppPackages")
        $pkgPath = $path 

        if($pkgPath)
        {
            Write-Status "Download file"

            $exportFile = $pkgPath + "\$($obj.FileName).encrypted"
            $contentFileObj = Start-DownloadAppContent $obj $exportFile -GetContentFileInfoOnly
            $encryptionFile = Find-AppEncryptionFile $obj $contentFileObj $encryptionSource            
            if($encryptionFile -and [IO.File]::Exists($encryptionFile))
            {
                Start-DownloadFile $contentFileObj.azureStorageUri $exportFile

                if([IO.File]::Exists($exportFile))
                {
                    Write-Status "Decrypting file"
                    $encryptionInfo = ConvertFrom-Json (Get-Content -Path $encryptionFile -Raw)
                    if($encryptionInfo.fileEncryptionInfo)
                    {
                        $encryptionInfo = $encryptionInfo.fileEncryptionInfo
                    }                    
                    $destination = $pkgPath + "\$($obj.FileName)"
                    Start-DecryptFile $exportFile $destination $encryptionInfo.encryptionKey $encryptionInfo.initializationVector
                }

                try { [IO.File]::Delete($exportFile) }
                catch {
                    Write-LogError "Failed to delete exported encrypted file" $_.Exception
                }
            }
            else
            {
                Write-Log "Cound not file encryption file"
            }
        }
    }
}

function Start-PostListApplications
{
    param($objList, $objectType)

    foreach($obj in ($objList | Where { $_.Object."@OData.Type" -eq "#microsoft.graph.winGetApp"}))
    {
        if($obj.Object.packageIdentifier -like "9*")
        {
            $installerType = "UWP"
        }
        elseif($obj.Object.packageIdentifier -like "X*")
        {
            $installerType = "Win32"
        }
        else
        {
            $objName = Get-GraphObjectName $obj.Object $objectType
            Write-Log "Unknown package identifier for app $($objName): $($obj.Object.packageIdentifier)" 2
            $installerType = "Unknown"
        }
        $obj.Object | Add-Member -MemberType NoteProperty -Name "InstallerType" -Value $installerType
    }
    $objList   
}

function Add-ScriptExportApplications
{
    param($form, $buttonPanel, $index = 0)

    Add-ScriptExportExtensions $form $buttonPanel $index

    $ctrl = $form.FindName("chkExportApplicationFile")
    if(-not $ctrl)
    {
        $xaml =  @"
<StackPanel $($global:wpfNS) Orientation="Horizontal" Margin="0,0,5,0">
<Label Content="Export application file" />
<Rectangle Style="{DynamicResource InfoIcon}" ToolTip="Export the application file. Note: Application file will only be exported if ecryption file is found." />
</StackPanel>
"@
        $label = [Windows.Markup.XamlReader]::Parse($xaml)

        $global:chkExportApplicationFile = [System.Windows.Controls.CheckBox]::new()
        $global:chkExportApplicationFile.IsChecked = ((Get-Setting "Intune" "ExportAppFile" "false") -eq "true")
        $global:chkExportApplicationFile.VerticalAlignment = "Center" 
        $global:chkExportApplicationFile.Name = "chkExportApplicationFile" 

        @($label, $global:chkExportApplicationFile)
    }    
}

function Start-PostGetApplications {
    param($obj, $objectType)

    if($obj.Object.dependentAppCount -is [Int] -and ($obj.Object.dependentAppCount -gt 0 -or $obj.Object.supersededAppCount -gt 0)) {
        $relationships = (Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$($obj.Id)/relationships?`$filter=targetType%20eq%20microsoft.graph.mobileAppRelationshipType%27child%27").value
        $dependencyApps = @()
        $supersededApps = @()
        foreach ($rel in $relationships) {
            if ($rel."@odata.type" -eq "#microsoft.graph.mobileAppDependency") {
                $dependencyApps += "$($rel.targetDisplayName)|!|$($rel.targetDisplayVersion)|!|$($rel.targetId)|!|$($rel.dependencyType)"
            }
            elseif ($rel."@odata.type" -eq "#microsoft.graph.mobileAppSupersedence") {
                $supersededApps += "$($rel.targetDisplayName)|!|$($rel.targetDisplayVersion)|!|$($rel.targetId)|!|$($rel.supersedenceType)"
            }
        }
        if ($dependencyApps.Count -gt 0) {
            $obj.Object | Add-Member -MemberType NoteProperty -Name "#CustomRefDependency" -Value ($dependencyApps -join "|*|")
        }
        
        if ($supersededApps.Count -gt 0) {
            $obj.Object | Add-Member -MemberType NoteProperty -Name "#CustomRefSupersedence" -Value ($supersededApps -join "|*|")
        }
    }
}

function Start-PostImportApplications
{
    param($obj, $objectType, $file)

    #$tmpObj = Get-GraphObjectFromFile $file
}

function Start-PostFilesImportApplications
{
    param($objType, $importedObjects, $importedFiles)
 
    $refObjects = $importedFiles | Where { $null -ne $_.Object."#CustomRefDependency" -or $null -ne $_.Object."#CustomRefSupersedence" }
    
    if(($refObjects | measure).Count -gt 0)
    {
        Write-Log "Applicetions with Dependency or Supersedence detected"
        foreach($file in $refObjects)
        {
            Add-ApplicationReferences $file.ImportedObject $file.Object
        }    
    }
}

function local:Add-ApplicationReferences 
{
    param($obj, $fileObj)

    if($fileObj."#CustomRefDependency" -or $fileObj."#CustomRefSupersedence")
    {
        Write-Log "Adding app references for $($obj.displayName)"

        $depAppsInfo = $fileObj."#CustomRefDependency"
        $supAppsInfo = $fileObj."#CustomRefSupersedence"

        $releationShips = [PSCustomObject]@{
            relationships = @()
        }

        if($depAppsInfo)
        {
            foreach($depApp in ($depAppsInfo -split "[|][*][|]"))
            {
                $appName, $appVer, $appId, $appType = $depApp -split "[|][!][|]"
                if(-not $appName -or -not $appVer)
                {
                    Write-Log "Could not get Name and Version from string: $appApp" 2
                    continue
                }
                $tmpApps = (Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps?`$filter=displayName eq '$appName'").value
                if(-not $tmpApps)
                {
                    Write-Log "No application found with name $appName" 2
                    continue
                }
                $tmpApp = $tmpApps | Where displayVersion -eq $appVer
                if(-not $tmpApp)
                {
                    Write-Log "No $appName application found with version $appVer" 2
                    continue
                }
                elseif(-not ($tmpApp | measure).Count -gt 1)
                {
                    Write-Log "Multiple $appName application found with version $appVer" 2
                    continue
                }
                Write-Log "Add $appName ($appVer) to Dependency list"
                $releationShips.relationships += [PSCustomObject]@{
                    "@odata.type" = "#microsoft.graph.mobileAppDependency"
                    targetId = $tmpApp.Id
                    dependencyType = $appType
                }
            }
        }

        if($supAppsInfo)
        {
            foreach($suppApp in ($supAppsInfo -split "[|][*][|]"))
            {
                $appName, $appVer, $appId, $appType = $suppApp -split "[|][!][|]"
                if(-not $appName -or -not $appVer)
                {
                    Write-Log "Could not get Name and Version from string: $appApp" 2
                    continue
                }
                $tmpApps = (Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps?`$filter=displayName eq '$appName'").value
                if(-not $tmpApps)
                {
                    Write-Log "No application found with name $appName" 2
                    continue
                }
                $tmpApp = $tmpApps | Where displayVersion -eq $appVer
                if(-not $tmpApp)
                {
                    Write-Log "No $appName application found with version $appVer" 2
                    continue
                }
                elseif(-not ($tmpApp | measure).Count -gt 1)
                {
                    Write-Log "Multiple $appName application found with version $appVer" 2
                    continue
                }
                Write-Log "Add $appName ($appVer) to Supersedence list"
                $releationShips.relationships += [PSCustomObject]@{
                    "@odata.type" = "#microsoft.graph.mobileAppSupersedence"
                    targetId = $tmpApp.Id
                    supersedenceType = $appType
                }
            }
        }

        if($releationShips.relationships.Count -gt 0)
        {
            $json = Update-JsonForEnvironment (ConvertTo-Json $releationShips -Depth 20)

            Write-Log "Update app references"
            Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$($obj.Id)/updateRelationships" -Method "POST" -Body $json
        }
    }    
}

#endregion

#region Group Policy/Administrative Templates functions
function Get-GPOObjectSettings
{
    param($GPOObj)

    $gpoSettings = @()

    if ($GPOObj.policyConfigurationIngestionType -eq "unknown") {
        $tmpObj = (Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations?`$filter=id eq '$($GPOObj.id)'").value[0]
        if ($tmpObj.policyConfigurationIngestionType) {
            $GPOObj.policyConfigurationIngestionType = $tmpObj.policyConfigurationIngestionType
        }
    }

    # Get all configured policies in the Administrative Templates profile 
    $GPODefinitionValues = Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations/$($GPOObj.id)/definitionValues?`$expand=definition" -ODataMetadata "skip"
    foreach($definitionValue in $GPODefinitionValues.value)
    {
        # Get presentation values for the current settings (with presentation object included)
        $presentationValues = Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations/$($GPOObj.id)/definitionValues/$($definitionValue.id)/presentationValues?`$expand=presentation"  -ODataMetadata "skip"

        # Set base policy settings
        $obj = @{
                "enabled" = $definitionValue.enabled
                "definition@odata.bind" = "$($global:graphURL)/deviceManagement/groupPolicyDefinitions('$($definitionValue.definition.id)')"
                }

        if($definitionValue.definition.categoryPath)
        {
            $obj.Add("#Definition_Id", $definitionValue.definition.id)
            $obj.Add("#Definition_displayName", $definitionValue.definition.displayName)
            $obj.Add("#Definition_classType", $definitionValue.definition.classType)
            $obj.Add("#Definition_categoryPath", $definitionValue.definition.categoryPath)            
        }                

        if($presentationValues.value)
        {
            # Policy presentation values set e.g. a drop down list, check box, text box etc.
            $obj.presentationValues = @()                        
            
            foreach ($presentationValue in $presentationValues.value) 
            {
                # Add presentation@odata.bind property that links the value to the presentation object
                $presentationValue | Add-Member -MemberType NoteProperty -Name "presentation@odata.bind" -Value "$($global:graphURL)/deviceManagement/groupPolicyDefinitions('$($definitionValue.definition.id)')/presentations('$($presentationValue.presentation.id)')"

                if($definitionValue.definition.categoryPath)
                {
                    $presentationValue | Add-Member -MemberType NoteProperty -Name "#Presentation_Id" -Value $presentationValue.presentation.id
                    $presentationValue | Add-Member -MemberType NoteProperty -Name "#Presentation_Label" -Value $presentationValue.presentation.label
                }
                #Remove presentation object so it is not included in the export
                Remove-ObjectProperty $presentationValue "presentation"
                
                #Optional removes. Import will igonre them
                Remove-ObjectProperty $presentationValue "id"
                Remove-ObjectProperty $presentationValue "lastModifiedDateTime"
                Remove-ObjectProperty $presentationValue "createdDateTime"

                # Add presentation value to the list
                $obj.presentationValues += $presentationValue
            }
        }
        $gpoSettings += $obj
    }
    $gpoSettings
}

function Import-GPOSetting
{
    param($obj, $settings)
    
    if($obj)
    {
        Write-Status "Import settings for $($obj.displayName)"

        $hasCustomADMX = $null -ne ($settings | Where { $null -ne $_.'#Definition_categoryPath' })

        if($hasCustomADMX)
        {
            Write-Status "Import custom ADMX settings"
            if(-not $script:CustomADMXDefinitions)
            {
                $tmpCustomCategories = Invoke-GraphRequest -Url "deviceManagement/groupPolicyCategories?`$expand=definitions(`$select=id, displayName, categoryPath, classType)&`$select=id, displayName&`$filter=ingestionSource eq 'custom'" -ODataMetadata "Minimal"
                if($tmpCustomCategories.Value)
                {
                    $script:CustomADMXDefinitions = @{}
                    foreach($tmpCat in $tmpCustomCategories.Value)
                    {
                        foreach($tmpDef in $tmpCat.definitions)
                        {
                            $key = ($tmpDef.displayName + $tmpDef.categoryPath + $tmpDef.classType).ToLower()
                            $val = [PSCustomObject]@{
                                Definition = $tmpDef
                                Category = $tmpCat
                                Presentations = $null
                            }
                            try {
                                $script:CustomADMXDefinitions.Add($key, $val)                                
                            }
                            catch {
                                Write-Log "Failed to add '$($tmpDef.displayName)' in category '$($tmpDef.categoryPath)' of class $($tmpDef.classType)" 3
                            }
                        }
                    }
                }
            }
        }        
        
        foreach($setting in $settings)
        {
            if($setting.'#Definition_categoryPath' -and $script:CustomADMXDefinitions -is [HashTable] -and $script:CustomADMXDefinitions.Count -gt 0)
            {
                $defVal = $null
                $key = ($setting.'#Definition_displayName' + $setting.'#Definition_categoryPath' + $setting.'#Definition_classType').ToLower()
                if($key -and $script:CustomADMXDefinitions.ContainsKey($key))
                {
                    $defVal = $script:CustomADMXDefinitions[$key]
                }
                elseif($key)
                {
                    Write-Log "No custom ADMX definitiona found for setting $($setting.'#Definition_displayName')" 2                    
                }
                else
                {
                    Write-Log "Setting $($setting.'#Definition_displayName') does not have information to be imported in the environment"
                }

                if($defVal)
                {
                    $setting.'definition@odata.bind' = $setting.'definition@odata.bind' -replace $setting.'#Definition_Id', $defVal.Definition.Id
                    if(($setting.presentationValues | measure).Count -gt 0)
                    {
                        if(-not $defVal.Presentations)
                        {
                            $tmpPresentation = Invoke-GraphRequest -Url "deviceManagement/groupPolicyDefinitions/$($defVal.Definition.Id)/presentations" -ODataMetadata "Minimal"
                            if($tmpPresentation.value)
                            {
                                foreach($settingPresentation in $setting.presentationValues)
                                {
                                    $tmpPresentationVal = $tmpPresentation.value | Where label -eq $settingPresentation.'#Presentation_Label'
                                    if($tmpPresentationVal)
                                    {
                                        $settingPresentation.'presentation@odata.bind' = $settingPresentation.'presentation@odata.bind' -replace $setting.'#Definition_Id', $defVal.Definition.Id
                                        $settingPresentation.'presentation@odata.bind' = $settingPresentation.'presentation@odata.bind' -replace $settingPresentation.'#Presentation_Id', $tmpPresentationVal.Id
                                    }
                                    else
                                    {
                                        Write-Log "Could not find a presentation value with label $($settingPresentation.'#Presentation_Label'). Setting will not be configured" 2
                                        continue
                                    }
                                }
                            }
                            else
                            {
                                Write-Log "Could not find presentation for setting $($settingPresentation.'#Presentation_Label'). Setting will not be configured." 2
                                continue
                            }
                        }
                    }
                }
                else
                {
                    Write-Log "Settings might not be available if imported in another environment" 3
                }
            }
            elseif($setting.'#Definition_categoryPath')
            {
                Write-Log "Custom AMDX settings cannot be imported without ADMX file imported. Definitions not found" 2
                continue
            }

            Start-GraphPreImport $setting

            if($true) 
            {
                foreach($tmpProp in (($setting.PSObject.Properties | Where Name -like "#*").Name))
                {
                    Remove-Property $setting $tmpProp
                }
                
                foreach($settingPresentation in $setting.presentationValues)
                {
                    foreach($tmpProp in (($settingPresentation.PSObject.Properties | Where Name -like "#*").Name))
                    {
                        Remove-Property $settingPresentation $tmpProp
                    }
                }
            }

            # Import each setting for the Administrative Template profile
            Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations/$($obj.id)/definitionValues" -Content (ConvertTo-Json $setting -Depth 20) -HttpMethod POST | Out-Null
        }
    }
}

function Start-PostExportAdministrativeTemplate
{
    param($obj, $objectType, $path)

    $fileName = (Get-GraphObjectName $obj $objectType).Trim('.')
    if((Get-SettingValue "AddIDToExportFile") -eq $true -and $obj.Id)
    {
        $fileName = ($fileName + "_" + $obj.Id)
    }
    
    if($obj.definitionValues)
    {
        $settings =  $obj.definitionValues
    }
    else
    {
        $settings = Get-GPOObjectSettings $obj
    }

    $fileName = "$path\$((Remove-InvalidFileNameChars $fileName))_Settings.json"
    Save-GraphObjectToFile $settings $fileName
}

function Start-PostCopyAdministrativeTemplate
{
    param($objCopyFrom, $objNew, $objectType)

    $settings = Get-GPOObjectSettings $objCopyFrom
    if($settings)
    {
        Import-GPOSetting $objNew $settings
    }
}

function Start-PostFileImportAdministrativeTemplate
{
    param($obj, $objectType, $file)

    $settings = Get-EMSettingsObject $obj $objectType $file -settingsProperty "definitionValues" -SettingsArray
    if($settings)
    {
        $tmpObj = Get-GraphObjectFromFile $file

        Import-GPOSetting $obj $settings
    }    
}

function Start-LoadAdministrativeTemplate
{
    param($fileName)

    if(-not $fileName) { return $null }

    $fi = [IO.FileInfo]$fileName
    if($fi.Exists -eq $false) { return }
 
    $obj = Get-GraphObjectFromFile $fi.FullName

    if($obj.definitionValues)
    {
        return $obj
    }

    $settingsFile = $fi.DirectoryName + "\" + $fi.BaseName + "_Settings.json"

    if([IO.File]::Exists($settingsFile))
    {
        $definitionValues = Get-GraphObjectFromFile $settingsFile

        $obj | Add-Member Noteproperty -Name "definitionValues" -Value $definitionValues -Force  
    }
    $obj
}

function Start-PostGetAdministrativeTemplate
{
    param($obj, $objectType)

    $definitionValues = Get-GPOObjectSettings $obj.Object
    if($definitionValues)
    {
        $obj.Object | Add-Member Noteproperty -Name "definitionValues" -Value $definitionValues -Force 
    }    
    <#
    # Leave for now. This only loads the configured definition values and not the values specified.
    # That would require enumerating each definition value which takes time. 
    $definitionValues = (Invoke-GraphRequest "deviceManagement/groupPolicyConfigurations('$($obj.Id)')/definitionValues?`$expand=definition(`$select=id,classType,displayName,policyType,groupPolicyCategoryId)" -ODataMetadata "minimal").value

    if($definitionValues)
    {
        $obj.Object | Add-Member Noteproperty -Name "definitionValues" -Value $definitionValues -Force 
    }
    #>
}

function Start-PreImportAdministrativeTemplate
{
    param($obj, $objectType, $file, $assignments)

    
}

#endregion

#region Policy Sets function

function Start-PreImportAssignmentsPolicySets
{
    param($obj, $objectType, $file, $assignments)

    @{"API"="$($objectType.API)/$($obj.Id)/Update"}
}

function Start-PreImportPolicySets
{
    param($obj, $objectType)

    @("items@odata.context","status","errorCode") | foreach { Remove-Property $obj $_ }

     # Properties to keep for items
    $keepProperties = @("@odata.type","payloadId","intent","settings")
    foreach($item in $obj.Items)
    {
        foreach($prop in ($item.PSObject.Properties | Where {$_.Name -notin $keepProperties}))
        {
            Remove-Property $item $prop.Name
        }
        #@("itemType","displayName","status","errorCode") | foreach { Remove-Property $item $_ }
    }
}

function Start-PreUpdatePolicySets
{
    param($obj, $objectType, $curObject, $fromObj)

    Start-PreImportPolicySets $obj $objectType

    $curObject = Get-GraphObject $curObject.Object $objectType

    # Update ref object in the json
    # Used when importing in a different environment
    $jsonObj = ConvertTo-Json $obj -Depth 15
    $updateObj = Update-JsonForEnvironment $jsonObj | ConvertFrom-Json

    $addedItems = @()
    $updatedItems = @()
    $deletedItems = @()

    foreach($item in $updateObj.items)
    {
        if(($curObject.Object.items | Where payloadId -eq $item.payloadId))
        {
            $updatedItems += $item
        }
        else
        {
            $addedItems += $item
        }
    }

    foreach($item in $curObject.Object.items)
    {
        if(-not ($updateObj.Items | Where payloadId -eq $item.payloadId))
        {
            $deletedItems += $item.id
        }
    }

    $updateItemObj = [PSCustomObject]@{
        addedPolicySetItems = $addedItems
        deletedPolicySetItems = $deletedItems
        updatedPolicySetItems = $updatedItems
    }

    Write-Log "Update Policy Set items. Add: $($addedItems.Count), Update: $($updatedItems.Count), Delete: $($deletedItems.Count)"

    $updateApi = "/deviceAppManagement/policySets/$($curObject.Object.Id)/update"
    $json = $updateItemObj | ConvertTo-Json -Depth 15

    Invoke-GraphRequest -Url $updateApi -HttpMethod "POST" -Content $json
    Remove-Property $obj "items"
}

function Update-EMPolicySetAssignment
{
    param($assignment, $sourceObject, $newObject, $objectType)

    $api = "/deviceAppManagement/policySets/$($assignment.SourceId)?`$expand=assignments,items"

    $psObj = Invoke-GraphRequest -Url $api -ODataMetadata "Minimal"

    if(-not $psObj)
    {
        return
    }

    $curItem = $psObj.Items | Where payloadId -eq $sourceObject.Id

    if(-not $curItem)
    {
        return
    }

    $api = "/deviceAppManagement/policySets/$($assignment.SourceId)/update"

    $curItemClone = $curItem | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $newItem = $curItem | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $newItem.payloadId = $newObject.Id
    if($newItem.guidedDeploymentTags -is [String] -and [String]::IsNullOrEmpty($newItem.guidedDeploymentTags))
    {
        $newItem.guidedDeploymentTags = @()
    }

    $keepProperties = @('@odata.type','payloadId','Settings','guidedDeploymentTags') 
    #itemType? e.g. #microsoft.graph.iosManagedAppProtection
    #priority?

    foreach($prop in ($newItem.PSObject.Properties | Where {$_.Name -notin $keepProperties}))
    {
        Remove-Property $newItem $prop.Name
    }

    $update = @{}
    $update.Add('addedPolicySetItems',@($newItem))
    $update.Add('updatedPolicySetItems', @())
    $update.Add('deletedPolicySetItems',@($curItemClone.Id))

    $json = $update | ConvertTo-Json -Depth 20

    Write-Log "Update PolicySet $($psObj.displayName) - Replace: $((Get-GraphObjectName $newObject $objectType))"

    Invoke-GraphRequest -Url $api -HttpMethod "POST" -Content $json 
}

function Start-PostListPolicySets
{
    param($objList, $objectType)

    foreach($obj in $objList)
    {
        $obj | Add-Member -MemberType NoteProperty -Name "IsAssigned" -Value ($obj.Object.status -ne "notAssigned")     
    }
    $objList    
}
#endregion

#endregion Locations
function Start-PreImportLocations
{
    param($obj, $objectType)

    if($obj.uniqueName)
    {
        $arr = $obj.uniqueName.Split('_')
        if($arr.Length -ge 3)
        {
            # Locations requires a unique name so generate a new guid and change the uniqueName property
            $obj.uniqueName = ($obj.uniqueName.Substring(0,$obj.uniqueName.Length-$arr[-1].Length) + [Guid]::NewGuid().Tostring("n"))
        }
    }
}
#endregion

#region RoleDefinitions
function Start-PostExportRoleDefinitions
{
    param($obj, $objectType, $path)

    $fileName = (Get-GraphObjectName $obj $objectType).Trim('.')
    if((Get-SettingValue "AddIDToExportFile") -eq $true -and $obj.Id)
    {
        $fileName = ($fileName + "_" + $obj.Id)
    }
    $tmpObj = $null
    $fileName = "$path\$((Remove-InvalidFileNameChars $fileName)).json"
    if([IO.File]::Exists($fileName))
    {
        $tmpObj = Get-GraphObjectFromFile $fileName
    }
    else
    {
        Write-Log "File not found: $fileName. Could not get role assignments" 3
    }

    if(($tmpObj.RoleAssignments | measure).Count -gt 0)
    {        
        $roleAssignmentsArr = @()
        foreach($roleAssignment in $tmpObj.RoleAssignments)
        {            
            $raObj = Invoke-GraphRequest -Url "/deviceManagement/roleAssignments/$($roleAssignment.Id)?`$expand=microsoft.graph.deviceAndAppManagementRoleAssignment/roleScopeTags" -ODataMetadata "Minimal"
            if($raObj) 
            {
                foreach($groupId in $raObj.resourceScopes) { Add-GroupMigrationObject $groupId }
                foreach($groupId in $raObj.members) { Add-GroupMigrationObject $groupId }
                $roleAssignmentsArr += $raObj 
            }
        }

        if($roleAssignmentsArr.Count -gt 0)
        {
            $tmpObj.RoleAssignments = $roleAssignmentsArr
            Save-GraphObjectToFile $tmpObj $fileName
        }
    }
}

function Start-PreImportRoleDefinitions
{
    param($obj, $objectType)

    Remove-Property $obj "RoleAssignments"
    Remove-Property $obj "RoleAssignments@odata.context"
}

function Start-PostFileImportRoleDefinitions
{
    param($obj, $objectType, $file)

    $tmpObj = Get-GraphObjectFromFile $file

    $loadedScopeTags = $global:LoadedDependencyObjects["ScopeTags"]
    if(($tmpObj.RoleAssignments | measure).Count -gt 0 -and ($loadedScopeTags | measure).Count -gt 0)
    {
        # Documentation way did not work so use the same way as the portal
        # Should be created with /deviceManagement/roleDefinitions/{roleDefinitionId}/roleAssignments
        foreach($roleAssignment in $tmpObj.RoleAssignments)
        {
            $roleAssignmentObj = New-object PSObject @{ 
                "description" = $roleAssignment.Description
                "displayName"= $roleAssignment.DisplayName
                "members" = $roleAssignment.members
                "resourceScopes" = $roleAssignment.resourceScopes
                "roleDefinition@odata.bind" = "https://graph.microsoft.com/beta/deviceManagement/roleDefinitions('$($obj.Id)')"
                "roleScopeTags@odata.bind" = @()
            }

            foreach($scopeTag in $roleAssignment.roleScopeTags)
            {
                $scopeMigObj = $loadedScopeTags | Where OriginalId -eq $scopeTag.Id
                if(-not $scopeMigObj.Id) { continue }                
                $roleAssignmentObj."roleScopeTags@odata.bind" += "https://graph.microsoft.com/beta/deviceManagement/roleScopeTags('$($scopeMigObj.Id)')"
            }

            # This will update GroupIds
            $json = Update-JsonForEnvironment (ConvertTo-Json $roleAssignmentObj -Depth 20)

            Write-Log "Import Role Assignments"
            Invoke-GraphRequest -Url "/deviceManagement/roleAssignments"  -Body $json -Method "POST"
        }
    }    
}
#endregion

#region SettingsCatalog
function Start-PostExportSettingsCatalog
{
    param($obj, $objectType, $path)

    Add-EMAssignmentsToExportFile $obj $objectType $path
}

function Start-PreUpdateSettingsCatalog
{
    param($obj, $objectType, $curObject, $fromObj)

    @{"Method"="PUT"}
}

function Start-PostGetSettingsCatalog
{
    param($obj, $objectType)

    if(-not $obj.Object.Assignments)
    {
        $url = "$($objectType.API)/$($obj.id)/assignments"
        $assignments = (Invoke-GraphRequest -Url $url).Value
        if($assignments)
        {
            $obj.Object.Assignments = $assignments
        }
    }
}

#endregion

#region Notification functions
function Start-PreImportNotifications
{
    param($obj, $objectType)

    Remove-Property $obj "defaultLocale"
    Remove-Property $obj "localizedNotificationMessages"
    Remove-Property $obj "localizedNotificationMessages@odata.context"
}

function Start-PostFileImportNotifications
{
    param($obj, $objectType, $file)

    $tmpObj = Get-GraphObjectFromFile $file

    foreach($localizedNotificationMessage in $tmpObj.localizedNotificationMessages)
    {
        Start-GraphPreImport $localizedNotificationMessage $objectType
        Invoke-GraphRequest -Url "$($objectType.API)/$($obj.id)/localizedNotificationMessages" -Body ($localizedNotificationMessage | ConvertTo-Json -Depth 20) -Method "POST"
    }
}

function Start-PostCopyNotifications
{
    param($objCopyFrom, $objNew, $objectType)

    foreach($localizedNotificationMessage in $objCopyFrom.localizedNotificationMessages)
    {
        Start-GraphPreImport $localizedNotificationMessage $objectType
        Invoke-GraphRequest -Url "$($objectType.API)/$($objNew.id)/localizedNotificationMessages" -Body ($localizedNotificationMessage | ConvertTo-Json -Depth 20) -Method "POST"
    }
}
#endregion

#region Enrollment Status Page functions
function Start-PreImportESP
{
    param($obj, $objectType)

    if($obj.Priority -eq 0)
    {
        $ret = @{}
        $ret.Add("API","$($objectType.API)/$($obj.Id)")
        $ret.Add("Method","PATCH") # Default profile always exists so update them
        $ret
    }
    else
    {
        Remove-Property $obj "Id"    
    }
}

function Start-PostExportESP
{
    param($obj, $objectType, $path)

    if($obj.Priority -eq 0)
    {
        Save-EMDefaultPolicy $obj $objectType $path
    }
}

function Start-PostListESP
{
    param($objList, $objectType)

    # endswith not working so filter them out
    $objList | Where { $_.Object.'@OData.Type' -eq '#microsoft.graph.windows10EnrollmentCompletionPageConfiguration' }
}
#endregion

#region Enrollment Restriction functions

function Start-PostExportEnrollmentRestrictions
{
    param($obj, $objectType, $path)

    if($obj.Priority -eq 0)
    {
        Save-EMDefaultPolicy $obj $objectType $path        
    }   
}

function Start-PreImportEnrollmentRestrictions
{
    param($obj, $objectType)

    if($obj.Priority -eq 0)
    {
        $ret = @{}
        $ret.Add("API","$($objectType.API)/$($obj.Id)")
        $ret.Add("Method","PATCH") # Default profile always exists so update them
        $ret
    }
    else
    {
        Remove-Property $obj "Id"    
    }

    if($obj.windowsMobileRestriction)
    {
        # Windows Phone operations are no longer supported
        Remove-Property $obj "windowsMobileRestriction" 
    }
}

function Start-PreDeleteEnrollmentRestrictions
{
    param($obj, $objectType)

    if($obj.Priority -eq 0)
    {
        @{ "Delete" = $false }
    }
}

function Start-PreReplaceEnrollmentRestrictions
{
    param($obj, $objectType, $sourceObj, $fromFile)

    if($sourceObj.Priority -eq 0) { @{ "Replace" = $false } }
}

function Start-PostReplaceEnrollmentRestrictions
{
    param($obj, $objectType, $sourceObj, $fromFile)

    if($sourceObj.Priority -eq 0) { return }

    $api = "/deviceManagement/deviceEnrollmentConfigurations/$($obj.id)/setpriority"

    $priority = [PSCustomObject]@{
        priority = $sourceObj.Priority
    }
    $json = $priority | ConvertTo-Json -Depth 20

    Write-Log "Update priority for $($obj.displayName) to $($sourceObj.Priority)"
    Invoke-GraphRequest $api -HttpMethod "POST" -Content $json
}

function Start-PreFilesImportEnrollmentRestrictions
{
    param($objectType, $filesToImport)

    $filesToImport | sort-object -property @{e={$_.Object.priority}} 
}

function Start-PreUpdateEnrollmentRestrictions
{
    param($obj, $objectType, $curObject, $fromObj)

    Remove-Property $obj "priority"
}

function Start-PostListEnrollmentRestrictions
{
    param($objList, $objectType)

    # endswith not working so filter them out
    $objList | Where { 
        ($_.Object.'@OData.Type' -eq '#microsoft.graph.deviceEnrollmentPlatformRestrictionConfiguration' -or 
        $_.Object.'@OData.Type' -eq '#microsoft.graph.deviceEnrollmentLimitConfiguration' -or 
        $_.Object.'@OData.Type' -eq '#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration') -and
        $_.Object.id -notlike "*_PlatformRestrictions" -and $_.Object.platformType -ne "WindowsPhone" -and $_.Object.platformType -ne "AndroidAosp"
    }
}

function Start-PreImportAssignmentsEnrollmentRestrictions
{
    param($obj, $objectType, $file, $assignments)

    if($obj.Priority -eq 0)
    {
        # Skip Assignment for Default Policy
        @{ "Import" = $false }
    }
}

#endregion

#region 
function Start-PostListCoManagementSettings
{
    param($objList, $objectType)

    # endswith not working so filter them out
    $objList | Where { $_.Object.'@OData.Type' -eq '#microsoft.graph.deviceComanagementAuthorityConfiguration' }
}
#endregion

#region ScopeTags
function Start-PostExportScopeTags
{
    param($obj, $objectType, $path)

    Add-EMAssignmentsToExportFile $obj $objectType $path
}

function Start-PostGetScopeTags
{
    param($obj, $objectType)

    $strAPI = "$($objectType.API)/$($obj.Object.Id)/assignments"
    $tmpObj = Invoke-GraphRequest -Url $strAPI

    if(($tmpObj.value | measure).count -gt 0)
    {
        $obj.Object.assignments = $tmpObj.value
    }  
}
#endregion

#region AutoPilot
function Start-PreImportAssignmentsAutoPilot
{
    param($obj, $objectType, $file, $assignments)

    Add-EMAssignmentsToObject $obj $objectType $file $assignments
}

function Start-PreDeleteAutoPilot
{
    param($obj, $objectType)

    Write-Log "Delete AutoPilot profile assignments"

    if(-not $obj.Assignments)
    {
        $tmpObj = (Get-GraphObject $obj $objectType).Object
    }
    else
    {
        $tmpObj = $obj
    }

    foreach($assignment in $tmpObj.Assignments)
    {
        if($assignment.Source -ne "direct") { continue }

        $api = "/deviceManagement/windowsAutopilotDeploymentProfiles/$($obj.Id)/assignments/$($assignment.Id)"

        Invoke-GraphRequest $api -HttpMethod "DELETE"
    }
}

#endregion

#region Health Scripts

function Start-PreDeleteDeviceHealthScripts
{
    param($obj, $objectType)

    if($obj.isGlobalScript -eq $true)
    {
        @{ "Delete" = $false }
    }
}

function Start-PreImportDeviceHealthScripts
{
    param($obj, $objectType, $file, $assignments)

    if($obj.isGlobalScript -eq $true)
    {
        @{ "Import" = $false }
    }
}

function Start-PreUpdateDeviceHealthScripts
{
    param($obj, $objectType, $curObject, $fromObj)

    if($curObject.Object.isGlobalScript -eq $true)
    {
        @{ "Import" = $false }
    }
}

function Start-PostExportDeviceHealthScripts
{
    param($obj, $objectType, $path)

    if($global:chkExportScript.IsChecked)
    {
        $fileName = Get-GraphObjectFile $obj $objectType
        $fi = [IO.FileInfo]"$path\$fileName"

        try
        {
            if($obj.detectionScriptContent)
            {
                [IO.File]::WriteAllBytes(("$path\$($fi.BaseName)_DetectionScript.ps1"), ([System.Convert]::FromBase64String($obj.detectionScriptContent)))
            }

            if($obj.remediationScriptContent)
            {
                [IO.File]::WriteAllBytes(("$path\$($fi.BaseName)_RemediationScript.ps1"), ([System.Convert]::FromBase64String($obj.remediationScriptContent)))
            }
        }
        catch
        {
            Write-LogError "Failed to export scripts" $_.Exception
        }
    }
}

#endregion

#region Generic functions

function Save-EMDefaultPolicy
{
    param($obj, $objectType, $path)

    if($obj.Priority -eq 0)
    {
        try
        {
            $fileName = $obj.Id.Split('_')[1]

            if($fileName)
            {
                $oldFile = "$path\$((Get-GraphObjectName $obj $objectType)).json"
                if([IO.File]::Exists($oldFile))
                {
                    # Clean up from old version of the script that used the wrong name for Default policies
                    try { [IO.File]::Delete($oldFile) | Out-Null } Catch {}
                }
                Save-GraphObjectToFile $obj "$path\$((Remove-InvalidFileNameChars $fileName)).json"
            }
        }
        catch {}
    }   
}
function Get-EMSettingsObject
{
    param($obj, $objectType, $file, $settingsProperty = "settings", [switch]$SettingsArray)

    if($obj.$settingsProperty) { return $obj.$settingsProperty }

    $fi = [IO.FileInfo]$file
    if($fi.Exists)
    {
        # Settings property removed during import so lets try exported file first
        $tmpObj = Get-GraphObjectFromFile $fi.FullName
        if($SettingsArray -eq $true)
        {
            # Only the an array of settings is expected
            return $tmpObj.$settingsProperty
        }
        else
        {
            if($tmpObj.$settingsProperty) 
            {
                # A property with the an array of settings is expected
                return ([PSCustomObject]@{
                    $settingsProperty = $tmpObj.$settingsProperty
                })
            } 
        }

        Write-Log "Settings not included in export file. Try import from _Settings.json file" 2
        $settingsFile = $fi.DirectoryName + "\" + $fi.BaseName + "_Settings.json"
        $fiSettings = [IO.FileInfo]$settingsFile
        if($fiSettings.Exists -eq $false)
        {
            Write-Log "Settings file '$($fiSettings.FullName)' was not found" 2
            return
        }        
        Get-GraphObjectFromFile $fiSettings.FullName
    }
    else
    {
        Write-Log "Settings not included in export file and _Settings.json file is missing." 3
    }
}

function Add-EMAssignmentsToExportFile
{
    param($obj, $objectType, $path, $Url = "")

    if($global:chkExportAssignments.IsChecked -ne $true) { return }

    $fileName = (Get-GraphObjectName $obj $objectType).Trim('.')
    if((Get-SettingValue "AddIDToExportFile") -eq $true -and $obj.Id)
    {
        $fileName = ($fileName + "_" + $obj.Id)
    }
    $fileName = "$path\$((Remove-InvalidFileNameChars $fileName)).json"
    if([IO.File]::Exists($fileName) -eq $false)
    {
        Write-Log "File not found: $fileName. Could not add assignments to file" 3
        return
    }
    
    $tmpObj = Get-GraphObjectFromFile $fileName

    if(-not $url)
    {
        $url = "$($objectType.API)/$($obj.id)/assignments"
    }
    $assignments = (Invoke-GraphRequest -Url $url -ODataMetadata "Minimal").Value
    if($assignments)
    {
        if(-not ($tmpObj.PSObject.Properties | Where Name -eq "assignments"))
        {
            $tmpObj | Add-Member -MemberType NoteProperty -Name "assignments" -Value $assignments
        }
        else
        {
            $tmpObj.Assignments = $assignments
        }
        Save-GraphObjectToFile $tmpObj $fileName
    }
}

function Add-EMAssignmentsToObject
{
    param($obj, $objectType, $file, $assignments)

    # AutoPilot and TaC are using assignments and not assign like other object types
    $api = "$($objectType.API)/$($obj.Id)/assignments"

    # These profiles don't support importing of multiple assignments with { "assignment" [...]}
    # Each assignment must be imported separately 

    foreach($assignment in $assignments)
    {
        if($assignment.Source -and $assignment.Source -ne "direct") { continue }

        foreach($prop in $assignment.PSObject.Properties)
        {
            if($prop.Name -in @("Target")) { continue }
            Remove-Property $assignment $prop.Name
        }

        foreach($prop in $assignment.target.PSObject.Properties)
        {
            if($prop.Name -in @("@odata.type","groupId")) { continue }
            Remove-Property $assignment.target $prop.Name
        }

        $json = Update-JsonForEnvironment ($assignment | ConvertTo-Json -Depth 20)
        Invoke-GraphRequest -Url $api -Body $json -Method "POST" | Out-Null
    }
    @{"Import"=$false}
}

#endregion

#region Mac Custom Scripts

function Start-PreUpdateMacCustomAttributes
{
    param($obj, $objectType, $curObject, $fromObj)

    foreach($prop in @('customAttributeName','customAttributeType','displayName'))
    {
        Remove-Property $obj $prop
    }
}

#endregion

#region Mac Feature Updates
function Start-PreUpdateFeatureUpdates
{
    param($obj, $objectType, $curObject, $fromObj)

    foreach($prop in @('deployableContentDisplayName','endOfSupportDate'))
    {
        Remove-Property $obj $prop
    }
}
#endregion

#region Conditional Access
function Add-ConditionalAccessImportExtensions
{
    param($form, $buttonPanel, $index = 0)

    $xaml =  @"
<StackPanel $($global:wpfNS) Orientation="Horizontal" Margin="0,0,5,0">
<Label Content="Conditional Access State" />
<Rectangle Style="{DynamicResource InfoIcon}" ToolTip="Specifies the enable state of Conditional Access policies" />
</StackPanel>
"@
    $label = [Windows.Markup.XamlReader]::Parse($xaml)

    $CAStates = @()
    $CAStates += [PSCustomObject]@{
        Name  = "As Exported - Change On to Report-only"
        Value = "AsExportedReportOnly"
    }

    $CAStates += [PSCustomObject]@{
        Name = "As Exported"
        Value = "AsExported"
    }

    $CAStates += [PSCustomObject]@{
        Name = "Report-only"
        Value = "enabledForReportingButNotEnforced"
    }

    $CAStates += [PSCustomObject]@{
        Name = "On"
        Value = "enabled"
    }
    
    $CAStates += [PSCustomObject]@{
        Name = "Off"
        Value = "disabled"
    }    

    $global:cbImportCAState = [System.Windows.Controls.ComboBox]::new()
    $global:cbImportCAState.DisplayMemberPath = "Name"
    $global:cbImportCAState.SelectedValuePath = "Value"
    $global:cbImportCAState.ItemsSource = $CAStates
    $global:cbImportCAState.SelectedValue = "disabled"
    $global:cbImportCAState.Margin="0,5,0,0"
    $global:cbImportCAState.HorizontalAlignment="Left"
    $global:cbImportCAState.Width=250
    $global:cbImportCAState.Name = "cbImportCAState"

    @($label, $global:cbImportCAState)
}

function Start-PreImportConditionalAccess
{
    param($obj, $objectType, $file, $assignments)

    if ($global:cbImportCAState.SelectedValue -and $global:cbImportCAState.SelectedValue -ne "AsExported") {
        if ($global:cbImportCAState.SelectedValue -eq "AsExportedReportOnly" -and $obj.state -eq "enabled") {
            Write-Log "Change Enabled policy to Report-only"
            $obj.state = "enabledForReportingButNotEnforced"
        }
        else {
            $obj.state = $global:cbImportCAState.SelectedValue
        }
    }

    if($obj.grantControls.authenticationStrength)
    {
        $obj.grantControls.operator = "AND"
        $tmpObj = Get-GraphObjectFromFile $file

        $authSetting = [PSCustomObject]@{
            id = $tmpObj.grantControls.authenticationStrength.id
        }
        $obj.grantControls.authenticationStrength = $authSetting
    }

    if($obj.sessionControls.disableResilienceDefaults -eq $false)
    {
        $obj.sessionControls.disableResilienceDefaults = $null
    }
}

function Start-PostExportConditionalAccess
{
    param($obj, $objectType, $path)

    $ids = @()
    foreach($id in ($obj.conditions.users.includeGroups + $obj.conditions.users.excludeGroups))
    {
        if($id -in $ids) { continue }
        elseif($id -eq "GuestsOrExternalUsers") { continue }
        elseif($id -eq "All") { continue }
        elseif($id -eq "None") { continue }
        
        $ids += $id
        Add-GraphMigrationObject $id "/groups" "Group"
    }
    
    foreach($id in ($obj.conditions.users.includeUsers +$obj.conditions.users.excludeUsers))
    {
        if($id -in $ids) { continue }
        elseif($id -eq "GuestsOrExternalUsers") { continue }
        elseif($id -eq "All") { continue }
        elseif($id -eq "None") { continue }
        
        $ids += $id
        Add-GraphMigrationObject $id "/users" "User"
    }    

    <#
    $roleIds = @()
    foreach($id in ($obj.conditions.users.includeRoles + $obj.conditions.users.excludeRoles))
    {
        if($id -in $ids) { continue }
        $roleIds += $id
    }
    #>
}
#endregion

#region Terms of use
function Start-PreImportTermsOfUse
{
    param($obj, $objectType, $file, $assignments)

    $pkgPath = Get-SettingValue "EMIntuneAppPackages"

    if(-not $pkgPath -or [IO.Directory]::Exists($pkgPath) -eq $false) 
    {
        Write-Log "Intune app directory is either missing or does not exist" 2        
    }

    try
    {
        $fi = [IO.FileInfo]$file
    } catch {}

    foreach($file in $obj.Files)
    {
        $pdfFile = $null

        if($fi.Directory.FullName)
        {
            $pdfFile = "$($fi.Directory.FullName)\$($file.fileName)"
        }
        
        if($null -eq $pdfFile -or [IO.File]::Exists($pdfFile) -eq $false) 
        {
            $pdfFile = "$($pkgPath)\$($file.fileName)"
        }        

        if([IO.File]::Exists($pdfFile) -eq $false) 
        {
            Write-Log "Terms of use file $($file.fileName) not found. The Terms of Use object will not be imported." 2
            @{"Import" = $false}
            return 
        }

        Write-Log "Add file data: $pdfFile"

        $bytes = [IO.File]::ReadAllBytes($pdfFile)        
        $file.fileData = [PSCustomObject]@{
            data = [Convert]::ToBase64String($bytes)
        }
    }
}

function Start-PostExportTermsOfUse
{
    param($obj, $objectType, $path)

    foreach($file in $obj.Files)
    {
        $url = "agreements/$($obj.id)/file/localizations('$($file.id)')/fileData/data"
        $data = (Invoke-GraphRequest -Url $url -ODataMetadata "Minimal").Value
        if($data)
        {
            Write-Log "Save file $($file.FileName)"
            $fileName = "$path\$($file.FileName)" 
            [IO.File]::WriteAllBytes($fileName, [System.Convert]::FromBase64String($data))
        } 
    }
}

#endregion

#region ADMXFiles

function Start-PreFilesImportADMXFiles
{
    param($objectType, $filesToImport)

    $filesToImport | sort-object -property @{e={$_.Object.lastModifiedDateTime}} 
}

function Start-PreImportADMXFiles
{
    param($obj, $objectType, $file, $assignments)

    $pkgPath = Get-SettingValue "EMIntuneAppPackages"

    if(-not $pkgPath -or [IO.Directory]::Exists($pkgPath) -eq $false) 
    {
        Write-Log "Intune app directory is either missing or does not exist" 2
        $pkgPath = $null
    }

    try
    {
        $fi = [IO.FileInfo]$file
    } catch {}

    $admxFile = $null

    if($fi.Directory.FullName)
    {
        $admxFile = "$($fi.Directory.FullName)\$($obj.fileName)"
        $admlFile = "$($fi.Directory.FullName)\$([io.path]::GetFileNameWithoutExtension($obj.fileName)).adml"
    }
    
    if($null -ne $pkgPath -and ($null -eq $admxFile -or [IO.File]::Exists($admxFile) -eq $false -or [IO.File]::Exists($admxFile) -eq $false))
    {
        Write-Log "$($obj.fileName) not foud in Export folder. Look in package path: $pkgPath"
        $admxFile = "$($pkgPath)\$($obj.fileName)"
        $admlFile = "$($pkgPath)\$([io.path]::GetFileNameWithoutExtension($obj.fileName)).adml"
    }        

    if([IO.File]::Exists($admxFile) -eq $false) 
    {
        Write-Log "ADMX (or ADML) file $($obj.fileName) not found. The ADMXFile object will not be imported." 2
        @{"Import" = $false}
        return 
    }

    #$bytes = [IO.File]::ReadAllBytes($admxFile)
    $bytes = Get-ASCIIBytes ([IO.File]::ReadAllText($admxFile))
    $obj.content = [Convert]::ToBase64String($bytes)

    #$bytes = [IO.File]::ReadAllBytes($admlFile)
    $bytes = Get-ASCIIBytes ([IO.File]::ReadAllText($admlFile))

    $obj.groupPolicyUploadedLanguageFiles += [PSCustomObject]@{
        fileName = [io.path]::GetFileName($admlFile)
        content = [Convert]::ToBase64String($bytes)
        languageCode = (?? $obj.defaultLanguageCode "en-US")
    }
    $obj.defaultLanguageCode = ""
}

function Start-PostImportADMXFiles
{
    param($obj, $objectType, $file)

    $script:CustomADMXDefinitions = $null
}

function Start-PreDeleteADMXFiles
{
    param($obj, $objectType)

    Write-Status "Delete $($obj.fileName)"
    $strAPI = ($objectType.API + "/$($obj.Id)/remove")
    Write-Log "Delete $($objectType.Title) object $($obj.fileName)"
    Invoke-GraphRequest -Url $strAPI -HttpMethod "POST" -ODataMetadata "none" | Out-Null
    
    @{ "Delete" = $false }
}

#endregion

#region Reusable Groups
function Start-PostGetReusableSettings
{
    param($obj, $objectType)

    $strAPI = "$($objectType.API)/$($obj.Object.Id)?`$select=settinginstance,displayname,description"
    $tmpObj = Invoke-GraphRequest -Url $strAPI

    if($tmpObj.settingInstance)
    {
        $obj.Object | Add-Member Noteproperty -Name "settingInstance" -Value $tmpObj.settingInstance -Force 
    }
}

#endregon

#region Authentication Strength
function Start-PreImportCommandAuthenticationStrengths
{
    param($obj, $objectType, $file, $assignments)

    if($obj.policyType -ne "custom")
    {
        Write-Log "Built-in Authentication Strength objects cannot be imported" 2
        @{ "Import" = $false }
    }
}
#endregion

#region Authentication Strength
function Start-PreImportCommandAuthenticationContext
{
    param($obj, $objectType, $file, $assignments)

    #@{ "Method" = "PATCH" }
    
}
#endregion


Export-ModuleMember -alias * -function *