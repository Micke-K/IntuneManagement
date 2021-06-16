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
    '3.1.0'
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

    $viewPanel = Get-XamlObject ($global:AppRootFolder + "\Xaml\EndpointManagerPanel.xaml") -AddVariables
    
    Set-EMViewPanel $viewPanel

    #Add menu group and items
    $global:EMViewObject = (New-Object PSObject -Property @{ 
        Title = "Intune Manager"
        Description = "Manages Intune environments. This view can be used for copying objects in an Intune environment. It can also be used for backing up an entire Intune environment and cloning the Intune environment into another tenant."
        ID="IntuneGraphAPI" 
        ViewPanel = $viewPanel 
        ItemChanged = { Show-GraphObjects; Write-Status ""}
        Deactivating = { Invoke-EMDeactivateView }
        Activating = { Invoke-EMActivatingView  }
        Authentication = (Get-MSALAuthenticationObject)
        Authenticate = { Invoke-EMAuthenticateToMSAL }
        AppInfo = (Get-GraphAppInfo "EMAzureApp" "d1ddf0e4-d672-4dae-b554-9d5bdfd93547")
        SaveSettings = { Invoke-EMSaveSettings }
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
        GroupId = "DeviceConfiguration"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Conditional Access"
        Id = "ConditionalAccess"
        ViewID = "IntuneGraphAPI"
        API = "/identity/conditionalAccess/policies"
        Permissons=@("Policy.Read.All","Policy.ReadWrite.ConditionalAccess","Application.Read.All")
        Dependencies = @("NamedLocations","Applications")
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
        #PreCopyCommand = { Start-PreCopyEndpointSecurity @args }
        PostCopyCommand = { Start-PostCopyEndpointSecurity @args }
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        GroupId = "EndpointSecurity"
    })    

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Compliance Policies"
        Id = "CompliancePolicies"
        ViewID = "IntuneGraphAPI"
        Expand = "scheduledActionsForRule(`$expand=scheduledActionConfigurations)"
        API = "/deviceManagement/deviceCompliancePolicies"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Dependencies = @("Locations","Notifications")
        PostExportCommand = { Start-PostExportCompliancePolicies @args }
        GroupId = "CompliancePolicies"
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
        Permissons=@("DeviceManagementApps.ReadWrite.All")
        Icon = "Branding"
        SkipRemoveProperties = @('Id') # Id is removed by PreImport. Required for default profile
        GroupId = "TenantAdmin"
    })

    <#
    # BUG in Graph? Cannot create default branding. Can only create it when importing another object
    # Header required Accept-Language: sv-SE
    # Documentation says to use Content-Language but that doesn't work

    # Could work with https://main.iam.ad.ext.azure.com/api/LoginTenantBrandings
    

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Azure Branding"
        Id = "AzureBranding"
        API = "/organization/%OrganizationId%/branding/localizations"
        ViewID = "IntuneGraphAPI"
        ViewProperties = @("Id")
        PreImportCommand = { Start-PreImportAzureBranding  @args }
        PostListCommand = { Start-PostListAzureBranding @args }
        NameProperty = "Id"
        Permissons=@("Organization.ReadWrite.All")
        Icon = "Branding"
        SkipRemoveProperties = @('Id')
    })
    #>

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Enrollment Status Page"
        Id = "EnrollmentStatusPage"
        API = "/deviceManagement/deviceEnrollmentConfigurations"
        ViewID = "IntuneGraphAPI"
        PreImportCommand = { Start-PreImportESP @args }
        PostExportCommand = { Start-PostExportESP @args }
        QUERYLIST = "`$filter=endsWith(id,'Windows10EnrollmentCompletionPageConfiguration')"
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        SkipRemoveProperties = @('Id')
        AssignmentsType = "enrollmentConfigurationAssignments"
        GroupId = "WinEnrollment"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Enrollment Restrictions"
        Id = "EnrollmentRestrictions"
        API = "/deviceManagement/deviceEnrollmentConfigurations"
        ViewID = "IntuneGraphAPI"
        QUERYLIST = "`$filter=not endsWith(id,'Windows10EnrollmentCompletionPageConfiguration')"
        PostExportCommand = { Start-PostExportEnrollmentRestrictions @args }
        PreImportCommand = { Start-PreImportEnrollmentRestrictions @args }
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        SkipRemoveProperties = @('Id')
        AssignmentsType = "enrollmentConfigurationAssignments"
        GroupId = "EnrollmentRestrictions"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Administrative Templates"
        Id = "AdministrativeTemplates"
        API = "/deviceManagement/groupPolicyConfigurations"
        ViewID = "IntuneGraphAPI"
        PostExportCommand = { Start-PostExportAdministrativeTemplate @args }
        PostCopyCommand = { Start-PostCopyAdministrativeTemplate @args }
        PostFileImportCommand = { Start-PostFileImportAdministrativeTemplate @args }
        LoadObject = { Start-LoadAdministrativeTemplate @args }
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Icon="DeviceConfiguration"
        GroupId = "DeviceConfiguration"
        CompareValue = "CombinedValueWithLabel"
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
        ExportFullObject = $true
        PropertiesToRemove = @('exemptAppLockerFiles')
        Permissons=@("DeviceManagementApps.ReadWrite.All")
        Dependencies = @("Applications")
        GroupId = "AppProtection"
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
        Permissons=@("DeviceManagementApps.ReadWrite.All")
        Dependencies = @("Applications")
        Icon = "AppConfiguration"
        GroupId = "AppConfiguration"
    })    

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "App Configuration (Device)"
        Id = "AppConfigurationManagedDevice"
        API = "/deviceAppManagement/mobileAppConfigurations"
        QUERYLIST = "`$filter=microsoft.graph.androidManagedStoreAppConfiguration/appSupportsOemConfig%20eq%20false%20or%20isof(%27microsoft.graph.androidManagedStoreAppConfiguration%27)%20eq%20false"
        ViewID = "IntuneGraphAPI"
        Permissons=@("DeviceManagementApps.ReadWrite.All")
        Dependencies = @("Applications")
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsAppConfiguration @args }
        #PostExportCommand = { Start-PostExportAppConfiguration @args }
        Icon = "AppConfiguration"
        GroupId = "AppConfiguration"
    })
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Applications"
        Id = "Applications"
        API = "/deviceAppManagement/mobileApps"
        ViewID = "IntuneGraphAPI"
        PropertiesToRemove = @('uploadState','publishingState','isAssigned','dependentAppCount','supersedingAppCount','supersededAppCount','committedContentVersion','isFeatured','size','categories')
        QUERYLIST = "`$filter=(microsoft.graph.managedApp/appAvailability%20eq%20null%20or%20microsoft.graph.managedApp/appAvailability%20eq%20%27lineOfBusiness%27%20or%20isAssigned%20eq%20true)&`$orderby=displayName"
        Permissons=@("DeviceManagementApps.ReadWrite.All")
        AssignmentsType="mobileAppAssignments"
        AssignmentProperties = @("@odata.type","target","settings","intent")
        AssignmentTargetProperties = @("@odata.type","groupId","deviceAndAppManagementAssignmentFilterId","deviceAndAppManagementAssignmentFilterType")
        ImportOrder = 60
        Expand="categories,assignments" # ODataMetadata is set to minimal so assignments can't be autodetected
        ODataMetadata="minimal" # categories property not supported with ODataMetadata full
        PostFileImportCommand = { Start-PostFileImportApplications @args }
        GroupId = "Apps"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "AutoPilot"
        Id = "AutoPilot"
        API = "/deviceManagement/windowsAutopilotDeploymentProfiles"
        ViewID = "IntuneGraphAPI"
        CopyDefaultName = "%displayName% Copy" # '-' is not allowed in the name
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        PreImportAssignmentsCommand = { Start-PreImportAssignmentsAutoPilot @args }
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
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        ImportOrder = 2000 # Policy Sets reference other objects so make sure it is imported last
        Dependencies = @("Applications","AppConfiguration","AppProtection","AutoPilot","EnrollmentRestrictions","EnrollmentStatusPage","DeviceConfiguration","AdministrativeTemplates","SettingsCatalog","CompliancePolicies")
        GroupId = "PolicySets"
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
    })
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Feature Updates"
        Id = "FeatureUpdates"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/windowsFeatureUpdateProfiles"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        GroupId = "WinFeatureUpdates"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Quality Updates"
        Id = "QualityUpdates"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/windowsQualityUpdateProfiles"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Icon = "UpdatePolicies"
        GroupId = "WinQualityUpdates"
    })    

    # Locations are not FULLY supported 
    # They will be imported but Compliance Policies will not be updated with new Location object after import
    # ToDo: Add support Export/Import Location Settings
    # Location object - Only used by Android Device Admins Compliance Policies 
    # - These should probably be migrated to Android Enterprise anyway. That is the recommendation by Google
    # Property that needs to be updated on the Compliance Policy
    # deviceManagement/managementConditionStatements/$obj.conditionStatementId
    
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
        GroupId = "TenantAdmin"
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Scope (Tags)"
        Id = "ScopeTags"
        ViewID = "IntuneGraphAPI"
        API = "/deviceManagement/roleScopeTags"
        QUERYLIST = "`$filter=isBuiltIn%20eq%20false"
        Permissons=@("DeviceManagementRBAC.ReadWrite.All")
        PostExportCommand = { Start-PostExportScopeTags @args }
        ImportOrder = 10
        GroupId = "TenantAdmin"
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
        GroupId = "CompliancePolicies"
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
    })
}

function Invoke-EMAuthenticateToMSAL
{
    $global:EMViewObject.AppInfo = Get-GraphAppInfo "EMAzureApp" "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"
    Set-MSALCurrentApp $global:EMViewObject.AppInfo
    & $global:msalAuthenticator.Login -Account (?? $global:MSALToken.Account.UserName (Get-Setting "" "LastLoggedOnUser"))
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
    $global:EMViewObject.AppInfo = (Get-GraphAppInfo "EMAzureApp" "d1ddf0e4-d672-4dae-b554-9d5bdfd93547")
    if(-not $global:EMViewObject.Authentication)
    {
        $global:EMViewObject.Authentication = Get-MSALAuthenticationObject    
    }

    # Add View specific menus
    Add-GraphBulkMenu
}

function Invoke-EMSaveSettings
{
    $tmpApp = Get-GraphAppInfo "EMAzureApp" "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"

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
}

function Set-EMViewPanel
{
    param($panel)
    # ToDo: Create View specific pannel and move this to graph
    Add-XamlEvent $panel "btnView" "Add_Click" -scriptBlock ([scriptblock]{ 
        Show-GraphObjectInfo
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
        Invoke-FiterBoxChanged $this
        #$e.Handled = $true
    })
    
    Add-XamlEvent $panel "txtFilter" "Add_GotFocus" ({
        if($this.Tag -eq "1" -and $this.Text -eq "Filter") { $this.Text = "" }
        Invoke-FiterBoxChanged $this
    })
    
    Add-XamlEvent $panel "txtFilter" "Add_TextChanged" ({
        Invoke-FiterBoxChanged $this
    })

    Invoke-FiterBoxChanged ($panel.FindName("txtFilter"))
    

    $global:dgObjects.add_selectionChanged({        
        Set-XamlProperty $this.Parent "btnView" "IsEnabled" (?: ($global:dgObjects.SelectedItem -eq $null) $false $true)
        Set-XamlProperty $this.Parent "btnCopy" "IsEnabled" (?: ($global:dgObjects.SelectedItem -eq $null) $false $true)
    })

    # ToDo: Move this to the view object
    $dpd = [System.ComponentModel.DependencyPropertyDescriptor]::FromProperty([System.Windows.Controls.ItemsControl]::ItemsSourceProperty, [System.Windows.Controls.DataGrid])
    if($dpd)
    {
        $dpd.AddValueChanged($global:dgObjects, {
            Set-XamlProperty $global:dgObjects.Parent "txtFilter" "Text" ""
            $enabled = (?: ($this.ItemsSource -eq $null -or ($this.ItemsSource | measure).Count -eq 0) $false $true)
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
        $btnRefresh.Focusable = $false
        $grdTitle.Children.Add($btnRefresh) | Out-Null

        $tooltip = [System.Windows.Controls.ToolTip]::new()
        $tooltip.Content = "Refresh"
        [System.Windows.Controls.ToolTipService]::SetToolTip($btnRefresh, $tooltip)

        $btnRefresh.Add_Click({
            # ToDo: Move this to view view object
            $txtFilter = $this.Parent.FindName("txtFilter")
            if($txtFilter) { $txtFilter.Text = "" }
            Show-GraphObjects
            Write-Status ""
        })
    }        
}

function Invoke-FiterBoxChanged 
{ 
    param($txtBox)

    $filter = $null
    
    if($txtBox.Text.Trim() -eq "" -and $txtBox.IsFocused -eq $false)
    {
        $txtBox.FontStyle = "Italic"
        $txtBox.Tag = 1
        $txtBox.Text = "Filter"
        $txtBox.Foreground="Lightgray"        
    }
    else
    {
        if($txtBox.Tag -eq "1" -and $txtBox.Text -eq "Filter" -and $txtBox.IsFocused -eq $false) { return }
        $txtBox.FontStyle = "Normal"
        $txtBox.Tag = $null
        $txtBox.Foreground="Black"
        $txtBox.Background="White"

        if($txtBox.Text)
        {
            $filter = {
                param ($item)
                foreach($prop in ($item.PSObject.Properties | Where {$_.Name -notin @("IsSelected","Object")}))
                {
                    if($prop.Value -match $txtBox.Text) { return $true }
                }
                $false
            }
        }         
    }

    if($dgObjects.ItemsSource -is [System.Windows.Data.ListCollectionView])
    {
        # This causes odd behaviour with focus e.g. and item has to be clicked twice to be selected 
        $dgObjects.ItemsSource.Filter = $filter
        $dgObjects.ItemsSource.Refresh()
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

    $settings = Invoke-GraphRequest -Url "$($objectType.API)/$($obj.id)/settings"
    $settingsJson = "{ `"settings`": $((ConvertTo-Json  $settings.value -Depth 10 ))`n}"
    $settingsJson | Out-File "$path\$((Get-GraphObjectName $obj $objectType))_Settings.json" -Force
}

function Start-PostFileImportEndpointSecurity
{
    param($obj, $objectType, $file)

    $settings = Get-EMSettingsObject $obj $objectType $file
    if($settings)
    {
        Start-GraphPreImport $settings
        Invoke-GraphRequest -Url "$($objectType.API)/$($obj.id)/updateSettings" -Body ($settings | ConvertTo-Json -Depth 10) -Method "POST"
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
        Invoke-GraphRequest -Url "$($objectType.API)/$($objNew.id)/updateSettings" -Body ($settingsObj | ConvertTo-Json -Depth 10) -Method "POST"
    }
}

#endregion

#region 

function Start-PostFileImportDeviceConfiguration
{
    param($obj, $objectType, $importFile)

    if($obj.'@OData.Type' -like "#microsoft.graph.windows10GeneralConfiguration")
    {
        $tmpObj = Get-Content $importFile | ConvertFrom-Json

        if(($tmpObj.privacyAccessControls | measure).Count -gt 0)
        {
            $privacyObj = [PSCustomObject]@{
                windowsPrivacyAccessControls = $tmpObj.privacyAccessControls
            }
            $json =  $privacyObj | ConvertTo-Json -Depth 10
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
            $json =  $privacyObj | ConvertTo-Json -Depth 10
            $ret = Invoke-GraphRequest -Url "deviceManagement/deviceConfigurations('$($objNew.Id)')/windowsPrivacyAccessControls" -Body $json -Method "POST"
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

        $global:brandingClone = $obj | ConvertTo-Json -Depth 10 | ConvertFrom-Json

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
    $json = ($global:brandingClone | ConvertTo-Json -Depth 10)
    $ret = Invoke-GraphRequest -Url "$($objectType.API)/$($obj.Id)" -Body $json -Method "PATCH"
}

function Start-PostGetIntuneBranding
{
    param($obj, $objectType)

    foreach($imgType in @("themeColorLogo","lightBackgroundLogo","landingPageCustomizedImage"))
    {
        Write-LogDebug "Get $imgType for $($obj.profileName)"
        $imgJson = Invoke-GraphRequest -Url "$($objectType.API)/$($obj.Id)/$imgType"
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
            $fileName = "$path\$((Remove-InvalidFileNameChars  (Get-GraphObjectName $obj $objectType)))_$imgType.jpg" 
            [IO.File]::WriteAllBytes($fileName, [System.Convert]::FromBase64String($obj.$imgType.Value))
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
}

function Add-ScriptExportExtensions
{
    param($form, $buttonPanel, $index = 0)

    $xaml =  @"
<StackPanel $($global:wpfNS) Orientation="Horizontal" Margin="0,0,5,0">
<Label Content="Export script" />
<Rectangle Style="{DynamicResource InfoIcon}" ToolTip="Export the powershell script to a ps1 file" />
</StackPanel>
"@
    $label = [Windows.Markup.XamlReader]::Parse($xaml)

    $global:chkExportScript = [System.Windows.Controls.CheckBox]::new()
    $global:chkExportScript.IsChecked = $true
    $global:chkExportScript.VerticalAlignment = "Center" 

    @($label, $global:chkExportScript)
}

function Start-PostExportScripts
{
    param($obj, $objectType, $exportPath)

    if($obj.scriptContent -and $global:chkExportScript.IsChecked)
    {
        Write-Log "Export script $($obj.FileName)"
        $fileName = [IO.Path]::Combine($exportPath, $obj.FileName)
        [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($obj.scriptContent)) | Out-File $fileName -Force
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
    # For some reason, the some $filter options is not supported to filter out these objects
    # e.g. not isof(...) to excluded the type, not startsWith(id, 'A_') to exlude based on Id
    # These filters generates a request error so fiter them out manually in this function instead
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

                $response = Invoke-GraphRequest -Url "/deviceAppManagement/$objectClass/$($obj.Id)/targetApps" -Content "{ apps: $(ConvertTo-Json $global:ImportObjectInfo.Apps -Depth 10)}" -HttpMethod POST
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
#endregion

#region App Configuration
function Start-PostExportAppConfiguration
{
    param($obj, $objectType, $path)

    Add-EMAssignmentsToExportFile $obj $objectType $path 
}

function Start-PreImportAssignmentsAppConfiguration
{
    param($obj, $objectType, $file, $assignments)

    @{"API"="/deviceAppManagement/mobileAppConfigurations/$($obj.Id)/microsoft.graph.managedDeviceMobileAppConfiguration/assign"}
}
#endregon

#region Applications
function Start-PostFileImportApplications
{
    param($obj, $objectType, $file)

    $tmpObj = Get-Content $file | ConvertFrom-Json

    if(-not $tmpObj.'@odata.type') { return }

    $pkgPath = Get-SettingValue "EMIntuneAppPackages"

    if(-not $pkgPath -or [IO.Directory]::Exists($pkgPath) -eq $false) 
    {
        Write-LogDebug "Package source directory is either missing or does not exist" 2
        return 
    }

    $packageFile = "$($pkgPath)\$($obj.fileName)"

    if([IO.File]::Exists($packageFile) -eq $false) 
    {
        Write-LogDebug "Package source file $packageFile not found" 2
        return 
    }

    Write-Status "Import appliction package file $($obj.fileName)"
    Write-Log "Import application file '$($packageFile)' for $($obj.displayName)"

    if(-not ($obj.PSObject.Properties | Where Name -eq '@odata.type'))
    {
        # Add @odata.type property if it is missing. Required by app package import
        $obj | Add-Member -MemberType NoteProperty -Name '@odata.type' -Value $tmpObj.'@odata.type'
    }

    $appType = $tmpObj.'@odata.type'.Trim('#')

    if($appType -eq "microsoft.graph.win32LobApp")
    {
        Copy-Win32LOBPackage $packageFile $obj
    }
    elseif($appType -eq "microsoft.graph.windowsMobileMSI")
    {
        Copy-MSILOB $packageFile $obj
    }
    elseif($appType -eq "microsoft.graph.iosLOBApp")
    {
        Copy-iOSLOB $packageFile $obj
    }
    elseif($appType -eq "microsoft.graph.androidLOBApp")
    {
        Copy-AndroidLOB $packageFile $obj
    }
    else
    {
        Write-Log "Unsupported application type $appType. File will not be uploaded" 2    
    }
}
#endregion

#region Group Policy/Administrative Templates functions
function Get-GPOObjectSettings
{
    param($GPOObj)

    $gpoSettings = @()

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

        if($presentationValues.value)
        {
            # Policy presentation values set e.g. a drop down list, check box, text box etc.
            $obj.presentationValues = @()                        
            
            foreach ($presentationValue in $presentationValues.value) 
            {
                # Add presentation@odata.bind property that links the value to the presentation object
                $presentationValue | Add-Member -MemberType NoteProperty -Name "presentation@odata.bind" -Value "$($global:graphURL)/deviceManagement/groupPolicyDefinitions('$($definitionValue.definition.id)')/presentations('$($presentationValue.presentation.id)')"

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
        
        foreach($setting in $settings)
        {
            Start-GraphPreImport $setting

            # Import each setting for the Administrative Template profile
            Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations/$($obj.id)/definitionValues" -Content (ConvertTo-Json $setting -Depth 10) -HttpMethod POST | Out-Null
        }
    }
}

function Start-PostExportAdministrativeTemplate
{
    param($obj, $objectType, $path)

    # Collect and save all the settings of the Administrative Templates profile
    $settings = Get-GPOObjectSettings $obj
    ConvertTo-Json $settings -Depth 10 | Out-File "$path\$((Get-GraphObjectName $obj $objectType))_Settings.json" -Force
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

    $settings = Get-EMSettingsObject $obj $objectType $file
    if($settings)
    {        
        Import-GPOSetting $obj $settings
    }    
}

function Start-LoadAdministrativeTemplate
{
    param($fileName)

    if(-not $fileName) { return $null }

    $fi = [IO.FileInfo]$fileName
    $obj = Get-Content $global:txtCompareFile.Text | ConvertFrom-Json 

    if($obj.definitionValues)
    {
        return $obj
    }

    $settingsFile = $fi.DirectoryName + "\" + $fi.BaseName + "_Settings.json"

    if([IO.File]::Exists($settingsFile))
    {
        $definitionValues = Get-Content $settingsFile | ConvertFrom-Json

        $obj | Add-Member Noteproperty -Name "definitionValues" -Value $definitionValues -Force  
    }
    $obj
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
            #if($prop.Name -in $keepProperties) { continue }
            Remove-Property $item $prop.Name
        }
        #@("itemType","displayName","status","errorCode") | foreach { Remove-Property $item $_ }
    }
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

    $fileName = "$path\$((Get-GraphObjectName $obj $objectType)).json"
    $tmpObj = Get-Content $fileName | ConvertFrom-Json

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
            $tmpObj | ConvertTo-Json -Depth 10 | Out-File $fileName
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

    $tmpObj = Get-Content $file | ConvertFrom-Json

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
            $json = Update-JsonForEnvironment (ConvertTo-Json $roleAssignmentObj -Depth 10)

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

    $tmpObj = Get-Content $file | ConvertFrom-Json

    foreach($localizedNotificationMessage in $tmpObj.localizedNotificationMessages)
    {
        Start-GraphPreImport $localizedNotificationMessage $objectType
        Invoke-GraphRequest -Url "$($objectType.API)/$($obj.id)/localizedNotificationMessages" -Body ($localizedNotificationMessage | ConvertTo-Json -Depth 10) -Method "POST"
    }
}

function Start-PostCopyNotifications
{
    param($objCopyFrom, $objNew, $objectType)

    foreach($localizedNotificationMessage in $objCopyFrom.localizedNotificationMessages)
    {
        Start-GraphPreImport $localizedNotificationMessage $objectType
        Invoke-GraphRequest -Url "$($objectType.API)/$($objNew.id)/localizedNotificationMessages" -Body ($localizedNotificationMessage | ConvertTo-Json -Depth 10) -Method "POST"
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
}
#endregion

#region ScopeTags
function Start-PostExportScopeTags
{
    param($obj, $objectType, $path)

    Add-EMAssignmentsToExportFile $obj $objectType $path
}
#endregion

#region AutoPilot
function Start-PreImportAssignmentsAutoPilot
{
    param($obj, $objectType, $file, $assignments)

    Add-EMAssignmentsToObject $obj $objectType $file $assignments
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
                try { [IO.File]::Delete($oldFile) | Out-Null } Catch {}

                $obj | ConvertTo-Json -Depth 10 | Out-File "$path\$fileName.json"
            }
        }
        catch {}
    }   
}
function Get-EMSettingsObject
{
    param($obj, $objectType, $file)

    if($obj.Settings) { $obj.Settings }

    $fi = [IO.FileInfo]$file
    if($fi.Exists)
    {
        Write-Log "Settings not included in export file. Try import from _Settings.json file" 2
        $settingsFile = $fi.DirectoryName + "\" + $fi.BaseName + "_Settings.json"
        $fiSettings = [IO.FileInfo]$settingsFile
        if($fiSettings.Exists -eq $false)
        {
            Write-Log "Settings file '$($fiSettings.FullName)' was not found" 2
            return
        }

        (Get-Content $fiSettings.FullName) | ConvertFrom-Json
    }
    else
    {
        Write-Log "Settings not included in export file and _Settings.json file is missing." 3
    }
}

function Add-EMAssignmentsToExportFile
{
    param($obj, $objectType, $path, $Url = "")

    $fileName = "$path\$((Get-GraphObjectName $obj $objectType)).json"
    $tmpObj = Get-Content $fileName | ConvertFrom-Json

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
        ConvertTo-Json $tmpObj -Depth 10 | Out-File $fileName -Force
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

        $json = Update-JsonForEnvironment ($assignment | ConvertTo-Json -Depth 10)
        Invoke-GraphRequest -Url $api -Body $json -Method "POST" | Out-Null
    }
    @{"Import"=$false}
}

#endregion

Export-ModuleMember -alias * -function *