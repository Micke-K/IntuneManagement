<#
.SYNOPSIS
Module for read-only Intune objects

.DESCRIPTION
This module is for the Endpoint Info View. It shows read-only objects in Intune

.NOTES
  Author:         Mikael Karlsson
#>
function Get-ModuleVersion
{
    '3.9.6'
}

function Invoke-InitializeModule
{
    #Add menu group and items
    $global:EMInfoViewObject = (New-Object PSObject -Property @{ 
        Title = "Intune Info"
        Description = "Displays read-only information in Intune."
        ID = "EMInfoGraphAPI" 
        ViewPanel = $viewPanel 
        AuthenticationID = "MSAL"
        AllowDelete = $false
        ItemChanged = { Show-GraphObjects -ObjectTypeChanged; Invoke-ModuleFunction "Invoke-GraphObjectsChanged"; Write-Status ""}
        Activating = { Invoke-EMInfoActivatingView }
        Authentication = (Get-MSALAuthenticationObject)
        Authenticate = { Invoke-EMInfoAuthenticateToMSAL }
        AppInfo = (Get-GraphAppInfo "EMAzureApp" $global:DefaultAzureApp "EM")
        SaveSettings = { Invoke-EMSaveSettings }
        Permissions = @()
    })

    Add-ViewObject $global:EMInfoViewObject

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Baseline Templates - Intent"
        Id = "BaselineTemplates"
        ViewID = "EMInfoGraphAPI"
        API = "/deviceManagement/templates"        
        ShowButtons = @("Export","View")
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Icon="EndpointSecurity"
        ExpandAssignmentsList = $false
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Baseline Templates - Settings Catalog"
        Id = "BaselineTemplatesSettingsCatalog"
        ViewID = "EMInfoGraphAPI"
        API = "/deviceManagement/configurationPolicyTemplates"
        QUERYLIST = "`$filter=(templateFamily eq 'Baseline')"
        ShowButtons = @("Export","View")
        DefaultColumns = "0,displayName=Template Name,displayVersion=Version,lifecycleState=State,baseId=Template Id,id"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Icon="EndpointSecurity"
        ExpandAssignmentsList = $false
    })    

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Android Google Play"
        Id = "AndroidGooglePlay"
        ViewID = "EMInfoGraphAPI"
        ViewProperties = @("bindStatus", "lastAppSyncDateTime", "ownerUserPrincipalName")
        API = "/deviceManagement/androidManagedStoreAccountEnterpriseSettings"
        ShowButtons = @("Export","View")
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        ExpandAssignmentsList = $false   
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Android Enrolment Profiles"
        Id = "AndroidEnrolmentProfiles"
        ViewID = "EMInfoGraphAPI"
        API = "deviceManagement/androidDeviceOwnerEnrollmentProfiles"
        ShowButtons = @("Export","View")
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Icon = "AndroidCOWP"
        ExpandAssignmentsList = $false
    })    
    
    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Apple VPP Tokens"
        Id = "AppleVPPTokens"
        ViewID = "EMInfoGraphAPI"
        ViewProperties = @("appleId", "state", "appleId", "id")
        API = "/deviceAppManagement/vppTokens"
        ShowButtons = @("Export","View")
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        ExpandAssignmentsList = $false
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Apple Enrollment Tokens"
        Id = "AppleEnrollmentTokens"
        ViewID = "EMInfoGraphAPI"
        ViewProperties = @("tokenName", "appleIdentifier", "tokenExpirationDateTime", "id")
        API = "/deviceManagement/depOnboardingSettings/?`$top=100"
        ShowButtons = @("Export","View")
        Permissons=@("DeviceManagementServiceConfig.ReadWrite.All")
        ExpandAssignmentsList = $false
    })

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Tenant Settings"
        Id = "TenantSettings"
        ViewID = "EMInfoGraphAPI"
        API = "deviceManagement/settings"
        NameProperty = "Name"
        AlwaysImport = $true
        #ExportFullObject = $true
        ViewProperties = @("Name")
        ShowButtons = @("Import","Export","View")
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        PreImportCommand = { Start-PreImportTenantSettings @args }
        GetObjectName = { Start-GetObjectNameTenantSettings @args }
        PostListCommand = { Start-PostListTenantSettings @args }
        Icon="TenantSettings"
        ExpandAssignmentsList = $false
    })    
}

function Invoke-EMInfoActivatingView
{
    if(-not $global:EMInfoViewObject.ViewPanel)
    {
        # Use the same view panel as Intune Manager
        $global:EMInfoViewObject.ViewPanel = $global:EMViewObject.ViewPanel
    }
}

function Invoke-EMInfoAuthenticateToMSAL
{
    $global:EMInfoViewObject.AppInfo = Get-GraphAppInfo "EMAzureApp" $global:DefaultAzureApp "EM"
    Set-MSALCurrentApp $global:EMInfoViewObject.AppInfo
    $usr = (?? $global:MSALToken.Account.UserName (Get-Setting "" "LastLoggedOnUser"))
    if($usr)
    {
        & $global:msalAuthenticator.Login -Account $usr
    }
}

function Start-PreImportTenantSettings
{
    param($obj, $objectType)

    $objClone = $obj | ConvertTo-Json -Depth 50 | ConvertFrom-Json
    if($objClone.deviceComplianceCheckinThresholdDays -lt 1)
    {
        $objClone.deviceComplianceCheckinThresholdDays = 30
    }
    Remove-Property $objClone "@odata.type"
    $json = @{ "settings" = $objClone } | ConvertTo-Json -Depth 50
    (Invoke-GraphRequest -Url "deviceManagement" -Content $json -HttpMethod "PATCH") | Out-Null

    return (@{"Import"=$false})
}

function Start-GetObjectNameTenantSettings
{
    param($objList, $objectType)

    return "Tenant Settings"
}

function Start-PostListTenantSettings
{
    param($objList, $objectType)

    if(($objList | measure).Count -eq 1) 
    {
        $objList[0].Name = "Tenant Settings"
        #$objList[0] | Add-Member -MemberType NoteProperty -Name "SettingName" -Value "Tenant Settings"
    }
    $objList
}