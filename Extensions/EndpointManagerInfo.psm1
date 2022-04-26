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
    '3.5.0'
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
        ItemChanged = { Show-GraphObjects -ObjectTypeChanged; Invoke-ModuleFunction "Invoke-GraphObjectsChanged"; Write-Status ""}
        Activating = { Invoke-EMInfoActivatingView }
        Authentication = (Get-MSALAuthenticationObject)
        Authenticate = { Invoke-EMInfoAuthenticateToMSAL }
        AppInfo = (Get-GraphAppInfo "EMAzureApp" "d1ddf0e4-d672-4dae-b554-9d5bdfd93547" "EM")
        SaveSettings = { Invoke-EMSaveSettings }
        Permissions = @()
    })

    Add-ViewObject $global:EMInfoViewObject

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Baseline Templates"
        Id = "BaselineTemplates"
        ViewID = "EMInfoGraphAPI"
        API = "/deviceManagement/templates"
        ShowButtons = @("Export","View")
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
    $global:EMInfoViewObject.AppInfo = Get-GraphAppInfo "EMAzureApp" "d1ddf0e4-d672-4dae-b554-9d5bdfd93547" "EM"
    Set-MSALCurrentApp $global:EMInfoViewObject.AppInfo
    $usr = (?? $global:MSALToken.Account.UserName (Get-Setting "" "LastLoggedOnUser"))
    if($usr)
    {
        & $global:msalAuthenticator.Login -Account $usr
    }
}