<#
.SYNOPSIS
Module for Authentication

.DESCRIPTION
This module manages Authentication for the application with MSAL. It is also responsible for displaying the Profile Picture control of he logged in user.

.NOTES
  Author:         Mikael Karlsson
#>
function Get-ModuleVersion
{
    '3.9.8a'
}

$global:msalAuthenticator = $null
function Invoke-InitializeModule
{
    $script:MSALAllApps = @()
    $global:MSALToken = $null
    $global:MSALTenantId = $null
    $script:AccessableTenants = $null
    $global:SkipTokenCacheHelperEx = $null

    $script:lstAADEnvironments = @(
        [PSCustomObject]@{
            Name = "Azure AD Public"
            Value = "public"
            URL = "login.microsoftonline.com"
        },
        [PSCustomObject]@{
            Name = "Azure AD US Government"
            Value = "usGov"
            URL = "login.microsoftonline.us"
        },
        [PSCustomObject]@{
            Name = "Azure AD China"
            Value = "china"
            URL = "login.partner.microsoftonline.cn"
            GraphURL = "microsoftgraph.chinacloudapi.cn"
        }
    )

    $script:lstGCCEnvironments = @(
        [PSCustomObject]@{
            Name = "GCC"
            Value = "gcc"
            URL = "graph.microsoft.com"
        },
        [PSCustomObject]@{
            Name = "GCC High"
            Value = "gcgHigh"
            URL = "graph.microsoft.us"
        },
        [PSCustomObject]@{
            Name = "GCC DoD"
            Value = "gccDoD"
            URL = "dod-graph.microsoft.us"
        }
    )

    $global:appSettingSections += (New-Object PSObject -Property @{
        Title = "MSAL"
        Id = "MSAL"
        Values = @()
        Priority = 8
    })

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "MSAL Library File"
        Key = "MSALDLL"
        Type = "File" 
        Description = "Full path to the Microsoft.Identity.Client.dll file"
    }) "MSAL"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Remember Login"
        Key = "CacheMSALToken"
        Type = "Boolean"
        DefaultValue = $true
        Description = "Store the MSAL token in an encrypted file and automatically log when the script starts. The token is stored in the users profile and can only be decrypted by the user that created it. Note: Requires restart"
    }) "MSAL"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Get Tenant List"
        Key = "GetTenantList"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Get a list of all tenants the current user has access to. Only used when the user has access to multiple tenants. This may cause duplicate login/consent prompts first time"
    }) "MSAL"
    
    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Use Default Permissions"
        Key = "UseDefaultPermissions"
        Type = "Boolean"
        DefaultValue = $true
        Description = "Default permissions of the selected app will be used when logging on. Some objects might not be accessable"
    }) "MSAL" 

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Add Azure Role Read permissions"
        Key = "AzureADRoleRead"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Request Azure AD Role read permission when getting the token. This can be use to resolve the SIDs to Azure Roles for the wids property on the Access Token. Note: This might trigger a consent prompt"
    }) "MSAL"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Show Azure AD login menu"
        Key = "AzureADLoginMenu"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Use this to login to US Government or China cloud. If not enabled, it will directly prompt for Public cloud login"
    }) "MSAL"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Sort Account List"
        Key = "SortAccountList"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Sort the list of cached accounts based on user name. Updated at restart or account change"
    }) "MSAL"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Sort Tenant List"
        Key = "SortTenantList"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Sort the list of available tenants based on Tenant name. Updated at restart or account change"
    }) "MSAL"

    Add-MSALPrereq
}

function Get-MSALAuthenticationObject
{
    if(-not $global:msalAuthenticator)
    {
        $global:msalAuthenticator = New-Object PSObject -Property @{
            Title = "MSAL"
            ID = "MSAL"
            SilentLogin = { Connect-MSALUser -Silent @args; } 
            Login = { Connect-MSALUser @args } 
            Logout = { Disconnect-MSALUser } 
            ProfilePicture = { Get-MSALProfileEllipse @args }
            ShowErrors = { Show-MSALError }
            Permissions = @("openid","profile","email","User.ReadWrite.All","Group.ReadWrite.All") #"RoleManagement.Read.Directory"
        }
    }

    $global:msalAuthenticator
}

function Invoke-SettingsUpdated
{
    Initialize-MSALSettings
    Invoke-MSALCheckObjectViewAccess 
}

function Initialize-MSALSettings
{

}

function Clear-MSALCurentUserVaiables
{
    $global:MSALTenantId = $null
    $global:MSALGraphEnvironment = $null

    $script:jwtAccessToken = $null
    $script:jwtIdToken = $null

}

function Get-MSALCurrentApp
{
    $global:appObj
}

function Set-MSALCurrentApp
{
    param($appInfoObj)

    $global:appObj = $appInfoObj
}

function Set-MSALGraphEnvironment
{
    param($user, $tenantId)
    
    if($global:MSALGraphEnvironment)
    {
        return
    }

    $graphEnv = "graph.microsoft.com"
    
    
    if($user)
    {
        $curAADEnv = $script:lstAADEnvironments | Where URL -eq $user.Environment
    }
    else
    {
        $curAADEnv = $script:lstAADEnvironments | Where value -eq (Get-Setting "" "MSALCloudType" "public")
    }    

    if($curAADEnv.Value -eq "usGov")
    {
        $gccEnv = (Get-Setting "" "MSALGCCType" "gcc")
        if($gccEnv)
        {
            $GCCEnvObj = $script:lstGCCEnvironments | Where Value -eq $gccEnv
            if($GCCEnvObj.URL)
            {
                $graphEnv = $GCCEnvObj.URL
            }
            else
            {
                Write-Log "Could not find GCC environment based on $gccEnv. Default will be used" 2
            }
        }
    }
    elseif($curAADEnv.GraphURL)
    {
        $graphEnv = $curAADEnv.GraphURL
    }
    
    Write-Log "Use Graph environment: $graphEnv"
    $global:MSALGraphEnvironment = $graphEnv
}

function Get-MSALUserInfo
{
    if($global:MSALToken)
    {
        Write-Log "Get current user"
        
        if($script:jwtAccessToken.Payload.idtyp -ne "app")
        {
            $tmpMe = MSGraph\Invoke-GraphRequest -Url "ME" -SkipAuthentication -ODataMetadata "Skip"
            if($null -ne $tmpMe -and $tmpMe.creationType -ne "Invitation")
            {
                ### Only get user info from home tenant
                $global:Me = $tmpMe
                Write-Log "Get profile picture"
                $global:profilePhoto = "$($env:LOCALAPPDATA)\CloudAPIPowerShellManagement\$($global:Me.Id).jpeg"
                MSGraph\Invoke-GraphRequest "me/photos/48x48/`$value" -OutFile $global:profilePhoto -SkipAuthentication -NoError | Out-Null
            }
        }
        else
        {
            $global:profilePhoto = $null
            $global:me = $script:jwtAccessToken.Payload.app_displayname
        }

        Write-Log "Get organization info"
        $global:Organization = (MSGraph\Invoke-GraphRequest -Url "Organization" -SkipAuthentication -ODataMetadata "Skip").Value
        if($global:Organization)
        {
            if($global:Organization -is [array]) { $global:Organization = $global:Organization[0]}
            Save-Setting $global:Organization.Id "_Name" $global:Organization.displayName
        }
        Set-EnvironmentInfo $global:Organization.displayName
    }
    else 
    {
        Set-EnvironmentInfo 
        $global:Me = $null
        $global:profilePhoto = $null
        $global:Organization = $null
    }
    Show-AuthenticationInfo
}

function Show-MSALError
{
    if($script:MSALDLLMissing -ne $true) { return }

    $script:msalPreReqForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\MSALPreReqForm.xaml") -AddVariables

    $isAdmin = Get-IsAdmin

    if($isAdmin)
    {
        Set-XamlProperty $script:msalPreReqForm "chkCurrentUser" "IsChecked" $false
    }

    $powerShellGet = Get-Module "PowerShellGet" -ListAvailable -ErrorAction SilentlyContinue | Sort -Property Version -Descending | Select -First 1

    if($powerShellGet -and $powerShellGet.Version -gt [Version]"2.0.0")
    {
        $global:installPowerShellGet = $false
        Write-Log "PowerShellGet $($powerShellGet.Version) detected. No need to install package"
        Set-XamlProperty $script:msalPreReqForm "spPowerShellGet" "Visibility" "Collapsed"
    }
    else
    {
        if($powerShellGet)
        {
            Write-Log "PowerShellGet $($powerShellGet.Version) detected. Module needs to be updated" 2
        }
        else
        {
            Write-Log "PowerShellGet is missing. It needs to be installed" 2
        }            

        $global:installPowerShellGet = $true    
    }

    $pkgNuGet = Get-PackageProvider | Where Name -eq "NuGet"

    if($pkgNuGet -and $pkgNuGet.Version -ge [Version]"2.8.5.201")
    {
        Write-Log "NuGet $($pkgNuGet.Version) detected. No need to install package"
        Set-XamlProperty $script:msalPreReqForm "spNuGet" "Visibility" "Collapsed"
        $global:installNuGet = $false
    }
    else
    {
        $global:installNuGet = $true

        if($isAdmin)
        {
            Set-XamlProperty $script:msalPreReqForm "txtNotAdmin" "Visibility" "Collapsed"
        }
        else
        {
            Set-XamlProperty $script:msalPreReqForm "chkInstallNuGet" "Visibility" "Collapsed"
            Set-XamlProperty $script:msalPreReqForm "btnInstallMSALPS" "IsEnabled" $false
            Set-XamlProperty $script:msalPreReqForm "btnInstallAz" "IsEnabled" $false
        }

        if($pkgNuGet)
        {
            Write-Log "NuGet $($pkgNuGet.Version) detected. Pakage needs to be updated" 2
        }
        else
        {
            Write-Log "NuGet is missing. It needs to be installed" 2
        }        
    }

    Add-XamlEvent $script:msalPreReqForm "btnInstallMSALPS" "add_click" {
        Install-MSALDependencyModule "MSAL.PS" $this
    }

    Add-XamlEvent $script:msalPreReqForm "btnInstallAz" "add_click" {
        Install-MSALDependencyModule "Az" $this
    }
    Show-ModalForm "MSAL Errors" $script:msalPreReqForm
}

function Install-MSALDependencyModule
{
    param($moduleToInstall, $button)

    if($global:hideUI -eq $true)
    {
        Write-Log "Cannot install MSAL module in Silent mode"
        return
    }

    $forceUserInstallation = $false

    if($global:chkCurrentUser.IsChecked -eq $false -and (Get-IsAdmin) -eq $false) 
    { 
        if([System.Windows.MessageBox]::Show("Module will be install for system but current user is not Admin`n`nDo you want to install modules as user instead?`n`nNo will abort the installation", "Not admin!", "YesNo", "Warning") -eq "Yes")
        {
            $forceUserInstallation = $true
        }
        else
        {
            return    
        }
        
    }

    $installExtra = ""
    if($global:installNuGet)
    {
        $installExtra += "`nNuGet will also be installed"
    }

    if($global:installPowerShellGet)
    {
        $installExtra += "`nPowerShellGet module will also be installed"
    }
    if($installExtra) { $installExtra = "`nAdditional installs:`n$($installExtra)" }

    if([System.Windows.MessageBox]::Show("Are you sure you want to install the $moduleToInstall module?$($installExtra)", "Install module?", "YesNo", "Question") -ne "Yes") { return }

    # Force TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if($global:installNuGet)
    {
        Write-Status "Install NuGet package provider"
        try 
        {
            Install-PackageProvider -Name NuGet -Force #-MinimumVersion 2.8.5.201
            if($? -eq $false)
            {
                throw $global:error[0]
            }
        }
        catch 
        {
            Write-Log "Failed to install package provider NuGet. Error: $($_.Exception.Message)" 3
            [System.Windows.MessageBox]::Show("Failed to install Nuget`n`nThe $moduleToInstall module cannot be installed.`n`nTry installing the module manually and then restart the appliaction.", "Failed!", "OK", "Error") | Out-Null
            return
        }
    }

    if($global:installPowerShellGet)
    {
        Write-Status "Install PowerShellGet module"
        try 
        {
            $params = @{}
            if($global:chkCurrentUser.IsChecked -or $forceUserInstallation) { $params.Add("Scope", "CurrentUser") }

            Install-Module -Name PowerShellGet -Force -AllowClobber @params #-MinimumVersion 2.8.5.201                

            if($? -eq $false)
            {
                throw $global:error[0]
            }
        }
        catch 
        {
            Write-Log "Failed to install PowerShellGet module. Error: $($_.Exception.Message)" 3
            [System.Windows.MessageBox]::Show("Failed to install PowerShellGet module`n`nThe $moduleToInstall module cannot be installed.`n`nTry installing the module manually and then restart the appliaction.", "Failed!", "OK", "Error") | Out-Null
            return
        }
    }

    $params = @{}
    $params.Add("Name", $moduleToInstall)
    $params.Add("Force", $true)

    if($global:chkAllowClobber.IsChecked) { $params.Add("AllowClobber", $true) }
    if($global:chkCurrentUser.IsChecked -or $forceUserInstallation) { $params.Add("Scope", "CurrentUser") }
    if($global:chkSkipPublisherCheck.IsChecked) { $params.Add("SkipPublisherCheck", $true) }
    #if($global:chkAcceptLicense.IsChecked) { $params.Add("AcceptLicense", $true) }
    $installError = ""
    Write-Status "Install module $moduleToInstall"
    try 
    {
        $mod = Install-Module @params -ErrorAction SilentlyContinue #-PassThru 
        if($? -eq $false)
        {
            throw $global:error[0]
        }
    }
    catch 
    {
        $installError = "Error: $($_.Exception.Message)`n`n"        
        Write-Log "Failed to install module. Error: $($_.Exception.Message)" 3
    }
    
    $checkModule = ?: ($moduleToInstall -eq "Az") "Az.Accounts" $moduleToInstall

    if(-not $mod) { $mod = Get-Module $checkModule -ListAvailable }

    Write-Status ""

    if($mod)
    {
        $script:MSALDLLMissing = $false
        Add-MSALPrereq
    }

    if(-not $mod)
    {
        [System.Windows.MessageBox]::Show("Failed to install the $moduleToInstall module`n`n$($installError)Try installing the module manually and then restart the appliaction.", "Failed!", "OK", "Error") | Out-Null
    }
    elseif($mod -and $script:MSALDLLMissing)
    {
        [System.Windows.MessageBox]::Show("The $moduleToInstall module was installed successfully`n`nBut the app failed to load the MSAL DLL.`n`nPlease reastart the app and try again", "Failed!", "OK", "Warning") | Out-Null
    }
    else
    {
        [System.Windows.MessageBox]::Show("The $moduleToInstall was installed successfully!`n", "Success!", "OK", "Info") | Out-Null
    }
    Show-ModalObject
}

function Add-MSALPrereq
{
    $msalPath = ""
    
    # Path stored in settings
    $msalPath = (Get-SettingValue "MSALDLL")
    if($msalPath -and ([IO.File]::Exists($msalPath)) -eq $false -and ([IO.Path]::GetFileName($msalPath)) -ne "Microsoft.Identity.Client.dll")
    {
        Write-Log "Microsoft.Identity.Client.dll file is either missing or pointing to the wrong file name"
        $msalPath = ""
    }

    # Check if located in app folder
    $tmpPath = "$($global:AppRootFolder)\Microsoft.Identity.Client.dll"
    if(-not $msalPath -and ([IO.File]::Exists($tmpPath)))
    {
        $msalPath = $tmpPath
    }

    # Check Az module
    if(-not $msalPath)
    {
        $module = Get-Module Az.Accounts -ListAvailable
        if($module)
        {
            # Use the latest version and first path in case it is install for both the user and device
            $module = $module | Sort -Property Version -Descending | Select -First 1
            $tmpPath = (([IO.Path]::GetDirectoryName(($module.Path | Select -First 1 ))) + "\PreloadAssemblies\Microsoft.Identity.Client.dll")
            if(([IO.File]::Exists($tmpPath)))
            {
                $msalPath = $tmpPath
            }
        }
    }

    # Check MSAL.PS module
    if(-not $msalPath)
    {
        $module = Get-Module MSAL.PS -ListAvailable
        $module = $module | Sort -Property Version -Descending | Select -First 1
        $folderMSAL = Get-ChildItem -Path ([IO.Path]::GetDirectoryName($module.Path)) -Filter "Microsoft.Identity.Client*" | ?{ $_.PSIsContainer } | Sort -Property Version -Descending | Select -First 1
        if([IO.File]::Exists(($folderMSAL.FullName + "\net45\Microsoft.Identity.Client.dll")))
        {
            $msalPath = ($folderMSAL.FullName + "\net45\Microsoft.Identity.Client.dll")
        }
    }

    if(-not $msalPath)
    {
        $script:MSALDLLMissing = $true
        Write-Log "Could not find Microsoft.Identity.Client.dll. Install the latest Az or MSAL.PS module or download MSAL library" 3
        return 
    }

    $fiLoaded = $null
    if(("Microsoft.Identity.Client.TokenCache" -as [type]))
    {
        $fiLoaded = [IO.FileInfo]"$([Microsoft.Identity.Client.TokenCache].Assembly.Location)"
    } 

    $fi = [IO.FileInfo]$msalPath

    [System.Collections.Generic.List[string]] $RequiredAssemblies = New-Object System.Collections.Generic.List[string]
    if($fiLoaded -and $fiLoaded.VersionInfo.FileVersion -ne $fi.VersionInfo.FileVersion)
    {
        Write-Log "Wrong version of MSAL.DLL is loaded - Version $($fiLoaded.VersionInfo.FileVersion). DLL: $($fiLoaded.FullName)" 3
        Write-Log "Expected version: $($fi.VersionInfo.FileVersion) from DLL $msalPath" 3
        Write-Log "Some MSAL features might not work!" 3
        Write-Log "This could happen if another version of MSAL.DLL was loaded beforethe script tried to load it" 3
        $RequiredAssemblies.Add($fiLoaded.FullName)
        $script:msalFile = $fiLoaded.FullName
    }
    else
    {    
        Write-Log "Using MSAL file $msalPath. Version: $($fi.VersionInfo.FileVersion)"
        [void][System.Reflection.Assembly]::LoadFile($msalPath)
        $RequiredAssemblies.Add($msalPath)
        $script:msalFile = $msalPath
    }
    $RequiredAssemblies.Add('System.Security.dll')

    try
    {
        Add-Type -Path ($global:AppRootFolder + "\CS\TokenCacheHelperEx.cs") -ReferencedAssemblies $RequiredAssemblies
    }
    catch
    {
        $global:SkipTokenCacheHelperEx = $true
        Write-LogError "Failed to compile TokenCacheHelperEx. The access token will not be cached. Check write access to the CS folder and ASR policies" $_.Exception
    }
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
}

function Connect-MSALClientApp
{
    param($clientId, $tenantId, $secret, $Certificate)
    $scopes = [String[]]".default"
        
    if(-not $script:MSALApp)
    {
        $authority = "https://login.microsoftonline.com/$tenantId"
        #$redirectUri = "http://localhost"
        
        if($secret)
        {
            $ClientApplicationBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($clientId).WithClientSecret($secret).WithAuthority([URI]::new($authority)) #.WithRedirectUri($redirectUri)
        }
        elseif($Certificate)
        {
            $f = [System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly
            $cert = $null
            # Try LocalMachine store first, if not found try also CurrentUser store
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "LocalMachine")
            $null = $store.Open($f)
            $cert = $store.Certificates | Where-Object {$_.Thumbprint -eq $Certificate}
            $null = $store.Close()
            if($null -eq $cert)
            {
                $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
                $null = $store.Open($f)
                $cert = $store.Certificates | Where-Object {$_.Thumbprint -eq $Certificate}
                $null = $store.Close()
            }

            if($null -eq $cert)
            {
                Write-LogError "Could not find a certificate with thumbprint '$($Certificate)' in LocalMachine or CurrentUser store"
                return
            }  
            $ClientApplicationBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($clientId).WithCertificate($cert).WithAuthority([URI]::new($authority)) #.WithRedirectUri($redirectUri)
        }
        else 
        {
            return
        }
        Add-MSALProxy $ClientApplicationBuilder
        $script:MSALApp = $ClientApplicationBuilder.Build()
    }

    if($script:MSALApp)
    {        
        $accessTokenRequest = $script:MSALApp.AcquireTokenForClient($scopes)
        $global:MSALToken = Get-MsalAuthenticationToken $accessTokenRequest        
    }
}

function Get-MsalAuthenticationToken
{
    param($aquireTokenObj)

    $script:authenticationFailure = $null
    $script:errorInfo = $null
    $authResult = $null
    try 
    {        
        $tokenSource = New-Object System.Threading.CancellationTokenSource
        $taskAuthenticationResult = $aquireTokenObj.ExecuteAsync($tokenSource.Token)
        try 
        {
            while (!$taskAuthenticationResult.IsCompleted) 
            {
                # Login hung on rare occations
                # Workaround: Added DoEvents
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Seconds 1
            }
        }
        finally 
        {
            if (-not $taskAuthenticationResult.IsCompleted) 
            {
                $tokenSource.Cancel()
            }
            $tokenSource.Dispose()
        }

        ## Parse task results
        if ($taskAuthenticationResult.IsFaulted) 
        {
            # ToDo Check if: $taskAuthenticationResult.Exception.InnerException -is [Microsoft.Identity.Client.MsalUiRequiredException]
            if($taskAuthenticationResult.Exception.InnerException.ResponseBody)
            {
                try 
                {
                    $script:errorInfo = $taskAuthenticationResult.Exception.InnerException.ResponseBody | ConvertFrom-Json
                }
                catch { }
            }
            $script:authenticationFailure = ?? $taskAuthenticationResult.Exception.InnerException $taskAuthenticationResult.Exception
            if($script:errorInfo.error_description)
            {
                Write-LogError "Failed to login. Error: $($script:errorInfo.error). Description: $($script:errorInfo.error_description)" 3
            }
            else
            {
                Write-LogError "Failed to login" (?? $taskAuthenticationResult.Exception.InnerException $taskAuthenticationResult.Exception)
            }
        }
        if ($taskAuthenticationResult.IsCanceled) 
        {
            Write-Log "The login was canceled" 2
        }
        else 
        {
            $authResult = $taskAuthenticationResult.Result
        }
    }
    catch 
    {
        $script:authenticationFailure = ?? $_.Exception.InnerException $_.Exception
        Write-LogError "Failed to authenticate" (?? $_.Exception.InnerException $_.Exception)
    }
    $authResult
}

function Add-MSALProxy
{
    param($appBuilder)

    $proxy = Get-SettingValue "ProxyURI"
    if($proxy) 
    {    
        Write-Log "Use proxy $proxy"        
        if(-not ("HttpFactoryWithProxy" -as [type]))
        {                
            try
            {
                Write-Log "Add type HttpFactoryWithProxy"
                [System.Collections.Generic.List[string]] $RequiredAssemblies = New-Object System.Collections.Generic.List[string]
                $RequiredAssemblies.Add($script:msalFile)
                $RequiredAssemblies.Add('System.Net.Http.dll')
                $RequiredAssemblies.Add('System.Net.Primitives.dll')

                Add-Type -Path ($global:AppRootFolder + "\CS\HttpFactoryWithProxy.cs") -ReferencedAssemblies $RequiredAssemblies
            }
            catch
            {
                Write-LogError "Failed to compile HttpFactoryWithProxy" $_.Exception
            }        
        }

        try
        {
            $hcf = [HttpFactoryWithProxy]::new($proxy)
            [void] $appBuilder.WithHttpClientFactory($hcf)
        }
        catch
        {
            Write-LogError "Failed to set proxy for MSAL" $_.Exception
        }
    }
}
function Get-MSALLoginEnvironment
{
    $loginEnv = $script:lstAADEnvironments | Where value -eq (Get-Setting "" "MSALCloudType" "public")
    return (?? $loginEnv.URL "login.microsoftonline.com")
}
function Get-MSALApp
{
    param($appInfo, $loginHint)

    $msalApp = $script:MSALAllApps | Where { $_.ClientId -eq  $appInfo.ClientID  -and (-not $appInfo.RedirectUri -or $_.AppConfig.RedirectUri -eq $appInfo.RedirectUri)}
    
    $tenant = ?? $appInfo.TenantId "organizations"
    
    if($loginHint.Environment)
    {
        $authority = "https://$($loginHint.Environment)/$tenant/"
    }
    elseif($appInfo.Authority)
    {
        $authority = $appInfo.Authority
    }
    else
    {
        $authority = "https://$((Get-MSALLoginEnvironment))/$tenant/"
    }

    if(-not $msalApp -or $msalApp.Authority -ne $authority)
    {
        Write-Log "Add MSAL App $($appInfo.ClientID) $authority"
        $appBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($appInfo.ClientID)

        [void]$appBuilder.WithAuthority($authority)
        #if($appInfo.TenantId) { [void]$appBuilder.WithAuthority("https://$((?? $loginHint.Environment (Get-MSALLoginEnvironment)))/$($appInfo.TenantId)/") }
        #elseif ($appInfo.Authority) { [void]$appBuilder.WithAuthority($appInfo.Authority) }

        if($appInfo.RedirectUri) { [void]$appBuilder.WithRedirectUri($appInfo.RedirectUri) }

        [void] $appBuilder.WithClientName("CloudAPIPowerShellManagement") 
        [void] $appBuilder.WithClientVersion($PSVersionTable.PSVersion)

        Add-MSALProxy $appBuilder   
        
        # Ceck if correct version...
        #$appBuilder.WithMultiCloudSupport($true)        
        
        $msalApp = $appBuilder.Build()

        if($global:SkipTokenCacheHelperEx -ne $true -and (Get-SettingValue "CacheMSALToken"))
        {
            [TokenCacheHelperEx]::EnableSerialization($msalApp.UserTokenCache, "%LOCALAPPDATA%\CloudAPIPowerShellManagement\msalcahce.bin3")
        }
        $script:MSALAllApps += $msalApp
    }
    return $msalApp
}

function Get-MSALAppAuthority
{
    try
    {
        ([uri]$global:MSALApp.Authority).Authority
    }
    catch
    {
        Get-MSALLoginEnvironment
    }
}

function Connect-MSALUser
{
    param(
        #[Parameter(Mandatory = $false, ParameterSetName = 'Silent')]
        [switch]
        $Silent,

        #[Parameter(Mandatory = $false, ParameterSetName = 'Silent')]
        [switch]
        $ForceRefresh,

        #[Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [switch]
        $Interactive,

        [switch]
        $ClientApp,

        $Account,

        [switch]
        $ShowMenu,

        $Tenant
    )

    if($global:hideUI -eq $true)
    {
        if($global:AzureAppId -and $global:ClientSecret -and $global:TenantId)
        {
            Connect-MSALClientApp $global:AzureAppId $global:TenantId -secret $global:ClientSecret 
        }
        elseif($global:AzureAppId -and $global:ClientCert -and $global:TenantId)
        {
            Connect-MSALClientApp $global:AzureAppId $global:TenantId -certificate $global:ClientCert
        }
        else
        {
            Write-Log "Azure AppId, Tenant Id and Sercret/Cert must be specified for batch jobs" 3
        }
        
        return 
    }

    # No login during first time the app is started
    if($global:FirstTimeRunning -and $global:MainAppStarted -eq $false) { return }

    Write-LogDebug "Authenticate"

    if($global:MainAppStarted -eq $false) 
    { 
        $script:AppLogin = (Get-SettingValue "GraphAzureAppLogin") -or ($global:TenantId -and $global:AzureAppId -and ($global:ClientSecret -or $global:ClientCert))
    }

    if($script:AppLogin) 
    {
        if($global:MSALToken -and $global:MSALToken.ExpiresOn.LocalDateTime.Ticks -gt ((Get-Date).AddMinutes(-5)).Ticks)
        {
            return
        }

        # Get login info for silent job from settings
        if(-not $global:AzureAppId) { $global:AzureAppId = Get-SettingValue "GraphAzureAppId" -TenantID $global:TenantId }
        if(-not $global:ClientSecret -and -not $global:ClientCert) { $global:ClientSecret = Get-SettingValue "GraphAzureAppSecret" -TenantID $global:TenantId }
        if(-not $global:ClientSecret -and -not $global:ClientCert) { $global:ClientCert = Get-SettingValue "GraphAzureAppCert" -TenantID $global:TenantId }
        
        if($global:AzureAppId -and $global:ClientSecret -and $global:TenantId)
        {
            Connect-MSALClientApp $global:AzureAppId $global:TenantId -secret $global:ClientSecret 
        }
        elseif($global:AzureAppId -and $global:ClientCert -and $global:TenantId)
        {
            Connect-MSALClientApp $global:AzureAppId $global:TenantId -certificate $global:ClientCert
        }
        else
        {
            Write-Log "Azure AppId, Tenant Id and Sercret/Cert must be specified for App logins" 3
        }

        Invoke-MSALAuthenticationUpdated $global:MSALToken
        
        return
    }


    if($ShowMenu -eq $true -and ((Get-SettingValue "AzureADLoginMenu") -eq $true))
    {
        if((Show-MSALLoginMenu) -eq $false) { return }
        $global:MSALGraphEnvironment = $null
    }

    if(-not $global:appObj.ClientId)
    {
        Write-Log "Application id is missing. Cannot authenticate" 3
        return
    }

    if ($global:SkipTokenCacheHelperEx -ne $true -and -not ("TokenCacheHelperEx" -as [type])) 
    {
        Add-MSALPrereq
    }

    $curTicks = $global:MSALToken.ExpiresOn.LocalDateTime.Ticks

    $currentLoggedInUserApp = ($global:MSALToken.Account.HomeAccountId.Identifier + $global:MSALToken.TenantId + $global:MSALApp.ClientId)
    $currentLoggedInUserId = $global:MSALToken.Account.HomeAccountId.Identifier
    if($Interactive -eq $true)
    {
        Clear-MSALCurentUserVaiables
        $global:MSALToken = $null
    }

    $global:MSALApp = Get-MSALApp $global:appObj $Account
    $loginHint = ""

    $global:MSALAccounts = $global:MSALApp.GetAccountsAsync().GetAwaiter().GetResult()
    if($Account)
    {
        $userName = ?? $Account.UserName $Account
        $loginHint = $global:MSALAccounts | Where UserName -eq $userName
        if($global:MSALToken -and $global:MSALToken.Account.UserName -ne $userName)
        {
            # We're logging in with someone else...
            Clear-MSALCurentUserVaiables
            $global:MSALToken = $null
        }
    }

    # If we force interactive login then skip setting loginHint to force the user to select account
    if(-not $loginHint -and $Interactive -ne $true)
    {
        if($global:MSALAccounts)
        {
            if($global:MSALToken)
            {
                # Make sure we are logging in with the current user
                $loginHint = $global:MSALAccounts | Where { $_.HomeAccountId.Identifier -eq $global:MSALToken.Account.HomeAccountId.Identifier }
            }
            else
            {
                $lastUser = 
                $lastUserId = Get-Setting "" "LastLoggedOnUserId"
                if($lastUserId)
                {
                    # Try to get user based on Id - to allow alias login...
                    $loginHint = $global:MSALAccounts | Where { $_.HomeAccountId.ObjectId -eq $lastUserId }
                }
                if(-not $loginHint)
                {   
                    $lastUser = Get-Setting "" "LastLoggedOnUser"
                    if($lastUser)
                    {
                        $loginHint = $global:MSALAccounts | Where { $_.HomeAccountId.Identifier -eq $lastUser }
                    }
                }
                if(-not $loginHint)
                {   
                    # Try with the first user in the list
                    $loginHint = $global:MSALAccounts | Select -First 1
                }
            }
        }
    }

    if($ForceRefresh -eq $true)
    {
        $global:MSALGraphEnvironment = $null
    }
    
    $tenantId = ?? $global:MSALTenantId $global:appObj.TenantId

    Set-MSALGraphEnvironment $loginHint $tenantId
    $useDefaultPermissions = (Get-SettingValue "UseDefaultPermissions" -TenantID (?? $tenantId $loginHint.HomeAccountId.TenantId)) 

    # Always login with default scopes
    # The app will check if there are any missing scopes
    # Full Consent prompt will be triggered if app is not approved in the environment
    # Consent prompt for additional scopes will be forced or available depending on the 'Use Default Permissions' setting
    [string[]] $Scopes = "https://$($global:MSALGraphEnvironment)/.default"
    $useDefaultPermissions = ($useDefaultPermissions -eq $true -or ($global:currentViewObject.ViewInfo.Permissions | measure).Count -eq 0)

    $prompConsent = $false
    $authResult  = $null

    try
    {
        #########################################################################################################
        ### Silent Login
        #########################################################################################################
        if($loginHint -and $Interactive -ne $true)
        {
            $aquireTokenObj = $global:MSALApp.AcquireTokenSilent($Scopes, $loginHint)
            if($ForceRefresh) { [void]$aquireTokenObj.WithForceRefresh($ForceRefresh) }
            if ($tenantId) { [void]$aquireTokenObj.WithAuthority("https://$((Get-MSALAppAuthority))/$($tenantId)/")  } 
            else { [void]$aquireTokenObj.WithAuthority($global:MSALApp.Authority) }

            $authResult = Get-MsalAuthenticationToken $aquireTokenObj

            if($script:authenticationFailure -and $script:authenticationFailure -isnot [Microsoft.Identity.Client.MsalUiRequiredException])
            {
                Write-Log "Authentication failed but not with UI required. Skipping futher authentication"
                # Force the user to click Login
                # Is this the best way to handle this? Could happen when connection is lost etc.
                $global:MSALToken = $null                
                Get-MSALUserInfo 
                Clear-MSALCurentUserVaiables
                return
            }
            
            if($authResult -and $authResult.ExpiresOn.LocalDateTime.Ticks -ne $curTicks)
            {        
                Write-Log "$($authResult.Account.UserName) authenticated successfully (Silent). CorrelationId: $($global:MSALToken.CorrelationId)"
            }
            else
            {        
                Write-LogDebug "$($authResult.Account.UserName) authenticated successfully (Silent). CorrelationId: $($global:MSALToken.CorrelationId)"
            }

            #AADSTS65001
            if($script:authenticationFailure.Classification -eq "ConsentRequired")
            {
                # Will this ever happen? Cached credentials but original app consent removed...
                $prompConsent = $true
            }
            elseif($authResult -and $useDefaultPermissions -eq $false -and ($null -eq $currentLoggedInUserId -or $currentLoggedInUserId -ne $global:MSALToken.Account.HomeAccountId.Identifier))
            {
                # If 'Use Default Permissions' is not checked and new user logs in...
                # Check if there are any missing permissions
                # Only prompt for consent if the app is fully started
                # "Cancel" login if the app is starting
                Get-MSALMissingScopes $authResult
                if(($script:missingPermissions | measure).Count -gt 0)
                {                    
                    if($global:MainAppStarted)
                    {                    
                        # App started...force full consent prompt
                        $tmpAuthResult = Start-MSALConsentPrompt -PassThru -authToken $authResult
                        if($tmpAuthResult)
                        {
                            # Consent successfull so update the authResult with new token
                            $authResult = $tmpAuthResult
                        }
                        else
                        {
                            # Consent cancelled by the user or failed...continue with successful silent login
                            # Note: This will not have full permissions
                            if($currentLoggedInUserApp -eq ($global:MSALToken.Account.HomeAccountId.Identifier + $global:MSALToken.TenantId + $global:MSALApp.ClientId))
                            {
                                # Only need to update if it is the same user
                                # If it is a new user, the code below will take care of this e.g. this could happen if clicked on cached user
                                Invoke-MSALCheckObjectViewAccess $authResult
                            }
                        }
                    }
                    else
                    {
                        # Silent login was successfull but cancel it since Use Default Permission is set to false
                        # One or more permissions are missing
                        # This will force an additional login that will prompt for consent
                        Write-Log "Silent login was successful but one or more scopes are missing. Aborting login to force Consent Prompt" 2
                        $authResult = $null
                        Get-MSALUserInfo 
                        Clear-MSALCurentUserVaiables
                        return
                    }
                }                
            }
        }
    }
    catch 
    {
        Write-LogError "Failed to perform silent login" $_.Exception
    }

    # Interactive login is only allowed once the app has started. Skip if silent login failed during startup
    if($global:MainAppStarted -and ((-not $authResult -and $Silent -ne $true) -or $prompConsent))
    {
        #########################################################################################################
        ### Interactive Login
        #########################################################################################################
        Write-Log "Initiate interactive logon"

        if($useDefaultPermissions -eq $false) 
        {
            [string[]]$Scopes = Get-MSALRequiredScopes            
        }

        Write-Log "Scopes: $(($Scopes -join ","))"
        $loginHintName = (?? $loginHint.Username $loginHint)

        $aquireTokenObj = $global:MSALApp.AcquireTokenInteractive($Scopes)

        if ($tenantId)
        {
            Write-Log "Tenant id: $tenantId"
            [void]$aquireTokenObj.WithAuthority("https://$((Get-MSALAppAuthority))/$tenantId/")
        }
        else
        {
            Write-Log "Authority: $($global:MSALApp.Authority)"
            [void]$aquireTokenObj.WithAuthority($global:MSALApp.Authority)
        }

        if($loginHintName) 
        {
            Write-Log "Login hint: $loginHintName" 
            [void]$AquireTokenObj.WithLoginHint($loginHintName) 
        }
        
        if($script:authenticationFailure.Claims) 
        {
            Write-Log "Login claims: $($script:authenticationFailure.Claims))" 
            [void]$AquireTokenObj.WithClaims($script:authenticationFailure.Claims) 
        }

        [IntPtr]$ParentWindow = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle
        if ($ParentWindow)
        {
            [void]$aquireTokenObj.WithParentActivityOrWindow($ParentWindow)            
        }

        # If we need a consent (e.g. App is not approved in the environment)
        if ($script:authenticationFailure.Classification -eq "ConsentRequired") 
        {
            Write-Log "Interactive login with Consent prompt" 
            [void]$aquireTokenObj.WithPrompt([Microsoft.Identity.Client.Prompt]::Consent) 
        }
        elseif(-not $loginHintName)
        {
            Write-Log "Interactive login with Select account prompt" 
            [void]$AquireTokenObj.WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount)
        }

        $authResult = Get-MsalAuthenticationToken $aquireTokenObj
        if($authResult)
        {        
            Write-Log "$($authResult.Account.UserName) authenticated successfully (Interactively). CorrelationId: $($authResult.CorrelationId)"
        }
    }

    if($currentLoggedInUserId -ne $authResult.Account.HomeAccountId.Identifier)
    {
        $script:AccessableTenants = $null
        if($authResult -and (Get-SettingValue "GetTenantList" -TenantID $authResult.Account.HomeAccountId.TenantId) -eq $true)
        {
            #########################################################################################################
            ### Get tenant list
            #########################################################################################################
            try 
            {
                Write-Log "Get tenant list"
                
                # Can we reuse the app used for login?
                $appBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($global:appObj.ClientID)
                if($tenantId) { [void]$appBuilder.WithAuthority("https://$((Get-MSALAppAuthority))/$($tenantId)") }
                else { [void]$appBuilder.WithAuthority($global:MSALApp.Authority) }
                if($global:appObj.RedirectUri) { [void]$appBuilder.WithRedirectUri($global:appObj.RedirectUri) }     
                
                Add-MSALProxy $appBuilder

                $app = $appBuilder.Build()

                if((Get-SettingValue "CacheMSALToken"))
                {
                    [TokenCacheHelperEx]::EnableSerialization($app.UserTokenCache, "%LOCALAPPDATA%\CloudAPIPowerShellManagement\msalcahce.bin3")
                }

                ### Silent login
                $tmpScope = [string[]]"https://management.azure.com/user_impersonation"
                $tmpResults = Get-MsalAuthenticationToken ($app.AcquireTokenSilent($tmpScope, $authResult.Account))
                if(-not $tmpResults -and $global:MainAppStarted -and $script:authenticationFailure -and $script:authenticationFailure -is [Microsoft.Identity.Client.MsalUiRequiredException])
                {
                    ### Interactive login
                    $AquireTokenObj = $app.AcquireTokenInteractive($tmpScope)
                    #[void]$AquireTokenObj.WithAccount($authResult.Account)
                    [void]$AquireTokenObj.WithLoginHint($authResult.Account.Username)
                    [void]$AquireTokenObj.WithPrompt([Microsoft.Identity.Client.Prompt]::NoPrompt) 
                    $tmpResults = Get-MsalAuthenticationToken $AquireTokenObj
                }

                if($tmpResults)
                {
                    $Headers = @{
                        'Content-Type' = 'application/json'
                        'Authorization' = "Bearer " + $tmpResults.AccessToken
                        'ExpiresOn' = $tmpResults.ExpiresOn
                    }

                    $params = @{}
                    $proxyURI = Get-ProxyURI
                    if($proxyURI)
                    {
                        $params.Add("proxy", $proxyURI)
                        $params.Add("UseBasicParsing", $true)
                    }

                    $ret = Invoke-RestMethod "https://management.azure.com/tenants?api-version=2020-01-01" -Headers $Headers @params
                    if($ret)
                    {
                        $script:AccessableTenants = $ret.Value
                    }
                }
            }
            catch { }
        }
    }
    
    Write-LogDebug "Authentication finished $($authResult.Account.UserName)"

    $global:MSALToken = $authResult

    if($currentLoggedInUserApp -ne ($global:MSALToken.Account.HomeAccountId.Identifier + $global:MSALToken.TenantId + $global:MSALApp.ClientId))
    {
        if($authResult) 
        {
            Save-Setting "" "LastLoggedOnUser" $authResult.Account.UserName
            Save-Setting "" "LastLoggedOnUserId" $authResult.Account.HomeAccountId.ObjectId
        }
        Invoke-MSALAuthenticationUpdated $authResult
        <#
        Write-LogDebug "User, tenant or app has changed"
        Get-MSALUserInfo
        if($authResult)
        {
            Invoke-MSALCheckObjectViewAccess $authResult
        }        
        Invoke-ModuleFunction "Invoke-GraphAuthenticationUpdated"
        #>
    }
}

function local:Invoke-MSALAuthenticationUpdated
{
    param($authResult)

    Write-LogDebug "User, tenant or app has changed"
    $script:jwtAccessToken = Get-JWTtoken $global:MSALToken.AccessToken
    $script:jwtIdToken = Get-JWTtoken $global:MSALToken.IdToken

    Get-MSALUserInfo
    if($authResult)
    {
        Invoke-MSALCheckObjectViewAccess $authResult
    }        
    Invoke-ModuleFunction "Invoke-GraphAuthenticationUpdated"
}

function Start-MSALConsentPrompt
{
    param([switch]$PassThru, $authToken)
    Write-Log "Initiate consent prompt"

    if(($script:missingPermissions | measure).Count -eq 0) { return } 

    if(-not $authToken -and $global:MSALToken)
    {
        $authToken = $global:MSALToken
    }

    [string[]] $Scopes = $script:missingPermissions

    $loginHintName = $authToken.Account.UserName
    $tenantId = $authToken.TenantId

    if(-not $loginHintName -or -not $tenantId)
    {
        return
    }    

    $aquireTokenObj = $global:MSALApp.AcquireTokenInteractive($Scopes)
    [void]$aquireTokenObj.WithAuthority("https://$((Get-MSALAppAuthority))/$tenantId/")
    [void]$AquireTokenObj.WithLoginHint($loginHintName) 
    
    [IntPtr]$ParentWindow = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle
    if ($ParentWindow)
    {
        [void]$aquireTokenObj.WithParentActivityOrWindow($ParentWindow) 
    }
    [void]$aquireTokenObj.WithPrompt([Microsoft.Identity.Client.Prompt]::NoPrompt)
    
    Write-Log "Consent prompt for the following scopes: $(($script:missingPermissions -join ","))" 
    [void] $aquireTokenObj.WithExtraScopesToConsent(([string[]]$script:missingPermissions))
    
    $authResult = Get-MsalAuthenticationToken $aquireTokenObj
    if($authResult)
    {
        Write-Log "Consent for additional scopes added successfully"
        if($PassThru -eq $true)
        {
            # The calling function will take care of the authResult
            $authResult
        }
        else
        {
            $global:MSALToken = $authResult
            Invoke-MSALCheckObjectViewAccess $authResult
        }
    }
}

function Invoke-MSALCheckObjectViewAccess
{
    param($authToken)

    if(-not $authToken -and $global:MSALToken)
    {
        $authToken = $global:MSALToken
    }

    if($authToken)
    {
        Get-MSALMissingScopes $authToken
    }

    if($script:userEllipsGrid -and ($script:missingPermissions | measure).Count -eq 0)
    {
        Set-XamlProperty $script:userEllipsGrid "lnkRequestConsent" "Visibility" "Collapsed"
    }
    elseif($script:userEllipsGrid)
    {
        Set-XamlProperty $script:userEllipsGrid "lnkRequestConsent" "Visibility" "Visible"
    }

    $accessToken = $null
    if($authToken)
    {
        $accessToken = Get-JWTtoken $authToken.AccessToken
    }

    $curPermissions = $null
    if($accessToken.Payload.idtyp -eq "app")
    {
        $curPermissions = $accessToken.Payload.roles
    }
    elseif($accessToken.Payload.scp)
    {
        $curPermissions = $accessToken.Payload.scp.Split(" ")
    }    

    foreach($viewObjInfo in ($global:viewObjects | Where { $_.ViewInfo.AuthenticationID -eq "MSAL" }))
    {
        $viewObjInfo = $global:viewObjects | Where { $_.ViewInfo.Id -eq $global:EMViewObject.Id }
        
        if($viewObjInfo)
        {
            if($authToken)
            {
                if($curPermissions)
                {
                    foreach($viewItem in $viewObjInfo.ViewItems)
                    {
                        $full = 0
                        $partial = 0
                        $missingScopes = @()
                        

                        foreach($permission in $viewItem.Permissons)
                        {
                            if($curPermissions -contains $permission)
                            {
                                $full++
                                continue
                            }
                            # Check read access
                            $permissionRead = $null 
                            $arrTemp = $permission.Split('.') 
                            if($arrTemp[1] -eq "ReadWrite")
                            {
                                $arrTemp[1] = "Read"
                                $permissionRead = $arrTemp -join "."
                                if($null -ne $permissionRead -and $curPermissions -contains $permissionRead)
                                {
                                    # ReadWrite permission required but the user only has Read
                                    $partial++
                                    continue                                
                                }
                            }
                            elseif($arrTemp[1] -eq "Read")
                            {                            
                                $arrTemp[1] = "ReadWrite"
                                $permissionRW = $arrTemp -join "."
                                if($null -ne $permissionRW -and $curPermissions -contains $permissionRW)
                                {
                                    # Only Read permission required but the user has ReadWrite
                                    $full++
                                    continue
                                }
                            }
                            $missingScopes += $permission
                        }
                        $hasAccess = $false
                        if($viewItem.Permissons.Count -eq $full)
                        {
                            $accessType = "Full"
                            $hasAccess = $true
                        }
                        elseif($partial -gt 0)
                        {
                            $accessType = "Limited"
                        }
                        else
                        {
                            $accessType = "None"
                        }

                        if(-not ($viewItem.PSObject.Properties | Where Name -eq "@HasPermissions"))
                        {
                            $viewItem | Add-Member -NotePropertyName "@HasPermissions" -NotePropertyValue $hasAccess
                            $viewItem | Add-Member -NotePropertyName "@AccessType" -NotePropertyValue $accessType
                            $viewItem | Add-Member -NotePropertyName "@MissingScopes" -NotePropertyValue ($missingScopes -join ",")
                            
                        }
                        else
                        {
                            $viewItem."@HasPermissions" = $hasAccess
                            $viewItem."@AccessType" = $accessType
                            $viewItem."@MissingScopes" = ($missingScopes -join ",")
                        }
                    }
                }
            }
            else
            {
                foreach($viewItem in $viewObjInfo.ViewItems)
                {
                    if(($viewItem.PSObject.Properties | Where Name -eq "@HasPermissions"))
                    {
                        $viewItem."@HasPermissions" = $null
                        $viewItem."@AccessType" = $null
                        $viewItem."@MissingScopes" = $null
                    }                    
                }
            }
        }
    }
    Show-ViewMenu
}

function Show-MSALLoginMenu
{
    $script:loginMenuForm = Initialize-Window ($global:AppRootFolder + "\Xaml\MSALLoginMenu.xaml")
    if(-not $script:loginMenuForm) { return }
    
    Set-XamlProperty  $script:loginMenuForm "cbMSALCloudType" "ItemsSource" $script:lstAADEnvironments
    Set-XamlProperty  $script:loginMenuForm "cbMSALGCCType" "ItemsSource" $script:lstGCCEnvironments

    Set-XamlProperty $script:loginMenuForm "cbMSALCloudType" "SelectedValue" (Get-Setting "" "MSALCloudType" "public")
    Set-XamlProperty $script:loginMenuForm "cbMSALGCCType" "SelectedValue" (Get-Setting "" "MSALGCCType" "gcc")

    Set-XamlProperty $script:loginMenuForm "cbMSALGCCType" "IsEnabled" ((Get-Setting "" "MSALCloudType" "public") -eq "usGov")

    Add-XamlEvent $script:loginMenuForm "cbMSALCloudType" "add_selectionChanged" {
        Set-XamlProperty $script:loginMenuForm "cbMSALGCCType" "IsEnabled" ($this.SelectedValue -eq "usGov")
    }

    Add-XamlEvent $script:loginMenuForm "btnLogin" "Add_Click" -scriptBlock ([scriptblock]{
        Save-Setting "" "MSALCloudType" (Get-XamlProperty $script:loginMenuForm "cbMSALCloudType" "SelectedValue")
        Save-Setting "" "MSALGCCType" (Get-XamlProperty $script:loginMenuForm "cbMSALGCCType" "SelectedValue")

        $script:loginMenuForm.DialogResult = $true
        $script:loginMenuForm.Close()
    })

    Add-XamlEvent $script:loginMenuForm "btnCancel" "Add_Click" -scriptBlock ([scriptblock]{
        $script:loginMenuForm.Close()
    })    

    $script:loginMenuForm.Owner = $global:window
    $script:loginMenuForm.Icon = $Window.Icon 
    return ($script:loginMenuForm.ShowDialog())
}

function Disconnect-MSALUser
{
    param($user, [switch]$force, [switch]$PassThru)

    $logout = $false
    $userLoggedOut = $false
    if(-not $user) 
    {
        $logout = $true
        if(-not $global:MSALToken.Account) { return }
        $user = $global:MSALToken.Account # Logout current user
        $global:MSALToken = $null        
        Clear-MSALCurentUserVaiables  # Only clear variables for current user
        $msg = "Do you want to remove the token from the cache?"
        $title = "Remove token?"
    }
    else
    {
        $msg = "Are you sure you want to forget user $($user.UserName)?"
        $title = "Forget user?"
    }

    # ToDo: Clear browser cache

    if($user -and $global:MSALApp -and (Get-SettingValue "CacheMSALToken"))
    {
        if($force -eq $true -or [System.Windows.MessageBox]::Show($msg, $title, "YesNo", "Question") -eq "Yes")
        {
            try 
            {
                [void]$global:MSALApp.RemoveAsync($user).GetAwaiter().GetResult()
                if($logout -eq $false)
                {
                    Write-Log "User $($user.UserName) removed from cache"
                }
                $userLoggedOut = $true
            }
            catch 
            {
                Write-LogError "Failed to remove $($user.UserName) from cache" $_.Exception
            }
        }
    }

    if($logout)
    {
        Get-MSALUserInfo
    }

    if($PassThru -eq $true)
    {
        $userLoggedOut
    }
}

function Get-MSALProfileEllipse
{
    param($size = 32, $fontSize = 20, $Color = "Blue", [Switch]$Popup, $AuthenticationProvider)

    Write-LogDebug "Create Profile Ellipse"

    if(-not $global:MSALToken -or -not $global:me)
    {
        #########################################################################################################
        ### Build login button when no user is logged on
        #########################################################################################################

        Write-LogDebug "Add login button"
        $grd = [System.Windows.Controls.Border]::new()
        $icon = Get-XamlObject ($global:AppRootFolder + "\Xaml\Icons\Logon.xaml")
        $icon.Width = $size
        $icon.Height = $size
        $grd.Background = "#01000000"
        $grd.Child = $icon

        $lnkButton = [System.Windows.Controls.Button]::new()
        $lnkButton.Content = $grd
        $lnkButton.Cursor = "Hand"
        $lnkButton.Style = $window.TryFindResource("ContentButton") 
        $lnkButton.add_Click({
            if($script:MSALDLLMissing)
            {
                Show-MSALError
                return
            }
            
            if(($global:MSALAccounts | measure).Count -eq 0)
            {
                # No cached users
                Connect-MSALUser -Interactive -ShowMenu
                if($global:curObjectType)
                {
                    Show-GraphObjects 
                }
                Write-Status ""
            }
            else
            {
                # Add list of cached users + a 'Sign in with a different account' option

                $xaml = Get-Content ($global:AppRootFolder + "\Xaml\LoginPanel.Xaml")
                $loginPanel = [Windows.Markup.XamlReader]::Parse($xaml)
                $otherLogins = $loginPanel.FindName("grdAccounts")
                foreach($account in $global:MSALAccounts)
                {
                    Add-CachedUser $account $otherLogins                             
                }           
                
                #########################################################################################################
                ### Add login button
                #########################################################################################################
                $grdAccount = [System.Windows.Controls.Grid]::new()
                $cd = [System.Windows.Controls.ColumnDefinition]::new()                
                $grdAccount.ColumnDefinitions.Add($cd)
                $cd = [System.Windows.Controls.ColumnDefinition]::new()
                $cd.Width = [double]::NaN   
                $grdAccount.ColumnDefinitions.Add($cd)

                $icon = Get-XamlObject ($global:AppRootFolder + "\Xaml\Icons\Logon.xaml")
                $icon.Width = 24
                $icon.Height = 24
                $icon.Margin = "0,0,5,0"
                $grdAccount.Children.Add($icon) | Out-Null

                $lbObj = [Windows.Markup.XamlReader]::Parse("<TextBlock $wpfNS>Sign in with a different account</TextBlock>")
                $lbObj.SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
                #$lbObj.Style = $window.TryFindResource("HoverUnderlineStyle") 
                $grdAccount.Children.Add($lbObj) | Out-Null

                $lnkButton = [System.Windows.Controls.Button]::new()
                $lnkButton.Content = $grdAccount
                $lnkButton.Style = $window.TryFindResource("LinkButton") 
                $lnkButton.Margin = "0,5,0,0"
                $lnkButton.Cursor = "Hand"
                $lnkButton.add_Click({
                    Write-Status "Logging in..."
                    Hide-Popup
                    Connect-MSALUser -Interactive -ShowMenu
                    if($global:curObjectType)
                    {
                        Show-GraphObjects 
                    }
                    Write-Status ""
                })
                
                Add-GridObject $otherLogins $lnkButton

                $loginPanel.Tag = $this.Content

                $loginPanel.Add_Loaded({param($obj, $e)
                    $point = $obj.Tag.TransformToAncestor($window).Transform([System.Windows.Point]::new(0,0));
                    [System.Windows.Controls.Canvas]::SetLeft($obj,($point.X - $obj.ActualWidth + $obj.Tag.ActualWidth)) 
                    [System.Windows.Controls.Canvas]::SetTop($obj,($point.Y + $obj.Tag.ActualHeight))
                })                
                
                Show-Popup $loginPanel
            }
        })

        return $lnkButton
    }

    #########################################################################################################
    ### Build the ellipse image for the Profile Info
    #########################################################################################################

    if($global:me.givenName -and $global:me.surname)
    {
        $initials = "$($global:me.givenName[0])$($global:me.surname[0])".ToUpper()    
    }
    elseif($global:me.userPrincipalName)
    {
        $initials = "$($global:me.userPrincipalName[0])".ToUpper()    
    }
    elseif($script:jwtAccessToken.Payload.idtyp -eq "app")
    {
        $initials = "APP"
    }

    $grd = Get-MSALUserPhotoEllips -size $size -fontSize $fontSize -Color $Color 
    
    if($Popup)
    {
        # Hide the popup when mouse button is clicked anywhere
        $grd.add_MouseLeftButtonDown(({param($obj, $e) 
            if(-not $global:grdProfileInfo) { return }
            Show-Popup $global:grdProfileInfo 
        }))

        try 
        {
            #########################################################################################################
            ### Build Profile Info for current user
            #########################################################################################################

            $global:grdProfileInfo = $null
            $xaml = Get-Content ($global:AppRootFolder + "\Xaml\ProfileInfo.Xaml")
            $global:grdProfileInfo = [Windows.Markup.XamlReader]::Parse($xaml)
            $global:grdProfileInfo.Tag = $grd
            $grd.Tag = $global:grdProfileInfo
            Set-XamlProperty $global:grdProfileInfo "txtOrganization" "Text" $global:Organization.displayName
            if($script:jwtAccessToken.Payload.idtyp -eq "app")
            {
                Set-XamlProperty $global:grdProfileInfo "txtUsername" "Text" "App Login"
            }
            else
            {
                Set-XamlProperty $global:grdProfileInfo "txtUsername" "Text" $global:me.displayName
                Set-XamlProperty $global:grdProfileInfo "txtLogonName" "Text" $global:me.userPrincipalName 
            }

            $global:tokenInfo =  Get-JWTtoken $global:MSALToken.AccessToken
            if($global:tokenInfo)
            {
                Write-LogDebug "App $($global:tokenInfo.Payload.app_displayname)"
                Set-XamlProperty $global:grdProfileInfo "txtAppName" "Text" $global:tokenInfo.Payload.app_displayname
                Set-XamlProperty $global:grdProfileInfo "txtAppId" "Text" $global:tokenInfo.Payload.appid 
            }

            # Get the elips with only the photo in a larger size for the popup info
            $tmpObj = Get-MSALUserPhotoEllips -size 64 -fontSize 32
            #$tmpObj = Get-MSALProfileEllipse -size 64 -fontSize 32
            $profileGrid =  $global:grdProfileInfo.FindName("ProfileInfo")
            if($tmpObj -and $profileGrid)
            {
                $tmpObj.SetValue([System.Windows.Controls.Grid]::RowProperty,1)
                $tmpObj.SetValue([System.Windows.Controls.Grid]::RowSpanProperty,2)
            }

            if($tmpObj)
            {
                $profileGrid.Children.Add($tmpObj) | Out-Null
            }

            if($script:jwtAccessToken.Payload.idtyp -eq "app")
            {
                $tmpObj.Visibility = "Collapsed"
            }            
        
            $global:grdProfileInfo.Add_Loaded({param($obj, $e)
                $point = $obj.Tag.TransformToAncestor($window).Transform([System.Windows.Point]::new(0,0));
                [System.Windows.Controls.Canvas]::SetLeft($obj,($point.X - $obj.ActualWidth + $obj.Tag.ActualWidth)) 
                [System.Windows.Controls.Canvas]::SetTop($obj,($point.Y + $obj.Tag.ActualHeight))
            })

            if($script:jwtAccessToken.Payload.idtyp -ne "app")
            {
                #########################################################################################################
                ### Show / Hide consent button
                #########################################################################################################
                $script:userEllipsGrid = $tmpObj
                if(($script:missingPermissions | measure).Count -eq 0)
                {
                    Set-XamlProperty $script:userEllipsGrid "lnkRequestConsent" "Visibility" "Collapsed"
                }
                Add-XamlEvent $script:userEllipsGrid "lnkRequestConsent" "add_Click" {
                    Start-MSALConsentPrompt
                }
                
                $otherLogins = $global:grdProfileInfo.FindName("grdCachedAccounts")

                #########################################################################################################
                ### Add cached users
                #########################################################################################################
                if((Get-SettingValue "SortAccountList") -eq $true)
                {
                    $accounts = $global:MSALAccounts | Sort -Property Username
                }
                else
                {
                    $accounts = $global:MSALAccounts
                }

                foreach($account in $accounts)
                {
                    # Skip current logged on user
                    if($global:MSALToken.Account.Username -eq $Account.Username -or 
                    $global:MSALToken.Account.HomeAccountId.ObjectId -eq $Account.HomeAccountId.ObjectId) { continue }

                    Add-CachedUser $account $otherLogins                
                }           
                
                #########################################################################################################
                ### Add login with another user
                #########################################################################################################
                $grdAccount = [System.Windows.Controls.Grid]::new()
                $cd = [System.Windows.Controls.ColumnDefinition]::new()                
                $grdAccount.ColumnDefinitions.Add($cd)
                $cd = [System.Windows.Controls.ColumnDefinition]::new()
                $cd.Width = [double]::NaN   
                $grdAccount.ColumnDefinitions.Add($cd)

                $icon = Get-XamlObject ($global:AppRootFolder + "\Xaml\Icons\Logon.xaml")
                $icon.Width = 24
                $icon.Height = 24
                $icon.Margin = "0,0,5,0"
                $grdAccount.Children.Add($icon) | Out-Null

                $lbObj = [Windows.Markup.XamlReader]::Parse("<TextBlock $wpfNS>Sign in with a different account</TextBlock>")
                $lbObj.SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
                #$lbObj.Style = $window.TryFindResource("HoverUnderlineStyle") 
                $grdAccount.Children.Add($lbObj) | Out-Null

                $lnkButton = [System.Windows.Controls.Button]::new()
                $lnkButton.Content = $grdAccount
                $lnkButton.Style = $window.TryFindResource("LinkButton") 
                $lnkButton.Margin = "0,5,0,0"
                $lnkButton.Cursor = "Hand"
                $lnkButton.Tag = $account
                $lnkButton.add_Click({
                    Write-Status "Logging in..."
                    Hide-Popup
                    Connect-MSALUser -Interactive -ShowMenu
                    if($global:curObjectType)
                    {
                        Show-GraphObjects 
                    }
                    Write-Status ""
                })                    

                $otherLogins = $global:grdProfileInfo.FindName("grdLoginAccount")

                Add-GridObject $otherLogins $lnkButton

                $otherLogins = $global:grdProfileInfo.FindName("grdTenantAccounts")
                
                if(($script:AccessableTenants | measure).Count -gt 1)
                {
                    #########################################################################################################
                    ### Add switch to another tenant
                    #########################################################################################################
                    $lbObj = [Windows.Markup.XamlReader]::Parse("<TextBlock $wpfNS><Bold>Tenants:</Bold></TextBlock>")
                    $lbObj.Margin = "0,5,0,0"

                    if((Get-SettingValue "SortTenantList") -eq $true)
                    {
                        $tenants = $script:AccessableTenants | Sort -Property DisplayName
                    }
                    else
                    {
                        $tenants = $script:AccessableTenants
                    }
    

                    Add-GridObject $otherLogins $lbObj
                    foreach($tenant in $tenants)
                    {
                        try
                        {
                            $tenantName = [System.Web.HttpUtility]::HtmlEncode($tenant.DisplayName)
                            $lbObj = [Windows.Markup.XamlReader]::Parse("<TextBlock $wpfNS  HorizontalAlignment=`"Stretch`"><Bold>$($tenantName)</Bold><LineBreak/>$($tenant.defaultDomain)<LineBreak/>$($tenant.tenantId)</TextBlock>")

                            if($tenant.tenantId -ne $global:MSALToken.TenantId)
                            {
                                $lbObj.Style = $window.TryFindResource("HoverUnderlineStyleWithBackground")
                                $lbObj.HorizontalAlignment = "Stretch"
                                $lnkButton = [System.Windows.Controls.Button]::new()
                                $lnkButton.Content = $lbObj
                                $lnkButton.HorizontalAlignment = "Stretch"
                                $lnkButton.Style = $window.TryFindResource("ContentButton") 
                                $lnkButton.Margin = "0,5,0,0"
                                $lnkButton.Cursor = "Hand"
                                $lnkButton.Tag = $tenant
                                $lnkButton.add_Click({
                                    Write-Status "Logging in to $($this.Tag.DisplayName)"
                                    # Set authority to selected tenant
                                    $global:MSALTenantId = $this.Tag.tenantId
                                    Hide-Popup                        
                                    Connect-MSALUser -Account ($global:MSALAccounts | Where UserName -eq $global:MSALToken.Account.Username)

                                    if($global:curObjectType)
                                    {
                                        Show-GraphObjects
                                    }
                                    Write-Status ""
                                })
                                Add-GridObject $otherLogins $lnkButton
                            }
                            else
                            {
                                $lbObj.Background = $window.TryFindResource("SelectedRowBackgroundColor")
                                $lbObj.Margin = "0,5,0,0"
                                Add-GridObject $otherLogins $lbObj
                            }                        
                        }
                        catch {}
                    }
                }
            }

            #########################################################################################################
            ### Add event handling
            #########################################################################################################
            Add-XamlEvent $tmpObj "lnkTokeninfo" "add_Click" {
                #Hide-Popup
                $tokenArr = @()
                foreach($prop in ($global:MSALToken | GM | Where MemberType -eq Property))
                {
                    if($prop.Name -in @("AccessToken", "IdToken")) { continue }
                    elseif($prop.Name -eq "Scopes") { $value = ($global:MSALToken.Scopes -join "`n")}
                    elseif($prop.Name -in @("ExpiresOn", "ExtendedExpiresOn")) { $value = $global:MSALToken."$($prop.Name)".LocalDateTime }
                    else { $value = $global:MSALToken."$($prop.Name)"}


                    $tokenArr += New-Object PSObject -Property @{
                        Name=$prop.Name
                        Value=$value
                    }
                }
            
                $dg = [System.Windows.Controls.DataGrid]::new()
                $dg.ItemsSource = ($tokenArr | Select Name, Value)
                Show-ModalForm "Token info" $dg
            }

            Add-XamlEvent $tmpObj "lnkAccessTokenInfo" "add_Click" {
                #Hide-Popup
                Show-MSALDecodedToken (Get-JWTtoken $global:MSALToken.AccessToken) "Access Token Info"
            }

            Add-XamlEvent $tmpObj "lnkIdTokenInfo" "add_Click" {
                #Hide-Popup
                Show-MSALDecodedToken (Get-JWTtoken $global:MSALToken.IdToken) "Id Token Info"
            }

            Add-XamlEvent $tmpObj "lnkForceRefresh" "add_Click" {
                Write-Status "Refreshing the token"
                Connect-MSALUser -ForceRefresh
                if(-not $global:MSALToken)
                {
                    # Refresh failed. User was logged out
                    Show-GraphObjects
                    Hide-Popup                    
                }
                else
                {                    
                    Invoke-MSALCheckObjectViewAccess $global:MSALToken
                }
                Write-Status ""
            }
            
            Add-XamlEvent $tmpObj "lnkLogout" "add_Click" {
                Hide-Popup
                Disconnect-MSALUser
                if($global:curObjectType)
                {
                    Show-GraphObjects
                }                
            }
            
            if($script:jwtAccessToken.Payload.idtyp -eq "app")
            {
                Set-XamlProperty $tmpObj "lnkLogout" "Visibility" "Collapsed"
            }
        }
        catch {
            Write-LogError "Failed to create profile information object. Error: " $_.Exception   
        }
    }

    $grd
}

function local:Add-CachedUser 
{
    param($account, $parentObj)

    try
    {
        $grdAccount = [System.Windows.Controls.Grid]::new()

        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        $grdAccount.ColumnDefinitions.Add($cd) # Login

        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        $cd.Width = [double]::NaN   
        $grdAccount.ColumnDefinitions.Add($cd) # Forget

        $grdLogin = [System.Windows.Controls.Grid]::new()
        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        $grdLogin.ColumnDefinitions.Add($cd)
        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        $cd.Width = [double]::NaN   
        $grdLogin.ColumnDefinitions.Add($cd)

        $icon = Get-XamlObject ($global:AppRootFolder + "\Xaml\Icons\LoggedOnUser.xaml")
        $icon.Width = 24
        $icon.Height = 24
        $icon.Margin = "0,0,5,0"
        $grdLogin.Children.Add($icon) | Out-Null

        $tenantName = [System.Web.HttpUtility]::HtmlEncode((Get-Setting $account.HomeAccountId.TenantId "_Name" $account.HomeAccountId.TenantId))

        $lbObj = [Windows.Markup.XamlReader]::Parse("<TextBlock $wpfNS>$($account.UserName)<LineBreak/>$($tenantName)</TextBlock>")
        $lbObj.SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
        $grdLogin.Children.Add($lbObj) | Out-Null

        $lnkButton = [System.Windows.Controls.Button]::new()
        $lnkButton.Content = $grdLogin
        $lnkButton.Style = $window.TryFindResource("LinkButton") 
        $lnkButton.Margin = "0,5,0,0"
        $lnkButton.Cursor = "Hand"
        $lnkButton.Tag = $account
        $lnkButton.add_Click({
            Write-Status "Logging in with $($this.Tag.UserName)"
            Hide-Popup
            Clear-MSALCurentUserVaiables
            Connect-MSALUser -Account $this.Tag

            if($global:curObjectType)
            {
                Show-GraphObjects
            }
            Write-Status ""
        })

        $grdAccount.Children.Add($lnkButton) | Out-Null

        # Add Forget user icon
        $icon = Get-XamlObject ($global:AppRootFolder + "\Xaml\Icons\Bin.xaml")
        $icon.Width = 16
        $icon.Height = 16
        $icon.Margin = "5,5,0,0"
        
        $lnkButton = [System.Windows.Controls.Button]::new()
        $lnkButton.ToolTip = "Forget"
        $lnkButton.Content = $icon
        $lnkButton.Style = $window.TryFindResource("LinkButton") 
        $lnkButton.Margin = "0,5,0,0"
        $lnkButton.Cursor = "Hand"
        $lnkButton.Tag = $account
        $lnkButton.SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
        $lnkButton.add_Click({
            Write-Status "Logging out $($this.Tag.UserName)"
            if((Disconnect-MSALUser $this.Tag -PassThru))
            {
                $this.Parent.Parent.Children.Remove($this.Parent)
            }
            
            Write-Status ""
        })
                            
        $grdAccount.Children.Add($lnkButton) | Out-Null                    
        
        Add-GridObject $parentObj $grdAccount
    }
    catch {}     
}

function local:Get-MSALUserPhotoEllips
{
    param($size = 32, $fontSize = 20, $Color = "Blue")

    $grd = [System.Windows.Controls.Grid]::new()

    $ellipse = [System.Windows.Shapes.Ellipse]::new()
    $ellipse.Width = $size
    $ellipse.Height = $size
    $ellipse.Fill = $Color
    $ellipse.Stroke = "#FFFF00FF"
    $ellipse.StrokeThickness = "0"

    $grd.Children.Add($ellipse) | Out-Null

    $tb = [System.Windows.Controls.TextBlock]::new()
    $tb.FontSize = $fontSize
    $tb.Foreground = "White"
    #$tb.FontFamily=""
    $tb.FontWeight = "Bold"
    #$tb.TextLineBounds="Tight"
    $tb.VerticalAlignment="Center"
    $tb.HorizontalAlignment="Center"
    #$tb.IsTextScaleFactorEnabled="False"
    $tb.Text = $initials

    $grd.Children.Add($tb) | Out-Null

    if($global:profilePhoto -and [IO.File]::Exists($global:profilePhoto))   
    {        
        Write-LogDebug "Create image"
        $img = [System.Windows.Media.Imaging.BitmapImage]::new()
        $img.BeginInit()
        $img.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $img.UriSource = [System.Uri]::new($global:profilePhoto)
        $img.EndInit()        
        $ib = [System.Windows.Media.ImageBrush]::new()
        $ib.ImageSource = $img

        $ellipse = [System.Windows.Shapes.Ellipse]::new()
        $ellipse.Width = $size
        $ellipse.Height = $size
        $ellipse.FlowDirection="LeftToRight"
        $ellipse.Fill = $ib
        $grd.Children.Add($ellipse) | Out-Null
    }

    $grd
}

function Get-MSALRequiredScopes
{
    $reqScopes = [string[]]$global:msalAuthenticator.Permissions

    $resolveRoles = ((Get-SettingValue "AzureADRoleRead" $false -TenantID (?? $tenantId $loginHint.HomeAccountId.TenantId)) -eq $true)

    if($resolveRoles -and $global:msalAuthenticator.Permissions -notcontains "RoleManagement.Read.Directory")
    {
        # Adds the required permission for reading AAD directory roles
        $reqScopes += "RoleManagement.Read.Directory"
    }

    $script:curViewPermissions = $global:currentViewObject.ViewInfo.Permissions

    foreach($tmpScope in $script:curViewPermissions)
    {
        if($reqScopes -notcontains $tmpScope) { $reqScopes += $tmpScope }
    }
    $reqScopes
}
function Get-MSALMissingScopes
{
    param($authToken)

    $reqScopes = Get-MSALRequiredScopes
    
    if(($reqScopes | measure).Count -eq 0) { return }

    $script:missingPermissions = @()

    if($script:jwtAccessToken.Payload.idtyp  -eq "app")
    {
        $curScopes = $script:jwtAccessToken.Payload.roles
    }
    else
    {
        $curScopes = $authToken.Scopes
    }

    foreach($scope in $reqScopes)
    {
        $tmpScope = $scope.Split('/')[-1]
        if($tmpScope -eq ".default") { continue }
        if($curScopes -contains $tmpScope) { continue }
        if(($curScopes -like "*/$tmpScope")) { continue }
        $arrTemp = $tmpScope.Split(".")
        if($arrTemp[1] -eq "Read")
        {
            # Check if we have "more" permissions than required eg ReadWrite when only Read is required                            
            $arrTemp[1] = "ReadWrite"
            $permissionRW = $arrTemp -join "."
            if($authToken.Scopes -contains $permissionRW) { continue }
            if(($authToken.Scopes -like "*/$permissionRW")) { continue }            
        }        
        $script:missingPermissions += $tmpScope
    }

    if($script:missingPermissions.Count -gt 0)
    {
        Write-Log "Missing scopes: $(($script:missingPermissions -join ","))" 2
    }
}

function Show-MSALDecodedToken {
    param (
        $tokenData,
        $title
    )

    if(-not $tokenData.Header) { return }

    $tokenArr = @()
    foreach($prop in ($tokenData.Header | GM | Where MemberType -eq NoteProperty))
    {
        $tokenArr += New-Object PSObject -Property @{
            Name=$prop.Name
            Value=$tokenData.Header."$($prop.Name)"
        }
    }

    foreach($prop in ($tokenData.Payload | GM | Where MemberType -eq NoteProperty))
    {
        if($prop.Name -in @("exp","iat","nbf","xms_tcdt"))
        {
            $value =[datetime]::new(1970, 1, 1, 0, 0, 0, 0, "UTC").AddSeconds(($tokenData.Payload."$($prop.Name)")).ToLocalTime()
        }
        elseif($prop.Name -in @("acrs","amr"))
        {
            $value = $tokenData.Payload."$($prop.Name)" -join ";"
        }
        elseif($prop.Name -in @("wids"))
        {
            if(-not $script:aadRoles)
            {
                # This will fail if RoleManagement.Read.Directory permission is not granted. Use -NoError to hide any problems
                $script:aadRoles = (Invoke-GraphRequest -url "/directoryRoles?`$select=roleTemplateId,displayName" -ODataMetadata "minimal" -Noerror).value
            }
            $wids = @()
            foreach($wid in $tokenData.Payload."$($prop.Name)")
            {
                $text = $wid
                $role = ($script:aadRoles | where roleTemplateId -eq $wid)
                if($role)
                {
                    $text = ($text + " ($($role.displayName))")
                }
                $wids += $text
            }
            $value = $wids -join "`n"
            #$value = $tokenData.Payload."$($prop.Name)" -join "`n"
        }
        elseif($prop.Name -in @("scp"))
        {
            $value = $tokenData.Payload."$($prop.Name)" -replace " ","`n"
        }
        else
        {
            $value = $tokenData.Payload."$($prop.Name)"   
        }
        $tokenArr += New-Object PSObject -Property @{
            Name=$prop.Name
            Value=$value
        }
    }
    $dg = [System.Windows.Controls.DataGrid]::new()
    $dg.ItemsSource = ($tokenArr | Select Name, Value)
    Show-ModalForm $title $dg    
}

Export-ModuleMember -alias * -function *