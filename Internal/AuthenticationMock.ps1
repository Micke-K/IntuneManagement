# Registration and token helpers for the offline mock tenant provider
# (Classes/AuthenticationMock.ps1). The provider only exists in a session
# started with IM_MOCK_DATA pointing at a data folder - see
# Docs/MockTenant.md and Tools/Start-MockTenant.ps1.

# Delegated scopes stamped into the mock token: every permission a policy type
# declares, so access marking shows the whole menu as usable.
$script:MockTokenScopes = @(
    "Agreement.ReadWrite.All", "CloudPC.ReadWrite.All",
    "DeviceManagementApps.ReadWrite.All", "DeviceManagementConfiguration.ReadWrite.All",
    "DeviceManagementManagedDevices.ReadWrite.All", "DeviceManagementRBAC.ReadWrite.All",
    "DeviceManagementScripts.ReadWrite.All", "DeviceManagementServiceConfig.ReadWrite.All",
    "Directory.Read.All", "Group.ReadWrite.All", "Organization.ReadWrite.All",
    "Policy.Read.All", "Policy.ReadWrite.ConditionalAccess", "User.Read", "User.Read.All",
    "openid", "profile", "offline_access"
)

# Intune Administrator role template id: access marking skips the RBAC lookup
# for a token carrying it (Internal/EffectivePermissions.ps1).
$script:MockTokenIntuneAdminRoleId = "3a2c62db-5318-420d-8d74-23affee5d9d5"

# tenant.json merged over defaults, so a data folder only has to say what differs.
function Get-MockTenantProfile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)

    $profile = [ordered]@{
        TenantId    = "11111111-2222-3333-4444-555555555555"
        TenantName  = "MyLabTenant"
        Domain      = "mylabtenant.onmicrosoft.com"
        UPN         = "admin@mylabtenant.onmicrosoft.com"
        DisplayName = "Lab Admin"
        UserId      = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
        AppId       = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
        AppName     = "Microsoft Graph Command Line Tools"
    }
    $file = Join-Path $Root "tenant.json"
    if(Test-Path -LiteralPath $file) {
        try {
            $data = [IO.File]::ReadAllText($file) | ConvertFrom-Json
            foreach($prop in $data.PSObject.Properties) {
                if($profile.Contains($prop.Name) -and $null -ne $prop.Value -and "$($prop.Value)" -ne "") { $profile[$prop.Name] = [string]$prop.Value }
            }
        }
        catch { Write-Log "Mock provider: cannot read $file - using defaults ($($_.Exception.Message))" 2 }
    }
    return [PSCustomObject]$profile
}

function ConvertTo-MockBase64Url {
    param([string]$Text)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    return [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

# An unsigned JWT with the claims the product reads (scp, wids, tid, upn, oid,
# name, iat, exp, idtyp). Nothing verifies the signature - the mock never
# leaves the process - so the third segment is a placeholder.
function New-MockAccessToken {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Tenant)

    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $header = [ordered]@{ typ = "JWT"; alg = "none" }
    $payload = [ordered]@{
        aud                = "https://graph.microsoft.com"
        iss                = "https://sts.windows.net/$($Tenant.TenantId)/"
        iat                = $now
        nbf                = $now
        exp                = $now + 31536000
        app_displayname    = $Tenant.AppName
        appid              = $Tenant.AppId
        idtyp              = "user"
        name               = $Tenant.DisplayName
        oid                = $Tenant.UserId
        preferred_username = $Tenant.UPN
        scp                = ($script:MockTokenScopes -join " ")
        tid                = $Tenant.TenantId
        unique_name        = $Tenant.UPN
        upn                = $Tenant.UPN
        ver                = "1.0"
        wids               = @($script:MockTokenIntuneAdminRoleId)
    }
    $h = ConvertTo-MockBase64Url ($header | ConvertTo-Json -Compress)
    $p = ConvertTo-MockBase64Url ($payload | ConvertTo-Json -Compress)
    return "$h.$p.mock"
}

function Invoke-MockProviderInitialize {
    [CmdletBinding()]
    param()

    if(-not $env:IM_MOCK_DATA) { return }
    if(-not (Get-Command -Name Register-AuthProvider -ErrorAction SilentlyContinue)) { return }
    if(-not (Test-Path -LiteralPath $env:IM_MOCK_DATA -PathType Container)) {
        Write-Log "IM_MOCK_DATA is set but '$($env:IM_MOCK_DATA)' is not a folder - mock provider not registered" 2
        return
    }

    # A mock session is only ever meant to run on the mock, so make it the
    # session's provider unless the launcher already chose one. MSAL registers
    # later with -SetActive; the AppInitialized handler applies this variable
    # after every provider has registered, which is what makes it stick.
    if(-not $env:IM_AUTH_PROVIDER) { $env:IM_AUTH_PROVIDER = "Mock" }

    try { Register-AuthProvider -Provider ([AuthenticationMock]::new()) }
    catch { Write-LogError "Failed to register the mock tenant provider" $_.Exception }
}

Invoke-MockProviderInitialize
