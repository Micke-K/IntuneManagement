# Microsoft.Graph PowerShell SDK provider — initialization & registration.
#
# Mirrors the pattern used by Internal/AuthenticationMSALHelpers.ps1: the auth backend
# has its own init file that the module loader picks up. The init function checks
# that the SDK is installed and, if so, registers the provider into the multi-
# provider auth core.
#
# We deliberately do NOT Import-Module Microsoft.Graph.Authentication here — that
# load is expensive (hundreds of ms) and the user may not actually use this
# provider in a given session. Lazy import happens in AuthenticationMgGraph.Connect().

# Required modules. Today only Microsoft.Graph.Authentication is needed — it covers
# Connect-MgGraph / Disconnect-MgGraph / Get-MgContext. If we later add features that
# need Microsoft.Graph.Identity.DirectoryManagement (Get-MgOrganization for tenant
# display name) etc., add them here and they'll be included in the install prompt.
$script:MgGraphRequiredModules = @('Microsoft.Graph.Authentication')

# Internal: check that all required SDK modules are installed. Returns missing names.
function Get-MgGraphMissingModules {
    $missing = @()
    foreach($name in $script:MgGraphRequiredModules) {
        $found = Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if(-not $found) { $missing += $name }
    }
    return $missing
}

# Ensures every required SDK module is installed. If any are missing, prompts the
# user. Install runs at CurrentUser scope (no admin). Returns $true on success,
# $false on user-decline or install failure (caller falls back to MSAL).
function Resolve-MgGraphModule {
    [CmdletBinding()]
    param()

    $missing = Get-MgGraphMissingModules
    if($missing.Count -eq 0) { return $true }

    Write-Log "MgGraph provider needs the following PowerShell module(s): $($missing -join ', ')"

    $msg = "The Microsoft Graph PowerShell SDK module(s) below are not installed:`n`n  " +
           ($missing -join "`n  ") +
           "`n`nInstall now (Install-Module ... -Scope CurrentUser)?"
    $accepted = Confirm-UserAction -Message $msg -Caption "Install MgGraph module?"

    if(-not $accepted) {
        Write-Log "User declined MgGraph module install - MgGraph provider unavailable" 2
        return $false
    }

    foreach($name in $missing) {
        try {
            Write-Log "Installing module '$name' (Scope=CurrentUser)..."
            Install-Module -Name $name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "Installed '$name' successfully"
        }
        catch {
            Write-LogError "Failed to install '$name'. Run 'Install-Module $name -Scope CurrentUser' manually." $_.Exception
            return $false
        }
    }

    # Re-verify after install
    $stillMissing = Get-MgGraphMissingModules
    if($stillMissing.Count -gt 0) {
        Write-Log "After install attempt, still missing: $($stillMissing -join ', ')" 3
        return $false
    }

    return $true
}

function Invoke-MgGraphProviderInitialize {
    [CmdletBinding()]
    param()

    if(-not (Get-Command -Name Register-AuthProvider -ErrorAction SilentlyContinue)) {
        # AuthenticationCore.ps1 didn't load (shouldn't happen, but be defensive)
        Write-LogDebug "MgGraph provider: AuthenticationCore not available, skipping registration"
        return
    }

    # Always register — even if the SDK module isn't installed yet. The provider's
    # Connect() will prompt the user to install at first use. This gives users the
    # ability to choose MgGraph in the settings UI without first installing modules
    # manually.
    try {
        Register-AuthProvider -Provider ([AuthenticationMgGraph]::new())
    }
    catch {
        Write-LogError "Failed to register AuthenticationMgGraph provider" $_.Exception
        return
    }

    # Log the install status so it's visible in the app log without prompting.
    $missing = Get-MgGraphMissingModules
    if($missing.Count -eq 0) {
        Write-LogDebug "MgGraph provider: all required SDK modules present"
    }
    else {
        Write-Log "MgGraph provider registered but the following modules need installing on first use: $($missing -join ', ')"
    }
}

# Extracts the MSAL-v3 cache byte[] from the SDK's private InMemoryTokenCache and
# rehydrates it into a transient PublicClientApplication so we can enumerate
# cached accounts. Lives here (regular function, late-bound) rather than as a
# class method because PowerShell classes parse type references at parse time
# and would fail before MSAL DLLs are loaded.
#
# Source pattern: github.com/microsoftgraph/msgraph-sdk-powershell
#   src/Authentication/Authentication/Common/InMemoryTokenCache.cs
#
# NOTE: the cache is in-memory only — only accounts seen during the current
# PowerShell session show up. Connect-MgGraph has no -LoginHint so clicking an
# entry can't directly switch to it; the list is informational.
function Get-MgGraphCachedMsalAccounts {
    [CmdletBinding()]
    param([string]$ProviderId = "MgGraph")

    $sessionType = "Microsoft.Graph.PowerShell.Authentication.GraphSession" -as [type]
    if(-not $sessionType) { return @() }
    $session = $sessionType::Instance
    if(-not $session -or -not $session.InMemoryTokenCache) { return @() }

    # Pull the private _tokenCache byte[] via reflection.
    $cacheObj  = $session.InMemoryTokenCache
    $cacheType = $cacheObj.GetType()
    $field = $cacheType.GetField('_tokenCache',
        [System.Reflection.BindingFlags]::NonPublic -bor
        [System.Reflection.BindingFlags]::Instance)
    if(-not $field) {
        Write-LogDebug "Get-MgGraphCachedMsalAccounts: _tokenCache field not present on $($cacheType.FullName)"
        return @()
    }
    $cacheBytes = $field.GetValue($cacheObj)
    if(-not $cacheBytes -or $cacheBytes.Length -eq 0) { return @() }

    # Make sure the MSAL types are available — Add-MSALPrereq loads them, but on a
    # pure MgGraph-only setup they might not have been loaded yet. The Microsoft.Graph
    # SDK ships Microsoft.Identity.Client itself though, so usually fine.
    $msalBuilderType = "Microsoft.Identity.Client.PublicClientApplicationBuilder" -as [type]
    if(-not $msalBuilderType) {
        Write-LogDebug "Get-MgGraphCachedMsalAccounts: MSAL PublicClientApplicationBuilder type not available"
        return @()
    }

    # Use the same ClientId the SDK established so the deserialized cache contents match.
    $clientId = $null
    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if($ctx -and $ctx.ClientId) { $clientId = $ctx.ClientId }
    } catch { }
    if(-not $clientId) { $clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e" }

    $appBuilder = $msalBuilderType::Create($clientId)
    [void]$appBuilder.WithAuthority("https://login.microsoftonline.com/organizations/")
    $tempApp = $appBuilder.Build()

    # Deserialize the SDK's cache directly. Avoids SetBeforeAccess(scriptblock)
    # which would fire from MSAL's background thread (no PS Runspace there).
    $tempApp.UserTokenCache.DeserializeMsalV3($cacheBytes, $true)

    $accounts = $tempApp.GetAccountsAsync().GetAwaiter().GetResult()
    if(-not $accounts -or $accounts.Count -eq 0) { return @() }

    $rows = foreach($acc in $accounts) {
        [PSCustomObject]@{
            Provider = $ProviderId
            Username = $acc.Username
            UserId   = $acc.HomeAccountId.ObjectId
            TenantId = $acc.HomeAccountId.TenantId
            Native   = $acc
        }
    }
    return $rows
}

Invoke-MgGraphProviderInitialize
