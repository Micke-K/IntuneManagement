# Avalonia authentication actions that must start after their initiating control
# event returns. Authentication itself remains synchronous and on the UI thread;
# only its start is deferred by one dispatcher turn so Avalonia can release the
# initiating button's pointer capture before a cancellable wait begins.

$script:AvaloniaDeferredActions = [System.Collections.Generic.Queue[object]]::new()

# Schedule module code for the next dispatcher turn. Arguments are stored
# explicitly because NewBoundScriptBlock intentionally does not retain
# function-local captures.
function Invoke-AvaloniaDeferredAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        $Argument
    )

    $boundAction = ConvertTo-AvaloniaEventScriptBlock $Action
    $runner = (ConvertTo-AvaloniaEventScriptBlock {
        Invoke-NextAvaloniaDeferredAction
    }) -as [Action]
    if(-not $runner) { throw 'Could not create deferred Avalonia dispatcher action.' }

    $script:AvaloniaDeferredActions.Enqueue([PSCustomObject]@{
        Action   = $boundAction
        Argument = $Argument
    })
    [Avalonia.Threading.Dispatcher]::UIThread.Post($runner)
}

function Invoke-NextAvaloniaDeferredAction {
    if(-not $script:AvaloniaDeferredActions -or $script:AvaloniaDeferredActions.Count -eq 0) { return }

    $work = $script:AvaloniaDeferredActions.Dequeue()
    try { & $work.Action $work.Argument }
    catch { Write-LogError 'Deferred Avalonia action failed' $_.Exception }
}

function Invoke-AvaloniaInteractiveSignIn {
    param([string]$Cloud)

    if($Cloud) { Write-Status "Signing in to $Cloud cloud..." }
    else { Write-Status 'Signing in...' }
    try {
        $tokenInfo = if($Cloud) {
            Invoke-AuthProviderInteractiveLogin -Cloud $Cloud
        } else {
            Invoke-AuthProviderInteractiveLogin
        }
        # Successful providers register the token and synchronously fire
        # AuthenticatedNewToken. Avalonia's handler owns profile enrichment and the
        # auth-info redraw; repeating Get-MSALUserInfo here would issue /me and photo
        # requests twice for every MSAL sign-in.
        if (-not $tokenInfo) {
            Write-Log 'Sign-in returned nothing (user cancelled or auth failed)' 2
        }
    }
    catch {
        Write-LogError 'Deferred sign-in failed' $_.Exception
    }
    finally {
        Write-Status ''
    }
}

function Invoke-AvaloniaConsentPrompt {
    try {
        if (Get-Command Start-MSALConsentPrompt -ErrorAction SilentlyContinue) {
            Start-MSALConsentPrompt
        }
    }
    catch {
        Write-LogError 'Start-MSALConsentPrompt failed' $_.Exception
    }
}

function Invoke-AvaloniaProfileRefresh {
    Write-Status 'Refreshing the token'
    try {
        $providerNow = $null
        try { $providerNow = Get-AuthProvider } catch { }

        $refreshed = $false
        if ($providerNow -and $providerNow.UsesBuiltInConnectPath) {
            $refreshed = [bool](Connect-EntraEnvironment -ForceRefresh -TokenId (Get-DefaultAuthTokenId))
        } elseif ($providerNow) {
            # The real default token id, not a literal 0 - see the same call in
            # UI/WPF/Extensions/MSGraphAuthenticationUIWPF.ps1: ids start at 1, so 0
            # matched nothing in OAuth's token table and refresh was a silent no-op.
            $refreshed = [bool]$providerNow.Refresh((Get-DefaultAuthTokenId))
        }

        if (-not $refreshed) {
            $script:UIProvider.HidePopup()
        } else {
            try { Get-MSALUserInfo } catch { Write-LogError 'Get-MSALUserInfo failed after refresh' $_.Exception }
            if (Get-Command Show-ViewMenu -ErrorAction SilentlyContinue) { Show-ViewMenu }
        }
    }
    catch {
        Write-LogError 'Token refresh failed' $_.Exception
    }
    finally {
        Write-Status ''
    }
}

function Invoke-AvaloniaTenantSwitch {
    param($Tenant)
    if(-not $Tenant) { return }

    Write-Status "Logging in to $($Tenant.DisplayName)"
    try {
        $userName = $script:CurrentUser.userPrincipalName
        if ($userName) {
            Connect-EntraEnvironment -User $userName -TenantId $Tenant.tenantId -DefaultToken | Out-Null
        } else {
            Connect-EntraEnvironment -TenantId $Tenant.tenantId -DefaultToken | Out-Null
        }
    }
    catch {
        Write-LogError 'Tenant switch failed' $_.Exception
    }
    finally {
        Write-Status ''
    }
}
