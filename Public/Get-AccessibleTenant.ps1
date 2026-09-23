function Get-AccessibleTenant {
    <#
    .SYNOPSIS
        List the tenants the signed-in account can reach.

    .DESCRIPTION
        Returns one row per tenant the current account has access to - its home
        tenant plus every tenant it is a guest in - so you can confirm a tenant is
        reachable before connecting to it, or feed a tenant picker.

        Microsoft Graph cannot answer this question. Its tenant APIs resolve a
        single tenant you already know (findTenantInformationByTenantId) or list
        the members of a configured multi-tenant organization. The list of tenants
        an account can sign in to comes from Azure Resource Manager, which means a
        token for a second audience, so the Entra app registration must hold the
        delegated permission 'Azure Service Management / user_impersonation'. When
        it does not, this command writes a warning saying so and returns nothing.

        Only providers that can mint that second token implement this; today that
        is MSAL. With another provider active the command warns and returns
        nothing rather than failing.

        Switching to one of these tenants does not need a separate command: pass
        its id to Connect-IntuneManagement. With a cached account the acquire is
        silent, and the new tenant is registered as an additional token, so both
        stay live and -TokenId can address either.

        Exported as Get-IMAccessibleTenant (the module applies the 'IM' prefix).

    .PARAMETER TokenId
        Ask the provider that owns this token. Defaults to the current token.

    .EXAMPLE
        # Every tenant the signed-in account can reach
        Get-IMAccessibleTenant

    .EXAMPLE
        # Confirm a guest tenant is reachable, then connect to it silently
        $guest = Get-IMAccessibleTenant | Where-Object displayName -eq 'Fabrikam'
        Connect-IMIntuneManagement -Interactive -TenantId $guest.tenantId

    .EXAMPLE
        # Export from two tenants in one script
        Connect-IMIntuneManagement -Interactive
        $home = (Get-IMAuthToken)[0].TokenId
        Connect-IMIntuneManagement -Interactive -TenantId (Get-IMAccessibleTenant)[1].tenantId
        $guest = (Get-IMAuthToken | Sort-Object TokenId)[-1].TokenId
        Start-IMGraphBulkExport -ExportFolder 'C:\Export\Home'  -TokenId $home
        Start-IMGraphBulkExport -ExportFolder 'C:\Export\Guest' -TokenId $guest
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $false)]
        [int]$TokenId = 0
    )

    # Same owner-first routing as Invoke-MSGraphAPI, in the same order: resolve the
    # default id 0 to the token that owns the session FIRST, then find that token's
    # provider. The provider that minted the token is the one that can mint a second
    # one for the same account. Resolving 0 directly returns $null and would fall
    # back to whichever provider is active - not necessarily the default token's
    # owner, since a second login only re-points the default when asked to. An
    # unregistered id keeps the active-provider fallback.
    if ($TokenId -le 0) { $TokenId = Get-DefaultAuthTokenId }

    $authProvider = Resolve-AuthTokenProvider $TokenId
    if (-not $authProvider) { $authProvider = Get-AuthProvider }

    if (-not $authProvider) {
        Write-Warning "Not signed in. Run Connect-IMIntuneManagement first."
        return
    }

    $result = $authProvider.GetAccessibleTenants($TokenId)

    if ($null -eq $result) {
        Write-Warning "The '$($authProvider.Id)' provider cannot list tenants. Listing them needs an Azure Service Management token for the signed-in account, which only the MSAL provider acquires. Connect with -Provider MSAL, or pass a known tenant id straight to Connect-IMIntuneManagement."
        return
    }

    if ($result.ConsentMissing) {
        Write-Warning $result.Message
        return
    }

    if ($result.Message) { Write-Warning $result.Message }

    return @($result.Tenants)
}
