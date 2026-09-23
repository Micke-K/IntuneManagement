function Get-AuthToken {
    <#
    .SYNOPSIS
        List the authentication tokens currently held, across every provider.

    .DESCRIPTION
        Returns one [IMAuthToken] per live token from the central token registry,
        regardless of which provider (MSAL / OAuth / MgGraph) acquired it. Each
        token is tagged with the environment it belongs to (Provider, TenantId,
        TenantName, Cloud) plus the account/app identity and expiry, and a global
        TokenId that uniquely identifies it.

        Use the TokenId with -TokenId on other commands (e.g. Invoke-MSGraphAPI,
        Copy-GraphPolicy) to route a call to that specific environment. This lets
        you sign into several environments at once - even across different
        providers - and copy policies between them.

        Exported as Get-IMAuthToken (the module applies the 'IM' command prefix).

    .PARAMETER TokenId
        Return only the token with this global id.

    .PARAMETER Provider
        Return only tokens owned by this provider (MSAL / OAuth / MgGraph).

    .PARAMETER TenantId
        Return only tokens for this tenant id.

    .EXAMPLE
        # Every environment you're signed into
        Get-IMAuthToken

    .EXAMPLE
        # Pick the production tenant's token and document only that environment
        $prod = Get-IMAuthToken | Where-Object TenantName -eq 'Contoso Prod'
        Get-IMGraphPolicies -TokenId $prod.TokenId
    #>
    [CmdletBinding()]
    [OutputType([IMAuthToken[]])]
    param(
        [int]$TokenId,
        # Completion instead of [ValidateSet([AuthProviderValues])] for PS5.1 import
        # compatibility (see Connect-IntuneManagement); unknown values just filter to none.
        [ArgumentCompleter({ param($commandName, $parameterName, $wordToComplete) @(& (Get-Module IntuneManagement) { Get-AuthProviderValues }) | Where-Object { $_ -like "$wordToComplete*" } })]
        [string]$Provider,
        [string]$TenantId
    )

    $tokens = Get-AuthTokenList

    if ($PSBoundParameters.ContainsKey('TokenId')) {
        $tokens = @($tokens | Where-Object { $_.TokenId -eq $TokenId })
    }
    if ($Provider) {
        $tokens = @($tokens | Where-Object { $_.Provider -eq $Provider })
    }
    if ($TenantId) {
        $tokens = @($tokens | Where-Object { $_.TenantId -eq $TenantId })
    }

    return [IMAuthToken[]]@($tokens)
}
