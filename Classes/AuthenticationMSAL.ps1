#ImportOrder 26

# MSAL implementation of AuthenticationProvider.
#
# Naming convention: every concrete auth provider is named Authentication<Backend>
# so they sort together in the Classes/ folder (AuthenticationMgGraph, AuthenticationOAuth).
#
# This class is a facade over the MSAL functions in Internal/AuthenticationMSALHelpers.ps1
# (Connect-EntraEnvironment / Connect-WithClientCredentials / Get-FullToken). Consumers
# reach MSAL through the provider abstraction: Invoke-MSGraphAPI resolves a call's owning
# provider by TokenId and asks it for the bearer via GetAccessToken.
class AuthenticationMSAL : AuthenticationProvider {

    # TokenIds currently inside a pre-flight silent refresh. Used by GetAccessToken
    # to break the re-entrancy cycle: Connect-EntraEnvironment itself calls back
    # into Graph (Organization / ME / photo) and those calls land here for the
    # bearer header. Without a guard the nested call sees the still-cached
    # expired token and recurses into Connect-EntraEnvironment forever — caught
    # in the wild as a "call depth overflow" crash. The first-in caller drives
    # the refresh; nested calls return the cached bearer (which Connect's inner
    # Invoke-MSGraphAPI -SkipAuthentication can already cope with).
    static [hashtable]$Refreshing = @{}

    AuthenticationMSAL() {
        $this.Id           = "MSAL"
        $this.DisplayName  = "Microsoft Authentication Library"

        # Required capabilities — all true (Interactive / ClientSecret / Certificate
        # default to $true on the base class). Set explicitly here as a contract
        # marker so a future edit can't accidentally turn one off.
        $this.SupportsInteractive  = $true
        $this.SupportsClientSecret = $true
        $this.SupportsCertificate  = $true

        # Optional capability: MSAL.NET supports federated credentials via
        # WithClientAssertion + managed identity via WithAzureMSI, but
        # Connect-IntuneManagement does NOT expose those paths yet. Flag stays
        # $false until the parameter sets are added.
        $this.SupportsIdentityProvider = $false

        # MSAL accepts BYO bearer tokens through Add-BYOTokenInfo.
        $this.SupportsBYOToken     = $true

        # MSAL re-mints tokens for CAE claims challenges (silent, escalating to
        # interactive when the caller allows it). See GetClaimsToken.
        $this.SupportsClaimsChallenge = $true

        # This provider's flows run through the built-in Connect-EntraEnvironment /
        # Connect-WithClientCredentials entry points (see UsesBuiltInConnectPath on the
        # base): Connect-IntuneManagement, the interactive-login helper, and the profile
        # Refresh action drive MSAL through those functions directly, keeping the rich
        # cloud/token-id handling and behaviour identical to the pre-abstraction path.
        $this.UsesBuiltInConnectPath  = $true

        # MSAL can launch an interactive consent prompt via Start-MSALConsentPrompt.
        $this.SupportsConsentPrompt   = $true
    }

    # No initialization work — MSAL DLLs and settings are wired by Invoke-MSALInitialize
    # which runs from AuthenticationMSALHelpers.ps1 at module load.
    [void] Initialize() { }

    [PSCustomObject] Connect([hashtable]$Arguments) {
        # Two entry points historically:
        #   Connect-IntuneManagement   : public API with explicit Secret / Certificate / Token
        #   Connect-EntraEnvironment   : internal — interactive, silent, refresh
        # Pick based on which fields are present in $Arguments.

        # Copy first so any Cloud→legacy translation we do here doesn't mutate the
        # caller's hashtable. We translate Cloud→GraphEnvironment+GCCType because the
        # downstream MSAL functions haven't migrated to the new enum yet (Phase 4).
        $local = @{}
        foreach($key in $Arguments.Keys) { $local[$key] = $Arguments[$key] }
        if($local.ContainsKey('Cloud') -and $local.Cloud -and -not $local.ContainsKey('GraphEnvironment')) {
            $entry = Get-CloudByValue ([string]$local.Cloud)
            if($entry) {
                $local['GraphEnvironment'] = $entry.LegacyEnv
                if($entry.LegacyGCC) { $local['GCCType'] = $entry.LegacyGCC }
            }
        }

        if($local.ContainsKey('Secret') -or $local.ContainsKey('Certificate') -or
           $local.ContainsKey('CertificatePath') -or $local.ContainsKey('Token')) {
            # Connect-IntuneManagement now understands -Cloud directly; we still pass
            # the legacy params for compatibility (Connect-IntuneManagement re-resolves
            # them with -Cloud taking precedence).
            return (Connect-IntuneManagement @local)
        }

        # Connect-EntraEnvironment uses the internal -Environment parameter (the MSAL
        # function predates our public taxonomy). Translate before splatting so the
        # cloud picker dialog's choice actually reaches MSAL. GCCType is NOT a
        # Connect-EntraEnvironment parameter (only Connect-WithClientCredentials
        # / Add-BYOTokenInfo take it) — splatting it triggers "Cannot bind
        # positional parameters"; drop it. Cloud IS a parameter on
        # Connect-EntraEnvironment now, so we no longer need to drop it either.
        if($local.ContainsKey('GraphEnvironment')) {
            $local['Environment'] = $local['GraphEnvironment']
            $local.Remove('GraphEnvironment') | Out-Null
        }
        if($local.ContainsKey('GCCType')) { $local.Remove('GCCType') | Out-Null }
        return (Connect-EntraEnvironment @local)
    }

    [bool] Disconnect([int]$TokenId) {
        try {
            Disconnect-EntraEnvironment -TokenID $TokenId
            return $true
        }
        catch {
            Write-LogError "MSAL Disconnect failed for TokenId $TokenId" $_.Exception
            return $false
        }
    }

    [bool] Refresh([int]$TokenId) {
        try {
            return [bool](Connect-EntraEnvironment -TokenId $TokenId -ForceRefresh)
        }
        catch {
            Write-LogError "MSAL Refresh failed for TokenId $TokenId" $_.Exception
            return $false
        }
    }

    # MSAL is the one provider that can do this: the same account can mint a second
    # token for the Azure Resource Manager audience, which is the only API that
    # enumerates a user's tenants (see Internal/EntraTenantList.ps1 for why Graph
    # cannot). Returns the result object that file documents, or $null when there
    # is no usable MSAL session to ask with.
    [PSCustomObject] GetAccessibleTenants([int]$TokenId) {
        $tokenInfo = Get-FullToken $TokenId
        if(-not $tokenInfo -or -not $tokenInfo.App -or -not $tokenInfo.Token -or -not $tokenInfo.Token.Account) {
            Write-LogDebug "MSAL GetAccessibleTenants: no usable token for id $TokenId"
            return $null
        }

        # Only offer the no-prompt interactive fallback once the UI is up; in a
        # script there is nobody to answer a window that may appear.
        $interactive = ($script:MainAppStarted -eq $true)

        return (Get-EntraAccessibleTenant -App $tokenInfo.App -Account $tokenInfo.Token.Account `
                    -TenantId $tokenInfo.Token.TenantId -Cloud $tokenInfo.Cloud -AllowInteractive:$interactive)
    }

    # Silent startup resume — re-establish the last session from the persisted MSAL
    # cache without prompting. Invoked once from Invoke-AuthCoreOnAppInitialized for
    # the active provider; restores the old Connect-MSALUser -Silent startup logon
    # that made the app auto-sign-in on launch. The fresh (no -TokenId) path resolves
    # the account from the on-disk cache via GetAccountsAsync matched against the
    # persisted LastLoggedOnUserId. -ForceSilent guarantees no interactive prompt: if
    # there is no cached account (or the broker can't silently reissue), it simply
    # returns $false and the user signs in manually. -DefaultToken makes the resumed
    # session the active default so the UI shows signed-in.
    [bool] TryResumeSession() {
        # Respect the "Remember Login" toggle — if the user disabled caching there is
        # nothing to resume and we should not touch the account cache.
        if(-not (Get-SettingValue "CacheMSALToken")) { return $false }
        # Cheap probe (settings + file existence only) BEFORE the MSAL runtime loads:
        # a fresh box with no cached session skips the DLL load entirely.
        if(-not (Test-MSALResumeLikely)) {
            Write-LogDebug "MSAL TryResumeSession skipped - no cached session to resume"
            return $false
        }
        try {
            $result = Connect-EntraEnvironment -ForceSilent -DefaultToken
            return [bool]$result
        }
        catch {
            Write-LogError "MSAL TryResumeSession failed" $_.Exception
            return $false
        }
    }

    # Silent ambient refresh on view activation. Gated on the base no-op for other
    # providers so MgGraph mode isn't hijacked (a successful silent MSAL auth here would
    # flip the active provider). No -DefaultToken: this only refreshes, never promotes.
    [void] RefreshAmbientSession() {
        Connect-EntraEnvironment -ForceSilent | Out-Null
    }

    # Native session inspector rows for the profile "Session Info" dialog: the MSAL
    # AuthenticationResult fields (minus the raw tokens).
    # Decoded id-token JWT for the profile popup's Id Token inspector. Reads MSAL's
    # own token entry from the registry; returns $null when there's no id token so the
    # UI hides the button. This keeps the id-token JWT confined to the MSAL provider.
    [object] GetIdTokenJwt([int]$TokenId) {
        $t = Get-FullToken $TokenId
        if($t -and $t.JWTIdToken) { return $t.JWTIdToken }
        return $null
    }

    [PSCustomObject[]] GetSessionInfoRows() {
        $rows = @()
        if($script:MSALDefaultToken -and $script:MSALDefaultToken.Token) {
            foreach($prop in ($script:MSALDefaultToken.Token | Get-Member | Where-Object MemberType -eq Property)) {
                if($prop.Name -in @("AccessToken", "IdToken")) { continue }
                $value = if($prop.Name -eq "Scopes") { ($script:MSALDefaultToken.Token.Scopes -join "`n") }
                         elseif($prop.Name -in @("ExpiresOn", "ExtendedExpiresOn")) { $script:MSALDefaultToken.Token."$($prop.Name)".LocalDateTime }
                         else { $script:MSALDefaultToken.Token."$($prop.Name)" }
                $rows += [PSCustomObject]@{ Name = $prop.Name; Value = $value }
            }
        }
        return [PSCustomObject[]]$rows
    }

    [bool] ForgetAccount([string]$AccountIdentifier) {
        if(-not $script:MSALAccounts) { return $false }
        $account = $script:MSALAccounts | Where-Object {
            $_.Username -eq $AccountIdentifier -or
            $_.HomeAccountId.Identifier -eq $AccountIdentifier
        } | Select-Object -First 1
        if(-not $account) { return $false }
        Remove-MSALAccount -Account $account
        return $true
    }

    [string] GetAccessToken([int]$TokenId, [string]$Resource) {
        $tok = Get-FullToken $TokenId
        if(-not $tok -or -not $tok.Token) { return $null }

        # Pre-flight refresh near expiry to avoid mid-batch 401s. Today the resource is
        # always Graph (the MSAL token caches a single audience); when other resources
        # are supported in a future phase, the provider should mint per-resource tokens
        # here via AcquireTokenSilent.WithScopes(resource/.default).
        #
        # Re-entrancy guard: Connect-EntraEnvironment internally calls Invoke-MSGraphAPI
        # ('Organization', 'ME', photo) before its own returns; those calls land back
        # in this method for the bearer header. The cached token is still the expired
        # one at that point — the new token isn't installed until Connect's
        # Add-MSALTokenInfo runs at the tail. Without a guard the nested call retries
        # the refresh, which calls Invoke-MSGraphAPI, which re-enters here, etc., until
        # PowerShell's call-depth limit aborts.
        if($tok.Token.ExpiresOn -lt [DateTimeOffset]::UtcNow.AddMinutes(5) -and
           -not [AuthenticationMSAL]::Refreshing.ContainsKey($TokenId)) {
            [AuthenticationMSAL]::Refreshing[$TokenId] = $true
            # Only let a refresh miss raise AuthenticationFailed once the token has
            # ACTUALLY expired. Within the 5-minute pre-flight window the current token
            # is still usable, so a silent-refresh miss (e.g. the WAM broker failing on
            # the refresh round-trip) must stay quiet - otherwise every near-expiry Graph
            # call falsely reports a failed login even though the call then succeeds on
            # the still-valid token. When truly expired, stay loud so the UI signs out.
            $tokenStillValid = $tok.Token.ExpiresOn -gt [DateTimeOffset]::UtcNow
            try {
                # Plain silent acquire (NO -ForceRefresh). AcquireTokenSilent already
                # renews an expired/near-expired access token from the refresh token
                # (or via the WAM broker) on its own. Forcing a refresh here made the
                # broker re-contact Entra ~5 min before every expiry, and WAM can surface
                # an interactive window on that forced round-trip - that was the ~70-min
                # re-login. The old 3.9.6 build never force-refreshed routinely (only on
                # an explicit user "Force refresh" link); this matches it. -ForceRefresh
                # is still used by Refresh() and the UI Refresh button where it is wanted.
                [void](Connect-EntraEnvironment -TokenId $TokenId -ForceSilent -SuppressFailedEvent:$tokenStillValid)
            }
            finally {
                [AuthenticationMSAL]::Refreshing.Remove($TokenId) | Out-Null
            }
            $tok = Get-FullToken $TokenId
            if(-not $tok -or -not $tok.Token) { return $null }
        }
        return $tok.Token.AccessToken
    }

    # Satisfy a CAE claims challenge. Silent re-acquire first (broker/WAM can often
    # satisfy a CAE / sign-in-frequency challenge without a visible prompt); if that
    # fails and the caller allows interaction, escalate to an interactive acquire with
    # the same claims so the challenge is met with a single prompt instead of a dead
    # 401. When $AllowInteractive is $false (headless / nested auth-flow call) this
    # stays silent-only and returns $null if the challenge can't be met.
    [string] GetClaimsToken([int]$TokenId, [string]$Resource, [string]$ClaimsChallenge, [bool]$AllowInteractive) {
        [void](Connect-EntraEnvironment -TokenId $TokenId -ClaimsChallenge $ClaimsChallenge -ForceSilent)
        $token = $this.GetAccessToken($TokenId, $Resource)

        if(-not $token -and $AllowInteractive) {
            Write-Log "CAE challenge could not be satisfied silently. Escalating to interactive login." 2
            [void](Connect-EntraEnvironment -TokenId $TokenId -ClaimsChallenge $ClaimsChallenge -ForceInteractive)
            $token = $this.GetAccessToken($TokenId, $Resource)
        }
        return $token
    }

    [datetime] GetAccessTokenExpiry([int]$TokenId, [string]$Resource) {
        $tok = Get-FullToken $TokenId
        if(-not $tok -or -not $tok.Token -or -not $tok.Token.ExpiresOn) {
            return [datetime]::MaxValue
        }
        return $tok.Token.ExpiresOn.LocalDateTime
    }

    [PSCustomObject] GetUserInfo([int]$TokenId) {
        $tok = Get-FullToken $TokenId
        if(-not $tok) { return $null }

        $authType = if($tok.AuthType) { $tok.AuthType } else { "Interactive" }
        $expires  = $null
        if($tok.Token -and $tok.Token.ExpiresOn) { $expires = $tok.Token.ExpiresOn.LocalDateTime }

        $upn = $null
        if($tok.Token -and $tok.Token.Account) { $upn = $tok.Token.Account.Username }

        $userId = $null
        if($tok.Token -and $tok.Token.Account -and $tok.Token.Account.HomeAccountId) {
            $userId = $tok.Token.Account.HomeAccountId.ObjectId
        }

        return [PSCustomObject]@{
            Provider    = $this.Id
            DisplayName = $upn
            UPN         = $upn
            UserId      = $userId
            TenantId    = (?: $tok.Token $tok.Token.TenantId $null)
            TenantName  = (?: $tok.Organization $tok.Organization.displayName $null)
            AppId       = (?: $tok.EntraApp $tok.EntraApp.ClientId $null)
            AppName     = (?: $tok.EntraApp $tok.EntraApp.Name $null)
            AuthType    = $authType
            ExpiresOn   = $expires
        }
    }

    [PSCustomObject[]] GetCachedAccounts() {
        # Lazy-refresh from the on-disk MSAL cache. Same trick the UI already does in
        # Get-MSALUserProfile, but exposed at the provider level so any consumer
        # (CLI scripts, automation) sees the same list.
        if(($script:MSALAccounts | Measure-Object).Count -eq 0) {
            try {
                $app = $script:MSALApps | Select-Object -First 1
                if(-not $app) { $app = New-MSALApp }
                if($app) {
                    $script:MSALAccounts = $app.GetAccountsAsync().GetAwaiter().GetResult()
                }
            }
            catch {
                Write-LogDebug "MSAL GetCachedAccounts refresh failed: $($_.Exception.Message)"
            }
        }
        if(-not $script:MSALAccounts) { return [PSCustomObject[]]@() }

        $rows = foreach($acc in $script:MSALAccounts) {
            [PSCustomObject]@{
                Provider = $this.Id
                Username = $acc.Username
                UserId   = $acc.HomeAccountId.ObjectId
                TenantId = $acc.HomeAccountId.TenantId
                Native   = $acc
            }
        }
        return [PSCustomObject[]]@($rows)
    }

    [PSCustomObject[]] GetAvailableTenants([int]$TokenId) {
        $tok = Get-FullToken $TokenId
        if(-not $tok -or -not $tok.Tenants) { return [PSCustomObject[]]@() }

        $rows = foreach($t in $tok.Tenants) {
            [PSCustomObject]@{
                Provider   = $this.Id
                TenantId   = $t.tenantId
                TenantName = $t.displayName
                Native     = $t
            }
        }
        return [PSCustomObject[]]@($rows)
    }
}
