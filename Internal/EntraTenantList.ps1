# Enumerating the tenants a signed-in account can reach.
#
# Microsoft Graph still has no API for this. tenantRelationships
# findTenantInformationByTenantId / ByDomainName resolve ONE tenant you already
# know the id or domain of, and the multi-tenant-organization APIs only list the
# members of a configured MTO. Neither answers "which tenants can I sign in to".
# Azure Resource Manager's /tenants does, and remains the endpoint the portal's
# own directory switcher uses.
#
# The catch is the audience: ARM tokens are not Graph tokens, so the account has
# to mint a second one for the ARM resource. That is why only MSAL implements
# AuthenticationProvider.GetAccessibleTenants today - the MgGraph SDK hands out
# Graph-scoped tokens only. It is also why this fails on app registrations that
# were never granted the delegated Azure Service Management permission, which is
# the case this file reports clearly instead of swallowing (Get-TenantList's
# original catch{} left the list silently empty).

# Delegated permission on the "Windows Azure Service Management API" service
# principal. The resource it is requested on is per-cloud, so the scope is built at
# call time from Get-EntraArmAudience.
$script:EntraArmScopeSuffix = "/user_impersonation"

# The ARM host the tenant list is READ from. Kept beside its only consumer; the flat
# cloud table in AuthenticationCore.ps1 owns the value itself (ArmHost).
function Get-EntraArmHost {
    param([string]$Cloud)

    $entry = Get-CloudByValue $Cloud
    if($entry -and $entry.ArmHost) { return $entry.ArmHost }
    return "management.azure.com"
}

# The OAuth resource the ARM token is REQUESTED on - a separate value from the host
# above, even though the two are the same string in the public cloud. Reusing the
# REST endpoint as the audience is what made sovereign clouds wrong: USGov and China
# have their own Azure Service Management identifiers, and Entra answers a request
# for a resource its tenant does not know with AADSTS500011, not with anything that
# points at the resource string.
#
# An ordered list, because which identifier a given tenant accepts is not something
# this module can decide from here: current Azure SDK cloud metadata uses the ARM
# endpoint itself as the audience, while Az and the CLI still carry the older
# management.core.* ASM identifier, and sovereign tenants differ in which is
# provisioned. First entry is tried first, and Get-EntraAccessibleTenant only moves
# on when the failure says the resource itself was rejected.
function Get-EntraArmAudience {
    param([string]$Cloud)

    $entry = Get-CloudByValue $Cloud
    if($entry -and $entry.ArmAudiences) { return @($entry.ArmAudiences) }
    return @("https://$(Get-EntraArmHost $Cloud)")
}

# Classify an acquire failure. "The app was never granted Azure Service Management"
# is one of several things that can go wrong here, and reporting the others as it
# sends the user after the wrong thing: a prompt that timed out, a refresh token that
# was revoked, conditional access or an MFA requirement are not fixed by granting a
# permission. Get-MsalAuthenticationToken reports a timeout as a plain string, a user
# cancel as OperationCanceledException and anything else as the MSAL exception, so
# all three shapes arrive here.
#
#   None        - nothing captured; the silent acquire simply found no cached token
#   Resource    - Entra did not accept the requested resource identifier, so another
#                 audience is worth trying
#   Consent     - the app or the user has not consented to Azure Service Management
#   Interaction - UI required, but NOT specifically consent: an expired or revoked
#                 grant, conditional access, MFA, login_required, an expired
#                 password. MSAL reports all of these as MsalUiRequiredException and
#                 Entra as invalid_grant, so neither of those on its own means consent
#   Cancelled   - the user closed the prompt
#   TimedOut    - the prompt was never answered
#   Other       - anything else; its own message is all the caller can act on
function Get-EntraArmFailureKind {
    param($Failure)

    if(-not $Failure) { return "None" }
    if($Failure -is [string]) { return "TimedOut" }
    if($Failure -is [System.OperationCanceledException]) { return "Cancelled" }

    $text = "$($Failure.ErrorCode) $($Failure.Message)"

    # Rule out a rejected resource identifier before anything else: AADSTS500011
    # ("resource principal not found in the tenant") is exactly what a tenant says
    # about an audience it does not know, which includes an audience this module
    # picked wrongly, and AADSTS650057 names an invalid resource outright.
    if($text -match 'AADSTS500011|AADSTS650057|invalid_resource') { return "Resource" }

    # AADSTS65001: no consent recorded for the app. AADSTS65004: the user declined it.
    if($text -match 'AADSTS65001|AADSTS65004') { return "Consent" }

    # Type names, not a type literal: MSAL may not be loaded in the caller's session,
    # and PSObject.TypeNames carries the whole inheritance chain of a real exception.
    if(@($Failure.PSObject.TypeNames) -like '*MsalUiRequiredException') {
        # MSAL's own verdict when it has one. Without it, this exception says no more
        # than "a prompt is needed", which is true of every interaction state.
        if("$($Failure.Classification)" -eq 'ConsentRequired') { return "Consent" }
        return "Interaction"
    }

    return "Other"
}

# The no-prompt interactive round, kept in its own function so the decision about
# which failure to report is testable without MSAL: this is the only part of the
# acquire that needs the assembly loaded. It succeeds when the existing session
# already carries the consent, and fails fast otherwise.
#
# "No prompt" is about what Entra will ASK, not about what appears on screen, and
# the difference matters to anyone reusing this: Prompt.NoPrompt sends prompt=none,
# so the STS answers login_required / interaction_required instead of asking for
# credentials or consent - but this is still AcquireTokenInteractive, so MSAL starts
# the interactive flow and the system browser or the broker can open a window (and
# on some platforms flash one) before that answer arrives. It is therefore only
# safe where a user is present, which is what the caller's -AllowInteractive gate
# means: the MSAL provider sets it only once the UI is up.
#
# Returns $authResult, $failure like Get-MsalAuthenticationToken.
function Get-EntraArmInteractiveToken {
    param(
        [Parameter(Mandatory = $true)]$App,
        [Parameter(Mandatory = $true)]$Account,
        [Parameter(Mandatory = $true)][string[]]$Scope,
        [string]$TenantId
    )

    $interactive = $App.AcquireTokenInteractive($Scope)
    [void]$interactive.WithLoginHint($Account.Username)
    # Suppresses the credential and consent prompts, not the browser or broker itself.
    [void]$interactive.WithPrompt([Microsoft.Identity.Client.Prompt]::NoPrompt)
    if($TenantId) { [void]$interactive.WithTenantId($TenantId) }

    return @(Get-MsalAuthenticationToken $interactive -Interactive)
}

# One acquire round for one audience: silent first and, only with a user present, a
# no-prompt interactive round. Returns $authResult, $failure - the same pair shape as
# Get-MsalAuthenticationToken.
function Get-EntraArmToken {
    param(
        [Parameter(Mandatory = $true)]$App,
        [Parameter(Mandatory = $true)]$Account,
        [Parameter(Mandatory = $true)][string[]]$Scope,
        [string]$TenantId,
        [switch]$AllowInteractive
    )

    $authResult = $null
    $failure    = $null

    try {
        $silentBuilder = $App.AcquireTokenSilent($Scope, $Account)
        if($TenantId) { [void]$silentBuilder.WithTenantId($TenantId) }
        $authResult, $failure = Get-MsalAuthenticationToken $silentBuilder

        # A bare "UI required" is the one failure a no-prompt round can still resolve:
        # the cache has nothing, but the browser session may. Consent, a rejected
        # resource, a cancel or a timeout are all answered already - prompting again
        # only risks a window nobody asked for.
        if(-not $authResult -and $AllowInteractive -and (Get-EntraArmFailureKind $failure) -eq "Interaction") {
            $authResult, $interactiveFailure = Get-EntraArmInteractiveToken -App $App -Account $Account `
                                                    -Scope $Scope -TenantId $TenantId

            # The interactive reason is the newer and, in every kind but one, the more
            # specific one: a cancel or timeout is the user's decision, a named consent
            # or resource error is the answer this round was for, and a proxy, broker or
            # service error is the thing that actually stopped it - reporting the silent
            # "a prompt is needed" over any of those sends the user off to sign in again
            # for a problem signing in cannot fix.
            #
            # The exception is another UI-required state: a no-prompt round says that
            # whenever the browser session cannot satisfy it, which is no more than the
            # silent failure already said, so the silent reason stays.
            $interactiveKind = Get-EntraArmFailureKind $interactiveFailure
            if(-not $authResult -and $interactiveKind -notin @("None", "Interaction")) {
                $failure = $interactiveFailure
            }
        }
    }
    catch {
        $failure = $_.Exception
        Write-LogDebug "Tenant list: ARM token acquire threw - $($_.Exception.Message)"
    }

    return @($authResult, $failure)
}

# Ask Azure Resource Manager which tenants this account can reach.
#
# Returns a result object rather than a bare list so callers can tell the three
# outcomes apart:
#   Tenants        - the rows ARM returned (possibly empty)
#   ConsentMissing - the app registration lacks the Azure Service Management
#                    permission (or ARM refused the token with 401/403). Only set
#                    for consent-shaped failures: a cancelled prompt, an expired
#                    grant, conditional access or a network error leave it false and
#                    report their own reason in Message.
#   Message        - one sentence describing what to do about it
#   ArmHost        - the host the list was read from
#   ArmAudience    - the resource identifier the token was issued for
function Get-EntraAccessibleTenant {
    [CmdletBinding()]
    param(
        # MSAL application object able to acquire for this account.
        [Parameter(Mandatory = $true)]$App,
        # MSAL account (TokenInfo.Token.Account).
        [Parameter(Mandatory = $true)]$Account,
        [string]$TenantId,
        [string]$Cloud,
        # Allow a no-prompt interactive round when the silent acquire needs one. Off for
        # scripted callers: prompt=none stops Entra asking for credentials or consent,
        # but the round is still an interactive acquire, so a browser or broker window
        # can appear and there would be nobody to deal with it.
        [switch]$AllowInteractive
    )

    $armHost   = Get-EntraArmHost $Cloud
    $audiences = Get-EntraArmAudience $Cloud
    $result = [PSCustomObject]@{
        Tenants        = @()
        ConsentMissing = $false
        Message        = $null
        ArmHost        = $armHost
        ArmAudience    = $audiences[0]
    }

    $authResult            = $null
    $authenticationFailure = $null
    $failureKind           = "None"

    foreach($audience in $audiences) {
        $scope = [string[]]("$audience$($script:EntraArmScopeSuffix)")
        $authResult, $authenticationFailure = Get-EntraArmToken -App $App -Account $Account `
                                                -Scope $scope -TenantId $TenantId -AllowInteractive:$AllowInteractive
        if($authResult) {
            $result.ArmAudience = $audience
            break
        }

        $failureKind = Get-EntraArmFailureKind $authenticationFailure

        # A rejected resource identifier is the only failure another audience can fix.
        # Every other kind would fail identically on the next one, and each extra round
        # is another chance of an unwanted prompt.
        if($failureKind -ne "Resource") { break }
        Write-Log "Tenant list: $audience was not accepted as an Azure Service Management resource for this tenant - trying the next identifier" 2
    }

    if(-not $authResult) {
        $reason = if($authenticationFailure -is [string]) { $authenticationFailure }
                  elseif($authenticationFailure) { $authenticationFailure.Message }
                  else { "no reason was reported" }

        if($failureKind -eq "None" -or $failureKind -eq "Consent") {
            # Nothing cached for this resource, or Entra naming consent outright: the
            # never-granted app registration, which is the historical verdict here.
            $result.ConsentMissing = $true
            $result.Message = "The application could not get an Azure Service Management token for this account, so the tenant list is unavailable. Grant the Entra app registration the delegated permission 'Azure Service Management / user_impersonation' and consent to it, then sign in again."
            Write-Log "Tenant list: no ARM token for $armHost - the app is most likely missing the Azure Service Management delegated permission" 2
        }
        elseif($failureKind -eq "Resource") {
            # Every known identifier for this cloud was refused. Either the Azure
            # Service Management principal is not in the tenant, or none of the
            # identifiers this module knows is the one it wants - say both rather than
            # picking one.
            $result.Message = "Entra did not accept $($audiences -join ' or ') as an Azure Service Management resource for this tenant, so the tenant list is unavailable. Either the 'Windows Azure Service Management API' service principal does not exist in the tenant, or this cloud uses a different resource identifier - $reason"
            Write-Log "Tenant list: no audience accepted for $armHost - $reason" 2
        }
        elseif($failureKind -eq "Interaction") {
            # UI required for a reason that is not consent. Naming a permission grant
            # here would be a guess, and the wrong one whenever a sign-in is what is
            # actually needed.
            $result.Message = "Could not get an Azure Service Management token for $armHost without prompting, so the tenant list is unavailable. The account may need to sign in again - an expired or revoked grant, conditional access and an MFA requirement all end here - $reason"
            Write-Log "Tenant list: ARM token needs interaction for $armHost - $reason" 2
        }
        else {
            # Cancelled, timed out, or something else entirely. Its own message is the
            # only thing the caller can act on.
            $result.Message = "Could not get an Azure Service Management token for $armHost, so the tenant list is unavailable - $reason"
            Write-Log "Tenant list: ARM token acquire failed for $armHost - $reason" 2
        }
        return $result
    }

    $params = @{}
    $proxyURI = Get-ProxyURI
    if($proxyURI) {
        $params.Add("proxy", $proxyURI)
        $params.Add("UseBasicParsing", $true)
    }

    try {
        $headers = @{
            'Content-Type'  = 'application/json'
            'Authorization' = "Bearer " + $authResult.AccessToken
        }
        $response = Invoke-RestMethod "https://$armHost/tenants?api-version=2020-01-01" -Headers $headers @params
        if($response) { $result.Tenants = @($response.Value) }
        Write-Log "Tenant list: $($result.Tenants.Count) tenant(s) returned by $armHost"
    }
    catch {
        $status = $null
        try { $status = [int]$_.Exception.Response.StatusCode } catch { }
        if($status -in @(401, 403)) {
            $result.ConsentMissing = $true
            $result.Message = "Azure Resource Manager rejected the token with HTTP $status. The Entra app registration needs the delegated permission 'Azure Service Management / user_impersonation', with consent granted for this account."
        }
        else {
            $result.Message = "Could not read the tenant list from $armHost - $($_.Exception.Message)"
        }
        Write-LogError "Tenant list: request to $armHost failed" $_.Exception
    }

    return $result
}
