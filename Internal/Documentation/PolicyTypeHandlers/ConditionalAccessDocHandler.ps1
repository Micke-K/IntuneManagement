# Conditional Access policy documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:2476 (Invoke-CDDocument-
# ConditionalAccess, the largest of the ~13 custom handlers at ~944 LOC).
# Phase-4 work: this is the FIRST handler ported, validated against a golden
# fixture under C:/Intune/OldDocumentation/ConditionalAccess/.
#
# 2026-08-16: reconciled against the decompiled portal blade bundles (PolicyBlade
# converters + selectors + PoliciesResources). Language keys prefixed with "###"
# do NOT exist in Config/LanguageStrings yet (pipeline refresh needed); each such
# call carries an English fallback via Get-LanguageString -IgnoreMissing.
#
# Status — covered:
#   BasicInfo:   Name, Profile type, Enable policy state (incl. staged rollout +
#                target regions)
#   Assignments: users / groups / roles (incl. custom roles), guests + external
#                tenants (names resolved), workload identities (all radio tokens +
#                SP filter), agent identities + agent users (incl. legacy
#                AllAgentIdUsers token)
#   Target:      cloud apps (Office365 / admin portals / agentic resources / GSA
#                well-known apps), excluded apps, application filter, user actions
#                (incl. account recovery), authentication context (names resolved,
#                urn:microsoft:req1-3), GSA traffic profiles
#   Conditions:  user / sign-in / insider (flags string) / service-principal /
#                agent-id / agent-session risk, sign-in risk detections, agent
#                context, device platforms, locations (named + trusted), client
#                app types (incl. easSupported), time (preview), device states +
#                device filter, authentication flows (transfer methods + protocol
#                flows), Purview DLP
#   Grant:       block / mfa / authentication strength / compliant device /
#                domain-joined / approved app / app protection policy / password
#                change / risk remediation / identity verification, terms of use,
#                custom controls, AND/OR operator
#   Session:     app-enforced restrictions, Conditional Access App Control (4
#                types), sign-in frequency (time-based / every time / strict +
#                secondary-auth sub-option), persistent browser, continuous access
#                evaluation (3 modes), disable resilience defaults, token
#                protection (sign-in + app sessions), block sensitive actions,
#                Global Secure Access security profile (name resolved)
#
# Still missing:
#   - per-detection labels for signInRiskDetections (the checkbox labels live in
#     the lazy-loaded ConditionsSignInRiskDetectionsBlade, not captured; raw Graph
#     values are rendered)
#   - customAuthenticationFactors name resolution (portal parses
#     /identity/conditionalAccess/claimProviders definition JSON; raw ids rendered)

class ConditionalAccessDocHandler : DocumentationHandlerBase {
    ConditionalAccessDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.conditionalAccessPolicy')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        Invoke-CADocBasicInfo $obj
        Invoke-CADocUsersAndGroups $obj
        Invoke-CADocAgents $obj
        Invoke-CADocCloudApps $obj
        Invoke-CADocConditions $obj
        Invoke-CADocGrantControls $obj
        Invoke-CADocSessionControls $obj
    }
}

function Invoke-CADocBasicInfo {
    param($obj)

    Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.nameName')        $obj.displayName 'displayName'
    Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'AzureCA.conditionalAccessBladeTitle') '@odata.type'

    $state = switch ($obj.state) {
        'enabledForReportingButNotEnforced' { Get-LanguageString 'AzureCA.PolicyState.reportOnly' }
        'disabled'                          { Get-LanguageString 'AzureCA.PolicyState.off' }
        'partiallyEnabled'                  { Get-LanguageString 'AzureCA.policyStagedRollout' }
        default                             { Get-LanguageString 'AzureCA.PolicyState.on' }
    }
    Add-BasicPropertyValue (Get-LanguageString 'AzureCA.policyEnforceLabel') $state 'state'

    # Staged rollout target regions. Comma-joined tokens (ukSouth, usCentral, ...,
    # allNonListedRegions); the portal dropdown shows the raw tokens too - the
    # resource bundle has no localized labels for them.
    if ($obj.partialEnablementStrategy.targetRegions) {
        $ctx = Get-CurrentDocumentationContext
        $regions = @(($obj.partialEnablementStrategy.targetRegions -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        Add-BasicPropertyValue (Get-LanguageString 'AzureCA.targetRegions') ($regions -join $ctx.ObjectSeparator) 'targetRegions'
    }

    Add-BasicAdditionalValues $PolicyObject
}

function Invoke-CADocUsersAndGroups {
    param($obj)

    $ctx = Get-CurrentDocumentationContext
    $includeLabel = Get-LanguageString 'AzureCA.userSelectionBladeIncludeTabTitle'
    $excludeLabel = Get-LanguageString 'AzureCA.userSelectionBladeExcludeTabTitle'

    # Workload-identity / agent-identity policies target service principals via
    # conditions.clientApplications instead of conditions.users.
    if ($obj.conditions.clientApplications.includeServicePrincipals -or
        $obj.conditions.clientApplications.excludeServicePrincipals -or
        $obj.conditions.clientApplications.includeAgentIdServicePrincipals -or
        $obj.conditions.clientApplications.excludeAgentIdServicePrincipals) {
        ###################################################
        # Workload
        ###################################################

        # Well-known selection tokens (portal radio buttons) - never resolvable ids
        $spTokens = @('ServicePrincipalsInMyTenant', 'MicrosoftServicePrincipalsInMyTenant', 'ManagedIdentityServicePrincipalsInMyTenant', 'AllServicePrincipalsInMyTenant', 'None', 'All')
        $ids = @($obj.conditions.clientApplications.includeServicePrincipals +
                 $obj.conditions.clientApplications.excludeServicePrincipals +
                 $obj.conditions.clientApplications.includeAgentIdServicePrincipals +
                 $obj.conditions.clientApplications.excludeAgentIdServicePrincipals) | Where-Object { $_ -and $_ -notin $spTokens } | Get-Unique

        $category = Get-LanguageString "AzureCA.workloadIdentities"

        # Workload-identity service principals are source-tenant objects; resolve via
        # the shared helper (offline / source-unavailable guarded, returns @() offline).
        $idInfo = Resolve-CADirectoryObjectsById $ids

        if(($null -ne ($obj.conditions.clientApplications.includeServicePrincipals | Where-Object { $_ -eq "ServicePrincipalsInMyTenant"})))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.ServicePrincipalsCA.Radio.all"
                Category = $category
                SubCategory = $includeLabel
                EntityKey = "includeServicePrincipals"
            })
        }
        elseif(($null -ne ($obj.conditions.clientApplications.includeServicePrincipals | Where-Object { $_ -eq "MicrosoftServicePrincipalsInMyTenant"})))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.ServicePrincipalsCA.Radio.allFirstPartyApps" "All Microsoft service principals"
                Category = $category
                SubCategory = $includeLabel
                EntityKey = "includeServicePrincipals"
            })
        }
        elseif(($null -ne ($obj.conditions.clientApplications.includeServicePrincipals | Where-Object { $_ -eq "ManagedIdentityServicePrincipalsInMyTenant"})))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.ServicePrincipalsCA.Radio.allManagedIdentities" "All managed identities"
                Category = $category
                SubCategory = $includeLabel
                EntityKey = "includeServicePrincipals"
            })
        }
        elseif(($null -ne ($obj.conditions.clientApplications.includeServicePrincipals | Where-Object { $_ -eq "AllServicePrincipalsInMyTenant"})))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.ServicePrincipalsCA.Radio.allThirdPartyApps" "All owned single and multi tenant service principals"
                Category = $category
                SubCategory = $includeLabel
                EntityKey = "includeServicePrincipals"
            })
        }
        elseif(($null -ne ($obj.conditions.clientApplications.includeServicePrincipals | Where-Object { $_ -eq "None"})))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.chooseApplicationsNone"
                Category = $category
                SubCategory = $includeLabel
                EntityKey = "includeServicePrincipals"
            })
        }
        elseif($ids.Count -gt 0 -and $obj.conditions.clientApplications.includeServicePrincipals)
        {
            $tmpObjs = @()
            foreach($id in ($obj.conditions.clientApplications.includeServicePrincipals))
            {
                $idObj = $idInfo | Where-Object { $_.Id -eq $id }
                $tmpObjs += ?? $idObj.displayName $id
            }

            if($tmpObjs.count -gt 0)
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = $category
                    Value = $tmpObjs -join $ctx.ObjectSeparator
                    Category = $category
                    SubCategory = $includeLabel
                    EntityKey = "includeServicePrincipals"
                })
            }
        }

        if($obj.conditions.clientApplications.servicePrincipalFilter)
        {
            if($obj.conditions.clientApplications.servicePrincipalFilter.mode -eq "include")
            {
                $filterMode = "included"
            }
            else
            {
                $filterMode = "excluded"
            }

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.AppliesTo.$filterMode"
                Value = $obj.conditions.clientApplications.servicePrincipalFilter.rule
                Category = $category
                SubCategory = Get-LanguageString "AzureCA.CloudappsSelectionBlade.Filter.titleSP"
                EntityKey = "servicePrincipalFilter"
            })
        }

        if(($null -ne ($obj.conditions.clientApplications.excludeServicePrincipals | Where-Object { $_ -eq "ServicePrincipalsInMyTenant"})))
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $excludeLabel
                Value = Get-LanguageString "AzureCA.ServicePrincipalsCA.Radio.all"
                Category = $category
                SubCategory = $excludeLabel
                EntityKey = "excludeServicePrincipals"
            })
        }
        elseif($ids.Count -gt 0)
        {
            $tmpObjs = @()
            foreach($id in ($obj.conditions.clientApplications.excludeServicePrincipals))
            {
                $idObj = $idInfo | Where-Object { $_.Id -eq $id }
                $tmpObjs += ?? $idObj.displayName $id
            }

            if($tmpObjs.count -gt 0)
            {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = $category
                    Value = $tmpObjs -join $ctx.ObjectSeparator
                    Category = $category
                    SubCategory = $excludeLabel
                    EntityKey = "excludeServicePrincipals"
                })
            }
        }

        ###################################################
        # Agent identities (agentic service principals)
        ###################################################

        foreach ($side in @(
                @{ Vals = @($obj.conditions.clientApplications.includeAgentIdServicePrincipals); Sub = $includeLabel; Key = 'includeAgentIdServicePrincipals' },
                @{ Vals = @($obj.conditions.clientApplications.excludeAgentIdServicePrincipals); Sub = $excludeLabel; Key = 'excludeAgentIdServicePrincipals' })) {
            $vals = @($side.Vals | Where-Object { $_ })
            if ($vals.Count -eq 0) { continue }
            if ($vals -contains 'All') {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = $side.Sub
                    Value = Get-LanguageString 'AzureCA.ServicePrincipalsCA.AgenticResources.allSelectedGA' 'All agent identities (Preview)'
                    Category = $category; SubCategory = $side.Sub; EntityKey = $side.Key
                })
            }
            else {
                $names = foreach ($id in ($vals | Where-Object { $_ -notin $spTokens })) { $o = $idInfo | Where-Object { $_.Id -eq $id }; ?? $o.displayName $id }
                if (@($names).Count -gt 0) {
                    Add-CustomSettingObject ([PSCustomObject]@{
                        Name = $category
                        Value = (@($names) -join $ctx.ObjectSeparator)
                        Category = $category; SubCategory = $side.Sub; EntityKey = $side.Key
                    })
                }
            }
        }

        if ($obj.conditions.clientApplications.agentIdServicePrincipalFilter) {
            $mode = if ($obj.conditions.clientApplications.agentIdServicePrincipalFilter.mode -eq 'include') { 'included' } else { 'excluded' }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.AppliesTo.$mode"
                Value = $obj.conditions.clientApplications.agentIdServicePrincipalFilter.rule
                Category = $category
                SubCategory = Get-LanguageString 'AzureCA.CloudappsSelectionBlade.Filter.titleAgents'
                EntityKey = 'agentIdServicePrincipalFilter'
            })
        }
    }
    else {
        ###################################################
        # Users and groups (no workload identities)
        ###################################################

        $ids = @($obj.conditions.users.includeUsers + $obj.conditions.users.includeGroups + $obj.conditions.users.excludeUsers + $obj.conditions.users.excludeGroups) | Where-Object { $_ -and $_ -notin @('All', 'None', 'GuestsOrExternalUsers', 'AllAgentIdUsers') } | Get-Unique
        $directoryObjects = Resolve-CADirectoryObjectsById $ids

        $category     = Get-LanguageString 'AzureCA.usersGroupsLabel'
        $includeLabel = Get-LanguageString 'AzureCA.userSelectionBladeIncludeTabTitle'
        $excludeLabel = Get-LanguageString 'AzureCA.userSelectionBladeExcludeTabTitle'
        $users        = $obj.conditions.users

        $roleIds = @(@($users.includeRoles) + @($users.excludeRoles) | Where-Object { $_ } | Select-Object -Unique)
        $roles = Resolve-CADirectoryRolesById $roleIds

        # --- Include ---
        if ($users.includeUsers -contains 'All') {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel; Value = (Get-LanguageString 'AzureCA.allUsersString')
                Category = $category; SubCategory = $includeLabel; EntityKey = 'includeUsers'
            })
        }
        elseif ($users.includeUsers -contains 'None') {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel; Value = (Get-LanguageString 'AzureCA.chooseApplicationsNone')
                Category = $category; SubCategory = $includeLabel; EntityKey = 'includeUsers'
            })
        }
        elseif ($users.includeUsers -contains 'AllAgentIdUsers') {
            # Legacy all-agent-users shape; the portal migrates it to
            # conditions.agents.includeAgentUsers=["All"] on save.
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel; Value = (Get-LanguageString 'AzureCA.allAgentUsersGA')
                Category = $category; SubCategory = $includeLabel; EntityKey = 'includeUsers'
            })
        }
        else {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel; Value = (Get-LanguageString 'AzureCA.userSelectionBladeSelectedUsers')
                Category = $category; SubCategory = $includeLabel; EntityKey = 'includeUsers'
            })

            if ($users.includeUsers -contains 'GuestsOrExternalUsers') {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = (Get-LanguageString 'AzureCA.allGuestUserLabel')
                    Value = (Get-LanguageString 'Inputs.enabled')
                    Category = $category; SubCategory = $includeLabel; EntityKey = 'includeGuestsOrExternalUsers'
                })
            }

            Add-CADirectoryRoleRow -Ids @($users.includeRoles) -Roles $roles -Category $category -SubCategory $includeLabel -EntityKey 'includeRoles'
            Add-CADirectoryObjectRow -Ids @(@($users.includeUsers) + @($users.includeGroups)) -Objects $directoryObjects -Category $category -SubCategory $includeLabel -EntityKey 'includeUsersGroups'
        }

        # Structured guest / external-user targeting (guest types + external tenants).
        # The current portal experience populates includeGuestsOrExternalUsers /
        # excludeGuestsOrExternalUsers and does NOT set the legacy 'GuestsOrExternalUsers'
        # token handled below. The INCLUDE row must be emitted here, with the other
        # include rows and before the Exclude section: the output groups by row order
        # (a SubCategory header is emitted only when it changes), so an include row that
        # trails the exclude rows renders as a second, misplaced "Include" block.
        Add-CAGuestsOrExternalUsersRows $users.includeGuestsOrExternalUsers $category $includeLabel 'includeGuestsOrExternalUsers'

        # --- Exclude ---
        if ($users.excludeUsers -contains 'GuestsOrExternalUsers') {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'AzureCA.allGuestUserLabel')
                Value = (Get-LanguageString 'Inputs.enabled')
                Category = $category; SubCategory = $excludeLabel; EntityKey = 'excludeGuestsOrExternalUsers'
            })
        }

        Add-CADirectoryRoleRow -Ids @($users.excludeRoles) -Roles $roles -Category $category -SubCategory $excludeLabel -EntityKey 'excludeRoles'
        Add-CADirectoryObjectRow -Ids @(@($users.excludeUsers) + @($users.excludeGroups)) -Objects $directoryObjects -Category $category -SubCategory $excludeLabel -EntityKey 'excludeUsersGroups'
        Add-CAGuestsOrExternalUsersRows $users.excludeGuestsOrExternalUsers $category $excludeLabel 'excludeGuestsOrExternalUsers'
    }
}

function Resolve-CADirectoryObjectsById {
    param([object[]]$Ids)

    $idsToResolve = @($Ids | Where-Object { $_ } | Select-Object -Unique)
    if ($idsToResolve.Count -eq 0) { return @() }

    $ctx = Get-CurrentDocumentationContext
    if ($ctx.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) { return @() }

    try {
        $body = @{ ids = $idsToResolve } | ConvertTo-Json -Compress
        $resp = Invoke-MSGraphAPI -Url '/directoryObjects/getByIds?$select=displayName,id' -Content $body -HttpMethod 'POST'
        return @($resp.Value)
    }
    catch {
        Write-LogError 'ConditionalAccess: failed to resolve selected users and groups' $_.Exception
        return @()
    }
}

function Resolve-CADirectoryRolesById {
    param([object[]]$Ids)

    if (@($Ids | Where-Object { $_ }).Count -eq 0) { return @() }
    if ($script:_caDirectoryRoles) { return @($script:_caDirectoryRoles) }
    if ($script:_caDirectoryRolesLookupFailed) { return @() }

    $ctx = Get-CurrentDocumentationContext
    if ($ctx.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) { return @() }

    try {
        # roleDefinitions instead of directoryRoleTemplates: the portal blade uses role
        # definitions, which include CUSTOM roles (built-in definition ids equal the
        # legacy roleTemplateIds, so this is a strict superset).
        $resp = Invoke-MSGraphAPI -Url '/roleManagement/directory/roleDefinitions?$select=id,displayName' -ODataMetadata 'minimal' -AllPages
        if (-not $resp) {
            $script:_caDirectoryRolesLookupFailed = $true
            return @()
        }
        $script:_caDirectoryRoles = @($resp.Value)
        return @($script:_caDirectoryRoles)
    }
    catch {
        $script:_caDirectoryRolesLookupFailed = $true
        Write-LogError 'ConditionalAccess: failed to resolve directory roles' $_.Exception
        return @()
    }
}

function Add-CADirectoryObjectRow {
    param([object[]]$Ids, [object[]]$Objects, [string]$Category, [string]$SubCategory, [string]$EntityKey)

    $values = foreach ($id in @($Ids)) {
        if (-not $id -or $id -in @('All', 'None', 'GuestsOrExternalUsers', 'AllAgentIdUsers')) { continue }
        $resolved = $Objects | Where-Object Id -EQ $id | Select-Object -First 1
        if ($resolved.displayName) { $resolved.displayName } else { $id }
    }
    if (@($values).Count -eq 0) { return }

    $ctx = Get-CurrentDocumentationContext
    Add-CustomSettingObject ([PSCustomObject]@{
        Name = $Category; Value = ($values -join $ctx.ObjectSeparator)
        Category = $Category; SubCategory = $SubCategory; EntityKey = $EntityKey
    })
}

function Add-CADirectoryRoleRow {
    param([object[]]$Ids, [object[]]$Roles, [string]$Category, [string]$SubCategory, [string]$EntityKey)

    $values = foreach ($id in @($Ids)) {
        if (-not $id) { continue }
        $resolved = $Roles | Where-Object Id -EQ $id | Select-Object -First 1
        if ($resolved.displayName) { $resolved.displayName } else { $id }
    }
    if (@($values).Count -eq 0) { return }

    $ctx = Get-CurrentDocumentationContext
    Add-CustomSettingObject ([PSCustomObject]@{
        Name = (Get-LanguageString 'AzureCA.directoryRolesLabel'); Value = ($values -join $ctx.ObjectSeparator)
        Category = $Category; SubCategory = $SubCategory; EntityKey = $EntityKey
    })
}

# Render the structured guest / external-user selection (guestOrExternalUserTypes
# + externalTenants) for an include or exclude side. Newer than the legacy
# 'GuestsOrExternalUsers' token, which only ever meant "all guests".
function Add-CAGuestsOrExternalUsersRows {
    param($Guests, [string]$Category, [string]$SubCategory, [string]$EntityKey)

    if (-not $Guests -or -not $Guests.guestOrExternalUserTypes) { return }
    $ctx = Get-CurrentDocumentationContext

    $typeMap = @{
        'internalGuest'          = 'internalGuestLabel'
        'b2bCollaborationGuest'  = 'b2bCollaborationGuestLabel'
        'b2bCollaborationMember' = 'b2bCollaborationMemberLabel'
        'b2bDirectConnectUser'   = 'b2bDirectConnectUserLabel'
        'otherExternalUser'      = 'otherExternalUserLabel'
        'serviceProvider'        = 'serviceProviderUsersLabel'
    }

    $types = @(($Guests.guestOrExternalUserTypes -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne 'none' })
    $typeNames = foreach ($t in $types) {
        $key = $typeMap[$t]
        if ($key) { Get-LanguageString "AzureCA.GuestsOrExternalUsers.$key" } else { $t }
    }
    if (@($typeNames).Count -gt 0) {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString 'AzureCA.GuestsOrExternalUsers.guestOrExternalUsersLabel')
            Value = ($typeNames -join $ctx.ObjectSeparator)
            Category = $Category; SubCategory = $SubCategory; EntityKey = $EntityKey
        })
    }

    # External tenants: all, or an enumerated list of tenant IDs. The portal resolves
    # each tenant ID to its organization display name; fall back to the raw GUID.
    $et = $Guests.externalTenants
    if ($et -and $et.membershipKind) {
        if ($et.membershipKind -eq 'all') {
            $etValue = Get-LanguageString 'AzureCA.GuestsOrExternalUsers.allExternalTenantsLabel'
        }
        else {
            $members = @($et.members | Where-Object { $_ } | ForEach-Object { Resolve-CATenantDisplayName $_ })
            $etValue = if ($members.Count -gt 0) { ($members -join $ctx.ObjectSeparator) }
                       else { Get-LanguageString 'AzureCA.GuestsOrExternalUsers.enumeratedExternalTenantsLabel' }
        }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString 'AzureCA.GuestsOrExternalUsers.externalTenantsLabel')
            Value = $etValue
            Category = $Category; SubCategory = $SubCategory; EntityKey = ($EntityKey + '_externalTenants')
        })
    }
}

# Resolve an external tenant ID to its organization display name, like the portal's
# SelectOrganizationsGrid. Per-run cached; offline / source-unavailable / lookup
# failure falls back to the raw tenant ID.
function Resolve-CATenantDisplayName {
    param([string]$TenantId)

    if (-not $TenantId) { return $TenantId }
    if (-not $script:_caTenantNames) { $script:_caTenantNames = @{} }
    if ($script:_caTenantNames.ContainsKey($TenantId)) { return $script:_caTenantNames[$TenantId] }

    $name = $TenantId
    $ctx = Get-CurrentDocumentationContext
    if (-not $ctx.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
        try {
            $resp = Invoke-MSGraphAPI -Url "/tenantRelationships/findTenantInformationByTenantId(tenantId='$TenantId')" -ODataMetadata 'minimal' -NoError
            if ($resp.displayName) { $name = $resp.displayName }
        }
        catch {
            Write-LogError "ConditionalAccess: failed to resolve external tenant $TenantId" $_.Exception
        }
    }
    $script:_caTenantNames[$TenantId] = $name
    return $name
}

# Agents-acting-as-users targeting (conditions.agents = conditionalAccessAgents:
# includeAgentUsers / excludeAgentUsers / agentFilter). Distinct from workload
# identities (service principals) handled in Invoke-CADocUsersAndGroups.
function Invoke-CADocAgents {
    param($obj)

    $agents = $obj.conditions.agents
    if (-not $agents) { return }
    if (@($agents.includeAgentUsers).Count -eq 0 -and @($agents.excludeAgentUsers).Count -eq 0 -and -not $agents.agentFilter) { return }

    $ctx = Get-CurrentDocumentationContext
    $category     = Get-LanguageString 'AzureCA.agents'
    $includeLabel = Get-LanguageString 'AzureCA.userSelectionBladeIncludeTabTitle'
    $excludeLabel = Get-LanguageString 'AzureCA.userSelectionBladeExcludeTabTitle'

    $ids = @($agents.includeAgentUsers + $agents.excludeAgentUsers) | Where-Object { $_ -and $_ -notin @('All', 'None') } | Select-Object -Unique
    $idInfo = Resolve-CADirectoryObjectsById $ids

    foreach ($side in @(
            @{ Vals = @($agents.includeAgentUsers); Sub = $includeLabel; Key = 'includeAgentUsers' },
            @{ Vals = @($agents.excludeAgentUsers); Sub = $excludeLabel; Key = 'excludeAgentUsers' })) {
        $vals = @($side.Vals)
        if ($vals.Count -eq 0) { continue }
        if ($vals -contains 'All') {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $side.Sub; Value = (Get-LanguageString 'AzureCA.allAgentUsersGA')
                Category = $category; SubCategory = $side.Sub; EntityKey = $side.Key
            })
        }
        elseif ($vals -contains 'None') {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $side.Sub; Value = (Get-LanguageString 'AzureCA.chooseApplicationsNone')
                Category = $category; SubCategory = $side.Sub; EntityKey = $side.Key
            })
        }
        else {
            $names = foreach ($id in $vals) { $o = $idInfo | Where-Object { $_.Id -eq $id }; ?? $o.displayName $id }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $category; Value = ($names -join $ctx.ObjectSeparator)
                Category = $category; SubCategory = $side.Sub; EntityKey = $side.Key
            })
        }
    }

    if ($agents.agentFilter) {
        $mode = if ($agents.agentFilter.mode -eq 'include') { 'included' } else { 'excluded' }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString "AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.AppliesTo.$mode")
            Value = $agents.agentFilter.rule
            Category = $category
            SubCategory = (Get-LanguageString 'AzureCA.CloudappsSelectionBlade.Filter.titleAgentUsers')
            EntityKey = 'agentFilter'
        })
    }
}

function Invoke-CADocCloudApps {
    param($obj)

    # Review properties: applicationFilter

    $category       = Get-LanguageString 'AzureCA.NetworkAccess.targetResourcesSelectorTitle'
    $cloudAppsLabel = Get-LanguageString 'AzureCA.policyResourcesFormerlyCloudAppsLabel'
    $includeLabel   = Get-LanguageString 'AzureCA.userSelectionBladeIncludeTabTitle'
    $excludeLabel   = Get-LanguageString 'AzureCA.userSelectionBladeExcludeTabTitle'

    ### Include
    if ($obj.conditions.applications.includeApplications -contains 'All') {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value = (Get-LanguageString 'AzureCA.cloudappsSelectionBladeAllResources')
            Category = $category; SubCategory = $cloudAppsLabel; EntityKey = 'includeApplications'
        })
    }
    elseif ($obj.conditions.applications.includeApplications -contains 'None') {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value = (Get-LanguageString 'AzureCA.chooseApplicationsNone')
            Category = $category; SubCategory = $cloudAppsLabel; EntityKey = 'includeApplications'
        })
    }
    elseif ($obj.conditions.applications.includeApplications.Count -gt 0) {
        $ctx = Get-CurrentDocumentationContext
        $names = @(Get-CACloudAppNames $obj.conditions.applications.includeApplications)
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value = ($names -join $ctx.ObjectSeparator)
            Category = $category; SubCategory = $cloudAppsLabel; EntityKey = 'includeApplications'
        })
    }

    ### Exclude
    if ($obj.conditions.applications.excludeApplications.Count -gt 0) {
        $ctx = Get-CurrentDocumentationContext
        $names = @(Get-CACloudAppNames $obj.conditions.applications.excludeApplications)
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $excludeLabel
            Value = ($names -join $ctx.ObjectSeparator)
            Category = $category; SubCategory = $cloudAppsLabel; EntityKey = 'excludeApplications'
        })
    }

    ### Filter for applications
    if ($obj.conditions.applications.applicationFilter) {
        $appFilterMode = if ($obj.conditions.applications.applicationFilter.mode -eq 'include') { 'included' } else { 'excluded' }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString "AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.AppliesTo.$appFilterMode")
            Value = $obj.conditions.applications.applicationFilter.rule
            Category = $category; SubCategory = $cloudAppsLabel; EntityKey = 'applicationFilter'
        })
    }

    if ($obj.conditions.applications.includeUserActions.Count -gt 0) {
        $values = foreach ($action in @($obj.conditions.applications.includeUserActions)) {
            switch ($action) {
                'urn:user:registersecurityinfo' { Get-LanguageString 'AzureCA.UserActions.registerSecurityInfo' }
                'urn:user:registerdevice'       { Get-LanguageString 'AzureCA.UserActions.registerOrJoinDevices' }
                'urn:user:accountrecovery'      { Get-LanguageString 'AzureCA.UserActions.accountRecovery' }
                default                         { $action }
            }
        }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString 'AzureCA.UserActions.selectionInfo')
            Value = ($values -join (Get-CurrentDocumentationContext).ObjectSeparator)
            Category = $category; SubCategory = (Get-LanguageString 'AzureCA.UserActions.label'); EntityKey = 'includeUserActions'
        })
    }

    if ($obj.conditions.applications.includeAuthenticationContextClassReferences.Count -gt 0) {
        $ctx = Get-CurrentDocumentationContext
        $authContexts = Resolve-CAAuthContextClassRefs
        $values = foreach ($ref in @($obj.conditions.applications.includeAuthenticationContextClassReferences)) {
            # urn:microsoft:req1..3 = the "Accessing secured app data" Level 1-3 checkboxes
            if ($ref -match '^urn:microsoft:req([1-3])$') {
                '{0}: {1}' -f (Get-LanguageString 'AzureCA.UserActions.accessRequirementsLabel'), (Get-LanguageString "AzureCA.UserActions.accessRequirement$($Matches[1])")
            }
            else {
                $refObj = $authContexts | Where-Object { $_.Id -eq $ref } | Select-Object -First 1
                ?? $refObj.displayName $ref
            }
        }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name        = (Get-LanguageString 'AzureCA.AuthContext.label')
            Value       = ($values -join $ctx.ObjectSeparator)
            Category    = $category
            SubCategory = $cloudAppsLabel
            EntityKey   = 'includeAuthenticationContextClassReferences'
        })
    }

    ### Global Secure Access traffic profiles (legacy shape; newer saves translate the
    ### selection into the well-known GSA appIds and null this out)
    $trafficProfiles = @(($obj.conditions.applications.globalSecureAccess.includeTrafficProfiles -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($trafficProfiles.Count -gt 0) {
        $ctx = Get-CurrentDocumentationContext
        $values = foreach ($p in $trafficProfiles) {
            switch ($p) {
                'M365'     { Get-LanguageString 'AzureCA.NetworkAccess.m365OptionText' }
                'Internet' { Get-LanguageString 'AzureCA.NetworkAccess.internetOptionText' }
                'Private'  { Get-LanguageString 'AzureCA.NetworkAccess.privateOptionText' }
                default    { $p }
            }
        }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name        = (Get-LanguageString 'AzureCA.NetworkAccess.selectTrafficProfilesLabel')
            Value       = ($values -join $ctx.ObjectSeparator)
            Category    = $category
            SubCategory = $cloudAppsLabel
            EntityKey   = 'globalSecureAccess'
        })
    }
}

# Resolve authentication context class references (c1..c25) to display names like the
# portal's AuthContextService. Per-run cached; offline falls back to the raw id.
function Resolve-CAAuthContextClassRefs {
    if ($null -ne $script:_caAuthContexts) { return @($script:_caAuthContexts) }

    $ctx = Get-CurrentDocumentationContext
    if ($ctx.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) {
        $script:_caAuthContexts = @()
        return @()
    }

    try {
        # NB: this endpoint rejects $select ("Query option 'Select' is not allowed")
        $resp = Invoke-MSGraphAPI -Url '/identity/conditionalAccess/authenticationContextClassReferences' -ODataMetadata 'minimal'
        $script:_caAuthContexts = @($resp.Value)
    }
    catch {
        Write-LogError 'ConditionalAccess: failed to resolve authentication contexts' $_.Exception
        $script:_caAuthContexts = @()
    }
    return @($script:_caAuthContexts)
}

# Client apps sub-section under the "Cloud apps or actions" category. Critical
# for legacy-auth-blocking policies whose whole purpose is restricting
# clientAppTypes to exchangeActiveSync + other.
# Ported from old DocumentationCustom.psm1:3100-3127. `all` is the catch-all
# "not configured" sentinel: when present, the rest of the array is meaningless
# and the row is suppressed entirely (matches portal behaviour).
function Invoke-CADocClientApps {
    param($obj)

    $appTypes = @($obj.conditions.clientAppTypes)
    if ($appTypes.Count -eq 0) { return }
    if ($appTypes -contains 'all') { return }

    # Client apps live INSIDE the Conditions pane in the portal; keep the same
    # category as the surrounding condition rows so grouping doesn't flip mid-section.
    $category    = Get-LanguageString 'AzureCA.policyTriggersSelectorLabel'
    $subCategory = Get-LanguageString 'AzureCA.policyConditioniClientApp'
    $includeLabel = Get-LanguageString 'AzureCA.userSelectionBladeIncludeTabTitle'

    $rendered = foreach ($id in $appTypes) {
        switch ($id) {
            'browser'                    { Get-LanguageString 'AzureCA.clientAppWebBrowser' }
            'mobileAppsAndDesktopClients'{ Get-LanguageString 'AzureCA.clientAppMobileDesktop' }
            'exchangeActiveSync'         { Get-LanguageString 'AzureCA.clientAppExchangeActiveSync' }
            # easSupported replaces exchangeActiveSync when "Apply policy only to
            # supported platforms" is checked; easUnsupported is its legacy sibling
            'easSupported'               { Get-LanguageString 'AzureCA.WhatIfBlade.ClientApp.easSupported' }
            'easUnsupported'             { Get-LanguageString 'AzureCA.WhatIfBlade.ClientApp.easUnsupported' }
            'other'                      { Get-LanguageString 'AzureCA.clientTypeOtherClients' }
            default                      { Write-Log "ConditionalAccess: unsupported clientAppType '$id'" 2; $id }
        }
    }

    if ($rendered.Count -eq 0) { return }
    $ctx = Get-CurrentDocumentationContext
    Add-CustomSettingObject ([PSCustomObject]@{
        Name        = $includeLabel
        Value       = ($rendered -join $ctx.ObjectSeparator)
        Category    = $category
        SubCategory = $subCategory
        EntityKey   = 'clientAppTypes'
    })
}

function Resolve-CACloudAppsByIds {
    param([object[]]$AppIds)
    # Cache the FULL service-principal list once per run (matches old Get-CDAllCloudApps).
    # The previous per-policy "appId in (...)" filter was fragile AND cached only the
    # first policy's apps, so every later policy fell back to raw GUIDs. $AppIds is kept
    # for call-site symmetry but the whole list is fetched/cached and looked up by caller.
    if ($null -ne $script:_caCloudApps) { return @($script:_caCloudApps) }

    $ctx = Get-CurrentDocumentationContext
    if ($ctx.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) {
        $script:_caCloudApps = @()   # negative-cache so we don't retry per policy
        return @()
    }

    try {
        $resp = Invoke-MSGraphAPI -Url "/servicePrincipals?`$select=displayName,appId&`$top=999" -ODataMetadata 'minimal' -AllPages
        $script:_caCloudApps = @($resp.Value)
    }
    catch {
        Write-LogError 'ConditionalAccess: failed to resolve cloud apps' $_.Exception
        $script:_caCloudApps = @()
    }
    return @($script:_caCloudApps)
}

# Resolve an include/exclude application-id list to display names. Well-known
# selection tokens map to their localized labels; everything else resolves via the
# cached service-principal list, falling back to the raw id when not found.
function Get-CACloudAppNames {
    param([object[]]$AppIds)
    $resolved = Resolve-CACloudAppsByIds $AppIds
    foreach ($id in $AppIds) {
        switch ($id) {
            'Office365'             { Get-LanguageString 'AzureCA.PolicyTemplates.Summary.CloudApps.office365' -DefaultValue $id; break }
            'MicrosoftAdminPortals' { Get-LanguageString 'AzureCA.microsoftAdminPortals' -DefaultValue $id; break }
            'AllAgentIdResources'   { Get-LanguageString 'AzureCA.CloudappsSelectionBlade.AgenticResources.allAgenticResourcesGA' -DefaultValue $id; break }
            default {
                $sp = $resolved | Where-Object { $_.appId -eq $id } | Select-Object -First 1
                if ($sp.displayName) { $sp.displayName }
                else {
                    # Well-known Global Secure Access appIds ("All internet resources with
                    # Global Secure Access" writes these as ordinary GUID entries; they may
                    # not resolve as tenant service principals, especially offline)
                    switch ($id) {
                        'c08f52c9-8f03-4558-a0ea-9a4c878cf343' { Get-LanguageString 'AzureCA.NetworkAccess.m365OptionText' }
                        '5dc48733-b5df-475c-a49b-fa307ef00853' { Get-LanguageString 'AzureCA.NetworkAccess.internetOptionText' }
                        'e92b9b37-1b47-4c01-9fbc-91d84450870e' { Get-LanguageString 'AzureCA.NetworkAccess.privateOptionText' }
                        default { $id }
                    }
                }
            }
        }
    }
}

function Invoke-CADocConditions {
    param($obj)

    $ctx = Get-CurrentDocumentationContext
    # "Conditions" section title. The old key AzureCA.helpConditionsTitle no longer
    # exists under AzureCA (it moved to the AzureIAM namespace); policyTriggersSelectorLabel
    # is the reachable AzureCA equivalent (= "Conditions"). Without this the whole
    # Conditions category rendered blank for every CA policy.
    $category     = Get-LanguageString "AzureCA.policyTriggersSelectorLabel"
    $includeLabel = Get-LanguageString 'AzureCA.userSelectionBladeIncludeTabTitle'
    $excludeLabel = Get-LanguageString 'AzureCA.userSelectionBladeExcludeTabTitle'

    #$category = Get-LanguageString "AzureCA.policyConditionUserRisk"

    ### User risk
    if($obj.conditions.userRiskLevels.Count -gt 0)
    {
        $tmpObjs = @()
        foreach($id in ($obj.conditions.userRiskLevels))
        {
            $tmpObjs += Get-LanguageString "AzureCA.$($id)Risk"
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.policyConditionUserRisk"
            EntityKey = "userRiskLevels"
        })
    }

    ### Sign-in risk
    if($obj.conditions.signInRiskLevels.Count -gt 0)
    {
        $tmpObjs = @()
        foreach($id in ($obj.conditions.signInRiskLevels))
        {
            $tmpObjs += Get-LanguageString "AzureCA.$($id)Risk"
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.policyConditionSigninRisk"
            EntityKey = "signInRiskLevels"
        })
    }

    ### Insider risk
    ### NB: unlike user/sign-in/SP risk (arrays), insiderRiskLevels is a FLAGS enum -
    ### Graph returns one comma-joined string ("minor,moderate")
    $insiderRisks = @(($obj.conditions.insiderRiskLevels -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if($insiderRisks.Count -gt 0)
    {
        $tmpObjs = @()
        foreach($id in $insiderRisks)
        {
            $tmpObjs += Get-LanguageString "AzureCA.$($id)Risk"
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.PolicyCondition.InsiderRisk.label"
            EntityKey = "insiderRiskLevels"
        })
    }

    ### Sign-in risk detections
    ### NB: the per-detection checkbox labels live in the lazy-loaded
    ### ConditionsSignInRiskDetectionsBlade, which is not part of the captured portal
    ### bundles - no language ids are derivable, so the raw Graph values are rendered.
    if($obj.conditions.signInRiskDetections.includeDetections.Count -gt 0)
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value = @($obj.conditions.signInRiskDetections.includeDetections) -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.PolicyCondition.SignInRiskDetections.label"
            EntityKey = "signInRiskDetections"
        })
    }

    ### Service principal risk (workload-identity policies)
    if($obj.conditions.servicePrincipalRiskLevels.Count -gt 0)
    {
        $tmpObjs = @()
        foreach($id in ($obj.conditions.servicePrincipalRiskLevels))
        {
            $tmpObjs += Get-LanguageString "AzureCA.$($id)Risk"
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.Policy.Condition.ServicePrincipalRisk.title"
            EntityKey = "servicePrincipalRiskLevels"
        })
    }

    ### Agent ID risk (flags string: high/medium/low)
    $agentIdRisks = @(($obj.conditions.agentIdRiskLevels -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if($agentIdRisks.Count -gt 0)
    {
        $tmpObjs = foreach($id in $agentIdRisks) { Get-LanguageString "AzureCA.$($id)Risk" }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value = $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.ServicePrincipalsCA.AgenticResources.agentRisk" "Agent risk (Preview)"
            EntityKey = "agentIdRiskLevels"
        })
    }

    ### Agent session risk (flags string: high/medium/low)
    $agentSessionRisks = @(($obj.conditions.agentSessionRiskLevels -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if($agentSessionRisks.Count -gt 0)
    {
        $tmpObjs = foreach($id in $agentSessionRisks) { Get-LanguageString "AzureCA.$($id)Risk" }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value = $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.PolicyBlade.Conditions.AgentSessionRisk.label"
            EntityKey = "agentSessionRiskLevels"
        })
    }

    ### Assistive agent protections (agent context)
    foreach ($side in @(
            @{ Vals = @($obj.conditions.agentContext.includeAgentContexts); Sub = $includeLabel; Key = 'includeAgentContexts' },
            @{ Vals = @($obj.conditions.agentContext.excludeAgentContexts); Sub = $excludeLabel; Key = 'excludeAgentContexts' })) {
        $vals = @($side.Vals | Where-Object { $_ })
        if ($vals.Count -eq 0) { continue }
        $tmpObjs = foreach ($id in $vals) {
            switch ($id) {
                'allAgentAssistedFlows'      { Get-LanguageString 'AzureCA.PolicyBlade.Conditions.AgentContext.ContextPane.allAgentAssistedFlows' }
                'allHostingApplicationFlows' { Get-LanguageString 'AzureCA.PolicyBlade.Conditions.AgentContext.ContextPane.allHostingApplicationFlows' }
                default                      { $id }
            }
        }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $side.Sub
            Value = $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString 'AzureCA.PolicyBlade.Conditions.AgentContext.label'
            EntityKey = $side.Key
        })
    }

    ### Device platforms
    if($obj.conditions.platforms.includePlatforms.Count -gt 0)
    {
        $tmpObjs = @()
        foreach($id in ($obj.conditions.platforms.includePlatforms))
        {
            if($id -eq "all")
            {
                $tmpObjs += Get-LanguageString "AzureCA.allDevicePlatforms"
            }
            else
            {
                $tmpObjs += Get-LanguageString "AzureCA.$($id)DisplayName"
            }
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.policyConditionDevicePlatform"
            EntityKey = "includePlatforms"
        })
    }

    if($obj.conditions.platforms.excludePlatforms.Count -gt 0)
    {
        $tmpObjs = @()
        foreach($id in ($obj.conditions.platforms.excludePlatforms))
        {
            $tmpObjs += Get-LanguageString "AzureCA.$($id)DisplayName"
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $excludeLabel
            Value =  $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.policyConditionDevicePlatform"
            EntityKey = "excludePlatforms"
        })
    }

    ### Locations
    # ToDo: Create function and should be moved to Netwok category
    if(-not $script:allNamedLocations -and ($obj.conditions.locations.includeLocations.Count -gt 0 -or $obj.conditions.locations.excludeLocations.Count))
    {
        # Named locations are source-tenant objects; offline / source-unavailable falls
        # back to the raw location IDs (the lookups below use `?? displayName id`).
        if(-not $ctx.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable))
        {
            try {
                $script:allNamedLocations = (Invoke-MSGraphAPI -url "/identity/conditionalAccess/namedLocations?`$select=displayName,Id&top=999" -ODataMetadata "minimal").value
            } catch {
                Write-LogError 'ConditionalAccess: failed to resolve named locations' $_.Exception
            }
        }
        if(-not $script:allNamedLocations) {  $script:allNamedLocations = @()}
        elseif($script:allNamedLocations -isnot [Object[]]) {  $script:allNamedLocations = @($script:allNamedLocations) }

        $script:allNamedLocations += [PSCustomObject]@{
            displayName = Get-LanguageString "AzureCA.chooseLocationTrustedIpsItem"
            id =  "00000000-0000-0000-0000-000000000000"
        }
    }

    if($obj.conditions.locations.includeLocations.Count -gt 0)
    {
        $tmpObjs = @()
        foreach($id in ($obj.conditions.locations.includeLocations))
        {
            if($id -eq "AllTrusted")
            {
                $tmpObjs += Get-LanguageString "AzureCA.allTrustedLocationLabel"
            }
            elseif($id -eq "All")
            {
                $tmpObjs += Get-LanguageString "AzureCA.locationsAllLocationsLabel"
            }
            else
            {
                $idObj = $script:allNamedLocations | Where-Object { $_.Id -eq $id }
                $tmpObjs += ?? $idObj.displayName $id
            }
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.policyConditionLocation"
            EntityKey = "includeLocations"
        })
    }

    if($obj.conditions.locations.excludeLocations.Count -gt 0)
    {
        $tmpObjs = @()
        foreach($id in ($obj.conditions.locations.excludeLocations))
        {
            if($id -eq "AllTrusted")
            {
                $tmpObjs += Get-LanguageString "AzureCA.allTrustedLocationLabel"
            }
            elseif($id -eq "All")
            {
                $tmpObjs += Get-LanguageString "AzureCA.locationsAllLocationsLabel"
            }
            else
            {
                $idObj = $script:allNamedLocations | Where-Object { $_.Id -eq $id }
                $tmpObjs += ?? $idObj.displayName $id
            }
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $excludeLabel
            Value =  $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.policyConditionLocation"
            EntityKey = "excludeLocations"
        })
    }

    ### Client apps
    Invoke-CADocClientApps $obj

    ### Time (Preview)
    if($obj.conditions.times)
    {
        $times = $obj.conditions.times
        $timesSub = Get-LanguageString "AzureCA.timeConditionSelectorLabel"

        if($times.includeAllTimes -eq $true)
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $includeLabel
                Value = Get-LanguageString "AzureCA.timeSelectorAllTimesText"
                Category = $category
                SubCategory = $timesSub
                EntityKey = "includeAllTimes"
            })
        }

        foreach ($side in @(
                @{ Days = $times.includeDays; Range = $times.includeRange; Sub = $includeLabel; Key = 'includeTimes' },
                @{ Days = $times.excludeDays; Range = $times.excludeRange; Sub = $excludeLabel; Key = 'excludeTimes' })) {
            $parts = @()
            if ($side.Days.daysOfWeek) {
                $dayNames = @($side.Days.daysOfWeek | ForEach-Object { Get-LanguageString "SettingDetails.$_" })
                $txt = $dayNames -join $ctx.PropertySeparator
                if ($side.Days.startTime -and $side.Days.endTime) { $txt = '{0} {1}-{2}' -f $txt, $side.Days.startTime, $side.Days.endTime }
                if ($side.Days.timeZone) { $txt = '{0} ({1})' -f $txt, $side.Days.timeZone }
                $parts += $txt
            }
            if ($side.Range.startDateTime -and $side.Range.endDateTime) {
                $txt = '{0} - {1}' -f $side.Range.startDateTime, $side.Range.endDateTime
                if ($side.Range.timeZone) { $txt = '{0} ({1})' -f $txt, $side.Range.timeZone }
                $parts += $txt
            }
            if ($parts.Count -gt 0) {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = $side.Sub
                    Value = $parts -join $ctx.ObjectSeparator
                    Category = $category
                    SubCategory = $timesSub
                    EntityKey = $side.Key
                })
            }
        }
    }

    ###Filter for devices
    if($obj.conditions.devices.includeDevices.Count -gt 0)
    {
        # The portal only ever writes ["All"], but Graph accepts free strings for
        # API-authored policies - map known values, pass anything else through
        $tmpObjs = @()
        foreach($id in ($obj.conditions.devices.includeDevices))
        {
            $tmpObjs += switch ($id) {
                'All'          { Get-LanguageString "AzureCA.deviceStateAll" }
                'Compliant'    { Get-LanguageString "AzureCA.deviceStateCompliant" }
                'DomainJoined' { Get-LanguageString "AzureCA.deviceStateDomainJoined" }
                default        { $id }
            }
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value =  $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.deviceStateConditionSelectorLabel"
            EntityKey = "includeDevices"
        })
    }

    if($obj.conditions.devices.excludeDevices.Count -gt 0)
    {
        # Blade wording ("Device marked as compliant" / "Device Microsoft Entra hybrid
        # joined"), not the classic-policy grant-control strings
        $tmpObjs = @()
        foreach($id in ($obj.conditions.devices.excludeDevices))
        {
            $tmpObjs += switch ($id) {
                'Compliant'    { Get-LanguageString "AzureCA.deviceStateCompliant" }
                'DomainJoined' { Get-LanguageString "AzureCA.deviceStateDomainJoined" }
                default        { $id }
            }
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $excludeLabel
            Value =  $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.deviceStateConditionSelectorLabel"
            EntityKey = "excludeDevices"
        })
    }

    if($obj.conditions.devices.deviceFilter)
    {
        if($obj.conditions.devices.deviceFilter.mode -eq "include")
        {
            $filterMode = "included"
        }
        else
        {
            $filterMode = "excluded"
        }

        #AzureCA.PolicyBlade.Conditions.DeviceAttributes.AssignmentFilter.Blade
        #AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.title
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.AppliesTo.$filterMode"
            Value = $obj.conditions.devices.deviceFilter.rule
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.PolicyBlade.Conditions.DeviceAttributes.Blade.title"
            EntityKey = "deviceFilter"
        })
    }

    ### Authentication flow (transferMethods is a flags string: deviceCodeFlow / authenticationTransfer)
    $transferMethods = @(($obj.conditions.authenticationFlows.transferMethods -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne 'none' })
    if ($transferMethods.Count -gt 0)
    {
        $tmpObjs = foreach ($m in $transferMethods) { Get-LanguageString "AzureCA.PolicyBlade.Conditions.AuthenticationFlows.Selector.$m" }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString "AzureCA.PolicyBlade.Conditions.AuthenticationFlows.Selector.label")
            Value = $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.PolicyBlade.Conditions.AuthenticationFlows.Selector.label"
            EntityKey = "authenticationFlows"
        })
    }

    ### Protocol and flows (second flags string on authenticationFlows, next to
    ### transferMethods; member values are not in the CSDL yet - raw-value fallback)
    $protocolFlows = @(($obj.conditions.authenticationFlows.protocolFlows -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne 'none' })
    if ($protocolFlows.Count -gt 0)
    {
        $tmpObjs = foreach ($m in $protocolFlows) { Get-LanguageString "AzureCA.PolicyBlade.Conditions.AuthenticationFlows.Selector.$m" $m -IgnoreMissing }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString "AzureCA.PolicyBlade.Conditions.AuthenticationFlows.Selector.protocolFlows")
            Value = $tmpObjs -join $ctx.ObjectSeparator
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.PolicyBlade.Conditions.AuthenticationFlows.Selector.label"
            EntityKey = "protocolFlows"
        })
    }

    ### Purview Data Loss Prevention (value is the single sentinel
    ### "purviewDataLossPreventionBrowser" - the portal shows just "Configured")
    if ($obj.conditions.purviewDataLossPreventionRules)
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = $includeLabel
            Value = Get-LanguageString "AzureCA.PolicyBlade.Conditions.PurviewDLP.Selector.configured"
            Category = $category
            SubCategory = Get-LanguageString "AzureCA.PolicyBlade.Conditions.PurviewDLP.Selector.label"
            EntityKey = "purviewDataLossPreventionRules"
        })
    }
}

function Invoke-CADocGrantControls {
    param($obj)

    if (-not $obj.grantControls) { return }

    $ctx = Get-CurrentDocumentationContext
    $category = Get-LanguageString 'AzureCA.policyControlBladeTitle'

    $isBlock = $obj.grantControls.builtInControls -contains 'block'
    $controlValueKey = if ($isBlock) { 'AzureCA.policyControlBlockAccessDisplayedName' } else { 'AzureCA.policyControlAllowAccessDisplayedName' }
    Add-CustomSettingObject ([PSCustomObject]@{
        Name = (Get-LanguageString 'AzureCA.policyControlContentDescription')
        Value = (Get-LanguageString $controlValueKey)
        Category = $category; SubCategory = ''; EntityKey = 'policyControl'
    })

    # Block-access ends the section — nothing else applies
    if ($isBlock) { return }

    # Per-control rows (each shows "Enabled" when present).
    # NB despite the key names: approvedApplication/policyControlRequireMamDisplayedName
    # = "Require approved client app" and compliantApplication/
    # policyControlRequireCompliantAppDisplayedName = "Require app protection policy".
    $controlMap = [ordered]@{
        'mfa'                = 'AzureCA.policyControlMfaChallengeDisplayedName'
        'authenticationStrength' = 'AzureCA.policyControlAuthenticationStrengthDisplayedName'
        'compliantDevice'    = 'AzureCA.policyControlCompliantDeviceDisplayedName'
        'domainJoinedDevice' = 'AzureCA.policyControlRequireDomainJoinedDisplayedName'
        'approvedApplication'  = 'AzureCA.policyControlRequireMamDisplayedName'
        'compliantApplication' = 'AzureCA.policyControlRequireCompliantAppDisplayedName'
        'passwordChange'     = 'AzureCA.policyControlRequiredPasswordChangeDisplayedName'
        'riskRemediation'    = 'AzureCA.policyControlRequireRiskRemediationDisplayedName'
        'verifiedID'         = 'AzureCA.policyControlVerifiedIdDisplayedName'
    }

    foreach ($key in $controlMap.Keys) {

        if ($key -eq 'authenticationStrength') {
            if($obj.grantControls.authenticationStrength.displayName) {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = (Get-LanguageString 'AzureCA.policyControlAuthenticationStrengthDisplayedName')
                    Value = $obj.grantControls.authenticationStrength.displayName
                    Category = $category; SubCategory = ''; EntityKey = 'authenticationStrength'
                })
            }
        }
        elseif ($obj.grantControls.builtInControls -contains $key) {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString $controlMap[$key])
                Value = (Get-LanguageString 'Inputs.enabled')
                Category = $category; SubCategory = ''; EntityKey = $key
            })
        }
    }

    ### Terms of use
    if (($obj.grantControls.termsOfUse | Measure-Object).Count -gt 0) {
        # Terms of Use are source-tenant objects; offline / source-unavailable falls back to raw IDs.
        if (-not $script:allTermsOfUse -and -not $ctx.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
            try {
                $script:allTermsOfUse = (Invoke-MSGraphAPI -Url "/identityGovernance/termsOfUse/agreements?`$select=displayName,Id&top=999" -ODataMetadata 'minimal').value
            } catch {
                Write-LogError 'ConditionalAccess: failed to resolve terms of use' $_.Exception
            }
        }
        if (-not $script:allTermsOfUse) { $script:allTermsOfUse = @() }
        elseif ($script:allTermsOfUse -isnot [Object[]]) { $script:allTermsOfUse = @($script:allTermsOfUse) }

        $names = foreach ($id in $obj.grantControls.termsOfUse) {
            $touObj = $script:allTermsOfUse | Where-Object { $_.Id -eq $id }
            if ($touObj.displayName) { $touObj.displayName } else { $id }
        }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name  = (Get-LanguageString 'AzureCA.menuItemTermsOfUse')
            Value = ($names -join $ctx.ObjectSeparator)
            Category = $category; SubCategory = ''; EntityKey = 'termsOfUse'
        })
    }

    ### Custom controls (legacy claim-provider controls)
    if (@($obj.grantControls.customAuthenticationFactors).Count -gt 0) {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name  = (Get-LanguageString 'AzureCA.menuItemClaimProviderControls')
            Value = ($obj.grantControls.customAuthenticationFactors -join $ctx.ObjectSeparator)
            Category = $category; SubCategory = ''; EntityKey = 'customAuthenticationFactors'
        })
    }

    # AND / OR operator
    $controlCount = @($obj.grantControls.builtInControls).Count + [int][bool]$obj.grantControls.authenticationStrength + @($obj.grantControls.customAuthenticationFactors).Count
    if ($controlCount -ge 1) {
        $operatorKey = if ($obj.grantControls.operator -eq 'OR') { 'AzureCA.requireOneControlText' } else { 'AzureCA.requireAllControlsText' }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString 'AzureCA.descriptionContentForControlsAndOr')
            Value = (Get-LanguageString $operatorKey)
            Category = $category; SubCategory = ''; EntityKey = 'grantOperator'
        })
    }
}

function Invoke-CADocSessionControls
{
    param($obj)

    if(-not $obj.sessionControls) { return }

    $category = Get-LanguageString "AzureCA.sessionControlBladeTitle"

    # Use app enfore restrictions
    if($obj.sessionControls.applicationEnforcedRestrictions.isEnabled -eq $true)
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.sessionControlsAppEnforcedLabel"
            Value = Get-LanguageString "Inputs.enabled"
            Category = $category
            SubCategory = ""
            EntityKey = "applicationEnforcedRestrictions"
        })
    }

    #Use Conditional Access App Control
    if($obj.sessionControls.cloudAppSecurity.isEnabled -eq $true)
    {
        $strId = $null
        if($obj.sessionControls.cloudAppSecurity.cloudAppSecurityType -eq "mcasConfigured") { $strId = "useCustomControls" }
        elseif($obj.sessionControls.cloudAppSecurity.cloudAppSecurityType -eq "monitorOnly") { $strId = "monitorOnly" }
        elseif($obj.sessionControls.cloudAppSecurity.cloudAppSecurityType -eq "blockDownloads") { $strId = "blockDownloads" }
        elseif($obj.sessionControls.cloudAppSecurity.cloudAppSecurityType -eq "protectDownloads") { $strId = "protectDownloads" }

        $casValue = if ($strId) { Get-LanguageString "AzureCA.CAS.BuiltinPolicy.Option.$strId" }
                    else { $obj.sessionControls.cloudAppSecurity.cloudAppSecurityType }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.sessionControlsCasLabel"
            Value =  $casValue
            Category = $category
            SubCategory = ""
            EntityKey = "cloudAppSecurity"
        })
    }

    # Sign-in frequency. frequencyInterval decides the shape: 'timeBased' carries
    # type (hours/days) + value; 'everyTime' and 'everyTimeStrict' both null out
    # type/value, so they MUST be told apart on frequencyInterval, not type.
    if($obj.sessionControls.signInFrequency.isEnabled -eq $true)
    {
        $sif = $obj.sessionControls.signInFrequency
        if($sif.frequencyInterval -eq "everyTimeStrict")
        {
            $value = Get-LanguageString "AzureCA.SessionControls.SignInFrequency.everyTimeStrict"
        }
        elseif($sif.type -eq "hours")
        {
            if($sif.value -gt 1)
            {
                $value = (Get-LanguageString "AzureCA.SessionLifetime.SignInFrequency.Option.Hour.plural") -f $sif.value
            }
            else
            {
                $value = Get-LanguageString "AzureCA.SessionLifetime.SignInFrequency.Option.Hour.singular"
            }
        }
        elseif($sif.type -eq "days")
        {
            if($sif.value -gt 1)
            {
                $value = (Get-LanguageString "AzureCA.SessionLifetime.SignInFrequency.Option.Day.plural") -f $sif.value
            }
            else
            {
                $value = Get-LanguageString "AzureCA.SessionLifetime.SignInFrequency.Option.Day.singular"
            }
        }
        else
        {
            $value = Get-LanguageString "AzureCA.SessionControls.SignInFrequency.everytime"
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.SessionLifetime.SignInFrequency.Option.label"
            Value =  $value
            Category = $category
            SubCategory = ""
            EntityKey = "SignInFrequency"
        })

        # Every-time sub-option: apply to secondary authentication methods only
        if($sif.authenticationType -eq "secondaryAuthentication")
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "AzureCA.SessionControls.SignInFrequency.EverytimeCA.secondaryOnly" "Secondary authentication methods only" -IgnoreMissing
                Value = Get-LanguageString "Inputs.enabled"
                Category = $category
                SubCategory = ""
                EntityKey = "signInFrequencyAuthenticationType"
            })
        }
    }

    # Persistent browser session
    if($obj.sessionControls.persistentBrowser.isEnabled -eq $true)
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.SessionLifetime.PersistentBrowser.Option.label"
            Value =  Get-LanguageString "AzureCA.SessionLifetime.PersistentBrowser.Option.$($obj.sessionControls.persistentBrowser.mode)"
            Category = $category
            SubCategory = ""
            EntityKey = "persistentBrowser"
        })
    }

    # Customize continued access evaluation (CAE). Gate on .mode: the Graph default
    # object is {mode:null} and the portal treats that as not configured.
    if($obj.sessionControls.continuousAccessEvaluation.mode)
    {
        $value = switch ($obj.sessionControls.continuousAccessEvaluation.mode)
        {
            "strictLocation"    { Get-LanguageString "AzureCA.SessionControls.Cae.strictLocation" }
            "strictEnforcement" { Get-LanguageString "AzureCA.SessionControls.Cae.strictEnforcement" }
            default             { Get-LanguageString "AzureCA.SessionControls.Cae.disable" }
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.SessionControls.Cae.checkboxLabel"
            Value =  $value
            Category = $category
            SubCategory = ""
            EntityKey = "continuousAccessEvaluation"
        })
    }

    # Disable resilience defaults
    if ($obj.sessionControls.disableResilienceDefaults -eq $true)
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.SessionControls.ResiliencyDefaults.checkboxLabel"
            Value = Get-LanguageString "Inputs.enabled"
            Category = $category
            SubCategory = ""
            EntityKey = "disableResilienceDefaults"
        })
    }

    # Require token protection for sign-in sessions
    if ($obj.sessionControls.secureSignInSession.isEnabled -eq $true)
    {
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.SessionControls.SecureSignIn.checkboxLabel"
            Value = Get-LanguageString "Inputs.enabled"
            Category = $category
            SubCategory = ""
            EntityKey = "secureSignInSession"
        })

        # Sub-option: require token protection for app sessions
        if ($obj.sessionControls.secureSignInSession.secureAppSessionMode -eq 'enforced')
        {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = Get-LanguageString "AzureCA.SessionControls.SecureApp.checkboxLabel"
                Value = Get-LanguageString "Inputs.enabled"
                Category = $category
                SubCategory = ""
                EntityKey = "secureAppSessionMode"
            })
        }
    }

    # Block sensitive actions (agent-context policies only)
    if ($obj.sessionControls.blockSensitiveActions.isEnabled -eq $true)
    {
        # includeSensitiveActions is always ["all"] today; render any future
        # non-all list instead of a bare "Enabled"
        $actions = @($obj.sessionControls.blockSensitiveActions.includeSensitiveActions | Where-Object { $_ -and $_ -ne 'all' })
        $bsaValue = if ($actions.Count -gt 0) { $actions -join (Get-CurrentDocumentationContext).ObjectSeparator }
                    else { Get-LanguageString "Inputs.enabled" }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.SessionControls.BlockSensitiveActions.checkboxLabel"
            Value = $bsaValue
            Category = $category
            SubCategory = ""
            EntityKey = "blockSensitiveActions"
        })
    }

    # Use Global Secure Access security profile; resolve the profile name like the
    # portal's security-profiles dropdown, falling back to the raw id
    if ($obj.sessionControls.globalSecureAccessFilteringProfile.isEnabled -eq $true)
    {
        $gsaValue = $null
        $profileId = $obj.sessionControls.globalSecureAccessFilteringProfile.profileId
        if ($profileId) {
            $profiles = Resolve-CAGsaFilteringProfiles
            $profileObj = $profiles | Where-Object { $_.id -eq $profileId } | Select-Object -First 1
            $gsaValue = ?? $profileObj.name $profileId
        }
        if (-not $gsaValue) { $gsaValue = Get-LanguageString "Inputs.enabled" }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = Get-LanguageString "AzureCA.SessionControls.NetworkAccessSecurity.checkboxLabel"
            Value = $gsaValue
            Category = $category
            SubCategory = ""
            EntityKey = "globalSecureAccessFilteringProfile"
        })
    }
}

# Resolve Global Secure Access security profiles for the session control. The API is
# beta-only and absent in some sovereign clouds, so failures negative-cache and the
# caller falls back to the raw profileId.
function Resolve-CAGsaFilteringProfiles {
    if ($null -ne $script:_caGsaProfiles) { return @($script:_caGsaProfiles) }

    $ctx = Get-CurrentDocumentationContext
    if ($ctx.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) {
        $script:_caGsaProfiles = @()
        return @()
    }

    try {
        # filteringProfile inherits its display property as 'name' (not displayName)
        $resp = Invoke-MSGraphAPI -Url '/networkAccess/filteringProfiles?$select=id,name' -ODataMetadata 'minimal' -NoError
        $script:_caGsaProfiles = @($resp.Value)
    }
    catch {
        Write-LogError 'ConditionalAccess: failed to resolve Global Secure Access security profiles' $_.Exception
        $script:_caGsaProfiles = @()
    }
    return @($script:_caGsaProfiles)
}

# Register at file-load time
[DocumentationRegistry]::RegisterHandler([ConditionalAccessDocHandler]::new())
