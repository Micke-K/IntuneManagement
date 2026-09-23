# Role Definition documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:4513. Resolves
# permissions to resource/action names via /deviceManagement/resourceOperations
# (generic schema - resolved from any connected tenant) and enriches assignments
# with directory display names (source-tenant-specific - skipped when the source
# tenant is unavailable, emitting raw IDs).

class RoleDefinitionDocHandler : DocumentationHandlerBase {
    RoleDefinitionDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.deviceAndAppManagementRoleDefinition')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        Add-BasicDefaultValues $PolicyObject
        Add-BasicAdditionalValues $PolicyObject
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'RoleAssignment.rolesMenuTitle') '@odata.type'

        # Built-in vs custom role. isBuiltIn is true for Microsoft-supplied roles.
        if ($null -ne $obj.isBuiltIn) {
            $builtInValue = if ($obj.isBuiltIn) { Get-LanguageString 'SettingDetails.yes' } else { Get-LanguageString 'SettingDetails.no' }
            Add-BasicPropertyValue (Get-LanguageString 'DisplayRoleTypes.builtInRole') $builtInValue 'isBuiltIn'
        }

        # --- Permissions section: resolve action IDs to resource/action names ---
        $roleResources = @()
        # resourceOperations is a GENERIC catalog (resource/action names) - same on
        # every tenant - so resolved from any connected tenant.
        if (Test-DocumentationGraphAvailable) {
            try {
                $resp = Invoke-MSGraphAPI -Url '/deviceManagement/resourceOperations'
                $roleResources = @($resp.Value)
            }
            catch {
                Write-LogError 'Failed to fetch /deviceManagement/resourceOperations for role permissions' $_.Exception
            }
        }

        $permissionsCategory = Get-LanguageString 'Titles.permissions'

        # Prefer the modern rolePermissions structure: union allowedResourceActions
        # across ALL rolePermissions/resourceActions. Fall back to the legacy flat
        # permissions[0].actions list when rolePermissions is absent (older payloads
        # and some built-in roles only populate the legacy list).
        $actionIds = @()
        if ($obj.rolePermissions) {
            foreach ($rolePermission in @($obj.rolePermissions)) {
                foreach ($resourceAction in @($rolePermission.resourceActions)) {
                    foreach ($allowed in @($resourceAction.allowedResourceActions)) {
                        if ($allowed -and $actionIds -notcontains $allowed) { $actionIds += $allowed }
                    }
                }
            }
        }
        if ($actionIds.Count -eq 0 -and $obj.permissions -and $obj.permissions[0]) {
            $actionIds = @($obj.permissions[0].actions)
        }

        if ($roleResources.Count -gt 0 -and $actionIds.Count -gt 0) {
            # Resolved live: group resolved actions by resourceName
            $assignedActions = @()
            foreach ($id in $actionIds) {
                $r = $roleResources | Where-Object Id -EQ $id | Select-Object -First 1
                if ($r) { $assignedActions += $r }
            }

            $byResource = $assignedActions | Select-Object resourceName -Unique | Sort-Object -Property resourceName
            foreach ($rn in $byResource.resourceName) {
                $actions = @($assignedActions | Where-Object resourceName -EQ $rn)
                $resourceId = $actions[0].resource
                $actionNames = ($actions | ForEach-Object { $_.actionName })

                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = $rn
                    Value = ($actionNames -join $Context.ObjectSeparator)
                    EntityKey = $resourceId
                    Category = $permissionsCategory
                    SubCategory = $null
                })
            }
        }
        elseif ($actionIds.Count -gt 0) {
            # Offline: emit a single row with raw action IDs so the row count
            # is non-zero and compare-style downstream tools have something
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'Titles.permissions')
                Value = ($actionIds -join $Context.ObjectSeparator)
                EntityKey = 'actions'
                Category = $permissionsCategory
                SubCategory = $null
            })
        }

        # --- Assignments section ---
        # roleAssignments are enriched in place (full assignment + roleScopeTags)
        # by RoleDefinitionObject's sub-resource contract during hydration, so
        # each entry already carries displayName/description/members/scopeMembers/
        # roleScopeTags — no per-assignment Graph fetch here. The
        # deviceManagement/roleAssignments/<id> API now lives only on the class.
        # Offline runs still skip this section (matching prior behavior); the
        # member display-name resolution (getByIds) is shared reference data and
        # stays online-only.
        if ($Context.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) { return }

        $assignmentsCategory = Get-LanguageString 'TableHeaders.assignments'
        foreach ($info in @($obj.roleAssignments)) {
            if (-not $info -or [string]::IsNullOrWhiteSpace([string]$info.id)) {
                Write-Log 'RoleDefinition: skipping role assignment without an id' 2
                continue
            }

            # Resolve member + scope IDs to displayNames in one batch
            $ids = @()
            foreach ($id in @($info.members + $info.scopeMembers)) {
                if ($id -and $ids -notcontains $id) { $ids += $id }
            }
            $idInfo = @()
            if ($ids.Count -gt 0) {
                try {
                    $body = @{ ids = $ids } | ConvertTo-Json
                    $resp = Invoke-MSGraphAPI -Url "/directoryObjects/getByIds?`$select=displayName,id" -Content $body -Method POST
                    $idInfo = @($resp.Value)
                }
                catch { Write-LogError 'Failed to resolve role-assignment member display names' $_.Exception }
            }

            $sub = $info.displayName
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'SettingDetails.nameName')
                Value = $info.displayName
                EntityKey = 'displayName'
                Category = $assignmentsCategory; SubCategory = $sub
            })
            if ($info.description) {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = (Get-LanguageString 'SettingDetails.descriptionName')
                    Value = $info.description
                    EntityKey = 'description'
                    Category = $assignmentsCategory; SubCategory = $sub
                })
            }

            $admins = @()
            foreach ($id in @($info.members)) {
                $resolved = $idInfo | Where-Object Id -EQ $id | Select-Object -First 1
                $admins += if ($resolved.displayName) { $resolved.displayName } else { $id }
            }
            if ($admins.Count -gt 0) {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = (Get-LanguageString 'RoleAssignment.RoleAssignmentAdmin')
                    Value = ($admins -join $Context.ObjectSeparator)
                    EntityKey = 'members'
                    Category = $assignmentsCategory; SubCategory = $sub
                })
            }

            $scopeMembers = @()
            foreach ($id in @($info.scopeMembers)) {
                $resolved = $idInfo | Where-Object Id -EQ $id | Select-Object -First 1
                $scopeMembers += if ($resolved.displayName) { $resolved.displayName } else { $id }
            }
            if ($scopeMembers.Count -gt 0) {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = (Get-LanguageString 'RoleAssignment.RoleAssignmentScope')
                    Value = ($scopeMembers -join $Context.ObjectSeparator)
                    EntityKey = 'scopeMembers'
                    Category = $assignmentsCategory; SubCategory = $sub
                })
            }

            $scopeTags = @($info.roleScopeTags | ForEach-Object { $_.displayName })
            if ($scopeTags.Count -gt 0) {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = (Get-LanguageString 'TableHeaders.scopeTags')
                    Value = ($scopeTags -join $Context.ObjectSeparator)
                    EntityKey = 'scopeTags'
                    Category = $assignmentsCategory; SubCategory = $sub
                })
            }
        }
    }
}

[DocumentationRegistry]::RegisterHandler([RoleDefinitionDocHandler]::new())
