# Cross-tenant Policy Set import support (architecture rule R9 - new feature
# in a new file; the PolicySet helpers used by Replace-mode import live in
# Internal/IntuneManager.ps1).
#
# A policySet bundles references to other objects via items[].payloadId.
# Importing a set exported from ANOTHER tenant carries source-tenant ids that
# fail the POST. Resolve-IntunePolicySetItems rewrites each item so the set
# imports cleanly:
#   1. payloadId still valid in the target tenant -> keep (same-tenant import)
#   2. else match the item's displayName against the mapped PolicyType's
#      target list (client-side - Intune ignores server-side name filters)
#   3. else DROP the item with a warning - one dangling reference fails the
#      whole set.
# Items are then stripped to the POST-safe shape (the GET carries per-item
# ids, links, status and timestamps that Graph rejects on create) - same rule
# as the original project's Start-PreImportPolicySets.

# policySetItem @odata.type -> owning PolicyType Id(s) for name resolution.
# Unmapped item types keep their payloadId untouched.
$script:PolicySetItemTypeMap = @{
    '#microsoft.graph.mobileAppPolicySetItem'                                      = @('Applications')
    '#microsoft.graph.targetedManagedAppConfigurationPolicySetItem'                = @('AppConfigurationManagedApp')
    '#microsoft.graph.managedAppProtectionPolicySetItem'                           = @('AppProtection')
    '#microsoft.graph.mdmWindowsInformationProtectionPolicyPolicySetItem'          = @('AppProtection')
    '#microsoft.graph.windowsManagedAppProtectionPolicySetItem'                    = @('AppProtection')
    '#microsoft.graph.managedDeviceMobileAppConfigurationPolicySetItem'            = @('AppConfigurationManagedDevice')
    '#microsoft.graph.iosLobAppProvisioningConfigurationPolicySetItem'             = @('IosLobAppProvisioningConfigurations')
    '#microsoft.graph.deviceCompliancePolicyPolicySetItem'                         = @('CompliancePolicies')
    '#microsoft.graph.deviceConfigurationPolicySetItem'                            = @('DeviceConfiguration')
    '#microsoft.graph.deviceManagementConfigurationPolicyPolicySetItem'            = @('SettingsCatalog', 'EndpointSecuritySettingsCatalog')
    '#microsoft.graph.groupPolicyConfigurationPolicySetItem'                       = @('AdministrativeTemplates')
    '#microsoft.graph.deviceManagementScriptPolicySetItem'                         = @('PowerShellScripts')
    '#microsoft.graph.enrollmentRestrictionsConfigurationPolicySetItem'            = @('EnrollmentLimit')
    '#microsoft.graph.windows10EnrollmentCompletionPageConfigurationPolicySetItem' = @('EnrollmentStatusPage')
    '#microsoft.graph.windowsAutopilotDeploymentProfilePolicySetItem'              = @('AutoPilot')
}

# Update-mode import: policySet items are a navigation property Graph refuses
# on PATCH - membership changes go through the dedicated /update action with
# added/updated/deleted item deltas (same flow as the original project's
# Start-PreUpdatePolicySets). The caller's PATCH then carries metadata only
# (_PropertiesToRemoveForUpdate strips 'items').
function Update-IntunePolicySetItems
{
    param($PolicyObject, $ExistingObject, [int]$TokenId)

    if($TokenId -eq 0) { $TokenId = [int](Get-DefaultTokenId) }

    # Same cross-tenant re-pointing rules as create-mode import.
    Resolve-IntunePolicySetItems -PolicyObject $PolicyObject -TokenId $TokenId

    $json = $PolicyObject.JsonObject
    $targetId = [string]$ExistingObject.Id
    if(-not $targetId) { return }

    $current = Invoke-MSGraphAPI -Url "deviceAppManagement/policySets/$($targetId)?`$expand=items" -TokenId $TokenId
    $currentItems = @($current.items)

    $addedItems   = @()
    $updatedItems = @()
    $deletedItems = @()

    foreach($item in @($json.items)) {
        if(@($currentItems | Where-Object { [string]$_.payloadId -eq [string]$item.payloadId }).Count -gt 0) {
            $updatedItems += $item
        }
        else {
            $addedItems += $item
        }
    }

    foreach($currentItem in $currentItems) {
        if(@($json.items | Where-Object { [string]$_.payloadId -eq [string]$currentItem.payloadId }).Count -eq 0) {
            $deletedItems += [string]$currentItem.id
        }
    }

    Write-Log "Policy set '$($json.displayName)': update items. Add: $($addedItems.Count), Update: $($updatedItems.Count), Delete: $($deletedItems.Count)"

    $updateBody = [PSCustomObject]@{
        addedPolicySetItems   = $addedItems
        updatedPolicySetItems = $updatedItems
        deletedPolicySetItems = $deletedItems
    }

    $result = Invoke-MSGraphAPI -Url "deviceAppManagement/policySets/$targetId/update" -HttpMethod "POST" -Content (ConvertTo-Json $updateBody -Depth 15) -TokenId $TokenId -FullResponseObject
    if(-not $result.Success) {
        Write-Log "Policy set '$($json.displayName)': item update action failed - $($result.StatusDescription)" 2
    }
}

function Resolve-IntunePolicySetItems
{
    param($PolicyObject, [int]$TokenId)

    # A file-loaded import source has no token; fall back to the default.
    # (Passing 0 into Get-GraphPolicies resolves ALL registered tokens when
    # more than one provider is connected.)
    if($TokenId -eq 0) { $TokenId = [int](Get-DefaultTokenId) }

    $json = $PolicyObject.JsonObject
    if(-not $json -or -not $json.items) { return }

    Remove-Property $json 'items@odata.context'

    $keepProps = @('@odata.type', 'payloadId', 'intent', 'settings')
    $typeListCache = @{}
    $resolvedItems = [System.Collections.Generic.List[object]]::new()

    foreach($item in @($json.items)) {
        $odata     = [string]$item.'@odata.type'
        $itemName  = [string]$item.displayName
        $payloadId = [string]$item.payloadId

        $typeIds = $script:PolicySetItemTypeMap[$odata]
        if($typeIds) {
            $candidates = @()
            foreach($typeId in $typeIds) {
                if(-not $typeListCache.ContainsKey($typeId)) {
                    $typeListCache[$typeId] = @(Get-GraphPolicies -PolicyType $typeId -TokenId $TokenId)
                }
                $candidates += $typeListCache[$typeId]
            }

            if($payloadId -and @($candidates | Where-Object { [string]$_.Id -eq $payloadId }).Count -gt 0) {
                # Reference still valid in the target tenant - keep as-is.
            }
            else {
                $byName = @($candidates | Where-Object { $itemName -and [string]$_.Name -eq $itemName })
                if($byName.Count -eq 1) {
                    Write-Log "Policy set '$($json.displayName)': item '$itemName' re-pointed from '$payloadId' to '$($byName[0].Id)'"
                    $item.payloadId = [string]$byName[0].Id
                }
                elseif($byName.Count -gt 1) {
                    Write-Log "Policy set '$($json.displayName)': multiple objects named '$itemName' in the target tenant - dropping item (ambiguous reference)" 2
                    continue
                }
                else {
                    Write-Log "Policy set '$($json.displayName)': no object named '$itemName' in the target tenant - dropping item (a dangling reference fails the whole set)" 2
                    continue
                }
            }
        }

        foreach($prop in @($item.PSObject.Properties | Where-Object { $_.Name -notin $keepProps })) {
            Remove-Property $item $prop.Name
        }
        [void]$resolvedItems.Add($item)
    }

    $json.items = $resolvedItems.ToArray()
}
