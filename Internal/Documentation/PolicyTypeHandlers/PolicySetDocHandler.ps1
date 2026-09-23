# Policy Set documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:3422. Categorizes the
# policy-set items (apps / device configs / enrollment) into 3 sections, then
# emits one row per item with the item's displayName and a type-specific
# value (priority number for ordered items, AAD/AD for autopilot, etc.).

class PolicySetDocHandler : DocumentationHandlerBase {
    PolicySetDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.policySet')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        Add-BasicDefaultValues $PolicyObject
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'SettingDetails.appConfiguration') '@odata.type'

        $sections = @(
            [PSCustomObject]@{
                Category = (Get-LanguageString 'PolicySet.appManagement')
                Types = @(
                    @{ ODataType = '#microsoft.graph.mobileAppPolicySetItem'; SubKey = 'appTitle' }
                    @{ ODataType = '#microsoft.graph.targetedManagedAppConfigurationPolicySetItem'; SubKey = 'appConfigurationTitle' }
                    @{ ODataType = '#microsoft.graph.managedDeviceMobileAppConfigurationPolicySetItem'; SubKey = 'appConfigurationTitle' }
                    @{ ODataType = '#microsoft.graph.managedAppProtectionPolicySetItem'; SubKey = 'appProtectionTitle' }
                    @{ ODataType = '#microsoft.graph.iosLobAppProvisioningConfigurationPolicySetItem'; SubKey = 'iOSAppProvisioningTitle' }
                )
            }
            [PSCustomObject]@{
                Category = (Get-LanguageString 'PolicySet.deviceManagement')
                Types = @(
                    @{ ODataType = '#microsoft.graph.deviceConfigurationPolicySetItem'; SubKey = 'deviceConfigurationTitle' }
                    @{ ODataType = '#microsoft.graph.deviceManagementConfigurationPolicyPolicySetItem'; SubKey = 'SettingDetails.settingsCatalog' }
                    @{ ODataType = '#microsoft.graph.deviceCompliancePolicyPolicySetItem'; SubKey = 'deviceComplianceTitle' }
                    @{ ODataType = '#microsoft.graph.deviceManagementScriptPolicySetItem'; SubKey = 'powershellScriptTitle' }
                )
            }
            [PSCustomObject]@{
                Category = (Get-LanguageString 'PolicySet.deviceEnrollment')
                Types = @(
                    @{ ODataType = '#microsoft.graph.enrollmentRestrictionsConfigurationPolicySetItem'; SubKey = 'deviceTypeRestrictionTitle' }
                    @{ ODataType = '#microsoft.graph.windowsAutopilotDeploymentProfilePolicySetItem'; SubKey = 'windowsAutopilotDeploymentProfileTitle' }
                    @{ ODataType = '#microsoft.graph.windows10EnrollmentCompletionPageConfigurationPolicySetItem'; SubKey = 'enrollmentStatusSettingTitle' }
                )
            }
        )

        foreach ($section in $sections) {
            foreach ($subType in $section.Types) {
                foreach ($item in ($obj.items | Where-Object { $_.'@OData.Type' -eq $subType.ODataType -or $_.'@odata.type' -eq $subType.ODataType })) {
                    if ($item.status -eq 'error') {
                        Write-Log "Skipping missing $($subType.ODataType) type with id $($item.id). Error code: $($item.errorCode)" 2
                        continue
                    }

                    # SubKey is a bare key under PolicySet.* unless it already
                    # carries a namespace (dotted), letting new item types reuse
                    # strings that live outside the PolicySet section.
                    $subKeyFull = if ($subType.SubKey -like '*.*') { $subType.SubKey } else { "PolicySet.$($subType.SubKey)" }

                    Add-CustomSettingObject ([PSCustomObject]@{
                        Name        = $item.displayName
                        Value       = (Get-PolicySetItemValue $item)
                        EntityKey   = $item.id
                        Category    = $section.Category
                        SubCategory = (Get-LanguageString $subKeyFull)
                    })
                }
            }
        }
    }
}

function Get-PolicySetItemValue {
    param($item)

    $odata = if ($item.PSObject.Properties['@OData.Type']) { $item.'@OData.Type' } else { $item.'@odata.type' }

    if ($odata -in @(
        '#microsoft.graph.enrollmentRestrictionsConfigurationPolicySetItem',
        '#microsoft.graph.windows10EnrollmentCompletionPageConfigurationPolicySetItem'
    )) {
        return $item.Priority
    }

    if ($odata -eq '#microsoft.graph.windowsAutopilotDeploymentProfilePolicySetItem') {
        if ($item.itemType -eq '#microsoft.graph.azureADWindowsAutopilotDeploymentProfile') {
            return (Get-LanguageString 'Autopilot.DirectoryService.azureAD')
        }
        if ($item.itemType -eq '#microsoft.graph.activeDirectoryWindowsAutopilotDeploymentProfile') {
            return (Get-LanguageString 'Autopilot.DirectoryService.activeDirectoryAD')
        }
    }

    # TODO phase-4-followup: other PolicySet item types as fixtures arrive
    return $null
}

[DocumentationRegistry]::RegisterHandler([PolicySetDocHandler]::new())
