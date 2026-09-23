# Co-Management Settings documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:3862. Hardcoded
# Platform = Windows 10 (Co-Management is Windows-only).

class CoManagementDocHandler : DocumentationHandlerBase {
    CoManagementDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.deviceComanagementAuthorityConfiguration')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        Add-BasicDefaultValues $PolicyObject
        Add-BasicAdditionalValues $PolicyObject
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') ((Get-LanguageString 'WindowsEnrollment.coManagementAuthorityTitle').Trim()) '@odata.type'
        Add-BasicPropertyValue (Get-LanguageString 'Inputs.platformLabel') (Get-LanguageString 'Platform.Windows10') 'platform'

        $category = Get-LanguageString 'TableHeaders.settings'
        $yes = Get-LanguageString 'BooleanActions.yes'
        $no  = Get-LanguageString 'SettingDetails.no'

        $installValue = if ($obj.installConfigurationManagerAgent -eq $true) { $yes } else { $no }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString 'CoManagementAuthority.installAgent')
            Value = $installValue
            EntityKey = 'installConfigurationManagerAgent'
            Category = $category
            SubCategory = $null
        })

        if ($obj.installConfigurationManagerAgent -eq $true) {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'CoManagementAuthority.commandLineArgs')
                Value = $obj.configurationManagerAgentCommandLineArgument
                EntityKey = 'configurationManagerAgentCommandLineArgument'
                Category = $category
                SubCategory = $null
            })
        }

        $ownershipValue = if ($obj.managedDeviceAuthority -eq 1) { $yes } else { $no }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString 'CoManagementAuthority.managedDeviceOwnership')
            Value = $ownershipValue
            EntityKey = 'managedDeviceAuthority'
            Category = $category
            SubCategory = (Get-LanguageString 'CoManagementAuthority.advancedProperty')
        })
    }
}

[DocumentationRegistry]::RegisterHandler([CoManagementDocHandler]::new())
