# Reusable setting (deviceManagementReusablePolicySetting) documentation
# handler.
#
# Claims @odata.type='#microsoft.graph.deviceManagementReusablePolicySetting'.
# Covers the reusable settings surfaced as policy types (currently the Linux
# custom-compliance discovery script). The object is a name + description +
# settingDefinitionId wrapper around a single settings-catalog setting
# instance whose simpleSettingValue carries the payload (base64 script for
# the discovery-script definition). No old-project documenter existed.

class ReusableSettingDocHandler : DocumentationHandlerBase {
    ReusableSettingDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.deviceManagementReusablePolicySetting')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        if ($PolicyObject.Name) {
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.nameName') $PolicyObject.Name 'displayName'
        }
        $descValue = if ($obj.description) { $obj.description } else { '' }
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.descriptionName') $descValue 'description'

        $category = Get-LanguageString 'TableHeaders.settings'

        if ($obj.settingDefinitionId) {
            $definitionLabel = Get-LanguageString 'SettingDetails.settingIdName' -IgnoreMissing
            if (-not $definitionLabel) { $definitionLabel = 'Setting definition' }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name        = $definitionLabel
                Value       = [string]$obj.settingDefinitionId
                EntityKey   = 'settingDefinitionId'
                Category    = $category
                SubCategory = $null
            })
        }

        $rawValue = [string]$obj.settingInstance.simpleSettingValue.value
        if ($rawValue -and -not ($Context.Options -and $Context.Options['IncludeScripts'] -eq $false)) {
            # The discovery-script definition stores the script base64-encoded;
            # fall back to the raw value for definitions that don't.
            $value = $rawValue
            try {
                $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($rawValue))
                if ($decoded) { $value = $decoded }
            }
            catch { }

            $valueLabel = if ([string]$obj.settingDefinitionId -like '*discoveryscript*') {
                Get-LanguageString 'ProactiveRemediations.Create.Settings.DetectionScriptMultiLineTextBox.label'
            }
            else {
                $lbl = Get-LanguageString 'SettingDetails.valueName' -IgnoreMissing
                if ($lbl) { $lbl } else { 'Value' }
            }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name        = $valueLabel
                Value       = $value
                EntityKey   = 'settingInstanceValue'
                Category    = $category
                SubCategory = $null
            })
        }
    }
}

[DocumentationRegistry]::RegisterHandler([ReusableSettingDocHandler]::new())
