# Custom compliance script documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:4788
# (Invoke-CDDocumentDeviceComplianceScript). Claims
# @odata.type='#microsoft.graph.deviceComplianceScript'.
#
# deviceComplianceScript has no ObjectCategories entry, so - like the Scope
# Tag handler - basic info rows are emitted manually instead of via
# Add-BasicDefaultValues (which would add blank Platform/Profile rows).

class ComplianceScriptDocHandler : DocumentationHandlerBase {
    ComplianceScriptDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.deviceComplianceScript')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        if ($PolicyObject.Name) {
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.nameName') $PolicyObject.Name 'displayName'
        }
        $descValue = if ($obj.description) { $obj.description } else { '' }
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.descriptionName') $descValue 'description'
        if ($obj.publisher) {
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.publisher') $obj.publisher 'publisher'
        }
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'Titles.complianceScriptManagementPreview') 'configurationType'

        $category = Get-LanguageString 'TableHeaders.settings'
        $valueYes = Get-LanguageString 'BooleanActions.yes'
        $valueNo  = Get-LanguageString 'SettingDetails.no'

        if ($obj.detectionScriptContent -and -not ($Context.Options -and $Context.Options['IncludeScripts'] -eq $false)) {
            $scriptBody = ''
            try { $scriptBody = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.detectionScriptContent)) } catch { }
            if ($scriptBody) {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name        = Get-LanguageString 'ProactiveRemediations.Create.Settings.DetectionScriptMultiLineTextBox.label'
                    Value       = $scriptBody
                    EntityKey   = 'detectionScriptContent'
                    Category    = $category
                    SubCategory = $null
                })
            }
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name        = Get-LanguageString 'WindowsManagement.scriptContextLabel'
            Value       = $(if ($obj.runAsAccount -eq 'system') { $valueNo } else { $valueYes })
            EntityKey   = 'runAsAccount'
            Category    = $category
            SubCategory = $null
        })

        Add-CustomSettingObject ([PSCustomObject]@{
            Name        = Get-LanguageString 'WindowsManagement.enforceSignatureCheckLabel'
            Value       = $(if ($obj.enforceSignatureCheck -eq $false) { $valueNo } else { $valueYes })
            EntityKey   = 'enforceSignatureCheck'
            Category    = $category
            SubCategory = $null
        })

        Add-CustomSettingObject ([PSCustomObject]@{
            Name        = Get-LanguageString 'WindowsManagement.runAs64BitLabel'
            Value       = $(if ($obj.runAs32Bit -eq $true) { $valueNo } else { $valueYes })
            EntityKey   = 'runAs32Bit'
            Category    = $category
            SubCategory = $null
        })
    }
}

[DocumentationRegistry]::RegisterHandler([ComplianceScriptDocHandler]::new())
