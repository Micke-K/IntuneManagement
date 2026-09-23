# Assignment Filter documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:3823 (Invoke-CDDocument-
# AssignmentFilter, ~35 LOC). Claims @odata.type='#microsoft.graph.deviceAnd
# AppManagementAssignmentFilter' and produces a 7-row BasicInfo + 1-row
# Settings table (the rule syntax).
#
# Platform value: app-management platforms (androidMobileApplicationManagement
# etc.) resolve to empty strings in Strings-en.json which Get-LanguageString
# returns as $null — so BasicInfo emits the row with Value=null, matching the
# golden's `"Platform": null` for app filters.

class AssignmentFilterDocHandler : DocumentationHandlerBase {
    AssignmentFilterDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.deviceAndAppManagementAssignmentFilter')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        # BasicInfo order: Name, Description, Created, Last modified, Profile type, Platform
        # (Scope tags appended automatically by the engine's post-step.)
        Add-BasicDefaultValues $PolicyObject
        Add-BasicAdditionalValues $PolicyObject
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'Filters.filters') '@odata.type'
        Add-BasicPropertyValue (Get-LanguageString 'Inputs.platformLabel') (Get-LanguageString "Platform.$($obj.platform)") 'platform'

        # Filter scope: devices vs apps. Disambiguates app-management filters, whose
        # platform row resolves to null (see header note).
        $mgmtType = switch ($obj.assignmentFilterManagementType) {
            'devices' { Get-LanguageString 'Titles.devices' }
            'apps'    { Get-LanguageString 'Titles.apps' }
            default   { $obj.assignmentFilterManagementType }
        }
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.managementType') $mgmtType 'assignmentFilterManagementType'

        # Settings: a single Rule syntax row
        Add-CustomSettingObject ([PSCustomObject]@{
            Name        = Get-LanguageString 'Filters.ruleSyntax'
            Value       = $obj.rule
            EntityKey   = 'rule'
            Category    = Get-LanguageString 'SettingDetails.rules'
            SubCategory = $null
        })
    }
}

[DocumentationRegistry]::RegisterHandler([AssignmentFilterDocHandler]::new())
