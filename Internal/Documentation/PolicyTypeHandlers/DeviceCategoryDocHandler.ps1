# Device Category documentation handler.
#
# Claims @odata.type='#microsoft.graph.deviceCategory'. Device categories are
# name + description only; there is no ObjectCategories entry (so no
# Add-BasicDefaultValues - it would emit blank Platform/Profile rows) and no
# old-project documenter to port.

class DeviceCategoryDocHandler : DocumentationHandlerBase {
    DeviceCategoryDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.deviceCategory')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        if ($PolicyObject.Name) {
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.nameName') $PolicyObject.Name 'displayName'
        }
        $descValue = if ($obj.description) { $obj.description } else { '' }
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.descriptionName') $descValue 'description'
    }
}

[DocumentationRegistry]::RegisterHandler([DeviceCategoryDocHandler]::new())
