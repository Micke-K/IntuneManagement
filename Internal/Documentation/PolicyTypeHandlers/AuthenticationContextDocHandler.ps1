# Authentication Context documentation handler.
#
# Claims @odata.type='#microsoft.graph.authenticationContextClassReference'. Auth
# contexts (c1..c99) were previously only referenced by Conditional Access policies
# (ID -> displayName, see ConditionalAccessDocHandler); this handler documents the
# standalone object: display name, description and whether it is published to apps
# (isAvailable).
#
# Note: AuthenticationContextType strips @odata.type on export (_PropertiesToRemove),
# so this handler matches live documentation runs. File-based runs of an exported
# auth context lose the discriminator and fall through to NoProvider - a pre-existing
# export-cleanup limitation, not addressed here.

class AuthenticationContextDocHandler : DocumentationHandlerBase {
    AuthenticationContextDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.authenticationContextClassReference')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        # No plain-noun language string exists for the auth-context type, so use the
        # PolicyType title for the Profile type row.
        $nameValue = if ($obj.displayName) { $obj.displayName } else { $PolicyObject.Name }
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.nameName') $nameValue 'displayName'
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') $PolicyObject.PolicyType.Title '@odata.type'
        if ($obj.description) {
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.descriptionName') $obj.description 'description'
        }

        if ($null -ne $obj.isAvailable) {
            $availKey = if ($obj.isAvailable -eq $true) { 'Inputs.enabled' } else { 'Inputs.disabled' }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name        = (Get-LanguageString 'AzureCA.AuthContext.InfoBlade.publishLabel')
                Value       = (Get-LanguageString $availKey)
                Category    = $null
                SubCategory = $null
                EntityKey   = 'isAvailable'
            })
        }
    }
}

[DocumentationRegistry]::RegisterHandler([AuthenticationContextDocHandler]::new())
