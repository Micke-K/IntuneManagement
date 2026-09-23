# Role Scope Tag documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:4882 (Invoke-
# CDDocumentScopeTag). Claims @odata.type='#microsoft.graph.roleScopeTag'.
#
# Old handler accumulated per-tag rows into a cross-batch
# $script:ObjectTypeFullTable hashtable that the output providers flushed as a
# single consolidated "Scope Tags" table at PostProcess time. The new engine
# already has $ctx.ObjectTypeFullTable for the same purpose, but no output
# provider consumes it yet — until that's wired up, we just emit per-object
# BasicInfo so each tag at least documents independently rather than being
# dropped on the floor as NoProvider.
#
# Assignments are intentionally NOT translated here: Invoke-TranslateAssignments
# is a separate ~370-LOC port (DocumentationMigration.md risk #4) that no
# handler in the new project calls yet. When that lands, this handler can
# trivially call it after the BasicInfo block.

class ScopeTagDocHandler : DocumentationHandlerBase {
    ScopeTagDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.roleScopeTag')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        # Plain Name + Description rows. Scope tags don't have Platform supported
        # or Profile type rows in any old-engine path, so skip the
        # Add-BasicDefaultValues helper (which would emit blank Platform /
        # Profile rows from a missing ObjectCategories entry).
        if ($PolicyObject.Name) {
            $nameProp = if ($PolicyObject.PolicyType._NameProperty) { $PolicyObject.PolicyType._NameProperty } else { 'displayName' }
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.nameName') $PolicyObject.Name $nameProp
        }

        $descValue = if ($obj.description) { $obj.description } else { '' }
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.descriptionName') $descValue 'description'

        if ($null -ne $obj.isBuiltIn) {
            $val = if ($obj.isBuiltIn) { Get-LanguageString 'SettingDetails.yes' } else { Get-LanguageString 'SettingDetails.no' }
            Add-BasicPropertyValue (Get-LanguageString 'RoleScopeTag.isBuiltIn') $val 'isBuiltIn'
        }
    }
}

[DocumentationRegistry]::RegisterHandler([ScopeTagDocHandler]::new())
