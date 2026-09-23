# Avalonia UI extension for the Atlassian documentation output provider.
#
# Attaches file-browse metadata onto the registered [DocumentationRegistry]
# entry so the bulk-doc form's browse-button-wiring loop can render a Save
# picker for the Atlassian output file. The provider file itself
# (Internal/Documentation/OutputProviders/DocumentationOutputAtlassian.ps1)
# stays free of XAML control names and toolkit-specific dialog options —
# only the active UI backend's ClassExtensions load per session, so no
# WPF/Avalonia conflict is possible.
#
# Unlike HTML there is no CSS row: Confluence storage format carries no
# stylesheet, so the panel has a single browse control.

$provider = [DocumentationRegistry]::FindOutput('atlassian')
if ($provider) {
    Add-Member -InputObject $provider -MemberType NoteProperty -Force -Name FileBrowseControls -Value @(
        @{ Button='browseAtlassianDocumentName'; TextBox='txtAtlassianDocumentName'; Save=$true; Title='Save Atlassian documentation'; Extension='html'; FilterName='Confluence storage format'; FilterPattern='*.html' }
    )
}
