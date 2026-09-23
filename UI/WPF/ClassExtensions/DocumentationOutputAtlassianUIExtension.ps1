# WPF UI extension for the Atlassian documentation output provider.
#
# Attaches file-browse metadata onto the registered [DocumentationRegistry]
# entry so the bulk-doc form's browse-button-wiring loop can build a WinForms
# SaveFileDialog handler. Format matches Add-WpfDocumentationFileBrowse
# (Button/TextBox/Save/Filter). The provider file itself
# (Internal/Documentation/OutputProviders/DocumentationOutputAtlassian.ps1)
# stays free of XAML control names and Windows-Forms filter strings — only the
# active UI backend's ClassExtensions load per session, so the Avalonia
# counterpart won't conflict.
#
# Unlike HTML there is no CSS row: Confluence storage format carries no
# stylesheet, so the panel has a single browse control.

$provider = [DocumentationRegistry]::FindOutput('atlassian')
if ($provider) {
    Add-Member -InputObject $provider -MemberType NoteProperty -Force -Name FileBrowseControls -Value @(
        @{ Button='browseAtlassianDocumentName'; TextBox='txtAtlassianDocumentName'; Save=$true; Filter='Confluence storage format (*.html)|*.html|All files (*.*)|*.*' }
    )
}
