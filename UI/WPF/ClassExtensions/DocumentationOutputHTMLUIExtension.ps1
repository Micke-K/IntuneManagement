# WPF UI extension for the HTML documentation output provider.
#
# Attaches file-browse metadata onto the registered [DocumentationRegistry]
# entry so the bulk-doc form's browse-button-wiring loop can build WinForms
# SaveFileDialog / OpenFileDialog handlers. Format matches Add-WpfDocumentationFileBrowse
# (Button/TextBox/Save/Filter). The provider file itself
# (Internal/Documentation/OutputProviders/DocumentationOutputHTML.ps1) stays
# free of XAML control names and Windows-Forms filter strings — only the
# active UI backend's ClassExtensions load per session, so the Avalonia
# counterpart won't conflict.

$provider = [DocumentationRegistry]::FindOutput('html')
if ($provider) {
    Add-Member -InputObject $provider -MemberType NoteProperty -Force -Name FileBrowseControls -Value @(
        @{ Button='browseHTMLDocumentName'; TextBox='txtHTMLDocumentName'; Save=$true;  Filter='HTML files (*.html)|*.html|All files (*.*)|*.*' }
        @{ Button='browseHTMLCSSFile';      TextBox='txtHTMLCSSFile';      Save=$false; Filter='CSS files (*.css)|*.css|All files (*.*)|*.*'    }
    )
}
