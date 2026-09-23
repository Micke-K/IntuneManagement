# Avalonia UI extension for the HTML documentation output provider.
#
# Attaches file-browse metadata onto the registered [DocumentationRegistry]
# entry so the bulk-doc form's browse-button-wiring loop can render Save
# and Open pickers for the HTML output file and its CSS reference. The
# provider file itself (Internal/Documentation/OutputProviders/DocumentationOutputHTML.ps1)
# stays free of XAML control names and toolkit-specific dialog options —
# only the active UI backend's ClassExtensions load per session, so no
# WPF/Avalonia conflict is possible.

$provider = [DocumentationRegistry]::FindOutput('html')
if ($provider) {
    Add-Member -InputObject $provider -MemberType NoteProperty -Force -Name FileBrowseControls -Value @(
        @{ Button='browseHTMLDocumentName'; TextBox='txtHTMLDocumentName'; Save=$true;  Title='Save HTML documentation'; Extension='html'; FilterName='HTML files'; FilterPattern='*.html' }
        @{ Button='browseHTMLCSSFile';      TextBox='txtHTMLCSSFile';      Save=$false; Title='Select HTML CSS file';                       FilterName='CSS files';  FilterPattern='*.css'  }
    )
}
