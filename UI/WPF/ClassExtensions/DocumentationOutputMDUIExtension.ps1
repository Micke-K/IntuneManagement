# WPF UI extension for the Markdown documentation output provider.
# See DocumentationOutputHTMLUIExtension.ps1 for the rationale + pattern.

$provider = [DocumentationRegistry]::FindOutput('md')
if ($provider) {
    Add-Member -InputObject $provider -MemberType NoteProperty -Force -Name FileBrowseControls -Value @(
        @{ Button='browseMDDocumentName'; TextBox='txtMDDocumentName'; Save=$true;  Filter='Markdown files (*.md)|*.md|All files (*.*)|*.*' }
        @{ Button='browseMDCSSFile';      TextBox='txtMDCSSFile';      Save=$false; Filter='CSS files (*.css)|*.css|All files (*.*)|*.*'     }
    )
}
