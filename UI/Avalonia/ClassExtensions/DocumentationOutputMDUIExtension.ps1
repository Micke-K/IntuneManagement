# Avalonia UI extension for the Markdown documentation output provider.
# See DocumentationOutputHTMLUIExtension.ps1 for the rationale + pattern.

$provider = [DocumentationRegistry]::FindOutput('md')
if ($provider) {
    Add-Member -InputObject $provider -MemberType NoteProperty -Force -Name FileBrowseControls -Value @(
        @{ Button='browseMDDocumentName'; TextBox='txtMDDocumentName'; Save=$true;  Title='Save Markdown documentation'; Extension='md'; FilterName='Markdown files'; FilterPattern='*.md'  }
        @{ Button='browseMDCSSFile';      TextBox='txtMDCSSFile';      Save=$false; Title='Select Markdown CSS file';                    FilterName='CSS files';      FilterPattern='*.css' }
    )
}
