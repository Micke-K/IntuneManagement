# Avalonia UI extension for the Word documentation output provider.
# See DocumentationOutputHTMLUIExtension.ps1 for the rationale + pattern.

$provider = [DocumentationRegistry]::FindOutput('word')
if ($provider) {
    Add-Member -InputObject $provider -MemberType NoteProperty -Force -Name FileBrowseControls -Value @(
        @{ Button='browseWordDocumentName';     TextBox='txtWordDocumentName';     Save=$true;  Title='Save Word documentation'; Extension='docx'; FilterName='Word files';     FilterPattern='*.docx' }
        @{ Button='browseWordDocumentTemplate'; TextBox='txtWordDocumentTemplate'; Save=$false; Title='Select Word template';                      FilterName='Word templates'; FilterPattern='*.dotx;*.dotm' }
    )
}
