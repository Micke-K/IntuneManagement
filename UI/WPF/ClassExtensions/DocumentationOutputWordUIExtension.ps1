# WPF UI extension for the Word documentation output provider.
# See DocumentationOutputHTMLUIExtension.ps1 for the rationale + pattern.
#
# Note: WordDocumentName's original filter included a PDF entry because Word
# can save PDFs directly. Preserved here for parity with the pre-refactor
# Add-WpfDocumentationFileBrowse call.

$provider = [DocumentationRegistry]::FindOutput('word')
if ($provider) {
    Add-Member -InputObject $provider -MemberType NoteProperty -Force -Name FileBrowseControls -Value @(
        @{ Button='browseWordDocumentName';     TextBox='txtWordDocumentName';     Save=$true;  Filter='Word files (*.docx)|*.docx|PDF files (*.pdf)|*.pdf|All files (*.*)|*.*' }
        @{ Button='browseWordDocumentTemplate'; TextBox='txtWordDocumentTemplate'; Save=$false; Filter='Word templates (*.dotx;*.dotm)|*.dotx;*.dotm|All files (*.*)|*.*'      }
    )
}
