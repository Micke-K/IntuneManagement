# Render parsed Markdown into a WPF TextBlock.
#
# The parse lives in Internal/MarkdownSimple.ps1 and is shared with Avalonia
# (R10); this file only turns blocks into WPF inlines, so the toolkit types stay
# inside UI/WPF (R12). The Avalonia counterpart is
# UI/Avalonia/Extensions/MarkdownRenderAvalonia.ps1 and is deliberately a
# near-mirror of this - same block handling, different inline construction.
#
# Everything goes into a single TextBlock rather than a document/panel per block:
# Run exposes FontSize, FontWeight, FontFamily and Background, so headings, bold
# spans, code and indented bullets all fit in one control, and the existing
# ScrollViewer keeps working unchanged.

function Set-WpfMarkdownText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $TextBlock,
        [string]$Markdown
    )

    if (-not $TextBlock) { return }

    $TextBlock.Inlines.Clear()
    if ([string]::IsNullOrEmpty($Markdown)) { return }

    $baseSize = if ($TextBlock.FontSize -gt 0) { [double]$TextBlock.FontSize } else { 12.0 }
    $style    = Get-SimpleMarkdownStyle
    $codeBrush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($style.CodeBackground))
    $codeFont  = New-Object System.Windows.Media.FontFamily ($style.CodeFontFamily)

    foreach ($block in (ConvertFrom-SimpleMarkdown $Markdown)) {

        if ($block.Kind -eq 'Blank') {
            $TextBlock.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
            continue
        }

        # Lead-in for list items and quotes. Leading whitespace in a Run is
        # preserved, so indentation works without a panel per level.
        $lead = switch ($block.Kind) {
            'ListItem' { (' ' * ($block.Indent * 4)) + $block.Marker + ' ' }
            'Quote'    { $style.QuoteMarker }
            default    { $null }
        }
        if ($lead) {
            $TextBlock.Inlines.Add((New-Object System.Windows.Documents.Run -ArgumentList $lead))
        }

        foreach ($inline in $block.Inlines) {
            if ($inline.LineBreak) {
                $TextBlock.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
                # Keep continuation lines of a bullet aligned under its text.
                if ($block.Kind -eq 'ListItem') {
                    $pad = ' ' * (($block.Indent * 4) + $block.Marker.Length + 1)
                    $TextBlock.Inlines.Add((New-Object System.Windows.Documents.Run -ArgumentList $pad))
                }
                continue
            }

            # Links become real Hyperlinks; RequestNavigate opens the default
            # browser, matching how the About and Welcome dialogs handle theirs.
            if ($inline.Url) {
                $link = New-Object System.Windows.Documents.Hyperlink
                $link.Inlines.Add((New-Object System.Windows.Documents.Run -ArgumentList ([string]$inline.Text)))
                try { $link.NavigateUri = New-Object System.Uri ($inline.Url) } catch { }
                $link.Add_RequestNavigate({ Open-ExternalUri $_.Uri.AbsoluteUri; $_.Handled = $true })
                if ($inline.Bold) { $link.FontWeight = [System.Windows.FontWeights]::Bold }
                $TextBlock.Inlines.Add($link)
                continue
            }

            $run = New-Object System.Windows.Documents.Run -ArgumentList ([string]$inline.Text)
            if ($inline.Bold -or $block.Kind -eq 'Heading') {
                $run.FontWeight = [System.Windows.FontWeights]::Bold
            }
            if ($block.Kind -eq 'Heading') {
                $run.FontSize = Get-SimpleMarkdownHeadingSize -Level $block.Level -BaseSize $baseSize
            }
            if ($inline.Code -or $block.Kind -eq 'CodeBlock') {
                $run.FontFamily = $codeFont
                $run.Background = $codeBrush
            }
            if ($block.Kind -eq 'Quote') {
                $run.FontStyle = [System.Windows.FontStyles]::Italic
            }
            $TextBlock.Inlines.Add($run)
        }

        $TextBlock.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
    }
}
