# Render parsed Markdown into an Avalonia TextBlock.
#
# Mirror of UI/WPF/Extensions/MarkdownRenderWPF.ps1. The parse is shared via
# Internal/MarkdownSimple.ps1 (R10); only the inline construction differs, so the
# Avalonia types stay inside UI/Avalonia (R12).
#
# Translation notes vs the WPF version:
#  - Inlines live in Avalonia.Controls.Documents, not System.Windows.Documents.
#  - No Bold element; Run.FontWeight carries it, which is what WPF uses too.
#  - Run has no text constructor overload, so Text is set as a property.
#  - No Hyperlink inline. Links use the pattern this repo already uses for inline
#    links in Welcome.axaml: a Button with Classes="link" (styled in
#    Themes/Styles.axaml) wrapped in an InlineUIContainer.

function Set-AvaloniaMarkdownText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $TextBlock,
        [string]$Markdown
    )

    if (-not $TextBlock) { return }

    if ($TextBlock.Inlines) { $TextBlock.Inlines.Clear() }
    if ([string]::IsNullOrEmpty($Markdown)) { return }

    $baseSize = if ($TextBlock.FontSize -gt 0) { [double]$TextBlock.FontSize } else { 12.0 }
    $style     = Get-SimpleMarkdownStyle
    $codeBrush = [Avalonia.Media.SolidColorBrush]::new([Avalonia.Media.Color]::Parse($style.CodeBackground))
    $codeFont  = [Avalonia.Media.FontFamily]::new($style.CodeFontFamily)

    foreach ($block in (ConvertFrom-SimpleMarkdown $Markdown)) {

        if ($block.Kind -eq 'Blank') {
            $TextBlock.Inlines.Add([Avalonia.Controls.Documents.LineBreak]::new())
            continue
        }

        $lead = switch ($block.Kind) {
            'ListItem' { (' ' * ($block.Indent * 4)) + $block.Marker + ' ' }
            'Quote'    { $style.QuoteMarker }
            default    { $null }
        }
        if ($lead) {
            $leadRun = [Avalonia.Controls.Documents.Run]::new()
            $leadRun.Text = $lead
            $TextBlock.Inlines.Add($leadRun)
        }

        foreach ($inline in $block.Inlines) {
            if ($inline.LineBreak) {
                $TextBlock.Inlines.Add([Avalonia.Controls.Documents.LineBreak]::new())
                if ($block.Kind -eq 'ListItem') {
                    $pad = [Avalonia.Controls.Documents.Run]::new()
                    $pad.Text = ' ' * (($block.Indent * 4) + $block.Marker.Length + 1)
                    $TextBlock.Inlines.Add($pad)
                }
                continue
            }

            if ($inline.Url) {
                # Captured by value into the handler: $inline is the loop variable
                # and would otherwise be whatever the loop ended on by click time.
                $linkUrl = [string]$inline.Url
                $button = [Avalonia.Controls.Button]::new()
                $button.Content = [string]$inline.Text
                $button.Classes.Add('link')
                $button.Add_Click({
                    param($clickSender, $clickArgs)
                    Open-ExternalUri ([string]$clickSender.Tag)
                })
                # Tag carries the URL so the handler reads it off the sender rather
                # than closing over a loop variable.
                $button.Tag = $linkUrl

                $container = [Avalonia.Controls.Documents.InlineUIContainer]::new()
                $container.Child = $button
                $TextBlock.Inlines.Add($container)
                continue
            }

            $run = [Avalonia.Controls.Documents.Run]::new()
            $run.Text = [string]$inline.Text
            if ($inline.Bold -or $block.Kind -eq 'Heading') {
                $run.FontWeight = [Avalonia.Media.FontWeight]::Bold
            }
            if ($block.Kind -eq 'Heading') {
                $run.FontSize = Get-SimpleMarkdownHeadingSize -Level $block.Level -BaseSize $baseSize
            }
            if ($inline.Code -or $block.Kind -eq 'CodeBlock') {
                $run.FontFamily = $codeFont
                $run.Background = $codeBrush
            }
            if ($block.Kind -eq 'Quote') {
                $run.FontStyle = [Avalonia.Media.FontStyle]::Italic
            }
            $TextBlock.Inlines.Add($run)
        }

        $TextBlock.Inlines.Add([Avalonia.Controls.Documents.LineBreak]::new())
    }
}
