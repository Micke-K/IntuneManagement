# Minimal Markdown reader for the Release Notes dialog.
#
# Deliberately NOT a Markdown implementation. It handles exactly what
# ReleaseNotes.md actually contains, which is: one '#' heading, 42 '##' headings,
# 555 bullets nested up to three levels, 247 '**bold**' spans, 163 '[text](url)'
# links, 652 '<br />' tags and 6 '<b>' pairs. No numbered lists, code fences,
# tables, blockquotes or rules appear in the file - numbered lists are supported
# anyway because it costs one regex and future notes may use them.
#
# Output is UI-toolkit-free so WPF and Avalonia can share it (R10/R12): a flat
# list of blocks, each carrying inline runs. Both backends render into a
# TextBlock, which supports per-run FontSize/FontWeight, so a single control can
# present the whole document.
#
# Block : Kind = Heading | ListItem | Paragraph | Blank
#         Level  (Heading: 1-6)
#         Indent (ListItem: nesting depth, 0-based)
#         Marker (ListItem: '-' for bullets, or the literal '1.' style number)
#         Inlines
# Inline: Text, Bold, LineBreak

# Split one line of markdown into inline runs.
function ConvertFrom-SimpleMarkdownInline {
    param([string]$Text)

    $inlines = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $Text) { return $inlines }

    # One pass over the constructs that can appear mid-line. Alternation order
    # matters: '<br />' before '<b>' so the break is not mistaken for bold, and
    # `code` before the rest so markup inside a code span stays literal.
    $pattern = '(?<br><br\s*/?>)|(?<code>`(?<ctext>[^`]+)`)|(?<bold>\*\*(?<b1>.+?)\*\*)|(?<htmlbold><b>(?<b2>.*?)</b>)|(?<link>\[(?<ltext>[^\]]+)\]\((?<lurl>[^)]*)\))'
    $pos = 0

    foreach ($m in [regex]::Matches($Text, $pattern)) {
        if ($m.Index -gt $pos) {
            $inlines.Add((New-SimpleMarkdownInline -Text $Text.Substring($pos, $m.Index - $pos)))
        }

        if ($m.Groups['br'].Success) {
            $inlines.Add((New-SimpleMarkdownInline -LineBreak))
        }
        elseif ($m.Groups['code'].Success) {
            # Code spans are literal: no recursion, so **stars** and [links] inside
            # them survive as typed.
            $inlines.Add((New-SimpleMarkdownInline -Text $m.Groups['ctext'].Value -Code))
        }
        elseif ($m.Groups['bold'].Success -or $m.Groups['htmlbold'].Success) {
            # Recurse into the bold text so constructs nested inside it are still
            # handled - the file has a bold bullet containing a link, which would
            # otherwise render with its raw [text](url) markup showing. The inner
            # match is non-greedy and cannot contain another bold opener, so this
            # terminates after one level.
            $boldText = if ($m.Groups['bold'].Success) { $m.Groups['b1'].Value } else { $m.Groups['b2'].Value }
            foreach ($inner in (ConvertFrom-SimpleMarkdownInline $boldText)) {
                $inner.Bold = $true
                $inlines.Add($inner)
            }
        }
        elseif ($m.Groups['link'].Success) {
            $inlines.Add((New-SimpleMarkdownInline -Text $m.Groups['ltext'].Value -Url $m.Groups['lurl'].Value))
        }

        $pos = $m.Index + $m.Length
    }

    if ($pos -lt $Text.Length) {
        $inlines.Add((New-SimpleMarkdownInline -Text $Text.Substring($pos)))
    }

    return $inlines
}

# One inline run. Kept as a factory so every producer emits the same shape and a
# renderer can rely on every field existing.
function New-SimpleMarkdownInline {
    param(
        [string]$Text = '',
        [switch]$Bold,
        [switch]$Code,
        [switch]$LineBreak,
        [string]$Url = $null
    )
    return [PSCustomObject]@{
        Text      = $Text
        Bold      = [bool]$Bold
        Code      = [bool]$Code
        LineBreak = [bool]$LineBreak
        Url       = $Url
    }
}

# Parse a markdown document into renderable blocks.
function ConvertFrom-SimpleMarkdown {
    [CmdletBinding()]
    param([string]$Markdown)

    $blocks = [System.Collections.Generic.List[object]]::new()
    if ([string]::IsNullOrEmpty($Markdown)) { return $blocks }

    $inCodeFence = $false

    foreach ($line in ($Markdown -split "`r?`n")) {

        # Fenced code. The opening fence may carry a language tag; it is read and
        # discarded - no syntax handling, the block is styled as plain monospace.
        $fence = [regex]::Match($line, '^\s*(```|~~~)\s*(\S*)\s*$')
        if ($fence.Success) {
            $inCodeFence = -not $inCodeFence
            continue
        }
        if ($inCodeFence) {
            # Literal: no inline parsing, so markup inside code stays as typed.
            $blocks.Add([PSCustomObject]@{
                Kind = 'CodeBlock'; Level = 0; Indent = 0; Marker = ''
                Inlines = @((New-SimpleMarkdownInline -Text $line -Code))
            })
            continue
        }

        if ([string]::IsNullOrWhiteSpace($line)) {
            $blocks.Add([PSCustomObject]@{ Kind = 'Blank'; Level = 0; Indent = 0; Marker = ''; Inlines = @() })
            continue
        }

        $quote = [regex]::Match($line, '^\s*>\s?(.*)$')
        if ($quote.Success) {
            $blocks.Add([PSCustomObject]@{
                Kind = 'Quote'; Level = 0; Indent = 0; Marker = ''
                Inlines = (ConvertFrom-SimpleMarkdownInline $quote.Groups[1].Value)
            })
            continue
        }

        $heading = [regex]::Match($line, '^(#{1,6})\s+(.*)$')
        if ($heading.Success) {
            $blocks.Add([PSCustomObject]@{
                Kind    = 'Heading'
                Level   = $heading.Groups[1].Value.Length
                Indent  = 0
                Marker  = ''
                Inlines = (ConvertFrom-SimpleMarkdownInline $heading.Groups[2].Value)
            })
            continue
        }

        $bullet = [regex]::Match($line, '^(\s*)[-*+]\s+(.*)$')
        if ($bullet.Success) {
            # The file indents nested bullets by 2 spaces (with one stray 1-space
            # and a 4-space third level), so integer-divide by 2 and let the odd
            # one fall to level 0.
            $blocks.Add([PSCustomObject]@{
                Kind    = 'ListItem'
                Level   = 0
                Indent  = [int]([Math]::Floor($bullet.Groups[1].Value.Length / 2))
                Marker  = '-'
                Inlines = (ConvertFrom-SimpleMarkdownInline $bullet.Groups[2].Value)
            })
            continue
        }

        $numbered = [regex]::Match($line, '^(\s*)(\d+)[\.\)]\s+(.*)$')
        if ($numbered.Success) {
            $blocks.Add([PSCustomObject]@{
                Kind    = 'ListItem'
                Level   = 0
                Indent  = [int]([Math]::Floor($numbered.Groups[1].Value.Length / 2))
                Marker  = "$($numbered.Groups[2].Value)."
                Inlines = (ConvertFrom-SimpleMarkdownInline $numbered.Groups[3].Value)
            })
            continue
        }

        $blocks.Add([PSCustomObject]@{
            Kind    = 'Paragraph'
            Level   = 0
            Indent  = 0
            Marker  = ''
            Inlines = (ConvertFrom-SimpleMarkdownInline $line.Trim())
        })
    }

    return $blocks
}

# Presentation constants shared by both renderers, so WPF and Avalonia produce
# the same look. Colours are given as ARGB hex.
#
# CodeBackground is a low-alpha grey rather than a fixed light or dark colour:
# both backends have light and dark themes, and a translucent grey reads as a
# "slightly different background" over either without needing theme lookups.
function Get-SimpleMarkdownStyle {
    return @{
        CodeBackground = '#22808080'
        CodeFontFamily = 'Consolas, Cascadia Mono, Courier New, monospace'
        QuoteMarker    = '| '
    }
}

# Font size for a heading level, relative to the control's base size. Shared so
# both backends size headings identically.
function Get-SimpleMarkdownHeadingSize {
    param([int]$Level, [double]$BaseSize = 12)

    switch ($Level) {
        1       { return $BaseSize * 1.6 }
        2       { return $BaseSize * 1.3 }
        3       { return $BaseSize * 1.15 }
        default { return $BaseSize * 1.05 }
    }
}
