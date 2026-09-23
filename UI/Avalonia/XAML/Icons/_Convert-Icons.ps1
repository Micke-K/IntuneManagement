# One-shot icon converter — translates the WPF icon shape XAML under
# UI/WPF/XAML/Icons/*.xaml into Avalonia 11 .axaml files in this folder.
#
# Two transforms are needed (both verified against the WPF originals):
#   1. WPF presentation namespace → Avalonia namespace.
#   2. LinearGradientBrush.MappingMode:
#        "RelativeToBoundingBox"  → drop the attribute (Avalonia's default
#                                   matches; StartPoint/EndPoint stay 0..1).
#        "Absolute"               → drop the attribute + suffix
#                                   StartPoint/EndPoint with ",Absolute" so
#                                   Avalonia's RelativePoint parser reads
#                                   them as absolute device-independent units.
#
# All other shape primitives (Canvas, Path Data, Rectangle RadiusX/Y,
# Canvas.Top / Canvas.Left attached props, GradientStop Offset/Color,
# StaticResource keys) carry over unchanged.
#
# Run from anywhere; emits sibling .axaml files into this Icons folder.

$wpfRoot = Join-Path $PSScriptRoot '..\..\..\UI\XAML/Icons'
$wpfRoot = (Resolve-Path $wpfRoot).Path
$dstRoot = $PSScriptRoot

$converted = 0
foreach ($file in (Get-ChildItem -Path $wpfRoot -Filter '*.xaml')) {
    $xaml = [IO.File]::ReadAllText($file.FullName)

    $xaml = $xaml.Replace(
        'http://schemas.microsoft.com/winfx/2006/xaml/presentation',
        'https://github.com/avaloniaui')

    $xaml = $xaml -replace '\s+MappingMode="RelativeToBoundingBox"', ''

    if ($xaml -match 'MappingMode="Absolute"') {
        $xaml = [regex]::Replace($xaml, '<LinearGradientBrush\b[^>]*MappingMode="Absolute"[^>]*>', {
            param($m)
            $tag = $m.Value
            $tag = $tag -replace '\s+MappingMode="Absolute"', ''
            $tag = [regex]::Replace($tag, 'StartPoint="([^"]+)"', { param($mm) 'StartPoint="' + $mm.Groups[1].Value + ',Absolute"' })
            $tag = [regex]::Replace($tag, 'EndPoint="([^"]+)"',   { param($mm) 'EndPoint="'   + $mm.Groups[1].Value + ',Absolute"' })
            return $tag
        })
    }

    $dst = Join-Path $dstRoot ($file.BaseName + '.axaml')
    [IO.File]::WriteAllText($dst, $xaml, [Text.UTF8Encoding]::new($false))
    $converted++
}
Write-Host "Converted $converted icons → $dstRoot"
