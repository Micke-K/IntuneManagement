# Text columns for WPF DataGrids: one line per row, capped length, pixel scrolling.
#
# A DataGridTextColumn renders every line of a multi-line value, so a store app
# description makes its row forty lines tall and turns the grid's row-based
# scrolling into a lurch. TextWrapping=NoWrap does not help: it only stops soft
# wrapping, and hard line breaks still render. And a single very long line has
# the opposite problem - it widens its column until the others are pushed out of
# view. FirstLineConverter trims the bound value to its first non-blank line,
# caps it at MaxLength characters (0 = no cap), appends an ellipsis when either
# happened, and with ConverterParameter "Full" returns the whole text only when
# something was cut - so the tooltip appears exactly on the cells that hide it.

if(-not ("FirstLineConverter" -as [type]))
{
    # References resolved from loaded types, so the same call compiles on
    # .NET Framework (PS5.1) and .NET (PS7).
    Add-Type -ReferencedAssemblies @(
        [System.Windows.Data.IValueConverter].Assembly.Location,       # PresentationFramework
        [System.Windows.DependencyObject].Assembly.Location,           # WindowsBase
        [System.Windows.Markup.MarkupExtension].Assembly.Location      # System.Xaml - Binding.DoNothing needs it on .NET Framework
    ) -TypeDefinition @'
using System;
using System.Globalization;
using System.Windows.Data;

public class FirstLineConverter : IValueConverter
{
    // Maximum characters kept from the first line. 0 = unlimited.
    public int MaxLength { get; set; }

    // First non-blank line, capped at maxLength; "\u2026" appended when any
    // later non-blank line exists or the cap cut something off.
    public static string FirstLine(string text, int maxLength, out bool trimmed)
    {
        trimmed = false;
        if (text == null) return null;
        string[] lines = text.Split(new[] { '\r', '\n' });
        string first = null;
        foreach (string line in lines)
        {
            string t = line.Trim();
            if (t.Length == 0) continue;
            if (first == null) { first = t; continue; }
            trimmed = true;
            break;
        }
        if (first == null) return string.Empty;
        if (maxLength > 0 && first.Length > maxLength)
        {
            // Prefer a word boundary when one falls in the second half of the
            // cap, so the cell ends "...built in" rather than "...with Cop".
            int cut = first.LastIndexOf(' ', maxLength);
            if (cut < maxLength / 2) cut = maxLength;
            first = first.Substring(0, cut).TrimEnd();
            trimmed = true;
        }
        return trimmed ? first + "\u2026" : first;
    }

    // A binding to a PowerShell SCRIPT property (the row's Description, Platform,
    // LastModified ...) delivers the value wrapped in a PSObject; a note property
    // (Object.description) delivers the raw string. Unwrap by reflection so the
    // converter sees the string either way without referencing the PowerShell
    // assembly - this is what left Description rendering all its lines while
    // Object.description was trimmed.
    private static object Unwrap(object value)
    {
        if (value == null) return null;
        var t = value.GetType();
        if (t.FullName != "System.Management.Automation.PSObject") return value;
        var p = t.GetProperty("BaseObject");
        return p == null ? value : p.GetValue(value, null);
    }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        value = Unwrap(value);
        string text = value as string;
        if (text == null) return value;   // dates, numbers: leave the column's own formatting alone
        bool trimmed;
        string first = FirstLine(text, MaxLength, out trimmed);
        if (parameter != null && string.Equals(parameter.ToString(), "Full", StringComparison.OrdinalIgnoreCase))
        {
            return trimmed ? text : null; // tooltip only where something was hidden
        }
        return first;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return Binding.DoNothing;
    }
}
'@
}

# Builds a read-only text column. With -FirstLineOnly the cell shows the value's
# first line (capped at -MaxLength characters when that is above 0) and carries
# the full text as a tooltip. Sorting and the grid filter still use the
# underlying value: both read the binding path, not what is displayed.
function New-GridTextColumn
{
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Header,
        [switch]$FirstLineOnly,
        [int]$MaxLength = 0
    )

    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header     = if($Header) { $Header } else { $Path.Split('.')[-1] }
    $column.IsReadOnly = $true

    $binding = [System.Windows.Data.Binding]::new($Path)
    if($FirstLineOnly)
    {
        $binding.Converter = [FirstLineConverter]::new()
        $binding.Converter.MaxLength = $MaxLength

        $tooltip = [System.Windows.Data.Binding]::new($Path)
        $tooltip.Converter           = [FirstLineConverter]::new()
        $tooltip.Converter.MaxLength = $MaxLength
        $tooltip.ConverterParameter  = 'Full'

        $style = [System.Windows.Style]::new([System.Windows.Controls.TextBlock])
        $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.TextBlock]::TextTrimmingProperty, [System.Windows.TextTrimming]::CharacterEllipsis))
        $style.Setters.Add([System.Windows.Setter]::new([System.Windows.FrameworkElement]::ToolTipProperty, $tooltip))
        $column.ElementStyle = $style
    }
    $column.Binding = $binding

    return $column
}

# Pixel-based scrolling for a DataGrid built in code; the XAML grids declare
# VirtualizingPanel.ScrollUnit="Pixel" directly. WPF's default scrolls by whole
# rows and snaps the top row to the edge, which lurches as soon as rows differ in
# height. Row virtualization stays on.
function Set-GridPixelScrolling
{
    param([Parameter(Mandatory = $true)]$Grid)

    [System.Windows.Controls.VirtualizingPanel]::SetScrollUnit($Grid, [System.Windows.Controls.ScrollUnit]::Pixel)
    [System.Windows.Controls.VirtualizingPanel]::SetVirtualizationMode($Grid, [System.Windows.Controls.VirtualizationMode]::Recycling)
}
