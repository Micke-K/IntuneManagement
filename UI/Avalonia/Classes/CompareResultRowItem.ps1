#ImportOrder 11

# CLR-typed row shape for the Compare DataGrid. Compare-PolicyObjects
# returns PSCustomObject rows, which Avalonia's reflection-based binder
# can't traverse (see [[avalonia-binding-needs-clr-types]]). Project each
# into one of these before assigning DataGrid.ItemsSource.
#
# Match is nullable bool (e.g. ignored core properties produce $null).
# Slice 3d3 uses it to drive row foreground color.
#
# MatchGlyph is the precomputed tri-state glyph the "Match" DataGrid column
# binds to (checkmark / ballot-X / em-dash), mirroring the WPF CompareForm.xaml
# DataTrigger chain. Avalonia has no WPF-style per-value DataTrigger, so the
# glyph is computed at projection time instead of in XAML.

class CompareResultRowItem
{
    [string] $PropertyName
    [string] $Category
    [string] $SubCategory
    [string] $Object1Value
    [string] $Object2Value
    [object] $Match
    [string] $MatchGlyph

    CompareResultRowItem() {}
}
