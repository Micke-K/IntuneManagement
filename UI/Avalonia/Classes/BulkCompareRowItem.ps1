#ImportOrder 10

# CLR row shape for Show-GraphBulkCompareForm's dgObjectsToCompare. Same
# reason as [[BulkExportRowItem]] / [[avalonia-binding-needs-clr-types]] —
# Avalonia's binder ignores PSObject NoteProperties. Default Selected is
# true (matches WPF original — every object type pre-selected for a
# bulk compare run).

class BulkCompareRowItem
{
    [bool]   $Selected = $true
    [string] $Title
    [object] $ObjectGroup

    BulkCompareRowItem() {}
}
