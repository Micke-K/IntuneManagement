#ImportOrder 10

# CLR row shape for Show-GraphBulkDeleteForm's dgBulkDeleteObjects. Same
# reason as [[BulkExportRowItem]] / [[avalonia-binding-needs-clr-types]] —
# Avalonia's binder ignores PSObject NoteProperties. Default Selected is
# false (matches WPF; bulk delete is destructive and shouldn't pre-select).

class BulkDeleteRowItem
{
    [bool]   $Selected = $false
    [string] $Title
    [object] $ObjectGroup

    BulkDeleteRowItem() {}
}
