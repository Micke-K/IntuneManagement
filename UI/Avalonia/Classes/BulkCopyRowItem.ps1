#ImportOrder 10

# CLR row shape for Show-GraphBulkCopyForm's dgBulkCopyObjects. Same reason
# as [[BulkDeleteRowItem]] / [[avalonia-binding-needs-clr-types]] — Avalonia's
# binder ignores PSObject NoteProperties. Default Selected is true (matches
# the original Bulk Copy where every type was pre-checked).

class BulkCopyRowItem
{
    [bool]   $Selected = $true
    [string] $Title
    [object] $ObjectType

    BulkCopyRowItem() {}
}
