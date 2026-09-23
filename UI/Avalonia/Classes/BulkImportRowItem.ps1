#ImportOrder 10

# CLR row shape for Show-GraphBulkImportForm's dgObjectsToImport. Same reason
# as [[BulkExportRowItem]] / [[avalonia-binding-needs-clr-types]] — Avalonia's
# binder ignores PSObject NoteProperties, so Selected and Title would render
# blank if we used [PSCustomObject]@{...} like the WPF original. ObjectGroup
# keeps a reference to the underlying IntuneGroup so the import handler can
# pull .PolicyTypes off it.

class BulkImportRowItem
{
    [bool]   $Selected = $true
    [string] $Title
    [object] $ObjectGroup

    BulkImportRowItem() {}
}
