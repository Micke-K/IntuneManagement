#ImportOrder 10

# CLR row shape for Show-GraphBulkExportForm's dgObjectsToExport. WPF used
# [PSCustomObject]@{...} but Avalonia's binder ignores PSObject NoteProperties
# (see [[avalonia-binding-needs-clr-types]]) — Selected and Title would render
# blank. ObjectGroup / ObjectType keep references to the underlying
# IntuneGroup / IntuneType so Get-BulkExportSelectedIds can pull .Id off them.

class BulkExportRowItem
{
    [bool]   $Selected = $true
    [string] $Title
    [object] $ObjectGroup
    [object] $ObjectType

    BulkExportRowItem() {}
}
