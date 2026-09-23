#ImportOrder 10

# CLR row shape for Show-GraphBulkDocumentationForm's dgObjectsToDocument.
# Same reason as [[BulkExportRowItem]] / [[avalonia-binding-needs-clr-types]] —
# Avalonia's binder ignores PSObject NoteProperties, so Selected and Title
# would render blank if we used PSCustomObject. ObjectGroup / ObjectType
# hold references to the underlying IntuneGroup / IntuneType / policy so the
# click handler can pass the appropriate selection into Start-GraphBulkDocumentation.

class BulkDocumentationRowItem
{
    [bool]   $Selected = $true
    [string] $Title
    [object] $ObjectGroup
    [object] $ObjectType
    [object] $PolicyObject

    BulkDocumentationRowItem() {}
}
