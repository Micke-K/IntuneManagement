#ImportOrder 10

# CLR row shape for Show-GraphBulkAssignmentsForm's dgBulkAssignObjects
# DataGrid (the left-hand "Objects to update" grid). Mirrors
# [[BulkScopeTagRowItem]] / [[BulkExportRowItem]] — Group/API mode means
# exactly one of ObjectGroup / ObjectType is populated, the other is $null.
# Avalonia's binder ignores PSObject NoteProperties, so PSCustomObject would
# render Selected / Title blank — see [[avalonia-binding-needs-clr-types]].

class BulkAssignRowItem
{
    [bool]   $Selected = $true
    [string] $Title
    [object] $ObjectGroup
    [object] $ObjectType

    BulkAssignRowItem() {}
}
