#ImportOrder 10

# CLR row shape for Show-GraphBulkScopeTagForm's dgBulkScopeTagObjects.
# Same reasoning as [[BulkExportRowItem]] / [[avalonia-binding-needs-clr-types]] —
# Avalonia's binder ignores PSObject NoteProperties, so Selected / Title would
# render blank if we used PSCustomObject. Mirrors BulkExportRowItem because
# this form also offers Group / API mode (one of ObjectGroup / ObjectType is
# populated, the other is $null).

class BulkScopeTagRowItem
{
    [bool]   $Selected = $true
    [string] $Title
    [object] $ObjectGroup
    [object] $ObjectType

    BulkScopeTagRowItem() {}
}
