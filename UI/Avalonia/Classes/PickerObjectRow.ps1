#ImportOrder 10

# CLR row shape for Show-ObjectPickerDialog's dgPickerObjects DataGrid.
# Avalonia's binder can't reach PSObject NoteProperties (see
# [[avalonia-binding-needs-clr-types]]), so each policy fetched from
# Get-GraphPolicies is projected into one of these. Source carries the
# original PSCustomObject so callers receive it back via SelectedItem.Source.
# Currently the dialog is called with a "Name" + "PolicyType.Title" column
# pair — mapped to the Name / PolicyTypeTitle properties below. Add a new
# property + binding alias in Show-ObjectPickerDialog when a future caller
# needs another column.

class PickerObjectRow
{
    [string] $Name
    [string] $PolicyTypeTitle
    [string] $Description
    [object] $Source

    PickerObjectRow() {}
}
