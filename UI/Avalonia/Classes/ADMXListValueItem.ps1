#ImportOrder 6

# CLR-typed row for the ADMX listBox element editor (Setting Properties
# dialog). Was previously a `ListValue` class defined via Invoke-Expression
# inside Set-ADMXElementsPanel - a type created that way is not resolvable
# from the module's event handlers at click time, which killed the list
# editor's Add button. A real class file loads with the module and is
# visible everywhere (and, being CLR-typed, binds in the DataGrid - see
# [[avalonia-binding-needs-clr-types]]).

class ADMXListValueItem
{
    [string]$Key   = ''
    [string]$Value = ''

    ADMXListValueItem() {}
}
