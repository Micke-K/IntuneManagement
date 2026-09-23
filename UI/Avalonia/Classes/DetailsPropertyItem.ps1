#ImportOrder 10

# CLR row shape for the Object Details "Columns" tab — used by both
# lstBasicProperties and lstObjectProperties. WPF sets DisplayMemberPath="Name"
# on a [PSCustomObject]@{ Name; Value; Source }, but Avalonia has no
# DisplayMemberPath and its binder doesn't reflect into PSCustomObject
# NoteProperties (see [[avalonia-binding-needs-clr-types]]) — so we project
# into this typed shape and wire an inline ItemTemplate.

class DetailsPropertyItem
{
    [string] $Name
    [object] $Value
    [string] $Source

    DetailsPropertyItem() {}

    [string] ToString()
    {
        return [string]$this.Name
    }
}
