#ImportOrder 10

# CLR-typed shape for items inside the Copy dialog's scope-tag ListBox
# controls. Get-GraphPolicies / Get-CacheObject return PSCustomObject tags
# whose NoteProperties don't bind through Avalonia's reflection-based
# ItemTemplate (see [[avalonia-binding-needs-clr-types]]). The Copy dialog
# projects each PSCustomObject tag into one of these before assigning to
# the ObservableCollection bound to the ListBox.

class CopyDialogScopeTagItem
{
    [string] $Id
    [string] $Name

    CopyDialogScopeTagItem() {}

    # Backstop only - the control selects the field via DisplayMemberBinding.
    [string] ToString() { return [string]$this.Name }
}
