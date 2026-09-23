#ImportOrder 10

# CLR-typed item shape for the Copy dialog's destination-tenant ComboBox.
# Get-TokenInfo returns PSObjects whose NoteProperties don't bind through
# Avalonia's reflection-based ItemTemplate (see [[avalonia-binding-needs-clr-types]]).
# Project the fields we actually need into this class; the original token
# stays on .Source for the Copy call.

class CopyDialogTenantItem
{
    [string] $TenantNameEx
    [int]    $Id
    [bool]   $IsDefault
    [object] $Source

    CopyDialogTenantItem() {}

    # Backstop only - the control selects the field via DisplayMemberBinding.
    [string] ToString() { return [string]$this.TenantNameEx }
}
