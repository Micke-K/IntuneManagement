#ImportOrder 10

# CLR-typed row for the Permissions dialog's DataGrid
# (Extensions/EffectivePermissionsUIAvalonia.ps1). Get-GraphEffectivePermissions
# returns PSCustomObjects, which the Avalonia binder renders blank, so the
# dialog converts each row to an instance of this class.

# Property declaration order is the column order (AutoGenerateColumns).
class EffectivePermissionRowItem
{
    [string] $Type
    [string] $Category
    [string] $Required
    [string] $Token
    [string] $IntuneRole
    [string] $Effective
    [string] $Result
    [string] $Reason

    EffectivePermissionRowItem() {}
}
