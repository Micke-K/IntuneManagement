#ImportOrder 10

# CLR-typed row shape for the Intune Manager DataGrid. The WPF implementation
# wraps each IntunePolicyBase via [PSCustomObject]$policy + Add-Member; that
# survives WPF's PSObject special-case binding but renders empty in Avalonia
# (see [[avalonia-binding-needs-clr-types]]). Show-SelectedGraphPolicies
# projects the canonical view properties onto this class before binding.
#
# Source / JsonObject / TokenId mirror the fields the bulk forms read off
# the row directly. Custom column overrides (per-type `ObjectColumns`
# setting) reach paths beyond these canonical fields and bind through
# [IntuneManagement.AvaloniaHost.IntunePolicyPathConverter] (declared in
# Bootstrap/AvaloniaHost.cs, compiled at runtime) against `Source` so
# PSObject NoteProperties on the underlying IntunePolicyBase resolve.

class IntuneObjectRowItem
{
    [bool]   $IsSelected = $false
    [string] $Name
    [string] $PolicyName
    [string] $Platform
    [string] $Description
    [string] $LastModified
    [string] $Created
    [string] $ID
    [string] $ScopeTags

    [object] $JsonObject
    [object] $Source
    [int]    $TokenId

    IntuneObjectRowItem() {}
}
