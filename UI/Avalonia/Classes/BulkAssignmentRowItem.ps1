#ImportOrder 10

# CLR row shape for the dgBulkAssignList DataGrid in Show-GraphBulkAssignmentsForm.
# The WPF original uses PSCustomObject here; Avalonia's binder needs real CLR
# properties (see [[avalonia-binding-needs-clr-types]]).
#
# Set-GraphBulkAssignments accesses these by name (.TargetType, .GroupId,
# .Intent, .Settings as [Hashtable]) so the public command consumes class
# instances the same way it consumes PSCustomObject — confirmed by reading
# Public/Set-GraphBulkAssignments.ps1.
#
# Settings is keyed by Graph settings type name (e.g.
# "win32LobAppAssignmentSettings"). Absent key = no override for that type.
# SettingsDisplay is a derived count string ("(none)" / "N platform(s)") that
# the App settings... dialog refreshes on save.

class BulkAssignmentRowItem
{
    [string]    $TargetType
    [string]    $TargetTypeDisplay
    [string]    $GroupId
    [string]    $GroupName
    [string]    $FilterId
    [string]    $FilterName
    [string]    $FilterType      = "include"
    [string]    $FilterDisplay   = ""
    [string]    $Intent          = "required"
    [Hashtable] $Settings        = @{}
    [string]    $SettingsDisplay = "(none)"

    BulkAssignmentRowItem() {}
}
