#ImportOrder 8

# CLR-typed row shape for the Intune Assignments DataGrid. Get-IntuneAssignmentRow
# in Internal/IntuneAssignments.ps1 returns [PSCustomObject]; Avalonia's binder
# can't traverse those (see [[avalonia-binding-needs-clr-types]]), so the
# Avalonia tool converts each PSCustomObject into one of these before assigning
# ItemsSource. Properties mirror what the DataGrid columns + filter + CSV
# export read.

class IntuneAssignmentRowItem
{
    [string]   $Name
    [string]   $Type
    [int]      $AssignmentCount
    [bool]     $HasFilters
    [string]   $IncludedString
    [string]   $ExcludedString
    [string]   $IncludedFilterString
    [string]   $ExcludedFilterString
    [string]   $Id
    [object]   $Object
    [object[]] $Included
    [object[]] $Excluded
    [object[]] $IncludedFilters
    [object[]] $ExcludedFilters

    IntuneAssignmentRowItem() {}
}
