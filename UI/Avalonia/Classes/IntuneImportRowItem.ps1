#ImportOrder 10

# CLR-typed row shape for the Import dialog's dgObjectsToImport DataGrid.
# Get-PoliciesFromFolder returns IntunePolicyBase / PSCustomObject hybrids
# whose nested-property bindings (Object.Name, Object.PolicyName, etc.)
# don't survive Avalonia's reflection-based binder
# (see [[avalonia-binding-needs-clr-types]]). Show-IntuneManagerImportForm
# projects each policy into one of these before assigning ItemsSource.
# Source/ImportTarget remain on .Source / .ImportTarget so the import
# pass can recover the originals.

class IntuneImportRowItem
{
    [bool]   $Selected = $true
    [string] $ObjectName
    [string] $PolicyType
    [string] $PolicyBase
    [string] $Platform
    [string] $FileName
    [string] $ImportAction
    [string] $ImportMatch
    [string] $ImportMatchStrategy
    [object] $Source        # underlying IntunePolicyBase (the .Object the WPF code reads)
    [object] $ImportTarget  # existing policy resolved by Resolve-IntuneImportUpdateTarget

    IntuneImportRowItem() {}
}
