#ImportOrder 9

# CLR-typed row shape for the Intune Filter Usage DataGrid. Get-IntuneFilterUsageData
# returns [PSCustomObject] rows; Avalonia binder can't traverse those, so
# Start-IntuneFilterUsageLoad converts each into one of these before assigning
# DataGrid.ItemsSource.

class IntuneFilterUsageRowItem
{
    [string] $FilterName
    [string] $Platform
    [string] $FilterType
    [string] $PolicyName
    [string] $PayloadType
    [string] $Mode
    [string] $GroupId
    [string] $GroupName

    IntuneFilterUsageRowItem() {}
}
