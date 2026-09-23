# Null-free array for anything assigned to a WPF ItemsSource.
#
# A $null element in an ItemsControl kills the host: WPF's layout and automation
# passes index the items, and both throw ArgumentNullException on a null one, on
# every render, until the process dies. @($null) is a ONE-element array containing
# null, so an omitted parameter or an empty result lands there via the usual @($x).
# Full write-up and the regression proof: Tests/UIItemsArray.Tests.ps1.
#
# The leading comma is load-bearing. A bare `return @(...)` is unrolled by the
# pipeline, so zero items would come back as $null and ONE item as the bare
# object - which cannot be assigned to an ItemsSource at all.
function Get-UIItemsArray
{
    param($Items)

    return ,@(@($Items) | Where-Object { $null -ne $_ })
}
