#ImportOrder 8

# CLR-typed row shape for the nested batch-request DataGrid in
# GraphCallsPanel.axaml. Lives one level below GraphCallRowItem and is bound
# via {Binding SelectedItem.BatchRequests, ElementName=dgGraphCalls}.
#
# ImportOrder 8 (loads before GraphCallRowItem at 9) so the
# List[GraphBatchRequestRowItem] type reference on GraphCallRowItem resolves.

class GraphBatchRequestRowItem
{
    [string] $Id
    [double] $KB
    [int]    $ObjectCount
    [int]    $PageCount
    [string] $StatusCode
    [string] $Method
    [string] $URL

    GraphBatchRequestRowItem() {}
}
