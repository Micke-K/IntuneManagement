#ImportOrder 9

# CLR-typed row shape for the Graph Calls DataGrid (GraphCallsPanel.axaml).
# Internal/MSGraph.ps1 stores entries in $script:AllGraphCalls as
# PSCustomObject; Update-GraphCallsView projects them into these so the
# Avalonia binder can resolve column paths and the BatchRequests sub-grid.
#
# BatchRequests is exposed as List[GraphBatchRequestRowItem] so the nested
# DataGrid (bound via SelectedItem.BatchRequests) gets a typed enumerable.

class GraphCallRowItem
{
    [string]   $Provider
    [datetime] $Time
    [double]   $Duration
    [double]   $KB
    [int]      $ObjectCount
    [int]      $PageCount
    [string]   $StatusCode
    [string]   $BatchErrorSummary
    [string]   $Method
    [string]   $URL
    [bool]     $IsBatch
    [System.Collections.Generic.List[GraphBatchRequestRowItem]] $BatchRequests

    GraphCallRowItem() {
        $this.BatchRequests = [System.Collections.Generic.List[GraphBatchRequestRowItem]]::new()
    }
}
