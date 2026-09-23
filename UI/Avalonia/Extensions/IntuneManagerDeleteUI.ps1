# Per-view Delete confirmation (Remove-GraphObjectsUI).
# Split out of IntuneManagerUIAvalonia.ps1 on 2026-06-03 to match the WPF tree's
# per-feature layout (architecture rule R9 — keep the Avalonia shell from growing further).

function Remove-GraphObjectsUI
{
    $ui = $script:UIProvider
    if (-not $script:dgIntuneManagerObjects -or $null -eq $script:IntuneManagerAllRows) { return }

    # Prefer the IsSelected-checked rows; fall back to the highlighted row.
    # Scan the grid's CURRENT (filtered) ItemsSource - never the unfiltered
    # backing list, or a checked row hidden by the filter would be deleted
    # without ever being visible to the user.
    $selectedRows = @(@($script:dgIntuneManagerObjects.ItemsSource) | Where-Object { $_ -and $_.IsSelected })
    if ($selectedRows.Count -eq 0 -and $script:dgIntuneManagerObjects.SelectedItem) {
        $selectedRows = @($script:dgIntuneManagerObjects.SelectedItem)
    }

    if ($selectedRows.Count -eq 0) {
        $ui.ShowMessageBox("No object selected`n`nSelect items you want to delete", "Error", "OK", "Error")
        return
    }

    $confirm = $ui.ShowMessageBox("Are you sure you want to delete $($selectedRows.Count) object(s)?`n`nEnvironment: $($script:OrganizationName)", "Delete Objects?", "YesNo", "Warning")
    if ($confirm -ne "Yes") { return }

    # Remove-GraphPolicy expects the underlying IntunePolicyBase, not the
    # CLR row wrapper — pull .Source off each row before piping in.
    $policiesToRemove = @($selectedRows | ForEach-Object { $_.Source } | Where-Object { $_ })
    if ($policiesToRemove.Count -eq 0) {
        Write-Log "Delete: $($selectedRows.Count) row(s) confirmed but none carried a policy object" 3
        $ui.ShowMessageBox("No object selected`n`nSelect items you want to delete", "Error", "OK", "Error")
        return
    }

    $deletedPolicies = @($policiesToRemove | Remove-GraphPolicy)

    if ($deletedPolicies.Count -gt 0) {
        $deletedIds = @($deletedPolicies | ForEach-Object { [string]$_.Id })
        # Drop deleted rows from the cached unfiltered list, then refresh the
        # filtered ItemsSource so the grid reflects the new state.
        $remaining = [System.Collections.Generic.List[object]]::new()
        foreach ($row in $script:IntuneManagerAllRows) {
            if ($deletedIds -notcontains [string]$row.ID) { [void]$remaining.Add($row) }
        }
        $script:IntuneManagerAllRows = $remaining

        Invoke-FilterBoxChanged $script:IntuneManagementFilterTextBox `
            $script:dgIntuneManagerObjects -ForceUpdate `
            -ObjectsCountTextBox $script:IntuneManagementObjectsCount
    }

    Write-Status ""
}
