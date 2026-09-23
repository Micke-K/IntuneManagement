# Per-view Delete confirmation (Remove-GraphObjectsUI).
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Remove-GraphObjectsUI
{
    $policiesToRemove = @()
    if(($script:dgIntuneManagerObjects.ItemsSource | Where-Object IsSelected -eq $true).Count -gt 0)
    {
        $policiesToRemove += $script:dgIntuneManagerObjects.ItemsSource | Where-Object IsSelected -eq $true
    }
    elseif($script:dgIntuneManagerObjects.SelectedItem)
    {
        $policiesToRemove += $script:dgIntuneManagerObjects.SelectedItem
    }

    if($policiesToRemove.Count -eq 0)
    {
        $script:UIProvider.ShowMessageBox("No object selected`n`nSelect items you want to delete", "Error", "OK", "Error")
        return
    }

    if(($script:UIProvider.ShowMessageBox("Are you sure you want to delete $($policiesToRemove.Count) object(s)?`n`nEnvironment: $($script:OrganizationName)", "Delete Objects?", "YesNo", "Warning")) -ne "Yes")
    {
        return
    }

    $deletedPolicies = $policiesToRemove | Remove-GraphPolicy

    if($deletedPolicies) {

        $script:dgIntuneManagerObjects.ItemsSource

        $deletedPolicies | ForEach-Object {
            try {
                $index = [Array]::IndexOf($script:dgIntuneManagerObjects.ItemsSource.Id, $_.Id)
                if($index -gt -1) {
                    ($script:dgIntuneManagerObjects.ItemsSource).RemoveAt($index)
                }
            }
            catch { } 
        }
    }
    
    Write-Status ""
}

