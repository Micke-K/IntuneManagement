# Bulk Delete form.
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Show-GraphBulkDeleteForm
{
    $script:bulkDeleteForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\BulkDeleteForm.xaml"), $true)
    if(-not $script:bulkDeleteForm) { return }

    $script:dgBulkDeleteObjects = $script:bulkDeleteForm.FindName("dgBulkDeleteObjects")

    $script:UIProvider.SetXamlProperty($script:bulkDeleteForm, "txtDeleteNameFilter", "Text", (Get-SettingStoreValue "" "DeleteNameFilter"))

    $script:deleteObjects = @()
    foreach($intuneGroup in $script:IntuneGroups)
    {
        if(-not $intuneGroup.Title) { continue }
        if($intuneGroup.ShowButtons -is [Object[]] -and $intuneGroup.ShowButtons -notcontains "Delete") { continue }

        $script:deleteObjects += New-Object PSObject -Property @{
            Title       = $intuneGroup.Title
            Selected    = $false
            ObjectGroup = $intuneGroup
        }
    }

    $column = Get-GridCheckboxColumn "Selected"
    $script:dgBulkDeleteObjects.Columns.Add($column)
    $column.Header.IsChecked = $false
    $column.Header.add_Click({
        foreach($Item in $script:dgBulkDeleteObjects.ItemsSource)
        {
            $Item.Selected = $this.IsChecked
        }
        $script:dgBulkDeleteObjects.Items.Refresh()
    })

    $binding = [System.Windows.Data.Binding]::new("Title")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Object type"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:dgBulkDeleteObjects.Columns.Add($column)

    $script:dgBulkDeleteObjects.ItemsSource = $script:deleteObjects

    $script:UIProvider.AddXamlEvent($script:bulkDeleteForm, "btnClose", "add_click", {
        $script:bulkDeleteForm = $null
        Show-ModalObject
    })

    $script:UIProvider.AddXamlEvent($script:bulkDeleteForm, "btnDelete", "add_click", {
        $selectedGroups = @($script:deleteObjects | Where-Object Selected -eq $true)

        if($selectedGroups.Count -eq 0)
        {
            $script:UIProvider.ShowMessageBox("No object types selected.`n`nSelect the types you want to delete.", "Error", "OK", "Error")
            return
        }

        $nameFilter = ($script:UIProvider.GetXamlProperty($script:bulkDeleteForm, "txtDeleteNameFilter", "Text")).Trim()
        Save-SettingStoreValue "" "DeleteNameFilter" $nameFilter

        $confirmMsg = "Are you sure you want to delete all objects of the selected type(s)?`n`n$($selectedGroups.Count) type(s) selected`n`nEnvironment: $script:OrganizationName"
        if($nameFilter) { $confirmMsg += "`nName filter: $nameFilter" }

        if(($script:UIProvider.ShowMessageBox($confirmMsg, "Delete Objects?", "YesNo", "Warning")) -ne "Yes")
        {
            return
        }

        # The UI is a thin caller of the public driver (R10) — the confirm
        # prompt above is the UI's job; the delete loop itself runs headless
        # via Start-GraphBulkDelete.
        Write-Status "Bulk delete" -Block
        try
        {
            $summary = Start-GraphBulkDelete -Filter $nameFilter `
                -PolicyGroup @($selectedGroups | ForEach-Object { $_.ObjectGroup.ID })
            # Delete shows nothing on plain success; an id that matched nothing is
            # the one outcome the user must not miss.
            $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
            if($unknownNote) {
                $script:UIProvider.ShowMessageBox($unknownNote, "Bulk Delete", "OK", "Warning") | Out-Null
            }
        }
        catch
        {
            Write-LogError "Bulk delete failed" $_.Exception
            $script:UIProvider.ShowMessageBox("Bulk delete failed:`n`n$($_.Exception.Message)", "Bulk Delete", "OK", "Error") | Out-Null
        }

        Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject | Out-Null
        Write-Status ""
    })

    $script:UIProvider.ShowModalForm("Bulk Delete", $script:bulkDeleteForm, $true)
}

#region Bulk Documentation
# Thin WPF wrapper around Start-GraphBulkDocumentation. The same public driver
# powers the silent batch path (Tests/Setup/Invoke-AutomationOrchestrator) so
# every code path goes through the same engine.

