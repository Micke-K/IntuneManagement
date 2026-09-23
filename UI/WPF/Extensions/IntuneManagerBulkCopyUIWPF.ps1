# Bulk Copy form (ported from the original project's Copy extension).
# Thin WPF wrapper around the public Start-GraphBulkCopy driver — the UI only
# collects the two name patterns and the selected policy types (R10).

function Show-GraphBulkCopyForm
{
    $script:bulkCopyForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\BulkCopy.xaml"), $true)
    if(-not $script:bulkCopyForm) { return }

    $script:dgBulkCopyObjects = $script:bulkCopyForm.FindName("dgBulkCopyObjects")

    $script:UIProvider.SetXamlProperty($script:bulkCopyForm, "txtCopyFromPattern", "Text", (Get-SettingStoreValue "Copy" "CopyFromPattern"))
    $script:UIProvider.SetXamlProperty($script:bulkCopyForm, "txtCopyToPattern", "Text", (Get-SettingStoreValue "Copy" "CopyToPattern"))

    # Rows are policy TYPES (matching the original Bulk Copy, which listed every
    # type in the active view), all selected by default.
    $script:copyObjects = @()
    foreach($intuneType in $script:IntuneTypes)
    {
        if(-not $intuneType.Title) { continue }
        if($intuneType.ShowButtons -is [Object[]] -and $intuneType.ShowButtons -notcontains "Copy") { continue }

        $script:copyObjects += New-Object PSObject -Property @{
            Title      = $intuneType.Title
            Selected   = $true
            ObjectType = $intuneType
        }
    }
    $script:copyObjects = @($script:copyObjects | Sort-Object -Property Title)

    $column = Get-GridCheckboxColumn "Selected"
    $script:dgBulkCopyObjects.Columns.Add($column)
    $column.Header.IsChecked = $true
    $column.Header.add_Click({
        foreach($Item in $script:dgBulkCopyObjects.ItemsSource)
        {
            $Item.Selected = $this.IsChecked
        }
        $script:dgBulkCopyObjects.Items.Refresh()
    })

    $binding = [System.Windows.Data.Binding]::new("Title")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Object type"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:dgBulkCopyObjects.Columns.Add($column)

    $script:dgBulkCopyObjects.ItemsSource = $script:copyObjects

    $script:UIProvider.AddXamlEvent($script:bulkCopyForm, "btnClose", "add_click", {
        $script:bulkCopyForm = $null
        Show-ModalObject
    })

    $script:UIProvider.AddXamlEvent($script:bulkCopyForm, "btnStartCopy", "add_click", {
        $copyFrom = ($script:UIProvider.GetXamlProperty($script:bulkCopyForm, "txtCopyFromPattern", "Text")).Trim()
        $copyTo   = ($script:UIProvider.GetXamlProperty($script:bulkCopyForm, "txtCopyToPattern", "Text")).Trim()

        if(-not $copyFrom -or -not $copyTo)
        {
            $script:UIProvider.ShowMessageBox("Both name patterns must be specified", "Error", "OK", "Error")
            return
        }

        $selectedTypes = @($script:copyObjects | Where-Object Selected -eq $true)
        if($selectedTypes.Count -eq 0)
        {
            $script:UIProvider.ShowMessageBox("No object types selected.`n`nSelect the types you want to copy.", "Error", "OK", "Error")
            return
        }

        Save-SettingStoreValue "Copy" "CopyFromPattern" $copyFrom
        Save-SettingStoreValue "Copy" "CopyToPattern" $copyTo

        Write-Status "Copy objects" -Block
        try
        {
            $summary = Start-GraphBulkCopy -CopyFromPattern $copyFrom -CopyToPattern $copyTo `
                -PolicyType @($selectedTypes | ForEach-Object { $_.ObjectType.Id })
            Write-Status $null
            $msg = ("Copied {0} object(s) across {1} type(s) ({2} skipped because the target name exists) in {3:hh\:mm\:ss}." -f `
                $summary.Copied, $summary.Types, $summary.Skipped, $summary.Duration)
            $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
            if($unknownNote) { $msg += "`n`n$unknownNote" }
            $script:UIProvider.ShowMessageBox($msg, "Bulk Copy", "OK", "Information") | Out-Null
        }
        catch
        {
            Write-Status $null
            Write-LogError "Bulk copy failed" $_.Exception
            $script:UIProvider.ShowMessageBox("Bulk copy failed:`n`n$($_.Exception.Message)", "Bulk Copy", "OK", "Error") | Out-Null
        }

        # Refresh the current view so any copies in the active type show up.
        Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject | Out-Null
    })

    $script:UIProvider.ShowModalForm("Bulk Copy Objects", $script:bulkCopyForm, $true)
}
