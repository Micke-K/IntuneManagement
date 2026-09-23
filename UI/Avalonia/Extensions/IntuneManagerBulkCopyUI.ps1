# Avalonia Bulk Copy dialog (ported from the old project's Copy extension via
# the WPF UI/WPF/Extensions/IntuneManagerBulkCopyUIWPF.ps1). New file per
# architecture rule R9. Thin caller of the public Start-GraphBulkCopy driver
# (R10) — the UI only collects the two name patterns and the selected types.

function Show-GraphBulkCopyForm
{
    $ui = $script:UIProvider
    $form = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/BulkCopyForm.axaml'))
    if (-not $form) { return }

    $hostType = Get-AvaloniaHost
    $script:dgBulkCopyObjects = $hostType::FindByName($form, 'dgBulkCopyObjects')
    $script:txtCopyFromPattern = $hostType::FindByName($form, 'txtCopyFromPattern')
    $script:txtCopyToPattern   = $hostType::FindByName($form, 'txtCopyToPattern')
    $btnStartCopy = $hostType::FindByName($form, 'btnStartCopy')
    $btnClose     = $hostType::FindByName($form, 'btnClose')

    if ($script:txtCopyFromPattern) { $script:txtCopyFromPattern.Text = [string](Get-SettingStoreValue "Copy" "CopyFromPattern") }
    if ($script:txtCopyToPattern)   { $script:txtCopyToPattern.Text   = [string](Get-SettingStoreValue "Copy" "CopyToPattern") }

    # Rows are policy TYPES (matching the original Bulk Copy), all selected.
    $rows = [System.Collections.Generic.List[BulkCopyRowItem]]::new()
    $copyTypes = @($script:IntuneTypes | Where-Object {
        $_.Title -and (-not ($_.ShowButtons -is [Object[]]) -or $_.ShowButtons -contains "Copy")
    } | Sort-Object Title)
    foreach ($intuneType in $copyTypes) {
        $rows.Add([BulkCopyRowItem]@{
            Title      = [string]$intuneType.Title
            Selected   = $true
            ObjectType = $intuneType
        })
    }
    $script:bulkCopyRows = @($rows)
    if ($script:dgBulkCopyObjects) { $script:dgBulkCopyObjects.ItemsSource = $script:bulkCopyRows }

    Initialize-AvaloniaGridSelectAllHeader -Grid $script:dgBulkCopyObjects -BindingProperty 'Selected' -InitiallyChecked $true | Out-Null

    if ($btnStartCopy) {
        $btnStartCopy.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $copyFrom = if ($script:txtCopyFromPattern) { ([string]$script:txtCopyFromPattern.Text).Trim() } else { '' }
            $copyTo   = if ($script:txtCopyToPattern)   { ([string]$script:txtCopyToPattern.Text).Trim() }   else { '' }

            if (-not $copyFrom -or -not $copyTo) {
                $script:UIProvider.ShowMessageBox("Both name patterns must be specified", "Bulk Copy", "OK", "Error") | Out-Null
                return
            }

            $selectedTypes = @($script:bulkCopyRows | Where-Object { $_.Selected -and $_.ObjectType })
            if ($selectedTypes.Count -eq 0) {
                $script:UIProvider.ShowMessageBox("No object types selected.`n`nSelect the types you want to copy.", "Bulk Copy", "OK", "Error") | Out-Null
                return
            }

            Save-SettingStoreValue "Copy" "CopyFromPattern" $copyFrom
            Save-SettingStoreValue "Copy" "CopyToPattern" $copyTo

            Write-Status "Copy objects" -Block
            try {
                $summary = Start-GraphBulkCopy -CopyFromPattern $copyFrom -CopyToPattern $copyTo `
                    -PolicyType @($selectedTypes | ForEach-Object { $_.ObjectType.Id })
                Write-Status $null
                $msg = ("Copied {0} object(s) across {1} type(s) ({2} skipped because the target name exists) in {3:hh\:mm\:ss}." -f `
                    $summary.Copied, $summary.Types, $summary.Skipped, $summary.Duration)
                $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
                if ($unknownNote) { $msg += "`n`n$unknownNote" }
                $script:UIProvider.ShowMessageBox($msg, "Bulk Copy", "OK", "Information") | Out-Null
            }
            catch {
                Write-Status $null
                Write-LogError "Bulk copy failed" $_.Exception
                $script:UIProvider.ShowMessageBox("Bulk copy failed:`n`n$($_.Exception.Message)", "Bulk Copy", "OK", "Error") | Out-Null
            }

            if ($script:IntuneManagerSelectedObject) {
                Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject | Out-Null
            }
        }))
    }

    if ($btnClose) {
        $btnClose.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $script:dgBulkCopyObjects = $null
            $script:bulkCopyRows = $null
            Show-ModalObject
        }))
    }

    $form.add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            $script:dgBulkCopyObjects = $null
            $script:bulkCopyRows = $null
            Show-ModalObject
            $e.Handled = $true
        }
    }))

    $ui.ShowModalForm("Bulk Copy Objects", $form, $true)
}
