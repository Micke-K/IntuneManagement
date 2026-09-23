# Avalonia port of the Bulk Delete dialog from
# UI/WPF/Extensions/IntuneManagerUI.ps1 (Show-GraphBulkDeleteForm).
# New file because Bulk Delete is a self-contained subsystem and the WPF
# original sits in the giant IntuneManagerUI file we're trying not to grow
# further (architecture rule R9).
#
# Slice 4c: core Bulk Delete form. Simpler than 4a/4b — no settings-driven
# defaults, just a name filter + per-group selection + a confirm dialog. The
# btnDelete handler is a thin caller of the public Start-GraphBulkDelete driver
# (R10); the confirm prompt is the UI's job. Event handlers use the module-scope
# $script:_* state pattern + bare module functions (no GetNewClosure / captured
# locals / $script:UIProvider) so they resolve correctly at click time
# (see [[avalonia-closure-dynamic-module]]).

function Update-BulkDeleteObjectList
{
    if (-not $script:dgBulkDeleteObjects) { return }

    $rows = [System.Collections.Generic.List[BulkDeleteRowItem]]::new()
    $sortedGroups = $script:bulkDeleteDeletableGroups | Sort-Object Title
    foreach ($intuneGroup in $sortedGroups) {
        $rows.Add([BulkDeleteRowItem]@{
            Title       = [string]$intuneGroup.Title
            Selected    = $false
            ObjectGroup = $intuneGroup
        })
    }

    $script:bulkDeleteRows = @($rows)
    $script:dgBulkDeleteObjects.ItemsSource = $script:bulkDeleteRows
}

function Show-GraphBulkDeleteForm
{
    $ui = $script:UIProvider
    $form = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/BulkDeleteForm.axaml'))
    if (-not $form) { return }

    $hostType = Get-AvaloniaHost
    $script:dgBulkDeleteObjects = $hostType::FindByName($form, 'dgBulkDeleteObjects')

    $txtFilter    = $hostType::FindByName($form, 'txtDeleteNameFilter')
    $btnDelete    = $hostType::FindByName($form, 'btnDelete')
    $btnClose     = $hostType::FindByName($form, 'btnClose')

    $script:bulkDeleteDeletableGroups = @($script:IntuneGroups | Where-Object {
        $_.Title -and (-not ($_.ShowButtons -is [Object[]]) -or $_.ShowButtons -contains "Delete")
    })

    if ($txtFilter) { $txtFilter.Text = [string](Get-SettingStoreValue "" "DeleteNameFilter") }
    $script:_bulkDeleteTxtFilter = $txtFilter

    Update-BulkDeleteObjectList

    # Select/deselect-all moved into the DataGrid column header.
    # Delete defaults to "nothing selected" to keep the destructive action
    # explicit, matching the AXAML IsChecked="False" on the header CheckBox.
    Initialize-AvaloniaGridSelectAllHeader -Grid $script:dgBulkDeleteObjects -BindingProperty 'Selected' -InitiallyChecked $false | Out-Null

    if ($btnDelete) {
        $btnDelete.add_Click({
            $txtFilter = $script:_bulkDeleteTxtFilter

            $selectedGroups = @($script:bulkDeleteRows | Where-Object { $_.Selected -and $_.ObjectGroup })
            if ($selectedGroups.Count -eq 0) {
                Show-MessageBox "No object types selected.`n`nSelect the types you want to delete." "Bulk Delete" "OK" "Error" | Out-Null
                return
            }

            $nameFilter = if ($txtFilter) { ([string]$txtFilter.Text).Trim() } else { '' }
            Save-SettingStoreValue "" "DeleteNameFilter" $nameFilter

            $confirmMsg = "Are you sure you want to delete all objects of the selected type(s)?`n`n$($selectedGroups.Count) type(s) selected`n`nEnvironment: $script:OrganizationName"
            if ($nameFilter) { $confirmMsg += "`nName filter: $nameFilter" }

            if ((Show-MessageBox $confirmMsg "Delete Objects?" "YesNo" "Warning") -ne "Yes") {
                return
            }

            # Thin caller of the public driver (R10) — the confirm prompt above is
            # the UI's job; the delete loop runs headless via Start-GraphBulkDelete.
            Write-Status "Bulk delete" -Block
            try {
                $summary = Start-GraphBulkDelete -Filter $nameFilter `
                    -PolicyGroup @($selectedGroups | ForEach-Object { $_.ObjectGroup.ID })
                # Delete shows nothing on plain success; an id that matched nothing is
                # the one outcome the user must not miss.
                $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
                if ($unknownNote) { Show-MessageBox $unknownNote "Bulk Delete" "OK" "Warning" | Out-Null }
            }
            catch {
                Write-LogError "Bulk delete failed" $_.Exception
                Show-MessageBox "Bulk delete failed:`n`n$($_.Exception.Message)" "Bulk Delete" "OK" "Error" | Out-Null
            }

            # Match WPF (IntuneManagerBulkDeleteUIWPF.ps1): refresh the main view's
            # object list and leave the Bulk Delete form open so the user can run
            # another delete or close it explicitly. btnClose / Escape tear down the
            # $script: state - don't null it here while the form is still live.
            if ($script:IntuneManagerSelectedObject) {
                Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject | Out-Null
            }
            Write-Status ""
        })
    }

    if ($btnClose) {
        $btnClose.add_Click({
            $script:dgBulkDeleteObjects = $null
            $script:bulkDeleteRows = $null
            $script:_bulkDeleteTxtFilter = $null
            Show-ModalObject
        })
    }

    $form.add_KeyDown({
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            $script:dgBulkDeleteObjects = $null
            $script:bulkDeleteRows = $null
            $script:_bulkDeleteTxtFilter = $null
            Show-ModalObject
            $e.Handled = $true
        }
    })

    $ui.ShowModalForm("Bulk Delete", $form, $true)
}
