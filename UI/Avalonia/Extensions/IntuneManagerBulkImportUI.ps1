# Avalonia port of the Bulk Import dialog from
# UI/WPF/Extensions/IntuneManagerUI.ps1 (Show-GraphBulkImportForm).
# New file because Bulk Import is a self-contained subsystem and the WPF
# original sits in the giant IntuneManagerUI file we're trying not to grow
# further (architecture rule R9).
#
# Slice 4b: core Bulk Import form. Per-policy-type Add-IntuneManagerImportUIExtensions
# wrapper is invoked but currently a no-op — no class extension defines
# AddUIImportExtensions in either tree.
#
# Thin caller of the public Start-GraphBulkImport driver (R10). The old
# "Save settings for batch job" button was removed 2026-06-12 per user
# decision — scheduled runs call the driver directly.

function Update-BulkImportObjectList
{
    if (-not $script:dgBulkImportObjects) { return }

    $rows = [System.Collections.Generic.List[BulkImportRowItem]]::new()
    $sortedGroups = $script:bulkImportImportableGroups | Sort-Object Title
    foreach ($intuneGroup in $sortedGroups) {
        $defaultSelected = ?? $intuneGroup.BulkImport $true
        $rows.Add([BulkImportRowItem]@{
            Title       = [string]$intuneGroup.Title
            Selected    = [bool]$defaultSelected
            ObjectGroup = $intuneGroup
        })
    }

    $script:bulkImportRows = @($rows)
    $script:dgBulkImportObjects.ItemsSource = $script:bulkImportRows
}

function Show-GraphBulkImportForm
{
    $ui = $script:UIProvider
    $form = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/BulkImportForm.axaml'))
    if (-not $form) { return }

    $hostType = Get-AvaloniaHost
    $script:dgBulkImportObjects = $hostType::FindByName($form, 'dgObjectsToImport')

    $txtPath        = $hostType::FindByName($form, 'txtImportPath')
    $btnBrowse      = $hostType::FindByName($form, 'browseImportPath')
    $txtFilter      = $hostType::FindByName($form, 'txtImportNameFilter')
    $lblMigInfo     = $hostType::FindByName($form, 'lblMigrationTableInfo')
    $chkScopes      = $hostType::FindByName($form, 'chkImportScopes')
    $chkAssign      = $hostType::FindByName($form, 'chkImportAssignments')
    $chkReplaceDeps = $hostType::FindByName($form, 'chkReplaceDependencyIDs')
    $cbImportType   = $hostType::FindByName($form, 'cbImportType')

    $btnImport      = $hostType::FindByName($form, 'btnImport')
    $btnClose       = $hostType::FindByName($form, 'btnClose')

    $importSettings = [IntuneManagerImportSettings]::new()

    # Handlers lose function-local captures when module-bound, so every
    # control they touch lives in module scope (Bulk Compare pattern).
    $script:_bulkImportFormState = [ordered]@{
        TxtPath        = $txtPath
        TxtFilter      = $txtFilter
        ChkScopes      = $chkScopes
        ChkAssign      = $chkAssign
        ChkReplaceDeps = $chkReplaceDeps
        LblMigInfo     = $lblMigInfo
        CbImportType   = $cbImportType
        ImportSettings = $importSettings
    }

    $script:bulkImportImportableGroups = @($script:IntuneGroups | Where-Object {
        $_.Title -and (-not ($_.ShowButtons -is [Object[]]) -or $_.ShowButtons -contains "Import")
    })

    # Per-type UI extensions (no-op today — no class defines
    # AddUIImportExtensions — but matches the WPF call site so future
    # extensions hook in automatically).
    foreach ($intuneGroup in $script:bulkImportImportableGroups) {
        $intuneGroup.PolicyTypes | Add-IntuneManagerImportUIExtensions -Form $form
    }

    # Push current values into controls. IntuneManagerImportSettings has no
    # INotifyPropertyChanged so two-way binding would render but never
    # round-trip — read back imperatively on Import click.
    if ($txtPath)        { $txtPath.Text          = [string]$importSettings.ImportFolder }
    if ($chkScopes)      { $chkScopes.IsChecked   = [bool]$importSettings.ImportScopeTags }
    if ($chkAssign)      { $chkAssign.IsChecked   = [bool]$importSettings.ImportAssignments }
    if ($chkReplaceDeps) { $chkReplaceDeps.IsChecked = [bool]$importSettings.ReplaceDependencyIDs }

    # Import-type ComboBox: NameValueObject -> SettingsListItem (same convention
    # as Slice 3e single-policy import).
    $importTypeItems = @()
    foreach ($opt in $script:IntuneManagerImportOptions) {
        $importTypeItems += [SettingsListItem]@{
            Name  = [string]$opt.Name
            Value = [string]$opt.Value
        }
    }
    if ($cbImportType) {
        $cbImportType.ItemsSource = $importTypeItems
        $current = $importTypeItems | Where-Object { $_.Value -eq [string]$importSettings.ImportType } | Select-Object -First 1
        if (-not $current -and $importTypeItems.Count -gt 0) { $current = $importTypeItems[0] }
        if ($current) { $cbImportType.SelectedItem = $current }
    }

    Update-BulkImportObjectList

    # Migration-info refresh: reads the Groups/MigrationTable.json under the
    # given folder, sets sameTenant -> ReplaceDependencyIDs default. Same
    # behaviour as Slice 3e's loadFromFolder, minus the policy-file scan.
    # Runs here (not earlier) because it reads the module-scope state.
    if ($importSettings.ImportFolder) {
        Update-BulkImportMigrationInfo -Folder $importSettings.ImportFolder
    }

    if ($btnBrowse) {
        $btnBrowse.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkImportFormState
            if (-not $st) { return }
            $folder = $script:UIProvider.ShowFolderPicker(([string]$st.TxtPath.Text), 'Select root folder for import')
            if ($folder) {
                $st.ImportSettings.ImportFolder = $folder
                if ($st.TxtPath) { $st.TxtPath.Text = $folder }
                Update-BulkImportMigrationInfo -Folder $folder
            }
        }))
    }

    if ($cbImportType) {
        $cbImportType.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $st = $script:_bulkImportFormState
            if ($st -and $s.SelectedItem) {
                $st.ImportSettings.ImportType = [string]$s.SelectedItem.Value
            }
        }))
    }

    # Select/deselect-all moved into the DataGrid column header.
    Initialize-AvaloniaGridSelectAllHeader -Grid $script:dgBulkImportObjects -BindingProperty 'Selected' -InitiallyChecked $true | Out-Null

    if ($btnImport) {
        $btnImport.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkImportFormState
            if (-not $st) { return }
            $importSettings = $st.ImportSettings
            $importFolder = if ($st.TxtPath) { [string]$st.TxtPath.Text } else { '' }
            if ($importFolder) { $importFolder = $importFolder.Trim().Trim('"').Trim("'") }
            try { $importFolder = Expand-FileName $importFolder } catch { }

            if (-not [IO.Directory]::Exists($importFolder)) {
                $ui.ShowMessageBox("Import folder not found:`n$importFolder", "Bulk Import", "OK", "Error") | Out-Null
                return
            }
            $importSettings.ImportFolder = $importFolder

            if ($st.ChkScopes)      { $importSettings.ImportScopeTags      = [bool]$st.ChkScopes.IsChecked }
            if ($st.ChkAssign)      { $importSettings.ImportAssignments    = [bool]$st.ChkAssign.IsChecked }
            if ($st.ChkReplaceDeps) { $importSettings.ReplaceDependencyIDs = [bool]$st.ChkReplaceDeps.IsChecked }

            $selectedGroups = @($script:bulkImportRows | Where-Object { $_.Selected -and $_.ObjectGroup })
            if ($selectedGroups.Count -eq 0) {
                $ui.ShowMessageBox("No object types selected.", "Bulk Import", "OK", "Warning") | Out-Null
                return
            }

            Save-SettingStoreValue "" "LastUsedRoot" $importFolder

            $nameFilter = if ($st.TxtFilter) { ([string]$st.TxtFilter.Text).Trim() } else { '' }

            # The UI is a thin caller of the public driver (R10) — the same
            # import runs headless via Start-GraphBulkImport, which also
            # persists the checkbox settings (the import pipeline reads them
            # via Get-SettingValue) and honours ClearCacheBeforeExportImport.
            Write-Status "Bulk import" -Block
            $totalImported = 0
            try {
                $summary = Start-GraphBulkImport -ImportFolder $importFolder -Filter $nameFilter `
                    -PolicyGroup @($selectedGroups | ForEach-Object { $_.ObjectGroup.ID }) `
                    -ImportAssignments $importSettings.ImportAssignments `
                    -ImportScopeTags $importSettings.ImportScopeTags `
                    -ReplaceDependencyIDs $importSettings.ReplaceDependencyIDs `
                    -ImportType $importSettings.ImportType
                $totalImported = $summary.Imported
            }
            catch {
                Write-LogError "Bulk import failed" $_.Exception
                $ui.ShowMessageBox("Bulk import failed:`n`n$($_.Exception.Message)", "Bulk Import", "OK", "Error") | Out-Null
            }
            Write-Status ""

            $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
            if ($totalImported -eq 0) {
                $msg = "No objects were imported.`nVerify the import folder and exported files."
                if ($unknownNote) { $msg += "`n`n$unknownNote" }
                $ui.ShowMessageBox($msg, "Bulk Import", "OK", "Warning") | Out-Null
            } else {
                $msg = ("Imported {0} object(s) from:`n{1}" -f $totalImported, $importFolder)
                if ($unknownNote) { $msg += "`n`n$unknownNote" }
                $ui.ShowMessageBox($msg, "Bulk Import", "OK", "Information") | Out-Null
                if ($script:IntuneManagerSelectedObject) {
                    Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject | Out-Null
                }
            }

            $script:dgBulkImportObjects = $null
            $script:bulkImportRows = $null
            $script:_bulkImportFormState = $null
            Show-ModalObject
        }))
    }

    if ($btnClose) {
        $btnClose.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $script:dgBulkImportObjects = $null
            $script:bulkImportRows = $null
            $script:_bulkImportFormState = $null
            Show-ModalObject
        }))
    }

    $form.add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            $script:dgBulkImportObjects = $null
            $script:bulkImportRows = $null
            $script:_bulkImportFormState = $null
            Show-ModalObject
            $e.Handled = $true
        }
    }))

    $ui.ShowModalForm("Bulk Import", $form, $true)
}

function Update-BulkImportMigrationInfo
{
    # Reads MigrationTable.json for the folder and reflects same-tenant state
    # on the form. Module function, not a captured scriptblock - handlers
    # cannot reach captures.
    param([string]$Folder)

    $st = $script:_bulkImportFormState
    if (-not $st) { return }

    if (-not $Folder -or -not [IO.Directory]::Exists($Folder)) {
        if ($st.LblMigInfo) { $st.LblMigInfo.Text = '' }
        return
    }

    try {
        $migrationInfo, $sameTenant = Get-MigrationTableInfo $Folder $script:OrganizationId
        $st.ImportSettings.SameTenant = [bool]$sameTenant
        if ($st.LblMigInfo) { $st.LblMigInfo.Text = [string]$migrationInfo }
        if ($st.ChkReplaceDeps) {
            $st.ChkReplaceDeps.IsEnabled = (-not $sameTenant)
            $st.ChkReplaceDeps.IsChecked = (-not $sameTenant)
        }
    } catch {
        Write-LogError "Failed to read migration table info" $_.Exception
    }
}
