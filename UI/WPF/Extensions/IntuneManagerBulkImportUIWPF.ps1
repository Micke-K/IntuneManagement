# Bulk Import form.
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Show-GraphBulkImportForm
{
    $script:bulkImportForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\BulkImportForm.xaml"), $true)
    if(-not $script:bulkImportForm) { return }

    $script:dgObjectsToImport = $script:bulkImportForm.FindName("dgObjectsToImport")

    $importSettings = [IntuneManagerImportSettings]::new()
    $script:bulkImportForm.DataContext = $importSettings

    $script:UIProvider.SetXamlProperty($script:bulkImportForm, "txtImportPath",          "Text",          $importSettings.ImportFolder)
    $script:UIProvider.SetXamlProperty($script:bulkImportForm, "chkImportAssignments",   "IsChecked",     $importSettings.ImportAssignments)
    $script:UIProvider.SetXamlProperty($script:bulkImportForm, "chkImportScopes",        "IsChecked",     $importSettings.ImportScopeTags)
    $script:UIProvider.SetXamlProperty($script:bulkImportForm, "chkReplaceDependencyIDs", "IsChecked",    $importSettings.ReplaceDependencyIDs)
    $script:UIProvider.SetXamlProperty($script:bulkImportForm, "cbImportType",           "ItemsSource",   $script:IntuneManagerImportOptions)
    $script:UIProvider.SetXamlProperty($script:bulkImportForm, "cbImportType",           "SelectedValue", $importSettings.ImportType)
    $script:UIProvider.SetXamlProperty($script:bulkImportForm, "lblImportType",          "Visibility",    "Visible")
    $script:UIProvider.SetXamlProperty($script:bulkImportForm, "cbImportType",           "Visibility",    "Visible")

    $script:importObjects = @()
    foreach($intuneGroup in $script:IntuneGroups)
    {
        if(-not $intuneGroup.Title) { continue }
        if($intuneGroup.ShowButtons -is [Object[]] -and $intuneGroup.ShowButtons -notcontains "Import") { continue }

        $script:importObjects += New-Object PSObject -Property @{
            Title       = $intuneGroup.Title
            Selected    = (?? $intuneGroup.BulkImport $true)
            ObjectGroup = $intuneGroup
        }

        $intuneGroup.PolicyTypes | Add-IntuneManagerImportUIExtensions -Form $script:bulkImportForm
    }

    $column = Get-GridCheckboxColumn "Selected"
    $script:dgObjectsToImport.Columns.Add($column)
    $column.Header.IsChecked = $true
    $column.Header.add_Click({
        foreach($Item in $script:dgObjectsToImport.ItemsSource)
        {
            $Item.Selected = $this.IsChecked
        }
        $script:dgObjectsToImport.Items.Refresh()
    })

    $binding = [System.Windows.Data.Binding]::new("Title")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Object type"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:dgObjectsToImport.Columns.Add($column)

    $script:dgObjectsToImport.ItemsSource = $script:importObjects

    if($importSettings.ImportFolder)
    {
        $migrationInfo, $sameTenant = Get-MigrationTableInfo $importSettings.ImportFolder $script:OrganizationId
        $script:UIProvider.SetXamlProperty($script:bulkImportForm, "chkReplaceDependencyIDs", "IsEnabled", ($sameTenant -eq $false))
        $script:UIProvider.SetXamlProperty($script:bulkImportForm, "chkReplaceDependencyIDs", "IsChecked", ($sameTenant -eq $false))
        $script:UIProvider.SetXamlProperty($script:bulkImportForm, "lblMigrationTableInfo",   "Content",   $migrationInfo)
    }

    $script:UIProvider.AddXamlEvent($script:bulkImportForm, "browseImportPath", "add_click", {
        $folder = Get-Folder ($script:UIProvider.GetXamlProperty($script:bulkImportForm, "txtImportPath", "Text")) "Select root folder for import"
        if($folder)
        {
            $script:UIProvider.SetXamlProperty($script:bulkImportForm, "txtImportPath", "Text", $folder)
            $script:bulkImportForm.DataContext.ImportFolder = $folder
            $migrationInfo, $sameTenant = Get-MigrationTableInfo $folder $script:OrganizationId
            $script:UIProvider.SetXamlProperty($script:bulkImportForm, "chkReplaceDependencyIDs", "IsEnabled", ($sameTenant -eq $false))
            $script:UIProvider.SetXamlProperty($script:bulkImportForm, "chkReplaceDependencyIDs", "IsChecked", ($sameTenant -eq $false))
            $script:UIProvider.SetXamlProperty($script:bulkImportForm, "lblMigrationTableInfo",   "Content",   $migrationInfo)
        }
    })

    $script:UIProvider.AddXamlEvent($script:bulkImportForm, "btnClose", "add_click", {
        $script:bulkImportForm = $null
        Show-ModalObject
    })

    $script:UIProvider.AddXamlEvent($script:bulkImportForm, "btnImport", "add_click", {
        $importFolder = Expand-FileName ($script:UIProvider.GetXamlProperty($script:bulkImportForm, "txtImportPath", "Text"))

        if(-not [IO.Directory]::Exists($importFolder))
        {
            $script:UIProvider.ShowMessageBox("Import folder not found:`n$importFolder", "Error", "OK", "Error")
            return
        }

        $selectedGroups = @($script:importObjects | Where-Object Selected -eq $true)
        if($selectedGroups.Count -eq 0)
        {
            $script:UIProvider.ShowMessageBox("No object types selected.", "Error", "OK", "Error")
            return
        }

        Save-SettingStoreValue "" "LastUsedRoot" $importFolder

        $importType = $script:UIProvider.GetXamlProperty($script:bulkImportForm, "cbImportType", "SelectedValue")
        $nameFilter = ($script:UIProvider.GetXamlProperty($script:bulkImportForm, "txtImportNameFilter", "Text")).Trim()

        # The UI is a thin caller of the public driver (R10) — the same import
        # runs headless via Start-GraphBulkImport, which also persists the
        # checkbox settings (the import pipeline reads them via Get-SettingValue)
        # and honours the ClearCacheBeforeExportImport setting.
        Write-Status "Bulk import" -Block
        try
        {
            $summary = Start-GraphBulkImport -ImportFolder $importFolder -Filter $nameFilter `
                -PolicyGroup @($selectedGroups | ForEach-Object { $_.ObjectGroup.ID }) `
                -ImportAssignments ([bool]($script:UIProvider.GetXamlProperty($script:bulkImportForm, "chkImportAssignments",    "IsChecked"))) `
                -ImportScopeTags ([bool]($script:UIProvider.GetXamlProperty($script:bulkImportForm, "chkImportScopes",         "IsChecked"))) `
                -ReplaceDependencyIDs ([bool]($script:UIProvider.GetXamlProperty($script:bulkImportForm, "chkReplaceDependencyIDs", "IsChecked"))) `
                -ImportType $importType

            $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
            if($summary.Imported -eq 0)
            {
                $msg = "No objects were imported.`nVerify the import folder and exported files."
                if($unknownNote) { $msg += "`n`n$unknownNote" }
                $script:UIProvider.ShowMessageBox($msg, "Import", "OK", "Warning")
            }
            else
            {
                # Success shows no dialog, so a group id that matched nothing gets its own.
                if($unknownNote) { $script:UIProvider.ShowMessageBox($unknownNote, "Import", "OK", "Warning") | Out-Null }
                Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject | Out-Null
            }
        }
        catch
        {
            Write-LogError "Bulk import failed" $_.Exception
            $script:UIProvider.ShowMessageBox("Bulk import failed:`n`n$($_.Exception.Message)", "Import", "OK", "Error") | Out-Null
        }

        Write-Status ""
    })

    $script:UIProvider.ShowModalForm("Bulk Import", $script:bulkImportForm, $true)
}

