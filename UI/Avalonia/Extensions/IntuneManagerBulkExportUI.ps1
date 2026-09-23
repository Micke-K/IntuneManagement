# Avalonia port of the Bulk Export dialog from
# UI/WPF/Extensions/IntuneManagerUI.ps1 (Show-GraphBulkExportForm + helpers).
# New file because Bulk Export is a self-contained subsystem and the WPF
# original sits in the giant IntuneManagerUI file we're trying not to grow
# further (architecture rule R9).

function Update-BulkExportObjectList
{
    if (-not $script:dgBulkExportObjects) { return }

    $rows = [System.Collections.Generic.List[BulkExportRowItem]]::new()

    if ($script:bulkExportMode -eq "Type") {
        $exportableGroupIds = @($script:bulkExportExportableGroups | ForEach-Object { $_.Id })
        $sortedTypes = $script:IntuneTypes |
            Where-Object { $_.PolicyGroup -and ($exportableGroupIds -contains $_.PolicyGroup.Id) } |
            Sort-Object Title
        foreach ($intuneType in $sortedTypes) {
            $rows.Add([BulkExportRowItem]@{
                Title       = [string]$intuneType.Title
                Selected    = $true
                ObjectGroup = $null
                ObjectType  = $intuneType
            })
        }
    } else {
        $sortedGroups = $script:bulkExportExportableGroups | Sort-Object Title
        foreach ($intuneGroup in $sortedGroups) {
            $defaultSelected = ?? $intuneGroup.BulkExport $true
            $rows.Add([BulkExportRowItem]@{
                Title       = [string]$intuneGroup.Title
                Selected    = [bool]$defaultSelected
                ObjectGroup = $intuneGroup
                ObjectType  = $null
            })
        }
    }

    $script:bulkExportRows = @($rows)
    $script:dgBulkExportObjects.ItemsSource = $script:bulkExportRows
}

function Get-BulkExportSelectedIds
{
    if ($null -eq $script:bulkExportRows) { return @() }
    if ($script:bulkExportMode -eq "Type") {
        return @($script:bulkExportRows |
            Where-Object { $_.Selected -and $_.ObjectType -and $_.ObjectType.Id } |
            ForEach-Object { $_.ObjectType.Id })
    }
    return @($script:bulkExportRows |
        Where-Object { $_.Selected -and $_.ObjectGroup -and $_.ObjectGroup.Id } |
        ForEach-Object { $_.ObjectGroup.Id })
}

function Show-GraphBulkExportForm
{
    # CORRECTION (2026-08-25): the note that used to sit here had it exactly
    # backwards. Inside Avalonia handlers, $script: variables and bare
    # module-function lookups DO resolve - it is the function-LOCAL captures
    # that are stripped by ConvertTo-AvaloniaEventScriptBlock. Handler state
    # therefore lives in $script:_bulkExportFormState.
    $ui = $script:UIProvider

    $form = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/BulkExportForm.axaml'))
    if (-not $form) { return }

    $hostType = Get-AvaloniaHost
    $script:dgBulkExportObjects = $hostType::FindByName($form, 'dgObjectsToExport')

    $txtPath        = $hostType::FindByName($form, 'txtExportPath')
    $btnBrowse      = $hostType::FindByName($form, 'browseExportPath')
    $txtFilter      = $hostType::FindByName($form, 'txtExportNameFilter')
    $chkAssign      = $hostType::FindByName($form, 'chkExportAssignments')
    $chkCompany     = $hostType::FindByName($form, 'chkAddCompanyName')
    $txtNested      = $hostType::FindByName($form, 'txtNestedGroupLevels')
    $rbGroup        = $hostType::FindByName($form, 'rbBulkExportViewGroup')
    $rbType         = $hostType::FindByName($form, 'rbBulkExportViewType')
    $btnExport      = $hostType::FindByName($form, 'btnExport')
    $btnClose       = $hostType::FindByName($form, 'btnClose')

    $exportSettings = [IntuneManagerExportSettings]::new()

    # Handlers are re-bound to the module and lose function-local captures, so
    # every control they touch lives here instead (same pattern as Bulk
    # Compare). Without this the Export, Close, Escape and Browse handlers all
    # threw on their first line and the dialog was unusable.
    $script:_bulkExportFormState = [ordered]@{
        TxtPath = $null; TxtFilter = $null; ChkAssign = $null; ChkCompany = $null
        TxtNested = $null; HeaderCb = $null
        ExportSettings = $exportSettings
    }

    # Register dynamic export-properties for every exportable policy type up
    # front so toggling Group/API doesn't re-register and we never miss a type
    # that only appears in the other view.
    $script:bulkExportExportableGroups = @($script:IntuneGroups | Where-Object {
        $_.Title -and (-not ($_.ShowButtons -is [Object[]]) -or $_.ShowButtons -contains "Export")
    })
    foreach ($intuneGroup in $script:bulkExportExportableGroups) {
        $intuneGroup.PolicyTypes | Add-IntuneManagerExportProperties -ExportSettings $exportSettings
        $intuneGroup.PolicyTypes | Add-IntuneManagerExportUIExtensions -Form $form
    }

    # Push current values into controls. IntuneManagerExportSettings has no
    # INotifyPropertyChanged so two-way binding would render but never
    # round-trip — read back imperatively on Export click.
    if ($txtPath)    { $txtPath.Text       = [string]$exportSettings.ExportFolder }
    if ($txtFilter)  { $txtFilter.Text     = [string]$exportSettings.Filter }
    if ($chkAssign)  { $chkAssign.IsChecked = [bool]$exportSettings.ExportAssignments }
    if ($chkCompany) { $chkCompany.IsChecked = [bool]$exportSettings.AddCompanyName }
    if ($txtNested)  { $txtNested.Text      = [string]$exportSettings.ExportNestedGroupLevels }

    $script:_bulkExportFormState.TxtPath     = $txtPath
    $script:_bulkExportFormState.TxtFilter   = $txtFilter
    $script:_bulkExportFormState.ChkAssign   = $chkAssign
    $script:_bulkExportFormState.ChkCompany  = $chkCompany
    $script:_bulkExportFormState.TxtNested   = $txtNested

    $script:bulkExportMode = "Group"
    Update-BulkExportObjectList

    if ($btnBrowse) {
        $btnBrowse.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkExportFormState
            if (-not $st) { return }
            $folder = $script:UIProvider.ShowFolderPicker(([string]$st.TxtPath.Text), 'Select root folder for export')
            if ($folder) {
                $st.ExportSettings.ExportFolder = $folder
                if ($st.TxtPath) { $st.TxtPath.Text = $folder }
            }
        }))
    }

    # Select/deselect-all column header — initialised here so the per-mode
    # rebind below can simply flip it back on with $headerCb.IsChecked = $true.
    $headerCb = Initialize-AvaloniaGridSelectAllHeader -Grid $script:dgBulkExportObjects -BindingProperty 'Selected' -InitiallyChecked $true
    $script:_bulkExportFormState.HeaderCb = $headerCb

    if ($rbGroup) {
        $rbGroup.add_IsCheckedChanged((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            if (-not $s.IsChecked) { return }
            $script:bulkExportMode = "Group"
            Update-BulkExportObjectList
            $st = $script:_bulkExportFormState
            if ($st -and $st.HeaderCb) { $st.HeaderCb.IsChecked = $true }
        }))
    }
    if ($rbType) {
        $rbType.add_IsCheckedChanged((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            if (-not $s.IsChecked) { return }
            $script:bulkExportMode = "Type"
            Update-BulkExportObjectList
            $st = $script:_bulkExportFormState
            if ($st -and $st.HeaderCb) { $st.HeaderCb.IsChecked = $true }
        }))
    }



    # NOTE: the "Save settings for batch job" button was removed 2026-06-12 per
    # user decision — scheduled/CI runs use the public drivers directly
    # (Start-GraphBulkExport -SettingsFile files can still be produced with
    # Save-GraphBulkExportSettings from a script).

    if ($btnExport) {
        $btnExport.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkExportFormState
            if (-not $st) { return }
            Set-BulkExportSettingsFromForm
            $exportSettings = $st.ExportSettings

            $selection = Get-BulkExportSelectedIds
            $unit = if ($script:bulkExportMode -eq "Type") { "policy type" } else { "object group" }

            if ($selection.Count -eq 0) {
                $ui.ShowMessageBox("Select at least one $unit to export.", "Bulk Export", "OK", "Warning") | Out-Null
                return
            }

            if ([string]::IsNullOrWhiteSpace($exportSettings.ExportFolder)) {
                $ui.ShowMessageBox("Select an export folder.", "Bulk Export", "OK", "Warning") | Out-Null
                return
            }

            try { $resolvedPath = [IO.Path]::GetFullPath($exportSettings.ExportFolder) }
            catch {
                $ui.ShowMessageBox("Invalid export folder path: $($_.Exception.Message)", "Bulk Export", "OK", "Error") | Out-Null
                return
            }
            $exportSettings.ExportFolder = $resolvedPath
            Write-Log "Bulk export starting. Folder: $resolvedPath"

            $startParams = @{ ExportSettings = $exportSettings }
            if ($script:bulkExportMode -eq "Type") {
                $startParams.PolicyType = $selection
            } else {
                $startParams.PolicyGroup = $selection
            }

            try {
                $summary = Start-GraphBulkExport @startParams
                Write-Status ""
                $msg = ("Exported {0} policies across {1} types ({2} failed) in {3:hh\:mm\:ss}.`n`nFolder:`n{4}" -f `
                    $summary.Policies, $summary.Types, $summary.Failed, $summary.Duration, $resolvedPath)
                $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
                if ($unknownNote) { $msg += "`n`n$unknownNote" }
                $ui.ShowMessageBox($msg, "Bulk Export", "OK", "Information") | Out-Null
            } catch {
                Write-LogError "Bulk export failed" $_.Exception
                Write-Status ""
                $ui.ShowMessageBox("Bulk export failed: $($_.Exception.Message)", "Bulk Export", "OK", "Error") | Out-Null
            }

            $script:dgBulkExportObjects = $null
            $script:bulkExportRows = $null
            $script:_bulkExportFormState = $null
            Close-TopModalObject
        }))
    }

    if ($btnClose) {
        $btnClose.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $script:dgBulkExportObjects = $null
            $script:bulkExportRows = $null
            $script:_bulkExportFormState = $null
            Close-TopModalObject
        }))
    }

    $form.add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            $script:dgBulkExportObjects = $null
            $script:bulkExportRows = $null
            $script:_bulkExportFormState = $null
            Close-TopModalObject
            $e.Handled = $true
        }
    }))

    $ui.ShowModalForm("Bulk Export", $form, $true)
}

function Set-BulkExportSettingsFromForm
{
    # Pulls the live form values into the settings object. A module function
    # rather than a captured scriptblock: handlers cannot reach captures.
    $st = $script:_bulkExportFormState
    if (-not $st -or -not $st.ExportSettings) { return }
    $s = $st.ExportSettings

    if ($st.TxtPath) {
        $raw = [string]$st.TxtPath.Text
        if ($raw) { $raw = $raw.Trim().Trim('"').Trim("'") }
        $s.ExportFolder = $raw
    }
    if ($st.TxtFilter)  { $s.Filter            = [string]$st.TxtFilter.Text }
    if ($st.ChkAssign)  { $s.ExportAssignments = [bool]$st.ChkAssign.IsChecked }
    if ($st.ChkCompany) { $s.AddCompanyName    = [bool]$st.ChkCompany.IsChecked }
    if ($st.TxtNested) {
        $parsed = 0
        if ([int]::TryParse([string]$st.TxtNested.Text, [ref]$parsed) -and $parsed -ge 1) {
            $s.ExportNestedGroupLevels = $parsed
        }
    }
}