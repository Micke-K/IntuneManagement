# Avalonia port of the Export dialog from UI/WPF/Extensions/IntuneManagerUI.ps1
# (Show-IntuneManagerExportForm). New file because Export is a self-contained
# subsystem and we don't want to grow the giant IntuneManagerUI files further
# (architecture rule R9).
#
# Same closure-free design rationale as IntuneManagerImportUI.ps1: form state
# lives in $script:_imExportFormState so click handlers can be plain
# scriptblocks (no .GetNewClosure() — that puts the SB into a dynamic module
# with no module-private function resolution).

function Show-IntuneManagerExportForm
{
    if (-not $script:dgIntuneManagerObjects) { return }

    $ui = $script:UIProvider

    $exportForm = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/ExportForm.axaml'))
    if (-not $exportForm) { return }

    $hostType  = Get-AvaloniaHost
    $txtPath   = $hostType::FindByName($exportForm, 'txtExportPath')
    $browse    = $hostType::FindByName($exportForm, 'browseExportPath')
    $chkObj    = $hostType::FindByName($exportForm, 'chkAddObjectType')
    $chkComp   = $hostType::FindByName($exportForm, 'chkAddCompanyName')
    $chkAssign = $hostType::FindByName($exportForm, 'chkExportAssignments')
    $btnSel    = $hostType::FindByName($exportForm, 'btnExportSelected')
    $btnAll    = $hostType::FindByName($exportForm, 'btnExportAll')
    $btnCancel = $hostType::FindByName($exportForm, 'btnCancel')
    $lblSel    = $hostType::FindByName($exportForm, 'lblSelectedObject')

    # Build the settings instance and let the data-side hooks add their own
    # NoteProperties (per-type extra export options). These reach onto the
    # CLR class via Add-Member so we can still read them after Save().
    $exportSettings = [IntuneManagerExportSettings]::new()
    $policyTypes = Get-IntuneManagerSelectedPolicyTypes
    $policyTypes | Add-IntuneManagerExportProperties -ExportSettings $exportSettings
    $policyTypes | Add-IntuneManagerExportUIExtensions -Form $exportForm

    # Avalonia bindings would render against IntuneManagerExportSettings
    # but the class has no INotifyPropertyChanged, so two-way round-trip
    # silently fails. Push current values into controls and read them back
    # imperatively on click.
    if ($txtPath)   { $txtPath.Text     = [string]$exportSettings.ExportFolder }
    if ($chkObj)    { $chkObj.IsChecked    = [bool]$exportSettings.AddObjectType }
    if ($chkComp)   { $chkComp.IsChecked   = [bool]$exportSettings.AddCompanyName }
    if ($chkAssign) { $chkAssign.IsChecked = [bool]$exportSettings.ExportAssignments }

    # Selection summary label + Export Selected enable state mirror the WPF
    # behaviour: prefer IsSelected-checked rows over the highlighted row.
    # Scan the grid's CURRENT (filtered) ItemsSource - a row checked before the
    # filter was typed must not be exported while invisible.
    $checkedRows = @(@($script:dgIntuneManagerObjects.ItemsSource) | Where-Object { $_ -and $_.IsSelected })
    $highlighted = $script:dgIntuneManagerObjects.SelectedItem

    if ($lblSel) {
        if ($checkedRows.Count -gt 0) {
            $lblSel.Content = "$($checkedRows.Count) selected object(s)"
        } elseif ($highlighted) {
            $lblSel.Content = "Selected object: $($highlighted.Name)"
        }
    }
    if ($btnSel) {
        $btnSel.IsEnabled = ($checkedRows.Count -gt 0) -or ($null -ne $highlighted)
    }

    # Module-scope state container. Click handlers reference by name —
    # no .GetNewClosure() so they survive Avalonia delegate dispatch.
    $script:_imExportFormState = @{
        Settings    = $exportSettings
        TxtPath     = $txtPath
        ChkObj      = $chkObj
        ChkComp     = $chkComp
        ChkAssign   = $chkAssign
        CheckedRows = $checkedRows
        Highlighted = $highlighted
        Action      = $null
    }

    if ($browse) {
        $browse.add_Click({
            $st  = $script:_imExportFormState
            $ui2 = $script:UIProvider
            if (-not $st) { return }
            $folder = $ui2.ShowFolderPicker('Select root folder for export')
            if ($folder) {
                $st.Settings.ExportFolder = $folder
                if ($st.TxtPath) { $st.TxtPath.Text = $folder }
            }
        })
    }

    if ($btnSel) {
        $btnSel.add_Click({
            $st = $script:_imExportFormState
            if (-not $st) { return }
            $st.Action = 'Selected'
            # Keep the dialog open when the export bailed out on validation
            # (e.g. no folder picked) - it used to close and discard the input.
            $ranOk = $false
            try { $ranOk = [bool](Invoke-IntuneManagerExportExecute) } catch { Write-LogError "Export execution failed" $_.Exception }
            if (-not $ranOk) { return }
            $script:UIProvider.CloseTopModalObject()
            $script:_imExportFormState = $null
            if ($script:dgIntuneManagerObjects) { $script:dgIntuneManagerObjects.Focus() }
        })
    }
    if ($btnAll) {
        $btnAll.add_Click({
            $st = $script:_imExportFormState
            if (-not $st) { return }
            $st.Action = 'All'
            $ranOk = $false
            try { $ranOk = [bool](Invoke-IntuneManagerExportExecute) } catch { Write-LogError "Export execution failed" $_.Exception }
            if (-not $ranOk) { return }
            $script:UIProvider.CloseTopModalObject()
            $script:_imExportFormState = $null
            if ($script:dgIntuneManagerObjects) { $script:dgIntuneManagerObjects.Focus() }
        })
    }
    if ($btnCancel) {
        $btnCancel.add_Click({
            $script:UIProvider.CloseTopModalObject()
            $script:_imExportFormState = $null
            if ($script:dgIntuneManagerObjects) { $script:dgIntuneManagerObjects.Focus() }
        })
    }

    $exportForm.add_KeyDown({
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            $script:UIProvider.CloseTopModalObject()
            $script:_imExportFormState = $null
            $e.Handled = $true
        }
    })

    # ShowModalForm is non-blocking; export runs inside the button handlers via
    # Invoke-IntuneManagerExportExecute while the form state is live.
    $ui.ShowModalForm("Export $($script:IntuneManagerSelectedObject.Title) objects", $exportForm, $true)
}

function Invoke-IntuneManagerExportExecute
{
    # Runs the export for the live $script:_imExportFormState. Called from the
    # Export button handlers (ShowModalForm does not block).
    $st = $script:_imExportFormState
    if (-not $st -or -not $st.Action -or $st.Action -eq 'Cancel') { return }
    $ui             = $script:UIProvider
    $exportSettings = $st.Settings
    $txtPath        = $st.TxtPath
    $chkObj         = $st.ChkObj
    $chkComp        = $st.ChkComp
    $chkAssign      = $st.ChkAssign
    $checkedRows    = @($st.CheckedRows)
    $highlighted    = $st.Highlighted

    # Pull form state back into the settings object before exporting.
    if ($txtPath)   { $exportSettings.ExportFolder      = [string]$txtPath.Text }
    if ($chkObj)    { $exportSettings.AddObjectType     = [bool]$chkObj.IsChecked }
    if ($chkComp)   { $exportSettings.AddCompanyName    = [bool]$chkComp.IsChecked }
    if ($chkAssign) { $exportSettings.ExportAssignments = [bool]$chkAssign.IsChecked }

    if ([string]::IsNullOrWhiteSpace($exportSettings.ExportFolder)) {
        $ui.ShowMessageBox("Pick an export root folder first.", "Export", "OK", "Error") | Out-Null
        return $false
    }

    $rowsToExport = @()
    if ($st.Action -eq 'Selected') {
        if ($checkedRows.Count -gt 0) {
            $rowsToExport = $checkedRows
        } elseif ($highlighted) {
            $rowsToExport = @($highlighted)
        }
    } else {
        # Export All — use the current ItemsSource (filtered) to match what
        # the user sees.
        if ($script:dgIntuneManagerObjects.ItemsSource) {
            $rowsToExport = @($script:dgIntuneManagerObjects.ItemsSource)
        }
    }

    if ($rowsToExport.Count -eq 0) {
        Write-Log "No rows to export" 2
        return $false
    }

    Write-Status "Export $($script:IntuneManagerSelectedObject.Title)"

    # Export-GraphPolicy expects IntunePolicyBase, not the CLR row wrapper —
    # pull .Source off each row before piping.
    $policiesToExport = @($rowsToExport | ForEach-Object { $_.Source } | Where-Object { $_ })

    try {
        $policiesToExport | Export-GraphPolicy -ExportSettings $exportSettings
    } catch {
        Write-LogError "Export-GraphPolicy failed" $_.Exception
        $ui.ShowMessageBox("Export failed: $($_.Exception.Message)", "Error", "OK", "Error") | Out-Null
        Write-Status ""
        return $false
    }

    $exportSettings.Save()

    Write-Status ""
    if ($script:dgIntuneManagerObjects) { $script:dgIntuneManagerObjects.Focus() }
    return $true
}
