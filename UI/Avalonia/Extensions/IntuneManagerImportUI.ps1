# Avalonia port of the Import dialog from UI/WPF/Extensions/IntuneManagerUI.ps1
# (Show-IntuneManagerImportForm + supporting helpers). New file because
# Import is a self-contained subsystem and the WPF original sits in the
# giant CoreUI/IntuneManagerUI files we're trying not to grow further
# (architecture rule R9).
#
# CLOSURE-FREE design (after 2026-05-30 fix): PowerShell's `.GetNewClosure()`
# puts a scriptblock into a brand-new DynamicModule that has no access to
# IntuneManagement's module-private functions (Save-SettingStoreValue,
# Get-MigrationTableInfo, Write-Log, ...). `ConvertTo-AvaloniaEventScriptBlock`
# then calls NewBoundScriptBlock which strips the closure entirely.
# Both effects = click handlers crash with "property not found" /
# "term not recognized". To avoid this, the form's state lives in
# `$script:_imImportFormState`, the helper logic is exposed as real
# module-scope functions (`Invoke-IntuneManagerImportLoad`,
# `Update-IntuneManagerImportMatch`), and click handlers are plain
# scriptblocks (no .GetNewClosure()) that reference the script-scope state
# and call the helper functions by name.
#
# The match-resolution helpers (Resolve-IntuneImportUpdateTarget,
# New-IntuneImportMatchResult, Get-IntuneImportPolicyReferenceTokens,
# Normalize-IntuneImportPolicyName) live in Internal/IntuneManager.ps1 and
# are shared with the WPF tree.

function Update-IntuneManagerImportMatch
{
    # Operates on $script:_imImportFormState. Recomputes ImportAction /
    # ImportMatch / ImportMatchStrategy / ImportTarget on every visible
    # row, then forces a DataGrid refresh (no INotifyPropertyChanged on
    # IntuneImportRowItem). Called from cbImportType selection-changed and
    # from the load helper after a fresh load.
    $st = $script:_imImportFormState
    if (-not $st -or -not $st.DG -or -not $st.DG.ItemsSource) { return }

    $importType = [string]$st.Settings.ImportType
    $sameTenant = [bool]$st.Settings.SameTenant
    $existing   = @()
    if ($script:dgIntuneManagerObjects) {
        $existing = @($script:dgIntuneManagerObjects.ItemsSource | ForEach-Object {
            if ($_ -and $_.Source) { $_.Source } else { $_ }
        } | Where-Object { $_ })
    }

    foreach ($row in @($st.DG.ItemsSource)) {
        if (-not $row -or -not $row.Source) { continue }
        $match = Resolve-IntuneImportUpdateTarget -ImportPolicy $row.Source -ExistingPolicies $existing -SameTenant $sameTenant -ImportType $importType
        $row.ImportAction        = [string]$match.Action
        $row.ImportMatch         = [string]$match.Message
        $row.ImportMatchStrategy = [string]$match.Strategy
        $row.ImportTarget        = $match.Target
    }

    $current = @($st.DG.ItemsSource)
    $st.DG.ItemsSource = $null
    $st.DG.ItemsSource = $current
}

function Invoke-IntuneManagerImportLoad
{
    # Reads $script:_imImportFormState. Loads files from $Folder into
    # the form's DataGrid + populates migration-table info. Surface any
    # error via message box so the user sees the cause; the previous
    # silent-return path swallowed Get-MigrationTableInfo / Get-PoliciesFromFolder
    # failures.
    param([string]$Folder, [bool]$KeepStatus, [switch]$Quiet)

    $st = $script:_imImportFormState
    if (-not $st) { return }
    $ui = $script:UIProvider

    if (-not $Folder) { return }
    Write-Status "Get policy objects from $Folder"

    try { $Folder = Expand-FileName $Folder } catch { }

    if (-not [IO.Directory]::Exists($Folder)) {
        if (-not $KeepStatus) { Write-Status "" }
        # -Quiet is used for the seeded-folder load that runs BEFORE the dialog
        # is shown: a popup there appears out of nowhere with no dialog behind
        # it. An explicit Browse/Get files still warns.
        if (-not $Quiet) {
            $ui.ShowMessageBox("Folder does not exist:`n$Folder", "Import", "OK", "Warning")
        } else {
            Write-Log "Import: seeded folder does not exist: $Folder" 2
        }
        return
    }

    try {
        $params = @{}
        $curPolicyTypes = Get-IntuneManagerSelectedPolicyTypes
        $subFolders = @()
        foreach ($pt in $curPolicyTypes) {
            if ([IO.Directory]::Exists([IO.Path]::Combine($Folder, $pt.Folder))) {
                $subFolders += $pt.Folder
            }
        }
        if ($subFolders.Count -gt 0) { $params['SubFolders'] = $subFolders }
        if ($curPolicyTypes)         { $params['PolicyTypes'] = $curPolicyTypes }

        # Collect into a typed List + sort in place. Routing through
        # `Sort-Object ObjectName` would stamp a PSObject ETS wrapper on each
        # IntuneImportRowItem and Avalonia's DataGrid text-column binder
        # reflects on the runtime type and renders blank cells against the
        # wrapper (see [[avalonia-binding-needs-clr-types]]).
        $rows = [System.Collections.Generic.List[IntuneImportRowItem]]::new()
        foreach ($policy in @(Get-PoliciesFromFolder $Folder @params)) {
            $obj = [PSCustomObject]$policy
            $row = [IntuneImportRowItem]::new()
            $row.Selected   = $true
            $row.Source     = $obj
            $row.ObjectName = [string]$obj.Name
            $row.PolicyType = [string]$obj.PolicyName
            $row.PolicyBase = [string]$obj.PolicyBaseName
            $row.Platform   = [string]$obj.Platform
            $row.FileName   = if ($obj.FileInfo) { [string]$obj.FileInfo.Name } else { '' }
            [void]$rows.Add($row)
        }
        $rows.Sort([Comparison[IntuneImportRowItem]]{
            param($a, $b)
            [string]::Compare($a.ObjectName, $b.ObjectName, [StringComparison]::OrdinalIgnoreCase)
        })

        if ($st.DG) { $st.DG.ItemsSource = $rows }

        Save-SettingStoreValue "" "LastUsedFullPath" $Folder

        $migrationInfo, $sameTenant = Get-MigrationTableInfo $Folder $script:organizationId
        $st.Settings.SameTenant = [bool]$sameTenant
        if ($st.LblMigInfo) { $st.LblMigInfo.Text = [string]$migrationInfo }
        if ($st.ChkRepl) {
            $st.ChkRepl.IsEnabled = (-not $sameTenant)
            $st.ChkRepl.IsChecked = (-not $sameTenant)
            $st.Settings.ReplaceDependencyIDs = (-not $sameTenant)
        }

        Update-IntuneManagerImportMatch
    } catch {
        # Log the internals; show the user only the message. The raw
        # script/line dump had no WPF counterpart.
        Write-LogError "Invoke-IntuneManagerImportLoad failed" $_.Exception
        $ui.ShowMessageBox("Could not read the import folder:`n`n$($_.Exception.Message)", "Import error", "OK", "Error")
    }

    if (-not $KeepStatus) { Write-Status "" }
}

function Show-IntuneManagerImportForm
{
    $ui = $script:UIProvider
    $policyTypes = Get-IntuneManagerSelectedPolicyTypes

    if (($policyTypes | Measure-Object).Count -eq 0) {
        $ui.ShowMessageBox("No object types selected.", "Error", "OK", "Error")
        return
    }

    $importForm = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/ImportForm.axaml'))
    if (-not $importForm) { return }

    Write-Status "Load import form"

    $hostType = Get-AvaloniaHost
    $txtFolder      = $hostType::FindByName($importForm, 'txtImportFolder')
    $browseFolder   = $hostType::FindByName($importForm, 'browseImportFolder')
    $lblMigInfo     = $hostType::FindByName($importForm, 'lblMigrationTableInfo')
    $chkScopes      = $hostType::FindByName($importForm, 'chkImportScopes')
    $chkAssign      = $hostType::FindByName($importForm, 'chkImportAssignments')
    $chkReplaceDeps = $hostType::FindByName($importForm, 'chkReplaceDependencyIDs')
    $cbImportType   = $hostType::FindByName($importForm, 'cbImportType')
    $dgImport       = $hostType::FindByName($importForm, 'dgObjectsToImport')
    $btnGetFiles    = $hostType::FindByName($importForm, 'btnGetFiles')
    $btnImport      = $hostType::FindByName($importForm, 'btnImportSelected')
    $btnCancel      = $hostType::FindByName($importForm, 'btnCancel')

    # Settings instance + per-type extras (data-side hooks).
    $importSettings = [IntuneManagerImportSettings]::new()
    $policyTypes | Add-IntuneManagerImportProperties -ImportSettings $importSettings
    $policyTypes | Add-IntuneManagerImportUIExtensions -Form $importForm

    # Prefer the last-used full path, but if it points into a per-type subfolder
    # of the currently-selected policy types, climb up one level so the parent
    # is used (matches WPF heuristic — recommended behaviour for multi-type
    # imports).
    $path = Get-SettingStoreValue "" "LastUsedFullPath"
    if ($path) {
        try {
            $di = [IO.DirectoryInfo]$path
            if ($policyTypes | Where-Object { $_.Folder -eq $di.Name }) {
                $path = $di.Parent.FullName
            }
            if (-not [IO.Directory]::Exists($path)) {
                $path = Get-SettingStoreValue "" "LastUsedRoot"
            }
        } catch { }
        if ($path) { $importSettings.ImportFolder = $path }
    }

    if ($txtFolder)      { $txtFolder.Text         = [string]$importSettings.ImportFolder }
    if ($chkScopes)      { $chkScopes.IsChecked    = [bool]$importSettings.ImportScopeTags }
    if ($chkAssign)      { $chkAssign.IsChecked    = [bool]$importSettings.ImportAssignments }
    if ($chkReplaceDeps) { $chkReplaceDeps.IsChecked = [bool]$importSettings.ReplaceDependencyIDs }

    # Import type ComboBox: NameValueObject -> SettingsListItem.
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

    # Module-scope state container. Avoids .GetNewClosure() in event
    # handlers (see file-level comment). Cleared when the form closes.
    $script:_imImportFormState = @{
        Settings   = $importSettings
        Form       = $importForm
        DG         = $dgImport
        LblMigInfo = $lblMigInfo
        ChkRepl    = $chkReplaceDeps
        TxtFolder  = $txtFolder
        Confirmed  = $false
    }

    # --- Event wiring (NO closures — references $script:_imImportFormState by name) ---

    if ($browseFolder) {
        $browseFolder.add_Click({
            $st  = $script:_imImportFormState
            $ui2 = $script:UIProvider
            if (-not $st) { return }
            $start = if ($st.TxtFolder) { [string]$st.TxtFolder.Text } else { $null }
            $folder = $ui2.ShowFolderPicker($start, 'Select root folder for import')
            if ($folder) {
                if ($st.TxtFolder) { $st.TxtFolder.Text = $folder }
                $st.Settings.ImportFolder = $folder
                Invoke-IntuneManagerImportLoad -Folder $folder -KeepStatus $false
            }
        })
    }

    if ($btnGetFiles) {
        $btnGetFiles.add_Click({
            $st = $script:_imImportFormState
            if (-not $st) { return }
            $folder = if ($st.TxtFolder) { [string]$st.TxtFolder.Text } else { $st.Settings.ImportFolder }
            $st.Settings.ImportFolder = $folder
            Invoke-IntuneManagerImportLoad -Folder $folder -KeepStatus $false
        })
    }

    if ($cbImportType) {
        $cbImportType.add_SelectionChanged({
            param($s, $e)
            $st = $script:_imImportFormState
            if (-not $st) { return }
            if ($s.SelectedItem) {
                $st.Settings.ImportType = [string]$s.SelectedItem.Value
            }
            Update-IntuneManagerImportMatch
        })
    }

    # Two-way mirror for the simple checkboxes.
    if ($chkScopes) {
        $chkScopes.add_IsCheckedChanged({
            param($s, $e)
            $st = $script:_imImportFormState
            if ($st) { $st.Settings.ImportScopeTags = [bool]$s.IsChecked }
        })
    }
    if ($chkAssign) {
        $chkAssign.add_IsCheckedChanged({
            param($s, $e)
            $st = $script:_imImportFormState
            if ($st) { $st.Settings.ImportAssignments = [bool]$s.IsChecked }
        })
    }
    if ($chkReplaceDeps) {
        $chkReplaceDeps.add_IsCheckedChanged({
            param($s, $e)
            $st = $script:_imImportFormState
            if ($st) { $st.Settings.ReplaceDependencyIDs = [bool]$s.IsChecked }
        })
    }

    # Select/deselect-all moved into the DataGrid column header.
    Initialize-AvaloniaGridSelectAllHeader -Grid $dgImport -BindingProperty 'Selected' -InitiallyChecked $true | Out-Null

    if ($btnImport) {
        $btnImport.add_Click({
            # Run the import while the form state is still live, then close + clear.
            $st = $script:_imImportFormState
            if (-not $st) { return }
            try { Invoke-IntuneManagerImportExecute }
            catch { Write-LogError "Import execution failed" $_.Exception }
            Show-ModalObject
            $script:_imImportFormState = $null
            if ($script:dgIntuneManagerObjects) { $script:dgIntuneManagerObjects.Focus() }
        })
    }
    if ($btnCancel) {
        $btnCancel.add_Click({
            Show-ModalObject
            $script:_imImportFormState = $null
            if ($script:dgIntuneManagerObjects) { $script:dgIntuneManagerObjects.Focus() }
        })
    }

    $importForm.add_KeyDown({
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            Show-ModalObject
            $script:_imImportFormState = $null
            $e.Handled = $true
        }
    })

    # Initial load against the seeded folder (if any). KeepStatus=true so
    # Write-Status "" runs after the modal opens, not before.
    if ($importSettings.ImportFolder) {
        Invoke-IntuneManagerImportLoad -Folder $importSettings.ImportFolder -KeepStatus $true -Quiet
    }

    # ShowModalForm is non-blocking, so the import cannot run after it returns
    # (the handlers fire later, on the dispatcher loop). It runs inside the
    # btnImport handler via Invoke-IntuneManagerImportExecute while state is live.
    $ui.ShowModalForm("Import objects", $importForm, $true)
    Write-Status ""
}

function Invoke-IntuneManagerImportExecute
{
    # Runs the actual import for the live $script:_imImportFormState. Called from
    # the Import button handler (not after ShowModalForm, which does not block).
    $st = $script:_imImportFormState
    if (-not $st) { return }
    $ui = $script:UIProvider

    Write-Status "Import selected objects"
    # Honour the ClearCacheBeforeExportImport setting (bit 8 = manual import).
    Invoke-GraphCacheClearBeforeOperation -Operation ManualImport

    $importType = [string]$st.Settings.ImportType
    # Persist the per-import toggles so the import pipeline (Get-SettingValue) honours them.
    $st.Settings.Save()
    $rows = @()
    if ($st.DG -and $st.DG.ItemsSource) {
        $rows = @($st.DG.ItemsSource | Where-Object { $_ -and $_.Selected })
    }

    if ($rows.Count -eq 0) {
        Write-Log "No files selected for import" 2
        Write-Status ""
        return
    }

    # Re-resolve match info up to the moment of import — the user may have
    # toggled rows or changed ImportType after the last refresh. Put state
    # back briefly so the helper can read it.
    $script:_imImportFormState = $st
    Update-IntuneManagerImportMatch
    $script:_imImportFormState = $null

    $existing = @()
    if ($script:dgIntuneManagerObjects) {
        $existing = @($script:dgIntuneManagerObjects.ItemsSource | ForEach-Object {
            if ($_ -and $_.Source) { $_.Source } else { $_ }
        } | Where-Object { $_ })
    }

    $policiesToImport = @()
    $updatedPolicies  = @()

    foreach ($row in $rows) {
        $importPolicy = $row.Source
        if (-not $importPolicy) { continue }

        $match = Resolve-IntuneImportUpdateTarget -ImportPolicy $importPolicy -ExistingPolicies $existing -SameTenant $st.Settings.SameTenant -ImportType $importType

        if ($match.Action -eq "Update" -and $importType -eq "update") {
            try {
                $updated = $importPolicy.UpdateObject($match.Target, (Get-DefaultTokenId))
                if ($updated) { $updatedPolicies += $updated }
            } catch { Write-LogError "UpdateObject failed for $($importPolicy.Name)" $_.Exception }
        }
        elseif ($match.Action -eq "Replace") {
            try {
                $replaced = Invoke-IntuneImportReplace -ImportPolicy $importPolicy -Target $match.Target -ImportType $importType
                if ($replaced) { $updatedPolicies += $replaced }
            } catch { Write-LogError "Replace failed for $($importPolicy.Name)" $_.Exception }
        }
        elseif ($match.Action -eq "Ambiguous") {
            Write-Log "Skip import/update for $($importPolicy.Name) ($($importPolicy.PolicyName)): $($match.Message)" 2
        }
        elseif ($match.Action -eq "Skip") {
            Write-Log "Skip import/update for $($importPolicy.Name) ($($importPolicy.PolicyName)): $($match.Message)"
        }
        else {
            $policiesToImport += $importPolicy
        }
    }

    $importedPolicies = @()
    if ($policiesToImport.Count -gt 0) {
        try {
            $importedPolicies = @($policiesToImport | Import-GraphPolicy)
        } catch {
            Write-LogError "Import-GraphPolicy failed" $_.Exception
            $ui.ShowMessageBox("Import failed: $($_.Exception.Message)", "Error", "OK", "Error")
            Write-Status ""
            return
        }
    }

    if ($importedPolicies.Count -gt 0 -or $updatedPolicies.Count -gt 0) {
        Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject | Out-Null
    }

    Write-Status ""
    if ($script:dgIntuneManagerObjects) { $script:dgIntuneManagerObjects.Focus() }
}
