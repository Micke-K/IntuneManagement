# Avalonia port of the Compare dialog from UI/WPF/Extensions/CompareUI.ps1
# (Show-GraphCompareForm). New file because Compare is a self-contained
# subsystem we don't want to graft onto the WPF CompareUI.ps1 file
# (architecture rule R9, R12).
#
# Slice 3d1 scope:
#   - Intune-vs-Intune compare path only. The user must check 2 rows in
#     dgIntuneManagerObjects before clicking Compare; if 1 is checked, the
#     "browse second Intune object" path is stubbed (3d2/3d3 will add the
#     ObjectPicker sub-modal).
#
# Slice 3d2 adds:
#   - File compare branch — CompareFileOptions.axaml mounted into
#     ccCompareInputOptions; browseCompareObject calls
#     (Get-AvaloniaHost)::OpenFilePicker. File compare path in
#     btnStartCompare loads the file via $source.PolicyType.GetObject
#     and passes the result to Compare-PolicyObjects. Mirrors WPF lines
#     475-525 of UI/WPF/Extensions/CompareUI.ps1.
#
# Slice 3d3 adds:
#   - Match-based row coloring via DataGrid.LoadingRow + Row.Foreground swap
#     (Avalonia has no DataTrigger; Foreground inheritance cascades into the
#     TextBlock cells, matching the WPF DataGridCell.Foreground swap).
#   - btnCompareSave wired through (Get-AvaloniaHost)::SaveFilePicker +
#     Get-CompareOutputProviderByExtension + $provider.FormatRows pipeline.
#   - btnCompareCopy wired through Get-CompareCsvInfo | Set-Clipboard.
#
# Still deferred:
#   - Browse-pick second Intune object (ObjectPicker sub-modal) — later slice.
#
# Data layer notes:
#   - $script:comparisonTypes / $script:compareOutputProviders are live views onto
#     [CompareRegistry] (registered at module load in Internal/Compare.ps1; the
#     "doc" type self-registers from Internal/Documentation.ps1). Not bindable
#     in Avalonia (avalonia-binding-needs-clr-types), so we project each
#     into [SettingsListItem] (Name + Value) for ComboBox display and
#     keep a parallel hashtable that maps Value -> original PSCustomObject
#     so we can hand the source object back to Set-CompareRuntimeOptions.
#   - Compare-PolicyObjects returns PSCustomObject rows; we project those
#     into [CompareResultRowItem] before assigning DataGrid.ItemsSource.

function Invoke-IntuneManagerCompare
{
    # Mirror of the WPF btnCompare handler in UI/WPF/Extensions/IntuneManagerUI.ps1:
    # prefer rows the user has check-boxed; fall back to the highlighted row.
    if (-not $script:dgIntuneManagerObjects) { return }

    # Scan the grid's CURRENT (filtered) ItemsSource so a checked row hidden
    # by the filter is never compared; fall back to the highlighted row.
    $selectedRows = @(@($script:dgIntuneManagerObjects.ItemsSource) | Where-Object { $_ -and $_.IsSelected })
    if ($selectedRows.Count -eq 0 -and $script:dgIntuneManagerObjects.SelectedItem) {
        $selectedRows = @($script:dgIntuneManagerObjects.SelectedItem)
    }
    if ($selectedRows.Count -eq 0) { return }

    # Show-GraphCompareForm wants the underlying IntunePolicyBase objects,
    # not the IntuneObjectRowItem wrappers.
    $policies = @($selectedRows | ForEach-Object { $_.Source } | Where-Object { $_ })
    if ($policies.Count -eq 0) {
        Write-Log "Compare: $($selectedRows.Count) row(s) selected but none carried a policy object" 3
        return
    }

    Show-GraphCompareForm $policies
}

function Show-GraphCompareForm
{
    param([object[]]$Policies)

    $ui = $script:UIProvider
    if (-not $Policies -or $Policies.Count -eq 0) { return }

    $compareForm = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/CompareForm.axaml'))
    if (-not $compareForm) { return }

    $hostType   = Get-AvaloniaHost
    $tabCompare         = $hostType::FindByName($compareForm, 'tabCompare')
    $txtIntuneObject    = $hostType::FindByName($compareForm, 'txtIntuneObject')
    $cbCompareInput     = $hostType::FindByName($compareForm, 'cbCompareInput')
    $ccCompareInputOpts = $hostType::FindByName($compareForm, 'ccCompareInputOptions')
    $cbCompareType      = $hostType::FindByName($compareForm, 'cbCompareType')
    $cbOutputFormat     = $hostType::FindByName($compareForm, 'cbSingleCompareOutputFormat')
    $chkIgnoreCore      = $hostType::FindByName($compareForm, 'chkIgnoreCoreProperties')
    $cbCompareFilter    = $hostType::FindByName($compareForm, 'cbCompareFilter')
    $cbMultiValueDelim  = $hostType::FindByName($compareForm, 'cbCompareMultiValueDelimiter')
    $chkSkipAssignments = $hostType::FindByName($compareForm, 'chkSkipCompareAssignments')
    $txtSummary         = $hostType::FindByName($compareForm, 'txtCompareSummary')
    $dgCompareInfo      = $hostType::FindByName($compareForm, 'dgCompareInfo')
    $btnSwap            = $hostType::FindByName($compareForm, 'btnSwap')
    $btnStartCompare    = $hostType::FindByName($compareForm, 'btnStartCompare')
    $btnClose           = $hostType::FindByName($compareForm, 'btnClose')
    $btnCompareSave     = $hostType::FindByName($compareForm, 'btnCompareSave')
    $btnCompareCopy     = $hostType::FindByName($compareForm, 'btnCompareCopy')

    # State for this dialog instance. Lives in MODULE scope: handlers are
    # re-bound by ConvertTo-AvaloniaEventScriptBlock, which strips
    # function-local captures - the previous captured $state left Compare,
    # Swap, Save, Copy and both browse handlers throwing on their first line.
    $state = [ordered]@{
        Source         = $Policies[0]
        Target         = if ($Policies.Count -ge 2) { $Policies[1] } else { $null }
        AllRows        = @()
        InputType      = 'intune'
        TypeMap        = @{}
        OutputMap      = @{}
        InputMap       = @{}
        IntunePanel    = $null
        IntunePanelTxt = $null
        FilePanel      = $null
        FilePanelTxt   = $null
        HostType           = $hostType
        Form               = $compareForm
        TabCompare         = $tabCompare
        TxtIntuneObject    = $txtIntuneObject
        CcCompareInputOpts = $ccCompareInputOpts
        CbCompareType      = $cbCompareType
        CbOutputFormat     = $cbOutputFormat
        CbCompareInput     = $cbCompareInput
        ChkIgnoreCore      = $chkIgnoreCore
        CbCompareFilter    = $cbCompareFilter
        CbMultiValueDelim  = $cbMultiValueDelim
        ChkSkipAssignments = $chkSkipAssignments
        TxtSummary         = $txtSummary
        DgCompareInfo      = $dgCompareInfo
    }
    $script:_compareFormState = $state

    if ($txtIntuneObject) { $txtIntuneObject.Text = [string]$state.Source.Name }

    # --- Match-based row coloring ------------------------------------------------
    # Avalonia has no DataTrigger; emulate the WPF DataGridCell.Foreground swap
    # by hooking LoadingRow and setting Row.Foreground based on the bound
    # CompareResultRowItem.Match. Foreground cascades into TextBlock cells via
    # property inheritance, matching the WPF Foreground={DynamicResource
    # MismatchColor} effect. Match=$null (skipped/ignored) leaves the row at
    # the default Foreground.
    if ($dgCompareInfo) {
        $dgCompareInfo.add_LoadingRow((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $st = $script:_compareFormState
            if (-not $st) { return }
            $row  = $e.Row
            $item = $row.DataContext
            if ($null -eq $item) { return }
            if ($item.Match -eq $false) {
                $brush = $null
                if ([Avalonia.Application]::Current -and [Avalonia.Application]::Current.Resources) {
                    [void][Avalonia.Application]::Current.TryGetResource('MismatchColor', $row.ActualThemeVariant, [ref]$brush)
                }
                if ($brush) { $row.Foreground = $brush }
            } else {
                $row.ClearValue([Avalonia.Controls.DataGridRow]::ForegroundProperty)
            }
        }))
    }

    # --- Combo: Comparison Type --------------------------------------------------
    if ($cbCompareType) {
        $items = New-Object 'System.Collections.Generic.List[SettingsListItem]'
        foreach ($t in @($script:comparisonTypes)) {
            $li = [SettingsListItem]@{ Name = [string]$t.Name; Value = [string]$t.Value }
            $items.Add($li) | Out-Null
            $state.TypeMap[[string]$t.Value] = $t
        }
        $cbCompareType.ItemsSource  = $items
        $defaultType = (Get-SettingStoreValue "Compare" "Type" "property")
        $sel = $items | Where-Object { "$($_.Value)" -eq "$defaultType" } | Select-Object -First 1
        if (-not $sel -and $items.Count -gt 0) { $sel = $items[0] }
        if ($sel) { $cbCompareType.SelectedItem = $sel }
    }

    # --- Combo: Output Format ----------------------------------------------------
    if ($cbOutputFormat) {
        $items = New-Object 'System.Collections.Generic.List[SettingsListItem]'
        foreach ($p in @($script:compareOutputProviders)) {
            $li = [SettingsListItem]@{ Name = [string]$p.Name; Value = [string]$p.Value }
            $items.Add($li) | Out-Null
            $state.OutputMap[[string]$p.Value] = $p
        }
        $cbOutputFormat.ItemsSource  = $items
        $defaultOut = (Get-SettingStoreValue "Compare" "SingleOutputFormat" "csv")
        $sel = $items | Where-Object { "$($_.Value)" -eq "$defaultOut" } | Select-Object -First 1
        if (-not $sel -and $items.Count -gt 0) { $sel = $items[0] }
        if ($sel) { $cbOutputFormat.SelectedItem = $sel }
    }

    # --- Combo: Compare Input (file vs intune) ----------------------------------
    if ($cbCompareInput) {
        $items = New-Object 'System.Collections.Generic.List[SettingsListItem]'
        $items.Add([SettingsListItem]@{ Name = 'File';          Value = 'file'   }) | Out-Null
        $items.Add([SettingsListItem]@{ Name = 'Intune Object'; Value = 'intune' }) | Out-Null
        $state.InputMap['file']   = 'file'
        $state.InputMap['intune'] = 'intune'

        $cbCompareInput.ItemsSource  = $items
        # Default: prefer "intune" if a target is already populated;
        # otherwise honour the saved setting.
        $defaultInput = if ($state.Target) { 'intune' } else { (Get-SettingStoreValue "Compare" "InputType" "file") }
        $sel = $items | Where-Object { "$($_.Value)" -eq "$defaultInput" } | Select-Object -First 1
        if (-not $sel) { $sel = $items[1] }  # 'intune'
        if ($sel) {
            $cbCompareInput.SelectedItem = $sel
            $state.InputType = [string]$sel.Value
        }

        $cbCompareInput.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $st = $script:_compareFormState
            if (-not $st) { return }
            $val = if ($s.SelectedItem) { [string]$s.SelectedItem.Value } else { 'intune' }
            $st.InputType = $val
            Set-CompareInputPanel -State $st -HostType $st.HostType -Container $st.CcCompareInputOpts -Type $val
        }))

        # Initial panel render.
        Set-CompareInputPanel -State $state -HostType $hostType -Container $ccCompareInputOpts -Type $state.InputType
    }

    # --- Combo: Multi-Value Delimiter (doc-compare ObjectSeparator) ---------------
    if ($cbMultiValueDelim) {
        $items = New-Object 'System.Collections.Generic.List[SettingsListItem]'
        $items.Add([SettingsListItem]@{ Name = 'New line'; Value = [Environment]::NewLine }) | Out-Null
        $items.Add([SettingsListItem]@{ Name = ';';        Value = ';' }) | Out-Null
        $items.Add([SettingsListItem]@{ Name = '|';        Value = '|' }) | Out-Null
        $cbMultiValueDelim.ItemsSource  = $items
        $defaultSep = (Get-SettingStoreValue "Compare" "ObjectSeparator" ([Environment]::NewLine))
        $sel = $items | Where-Object { "$($_.Value)" -eq "$defaultSep" } | Select-Object -First 1
        if (-not $sel) { $sel = $items[0] }
        $cbMultiValueDelim.SelectedItem = $sel
    }

    # --- Checkboxes --------------------------------------------------------------
    if ($chkIgnoreCore) {
        $chkIgnoreCore.IsChecked = ((Get-SettingStoreValue "Compare" "IgnoreCoreProperties" "true") -eq "true")
    }
    if ($chkSkipAssignments) {
        $chkSkipAssignments.IsChecked = ((Get-SettingStoreValue "Compare" "SkipCompareAssignments" "false") -eq "true")
    }
    if ($cbCompareFilter) {
        # 3-way result filter (All / Mismatch / Match). Mirror the other combos in
        # this file (SettingsListItem + LoadXaml template) - Avalonia's binder needs
        # CLR-typed items.
        $filterItems = New-Object 'System.Collections.Generic.List[SettingsListItem]'
        $filterItems.Add([SettingsListItem]@{ Name = 'All properties';  Value = 'All' })      | Out-Null
        $filterItems.Add([SettingsListItem]@{ Name = 'Mismatches only'; Value = 'Mismatch' }) | Out-Null
        $filterItems.Add([SettingsListItem]@{ Name = 'Matches only';    Value = 'Match' })     | Out-Null
        $cbCompareFilter.ItemsSource  = $filterItems
        $savedFilter = Get-SettingStoreValue "Compare" "ResultFilter" "All"
        $sel = $filterItems | Where-Object { "$($_.Value)" -eq "$savedFilter" } | Select-Object -First 1
        if (-not $sel) { $sel = $filterItems[0] }
        $cbCompareFilter.SelectedItem = $sel
        $cbCompareFilter.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $st = $script:_compareFormState
            if (-not $st) { return }
            $mode = if ($s.SelectedItem) { [string]$s.SelectedItem.Value } else { 'All' }
            Update-CompareGridRows -State $st -Grid $st.DgCompareInfo -Filter $mode
        }))
    }

    # --- Buttons -----------------------------------------------------------------
    if ($btnClose) {
        $btnClose.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            Show-ModalObject
            # Holds both hydrated policies plus every compare row - release it.
            $script:_compareFormState = $null
        }))
    }

    if ($btnSwap) {
        $btnSwap.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_compareFormState
            if (-not $st) { return }
            if (-not $st.Source -or -not $st.Target) { return }
            $tmp           = $st.Source
            $st.Source  = $st.Target
            $st.Target  = $tmp
            if ($st.TxtIntuneObject) { $st.TxtIntuneObject.Text = [string]$st.Source.Name }
            if ($st.IntunePanelTxt) { $st.IntunePanelTxt.Text = [string]$st.Target.Name }
        }))
    }

    if ($btnStartCompare) {
        $btnStartCompare.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_compareFormState
            if (-not $st) { return }
            $inputType = $st.InputType

            # Resolve the second object early so we fail fast with a useful
            # message before persisting settings.
            $secondPolicy = $null
            if ($inputType -eq 'file') {
                $compareFile = if ($st.FilePanelTxt) { [string]$st.FilePanelTxt.Text } else { '' }
                if ([string]::IsNullOrWhiteSpace($compareFile)) {
                    $ui.ShowMessageBox("No file selected for comparison.", "Compare", "OK", "Error")
                    return
                }
                if (-not [IO.File]::Exists($compareFile)) {
                    $ui.ShowMessageBox("File '$compareFile' not found.", "Compare", "OK", "Error")
                    return
                }
                if (-not $st.Source -or -not $st.Source.PolicyType) {
                    $ui.ShowMessageBox("Source object has no PolicyType - cannot load file.", "Compare", "OK", "Error")
                    return
                }
                try {
                    $secondPolicy = $st.Source.PolicyType.GetObject([IO.FileInfo]$compareFile)
                } catch {
                    Write-LogError "Failed to load compare file '$compareFile'" $_.Exception
                    $ui.ShowMessageBox("Failed to load file '$compareFile': $($_.Exception.Message)", "Compare", "OK", "Error")
                    return
                }
                if (-not $secondPolicy) {
                    $ui.ShowMessageBox("Failed to load file '$compareFile'. Object type may not match.", "Compare", "OK", "Error")
                    return
                }
                Save-SettingStoreValue "Compare" "LastFile" $compareFile
            } else {
                if (-not $st.Target) {
                    $ui.ShowMessageBox("No second Intune object selected. Pre-select two objects in the list (check the boxes) before opening Compare. Browse-pick is not yet ported.", "Compare", "OK", "Error")
                    return
                }
                $secondPolicy = $st.Target
            }

            # Persist user choices.
            if ($st.CbCompareType -and $st.CbCompareType.SelectedItem) {
                Save-SettingStoreValue "Compare" "Type" $st.CbCompareType.SelectedItem.Value
            }
            if ($st.CbCompareInput -and $st.CbCompareInput.SelectedItem) {
                Save-SettingStoreValue "Compare" "InputType" $st.CbCompareInput.SelectedItem.Value
            }
            if ($st.CbOutputFormat -and $st.CbOutputFormat.SelectedItem) {
                Save-SettingStoreValue "Compare" "SingleOutputFormat" $st.CbOutputFormat.SelectedItem.Value
            }
            if ($st.ChkIgnoreCore) {
                Save-SettingStoreValue "Compare" "IgnoreCoreProperties" $(if ($st.ChkIgnoreCore.IsChecked) { "true" } else { "false" })
            }
            if ($st.CbCompareFilter -and $st.CbCompareFilter.SelectedItem) {
                Save-SettingStoreValue "Compare" "ResultFilter" ([string]$st.CbCompareFilter.SelectedItem.Value)
            }
            if ($st.CbMultiValueDelim -and $st.CbMultiValueDelim.SelectedItem) {
                Save-SettingStoreValue "Compare" "ObjectSeparator" ([string]$st.CbMultiValueDelim.SelectedItem.Value)
            }
            if ($st.ChkSkipAssignments) {
                Save-SettingStoreValue "Compare" "SkipCompareAssignments" $(if ($st.ChkSkipAssignments.IsChecked) { "true" } else { "false" })
            }

            # Push runtime options into the Internal/Compare.ps1 module so the
            # comparison engine sees the same values WPF would have set via
            # Set-CompareRuntimeOptionsFromUI.
            $runtimeArgs = @{}
            if ($st.CbCompareType -and $st.CbCompareType.SelectedItem) {
                $runtimeArgs.CompareType = [string]$st.CbCompareType.SelectedItem.Value
                $cd = $st.TypeMap[[string]$st.CbCompareType.SelectedItem.Value]
                if ($cd) { $runtimeArgs.CompareDefinition = $cd }
            }
            if ($st.ChkIgnoreCore) {
                $runtimeArgs.IgnoreCoreProperties = [bool]$st.ChkIgnoreCore.IsChecked
            }
            if ($st.CbOutputFormat -and $st.CbOutputFormat.SelectedItem) {
                $op = $st.OutputMap[[string]$st.CbOutputFormat.SelectedItem.Value]
                if ($op) { $runtimeArgs.OutputProvider = $op }
            }
            if ($st.CbMultiValueDelim -and $st.CbMultiValueDelim.SelectedItem) {
                $runtimeArgs.ObjectSeparator = [string]$st.CbMultiValueDelim.SelectedItem.Value
            }
            if ($st.ChkSkipAssignments) {
                $runtimeArgs.SkipAssignments = [bool]$st.ChkSkipAssignments.IsChecked
            }
            Set-CompareRuntimeOptions @runtimeArgs

            Write-Status "Compare objects"
            try {
                Resolve-FullPolicyForCompare $st.Source
                if ($inputType -eq 'intune') {
                    Resolve-FullPolicyForCompare $secondPolicy
                } else {
                    # File-loaded policy: warn if @odata.type doesn't match.
                    $sourceType = $null; $fileType = $null
                    try { $sourceType = $st.Source.Object.'@OData.Type' } catch { }
                    try { $fileType   = $secondPolicy.Object.'@OData.Type' } catch { }
                    if ($sourceType -and $fileType -and $sourceType -ne $fileType) {
                        $answer = $ui.ShowMessageBox("The object types do not match.`n`nDo you want to compare the objects anyway?", "Compare", "YesNo", "Warning")
                        if ($answer -ne "Yes") { Write-Status ""; return }
                    }
                }
                $rawRows = @(Compare-PolicyObjects @($st.Source, $secondPolicy))
            } catch {
                Write-LogError "Compare-PolicyObjects failed" $_.Exception
                $ui.ShowMessageBox("Compare failed: $($_.Exception.Message)", "Error", "OK", "Error")
                Write-Status ""
                return
            }

            $st.AllRows = @($rawRows | ForEach-Object {
                # Tri-state Match glyph, mirroring the WPF CompareForm.xaml column:
                # checkmark = match, ballot-X = mismatch, em-dash = ignored ($null).
                # Built from [char] codes so the source stays ASCII (see
                # [[ascii-only-string-literals]]) - the literal glyph exists only in
                # the rendered string, never in the .ps1 bytes. Mismatch rows are
                # already tinted red by the LoadingRow handler, which cascades to
                # this cell too (matching the red WPF ballot-X).
                $glyph = [char]0x2014
                if ($_.Match -eq $true)      { $glyph = [char]0x2713 }
                elseif ($_.Match -eq $false) { $glyph = [char]0x2717 }
                [CompareResultRowItem]@{
                    PropertyName = [string]$_.PropertyName
                    Category     = [string]$_.Category
                    SubCategory  = [string]$_.SubCategory
                    Object1Value = [string]$_.Object1Value
                    Object2Value = [string]$_.Object2Value
                    Match        = $_.Match
                    MatchGlyph   = [string]$glyph
                }
            })

            $mode = if ($st.CbCompareFilter -and $st.CbCompareFilter.SelectedItem) { [string]$st.CbCompareFilter.SelectedItem.Value } else { 'All' }
            Update-CompareGridRows -State $st -Grid $st.DgCompareInfo -Filter $mode
            Update-CompareSummaryText -State $st -TextBlock $st.TxtSummary

            if ($st.TabCompare) { $st.TabCompare.SelectedIndex = 1 }
            Write-Status ""
        }))
    }

    if ($btnCompareSave) {
        $btnCompareSave.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_compareFormState
            if (-not $st) { return }
            if (-not $st.AllRows -or $st.AllRows.Count -eq 0) { return }

            # Pick output provider from the SelectedItem (combobox model has the
            # underlying CompareOutputProviderBase via OutputMap), with CSV fallback.
            $outProv = $null
            if ($st.CbOutputFormat -and $st.CbOutputFormat.SelectedItem) {
                $outProv = $st.OutputMap[[string]$st.CbOutputFormat.SelectedItem.Value]
            }
            if (-not $outProv) { $outProv = [CompareCSVOutputProvider]::new() }

            $defaultExt = if ($outProv.Extension) { [string]$outProv.Extension } else { 'csv' }
            $filterName    = if ($defaultExt -eq 'json') { 'JSON files' } else { 'CSV files' }
            $filterPattern = "*.$defaultExt"

            $startDir = Get-SettingStoreValue "Compare" "LastSaveDirectory"
            if (-not $startDir) { $startDir = Get-SettingValue "RootFolder" }

            $picked = (Get-AvaloniaHost)::SaveFilePicker(
                $script:Window,
                'Save compare results',
                [string]$st.Source.Name,
                $defaultExt,
                $filterName,
                $filterPattern,
                [string]$startDir)
            if (-not $picked) { return }

            try {
                Save-SettingStoreValue "Compare" "LastSaveDirectory" ([IO.FileInfo]$picked).DirectoryName
                $ext      = [IO.Path]::GetExtension($picked).TrimStart('.')
                $writeProv = Get-CompareOutputProviderByExtension $ext
                if (-not $writeProv) { $writeProv = $outProv }

                # FormatRows wants the Compare-PolicyObjects shape (PSCustomObject
                # with PropertyName/Object1Value/...); CompareResultRowItem mirrors
                # those property names so duck-typing through Get-Member works.
                $propsArgs = @{}
                if ($st.CbCompareType -and $st.CbCompareType.SelectedItem) {
                    $propsArgs.CompareType = [string]$st.CbCompareType.SelectedItem.Value
                    $propsArgs.CompareDefinition = $st.TypeMap[[string]$st.CbCompareType.SelectedItem.Value]
                }
                $props = Get-CompareOutputProps @propsArgs
                $writeProv.FormatRows($st.AllRows, $props) | Out-File -LiteralPath $picked -Force -Encoding UTF8
            } catch {
                Write-LogError "Save compare results failed" $_.Exception
                $ui.ShowMessageBox("Save failed: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        }))
    }

    if ($btnCompareCopy) {
        $btnCompareCopy.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_compareFormState
            if (-not $st) { return }
            if (-not $st.AllRows -or $st.AllRows.Count -eq 0) { return }
            try {
                $compareType = if ($st.CbCompareType -and $st.CbCompareType.SelectedItem) {
                    [string]$st.CbCompareType.SelectedItem.Value
                } else { 'property' }
                $compareDef  = $st.TypeMap[$compareType]
                (Get-CompareCsvInfo $st.AllRows $st.Source $compareType $compareDef) | Set-Clipboard
            } catch {
                Write-LogError "Copy compare results failed" $_.Exception
                $ui.ShowMessageBox("Copy failed: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        }))
    }

    $compareForm.add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
        param($s, $e)
        $st = $script:_compareFormState
        if (-not $st) { return }
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            Show-ModalObject
            $e.Handled = $true
        }
    }))

    $ui.ShowModalForm("Compare Intune Objects", $compareForm, $true)
}

function Set-CompareInputPanel
{
    param($State, $HostType, $Container, [string]$Type)

    $ui = $script:UIProvider
    if (-not $Container) { return }

    if ($Type -eq 'intune') {
        if (-not $State.IntunePanel) {
            $State.IntunePanel = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/CompareIntuneObjectOptions.axaml'))
            if ($State.IntunePanel) {
                $State.IntunePanelTxt = $HostType::FindByName($State.IntunePanel, 'txtCompareIntuneObject')
                $browse                = $HostType::FindByName($State.IntunePanel, 'browseIntuneObject')
                if ($browse) {
                    $browse.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                        $st = $script:_compareFormState
                        if (-not $st) { return }

                        $sourceType = if ($st.Source) { $st.Source.PolicyType } else { $null }
                        if (-not $sourceType) {
                            Show-MessageBox "Select a source policy first." "Compare" "OK" "Warning" | Out-Null
                            return
                        }
                        $sourceId = if ($st.Source) { $st.Source.Id } else { $null }

                        # Seed with same-type policies already cached in the main view so the
                        # dialog opens with something in it and no Graph call. Searching or
                        # "Load all in scope" goes to Graph from there.
                        $initial = @()
                        if ($script:IntuneManagerAllRows) {
                            $initial = @(
                                $script:IntuneManagerAllRows |
                                    Where-Object { $_.Source -and $_.Source.PolicyType -and $_.Source.PolicyType.ID -eq $sourceType.ID -and (-not $sourceId -or $_.Source.Id -ne $sourceId) } |
                                    ForEach-Object { $_.Source }
                            )
                        }

                        $exclude = @()
                        if ($sourceId) { $exclude = @($sourceId) }

                        $picked = Show-PolicySearchDialog `
                            -Title ("Select Intune Object to Compare ({0})" -f $sourceType.Title) `
                            -DefaultPolicyTypeId $sourceType.ID `
                            -ExcludeIds $exclude `
                            -InitialItems $initial
                        if ($picked) {
                            $st.Target = $picked
                            if ($st.IntunePanelTxt) { $st.IntunePanelTxt.Text = [string]$picked.Name }
                        }
                    }))
                }
            }
        }
        if ($State.IntunePanelTxt -and $State.Target) {
            $State.IntunePanelTxt.Text = [string]$State.Target.Name
        }
        $Container.Content = $State.IntunePanel
    } else {
        if (-not $State.FilePanel) {
            $State.FilePanel = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/CompareFileOptions.axaml'))
            if ($State.FilePanel) {
                $State.FilePanelTxt = $HostType::FindByName($State.FilePanel, 'txtCompareFile')
                $browseFile         = $HostType::FindByName($State.FilePanel, 'browseCompareObject')

                # Pre-populate from saved last-file path so the user doesn't
                # have to re-browse on every reopen.
                $lastFile = Get-SettingStoreValue "Compare" "LastFile"
                if ($lastFile -and $State.FilePanelTxt) { $State.FilePanelTxt.Text = $lastFile }

                if ($browseFile) {
                    $browseFile.add_Click((ConvertTo-AvaloniaEventScriptBlock {
                        $st = $script:_compareFormState
                        if (-not $st) { return }
                        # Pick a sensible starting directory: Compare/LastFile's
                        # parent if set, else the source's PolicyType.Folder
                        # under LastUsedFullPath, else nothing (host falls back
                        # to default).
                        $startDir = $null
                        $cachedFile = Get-SettingStoreValue "Compare" "LastFile"
                        if ($cachedFile) {
                            try { $startDir = ([IO.FileInfo]$cachedFile).DirectoryName } catch { }
                        }
                        if (-not $startDir) {
                            $rootHint = Get-SettingStoreValue "" "LastUsedFullPath"
                            if ($rootHint) {
                                try { $rootHint = [IO.Directory]::GetParent($rootHint).FullName } catch { }
                            }
                            if ($rootHint -and $st.Source -and $st.Source.PolicyType -and $st.Source.PolicyType.Folder) {
                                $candidate = [IO.Path]::Combine($rootHint, $st.Source.PolicyType.Folder)
                                if ([IO.Directory]::Exists($candidate)) { $startDir = $candidate }
                                elseif ([IO.Directory]::Exists($rootHint)) { $startDir = $rootHint }
                            } elseif ($rootHint) {
                                $startDir = $rootHint
                            }
                        }

                        $picked = (Get-AvaloniaHost)::OpenFilePicker($script:Window, 'Select compare file', $startDir, 'JSON files', '*.json')
                        if ($picked) {
                            if ($st.FilePanelTxt) { $st.FilePanelTxt.Text = $picked }
                            Save-SettingStoreValue "Compare" "LastFile" $picked
                        }
                    }))
                }
            }
        }
        $Container.Content = $State.FilePanel
    }
}

function Update-CompareGridRows
{
    # $Filter: All (everything) / Mismatch (Match -eq $false) / Match (Match -eq $true).
    # The 'Match' view hides ignored ($null) rows, same as 'Mismatch' hides them.
    param($State, $Grid, [string]$Filter = 'All')

    if (-not $Grid) { return }
    # Build a TYPED list rather than piping through Where-Object: the
    # pipeline stamps a PSObject wrapper on every item, and Avalonia's
    # DataGrid binder reflects on the runtime type and renders blank cells
    # against the wrapper (same trap as the assignments filter - see
    # [[avalonia-binding-needs-clr-types]]).
    $rows = [System.Collections.Generic.List[CompareResultRowItem]]::new()
    foreach ($row in @($State.AllRows)) {
        if (-not $row) { continue }
        $keep = switch ($Filter) {
            'Mismatch' { $row.Match -eq $false }
            'Match'    { $row.Match -eq $true }
            default    { $true }
        }
        if ($keep) { $rows.Add($row) }
    }
    $Grid.ItemsSource = $rows
}

function Update-CompareSummaryText
{
    param($State, $TextBlock)

    if (-not $TextBlock) { return }
    $items      = @($State.AllRows)
    $total      = $items.Count
    $mismatches = @($items | Where-Object { $_.Match -eq $false }).Count
    $skipped    = @($items | Where-Object { $null -eq $_.Match }).Count
    $matched    = $total - $mismatches - $skipped

    $tail = if ($skipped -gt 0) { ", $skipped ignored" } else { '' }
    $TextBlock.Text = "$total properties - $matched matched, $mismatches mismatched$tail"
}

# Category / SubCategory columns are only meaningful for settings-based
# comparisons; Internal/Compare.ps1 raises CompareColumnVisibilityChanged to
# show or hide them. WPF wires this (IntuneManagerCompareUIWPF.ps1:1); the
# Avalonia port never did, so both columns stayed visible - and empty - for
# property compares.
Add-AppEventHandler "CompareColumnVisibilityChanged" "Set-CompareGridColumnVisibility"

function Set-CompareGridColumnVisibility
{
    param($ShowCategory = $false, $ShowSubCategory = $false)

    $st = $script:_compareFormState
    $grid = if ($st) { $st.DgCompareInfo } else { $null }
    if (-not $grid) { return }

    foreach ($col in $grid.Columns) {
        $header = [string]$col.Header
        if ($header -eq 'Category')    { $col.IsVisible = [bool]$ShowCategory }
        if ($header -eq 'SubCategory') { $col.IsVisible = [bool]$ShowSubCategory }
    }
}
