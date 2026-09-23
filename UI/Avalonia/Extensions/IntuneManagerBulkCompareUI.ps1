# Avalonia port of the Bulk Compare dialog from
# UI/WPF/Extensions/CompareUI.ps1 (Show-GraphBulkCompareForm + helpers).
# New file because Bulk Compare is a self-contained subsystem and the WPF
# original sits in CompareUI.ps1 alongside the single-policy compare; we
# already split that into Extensions/IntuneManagerCompareUI.ps1 (architecture
# rule R9, R12).
#
# Slice 4f: full Bulk Compare form for the four "list-driven" providers
# (export / IntuneWithExport / name / exportedFolders). A fifth, Direct
# policy-pair provider was removed in 2026-09 - a two-object compare belongs
# in the single Compare form, which already picks a second policy.
#
# Helpers used (all in Internal/Compare.ps1, loaded in both UI backends):
#   Set-CompareRuntimeOptions, Get-BulkCompareOutputProvider,
#   Get-CompareOutputType, Start-BulkCompare.
# Provider classes (Classes/CompareClasses.ps1) inherit CompareProviderBase
# and have real CLR `[string]Name` / `[string]Value` properties + Validate /
# SaveSettings / GetComparePairs methods, so they bind directly to control
# DataContext for the per-provider option panels' TwoWay TextBox bindings.

function Show-GraphBulkCompareForm
{
    $ui = $script:UIProvider
    $form = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/BulkCompareForm.axaml'))
    if (-not $form) { return }

    $hostType = Get-AvaloniaHost
    $cbProvider     = $hostType::FindByName($form, 'cbCompareProvider')
    $ccProviderOpts = $hostType::FindByName($form, 'ccContentProviderOptions')
    $cbSave         = $hostType::FindByName($form, 'cbCompareSave')
    $cbOutFmt       = $hostType::FindByName($form, 'cbCompareOutputFormat')
    $cbType         = $hostType::FindByName($form, 'cbCompareType')
    $cbDelimiter    = $hostType::FindByName($form, 'cbCompareCSVDelimiter')
    $cbMultiValueDelim  = $hostType::FindByName($form, 'cbCompareMultiValueDelimiter')
    $chkIgnoreCore      = $hostType::FindByName($form, 'chkIgnoreCoreProperties')
    $chkSkipAssignments = $hostType::FindByName($form, 'chkSkipCompareAssignments')
    $chkSkipMissingSrc  = $hostType::FindByName($form, 'chkSkipMissingSourcePolicies')
    $chkSkipMissingDst  = $hostType::FindByName($form, 'chkSkipMissingDestinationPolicies')
    $grdObjects     = $hostType::FindByName($form, 'grdObjectsToCompareSection')
    $dgObjects      = $hostType::FindByName($form, 'dgObjectsToCompare')
    $txtStatus      = $hostType::FindByName($form, 'txtBulkCompareStatus')
    $btnStart       = $hostType::FindByName($form, 'btnStartCompare')
    $btnClose       = $hostType::FindByName($form, 'btnClose')

    # Panel cache + per-provider Value->instance maps + control refs. Event
    # handlers reach ALL of this through $script:_bulkCompareFormState -
    # closure captures of function locals do NOT survive
    # ConvertTo-AvaloniaEventScriptBlock (NewBoundScriptBlock rebinds to the
    # module and drops them; only global/module variables fall through).
    $state = [ordered]@{
        ProviderPanelCache = @{}
        ProviderMap        = @{}
        SaveMap            = @{}
        OutputProviderMap  = @{}
        TypeMap            = @{}
        Rows               = @()
        Ui                 = $ui
        HostType           = $hostType
        CbProvider         = $cbProvider
        CcProviderOpts     = $ccProviderOpts
        CbSave             = $cbSave
        CbOutFmt           = $cbOutFmt
        CbType             = $cbType
        CbDelimiter        = $cbDelimiter
        CbMultiValueDelim  = $cbMultiValueDelim
        ChkIgnoreCore      = $chkIgnoreCore
        ChkSkipAssignments = $chkSkipAssignments
        ChkSkipMissingSrc  = $chkSkipMissingSrc
        ChkSkipMissingDst  = $chkSkipMissingDst
        GrdObjects         = $grdObjects
        DgObjects          = $dgObjects
        TxtStatus          = $txtStatus
        BtnStart           = $btnStart
        BtnClose           = $btnClose
    }
    $script:_bulkCompareFormState = $state


    # --- Combo: Provider ---------------------------------------------------------
    if ($cbProvider) {
        $items = New-Object 'System.Collections.Generic.List[SettingsListItem]'
        foreach ($p in @($script:compareProviders)) {
            if (-not $p) { continue }
            $items.Add([SettingsListItem]@{ Name = [string]$p.Name; Value = [string]$p.Value }) | Out-Null
            $state.ProviderMap[[string]$p.Value] = $p
        }
        $cbProvider.ItemsSource  = $items
        $defaultProvider = (Get-SettingStoreValue "Compare" "Provider" "export")
        $sel = $items | Where-Object { "$($_.Value)" -eq "$defaultProvider" } | Select-Object -First 1
        if (-not $sel -and $items.Count -gt 0) { $sel = $items[0] }
        if ($sel) { $cbProvider.SelectedItem = $sel }
    }

    # --- Combo: Save (one-file-per-type vs all) ---------------------------------
    if ($cbSave) {
        $items = New-Object 'System.Collections.Generic.List[SettingsListItem]'
        foreach ($t in @($script:compareOutputTypes)) {
            if (-not $t) { continue }
            $items.Add([SettingsListItem]@{ Name = [string]$t.Name; Value = [string]$t.Value }) | Out-Null
            $state.SaveMap[[string]$t.Value] = $t
        }
        $cbSave.ItemsSource  = $items
        $defaultSave = (Get-SettingStoreValue "Compare" "SaveType" "objectType")
        $sel = $items | Where-Object { "$($_.Value)" -eq "$defaultSave" } | Select-Object -First 1
        if (-not $sel -and $items.Count -gt 0) { $sel = $items[0] }
        if ($sel) { $cbSave.SelectedItem = $sel }
    }

    # --- Combo: Output Format (CSV / JSON) --------------------------------------
    if ($cbOutFmt) {
        $items = New-Object 'System.Collections.Generic.List[SettingsListItem]'
        foreach ($p in @($script:compareOutputProviders)) {
            if (-not $p) { continue }
            $items.Add([SettingsListItem]@{ Name = [string]$p.Name; Value = [string]$p.Value }) | Out-Null
            $state.OutputProviderMap[[string]$p.Value] = $p
        }
        $cbOutFmt.ItemsSource  = $items
        $defaultOut = (Get-SettingStoreValue "Compare" "OutputFormat" "csv")
        $sel = $items | Where-Object { "$($_.Value)" -eq "$defaultOut" } | Select-Object -First 1
        if (-not $sel -and $items.Count -gt 0) { $sel = $items[0] }
        if ($sel) { $cbOutFmt.SelectedItem = $sel }
    }

    # --- Combo: Comparison Type --------------------------------------------------
    if ($cbType) {
        $items = New-Object 'System.Collections.Generic.List[SettingsListItem]'
        foreach ($t in @($script:comparisonTypes)) {
            if (-not $t) { continue }
            $items.Add([SettingsListItem]@{ Name = [string]$t.Name; Value = [string]$t.Value }) | Out-Null
            $state.TypeMap[[string]$t.Value] = $t
        }
        $cbType.ItemsSource  = $items
        $defaultType = (Get-SettingStoreValue "Compare" "Type" "property")
        $sel = $items | Where-Object { "$($_.Value)" -eq "$defaultType" } | Select-Object -First 1
        if (-not $sel -and $items.Count -gt 0) { $sel = $items[0] }
        if ($sel) { $cbType.SelectedItem = $sel }
    }

    # --- Combo: CSV Delimiter ----------------------------------------------------
    if ($cbDelimiter) {
        $cbDelimiter.ItemsSource = @("", ",", ";", "-", "|")
        $defaultDelim = (Get-SettingStoreValue "Compare" "Delimiter" ";")
        $cbDelimiter.SelectedItem = $defaultDelim
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

    # --- Checkboxes ----------------------------------------------------------------
    if ($chkIgnoreCore) {
        $chkIgnoreCore.IsChecked = ((Get-SettingStoreValue "Compare" "IgnoreCoreProperties" "true") -eq "true")
    }
    if ($chkSkipAssignments) {
        $chkSkipAssignments.IsChecked = ((Get-SettingStoreValue "Compare" "SkipCompareAssignments" "false") -eq "true")
    }
    if ($chkSkipMissingSrc) {
        $chkSkipMissingSrc.IsChecked = ((Get-SettingStoreValue "Compare" "SkipMissingSourcePolicies" "false") -eq "true")
    }
    if ($chkSkipMissingDst) {
        $chkSkipMissingDst.IsChecked = ((Get-SettingStoreValue "Compare" "SkipMissingDestinationPolicies" "false") -eq "true")
    }

    # --- Objects-to-compare grid -------------------------------------------------
    $rows = New-Object 'System.Collections.Generic.List[BulkCompareRowItem]'
    foreach ($intuneGroup in @($script:IntuneGroups)) {
        if (-not $intuneGroup -or -not $intuneGroup.Title) { continue }
        $rows.Add([BulkCompareRowItem]@{
            Title       = [string]$intuneGroup.Title
            Selected    = $true
            ObjectGroup = $intuneGroup
        }) | Out-Null
    }
    $state.Rows = @($rows)
    if ($dgObjects) { $dgObjects.ItemsSource = $state.Rows }

    # Select/deselect-all moved into the DataGrid column header.
    Initialize-AvaloniaGridSelectAllHeader -Grid $dgObjects -BindingProperty 'Selected' -InitiallyChecked $true | Out-Null

    # --- Provider panel swap -----------------------------------------------------
    $script:_bulkCompareApplyProvider = {
        param($Provider)

        $st = $script:_bulkCompareFormState
        if (-not $st -or -not $st.CcProviderOpts) { return }
        if (-not $Provider) {
            $st.CcProviderOpts.Content   = $null
            $st.CcProviderOpts.IsVisible = $false
            if ($st.GrdObjects) { $st.GrdObjects.IsVisible = $true }
            return
        }

        $panel = $null
        if ($Provider.OptionsXaml) {
            if ($st.ProviderPanelCache.ContainsKey($Provider.Value)) {
                $panel = $st.ProviderPanelCache[$Provider.Value]
            } else {
                $panel = $st.Ui.GetXamlObject((Join-Path $script:AppUIRootFolder "XAML/$($Provider.OptionsXaml).axaml"))
                if ($panel) {
                    $st.ProviderPanelCache[$Provider.Value] = $panel
                    Register-BulkCompareProviderBrowseEvents -Panel $panel -Provider $Provider -HostType $st.HostType
                }
            }
            if ($panel) { $panel.DataContext = $Provider }
        }

        $st.CcProviderOpts.Content   = $panel
        $st.CcProviderOpts.IsVisible = ($null -ne $panel)

        if ($st.GrdObjects) {
            $st.GrdObjects.IsVisible = -not [bool]$Provider.IgnoreGroups
        }
    }

    if ($cbProvider) {
        $cbProvider.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $st = $script:_bulkCompareFormState
            if (-not $st) { return }
            $sel = $s.SelectedItem
            $provider = if ($sel) { $st.ProviderMap[[string]$sel.Value] } else { $null }
            & $script:_bulkCompareApplyProvider $provider
        }))

        # Initial render.
        $initial = $null
        if ($cbProvider.SelectedItem) {
            $initial = $state.ProviderMap[[string]$cbProvider.SelectedItem.Value]
        }
        & $script:_bulkCompareApplyProvider $initial
    }

    # --- Output format → CSV-delimiter enable toggle ----------------------------
    $script:_bulkCompareApplyDelimiterEnable = {
        $st = $script:_bulkCompareFormState
        if (-not $st -or -not $st.CbDelimiter -or -not $st.CbOutFmt -or -not $st.CbOutFmt.SelectedItem) { return }
        $sel = $st.OutputProviderMap[[string]$st.CbOutFmt.SelectedItem.Value]
        $st.CbDelimiter.IsEnabled = ($sel -is [CompareCSVOutputProvider])
    }
    if ($cbOutFmt) {
        $cbOutFmt.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock { param($s, $e) & $script:_bulkCompareApplyDelimiterEnable }))
        & $script:_bulkCompareApplyDelimiterEnable
    }

    # --- Buttons -----------------------------------------------------------------
    if ($btnClose) {
        $btnClose.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            Show-ModalObject
            $script:_bulkCompareFormState = $null
        }))
    }

    if ($btnStart) {
        $btnStart.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkCompareFormState
            if (-not $st) { return }
            $providerSel = if ($st.CbProvider) { $st.CbProvider.SelectedItem } else { $null }
            $provider = if ($providerSel) { $st.ProviderMap[[string]$providerSel.Value] } else { $null }
            if (-not $provider) {
                $st.Ui.ShowMessageBox("No compare provider selected.", "Compare", "OK", "Error")
                return
            }

            # Persist all five user choices (matches WPF order).
            if ($st.CbProvider -and $st.CbProvider.SelectedItem)   { Save-SettingStoreValue "Compare" "Provider"     $st.CbProvider.SelectedItem.Value }
            if ($st.CbType     -and $st.CbType.SelectedItem)       { Save-SettingStoreValue "Compare" "Type"         $st.CbType.SelectedItem.Value }
            if ($st.CbDelimiter)                                   { Save-SettingStoreValue "Compare" "Delimiter"    ([string]$st.CbDelimiter.SelectedItem) }
            if ($st.CbOutFmt   -and $st.CbOutFmt.SelectedItem)     { Save-SettingStoreValue "Compare" "OutputFormat" $st.CbOutFmt.SelectedItem.Value }
            if ($st.CbSave     -and $st.CbSave.SelectedItem)       { Save-SettingStoreValue "Compare" "SaveType"     $st.CbSave.SelectedItem.Value }
            if ($st.CbMultiValueDelim -and $st.CbMultiValueDelim.SelectedItem) { Save-SettingStoreValue "Compare" "ObjectSeparator" ([string]$st.CbMultiValueDelim.SelectedItem.Value) }
            if ($st.ChkIgnoreCore)      { Save-SettingStoreValue "Compare" "IgnoreCoreProperties"   $(if ($st.ChkIgnoreCore.IsChecked)      { "true" } else { "false" }) }
            if ($st.ChkSkipAssignments) { Save-SettingStoreValue "Compare" "SkipCompareAssignments" $(if ($st.ChkSkipAssignments.IsChecked) { "true" } else { "false" }) }
            if ($st.ChkSkipMissingSrc)  { Save-SettingStoreValue "Compare" "SkipMissingSourcePolicies"      $(if ($st.ChkSkipMissingSrc.IsChecked) { "true" } else { "false" }) }
            if ($st.ChkSkipMissingDst)  { Save-SettingStoreValue "Compare" "SkipMissingDestinationPolicies" $(if ($st.ChkSkipMissingDst.IsChecked) { "true" } else { "false" }) }
            if ($st.ChkSkipMissingSrc)  { $provider.SkipMissingSourcePolicies      = [bool]$st.ChkSkipMissingSrc.IsChecked }
            if ($st.ChkSkipMissingDst)  { $provider.SkipMissingDestinationPolicies = [bool]$st.ChkSkipMissingDst.IsChecked }

            if ($st.BtnStart) { $st.BtnStart.IsEnabled = $false }
            if ($st.BtnClose) { $st.BtnClose.IsEnabled = $false }
            if ($st.TxtStatus) { $st.TxtStatus.Text = "Comparing..." }
            Write-Status "Compare objects"
            try {
                # Push runtime options into Internal/Compare.ps1 — mirrors WPF
                # Set-CompareRuntimeOptionsFromUI.
                $runtimeArgs = @{}
                if ($st.CbType -and $st.CbType.SelectedItem) {
                    $runtimeArgs.CompareType = [string]$st.CbType.SelectedItem.Value
                    $cd = $st.TypeMap[[string]$st.CbType.SelectedItem.Value]
                    if ($cd) { $runtimeArgs.CompareDefinition = $cd }
                }
                if ($st.CbOutFmt -and $st.CbOutFmt.SelectedItem) {
                    $op = $st.OutputProviderMap[[string]$st.CbOutFmt.SelectedItem.Value]
                    if ($op) { $runtimeArgs.OutputProvider = $op }
                }
                if ($st.CbSave -and $st.CbSave.SelectedItem) {
                    $runtimeArgs.SaveType = [string]$st.CbSave.SelectedItem.Value
                }
                if ($st.CbDelimiter -and $null -ne $st.CbDelimiter.SelectedItem) {
                    $runtimeArgs.CsvDelimiter = [string]$st.CbDelimiter.SelectedItem
                }
                if ($st.CbMultiValueDelim -and $st.CbMultiValueDelim.SelectedItem) {
                    $runtimeArgs.ObjectSeparator = [string]$st.CbMultiValueDelim.SelectedItem.Value
                }
                if ($st.ChkIgnoreCore) {
                    $runtimeArgs.IgnoreCoreProperties = [bool]$st.ChkIgnoreCore.IsChecked
                }
                if ($st.ChkSkipAssignments) {
                    $runtimeArgs.SkipAssignments = [bool]$st.ChkSkipAssignments.IsChecked
                }
                Set-CompareRuntimeOptions @runtimeArgs

                $selectedGroups = @()
                if (-not $provider.IgnoreGroups) {
                    $selectedGroups = @($st.Rows |
                        Where-Object { $_ -and $_.Selected -and $_.ObjectGroup } |
                        ForEach-Object { $_.ObjectGroup })
                    if ($selectedGroups.Count -eq 0) {
                        $st.Ui.ShowMessageBox("No object types selected.", "Compare", "OK", "Warning")
                        return
                    }
                }

                Start-BulkCompare $provider -SelectedGroups $selectedGroups
                if ($st.TxtStatus) { $st.TxtStatus.Text = "Compare finished." }
            } catch {
                Write-LogError "Bulk compare failed" $_.Exception
                $st.Ui.ShowMessageBox("Compare failed: $($_.Exception.Message)", "Compare", "OK", "Error")
                if ($st.TxtStatus) { $st.TxtStatus.Text = "Compare failed." }
            } finally {
                Write-Status ""
                if ($st.BtnStart) { $st.BtnStart.IsEnabled = $true }
                if ($st.BtnClose) { $st.BtnClose.IsEnabled = $true }
            }
        }))
    }

    $form.add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            Show-ModalObject
            $e.Handled = $true
        }
    }.GetNewClosure()))

    $ui.ShowModalForm("Bulk Compare Objects", $form, $true)
}

function Register-BulkCompareProviderBrowseEvents
{
    # Wires the four Browse buttons that may appear in the per-provider option
    # panels (browseExportPath / browseSavePath / browseExportPathSource /
    # browseExportPathCompare). Handler context rides on each button's Tag
    # (sender-based) - function locals are not available at click time
    # (ConvertTo-AvaloniaEventScriptBlock strips closures).
    # NOTE: param() must stay the FIRST statement - a $ui assignment above it
    # once made 'param' unparseable at call time and broke every panel swap.
    param($Panel, $Provider, $HostType)

    if (-not $Panel -or -not $Provider) { return }

    # Text boxes -> provider properties. The panels' XAML has no bindings and
    # nothing synced them, so everything the user typed was discarded - which
    # made the "Named Objects in Intune" provider (whose only inputs are two
    # text boxes) impossible to use at all. Each box pushes on TextChanged;
    # the current provider value seeds the box up front.
    $textPropertyPairs = @(
        @{ Name = 'txtCompareNameFilter';  Property = 'NameFilter' },
        @{ Name = 'txtExportPath';         Property = 'ExportPath' },
        @{ Name = 'txtExportPathSource';   Property = 'SourcePath' },
        @{ Name = 'txtExportPathCompare';  Property = 'ComparePath' },
        @{ Name = 'txtCompareSource';      Property = 'SourcePattern' },
        @{ Name = 'txtCompareWith';        Property = 'ComparePattern' },
        @{ Name = 'txtSavePath';           Property = 'SavePath' }
    )
    foreach ($pair in $textPropertyPairs) {
        $tb = $HostType::FindByName($Panel, $pair.Name)
        if (-not $tb) { continue }
        if (-not $Provider.PSObject.Properties[$pair.Property]) { continue }

        $tb.Text = [string]$Provider.($pair.Property)
        $tb.Tag  = @{ Provider = $Provider; Property = $pair.Property }
        $tb.add_TextChanged((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            $tag = $S.Tag
            if ($tag -and $tag.Provider) { $tag.Provider.($tag.Property) = [string]$S.Text }
        }))
    }

    $browsePathPairs = @(
        @{ Name = 'browseExportPath';        Property = 'ExportPath';   TextName = 'txtExportPath';        Title = 'Select root folder for compare' },
        @{ Name = 'browseSavePath';          Property = 'SavePath';     TextName = 'txtSavePath';          Title = 'Select save folder' },
        @{ Name = 'browseExportPathSource';  Property = 'SourcePath';   TextName = 'txtExportPathSource';  Title = 'Select source root folder' },
        @{ Name = 'browseExportPathCompare'; Property = 'ComparePath';  TextName = 'txtExportPathCompare'; Title = 'Select compare root folder' }
    )

    foreach ($pair in $browsePathPairs) {
        $btn = $HostType::FindByName($Panel, $pair.Name)
        if (-not $btn) { continue }
        # Provider classes don't expose every property; skip wiring for
        # buttons whose target property doesn't exist on this provider.
        if (-not $Provider.PSObject.Properties[$pair.Property]) { continue }

        $btn.Tag = @{
            Provider = $Provider; Property = $pair.Property; Title = $pair.Title
            TextBox  = $HostType::FindByName($Panel, $pair.TextName)
        }
        $btn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            $tag = $s.Tag
            $current = if ($tag.TextBox) { [string]$tag.TextBox.Text } else { '' }
            $folder = $script:UIProvider.ShowFolderPicker($current, [string]$tag.Title)
            if ($folder) {
                $tag.Provider.($tag.Property) = $folder
                # Write the TextBox directly: these panels have no bindings, so
                # the old DataContext null/restore refresh changed nothing on
                # screen and picking a folder looked like it did nothing.
                if ($tag.TextBox) { $tag.TextBox.Text = $folder }
            }
        }))
    }
}
