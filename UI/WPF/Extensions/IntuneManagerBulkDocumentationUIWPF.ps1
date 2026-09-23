# Bulk Documentation form + per-view Show-IntuneManagerDocumentForm caller.
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Set-WpfDocumentationCombo {
    param($Form, [string]$Name, $Items, $SelectedValue)
    $control = $Form.FindName($Name)
    if (-not $control) { return }
    $control.ItemsSource = Get-UIItemsArray $Items
    if ($null -ne $SelectedValue) { $control.SelectedValue = $SelectedValue }
    if ($control.SelectedIndex -lt 0 -and $control.Items.Count -gt 0) { $control.SelectedIndex = 0 }
}

# Show/hide $TargetNames whenever $TriggerName's SelectedValue matches
# $VisibleWhen. Mirrors the OLD project's per-output-form selectionChanged
# wiring (DocumentationWord.psm1 ~L126, Documentation.psm1 ~L5029) so the
# Custom Properties row only appears for Output Properties = "custom" and
# the limit-length grid only appears for Output Level = "limited".
function Set-WpfDocumentationConditionalVisibility {
    param(
        $Form,
        [string]$TriggerName,
        [string]$VisibleWhen,
        [string[]]$TargetNames
    )
    $trigger = $Form.FindName($TriggerName)
    if (-not $trigger) { return }
    $targets = @($TargetNames | ForEach-Object { $Form.FindName($_) } | Where-Object { $_ })
    if ($targets.Count -eq 0) { return }

    $apply = {
        param($selectedValue)
        $vis = if ([string]$selectedValue -eq $VisibleWhen) { 'Visible' } else { 'Collapsed' }
        foreach ($t in $targets) { $t.Visibility = $vis }
    }.GetNewClosure()

    & $apply $trigger.SelectedValue
    $trigger.Add_SelectionChanged({
        param($s, $e)
        & $apply $s.SelectedValue
    }.GetNewClosure())
}

function Initialize-WpfDocumentationOptions {
    param($Form)

    Set-WpfDocumentationCombo $Form 'cbDocumentationLanguage' (Get-DocumentationLanguages | Sort-Object EnglishName) (Get-DocumentationSetting 'Language' 'en')
    Set-WpfDocumentationCombo $Form 'cbDocumentationPropertySeparator' @(',',';','-','|') (Get-DocumentationSetting 'PropertySeparator' ';')
    Set-WpfDocumentationCombo $Form 'cbDocumentationObjectSeparator' @(
        [PSCustomObject]@{ Name='New line'; Value=[Environment]::NewLine },
        [PSCustomObject]@{ Name=';'; Value=';' },
        [PSCustomObject]@{ Name='|'; Value='|' }
    ) (Get-DocumentationSetting 'ObjectSeparator' ([Environment]::NewLine))
    Set-WpfDocumentationCombo $Form 'cbNotConfiguredText' @(
        [PSCustomObject]@{Name='Not configured (localized)';Value='notConfigured'},
        [PSCustomObject]@{Name='Empty';Value='empty'},
        [PSCustomObject]@{Name="Don't change";Value='asis'}
    ) (Get-DocumentationSetting 'NotConfiguredText' 'notConfigured')
    Set-WpfDocumentationCombo $Form 'cbValueOutputProperty' @(
        [PSCustomObject]@{Name='Value';Value='value'},
        [PSCustomObject]@{Name='Value with label';Value='valueWithLabel'}
    ) (Get-DocumentationSetting 'ValueOutputProperty' 'value')
    $propertyModes = @([PSCustomObject]@{Name='Simple';Value='simple'},[PSCustomObject]@{Name='Extended';Value='extended'},[PSCustomObject]@{Name='Custom';Value='custom'})
    $fileModes = @([PSCustomObject]@{Name='Single file';Value='Full'},[PSCustomObject]@{Name='One file per object';Value='Object'})
    Set-WpfDocumentationCombo $Form 'cbCSVDocumentationProperties' $propertyModes (Get-DocumentationSetting 'CSVExportProperties' 'simple')
    Set-WpfDocumentationCombo $Form 'cbCSVDelimiter' @('',',',';','-','|') (Get-DocumentationSetting 'CSVDelimiter' '')
    Set-WpfDocumentationCombo $Form 'cbJsonOutputFileType' @([PSCustomObject]@{Name='Single file';Value='Full'},[PSCustomObject]@{Name='One file per object type';Value='ObjectType'}) (Get-DocumentationSetting 'JSONOutputFileType' 'Full')
    Set-WpfDocumentationCombo $Form 'cbHTMLDocumentFileType' $fileModes (Get-DocumentationSetting 'HTMLDocumentFileType' 'Full')
    Set-WpfDocumentationCombo $Form 'cbAtlassianDocumentFileType' $fileModes (Get-DocumentationSetting 'AtlassianDocumentFileType' 'Full')
    Set-WpfDocumentationCombo $Form 'cbMDDocumentFileType' $fileModes (Get-DocumentationSetting 'MDDocumentFileType' 'Full')
    Set-WpfDocumentationCombo $Form 'cbWordExportProperties' $propertyModes (Get-DocumentationSetting 'WordExportProperties' 'simple')
    Set-WpfDocumentationCombo $Form 'cbWordDocumentFormat' @([PSCustomObject]@{Name='DOCX';Value='wdFormatDocumentDefault'},[PSCustomObject]@{Name='Strict Open XML';Value='wdFormatStrictOpenXMLDocument'},[PSCustomObject]@{Name='PDF';Value='wdFormatPDF'}) (Get-DocumentationSetting 'WordDocumentFormat' 'wdFormatDocumentDefault')
    Set-WpfDocumentationCombo $Form 'cbWordDocumentationLevel' @([PSCustomObject]@{Name='Full';Value='full'},[PSCustomObject]@{Name='Limited';Value='limited'},[PSCustomObject]@{Name='Basic';Value='basic'}) (Get-DocumentationSetting 'WordDocumentationLevel' 'full')
    Set-WpfDocumentationCombo $Form 'cbWordTableCaptionPosition' @([PSCustomObject]@{Name='Below table';Value='below'},[PSCustomObject]@{Name='Above table';Value='above'}) (Get-DocumentationSetting 'WordTableCaptionPosition' 'below')

    $defaults = @{
        txtCSVDocumentationPath='CSVDocumentationPath'; txtCSVCustomProperties='CSVCustomDisplayProperties'; txtJsonDocumentName='JSONDocumentName'
        txtHTMLDocumentName='HTMLDocumentName'; txtHTMLCSSFile='HTMLCSSFile'; txtHTMLTitleProperty='HTMLTitleProperty'
        txtAtlassianDocumentName='AtlassianDocumentName'; txtAtlassianTitleProperty='AtlassianTitleProperty'
        txtMDDocumentName='MDDocumentName'; txtMDCSSFile='MDCSSFile'; txtMDTitleProperty='MDTitleProperty'
        txtWordDocumentName='WordDocumentName'; txtWordDocumentTemplate='WordDocumentTemplate'; txtWordCustomDisplayProperties='WordCustomDisplayProperties'
        txtWordDocumentationLimitMaxLength='WordDocumentationLimitMaxLength'; txtWordDocumentationLimitTruncateLength='WordDocumentationLimitTruncateLength'
        txtWordTitleProperty='WordTitleProperty'; txtWordSubjectProperty='WordSubjectProperty'; txtWordContentControls='WordContentControls'
        txtWordCoverPage='WordCoverPage'; txtWordHeader1Style='WordHeader1Style'; txtWordHeader2Style='WordHeader2Style'; txtWordHeader3Style='WordHeader3Style'
        txtWordTableStyle='WordTableStyle'; txtWordTableHeaderStyle='WordTableHeaderStyle'; txtWordCategoryHeaderStyle='WordCategoryHeaderStyle'
        txtWordSubCategoryHeaderStyle='WordSubCategoryHeaderStyle'; txtWordTableTextStyle='WordTableTextStyle'
        txtWordScriptTableStyle='WordScriptTableStyle'; txtWordScriptStyle='WordScriptStyle'
    }
    foreach($name in $defaults.Keys) {
        $c = $Form.FindName($name)
        # Some output XAML files don't carry every control yet (e.g. the
        # WPF MD form omits TitleProperty); skip silently rather than
        # NullReferenceException-on-property-set.
        if ($c) { $c.Text = [string](Get-DocumentationSetting $defaults[$name] '') }
    }
    foreach($pair in @(
        @('chkSkipNotConfigured','SkipNotConfigured',$false),@('chkSkipDefaultValues','SkipDefaultValues',$false),@('chkSkipDisabled','SkipDisabled',$true),
        @('chkSetUnconfiguredValue','SetUnconfiguredValue',$true),@('chkSetDefaultValue','SetDefaultValue',$false),@('chkIncludeScripts','IncludeScripts',$true),
        @('chkExcludeScriptSignature','ExcludeScriptSignature',$false),@('chkExcludeAssignments','ExcludeAssignments',$false),@('chkIncludePolicyId','IncludePolicyId',$false),@('chkSkipDocumentInfo','SkipDocumentInfo',$false),@('chkFallbackDocumentation','FallbackDocumentation',$true),
        @('chkCSVAddObjectType','CSVAddObjectType',$true),@('chkCSVAddCompanyName','CSVAddCompanyName',$false),@('chkJsonOpenFile','JSONOpenFile',$true),
        @('chkHTMLOpenFile','HTMLOpenFile',$true),@('chkAtlassianOpenFile','AtlassianOpenFile',$true),
        @('chkMDIncludeCSS','MDIncludeCSS',$true),@('chkMDOpenFile','MDOpenFile',$true),@('chkMDDocumentSkipDate','MDDocumentSkipDate',$false),
        @('chkWordDocumentationLimitAttach','WordDocumentationLimitAttach',$false),@('chkWordAddCategories','WordAddCategories',$true),@('chkWordAddSubCategories','WordAddSubCategories',$true),
        @('chkWordOpenDocument','WordOpenDocument',$true),@('chkWordAttachJsonFile','WordAttachJsonFile',$false)
    )) {
        $c = $Form.FindName($pair[0])
        if ($c) { $c.IsChecked = (Get-DocumentationSetting $pair[1] $pair[2]) -eq $true }
    }

    # Conditional-visibility wiring: matches the OLD project's per-output-form
    # selectionChanged handlers. The Custom Properties row only appears for
    # Output Properties = "custom"; the limit-length grid only appears for
    # Output Level = "limited".
    Set-WpfDocumentationConditionalVisibility -Form $Form -TriggerName 'cbWordExportProperties' -VisibleWhen 'custom' `
        -TargetNames @('spWordCustomProperties','txtWordCustomDisplayProperties')
    Set-WpfDocumentationConditionalVisibility -Form $Form -TriggerName 'cbWordDocumentationLevel' -VisibleWhen 'limited' `
        -TargetNames @('gdWordDocumentationLimitOptions')
    Set-WpfDocumentationConditionalVisibility -Form $Form -TriggerName 'cbCSVDocumentationProperties' -VisibleWhen 'custom' `
        -TargetNames @('spCSVCustomProperties','txtCSVCustomProperties')
}

function Get-WpfDocumentationOptions {
    param($Form)
    # Missing-control safety: if a XAML file doesn't carry a referenced
    # control (e.g. legacy form without the new field), return $null/empty
    # rather than tripping NullReferenceException. Save-DocumentationOptionsDefaults
    # later filters $null so we don't overwrite the user's persisted setting
    # with garbage on click-Export.
    function Txt($n) { $c = $Form.FindName($n); if ($c) { [string]$c.Text } else { $null } }
    function Sel($n) { $c = $Form.FindName($n); if ($c) { $c.SelectedValue } else { $null } }
    function Chk($n) { $c = $Form.FindName($n); if ($c) { [bool]$c.IsChecked } else { $null } }
    @{
        Language=Sel 'cbDocumentationLanguage'; PropertySeparator=Sel 'cbDocumentationPropertySeparator'; ObjectSeparator=Sel 'cbDocumentationObjectSeparator'; NameFilter=Txt 'txtDocumentFilter'; SourceFolder=Txt 'txtDocumentFromFolder'
        SkipNotConfigured=Chk 'chkSkipNotConfigured'; SkipDefaultValues=Chk 'chkSkipDefaultValues'; SkipDisabled=Chk 'chkSkipDisabled'
        SetUnconfiguredValue=Chk 'chkSetUnconfiguredValue'; SetDefaultValue=Chk 'chkSetDefaultValue'; IncludeScripts=Chk 'chkIncludeScripts'
        ExcludeScriptSignature=Chk 'chkExcludeScriptSignature'; ExcludeAssignments=Chk 'chkExcludeAssignments'; IncludePolicyId=Chk 'chkIncludePolicyId'
        SkipDocumentInfo=Chk 'chkSkipDocumentInfo'; FallbackDocumentation=Chk 'chkFallbackDocumentation'
        NotConfiguredText=Sel 'cbNotConfiguredText'; ValueOutputProperty=Sel 'cbValueOutputProperty'
        Outputs=@{
            csv=@{CSVDocumentationPath=Txt 'txtCSVDocumentationPath';CSVExportProperties=Sel 'cbCSVDocumentationProperties';CSVCustomDisplayProperties=Txt 'txtCSVCustomProperties';CSVDelimiter=Sel 'cbCSVDelimiter';CSVAddObjectType=Chk 'chkCSVAddObjectType';CSVAddCompanyName=Chk 'chkCSVAddCompanyName'}
            json=@{JSONDocumentName=Txt 'txtJsonDocumentName';JSONOpenFile=Chk 'chkJsonOpenFile';JSONOutputFileType=Sel 'cbJsonOutputFileType'}
            html=@{HTMLDocumentName=Txt 'txtHTMLDocumentName';HTMLCSSFile=Txt 'txtHTMLCSSFile';HTMLTitleProperty=Txt 'txtHTMLTitleProperty';HTMLOpenFile=Chk 'chkHTMLOpenFile';HTMLDocumentFileType=Sel 'cbHTMLDocumentFileType'}
            atlassian=@{AtlassianDocumentName=Txt 'txtAtlassianDocumentName';AtlassianTitleProperty=Txt 'txtAtlassianTitleProperty';AtlassianOpenFile=Chk 'chkAtlassianOpenFile';AtlassianDocumentFileType=Sel 'cbAtlassianDocumentFileType'}
            md=@{MDDocumentName=Txt 'txtMDDocumentName';MDCSSFile=Txt 'txtMDCSSFile';MDTitleProperty=Txt 'txtMDTitleProperty';MDIncludeCSS=Chk 'chkMDIncludeCSS';MDOpenFile=Chk 'chkMDOpenFile';MDDocumentSkipDate=Chk 'chkMDDocumentSkipDate';MDDocumentFileType=Sel 'cbMDDocumentFileType'}
            word=@{WordDocumentName=Txt 'txtWordDocumentName';WordDocumentTemplate=Txt 'txtWordDocumentTemplate';WordExportProperties=Sel 'cbWordExportProperties';WordCustomDisplayProperties=Txt 'txtWordCustomDisplayProperties';WordDocumentFormat=Sel 'cbWordDocumentFormat';WordDocumentationLevel=Sel 'cbWordDocumentationLevel';WordDocumentationLimitMaxLength=Txt 'txtWordDocumentationLimitMaxLength';WordDocumentationLimitTruncateLength=Txt 'txtWordDocumentationLimitTruncateLength';WordDocumentationLimitAttach=Chk 'chkWordDocumentationLimitAttach';WordAddCategories=Chk 'chkWordAddCategories';WordAddSubCategories=Chk 'chkWordAddSubCategories';WordOpenDocument=Chk 'chkWordOpenDocument';WordAttachJsonFile=Chk 'chkWordAttachJsonFile';WordTitleProperty=Txt 'txtWordTitleProperty';WordSubjectProperty=Txt 'txtWordSubjectProperty';WordContentControls=Txt 'txtWordContentControls';WordCoverPage=Txt 'txtWordCoverPage';WordHeader1Style=Txt 'txtWordHeader1Style';WordHeader2Style=Txt 'txtWordHeader2Style';WordHeader3Style=Txt 'txtWordHeader3Style';WordTableStyle=Txt 'txtWordTableStyle';WordTableHeaderStyle=Txt 'txtWordTableHeaderStyle';WordCategoryHeaderStyle=Txt 'txtWordCategoryHeaderStyle';WordSubCategoryHeaderStyle=Txt 'txtWordSubCategoryHeaderStyle';WordTableTextStyle=Txt 'txtWordTableTextStyle';WordTableCaptionPosition=Sel 'cbWordTableCaptionPosition';WordScriptTableStyle=Txt 'txtWordScriptTableStyle';WordScriptStyle=Txt 'txtWordScriptStyle'}
        }
    }
}

function Add-WpfDocumentationFileBrowse {
    param($Form, [string]$ButtonName, [string]$TextName, [switch]$Save, [string]$Filter = 'All files (*.*)|*.*')
    $script:UIProvider.AddXamlEvent($Form, $ButtonName, 'add_click', {
        $dialog = if($Save) { [System.Windows.Forms.SaveFileDialog]::new() } else { [System.Windows.Forms.OpenFileDialog]::new() }
        $dialog.Filter = $Filter
        $current = [string]$Form.FindName($TextName).Text
        if($current) { $dialog.FileName = $current }
        if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $Form.FindName($TextName).Text = $dialog.FileName
        }
    }.GetNewClosure())
}

function Set-WpfDocumentationOutputPanel {
    param($Form, [string]$OutputValue)
    foreach($output in [DocumentationRegistry]::Outputs) {
        $panel = $Form.FindName("pnlDocumentationOutput$($output.Name)")
        if($panel) { $panel.Visibility = if($output.Value -eq $OutputValue) { 'Visible' } else { 'Collapsed' } }
    }
}

function Show-GraphBulkDocumentationForm
{
    param([object[]]$PolicyObject)

    $script:bulkDocForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\BulkDocumentationForm.xaml"), $true)
    if(-not $script:bulkDocForm) { return }

    $script:dgObjectsToDocument = $script:bulkDocForm.FindName("dgObjectsToDocument")

    # ---- Output-format combo populated from the registry ----
    $cbType = $script:bulkDocForm.FindName("cbDocumentationType")
    if($cbType) {
        $outputs = @([DocumentationRegistry]::Outputs |
            Sort-Object Name |
            ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Value = $_.Value } })
        $cbType.ItemsSource = $outputs
        if($outputs.Count -gt 0) {
            $last = Get-SettingStoreValue "Documentation" "OutputType" "json"
            $match = $outputs | Where-Object Value -eq $last | Select-Object -First 1
            if($match) { $cbType.SelectedValue = $match.Value }
            else { $cbType.SelectedIndex = 0 }
        }
        Set-WpfDocumentationOutputPanel $script:bulkDocForm ([string]$cbType.SelectedValue)
        $cbType.Add_SelectionChanged({
            param($source, $e)
            Set-WpfDocumentationOutputPanel $script:bulkDocForm ([string]$source.SelectedValue)
        })
    }

    # ---- Restore persisted form values ----
    $script:UIProvider.SetXamlProperty($script:bulkDocForm, "txtDocumentationOutputFolder", "Text", [string](Get-SettingStoreValue "Documentation" "OutputFolder" ""))
    $script:UIProvider.SetXamlProperty($script:bulkDocForm, "txtDocumentFromFolder", "Text", [string](Get-DocumentationSetting 'SourceFolder' ''))
    $script:UIProvider.SetXamlProperty($script:bulkDocForm, "txtDocumentFilter", "Text", [string](Get-DocumentationSetting 'NameFilter' ''))
    Initialize-WpfDocumentationOptions $script:bulkDocForm

    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "browseDocumentationOutputFolder", "add_click", {
        $current = $script:UIProvider.GetXamlProperty($script:bulkDocForm, "txtDocumentationOutputFolder", "Text")
        $folder = Get-Folder $current "Select root folder for documentation"
        if($folder) {
            $script:UIProvider.SetXamlProperty($script:bulkDocForm, "txtDocumentationOutputFolder", "Text", $folder)
        }
    })
    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "browseDocumentFromFolder", "add_click", {
        $current = $script:UIProvider.GetXamlProperty($script:bulkDocForm, "txtDocumentFromFolder", "Text")
        $folder = Get-Folder $current "Select exported source folder"
        if($folder) {
            $script:UIProvider.SetXamlProperty($script:bulkDocForm, "txtDocumentFromFolder", "Text", $folder)
        }
    })
    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "browseCSVDocumentationPath", "add_click", {
        $current = $script:UIProvider.GetXamlProperty($script:bulkDocForm, "txtCSVDocumentationPath", "Text")
        $folder = Get-Folder $current "Select CSV documentation folder"
        if($folder) { $script:UIProvider.SetXamlProperty($script:bulkDocForm, "txtCSVDocumentationPath", "Text", $folder) }
    })
    # Registry-driven browse-button wiring. Each output provider that has
    # file-picker inputs on the bulk-doc form attaches a FileBrowseControls
    # array to its [DocumentationRegistry] entry from a matching WPF
    # ClassExtensions file (UI/WPF/ClassExtensions/DocumentationOutput<X>UIExtension.ps1).
    # Adding a new provider with browse buttons requires no edit here — just
    # drop a ClassExtensions file that Add-Member's onto the provider.
    foreach($output in [DocumentationRegistry]::Outputs) {
        if(-not $output.PSObject.Properties['FileBrowseControls']) { continue }
        foreach($fb in @($output.FileBrowseControls)) {
            $params = @{ Form = $script:bulkDocForm; ButtonName = $fb.Button; TextName = $fb.TextBox }
            if($fb.Save)   { $params['Save']   = $true }
            if($fb.Filter) { $params['Filter'] = $fb.Filter }
            Add-WpfDocumentationFileBrowse @params
        }
    }

    # ---- Object list (same filter as Bulk Export) ----
    $script:bulkDocDocumentableGroups = @($script:IntuneGroups | Where-Object {
        $_.Title -and (-not ($_.ShowButtons -is [Object[]]) -or $_.ShowButtons -contains "Document")
    })
    $script:bulkDocPolicyObjects = Get-UIItemsArray $PolicyObject

    # Selected column with header-checkbox toggle-all
    $column = Get-GridCheckboxColumn "Selected"
    $script:dgObjectsToDocument.Columns.Add($column)
    $column.Header.IsChecked = $true
    $column.Header.add_Click({
        foreach($item in $script:dgObjectsToDocument.ItemsSource) {
            $item.Selected = $this.IsChecked
        }
        $script:dgObjectsToDocument.Items.Refresh()
    })

    # Title column
    $binding = [System.Windows.Data.Binding]::new("Title")
    $titleCol = [System.Windows.Controls.DataGridTextColumn]::new()
    $titleCol.Header = if($script:bulkDocPolicyObjects.Count -gt 0) { "Title" } else { "Object type" }
    $titleCol.IsReadOnly = $true
    $titleCol.Binding = $binding
    $script:dgObjectsToDocument.Columns.Add($titleCol)

    $script:bulkDocMode = if($script:bulkDocPolicyObjects.Count -gt 0) { "Object" } else { "Group" }
    $listBy = $script:bulkDocForm.FindName("pnlDocumentationListBy")
    if($listBy) { $listBy.Visibility = if($script:bulkDocMode -eq "Object") { "Collapsed" } else { "Visible" } }
    foreach($bulkOnlyControlName in @("lblDocumentFilter","txtDocumentFilter","lblDocumentFromFolder","pnlDocumentFromFolder")) {
        $bulkOnlyControl = $script:bulkDocForm.FindName($bulkOnlyControlName)
        if($bulkOnlyControl) { $bulkOnlyControl.Visibility = if($script:bulkDocMode -eq "Object") { "Collapsed" } else { "Visible" } }
    }
    Update-BulkDocumentationObjectList

    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "rbBulkDocViewGroup", "add_Checked", {
        $script:bulkDocMode = "Group"
        Update-BulkDocumentationObjectList
    })
    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "rbBulkDocViewType", "add_Checked", {
        $script:bulkDocMode = "Type"
        Update-BulkDocumentationObjectList
    })

    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "btnClose", "add_click", {
        $script:bulkDocForm = $null
        $script:bulkDocPolicyObjects = $null
        Show-ModalObject
    })

    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "btnDocument", "add_click", {
        try {
            Write-Status "Initializing documentation"
            $selection = if($script:bulkDocMode -eq "Object") { Get-BulkDocumentationSelectedObjects } else { Get-BulkDocumentationSelectedIds }
            $unit = if($script:bulkDocMode -eq "Object") { "policy" } elseif($script:bulkDocMode -eq "Type") { "policy type" } else { "object group" }
            if($selection.Count -eq 0) {
                $script:UIProvider.ShowMessageBox("Select at least one $unit to document.", "Bulk Documentation", "OK", "Warning") | Out-Null
                return
            }

            $cb = $script:bulkDocForm.FindName("cbDocumentationType")
            $outputValue = if($cb) { [string]$cb.SelectedValue } else { $null }
            if(-not $outputValue) {
                $script:UIProvider.ShowMessageBox("Select an output format.", "Bulk Documentation", "OK", "Warning") | Out-Null
                return
            }

            $rawFolder = $script:UIProvider.GetXamlProperty($script:bulkDocForm, "txtDocumentationOutputFolder", "Text")
            if($rawFolder) { $rawFolder = ([string]$rawFolder).Trim().Trim('"').Trim("'") }
            $options  = Get-WpfDocumentationOptions $script:bulkDocForm
            if(-not $rawFolder) {
                # Registry-driven fallback: consult the provider's declared
                # PrimaryPathOption to find its persisted path, and PathIsFolder to
                # decide whether the persisted value is already a folder (CSV) or a
                # file whose parent we need to take. Adding a provider requires no
                # change here — it just declares the two fields on registration.
                $provider = [DocumentationRegistry]::FindOutput($outputValue)
                $selectedPath = if($provider -and $provider.PSObject.Properties['PrimaryPathOption'] -and $provider.PrimaryPathOption) {
                    $providerOpts = $options.Outputs[$outputValue]
                    if($providerOpts) { $providerOpts[$provider.PrimaryPathOption] }
                }
                if($selectedPath) {
                    $rawFolder = if($provider -and $provider.PSObject.Properties['PathIsFolder'] -and $provider.PathIsFolder) { $selectedPath }
                                else { Split-Path -Parent $selectedPath }
                }
                if(-not $rawFolder) { $rawFolder = [Environment]::GetFolderPath('MyDocuments') }
            }
            try { $resolvedFolder = [IO.Path]::GetFullPath($rawFolder) }
            catch {
                $script:UIProvider.ShowMessageBox("Invalid output folder path: $($_.Exception.Message)", "Bulk Documentation", "OK", "Error") | Out-Null
                return
            }

            $language = [string]$options.Language
            if(-not $language) { $language = 'en' }
            $defaultName = "%Organization%-%Date%"
            if(-not $options.Outputs.csv.CSVDocumentationPath) { $options.Outputs.csv.CSVDocumentationPath = $resolvedFolder }
            if(-not $options.Outputs.json.JSONDocumentName) { $options.Outputs.json.JSONDocumentName = Join-Path $resolvedFolder "$defaultName.json" }
            if(-not $options.Outputs.html.HTMLDocumentName) { $options.Outputs.html.HTMLDocumentName = Join-Path $resolvedFolder "$defaultName.html" }
            # Confluence storage format is XHTML, so .html keeps it openable in a browser.
            if(-not $options.Outputs.atlassian.AtlassianDocumentName) { $options.Outputs.atlassian.AtlassianDocumentName = Join-Path $resolvedFolder "$defaultName.html" }
            if(-not $options.Outputs.md.MDDocumentName) { $options.Outputs.md.MDDocumentName = Join-Path $resolvedFolder "$defaultName.md" }
            if(-not $options.Outputs.word.WordDocumentName) { $options.Outputs.word.WordDocumentName = Join-Path $resolvedFolder "$defaultName.docx" }

            # Persist form state for next time
            Save-SettingStoreValue "Documentation" "OutputType"            $outputValue
            Save-SettingStoreValue "Documentation" "OutputFolder"          $resolvedFolder
            Save-DocumentationOptionsDefaults $options
            $startParams = @{
                OutputFormat = $outputValue
                Language     = $language
                Options      = $options
            }
            if($script:bulkDocMode -ne "Object" -and $options.SourceFolder) { $startParams.SourceFolder = $options.SourceFolder }
            if($script:bulkDocMode -eq "Type") {
                $startParams.PolicyType = $selection
            }
            elseif($script:bulkDocMode -eq "Group") {
                $startParams.PolicyGroup = $selection
            }

            Write-Log "Documentation starting. Mode=$script:bulkDocMode Format=$outputValue Folder=$resolvedFolder Selection=$($selection -join ',')"
            try {
                if($script:bulkDocMode -eq "Object") {
                    [void]($selection | Start-GraphBulkDocumentation @startParams)
                }
                else {
                    [void](Start-GraphBulkDocumentation @startParams)
                }
                Write-Status $null
                $script:UIProvider.ShowMessageBox(("Documentation finished. Output folder:`n{0}" -f $resolvedFolder), "Bulk Documentation", "OK", "Information") | Out-Null
            }
            catch {
                Write-LogError "Bulk documentation failed" $_.Exception
                $script:UIProvider.ShowMessageBox("Bulk documentation failed: $($_.Exception.Message)", "Bulk Documentation", "OK", "Error") | Out-Null
            }
        }
        finally {
            Write-Status $null
        }
    })

    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "btnPreviewDocumentation", "add_click", {
        $selection = if($script:bulkDocMode -eq "Object") { Get-BulkDocumentationSelectedObjects } else { Get-BulkDocumentationSelectedIds }
        if($selection.Count -eq 0) { return }
        $options = Get-WpfDocumentationOptions $script:bulkDocForm
        $params = if($script:bulkDocMode -eq 'Type') { @{PolicyType=$selection} } elseif($script:bulkDocMode -eq 'Group') { @{PolicyGroup=$selection} } else { @{} }
        if($script:bulkDocMode -eq 'Object') {
            $policies = $selection
        }
        elseif($options.SourceFolder) {
            $options.SourceTenantUnavailable = $true
            Initialize-DocumentationSourceTenantContext -SourceFolder $options.SourceFolder
            $policies = Get-DocumentationPoliciesFromSourceFolder -SourceFolder $options.SourceFolder @params
        } else {
            $policies = Get-GraphPolicies @params
        }
        $script:bulkDocPreview = @($policies | Select-Object -First 5 | Get-GraphDocumentation -Language $options.Language -Options $options)
        $script:bulkDocForm.FindName('txtDocumentationRawData').Text = ($script:bulkDocPreview | ConvertTo-Json -Depth 10)
    })
    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "btnCopyBasicDocumentation", "add_click", {
        @($script:bulkDocPreview | ForEach-Object BasicInfo) | Select-Object Name,Value | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
    })
    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "btnCopySettingsDocumentation", "add_click", {
        @($script:bulkDocPreview | ForEach-Object FilteredSettings) | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
    })
    $script:UIProvider.AddXamlEvent($script:bulkDocForm, "btnCopyRawDocumentation", "add_click", {
        $script:bulkDocForm.FindName('txtDocumentationRawData').Text | Set-Clipboard
    })

    $title = if($script:bulkDocMode -eq 'Object') { 'Document' } else { 'Bulk Documentation' }
    $script:UIProvider.ShowModalForm($title, $script:bulkDocForm, $true)
}

function Update-BulkDocumentationObjectList
{
    if(-not $script:dgObjectsToDocument) { return }

    $script:documentObjects = @()

    if($script:bulkDocMode -eq "Object") {
        foreach($policy in ($script:bulkDocPolicyObjects | Sort-Object Name)) {
            $script:documentObjects += New-Object PSObject -Property @{
                Title        = $policy.Name
                Selected     = $true
                ObjectGroup  = $null
                ObjectType   = $policy.PolicyType
                PolicyObject = $policy
            }
        }
    }
    elseif($script:bulkDocMode -eq "Type") {
        $documentableGroupIds = @($script:bulkDocDocumentableGroups | ForEach-Object { $_.Id })
        $sortedTypes = $script:IntuneTypes |
            Where-Object { $_.PolicyGroup -and ($documentableGroupIds -contains $_.PolicyGroup.Id) } |
            Sort-Object Title
        foreach($intuneType in $sortedTypes) {
            $script:documentObjects += New-Object PSObject -Property @{
                Title       = $intuneType.Title
                Selected    = $true
                ObjectGroup = $null
                ObjectType  = $intuneType
            }
        }
    }
    else {
        $sortedGroups = $script:bulkDocDocumentableGroups | Sort-Object Title
        foreach($intuneGroup in $sortedGroups) {
            $script:documentObjects += New-Object PSObject -Property @{
                Title       = $intuneGroup.Title
                Selected    = $true
                ObjectGroup = $intuneGroup
                ObjectType  = $null
            }
        }
    }

    $script:dgObjectsToDocument.ItemsSource = $script:documentObjects
}

function Get-BulkDocumentationSelectedIds
{
    if($script:bulkDocMode -eq "Type") {
        return @($script:documentObjects |
            Where-Object { $_.Selected -eq $true -and $_.ObjectType -and $_.ObjectType.Id } |
            ForEach-Object { $_.ObjectType.Id })
    }
    return @($script:documentObjects |
        Where-Object { $_.Selected -eq $true -and $_.ObjectGroup -and $_.ObjectGroup.Id } |
        ForEach-Object { $_.ObjectGroup.Id })
}

function Get-BulkDocumentationSelectedObjects
{
    return @($script:documentObjects |
        Where-Object { $_.Selected -eq $true -and $_.PolicyObject } |
        ForEach-Object { $_.PolicyObject })
}

# Compatibility wrapper for the main-view Document button. The shared form
# switches to Object mode and displays the selected policies.
function Show-IntuneManagerDocumentForm
{
    param([object[]]$PolicyObject)

    if (-not $PolicyObject -or $PolicyObject.Count -eq 0) {
        $script:UIProvider.ShowMessageBox("No items to document.", "Document", "OK", "Warning") | Out-Null
        return
    }

    Show-GraphBulkDocumentationForm -PolicyObject $PolicyObject
}

#endregion

#endregion
