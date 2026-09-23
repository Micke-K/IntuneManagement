# Avalonia bulk-documentation dialog. New file because Bulk Documentation is a
# self-contained subsystem (R9 — keep IntuneManagerUIAvalonia from growing
# further) and the WPF original at UI/WPF/Extensions/IntuneManagerUI.ps1 isn't
# the only consumer; the same Start-GraphBulkDocumentation function it calls
# also drives the silent batch path (Tests/Setup/Invoke-AutomationOrchestrator)
# so the UI logic stays thin.

function Update-BulkDocumentationObjectList
{
    if (-not $script:dgBulkDocObjects) { return }

    $rows = [System.Collections.Generic.List[BulkDocumentationRowItem]]::new()

    if ($script:bulkDocMode -eq "Object") {
        foreach ($policy in @($script:bulkDocPolicyObjects | Sort-Object Name)) {
            $rows.Add([BulkDocumentationRowItem]@{
                Title        = [string]$policy.Name
                Selected     = $true
                ObjectGroup  = $null
                ObjectType   = $policy.PolicyType
                PolicyObject = $policy
            })
        }
    } elseif ($script:bulkDocMode -eq "Type") {
        $documentableGroupIds = @($script:bulkDocDocumentableGroups | ForEach-Object { $_.Id })
        $sortedTypes = $script:IntuneTypes |
            Where-Object { $_.PolicyGroup -and ($documentableGroupIds -contains $_.PolicyGroup.Id) } |
            Sort-Object Title
        foreach ($intuneType in $sortedTypes) {
            $rows.Add([BulkDocumentationRowItem]@{
                Title       = [string]$intuneType.Title
                Selected    = $true
                ObjectGroup = $null
                ObjectType  = $intuneType
            })
        }
    } else {
        $sortedGroups = $script:bulkDocDocumentableGroups | Sort-Object Title
        foreach ($intuneGroup in $sortedGroups) {
            $rows.Add([BulkDocumentationRowItem]@{
                Title       = [string]$intuneGroup.Title
                Selected    = $true
                ObjectGroup = $intuneGroup
                ObjectType  = $null
            })
        }
    }

    $script:bulkDocRows = @($rows)
    $script:dgBulkDocObjects.ItemsSource = $script:bulkDocRows
}

function Get-BulkDocumentationSelectedIds
{
    if ($null -eq $script:bulkDocRows) { return @() }
    if ($script:bulkDocMode -eq "Type") {
        return @($script:bulkDocRows |
            Where-Object { $_.Selected -and $_.ObjectType -and $_.ObjectType.Id } |
            ForEach-Object { $_.ObjectType.Id })
    }
    return @($script:bulkDocRows |
        Where-Object { $_.Selected -and $_.ObjectGroup -and $_.ObjectGroup.Id } |
        ForEach-Object { $_.ObjectGroup.Id })
}

function Get-BulkDocumentationSelectedObjects
{
    if ($null -eq $script:bulkDocRows) { return @() }
    return @($script:bulkDocRows |
        Where-Object { $_.Selected -and $_.PolicyObject } |
        ForEach-Object { $_.PolicyObject })
}

function ConvertTo-AvaloniaComboItem {
    # Wrap whatever the caller passes (string / PSCustomObject{Name,Value} /
    # PSCustomObject{EnglishName,Name} for languages) in a NameValueComboItem
    # so DisplayMemberBinding="{Binding Name|EnglishName}" finds a real CLR
    # property to render. SelectedItem.Value is preserved for the existing
    # readers; .Source carries the original object for callers that need
    # extra fields (Get-DocumentationLanguages returns CultureInfo).
    param($Item)
    if ($null -eq $Item) { return $null }
    if ($Item -is [NameValueComboItem]) { return $Item }

    $row = [NameValueComboItem]::new()
    $row.Source = $Item
    if ($Item -is [string]) {
        $row.Name = $Item; $row.EnglishName = $Item; $row.Value = $Item
        return $row
    }
    if ($Item.PSObject.Properties['EnglishName']) {
        $row.Name        = [string]$Item.EnglishName
        $row.EnglishName = [string]$Item.EnglishName
        $row.Value       = if ($Item.PSObject.Properties['Value']) { $Item.Value } elseif ($Item.PSObject.Properties['Name']) { $Item.Name } else { $Item.EnglishName }
        return $row
    }
    if ($Item.PSObject.Properties['Name']) {
        $row.Name        = [string]$Item.Name
        $row.EnglishName = [string]$Item.Name
        $row.Value       = if ($Item.PSObject.Properties['Value']) { $Item.Value } else { $Item.Name }
        return $row
    }
    $row.Name = [string]$Item; $row.EnglishName = [string]$Item; $row.Value = $Item
    return $row
}

function Set-AvaloniaDocumentationCombo {
    param($HostType, $Form, [string]$Name, $Items, $SelectedValue)
    $control = $HostType::FindByName($Form, $Name)
    if (-not $control) { return }
    $projected = @($Items | ForEach-Object { ConvertTo-AvaloniaComboItem $_ })
    $control.ItemsSource = $projected
    # Display field comes from DisplayMemberBinding on the control in
    # BulkDocumentationForm.axaml - nothing to set here.
    $match = @($projected | Where-Object { $_.Value -eq $SelectedValue }) | Select-Object -First 1
    $control.SelectedItem = if ($match) { $match } elseif ($projected.Count -gt 0) { $projected[0] } else { $null }
}

# Avalonia parity for Set-WpfDocumentationConditionalVisibility — show/hide
# $TargetNames whenever $TriggerName's selected value equals $VisibleWhen.
# State goes in $script:_docCondVisibility so the event handler can re-resolve
# controls by name; closures captured via GetNewClosure() get scrubbed by
# ConvertTo-AvaloniaEventScriptBlock (see feedback_avalonia_closure_dynamic_module).
function Set-AvaloniaDocumentationConditionalVisibility {
    param(
        $HostType, $Form,
        [string]$TriggerName,
        [string]$VisibleWhen,
        [string[]]$TargetNames
    )
    $trigger = $HostType::FindByName($Form, $TriggerName)
    if (-not $trigger) { return }
    $targets = @($TargetNames | ForEach-Object { $HostType::FindByName($Form, $_) } | Where-Object { $_ })
    if ($targets.Count -eq 0) { return }

    if (-not $script:_docCondVisibility) { $script:_docCondVisibility = @{} }
    $script:_docCondVisibility[$TriggerName] = [PSCustomObject]@{
        VisibleWhen = $VisibleWhen
        TargetNames = $TargetNames
        HostType    = $HostType
        Form        = $Form
    }

    Update-AvaloniaDocumentationConditionalVisibility -TriggerName $TriggerName
    $trigger.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
        # $this is the trigger control in Avalonia event handlers.
        Update-AvaloniaDocumentationConditionalVisibility -TriggerName $this.Name
    }))
}

function Update-AvaloniaDocumentationConditionalVisibility {
    param([string]$TriggerName)
    if (-not $script:_docCondVisibility -or -not $script:_docCondVisibility.ContainsKey($TriggerName)) { return }
    $entry   = $script:_docCondVisibility[$TriggerName]
    $trigger = $entry.HostType::FindByName($entry.Form, $TriggerName)
    if (-not $trigger) { return }
    $item = $trigger.SelectedItem
    $current = if ($null -eq $item) { $null }
               elseif ($item.PSObject.Properties['Value']) { $item.Value }
               elseif ($item.PSObject.Properties['Name'])  { $item.Name }
               else { $item }
    $visible = [string]$current -eq $entry.VisibleWhen
    foreach ($name in $entry.TargetNames) {
        $t = $entry.HostType::FindByName($entry.Form, $name)
        if ($t) { $t.IsVisible = $visible }
    }
}

function Initialize-AvaloniaDocumentationOptions {
    param($HostType, $Form)
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbDocumentationLanguage' (Get-DocumentationLanguages | Sort-Object EnglishName) (Get-DocumentationSetting 'Language' 'en')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbDocumentationPropertySeparator' @(',',';','-','|') (Get-DocumentationSetting 'PropertySeparator' ';')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbDocumentationObjectSeparator' @([PSCustomObject]@{Name='New line';Value=[Environment]::NewLine},[PSCustomObject]@{Name=';';Value=';'},[PSCustomObject]@{Name='|';Value='|'}) (Get-DocumentationSetting 'ObjectSeparator' ([Environment]::NewLine))
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbNotConfiguredText' @([PSCustomObject]@{Name='Not configured (localized)';Value='notConfigured'},[PSCustomObject]@{Name='Empty';Value='empty'},[PSCustomObject]@{Name="Don't change";Value='asis'}) (Get-DocumentationSetting 'NotConfiguredText' 'notConfigured')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbValueOutputProperty' @([PSCustomObject]@{Name='Value';Value='value'},[PSCustomObject]@{Name='Value with label';Value='valueWithLabel'}) (Get-DocumentationSetting 'ValueOutputProperty' 'value')
    $propertyModes = @([PSCustomObject]@{Name='Simple';Value='simple'},[PSCustomObject]@{Name='Extended';Value='extended'},[PSCustomObject]@{Name='Custom';Value='custom'})
    $fileModes = @([PSCustomObject]@{Name='Single file';Value='Full'},[PSCustomObject]@{Name='One file per object';Value='Object'})
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbCSVDocumentationProperties' $propertyModes (Get-DocumentationSetting 'CSVExportProperties' 'simple')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbCSVDelimiter' @('',',',';','-','|') (Get-DocumentationSetting 'CSVDelimiter' '')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbJsonOutputFileType' @([PSCustomObject]@{Name='Single file';Value='Full'},[PSCustomObject]@{Name='One file per object type';Value='ObjectType'}) (Get-DocumentationSetting 'JSONOutputFileType' 'Full')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbHTMLDocumentFileType' $fileModes (Get-DocumentationSetting 'HTMLDocumentFileType' 'Full')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbAtlassianDocumentFileType' $fileModes (Get-DocumentationSetting 'AtlassianDocumentFileType' 'Full')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbMDDocumentFileType' $fileModes (Get-DocumentationSetting 'MDDocumentFileType' 'Full')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbWordExportProperties' $propertyModes (Get-DocumentationSetting 'WordExportProperties' 'simple')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbWordDocumentFormat' @([PSCustomObject]@{Name='DOCX';Value='wdFormatDocumentDefault'},[PSCustomObject]@{Name='Strict Open XML';Value='wdFormatStrictOpenXMLDocument'},[PSCustomObject]@{Name='PDF';Value='wdFormatPDF'}) (Get-DocumentationSetting 'WordDocumentFormat' 'wdFormatDocumentDefault')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbWordDocumentationLevel' @([PSCustomObject]@{Name='Full';Value='full'},[PSCustomObject]@{Name='Limited';Value='limited'},[PSCustomObject]@{Name='Basic';Value='basic'}) (Get-DocumentationSetting 'WordDocumentationLevel' 'full')
    Set-AvaloniaDocumentationCombo $HostType $Form 'cbWordTableCaptionPosition' @([PSCustomObject]@{Name='Below table';Value='below'},[PSCustomObject]@{Name='Above table';Value='above'}) (Get-DocumentationSetting 'WordTableCaptionPosition' 'below')
    $text = @{txtCSVDocumentationPath='CSVDocumentationPath';txtCSVCustomProperties='CSVCustomDisplayProperties';txtJsonDocumentName='JSONDocumentName';txtHTMLDocumentName='HTMLDocumentName';txtHTMLCSSFile='HTMLCSSFile';txtHTMLTitleProperty='HTMLTitleProperty';txtAtlassianDocumentName='AtlassianDocumentName';txtAtlassianTitleProperty='AtlassianTitleProperty';txtMDDocumentName='MDDocumentName';txtMDCSSFile='MDCSSFile';txtMDTitleProperty='MDTitleProperty';txtWordDocumentName='WordDocumentName';txtWordDocumentTemplate='WordDocumentTemplate';txtWordCustomDisplayProperties='WordCustomDisplayProperties';txtWordDocumentationLimitMaxLength='WordDocumentationLimitMaxLength';txtWordDocumentationLimitTruncateLength='WordDocumentationLimitTruncateLength';txtWordTitleProperty='WordTitleProperty';txtWordSubjectProperty='WordSubjectProperty';txtWordContentControls='WordContentControls';txtWordCoverPage='WordCoverPage';txtWordHeader1Style='WordHeader1Style';txtWordHeader2Style='WordHeader2Style';txtWordHeader3Style='WordHeader3Style';txtWordTableStyle='WordTableStyle';txtWordTableHeaderStyle='WordTableHeaderStyle';txtWordCategoryHeaderStyle='WordCategoryHeaderStyle';txtWordSubCategoryHeaderStyle='WordSubCategoryHeaderStyle';txtWordTableTextStyle='WordTableTextStyle';txtWordScriptTableStyle='WordScriptTableStyle';txtWordScriptStyle='WordScriptStyle'}
    foreach($name in $text.Keys) {
        $c = $HostType::FindByName($Form, $name)
        # Skip when an AXAML file doesn't carry the field yet — assigning to a
        # null control would throw NullReferenceException and abort the whole
        # form load.
        if ($c) { $c.Text = [string](Get-DocumentationSetting $text[$name] '') }
    }
    foreach($pair in @(@('chkSkipNotConfigured','SkipNotConfigured',$false),@('chkSkipDefaultValues','SkipDefaultValues',$false),@('chkSkipDisabled','SkipDisabled',$true),@('chkSetUnconfiguredValue','SetUnconfiguredValue',$true),@('chkSetDefaultValue','SetDefaultValue',$false),@('chkIncludeScripts','IncludeScripts',$true),@('chkExcludeScriptSignature','ExcludeScriptSignature',$false),@('chkExcludeAssignments','ExcludeAssignments',$false),@('chkIncludePolicyId','IncludePolicyId',$false),@('chkSkipDocumentInfo','SkipDocumentInfo',$false),@('chkFallbackDocumentation','FallbackDocumentation',$true),@('chkCSVAddObjectType','CSVAddObjectType',$true),@('chkCSVAddCompanyName','CSVAddCompanyName',$false),@('chkJsonOpenFile','JSONOpenFile',$true),@('chkHTMLOpenFile','HTMLOpenFile',$true),@('chkAtlassianOpenFile','AtlassianOpenFile',$true),@('chkMDIncludeCSS','MDIncludeCSS',$true),@('chkMDOpenFile','MDOpenFile',$true),@('chkMDDocumentSkipDate','MDDocumentSkipDate',$false),@('chkWordDocumentationLimitAttach','WordDocumentationLimitAttach',$false),@('chkWordAddCategories','WordAddCategories',$true),@('chkWordAddSubCategories','WordAddSubCategories',$true),@('chkWordOpenDocument','WordOpenDocument',$true),@('chkWordAttachJsonFile','WordAttachJsonFile',$false))) {
        $c = $HostType::FindByName($Form, $pair[0])
        if ($c) { $c.IsChecked = (Get-DocumentationSetting $pair[1] $pair[2]) -eq $true }
    }

    # Conditional-visibility wiring (Word custom-properties row, Word
    # limit-options grid, CSV custom-properties row). Mirrors the WPF
    # Set-WpfDocumentationConditionalVisibility calls.
    Set-AvaloniaDocumentationConditionalVisibility -HostType $HostType -Form $Form -TriggerName 'cbWordExportProperties' -VisibleWhen 'custom' `
        -TargetNames @('spWordCustomProperties','txtWordCustomDisplayProperties')
    Set-AvaloniaDocumentationConditionalVisibility -HostType $HostType -Form $Form -TriggerName 'cbWordDocumentationLevel' -VisibleWhen 'limited' `
        -TargetNames @('gdWordDocumentationLimitOptions')
    Set-AvaloniaDocumentationConditionalVisibility -HostType $HostType -Form $Form -TriggerName 'cbCSVDocumentationProperties' -VisibleWhen 'custom' `
        -TargetNames @('spCSVCustomProperties','txtCSVCustomProperties')
}

function Get-AvaloniaDocumentationOptions {
    param($HostType, $Form)
    function C($n) { $HostType::FindByName($Form,$n) }
    # Return $null when the control is missing so Save-DocumentationOptionsDefaults
    # skips writing — avoids erasing a real persisted value with an empty string
    # from a non-existent XAML field.
    function Txt($n) { $c = C $n; if ($c) { [string]$c.Text } else { $null } }
    function Sel($n) { $c = C $n; if (-not $c) { return $null }; $v=$c.SelectedItem; if($v -and $v.PSObject.Properties['Value']){$v.Value}elseif($v -and $v.PSObject.Properties['Name']){$v.Name}else{$v} }
    function Chk($n) { $c = C $n; if ($c) { [bool]$c.IsChecked } else { $null } }
    @{Language=Sel 'cbDocumentationLanguage';PropertySeparator=Sel 'cbDocumentationPropertySeparator';ObjectSeparator=Sel 'cbDocumentationObjectSeparator';NameFilter=Txt 'txtDocumentFilter';SourceFolder=Txt 'txtDocumentFromFolder';SkipNotConfigured=Chk 'chkSkipNotConfigured';SkipDefaultValues=Chk 'chkSkipDefaultValues';SkipDisabled=Chk 'chkSkipDisabled';SetUnconfiguredValue=Chk 'chkSetUnconfiguredValue';SetDefaultValue=Chk 'chkSetDefaultValue';IncludeScripts=Chk 'chkIncludeScripts';ExcludeScriptSignature=Chk 'chkExcludeScriptSignature';ExcludeAssignments=Chk 'chkExcludeAssignments';IncludePolicyId=Chk 'chkIncludePolicyId';SkipDocumentInfo=Chk 'chkSkipDocumentInfo'; FallbackDocumentation=Chk 'chkFallbackDocumentation';NotConfiguredText=Sel 'cbNotConfiguredText';ValueOutputProperty=Sel 'cbValueOutputProperty';Outputs=@{csv=@{CSVDocumentationPath=Txt 'txtCSVDocumentationPath';CSVExportProperties=Sel 'cbCSVDocumentationProperties';CSVCustomDisplayProperties=Txt 'txtCSVCustomProperties';CSVDelimiter=Sel 'cbCSVDelimiter';CSVAddObjectType=Chk 'chkCSVAddObjectType';CSVAddCompanyName=Chk 'chkCSVAddCompanyName'};json=@{JSONDocumentName=Txt 'txtJsonDocumentName';JSONOpenFile=Chk 'chkJsonOpenFile';JSONOutputFileType=Sel 'cbJsonOutputFileType'};html=@{HTMLDocumentName=Txt 'txtHTMLDocumentName';HTMLCSSFile=Txt 'txtHTMLCSSFile';HTMLTitleProperty=Txt 'txtHTMLTitleProperty';HTMLOpenFile=Chk 'chkHTMLOpenFile';HTMLDocumentFileType=Sel 'cbHTMLDocumentFileType'};atlassian=@{AtlassianDocumentName=Txt 'txtAtlassianDocumentName';AtlassianTitleProperty=Txt 'txtAtlassianTitleProperty';AtlassianOpenFile=Chk 'chkAtlassianOpenFile';AtlassianDocumentFileType=Sel 'cbAtlassianDocumentFileType'};md=@{MDDocumentName=Txt 'txtMDDocumentName';MDCSSFile=Txt 'txtMDCSSFile';MDTitleProperty=Txt 'txtMDTitleProperty';MDIncludeCSS=Chk 'chkMDIncludeCSS';MDOpenFile=Chk 'chkMDOpenFile';MDDocumentSkipDate=Chk 'chkMDDocumentSkipDate';MDDocumentFileType=Sel 'cbMDDocumentFileType'};word=@{WordDocumentName=Txt 'txtWordDocumentName';WordDocumentTemplate=Txt 'txtWordDocumentTemplate';WordExportProperties=Sel 'cbWordExportProperties';WordCustomDisplayProperties=Txt 'txtWordCustomDisplayProperties';WordDocumentFormat=Sel 'cbWordDocumentFormat';WordDocumentationLevel=Sel 'cbWordDocumentationLevel';WordDocumentationLimitMaxLength=Txt 'txtWordDocumentationLimitMaxLength';WordDocumentationLimitTruncateLength=Txt 'txtWordDocumentationLimitTruncateLength';WordDocumentationLimitAttach=Chk 'chkWordDocumentationLimitAttach';WordAddCategories=Chk 'chkWordAddCategories';WordAddSubCategories=Chk 'chkWordAddSubCategories';WordOpenDocument=Chk 'chkWordOpenDocument';WordAttachJsonFile=Chk 'chkWordAttachJsonFile';WordTitleProperty=Txt 'txtWordTitleProperty';WordSubjectProperty=Txt 'txtWordSubjectProperty';WordContentControls=Txt 'txtWordContentControls';WordCoverPage=Txt 'txtWordCoverPage';WordHeader1Style=Txt 'txtWordHeader1Style';WordHeader2Style=Txt 'txtWordHeader2Style';WordHeader3Style=Txt 'txtWordHeader3Style';WordTableStyle=Txt 'txtWordTableStyle';WordTableHeaderStyle=Txt 'txtWordTableHeaderStyle';WordCategoryHeaderStyle=Txt 'txtWordCategoryHeaderStyle';WordSubCategoryHeaderStyle=Txt 'txtWordSubCategoryHeaderStyle';WordTableTextStyle=Txt 'txtWordTableTextStyle';WordTableCaptionPosition=Sel 'cbWordTableCaptionPosition';WordScriptTableStyle=Txt 'txtWordScriptTableStyle';WordScriptStyle=Txt 'txtWordScriptStyle'}}}
}

function Add-AvaloniaDocumentationFileBrowse {
    param($HostType, $Form, [string]$ButtonName, [string]$TextName, [switch]$Save, [string]$Title, [string]$Extension, [string]$FilterName, [string]$FilterPattern)
    $button = $HostType::FindByName($Form, $ButtonName)
    $text = $HostType::FindByName($Form, $TextName)
    if(-not $button -or -not $text) { return }

    # Per-button context rides on the sender's Tag: this helper is called once
    # per registry-declared picker (7+ times), so a single module-scope slot
    # would be overwritten by each successive call - and the captured locals it
    # used before were stripped, leaving every provider's browse button dead.
    $button.Tag = [PSCustomObject]@{
        TextBox = $text; Save = [bool]$Save; Title = $Title
        Extension = $Extension; FilterName = $FilterName; FilterPattern = $FilterPattern
    }
    $button.add_Click((ConvertTo-AvaloniaEventScriptBlock {
        param($S, $E)
        $cfg = $S.Tag
        if(-not $cfg -or -not $cfg.TextBox) { return }
        $hostT = Get-AvaloniaHost
        $current = [string]$cfg.TextBox.Text
        if($cfg.Save) {
            $suggested = if($current) { [IO.Path]::GetFileName($current) } else { '' }
            $picked = $hostT::SaveFilePicker($script:Window, $cfg.Title, $suggested, $cfg.Extension, $cfg.FilterName, $cfg.FilterPattern)
        } else {
            $start = if($current) { [IO.Path]::GetDirectoryName($current) } else { '' }
            $picked = $hostT::OpenFilePicker($script:Window, $cfg.Title, $start, $cfg.FilterName, $cfg.FilterPattern)
        }
        if($picked) { $cfg.TextBox.Text = $picked }
    }))
}

function Set-AvaloniaDocumentationOutputPanel {
    param($HostType, $Form, [string]$OutputValue)
    foreach($output in [DocumentationRegistry]::Outputs) {
        $panel = $HostType::FindByName($Form, "pnlDocumentationOutput$($output.Name)")
        if($panel) { $panel.IsVisible = ($output.Value -eq $OutputValue) }
    }
}

function Show-GraphBulkDocumentationForm
{
    param([object[]]$PolicyObject)

    # $script: scope-prefixed references and bare module-private function
    # CORRECTION (2026-08-25): this note previously claimed the opposite.
    # $script: variables and bare module-function lookups DO resolve inside
    # Avalonia handlers; it is function-LOCAL captures that get stripped by
    # ConvertTo-AvaloniaEventScriptBlock. Handler state therefore lives in
    # $script:_bulkDocFormState, never in captured locals.
    $ui = $script:UIProvider

    $form = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML\BulkDocumentationForm.axaml'))
    if (-not $form) { return }

    $hostType = Get-AvaloniaHost
    $script:dgBulkDocObjects = $hostType::FindByName($form, 'dgObjectsToDocument')

    $cbType       = $hostType::FindByName($form, 'cbDocumentationType')
    $txtFolder    = $hostType::FindByName($form, 'txtDocumentationOutputFolder')
    $btnBrowse    = $hostType::FindByName($form, 'browseDocumentationOutputFolder')
    $txtSource    = $hostType::FindByName($form, 'txtDocumentFromFolder')
    $btnSource    = $hostType::FindByName($form, 'browseDocumentFromFolder')
    $txtLang      = $hostType::FindByName($form, 'cbDocumentationLanguage')
    $chkIncScr    = $hostType::FindByName($form, 'chkIncludeScripts')
    $chkExcSig    = $hostType::FindByName($form, 'chkExcludeScriptSignature')
    $chkExcAssign = $hostType::FindByName($form, 'chkExcludeAssignments')
    $chkIncPolId  = $hostType::FindByName($form, 'chkIncludePolicyId')
    $rbGroup      = $hostType::FindByName($form, 'rbBulkDocViewGroup')
    $rbType       = $hostType::FindByName($form, 'rbBulkDocViewType')
    $pnlListBy    = $hostType::FindByName($form, 'pnlDocumentationListBy')
    # The label rows are Label + info-icon inside a horizontal StackPanel, so the
    # wrapper is what has to be hidden - toggling the inner Label alone leaves the
    # panel occupying its row with the info icon still showing.
    $pnlFilterLbl = $hostType::FindByName($form, 'pnlDocumentFilterLabel')
    $txtFilter    = $hostType::FindByName($form, 'txtDocumentFilter')
    $pnlSourceLbl = $hostType::FindByName($form, 'pnlDocumentFromFolderLabel')
    $pnlSource    = $hostType::FindByName($form, 'pnlDocumentFromFolder')
    $btnDoc       = $hostType::FindByName($form, 'btnDocument')
    $btnClose     = $hostType::FindByName($form, 'btnClose')
    $btnPreview   = $hostType::FindByName($form, 'btnPreviewDocumentation')
    $btnCopyRaw   = $hostType::FindByName($form, 'btnCopyRawDocumentation')
    $btnCopyBasic = $hostType::FindByName($form, 'btnCopyBasicDocumentation')
    $btnCopySettings = $hostType::FindByName($form, 'btnCopySettingsDocumentation')
    $txtRaw       = $hostType::FindByName($form, 'txtDocumentationRawData')

    $script:_bulkDocFormState = [ordered]@{
        HostType = $hostType; Form = $form
        CbType = $cbType; TxtFolder = $txtFolder; TxtSource = $txtSource
        TxtLang = $txtLang; ChkIncScr = $chkIncScr; ChkExcSig = $chkExcSig
        ChkExcAssign = $chkExcAssign; ChkIncPolId = $chkIncPolId
        TxtFilter = $txtFilter; TxtRaw = $txtRaw; HeaderCb = $null
    }
    $script:bulkDocPreview = @()

    # ---- Populate output-format combo from the registry ----
    if ($cbType) {
        # Project to NameValueComboItem so Avalonia's DisplayMemberBinding can
        # read .Name off a real CLR property (PSCustomObject NoteProperties
        # don't bind — see [[avalonia-binding-needs-clr-types]]).
        $outputs = @([DocumentationRegistry]::Outputs |
            Sort-Object Name |
            ForEach-Object { ConvertTo-AvaloniaComboItem ([PSCustomObject]@{ Name = $_.Name; Value = $_.Value }) })
        $cbType.ItemsSource = $outputs
        if ($outputs.Count -gt 0) {
            $last = Get-SettingStoreValue "Documentation" "OutputType" "json"
            $match = $outputs | Where-Object Value -EQ $last | Select-Object -First 1
            $cbType.SelectedItem = if ($match) { $match } else { $outputs[0] }
        }
        $selectedOutput = if($cbType.SelectedItem) { [string]$cbType.SelectedItem.Value } else { '' }
        Set-AvaloniaDocumentationOutputPanel $hostType $form $selectedOutput
        $cbType.add_SelectionChanged((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkDocFormState
            if (-not $st) { return }
            $value = if($st.CbType -and $st.CbType.SelectedItem) { [string]$st.CbType.SelectedItem.Value } else { '' }
            Set-AvaloniaDocumentationOutputPanel $st.HostType $st.Form $value
        }))
    }

    # ---- Restore persisted form values ----
    if ($txtFolder)   { $txtFolder.Text   = [string](Get-SettingStoreValue "Documentation" "OutputFolder" "") }
    if ($txtSource)   { $txtSource.Text   = [string](Get-DocumentationSetting 'SourceFolder' '') }
    if ($txtFilter)   { $txtFilter.Text   = [string](Get-DocumentationSetting 'NameFilter' '') }
    Initialize-AvaloniaDocumentationOptions $hostType $form

    # ---- Populate object list (same Groups filter as Bulk Export) ----
    $script:bulkDocDocumentableGroups = @($script:IntuneGroups | Where-Object {
        $_.Title -and (-not ($_.ShowButtons -is [Object[]]) -or $_.ShowButtons -contains "Document")
    })
    $script:bulkDocPolicyObjects = @($PolicyObject | Where-Object { $null -ne $_ })
    $script:bulkDocMode = if($script:bulkDocPolicyObjects.Count -gt 0) { "Object" } else { "Group" }
    if($pnlListBy) { $pnlListBy.IsVisible = ($script:bulkDocMode -ne "Object") }
    foreach($bulkOnlyControl in @($pnlFilterLbl,$txtFilter,$pnlSourceLbl,$pnlSource)) {
        if($bulkOnlyControl) { $bulkOnlyControl.IsVisible = ($script:bulkDocMode -ne "Object") }
    }
    Update-BulkDocumentationObjectList

    if ($btnBrowse) {
        $btnBrowse.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkDocFormState
            if (-not $st) { return }
            $folder = $script:UIProvider.ShowFolderPicker(([string]$st.TxtFolder.Text), 'Select root folder for documentation')
            if ($folder -and $st.TxtFolder) { $st.TxtFolder.Text = $folder }
        }))
    }
    if ($btnSource) {
        $btnSource.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkDocFormState
            if (-not $st) { return }
            $folder = $script:UIProvider.ShowFolderPicker(([string]$st.TxtSource.Text), 'Select exported source folder')
            if ($folder -and $st.TxtSource) { $st.TxtSource.Text = $folder }
        }))
    }
    $btnCSVFolder = $hostType::FindByName($form, 'browseCSVDocumentationPath')
    if ($btnCSVFolder) {
        $btnCSVFolder.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkDocFormState
            if (-not $st) { return }
            $tb = $st.HostType::FindByName($st.Form, 'txtCSVDocumentationPath')
            $current = if ($tb) { [string]$tb.Text } else { '' }
            $folder = $script:UIProvider.ShowFolderPicker($current, 'Select CSV documentation folder')
            if ($folder -and $tb) { $tb.Text = $folder }
        }))
    }
    # Registry-driven file-browse wiring: each output provider optionally declares
    # a FileBrowseControls array in its Add-DocumentationOutputProvider hashtable
    # (see e.g. DocumentationOutputHTML.ps1). Any browse button in the XAML whose
    # name matches an entry gets wired here — adding a new provider does not
    # require editing this file. Buttons whose XAML controls are absent (e.g. a
    # provider that declares controls the current XAML doesn't include) are
    # ignored by Add-AvaloniaDocumentationFileBrowse's null-check.
    foreach($output in [DocumentationRegistry]::Outputs) {
        if(-not $output.PSObject.Properties['FileBrowseControls']) { continue }
        foreach($fb in @($output.FileBrowseControls)) {
            $params = @{
                HostType      = $hostType
                Form          = $form
                ButtonName    = $fb.Button
                TextName      = $fb.TextBox
                Title         = $fb.Title
                FilterName    = $fb.FilterName
                FilterPattern = $fb.FilterPattern
            }
            if($fb.Save)      { $params['Save']      = $true }
            if($fb.Extension) { $params['Extension'] = $fb.Extension }
            Add-AvaloniaDocumentationFileBrowse @params
        }
    }

    # Select/deselect-all moved into the DataGrid column header.
    $headerCb = Initialize-AvaloniaGridSelectAllHeader -Grid $script:dgBulkDocObjects -BindingProperty 'Selected' -InitiallyChecked $true
    $script:_bulkDocFormState.HeaderCb = $headerCb

    if ($rbGroup) {
        $rbGroup.add_IsCheckedChanged((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            if (-not $s.IsChecked) { return }
            $script:bulkDocMode = "Group"
            Update-BulkDocumentationObjectList
            $st = $script:_bulkDocFormState
            if ($st -and $st.HeaderCb) { $st.HeaderCb.IsChecked = $true }
        }))
    }
    if ($rbType) {
        $rbType.add_IsCheckedChanged((ConvertTo-AvaloniaEventScriptBlock {
            param($s, $e)
            if (-not $s.IsChecked) { return }
            $script:bulkDocMode = "Type"
            Update-BulkDocumentationObjectList
            $st = $script:_bulkDocFormState
            if ($st -and $st.HeaderCb) { $st.HeaderCb.IsChecked = $true }
        }))
    }

    if ($btnDoc) {
        $btnDoc.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkDocFormState
            if (-not $st) { return }
            $selection = if($script:bulkDocMode -eq "Object") { Get-BulkDocumentationSelectedObjects } else { Get-BulkDocumentationSelectedIds }
            $unit = if ($script:bulkDocMode -eq "Object") { "policy" } elseif ($script:bulkDocMode -eq "Type") { "policy type" } else { "object group" }

            if ($selection.Count -eq 0) {
                $ui.ShowMessageBox("Select at least one $unit to document.", "Bulk Documentation", "OK", "Warning") | Out-Null
                return
            }

            # Read the selected output format
            $outputValue = $null
            if ($st.CbType -and $st.CbType.SelectedItem) { $outputValue = [string]$st.CbType.SelectedItem.Value }
            if (-not $outputValue) {
                $ui.ShowMessageBox("Select an output format.", "Bulk Documentation", "OK", "Warning") | Out-Null
                return
            }

            $folder = if ($st.TxtFolder) { [string]$st.TxtFolder.Text } else { $null }
            if ($folder) { $folder = $folder.Trim().Trim('"').Trim("'") }
            $options = Get-AvaloniaDocumentationOptions $st.HostType $st.Form
            if (-not $folder) {
                # Registry-driven fallback: ask the provider where it stores
                # its output path and whether that path is a folder. Same
                # pattern as the WPF twin — adding a provider needs no change
                # here, just PrimaryPathOption/PathIsFolder on registration.
                $provider = [DocumentationRegistry]::FindOutput($outputValue)
                $selectedPath = if($provider -and $provider.PSObject.Properties['PrimaryPathOption'] -and $provider.PrimaryPathOption) {
                    $providerOpts = $options.Outputs[$outputValue]
                    if($providerOpts) { $providerOpts[$provider.PrimaryPathOption] }
                }
                if($selectedPath) {
                    $folder = if($provider -and $provider.PSObject.Properties['PathIsFolder'] -and $provider.PathIsFolder) { $selectedPath }
                              else { Split-Path -Parent $selectedPath }
                }
                if(-not $folder) { $folder = [Environment]::GetFolderPath('MyDocuments') }
            }
            try { $folder = [IO.Path]::GetFullPath($folder) }
            catch {
                $ui.ShowMessageBox("Invalid output folder path: $($_.Exception.Message)", "Bulk Documentation", "OK", "Error") | Out-Null
                return
            }

            $language = [string]$options.Language
            if (-not $language) { $language = 'en' }
            $defaultName = "%Organization%-%Date%"
            if(-not $options.Outputs.csv.CSVDocumentationPath) { $options.Outputs.csv.CSVDocumentationPath = $folder }
            if(-not $options.Outputs.json.JSONDocumentName) { $options.Outputs.json.JSONDocumentName = Join-Path $folder "$defaultName.json" }
            if(-not $options.Outputs.html.HTMLDocumentName) { $options.Outputs.html.HTMLDocumentName = Join-Path $folder "$defaultName.html" }
            # Confluence storage format is XHTML, so .html keeps it openable in a browser.
            if(-not $options.Outputs.atlassian.AtlassianDocumentName) { $options.Outputs.atlassian.AtlassianDocumentName = Join-Path $folder "$defaultName.html" }
            if(-not $options.Outputs.md.MDDocumentName) { $options.Outputs.md.MDDocumentName = Join-Path $folder "$defaultName.md" }
            if(-not $options.Outputs.word.WordDocumentName) { $options.Outputs.word.WordDocumentName = Join-Path $folder "$defaultName.docx" }

            # Persist UI defaults for next time. The public call below receives
            # the complete options object and never depends on these settings.
            Save-SettingStoreValue "Documentation" "OutputType"            $outputValue
            Save-SettingStoreValue "Documentation" "OutputFolder"          $folder
            Save-DocumentationOptionsDefaults $options

            $startParams = @{
                OutputFormat = $outputValue
                Language     = $language
                Options      = $options
            }
            if ($script:bulkDocMode -ne "Object" -and $options.SourceFolder) { $startParams.SourceFolder = $options.SourceFolder }
            if ($script:bulkDocMode -eq "Type") {
                $startParams.PolicyType = $selection
            } elseif ($script:bulkDocMode -eq "Group") {
                $startParams.PolicyGroup = $selection
            }

            Write-Log "Documentation starting. Mode=$script:bulkDocMode Format=$outputValue Folder=$folder Selection=$($selection -join ',')"
            try {
                if($script:bulkDocMode -eq "Object") {
                    [void]($selection | Start-GraphBulkDocumentation @startParams)
                } else {
                    [void](Start-GraphBulkDocumentation @startParams)
                }
                Write-Status ""
                $ui.ShowMessageBox(("Documentation finished. Output folder:`n{0}" -f $folder), "Bulk Documentation", "OK", "Information") | Out-Null
            } catch {
                Write-LogError "Bulk documentation failed" $_.Exception
                Write-Status ""
                $ui.ShowMessageBox("Bulk documentation failed: $($_.Exception.Message)", "Bulk Documentation", "OK", "Error") | Out-Null
            }
        }))
    }

    if ($btnPreview) {
        $btnPreview.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $selection = if($script:bulkDocMode -eq "Object") { Get-BulkDocumentationSelectedObjects } else { Get-BulkDocumentationSelectedIds }
            if ($selection.Count -eq 0) { return }
            $st = $script:_bulkDocFormState
            if (-not $st) { return }
            $options = Get-AvaloniaDocumentationOptions $st.HostType $st.Form
            $params = if ($script:bulkDocMode -eq 'Type') { @{PolicyType=$selection} } elseif($script:bulkDocMode -eq 'Group') { @{PolicyGroup=$selection} } else { @{} }
            if($script:bulkDocMode -eq 'Object') {
                $policies = $selection
            } elseif ($options.SourceFolder) {
                $options.SourceTenantUnavailable = $true
                Initialize-DocumentationSourceTenantContext -SourceFolder $options.SourceFolder
                $policies = Get-DocumentationPoliciesFromSourceFolder -SourceFolder $options.SourceFolder @params
            } else {
                $policies = Get-GraphPolicies @params
            }
            $script:bulkDocPreview = @($policies | Select-Object -First 5 | Get-GraphDocumentation -Language $options.Language -Options $options)
            if ($st.TxtRaw) { $st.TxtRaw.Text = ($script:bulkDocPreview | ConvertTo-Json -Depth 10) }
        }))
    }
    if ($btnCopyBasic) {
        $btnCopyBasic.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            @($script:bulkDocPreview | ForEach-Object BasicInfo) | Select-Object Name,Value | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
        }))
    }
    if ($btnCopySettings) {
        $btnCopySettings.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            @($script:bulkDocPreview | ForEach-Object FilteredSettings) | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
        }))
    }
    if ($btnCopyRaw) {
        $btnCopyRaw.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_bulkDocFormState
            if ($st -and $st.TxtRaw) { [string]$st.TxtRaw.Text | Set-Clipboard }
        }))
    }

    if ($btnClose) {
        $btnClose.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $script:dgBulkDocObjects = $null
            $script:bulkDocRows = $null
            $script:bulkDocPolicyObjects = $null
            $script:_bulkDocFormState = $null
            Close-TopModalObject
        }))
    }

    $form.add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            $script:dgBulkDocObjects = $null
            $script:bulkDocRows = $null
            $script:bulkDocPolicyObjects = $null
            $script:_bulkDocFormState = $null
            Close-TopModalObject
            $e.Handled = $true
        }
    }))

    $title = if($script:bulkDocMode -eq 'Object') { 'Document' } else { 'Bulk Documentation' }
    $ui.ShowModalForm($title, $form, $true)
}

# Compatibility wrapper for the main-view Document button. The shared form
# switches to Object mode and displays the selected policies.
function Show-IntuneManagerDocumentForm
{
    param([object[]]$PolicyObject)

    Show-GraphBulkDocumentationForm -PolicyObject $PolicyObject
}
