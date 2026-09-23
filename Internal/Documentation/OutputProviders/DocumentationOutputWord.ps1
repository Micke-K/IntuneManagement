# Word output provider.
#
# Consumes the per-object documentation result (see DocumentationOutputJson.ps1
# header for the field list). Writes a .docx, .docm/.xml (strict), or .pdf via
# Microsoft.Office.Interop.Word COM automation. Word must be installed locally.
#
# https://docs.microsoft.com/en-us/office/vba/api/overview/word
#
# Differences vs old DocumentationWord.psm1:
#   - Uses the typed-object API: $PolicyObject.Name + $PolicyObject.PolicyType
#     instead of Get-GraphObjectName / Get-ObjectTypeString taking $objectType
#   - V1 NewObjectGroup/NewObjectType ($obj-arg) replaced by V2 (string-arg);
#     the old V1 versions were unreachable (registration used V2)
#   - The "Attach raw object JSON" Word feature is stubbed out — it depends on
#     Export-GraphObject which lives in old MSGraph.psm1 and hasn't been ported
#     to the new project. A "feature unavailable" log message replaces it; phase 2
#     can re-enable once the equivalent exporter is wired up.
#   - Invoke-WordProcessAllObjects ScopeTags consolidated table is stubbed too,
#     same reason as the HTML provider (Get-TableObjects deferred to phase 2).
#   - All other COM logic — cover page, ToC, building blocks, style hashtable,
#     option snapshot/restore, save & close — preserved verbatim.

# Load the Word primary interop assembly. Deliberately NOT called at module
# import: Add-Type -AssemblyName fails on PS7 (it resolves against the current
# directory, not the GAC), so this always fell through to a recursive scan of
# %windir%\assembly\GAC_MSIL - ~77ms on every single Import-Module, for a
# feature most sessions never use. It also put an assembly in the AppDomain
# whose GetExportedTypes() throws, which is one of the two things that used to
# take down Avalonia's XAML loader (see Host.SanitizeXamlTypeSystem).
#
# Idempotent: returns $true as soon as the interop types are resolvable.
#
# The readiness probe is WdSaveFormat, not the Application coclass. Everything
# this provider needs from the interop assembly is enums - the Word instance
# itself comes from late-bound `New-Object -ComObject Word.Application` - and on
# PS7 the enums resolve while the coclass does not. Probing Application would
# therefore report "not loaded" forever on PS7, which is also why the previous
# code's short-circuit never fired there and re-scanned the GAC on every import.
function Initialize-WordInteropAssembly {
    if ("Microsoft.Office.Interop.Word.WdSaveFormat" -as [Type]) { return $true }

    try {
        Add-Type -AssemblyName Microsoft.Office.Interop.Word -ErrorAction Stop
        if ("Microsoft.Office.Interop.Word.WdSaveFormat" -as [Type]) { return $true }
    }
    catch { }

    try {
        $wordFile = Get-ChildItem -Path "$($env:windir)\assembly\GAC_MSIL" -Filter "Microsoft.Office.Interop.Word.dll" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($wordFile -and $wordFile.Exists) {
            Add-Type -Path $wordFile.FullName
            if ("Microsoft.Office.Interop.Word.WdSaveFormat" -as [Type]) { return $true }
        }
    }
    catch {
        Write-LogError "Failed to add Word Interop type. Cannot create Word documents. Verify that Word is installed properly." $_.Exception
        return $false
    }

    Write-LogError "Word Interop type not found. Cannot create Word documents. Verify that Word is installed properly." $null
    return $false
}

function Invoke-InitializeWordOutput {
    # Word output relies on Microsoft.Office.Interop.Word COM automation, which only
    # exists on Windows with Word installed. Skip the provider entirely elsewhere.
    if (-not $script:IsWindowsOS) {
        Write-Log "Word documentation output is not available on this platform (requires Windows + Word). Skipping."
        return
    }

    # Registration stays gated on Word being installed, as before - but probed by
    # reading the COM registration straight out of the registry, which loads
    # nothing into the AppDomain. [Type]::GetTypeFromProgID looks like the natural
    # call here and is correct on PS7, but on PS5.1 the .NET Framework resolves a
    # ProgID to its primary interop assembly and loads it - which would reintroduce
    # exactly the eager load this is meant to remove, on the one shell where the
    # old code's short-circuit actually worked.
    #
    # The interop itself is loaded lazily by Invoke-WordActivate, i.e. only when a
    # documentation run actually selects Word output.
    $wordRegistered = (Test-Path 'HKLM:\SOFTWARE\Classes\Word.Application') -or
                      (Test-Path 'HKCU:\SOFTWARE\Classes\Word.Application')
    if (-not $wordRegistered) {
        Write-Log "Word is not registered on this machine. Word documentation output will not be available." 2
        return
    }

    Add-DocumentationOutputProvider ([PSCustomObject]@{
        Name              = "Word"
        Value             = "word"
        # Path metadata (see DocumentationOutputHTML.ps1 header comment).
        # UI browse-button wiring lives in
        # UI/<backend>/ClassExtensions/DocumentationOutputWordUIExtension.ps1.
        PrimaryPathOption = "WordDocumentName"
        PathIsFolder      = $false
        Activate          = { Invoke-WordActivate @args }
        PreProcess        = { Invoke-WordPreProcessItems @args }
        NewObjectGroup    = { Invoke-WordNewObjectGroup @args }
        NewObjectType     = { Invoke-WordNewObjectType @args }
        Process           = { Invoke-WordProcessItem @args }
        PostProcess       = { Invoke-WordPostProcessItems @args }
        ProcessAllObjects = { Invoke-WordProcessAllObjects @args }
    })
}

function Invoke-WordActivate {
    # Lazy load point for the interop assembly. Activate is the first lifecycle
    # hook the engine runs for a selected provider, so this happens only when a
    # documentation run actually asks for Word output. Throwing here is
    # deliberate: the engine wraps Activate in its $recordFailure handler, so the
    # run reports "Activate failed for Word: ..." instead of failing later and
    # less clearly when Process touches an interop enum.
    if (-not (Initialize-WordInteropAssembly)) {
        throw "Word Interop assembly could not be loaded. Cannot create Word documents. Verify that Word is installed properly."
    }
}

function Invoke-WordPreProcessItems {
    # Validate Limit-mode bounds
    $script:limitMaxValue        = 100
    $script:truncateValueLength  = $script:limitMaxValue

    if ((Get-DocumentationOutputOption word "WordDocumentationLevel" "full") -eq "limited") {
        $maxText      = Get-DocumentationOutputOption word "WordDocumentationLimitMaxLength"      ""
        $truncateText = Get-DocumentationOutputOption word "WordDocumentationLimitTruncateLength" ""

        if ($maxText) {
            try { $script:limitMaxValue = [int]::Parse($maxText) }
            catch { Write-LogError "Failed to parse '$maxText' to int. Max value length will be set to 100." $_.Exception }
        }
        if ($truncateText) {
            try { $script:truncateValueLength = [int]::Parse($truncateText) }
            catch { Write-LogError "Failed to parse '$truncateText' to int. Truncate length will be set to $script:limitMaxValue." $_.Exception }
        }

        if ($script:limitMaxValue -lt 20) {
            Write-Log "Max value length must be 20 or more. Changed to 0" 2
            $script:limitMaxValue = 0
        }
        if ($script:truncateValueLength -lt 0) {
            Write-Log "Truncate length must be 0 or more. Changed to 0" 2
            $script:truncateValueLength = 0
        }
        elseif ($script:truncateValueLength -gt $script:limitMaxValue) {
            Write-Log "Truncate length cannot be larger than Max value length. Changed to: $script:limitMaxValue" 2
            $script:truncateValueLength = $script:limitMaxValue
        }
    }

    # Create Word COM app
    try {
        $script:wordApp = New-Object -ComObject Word.Application
    }
    catch {
        Write-LogError "Failed to create Word App object. Word documentation aborted..." $_.Exception
        return $false
    }

    # Performance: suppress UI redraw and background processing while filling the document.
    # Application-level Options persist to the user profile, so we snapshot and restore them in PostProcess.
    $script:wordApp.ScreenUpdating = $false
    $script:wordApp.DisplayAlerts  = 0  # wdAlertsNone

    $script:wordOptionsBackup = $null
    try {
        $script:wordOptionsBackup = @{
            Pagination             = $script:wordApp.Options.Pagination
            CheckGrammarAsYouType  = $script:wordApp.Options.CheckGrammarAsYouType
            CheckSpellingAsYouType = $script:wordApp.Options.CheckSpellingAsYouType
            BackgroundSave         = $script:wordApp.Options.BackgroundSave
        }
        $script:wordApp.Options.Pagination             = $false
        $script:wordApp.Options.CheckGrammarAsYouType  = $false
        $script:wordApp.Options.CheckSpellingAsYouType = $false
        $script:wordApp.Options.BackgroundSave         = $false
    }
    catch { }

    $template = Get-DocumentationOutputOption word "WordDocumentTemplate" ""
    if ($template) {
        try {
            $script:doc = $script:wordApp.Documents.Add($template)
        }
        catch {
            Write-LogError "Failed to create document based on template: $template" $_.Exception
        }
    }
    else {
        $script:doc = $script:wordApp.Documents.Add()
    }

    # Get BuiltIn properties
    $script:builtInProps = [System.Collections.Generic.List[object]]::new()
    $script:doc.BuiltInDocumentProperties | ForEach-Object {
        $name = [System.__ComObject].InvokeMember("name",  [System.Reflection.BindingFlags]::GetProperty, $null, $_, $null)
        $value = $null
        try { $value = [System.__ComObject].InvokeMember("value", [System.Reflection.BindingFlags]::GetProperty, $null, $_, $null) } catch {}
        if ($name) {
            $script:builtInProps.Add([PSCustomObject]@{ Name = $name; Value = $value })
        }
    }

    # Get Custom properties
    $script:customProps = [System.Collections.Generic.List[object]]::new()
    $script:doc.CustomDocumentProperties | ForEach-Object {
        $name = [System.__ComObject].InvokeMember("name",  [System.Reflection.BindingFlags]::GetProperty, $null, $_, $null)
        $value = $null
        try { $value = [System.__ComObject].InvokeMember("value", [System.Reflection.BindingFlags]::GetProperty, $null, $_, $null) } catch {}
        if ($name) {
            $script:customProps.Add([PSCustomObject]@{ Name = $name; Value = $value })
        }
    }

    # Style cache: O(1) lookup by NameLocal (replaces per-call linear scan in Get-DocStyle / Set-DocObjectStyle)
    $script:wordStyles = @{}
    $script:doc.Styles | ForEach-Object {
        if ($_.NameLocal -and -not $script:wordStyles.ContainsKey($_.NameLocal)) {
            $script:wordStyles[$_.NameLocal] = [PSCustomObject]@{
                Name = $_.NameLocal; Type = $_.Type; Style = $_
            }
        }
    }

    # Built-in style cache: same O(1) treatment for the ~376 enum names
    $script:builtinStyles = @{}
    foreach ($builtinName in [Enum]::GetNames([Microsoft.Office.Interop.Word.wdBuiltinStyle])) {
        $script:builtinStyles[$builtinName] = $true
    }

    if (-not $template) {
        $script:doc.Application.Templates.LoadBuildingBlocks()
        $bb = $script:doc.Application.Templates | Where-Object { $_.Name -eq 'Built-In Building Blocks.dotx' }
        if ($bb) {
            $coverPageName = Get-DocumentationOutputOption word "WordCoverPage" "Ion (Dark)"
            if (-not $coverPageName) { $coverPageName = 'Ion (Dark)' }

            try {
                $blocks = @()
                for ($i = 1; $i -le $bb.BuildingBlockEntries.Count; $i++) {
                    $blocks += $bb.BuildingBlockEntries.Item($i)
                }
                $coverPages = ($blocks | Where-Object { $_.Type.Index -eq 2 } | Select-Object Name | Sort-Object -Property Name).Name

                if (($coverPages | Measure-Object).Count -gt 0) {
                    if ($coverPageName -notin $coverPages) {
                        Write-Log "$coverPageName not found in available Cover Page list. Using: $($coverPages[0])"
                        Write-Log "Available Cover Pages: $($coverPages -join ',')"
                        $coverPageName = $coverPages[0]
                    }
                    else {
                        Write-Log "Add Cover Page: $coverPageName"
                    }
                }

                $coverPage = $bb.BuildingBlockEntries.Item($coverPageName)
                $coverPage.Insert($script:wordApp.Selection.Range, $true) | Out-Null
                $script:wordApp.Selection.InsertNewPage()
            }
            catch { Write-LogError "Failed to create Cover Page" $_.Exception }

            try {
                $coverPageProps = $script:doc.CustomXMLParts | Where-Object { $_.NamespaceURI -match "coverPageProps$" }
                if ($coverPageProps) {
                    Write-Log "Available Cover Page properties for $($coverPageName): $((([xml]$coverPageProps.DocumentElement.XML).ChildNodes[0].ChildNodes).Name -join ',')"
                }
            }
            catch { }

            try {
                $script:doc.TablesOfContents.Add($script:wordApp.Selection.Range) | Out-Null
                $script:wordApp.Selection.InsertNewPage()
            }
            catch { Write-LogError "Failed to create Table of Contents" $_.Exception }
        }
    }
    else {
        Invoke-DocGoToEnd
        $script:wordApp.Selection.InsertNewPage()
    }
}

function Invoke-WordPostProcessItems {
    $userName = $null
    $me = Get-CurrentUser
    if ($me) {
        if ($me.givenName -and $me.surname) {
            $userName = "$($me.givenName) $($me.surname)"
        }
        else {
            $userName = $me.displayName
        }
    }

    $titleProp   = Get-DocumentationOutputOption word "WordTitleProperty"   "Intune documentation"
    $subjectProp = Get-DocumentationOutputOption word "WordSubjectProperty" "Intune documentation"
    if (-not $titleProp)   { $titleProp   = "Intune documentation" }
    if (-not $subjectProp) { $subjectProp = "Intune documentation" }

    Set-WordDocBuiltInProperty "wdPropertyTitle"    $titleProp
    Set-WordDocBuiltInProperty "wdPropertySubject"  $subjectProp
    # Author + Company are the "who generated this" document info. Word writes them
    # as built-in file metadata (not a visible top-of-document block like the other
    # providers), but they are the same info the generic SkipDocumentInfo flag hides.
    if (-not ((Get-DocumentationOption "SkipDocumentInfo" $false) -eq $true)) {
        Set-WordDocBuiltInProperty "wdPropertyAuthor" $userName
        $orgName = Get-CurrentOrganizationName
        if ($orgName) {
            Set-WordDocBuiltInProperty "wdPropertyCompany" $orgName
        }
    }
    Set-WordDocBuiltInProperty "wdPropertyKeywords" "Intune,Endpoint Manager,MEM"

    try {
        $controls = Get-DocumentationOutputOption word "WordContentControls" ""
        foreach ($ccObj in $controls.Split(';')) {
            $ccName, $ccVal = $ccObj.Split('=')
            Set-WordContentControlText $ccName $ccVal
        }
    }
    catch { }

    foreach ($field in @("Fields","TablesOfContents","TablesOfFigures","TablesOfAuthorities")) {
        try { $script:doc.$field | ForEach-Object { $_.Update() | Out-Null } }
        catch { Write-LogError "Failed to update document $field" $_.Exception }
    }

    # Restore Application Options before saving so user settings aren't permanently changed.
    try {
        if ($script:wordOptionsBackup) {
            $script:wordApp.Options.Pagination             = $script:wordOptionsBackup.Pagination
            $script:wordApp.Options.CheckGrammarAsYouType  = $script:wordOptionsBackup.CheckGrammarAsYouType
            $script:wordApp.Options.CheckSpellingAsYouType = $script:wordOptionsBackup.CheckSpellingAsYouType
            $script:wordApp.Options.BackgroundSave         = $script:wordOptionsBackup.BackgroundSave
        }
        $script:wordApp.ScreenUpdating = $true
        $script:doc.Repaginate()
    }
    catch { }

    $formatStr = Get-DocumentationOutputOption word "WordDocumentFormat" "wdFormatDocumentDefault"
    if     ($formatStr -eq "pdf")  { $formatStr = "wdFormatPDF" }
    elseif ($formatStr -eq "docx") { $formatStr = "wdFormatDocumentDefault" }
    Write-Log "Using document format: $formatStr"
    $format = $null
    try {
        $format = [Microsoft.Office.Interop.Word.WdSaveFormat]$formatStr
    }
    catch {
        Write-LogError "Document format validation failed; defaulting to wdFormatDocumentDefault" $_.Exception
        $format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault
    }

    $fileName = Get-DocumentationOutputOption word "WordDocumentName" ""
    if (-not $fileName) { $fileName = "%MyDocuments%\%Organization%-%Date%.docx" }
    $fileName = Expand-FileName $fileName

    if ($format -eq [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF -and $fileName -notlike "*.pdf") {
        $fileName = [IO.Path]::ChangeExtension($fileName, ".pdf")
    }

    try {
        $script:doc.SaveAs2([ref]$fileName, [ref]$format)
        Write-Log "Document $fileName saved successfully"
    }
    catch {
        Write-LogError "Failed to save file $fileName" $_.Exception
    }

    try {
        $openDocSetting = Get-DocumentationOutputOption word "WordOpenDocument" "true"
        $openDoc = ($openDocSetting -ne "false") -and ($openDocSetting -ne $false)
        # Only pop Word visible in interactive UI mode; headless / silent / bulk runs
        # close it (the .docx is already saved). ($global:hideUI was a dead old-project
        # global, never assigned -> Word always opened, even during automation.)
        $hideUI  = (Get-CacheObject "ShowUI") -ne $true
        if ($openDoc -and -not $hideUI) {
            $script:wordApp.Visible = $true
            $script:wordApp.WindowState = [Microsoft.Office.Interop.Word.WdWindowState]::wdWindowStateMaximize
            $script:wordApp.Activate()
        }
        else {
            $script:doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
            $script:wordApp.Quit()
        }
    }
    catch {
        Write-LogError "Failed to close the Word application" $_.Exception
    }
    finally {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:doc) } catch { }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:wordApp) } catch { }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Set-WordContentControlText {
    param([string]$ControlName, $Value)

    if (-not $ControlName) { return }

    try {
        $ctrl = $script:doc.SelectContentControlsByTitle($ControlName)
        if ($ctrl) {
            Write-LogDebug "Update ContentControl $ControlName (Type: $($ctrl[1].Type))"
            if ($ctrl[1].Type -eq 6) {
                if ($ctrl[1].DateDisplayFormat) {
                    $ctrl[1].Range.Text = (Get-Date).ToString($ctrl[1].DateDisplayFormat)
                }
                else {
                    $ctrl[1].Range.Text = (Get-Date).ToShortDateString()
                }
            }
            else {
                if (-not $Value) { return }
                $ctrl[1].Range.Text = $Value
            }
        }
    }
    catch {
        Write-LogError "Failed to set ContentControl $ControlName" $_.Exception
    }
}

function Invoke-WordNewObjectGroup {
    param($groupId)

    $header1 = Get-DocumentationOutputOption word "WordHeader1Style" "Heading 1"
    if (-not $header1) { $header1 = "Heading 1" }
    Add-DocText (Get-DocObjectTypeString $groupId) $header1
}

function Invoke-WordNewObjectType {
    param($objectTypeName)

    $script:objectHeaderLevel = 2
    Add-DocText $objectTypeName (Get-ObjectLevelHeader)
    $script:objectHeaderLevel = 3
}

function Get-ObjectLevelHeader {
    if ($script:objectHeaderLevel -eq 3) {
        $h3 = Get-DocumentationOutputOption word "WordHeader3Style" ""
        if ($h3) { return $h3 }
    }
    $h2 = Get-DocumentationOutputOption word "WordHeader2Style" "Heading 2"
    if (-not $h2) { $h2 = "Heading 2" }
    return $h2
}

function Invoke-WordProcessItem {
    param($PolicyObject, $documentedObj)

    if (-not $documentedObj -or -not $PolicyObject) { return }

    # A documented object may ask to be titled by something other than its display
    # name (see Get-DocumentationDisplayName). Headings and captions follow it; the
    # file name below deliberately does not.
    $objName   = Get-DocumentationDisplayName $PolicyObject $documentedObj
    $script:docDisplayName = $objName
    $typeTitle = $PolicyObject.PolicyType.Title

    Add-DocText $objName (Get-ObjectLevelHeader)
    $script:doc.Application.Selection.TypeParagraph()

    $propMode      = Get-DocumentationOutputOption word "WordExportProperties"        "simple"
    $customProps   = Get-DocumentationOutputOption word "WordCustomDisplayProperties" ""
    $docLevel      = Get-DocumentationOutputOption word "WordDocumentationLevel"      "full"
    $addCategories = (Get-DocumentationOutputOption word "WordAddCategories"    $true)  -eq $true
    $addSubCats    = (Get-DocumentationOutputOption word "WordAddSubCategories" $true)  -eq $true
    $attachJson    = (Get-DocumentationOutputOption word "WordAttachJsonFile"   $false) -eq $true

    try {
        foreach ($tableType in @("BasicInfo","FilteredSettings")) {
            if ($tableType -eq "BasicInfo") {
                $properties = @("Name","Value")
            }
            elseif ($propMode -eq 'extended' -and $documentedObj.DisplayProperties) {
                $properties = @("Name","Value","Description")
            }
            elseif ($propMode -eq 'custom' -and $customProps) {
                $properties = @()
                foreach ($prop in $customProps.Split(",")) {
                    $propInfo = $prop.Split('=')
                    if (($propInfo | Measure-Object).Count -gt 1) {
                        $properties += $propInfo[0]
                        Set-DocColumnHeaderLanguageId $propInfo[0] $propInfo[1]
                    }
                    else {
                        $properties += $prop
                    }
                }
            }
            else {
                if ($documentedObj.DefaultDocumentationProperties) {
                    $properties = $documentedObj.DefaultDocumentationProperties
                }
                else {
                    $properties = @("Name","Value")
                }
            }

            if ($docLevel -eq "basic" -and $tableType -ne "BasicInfo") { continue }

            # Custom tables with a negative Order belong ABOVE the settings
            # table: the portal shows a MAM app config's "Settings catalog"
            # blade above its "Settings" blade.
            if ($tableType -eq "FilteredSettings") {
                foreach ($customTable in ($documentedObj.CustomTables | Where-Object { $_.Order -lt 0 } | Sort-Object -Property Order)) {
                    Add-DocTableItems $PolicyObject $typeTitle $customTable.Values $customTable.Columns -LngId $customTable.LanguageId -AddCategories -AddSubcategories
                }
            }

            if (($documentedObj.$tableType | Measure-Object).Count -gt 0) {
                Add-DocTableItems $PolicyObject $typeTitle $documentedObj.$tableType $properties -AddCategories:$addCategories -AddSubcategories:$addSubCats -ForceFullValue:($tableType -eq "BasicInfo")
            }
        }

        if ($docLevel -ne "basic") {
            if (($documentedObj.ComplianceActions | Measure-Object).Count -gt 0) {
                Add-DocTableItems $PolicyObject $typeTitle $documentedObj.ComplianceActions @("Action","Schedule","MessageTemplate","EmailCC") -LngId "Category.complianceActionsLabel"
            }

            if (($documentedObj.ApplicabilityRules | Measure-Object).Count -gt 0) {
                Add-DocTableItems $PolicyObject $typeTitle $documentedObj.ApplicabilityRules @("Rule","Property","Value") -LngId "SettingDetails.applicabilityRules"
            }

            Add-DocObjectScripts $documentedObj

            # Negative Order already rendered above the settings table.
            foreach ($customTable in ($documentedObj.CustomTables | Where-Object { $_.Order -ge 0 } | Sort-Object -Property Order)) {
                Add-DocTableItems $PolicyObject $typeTitle $customTable.Values $customTable.Columns -LngId $customTable.LanguageId -AddCategories -AddSubcategories
            }
        }

        if (($documentedObj.Assignments | Measure-Object).Count -gt 0) {
            $settingProps = $null
            if ($documentedObj.Assignments[0].RawIntent) {
                $properties = @("GroupMode","Group","Filter","FilterMode")
                $settingProps = @("Filter","FilterMode")
                $settingsObj  = $documentedObj.Assignments | Where-Object { $_.Settings -ne $null } | Select-Object -First 1
                if ($settingsObj) {
                    foreach ($objProp in $settingsObj.Settings.Keys) {
                        if ($objProp -in $properties)               { continue }
                        if ($objProp -in @("Category","RawIntent")) { continue }
                        $settingProps += "Settings.$objProp"
                    }
                }
            }
            else {
                $hasFilter = $false
                foreach ($a in $documentedObj.Assignments) {
                    if ($a.PSObject.Properties.Name -contains "FilterMode") { $hasFilter = $true; break }
                }
                $properties = @("Group")
                if ($hasFilter) { $properties += @("Filter","FilterMode") }
            }

            Add-DocTableItems $PolicyObject $typeTitle $documentedObj.Assignments $properties -LngId "TableHeaders.assignments" -AddCategories

            if ($null -ne $settingProps) {
                # Adds additional values to the assignments table for Apps assignments
                Set-DocTableSettingsItems $documentedObj.Assignments $settingProps 3
            }
        }

        if ($attachJson) {
            # Embed the full raw object JSON as an OLE object (old feature gated on
            # chkWordAttachJsonFile). The full object is already in hand, so write it
            # to a temp file directly rather than depending on an external exporter.
            try {
                # The policy's real name, not the heading: the heading may be a
                # DocumentName override, and this is the attached object's label.
                $safeName = ([string]$PolicyObject.Name)
                foreach ($ch in [IO.Path]::GetInvalidFileNameChars()) { $safeName = $safeName.Replace($ch, '_') }
                if ([string]::IsNullOrEmpty($safeName)) { $safeName = 'object' }
                $fi = [IO.FileInfo](Join-Path ([IO.Path]::GetTempPath()) "$safeName.json")
                ($PolicyObject.JsonObject | ConvertTo-Json -Depth 50) | Out-File -LiteralPath $fi.FullName -Encoding UTF8
                $fi.Refresh()
                if ($fi.Exists) {
                    $script:doc.Application.Selection.InlineShapes.AddOLEObject("", $fi.FullName, $false, $true, "$($env:WinDir)\System32\Notepad.exe", 0, $fi.Name)
                    $script:doc.Application.Selection.TypeParagraph()
                    try { $fi.Delete() } catch { }
                }
            }
            catch {
                Write-LogError "Failed to attach JSON for $objName" $_.Exception
            }
        }
    }
    catch {
        Write-LogError "Failed to process object $objName" $_.Exception
    }
}

function Set-DocTableSettingsItems {
    param($items, $properties, [int]$firstColumn)

    $secondColumn = $firstColumn + 1

    $script:docTable.Cell(1, $firstColumn).Range.Text  = (Invoke-DocTranslateColumnHeader "Settings")
    $script:docTable.Cell(1, $secondColumn).Range.Text = ""

    $row = 2
    foreach ($itemObj in $items) {
        while ($script:docTable.Cell($row, 1).Next.RowIndex -gt $row) {
            # Category / Sub-category row — skip
            $row++
        }
        $script:docTable.Cell($row, $firstColumn).Range.Text  = ""
        $script:docTable.Cell($row, $secondColumn).Range.Text = ""
        $script:docTable.Cell($row, $firstColumn).Split($properties.Count, 1)
        $script:docTable.Cell($row, $secondColumn).Split($properties.Count, 1)

        $cellRow = $row
        foreach ($settingProp in $properties) {
            if ([string]::IsNullOrEmpty($settingProp)) { continue }

            $script:docTable.Cell($cellRow, $firstColumn).Range.Text = (Invoke-DocTranslateColumnHeader ($settingProp.Split('.')[-1]))

            $propArr = $settingProp.Split('.')
            $tmpObj  = $itemObj
            $propName = $propArr[-1]
            for ($x = 0; $x -lt ($propArr.Count - 1); $x++) {
                $tmpObj = $tmpObj."$($propArr[$x])"
            }

            $script:docTable.Cell($cellRow, $secondColumn).Range.Text = "$($tmpObj.$propName)"
            $cellRow++
        }
        $row = $row + $properties.Count
    }
}

function Invoke-WordProcessAllObjects {
    param($allObjectTypeObjects)
    # ScopeTags consolidated table is deferred — depends on Get-TableObjects-style
    # cross-object accumulation that hasn't been ported yet. Phase 2 re-enables.
}

function Add-DocTableItems {
    param(
        $PolicyObject,
        [string]$TypeTitle,
        $Items,
        [string[]]$Properties,
        [string]$LngId,
        [switch]$AddCategories,
        [switch]$AddSubcategories,
        $CaptionOverride,
        [switch]$ForceFullValue
    )

    if (($Items | Measure-Object).Count -eq 0) { return }

    $tblHeaderStyle      = Get-DocumentationOutputOption word "WordTableHeaderStyle"       ""
    $tblCategoryStyle    = Get-DocumentationOutputOption word "WordCategoryHeaderStyle"    ""
    $tblSubCategoryStyle = Get-DocumentationOutputOption word "WordSubCategoryHeaderStyle" ""
    $tblTextStyle        = Get-DocumentationOutputOption word "WordTableTextStyle"         ""
    $tblStyle            = Get-DocumentationOutputOption word "WordTableStyle"             "Grid table 4 - Accent 3"
    $captionPos          = Get-DocumentationOutputOption word "WordTableCaptionPosition"   "below"
    $docLevel            = Get-DocumentationOutputOption word "WordDocumentationLevel"     "full"
    $limitAttach         = (Get-DocumentationOutputOption word "WordDocumentationLimitAttach" $false) -eq $true

    $range = $script:doc.Application.Selection.Range

    # Pre-pass: count category / sub-category rows so the table can be allocated at its final size.
    # Boundary logic MUST stay in sync with the main fill loop below.
    $extraRows = 0
    $preCat    = ""
    $preSubCat = ""
    foreach ($itemObj in $Items) {
        if ($itemObj.Category -and $preCat -ne $itemObj.Category -and $AddCategories) {
            $extraRows++
            $preCat    = $itemObj.Category
            $preSubCat = ""
        }
        if ($itemObj.SubCategory -and $preSubCat -ne $itemObj.SubCategory -and $AddSubcategories) {
            $extraRows++
            $preSubCat = $itemObj.SubCategory
        }
    }

    $totalRows = @($Items).Count + 1 + $extraRows

    # Create with wdAutoFitFixed during fill — wdAutoFitWindow recalculates column widths after every cell write.
    # We re-enable wdAutoFitWindow once after the rows are populated.
    $script:docTable = $script:doc.Tables.Add($range, $totalRows, $Properties.Count, [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior, [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitFixed)
    $script:docTable.ApplyStyleHeadingRows = $true
    Set-DocObjectStyle $script:docTable $tblStyle | Out-Null

    if ($CaptionOverride) {
        $caption = $CaptionOverride
    }
    elseif ($LngId -and $PolicyObject) {
        $caption = "$((Get-LanguageString $LngId)) - $(Get-DocCaptionName $PolicyObject)"
    }
    elseif ($PolicyObject) {
        $caption = "$(Get-DocCaptionName $PolicyObject) ($TypeTitle)"
    }
    else {
        $caption = $TypeTitle
    }

    $i = 1
    foreach ($prop in $Properties) {
        if ([string]::IsNullOrEmpty($prop)) { continue }
        $script:docTable.Cell(1, $i).Range.Text = (Invoke-DocTranslateColumnHeader ($prop.Split('.')[-1]))
        $i++
    }

    if (-not (Set-DocObjectStyle $script:docTable.Rows(1).Range $tblHeaderStyle)) {
        $script:docTable.Rows(1).Range.Font.Size += 2
        $script:docTable.Rows(1).Range.Font.Bold = $true
    }

    $curCategory    = ""
    $curSubCategory = ""

    $row = 2
    foreach ($itemObj in $Items) {
        try {
            if ($itemObj.Category -and $curCategory -ne $itemObj.Category -and $AddCategories) {
                try { $script:docTable.Rows.Item($row).Cells.Merge() } catch { }
                $script:docTable.Cell($row, 1).Range.Text = $itemObj.Category

                if (-not (Set-DocObjectStyle $script:docTable.Rows($row).Range $tblCategoryStyle)) {
                    $script:docTable.Rows($row).Range.Font.Size += 2
                    $script:docTable.Rows($row).Range.Font.Italic = $true
                }

                $row++
                $curCategory    = $itemObj.Category
                $curSubCategory = ""
            }

            if ($itemObj.SubCategory -and $curSubCategory -ne $itemObj.SubCategory -and $AddSubcategories) {
                try { $script:docTable.Rows.Item($row).Cells.Merge() } catch { }
                $script:docTable.Cell($row, 1).Range.Text = $itemObj.SubCategory

                if (-not (Set-DocObjectStyle $script:docTable.Rows($row).Range $tblSubCategoryStyle)) {
                    $script:docTable.Rows($row).Range.Font.Italic = $true
                }

                $row++
                $curSubCategory = $itemObj.SubCategory
            }

            $i = 1
            foreach ($prop in $Properties) {
                try {
                    $propArr  = $prop.Split('.')
                    $tmpObj   = $itemObj
                    $propName = $propArr[-1]
                    for ($x = 0; $x -lt ($propArr.Count - 1); $x++) {
                        $tmpObj = $tmpObj."$($propArr[$x])"
                    }
                    $propValue     = "$($tmpObj.$propName)"
                    $propValueFull = $null

                    if (-not $ForceFullValue -and $docLevel -eq "limited" -and $propValue.Length -gt $script:limitMaxValue) {
                        $propValueFull = $propValue
                        if ($script:truncateValueLength -gt 0) {
                            $propValue = $propValue.Substring(0, $script:truncateValueLength) + "..."
                            if ($limitAttach) { $propValue = "`r`n" + $propValue }
                        }
                        else {
                            $propValue = $null
                        }
                    }

                    $levelExtra = ""
                    if ($i -eq 1 -and $itemObj.Level) {
                        try {
                            $level = [int]$itemObj.Level
                            if ($level -lt 0) { $level = 0 }
                            if ($level -gt 0) {
                                $levelExtra = [string]::new(" ", ($level * 2))
                            }
                        }
                        catch { }
                    }

                    if ($null -ne $propValue) {
                        $script:docTable.Cell($row, $i).Range.Text = "$levelExtra$propValue"
                    }

                    if ($propValueFull -and $limitAttach) {
                        $tmpName = "$($PolicyObject.Name)-$propName"
                        $tmpFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "$tmpName.txt")
                        $tmpFile = Remove-InvalidFileNameChars $tmpFile
                        $propValueFull | Out-File -LiteralPath $tmpFile -Force
                        $fi = [IO.FileInfo]$tmpFile
                        [void]$script:docTable.Cell($row, $i).Range.InlineShapes.AddOLEObject("", $fi.FullName, $false, $true, "$($env:WinDir)\System32\Notepad.exe", 0, "Full value")
                        try { $fi.Delete() } catch { }
                    }
                }
                catch {
                    Write-LogError "Failed to add property value for $prop" $_.Exception
                }
                $i++
            }

            Set-DocObjectStyle $script:docTable.Rows($row).Range $tblTextStyle | Out-Null
        }
        catch {
            Write-Log "Failed to process property" 2
        }

        $row++
    }

    try { $script:docTable.AutoFitBehavior([Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow) } catch { }

    # -2 = Table caption, 1 = Below / 0 = Above
    $capPos = if ($captionPos -eq "above") { 0 } else { 1 }
    $script:docTable.Application.Selection.InsertCaption(-2, ". $caption", $null, $capPos)

    Invoke-DocGoToEnd
    $script:doc.Application.Selection.TypeParagraph()
}

function Add-DocTableScript {
    param([string]$Caption, [string]$Header, [string]$ScriptText)

    if (-not $ScriptText) { return }

    $primary = Get-DocumentationOutputOption word "WordScriptTableStyle" ""
    if (-not $primary) {
        $primary = Get-DocumentationOutputOption word "WordTableStyle" "Grid table 4 - Accent 3"
    }
    $scriptStyle = Get-DocumentationOutputOption word "WordScriptStyle" ""

    $range = $script:doc.Application.Selection.Range
    $scriptTable = $script:doc.Tables.Add($range, 2, 1, [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior, [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitFixed)
    $scriptTable.ApplyStyleHeadingRows = $true
    Set-DocObjectStyle $scriptTable $primary | Out-Null

    if ($Header) {
        $scriptTable.Cell(1, 1).Range.Text = $Header
    }

    $scriptTable.Cell(2, 1).Range.Font.Bold = $false
    $scriptTable.Cell(2, 1).Range.Text      = $ScriptText
    if ($scriptStyle) {
        Set-DocObjectStyle $scriptTable.Rows(2).Range $scriptStyle | Out-Null
    }
    else {
        $tmp = $script:wordStyles["HTML Code"]
        if ($tmp) {
            $scriptTable.Cell(2, 1).Range.Font = $tmp.Style.Font
        }
        $scriptTable.Cell(2, 1).Range.Font.Bold = $false
    }
    $scriptTable.Cell(2, 1).Range.NoProofing = $true

    try { $scriptTable.AutoFitBehavior([Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow) } catch { }
    $scriptTable.Application.Selection.InsertCaption(-2, ". $Caption", $null, 1)
    $script:doc.Application.Selection.TypeParagraph()
}

function Get-DocStyle {
    param([string]$StyleName)

    $tmpStyle = $null
    if ($StyleName -and $script:wordStyles.ContainsKey($StyleName)) {
        $tmpStyle = $script:wordStyles[$StyleName].Style
    }
    if (-not $tmpStyle) { Write-Log "Style $StyleName not found" }
    $tmpStyle
}

function Add-DocText {
    param([string]$Text, [string]$Style, [switch]$SkipAddParagraph)

    Set-DocObjectStyle $script:doc.Application.Selection $Style | Out-Null
    $script:doc.Application.Selection.TypeText($Text)
    if (-not $SkipAddParagraph) {
        $script:doc.Application.Selection.TypeParagraph()
    }
}

function Invoke-DocGoToEnd {
    $script:doc.Application.Selection.GoTo([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToBookmark, $null, $null, '\EndOfDoc') | Out-Null
}

function Set-WordDocBuiltInProperty {
    param([string]$PropertyName, $Value)

    try {
        $script:doc.BuiltInDocumentProperties([Microsoft.Office.Interop.Word.WdBuiltInProperty]$PropertyName) = $Value
    }
    catch {
        Write-LogError "Failed to set built in property $PropertyName to $Value" $_.Exception
    }
}

function Set-DocObjectStyle {
    param($DocObj, [string]$ObjStyle)

    $styleSet = $false
    if ($DocObj -and $ObjStyle) {
        try {
            if ($script:builtinStyles.ContainsKey($ObjStyle)) {
                $DocObj.style = [Microsoft.Office.Interop.Word.wdBuiltinStyle]$ObjStyle
            }
            else {
                $DocObj.style = $ObjStyle
            }
            $styleSet = $true
        }
        catch {
            Write-Log "Failed to set style: $ObjStyle" 3
        }
    }
    $styleSet
}

function Add-DocObjectScripts {
    param($documentedObj)

    foreach ($scriptItem in $documentedObj.Scripts) {
        if (-not $scriptItem.ScriptContent -or -not $scriptItem.Caption) { continue }
        Add-DocTableScript $scriptItem.Caption $scriptItem.Header $scriptItem.ScriptContent
    }
}

Invoke-InitializeWordOutput
