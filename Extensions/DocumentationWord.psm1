# Documentation Output Provider for Word 

#https://docs.microsoft.com/en-us/office/vba/api/overview/word
function Get-ModuleVersion
{
    '1.7.0'
}

function Invoke-InitializeModule
{
    if(!("Microsoft.Office.Interop.Word.Application" -as [Type]))
    {
        try
        {
            Add-Type -AssemblyName Microsoft.Office.Interop.Word
        }
        catch
        {
            Write-LogError "Failed to add Word Interop type. Cannot create word documents. Verify that Word is installed properly." $_.Exception
            return
        }
    }

    Add-OutputType ([PSCustomObject]@{
        Name="Word"
        Value="word"
        OutputOptions = (Add-WordOptionsControl)
        Activate = { Invoke-WordActivate @args }
        PreProcess = { Invoke-WordPreProcessItems @args }
        NewObjectGroup = { Invoke-WordNewObjectGroup @args }
        NewObjectType = { Invoke-WordNewObjectType @args }
        Process = { Invoke-WordProcessItem @args }
        PostProcess = { Invoke-WordPostProcessItems @args }
        ProcessAllObjects = { Invoke-WordProcessAllObjects @args }
    })    
}

function Add-WordOptionsControl
{
    $script:wordForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\DocumentationWordOptions.xaml") -AddVariables

    $global:cbWordDocumentationProperties.ItemsSource = ("[ { Name: `"Simple (Name and Value)`",Value: `"simple`" }, { Name: `"Extended`",Value: `"extended`" }, { Name: `"Custom`",Value: `"custom`" }]" | ConvertFrom-Json)
    $global:cbWordDocumentationProperties.SelectedValue = (Get-Setting "Documentation" "WordExportProperties" "simple")
    $global:txtWordCustomProperties.Text = (Get-Setting "Documentation" "WordCustomDisplayProperties" "Name,Value,Category,SubCategory")

    $global:spWordCustomProperties.Visibility = (?: ($global:cbWordDocumentationProperties.SelectedValue -ne "custom") "Collapsed" "Visible")
    $global:txtWordCustomProperties.Visibility = (?: ($global:cbWordDocumentationProperties.SelectedValue -ne "custom") "Collapsed" "Visible")

    $global:cbWordDocumentationLevel.ItemsSource = ("[ { Name: `"Full`",Value: `"full`" }, { Name: `"Limited`",Value: `"limited`" }, { Name: `"Basic`",Value: `"basic`" }]" | ConvertFrom-Json)
    $global:cbWordDocumentationLevel.SelectedValue = (Get-Setting "Documentation" "WordDocumentationLevel" "full")

    $global:gdWordDocumentationLimitOptions.Visibility = (?: ($global:cbWordDocumentationLevel.SelectedValue -ne "limited") "Collapsed" "Visible")
    $global:txtWordDocumentationLimitMaxLength.Text = Get-Setting "Documentation" "WordDocumentationLimitMaxLength" ""
    $global:txtWordDocumentationLimitTruncateLength.Text = Get-Setting "Documentation" "WordDocumentationLimitTruncateLength" ""
    $global:chkWordDocumentationLimitAttach.IsChecked = ((Get-Setting "Documentation" "WordDocumentationLimitAttatch" "true") -ne "false")
    
    $global:txtWordDocumentTemplate.Text = Get-Setting "Documentation" "WordDocumentTemplate" ""
    $global:txtWordDocumentName.Text = (Get-Setting "Documentation" "WordDocumentName" "%MyDocuments%\%Organization%-%Date%.docx")
    
    $global:chkWordAddCategories.IsChecked = ((Get-Setting "Documentation" "WordAddCategories" "true") -ne "false")
    $global:chkWordAddSubCategories.IsChecked = ((Get-Setting "Documentation" "WordAddSubCategories" "true") -ne "false")
    
    $global:txtWordHeader1Style.Text = Get-Setting "Documentation" "WordHeader1Style" "Heading 1"
    $global:txtWordHeader2Style.Text = Get-Setting "Documentation" "WordHeader2Style" "Heading 2"
    $global:txtWordHeader3Style.Text = Get-Setting "Documentation" "WordHeader3Style" "Heading 3" 
    $global:txtWordTableStyle.Text = Get-Setting "Documentation" "WordTableStyle" "Grid table 4 - Accent 3"
    $global:txtWordTableHeaderStyle.Text = Get-Setting "Documentation" "WordTableHeaderStyle" ""
    $global:txtWordCategoryHeaderStyle.Text = Get-Setting "Documentation" "WordCategoryHeaderStyle" ""
    $global:txtWordSubCategoryHeaderStyle.Text = Get-Setting "Documentation" "WordSubCategoryHeaderStyle" ""
    $global:txtWordTableTextStyle.Text = Get-Setting "Documentation" "WordTableTextStyle" ""

    $global:cbWordTableCaptionPosition.ItemsSource = ("[ { Name: `"Above`",Value: `"above`" }, { Name: `"Below`",Value: `"below`" }]" | ConvertFrom-Json)
    $global:cbWordTableCaptionPosition.SelectedValue = (Get-Setting "Documentation" "WordTableCaptionPosition" "below")
    
    $global:txtWordContentControls.Text = Get-Setting "Documentation" "WordContentControls" "Year=;Address="
    $global:txtWordTitleProperty.Text = Get-Setting "Documentation" "WordTitleProperty" "Intune documentation"
    $global:txtWordSubjectProperty.Text = Get-Setting "Documentation" "WordSubjectProperty" "Intune documentation"
    
    $global:chkWordAttachJsonFile.IsChecked = ((Get-Setting "Documentation" "WordAttatchJsonFile" "false") -ne "false")
    
    $global:txtWordScriptTableStyle.Text = Get-Setting "Documentation" "WordScriptTableStyle" ""
    $global:txtWordScriptStyle.Text = Get-Setting "Documentation" "WordScriptStyle" 

    $global:chkWordOpenDocument.IsChecked = ((Get-Setting "Documentation" "WordOpenDocument" "true") -ne "false")

    Add-XamlEvent $script:wordForm "browseWordDocumentTemplate" "add_click" {
        $of = [System.Windows.Forms.OpenFileDialog]::new()
        $of.Multiselect = $false
        $of.Filter = "Word Templates (*.dotx)|*.dotx"
        if($of.ShowDialog())
        {
            Set-XamlProperty $script:wordForm "txtWordDocumentTemplate" "Text" $of.FileName
            Save-Setting "Documentation" "WordDocumentTemplate" $of.FileName
        }                
    }
    
    Add-XamlEvent $script:wordForm "cbWordDocumentationProperties" "add_selectionChanged" {
        $global:spWordCustomProperties.Visibility = (?: ($this.SelectedValue -ne "custom") "Collapsed" "Visible")
        $global:txtWordCustomProperties.Visibility = (?: ($this.SelectedValue -ne "custom") "Collapsed" "Visible")
    }

    Add-XamlEvent $script:wordForm "cbWordDocumentationLevel" "add_selectionChanged" {
        $global:gdWordDocumentationLimitOptions.Visibility = (?: ($this.SelectedValue -ne "limited") "Collapsed" "Visible")
    }
    
    $script:wordForm
}
function Invoke-WordActivate
{
    #$global:chkWordAddCompanyName.IsChecked = (Get-SettingValue "AddCompanyName")
}

function Invoke-WordPreProcessItems
{
    Save-Setting "Documentation" "WordExportProperties" $global:cbWordDocumentationProperties.SelectedValue
    Save-Setting "Documentation" "WordCustomDisplayProperties" $global:txtWordCustomProperties.Text
    Save-Setting "Documentation" "WordDocumentTemplate" $global:txtWordDocumentTemplate.Text
    Save-Setting "Documentation" "WordDocumentName" $global:txtWordDocumentName.Text

    Save-Setting "Documentation" "WordDocumentationLevel" $global:cbWordDocumentationLevel.SelectedValue
    Save-Setting "Documentation" "WordDocumentationLimitMaxLength" $global:txtWordDocumentationLimitMaxLength.Text
    Save-Setting "Documentation" "WordDocumentationLimitTruncateLength" $global:txtWordDocumentationLimitTruncateLength.Text
    Save-Setting "Documentation" "WordDocumentationLimitAttatch" $global:chkWordDocumentationLimitAttach.IsChecked

    Save-Setting "Documentation" "WordAddCategories" $global:chkWordAddCategories.IsChecked
    Save-Setting "Documentation" "WordAddSubCategories" $global:chkWordAddSubCategories.IsChecked
    Save-Setting "Documentation" "WordOpenDocument" $global:chkWordOpenDocument.IsChecked

    Save-Setting "Documentation" "WordHeader1Style" $global:txtWordHeader1Style.Text
    Save-Setting "Documentation" "WordHeader2Style" $global:txtWordHeader2Style.Text
    Save-Setting "Documentation" "WordHeader3Style" $global:txtWordHeader3Style.Text
    Save-Setting "Documentation" "WordTableStyle" $global:txtWordTableStyle.Text
    Save-Setting "Documentation" "WordTableHeaderStyle" $global:txtWordTableHeaderStyle.Text
    Save-Setting "Documentation" "WordCategoryHeaderStyle" $global:txtWordCategoryHeaderStyle.Text
    Save-Setting "Documentation" "WordSubCategoryHeaderStyle" $global:txtWordSubCategoryHeaderStyle.Text
    Save-Setting "Documentation" "WordTableTextStyle" $global:txtWordTableTextStyle.Text
    Save-Setting "Documentation" "WordTableCaptionPosition" $global:cbWordTableCaptionPosition.SelectedValue
    
    Save-Setting "Documentation" "WordContentControls" $global:txtWordContentControls.Text
    Save-Setting "Documentation" "WordTitleProperty" $global:txtWordTitleProperty.Text
    Save-Setting "Documentation" "WordSubjectProperty" $global:txtWordSubjectProperty.Text

    Save-Setting "Documentation" "WordAttatchJsonFile" $global:chkWordAttachJsonFile.IsChecked

    Save-Setting "Documentation" "WordScriptTableStyle" $global:txtWordScriptTableStyle.Text
    Save-Setting "Documentation" "WordScriptStyle" $global:txtWordScriptStyle.Text

    $script:limitMaxValue = 100
    $script:truncateValueLength = $script:limitMaxValue

    if($global:cbWordDocumentationLevel.SelectedValue -eq "limited")
    {
        if($global:txtWordDocumentationLimitMaxLength.Text)    
        {
            try
            {
                $script:limitMaxValue = [int]::Parse($global:txtWordDocumentationLimitMaxLength.Text)
            }
            catch
            {
                Write-LogError "Failed to parse $($global:txtWordDocumentationLimitMaxLength.Text) to int. Max value length will be set to 100." $_.Exception
            }
        }

        if($global:txtWordDocumentationLimitTruncateLength.Text)    
        {
            try
            {
                $script:truncateValueLength = [int]::Parse($global:txtWordDocumentationLimitTruncateLength.Text)
            }
            catch
            {
                Write-LogError "Failed to parse $($global:txtWordDocumentationLimitTruncateLength.Text) to int. Truncat length will be set to $script:limitMaxValue." $_.Exception
            }
        }
        
        if($script:limitMaxValue -lt 20)
        {
            Write-Log "Max value length must be 20 or more. Changed to 20" 2
            $script:limitMaxValue = 0
        }

        if($script:truncateValueLength -lt 0)
        {
            Write-Log "Truncate length must be 0 or more. Changed to 0" 2
            $script:truncateValueLength = 0
        }        
        elseif($script:truncateValueLength -gt $script:limitMaxValue)
        {
            Write-Log "Truncate length cannot be larger than Max value length. Canged to: $($script:limitMaxValue)" 2
            $script:truncateValueLength = $script:limitMaxValue
        }
    }

    try
    {
        $script:wordApp = New-Object -ComObject Word.Application
    }
    catch
    {
        Write-LogError "Failed to create Word App object. Word documentation aborted..." $_.Exception
        return $false
    }
    
    #$wordApp.Visible = $true

    if($global:txtWordDocumentTemplate.Text)
    {
        try
        {
            $script:doc = $script:wordApp.Documents.Add($global:txtWordDocumentTemplate.Text) 
        }
        catch
        {
            Write-LogError "Failed to create document based on template: $($global:txtWordDocumentTemplate.Text)" $_.Exception
        }
    }
    else
    {
        $script:doc = $script:wordApp.Documents.Add() 
    }

    #Get BuiltIn properties
    $script:builtInProps = @()
    $script:doc.BuiltInDocumentProperties | foreach-object { 

        $name = [System.__ComObject].invokemember("name",[System.Reflection.BindingFlags]::GetProperty,$null,$_,$null)
        try
        {
        $value = [System.__ComObject].invokemember("value",[System.Reflection.BindingFlags]::GetProperty,$null,$_,$null)
        }
        catch{}

        if($name)
        {
            $script:builtInProps += [PSCustomObject]@{
            Name = $name
            Value = $value
            }
        }
    }

    #Get Custom properties
    $script:customProps = @()
    $script:doc.CustomDocumentProperties | foreach-object { 

        $name = [System.__ComObject].invokemember("name",[System.Reflection.BindingFlags]::GetProperty,$null,$_,$null)
        try
        {
            $value = [System.__ComObject].invokemember("value",[System.Reflection.BindingFlags]::GetProperty,$null,$_,$null)
        }
        catch{}

        if($name)
        {
            $script:customProps += [PSCustomObject]@{
            Name = $name
            Value = $value
            }
        }
    }

    $script:wordStyles = @()
    $script:doc.Styles | foreach { 
        $script:wordStyles += [PSCUstomObject]@{
            Name=$_.NameLocal
            Type=$_.Type
            Style=$_
            }
    }

    $script:builtinStyles = [Enum]::GetNames([Microsoft.Office.Interop.Word.wdBuiltinStyle])

    if(-not $global:txtWordDocumentTemplate.Text)
    {
        $script:doc.Application.Templates.LoadBuildingBlocks()
        $BuildingBlocks = $script:doc.Application.Templates | Where {$_.name -eq 'Built-In Building Blocks.dotx'}
        if($BuildingBlocks)
        {    
            $coverPageName = ?? $global:txtWordCoverPage.Text 'Ion (Dark)'  
            try
            {
                $blocks = @()

                for($i = 1;$i -le $BuildingBlocks.BuildingBlockEntries.Count;$i++)
                {
                    $blocks += $BuildingBlocks.BuildingBlockEntries.Item($i)
                }
                
                $coverPages = (($blocks | Where { $_.Type.Index -eq 2 } | Select Name) | Sort -Property Name).Name

                if(($coverPages | measure).Count -gt 0)
                {
                    if($coverPageName -notin $coverPages)
                    {
                        Write-Log "$coverPageName not found in available Cover Page list. Using: $($coverPages[0])"
                        Write-Log "Available Cover Pages: $(($coverPages -join ","))"
                        $coverPageName = $coverPages[0]
                    }
                    else
                    {
                        Write-Log "Add Cover Page: $coverPageName"
                    }
                }

                $coverPage = $BuildingBlocks.BuildingBlockEntries.Item($coverPageName)
                $coverPage.Insert($script:wordApp.Selection.Range,$true) | Out-Null
                $script:wordApp.Selection.InsertNewPage()
            }
            catch 
            {
                Write-LogError "Failed to create Cover Page" $_.Exception
            }

            try
            {
                $coverPageProps = $script:doc.CustomXMLParts | where { $_.NamespaceURI -match "coverPageProps$" }
                if($coverPageProps)
                {
                    Write-Log "Available Cover Page properties for $($coverPageName): $(((([xml]$coverPageProps.DocumentElement.XML).ChildNodes[0].ChildNodes).Name -join ","))"
                }
            }
            catch{}

            try
            {
                $script:doc.TablesOfContents.Add($script:wordApp.Selection.Range) | out-null
                $script:wordApp.Selection.InsertNewPage()
            }
            catch
            {
                Write-LogError "Failed to create Table of Contents" $_.Exception
            }
        }
    }
    else
    {
        if(($script:doc.TablesOfContents | measure).Count -eq 0)
        {
            # Where should it be added?
            # $script:doc.TablesOfContents.Add($script:wordApp.Selection.Range) | out-null            
        }
        
        Invoke-DocGoToEnd
        $script:wordApp.Selection.InsertNewPage()
    }
}

function Invoke-WordPostProcessItems
{
    $userName = $global:me.displayName
    if($global:me.givenName -and $global:me.surname)
    {
        $userName = ($global:me.givenName + " " + $global:me.surname)
    }
    
    #Add properties - ToDo: This is static...
    Set-WordDocBuiltInProperty "wdPropertyTitle" (?? $global:txtWordTitleProperty.Text "Intune documentation")
    Set-WordDocBuiltInProperty "wdPropertySubject" (?? $global:txtWordSubjectProperty.Text "Intune documentation")
    Set-WordDocBuiltInProperty "wdPropertyAuthor" $userName
    Set-WordDocBuiltInProperty "wdPropertyCompany" $global:Organization.displayName
    Set-WordDocBuiltInProperty "wdPropertyKeywords" "Intune,Endpoint Manager,MEM"
    
    try
    {
        # ToDo: Add support for custom properties
        # Add: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.core.documentproperties.add?view=office-pia
        # Types: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.core.msodocproperties?view=office-pia
        #$coverPageProps = $script:doc.CustomXMLParts | where { $_.NamespaceURI -match "coverPageProps$" }
        #[System.__ComObject].InvokeMember("add",[System.Reflection.BindingFlags]::InvokeMethod,$null,$script:doc.CustomDocumentProperties,([array]("PropName", $false, 4, "PropValue")))
        
        $ContentControlProperties = $global:txtWordContentControls.Text #"Year=;Address=TestAddress"
        foreach($ccObj in $ContentControlProperties.Split(';'))
        {
            $ccName,$ccVal = $ccObj.Split('=')
            Set-WordContentControlText $ccName $ccVal
        }
    }
    catch {}

    #update fields, ToC etc.
    $script:doc.Fields | ForEach-Object -Process { $_.Update() | Out-Null } 
    $script:doc.TablesOfContents | ForEach-Object -Process { $_.Update() | Out-Null }
    $script:doc.TablesOfFigures | ForEach-Object -Process { $_.Update() | Out-Null }
    $script:doc.TablesOfAuthorities | ForEach-Object -Process { $_.Update() | Out-Null }

    $fileName = $global:txtWordDocumentName.Text
    if(-not $fileName)
    {
        $fileName = "%MyDocuments%\%Organization%-%Date%.docx"
    }

    $fileName = Expand-FileName $fileName

    $format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault

    try
    {
        $script:doc.SaveAs2([ref]$fileName,[ref]$format)
        Write-Log "Document $fileName saved successfully"
    }
    catch
    {
        Write-LogError "Failed to save file $fileName" $_.Excption
    }

    if($global:chkWordOpenDocument.IsChecked -eq $true -and $global:hideUI -ne $true)
    {
        $script:wordApp.Visible = $true
        $script:wordApp.WindowState = [Microsoft.Office.Interop.Word.WdWindowState]::wdWindowStateMaximize
        $script:wordApp.Activate()
        [Console.Window]::SetForegroundWindow($script:wordApp.ActiveWindow.Hwnd) | Out-Null
    }
    else
    {
        $script:doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
        $script:wordApp.Quit()
    }
}

function Set-WordContentControlText 
{
    param($controlName, $value)

    if(-not $controlName) { return }

    try
    {
        $ctrl = $script:doc.SelectContentControlsByTitle($controlName)

        if($ctrl) 
        {
            Write-LogDebug "Update ContentControl $controlName (Type: $($ctrl[1].Type))"
            if($ctrl[1].Type -eq 6)
            {
                if($ctrl[1].DateDisplayFormat)
                {
                    $ctrl[1].Range.Text = (Get-Date).ToString($ctrl[1].DateDisplayFormat)
                }
                else
                {
                    $ctrl[1].Range.Text = (Get-Date).ToShortDateString()
                }
            }
            else
            {
                if(-not $value) { return }

                $ctrl[1].Range.Text = $value
            }
        }
        else
        {
            #Write-Log "No ContentControl found with name $controlName" 2
        }
    }
    catch
    {
        Write-LogError "Failed to set ContentControl $controlName" $_.Exception
    }
}

function Invoke-WordNewObjectGroup
{
    param($groupId)

    $objectTypeString = Get-ObjectTypeString -ObjectType $groupId

    Add-DocText $objectTypeString $global:txtWordHeader1Style.Text
}

function Invoke-WordNewObjectType
{
    param($objectTypeName)

    $script:objectHeaderLevel = 2

    Add-DocText $objectTypeName (Get-ObjectLevelHeader)

    $script:objectHeaderLevel = 3
}

function local:Get-ObjectLevelHeader
{
    if($script:objectHeaderLevel -eq 3 -and $global:txtWordHeader3Style.Text)
    {
        return $global:txtWordHeader3Style.Text
    }
    return $global:txtWordHeader2Style.Text
}

function Invoke-WordProcessItem
{
    param($obj, $objectType, $documentedObj)

    if(!$documentedObj -or !$obj -or !$objectType) { return }

    $objName = Get-GraphObjectName $obj $objectType

    Add-DocText $objName (Get-ObjectLevelHeader)
    
    $script:doc.Application.Selection.TypeParagraph()

    try 
    {
        #TableHeaders.value
        #Inputs.displayNameLabel

        foreach($tableType in @("BasicInfo","FilteredSettings"))
        {
            if($tableType -eq "BasicInfo")
            {
                $properties = @("Name","Value")
            }
            elseif($global:cbWordDocumentationProperties.SelectedValue -eq 'extended' -and $documentedObj.DisplayProperties)
            {
                $properties = @("Name","Value","Description")
            }
            elseif($global:cbWordDocumentationProperties.SelectedValue -eq 'custom' -and $global:txtWordCustomProperties.Text)
            {
                $properties = @()
                
                foreach($prop in $global:txtWordCustomProperties.Text.Split(","))
                {
                    # This will add language support for custom columns (or replacing existing header)
                    $propInfo = $prop.Split('=')
                    if(($propInfo | measure).Count -gt 1)
                    {
                        $properties += $propInfo[0] 
                        Set-DocColumnHeaderLanguageId $propInfo[0] $propInfo[1]
                    }
                    else
                    {
                        $properties += $prop
                    }
                }
            }        
            else
            {
                $properties = (?? $documentedObj.DefaultDocumentationProperties (@("Name","Value")))
            }

            if($global:cbWordDocumentationLevel.SelectedValue -eq "basic" -and $tableType -ne "BasicInfo")
            {
                continue
            }
            
            $lngId = ?: ($tableType -eq "BasicInfo") "SettingDetails.basics" "TableHeaders.settings" -AddCategories

            if(($documentedObj.$tableType).Count -gt 0) {
                Add-DocTableItems $obj $objectType ($documentedObj.$tableType) $properties $lngId `
                -AddCategories:($global:chkWordAddCategories.IsChecked -eq $true) `
                -AddSubcategories:($global:chkWordAddSubCategories.IsChecked -eq $true) `
                -ForceFullValue:($tableType -eq "BasicInfo")
            }
        }

        if($global:cbWordDocumentationLevel.SelectedValue -ne "basic")
        {
            if(($documentedObj.ComplianceActions | measure).Count -gt 0)
            {
                $properties = @("Action","Schedule","MessageTemplate","EmailCC")

                Add-DocTableItems $obj $objectType $documentedObj.ComplianceActions $properties "Category.complianceActionsLabel"
            }

            if(($documentedObj.ApplicabilityRules | measure).Count -gt 0)
            {
                $properties = @("Rule","Property","Value")

                Add-DocTableItems $obj $objectType $documentedObj.ApplicabilityRules $properties "SettingDetails.applicabilityRules"
            }

            Add-DocObjectSettings $obj $objectType $documentedObj

            foreach($customTable in ($documentedObj.CustomTables | Sort-Object -Property Order)) 
            {
                Add-DocTableItems $obj $objectType $documentedObj $customTable.Values $customTable.Columns $customTable.LanguageId -AddCategories -AddSubcategories
            }
        }

        if(($documentedObj.Assignments | measure).Count -gt 0)
        {
            $params = @{}
            $settingProps = $null
            if($documentedObj.Assignments[0].RawIntent)
            {
                $properties = @("GroupMode","Group","Filter","FilterMode")

                $settingProps = @("Filter","FilterMode")
            
                $settingsObj = $documentedObj.Assignments | Where { $_.Settings -ne $null } | Select -First 1

                if($settingsObj)
                {
                    foreach($objProp in $settingsObj.Settings.Keys)
                    {
                        if($objProp -in $properties) { continue }
                        if($objProp -in @("Category","RawIntent")) { continue }
                        $settingProps += ("Settings." + $objProp)
                    }
                }
                $params.Add("AddCategories", $true)
            }
            else
            {
                $isFilterAssignment = $false
                foreach($assignment in $documentedObj.Assignments)
                {
                    if(($assignment.PSObject.Properties | Where Name -eq "FilterMode"))
                    {
                        $isFilterAssignment = $true
                        break
                    }
                }
                $properties = @("Group")
                if($isFilterAssignment)
                {
                    $properties += @("Filter","FilterMode")
                }
                $params.Add("AddCategories", $true)
            }

            # Creates a standard assignments table 
            Add-DocTableItems $obj $objectType $documentedObj.Assignments $properties "TableHeaders.assignments" @params

            if($null -ne $settingProps)
            {
                # Adds additional values to the assignments table for Apps assignments 
                Set-DocTableSettingsItems $obj $objectType $documentedObj.Assignments $settingProps 3
            }
        }

        if($global:chkWordAttachJsonFile.IsChecked -eq $true)
        {
            $fileName = Export-GraphObject $obj $objectType ([IO.Path]::GetTempPath()) -IsFullObject -PassThru -SkipAddID
            if($fileName)
            {
                $fi = [IO.FileInfo]$fileName
                if($fi.Exists)
                {
                    $script:doc.Application.Selection.InlineShapes.AddOLEObject("",$fi.FullName,$false,$true,"$($env:WinDir)\System32\Notepad.exe",0,$fi.Name)
                    $script:doc.Application.Selection.TypeParagraph()
                    try { $fi.Delete() } catch {} # Cleanup
                }
            }
        }        
    }
    catch 
    {
        Write-LogError "Failed to process object $objName" $_.Exception
    }
}

function Set-DocTableSettingsItems
{
    param($obj, $objectType, $items, $properties, $firstColumn)

    $secondColumn = $firstColumn + 1

    $script:docTable.Cell(1, $firstColumn).Range.Text = (Invoke-DocTranslateColumnHeader "Settings")
    $script:docTable.Cell(1, $secondColumn).Range.Text = ""

    $row = 2
    foreach($itemObj in $items)
    {
        #if($script:docTable.Rows($row).Cells.Count -eq 1) { $row++;continue } # Category / Sub-category
        while($script:docTable.Cell($row,1).Next.RowIndex -gt $row) 
        { 
            # Category / Sub-category
            $row++;
        } 
        $script:docTable.Cell($row, $firstColumn).Range.Text = ""
        $script:docTable.Cell($row, $secondColumn).Range.Text = ""
        $script:docTable.Cell($row, $firstColumn).Split($properties.Count,1)
        $script:docTable.Cell($row, $secondColumn).Split($properties.Count,1)
        
        $cellRow = $row
        foreach($settingProp in $properties)
        {
            $script:docTable.Cell($cellRow, $firstColumn).Range.Text = (Invoke-DocTranslateColumnHeader ($settingProp.Split('.')[-1]))
            
            $propArr = $settingProp.Split('.')
            $tmpObj = $itemObj
            $propName = $propArr[-1]
            for($x = 0; $x -lt ($propArr.Count - 1);$x++)
            {
                $tmpObj = $tmpObj."$($propArr[$x])"
            }

            $script:docTable.Cell($cellRow, $secondColumn).Range.Text = "$($tmpObj.$propName)"

            $cellRow++
        }
        $row = $row + $properties.Count
    }
}

function Invoke-WordProcessAllObjects
{
    param($allObjectTypeObjects)

    if(($allObjectTypeObjects | measure).Count -eq 0) { return }

    $tmpObj = $allObjectTypeObjects | Select -First 1
    if(-not $tmpObj) { return }

    $objectType = $tmpObj.Object.ObjectType
    if($objectType.Id -eq "ScopeTags")
    {
        $objTypeName = Get-LanguageString "SettingDetails.scopeTags"

        Add-DocText $objTypeName (Get-ObjectLevelHeader)        

        $script:doc.Application.Selection.TypeParagraph()  

        $items = @()

        $nameLabel = Get-LanguageString "Inputs.displayNameLabel"
        $descriptionLable = Get-LanguageString "TableHeaders.description" 
        foreach($obj in $allObjectTypeObjects.Object.Object)
        {
            $items += [PSCustomObject]@{
                $nameLabel = $obj.displayName
                ID = $obj.Id
                $descriptionLable = $obj.Description
                Object = $obj
            }
        }

        $items = $items | Sort -Property $nameLabel

        $properties = @($nameLabel,"id",$descriptionLable)

        Add-DocTableItems $tmpObj.Object.Object $tmpObj.Object.ObjectType $items $properties -captionOverride (Get-LanguageString "SettingDetails.scopeTags")
    }
}

function Invoke-WordCustomProcessItems
{
    param($obj, $documentedObj)

}

function Add-DocTableItems
{
    param($obj, $objectType, $items, $properties, $lngId, [switch]$AddCategories, [switch]$AddSubcategories, $captionOverride, [switch]$forceFullValue)

    if(($items | measure).Count -eq 0)
    {
        return
    }

    $tblHeaderStyle = $global:txtWordTableHeaderStyle.Text
    $tblCategoryStyle = $global:txtWordCategoryHeaderStyle.Text
    $tblSubCategoryStyle = $global:txtWordSubCategoryHeaderStyle.Text
    $tblTextStyle = $global:txtWordTableTextStyle.Text
    $txtTableCaptionPosition = $global:cbWordTableCaptionPosition.SelectedValue

    $range = $script:doc.application.selection.range
    
    $script:docTable = $script:doc.Tables.Add($range, (($items | measure).Count + 1), $properties.Count, [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior, [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow)
    $script:docTable.ApplyStyleHeadingRows = $true
    Set-DocObjectStyle $script:docTable $global:txtWordTableStyle.Text | Out-null

    if($captionOverride)
    {
        $caption = $captionOverride
    }
    elseif($lngId)
    {
        $caption = "$((Get-LanguageString $lngId)) - $((Get-GraphObjectName $obj $objectType))"
    }
    else
    {
        $caption = "$((Get-GraphObjectName $obj $objectType)) ($($objectType.Title))"
    }

    $i = 1
    foreach($prop in $properties)
    {
        $script:docTable.Cell(1, $i).Range.Text = (Invoke-DocTranslateColumnHeader ($prop.Split(".")[-1]))
        $i++
    }

    if(!(Set-DocObjectStyle $script:docTable.Rows(1).Range $tblHeaderStyle))
    {
        $script:docTable.Rows(1).Range.Font.Size += 2
        $script:docTable.Rows(1).Range.Font.Bold = $true
    }    
    
    $curCategory = ""
    $curSubCategory = ""

    $row = 2
    $curentPropertyIndex = 0
    foreach($itemObj in $items)
    {
        try 
        {
            if(($itemObj.Category -and $curCategory -ne $itemObj.Category -and $AddCategories -eq $true) -or
                ($itemObj.SubCategory -and $curSubCategory -ne $itemObj.SubCategory -and $AddSubcategories -eq $true))
            {
                $curentPropertyIndex = 0
            }

            if($itemObj.PropertyIndex -is [int] -and $itemObj.PropertyIndex -gt 0 -and $itemObj.PropertyIndex -eq 1)
            {
                $curentPropertyIndex = $itemObj.PropertyIndex
                # !!! ToDo: Set style for new property
            }                

            $i = 1
            foreach($prop in $properties)
            {
                try
                {
                    # This adds support for properties like Settings.PropName
                    $propArr = $prop.Split('.')
                    $tmpObj = $itemObj
                    $propName = $propArr[-1]
                    for($x = 0; $x -lt ($propArr.Count - 1);$x++)
                    {
                        $tmpObj = $tmpObj."$($propArr[$x])"
                    }
                    $propValue = "$($tmpObj.$propName)"
                    $propValueFull = $null

                    if($forceFullValue -ne $true -and $global:cbWordDocumentationLevel.SelectedValue -eq "limited" -and $propValue.Length -gt $script:limitMaxValue)
                    {
                        $propValueFull = $propValue
                        if($script:truncateValueLength -gt 0)
                        {
                            $propValue = ($propValue.Substring(0, $script:truncateValueLength) + "...")
                            if($global:chkWordDocumentationLimitAttach.IsChecked -eq $true)
                            {
                                $propValue = ("`r`n" + $propValue)
                            }                            
                        }
                        else
                        {
                            $propValue = $null
                        }
                    }

                    $levelExtra = ""
                    if($i -eq 1 -and $itemObj.Level)
                    {
                        try
                        {
                            $level = ([int]$itemObj.Level) # - 1
                            if($level -lt 0)  { $level = 0 }
                            if($level -gt 0)
                            {
                                $levelExtra = [String]::new(" ", ($level * 2)) #Should probably use tab stops instead
                            }
                        }
                        catch{}
                    }

                    if($null -ne $propValue)
                    {
                        $script:docTable.Cell($row, $i).Range.Text = "$levelExtra$propValue"
                    }

                    if($propValueFull -and $global:chkWordDocumentationLimitAttach.IsChecked -eq $true)
                    {
                        if($null -ne $propValue)
                        {
                            #$script:doc.Application.Selection.TypeParagraph()
                        }

                        $tmpName = "$((Get-GraphObjectName $obj $objectType))-$propName"
                                                
                        $tmpFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "$($tmpName).txt")
                        $tmpFile = Remove-InvalidFileNameChars $tmpFile
                        $propValueFull | Out-File -LiteralPath $tmpFile -Force 
                        $fi = [IO.FileInfo]$tmpFile
                        [void]$script:docTable.Cell($row, $i).Range.InlineShapes.AddOLEObject("",$fi.FullName,$false,$true,"$($env:WinDir)\System32\Notepad.exe",0,"Full value")
                        try { $fi.Delete() } catch {}
                    }
                }
                catch
                {
                    Write-LogError "Failed to add property value for $prop" $_.Exception
                }
                $i++
            }

            Set-DocObjectStyle $script:docTable.Rows($row).Range $tblTextStyle | Out-Null
        
            if($itemObj.Category -and $curCategory -ne $itemObj.Category -and $AddCategories -eq $true)
            {
                # Insert row for the Category above the new row
                $script:docTable.Rows.Add($script:docTable.Rows.Item($row)) | Out-Null
                $script:docTable.Rows.Item($row).Cells.Merge()
                $script:docTable.Cell($row, 1).Range.Text = $itemObj.Category
                
                if(!(Set-DocObjectStyle $script:docTable.Rows($row).Range $tblCategoryStyle))
                {
                    $script:docTable.Rows($row).Range.Font.Size += 2
                    $script:docTable.Rows($row).Range.Font.Italic = $true
                }
                $row++
                $curCategory = $itemObj.Category
                $curSubCategory = ""
            }

            if($itemObj.SubCategory -and $curSubCategory -ne $itemObj.SubCategory -and $AddSubcategories -eq $true)
            {
                # Insert row for the SubCategory above the new row
                $script:docTable.Rows.Add($script:docTable.Rows.Item($row)) | Out-Null
                $script:docTable.Rows.Item($row).Cells.Merge()
                $script:docTable.Cell($row, 1).Range.Text = $itemObj.SubCategory
                
                if(!(Set-DocObjectStyle $script:docTable.Rows($row).Range $tblSubCategoryStyle))
                {
                    $script:docTable.Rows($row).Range.Font.Italic = $true
                }
                $row++
                $curSubCategory = $itemObj.SubCategory
            }
        }
        catch 
        {
            Write-Log "Failed to process property" 2    
        }

        $row++
    }
    
    # -2 = Table caption, 1 = Below / 0 = Above
    if($txtTableCaptionPosition -eq "above")
    {
        $capPos = 0
    }
    else
    {
        $capPos = 1
    }
    $script:docTable.Application.Selection.InsertCaption(-2, ". $caption", $null, $capPos)
    
    Invoke-DocGoToEnd

    # Add new row after the table
    #$script:doc.Application.Selection.InsertParagraphAfter()    
    $script:doc.Application.Selection.TypeParagraph()
}

function Add-DocTableScript
{
    param($caption, $header, $script)

    if(-not $script) { return }

    $tblScriptStyle = (?? $global:txtWordScriptTableStyle.Text $global:txtWordTableStyle.Text)

    $range = $script:doc.application.selection.range
    
    $scriptTable = $script:doc.Tables.Add($range, 2, 1, [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior, [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow)
    $scriptTable.ApplyStyleHeadingRows = $true
    Set-DocObjectStyle $scriptTable $tblScriptStyle | Out-Null

    if($header)
    {
        $scriptTable.Cell(1, 1).Range.Text = $header
    }

    $scriptTable.Cell(2,1).Range.Font.Bold = $false
    $scriptTable.Cell(2, 1).Range.Text = $script
    if($global:txtWordScriptStyle.Text)
    {
        Set-DocObjectStyle $scriptTable.Rows(2).Range $global:txtWordScriptStyle.Text  | Out-Null
    }
    else
    {
        $tmp = $script:wordStyles | Where Name -like "HTML Code"
        if($tmp)
        {
            $scriptTable.Cell(2,1).Range.Font = $tmp.Style.Font
        }
        $scriptTable.Cell(2,1).Range.Font.Bold = $false
    }
    $scriptTable.Cell(2,1).Range.NoProofing = $true

    # -2 = Table, 1 = Below
    $scriptTable.Application.Selection.InsertCaption(-2, ". $caption", $null, 1)

    # Add new row after the table
    $script:doc.Application.Selection.TypeParagraph()
}

function Get-DocStyle
{
    param($styleName)

    $tmpStyle = ($script:wordStyles | Where Name -like $styleName).Style
    
    # BuiltIn Styles
    #[Enum]::GetNames([Microsoft.Office.Interop.Word.wdBuiltinStyle])
    
    if(!$tmpStyle)
    {
        Write-Log "Style $styleName not found"
    }
    $tmpStyle
}

function Add-DocText
{
    param($text, $style, [switch]$SkipAddParagraph)

    Set-DocObjectStyle $script:doc.application.selection $style | Out-Null
    
    $script:doc.Application.Selection.TypeText($text)  

    if($SkipAddParagraph -ne $true)
    {
        # Add new paragraph by default
        $script:doc.Application.Selection.TypeParagraph()               
    }
}

function Invoke-DocGoToEnd
{
    $script:doc.Application.Selection.goto([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToBookmark, $null, $null, '\EndOfDoc') | Out-Null
}

function Set-WordDocBuiltInProperty
{
    param($propertyName, $value)

    try
    {
        $script:doc.BuiltInDocumentProperties([Microsoft.Office.Interop.Word.WdBuiltInProperty]$propertyName) = $value
    }
    catch
    {
        Write-LogError "Failed to set built in property $propertyName to $value"  $_.Exception
    }
}

function Set-DocObjectStyle
{
    param($docObj, $objStyle)

    $styleSet = $false
    if($docObj -and $objStyle)
    {
        try
        {
            if(($script:builtinStyles | Where { $_ -eq $objStyle }))
            {
                $docObj.style = [Microsoft.Office.Interop.Word.wdBuiltinStyle]$objStyle
            }
            else
            {
                $docObj.style = $objStyle
            }
            $styleSet = $true
        }
        catch
        {
            Write-Log "Failed to set style: $objStyle" 3
        }
    }
    $styleSet
}

function Add-DocObjectSettings
{
    param($obj, $objectType, $documentedObj)

    foreach($objectScript in $documentedObj.Scripts)
    {
        if(-not $objectScript.ScriptContent -or -not $objectScript.Caption) { continue }

        Add-DocTableScript $objectScript.Caption $objectScript.Header $objectScript.ScriptContent
    }   
}