function Get-ModuleVersion
{
    '1.2.0'
}

function Invoke-InitializeModule
{
    Add-OutputType ([PSCustomObject]@{
        Name="Markdown"
        Value="md"
        OutputOptions = (Add-MDOptionsControl)
        #Activate = { Invoke-MDActivate @args }
        PreProcess = { Invoke-MDPreProcessItems @args }
        NewObjectGroup = { Invoke-MDNewObjectGroup2 @args }
        NewObjectType = { Invoke-MDNewObjectType2 @args }
        Process = { Invoke-MDProcessItem @args }
        PostProcess = { Invoke-MDPostProcessItems @args }
        ProcessAllObjects = { Invoke-MDProcessAllObjects @args }
    })        
}

function Add-MDOptionsControl
{
    $script:mdForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\DocumentationMDOptions.xaml") -AddVariables
    
    Set-XamlProperty $script:mdForm "txtMDDocumentName" "Text" (Get-Setting "Documentation" "MDDocumentName" "")
    Set-XamlProperty $script:mdForm "txtMDCSSFile" "Text" (Get-Setting "Documentation" "MDCSSFile" "")
    Set-XamlProperty $script:mdForm "chkMDIncludeCSS" "IsChecked" (Get-Setting "Documentation" "MDIncludeCSS" $true)
    Set-XamlProperty $script:mdForm "chkMDOpenDocument" "IsChecked" (Get-Setting "Documentation" "MDOpenFile" $true)
    Set-XamlProperty $script:mdForm "cbMDDocumentOutputFile" "ItemsSource" ("[ { Name: `"Single file`",Value: `"Full`" }, { Name: `"One file per object`",Value: `"Object`" }]" | ConvertFrom-Json)
    Set-XamlProperty $script:mdForm "cbMDDocumentOutputFile" "SelectedValue" (Get-Setting "Documentation" "MDDocumentFileType" "Full")

    Add-XamlEvent $script:mdForm "browseMDDocumentName" "add_click" {
        $sf = [System.Windows.Forms.SaveFileDialog]::new()
        $sf.DefaultExt = "*.md"
        $sf.Filter = "MD (*.md)|*.md|All files (*.*)|*.*"
        if($sf.ShowDialog() -eq "OK")
        {
            Set-XamlProperty $script:MDForm "txtMDDocumentName" "Text" $sf.FileName
            Save-Setting "Documentation" "MDDocumentName" $sf.FileName
        }                
    }

    Add-XamlEvent $script:mdForm "browseMDCSSFile" "add_click" {
        $of = [System.Windows.Forms.OpenFileDialog]::new()
        $of.Multiselect = $false
        $of.Filter = "CSS Files (*.css)|*.css|All files (*.*)|*.*"
        if($of.ShowDialog())
        {
            Set-XamlProperty $script:mdForm "txtMDCSSFile" "Text" $of.FileName
            Save-Setting "Documentation" "txtMDCSSFile" $of.FileName
        }                
    }
    
    $script:mdForm
}

function Invoke-MDProcessAllObjects
{
    param($allObjectTypeObjects, $objectType)
}

function Invoke-MDPreProcessItems
{
    $script:sectionAnchors = @()
    $script:totAnchors = @()
    $script:mdStrings = $null
    $script:currentItemFileName = $null

    Save-Setting "Documentation" "MDDocumentName" (Get-XamlProperty $script:mdForm "txtMDDocumentName" "Text" "")
    Save-Setting "Documentation" "MDIncludeCSS" (Get-XamlProperty $script:mdForm "chkMDIncludeCSS" "IsChecked")
    Save-Setting "Documentation" "MDCSSFile" (Get-XamlProperty $script:mdForm "txtMDCSSFile" "Text" "")
    Save-Setting "Documentation" "MDOpenFile" (Get-XamlProperty $script:mdForm "chkMDOpenDocument" "IsChecked")
    Save-Setting "Documentation" "MDDocumentFileType" (Get-XamlProperty $script:mdForm "cbMDDocumentOutputFile" "SelectedValue" '')

    $defaultCSSFile = $global:AppRootFolder + "\Documentation\DefaultMDStyle.css"
    $MDCssFile = Get-XamlProperty $script:mdForm "txtMDCSSFile" "Text" $defaultCSSFile
 
    if(-not $MDCssFile)
    {
        Write-Log "CSS file not specified. Using default" 2
        $MDCssFile = $defaultCSSFile
    }
    elseif([IO.File]::Exists($MDCssFile) -eq -$false)
    {
        Write-Log "CSS file $($MDCssFile) not found. Using default" 2
        $MDCssFile = $defaultCSSFile
    }

    $cssStyle = ""
    if([IO.File]::Exists($MDCssFile))
    {
        Write-Log "Using CSS file $($MDCssFile)"
        $cssStyle = Get-Content -Raw -Path $MDCssFile
        $cssStyle += [System.Environment]::NewLine
    }
    else
    {
        Write-Log "CSS file $($MDCssFile) not found. No styles applied" 2
    }
    $script:cssStyle = $cssStyle

    $fileName = Expand-FileName (Get-XamlProperty $script:mdForm "txtMDDocumentName" "Text" "%MyDocuments%\%Organization%-%Date%.md")

    $script:outFile = $fileName
    $script:documentPath = [io.path]::GetDirectoryName($fileName)

    $script:outputType = (Get-XamlProperty $script:mdForm "cbMDDocumentOutputFile" "SelectedValue" "Full")

    if($script:outputType -eq "Object")
    {
        Write-Log "Document one file for each object + index file"
    }
    else
    {
        Write-Log "Document one single file for all objects"
        $script:outputType = "Full"
        $script:mdStrings = [System.Text.StringBuilder]::new()        
    }    
}

function Invoke-MDPostProcessItems
{

    $userName = $global:me.displayName
    if($global:me.givenName -and $global:me.surname)
    {
        $userName = ($global:me.givenName + " " + $global:me.surname)
    }

    $script:mdContent = [System.Text.StringBuilder]::new()

    $script:mdContent.AppendLine("# $((?? $global:txtMDTitleProperty.Text "Intune documentation"))")
    $script:mdContent.AppendLine("")
    $script:mdContent.AppendLine("")

    $mail = ""
    if($global:me.mail)
    {
        $mail = " ($($global:me.mail))"
    }

    $script:mdContent.AppendLine("*Organization:* $($global:Organization.displayName)`n")
    $script:mdContent.AppendLine("*Generated by:* $userName$mail`n")
    $script:mdContent.AppendLine("*Generated:* $((Get-Date).ToShortDateString()) $((Get-Date).ToLongTimeString())`n")

    if($script:sectionAnchors.Count -gt 0)
    {
        $script:mdContent.AppendLine("")
        $script:mdContent.AppendLine("## Table of Contents")
    }

    foreach($header in $script:sectionAnchors)
    {
        $indent = [String]::new(" ", (($header.Level - 1) * 2))
        $script:mdContent.AppendLine("$($indent)- [$($header.Name)]($($header.FileName)#$($header.Anchor))`n")
    }

    $mdText = $script:cssStyle 

    $script:mdContent.AppendLine("")
    $mdText += $script:mdContent.ToString()
    if($script:outputType -eq "Full")
    {
        $mdText += $script:mdStrings.ToString()
    }    
    
    Save-DocumentationFile $mdText $script:outFile -OpenFile:((Get-Setting "Documentation" "MDOpenFile" $true) -eq $true)
    <#
    $fileName = Expand-FileName (Get-XamlProperty $script:mdForm "txtMDDocumentName" "Text" "%MyDocuments%\%Organization%-%Date%.md")
    
    try
    {
        $mdText | Out-File -FilePath $fileName -Force -Encoding utf8 -ErrorAction Stop
        Write-Log "Markdown document $fileName saved successfully"

        if((Get-Setting "Documentation" "MDOpenFile" $true) -eq $true)
        {
            Invoke-Item $fileName
        }
    }
    catch
    {
        Write-LogError "Failed to save Markdown file: $fileName." $_.Exception
    }
    #>
}

function Invoke-MDNewObjectGroup
{
    param($obj, $documentedObj)

    $objectTypeString = Get-ObjectTypeString $obj.Object $obj.ObjectType

    Add-MDHeader "$((?? $objectTypeString $obj.ObjectType.Title))" -Level 1 -USEHtml
}

function Invoke-MDNewObjectType
{
    param($obj, $documentedObj, [int]$groupCategoryCount = 0)

    if($obj.ObjectType.GroupId -eq "EndpointSecurity")
    {
        $objectTypeString = $obj.CategoryName
    }
    else
    {
        $objectTypeString = $obj.ObjectType.Title
    }

    Add-MDHeader "$((?? $objectTypeString $obj.ObjectType.Title))" -Level 2 -USEHtml

}

function Invoke-MDNewObjectGroup2
{
    param($groupId)

    $objectTypeString = Get-ObjectTypeString -ObjectType $groupId

    Add-MDHeader $objectTypeString -Level 1 -USEHtml
}

function Invoke-MDNewObjectType2
{
    param($objectTypeName)

    Add-MDHeader $objectTypeName -Level 2 -USEHtml
}

function Invoke-MDProcessItem
{
    param($obj, $objectType, $documentedObj)

    if(!$documentedObj -or !$obj -or !$objectType) { return }

    $objName = Get-GraphObjectName $obj $objectType

    if($script:outputType -eq "Object")
    {
        $script:totAnchors = @()
        $script:mdStrings = [System.Text.StringBuilder]::new()
        $script:currentItemFileName = "./$((Remove-InvalidFileNameChars "$($objName).md").Replace(" ","_"))"
    }    

    Add-MDHeader $objName -Level 3 -USEHtml

    $script:mdStrings.AppendLine("")

    try 
    {
        foreach($tableType in @("BasicInfo","FilteredSettings"))
        {
            if($tableType -eq "BasicInfo")
            {
                $properties = @("Name","Value")
            }
            elseif($global:cbMDDocumentationProperties.SelectedValue -eq 'extended' -and $documentedObj.DisplayProperties)
            {
                $properties = @("Name","Value","Description")
            }
            elseif($global:cbMDDocumentationProperties.SelectedValue -eq 'custom' -and $global:txtMDCustomProperties.Text)
            {
                $properties = @()
                
                foreach($prop in $global:txtMDCustomProperties.Text.Split(","))
                {
                    # This will add language support for custom columens (or replacing existing header)
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
            
            $lngId = ?: ($tableType -eq "BasicInfo") "SettingDetails.basics" "TableHeaders.settings" -AddCategories

            if(($documentedObj.$tableType).Count -gt 0) {
                Add-MDTableItems $obj $objectType ($documentedObj.$tableType) $properties $lngId -AddCategories -AddSubcategories
            }

            #Add-MDTableItems $obj $objectType ($documentedObj.$tableType) $properties $lngId `
            #    -AddCategories:($global:chkMDAddCategories.IsChecked -eq $true) `
            #    -AddSubcategories:($global:chkMDAddSubCategories.IsChecked -eq $true)                
        }

        if(($documentedObj.ComplianceActions | measure).Count -gt 0)
        {
            $properties = @("Action","Schedule","MessageTemplate","EmailCC")

            Add-MDTableItems $obj $objectType $documentedObj.ComplianceActions $properties "Category.complianceActionsLabel"
        }

        if(($documentedObj.ApplicabilityRules | measure).Count -gt 0)
        {
            $properties = @("Rule","Property","Value")

            Add-MDTableItems $obj $objectType $documentedObj.ApplicabilityRules $properties "SettingDetails.applicabilityRules"
        }

        Add-MDObjectSettings $obj $objectType $documentedObj

        foreach($customTable in ($documentedObj.CustomTables | Sort-Object -Property Order)) 
        {
            Add-MDTableItems $obj $objectType $documentedObj $customTable.Values $customTable.Columns $customTable.LanguageId -AddCategories -AddSubcategories
        }

        if(($documentedObj.Assignments | measure).Count -gt 0)
        {
            if($documentedObj.Assignments[0].RawIntent)
            {
                $properties = @("GroupMode","Group","Filter","FilterMode")
            
                $settingsObj = $documentedObj.Assignments | Where { $_.Settings -ne $null } | Select -First 1

                if($settingsObj)
                {
                    foreach($objProp in $settingsObj.Settings.Keys)
                    {
                        if($objProp -in $properties) { continue }
                        if($objProp -in @("Category","RawIntent")) { continue }
                        $properties += ("Settings." + $objProp)
                    }
                }
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
            }

            Add-MDTableItems $obj $objectType $documentedObj.Assignments $properties "TableHeaders.assignments" -AddCategories
        }
    }
    catch 
    {
        Write-LogError "Failed to process object $objName" $_.Exception
    }
    
    if($script:outputType -eq "Object")
    {
        $script:mdContent = [System.Text.StringBuilder]::new()
        $script:mdContent.AppendLine($script:cssStyle)        
        $mdText = $script:mdContent.ToString() 
        $mdText += $script:mdStrings.ToString()

        $fileName = "$($script:documentPath)\$($script:currentItemFileName)"
        Save-DocumentationFile $mdText $fileName
        $script:mdStrings = $null
    }    
}

function Add-MDTableItems
{
    param($obj, $objectType, $items, $properties, $lngId, [switch]$AddCategories, [switch]$AddSubcategories, $captionOverride)

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

    $tableText =  [System.Text.StringBuilder]::new()
    $tableText.AppendLine("<table class='table-settings'>")
    $tableText.AppendLine("<tr class='table-header1'>")

    $columnCount = 0
    foreach($prop in $properties)
    {
        $tableText.AppendLine("<td>$((Invoke-DocTranslateColumnHeader $prop.Split(".")[-1]))</td>")
        $columnCount++
    }
    $tableText.AppendLine("</tr>")

    $curCategory = ""
    $curSubCategory = ""

    $columnCategory = $null
    $columnSubCategory = $null

    foreach($itemObj in $items)
    {
        $additionalRowClass = ""
        if($itemObj.Category -and $curCategory -ne $itemObj.Category -and $AddCategories -eq $true)
        {
            $tableText.AppendLine("<tr>")
            $tableText.AppendLine("<td colspan=`"$($columnCount)`" class='category-level1'>$((Set-MDText $itemObj.Category))</td>")
            $tableText.AppendLine("</tr>")

            $curCategory = $itemObj.Category
            $curSubCategory = ""
            $curentPropertyIndex = 0
        }

        if($itemObj.SubCategory -and $curSubCategory -ne $itemObj.SubCategory -and $AddSubcategories -eq $true)
        {
            $tableText.AppendLine("<tr>")
            $tableText.AppendLine("<td colspan=`"$($columnCount)`" class='category-level2'>$((Set-MDText $itemObj.SubCategory))</td>")
            $tableText.AppendLine("</tr>")

            $curSubCategory = $itemObj.SubCategory
            $curentPropertyIndex = 0
        }
        
        if($itemObj.PropertyIndex -is [int] -and $itemObj.PropertyIndex -gt 0 -and $itemObj.PropertyIndex -eq 1)
        {
            $curentPropertyIndex = $itemObj.PropertyIndex
            $additionalRowClass = "row-new-property"
        }        

        try 
        {   
            $tableText.AppendLine("<tr class='$($additionalRowClass)'>")

            $curCol = 1
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

                    if($propName -eq "Value" -and ($itemObj.FullValueTable | measure).Count -gt 0)
                    {
                        $tableText.AppendLine("<td><table class='table-value'>")
                        $tableText.AppendLine("<tr>")
                        foreach($tableObjectProp in $itemObj.FullValueTable[0].PSObject.Properties)
                        {
                            $tableText.AppendLine("<td class='table-header1'>$($tableObjectProp.Name)</td>")
                        }
                        $tableText.AppendLine("</tr>")

                        foreach($tableValue in $itemObj.FullValueTable)
                        {
                            $tableText.AppendLine("<tr>")
                            foreach($tableObjectProp in $itemObj.FullValueTable[0].PSObject.Properties)
                            {
                                $tableText.AppendLine("<td>$($tableValue."$($tableObjectProp.Name)")</td>")
                            }
                            $tableText.AppendLine("</tr>")
                        }
                        $tableText.AppendLine("</table></td>")
                    }
                    else
                    {
                        $style = ""
                        if($curCol -eq 1 -and $itemObj.Level)
                        {
                            try
                            {
                                $level = [int]$itemObj.Level
                                $style = " style='padding-left:$((5 + ($level * 5)))px !important;'"
                            }
                            catch{}
                        }
                        $params = @{}
                        if($curCol -gt 0)
                        {
                            $params.Add("CodeBlock", $true)
                        }

                        $tableText.AppendLine("<td class='property-column$($curCol)'$style>$((Set-MDText $tmpObj.$propName @params))</td>")
                    }
                    
                    #$columnData += "$((Set-MDText "$($tmpObj.$propName)"))|"
                }
                catch
                {
                    #$columnData += "|"
                    Write-LogError "Failed to add property value for $prop" $_.Exception
                }
                $curCol++         
            }
                          
        }
        catch 
        {
            Write-Log "Failed to process property" 2    
        }
        finally 
        {
            $tableText.AppendLine("</tr>")
        }

        #Add-MDText $columnData 
    }

    $tableText.AppendLine("</table>")
    Add-MDText $tableText.ToString()
    
    Add-MDHeader $caption -Level 6 -TOT -AddParagraph
}

function Add-MDText
{
    param($text, [switch]$AddParagraph)

    $script:mdStrings.AppendLine($text)

    if($AddParagraph -eq $true)
    {
        # Add new paragraph by default
        $script:mdStrings.AppendLine("")
    }
}

function Set-MDText
{
    param([string]$text, [switch]$CodeBlock)

    if($null -eq $text) { return }

    $txtSummary = ""
    $textOut = ""

    if($text -and $text.Length -gt 250)
    {
        $summaryMax = 40
        # Show the first row or the first $max characters if first row is too short or too long
        $idx = $text.IndexOfAny(@("`r","`n"))
        if($idx -gt 10 -and $idx -lt 50)
        {
            $summaryMax = $idx
        }
        $txtSummary = $text.SubString(0,$summaryMax)
    }

    if($CodeBlock -eq $true)
    {
        $trimText = $text.Trim()
        if($trimText.StartsWith("<?xml") -or $trimText.StartsWith("<xml") -or ($trimText.StartsWith("<") -and $trimText.EndsWith(">")))
        {
            $textOut = ([Environment]::NewLine + [Environment]::NewLine + "``````xml" + [Environment]::NewLine + $text  + [Environment]::NewLine + "``````" + [Environment]::NewLine + [Environment]::NewLine)
        }
    }
    
    if($CodeBlock -eq $false -or -not $textOut)
    {
        $text = $text.Replace("|", '`|')
        $text = $text.Replace("*", '`*')
        $text = $text.Replace("$", '`$')
        $text = $text.Replace("`r`n", "<br />")
        $textOut = $text.Replace("`n", "<br />")
    }

    if($txtSummary)
    {
        "<details class='description'><summary data-open='Minimize' data-close='$($txtSummary)...expand'></summary>$textOut</details>"
    }
    else
    {
        $textOut
    }
}

function Add-MDHeader
{
    param($text, [int]$level = 1, [switch]$AddParagraph, [switch]$UseHTML, [switch]$ToT, [switch]$SkipTOC)

    if($script:mdStrings) 
    {
        $prefix = ""
        if($ToT -eq $true)
        {
            $prefix = "Table $(($script:totAnchors.Count + 1)). "
        }    

        if($UseHTML -eq $true)
        {
            if($ToT -eq $true)
            {
                $sectionAnchor = "table-$(($script:totAnchors.Count + 1))"            
            }
            else
            {
                $sectionAnchor = "section-$(($script:sectionAnchors.Count + 1))"
            }
            
            $script:mdStrings.AppendLine("<h$level id=`"$prefix$($sectionAnchor)`">$text</h$level>")
        }
        else 
        {
            # Warnig: Not complete! Use HTML if not working...
            $text = "$prefix$text"
            $sectionAnchor = $text.ToLower().Replace(" ","-").Replace("[","").Replace("]","")

            $mdHeader = [String]::new('#',$level)
            $script:mdStrings.AppendLine("$mdHeader $text")            
        }
        $FileName = $script:currentItemFileName
    }
    else
    {
        $sectionAnchor = $null
        $FileName = $null
    }
    
    if($ToT -eq $true)
    {
        $script:totAnchors += [PSCustomObject]@{
            Name = $text
            Anchor = $sectionAnchor
            FileName = $FileName
            Level = $level
        }
    }
    elseif($SkipTOC -ne $true)
    {
        $script:sectionAnchors += [PSCustomObject]@{
            Name = $text
            Anchor = $sectionAnchor
            FileName = $FileName
            Level = $level
        }
    }    

    if($AddParagraph -eq $true)
    {
        # Add new paragraph by default
        $script:mdStrings.AppendLine("`n")
    } 
}

function Add-MDObjectSettings 
{
    param($obj, $objectType, $documentedObj)
    
    foreach($objectScript in $documentedObj.Scripts)
    {
        if(-not $objectScript.ScriptContent -or -not $objectScript.Caption) { continue }

        $script:mdStrings.AppendLine("~~~powershell")
        $script:mdStrings.AppendLine($objectScript.ScriptContent)
        $script:mdStrings.AppendLine("~~~")
        Add-MDHeader $objectScript.Caption -Level 6 -SkipTOC -AddParagraph
    }
}