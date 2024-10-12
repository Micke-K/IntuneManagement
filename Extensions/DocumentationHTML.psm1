function Get-ModuleVersion
{
    '1.1.0'
}

function Invoke-InitializeModule
{
    Add-OutputType ([PSCustomObject]@{
        Name="HTML"
        Value="html"
        OutputOptions = (Add-HTMLOptionsControl)
        #Activate = { Invoke-HTMLActivate @args }
        PreProcess = { Invoke-HTMLPreProcessItems @args }
        NewObjectGroup = { Invoke-HTMLNewObjectGroup2 @args }
        NewObjectType = { Invoke-HTMLNewObjectType2 @args }
        Process = { Invoke-HTMLProcessItem @args }
        PostProcess = { Invoke-HTMLPostProcessItems @args }
        ProcessAllObjects = { Invoke-HTMLProcessAllObjects @args }
    })    
}

function Add-HTMLOptionsControl
{
    $script:htmlForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\DocumentationHTMLOptions.xaml") -AddVariables

    Set-XamlProperty $script:htmlForm "cbHTMLDocumentOutputFile" "ItemsSource" ("[ { Name: `"Single file`",Value: `"Full`" }, { Name: `"One file per object`",Value: `"Object`" }]" | ConvertFrom-Json)
    Set-XamlProperty $script:htmlForm "cbHTMLDocumentOutputFile" "SelectedValue" (Get-Setting "Documentation" "HTMLDocumentFileType" "Full")
    Set-XamlProperty $script:htmlForm "chkHTMLOpenDocument" "IsChecked" (Get-Setting "Documentation" "HTMLOpenFile" $true)

    Set-XamlProperty $script:htmlForm "txtHTMLDocumentName" "Text" (Get-Setting "Documentation" "HTMLDocumentName" "")
    Set-XamlProperty $script:htmlForm "txtHTMLCSSFile" "Text" (Get-Setting "Documentation" "HTMLCSSFile" "")
 
    Add-XamlEvent $script:htmlForm "browseHTMLDocumentName" "add_click" {
        $sf = [System.Windows.Forms.SaveFileDialog]::new()
        $sf.DefaultExt = "*.html"
        $sf.Filter = "HTML (*.html)|*.html|All files (*.*)|*.*"
        if($sf.ShowDialog() -eq "OK")
        {
            Set-XamlProperty $script:htmlForm "txtHTMLDocumentName" "Text" $sf.FileName
            Save-Setting "Documentation" "HTMLDocumentName" $sf.FileName
        }                
    }

    Add-XamlEvent $script:htmlForm "browseHTMLCSSFile" "add_click" {
        $of = [System.Windows.Forms.OpenFileDialog]::new()
        $of.Multiselect = $false
        $of.Filter = "CSS Files (*.css)|*.css|All files (*.*)|*.*"
        if($of.ShowDialog())
        {
            Set-XamlProperty $script:htmlForm "txtHTMLCSSFile" "Text" $of.FileName
            Save-Setting "Documentation" "HTMLCSSFile" $of.FileName
        }                
    }

    $script:htmlForm
}
function Invoke-HTMLPreProcessItems
{
    $script:sectionAnchors = @()
    $script:totAnchors = @()
    $script:htmlStrings = $null
    $script:currentItemFileName = $null
    
    Save-Setting "Documentation" "HTMLDocumentName" (Get-XamlProperty $script:htmlForm "txtHTMLDocumentName" "Text" "")
    Save-Setting "Documentation" "HTMLCSSFile" (Get-XamlProperty $script:htmlForm "txtHTMLCSSFile" "Text" "")
    Save-Setting "Documentation" "HTMLOpenFile" (Get-XamlProperty $script:htmlForm "chkHTMLOpenDocument" "IsChecked")
    Save-Setting "Documentation" "HTMLDocumentFileType" (Get-XamlProperty $script:htmlForm "cbHTMLDocumentOutputFile" "SelectedValue" '')

    $defaultCSSFile = $global:AppRootFolder + "\Documentation\DefaultHTMLStyle.css"
    $HTMLCssFile = Get-XamlProperty $script:htmlForm "txtHTMLCSSFile" "Text" $defaultCSSFile

    if(-not $HTMLCssFile)
    {
        Write-Log "CSS file not specified. Using default" 2
        $HTMLCssFile = $defaultCSSFile
    }
    elseif([IO.File]::Exists($HTMLCssFile) -eq -$false)
    {
        Write-Log "CSS file $($HTMLCssFile) not found. Using default" 2
        $HTMLCssFile = $defaultCSSFile
    }

    $cssStyle = ""
    if([IO.File]::Exists($HTMLCssFile))
    {
        Write-Log "Using CSS file $($HTMLCssFile)"
        $cssStyle = ((Get-Content -Raw -Path $HTMLCssFile) + [System.Environment]::NewLine)
    }
    else
    {
        Write-Log "CSS file $($HTMLCssFile) not found. No styles applied" 2
    }
    
    $script:cssStyle = $cssStyle
 
    $fileName = Expand-FileName (Get-XamlProperty $script:htmlForm "txtHTMLDocumentName" "Text" "%MyDocuments%\%Organization%-%Date%.html")

    $script:outFile = $fileName
    $script:documentPath = [io.path]::GetDirectoryName($fileName)

    $script:outputType = (Get-XamlProperty $script:htmlForm "cbHTMLDocumentOutputFile" "SelectedValue" "Full")

    if($script:outputType -eq "Object")
    {
        Write-Log "Document one file for each object + index file"
    }
    else
    {
        Write-Log "Document one single file for all objects"
        $script:outputType = "Full"
        $script:htmlStrings = [System.Text.StringBuilder]::new()        
    }
}

function Invoke-HTMLPostProcessItems
{
    $userName = $global:me.displayName
    if($global:me.givenName -and $global:me.surname)
    {
        $userName = ($global:me.givenName + " " + $global:me.surname)
    }

    $script:htmlContent = [System.Text.StringBuilder]::new()
    $script:htmlContent.AppendLine("<HTML>")

    $script:htmlContent.AppendLine($script:cssStyle)

    $script:htmlContent.AppendLine("<H1 class='header-level1'>$((?? $script:htmlDocTitle "Intune documentation"))</H1>")

    $mail = ""
    if($global:me.mail)
    {
        $mail = " ($($global:me.mail))"
    }

    $script:htmlContent.AppendLine("Organization: $($global:Organization.displayName)<br />")
    $script:htmlContent.AppendLine("Generated by: $userName$mail<br />")
    $script:htmlContent.AppendLine("Generated: $((Get-Date).ToShortDateString()) $((Get-Date).ToLongTimeString())<br />")
    
    if($script:sectionAnchors.Count -gt 0)
    {
        $script:htmlContent.AppendLine("<br />")
        $script:htmlContent.AppendLine("<H2 class='header-level2'>Table of Contents</H2>")
    }

    $tocMaxLevel = 4
    foreach($header in $script:sectionAnchors)
    {
        if($tocMaxLevel -gt 0 -and $header.Level -gt $tocMaxLevel)
        {
            continue
        }

        $script:htmlContent.AppendLine("<a href='$($header.FileName)#$($header.Anchor)' class='anchor-style anchor-level$($header.Level)'>$($header.Name)</a><br />")
    }
    
    if($script:sectionAnchors.Count -gt 0)
    {
        $script:htmlContent.AppendLine("<br />")
    }    
    
    $htmlText = $script:htmlContent.ToString() 
    if($script:outputType -eq "Full")
    {
        $htmlText += $script:htmlStrings.ToString()
    }
    $htmlText += "</HTML>"
    
    Save-DocumentationFile $htmlText $script:outFile -OpenFile:((Get-Setting "Documentation" "HTMLOpenFile" $true) -eq $true)
}

function Invoke-HTMLNewObjectGroup
{
    param($obj, $documentedObj)

    $script:objectHeaderLevel = 2

    $objectTypeString = Get-ObjectTypeString $obj.Object $obj.ObjectType

    Add-HTMLHeader (?? $objectTypeString $obj.ObjectType.Title)
}


function Invoke-HTMLNewObjectType
{
    param($obj, $documentedObj, [int]$groupCategoryCount = 0)

    $script:objectHeaderLevel = 3

    if($obj.ObjectType.GroupId -eq "EndpointSecurity")
    {
        $objectTypeString = $obj.CategoryName
    }
    else
    {
        $objectTypeString = $obj.ObjectType.Title
    }

    Add-HTMLHeader (?? $objectTypeString $obj.ObjectType.Title)

    $script:objectHeaderLevel = 4
}

function Invoke-HTMLNewObjectGroup2
{
    param($groupId)

    $script:objectHeaderLevel = 2

    $objectTypeString = Get-ObjectTypeString -ObjectType $groupId

    Add-HTMLHeader (?? $objectTypeString $obj.ObjectType.Title)
}

function Invoke-HTMLNewObjectType2
{
    param($objectTypeName)

    $script:objectHeaderLevel = 3

    Add-HTMLHeader $objectTypeName

    $script:objectHeaderLevel = 4
}

function Add-HTMLHeader
{
    param ($headerText, [int]$level = $script:objectHeaderLevel, [switch]$ToT, [switch]$SkipTOC)

    if($script:htmlStrings) 
    { 
        $prefix = ""
        if($ToT -eq $true)
        {
            $prefix = "Table $(($script:totAnchors.Count + 1)). "
        } 

        if($ToT -eq $true)
        {
            $sectionAnchor = "table-$(($script:totAnchors.Count + 1))"            
        }
        else
        {
            $sectionAnchor = "section-$(($script:sectionAnchors.Count + 1))"
        }

        $script:htmlStrings.AppendLine("<H$($level) id=`"$prefix$($sectionAnchor)`" class='header-level$($level)'>$headerText</H$($level)>")
        
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
            Name = $headerText
            Anchor = $sectionAnchor
            Level = $level
            FileName = $FileName
        }
    }
    elseif($SkipTOC -ne $true)
    {
        $script:sectionAnchors += [PSCustomObject]@{
            Name = $headerText
            Anchor = $sectionAnchor
            Level = $level
            FileName = $FileName
        }
    }     
}


function Invoke-HTMLProcessItem
{
    param($obj, $objectType, $documentedObj)

    if(!$documentedObj -or !$obj -or !$objectType) { return }

    $objName = Get-GraphObjectName $obj $objectType

    if($script:outputType -eq "Object")
    {
        $script:totAnchors = @()
        $script:htmlStrings = [System.Text.StringBuilder]::new()
        $script:currentItemFileName = (Remove-InvalidFileNameChars "$($objName).html")
    }

    Add-HTMLHeader $objName

    $script:htmlStrings.AppendLine("<br />")

    try 
    {
        foreach($tableType in @("BasicInfo","FilteredSettings"))
        {
            if($tableType -eq "BasicInfo")
            {
                $properties = @("Name","Value")
            }
            elseif($global:txtHTMLDocumentationProperties.SelectedValue -eq 'extended' -and $documentedObj.DisplayProperties)
            {
                $properties = @("Name","Value","Description")
            }
            elseif($global:txtHTMLDocumentationProperties.SelectedValue -eq 'custom' -and $global:txtHTMLCustomProperties.Text)
            {
                $properties = @()
                
                foreach($prop in $global:txtHTMLCustomProperties.Text.Split(","))
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
                Add-HTMLTableItems $obj $objectType ($documentedObj.$tableType) $properties $lngId -AddCategories -AddSubcategories
            }
        }

        if(($documentedObj.ComplianceActions | measure).Count -gt 0)
        {
            $properties = @("Action","Schedule","MessageTemplate","EmailCC")

            Add-HTMLTableItems $obj $objectType $documentedObj.ComplianceActions $properties "Category.complianceActionsLabel"
        }

        if(($documentedObj.ApplicabilityRules | measure).Count -gt 0)
        {
            $properties = @("Rule","Property","Value")

            Add-HTMLTableItems $obj $objectType $documentedObj.ApplicabilityRules $properties "SettingDetails.applicabilityRules"
        }

        Add-HTMLObjectSettings $obj $objectType $documentedObj

        foreach($customTable in ($documentedObj.CustomTables | Sort-Object -Property Order)) 
        {
            Add-HTMLTableItems $obj $objectType $customTable.Values $customTable.Columns $customTable.LanguageId -AddCategories -AddSubcategories
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

            Add-HTMLTableItems $obj $objectType $documentedObj.Assignments $properties "TableHeaders.assignments" -AddCategories
        }
    }
    catch 
    {
        Write-LogError "Failed to process object $objName" $_.Exception
    }
    
    if($script:outputType -eq "Object")
    {
        $script:htmlContent = [System.Text.StringBuilder]::new()
        $script:htmlContent.AppendLine("<HTML>")
        $script:htmlContent.AppendLine($script:cssStyle)        
        $htmlText = $script:htmlContent.ToString() 
        $htmlText += $script:htmlStrings.ToString()
        $htmlText += "</HTML>"

        $fileName = "$($script:documentPath)\$($script:currentItemFileName)"
        Save-DocumentationFile $htmlText $fileName
        $script:htmlStrings = $null
    }
}

function Add-HTMLTableItems
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
    # Add Header row
    $tableText.AppendLine("<tr>")
    $columnCount = 0    
    foreach($prop in $properties)
    {
        $tableText.AppendLine("<th>$((Invoke-DocTranslateColumnHeader $prop.Split(".")[-1]))</th>")
        $columnCount++
    }
    $tableText.AppendLine("</tr>")

    $curCategory = ""
    $curSubCategory = ""

    $columnCategory = $null
    $columnSubCategory = $null

    $row = 1

    foreach($itemObj in $items)
    {
        $newCategory = $false
        $newSubCategory = $false
        $additionalRowClass = ""
        if($itemObj.Category -and $curCategory -ne $itemObj.Category -and $AddCategories -eq $true)
        {
            # Add Category row
            $tableText.AppendLine("<tr>")
            $tableText.AppendLine("<td colspan=`"$($columnCount)`" class='category-level1'>$($itemObj.Category)</td>")
            $tableText.AppendLine("</tr>")

            $curCategory = $itemObj.Category
            $curSubCategory = ""
            $row = 1
            $newCategory = $true
            $curentPropertyIndex = 0
        }

        if($itemObj.SubCategory -and $curSubCategory -ne $itemObj.SubCategory -and $AddSubcategories -eq $true)
        {
            # Add Sub-category row
            $tableText.AppendLine("<tr>")
            $tableText.AppendLine("<td colspan=`"$($columnCount)`" class='category-level2'>$($itemObj.SubCategory)</td>")
            $tableText.AppendLine("</tr>")

            $curSubCategory = $itemObj.SubCategory
            $row = 1
            $newSubCategory = $true
            $curentPropertyIndex = 0
        }

        if($itemObj.PropertyIndex -is [int] -and $itemObj.PropertyIndex -gt 0 -and $itemObj.PropertyIndex -eq 1)
        {
            $curentPropertyIndex = $itemObj.PropertyIndex
            $additionalRowClass = "row-new-property"
        }

        try 
        {   
            if(($row % 2) -eq 1)
            {
                $rowClass = "row-odd"
            }
            else
            {
                $rowClass = "row-even"
            }

            $row++

            $tableText.AppendLine("<tr class='$($rowClass) $($additionalRowClass)'>")

            $curCol = 1
            foreach($prop in $properties)
            {
                try
                {
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
                            $tableText.AppendLine("<th>$($tableObjectProp.Name)</th>")
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
                                $style = " style='padding-left:$((5 + ($level * 5)))px;'"
                            }
                            catch{}
                        }
                        $tableText.AppendLine("<td class='property-column$($curCol)'$style>$((Set-HTMLText $tmpObj.$propName))</td>")
                    }
                }
                catch
                {
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
    }

    $tableText.AppendLine("</table>")
    $script:htmlStrings.Append($tableText.ToString())
    Add-HTMLHeader $caption -level 6
}

function Invoke-HTMLProcessAllObjects
{
    param($documentationInfo)

    $scopeTagInfo = Get-TableObjects "ScopeTags"

    if($allObjectTypeObjdocumentationInfoects.ObjectType.Id -eq "ScopeTags")
    {
        if(($documentationInfo.Items | measure).Count -gt 0)
        {
            Add-HTMLHeader $documentationInfo.TypeName

            Add-HTMLTableItems $null $documentationInfo.ObjectType $documentationInfo.Items -captionOverride (Get-LanguageString "SettingDetails.scopeTags")
        }
    }
}

function Set-HTMLText
{
    param([string]$text, [switch]$NoCodeBlock)

    if(-not $text)
    {
        return
    }

    $txtSummary = ""
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
    
    $code = $false
    if($NoCodeBlock -ne $true)
    {
        $trimText = $text.Trim()
        if($trimText.StartsWith("<") -and $trimText.EndsWith(">"))
        {
            $code = $true
            $text = "<pre class='code'>$($text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;'))</pre>"
            if($txtSummary)
            {
                $txtSummary = $txtSummary.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
            }
        }
    }

    if($code -eq $false)
    {
        $text = $text.Replace("`r`n", "<br />")
        $text = $text.Replace("`n", "<br />")
        $text = $text.Replace('&', '&amp;')
    }

    if($txtSummary)
    {
        "<details class='description'><summary data-open='Minimize' data-close='$($txtSummary)...expand'></summary>$text</details>"
    }
    else
    {
        $text
    }
}

function Add-HTMLObjectSettings
{
    param($obj, $objectType, $documentedObj) 

    foreach($objectScript in $documentedObj.Scripts)
    {
        if(-not $objectScript.ScriptContent -or -not $objectScript.Caption) { continue }
        $script:htmlStrings.AppendLine("<pre class='code'>")
        $script:htmlStrings.AppendLine($objectScript.ScriptContent)
        $script:htmlStrings.AppendLine("</pre>")
        Add-HTMLHeader $objectScript.Caption -Level 6 -SkipTOC
    }
}