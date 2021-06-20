<#
This moule extends the EnpointManager view with a Compare option.

This will compare an Intune object with an exported file.

The properties of the compared objects will be added to a DataGrid and the non-matching properties will be highlighted

Objects can be compared based on Properties or Documentatation info.

#>

function Get-ModuleVersion
{
    '1.0.1'
}

function Invoke-InitializeModule
{
    # Make sure we add the default Output types
    Add-OutputType
}

function Invoke-ShowMainWindow
{
    $button = [System.Windows.Controls.Button]::new()
    $button.Content = "Compare"
    $button.Name = "btnCompare"
    $button.MinWidth = 100
    $button.Margin = "0,0,5,0" 
    $button.IsEnabled = $false
    $button.ToolTip = "Compare object with exported file"
    $global:dgObjects.add_selectionChanged({
        Set-XamlProperty $global:dgObjects.Parent "btnCompare" "IsEnabled" (?: ($global:dgObjects.SelectedItem -eq $null) $false $true)
    })

    $button.Add_Click({ 
        Show-CompareForm $global:dgObjects.SelectedItem
    })    

    $global:spSubMenu.RegisterName($button.Name, $button)

    $global:spSubMenu.Children.Insert(0, $button)
}

function Show-CompareForm
{
    param($objInfo)

    $script:cmpForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\CompareForm.xaml") -AddVariables
    if(-not $script:cmpForm) { return }

    $script:copareSource = $objInfo

    $global:cbCompareType.ItemsSource = ("[ { Name: `"Property`",Value: `"property`" }, { Name: `"Documentation`",Value: `"doc`" }]" | ConvertFrom-Json)
    $global:cbCompareType.SelectedValue = (Get-Setting "Compare" "Type" "property")

    $global:txtIntuneObject.Text = (Get-GraphObjectName $objInfo.Object $objInfo.ObjectType)

    Add-XamlEvent $script:cmpForm "btnClose" "add_click" {
        $script:cmpForm = $null
        Show-ModalObject 
    }

    Add-XamlEvent $script:cmpForm "btnStartCompare" "add_click" {
        Write-Status "Compare objects"
        Save-Setting "Compare" "Type" $global:cbCompareType.SelectedValue
        $script:currentObjName = ""
        Invoke-CompareObjects
        Write-Status "" 
    }
    
    Add-XamlEvent $script:cmpForm "btnCompareSave" "add_click" {

        if(($global:dgCompareInfo.ItemsSource | measure).Count -eq 0) { return }

        $sf = [System.Windows.Forms.SaveFileDialog]::new()
        $sf.FileName = $script:currentObjName
        $sf.initialDirectory = (?: ($global:lastCompareFile -eq $null) (Get-Setting "" "LastUsedRoot") ([IO.FileInfo]$global:lastCompareFile).DirectoryName)
        $sf.DefaultExt = "*.csv"
        $sf.Filter = "CSV (*.csv)|*.csv|All files (*.*)| *.*"
        if($sf.ShowDialog() -eq "OK")
        {
            $csvInfo = $global:dgCompareInfo.ItemsSource | Select PropertyName,Object1Value,Object2Value,Category,SubCategory,Match | ConvertTo-Csv -NoTypeInformation
            $csvInfo | Out-File $sf.FileName -Force -Encoding UTF8
        }    
    }

    Add-XamlEvent $script:cmpForm "btnCompareCopy" "add_click" {
        
        $global:dgCompareInfo.ItemsSource | Select PropertyName,Object1Value,Object2Value,Category,SubCategory,Match | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
    }    

    Add-XamlEvent $script:cmpForm "browseCompareObject" "add_click" {
        $of = [System.Windows.Forms.OpenFileDialog]::new()
        $of.Multiselect = $false
        $of.Filter = "Json files (*.json)|*.json"
        $of.InitialDirectory = (?: ($global:lastCompareFile -eq $null) (Get-Setting "" "LastUsedRoot") ([IO.FileInfo]$global:lastCompareFile).DirectoryName)

        if($of.ShowDialog())
        {
            Set-XamlProperty $script:cmpForm "txtCompareFile" "Text" $of.FileName
            $global:lastCompareFile = $of.FileName
        }
    }

    #Add-XamlEvent $script:cmpForm "dgCompareInfo" "add_loaded" {

    #}

    Show-ModalForm "Compare Intune Objects" $script:cmpForm -HideButtons
}

function Invoke-CompareObjects
{
    if(-not $script:copareSource) { return }

    if(-not $global:txtCompareFile.Text)
    {
        [System.Windows.MessageBox]::Show("No file selected", "Comapre", "OK", "Error")
        return
    }
    elseif([IO.File]::Exists($global:txtCompareFile.Text) -eq $false)
    {
        [System.Windows.MessageBox]::Show("File '$($global:txtCompareFile.Text)' not found", "Comapre", "OK", "Error")
        return
    }

    try
    {
        if($script:copareSource.ObjectType.LoadObject)
        {
            $compareObj  = & $script:copareSource.ObjectType.LoadObject $global:txtCompareFile.Text
        }
        else
        {
            $compareObj = Get-Content $global:txtCompareFile.Text | ConvertFrom-Json 
        }
    }
    catch
    {
        [System.Windows.MessageBox]::Show("Failed to convert json file '$($global:txtCompareFile.Text)'", "Comapre", "OK", "Error")
        return
    }

    $obj = Get-GraphObject $script:copareSource.Object $script:copareSource.ObjectType

    $script:currentObjName = Get-GraphObjectName $script:copareSource.Object $script:copareSource.ObjectType

    if($obj.Object."@OData.Type" -ne $compareObj."@OData.Type")
    {
        if(([System.Windows.MessageBox]::Show("The object types does not match.`n`nDo you to compare the objects?", "Comapre", "YesNo", "Warning")) -eq "No")
        {
            return
        }
    }    
    
    $script:compareProperties = @()

    if($global:cbCompareType.SelectedValue -eq "property")
    {
        Compare-ObjectsBasedonProperty $obj.Object $compareObj $obj.ObjectType
    }
    elseif($global:cbCompareType.SelectedValue -eq "doc")
    {
        Compare-ObjectsBasedonDocumentation $obj $compareObj
    }
    $global:dgCompareInfo.ItemsSource = $script:compareProperties  
}
function Set-ColumnVisibility
{
    param($showCategory = $false, $showSubCategory = $false)

    $colTmp = $global:dgCompareInfo.Columns | Where { $_.Binding.Path.Path -eq "Category" }
    if($colTmp)
    {
        $colTmp.Visibility = (?: ($showCategory -eq $true) "Visible" "Collapsed")
    }

    $colTmp = $global:dgCompareInfo.Columns | Where { $_.Binding.Path.Path -eq "SubCategory" }
    if($colTmp)
    {
        $colTmp.Visibility = (?: ($showSubCategory -eq $true) "Visible" "Collapsed")
    }
}

function Add-CompareProperty
{
    param($name, $value1, $value2, $category, $subCategory, $match = $null)

    $value1 = if($value1 -eq $null) { "" } else { $value1.ToString().Trim("`"") }
    $value2 = if($value2 -eq $null) { "" } else {  $value2.ToString().Trim("`"") }
    if( ($value1 -eq $value2) -eq $false)
    {
        $dummy = 1
    }

    $script:compareProperties += [PSCustomObject]@{
        PropertyName = $name
        Object1Value = $value1 #if($value1 -ne $null) { $value1.ToString().Trim("`"") } else { "" }
        Object2Value = $value2 #if($value2 -ne $null) { $value2.ToString().Trim("`"") } else { "" }
        Category = $category
        SubCategory = $subCategory
        Match = ?? $match ($value1 -eq $value2)
    }
}

function Compare-ObjectsBasedonProperty
{
    param($obj1, $obj2, $objectType)

    Write-Status "Compare properties"

    Set-ColumnVisibility $false

    $coreProps = @((?? $objectType.NameProperty "displayName"), "Description", "Id", "createdDateTime", "lastModifiedDateTime", "version")
    $postProps = @("Advertisements")

    foreach ($propName in $coreProps)
    {
        if(-not ($obj1.PSObject.Properties | Where Name -eq $propName))
        {
            continue
        }
        $val1 = ($obj1.$propName | ConvertTo-Json -Depth 10)
        $val2 = ($obj2.$propName | ConvertTo-Json -Depth 10)
        Add-CompareProperty $propName $val1 $val2
    }    

    $addedProps = @()
    foreach ($propName in ($obj1.PSObject.Properties | Select Name).Name) 
    {
        if($propName -in $coreProps) { continue }
        if($propName -in $postProps) { continue }

        if($propName -like "*@OData*" -or $propName -like "#microsoft.graph*") { continue }

        $addedProps += $propName
        $val1 = ($obj1.$propName | ConvertTo-Json -Depth 10)
        $val2 = ($obj2.$propName | ConvertTo-Json -Depth 10)
        Add-CompareProperty $propName $val1 $val2
    }

    foreach ($propName in ($obj2.PSObject.Properties | Select Name).Name) 
    {
        if($propName -in $coreProps) { continue }
        if($propName -in $postProps) { continue }
        if($propName -in $addedProps) { continue }

        if($propName -like "*@OData*" -or $propName -like "#microsoft.graph*") { continue }

        $val1 = ($obj1.$propName | ConvertTo-Json -Depth 10)
        $val2 = ($obj2.$propName | ConvertTo-Json -Depth 10)
        Add-CompareProperty $propName $val1 $val2
    }    

    foreach ($propName in $postProps)
    {
        if(-not ($obj1.PSObject.Properties | Where Name -eq $propName))
        {
            continue
        }
        $val1 = ($obj1.$propName | ConvertTo-Json -Depth 10)
        $val2 = ($obj2.$propName | ConvertTo-Json -Depth 10)
        Add-CompareProperty $propName $val1 $val2
    }    
}

function Get-CompareCustomColumnsDoc
{
    param($objInfo)

    if($objInfo.Object.'@OData.Type' -eq "#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration")
    {
        Set-ColumnVisibility $true $true
    }
    else
    {
        Set-ColumnVisibility $true $false
    }
}

function Compare-ObjectsBasedonDocumentation
{
    param($obj1, $obj2)

    Get-CompareCustomColumnsDoc $obj

    # ToDo: set this based on configuration value
    $script:assignmentOutput = "simpleFullCompare"

    $docObj1 = Invoke-ObjectDocumentation $obj1
    
    $obj2 | Add-Member Noteproperty -Name "@CompareObject" -Value $true -Force      

    $docObj2 = Invoke-ObjectDocumentation ([PSCustomObject]@{
        Object = $obj2
        ObjectType = $obj1.ObjectType
    })

    $settingsValue = ?? $obj1.ObjectType.CompareValue "Value"

    foreach ($prop in $docObj1.BasicInfo)
    {
        $val1 = $prop.Value 
        $prop2 = $docObj2.BasicInfo | Where Name -eq $prop.Name
        $val2 = $prop2.Value 
        Add-CompareProperty $prop.Name $val1 $val2 $prop.Category
    }

    $addedProperties = @()

    if($docObj1.InputType -eq "Settings")
    {
        foreach ($prop in $docObj1.Settings)
        {
            if(($prop.SettingId + $prop.ParentSettingId) -in $addedProperties) { continue }

            $addedProperties += ($prop.SettingId + $prop.ParentSettingId)
            $val1 = $prop.Value 
            $prop2 = $docObj2.Settings | Where { $_.SettingId -eq $prop.SettingId -and $_.ParentSettingId -eq $prop.ParentSettingId }
            $val2 = $prop2.Value
            Add-CompareProperty $prop.Name $val1 $val2 $prop.Category

            # ToDo: fix lazy copy/past coding
            $children1 = $docObj1.Settings | Where ParentId -eq $prop.Id
            $children2 = $docObj2.Settings | Where ParentId -eq $prop2.Id
            
            # Add children defined on Object 1 property
            foreach ($childProp in $children1)
            {
                if(($childProp.SettingId + $childProp.ParentSettingId) -in $addedProperties) { continue }

                $addedProperties += ($childProp.SettingId + $childProp.ParentSettingId)
                $val1 = $childProp.Value 
                $prop2 = $docObj2.Settings | Where { $_.SettingId -eq $childProp.SettingId -and $_.ParentSettingId -eq $childProp.ParentSettingId }
                $val2 = $prop2.Value
                Add-CompareProperty $childProp.Name $val1 $val2 $prop.Category
            }
            
            # Add children defined only on Object 2 property e.g. Baseline Firewall profile was disable AFTER export.
            # This is to make sure all children are added under its parent and not last in the table
            foreach ($childProp in $children2)
            {
                if(($childProp.SettingId + $childProp.ParentSettingId) -in $addedProperties) { continue }

                $addedProperties += ($childProp.SettingId + $childProp.ParentSettingId)
                $val2 = $childProp.Value 
                $prop2 = $docObj1.Settings | Where { $_.SettingId -eq $childProp.SettingId -and $_.ParentSettingId -eq $childProp.ParentSettingId }
                $val1 = $prop2.Value
                Add-CompareProperty $childProp.Name $val1 $val2 $prop.Category
            }
        }
        
        # These objects are defined only on Object 2. They will be last in the table
        foreach ($prop in $docObj2.Settings)
        {
            if(($prop.SettingId + $prop.ParentSettingId) -in $addedProperties) { continue }

            $addedProperties += ($prop.SettingId + $prop.ParentSettingId)
            $val2 = $prop.Value    
            $prop2 = $docObj1.Settings | Where  { $_.SettingId -eq $prop.SettingId -and $_.ParentSettingId -eq $prop.ParentSettingId }
            $val1 = $prop2.Value   
            Add-CompareProperty $prop.Name $val1 $val2 $prop.Category
        }    
    }
    else
    {
        foreach ($prop in $docObj1.Settings)
        {
            if(($prop.EntityKey + $prop.Category + $prop.SubCategory) -in $addedProperties) { continue }

            $addedProperties += ($prop.EntityKey + $prop.Category + $prop.SubCategory)
            $val1 = $prop.$settingsValue 
            $prop2 = $docObj2.Settings | Where { $_.EntityKey -eq $prop.EntityKey -and $_.Category -eq $prop.Category -and $_.SubCategory -eq $prop.SubCategory }
            $val2 = $prop2.$settingsValue
            Add-CompareProperty $prop.Name $val1 $val2 $prop.Category $prop.SubCategory
        }
        
        # These objects are defined only on Object 2. They will be last in the table
        foreach ($prop in $docObj2.Settings)
        {
            if(($prop.EntityKey + $prop.Category + $prop.SubCategory) -in $addedProperties) { continue }

            $addedProperties += ($prop.EntityKey + $prop.Category + $prop.SubCategory)
            $val2 = $prop.$settingsValue
            $prop2 = $docObj1.Settings | Where  { $_.EntityKey -eq $prop.EntityKey -and $_.Category -eq $prop.Category -and $_.SubCategory -eq $prop.SubCategory }
            $val1 = $prop2.$settingsValue   
            Add-CompareProperty $prop.Name $val1 $val2 $prop.Category $prop.SubCategory
        }           
    }

    $applicabilityRulesAdded = @()
    #$properties = @("Rule","Property","Value")
    foreach($applicabilityRule in $docObj1.ApplicabilityRules)
    {
        $applicabilityRule2 = $docObj2.ApplicabilityRules | Where { $_.Id -eq $applicabilityRule.Id }
        $applicabilityRulesAdded += $applicabilityRule.Id
        $val1 = ($applicabilityRule.Rule + [environment]::NewLine + $applicabilityRule.Value)
        $val2 = ($applicabilityRule2.Rule + [environment]::NewLine + $applicabilityRule2.Value)

        Add-CompareProperty $applicabilityRule.Property $val1 $val2 $applicabilityRule.Category
    }

    foreach($applicabilityRule in $docObj2.ApplicabilityRules)
    {
        if(($applicabilityRule.Id) -in $applicabilityRulesAdded) { continue }
        $applicabilityRule2 = $docObj1.ApplicabilityRules | Where { $_.Id -eq $applicabilityRule.Id }
        $script:applicabilityRulesAdded += $applicabilityRule.Id
        $val2 = ($applicabilityRule.Rule + [environment]::NewLine + $applicabilityRule.Value)
        $val1 = ($applicabilityRule2.Rule + [environment]::NewLine + $applicabilityRule2.Value)

        Add-CompareProperty $applicabilityRule.Property $val1 $val2 $applicabilityRule.Category
    }    

    $script:assignmentStr = Get-LanguageString "TableHeaders.assignment"
    $script:groupsAdded = @()

    $assignmentType = $null
    $curType = $null

    foreach ($assignment in $docObj1.Assignments)
    {
        #if(-not $assignmentType)
        #{
        #    $assignmentType = (?: ($assignment.RawIntent -eq $null) "generic" "app") 
        #}
        
        $prevType = $null

        if($curType -ne $assignment.Category) 
        {
            if($curType) { $prevType = $curType}
            $curType = $assignment.Category
        }

        if($prevType)
        {
            # Add any additional missing intent in the same intent group
            foreach($tmpAssignment in $docObj2.Assignments | Where { $_.Category -eq $prevType })
            {
                Add-AssignmentInfo $docObj2 $docObj1 $tmpAssignment -ReversedValue
            }
        }
        Add-AssignmentInfo $docObj1 $docObj2 $assignment
    }

    # Add any missing assignments from Object 2
    foreach ($assignment in $docObj2.Assignments)
    {
        Add-AssignmentInfo $docObj2 $docObj1 $assignment -ReversedValue
    }
}

function Add-AssignmentInfo
{
    param($srcObj, $cmpObj, $assignment, [switch]$ReversedValue)
 
    if(($assignment.Group + $assignment.GroupMode + $assignment.RawIntent) -in $script:groupsAdded) { continue }

    $assignment2 = $cmpObj.Assignments | Where { $_.GroupMode -eq $assignment.GroupMode -and $_.Group -eq $assignment.Group -and $_.RawIntent -eq $assignment.RawIntent }
    $script:groupsAdded += ($assignment.Group + $assignment.GroupMode + $assignment.RawIntent)

    $match = $null    

    # To only show the group name
    if($script:assignmentOutput -eq "simple")
    {
        $val1 = $assignment.Group
        $val2 = $assignment2.Group
    }
    else
    {
        # Show full Assignment info
        # -Property @("Group","*") will generete error but will put the Group first and the rest of the properties after it. ErrorAction SilentlyContinue will ignore the error
        # Should be another way of doing this without generating an error. 
        $val1 = $assignment | Select -Property @("Group","*") -ExcludeProperty @("RawJsonValue","RawIntent","GroupMode","Category") -ErrorAction SilentlyContinue | ConvertTo-Json -Compress #$assignment.Group
        $val2 = $assignment2 | Select -Property @("Group","*") -ExcludeProperty @("RawJsonValue","RawIntent","GroupMode","Category") -ErrorAction SilentlyContinue | ConvertTo-Json -Compress #$assignment2.Group
        
        if($script:assignmentOutput -eq "simpleFullCompare")
        {
            # Full compare but show only the Group name. This could cause red for not matching even though the same group is used e.g. Filter is changed
            $match = ($val1 -eq $val2)
            $val1 = $assignment.Group
            $val2 = $assignment2.Group    
        }
    }

    if($ReversedValue -eq $true)
    {
        $tmpVal = $val1
        $val1 = $val2
        $val2 = $tmpVal
    }

    if($assignment.RawIntent)
    {
        Add-CompareProperty $assignment.Category $val1 $val2 -Category $assignment.GroupMode -match $match
    }
    else
    {
        Add-CompareProperty $assignmentStr $val1 $val2 -Category $assignment.GroupMode -match $match
    }
}