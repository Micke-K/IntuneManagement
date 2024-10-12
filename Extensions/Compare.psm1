<#
This moule extends the EnpointManager view with a Compare option.

This will compare an Intune object with an exported file.

The properties of the compared objects will be added to a DataGrid and the non-matching properties will be highlighted

Objects can be compared based on Properties or Documentatation info.

#>

function Get-ModuleVersion
{
    '1.2.0'
}

function Invoke-InitializeModule
{
    $global:comparisonTypes = $null
    $global:compareProviders = $null
    $script:CompareProviderOptionsCache = $null

    $script:defaultCompareProps = [Collections.Generic.List[String]]@('ObjectName', 'Id', 'Type', 'Category', 'SubCategory', 'Property', 'Value1', 'Value2', 'Match')
    
    # Make sure we add the default providers
    Add-CompareProvider
    Add-ComparisonTypes

    $script:saveType = @(
        [PSCustomObject]@{
            Name="One file for each object type"
            Value="objectType"
        },
        [PSCustomObject]@{
            Name="One file for all objects"
            Value="all"
        }
    )
}

function Add-CompareProvider
{
    param($compareProvider)

    if(-not $global:compareProviders)
    {
        $global:compareProviders = @()
    }

    if($global:compareProviders.Count -eq 0)
    {
        $global:compareProviders += [PSCustomObject]@{
            Name = "Exported Files with Intune Objects (Id)"
            Value = "export"
            ObjectCompare = { Compare-ObjectsBasedonProperty @args }
            BulkCompare = { Start-BulkCompareExportObjects @args }
            ProviderOptions = "CompareExportOptions"
            Activate = { Invoke-ActivateCompareWithExportObjects @args }
        }

        $global:compareProviders += [PSCustomObject]@{
            Name = "Intune Objects with Exported Files (Name)"
            Value = "IntuneWithExport"
            ObjectCompare = { Compare-ObjectsBasedonProperty @args }
            BulkCompare = { Start-BulkCompareExportIntuneToNamedExportedObjects @args }
            ProviderOptions = "CompareExportOptions"
            Activate = { Invoke-ActivateCompareWithExportObjects @args }
        }

        $global:compareProviders += [PSCustomObject]@{
            Name = "Named Objects in Intune"
            Value = "name"
            BulkCompare = { Start-BulkCompareNamedObjects @args }
            ProviderOptions = "CompareNamedOptions"
            Activate = { Invoke-ActivateCompareNamesObjects @args }
            RemoveProperties = @("Id")
        }        

        $global:compareProviders += [PSCustomObject]@{
            Name = "Files in Exported Folders"
            Value = "exportedFolders"
            ObjectCompare = { Compare-ObjectsBasedonProperty @args }
            BulkCompare = { Start-BulkCompareExportFolders @args }
            ProviderOptions = "CompareExportedFilesOptions"
            Activate = { Invoke-ActivateCompareExportedObjects @args }
        }

        $global:compareProviders += [PSCustomObject]@{
            Name = "Existing objects"
            Value = "existing"
            Compare = { Compare-ObjectsBasedonDocumentation @args }
        }
    }

    if(!$compareProvider) { return }

    $global:compareProviders += $compareProvider
}

function Add-ComparisonTypes
{
    param($comparisonType)

    if(-not $global:comparisonTypes)
    {
        $global:comparisonTypes = @()
    }

    if($global:comparisonTypes.Count -eq 0)
    {
        $global:comparisonTypes += [PSCustomObject]@{
            Name = "Property"
            Value = "property"
            Compare = { Compare-ObjectsBasedonProperty @args }
            RemoveProperties = @('Category','SubCategory')
        }

        $global:comparisonTypes += [PSCustomObject]@{
            Name = "Documentation"
            Value = "doc"
            Compare = { Compare-ObjectsBasedonDocumentation @args }
        }
    }

    if(!$comparisonType) { return }

    $global:comparisonTypes += $comparisonType
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

    $button.Add_Click({ 
        Show-CompareForm $global:dgObjects.SelectedItem
    })    

    $global:spSubMenu.RegisterName($button.Name, $button)

    $global:spSubMenu.Children.Insert(0, $button)
}

function Invoke-EMSelectedItemsChanged
{
    $hasSelectedItems = ($global:dgObjects.ItemsSource | Where IsSelected -eq $true) -or ($null -ne $global:dgObjects.SelectedItem)
    Set-XamlProperty $global:dgObjects.Parent "btnCompare" "IsEnabled" $hasSelectedItems
}

function Invoke-ViewActivated
{
    if($global:currentViewObject.ViewInfo.ID -ne "IntuneGraphAPI") { return }
    
    $tmp = $mnuMain.Items | Where Name -eq "EMBulk"
    if($tmp)
    {
        $tmp.AddChild(([System.Windows.Controls.Separator]::new())) | Out-Null
        $subItem = [System.Windows.Controls.MenuItem]::new()
        $subItem.Header = "_Compare"
        $subItem.Add_Click({Show-CompareBulkForm})
        $tmp.AddChild($subItem)
    }
}

function Show-CompareBulkForm
{
    $script:cmpForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\BulkCompare.xaml") -AddVariables
    if(-not $script:cmpForm) { return }

    $global:cbCompareProvider.ItemsSource = @($global:compareProviders)
    $global:cbCompareProvider.SelectedValue = (Get-Setting "Compare" "Provider" "export")

    $global:cbCompareSave.ItemsSource = @($script:saveType)
    $global:cbCompareSave.SelectedValue = (Get-Setting "Compare" "SaveType" "objectType")

    $global:cbCompareType.ItemsSource = $global:comparisonTypes | Where ShowOnBulk -ne $false 
    $global:cbCompareType.SelectedValue = (Get-Setting "Compare" "Type" "property")

    $global:cbCompareCSVDelimiter.ItemsSource = @("", ",",";","-","|")
    $global:cbCompareCSVDelimiter.SelectedValue = (Get-Setting "Compare" "Delimiter" ";")

    $global:chkSkipCompareBasicProperties.IsChecked = (Get-Setting "Compare" "SkipCompareBasicProperties") -eq "true"
    $global:chkSkipCompareAssignments.IsChecked = (Get-Setting "Compare" "SkipCompareAssignments") -eq "true"
    $global:chkSkipMissingSourcePolicies.IsChecked = (Get-Setting "Compare" "SkipMissingSourcePolicies") -eq "true"
    $global:chkSkipMissingDestinationPolicies.IsChecked = (Get-Setting "Compare" "SkipMissingDestinationPolicies") -eq "true"
    
    $script:compareObjects = @()
    foreach($objType in $global:lstMenuItems.ItemsSource)
    {
        if(-not $objType.Title) { continue }

        $script:compareObjects += New-Object PSObject -Property @{
            Title = $objType.Title
            Selected = $true
            ObjectType = $objType
        }
    }

    $column = Get-GridCheckboxColumn "Selected"
    $global:dgObjectsToCompare.Columns.Add($column)

    $column.Header.IsChecked = $true # All items are checked by default
    $column.Header.add_Click({
            foreach($item in $global:dgObjectsToCompare.ItemsSource)
            {
                $item.Selected = $this.IsChecked
            }
            $global:dgObjectsToCompare.Items.Refresh()
        }
    ) 

    # Add Object type column
    $binding = [System.Windows.Data.Binding]::new("Title")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Object type"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:dgObjectsToCompare.Columns.Add($column)

    $global:dgObjectsToCompare.ItemsSource = $script:compareObjects

    Add-XamlEvent $script:cmpForm "btnClose" "add_click" {
        $script:cmpForm = $null
        Show-ModalObject 
    }

    Add-XamlEvent $script:cmpForm "btnStartCompare" "add_click" {
        Write-Status "Compare objects"
        Save-Setting "Compare" "Provider" $global:cbCompareProvider.SelectedValue
        Save-Setting "Compare" "Type" $global:cbCompareType.SelectedValue
        Save-Setting "Compare" "Delimiter" $global:cbCompareCSVDelimiter.SelectedValue

        Save-Setting "Compare" "SkipCompareBasicProperties" $global:chkSkipCompareBasicProperties.IsChecked
        Save-Setting "Compare" "SkipCompareAssignments" $global:chkSkipCompareAssignments.IsChecked
        Save-Setting "Compare" "SkipMissingSourcePolicies" $global:chkSkipMissingSourcePolicies.IsChecked
        Save-Setting "Compare" "SkipMissingDestinationPolicies" $global:chkSkipMissingDestinationPolicies.IsChecked

        if($global:cbCompareProvider.SelectedItem.BulkCompare)
        {
            & $global:cbCompareProvider.SelectedItem.BulkCompare
        }
        Write-Status "" 
    }

    $global:cbCompareProvider.Add_SelectionChanged({        
        Set-CompareProviderOptions $this
    })
    
    Add-XamlEvent $script:cmpForm "btnExportSettingsForSilentCompare" "add_click" ({
        $sf = [System.Windows.Forms.SaveFileDialog]::new()
        $sf.FileName = "BulkCompare.json"
        $sf.DefaultExt = "*.json"
        $sf.Filter = "Json (*.json)|*.json|All files (*.*)|*.*"
        if($sf.ShowDialog() -eq "OK")
        {
            $tmp = [PSCustomObject]@{
                Name = "ObjectTypes"
                Type = "Custom"
                ObjectTypes = @()
            }
            foreach($ot in ($global:dgObjectsToCompare.ItemsSource | Where Selected -eq $true))
            {
                $tmp.ObjectTypes += $ot.ObjectType.Id
            }
            Export-GraphBatchSettings $sf.FileName $script:cmpForm "BulkCompare" @($tmp)
        }  
    })

    Set-CompareProviderOptions $global:cbCompareProvider

    Show-ModalForm "Bulk Compare Objects" $script:cmpForm -HideButtons
}

function Set-CompareProviderOptions
{
    param($control)

    $providerOptions = $null
    $firstTime = $false
    if($control.SelectedItem.ProviderOptions)
    {
        if($script:CompareProviderOptionsCache -isnot [Hashtable]) { $script:CompareProviderOptionsCache = @{} }
        if($script:CompareProviderOptionsCache.Keys -contains $control.SelectedValue)
        {
            $providerOptions = $script:CompareProviderOptionsCache[$control.SelectedValue]
        }
        else
        {
            $providerOptions = Get-XamlObject ($global:AppRootFolder + "\Xaml\$($control.SelectedItem.ProviderOptions).xaml") -AddVariables
            if($providerOptions)
            {
                $firstTime = $true            
                $script:CompareProviderOptionsCache.Add($control.SelectedValue, $providerOptions)
            }
            else
            {
                Write-Log "Failed to create options for $($control.SelectedItem.Name)" 3
            }
        }
        $global:ccContentProviderOptions.Content = $providerOptions
    }
    else
    {
        $global:ccContentProviderOptions.Content = $null 
    }
    $global:ccContentProviderOptions.Visibility = (?: ($global:ccContentProviderOptions.Content -eq $null) "Collapsed" "Visible")

    if($control.SelectedItem.Activate)
    {
        if($firstTime)
        {
            Write-Log "Initialize $($global:cbCompareProvider.SelectedItem.Name) provider options"
        }
    
        & $control.SelectedItem.Activate $providerOptions $firstTime
    }    
}

# Compare Intune object with exported folder

function Invoke-ActivateCompareWithExportObjects
{
    param($providerOptions, $firstTime)

    if($firstTime)
    {
        $path = Get-Setting "" "LastUsedFullPath"
        if($path) 
        {
            $path = [IO.Directory]::GetParent($path).FullName
        }        
        Set-XamlProperty $providerOptions "txtExportPath" "Text" (?? $path (Get-SettingValue "RootFolder"))

        Add-XamlEvent $providerOptions "browseExportPath" "add_click" ({
            $folder = Get-Folder (Get-XamlProperty $this.Parent "txtExportPath" "Text") "Select root folder for compare"
            if($folder)
            {
                Set-XamlProperty $this.Parent "txtExportPath" "Text" $folder
            }
        })
    }
}

# Compare two exported folders
function Invoke-ActivateCompareExportedObjects
{
    param($providerOptions, $firstTime)

    if($firstTime)
    {
        $path = Get-Setting "" "LastUsedFullPath"
        if($path) 
        {
            $path = [IO.Directory]::GetParent($path).FullName
        }        
        Set-XamlProperty $providerOptions "txtExportPathSource" "Text" (?? $path (Get-SettingValue "RootFolder"))
        Set-XamlProperty $providerOptions "txtExportPathCompare" "Text" (Get-SettingValue "ExportPathCompare")

        Add-XamlEvent $providerOptions "browseExportPathSource" "add_click" ({
            $folder = Get-Folder (Get-XamlProperty $this.Parent "txtExportPathSource" "Text") "Select root folder for source"
            if($folder)
            {
                Set-XamlProperty $this.Parent "txtExportPathSource" "Text" $folder
            }
        })

        Add-XamlEvent $providerOptions "browseExportPathCompare" "add_click" ({
            $folder = Get-Folder (Get-XamlProperty $this.Parent "txtExportPathCompare" "Text") "Select folder to compare the source with"
            if($folder)
            {
                Set-XamlProperty $this.Parent "txtExportPathCompare" "Text" $folder
            }
        })        
    }
}

function Invoke-ActivateCompareNamesObjects
{
    param($providerOptions, $firstTime)

    if($providerOptions -and $firstTime)
    {
        Set-XamlProperty $providerOptions "txtCompareSource" "Text" (Get-Setting "Compare" "CompareSource" "")
        Set-XamlProperty $providerOptions "txtCompareWith" "Text" (Get-Setting "Compare" "CompareWith" "")

        Set-XamlProperty $providerOptions "txtSavePath" "Text" (Get-Setting "Compare" "SavePath" "")
        Add-XamlEvent $providerOptions "browseSavePath" "add_click" ({
            $folder = Get-Folder (Get-XamlProperty $this.Parent "txtSavePath" "Text") "Select folder"
            if($folder)
            {
                Set-XamlProperty $this.Parent "txtSavePath" "Text" $folder
            }
        })
    }    
}

function Start-BulkCompareNamedObjects
{
    Write-Log "****************************************************************"
    Write-Log "Start bulk Named Objects compare"
    Write-Log "****************************************************************"
    $compareObjectsResult = @()

    $compareSource = (Get-XamlProperty $global:ccContentProviderOptions.Content "txtCompareSource" "Text")
    $compareWith = (Get-XamlProperty $global:ccContentProviderOptions.Content "txtCompareWith" "Text")    

    if(-not $compareSource -or -not $compareWith)
    {
        [System.Windows.MessageBox]::Show("Both source and compare name patterns must be specified", "Error", "OK", "Error")
        return
    }

    Save-Setting "Compare" "CompareSource" $compareSource
    Save-Setting "Compare" "CompareWith" $compareWith

    Invoke-BulkCompareNamedObjects $compareSource $compareWith

    Write-Log "****************************************************************"
    Write-Log "Bulk compare Named Objects finished"
    Write-Log "****************************************************************"
    Write-Status ""
}

function Invoke-BulkCompareNamedObjects
{
    param($sourcePattern, $comparePattern)

    $outputType = $global:cbCompareSave.SelectedValue

    Save-Setting "Compare" "SaveType" $outputType
    
    $compResultValues = @()
    $compareObjectsResult = @()
    
    $outputFolder = (Get-XamlProperty $global:ccContentProviderOptions.Content "txtSavePath" "Text")
    if(-not $outputFolder)
    {
        $outputFolder = Expand-FileName "%MyDocuments%"
    }

    $compareProps = $script:defaultCompareProps
    
    foreach($removeProp in $global:cbCompareProvider.SelectedItem.RemoveProperties)
    {
        $compareProps.Remove($removeProp) | Out-Null
    }

    foreach($removeProp in $global:cbCompareType.SelectedItem.RemoveProperties)
    {
        $compareProps.Remove($removeProp) | Out-Null
    }

    foreach($item in ($global:dgObjectsToCompare.ItemsSource | where Selected -eq $true))
    { 
        Write-Status "Compare $($item.ObjectType.Title) objects" -Force -SkipLog
        Write-Log "----------------------------------------------------------------"
        Write-Log "Compare $($item.ObjectType.Title) objects"
        Write-Log "----------------------------------------------------------------"
    
        [array]$graphObjects = Get-GraphObjects -property $item.ObjectType.ViewProperties -objectType $item.ObjectType

        $nameProp = ?? $item.ObjectType.NameProperty "displayName"

        foreach($graphObj in ($graphObjects | Where { $_.Object."$($nameProp)" -imatch [regex]::Escape($sourcePattern) }))
        {
            $sourceName = $graphObj.Object."$($nameProp)"
            $compareName  = $sourceName -ireplace [regex]::Escape($sourcePattern),$comparePattern

            $compareObj = $graphObjects | Where { $_.Object."$($nameProp)" -eq $compareName -and $_.Object.'@OData.Type' -eq $graphObj.Object.'@OData.Type' }
        
            if(($compareObj | measure).Count -gt 1)
            {
                Write-Log "Multiple objects found with name $compareName. Compare will not be performed" 2
                continue
            }
            elseif($compareObj)
            {
                $sourceObj = Get-GraphObject $graphObj.Object $graphObj.ObjectType 
                $compareObj = Get-GraphObject $compareObj.Object $compareObj.ObjectType 
                $compareProperties = Compare-Objects $sourceObj.Object $compareObj.Object $sourceObj.ObjectType                
            }
            else
            {
                $sourceObj = Get-GraphObject $graphObj.Object $graphObj.ObjectType 
                # Add objects that are exported but deleted/not imported etc.
                Write-Log "Object '$((Get-GraphObjectName $graphObj.Object $graphObj.ObjectType))' with id $($graphObj.Object.Id) has no matching object with the compare pattern" 2
                $compareProperties = @([PSCustomObject]@{
                        Object1Value = (Get-GraphObjectName $graphObj.Object $graphObj.ObjectType)
                        Object2Value = $null
                        Match = $false
                    })
            }

            $compareObjectsResult += [PSCustomObject]@{
                Object1 = $sourceObj.Object
                Object2 = $compareObj.Object
                ObjectType = $item.ObjectType
                Id = $sourceObj.Object.Id
                Result = $compareProperties
            }          
        }

        if($outputType -eq "objectType")
        {
            $compResultValues = @()
        }

        foreach($compObj in @($compareObjectsResult | Where { $_.ObjectType.Id -eq $item.ObjectType.Id }))
        {
            $objName = Get-GraphObjectName (?? $compObj.Object1 $compObj.Object2) $item.ObjectType
            foreach($compValue in $compObj.Result)
            {
                $compResultValues += [PSCustomObject]@{
                    ObjectName = $objName
                    Id = $compObj.Id
                    Type = $compObj.ObjectType.Title
                    ODataType = $compObj.Object1.'@OData.Type'
                    Property = $compValue.PropertyName
                    Value1 = $compValue.Object1Value
                    Value2 = $compValue.Object2Value
                    Category = $compValue.Category
                    SubCategory = $compValue.SubCategory
                    Match = $compValue.Match
                }
            }
        }

        if($outputType -eq "objectType")
        {
            $fileName = Remove-InvalidFileNameChars (Expand-FileName "Compare-$($graphObj.ObjectType.Id)-$sourcePattern-$comparePattern-%DateTime%.csv")
            Save-BulkCompareResults $compResultValues (Join-Path $outputFolder $fileName) $compareProps
        }        
    }
    #$fileName = Expand-FileName $fileName

    if($compareObjectsResult.Count -eq 0)
    {
        [System.Windows.MessageBox]::Show("No objects were comparced. Verify name patterns", "Error", "OK", "Error")
    }
    elseif($outputType -eq "all")
    {
        $fileName = Remove-InvalidFileNameChars (Expand-FileName "Compare-$sourcePattern-$comparePattern-%DateTime%.csv")
        Save-BulkCompareResults $compResultValues (Join-Path $outputFolder $fileName) $compareProps
    }       
}

function Start-BulkCompareExportObjects
{
    Write-Log "****************************************************************"
    Write-Log "Start bulk Exported Objects compare"
    Write-Log "****************************************************************"
    $compareObjectsResult = @()

    $txtNameFilter = (Get-XamlProperty $global:ccContentProviderOptions.Content "txtCompareNameFilter" "Text")
    if($txtNameFilter -is [String])
    {
        $txtNameFilter = $txtNameFilter.Trim()
    }
    $rootFolder = (Get-XamlProperty $global:ccContentProviderOptions.Content "txtExportPath" "Text")
    
    $compareProps = $script:defaultCompareProps
    
    foreach($removeProp in $global:cbCompareProvider.SelectedItem.RemoveProperties)
    {
        $compareProps.Remove($removeProp) | Out-Null
    }

    foreach($removeProp in $global:cbCompareType.SelectedItem.RemoveProperties)
    {
        $compareProps.Remove($removeProp) | Out-Null
    }

    if(-not $rootFolder)
    {
        [System.Windows.MessageBox]::Show("Root folder must be specified", "Error", "OK", "Error")
        return
    }

    if([IO.Directory]::Exists($rootFolder) -eq $false)
    {
        [System.Windows.MessageBox]::Show("Root folder $rootFolder does not exist", "Error", "OK", "Error")
        return
    }
    
    $outputType = $global:cbCompareSave.SelectedValue
    Save-Setting "Compare" "SaveType" $outputType

    $compResultValues = @()

    foreach($item in ($global:dgObjectsToCompare.ItemsSource | where Selected -eq $true))
    { 
        Write-Status "Compare $($item.ObjectType.Title) objects" -Force -SkipLog
        Write-Log "----------------------------------------------------------------"
        Write-Log "Compare $($item.ObjectType.Title) objects"
        Write-Log "----------------------------------------------------------------"

        $folder = Join-Path $rootFolder $item.ObjectType.Id
        
        if([IO.Directory]::Exists($folder))
        {
            Save-Setting "" "LastUsedFullPath" $folder
        
            [array]$graphObjects = Get-GraphObjects -property $item.ObjectType.ViewProperties -objectType $item.ObjectType

            foreach ($fileObj in @(Get-GraphFileObjects $folder -ObjectType $item.ObjectType))
            {                
                if(-not $fileObj.Object.Id)
                {
                    Write-Log "Object from file '$($fileObj.FullName)' has no Id property. Compare not supported" 2
                    continue
                }

                $objName = Get-GraphObjectName $fileObj.Object $fileObj.ObjectType

                if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                {
                    continue
                }
                                
                $curObject = $graphObjects | Where { $_.Object.Id -eq $fileObj.Object.Id }

                if(-not $curObject)
                {
                    if($global:chkSkipMissingDestinationPolicies.IsChecked -ne $true) {
                        # Add objects that are exported but deleted
                        Write-Log "Object '$($objName)' with id $($fileObj.Object.Id) not found in Intune. Deleted?" 2
                        $compareProperties = @([PSCustomObject]@{
                                Object1Value = $null
                                Object2Value = (Get-GraphObjectName $fileObj.Object $item.ObjectType)
                                Match = $false
                            })
                    }
                }
                else
                {
                    $sourceObj = Get-GraphObject $curObject.Object $curObject.ObjectType
                    $fileObj.Object | Add-Member Noteproperty -Name "@ObjectFromFile" -Value $true -Force
                    $fileObj.Object | Add-Member Noteproperty -Name "@ObjectFileName" -Value $fileObj.FileInfo.FullName -Force
                    $compareProperties = Compare-Objects $sourceObj.Object $fileObj.Object $item.ObjectType                    
                }

                $compareObjectsResult += [PSCustomObject]@{
                    Object1 = $curObject.Object
                    Object2 = $fileObj.Object
                    ObjectType = $item.ObjectType
                    Id = $fileObj.Object.Id
                    Result = $compareProperties
                }                
            }

            if($global:chkSkipMissingSourcePolicies.IsChecked -ne $true) {
                foreach($graphObj in $graphObjects)
                {
                    # Add objects that are in Intune but not exported
                    if(($compareObjectsResult | Where { $_.Id -eq $graphObj.Id})) { continue }

                    $objName = Get-GraphObjectName $graphObj.Object $item.ObjectType
                    if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                    {
                        continue
                    }                 

                    $compareObjectsResult += [PSCustomObject]@{
                        Object1 = $curObject.Object
                        Object2 = $null
                        ObjectType = $item.ObjectType
                        Id = $graphObj.Id
                        Result = @([PSCustomObject]@{
                            Object1Value = $objName
                            Object2Value = $null
                            Match = $false
                        })
                    }
                }
            }

            if($outputType -eq "objectType")
            {
                $compResultValues = @()
            }

            foreach($compObj in @($compareObjectsResult | Where { $_.ObjectType.Id -eq $item.ObjectType.Id }))
            {
                $objName = Get-GraphObjectName (?? $compObj.Object1 $compObj.Object2) $item.ObjectType
                foreach($compValue in $compObj.Result)
                {
                    $compResultValues += [PSCustomObject]@{
                        ObjectName = $objName
                        Id = $compObj.Id
                        Type = $compObj.ObjectType.Title
                        ODataType = $compObj.Object1.'@OData.Type'
                        Property = $compValue.PropertyName
                        Value1 = $compValue.Object1Value
                        Value2 = $compValue.Object2Value
                        Category = $compValue.Category
                        SubCategory = $compValue.SubCategory
                        Match = $compValue.Match
                    }
                }
            }

            if($outputType -eq "objectType")
            {
                Save-BulkCompareResults $compResultValues (Join-Path $folder "Compare_$(((Get-Date).ToString("yyyyMMdd-HHmm"))).csv") $compareProps
            }
        }
        else
        {
            Write-Log "Folder $folder not found. Skipping compare" 2    
        }
    }

    if($outputType -eq "all" -and $compResultValues.Count -gt 0)
    {
        Save-BulkCompareResults $compResultValues (Join-Path $rootFolder "Compare_$(((Get-Date).ToString("yyyyMMDD-HHmm"))).csv") $compareProps
    }    

    Write-Log "****************************************************************"
    Write-Log "Bulk compare Exported Objects finished"
    Write-Log "****************************************************************"
    Write-Status ""
    if($compareObjectsResult.Count -eq 0)
    {
        [System.Windows.MessageBox]::Show("No objects were comparced. Verify folder and exported files", "Error", "OK", "Error")
    }
}

function Start-BulkCompareExportIntuneToNamedExportedObjects
{
    Write-Log "****************************************************************"
    Write-Log "Start bulk compare Intune with Exported Objects compare"
    Write-Log "****************************************************************"
    $compareObjectsResult = @()

    $txtNameFilter = (Get-XamlProperty $global:ccContentProviderOptions.Content "txtCompareNameFilter" "Text")
    if($txtNameFilter -is [String])
    {
        $txtNameFilter = $txtNameFilter.Trim()
    }
    $rootFolder = (Get-XamlProperty $global:ccContentProviderOptions.Content "txtExportPath" "Text")
    
    $compareProps = $script:defaultCompareProps
    
    foreach($removeProp in $global:cbCompareProvider.SelectedItem.RemoveProperties)
    {
        $compareProps.Remove($removeProp) | Out-Null
    }

    foreach($removeProp in $global:cbCompareType.SelectedItem.RemoveProperties)
    {
        $compareProps.Remove($removeProp) | Out-Null
    }

    if(-not $rootFolder)
    {
        [System.Windows.MessageBox]::Show("Root folder must be specified", "Error", "OK", "Error")
        return
    }

    if([IO.Directory]::Exists($rootFolder) -eq $false)
    {
        [System.Windows.MessageBox]::Show("Root folder $rootFolder does not exist", "Error", "OK", "Error")
        return
    }
    
    $outputType = $global:cbCompareSave.SelectedValue
    Save-Setting "Compare" "SaveType" $outputType

    $compResultValues = @()

    foreach($item in ($global:dgObjectsToCompare.ItemsSource | where Selected -eq $true))
    { 
        Write-Status "Compare $($item.ObjectType.Title) objects" -Force -SkipLog
        Write-Log "----------------------------------------------------------------"
        Write-Log "Compare $($item.ObjectType.Title) objects"
        Write-Log "----------------------------------------------------------------"

        $folder = Join-Path $rootFolder $item.ObjectType.Id
        
        if([IO.Directory]::Exists($folder))
        {
            Save-Setting "" "LastUsedFullPath" $folder
        
            [array]$graphObjects = Get-GraphObjects -property $item.ObjectType.ViewProperties -objectType $item.ObjectType

            $fileObjects = @(Get-GraphFileObjects $folder -ObjectType $item.ObjectType)

            foreach ($graphObject in @($graphObjects))
            {                
                $objName = Get-GraphObjectName $graphObject.Object $graphObject.ObjectType

                if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                {
                    continue
                }
                                
                $fileObj = $fileObjects | Where { (Get-GraphObjectName $_.Object $_.ObjectType) -eq $objName }

                if(-not $fileObj)
                {
                    # Add objects that are exported but deleted
                    if($global:chkSkipMissingDestinationPolicies.IsChecked -ne $true) {
                        Write-Log "Object '$($objName)' with id $($fileObj.Object.Id) not found in exported folder. New Object?" 2
                        $compareProperties = @([PSCustomObject]@{
                                Object1Value = $objName
                                Object2Value = $null
                                Match = $false
                            })
                    }
                }
                elseif(($fileObj | measure).Count -gt 1)
                {
                    # Add objects that are exported but deleted
                    Write-Log "Multiple exported objects found with name '$($objName)" 2
                    $compareProperties = @([PSCustomObject]@{
                            Object1Value = $objName
                            Object2Value = $null
                            Match = $false
                        })
                }
                else
                {
                    $sourceObj = Get-GraphObject $graphObject.Object $graphObject.ObjectType
                    $fileObj.Object | Add-Member Noteproperty -Name "@ObjectFromFile" -Value $true -Force
                    $fileObj.Object | Add-Member Noteproperty -Name "@ObjectFileName" -Value $fileObj.FileInfo.FullName -Force
                    $compareProperties = Compare-Objects $sourceObj.Object $fileObj.Object $item.ObjectType
                }

                $compareObjectsResult += [PSCustomObject]@{
                    Object1 = $graphObject.Object
                    Object2 = $fileObj.Object
                    ObjectType = $item.ObjectType
                    Id = $graphObject.Object.Id
                    Result = $compareProperties
                }                
            }

            if($global:chkSkipMissingSourcePolicies.IsChecked -ne $true) {
                foreach ($fileObj in @($fileObjects))
                {
                    # Add objects that are exported but not in Intune
                    if(($compareObjectsResult | Where { $_.FileInfo.FullName -eq $fileObj.FileInfo.FullName})) { continue }

                    $objName = Get-GraphObjectName $fileObj.Object $item.ObjectType
                    if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                    {
                        continue
                    }                 

                    $compareObjectsResult += [PSCustomObject]@{
                        Object1 = $null
                        Object2 = $fileObj.Object
                        ObjectType = $item.ObjectType
                        Id = $fileObj.Object.Id
                        Result = @([PSCustomObject]@{
                            Object1Value = $objName
                            Object2Value = $null
                            Match = $false
                        })
                    }
                }
            }

            if($outputType -eq "objectType")
            {
                $compResultValues = @()
            }

            foreach($compObj in @($compareObjectsResult | Where { $_.ObjectType.Id -eq $item.ObjectType.Id }))
            {
                $objName = Get-GraphObjectName $item.Object $item.ObjectType
                foreach($compValue in $compObj.Result)
                {
                    $compResultValues += [PSCustomObject]@{
                        ObjectName = $objName
                        Id = $compObj.Id
                        Type = $compObj.ObjectType.Title
                        ODataType = $compObj.Object1.'@OData.Type'
                        Property = $compValue.PropertyName
                        Value1 = $compValue.Object1Value
                        Value2 = $compValue.Object2Value
                        Category = $compValue.Category
                        SubCategory = $compValue.SubCategory
                        Match = $compValue.Match
                    }
                }
            }

            if($outputType -eq "objectType")
            {
                Save-BulkCompareResults $compResultValues (Join-Path $folder "Compare_$(((Get-Date).ToString("yyyyMMdd-HHmm"))).csv") $compareProps
            }
        }
        else
        {
            Write-Log "Folder $folder not found. Skipping compare" 2    
        }
    }

    if($outputType -eq "all" -and $compResultValues.Count -gt 0)
    {
        Save-BulkCompareResults $compResultValues (Join-Path $rootFolder "Compare_$(((Get-Date).ToString("yyyyMMDD-HHmm"))).csv") $compareProps
    }    

    Write-Log "****************************************************************"
    Write-Log "Bulk compare Intune with Exported Objects finished"
    Write-Log "****************************************************************"
    Write-Status ""
    if($compareObjectsResult.Count -eq 0)
    {
        [System.Windows.MessageBox]::Show("No objects were comparced. Verify folder and exported files", "Error", "OK", "Error")
    }
}


function Start-BulkCompareExportFolders
{
    Write-Log "****************************************************************"
    Write-Log "Start bulk Exported Folders compare"
    Write-Log "****************************************************************"
    $compareObjectsResult = @()

    $txtNameFilter = (Get-XamlProperty $global:ccContentProviderOptions.Content "txtCompareNameFilter" "Text").Trim()
    $rootFolderSource = (Get-XamlProperty $global:ccContentProviderOptions.Content "txtExportPathSource" "Text")
    $rootFolderCompare = (Get-XamlProperty $global:ccContentProviderOptions.Content "txtExportPathCompare" "Text")
    
    $compareProps = $script:defaultCompareProps
    
    foreach($removeProp in $global:cbCompareProvider.SelectedItem.RemoveProperties)
    {
        $compareProps.Remove($removeProp) | Out-Null
    }

    foreach($removeProp in $global:cbCompareType.SelectedItem.RemoveProperties)
    {
        $compareProps.Remove($removeProp) | Out-Null
    }

    if(-not $rootFolderSource -or -not $rootFolderCompare)
    {
        [System.Windows.MessageBox]::Show("Both folders must be specified", "Error", "OK", "Error")
        return
    }

    if([IO.Directory]::Exists($rootFolderSource) -eq $false)
    {
        [System.Windows.MessageBox]::Show("Root folder $rootFolderSource does not exist", "Error", "OK", "Error")
        return
    }

    if([IO.Directory]::Exists($rootFolderCompare) -eq $false)
    {
        [System.Windows.MessageBox]::Show("Root folder $rootFolderCompare does not exist", "Error", "OK", "Error")
        return
    } 
    
    $outputType = $global:cbCompareSave.SelectedValue
    Save-Setting "Compare" "SaveType" $outputType

    $compResultValues = @()

    foreach($item in ($global:dgObjectsToCompare.ItemsSource | where Selected -eq $true))
    { 
        Write-Status "Compare $($item.ObjectType.Title) objects" -Force -SkipLog
        Write-Log "----------------------------------------------------------------"
        Write-Log "Compare $($item.ObjectType.Title) objects"
        Write-Log "----------------------------------------------------------------"

        $folderSource = Join-Path $rootFolderSource $item.ObjectType.Id
        $folderCompare = Join-Path $rootFolderCompare $item.ObjectType.Id

        if([IO.Directory]::Exists($folderSource))
        {
            Save-Setting "" "LastUsedFullPath" $folderSource
        
            $fileCompareObjs = @(Get-GraphFileObjects $folderCompare -ObjectType $item.ObjectType)      

            foreach ($fileSourceObj in @(Get-GraphFileObjects $folderSource -ObjectType $item.ObjectType))
            {
                $objName = Get-GraphObjectName $fileSourceObj.Object $item.ObjectType
                if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                {
                    continue
                }

                if(-not $fileSourceObj.Object.Id)
                {
                    Write-Log "Object from file '$($fileSourceObj.FullName)' has no Id property. Compare not supported" 2
                    continue
                }

                $compareObject = $fileCompareObjs | Where { $_.Object.Id -eq $fileSourceObj.Object.Id }

                if(-not $compareObject)
                {
                    if($global:chkSkipMissingDestinationPolicies.IsChecked -ne $true) {
                        # Add objects that are exported but deleted
                        Write-Log "Object '$($objName)' with id $($fileSourceObj.Object.Id) not found in Intune. Deleted?" 2
                        $compareProperties = @([PSCustomObject]@{
                                Object1Value = $null
                                Object2Value = (Get-GraphObjectName $fileSourceObj.Object $fileSourceObj.ObjectType)
                                Match = $false
                            })
                    }
                }
                else
                {
                    $fileSourceObj.Object | Add-Member Noteproperty -Name "@ObjectFromFile" -Value $true -Force
                    $fileSourceObj.Object | Add-Member Noteproperty -Name "@ObjectFileName" -Value $fileSourceObj.FileInfo.FullName -Force
                    $compareObject.Object | Add-Member Noteproperty -Name "@ObjectFromFile" -Value $true -Force 
                    $compareObject.Object | Add-Member Noteproperty -Name "@ObjectFileName" -Value $compareObject.FileInfo.FullName -Force
                    $compareProperties = Compare-Objects $compareObject.Object $fileSourceObj.Object $item.ObjectType                    
                }

                $compareObjectsResult += [PSCustomObject]@{
                    Object1 = $compareObject.Object
                    Object2 = $fileSourceObj.Object
                    ObjectType = $item.ObjectType
                    Id = $fileSourceObj.Object.Id
                    Result = $compareProperties
                }                
            }

            if($global:chkSkipMissingSourcePolicies.IsChecked -ne $true) {
                foreach($fileCompareObj in $fileCompareObjs)
                {                
                    # Add objects that were not exported in source folder
                    if(($compareObjectsResult | Where { $_.Id -eq $fileCompareObj.Object.Id})) { continue }

                    $objName = Get-GraphObjectName $fileCompareObj.Object $item.ObjectType
                    if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                    {
                        continue
                    }                

                    $compareObjectsResult += [PSCustomObject]@{
                        Object1 = $fileCompareObj.Object
                        Object2 = $null
                        ObjectType = $item.ObjectType
                        Id = $fileCompareObj.Object.Id
                        Result = @([PSCustomObject]@{
                            Object1Value = (Get-GraphObjectName $fileCompareObj.Object $item.ObjectType)
                            Object2Value = $null
                            Match = $false
                        })
                    }
                }
            }

            if($outputType -eq "objectType")
            {
                $compResultValues = @()
            }

            foreach($compObj in @($compareObjectsResult | Where { $_.ObjectType.Id -eq $item.ObjectType.Id }))
            {
                $objName = Get-GraphObjectName (?? $compObj.Object1 $compObj.Object2) $item.ObjectType
                foreach($compValue in $compObj.Result)
                {
                    $compResultValues += [PSCustomObject]@{
                        ObjectName = $objName
                        Id = $compObj.Id
                        Type = $compObj.ObjectType.Title
                        ODataType = $compObj.Object1.'@OData.Type'
                        Property = $compValue.PropertyName
                        Value1 = $compValue.Object1Value
                        Value2 = $compValue.Object2Value
                        Category = $compValue.Category
                        SubCategory = $compValue.SubCategory
                        Match = $compValue.Match
                    }
                }
            }

            if($outputType -eq "objectType")
            {
                Save-BulkCompareResults $compResultValues (Join-Path $folderSource "Compare_$(((Get-Date).ToString("yyyyMMdd-HHmm"))).csv") $compareProps
            }
        }
        else
        {
            Write-Log "Folder $folderSource not found. Skipping compare" 2    
        }
    }

    if($outputType -eq "all" -and $compResultValues.Count -gt 0)
    {
        Save-BulkCompareResults $compResultValues (Join-Path $rootFolderSource "Compare_$(((Get-Date).ToString("yyyyMMDD-HHmm"))).csv") $compareProps
    }    

    Write-Log "****************************************************************"
    Write-Log "Bulk compare Exported Folders finished"
    Write-Log "****************************************************************"
    Write-Status ""
    if($compareObjectsResult.Count -eq 0)
    {
        [System.Windows.MessageBox]::Show("No objects were comparced. Verify folder and exported files", "Error", "OK", "Error")
    }
}

function Save-BulkCompareResults
{
    param($compResultValues, $file, $props)

    if($compResultValues.Count -gt 0)
    {
        $params = @{}
        try
        {        
            if($global:cbCompareCSVDelimiter.Text)
            {
                $params.Add("Delimiter", [char]$global:cbCompareCSVDelimiter.Text)
            }
        }
        catch
        {
            
        }
        Write-Log "Save bulk comare results to $file"
        $compResultValues | Select -Property $props | ConvertTo-Csv -NoTypeInformation @params | Out-File -LiteralPath $file -Force -Encoding UTF8
    } 
}

function Show-CompareForm
{
    param($objInfo)

    $script:cmpForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\CompareForm.xaml") -AddVariables
    if(-not $script:cmpForm) { return }

    $script:cmpForm.Tag = $objInfo

    $script:copareSource = $objInfo

    $global:cbCompareType.ItemsSource = $global:comparisonTypes | Where ShowOnObject -ne $false
    $global:cbCompareType.SelectedValue = (Get-Setting "Compare" "Type" "property")

    $global:txtIntuneObject.Text = (Get-GraphObjectName $objInfo.Object $objInfo.ObjectType)
    $global:txtIntuneObject.Tag = $objInfo

    Add-XamlEvent $script:cmpForm "btnClose" "add_click" {
        $script:cmpForm = $null
        Show-ModalObject 
    }

    Add-XamlEvent $script:cmpForm "btnStartCompare" "add_click" {
        Write-Status "Compare objects"
        Save-Setting "Compare" "Type" $global:cbCompareType.SelectedValue
        $script:currentObjName = ""
        Start-CompareExportObject
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
            $csvInfo = Get-CompareCsvInfo $global:dgCompareInfo.ItemsSource $script:cmpForm.Tag
            $csvInfo | Out-File -LiteralPath $sf.FileName -Force -Encoding UTF8
        }    
    }

    Add-XamlEvent $script:cmpForm "btnCompareCopy" "add_click" {

        (Get-CompareCsvInfo $global:dgCompareInfo.ItemsSource $script:cmpForm.Tag) | Set-Clipboard
    }    

    Add-XamlEvent $script:cmpForm "browseCompareObject" "add_click" {

        $path = Get-Setting "" "LastUsedFullPath"
        if($path) 
        {
            $path = [IO.Directory]::GetParent($path).FullName
            if($global:txtIntuneObject.Tag.ObjectType)
            {
                $objectTypePath = [IO.Path]::Combine($path, $global:txtIntuneObject.Tag.ObjectType.Id)
                if([IO.Directory]::Exists($objectTypePath))
                {
                    $path = $objectTypePath
                }
            }
        }

        if([String]::IsNullOrEmpty($global:lastCompareFile) -eq $false)
        {
            $path = ([IO.FileInfo]$global:lastCompareFile).DirectoryName
        }

        $of = [System.Windows.Forms.OpenFileDialog]::new()
        $of.Multiselect = $false
        $of.Filter = "Json files (*.json)|*.json"
        if($path)
        {
            $of.InitialDirectory = $path
        }

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

function Get-CompareCsvInfo
{
    param($comareInfo, $objInfo)

    $compResultValues = @()
    $objName = Get-GraphObjectName $objInfo.Object $objInfo.ObjectType
    foreach($compValue in $comareInfo)
    {
        $compResultValues += [PSCustomObject]@{
            ObjectName = $objName
            Id =  $objInfo.Object.Id
            Type = $objInfo.ObjectType.Title
            ODataType = $objInfo.Object.'@OData.Type'
            Property = $compValue.PropertyName
            Value1 = $compValue.Object1Value
            Value2 = $compValue.Object2Value
            Category = $compValue.Category
            SubCategory = $compValue.SubCategory
            Match = $compValue.Match
        }
    }            

    $compareProps = $script:defaultCompareProps

    # !!! Not supported yet
    #foreach($removeProp in $global:cbCompareProvider.SelectedItem.RemoveProperties)
    #{
    #    $compareProps.Remove($removeProp) | Out-Null
    #}

    foreach($removeProp in $global:cbCompareType.SelectedItem.RemoveProperties)
    {
        $compareProps.Remove($removeProp) | Out-Null
    }
    $compResultValues | Select -Property $compareProps | ConvertTo-Csv -NoTypeInformation
}

function Start-CompareExportObject
{
    if(-not $script:copareSource) { return }

    if(-not $global:txtCompareFile.Text)
    {
        [System.Windows.MessageBox]::Show("No file selected", "Compare", "OK", "Error")
        return
    }
    elseif([IO.File]::Exists($global:txtCompareFile.Text) -eq $false)
    {
        [System.Windows.MessageBox]::Show("File '$($global:txtCompareFile.Text)' not found", "Compare", "OK", "Error")
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
            $compareObj = Get-Content -LiteralPath $global:txtCompareFile.Text | ConvertFrom-Json 
        }
    }
    catch
    {
        [System.Windows.MessageBox]::Show("Failed to convert json file '$($global:txtCompareFile.Text)'", "Compare", "OK", "Error")
        return
    }

    $obj = Get-GraphObject $script:copareSource.Object $script:copareSource.ObjectType

    $script:currentObjName = Get-GraphObjectName $script:copareSource.Object $script:copareSource.ObjectType

    if($obj.Object."@OData.Type" -ne $compareObj."@OData.Type")
    {
        if(([System.Windows.MessageBox]::Show("The object types does not match.`n`nDo you to compare the objects?", "Compare", "YesNo", "Warning")) -eq "No")
        {
            return
        }
    }

    $compareObj | Add-Member Noteproperty -Name "@ObjectFileName" -Value $global:txtCompareFile.Text -Force
    $compareObj | Add-Member Noteproperty -Name "@ObjectFromFile" -Value $true -Force

    $compareResult = Compare-Objects $obj.Object $compareObj $obj.ObjectType

    $global:dgCompareInfo.ItemsSource = $compareResult
}

function Compare-Objects
{
    param($obj1, $obj2, $objectType)

    $script:compareProperties = @()

    if($obj1.'@OData.Type' -eq "#microsoft.graph.deviceManagementConfigurationPolicy" -or 
        $obj1.'@OData.Type' -eq "#microsoft.graph.deviceManagementIntent" -or 
        $obj1.'@OData.Type' -eq "#microsoft.graph.groupPolicyConfiguration")
    {
        # Always use documentation for Settings Catalog, Endpoint Security and Administrative Template policies
        # These use Graph API for docummentation and all properties will be documented
        $compareResult = Compare-ObjectsBasedonDocumentation $obj1 $obj2 $objectType
    }
    elseif($global:cbCompareType.SelectedItem.Compare)
    {
        $compareResult = & $global:cbCompareType.SelectedItem.Compare $obj1 $obj2 $objectType
    }
    else
    {
        Write-Log "Selected comparison type ($($global:cbCompareType.SelectedItem.Name)) does not have a Compare property specified" 3
    }

    $compareResult
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
    param($name, $value1, $value2, $category, $subCategory, $match = $null, [switch]$skip)

    $value1 = if($value1 -eq $null) { "" } else { $value1.ToString().Trim("`"") }
    $value2 = if($value2 -eq $null) { "" } else {  $value2.ToString().Trim("`"") }

    $compare += [PSCustomObject]@{
        PropertyName = $name
        Object1Value = $value1 #if($value1 -ne $null) { $value1.ToString().Trim("`"") } else { "" }
        Object2Value = $value2 #if($value2 -ne $null) { $value2.ToString().Trim("`"") } else { "" }
        Category = $category
        SubCategory = $subCategory
        Match = ?? $match ($value1 -eq $value2)
    }
    if($skip -eq $true) {
        $compare.Match = $null
    }

    $script:compareProperties += $compare
}

function Compare-ObjectsBasedonProperty
{
    param($obj1, $obj2, $objectType)

    Write-Status "Compare objects based on property values"

    Set-ColumnVisibility $false

    $skipBasicProperties = Get-XamlProperty $script:cmpForm "chkSkipCompareBasicProperties" "IsChecked"    

    $coreProps = @((?? $objectType.NameProperty "displayName"), "Description", "Id", "createdDateTime", "lastModifiedDateTime", "version")
    $postProps = @("Advertisements")
    $skipProps = @("@ObjectFromFile","@ObjectFileName")
    $skipPropertiesToCompare = @()
    if($skipBasicProperties) {
        $skipPropertiesToCompare += "roleScopeTagIds"
        $skipPropertiesToCompare += "roleScopeTags"
    }


    foreach ($propName in $coreProps)
    {
        if(-not ($obj1.PSObject.Properties | Where Name -eq $propName))
        {
            continue
        }
        $val1 = ($obj1.$propName | ConvertTo-Json -Depth 10)
        $val2 = ($obj2.$propName | ConvertTo-Json -Depth 10)
        Add-CompareProperty $propName $val1 $val2 -Skip:($skipBasicProperties -eq $true)
    }    

    $addedProps = @()
    foreach ($propName in ($obj1.PSObject.Properties | Select Name).Name) 
    {
        if($propName -in $coreProps) { continue }
        if($propName -in $postProps) { continue }
        if($propName -in $skipProps) { continue }

        if($propName -like "*@OData*" -or $propName -like "#microsoft.graph*") { continue }

        $addedProps += $propName
        $val1 = ($obj1.$propName | ConvertTo-Json -Depth 10)
        $val2 = ($obj2.$propName | ConvertTo-Json -Depth 10)
        Add-CompareProperty $propName $val1 $val2 -Skip:($skipPropertiesToCompare -contains $propName)
    }

    foreach ($propName in ($obj2.PSObject.Properties | Select Name).Name) 
    {
        if($propName -in $coreProps) { continue }
        if($propName -in $postProps) { continue }
        if($propName -in $skipProps) { continue }
        if($propName -in $addedProps) { continue }

        if($propName -like "*@OData*" -or $propName -like "#microsoft.graph*") { continue }

        $val1 = ($obj1.$propName | ConvertTo-Json -Depth 10)
        $val2 = ($obj2.$propName | ConvertTo-Json -Depth 10)
        Add-CompareProperty $propName $val1 $val2 -Skip:($skipPropertiesToCompare -contains $propName)
    }    

    $skipAssignments = Get-XamlProperty $script:cmpForm "chkSkipCompareAssignments" "IsChecked"
    foreach ($propName in $postProps)
    {
        if(-not ($obj1.PSObject.Properties | Where Name -eq $propName))
        {
            continue
        }
        $val1 = ($obj1.$propName | ConvertTo-Json -Depth 10)
        $val2 = ($obj2.$propName | ConvertTo-Json -Depth 10)
        Add-CompareProperty $propName $val1 $val2 -Skip:($skipAssignments -eq $true)
    }
    
    $script:compareProperties
}

function Get-CompareCustomColumnsDoc
{
    param($obj)

    if($obj.'@OData.Type' -eq "#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration")
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
    param($obj1, $obj2, $objectType)

    Write-Status "Compare objects based on documentation values"

    Get-CompareCustomColumnsDoc $obj1

    # ToDo: set this based on configuration value
    $script:assignmentOutput = "simpleFullCompare"

    $docObj1 = Invoke-ObjectDocumentation ([PSCustomObject]@{
        Object = $obj1
        ObjectType = $objectType
    })
    

    $docObj2 = Invoke-ObjectDocumentation ([PSCustomObject]@{
        Object = $obj2
        ObjectType = $objectType
    })

    $settingsValue = ?? $objectType.CompareValue "Value"

    $skipBasicProperties = Get-XamlProperty $script:cmpForm "chkSkipCompareBasicProperties" "IsChecked"    

    if($docObj1.BasicInfo -and -not ($docObj1.BasicInfo | where Value -eq $obj1.Id))
    {
        # Make sure the Id property is included
        Add-CompareProperty "Id" $obj1.Id $obj2.Id $docObj1.BasicInfo[0].Category -Skip:($skipBasicProperties -eq $true)
    }

    foreach ($prop in $docObj1.BasicInfo)
    {
        $val1 = $prop.Value 
        $prop2 = $docObj2.BasicInfo | Where Name -eq $prop.Name
        $val2 = $prop2.Value 
        Add-CompareProperty $prop.Name $val1 $val2 $prop.Category -Skip:($skipBasicProperties -eq $true)
    }

    $addedProperties = @()

    if($docObj1.InputType -eq "Settings")
    {
        foreach ($prop in $docObj1.Settings)
        {
            if(($prop.SettingId + $prop.ParentSettingId + $prop.RowIndex) -in $addedProperties) { continue }

            $addedProperties += ($prop.SettingId + $prop.ParentSettingId + $prop.RowIndex)
            $val1 = $prop.Value 
            $prop2 = $docObj2.Settings | Where { $_.SettingId -eq $prop.SettingId -and $_.ParentSettingId -eq $prop.ParentSettingId -and $_.RowIndex -eq $prop.RowIndex }
            if($val1 -isnot [Array] -and $prop2.Value -is [Array])
            {
                Write-Log "Compare property for $($prop.SettingId) found based on value" 2
                $prop2 = $prop2 | Where Value -eq $val1
            }
            $val2 = $prop2.Value
            Add-CompareProperty $prop.Name $val1 $val2 $prop.Category

            # ToDo: fix lazy copy/past coding
            $children1 = $docObj1.Settings | Where ParentId -eq $prop.Id
            $children2 = $docObj2.Settings | Where ParentId -eq $prop2.Id
            
            # Add children defined on Object 1 property
            foreach ($childProp in $children1)
            {
                if(($childProp.SettingId + $childProp.ParentSettingId + $childProp.RowIndex) -in $addedProperties) { continue }

                $addedProperties += ($childProp.SettingId + $childProp.ParentSettingId + $childProp.RowIndex)
                $val1 = $childProp.Value 
                $prop2 = $docObj2.Settings | Where { $_.SettingId -eq $childProp.SettingId -and $_.ParentSettingId -eq $childProp.ParentSettingId -and $_.RowIndex -eq $childProp.RowIndex}
                if($val1 -isnot [Array] -and $prop2.Value -is [Array])
                {
                    Write-Log "Compare property for $($childProp.SettingId) found based on value" 2
                    $prop2 = $prop2 | Where Value -eq $val1
                }
                $val2 = $prop2.Value

                Add-CompareProperty $childProp.Name $val1 $val2 $prop.Category
            }
            
            # Add children defined only on Object 2 property e.g. Baseline Firewall profile was disable AFTER export.
            # This is to make sure all children are added under its parent and not last in the table
            foreach ($childProp in $children2)
            {
                if(($childProp.SettingId + $childProp.ParentSettingId + $childProp.RowIndex) -in $addedProperties) { continue }

                $addedProperties += ($childProp.SettingId + $childProp.ParentSettingId + $childProp.RowIndex)
                $val2 = $childProp.Value 
                $prop2 = $docObj1.Settings | Where { $_.SettingId -eq $childProp.SettingId -and $_.ParentSettingId -eq $childProp.ParentSettingId -and $_.RowIndex -eq $childProp.RowIndex }
                if($val2 -isnot [Array] -and $prop2.Value -is [Array])
                {
                    Write-Log "Compare property for $($childProp.SettingId) found based on value" 2
                    $prop2 = $prop2 | Where Value -eq $val1
                }
                $val1 = $prop2.Value
                Add-CompareProperty $childProp.Name $val1 $val2 $prop.Category
            }
        }
        
        # These objects are defined only on Object 2. They will be last in the table
        foreach ($prop in $docObj2.Settings)
        {
            if(($prop.SettingId + $prop.ParentSettingId + $prop.RowIndex) -in $addedProperties) { continue }

            $addedProperties += ($prop.SettingId + $prop.ParentSettingId + $prop.RowIndex)
            $val2 = $prop.Value    
            $prop2 = $docObj1.Settings | Where  { $_.SettingId -eq $prop.SettingId -and $_.ParentSettingId -eq $prop.ParentSettingId -and $_.RowIndex -eq $childProp.RowIndex }
            if($val2 -isnot [Array] -and $prop2.Value -is [Array])
            {
                Write-Log "Compare property for $($prop.SettingId) found based on value" 2
                $prop2 = $prop2 | Where Value -eq $val2
            }
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
            $prop2 = $docObj2.Settings | Where { $_.EntityKey -eq $prop.EntityKey -and $_.Category -eq $prop.Category -and $_.SubCategory -eq $prop.SubCategory -and $_.Enabled -eq $prop.Enabled }
            $val2 = $prop2.$settingsValue
            if($val1 -isnot [array] -and $val2 -is [array] -and $val2.Count -gt 1) 
            {
                Write-Log "Multiple compare results returend for $($prop.Name). Using first result" 2
                $val2 = $val2[0]
            }
            Add-CompareProperty $prop.Name $val1 $val2 $prop.Category $prop.SubCategory
        }
        
        # These objects are defined only on Object 2. They will be last in the table
        foreach ($prop in $docObj2.Settings)
        {
            if(($prop.EntityKey + $prop.Category + $prop.SubCategory) -in $addedProperties) { continue }

            $addedProperties += ($prop.EntityKey + $prop.Category + $prop.SubCategory)
            $val2 = $prop.$settingsValue
            $prop2 = $docObj1.Settings | Where  { $_.EntityKey -eq $prop.EntityKey -and $_.Category -eq $prop.Category -and $_.SubCategory -eq $prop.SubCategory -and $_.Enabled -eq $prop.Enabled  }
            $val1 = $prop2.$settingsValue
            if($val2 -isnot [array] -and $val1 -is [array] -and $val1.Count -gt 1) 
            {
                Write-Log "Multiple compare results returend for $($prop.Name). Using first result" 2
                $val1 = $val1[0]
            }

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

    $complianceActionsAdded = @()
    foreach($complianceAction in $docObj1.ComplianceActions)
    {
        $complianceAction2 = $docObj2.ComplianceActions | Where { $_.IdStr -eq $complianceAction.IdStr }
        $complianceActionsAdded += $complianceAction.IdStr
        $val1 = ($complianceAction.Action + [environment]::NewLine + $complianceAction.Schedule + [environment]::NewLine + $complianceAction.MessageTemplateId + [environment]::NewLine + $complianceAction.EmailCCIds)
        $val2 = ($complianceAction2.Action + [environment]::NewLine + $complianceAction2.Schedule + [environment]::NewLine + $complianceAction2.MessageTemplateId + [environment]::NewLine + $complianceAction2.EmailCCIds)

        Add-CompareProperty $complianceAction.Category $val1 $val2 
    }

    foreach($complianceAction in $docObj2.ComplianceActions)
    {
        if(($complianceAction.IdStr) -in $complianceActionsAdded) { continue }
        $complianceAction2 = $docObj1.ComplianceActions | Where { $_.IdStr -eq $complianceAction.IdStr }
        $complianceActionsAdded += $complianceAction.IdStr
        $val2 = ($complianceAction.Action + [environment]::NewLine + $complianceAction.Schedule + [environment]::NewLine + $complianceAction.MessageTemplateId + [environment]::NewLine + $complianceAction.EmailCCIds)
        $val1 = ($complianceAction2.Action + [environment]::NewLine + $complianceAction2.Schedule + [environment]::NewLine + $complianceAction2.MessageTemplateId + [environment]::NewLine + $complianceAction2.EmailCCIds)

        Add-CompareProperty $complianceAction.Category $val1 $val2 
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

    $script:compareProperties
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

    $skipAssignments = Get-XamlProperty $script:cmpForm "chkSkipCompareAssignments" "IsChecked"

    if($assignment.RawIntent)
    {
        Add-CompareProperty $assignment.Category $val1 $val2 -Category $assignment.GroupMode -match $match  -Skip:($skipAssignments -eq $true)
    }
    else
    {
        Add-CompareProperty $assignmentStr $val1 $val2 -Category $assignment.GroupMode -match $match -Skip:($skipAssignments -eq $true)
    }
}

function Invoke-SilentBatchJob
{
    param($settingsObj)

    if(-not $global:MSALToken) { return } # Skip if not authenticated

    if($settingsObj.BulkCompare)
    {
        $global:currentViewObject =  $global:viewObjects | Where { $_.ViewInfo.ID -eq "IntuneGraphAPI" }

        $script:cmpForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\BulkCompare.xaml") -AddVariables
        if(-not $script:cmpForm) { return }

        $script:compareObjects = Get-GraphBatchObjectTypes $settingsObj.BulkCompare        

        $global:cbCompareProvider.ItemsSource = @($global:compareProviders)
        $global:cbCompareProvider.SelectedValue = ($settingsObj.BulkCompare | Where Name -eq "cbCompareProvider").Value
    
        $global:cbCompareSave.ItemsSource = @($script:saveType)
        $global:cbCompareSave.SelectedValue = ($settingsObj.BulkCompare | Where Name -eq "cbCompareSave").Value
    
        $global:cbCompareType.ItemsSource = $global:comparisonTypes | Where ShowOnBulk -ne $false 
        $global:cbCompareType.SelectedValue = ($settingsObj.BulkCompare | Where Name -eq "cbCompareType").Value
    
        $global:cbCompareCSVDelimiter.ItemsSource = @("", ",",";","-","|")
        $global:cbCompareCSVDelimiter.SelectedValue = ($settingsObj.BulkCompare | Where Name -eq "cbCompareCSVDelimiter").Value
        
        Set-CompareProviderOptions $global:cbCompareProvider

        Set-BatchProperties $settingsObj.BulkCompare $script:cmpForm -SkipMissingControlWarning        
        Set-BatchProperties $settingsObj.BulkCompare $global:ccContentProviderOptions.Content -SkipMissingControlWarning

        $global:dgObjectsToCompare.ItemsSource = @($script:compareObjects)
    
        if($global:cbCompareProvider.SelectedItem.BulkCompare)
        {
            & $global:cbCompareProvider.SelectedItem.BulkCompare
        }    
    }
}

