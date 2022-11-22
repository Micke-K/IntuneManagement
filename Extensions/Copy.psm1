function Get-ModuleVersion
{
    '1.0.2'
}

function Invoke-InitializeModule
{

}


function Invoke-ViewActivated
{
    if($global:currentViewObject.ViewInfo.ID -ne "IntuneGraphAPI") { return }
    
    $tmp = $mnuMain.Items | Where Name -eq "EMBulk"
    if($tmp)
    {
        $tmp.AddChild(([System.Windows.Controls.Separator]::new())) | Out-Null
        $subItem = [System.Windows.Controls.MenuItem]::new()
        $subItem.Header = "Cop_y"
        $subItem.Add_Click({Show-CopyBulkForm})
        $tmp.AddChild($subItem)
    }
}

function Show-CopyBulkForm
{
    $script:form = Get-XamlObject ($global:AppRootFolder + "\Xaml\BulkCopy.xaml") -AddVariables
    if(-not $script:form) { return }

    $global:txtCopyFromPattern.Text = Get-Setting "Copy" "CopyFromPattern"    
    $global:txtCopyToPattern.Text = Get-Setting "Copy" "CopyToPattern"

    $script:copyObjects = @()
    foreach($objType in $global:lstMenuItems.ItemsSource)
    {
        if(-not $objType.Title) { continue }

        $script:copyObjects += New-Object PSObject -Property @{
            Title = $objType.Title
            Selected = $true
            ObjectType = $objType
        }
    }

    $column = Get-GridCheckboxColumn "Selected"
    $global:dgObjectsToCopy.Columns.Add($column)

    $column.Header.IsChecked = $true # All items are checked by default
    $column.Header.add_Click({
            foreach($item in $global:dgObjectsToCopy.ItemsSource)
            {
                $item.Selected = $this.IsChecked
            }
            $global:dgObjectsToCopy.Items.Refresh()
        }
    ) 

    # Add Object type column
    $binding = [System.Windows.Data.Binding]::new("Title")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Object type"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:dgObjectsToCopy.Columns.Add($column)

    $global:dgObjectsToCopy.ItemsSource = $script:copyObjects

    Add-XamlEvent $script:form "btnClose" "add_click" {
        $script:form = $null
        Show-ModalObject 
    }

    Add-XamlEvent $script:form "btnStartCopy" "add_click" {
        Write-Status "Copy objects"
        Start-BulkCopyObjects
        Write-Status "" 
    }

    Show-ModalForm "Bulk Copy Objects" $script:form -HideButtons
}

function Start-BulkCopyObjects
{
    Write-Log "****************************************************************"
    Write-Log "Start bulk copy"
    Write-Log "****************************************************************"

    $copyFrom =  $global:txtCopyFromPattern.Text
    $copyTo = $global:txtCopyToPattern.Text

    if(-not $copyFrom -or -not $copyTo)
    {
        [System.Windows.MessageBox]::Show("Both name patterns must be specified", "Error", "OK", "Error")
        return
    }

    Save-Setting "Copy" "CopyFromPattern" $global:txtCopyFromPattern.Text        
    Save-Setting "Copy" "CopyToPattern" $global:txtCopyToPattern.Text        

    foreach($item in ($global:dgObjectsToCopy.ItemsSource | where Selected -eq $true))
    { 
        Write-Status "Copy $($item.ObjectType.Title) objects" -Force -SkipLog
        Write-Log "----------------------------------------------------------------"
        Write-Log "Copy $($item.ObjectType.Title) objects"
        Write-Log "----------------------------------------------------------------"
    
        [array]$graphObjects = Get-GraphObjects -property $item.ObjectType.ViewProperties -objectType $item.ObjectType
        
        $nameProp = ?? $item.ObjectType.NameProperty "displayName"

        foreach($graphObj in ($graphObjects | Where { $_.Object."$($nameProp)" -imatch [regex]::Escape($copyFrom) }))
        {
            $sourceName = $graphObj.Object."$($nameProp)"
            $copyName  = $sourceName -ireplace [regex]::Escape($copyFrom),$copyTo

            $copyObj = $graphObjects | Where { $_.Object."$($nameProp)" -eq $copyName -and $_.Object.'@OData.Type' -eq $graphObj.Object.'@OData.Type' }
        
            if(($copyObj | measure).Count -gt 0)
            {
                Write-Log "Object with name $copyName already exists. $sourceName will not be copied" 2
                continue
            }
            else
            {
                Write-Status "Create $copyName from $sourceName" -Force
                
                if($graphObj.ObjectType.PreCopyCommand)
                {
                    if((& $graphObj.ObjectType.PreCopyCommand $graphObj.Object $graphObj.ObjectType $copyName))
                    {
                        continue
                    }
                }
        
                $copyFromObj = (Get-GraphObject $graphObj.Object $graphObj.ObjectType -SkipAssignments).Object
        
                # Convert to Json and back to clone the object
                $obj = ConvertTo-Json $copyFromObj -Depth 10 | ConvertFrom-Json
                if($obj)
                {
                    # Import new profile
                    Set-GraphObjectName $obj $graphObj.ObjectType $copyName
        
                    $newObj = Import-GraphObject $obj $graphObj.ObjectType
                    if($newObj)
                    {
                        if($graphObj.ObjectType.PostCopyCommand)
                        {
                            & $graphObj.ObjectType.PostCopyCommand $copyFromObj $newObj $graphObj.ObjectType
                        }
                    }
                    else
                    {
                        Write-log "Failed to copy $sourceName" 3
                    }
                }                
            }       
        }
    }

    Write-Log "****************************************************************"
    Write-Log "Bulk copy finished"
    Write-Log "****************************************************************"
    Write-Status ""
}