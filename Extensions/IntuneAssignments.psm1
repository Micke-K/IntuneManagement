<#
.SYNOPSIS
Module for listing Intune assignments

.DESCRIPTION

.NOTES
  Author:         Mikael Karlsson
#>
function Get-ModuleVersion
{
    '1.0.3'
}

function Invoke-InitializeModule
{
    Add-EMToolsViewItem (New-Object PSObject -Property @{
        Title = "Intune Assignments"
        Id = "IntuneAssignments"
        ViewID = "EMTools"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Icon="DeviceConfiguration"
        ShowViewItem = { Show-EMToolsIntuneAssignments }
    })
}

function Show-EMToolsIntuneAssignments
{
    if(-not $script:frmIntuneAssignments)
    {
        $script:frmIntuneAssignments = Get-XamlObject ($global:AppRootFolder + "\Xaml\EndpointManagerToolsIntuneAssignments.xaml") #-AddVariables

        if(-not $script:frmIntuneAssignments) { return }
        
        Add-XamlEvent $script:frmIntuneAssignments "btnBrowseIntuneAssignmentsExportPath" "add_click" ({
            $folder = Get-Folder (Get-XamlProperty $script:frmIntuneAssignments "txtIntuneAssignmentsExportPath" "Text") "Select root folder for exported files"
            if($folder)
            {
                Set-XamlProperty $script:frmIntuneAssignments "txtIntuneAssignmentsExportPath" "Text" $folder
            }
        })

        Add-XamlEvent $script:frmIntuneAssignments "btnGetIntuneAssignments" "add_click" ({
            $folder = Get-XamlProperty $script:frmIntuneAssignments "txtIntuneAssignmentsExportPath" "Text"
            if($folder)
            {
                Write-Status "Get Intune Assignments"
                Get-EMIntuneAssignments $folder
                Write-Status ""
            }
        })

        Add-XamlEvent $script:frmIntuneAssignments "btnIntuneAssignmentsCopy" "add_click" ({
            $script:objAssignments | Select Name, Type, IncludedString, ExcludedString | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
        })

        Add-XamlEvent $script:frmIntuneAssignments "btnIntuneAssignmentsSave" "add_click" ({
        
            $dlgSave = New-Object -Typename System.Windows.Forms.SaveFileDialog
            #$dlgSave.InitialDirectory = Get-SettingValue "IntuneRootFolder" $env:Temp
            $dlgSave.FileName = $obj.FileName
            $dlgSave.DefaultExt = "*.csv"
            $dlgSave.Filter = "CSV (*.csv)|*.csv|All files (*.*)| *.*"            
            if($dlgSave.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $dlgSave.Filename)
            {
                $script:objAssignments | Select Name, Type, IncludedString, ExcludedString | ConvertTo-Csv -NoTypeInformation | Out-File -LiteralPath $dlgSave.Filename -Encoding UTF8 -Force
            }            
        })
    }
    
    $global:grdToolsMain.Children.Clear()
    $global:grdToolsMain.Children.Add($frmIntuneAssignments)
}

function Get-EMIntuneAssignmentInfo
{
    param($rootDir)

    Write-Status "Gather Export Information"

    $path = "$rootDir\Groups"

    $script:htGroups = @{}

    foreach($file in (Get-Item -path "$path\*.json"))
    {
        $graphObj = (ConvertFrom-Json (Get-Content -LiteralPath $file.FullName -Raw))
        $htGroups.Add($graphObj.Id, $graphObj)
    }
    
    $script:fileArr = @()

    foreach($path in [IO.Directory]::EnumerateDirectories($rootDir))
    {    
        if($path -eq "$rootDir\Groups") { continue }        

        foreach($file in (Get-Item -path "$path\*.json" -Exclude @("*_settings.json","*_assignments.json")))
        {
            $graphObj = (ConvertFrom-Json (Get-Content -LiteralPath $file.FullName -Raw))
        
            $obj = New-Object PSObject -Property @{
                    FileName = $file.Name
                    FileInfo = $file
                    Selected = $SelectedStatus
                    Object = $graphObj
            }

            $script:fileArr += $obj
        }
    }
}

function Get-EMIntuneAssignments 
{ 
    param($folder)
    
    Set-XamlProperty $script:frmIntuneAssignments "dgIntuneAssignments" "ItemsSource" $null

    $folderDI = [IO.DirectoryInfo]$folder
    if(-not $folderDI.Exists) { return }

    Get-EMIntuneAssignmentInfo $folder

    Write-Status "Collect exported assignments"

    $intuneViewObj = $global:viewObjects | Where { $_.ViewInfo.ID -eq "IntuneGraphAPI" }

    $script:objAssignments = @()

    foreach($fileObj in $script:fileArr)
    {
        $objectType = $null
        $folderName = $fileObj.FileInfo.Directory.Name
        if($folderName)
        {
            $objectType = $intuneViewObj.ViewItems | Where Id -eq $folderName
        }

        $obj = New-Object PSObject -Property @{
            Object = $fileObj.Object
            Name = $fileObj.Object."$((?? $objectType.NameProperty "displayName"))"
            Type = $null
            Included = $null
            Excluded = $null
            IncludedString = ""
            ExcludedString = ""
        }
        $obj.Included = @()
        $obj.Excluded = @()
        if($fileObj.Object.'@OData.Type')
        {
            $obj.Type = $fileObj.Object.'@OData.Type'.Split('.')[-1]
        }
        else
        {
            $obj.Type = $file.Directory.Parent.Name
        }

        foreach($assignment in $fileObj.Object.assignments)
        {
            $assignmentObj = $null
            $included = $true

            if($assignment.target.'@odata.type' -eq "#microsoft.graph.groupAssignmentTarget" -or 
                $assignment.target.'@odata.type' -eq "#microsoft.graph.exclusionGroupAssignmentTarget")
            {
                if($script:htGroups.ContainsKey($assignment.target.groupId))
                {
                    $assignmentObj = $script:htGroups[$assignment.target.groupId].displayName
                }
                else
                {
                    $assignmentObj = $assignment.target.groupId
                    Write-Warning "Could not find a group with ID $($assignment.target.groupId)"
                }
                $included = $assignment.target.'@odata.type' -eq "#microsoft.graph.groupAssignmentTarget"
            }
            elseif($assignment.target.'@odata.type' -eq "#microsoft.graph.allDevicesAssignmentTarget") 
            {
                $assignmentObj = "All Devices"
            }
            elseif($assignment.target.'@odata.type' -eq "#microsoft.graph.allLicensedUsersAssignmentTarget") 
            {
                $assignmentObj = "All Users"
            }

            if($included)
            {
                $obj.Included += $assignmentObj
            }
            else
            {
                $obj.Excluded += $assignmentObj
            }
        }
        $obj.IncludedString = $obj.Included -join ";"
        $obj.ExcludedString = $obj.Excluded -join ";"

        $script:objAssignments += $obj
    }

    Add-XamlEvent $script:frmIntuneAssignments "txtIntuneAssignmentsFilter" "Add_LostFocus" ({
        Invoke-IntueAssignmentFilterBoxChanged $this
    })
    
    Add-XamlEvent $script:frmIntuneAssignments "txtIntuneAssignmentsFilter" "Add_GotFocus" ({
        if($this.Tag -eq "1" -and $this.Text -eq "Filter") { $this.Text = "" }
        Invoke-IntueAssignmentFilterBoxChanged $this ($script:frmIntuneAssignments.FindName("dgIntuneAssignments"))
    })    
    
    Add-XamlEvent $script:frmIntuneAssignments "txtIntuneAssignmentsFilter" "Add_TextChanged" ({
        Invoke-IntueAssignmentFilterBoxChanged $this ($script:frmIntuneAssignments.FindName("dgIntuneAssignments"))
    })

    Invoke-IntueAssignmentFilterBoxChanged ($script:frmIntuneAssignments.FindName("txtIntuneAssignmentsFilter")) ($script:frmIntuneAssignments.FindName("dgIntuneAssignments"))

    $ocList = [System.Collections.ObjectModel.ObservableCollection[object]]::new(@($script:objAssignments))

    Set-XamlProperty $script:frmIntuneAssignments "dgIntuneAssignments" "ItemsSource" ([System.Windows.Data.CollectionViewSource]::GetDefaultView($ocList))
}

function Invoke-IntueAssignmentFilterBoxChanged
{
    param($txtBox, $dgObject)

    $filter = $null
    
    if($txtBox.Text.Trim() -eq "" -and $txtBox.IsFocused -eq $false)
    {
        $txtBox.FontStyle = "Italic"
        $txtBox.Tag = 1
        $txtBox.Text = "Filter"
        $txtBox.Foreground="Lightgray"        
    }
    elseif($txtBox.Tag -eq "1" -and $txtBox.Text -eq "Filter" -and $txtBox.IsFocused -eq $false)
    {
        
    }
    else
    {
        $txtBox.FontStyle = "Normal"
        $txtBox.Tag = $null
        $txtBox.Foreground="Black"
        $txtBox.Background="White"

        if($txtBox.Text)
        {
            $filter = {
                param ($item)

                return ($item.Name -match [regex]::Escape($txtBox.Text) -or $item.IncludedString -match [regex]::Escape($txtBox.Text) -or $item.ExcludedString -match [regex]::Escape($txtBox.Text) )
            }
        }         
    }

    if($dgObject.ItemsSource -is [System.Windows.Data.ListCollectionView] -and $txtBox.IsFocused -eq $true)
    {
        # This causes odd behaviour with focus e.g. and item has to be clicked twice to be selected 
        $dgObject.ItemsSource.Filter = $filter
        #$dgObject.ItemsSource.Refresh()
    }
}
