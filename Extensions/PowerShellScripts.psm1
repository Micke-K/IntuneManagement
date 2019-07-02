########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-PowerShellScriptName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-PowerShellScripts}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-PowerShellScriptName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all powershell scripts"
                Import-AllPowerShellScriptObjects (Join-Path $rootFolder (Get-PowerShellScriptFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-PowerShellScriptName)
            Script = [ScriptBlock]{
                param($rootFolder)
                    Write-Status "Export all powershell scripts"
                    Get-PowerShellScriptObjects | ForEach-Object { Export-SinglePowerShellScript $PSItem.Object (Join-Path $rootFolder (Get-PowerShellScriptFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-PowerShellScriptFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-PowerShellScriptName
{
    return "PowerShell Script"
}

function Get-PowerShellScriptFolderName
{
    return "PowerShell"
}

function Get-PowerShellScripts
{
    Write-Status "Loading PowerShell objects" 
    $dgObjects.ItemsSource = @(Get-PowerShellScriptObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllPowerShellScripts $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedPowerShellScript $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllPowerShellScriptObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-PowerShellScriptObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    $exportExtension = (New-Object PSObject -Property @{            
            Xaml = @"
<Grid>
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>        
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />
        <ColumnDefinition Width="*" />
    </Grid.ColumnDefinitions>

    <StackPanel Orientation="Horizontal" Margin="0,0,5,0">
        <Label Content="Export script" />
        <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="Export the powershell script to a ps1 file" />
    </StackPanel>
    <CheckBox Grid.Column='1' Name='chkExportScript' VerticalAlignment="Center" IsChecked="true" />

</Grid>
"@
            Script = [ScriptBlock]{
                param($form)
                $script:chkExportScript = $form.FindName("chkExportScript")    
            }
            })

    $script:exportParams.Add("Extension", $exportExtension)

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -copy ([scriptblock]{Copy-PowerShellScript}) -ViewFullObject ([scriptblock]{Get-PowerShellScriptObject $global:dgObjects.SelectedItem.Object})           

    #Add download button
    $btnDownload = New-Object System.Windows.Controls.Button    
    $btnDownload.Content = 'Download'
    $btnDownload.Name = 'btnDownload'
    $btnDownload.Margin = "5,0,0,0"  
    $btnDownload.Width = "100"  
    $spSubMenu.Children.Insert(0, $btnDownload)

    $btnDownload.Add_Click({
        Invoke-DownloadScript
    })
}

function Get-PowerShellScriptObjects
{
    Get-GraphObjects -Url "/deviceManagement/deviceManagementScripts"
}

function Get-PowerShellScriptObject
{
    param($object, $additional = "")

    if(-not $Object.id) { return }
    Invoke-GraphRequest -Url "/deviceManagement/deviceManagementScripts/$($object.id)$additional"
}

function Export-AllPowerShellScripts
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SinglePowerShellScript $objTmp.Object $path            
        }
    }
}

function Export-SelectedPowerShellScript
{
    param($path = "$env:Temp")

    Export-SinglePowerShellScript $global:dgObjects.SelectedItem.Object $path
}

function Export-SinglePowerShellScript
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-PowerShellScriptFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.displayName)"
        $obj = Get-PowerShellScriptObject -object $psObj -additional "?`$expand=assignments"
        #$obj = Invoke-GraphRequest -Url "/deviceManagement/deviceManagementScripts/$($psObj.id)?`$expand=assignments"
        if($obj)
        {            
            $fileName = "$path\$((Remove-InvalidFileNameChars $obj.displayName)).json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force
            if($script:chkExportScript.IsChecked)
            {
                $fileName = "$path\$($obj.FileName)"
                [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($obj.scriptContent)) | Out-File $fileName -Force
            }
            Add-MigrationInfo $obj.assignments
            $global:exportedObjects++
        }        
    }
}

function Copy-PowerShellScript
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect PowerShell script you want to copy", "Error", "OK", "Error") 
        return 
    }

    $ret = Show-InputDialog "Copy PowerShell script" "Select name for the new object" "$($global:dgObjects.SelectedItem.displayName) - Copy"

    if($ret)
    {
        # Export profile
        Write-Status "Export $($dgObjects.SelectedItem.displayName)"        
        $obj = Get-PowerShellScriptObject -object $dgObjects.SelectedItem.Object
        #$obj = Invoke-GraphRequest -Url "/deviceManagement/deviceManagementScripts/$($dgObjects.SelectedItem.id)"
        if($obj)
        {            
            # Import new profile
            $obj.displayName = $ret
            Import-PowerShellScript $obj | Out-null

            $dgObjects.ItemsSource = @(Get-PowerShellScriptObjects)
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

function Import-PowerShellScript
{
    param($obj)

    Remove-ObjectProperty $obj "id"
    Remove-ObjectProperty $obj "createdDateTime"
    Remove-ObjectProperty $obj "lastModifiedDateTime"
    Remove-ObjectProperty $obj "assignments@odata.context"
    Remove-ObjectProperty $obj "assignments"    

    Write-Status "Import $($obj.displayName)"

    Invoke-GraphRequest -Url "/deviceManagement/deviceManagementScripts" -Content (ConvertTo-Json $obj -Depth 5) -HttpMethod POST

}

function Import-AllPowerShellScriptObjects
{
    param(
        $path = "$env:Temp"
    )    

    Import-PowerShellScriptObjects (Get-JsonFileObjects $path)
}

function Import-PowerShellScriptObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import PowerShell scripts"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import PowerShell script: $($obj.Object.displayName)"

        $assignments = Get-GraphAssignmentsObject $obj.Object ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_assignments.json")

        $response = Import-PowerShellScript $obj.Object

        if($response)
        {
            $global:importedObjects++
            Import-GraphAssignments $assignments "deviceManagementScriptAssignments" "/deviceManagement/deviceManagementScripts/$($response.Id)/assign"
        }        
    }
    
    $dgObjects.ItemsSource = @(Get-PowerShellScriptObjects)
    Write-Status ""
}

function Invoke-DownloadScript
{
    if(-not $global:dgObjects.SelectedItem.Object.id) { return }

    $obj = Get-PowerShellScriptObject -object $dgObjects.SelectedItem.Object
    #$obj = Invoke-GraphRequest -Url "/deviceManagement/deviceManagementScripts/$($global:dgObjects.SelectedItem.Object.id)"
    if($obj.scriptContent)
    {            
        Write-Log "Download PowerShell script '$($obj.FileName)' from $($obj.displayName)"
        $fileName = "$path\$($obj.FileName)"
        
        $dlgSave = New-Object -Typename System.Windows.Forms.SaveFileDialog
        $dlgSave.InitialDirectory = Get-SettingValue "IntuneRootFolder" $env:Temp
        $dlgSave.FileName = $obj.FileName    
        if($dlgSave.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $dlgSave.Filename)
        {            
            [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($obj.scriptContent)) | Out-File $dlgSave.Filename -Force
        }
    }    
}