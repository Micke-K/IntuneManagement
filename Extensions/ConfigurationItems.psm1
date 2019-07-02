########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-DeviceConfigurationName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-DeviceConfigurations}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-DeviceConfigurationName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all device configuration objects"
                Import-AllDeviceConfigurationObjects (Join-Path $rootFolder (Get-DeviceConfigurationFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-DeviceConfigurationName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all device configuration objects"
                    Get-DeviceConfigurationObjects | ForEach-Object { Export-SingleDeviceConfiguration $PSItem.Object (Join-Path $rootFolder (Get-DeviceConfigurationFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-DeviceConfigurationFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-DeviceConfigurationName
{
    return "Device Configurations"
}

function Get-DeviceConfigurationFolderName
{
    return "DeviceConfigurations"
}

function Get-DeviceConfigurations
{
    Write-Status "Loading device configurations" 
    $dgObjects.ItemsSource = @(Get-DeviceConfigurationObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllDeviceConfigurations $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedDeviceConfiguration $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })
    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllDeviceConfigurationObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-DeviceConfigurationObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -copy ([scriptblock]{Copy-DeviceConfiguration}) -ViewFullObject ([scriptblock]{Get-DeviceConfigurationObject $global:dgObjects.SelectedItem.Object})               
}

function Get-DeviceConfigurationObjects
{
    Get-GraphObjects -Url "/deviceManagement/deviceConfigurations"#,"/deviceManagement/groupPolicyConfigurations"
}

function Get-DeviceConfigurationObject
{
    param($object, $additional = "")

    if(-not $Object.id) { return }

    Invoke-GraphRequest -Url "/deviceManagement/deviceConfigurations/$($Object.id)$additional"
}

function Export-AllDeviceConfigurations
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleDeviceConfiguration $objTmp.Object $path            
        }
    }
}

function Export-SelectedDeviceConfiguration
{
    param($path = "$env:Temp")

    Export-SingleDeviceConfiguration $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleDeviceConfiguration
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }        
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-DeviceConfigurationFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.displayName)"
        $obj = Invoke-GraphRequest -Url "/deviceManagement/deviceConfigurations/$($psObj.id)?`$expand=assignments"
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

function Copy-DeviceConfiguration
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect device configuration item you want to copy", "Error", "OK", "Error") 
        return 
    }

    $ret = Show-InputDialog "Copy device configuration" "Select name for the new object" "$($dgObjects.SelectedItem.displayName) - Copy"

    if($ret)
    {
        # Export profile
        Write-Status "Export $($dgObjects.SelectedItem.displayName)"
        # Convert to Json and back to clone the object
        $obj = ConvertTo-Json $dgObjects.SelectedItem.Object -Depth 5 | ConvertFrom-Json
        if($obj)
        {            
            # Import new profile
            $obj.displayName = $ret
            Import-DeviceConfiguration $obj | Out-Null

            $dgObjects.ItemsSource = @(Get-DeviceConfigurationObjects)
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

function Import-DeviceConfiguration
{
    param($obj)

    if(($obj | GM -MemberType NoteProperty -Name "supportsScopeTags"))
    {
        # Remove read-only property
        $obj.PSObject.Properties.Remove('supportsScopeTags')
    }

    Write-Status "Import $($obj.displayName)"

    Invoke-GraphRequest -Url "/deviceManagement/deviceConfigurations" -Content (ConvertTo-Json $obj -Depth 5) -HttpMethod POST        
}

function Import-AllDeviceConfigurationObjects
{
    param($path = "$env:Temp")    

    Import-DeviceConfigurationObjects (Get-JsonFileObjects $path)
}

function Import-DeviceConfigurationObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import  device configuration profiles"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import device configuration policy: $($obj.Object.displayName)"

        $assignments = Get-GraphAssignmentsObject $obj.Object ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_assignments.json")

        $response = Import-DeviceConfiguration $obj.Object
        if($response)
        {
            $global:importedObjects++
            Import-GraphAssignments $assignments "assignments" "/deviceManagement/deviceConfigurations/$($response.Id)/assign"
        }        
    }
    $dgObjects.ItemsSource = @(Get-DeviceConfigurationObjects)
    Write-Status ""
}