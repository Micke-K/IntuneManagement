########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-AutoPilotName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-AutoPilots}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-AutoPilotName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all AutoPilot policies"
                Import-AllAutoPilotObjects (Join-Path $rootFolder (Get-AutoPilotFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-AutoPilotName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all AutoPilot policies"
                    Get-AutoPilotObjects | ForEach-Object { Export-SingleAutoPilotObject $PSItem.Object (Join-Path $rootFolder (Get-AutoPilotFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-AutoPilotFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-AutoPilotName
{
    (Get-AutoPilotFolderName)
}

function Get-AutoPilotFolderName
{
    "AutoPilot"
}

function Get-AutoPilots
{
    Write-Status "Loading AutoPilot profiles" 
    $dgObjects.ItemsSource = @(Get-AutoPilotObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllAutoPilots $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedAutoPilotObject $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllAutoPilotObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-AutoPilotObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -copy ([scriptblock]{Copy-AutoPilot}) -ViewFullObject ([scriptblock]{Get-AutoPilotObject $global:dgObjects.SelectedItem.Object})                
}

function Get-AutoPilotObjects
{
    Get-GraphObjects -Url "/deviceManagement/windowsAutopilotDeploymentProfiles"
}

function Get-AutoPilotObject
{
    param($object, $additional = "")

    if(-not $Object.id) { return }

    Invoke-GraphRequest -Url "/deviceManagement/windowsAutopilotDeploymentProfiles/$($Object.id)$additional"
}

function Export-AllAutoPilots
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleAutoPilotObject $objTmp.Object $path            
        }
    }
}

function Export-SelectedAutoPilotObject
{
    param($path = "$env:Temp")

    Export-SingleAutoPilotObject $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleAutoPilotObject
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-AutoPilotFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.displayName)"
        $obj = Invoke-GraphRequest -Url "/deviceManagement/windowsAutopilotDeploymentProfiles/$($psObj.id)?`$expand=assignments"
        if($obj)
        {            
            $fileName = "$path\$((Remove-InvalidFileNameChars $obj.displayName)).json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force

            Add-MigrationInfo $obj.assignments
        }
        $global:exportedObjects++
    }
}

function Copy-AutoPilot
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect AutoPilot item you want to copy", "Error", "OK", "Error") 
        return 
    }

    $ret = Show-InputDialog "Copy AutoPilot" "Select name for the new object" "$($dgObjects.SelectedItem.displayName) Copy"

    if($ret)
    {
        # Export profile
        Write-Status "Export $($dgObjects.SelectedItem.displayName)"
        # Convert to Json and back to clone the object
        $obj = ConvertTo-Json $dgObjects.SelectedItem.Object -Depth 5 | ConvertFrom-Json
        if($obj)
        {            
            # Import new profile
            $obj.displayName = Remove-InvalidFileNameChars $ret
            Import-AutoPilot $obj | Out-Null  

            $dgObjects.ItemsSource = @(Get-AutoPilotObjects)
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

function Import-AutoPilot
{
    param($obj)

    Write-Status "Import $($obj.displayName)"

    Start-PreImport $obj 

    Invoke-GraphRequest -Url "/deviceManagement/windowsAutopilotDeploymentProfiles" -Content (ConvertTo-Json $obj -Depth 5) -HttpMethod POST        
}

function Import-AllAutoPilotObjects
{
    param(
        $path = "$env:Temp"
    )    

    Import-AutoPilotObjects (Get-JsonFileObjects $path)
}

function Import-AutoPilotObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import AutoPilot profiles"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import AutoPilot profile: $($obj.Object.displayName)"

        $assignments = Get-GraphAssignmentsObject $obj.Object ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_assignments.json")

        $response = Import-AutoPilot $obj.Object

        if($response)
        {
            $global:importedObjects++
            Import-GraphAssignments2 $assignments "/deviceManagement/windowsAutopilotDeploymentProfiles/$($response.Id)/assignments"
        }               
    }
    $dgObjects.ItemsSource = @(Get-AutoPilotObjects)
    Write-Status ""
}