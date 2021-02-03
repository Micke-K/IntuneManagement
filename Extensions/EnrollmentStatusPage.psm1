########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-ESPName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-ESPs}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-ESPName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all enrollment status page settings"
                Import-AllESPObjects (Join-Path $rootFolder (Get-ESPFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-ESPName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all enrollment status page settings" 
                    Get-ESPObjects | ForEach-Object { Export-SingleESP $PSItem.Object (Join-Path $rootFolder (Get-ESPFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-ESPFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-ESPName
{
    return "Enrollment Status Page"
}


function Get-ESPFolderName
{
    return "EnrollmentStatusPage"
}

function Get-ESPs
{
    Write-Status "Loading enrollment status page objects" 
    $dgObjects.ItemsSource = @(Get-ESPObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllESPs $global:txtExportPath.Text
            Set-ObjectGrid
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedESP $global:txtExportPath.Text
            Set-ObjectGrid
            Write-Status ""
        })
    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllESPObjects $global:txtExportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-ESPObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -copy ([scriptblock]{Copy-ESP}) -ViewFullObject ([scriptblock]{Get-ESPObject $global:dgObjects.SelectedItem.Object})               
}

function Get-ESPObjects
{
    Get-GraphObjects -Url "/deviceManagement/deviceEnrollmentConfigurations"
}

function Get-ESPObject
{
    param($object, $additional = "")

    if(-not $Object.id) { return }

    Invoke-GraphRequest -Url "/deviceManagement/deviceEnrollmentConfigurations/$($Object.id)$additional"
}

function Export-AllESPs
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleESP $objTmp.Object $path            
        }
    }
}

function Export-SelectedESP
{
    param($path = "$env:Temp")

    Export-SingleESP $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleESP
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-ESPFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.displayName)"
        $obj = Invoke-GraphRequest -Url "/deviceManagement/deviceEnrollmentConfigurations/$($psObj.id)" #?`$expand=assignments"
        if($obj)
        {            
            if($obj.id -like "*_default*")
            {
                $idx = $obj.id.ToLower().IndexOf("_default")                
                $baseName = "Default_" + $obj.id.SubString($idx + "_default".Length)
            }
            else
            {
                # ?`$expand=assignments is not working so get assignments
                $assignments = Invoke-GraphRequest -Url "/deviceManagement/deviceEnrollmentConfigurations/$($obj.id)/assignments"
                if($assignments.value)
                {
                    $obj | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments.value                    
                }
                $baseName = Remove-InvalidFileNameChars $obj.displayName
            }
            $fileName = "$path\$baseName.json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force

            Add-MigrationInfo $obj.assignments
        }
        $global:exportedObjects++
    }
}

function Copy-ESP
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect enrollment status page item you want to copy", "Error", "OK", "Error") | Out-Null
        return 
    }

    if($dgObjects.SelectedItem.Object.id -like "*_default*")
    {
        [System.Windows.MessageBox]::Show("You cannot copy default items`n`nSelect custom entrollment status page item", "Error", "OK", "Error") | Out-Null
        return 
    }

    $ret = Show-InputDialog "Copy enrollment status page" "Select name for the new object" "$($dgObjects.SelectedItem.displayName) - Copy"

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
            Import-ESP $obj | Out-Null

            $dgObjects.ItemsSource = @(Get-ESPObjects)
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

function Import-ESP
{
    param($obj)

    Start-PreImport $obj

    if($obj.id -like "*_default*")
    {
        Write-Status "Update $($obj.displayName)"

        Invoke-GraphRequest -Url "/deviceManagement/deviceEnrollmentConfigurations/$($obj.id)" -Content (ConvertTo-Json $obj -Depth 5) -HttpMethod PATCH        
    }
    else
    {
        Write-Status "Import $($obj.displayName)"

        Invoke-GraphRequest -Url "/deviceManagement/deviceEnrollmentConfigurations" -Content (ConvertTo-Json $obj -Depth 5) -HttpMethod POST
    }
}

function Import-AllESPObjects
{
    param($path = "$env:Temp")    

    Import-ESPObjects (Get-JsonFileObjects $path)
}

function Import-ESPObjects
{
    param(
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import enrollment status page"

    foreach($obj in $Objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        if($obj.Object.id -like "*_default*")
        {
            $idx = $obj.Object.id.ToLower().IndexOf("_default")                
            $extInfo = " ($($obj.Object.id.SubString($idx + "_default".Length)))"
        }
        else
        {
            $extInfo = ""
        }

        Write-Log "Import Enrollment Status Page: $($obj.Object.displayName)$extInfo"

        $assignments = Get-GraphAssignmentsObject $obj.Object ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_assignments.json")

        $response = Import-ESP $obj.Object

        if($response)
        {
            $global:importedObjects++
            Import-GraphAssignments $assignments "enrollmentConfigurationAssignments" "/deviceManagement/deviceEnrollmentConfigurations/$($response.Id)/assign"
        }        
    }

    $dgObjects.ItemsSource = @(Get-ESPObjects)
    Write-Status ""
}