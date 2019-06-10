########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-BaselineTemplatesName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-BaselineTemplates}
    })

    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-BaselineName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-BaselineProfiles}
    })

}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-AppProtectionName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all baseline policies"
                Import-AllBaselineProfileObjects  (Join-Path $rootFolder (Get-BaselineFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-AppProtectionName)
            Folder = (Get-BaselineFolderName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Import all baseline policies"
                    Get-BaselineProfileObjects | ForEach-Object { Export-SingleBaselineProfile $PSItem.Object (Join-Path $rootFolder (Get-BaselineFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-BaselineFolderName } 

}
########################################################
#
# Object specific functions
#
########################################################
function Get-BaselineTemplatesName
{
    return "Baseline Templates"
}


function Get-BaselineName
{
    return "Baseline Profiles"
}

function Get-BaselineFolderName
{
    return "Baseline"
}

function Get-BaselineTemplates
{
    Write-Status "Loading baseline templates" -SkipLog
    $dgObjects.ItemsSource = @(Get-BaselineTemplateObjects)
}

function Get-BaselineTemplateObjects
{
    Get-GraphObjects -Url "/deviceManagement/templates"
}

function Get-BaselineProfiles
{
    Write-Status "Loading banding profiles" -SkipLog
    $dgObjects.ItemsSource = @(Get-BaselineProfileObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllBaselineProfiles $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedBaselineProfile $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllBaselineProfileObjects $global:txtImportPath.Text        
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-BaselineProfileObjects $global:lstFiles.ItemsSource -Selected        
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude @("*_Settings.json","*_assignments.json"))
    }
    
    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -copy ([scriptblock]{Copy-BaselineProfile})             
}

function Get-BaselineProfileObjects
{
    Get-GraphObjects -Url "/deviceManagement/intents"
}

function Export-AllBaselineProfiles
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleBaselineProfile $objTmp.Object $path            
        }
    }
}

function Export-SelectedBaselineProfile
{
    param($path = "$env:Temp")

    Export-SingleBaselineProfile $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleBaselineProfile
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-BaselineFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.displayName)"
        $obj = $psObj
        if($obj)
        {            
            $fileName = "$path\$((Remove-InvalidFileNameChars $obj.displayName)).json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force
            $settings = Invoke-GraphRequest -Url "/deviceManagement/intents/$($obj.id)/settings"
            ConvertTo-Json $settings.value -Depth 5 | Out-File "$path\$($obj.displayName)_Settings.json" -Force
        }
        $assignments = Invoke-GraphRequest -Url "/deviceManagement/intents/$($obj.id)/assignments"
        if(($assignments.Value | measure).Count -gt 0)
        {
            ConvertTo-Json $assignments.value -Depth 5| Out-File "$path\$($obj.displayName)_assignments.json" -Force
        }
        $global:exportedObjects++
    }
}

function Copy-BaselineProfile
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect baseline profile you want to copy", "Error", "OK", "Error") 
        return 
    }

    $ret = Show-InputDialog "Copy baseline profiles" "Select name for the new object" "$($dgObjects.SelectedItem.displayName) - Copy"

    if($ret)
    {
        # Export profile
        Write-Status "Export $($dgObjects.SelectedItem.displayName)"
        # Convert to Json and back to clone the object
        $obj = ConvertTo-Json $dgObjects.SelectedItem.Object -Depth 5 | ConvertFrom-Json
        $settings = Invoke-GraphRequest -Url "/deviceManagement/intents/$($obj.id)/settings"
        $intentSettings = ConvertTo-Json $settings.value -Depth 5

        if($obj)
        {            
            # Import new profile
            $obj.displayName = $ret
            Import-BaselineProfile $obj $intentSettings | Out-null

            $dgObjects.ItemsSource = @(Get-BaselineProfileObjects)            
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

function Import-BaselineProfile
{
    param($obj, $intentSettings, $templateId)

$json = @"
    {
        "displayName": "$($obj.displayName)",
        "description": "$($obj.description)",
        "settingsDelta": 
        $($intentSettings)
  
    }
"@

    if($templateId)
    {
        $tempId = $templateId
    }
    else
    {
        $tempId = $obj.templateId
    }

    Write-Status "Import $($obj.displayName)"

    return Invoke-GraphRequest -Url "/deviceManagement/templates/$($tempId)/createInstance" -Content $json -HttpMethod POST
}

function Import-AllBaselineProfileObjects
{
    param(
        $path = "$env:Temp"
    )    

    Import-BaselineProfileObjects (Get-JsonFileObjects $path)
}

function Import-BaselineProfileObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import terms and conditions"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import security baseline: $($obj.Object.displayName)"

        $assignments = Get-GraphAssignmentsObject $obj.Object ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_assignments.json")

        $settingsFile = $obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_settings.json"
        if(-not (Test-Path $settingsFile)) { continue }

        $intentSettings = Get-Content $settingsFile -Raw

        $response = Import-BaselineProfile $obj.Object $intentSettings
        if($response)
        {
            $global:importedObjects++
            Import-GraphAssignments $assignments "assignments" "/deviceManagement/intents/$($response.Id)/assign"
        }
    }
    $dgObjects.ItemsSource = @(Get-BaselineProfileObjects)
    Write-Status ""
}
