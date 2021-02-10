########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-GPOSettingName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-GPOSettings}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-GPOSettingName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all administrative templates"
                Import-AllGPOSettingObjects (Join-Path $rootFolder (Get-GPOSettingFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-GPOSettingName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all administrative templates"
                    Get-GPOSettingObjects | ForEach-Object { Export-SingleGPOSetting $PSItem.Object (Join-Path $rootFolder (Get-GPOSettingFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-GPOSettingFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-GPOSettingName
{
    return "Administrative Templates"
}


function Get-GPOSettingFolderName
{
    return "AdministrativeTemplates"
}

function Get-GPOSettings
{
    Write-Status "Loading administrative templates" 
    $dgObjects.ItemsSource = @(Get-GPOSettingObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllGPOSettings $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedGPOSetting $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllGPOSettingObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-GPOSettingObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -copy ([scriptblock]{Copy-GPOSetting}) -ViewFullObject ([scriptblock]{Get-GPOSettingObject $global:dgObjects.SelectedItem.Object})               
}

function Get-GPOSettingObjects
{
    Get-GraphObjects -Url "/deviceManagement/groupPolicyConfigurations"
}

function Get-GPOSettingObject
{
    param($object, $additional = "")

    if(-not $Object.id) { return }

    @((Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations/$($Object.id)$additional"),(Get-GPOObjectSettings $Object))
}

function Export-AllGPOSettings
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleGPOSetting $objTmp.Object $path            
        }
    }
}

function Export-SelectedGPOSetting
{
    param($path = "$env:Temp")

    Export-SingleGPOSetting $global:dgObjects.SelectedItem.Object $path
}

function Get-GPOObjectSettings
{
    param($GPOObj)

    $gpoSettings = @()

    # Get all configured policies in the Administrative Templates profile 
    $GPODefinitionValues = Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations/$($GPOObj.id)/definitionValues?`$expand=definition"
    foreach($definitionValue in $GPODefinitionValues.value)
    {
        # Get presentation values for the current settings (with presentation object included)
        $presentationValues = Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations/$($GPOObj.id)/definitionValues/$($definitionValue.id)/presentationValues?`$expand=presentation"

        # Set base policy settings
        $obj = @{
                "enabled" = $definitionValue.enabled
                "definition@odata.bind" = "$($global:graphURL)/deviceManagement/groupPolicyDefinitions('$($definitionValue.definition.id)')"
                }

        if($presentationValues.value)
        {
            # Policy presentation values set e.g. a drop down list, check box, text box etc.
            $obj.presentationValues = @()
                        
            $presentations = $null
            foreach ($presentationValue in $presentationValues.value) 
            {
                # Add presentation@odata.bind property that links the value to the presentation object
                $presentationValue | Add-Member -MemberType NoteProperty -Name "presentation@odata.bind" -Value "$($global:graphURL)/deviceManagement/groupPolicyDefinitions('$($definitionValue.definition.id)')/presentations('$($presentationValue.presentation.id)')"

                #Remove presentation object so it is not included in the export
                Remove-ObjectProperty $presentationValue "presentation"
                
                #Optional removes. Import will igonre them
                Remove-ObjectProperty $presentationValue "id"
                Remove-ObjectProperty $presentationValue "lastModifiedDateTime"
                Remove-ObjectProperty $presentationValue "createdDateTime"

                # Add presentation value to the list
                $obj.presentationValues += $presentationValue
            }
        }
        $gpoSettings += $obj
    }
    $gpoSettings
}

function Export-SingleGPOSetting
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-GPOSettingFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.displayName)"
        $obj = Invoke-GraphRequest -Url "deviceManagement/groupPolicyConfigurations/$($psObj.Id)?`$expand=assignments"

        if($obj)
        {            
            # Save Administrative Templates profile
            ConvertTo-Json $obj -Depth 5 | Out-File "$path\$((Remove-InvalidFileNameChars $obj.displayName)).json" -Force

            # Collect and save all the settings of the Administrative Templates profile
            $gpoSettings = Get-GPOObjectSettings $obj
            ConvertTo-Json $gpoSettings -Depth 5 | Out-File "$path\$($obj.displayName)_Settings.json" -Force

            # Export assignment info 
            Add-MigrationInfo $obj.assignments
        }
        $global:exportedObjects++
    }
}

function Copy-GPOSetting
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect administrative templates profile you want to copy", "Error", "OK", "Error") 
        return 
    }

    $ret = Show-InputDialog "Copy administrative template" "Select name for the new profile" "$($dgObjects.SelectedItem.displayName) - Copy"

    if($ret)
    {
        # Export profile 
        Write-Status "Export $($dgObjects.SelectedItem.displayName)"
        # Convert to Json and back to clone the object
        $obj = ConvertTo-Json $dgObjects.SelectedItem.Object -Depth 5 | ConvertFrom-Json
        if($obj)
        {            
            # Get the settings of the profile
            $gpoSettings = Get-GPOObjectSettings $obj

            # Import the new profile
            $obj.displayName = $ret
            Import-GPOSetting $obj $gpoSettings | Out-Null

            #Reload objects
            $dgObjects.ItemsSource = @(Get-GPOSettingObjects)
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

function Import-GPOSetting
{
    param($obj, $settings)

    Write-Status "Import $($obj.displayName)"

    Start-PreImport $obj

    # Import Administrative Template profile
    $response = Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations" -Content (ConvertTo-Json ($obj | Select-Object -Property * -ExcludeProperty createdDateTime, lastModifiedDateTime) -Depth 5) -HttpMethod POST
    
    if($response)
    {
        foreach($setting in $settings)
        {
            Start-PreImport $setting

            # Import each setting for the Administrative Template profile
            $response2 = Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations/$($response.id)/definitionValues" -Content (ConvertTo-Json $setting -Depth 5) -HttpMethod POST        
        }
    }

    $response
}

function Import-AllGPOSettingObjects
{
    param($path = "$env:Temp")    

    # Read json files and import all objects
    # Note: Each json file must match the object type being imported
    Import-GPOSettingObjects (Get-JsonFileObjects $path)
}

function Import-GPOSettingObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import administrative template profile"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import Administrative Template: $($obj.Object.displayName)"

        $gpoSettings = $null

        # Load settings from the <AdminTeplateName>_settings.json file
        $settingsFile = ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_settings.json")
        if(Test-Path $settingsFile)
        {
            $gpoSettings = (ConvertFrom-Json (Get-Content $settingsFile -Raw))
        }

        # Get assignment settings
        $assignments = Get-GraphAssignmentsObject $obj.Object ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_assignments.json")

        # Import Administrative Template object 
        $response = Import-GPOSetting $obj.Object $gpoSettings

        if($response)
        {
            $global:importedObjects++
            # Import assignments
            Import-GraphAssignments $assignments "assignments" "/deviceManagement/groupPolicyConfigurations/$($response.Id)/assign"
        }
    }

    #Reload list of objects
    $dgObjects.ItemsSource = @(Get-GPOSettingObjects)
    Write-Status ""
}
