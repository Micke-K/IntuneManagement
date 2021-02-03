########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-TermsAndConditionName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-TermsAndConditions}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-TermsAndConditionName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all terms and conditions"
                Import-AllTermsAndConditionObjects (Join-Path $rootFolder (Get-TermsAndConditionFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-TermsAndConditionName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all Intune terms and conditions"
                    Get-TermsAndConditionObjects | ForEach-Object { Export-SingleTermsAndCondition $PSItem.Object (Join-Path $rootFolder (Get-TermsAndConditionFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-TermsAndConditionFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-TermsAndConditionName
{
    return "Terms and Conditions"
}

function Get-TermsAndConditionFolderName
{
    return "TermsAndConditions"
}

function Get-TermsAndConditions
{
    Write-Status "Loading terms and conditions" 
    $dgObjects.ItemsSource = @(Get-TermsAndConditionObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllTermsAndConditions $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedTermsAndCondition $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllTermsAndConditionObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-TermsAndConditionObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -copy ([scriptblock]{Copy-TermsAndCondition}) -ViewFullObject ([scriptblock]{Get-TermsAndConditionObject $global:dgObjects.SelectedItem.Object})                  
}

function Get-TermsAndConditionObjects
{
    Get-GraphObjects -Url "/deviceManagement/termsAndConditions"
}

function Get-TermsAndConditionObject
{
    param($object, $additional = "")

    if(-not $Object.id) { return }

    Invoke-GraphRequest -Url "/deviceManagement/termsAndConditions/$($Object.id)$additional"
}

function Export-AllTermsAndConditions
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleTermsAndCondition $objTmp.Object $path            
        }
    }
}

function Export-SelectedTermsAndCondition
{
    param($path = "$env:Temp")

    Export-SingleTermsAndCondition $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleTermsAndCondition
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-TermsAndConditionFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.displayName)"
        $obj = Invoke-GraphRequest -Url "/deviceManagement/termsAndConditions/$($psObj.id)" #?`$expand=assignments"
        if($obj)
        {            
            # ?`$expand=assignments is not working so get assignments
            $assignments = Invoke-GraphRequest -Url "/deviceManagement/termsAndConditions/$($obj.id)/assignments"
            if($assignments.value)
            {
                $obj | Add-Member -NotePropertyName "assignments" -NotePropertyValue $assignments.value                    
            }

            $fileName = "$path\$((Remove-InvalidFileNameChars $obj.displayName)).json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force

            Add-MigrationInfo $obj.assignments
        }
        $global:exportedObjects++
    }
}

function Copy-TermsAndCondition
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect terms and conditions item you want to copy", "Error", "OK", "Error") 
        return 
    }

    $ret = Show-InputDialog "Copy terms and conditions" "Select name for the new object" "$($dgObjects.SelectedItem.displayName) - Copy"

    if($ret)
    {
        # Export profile
        Write-Status "Export $($dgObjects.SelectedItem.displayName)"        
        # Get full object for export
        $obj = Invoke-GraphRequest -Url "/deviceManagement/termsAndConditions/$($dgObjects.SelectedItem.Object.id)"        
        if($obj)
        {            
            # Import new profile
            $obj.displayName = $ret
            Import-TermsAndCondition $obj | Out-Null

            $dgObjects.ItemsSource = @(Get-TermsAndConditionObjects)
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

function Import-TermsAndCondition
{
    param($obj)

    Write-Status "Import $($obj.displayName)"

    Start-PreImport $obj 

    Invoke-GraphRequest -Url "/deviceManagement/termsAndConditions" -Content (ConvertTo-Json $obj -Depth 5) -HttpMethod POST        
}

function Import-AllTermsAndConditionObjects
{
    param($path = "$env:Temp")    

    Import-TermsAndConditionObjects (Get-JsonFileObjects $path)
}

function Import-TermsAndConditionObjects
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

        Write-Log "Import Terms and Conditions: $($obj.Object.displayName)"

        $assignments = Get-GraphAssignmentsObject $obj.Object ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_assignments.json")

        $response = Import-TermsAndCondition $obj.Object
        if($response)
        {
            $global:importedObjects++
            Import-GraphAssignments2 $assignments "/deviceManagement/termsAndConditions/$($response.Id)/assignments"
        }        
    }
    $dgObjects.ItemsSource = @(Get-TermsAndConditionObjects)
    Write-Status ""
}