########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-AppProtectionName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-AppProtections}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-AppProtectionName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all app protection/configuration policies"
                Import-AllAppProtectionObjects (Join-Path $rootFolder (Get-AppProtectionFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-AppProtectionName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all app protection/configuration policies"
                    Get-AppProtectionObjects | ForEach-Object { Export-SingleAppProtection $PSItem.Object (Join-Path $rootFolder (Get-AppProtectionFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-AppProtectionFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-AppProtectionName
{
    return "App Protection/Configuration"
}

function Get-AppProtectionFolderName
{
    return "AppProtection"
}

function Get-AppProtections
{
    Write-Status "Loading app protections and configurations" 
    $dgObjects.ItemsSource = @(Get-AppProtectionObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllAppProtections $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedAppProtection $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllAppProtectionObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-AppProtectionObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -copy ([scriptblock]{Copy-AppProtection}) -ViewFullObject ([scriptblock]{Get-AppProtectionObject $global:dgObjects.SelectedItem.Object}) -ForceFullObject
}

function Get-AppProtectionObjects
{    
    Get-GraphObjects -Url "/deviceAppManagement/managedAppPolicies"
}

function Get-AppProtectionObject
{
    param($object, $additional = "")

    if(-not $object.id) { return }

    $objType = Get-AppProtectionObjectType $object."@odata.type"

    $expand = ""
    if($objType -eq "targetedManagedAppConfigurations")
    { 
        $expand = "?`$expand=Apps" 
    }

    if($objType)
    {
        Invoke-GraphRequest -Url "/deviceAppManagement/$objType/$($object.id)$($expand)"
    }
}

function Export-AllAppProtections
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleAppProtection $objTmp.Object $path            
        }
    }
}

function Export-SelectedAppProtection
{
    param($path = "$env:Temp")

    Export-SingleAppProtection $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleAppProtection
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-AppProtectionFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.displayName)"        
        
        $obj = Get-AppProtectionObjectForExport $psObj

        if($obj)
        {            
            $fileName = "$path\$((Remove-InvalidFileNameChars $obj.displayName)).json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force

            Add-MigrationInfo $obj.assignments
        }
        $global:exportedObjects++
    }
}

function Get-AppProtectionObjectType
{
    param($odataType)

    if($odataType -like "*targetedManagedAppConfiguration*")
    { 
        "targetedManagedAppConfigurations"
         
    }
    elseif($odataType -like "*iosManagedAppProtection*")        
    { 
        "iosManagedAppProtections"            
    }
    elseif($odataType -like "*androidManagedAppProtection*")        
    { 
        "androidManagedAppProtections"            
    }
}

function Get-AppProtectionObjectForExport
{
    param($obj)

    $objType = Get-AppProtectionObjectType $obj."@odata.type"

    $expand = "?`$expand=assignments"
    if($objType -eq "targetedManagedAppConfigurations")
    { 
        $expand += ",Apps" 
    }

    if($objType)
    {
        Invoke-GraphRequest -Url "/deviceAppManagement/$objType/$($obj.id)$($expand)"
    }
}

function Copy-AppProtection
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect app protection/configuration item you want to copy", "Error", "OK", "Error") 
        return 
    }

    $ret = Show-InputDialog "Copy app protection/configuration" "Select name for the new policy" "$($dgObjects.SelectedItem.displayName) - Copy"

    if($ret)
    {
        # Export profile
        Write-Status "Export $($dgObjects.SelectedItem.displayName)"
        $obj = Get-AppProtectionObjectForExport $dgObjects.SelectedItem.Object
        if($obj)
        {       
            # Remove assignment properties
            Remove-ObjectProperty $obj "assignments"
            Remove-ObjectProperty $obj "assignments@odata.context"
             
            # Import new profile
            $obj.displayName = $ret
            Import-AppProtection $obj | Out-Null

            $dgObjects.ItemsSource = @(Get-AppProtectionObjects)
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

function Import-AppProtection
{
    param($obj)

    if(($obj | GM -MemberType NoteProperty -Name "Apps"))    
    {        
        $apps = $obj.Apps
        # Remove apps properties
        Remove-ObjectProperty $obj "apps"
        Remove-ObjectProperty $obj "apps@odata.context"
    }

    Write-Status "Import $($obj.displayName)"

    $objType = Get-AppProtectionObjectType $obj."@odata.context"

    if($objType)
    {
        #Import the app configuration policy
        $response = Invoke-GraphRequest -Url "/deviceAppManagement/$objType" -Content (ConvertTo-Json $obj -Depth 5) -HttpMethod POST
        if($response -and $apps)
        {
            # Import targeted apps
            $response2 = Invoke-GraphRequest -Url "/deviceAppManagement/$objType/$($response.Id)/targetApps" -Content "{ apps: $(ConvertTo-Json $apps -Depth 5)}" -HttpMethod POST
        }
        $response
    }
}

function Import-AllAppProtectionObjects
{
    param($path = "$env:Temp")    

    Import-AppProtectionObjects (Get-JsonFileObjects $path)
}

function Import-AppProtectionObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import app protection/configuration policies"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import App Protection/Configuration: $($obj.Object.displayName)"

        $assignments = Get-GraphAssignmentsObject $obj.Object ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_assignments.json")

        $response = Import-AppProtection $obj.Object

        if($response)
        {
            $global:importedObjects++

            $dataType = Get-AppProtectionObjectType $response."@odata.context"

            if($dataType)
            {
                Import-GraphAssignments $assignments "assignments" "/deviceAppManagement/$dataType/$($response.Id)/assign"
            }         
        }
    }
    $dgObjects.ItemsSource = @(Get-AppProtectionObjects)
    Write-Status ""
}