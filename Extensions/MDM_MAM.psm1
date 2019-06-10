########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-MDMMAMName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-MDMMAM}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-MDMMAMName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all MDM/MAM setting"
                Import-AllMDMMAMObjects (Join-Path $rootFolder (Get-MDMMAMFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-MDMMAMName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all MDM/MAM settings"
                    Get-MDMMAMObjects | ForEach-Object { Export-SingleMDMMAM $PSItem.Object (Join-Path $rootFolder (Get-MDMMAMFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-MDMMAMFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-MDMMAMName
{
    return "MDM/MAM"
}

function Get-MDMMAMFolderName
{
    return "MDMMAM"
}

function Get-MDMMAM
{
    Write-Status "Loading MDM/MAM object" 
    $dgObjects.ItemsSource = @(Get-MDMMAMObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllMDMMAM $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedMDMMAM $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })
    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllMDMMAMObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-MDMMAMObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles})            
}

function Get-MDMMAMObjects
{
    Get-AzureNativeObjects "MdmApplications" -property @('appDisplayName')
}

function Export-AllMDMMAM
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleMDMMAM $objTmp.Object $path            
        }
    }
}

function Export-SelectedMDMMAM
{
    param($path = "$env:Temp")

    Export-SingleMDMMAM $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleMDMMAM
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-MDMMAMFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.appDisplayName)"
        
        $obj = Invoke-AzureNativeRequest "MdmApplications/$($psObj.objectId)"

        if($obj)
        {            
            $fileName = "$path\$((Remove-InvalidFileNameChars $obj.appDisplayName)).json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force
        }
        
        if($obj.mdmAppliesToGroups)
        {
            $obj.mdmAppliesToGroups | ForEach-Object { Add-GroupMigrationObject $PSItem.objectId }
        }

        if($obj.mamAppliesToGroups)
        {
            $obj.mamAppliesToGroups | ForEach-Object { Add-GroupMigrationObject $PSItem.objectId }
        }

        $global:exportedObjects++
    }
}

function Import-MDMMAM
{
    param($obj)

    $argStr = "?"
    if($obj.enrollmentUrl) { $argStr += "mdmAppliesToChanged=true" }
    else{ $argStr += "mdmAppliesToChanged=false" }
    if($obj.mamEnrollmentUrl) { $argStr += "&mamAppliesToChanged=true" }
    else{ $argStr += "&mamAppliesToChanged=false" }

    $response = Invoke-AzureNativeRequest "MdmApplications/$($obj.objectId)$argStr" -Method PUT -Body (Update-JsonForEnvironment (ConvertTo-Json $obj -Depth 5))
}

function Import-AllMDMMAMObjects
{
    param($path = "$env:Temp")    

    Import-MDMMAMObjects (Get-JsonFileObjects $path)
}

function Import-MDMMAMObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import MDM/MAM settings"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import MDM/MAM app settings: $($obj.Object.appDisplayName)"

        Import-MDMMAM $obj.Object

        # No assignments for MDM/MAM
    }
    $dgObjects.ItemsSource = @(Get-MDMMAMObjects)
    Write-Status ""
}