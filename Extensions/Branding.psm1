########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-IntuneBrandingName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-IntuneBrandings}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-IntuneBrandingName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all Intune branding objects"
                Import-AllIntuneBrandingObjects (Join-Path $rootFolder (Get-IntuneBrandingFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-IntuneBrandingName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all Intune branding objects"
                    Get-IntuneBrandingObjects | ForEach-Object { Export-SingleIntuneBranding $PSItem.Object (Join-Path $rootFolder (Get-IntuneBrandingFolderName)) }                    
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-IntuneBrandingFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-IntuneBrandingName
{
    return "Intune Branding"
}

function Get-IntuneBrandingFolderName
{
    return "IntuneBranding"
}

function Get-IntuneBrandings
{
    Write-Status "Loading banding profiles" 
    $dgObjects.ItemsSource = @(Get-IntuneBrandingObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllIntuneBrandings $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    # Same as ExportAllScript since only one object is supported
    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-AllIntuneBrandings $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllIntuneBrandingObjects $global:txtImportPath.Text        
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-IntuneBrandingObjects $global:lstFiles.ItemsSource -Selected        
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")        
    }
    
    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles})             
}

function Get-IntuneBrandingObjects
{
    Get-GraphObjects -Url "/deviceManagement/intuneBrand" -property @("displayName")
}

function Export-AllIntuneBrandings
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleIntuneBranding $objTmp.Object $path            
        }
    }
}

function Export-SingleIntuneBranding
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-IntuneBrandingFolderName) }
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
            Save-IntuneBrandingFile $obj "lightBackgroundLogo" $path
            Save-IntuneBrandingFile $obj "darkBackgroundLogo" $path
            Save-IntuneBrandingFile $obj "landingPageCustomizedImage" $path
        }
        $global:exportedObjects++
    }
}

function Save-IntuneBrandingFile
{
    param($obj, $prop, $path)

    if(-not $obj.$prop.type) { return }

    $arr=$obj.$prop.type.Split('/')
    if($arr.Length -gt 1)
    {
        $fileType = $arr[1]
    }
    else
    {
        $fileType = ".jpg" # assume...
    }

    $fileName = "$path\$((Remove-InvalidFileNameChars "$($obj.displayName).$prop.$fileType"))"
    try
    {
        if(Test-Path $fileName)
        {
            Remove-Item -Path $fileName -Force
        }
        [IO.File]::WriteAllBytes($fileName, [System.Convert]::FromBase64String($obj.$prop.value))
    }
    catch {}
}


function Import-IntuneBranding
{
    param($obj)

    Start-PreImport $obj -RemoveProperties @("@odata.context")
    
    $newObject = @"
{
    "intuneBrand":$((ConvertTo-Json $obj -Depth 5))
}

"@
    Write-Status "Import $($obj.displayName)"

    # Note: Branding is imported to deviceManagement with JSON parent object intuneBrand
    Invoke-GraphRequest -Url "$URL/deviceManagement" -Content $newObject -HttpMethod PATCH
}

function Import-AllIntuneBrandingObjects
{
    param($path = "$env:Temp")    

    Import-IntuneBrandingObjects (Get-JsonFileObjects $path)
}

function Import-IntuneBrandingObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import Intune branding"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import Intune branding"
        
        $response = Import-IntuneBranding $obj.Object

        # Note: No assignments for branding. This is default branding for everyone

    }
    $dgObjects.ItemsSource = @(Get-IntuneBrandingObjects)
    Write-Status ""
}