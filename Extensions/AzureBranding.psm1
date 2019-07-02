########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-AZBrandingName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-AZBrandings}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-AZBrandingName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all Azure branding"
                Import-AllAZBrandingObjects (Join-Path $rootFolder (Get-AZBrandingFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-AZBrandingName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all Azure branding"
                    Get-AZBrandingObjects | ForEach-Object { Export-SingleAZBranding $PSItem.Object (Join-Path $rootFolder (Get-AZBrandingFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-AZBrandingFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-AZBrandingName
{
    return "Azure Branding"
}

function Get-AZBrandingFolderName
{
    return "AZBranding"
}

function Get-AZBrandings
{
    Write-Status "Loading Azure brandings" 
    $dgObjects.ItemsSource = @(Get-AZBrandingObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllAZBrandings $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedAZBranding $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    $script:exportParams.Add("DisplayColumn", "localeDisplayName")

    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllAZBrandingObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-AZBrandingObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -ViewFullObject ([scriptblock]{Get-AZBrandingObject $global:dgObjects.SelectedItem.Object})            
}

function Get-AZBrandingObjects
{
    $response =  Get-AzureNativeObjects "LoginTenantBrandings" -property @('locale', 'localeDisplayName')
    if($response)
    {
        $response | Where { $_.Object.isConfigured -eq $true }
    }
}

function Get-AZBrandingObject
{
    param($object, $additional = "")

    if(-not $Object.locale) { return }

    Invoke-AzureNativeRequest "LoginTenantBrandings/$($Object.locale)$additional"
}

function Export-AllAZBrandings
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleAZBranding $objTmp.Object $path            
        }
    }
}

function Export-SelectedAZBranding
{
    param($path = "$env:Temp")

    Export-SingleAZBranding $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleAZBranding
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-AZBrandingFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.localeDisplayName)"
        $obj = Invoke-AzureNativeRequest "LoginTenantBrandings/$($psObj.locale)"
        if($obj)
        {            
            $fileName = "$path\$((Remove-InvalidFileNameChars $obj.localeDisplayName)).json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force

            Save-AzureBrandingFile $obj "tileLogoUrl" $path
            Save-AzureBrandingFile $obj "bannerLogoUrl" $path
            Save-AzureBrandingFile $obj "illustrationUrl" $path
            Save-AzureBrandingFile $obj "squareLogoDarkUrl" $path

            $global:exportedObjects++
        }
        Set-ObjectPath $global:txtExportPath.Text
    }
}

function Save-AzureBrandingFile
{
    param($obj, $prop, $path)

    if(-not $obj.$prop) { return }

    $arr=$obj.$prop.Split('.')
    if($arr.Length -ne 1)
    {
        return
    }
    $fileType = "jpg"  # Assume...not OK. $arr[0] contains information about what kind of file it is

    $fileName = "$path\$((Remove-InvalidFileNameChars "$($obj.localeDisplayName).$prop.$fileType"))"
    try
    {
        if(Test-Path $fileName)
        {
            Remove-Item -Path $fileName -Force
        }
        [IO.File]::WriteAllBytes($fileName, [System.Convert]::FromBase64String($arr[1]))
    }
    catch {}
}

function Import-AZBranding
{
    param($obj)

    if($global:runningBulkImport -eq $true)
    {
        # Update Default and create the rest...
        $createNew = $obj.locale -ne 0 
    }
    else
    {
        $curObj = $global:lstFiles.ItemsSource | Where { $_.Object.locale -eq $obj.locale }

        if($curObj -and $obj.locale -ne 0)
        {
            return # Do not update existing object except default
        }
        elseif(-not $curObj)
        {
            $createNew = $true
        }
        else
        {
            $createNew = $false
        }
    }

    $json = "{"

    if($createNew) { $json += "`"locale`":`"$($obj.locale)`"," }

    if($obj.signInUserIdLabel) { $json += "`"userIdLabel`":  `"$($obj.signInUserIdLabel)`"," }
    if($obj.signInPageText) { $json += "`"boilerPlateText`":  `"$($obj.signInPageText)`"," }
    if($obj.signInBackColor) { $json += "`"backgroundColor`":  `"$($obj.signInBackColor)`"," }
    if($obj.tileLogoUrl) { $json += "`"tileLogoUrl`": `"$($obj.tileLogoUrl)`"," }
    if($obj.bannerLogoUrl) { $json += "`"bannerLogoUrl`": `"$($obj.bannerLogoUrl)`"," }
    if($obj.illustrationUrl) { $json += "`"illustrationUrl`": `"$($obj.illustrationUrl)`"," }
    if($obj.squareLogoDarkUrl) { $json += "`"squareLogoDarkUrl`": `"$($obj.squareLogoDarkUrl)`"," }    

    if($obj.hideKeepMeSignedIn -and $obj.locale -eq 0)  { $json += "`"keepMeSignedInDisabled`":  $($obj.hideKeepMeSignedIn.ToString().ToLower())," }
        
    if($createNew)
    {
        if($curObj.bannerLogoUrl -ne $curObj.bannerLogoUrl)
        {
            $json += "`"isTileLogoUpdated`":true,"
        }

        if($curObj.illustrationUrl -ne $curObj.illustrationUrl)
        {
            $json += "`"isIllustrationImageUpdated`":true,"
        }

        if($curObj.squareLogoDarkUrl -ne $curObj.squareLogoDarkUrl)
        {
            $json += "`"isSquareDarkLogoUpdated`":true,"
        }

        if($curObj.bannerLogoUrl -ne $curObj.bannerLogoUrl)
        {
            $json += "`"isBannerLogoUpdated`":true,"
        }
    }

    $json = $json.TrimEnd(',')
    $json += "}"

    Write-Status "Import $($obj.localeDisplayName)"

    if($createNew)
    {
        Invoke-AzureNativeRequest "LoginTenantBrandings" -Method POST -Body $json | Out-Null
    }
    else
    {
        Invoke-AzureNativeRequest "LoginTenantBrandings/$($obj.locale)" -Method PATCH -Body $json  | Out-Null
    }
}

function Import-AllAZBrandingObjects
{
    param($path = "$env:Temp")    

    Import-AZBrandingObjects (Get-JsonFileObjects $path)
}

function Import-AZBrandingObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import Azure brandings"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import Azure branding"

        $response = Import-AZBranding $obj.Object

        if($response)
        {
            $global:importedObjects++
        }
    }
    $dgObjects.ItemsSource = @(Get-AZBrandingObjects)
    Write-Status ""
}