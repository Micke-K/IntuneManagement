########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-ConditionalAccessName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-ConditionalAccess}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-ConditionalAccessName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all conditional access policies"
                Import-AllConditionalAccessObjects (Join-Path $rootFolder (Get-ConditionalAccessFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-ConditionalAccessName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all conditional access policies"
                    Get-ConditionalAccessObjects | ForEach-Object { Export-SingleConditionalAccess $PSItem.Object (Join-Path $rootFolder (Get-ConditionalAccessFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-ConditionalAccessFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-ConditionalAccessName
{
    return "Conditional Access"
}

function Get-ConditionalAccessFolderName
{
    return "ConditionalAccess"
}

function Get-ConditionalAccess
{
    Write-Status "Loading conditional access objects" 
    $dgObjects.ItemsSource = @(Get-ConditionalAccessObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllConditionalAccess $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedConditionalAccess $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })
    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllConditionalAccessObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-ConditionalAccessObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -ViewFullObject ([scriptblock]{Get-ConditionalAccessObject $global:dgObjects.SelectedItem.Object})           
}

function Get-ConditionalAccessObjects
{
    #https://main.iam.ad.ext.azure.com/api/Policies/Policies?top=10&nextLink=null&appId=&includeBaseline=true
    Get-AzureNativeObjects "Policies/Policies?top=10&appId=&includeBaseline=true" -property @('policyName') -allowPaging
}

function Get-ConditionalAccessObject
{
    param($object, $additional = "")

    if(-not $Object.policyId) { return }

    if($Object.baselineType -eq 0)
    {
        Invoke-AzureNativeRequest "Policies/$($Object.policyId)$additional"
    }
    else
    {
        Invoke-AzureNativeRequest "BaselinePolicies/$($Object.policyId)$additional"
    }
}

function Export-AllConditionalAccess
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleConditionalAccess $objTmp.Object $path            
        }
    }
}

function Export-SelectedConditionalAccess
{
    param($path = "$env:Temp")

    Export-SingleConditionalAccess $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleConditionalAccess
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-ConditionalAccessFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.policyName)"
        
        if($psObj.baselineType -eq 0)
        {
            $obj = Invoke-AzureNativeRequest "Policies/$($psObj.policyId)"
        }
        else
        {
            $obj = Invoke-AzureNativeRequest "BaselinePolicies/$($psObj.policyId)"
        }

        if($obj)
        {            
            $fileName = "$path\$((Remove-InvalidFileNameChars $psObj.policyName)).json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force
        }

        if($jsonObj.usersV2.included.groupIds)
        {
            $jsonObj.usersV2.included.groupIds | ForEach-Object { Add-GroupMigrationObject $PSItem }
        }

        if($jsonObj.usersV2.excluded.groupIds)
        {
            $jsonObj.usersV2.excluded.groupIds | ForEach-Object { Add-GroupMigrationObject $PSItem }
        }

        if($jsonObj.usersV2.included.userIds -or $jsonObj.usersV2.excluded.userIds)
        {
            Write-Log "Users are specified in $($psObj.policyName). User are not supported in this version. This conditional access policy might not be imported" 2 
        }
        
        if($jsonObj.usersV2.included.roleIds -or $jsonObj.usersV2.excluded.roleIds)
        {
            Write-Log "Roles are specified in $($psObj.policyName). Roles are not supported in this version. This conditional access policy might not be imported" 2
        }

        if($jsonObj.conditions.namedNetworks.includedNetworkIds -or $jsonObj.conditions.namedNetworks.excludedNetworkIds)
        {
            Write-Log "Networks are specified in $($psObj.policyName). Named networks are not supported in this version. This conditional access policy might not be imported" 2
        }

        # There might be a lot more to check here...
        
        $global:exportedObjects++
    }
}

function Import-ConditionalAccess
{
    param($obj)

    Start-PreImport $obj 

    $json = Update-JsonForEnvironment $json

    if($obj.baselineType -eq 0)
    {
        $obj.policyId = ""
        $obj.isAllProtocolsEnabled = $true
        $json = ConvertTo-Json $obj -Depth 10
        $json = Update-JsonForEnvironment $json

        if((Invoke-AzureNativeRequest "Policies/Validate" -Method POST -Body $json) -eq 11)
        {
            Invoke-AzureNativeRequest "Policies" -Method POST -Body $json | Out-Null
        }
        else
        {
            Write-Log "Policy validation of json data failed" 3
        }
    }
    else
    {
        Write-Log "Conditional Access Baseline Policies does not support import"
        #Invoke-AzureNativeRequest "BaselinePolicies/$($obj.id)" -Method PUT -Body (ConvertTo-Json $obj -Depth 5)  | Out-Null
    }
}

function Import-AllConditionalAccessObjects
{
    param($path = "$env:Temp")    

    Import-ConditionalAccessObjects (Get-JsonFileObjects $path)
}

function Import-ConditionalAccessObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import conditional access policies"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import Conditional Access: $($obj.Object.policyName)"

        $response = Import-ConditionalAccess $obj.Object

        if($response)
        {
            $global:importedObjects++
        }
        # No additionl assignments on conditional access policies
    }
    $dgObjects.ItemsSource = @(Get-ConditionalAccessObjects)
    Write-Status ""
}

<#    
    # Get all networks
    Get-AzureNativeObjects "NamedNetworksV2"

    # Network example
    #{"networkName":"Australia","cidrIpRanges":[],"categories":[],"applyToUnknownCountry":false,"countryIsoCodes":["AU"],"isTrustedLocation":false,"namedLocationsType":2}

    Get-AzureNativeObjects "NamedNetworksV2" -Method POST -Body $json | Out-Nul

    # Get all contry codes
    NamedNetworksV2/CountryCodes
#>