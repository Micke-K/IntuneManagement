########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-CompliancePolicyName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-CompliancePolicies}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-CompliancePolicyName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all Intune compliance policies"
                Import-AllCompliancePolicyObjects (Join-Path $rootFolder (Get-CompliancePolicyFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-CompliancePolicyName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all compliance policies"
                    Get-CompliancePolicyObjects | ForEach-Object { Export-SingleCompliancePolicy $PSItem.Object (Join-Path $rootFolder (Get-CompliancePolicyFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-CompliancePolicyFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-CompliancePolicyName
{
    return "Compliance Policies"
}

function Get-CompliancePolicyFolderName
{
    return "CompliancePolicies"
}

function Get-CompliancePolicies
{
    Write-Status "Loading compliance policies" 
    $dgObjects.ItemsSource = @(Get-CompliancePolicyObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
   $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllCompliancePolicies $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedCompliancePolicy $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllCompliancePolicyObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-CompliancePolicyObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles}) -copy ([scriptblock]{Copy-CompliancePolicy})                
}

function Get-CompliancePolicyObjects
{
    Get-GraphObjects -Url "/deviceManagement/deviceCompliancePolicies"
}

function Export-AllCompliancePolicies
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleCompliancePolicy $objTmp.Object $path            
        }
    }
}

function Export-SelectedCompliancePolicy
{
    param($path = "$env:Temp")

    Export-SingleCompliancePolicy $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleCompliancePolicy
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-CompliancePolicyFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.displayName)"
        $obj = Invoke-GraphRequest -Url "/deviceManagement/deviceCompliancePolicies/$($psObj.id)?`$expand=assignments"
        if($obj)
        {            
            $fileName = "$path\$((Remove-InvalidFileNameChars $obj.displayName)).json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force

            Add-MigrationInfo $obj.assignments
        }
        $global:exportedObjects++
    }
}

function Copy-CompliancePolicy
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect compliance policy item you want to copy", "Error", "OK", "Error") 
        return 
    }

    $ret = Show-InputDialog "Copy compliance policy" "Select name for the new object" "$($dgObjects.SelectedItem.displayName) - Copy"

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
            Import-CompliancePolicy $obj | Out-null      

            $dgObjects.ItemsSource = @(Get-CompliancePolicyObjects)
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

function Import-CompliancePolicy
{
    param($obj)

    $json = ConvertTo-Json $obj -Depth 5
    $json = $json.Trim().TrimEnd('}').Trim()
    $json += @"
,
    "scheduledActionsForRule":[{"ruleName":"PasswordRequired","scheduledActionConfigurations":[{"actionType":"block","gracePeriodHours":0,"notificationTemplateId":"","notificationMessageCCList":[]}]}]
}

"@

    Write-Status "Import $($obj.displayName)"

    Invoke-GraphRequest -Url "/deviceManagement/deviceCompliancePolicies" -Content $json -HttpMethod POST        
}

function Import-AllCompliancePolicyObjects
{
    param($path = "$env:Temp")    

    Import-CompliancePolicyObjects (Get-JsonFileObjects $path)
}

function Import-CompliancePolicyObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import compliance policies"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        Write-Log "Import Compliance Policy: $($obj.Object.displayName)"

        $assignments = Get-GraphAssignmentsObject $obj.Object ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_assignments.json")

        $response = Import-CompliancePolicy $obj.Object
        if($response)
        {
            $global:importedObjects++
            Import-GraphAssignments $assignments "assignments" "/deviceManagement/deviceCompliancePolicies/$($response.Id)/assign"
        }        
    }
    $dgObjects.ItemsSource = @(Get-CompliancePolicyObjects)
    Write-Status ""
}