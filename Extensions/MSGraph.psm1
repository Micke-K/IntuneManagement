<#
.SYNOPSIS
Module for MS Graph functions

.DESCRIPTION
This module manages Microsoft Grap fuctions like calling APIs, managing graph objects etc. This is common for all view using MS Graph

.NOTES
  Author:         Mikael Karlsson
#>
function Get-ModuleVersion
{
    '3.1.8'
}

$global:MSGraphGlobalApps = @(
    #Authority="https://login.microsoftonline.com/organizations/"
    (New-Object PSObject -Property @{Name="";ClientId="";RedirectUri="";Authority=""}),
    (New-Object PSObject -Property @{Name="Microsoft Intune PowerShell";ClientId="d1ddf0e4-d672-4dae-b554-9d5bdfd93547";RedirectUri="urn:ietf:wg:oauth:2.0:oob"; }),
    (New-Object PSObject -Property @{Name="Microsoft Graph PowerShell";ClientId="14d82eec-204b-4c2f-b7e8-296a70dab67e";RedirectUri="https://login.microsoftonline.com/common/oauth2/nativeclient";})
    )

function Invoke-InitializeModule
{
    $global:graphURL = "https://graph.microsoft.com/beta"

    $global:LoadedDependencyObject = $null
    $global:MigrationTableCache = $null

    $script:lstImportTypes = @(
        [PSCustomObject]@{
            Name = "Always import"
            Value = "alwaysImport"
        },
        [PSCustomObject]@{
            Name = "Skip if object exists"
            Value = "skipIfExist"
        },
        [PSCustomObject]@{
            Name = "Replace (Preview)"
            Value = "replace"
        },
        [PSCustomObject]@{
            Name = "Update (Preview)"
            Value = "update"
        }
    )

    # Make sure MS Graph settings are added before exiting before App Id and Tenant Id is missing
    Write-Log "Add settings and menu items"

    # Add settings
    $global:appSettingSections += (New-Object PSObject -Property @{
        Title = "Import/Export"
        Id = "ImportExport"
        Values = @()
    })

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Root folder"
        Key = "RootFolder"
        Type = "Folder"   
        Description = "Root folder for exporting/importing objects"         
    }) "ImportExport"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Add object type"
        Key = "AddObjectType"
        Type = "Boolean"
        DefaultValue = $true
        Description = "Default setting for adding object type to the export folder"
    }) "ImportExport"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Add company name"
        Key = "AddCompanyName"
        Type = "Boolean"
        DefaultValue = $true
        Description = "Default setting for adding company name to the export folder"
    }) "ImportExport"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Export Assignments"
        Key = "ExportAssignments"
        Type = "Boolean"
        DefaultValue = $true
        Description = "Default setting for exporting assignments"
    }) "ImportExport"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Create groups"
        Key = "CreateGroupOnImport"
        Type = "Boolean"
        DefaultValue = $true
        Description = "Default setting for creating groups during import"
    }) "ImportExport"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Convert synced groups"
        Key = "ConvertSyncedGroupOnImport"
        Type = "Boolean"
        DefaultValue = $true
        Description = "Convert AD synched groups to Azure AD group during import if the group does not exist"
    }) "ImportExport"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Import type"
        Key = "ImportType"
        Type = "List" 
        ItemsSource = $script:lstImportTypes
        DefaultValue = "alwaysImport"
    }) "ImportExport"    

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Import Assignments"
        Key = "ImportAssignments"
        Type = "Boolean"
        DefaultValue = $true
        Description = "Default value for Import assignments when importing objects"
    }) "ImportExport"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Import Scope (Tags)"
        Key = "ImportScopeTags"
        Type = "Boolean"
        DefaultValue = $true
        Description = "Default value for Import Scope (Tags) when importing objects"
    }) "ImportExport"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Show Delete button"
        Key = "EMAllowDelete"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Allow deleting individual objectes"
    }) "ImportExport"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Show Bulk Delete "
        Key = "EMAllowBulkDelete"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Allow using bulk delete to delete all objects of selected types"
    }) "ImportExport"
    

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Allow update on import (Preview)"
        Key = "AllowUpdate"
        Type = "Boolean"
        DefaultValue = $false
        Description = "This will enable the option to update/replace an existing object during import"
    }) "ImportExport"


    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Add ID to export file"
        Key = "AddIDToExportFile"
        Type = "Boolean"
        DefaultValue = $false
        Description = "This will add object ID to the export file to support objects with the same name e.g. ObjectName_ObjectId.json"
    }) "ImportExport"    
}

function Get-GraphAppInfo
{
    param($settingId, $defaultAppId)

    $graphAppId = Get-SettingValue $settingId

    if($graphAppId)
    {
        # Check if an app in the list is selected
        $appObj = $global:MSGraphGlobalApps | Where ClientId -eq $graphAppId
    }

    if(-not $appObj)
    {
        # Set app info from custom settings
        $appObj = New-Object PSObject -Property @{
            ClientId = Get-SettingValue "$($PreFix)CustomAppId"
            TenantId = Get-SettingValue "$($PreFix)CustomTenantId"
            RedirectUri = Get-SettingValue "$($PreFix)CustomAppRedirect"
            Authority = Get-SettingValue "$($PreFix)CustomAuthority"
        }
    }

    if(-not $appObj.ClientId -and $defaultAppId)
    {
        # No app info found. Use default
        $appObj = $global:MSGraphGlobalApps | Where ClientId -eq $defaultAppId
    }

    $appObj
}

function Invoke-GraphAuthenticationUpdated
{
    $global:MigrationTableCache = $null
    $global:MigrationTableCacheId = $null
    $global:LoadedDependencyObjects = $null
    $global:migFileObj = $null
}

function Invoke-SettingsUpdated
{
    Initialize-GraphSettings
}

function Initialize-GraphSettings
{
    
}

function Invoke-GraphRequest
{
    param (
            [Parameter(Mandatory)]
            $Url,

            [Alias("Body")]
            $Content,

            $Headers,

            [ValidateSet("GET","POST","OPTIONS","DELETE", "PATCH","PUT")]
            [Alias("Method")]
            $HttpMethod = "GET",

            $AdditionalHeaders,

            [string]$Outfile = "",

            [Switch]$SkipAuthentication,

            $ODataMetadata = "full", # full, minimal, none or skip

            [switch]
            $AllPages,

            [switch]
            $NoError
        )

    if($SkipAuthentication -ne $true)
    {
        Connect-MSALUser
    }

    $params = @{}

    $requestId = [Guid]::NewGuid().guid

    if(-not $Headers)
    {
        $Headers = @{
        'Content-Type' = 'application/json; charset=utf-8'
        'Authorization' = "Bearer " + $global:MSALToken.AccessToken
        'ExpiresOn' = $global:MSALToken.ExpiresOn
        'x-ms-client-request-id' = $requestId
        }
    }

    if($HttpMethod -eq "GET" -and $ODataMetadata -ne "Skip")
    {
        # Note: odata.metadata=full in Accept 
        # @odata.type is not always included with default (minimum). 
        # That is required to identify the object type in some functions
        # It does include a lot of info we don't need... 
        $Headers.Add("Accept","application/json;odata.metadata=$ODataMetadata")
    }
    #elseif($Content)
    #{
    #    # Upload content as UTF8 to support international and extended characters
    #    $Content = [System.Text.Encoding]::UTF8.GetBytes($Content)
    #}

    if($AdditionalHeaders -is [HashTable])
    {
        foreach($key in $AdditionalHeaders.Keys)
        {
            if($Headers.ContainsKey($key)) { continue }

            $Headers.Add($key, $AdditionalHeaders[$key])
        }
    }

    if($Content) { $params.Add("Body", [System.Text.Encoding]::UTF8.GetBytes($Content)) }
    if($Headers) { $params.Add("Headers", $Headers) }
    if($Outfile)
    {
        $dirName = [IO.Path]::GetDirectoryName($Outfile)
        try {
            [IO.Directory]::CreateDirectory($dirName)            
        }
        catch {
            
        }
        if([IO.Directory]::Exists($dirName))
        {
            $params.Add("OutFile", $OutFile)
        }
        else {
            Write-Log "Failed to create directory for OutFile $Outfile" 3
        }
    }

    if(($Url -notmatch "^http://|^https://"))
    {        
        $Url = $global:graphURL + "/" + $Url.TrimStart('/')
        $Url = $Url -replace "%OrganizationId%", $global:Organization.Id
    }

    <#
    if($AllPages)
    {    
        # Code to test paging - Force each page to size specified in top parameter below
        # Kept for reference

        if(($url.IndexOf('?')) -eq -1) 
        {
            $url = "$($url.Trim())?"
        }
        else
        {
            $url = "$($url.Trim())&"
        }
        $url = "$($url.Trim())`$top=20"
    }
    #>    

    $ret = $null
    try
    {
        Write-LogDebug "Invoke graph API: $Url (Request ID: $requestId)"
        $allValues = @()
        do 
        {
            $ret = Invoke-RestMethod -Uri $Url -Method $HttpMethod @params 
            if($? -eq $false) 
            {
                throw $global:error[0]
            }
    
            if($HttpMethod -eq "PATCH" -and [String]::IsNullOrempty($ret))
            {
                $ret = $true;
                break; 
            }
            elseif($AllPages -eq $true -and $HttpMethod -eq "GET" -and $ret.value -is [Array])
            {
                $allValues += $ret.value
                if($ret.'@odata.nextLink')
                {
                    $Url = $ret.'@odata.nextLink'
                }
            }
            else
            {
                break    
            }
        } while($ret.'@odata.nextLink')
        
        if($allValues.Count -gt 0 -and $ret.value -is [Array])
        {
            $ret.value = $allValues
        }
    }
    catch
    {
        if($NoError -eq $true) { return }
        Write-LogError "Failed to invoke MS Graph with URL $Url (Request ID: $requestId). Status code: $($_.Exception.Response.StatusCode)" $_.Excption
    }
    
    Write-Debug "$(($ret | Select *))"
    
    $ret
}

function Get-GraphObjects 
{
    param(
    [String]
    $Url,
    [Array]
    $property = $null,
    [Array]
    $exclude,
    $SortProperty = "displayName",
    $objectType,
    [switch]
    $SingleObject)

    $objects = @()
    
    if($property -isnot [Object[]]) { $property = @('displayName', 'description', 'id')}

    $params = @{}
    if($objectType.ODataMetadata)
    {
        $params.Add('ODataMetadata',$objectType.ODataMetadata)
    }

    if(-not $url)
    {
        $url = $objectType.API
    }

    if($SingleObject -ne $true -and $objectType.QUERYLIST)
    {
        if(($url.IndexOf('?')) -eq -1) 
        {
            $url = "$($url.Trim())?$($objectType.QUERYLIST.Trim())"
        }
        else
        {
            $url = "$($url.Trim())&$($objectType.QUERYLIST.Trim())" # Risky...does not check that the parameter is already in use
        }        
    }
    
    if($SingleObject -ne $true)
    {
        $params.Add('AllPages',$true)
    }

    $graphObjects = Invoke-GraphRequest -Url $url @params

    if($SingleObject -ne $true -and $objectType.PostListCommand)
    {
        $graphObjects = & $objectType.PostListCommand $graphObjects $objectType
    }
        
    if($graphObjects -and ($graphObjects | GM -Name Value -MemberType NoteProperty))
    {
        $retObjects = $graphObjects.Value            
    }
    else
    {
        $retObjects = $graphObjects
    }

    foreach($graphObject in $retObjects)
    {
        $params = @{}
        if($property) { $params.Add("Property", $property) }
        if($exclude) { $params.Add("ExcludeProperty", $exclude) }
        foreach($objTmp in ($graphObject | Select-Object @params))
        {
            $objTmp | Add-Member -NotePropertyName "IsSelected" -NotePropertyValue $false
            $objTmp | Add-Member -NotePropertyName "Object" -NotePropertyValue $graphObject
            $objTmp | Add-Member -NotePropertyName "ObjectType" -NotePropertyValue $objectType
            $objects += $objTmp
        }            
    }    
    $property = "IsSelected",$property

    if($objects.Count -gt 0 -and $SortProperty -and ($objects[0] | GM -MemberType NoteProperty -Name $SortProperty))
    {
        $objects = $objects | sort -Property $SortProperty
    }
    $objects
}

function Show-GraphObjects
{
    $global:curObjectType = $global:lstMenuItems.SelectedItem

    Clear-GraphObjects

    if(-not $global:MSALToken)
    {
        $global:grdNotLoggedIn.Visibility = "Visible"
        $global:grdData.Visibility = "Collapsed"
        return
    }
    $global:grdNotLoggedIn.Visibility = "Collapsed"
    $global:grdData.Visibility = "Visible"

    # Always show Import is an item is selected
    $global:btnImport.IsEnabled = $global:lstMenuItems.SelectedItem -ne $null

    if(-not $global:lstMenuItems.SelectedItem) { return }

    Write-Status "Loading $($global:curObjectType.Title) objects" 

    if($global:lstMenuItems.SelectedItem.ShowForm -ne $false)
    {
        $viewItem = $global:lstMenuItems.SelectedItem
        if($viewItem.Icon -or [IO.File]::Exists(($global:AppRootFolder + "\Xaml\Icons\$($viewItem.Id).xaml")))
        {
            $global:ccIcon.Content = Get-XamlObject ($global:AppRootFolder + "\Xaml\Icons\$((?? $viewItem.Icon $viewItem.Id)).xaml")
        }
    
        $global:txtFormTitle.Text = $global:lstMenuItems.SelectedItem.Title        
        $global:grdTitle.Visibility = "Visible"
    }

    $graphObjects = @(Get-GraphObjects -property $global:curObjectType.ViewProperties -objectType $global:curObjectType)

    $dgObjects.AutoGenerateColumns = $false
    $dgObjects.Columns.Clear()

    if(($graphObjects | measure).Count -gt 0)
    { 
        $tmpObj = $graphObjects | Select -First 1

        $prop = $tmpObj.PSObject.Properties | Where Name -eq "IsSelected"
        if($prop)
        {        
            $column = Get-GridCheckboxColumn "IsSelected"
            $dgObjects.Columns.Add($column)

            $column.Header.add_Click({
                foreach($item in $global:dgObjects.ItemsSource)
                { 
                    $item.IsSelected = $this.IsChecked
                }
                $global:dgObjects.Items.Refresh()
            })           
        }

        $tableColumns = @()

        $additionalColumns = @()
        $additionalColsStr = Get-Setting "EndpointManager\ObjectColumns" "$($global:curObjectType.Id)"
        if($additionalColsStr)
        {
            $additionalColumns += $additionalColsStr.Split(',')
        }

        if($additionalColumns.Count -eq 0 -or $additionalColumns[0] -ne "0")
        {
            # Add default columns
            foreach($prop in ($tmpObj.PSObject.Properties | Where {$_.Name -notin @("IsSelected","Object","ObjectType")}))
            {
                $binding = [System.Windows.Data.Binding]::new($prop.Name)
                $column = [System.Windows.Controls.DataGridTextColumn]::new()
                $column.Header = $prop.Name
                $column.IsReadOnly = $true
                $column.Binding = $binding

                $tableColumns += $prop.Name
                $dgObjects.Columns.Add($column)
            }
        }

        # Add custom columns
        foreach($additionalCol in $additionalColumns)
        {
            if($additionalCol -eq "0" -or $additionalCol -eq "1") { continue }

            $bindingProp,$colHeader = $additionalCol.Split('=')

            if(-not $colHeader)
            {
                $colHeader = $bindingProp
            }

            $binding = [System.Windows.Data.Binding]::new("Object.$($bindingProp)")
            $column = [System.Windows.Controls.DataGridTextColumn]::new()
            $column.Header = $colHeader
            $column.IsReadOnly = $true
            $column.Binding = $binding

            $tableColumns += $colHeader
            $dgObjects.Columns.Add($column)
        }

        $ocList = [System.Collections.ObjectModel.ObservableCollection[object]]::new($graphObjects)
        $dgObjects.ItemsSource = [System.Windows.Data.CollectionViewSource]::GetDefaultView($ocList)
    }
    else
    {
        $dgObjects.ItemsSource = $null
    }

    # Show/Hide buttons based on object type
    foreach($ctrl in $spSubMenu.Children)
    {
        if(-not $global:curObjectType.ShowButtons -or ($global:curObjectType.ShowButtons | Where-Object { $ctrl.Name -like "*$($_)" } ))
        {
            Write-LogDebug "Show $($ctrl.Name)"
            $ctrl.Visibility = "Visible"
        }
        else
        {
            Write-LogDebug "Hide $($ctrl.Name)"
            $ctrl.Visibility = "Collapsed"
        }
    }
}

function Clear-GraphObjects
{        
    $global:txtFormTitle.Text = ""
    $global:txtEMObjects.Text = ""
    $global:grdTitle.Visibility = "Collapsed"
    $global:grdObject.Children.Clear()
    $global:dgObjects.ItemsSource = $null
    Set-ObjectGrid
    
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-GraphObject
{
    param($obj, $objectType, [switch]$SkipAssignments)

    Write-Status "Loading $((Get-GraphObjectName $obj $objectType))" 

    if($objectType.PreGetCommand)
    {
        $preConfig  = & $objectType.PreGetCommand $obj $objectType
    }

    if($preConfig -isnot [Hashtable]) { $preConfig = @{} }

    if($preConfig.ContainsKey("API") -and $preConfig["API"])
    {
        $api = $preConfig["API"]
    }
    elseif(-not $objectType.APIGET)
    {
        $api = ("$($objectType.API)/$($obj.Id)")
    }
    else
    {
        $api = $graphObject.APIGET -replace "%id%", (Get-GraphObjectId $obj $objectType)
    }

    $expand = @()
    if($obj.'assignments@odata.navigationLink' -and $SkipAssignments -ne $true -and $objectType.ExpandAssignments -ne $false)
    {
        $expand += "assignments"
    }

    if($obj.'apps@odata.navigationLink')
    {
        $expand += "apps"
    }

    if($obj.'settings@odata.navigationLink')
    {
        $expand += "settings"
    }

    if($obj.'roleAssignments@odata.navigationLink')
    {
        $expand += "roleAssignments"
    }
    
    if($obj.'privacyAccessControls@odata.associationLink')
    {
        $expand += "microsoft.graph.windows10GeneralConfiguration/privacyAccessControls"
    }    
    
    if($objectType.Expand)
    {
        foreach($objExpand in $objectType.Expand.Split(","))
        {
            if($objExpand -notin $expand) { $expand += $objExpand}
        }
    }

    if($expand.Count -gt 0)
    {
        if($api.IndexOf('?') -eq -1) 
        {
            $api = ($api + "?`$expand=")
        }
        elseif($api.IndexOf("`$expand") -gt 1)
        {
            $api = ($api + ",") # A bit risky...assumes that expand is last in the existing query 
        }
        else
        {
            $api = ($api + "&`$expand=")
        }

        $api = ($api + ($expand -join ","))
    }

    $objInfo = Get-GraphObjects -Url $api -property $objectType.ViewProperties -objectType $objectType -SingleObject

    if($objInfo -and $objectType.PostGetCommand)
    {
        & $objectType.PostGetCommand $objInfo $objectType
    }
    $objInfo 
}

# Generic Pre-Import function for all imports
function Start-GraphPreImport
{
    param($obj, $objectType)

    if($objectType.SkipRemovingProperties -eq $true) { return }

    $removeProperties = $objectType.PropertiesToRemove

    if($removeProperties -isnot [Object[]])
    {
        $removeProperties = @()        
    }

    if($removeProperties.Count -eq 0 -or $objectType.SkipRemoveDefaultProperties -ne $true)
    {
        # Default properties to delete
        $removeProperties += @('lastModifiedDateTime','createdDateTime','supportsScopeTags','id','modifiedDateTime')
    }

    # Remove OData properties
    foreach($odataProp in ($obj.PSObject.Properties | Where { $_.Name -like "*@Odata*Link" -or $_.Name -like "*@odata.context" -or $_.Name -like "*@odata.id" -or ($_.Name -like "*@odata.type" -and $_.Name -ne "@odata.type")}))
    {        
        $removeProperties += $odataProp.Name
    }

    foreach($prop in $removeProperties)
    {
        # Allow override deleting default propeties e.g. some object types requires the Id property
        if($objectType.SkipRemoveProperties -is [Object[]] -and $prop -in $objectType.SkipRemoveProperties) { continue }
        Remove-Property $obj $prop
    }

    if($objectType.SkipRemovingChildProperties -ne $true)
    {
        foreach($prop in ($obj.PSObject.Properties))
        {
            if($obj."$($prop.Name)"."@odata.type")
            {
                foreach($childObj in ($obj."$($prop.Name)"))
                {
                    Start-GraphPreImport  $childObj $objectType         
                }
            }
        }
    }
}

function Get-GraphMetaData
{
    if(-not $global:metaDataXML)
    {
        # Graph metadata does not support Content-Length in response so size can not be used to check if it is updated
        # There also no other version information in response headers. Use file date to update every week
        $url = "https://graph.microsoft.com/beta/`$metadata"
        $fileFullPath = [Environment]::ExpandEnvironmentVariables("%LOCALAPPDATA%\CloudAPIPowerShellManagement\GraphMetaData.xml")
        $fi = [IO.FileInfo]$fileFullPath
        $maxAge = (Get-Date).AddDays(-7)
        if($fi.Exists -and ($fi.LastWriteTime -gt $maxAge -or $fi.CreationTime -gt $maxAge))
        {            
            try 
            {
                [xml]$global:metaDataXML = Get-Content $fi.FullName              
            }
            catch { }
        }

        if(-not $global:metaDataXML)
        {
            [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
            $wc = New-Object System.Net.WebClient
            $wc.Encoding = [System.Text.Encoding]::UTF8
            try 
            {
                [xml]$global:metaDataXML = $wc.DownloadString($url)
                # Download to string and then use Save to format the XML output
                $global:metaDataXML.Save($fi.FullName)
            }
            catch
            {
                Write-LogError "Failed to download Graph MetaData file" $_.Exception
            }
            finally
            {
                $wc.Dispose()
            }
        }
    }
}

function Get-GraphObjectClassName
{
    param($type)

    Get-GraphMetaData

    $objectClassName = $null
    
    $nodes = $global:metaDataXML.SelectNodes("//*[@Type='Collection(graph.$($type))']")
    if($nodes -ne $null -and $nodes.Count -gt 0)
    {
        foreach($node in $nodes)
        {
            if($node.ParentNode.Name -eq "deviceAppManagement")
            {
                $objectClassName = $node.Name
                break
            }
        }
    }

    $objectClassName
}

#region Export/Import dialogs

function Show-GraphExportForm
{
    $script:exportForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\ExportForm.xaml") -AddVariables
    if(-not $script:exportForm) { return }

    Set-XamlProperty $script:exportForm "txtExportPath" "Text" (?? (Get-Setting "" "LastUsedRoot") (Get-SettingValue "RootFolder"))
    Set-XamlProperty $script:exportForm "chkAddObjectType" "IsChecked" (Get-SettingValue "AddObjectType")
    Set-XamlProperty $script:exportForm "chkAddCompanyName" "IsChecked" (Get-SettingValue "AddCompanyName")
    Set-XamlProperty $script:exportForm "chkExportAssignments" "IsChecked" (Get-SettingValue "ExportAssignments")    

    Set-XamlProperty $script:exportForm "btnExportSelected" "IsEnabled" ($global:dgObjects.SelectedItem -ne $null)
    if(($global:dgObjects.ItemsSource | Where IsSelected -eq $true).Count -gt 0)
    {
        Set-XamlProperty $script:exportForm "lblSelectedObject" "Content" "$(($global:dgObjects.ItemsSource | Where IsSelected -eq $true).Count) selected object(s)" 
    }
    elseif($global:dgObjects.SelectedItem)
    {
        Set-XamlProperty $script:exportForm "lblSelectedObject" "Content" "Selected object: $((Get-GraphObjectName $global:dgObjects.SelectedItem $global:curObjectType))" 
    }
    Add-XamlEvent $script:exportForm "btnCancel" "add_click" {
        $script:exportForm = $null
        Show-ModalObject
    }

    Add-XamlEvent $script:exportForm "btnExportAll" "add_click" {
        
        Export-GraphObjects
        
        $script:exportForm = $null
        Show-ModalObject
    }

    Add-XamlEvent $script:exportForm "btnExportSelected" "add_click" {
        Export-GraphObjects -Selected
        
        $script:exportForm = $null
        Show-ModalObject
    }

    Add-XamlEvent $script:exportForm "browseExportPath" "add_click" {
        $folder = Get-Folder (Get-XamlProperty $script:exportForm "txtExportPath" "Text") "Select root folder for export"
        if($folder)
        {
            Set-XamlProperty $script:exportForm "txtExportPath" "Text" $folder
        }
    }

    Add-GraphExportExtensions $script:exportForm 1 $global:curObjectType
    
    Show-ModalForm "Export $($global:curObjectType.Title) objects" $script:exportForm -HideButtons
}

function Show-GraphBulkExportForm
{
    $script:exportForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\BulkExportForm.xaml") -AddVariables
    if(-not $script:exportForm) { return }

    Set-XamlProperty $script:exportForm "txtExportPath" "Text" (?? (Get-Setting "" "LastUsedRoot") (Get-SettingValue "RootFolder"))
    Set-XamlProperty $script:exportForm "chkAddCompanyName" "IsChecked" (Get-SettingValue "AddCompanyName")
    Set-XamlProperty $script:exportForm "chkExportAssignments" "IsChecked" (Get-SettingValue "ExportAssignments")
    #Set-XamlProperty $script:exportForm "txtExportNameFilter" "Text" (Get-Setting "" "ExportNameFilter")

    Add-XamlEvent $script:exportForm "browseExportPath" "add_click" ({
        $folder = Get-Folder (Get-XamlProperty $script:exportForm "txtExportPath" "Text") "Select root folder for export"
        if($folder)
        {
            Set-XamlProperty $script:exportForm "txtExportPath" "Text" $folder
        }
    })

    $script:exportObjects = @()
    foreach($objType in $global:lstMenuItems.ItemsSource)
    {
        if(-not $objType.Title) { continue }

        if($objType.ShowButtons -is [Object[]] -and $objType.ShowButtons -notcontains "Export") { continue }

        $script:exportObjects += New-Object PSObject -Property @{
            Title = $objType.Title
            Selected = (?? $objType.BulkExport $true)
            ObjectType = $objType
        }

        Add-GraphExportExtensions $script:exportForm 0 $objType
    }    

    $column = Get-GridCheckboxColumn "Selected"
    $global:dgObjectsToExport.Columns.Add($column)

    $column.Header.IsChecked = $true # All items are checked by default
    $column.Header.add_Click({
            foreach($item in $global:dgObjectsToExport.ItemsSource)
            {
                $item.Selected = $this.IsChecked
            }
            $global:dgObjectsToExport.Items.Refresh()
        }
    ) 

    # Add Object type column
    $binding = [System.Windows.Data.Binding]::new("Title")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Object type"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:dgObjectsToExport.Columns.Add($column)

    $global:dgObjectsToExport.ItemsSource = $script:exportObjects

    Add-XamlEvent $script:exportForm "btnClose" "add_click" ({
        $script:exportForm = $null
        Show-ModalObject
    })

    Add-XamlEvent $script:exportForm "btnExport" "add_click" ({
        Write-Status "Export objects" -Block
        Write-Log "****************************************************************"
        Write-Log "Start bulk export"
        Write-Log "****************************************************************"

        $global:AADObjectCache = $null

        foreach($item in $script:exportObjects)
        { 
            if($item.Selected -ne $true) { continue }

            Write-Log "----------------------------------------------------------------"
            Write-Log "Export $($item.ObjectType.Title) objects"
            Write-Log "----------------------------------------------------------------"
            
            $txtNameFilter = $global:txtExportNameFilter.Text.Trim()
            Save-Setting "" "ExportNameFilter" $txtNameFilter
            if($txtNameFilter) { Write-Log "Name filter: $txtNameFilter" }

            try 
            {
                $folder = Get-GraphObjectFolder $item.ObjectType (Get-XamlProperty $script:exportForm "txtExportPath" "Text") (Get-XamlProperty $script:exportForm "chkAddObjectType" "IsChecked") (Get-XamlProperty $script:exportForm "chkAddCompanyName" "IsChecked")

                $objects = @(Get-GraphObjects -property $item.ObjectType.ViewProperties -objectType $item.ObjectType)
                foreach($obj in $objects)
                {
                    $objName = Get-GraphObjectName $obj.Object $obj.ObjectType

                    if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                    {
                        continue
                    }

                    Write-Status "Export $($item.Title): $objName" -Force
                    Export-GraphObject $obj.Object $item.ObjectType $folder
                }
                Save-Setting "" "LastUsedFullPath" $folder
            }
            catch 
            {
                Write-LogError "Failed when exporting $($item.Title) objects" $_.Exception
            }
        }
        Save-Setting "" "LastUsedRoot" (Get-XamlProperty $script:exportForm "txtExportPath" "Text")

        Write-Log "****************************************************************"
        Write-Log "Bulk export finished"
        Write-Log "****************************************************************"
        Write-Status ""
    })

    Show-ModalForm "Bulk Export" $script:exportForm -HideButtons
}

function Show-GraphImportForm
{
    $script:importForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\ImportForm.xaml") -AddVariables
    if(-not $script:importForm) { return }

    $path = Get-Setting "" "LastUsedFullPath"
    if($path) 
    {
        $path = [IO.Path]::Combine([IO.Directory]::GetParent($path).FullName, $global:lstMenuItems.SelectedItem.Id)
        if([IO.Directory]::Exists($path) -eq $false)
        {
            $path = Get-Setting "" "LastUsedRoot"
        }
    }

    Set-XamlProperty $script:importForm "txtImportPath" "Text" (?? $path (Get-SettingValue "RootFolder"))
    Set-XamlProperty $script:importForm "chkImportAssignments" "IsChecked" (Get-SettingValue "ImportAssignments")
    Set-XamlProperty $script:importForm "chkImportScopes" "IsChecked" (Get-SettingValue "ImportScopeTags")
    Set-XamlProperty $script:importForm "cbImportType" "ItemsSource" $script:lstImportTypes
    Set-XamlProperty $script:importForm "cbImportType" "SelectedValue" (Get-SettingValue "ImportType" "alwaysImport")
    
    if((Get-SettingValue "AllowUpdate") -eq $true)
    {
        Set-XamlProperty  $script:importForm "lblImportType" "Visibility" "Visible"
        Set-XamlProperty  $script:importForm "cbImportType" "Visibility" "Visible"
    }

    $column = Get-GridCheckboxColumn "Selected"
    $global:dgObjectsToImport.Columns.Add($column)

    $column.Header.IsChecked = $true # All items are checked by default
    $column.Header.add_Click({
            foreach($item in $global:dgObjectsToImport.ItemsSource)
            {
                $item.Selected = $this.IsChecked
            }
            $global:dgObjectsToImport.Items.Refresh()
        }
    ) 

    # Add Object type column
    $binding = [System.Windows.Data.Binding]::new("fileName")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "File Name"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:dgObjectsToImport.Columns.Add($column)

    Add-XamlEvent $script:importForm "browseImportPath" "add_click" ({
        $folder = Get-Folder (Get-XamlProperty $script:importForm "txtImportPath" "Text") "Select root folder for import"
        if($folder)
        {
            Set-XamlProperty $script:importForm "txtImportPath" "Text" $folder
            $global:dgObjectsToImport.ItemsSource = @(Get-GraphFileObjects $folder)
            Save-Setting "" "LastUsedFullPath" $folder
            Set-XamlProperty $script:importForm "lblMigrationTableInfo" "Content" (Get-MigrationTableInfo)
        }
    })
    
    Add-XamlEvent $script:importForm "btnCancel" "add_click" {
        $script:importForm = $null
        Show-ModalObject
    }

    Add-XamlEvent $script:importForm "btnImportSelected" "add_click" {
        Write-Status "Import objects"
        Get-GraphDependencyDefaultObjects
        $allowUpdate = ((Get-SettingValue "AllowUpdate") -eq $true)
        $filesToImport = $global:dgObjectsToImport.ItemsSource | Where Selected -eq $true
        if($global:curObjectType.PreFilesImportCommand)
        {
            $filesToImport = & $global:curObjectType.PreFilesImportCommand $global:curObjectType $filesToImport
        }

        foreach ($fileObj in $filesToImport)
        {
            if($allowUpdate -and $global:cbImportType.SelectedValue -ne "alwaysImport" -and (Reset-GraphObjet $fileObj $global:dgObjects.ItemsSource))
            {
                continue
            }

            Import-GraphFile $fileObj
        }
        Show-GraphObjects
        Show-ModalObject
        Write-Status ""
    }

    Add-XamlEvent $script:importForm "btnGetFiles" "add_click" {
        # Used when the user manually updates the path and the press Get Files
        $global:dgObjectsToImport.ItemsSource = @(Get-GraphFileObjects $global:txtImportPath.Text)
        if([IO.Directory]::Exists($global:txtImportPath.Text))
        {
            Save-Setting "" "LastUsedFullPath" $global:txtImportPath.Text
            Set-XamlProperty $script:importForm "lblMigrationTableInfo" "Content" (Get-MigrationTableInfo)
        }
    }

    Add-GraphImportExtensions $script:importForm 1 $global:curObjectType

    if($global:txtImportPath.Text)
    {
        $global:dgObjectsToImport.ItemsSource = @(Get-GraphFileObjects $global:txtImportPath.Text)
        Set-XamlProperty $script:importForm "lblMigrationTableInfo" "Content" (Get-MigrationTableInfo)
    }
    
    Show-ModalForm "Import objects" $script:importForm -HideButtons
}

function Show-GraphBulkImportForm
{
    $script:importForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\BulkImportForm.xaml") -AddVariables
    if(-not $script:importForm) { return }

    $path = Get-Setting "" "LastUsedFullPath"
    if($path) 
    {
        $path = [IO.Directory]::GetParent($path).FullName
    }

    Set-XamlProperty $script:importForm "txtImportPath" "Text" (?? $path (Get-SettingValue "RootFolder"))
    Set-XamlProperty $script:importForm "chkImportAssignments" "IsChecked" (Get-SettingValue "ImportAssignments")
    Set-XamlProperty $script:importForm "chkImportScopes" "IsChecked" (Get-SettingValue "ImportScopeTags")
    Set-XamlProperty $script:importForm "cbImportType" "ItemsSource" $script:lstImportTypes
    Set-XamlProperty $script:importForm "cbImportType" "SelectedValue" (Get-SettingValue "ImportType" "alwaysImport")
    #Set-XamlProperty $script:importForm "txtImportNameFilter" "Text" (Get-Setting "" "ImportNameFilter")
    
    if((Get-SettingValue "AllowUpdate") -eq $true)
    {
        Set-XamlProperty  $script:importForm "lblImportType" "Visibility" "Visible"
        Set-XamlProperty  $script:importForm "cbImportType" "Visibility" "Visible"        
    }

    Add-XamlEvent $script:importForm "browseImportPath" "add_click" ({
        $folder = Get-Folder (Get-XamlProperty $script:importForm "txtImportPath" "Text") "Select root folder for import"
        if($folder)
        {
            Set-XamlProperty $script:importForm "txtImportPath" "Text" $folder            
            Set-XamlProperty $script:importForm "lblMigrationTableInfo" "Content" (Get-MigrationTableInfo)       
        }
    })

    $script:importObjects = @()
    foreach($objType in $global:lstMenuItems.ItemsSource)
    {
        if(-not $objType.Title) { continue }

        if($objType.ShowButtons -is [Object[]] -and $objType.ShowButtons -notcontains "Import") { continue }

        $script:importObjects += New-Object PSObject -Property @{
            Title = $objType.Title
            Selected = (?? $objType.BulkImport $true)
            ObjectType = $objType
        }

        Add-GraphImportExtensions $script:importForm 0 $objType
    }

    $column = Get-GridCheckboxColumn "Selected"
    $global:dgObjectsToImport.Columns.Add($column)

    $column.Header.IsChecked = $true # All items are checked by default
    $column.Header.add_Click({
            foreach($item in $global:dgObjectsToImport.ItemsSource)
            {
                $item.Selected = $this.IsChecked
            }
            $global:dgObjectsToImport.Items.Refresh()
        }
    ) 

    # Add Object type column
    $binding = [System.Windows.Data.Binding]::new("Title")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Object type"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:dgObjectsToImport.Columns.Add($column)

    # Add Order column
    $binding = [System.Windows.Data.Binding]::new("ObjectType.ImportOrder")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Import order"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:dgObjectsToImport.Columns.Add($column)
    
    $global:dgObjectsToImport.ItemsSource = $script:importObjects

    Add-XamlEvent $script:importForm "btnClose" "add_click" ({
        $script:importForm = $null
        Show-ModalObject
    })

    Add-XamlEvent $script:importForm "btnImport" "add_click" ({
        Write-Status "Import objects" -Block
        Write-Log "****************************************************************"
        Write-Log "Start bulk import"
        Write-Log "****************************************************************"
        Get-GraphDependencyDefaultObjects
        $importedObjects = 0

        $txtNameFilter = $global:txtImportNameFilter.Text.Trim()
        Save-Setting "" "ImportNameFilter" $txtNameFilter
        if($txtNameFilter) { Write-Log "Name filter: $txtNameFilter" }

        $allowUpdate = ((Get-SettingValue "AllowUpdate") -eq $true)

        foreach($item in ($script:importObjects | where Selected -eq $true | sort-object -property @{e={$_.ObjectType.ImportOrder}}))
        { 
            Write-Status "Import $($item.ObjectType.Title) objects" -Force
            Write-Log "----------------------------------------------------------------"
            Write-Log "Import $($item.ObjectType.Title) objects"
            Write-Log "----------------------------------------------------------------"
            $folder = Get-GraphObjectFolder $item.ObjectType (Get-XamlProperty $script:importForm "txtImportPath" "Text") (Get-XamlProperty $script:importForm "chkAddObjectType" "IsChecked")
            
            $graphObjects = $null

            if($allowUpdate -and $global:cbImportType.SelectedValue -ne "alwaysImport")
            {           
                try 
                {
                    Write-Status "Get $($item.Title) objects" -Force
                    $graphObjects = @(Get-GraphObjects -property $item.ObjectType.ViewProperties -objectType $item.ObjectType)
                }
                catch {}
            }

            if([IO.Directory]::Exists($folder))
            {
                $filesToImport = Get-GraphFileObjects $folder -ObjectType $item.ObjectType
                if($item.ObjectType.PreFilesImportCommand)
                {
                    $filesToImport = & $item.ObjectType.PreFilesImportCommand $item.ObjectType $filesToImport
                }
        
                foreach ($fileObj in @($filesToImport))
                {
                    $objName = Get-GraphObjectName $fileObj.Object $item.ObjectType

                    if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                    {
                        continue
                    }

                    if($allowUpdate -and $global:cbImportType.SelectedValue -ne "alwaysImport" -and $graphObjects -and (Reset-GraphObjet $fileObj $graphObjects))
                    {
                        $importedObjects++ 
                        continue
                    }
        
                    Import-GraphFile $fileObj
                    $importedObjects++
                }
                Save-Setting "" "LastUsedFullPath" $folder
            }
            else
            {
                Write-Log "Folder $folder not found. Skipping import" 2    
            }
        }

        Write-Log "****************************************************************"
        Write-Log "Bulk import finished"
        Write-Log "****************************************************************"
        Write-Status ""
        if($importedObjects -eq 0)
        {
            [System.Windows.MessageBox]::Show("No objects were imported. Verify folder and exported files", "Error", "OK", "Error")
        }
        else
        {
            Show-GraphObjects
            Write-Status ""
        }
    })

    if((Get-XamlProperty $script:importForm "txtImportPath" "Text"))
    {
        Set-XamlProperty $script:importForm "lblMigrationTableInfo" "Content" (Get-MigrationTableInfo)
    }

    Show-ModalForm "Bulk Import" $script:importForm -HideButtons
}

function Add-GraphExportExtensions
{
    param($form, $buttonIndex = 0, $objectTypes)
       
    #$global:curObjectType
    foreach($objectType in $objectTypes)
    {
        if($objectType.ExportExtension)
        {
            $grid = $form.FindName("grdExportProperties")
            $extraProperties = & $objectType.ExportExtension $form "spExportSubMenu" 1
            for($i=0;($i + 1) -lt (($extraProperties) | measure).Count;$i ++) 
            {            
                $rd = [System.Windows.Controls.RowDefinition]::new()
                $rd.Height = [double]::NaN            
                $grid.RowDefinitions.Add($rd)
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::RowProperty,$grid.RowDefinitions.Count)
                $grid.Children.Add($extraProperties[$i])

                $i++            
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::RowProperty,$grid.RowDefinitions.Count)
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
                $grid.Children.Add($extraProperties[$i])
                
            }
        }
    }    
}

function Add-GraphImportExtensions
{
    param($form, $buttonIndex = 0, $objectTypes)
    
    foreach($objectType in $objectTypes)
    {
        if($objectType.ImportExtension)
        {
            $grid = $form.FindName("grdImportProperties")
            $extraProperties = & $objectType.ImportExtension $form "spExportSubMenu" 1
            for($i=0;($i + 1) -lt (($extraProperties) | measure).Count;$i ++) 
            {            
                $rd = [System.Windows.Controls.RowDefinition]::new()
                $rd.Height = [double]::NaN            
                $grid.RowDefinitions.Add($rd)
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::RowProperty,$grid.RowDefinitions.Count)
                $grid.Children.Add($extraProperties[$i])

                $i++            
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::RowProperty,$grid.RowDefinitions.Count)
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
                $grid.Children.Add($extraProperties[$i])
                
            }
        }
    }
}

function Show-GraphBulkDeleteForm
{
    $script:deleteForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\BulkDeleteForm.xaml") -AddVariables
    if(-not $script:deleteForm) { return }

    Set-XamlProperty $script:deleteForm "txtDeleteNameFilter" "Text" (Get-Setting "" "txtDeleteNameFilter")

    $script:deleteObjects = @()
    foreach($objType in $global:lstMenuItems.ItemsSource)
    {
        if(-not $objType.Title) { continue }

        if($objType.ShowButtons -is [Object[]] -and $objType.ShowButtons -notcontains "Delete") { continue }

        $script:deleteObjects += New-Object PSObject -Property @{
            Title = $objType.Title
            Selected = $false
            ObjectType = $objType
        }
    }

    $column = Get-GridCheckboxColumn "Selected"
    $global:dgBulkDeleteObjects.Columns.Add($column)

    $column.Header.IsChecked = $false # All items are NOT checked by default
    $column.Header.add_Click({
            foreach($item in $global:dgBulkDeleteObjects.ItemsSource)
            {
                $item.Selected = $this.IsChecked
            }
            $global:dgBulkDeleteObjects.Items.Refresh()
        }
    ) 

    # Add title column
    $binding = [System.Windows.Data.Binding]::new("Title")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Title"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:dgBulkDeleteObjects.Columns.Add($column)    

    $ocList = [System.Collections.ObjectModel.ObservableCollection[object]]::new(@($script:deleteObjects))
    $global:dgBulkDeleteObjects.ItemsSource = [System.Windows.Data.CollectionViewSource]::GetDefaultView($ocList)    

    Add-XamlEvent $script:deleteForm "btnClose" "add_click" ({
        $script:deleteForm = $null
        Show-ModalObject
    })

    Add-XamlEvent $script:deleteForm "btnDelete" "add_click" ({

        $selCount = (($global:dgBulkDeleteObjects.ItemsSource | Where Selected -eq $true) | measure).Count

        if($selCount -eq 0)
        {
            [System.Windows.MessageBox]::Show("No object types selected`n`nSelect types you want to delete", "Error", "OK", "Error") 
            return 
        }

        if(([System.Windows.MessageBox]::Show("Are you sure you want to delete all objects of the selected type(s)?`n`n$selCount type(s) selected`n`nEnvironment: $($global:Organization.displayName)", "Delete Objects?", "YesNo", "Warning")) -ne "Yes")
        {
            return
        }
    
        Write-Status "Delete objects" -Block
        Write-Log "****************************************************************"
        Write-Log "Start bulk delete"
        Write-Log "****************************************************************"

        foreach($item in ($global:dgBulkDeleteObjects.ItemsSource | Where Selected -eq $true))
        {
            Write-Log "----------------------------------------------------------------"
            Write-Log "Delete $($item.ObjectType.Title) objects"
            Write-Log "----------------------------------------------------------------"
            
            $txtNameFilter = $global:txtDeleteNameFilter.Text.Trim()
            Save-Setting "" "DeleteNameFilter" $txtNameFilter
            if($txtNameFilter) { Write-Log "Name filter: $txtNameFilter" }
            
            try 
            {
                Write-Status "Get $($item.Title) objects" -Force
                $objects = @(Get-GraphObjects -property $item.ObjectType.ViewProperties -objectType $item.ObjectType)
                foreach($obj in $objects)
                {                    
                    $objName = Get-GraphObjectName $obj.Object $obj.ObjectType

                    if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                    {
                        continue
                    }

                    Write-Status "Delete $($item.Title): $objName" -Force -SkipLog
                    Remove-GraphObject $obj.Object $obj.ObjectType $folder 
                }
            }
            catch 
            {
                Write-LogError "Failed when deleting $($item.Title) objects" $_.Exception
            }
        }

        Write-Log "****************************************************************"
        Write-Log "Bulk delete finished"
        Write-Log "****************************************************************"
        Show-GraphObjects
        Write-Status ""    
    })


    Show-ModalForm "Bulk Delete" $script:deleteForm -HideButtons
}

function Get-GraphFileObjects
{
    param($path, $Exclude = @("*_settings.json","*_assignments.json"), $SelectedStatus = $true, $ObjectType = $global:curObjectType)

    if(-not $path -or (Test-Path $path) -eq $false) { return }

    $params = @{}
    if($exclude)
    {
        $params.Add("Exclude", $exclude)
    }

    $fileArr = @()
    foreach($file in (Get-Item -path "$path\*.json" @params))
    {
        if($ObjectType.LoadObject)
        {
            $graphObj  = & $ObjectType.LoadObject $file.FullName
        }
        else
        {
            $graphObj = (ConvertFrom-Json (Get-Content -LiteralPath $file.FullName -Raw))
        }

        $obj = New-Object PSObject -Property @{
                FileName = $file.Name
                FileInfo = $file
                Selected = $SelectedStatus
                Object = $graphObj
                ObjectType = $ObjectType
        }

        $fileArr += $obj
    }
    
    if(($fileArr | measure).Count -eq 1)
    {
        return @($fileArr)
    }
    return $fileArr
}

function Import-GraphFile
{
    param($file, $objectType) 

    if([IO.File]::Exists($file.FileInfo.FullName) -eq $false)
    {
        Write-Log "File '$($file.FileInfo.FullName)' not found. Cannot import object" 3
        return
    }

    if($global:chkImportAssignments -and $global:chkImportAssignments.IsChecked -eq $true)
    {
        Get-GraphMigrationObjectsFromFile
    }

    Get-GraphDependencyObjects $file.ObjectType
    
    try 
    {
        # Clone the object to keep original values
        $objClone = $file.Object | ConvertTo-Json -Depth 20 | ConvertFrom-Json

        if($objectType.PreFileImportCommand)
        {
            & $objectType.PreFileImportCommand $objectType $file
        }
        
        Set-ScopeTags $file.Object

        # Never import with assignments. Add them if requested
        Remove-Property $file.Object "Assignments"
        
        $newObj = Import-GraphObject $file.Object $file.ObjectType $file.FileInfo.FullName

        if($newObj -and $file.ObjectType.PostFileImportCommand)
        {
            & $file.ObjectType.PostFileImportCommand $newObj $file.ObjectType $file.FileInfo.FullName
        }
        
        if($newObj -and $objClone.Assignments -and $global:chkImportAssignments.IsChecked -eq $true)
        {
            Import-GraphObjectAssignment $newObj $file.ObjectType $objClone.Assignments $file.FileInfo.FullName | Out-Null
        }        
    } 
    catch 
    {
        Write-LogError "Failed to import file '$($file.FileInfo.Name)'" $_.Exception        
    }
}

function Reset-GraphObjet
{ 
    param($fileObj, $objectList)

    $nameProp = ?? $fileObj.ObjectType.NameProperty "displayName"
    $curObject = $objectList | Where { $_.Object.$nameProp -eq $fileObj.Object.$nameProp -and $_.Object.'@OData.Type' -eq $fileObj.Object.'@OData.Type' }
    
    if($global:cbImportType.SelectedValue -eq "skipIfExist" -and ($curObject | measure).Count -gt 0)
    {
        Write-Log "Objects with name $($fileObj.Object.$nameProp) already exists. Object will not be imported"
        return $true
    }
    elseif(($curObject | measure).Count -gt 1)
    {
        Write-Log "Multiple objects return with name $($fileObj.Object.$nameProp). Object will not be imported or replaced" 2
        return $true
    }
    elseif(($curObject | measure).Count -eq 1)
    {
        Write-Log "Update $((Get-GraphObjectName $fileObj.Object $fileObj.ObjectType)) with id $($curObject.Object.Id)"
        $objectType = $fileObj.ObjectType

        # Clone the object before removing properties
        $obj = $fileObj.Object | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        Start-GraphPreImport $obj $objectType
        Remove-Property $obj "Assignments"
        Remove-Property $obj "isAssigned"
    
        if($global:cbImportType.SelectedValue -eq "update")
        {
            foreach($prop in $objectType.PropertiesToRemoveForUpdate)
            {
                Remove-Property $obj $prop
            }

            $params = @{}
            $strAPI = (?? $objectType.APIPATCH $objectType.API) + "/$($curObject.Object.Id)"
            $method = "PATCH"
            if($objectType.PreUpdateCommand)
            {
                $ret = & $objectType.PreUpdateCommand $obj $objectType $curObject $fileObj.Object
                if($ret -is [HashTable])
                {
                    if($ret.ContainsKey("Import") -and $ret["Import"] -eq $false)
                    {
                        # Import handled manually 
                        return $true
                    }

                    if($ret.ContainsKey("API"))
                    {
                        $strAPI = $ret["API"]
                    }
                    
                    if($ret.ContainsKey("Method"))
                    {
                        $method = $ret["Method"]
                    }

                    if($ret.ContainsKey("AdditionalHeaders") -and $ret["AdditionalHeaders"] -is [HashTable])
                    {
                        $params.Add("AdditionalHeaders",$ret["AdditionalHeaders"])
                    }            
                }
            }

            $json = ConvertTo-Json $obj -Depth 15
            if($true) #$global:MigrationTableCacheId -ne $global:Organization.Id)
            {
                # Call Update-JsonForEnvironment before importing the object
                # E.g. PolicySets contains references, AppConfiguration policies reference apps etc.
                $json = Update-JsonForEnvironment $json
            }

            $objectUpdated = (Invoke-GraphRequest -Url $strAPI -Content $json -HttpMethod $method @params)

            if($objectUpdated)
            {
                Write-Log "Object updated successfully"
            }

            if($objectUpdated -and $objectType.PostUpdateCommand)
            {
                # Reload the updated object
                $updatedObject = Get-GraphObject $curObject.Object $objectType
                & $objectType.PostUpdateCommand $updatedObject $fileObj
            }
            return $true
        }
        elseif($global:cbImportType.SelectedValue -eq "replace")
        {           
            $replace = $true
            $import = $true
            $delete = $true

            if($objectType.PreReplaceCommand)
            {
                $ret = & $objectType.PreReplaceCommand $obj $objectType $curObject.Object $fileObj
                if($ret -is [Hashtable])
                {
                    if($ret["Replace"] -eq $false) { $replace = $false }

                    if($ret["Import"] -eq $false) { $import = $false }

                    if($ret["Delete"] -eq $false) { $delete = $false }
                }                
            }

            if($import)
            {
                $newObj = Import-GraphObject $obj $objectType $fileObj.FileInfo.FullName
            }

            if($newObj -and $replace)
            {
                if($objectType.PostReplaceCommand)
                {
                    $ret = & $objectType.PostReplaceCommand $newObj $objectType $curObject.Object $fileObj
                    if($ret -is [Hashtable])
                    {
                        if($ret["Delete"] -eq $false) { $delete = $false }
                    }                          
                }

                # Load all information about current object to include assignments
                $curObject = Get-GraphObject $curObject.Object $objectType

                $refAssignments = $curObject.Object.Assignments | Where { $_.Source -ne "direct" } 
                if($refAssignments)
                {
                    foreach($refAssignment in $refAssignments)
                    {
                        if($refAssignment.Source -eq "policySets")
                        {
                            Update-EMPolicySetAssignment $refAssignment $curObject $newObj $objectType                           
                        }
                    }
                }
    
                Import-GraphObjectAssignment $newObj $objectType $curObject.Object.Assignments $fileObj.FileInfo.FullName -CopyAssignments | Out-Null

                if($delete)
                {
                    Remove-GraphObject $curObject.Object $objectType
                }
            }
            elseif($replace -eq $false) # Might not be 100% correct. Replace -eq $false probably means that the object was patched and not imported eg default enrollment restrictions etc.
            {
                Write-Log "Failed to import file for $($fileObj.Object.$nameProp) ($($objectType.Title))" 2
            }
            return $true
        }
    }
    # No object to update. Import the file
    return $false
}

function Import-GraphObjectAssignment
{
    param($obj, $objectType, $assignments, $fromFile, [switch]$CopyAssignments)

    if(($assignments | measure).Count -eq 0) { return }

    $preConfig = $null
    $clonedAssignments = $assignments | ConvertTo-Json -Depth 20 | ConvertFrom-Json

    if($objectType.PreImportAssignmentsCommand)
    {
        $preConfig = & $objectType.PreImportAssignmentsCommand $obj $objectType $fromFile $clonedAssignments
    }

    if($preConfig -isnot [Hashtable]) { $preConfig = @{} }

    if($preConfig["Import"] -eq $false) { return } # Assignment managed manually so skip further processing

    $api = ?? $preConfig["API"] "$($objectType.API)/$($newObj.Id)/assign"

    $method = ?? $preConfig["Method"] "POST"

    $keepProperties = ?? $objectType.AssignmentProperties @("target")
    $keepTargetProperties = ?? $objectType.AssignmentTargetProperties @("@odata.type","groupId")
    
    $ObjectAssignments = @()
    foreach($assignment in $clonedAssignments)
    {
        if(($assignment.target.UserId -and $CopyAssignments -ne $true) -or ($assignment.Source -and $assignment.Source -ne "direct"))
        {
            # E.g. Source could be PolicySet...so should not be added here
            continue 
        }

        $assignment.Id = ""
        foreach($prop in $assignment.PSObject.Properties)
        {
            if($prop.Name -in $keepProperties) { continue }
            Remove-Property $assignment $prop.Name
        }

        foreach($prop in $assignment.target.PSObject.Properties)
        {
            if($prop.Name -in $keepTargetProperties) { continue }
            Remove-Property $assignment.target $prop.Name
        }
        
        $ObjectAssignments += $assignment
    }

    if($ObjectAssignments.Count -eq 0) { return } # No "Direct" assignments

    $htAssignments = @{}
    $htAssignments.Add((?? $objectType.AssignmentsType "assignments"), @($ObjectAssignments))

    $json = $htAssignments | ConvertTo-Json -Depth 20
    if($CopyAssignments -ne $true)
    {
        $json = Update-JsonForEnvironment $json
    }

    $objAssign = Invoke-GraphRequest $api -HttpMethod $method -Content $json

    if($objectType.PostImportAssignmentsCommand)
    {
        & $objectType.PostImportAssignmentsCommand $obj $objectType $fromFile $objAssign
    }
    
}
#endregion

#region Migration Info
########################################################################
#
# Migration functions
#
########################################################################
function Set-ScopeTags
{
    param($obj)
    # ToDo: Get values from exported json files instead of MigrationTable?

    if(-not $obj.roleScopeTagIds) { return }

    $scopesIds = @()
    $loadedScopeTags = $global:LoadedDependencyObjects["ScopeTags"]
    $usingDefault = (($obj.roleScopeTagIds | measure).Count -eq 1 -and $obj.roleScopeTagIds[0] -eq "0")
    if($loadedScopeTags -and $global:chkImportScopes.IsChecked -eq $true -and $usingDefault -eq $false -and $global:MigrationTableCache)
    {        
        foreach($scopeId in $obj.roleScopeTagIds)
        {
            if($scopeId -eq 0) { $scopesIds += "0"; continue } # Add default

            $scopeMigObj = $loadedScopeTags | Where OriginalId -eq $scopeId
            if($scopeMigObj -and $scopeMigObj.Id)
            {
                $scopesIds += "$($scopeMigObj.Id)"
            }
            elseif($scopeMigObj)
            {
                Write-Log "Could not find a ScopeTag for exported Id '$($obj.Id)' ($($scopeMigObj.Name)). Make sure all ScopeTags are imported into the environment" 2
            }            
        }
    }
    if($scopesIds.Count -eq 0)
    {
        $scopesIds += "0" # Import with Default ScopeTag as default.
    }
    $obj.roleScopeTagIds = $scopesIds
}

# Called during export to add group info for assignments
# $objAssignments is specified for objects who don't support getting the assgnment info with expand=assignments
function Add-GraphMigrationInfo
{
    param($obj, $objAssignments)

    if(-not $obj) { return }

    $assignments = ?? $objAssignments $obj.Assignments

    foreach($assignment in $assignments)
    {
        foreach($objInfo in $assignment.target)
        {        
            if(-not $objInfo."@odata.type") { continue }

            $objType = $objInfo."@odata.type"

            if($objType -eq "#microsoft.graph.groupAssignmentTarget" -or
                $objType -eq "#microsoft.graph.exclusionGroupAssignmentTarget")
            {
                Add-GroupMigrationObject $objInfo.groupid
            }
            elseif($objType -eq "#microsoft.graph.allLicensedUsersAssignmentTarget" -or
                $objType -eq "#microsoft.graph.allDevicesAssignmentTarget")
            {
                # No need to migrate All Users or All Devices
            }        
            else
            {
                Write-Log "Unsupported migration object: $objType" 3
            }
        }
    }
}

# Used during Import to display Migration Table info on the Import Form
function Get-MigrationTableInfo
{
    $fileName = Get-GraphMigrationTableForImport 

    $str = $null
    $sameTenant = $false
    if($fileName -and [IO.File]::Exists($fileName))
    {
        $migFileObj = ConvertFrom-Json (Get-Content $fileName -Raw)
        if($migFileObj.TenantId -and $migFileObj.TenantId -eq $global:organization.Id) 
        { 
            $sameTenant = $true
            $str = "Current tenant. Migration table will not be used"
        }
        elseif($migFileObj.Organization)
        {
            $str = "Objects exported from $($migFileObj.Organization) ($($migFileObj.TenantId))"
        }
    }
    $chkReplaceDependencyIDs.IsEnabled = $sameTenant -eq $false
    $chkReplaceDependencyIDs.IsChecked = $sameTenant -eq $false

    if(-not $str)
    {
        # Hide controls?
        $str = "No migration table found"
    }
    $str
}

function Get-GraphMigrationTableFile
{
    param($path)

    if(-not $path)
    {
        Write-Log "Export path not set" 3
        return
    }

    if($global:chkAddCompanyName.IsChecked)
    {
        $path = Join-Path $path $global:organization.displayName
    }
    $path
}

function Add-GroupMigrationObject
{
    param($groupId)

    if(-not $groupId) { return }

    $path = Get-GraphMigrationTableFile $global:txtExportPath.Text

    if(-not $path) { return }

    # Check if group is already processed
    $groupObj = Get-GraphMigrationObject $groupId
    if(-not $groupObj)
    {
        # Get group info
        $groupObj = Invoke-GraphRequest "/groups/$groupId" -ODataMetadata "none"
    }

    if($groupObj)
    {
        # Add group to cache
        if($global:AADObjectCache.ContainsKey($groupId) -eq $false) { $global:AADObjectCache.Add($groupId, $groupObj) }

        # Add group to migration file
        if((Add-GraphMigrationObject $groupObj $path "Group"))
        {
            # Export group info to json file for possible import
            $grouspPath = Join-Path $path "Groups"
            if(-not (Test-Path $grouspPath)) { mkdir -Path $grouspPath -Force -ErrorAction SilentlyContinue | Out-Null }
            $fileName = "$grouspPath\$((Remove-InvalidFileNameChars $groupObj.displayName)).json"
            ConvertTo-Json $groupObj -Depth 10 | Out-File $fileName -Force            
        }
    }
}

function Get-GraphMigrationObject
{
    param($objId)

    if(-not $global:AADObjectCache)
    {
        $global:AADObjectCache = @{}
    }

    if($global:AADObjectCache.ContainsKey($objId)) { return $global:AADObjectCache[$objId] }
}

# Adds an object to migration file if not added previously 
function Add-GraphMigrationObject
{
    param($obj, $path, $objType)

    if(-not $objType) { $objType = $obj."@odata.type" }

    $migFileName = Join-Path $path "MigrationTable.json"

    if($global:migFileObj -and $global:migFileObj.TenantId -ne $global:organization.Id)
    {
        $global:migFileObj = $null
    }

    if(-not $global:migFileObj -or ([IO.File]::Exists($migFileName) -eq $false))
    {
        if(-not ([IO.File]::Exists($migFileName)))
        {
            # Create new file
            $global:migFileObj = (New-Object PSObject -Property @{
                TenantId = $global:organization.Id
                Organization = $global:organization.displayName
                Objects = @()
            })
        }
        else
        {
            # Add to existing file
            $global:migFileObj = ConvertFrom-Json (Get-Content $migFileName -Raw) 
        }
    }

    # Make sure Objects property actually exists
    if(($global:migFileObj | GM -MemberType NoteProperty -Name "Objects") -eq $false)
    {
        $global:migFileObj | Add-Member -MemberType NoteProperty -Name "Objects" -Value (@())
    }

    # Get current object
    $curObj = $global:migFileObj.Objects | Where { $_.Id -eq $obj.Id -and $_.Type -eq $objType }

    if($curObj) { return $false } # Existing object found so return false to tell that the object was not added

    $global:migFileObj.Objects += (New-Object PSObject -Property @{
            Id = $obj.Id
            DisplayName = $obj.displayName
            Type = $objType
        })    

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }
    ConvertTo-Json $global:migFileObj -Depth 10 | Out-File $migFileName -Force

    $true # New object was added
}

function Get-GraphMigrationTableForImport
{
    $global:GraphMigrationTable = $null
    # Migration table must be located in the root of the import path
    $path = $global:txtImportPath.Text
    
    for($i = 0;$i -lt 2;$i++)
    {
        if($i -gt 0)
        {
            # Get parent directory
            $path = [io.path]::GetDirectoryName($path)
        }

        $migFileName = Join-Path $path "MigrationTable.json"
        try
        {
            if([IO.File]::Exists($migFileName))
            {
                $global:GraphMigrationTable = $migFileName
                return $migFileName
            }
        }
        catch {}
    }

    Write-Log "Could not find migration table" 2
}

# Cache the migration table and create all missing groups
function Get-GraphMigrationObjectsFromFile
{
    if($global:MigrationTableCache) { return }

    $migFileName = Get-GraphMigrationTableForImport
    if(-not $migFileName) { return }

    $migFileObj = ConvertFrom-Json (Get-Content $migFileName -Raw) 

    # No need to translate migrated objects in the same environment as exported 
    if($migFileObj.TenantId -eq $global:organization.Id) { return }

    $global:MigrationTableCache = @()
    $global:MigrationTableCacheId = $migFileObj.TenantId

    Write-Status "Loading migration objects"

    if($global:chkImportAssignments.IsChecked -eq $true)
    {
        # Only check groups if Assignments are imported
        # This will CREATE the group if it doesn't exist in the target environment
        foreach($migObj in $migFileObj.Objects)
        {
            if($migObj.Type -like "*group*")
            {                
                $obj = (Invoke-GraphRequest "/groups?`$filter=displayName eq '$($migObj.DisplayName)'").Value
                if(-not $obj)
                {
                    $groupFi = $null
                    if($global:GraphMigrationTable)
                    {
                        $fi = [IO.FileInfo]$global:GraphMigrationTable
                        $groupFi = [IO.FileInfo]($fi.DirectoryName + "\Groups\$((Remove-InvalidFileNameChars $migObj.DisplayName)).json")
                    }

                    if($groupFi.Exists -eq $true)
                    {
                        # ToDo: Create group from Json (could be a dynamic group)
                        # Warn if synched group
                        $groupObj = (Get-Content -LiteralPath $groupFi.FullName) | ConvertFrom-Json 

                        #isAssignableToRole - For Role assignment groupd.
                        $keepProps = @("displayName","description","mailEnabled","mailNickname","securityEnabled","membershipRule","groupTypes", "membershipRuleProcessingState")
                        foreach($prop in $groupObj.PSObject.Properties)
                        {
                            if($prop.Name -in $keepProps) { continue }
                            
                            Remove-Property $groupObj $prop.Name
                        }
                        $groupJson = ConvertTo-Json $groupObj -Depth 10
                    }
                    else
                    {
                        Write-Log "No group object found for $($migObj.DisplayName). Creating a cloud group with default settings" 2
                        $groupJson = @"
                        { 
                            "displayName": "$($migObj.DisplayName)",
                            "groupTypes": [
                                ],
                            "mailEnabled": false,
                            "mailNickname" "NotSet"
                            "securityEnabled": true
                        }
"@
                    }
                    Write-Log "Create AAD Group $($migObj.DisplayName)"
                
                    $obj = Invoke-GraphRequest "/groups" -HttpMethod "POST" -Content $groupJson
                }
                $global:MigrationTableCache += (New-Object PSObject -Property @{
                    OriginalId = $migObj.Id            
                    Id = $obj.Id
                    Type = $migObj.Type    
                })
            }
        }
    }
}
function Update-JsonForEnvironment
{
    param($json)

    # Load MigrationTable file unless previously loaded
    Get-GraphMigrationObjectsFromFile

    if($global:chkReplaceDependencyIDs.IsChecked -eq $true)
    {
        foreach($depObjType in $global:LoadedDependencyObjects.Keys)
        {
            foreach($depObj in $global:LoadedDependencyObjects[$depObjType])
            {
                if(-not $depObj.Id -or -not $depObj.OriginalId) { continue }
                if($depObj.OriginalId.Length -lt 36) { continue } # Skip non-guid IDs # ToDo: Verify...
                $json = $json -replace $depObj.OriginalId,$depObj.Id    
            }
        }
    }

    if(-not $global:MigrationTableCache -or $global:MigrationTableCache.Count -eq 0) { return $json }

    # Enumerate all objects in the migration table and replace all exported Id's to Id's in the new environment 
    foreach($migInfo in ($global:MigrationTableCache | Where Type -like "*group*"))
    {
        if(-not $migInfo.Id -or -not $migInfo.OriginalId) { continue }
        if($migInfo.OriginalId.Length -lt 36) { continue } # Skip non-guid IDs # ToDo: Verify...
        $json = $json -replace $migInfo.OriginalId,$migInfo.Id
    }

    #return updated json
    $json
}

#endregion

#region Dependency Functions
function Get-GraphDependencyDefaultObjects
{
    Add-GraphDependencyObjects @("ScopeTags")
}

function Get-GraphDependencyObjects
{
    param($objectType)

    if($global:chkReplaceDependencyIDs.IsChecked -ne $true -or -not $objectType -or -not $objectType.Dependencies -or (($objectType.Dependencies) | Measure).Count -eq 0) { return }
    
    $missingDeps = @()
    foreach($dep in $objectType.Dependencies)
    {
        if($global:LoadedDependencyObjects -isnot [HashTable] -or $global:LoadedDependencyObjects.ContainsKey($dep) -eq $false) 
        { 
            $missingDeps += $dep
        }
    }

    if($missingDeps.Count -eq 0) { return }

    Add-GraphDependencyObjects $missingDeps
}

function Add-GraphDependencyObjects
{
    param($DependencyIds)

    if($global:LoadedDependencyObjects -isnot [HashTable]) { $global:LoadedDependencyObjects = @{} }

    $importPath = $global:txtImportPath.Text
    $parentPath = [IO.Path]::GetDirectoryName($importPath)
    foreach($dep in $DependencyIds)
    {
        if($global:LoadedDependencyObjects.ContainsKey($dep)) { continue }

        $depObjectType = $global:currentViewObject.ViewItems | Where Id -eq $Dep

        if(-not $depObjectType)
        {
            Write-Log "No ViewItem found with Id $dep" 2
            continue
        }

        if([IO.Directory]::Exists(($importPath + "\" + $dep)))
        {
            $path = ($importPath + "\" + $dep)
        }
        elseif([IO.Directory]::Exists(($parentPath + "\" + $dep)))
        {
            $path = ($parentPath + "\" + $dep)
        }
        else
        {
            Write-Log "Export folder for dependency $dep not found" 2
            continue    
        }

        $depFiles = Get-GraphFileObjects $path -ObjectType $depObjectType
        
        $url = ($depObjectType.API + "?`$select=$((?? $depObjectType.IdProperty "Id")),$((?? $depObjectType.NameProperty "displayName"))")

        if($depObjectType.QUERYLIST)
        {
            $url = "$($url.Trim())&$($depObjectType.QUERYLIST.Trim())"
        }

        $depObjects = (Invoke-GraphRequest $url -ODataMetadata "none").Value
        $arrDepObjects = @()
        foreach($depObject in $depObjects)
        {
            $name = Get-GraphObjectName $depObject $depObjectType
            
            $fileObj = $depFiles | Where { (Get-GraphObjectName $_.Object $depObjectType) -eq $name }
            if(-not $fileObj)
            {
                Write-Log "Could not find an exported '$($depObjectType.Title)' object with name $name" 2
                continue
            }
            if(($fileObj | measure).Count -gt 1)
            {
                $fileObj = $fileObj[0]
                Write-Log "Multple files returned for object $name. Using first: $($fileObj.FileInfo.Name)" 2                
            }
            $arrDepObjects += New-Object PSObject -Property @{
                OriginalId = $fileObj.Object.Id
                Name = $name
                Id = Get-GraphObjectId $depObject $depObjectType
                Type = $depObjectType.Id
            }
        }

        if($arrDepObjects.Count -gt 0)
        {
            $global:LoadedDependencyObjects.Add($depObjectType.Id,$arrDepObjects)
        }
    }
}


#endregion

#region Import/Export/Copy functions

function Export-GraphObjects
{
    param([switch]$Selected)

    $objectType = $global:curObjectType
    Write-Status "Export $($objectType.Title)"

    $global:ExportRoot = (Get-XamlProperty $script:exportForm "txtExportPath" "Text")
    $folder = Get-GraphObjectFolder  $objectType $global:ExportRoot (Get-XamlProperty $script:exportForm "chkAddObjectType" "IsChecked") (Get-XamlProperty $script:exportForm "chkAddCompanyName" "IsChecked")

    $objectsToExport = @()
    if($Selected -ne $true)
    {
        # Export all
        $objectsToExport = $global:dgObjects.ItemsSource
    }
    elseif(($global:dgObjects.ItemsSource | Where IsSelected -eq $true).Count -gt 0)
    {
        # Export checked items
        $objectsToExport += ($global:dgObjects.ItemsSource | Where IsSelected -eq $true)
    }
    elseif($global:dgObjects.SelectedItem)
    {
        # Export selected item
        $objectsToExport += $global:dgObjects.SelectedItem
    }
    else 
    {
        return
    }

    foreach($obj in $objectsToExport)
    {
        Export-GraphObject $obj.Object $global:curObjectType $folder
    }

    Save-Setting "" "LastUsedFullPath" $folder
    Save-Setting "" "LastUsedRoot" $global:ExportRoot

    Write-Status ""
}

function Export-GraphObject
{
    param($objToExport, 
            $objectType, 
            $exportFolder)

    if(-not $exportFolder) { return }

    Write-Status "Export $((Get-GraphObjectName $objToExport $objectType))"

    $obj = Get-GraphExportObject $objToExport $objectType
    
    if(-not $obj)
    {
        Write-Log "No object to export" 3
        return
    }

    try 
    {
        if([IO.Directory]::Exists($exportFolder) -eq $false)
        {
            [IO.Directory]::CreateDirectory($exportFolder)
        }

        if($chkExportAssignments.IsChecked -ne $true -and $obj.Assignments)
        {
            Remove-Property $obj "Assignments"
        }

        $fileName = Get-GraphObjectName $obj $objectType
        if((Get-SettingValue "AddIDToExportFile") -eq $true -and $obj.Id)
        {
            $fileName = ($fileName + "_" + $obj.Id)
        }

        $obj | ConvertTo-Json -Depth 20 | Out-File -LiteralPath ([IO.Path]::Combine($exportFolder, (Remove-InvalidFileNameChars "$($fileName).json"))) -Force
        
        if($objectType.PostExportCommand)
        {
            & $objectType.PostExportCommand $obj $objectType $exportFolder
        }

        Add-GraphMigrationInfo $obj
    }
    catch 
    {
        Write-LogError "Failed to export object" $_.Exception
    }
}

function Get-GraphExportObject
{
    param($obj, $objectType)

    if($objectType.ExportFullObject -ne $false)
    {
        $exportObj = (Get-GraphObject $obj $objectType).Object
    }
    else
    {
        if($obj.Object)
        {
            $exportObj = $obj.Object
        }
        else
        {
            $exportObj = $obj    
        }
    }    
    $exportObj
}

function Import-GraphObject
{
    param($obj,
        $objectType,
        $fromFile)

    Write-Log "Import $($objectType.Title) object $((Get-GraphObjectName $obj $objectType))"
    
    # Clone the object before removing properties
    $objClone = $obj | ConvertTo-Json -Depth 20 | ConvertFrom-Json

    Start-GraphPreImport $obj $objectType

    $params = @{}
    $strAPI = (?? $objectType.APIPOST $objectType.API)
    $method = "POST"
    if($objectType.PreImportCommand)
    {
        $ret = & $objectType.PreImportCommand $obj $objectType $fromFile
        if($ret -is [HashTable])
        {
            if($ret.ContainsKey("Import") -and $ret["Import"] -eq $false)
            {
                # Import handled manually 
                return $false
            }

            if($ret.ContainsKey("API"))
            {
                $strAPI = $ret["API"]
            }
            
            if($ret.ContainsKey("Method"))
            {
                $method = $ret["Method"]
            }

            if($ret.ContainsKey("AdditionalHeaders") -and $ret["AdditionalHeaders"] -is [HashTable])
            {
                $params.Add("AdditionalHeaders",$ret["AdditionalHeaders"])
            }            
        }
    }

    $json = ConvertTo-Json $obj -Depth 20
    if($fromFile)
    {
        # Call Update-JsonForEnvironment before importing the object
        # E.g. PolicySets contains references, AppConfiguration policies reference apps etc.
        $json = Update-JsonForEnvironment $json
    }

    $newObj = (Invoke-GraphRequest -Url $strAPI -Content $json -HttpMethod $method @params)

    if($newObj -and $method -eq "POST")
    {
        Write-Log "$($objectType.Title) object imported successfully with id: $($newObj.Id)"
    }

    if($newObj -and $objectType.PostImportCommand)
    {
        & $objectType.PostImportCommand $newObj $objectType $fromFile
    }

    $newObj
}

function Remove-GraphObjects
{
    $objectsToDelete = @()
    if(($global:dgObjects.ItemsSource | Where IsSelected -eq $true).Count -gt 0)
    {
        # Delete checked items
        $objectsToDelete += ($global:dgObjects.ItemsSource | Where IsSelected -eq $true)
    }
    elseif($global:dgObjects.SelectedItem)
    {
        # Delete the selected item
        $objectsToDelete += $global:dgObjects.SelectedItem
    }

    if($objectsToDelete.Count -eq 0) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect items you want to delete", "Error", "OK", "Error") 
        return 
    }

    if(([System.Windows.MessageBox]::Show("Are you sure you want to delete $($objectsToDelete.Count) $($global:curObjectType.Title) object(s)?`n`nEnvironment: $($global:Organization.displayName)", "Delete Objects?", "YesNo", "Warning")) -ne "Yes")
    {
        return
    }

    foreach($tmpObj in $objectsToDelete)
    {
        Remove-GraphObject $tmpObj.Object $tmpObj.ObjectType
    }

    Show-GraphObjects
    Write-Status ""
}

function Remove-GraphObject
{
    param($objToRemove, $objectType)

    $strAPI = $null
    if($objectType.PreDeleteCommand)
    {
        $ret = & $objectType.PreDeleteCommand $objToRemove $objectType
        if($ret -is [HashTable])
        {
            if($ret.ContainsKey("Delete") -and $ret["Delete"] -eq $false)
            {
                # Delete handled manually or aborted
                return $false
            }

            if($ret.ContainsKey("API"))
            {
                $strAPI = $ret["API"]
            }
        }
    }

    if($strAPI)
    {
        $api = $strAPI
    }
    elseif($objectType.DELETEAPI)
    {
        $api = $objectType.DELETEAPI
    }
    else
    {
        $api = $objectType.API
    }

    Write-Status "Delete $((Get-GraphObjectName $objToRemove $objectType))"
    $strAPI = ($api + "/$($objToRemove.Id)")
    Write-Log "Delete $($objectType.Title) object $((Get-GraphObjectName $objToRemove $objectType))"
    Invoke-GraphRequest -Url $strAPI -HttpMethod "DELETE" -ODataMetadata "none"
}

function Copy-GraphObject
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect the $($global:curObjectType.Title) item you want to copy", "Error", "OK", "Error") 
        return 
    }

    $newName = "$((Get-GraphObjectName $dgObjects.SelectedItem $global:curObjectType)) - Copy"
    if($global:curObjectType.CopyDefaultName)
    {
        $newName = $global:curObjectType.CopyDefaultName
        $dgObjects.SelectedItem.PSObject.Properties | foreach { $newName =  $newName -replace "%$($_.Name)%", $dgObjects.SelectedItem."$($_.Name)" }
    }
    $ret = Show-InputDialog "Copy $($global:curObjectType.Title)" "Select name for the new object" $newName

    if($ret)
    {
        # Export profile
        Write-Status "Export $((Get-GraphObjectName $dgObjects.SelectedItem $global:curObjectType))"

        $exportObj = (Get-GraphObject $dgObjects.SelectedItem.Object $global:curObjectType -SkipAssignments).Object

        if($global:curObjectType.PreCopyCommand)
        {
            if((& $global:curObjectType.PreCopyCommand $exportObj $global:curObjectType $ret))
            {
                Show-GraphObjects
                Write-Status ""
                return
            }
        }

        # Convert to Json and back to clone the object
        $obj = ConvertTo-Json $exportObj -Depth 10 | ConvertFrom-Json
        if($obj)
        {
            # Import new profile
            Set-GraphObjectName $obj $global:curObjectType $ret

            $newObj = Import-GraphObject $obj $global:curObjectType
            if($newObj)
            {
                if($global:curObjectType.PostCopyCommand)
                {
                    & $global:curObjectType.PostCopyCommand $exportObj $newObj $global:curObjectType
                }
                Show-GraphObjects
            }
            else
            {
                [System.Windows.MessageBox]::Show("Failed to copy object. See log for more information", "Error", "OK", "Error") 
            }
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

#endregion

function Show-GraphObjectInfo
{
    param(
        $FormTitle = "",
        [switch]$NoLoadFull)

    if(-not $global:dgObjects.SelectedItem) { return }
    if(-not $global:dgObjects.SelectedItem.Object) { return }    
    
    $script:detailsForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\ObjectDetails.xaml")
    if(-not $script:detailsForm) { return }

    if(-not $FormTitle) { $FormTitle = $global:curObjectType.Title }
    $objName = Get-GraphObjectName $global:dgObjects.SelectedItem.Object $global:curObjectType
    if($objName)
    {
        $FormTitle = "$FormTitle - $objName"
    }

    if($global:curObjectType.DetailExtension)
    {
        & $global:curObjectType.DetailExtension $script:detailsForm "pnlButtons"
    }

    Set-XamlProperty  $script:detailsForm "txtValue" "Text" (ConvertTo-Json $global:dgObjects.SelectedItem.Object -Depth 10)
    
    if($global:curObjectType.AllowFullDetails -eq $false)
    {
        Set-XamlProperty  $script:detailsForm "btnFull" "Visibility" "Collapsed"
    }
    
    Add-XamlEvent $script:detailsForm "btnCopy" "Add_Click" -scriptBlock ([scriptblock]{ 
        $tmp = $script:detailsForm.FindName("txtValue")
        if($tmp.Text) { $tmp.Text | Set-Clipboard }
    })

    Add-XamlEvent $script:detailsForm "btnFull" "Add_Click" -scriptBlock ([scriptblock]{
        
        $obj = Get-GraphObject $global:dgObjects.SelectedItem.Object $global:curObjectType
        if($obj.Object)
        {
            Set-XamlProperty  $script:detailsForm "txtValue" "Text" (ConvertTo-Json $obj.Object -Depth 10)
            Set-XamlProperty  $script:detailsForm "btnFull" "IsEnabled" $false
        }
        Write-Status ""
    })

    Show-ModalForm $FormTitle $detailsForm
}

function Get-GraphObjectName
{
    param($obj, $objectType)

    $obj."$((?? ($objectType.NameProperty) "displayName"))"
}

function Set-GraphObjectName
{
    param($obj, $objectType, $value)

    $obj."$((?? ($objectType.NameProperty) "displayName"))" = $value
}

function Get-GraphObjectId
{
    param($obj, 
            $objectType)

    $obj."$((?? ($objectType.IdProperty) "Id"))"
}
function Get-GraphObjectFolder
{
    param($objectType, 
            $rootFolder,
            $addObjectType,
            $addOrganization)

    $path = $rootFolder

    if($addOrganization) { $path = Join-Path $path $global:organization.displayName }

    if($addObjectType -and $objectType.Id) { $path = Join-Path $path $objectType.Id }

    $path
}

function Add-GraphBulkMenu
{
    $menuItem = [System.Windows.Controls.MenuItem]::new()
    $menuItem.Header = "_Bulk"
    $menuItem.Name = "EMBulk"
    
    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "_Export"
    $subItem.Add_Click({Show-GraphBulkExportForm})  
    $menuItem.AddChild($subItem) | Out-Null
    
    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "_Import"
    $subItem.Add_Click({Show-GraphBulkImportForm})  
    $menuItem.AddChild($subItem) | Out-Null
    
    $subItem = [System.Windows.Controls.MenuItem]::new()
    $subItem.Header = "_Delete"
    $subItem.Name = "mnuBulkDelete"
    $allowBulkDelete = Get-SettingValue "EMAllowBulkDelete"
    # Add it hidden even if not enabled, the save settings will enable it
    $subItem.Visibility = (?: ($allowBulkDelete -eq $true) "Visible" "Collapsed")
    $subItem.Add_Click({Show-GraphBulkDeleteForm})
    $menuItem.AddChild($subItem) | Out-Null

    $mnuMain.Items.Insert(1,$menuItem) | Out-Null
}