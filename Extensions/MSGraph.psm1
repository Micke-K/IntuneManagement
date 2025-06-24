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
    '3.9.8a'
}

$global:MSGraphGlobalApps = @(
    (New-Object PSObject -Property @{Name="";ClientId="";RedirectUri="";Authority=""}),
    (New-Object PSObject -Property @{Name="Microsoft Graph PowerShell";ClientId="14d82eec-204b-4c2f-b7e8-296a70dab67e";RedirectUri="https://login.microsoftonline.com/common/oauth2/nativeclient";}),
    (New-Object PSObject -Property @{Name="Decomissioned - Don't use - Microsoft Intune PowerShell";ClientId="d1ddf0e4-d672-4dae-b554-9d5bdfd93547";RedirectUri="urn:ietf:wg:oauth:2.0:oob"; })
    )

$global:DefaultAzureApp = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

$global:OldAzureApps = @("d1ddf0e4-d672-4dae-b554-9d5bdfd93547")

function Invoke-InitializeModule
{
    $global:graphURL = "https://graph.microsoft.com/beta"

    $global:LoadedDependencyObjects = $null
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
            Name = "Replace with assignments (Preview)"
            Value = "replace_with_assignments"
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

    $global:appSettingSections += (New-Object PSObject -Property @{
        Title = "Silent/Batch Job"
        Id = "GraphSilent"
        Values = @()
    })

    $global:appSettingSections += (New-Object PSObject -Property @{
        Title = "MS Graph General"
        Id = "GraphGeneral"
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
        Title = "Add ID to export file"
        Key = "AddIDToExportFile"
        Type = "Boolean"
        DefaultValue = $false
        Description = "This will add object ID to the export file to support objects with the same name e.g. ObjectName_ObjectId.json"
    }) "ImportExport"
    
    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Use Batch API (Preview)"
        Key = "UseBatchAPI"
        Type = "Boolean"
        DefaultValue = $false
        Description = "This will use batch API to export up to 20 objects on each API call"
    }) "ImportExport"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Resolve reference info"
        Key = "ResolveReferenceInfo"
        Type = "Boolean"
        DefaultValue = $true
        Description = "This will export/import info for referenced/navigation properties eg certificates in VPN profiles etc."
    }) "ImportExport"
        
    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "ApplicationId"
        Key = "GraphAzureAppId"
        Type = "String"
        Description = "Azure App Id to use for log-in during silent operatons"
    }) "GraphSilent"
    
    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Secret"
        Key = "GraphAzureAppSecret"
        Type = "String"
        Description = "Secret for the Azure App"
    }) "GraphSilent"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Certificate"
        Key = "GraphAzureAppCert"
        Type = "String"
        Description = "Certificate for Azure App"
    }) "GraphSilent"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Login with App in UI (Preview)"
        Key = "GraphAzureAppLogin"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Login with specified app in the UI. Note: Change will require app restart"
    }) "GraphSilent"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Refresh Objects after copy"
        Key = "RefreshObjectsAfterCopy"
        Type = "Boolean"
        DefaultValue = $true
        Description = "This will refresh all objects when after a copy. If this is disabled, the list must be refreshed manually to see the new objects. Default is true"
    }) "GraphGeneral"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Show Delete button"
        Key = "EMAllowDelete"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Allow deleting individual objectes"
    }) "GraphGeneral"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Show Bulk Delete "
        Key = "EMAllowBulkDelete"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Allow using bulk delete to delete all objects of selected types"
    }) "GraphGeneral"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Expand assignments"
        Key = "ExpandAssignments"
        Type = "Boolean"
        DefaultValue = $false
        Description = "Expand assignments when listing objects. This can be used in custom columns based on assignment info"
    }) "GraphGeneral"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Use Graph 1.0 (Not Recommended)"
        Key = "UseGraphV1"
        Type = "Boolean"
        DefaultValue = $false
        Description = "This will use production verionof graph, v1.0. Note: Thot officially supported since this can have unpredicted results. Some parts will require Beta version of Graph."
    }) "GraphGeneral"

}

function Get-GraphAppInfo
{
    param($settingId, $defaultAppId, $prefix)

    if($global:hideUI -eq $true)
    {
        # Taken care of by authentication function
        return
    }

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
    Write-Log "Clear cached values"
    $global:MigrationTableCache = $null
    $global:MigrationTableCacheId = $null
    $global:LoadedDependencyObjects = $null
    $global:migFileObj = $null
    $global:AADObjectCache = $null
}

function Invoke-SettingsUpdated
{
    Initialize-GraphSettings
}

function Initialize-GraphSettings
{
    $script:defaultVersion = ""
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

            [ValidateSet("beta","v1.0")]
            $GraphVersion = "",

            [switch]
            $AllPages,

            [int]
            $PageSize = -1,

            [switch]
            $Batch,

            [switch]
            $NoError
        )

    if($SkipAuthentication -ne $true)
    {
        Connect-MSALUser
    }

    if(-not $GraphVersion) 
    {
        if(-not $script:defaultVersion)
        {
            if((Get-SettingValue "UseGraphV1") -eq $true)
            {
                $script:defaultVersion = "v1.0"
            }
            else
            {
                $script:defaultVersion = "beta"
            }
        }
        $GraphVersion = $script:defaultVersion 
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
            [IO.Directory]::CreateDirectory($dirName) | Out-Null
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
        $Url =  "https://$((?? $global:MSALGraphEnvironment "graph.microsoft.com"))/$GraphVersion/" + $Url.TrimStart('/')
        $Url = $Url -replace "%OrganizationId%", $global:Organization.Id
    }

    if($PageSize -gt 0 -and $url.IndexOf("`$top=") -eq -1)
    {
        if(($url.IndexOf('?')) -eq -1) 
        {
            $url = "$($url.Trim())?"
        }
        else
        {
            $url = "$($url.Trim())&"
        }
        $url = "$($url.Trim())`$top=$($PageSize)"
    }

    $proxyURI = Get-ProxyURI
    if($proxyURI)
    {
        $params.Add("proxy", $proxyURI)
        $params.Add("UseBasicParsing", $true)
    }

    $ret = $null
    
    $retryCount = 0
    $retryMax = 10
    do
    {
        $retryRequest = $false
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
            $retryCount++
            if($NoError -eq $true) { return }
            if($_.Exception.Response.StatusCode -eq 429 -and $retryCount -le $retryMax)
            {
                # NOT OK - Should use the date property but could not replicate the issue
                $retryCount++
                $retryRequest = $true
                Write-Log "429 - Too many requests received. Wait 5 s before retry" 2
                Start-Sleep -Seconds 5
            }
            else
            {
                $extMessage = $null
                try
                {
                    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $reader.BaseStream.Position = 0
                    $reader.DiscardBufferedData()
                    $response = $reader.ReadToEnd() | ConvertFrom-Json
                    if($response.Error.Message)
                    {
                        $extMessage = $response.Error.Message
                        try
                        {
                            if($response.Error.Message.StartsWith("{") -and $response.Error.Message.EndsWith("}"))
                            {
                                $message = $response.Error.Message | ConvertFrom-Json
                                if($message.Message)
                                {
                                    $extMessage = ". Response message: $($message.Message)"
                                }
                            }
                        }
                        catch {}

                        $extMessage = ". Response message: $($extMessage)"
                    }
                }
                catch{}

                Write-LogError "Failed to invoke MS Graph with URL $Url (Request ID: $requestId). Status code: $($_.Exception.Response.StatusCode)$extMessage" $_.Exception
            }            
        }
    } while($retryRequest -eq $true)
    
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
    [string]
    $select,
    [switch]
    $SinglePage,
    [switch]
    $AllPages,
    [switch]
    $SingleObject,
    [string]
    $filter)
        
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

    if(($url.IndexOf("`$select=")) -eq -1 -and $select)
    {
        $url += (?: (($url.IndexOf('?')) -eq -1) "?" "&")
        $url += "`$select=$select"
    }    
    
    if($SinglePage -eq $true)
    {
        #Use default page size or use below for a specific page size for testing
        #$params.Add("pageSize",5) #!!!
    }
    elseif($SingleObject -ne $true -and $SinglePage -ne $true)
    {
        $params.Add('AllPages',$true)
    }

    if($script:nextGraphPage -and ($SinglePage -eq $true -or $AllPages -eq $true))
    {
        $url = $script:nextGraphPage
    }

    if($SingleObject -ne $true -and (Get-SettingValue "ExpandAssignments") -eq $true -and $objectType.ExpandAssignmentsList -ne $false)
    {
        # Expand assignments so they can be used in custom columns
        if(($url.IndexOf('expand',[System.StringComparison]::InvariantCultureIgnoreCase)) -eq -1)
        {
            $url += (?: (($url.IndexOf('?')) -eq -1) "?" "&")
            $url = "$($url)`$expand=assignments"
        }
    }

    if($script:multipleGraphPages -eq $true -and $SingleObject -ne $true -and $filter -and $objectType.QuerySearch -eq $true)
    {
        # QuerySearch is only reqired when there are more pages to load
        if(($url.IndexOf('search',[System.StringComparison]::InvariantCultureIgnoreCase)) -eq -1)
        {
            $url += (?: (($url.IndexOf('?')) -eq -1) "?" "&")
            $url = "$($url)`$search=`"$($filter)`""
        }
    }     
    
    $graphObjects = Invoke-GraphRequest -Url $url @params
    if($SinglePage -eq $true -or $AllPages -eq $true)
    {
        $script:nextGraphPage = $graphObjects.'@odata.nextLink'
        if($null -eq $script:multipleGraphPages)
        {
            $script:multipleGraphPages = $null -ne $script:nextGraphPage
        }
    }

    if($graphObjects -and ($graphObjects | GM -Name Value -MemberType NoteProperty))
    {
        $retObjects = $graphObjects.Value
    }
    else
    {
        $retObjects = $graphObjects
    }

    if($retObjects)
    {
        $graphObjects = Add-GraphObjectProperties $retObjects $objectType $property $exclude $SortProperty

        if($SingleObject -ne $true -and $objectType.PostListCommand)
        {
            $graphObjects = & $objectType.PostListCommand $graphObjects $objectType
        }
    }
    else
    {
        $graphObjects = $null    
    }
    
    if(($graphObjects | measure).Count -gt 0)
    {
        $graphObjects
    }
}

function Add-GraphObjectProperties
{
    param($graphObjects, 
            $objectType, 
            [Array]
            $property = $null,
            [Array]
            $exclude = $null,
            $SortProperty = "displayName")

    if($property -isnot [Object[]]) { $property = @('displayName', 'description', 'id')}
    
    $objects = @()
    
    if($graphObjects -and ($graphObjects | GM -Name Value -MemberType NoteProperty))
    {
        $retObjects = $graphObjects.Value            
    }
    else
    {
        $retObjects = $graphObjects
    }

    $getAssignmentInfo = ((Get-SettingValue "ExpandAssignments") -eq $true -and $objectType.ExpandAssignmentsList -ne $false)

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
        
        if($null -ne $graphObject.isAssigned)
        {
            $objTmp | Add-Member -NotePropertyName "IsAssigned" -NotePropertyValue $graphObject.isAssigned
        }
        elseif($getAssignmentInfo)
        {
            $objTmp | Add-Member -NotePropertyName "IsAssigned" -NotePropertyValue (($graphObject.assignments | measure).Count -gt 0)
        }        
    }    

    if($objects.Count -gt 0 -and $SortProperty -and ($objects[0] | GM -MemberType NoteProperty -Name $SortProperty))
    {
        $objects = $objects | sort -Property $SortProperty
    }

    $objects
}

function Show-GraphObjects
{
    param($filter, [switch]$ObjectTypeChanged)

    $global:curObjectType = $global:lstMenuItems.SelectedItem

    if($ObjectTypeChanged -eq $true)
    {
        $script:multipleGraphPages = $null
    }

    Clear-GraphObjects

    if(-not $global:MSALToken)
    {
        $global:txtNotLoggedIn.Content = "Not logged in. Please login to view objects" 
        $global:grdNotLoggedIn.Visibility = "Visible"
        $global:grdData.Visibility = "Collapsed"
        return
    }
    elseif($global:curObjectType.'@AccessType' -eq "None")
    {
        $requiredPermissions = ($global:curObjectType.Permissons -join ",")
        $missingScopes = ?? $global:curObjectType.'@MissingScopes' $requiredPermissions
        if($requiredPermissions -ne $missingScopes)
        {
            $requiredPermissions = "`nRequired permissions: $requiredPermissions"
        }
        else
        {
            $requiredPermissions = ""
        }
        $global:txtNotLoggedIn.Content = "You don't have the required permissons to access $($global:curObjectType.Title).$($requiredPermissions)`n`Missing perimssons: $missingScopes`n`nRequest consent from the 'Request Consent' link in the user login info`nor`nDisable the 'Use Default Permissions' setting to trigger consent prompt.`nNote: Changing the 'Use Default Permissions' setting will require a restart of the app`nand a 'manual' login" 
        $global:grdNotLoggedIn.Visibility = "Visible"
        $global:grdData.Visibility = "Collapsed"
        return
    }    
    $global:grdNotLoggedIn.Visibility = "Collapsed"
    $global:grdData.Visibility = "Visible"

    # Always show Import if an item is selected
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

    $script:nextGraphPage = $null    

    [array]$graphObjects = Get-GraphObjects -property $global:curObjectType.ViewProperties -objectType $global:curObjectType -SinglePage -Filter $filter

    $dgObjects.AutoGenerateColumns = $false
    $dgObjects.Columns.Clear()

    if($graphObjects)
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
                Invoke-ModuleFunction "Invoke-EMSelectedItemsChanged"
            })           
        }

        $tableColumns = @()

        $additionalColumns = @()
        $additionalColsStr = ?? (Get-Setting "EndpointManager\ObjectColumns" "$($global:curObjectType.Id)") $global:curObjectType.DefaultColumns
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
        if($ctrl.Name -eq "btnDelete")
        {
            $allowDelete = Get-SettingValue "EMAllowDelete"
            if($global:currentViewObject.ViewInfo.AllowDelete -eq $false) { $allowDelete = $false }
            $ctrl.Visibility = (?: ($allowDelete -eq $true) "Visible" "Collapsed")
        }
        elseif(-not $global:curObjectType.ShowButtons -or ($global:curObjectType.ShowButtons | Where-Object { $ctrl.Name -like "*$($_)" } ))
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

    Set-GraphPagesButtonStatus
}

function Set-GraphPagesButtonStatus
{
    $global:btnLoadAllPages.Visibility = (?: ($script:nextGraphPage) "Visible" "Collapsed")
    $global:btnLoadNextPage.Visibility = (?: ($script:nextGraphPage) "Visible" "Collapsed")
    $global:btnLoadAllPages.Tag = $script:nextGraphPage
    $global:btnLoadNextPage.Tag = $script:nextGraphPage
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
    param($obj, $objectType, [switch]$SkipAssignments, [switch]$GetAPI)

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

    if($global:Organization.Id)
    {
        $api = $api -replace "%OrganizationId%", $global:Organization.Id
    }

    if($GetAPI -eq $true)
    {
        return $api
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
    foreach($odataProp in ($obj.PSObject.Properties | Where { $_.Name -like "*@Odata*Link" -or $_.Name -like "*@odata.context" -or $_.Name -like "*@odata.id" -or ($_.Name -like "*@odata.type" -and $_.Name -ne "@odata.type")})) # -or $_.Name -like "#CustomRef*"
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
        Write-Log "Load Graph MetaData file"
        $url = "https://graph.microsoft.com/beta/`$metadata"
        $fileFullPath = [Environment]::ExpandEnvironmentVariables("%LOCALAPPDATA%\CloudAPIPowerShellManagement\GraphMetaData.xml")
        $fi = [IO.FileInfo]$fileFullPath
        $maxAge = (Get-Date).AddDays(-14)
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
            Write-Log "Download Graph MetaData file"
            [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
            $wc = New-Object System.Net.WebClient
            $wc.Encoding = [System.Text.Encoding]::UTF8
            $proxyURI = Get-ProxyURI
            
            try 
            {
                if($proxyURI)
                {
                    $wc.Proxy = [System.Net.WebProxy]::new($proxyURI)
                }

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

        if(-not $global:metaDataXML -and $fi.Exists)
        {
            Write-Log "Using old version of Graph MetaData file" 2
            try 
            {
                [xml]$global:metaDataXML = Get-Content $fi.FullName              
            }
            catch { }
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

function Invoke-InitSilentBatchJob
{
    $global:MSALToken = $null

    if(-not $global:TenantId)
    {
        Write-Log "Tenant Id is missing. Use -TenantId <Tenant-guid> on the command line to run silent batch jobs" 3
        return
    }

    if(-not $global:AzureAppId -or (-not $global:ClientSecret -and -not $global:ClientCert))
    {
        # Get login info for silent job from settings
        $global:AzureAppId = Get-SettingValue "GraphAzureAppId" -TenantID $global:TenantId
        $global:ClientSecret = Get-SettingValue "GraphAzureAppSecret" -TenantID $global:TenantId
        $global:ClientCert = Get-SettingValue "GraphAzureAppCert" -TenantID $global:TenantId
    }

    if(-not $global:AzureAppId)
    {
        Write-Log "App Id is missing. Cannot run silent job without App Id. Either specify the AppId in Settings or Command Line (-AppId <AppId>)" 3
        return
    }    

    if(-not $global:ClientSecret -and -not $global:ClientCert)
    {
        Write-Log "Secret or Certificate must be specified. Either specify Secret/Certificate in Settings or Command Line" 3
        return
    }    
    Connect-MSALUser | Out-Null
    if(-not $global:MSALToken)
    {
        Write-Log "Not authenticated. Batch job will be skipped" 3
    }
    else
    {        
        $accessToken = Get-JWTtoken $global:MSALToken.AccessToken
        if($accessToken)
        {
            $global:Organization = (MSGraph\Invoke-GraphRequest -Url "Organization" -SkipAuthentication -ODataMetadata "Skip" -NoError).Value 
            if($global:Organization)
            {
                if($global:Organization -is [array]) { $global:Organization = $global:Organization[0]}
            }
            else
            {
                Write-Log "Could not get Organization info. Verify that the app has permission to read Organization info (at least Organization.Read.All). Organization name wil not be set" 2
                $global:Organization = [PSCustomObject]@{
                    Id = $accessToken.Payload.tid
                    displayName = ""
                }
            }
            if($global:Organization.displayName)
            {
                $tenantInfo = "$($global:Organization.displayName) ($($global:Organization.id))"
            }
            else
            {
                $tenantInfo = $accessToken.Payload.tid
            }

            Write-Log "Successfully authenticated to tenant: $tenantInfo"
            Write-Log "Azure App (for authentication): $($accessToken.Payload.app_displayname) ($($accessToken.Payload.appid))"
            Write-Log "Permissions: $(($accessToken.Payload.roles -join ","))"
        }
    }
}

function Invoke-SilentBatchJob
{
    param($settingsObj)

    if(-not $global:MSALToken) { return } # Skip if not authenticated

    if(-not $settingsObj -or (-not $settingsObj.BulkExport -and -not $settingsObj.BulkImport))
    {
        return 
    }

    if($settingsObj.BulkExport)
    {
        Start-GraphSilentBulkExport $settingsObj
    }

    if($settingsObj.BulkImport)
    {
        Start-GraphSilentBulkImport $settingsObj
    }
}

function Start-GraphSilentBulkExport
{
    param($settingsObj)

    $script:exportForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\BulkExportForm.xaml") -AddVariables
    if(-not $script:exportForm) { return }

    $script:exportObjects = Get-GraphBatchObjectTypes $settingsObj.BulkExport

    foreach($viewObj in $script:exportObjects)
    {
        if(-not $viewObj.Title) { continue }

        if($viewObj.ObjectType.ShowButtons -is [Object[]] -and $viewObj.ObjectType.ShowButtons -notcontains "Export") { continue }

        Add-GraphExportExtensions $script:exportForm 0 $viewObj.ObjectType
    }    

    Set-BatchProperties $settingsObj.BulkExport $script:exportForm

    $global:dgObjectsToExport.ItemsSource = @($script:exportObjects)

    <#
    # Select ObjectTypes based on batch config
    $objTypes = $settingsObj.BulkExport | Where Name -eq ObjectTypes
    if($objTypes)
    {        
        foreach($objTypeId in $objTypes.ObjectTypes)
        {
            $obj = $global:dgObjectsToExport.ItemsSource | Where { $_.ObjectType.Id -eq $objTypeId}
            if($obj)
            {
                $obj.Selected = $true
            }
            else
            {
                Write-Log "No Object Type with id $objTypeId found" 2                    
            }
        }
    }
    #>

    Start-GraphObjectExport
}

function Start-GraphSilentBulkImport
{
    param($settingsObj)

    $script:importForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\BulkImportForm.xaml") -AddVariables
    if(-not $script:importForm) { return }

    # Get all objects but not selected
    # This will allow dependencies
    $script:importObjects = Get-GraphBatchObjectTypes $settingsObj.BulkImport -NotSelected -All
    
    $objTypes = $settingsObj.BulkImport | Where Name -eq ObjectTypes
    if($objTypes)
    {        
        # Select object types from the batch file        
        foreach($objTypeId in $objTypes.ObjectTypes)
        {
            $obj = $script:importObjects | Where { $_.ObjectType.Id -eq $objTypeId}
            if($obj)
            {
                $obj.Selected = $true
            }
            else
            {
                Write-Log "No Object Type with id $objTypeId found" 2                    
            }
        }
    }
    
    foreach($viewObj in $script:importObjects)
    {
        if(-not $viewObj.Title) { continue }

        if($viewObj.ObjectType.ShowButtons -is [Object[]] -and $viewObj.ObjectType.ShowButtons -notcontains "Import") { continue }

        Add-GraphImportExtensions $script:importForm 0 $viewObj.ObjectType
    }    
    
    Set-BatchProperties $settingsObj.BulkImport $script:importForm

    $global:dgObjectsToImport.ItemsSource = @($script:importObjects)

    $importedObjects = Start-GraphObjectImport

    if($importedObjects -eq 0)
    {
        Write-Log "No objects were imported. Verify import batch file settings" 2
    }
}

function Get-GraphBatchObjectTypes
{
    param($settingsObj, [switch]$NotSelected, [switch]$All)

    $silentViewObjects = @()

    $intuneView = $global:viewObjects | Where { $_.ViewInfo.Id -eq "IntuneGraphAPI" }
    if($All -ne $true)
    {
        $arrObjectTypes = ($settingsObj | Where Name -eq ObjectTypes).ObjectTypes
    }
    else
    {
        $arrObjectTypes = $intuneView.ViewItems.Id
    }

    foreach($objTypeId in $arrObjectTypes)
    {
        $objType = $intuneView.ViewItems | Where Id -eq $objTypeId
        if(-not $objType) 
        {
            Write-Log "ViewObject with id $objTypeId not found" 2
            continue
        }

        $silentViewObjects += New-Object PSObject -Property @{
            Title = $objType.Title
            Selected = ($NotSelected.IsPresent -ne $true)
            ObjectType = $objType
        }
    } 
    
    $silentViewObjects
}

function Get-GraphObjectType
{
    param($objTypeId)

    $intuneView = $global:viewObjects | Where { $_.ViewInfo.Id -eq "IntuneGraphAPI" }

    if($intuneView)
    {
        ($intuneView.ViewItems | Where Id -eq $objTypeId)
    }
}

function Start-GraphObjectExport
{
    Write-Status "Export objects" -Block
    Write-Log "****************************************************************"
    Write-Log "Start bulk export"
    Write-Log "****************************************************************"

    $script:exportRoot = Expand-FileName (Get-XamlProperty $script:exportForm "txtExportPath" "Text")
    Write-Log "Export root folder: $script:exportRoot"

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
            $folder = Get-GraphObjectFolder $item.ObjectType $script:exportRoot (Get-XamlProperty $script:exportForm "chkAddObjectType" "IsChecked") (Get-XamlProperty $script:exportForm "chkAddCompanyName" "IsChecked")

            $folder = Expand-FileName $folder

            Write-Status "Get a list of all $($item.ObjectType.Title) objects" -SkipLog -Force
            [array]$objects = Get-GraphObjects -property $item.ObjectType.ViewProperties -objectType $item.ObjectType

            if((Get-SettingValue "UseBatchAPI") -eq $true)
            {
                # Use batch to get details of each object
                $batchObjects = Get-GraphBatchObjects $objects $txtNameFilter
                $i = 1
                $total = ($batchObjects | measure).Count
                foreach($batchResult in $batchObjects)
                {
                    if(-not $batchResult.Object) { continue }
                    $objName = Get-GraphObjectName $batchResult.Object $batchResult.ObjectType
                    Write-Status "Export $($item.Title): $objName ($($i)/$($total))" -Force
                    Export-GraphObject $batchResult.Object $batchResult.ObjectType $folder -IsFullObject
                    $i++
                }                
            }
            else        
            {
                foreach($obj in $objects)
                {
                    # Export objects one by one
                    $objName = Get-GraphObjectName $obj.Object $obj.ObjectType

                    if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                    {
                        continue
                    }

                    Write-Status "Export $($item.Title): $objName" -Force
                    Export-GraphObject $obj.Object $item.ObjectType $folder
                }
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

        Start-GraphObjectExport        
    })

    Add-XamlEvent $script:exportForm "btnExportSettingsForSilentExport" "add_click" ({
        $sf = [System.Windows.Forms.SaveFileDialog]::new()
        $sf.FileName = "BulkExport.json"
        $sf.DefaultExt = "*.json"
        $sf.Filter = "Json (*.json)|*.json|All files (*.*)|*.*"
        if($sf.ShowDialog() -eq "OK")
        {
            $tmp = [PSCustomObject]@{
                Name = "ObjectTypes"
                Type = "Custom"
                ObjectTypes = @()
            }
            foreach($ot in ($script:exportObjects | Where Selected -eq $true))
            {
                $tmp.ObjectTypes += $ot.ObjectType.Id
            }
            Export-GraphBatchSettings $sf.FileName $script:exportForm "BulkExport" @($tmp)
        }  
    })

    Show-ModalForm "Bulk Export" $script:exportForm -HideButtons
}

function Export-GraphBatchSettings
{
    param($fileName, $form, $batchType, $customProps)

    $script:childObjects = @()
    Get-XamlChildObjects $form

    $outputObj = [PSCustomObject]@{
        $batchType = @()
    }

    if($script:childObjects.Count -gt 0)
    {
        foreach($ctrl in $script:childObjects)
        {
            if(-not $ctrl.Name)
            {
                Write-Log  "Name not specified for a control with type: $(($ctrl.GetType().FullName)). Property skipped" 2
                continue
            }
            elseif($ctrl -is [System.Windows.Controls.TextBox])
            {
                $value = $ctrl.Text
            }
            elseif($ctrl -is [System.Windows.Controls.CheckBox])
            {
                $value = $ctrl.IsChecked
            }
            elseif($ctrl -is [System.Windows.Controls.ComboBox])
            {
                $value = $ctrl.SelectedValue
            }
            else
            {
                Write-Log "Unsupported control type: $(($ctrl.GetType().FullName)). Property skipped" 2
                continue    
            }
            #$focusable = $childObjects | Where Focusable -eq $true
            $outputObj.$batchType += [PSCustomObject]@{
                Name = $ctrl.Name
                Value =  $value
            }
        }
    }

    if(($customProps | measure).Count -gt 0)
    {
        $outputObj.$batchType += $customProps
    }

    $json = $outputObj | ConvertTo-Json -Depth 50
    $json | Out-File -LiteralPath $fileName -Force
}

function Get-XamlChildObjects
{
    param($parent)

    foreach($child in [System.Windows.LogicalTreeHelper]::GetChildren($parent))
    {
        if($child -is [System.Windows.DependencyObject])
        {
            if(($child.Focusable -eq $true) -and
                ($child -is [System.Windows.Controls.TextBox] -or 
                $child -is [System.Windows.Controls.CheckBox] -or
                $child -is [System.Windows.Controls.ComboBox]))
            {
                $script:childObjects += $child
            }
            Get-XamlChildObjects $child $collection
        }
    }
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
    
    Set-XamlProperty  $script:importForm "lblImportType" "Visibility" "Visible"
    Set-XamlProperty  $script:importForm "cbImportType" "Visibility" "Visible"

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
        #Get-GraphDependencyDefaultObjects
        $allowUpdate = $true 
        $filesToImport = $global:dgObjectsToImport.ItemsSource | Where Selected -eq $true
        if($global:curObjectType.PreFilesImportCommand)
        {
            $filesToImport = & $global:curObjectType.PreFilesImportCommand $global:curObjectType $filesToImport
        }

        $importedObjectsCurType = 0
        $navigationPropObjects = @()
        $arrImportedObjects = @()
        foreach ($fileObj in $filesToImport)
        {
            if($allowUpdate -and $global:cbImportType.SelectedValue -ne "alwaysImport" -and (Reset-GraphObject $fileObj $global:dgObjects.ItemsSource))
            {
                continue
            }

            $importedObj = Import-GraphFile $fileObj -PassThru
            if($importedObj -and $global:curObjectType.NavigationProperties -eq $true)
            {
                $navigationPropObjects += [PSCustomObject]@{
                    File =  $fileObj   
                    ImportedObject = $importedObj
                }
            }
            $arrImportedObjects += $importedObj
            $importedObjectsCurType++
        }

        if($global:curObjectType.PostFilesImportCommand)
        {
            & $global:curObjectType.PostFilesImportCommand $global:curObjectType $arrImportedObjects $filesToImport
        }

        if($importedObjectsCurType -gt 0 -and $global:LoadedDependencyObjects -is [HashTable] -and $global:LoadedDependencyObjects.ContainsKey($global:curObjectType.Id))
        {
            Write-Log "Remove $($global:curObjectType.Title) from dependency cache"
            $global:LoadedDependencyObjects.Remove($global:curObjectType.Id)
        }        

        if($navigationPropObjects)
        {
            foreach($navPropObj in $navigationPropObjects)
            {
                Set-GraphNavigationPropertiesFromFile $navPropObj
            }
        }
        Show-GraphObjects
        Show-ModalObject
        Write-Status ""
    }

    Add-XamlEvent $script:importForm "btnGetFiles" "add_click" {
        # Used when the user manually updates the path and the press Get Files
        $path = Expand-FileName $global:txtImportPath.Text
        $global:dgObjectsToImport.ItemsSource = @(Get-GraphFileObjects $path)
        if([IO.Directory]::Exists($path))
        {
            Save-Setting "" "LastUsedFullPath" $global:txtImportPath.Text
            Set-XamlProperty $script:importForm "lblMigrationTableInfo" "Content" (Get-MigrationTableInfo)
        }
    }

    Add-GraphImportExtensions $script:importForm 1 $global:curObjectType

    if($global:txtImportPath.Text)
    {
        $path = Expand-FileName $global:txtImportPath.Text
        $global:dgObjectsToImport.ItemsSource = @(Get-GraphFileObjects $path)
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
    
    Set-XamlProperty  $script:importForm "lblImportType" "Visibility" "Visible"
    Set-XamlProperty  $script:importForm "cbImportType" "Visibility" "Visible"        

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
        
        $importedObjects = Start-GraphObjectImport

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

    Add-XamlEvent $script:importForm "btnExportSettingsForSilentImport" "add_click" ({
        $sf = [System.Windows.Forms.SaveFileDialog]::new()
        $sf.FileName = "BulkImport.json"
        $sf.DefaultExt = "*.json"
        $sf.Filter = "Json (*.json)|*.json|All files (*.*)|*.*"
        if($sf.ShowDialog() -eq "OK")
        {
            $tmp = [PSCustomObject]@{
                Name = "ObjectTypes"
                Type = "Custom"
                ObjectTypes = @()
            }
            foreach($ot in ($script:importObjects | Where Selected -eq $true))
            {
                $tmp.ObjectTypes += $ot.ObjectType.Id
            }
            Export-GraphBatchSettings $sf.FileName $script:importForm "BulkImport" @($tmp)
        }  
    })    

    if((Get-XamlProperty $script:importForm "txtImportPath" "Text"))
    {
        Set-XamlProperty $script:importForm "lblMigrationTableInfo" "Content" (Get-MigrationTableInfo)
    }

    Show-ModalForm "Bulk Import" $script:importForm -HideButtons
}

function Start-GraphObjectImport
{
    Write-Status "Import objects" -Block
    Write-Log "****************************************************************"
    Write-Log "Start bulk import"
    Write-Log "****************************************************************"
    
    $tmpFolder = Expand-FileName (Get-XamlProperty $script:importForm "txtImportPath" "Text")
    Write-Log "Import root folder: $tmpFolder"

    $importedObjects = 0

    $txtNameFilter = $global:txtImportNameFilter.Text.Trim()
    Save-Setting "" "ImportNameFilter" $txtNameFilter
    if($txtNameFilter) { Write-Log "Name filter: $txtNameFilter" }

    $allowUpdate = $true
    
    foreach($item in ($script:importObjects | where Selected -eq $true | sort-object -property @{e={$_.ObjectType.ImportOrder}}))
    { 
        Write-Status "Import $($item.ObjectType.Title) objects" -Force
        Write-Log "----------------------------------------------------------------"
        Write-Log "Import $($item.ObjectType.Title) objects"
        Write-Log "----------------------------------------------------------------"
        $folder = Get-GraphObjectFolder $item.ObjectType (Get-XamlProperty $script:importForm "txtImportPath" "Text") (Get-XamlProperty $script:importForm "chkAddObjectType" "IsChecked")
        
        $folder = Expand-FileName $folder

        $graphObjects = $null        

        if($allowUpdate -and $global:cbImportType.SelectedValue -ne "alwaysImport")
        {           
            try 
            {
                Write-Status "Get $($item.Title) objects" -Force
                [array]$graphObjects = Get-GraphObjects -property $item.ObjectType.ViewProperties -objectType $item.ObjectType
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
            $navigationPropObjects = @()

            $importedObjectsCurType = 0

            $arrImportedObjects = @()

            foreach ($fileObj in @($filesToImport))
            {
                $objName = Get-GraphObjectName $fileObj.Object $item.ObjectType

                if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                {
                    continue
                }

                if($allowUpdate -and $global:cbImportType.SelectedValue -ne "alwaysImport" -and $graphObjects -and (Reset-GraphObject $fileObj $graphObjects))
                {
                    $importedObjects++ 
                    continue
                }
    
                $importedObj = Import-GraphFile $fileObj -PassThru
                if($importedObj -and $item.ObjectType.NavigationProperties -eq $true)
                {
                    $navigationPropObjects += [PSCustomObject]@{
                        File =  $fileObj   
                        ImportedObject = $importedObj
                    }
                }
                $arrImportedObjects = $importedObj

                $importedObjects++
                $importedObjectsCurType++
            }

            if($item.ObjectType.PostFilesImportCommand)
            {
                & $item.ObjectType.PostFilesImportCommand $item.ObjectType $arrImportedObjects $filesToImport
            }

            if($importedObjectsCurType -gt 0 -and $global:LoadedDependencyObjects -is [HashTable] -and $global:LoadedDependencyObjects.ContainsKey($item.ObjectType.Id))
            {
                Write-Log "Remove $($item.ObjectType.Title) from dependency cache"
                $global:LoadedDependencyObjects.Remove($item.ObjectType.Id)
            }
            Save-Setting "" "LastUsedFullPath" $folder
            if($navigationPropObjects)
            {
                foreach($navPropObj in $navigationPropObjects)
                {
                    Set-GraphNavigationPropertiesFromFile $navPropObj
                }
            }            
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

    $importedObjects
}

function Add-GraphExportExtensions
{
    param($form, $buttonIndex = 0, $objectTypes)
       
    #$global:curObjectType
    $grid = $form.FindName("grdExportProperties")

    foreach($objectType in $objectTypes)
    {
        if($objectType.ExportExtension)
        {            
            $extraProperties = & $objectType.ExportExtension $form "spExportSubMenu" 1
            for($i=0;($i + 1) -lt (($extraProperties) | measure).Count;$i ++) 
            {            
                $rd = [System.Windows.Controls.RowDefinition]::new()
                $rd.Height = [double]::NaN            
                $grid.RowDefinitions.Add($rd)
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::RowProperty,($grid.RowDefinitions.Count - 1))
                $grid.Children.Add($extraProperties[$i]) | Out-Null

                $i++            
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::RowProperty,($grid.RowDefinitions.Count - 1))
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
                $grid.Children.Add($extraProperties[$i]) | Out-Null

                if($extraProperties[$i].Name)
                {
                    $form.RegisterName($extraProperties[$i].Name, $extraProperties[$i])
                }
            }
        }
    }    
}

function Add-GraphImportExtensions
{
    param($form, $buttonIndex = 0, $objectTypes)
    
    $grid = $form.FindName("grdImportProperties")

    foreach($objectType in $objectTypes)
    {
        if($objectType.ImportExtension)
        {            
            $extraProperties = & $objectType.ImportExtension $form "spImportSubMenu" 1
            for($i=0;($i + 1) -lt (($extraProperties) | measure).Count;$i ++) 
            {            
                $rd = [System.Windows.Controls.RowDefinition]::new()
                $rd.Height = [double]::NaN            
                $grid.RowDefinitions.Add($rd)
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::RowProperty,$grid.RowDefinitions.Count - 1)
                $grid.Children.Add($extraProperties[$i]) | Out-Null

                $i++            
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::RowProperty,$grid.RowDefinitions.Count - 1)
                $extraProperties[$i].SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
                $grid.Children.Add($extraProperties[$i]) | Out-Null

                if($extraProperties[$i].Name)
                {
                    $form.RegisterName($extraProperties[$i].Name, $extraProperties[$i])
                }                
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

        foreach($item in ($global:dgBulkDeleteObjects.ItemsSource | Where Selected -eq $true | sort-object -property @{e={$_.ObjectType.ImportOrder}} -Descending))
        {
            Write-Log "----------------------------------------------------------------"
            Write-Log "Delete $($item.ObjectType.Title) objects"
            Write-Log "----------------------------------------------------------------"
            
            $txtNameFilter = $global:txtDeleteNameFilter.Text.Trim()
            Save-Setting "" "DeleteNameFilter" $txtNameFilter
            if($txtNameFilter) { Write-Log "Name filter: $txtNameFilter" }
            
            try 
            {
                Write-Status "Get $($item.ObjectType.Title) objects" -Force
                [array]$objects = Get-GraphObjects -property $item.ObjectType.ViewProperties -objectType $item.ObjectType
                foreach($obj in $objects)
                {                    
                    $objName = Get-GraphObjectName $obj.Object $obj.ObjectType

                    if($txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
                    {
                        continue
                    }

                    Write-Status "Delete $($item.ObjectType.Title): $objName" -Force -SkipLog
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
            $graphObj = Get-GraphObjectFromFile $file.FullName
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
    param($file, $objectType, [switch]$PassThru) 

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
        $objClone = $file.Object | ConvertTo-Json -Depth 50 | ConvertFrom-Json

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

        if($newObj)
        {
            $file | Add-Member -NotePropertyName "ImportedObject" -NotePropertyValue $newObj
        }

        if($PassThru -eq $true -and $newObj)
        {
            $newObj
        }
    } 
    catch 
    {
        Write-LogError "Failed to import file '$($file.FileInfo.Name)'" $_.Exception        
    }
}

function Reset-GraphObject
{ 
    param($fileObj, $objectList)

    $nameProp = ?? $fileObj.ObjectType.NameProperty "displayName"
    $curObject = $objectList | Where { $_.Object.$nameProp -eq $fileObj.Object.$nameProp -and $_.Object.'@OData.Type' -eq $fileObj.Object.'@OData.Type' }
    
    if($global:cbImportType.SelectedValue -eq "skipIfExist" -and ($curObject | measure).Count -gt 0 -and $fileObj.ObjectType.AlwaysImport -ne $true)
    {
        Write-Log "Object with name $($fileObj.Object.$nameProp) already exists. Object will not be imported"
        return $true
    }
    elseif(($curObject | measure).Count -gt 1)
    {
        Write-Log "Multiple objects return with name $($fileObj.Object.$nameProp). Object will not be imported or replaced" 2
        return $true
    }
    elseif(($curObject | measure).Count -eq 1)
    {
        $idInfo = ""
        if([String]::IsNullOrEmpty($curObject.Object.Id) -eq $false)
        {
            $idInfo = " with id $($curObject.Object.Id)"
        }
        Write-Log "Update $((Get-GraphObjectName $fileObj.Object $fileObj.ObjectType))$idInfo"
        $objectType = $fileObj.ObjectType

        # Clone the object before removing properties
        $obj = $fileObj.Object | ConvertTo-Json -Depth 50 | ConvertFrom-Json
        Start-GraphPreImport $obj $objectType
        if ($global:cbImportType.SelectedValue -ne "replace_with_assignments"){  
            # will use the assignments from the file for "replace_with_assignments" type
            Remove-Property $obj "Assignments"
        }
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

            $json = ConvertTo-Json $obj -Depth 50
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
        elseif($global:cbImportType.SelectedValue -in @("replace","replace_with_assignments"))
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
                if ($global:cbImportType.SelectedValue -eq "replace")
                {
                    Import-GraphObjectAssignment $newObj $objectType $curObject.Object.Assignments $fileObj.FileInfo.FullName -CopyAssignments | Out-Null
                }
                else {
                    Import-GraphObjectAssignment $newObj $objectType $obj.Assignments $file.FileInfo.FullName | Out-Null
                }

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
    $clonedAssignments = $assignments | ConvertTo-Json -Depth 50 | ConvertFrom-Json

    if($objectType.PreImportAssignmentsCommand)
    {
        $preConfig = & $objectType.PreImportAssignmentsCommand $obj $objectType $fromFile $clonedAssignments
    }

    if($preConfig -isnot [Hashtable]) { $preConfig = @{} }

    if($preConfig["Import"] -eq $false) { return } # Assignment managed manually so skip further processing

    $api = ?? $preConfig["API"] "$($objectType.API)/$($obj.Id)/assign"

    $method = ?? $preConfig["Method"] "POST"

    $clonedAssignments = ?? $preConfig["Assignments"] $clonedAssignments

    $keepProperties = ?? $objectType.AssignmentProperties @("target")
    $keepTargetProperties = ?? $objectType.AssignmentTargetProperties @("@odata.type","groupId","deviceAndAppManagementAssignmentFilterId","deviceAndAppManagementAssignmentFilterType")
    
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

    $json = $htAssignments | ConvertTo-Json -Depth 50
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

    if(($obj.PSObject.Properties | Where Name -eq "roleScopeTagIds"))
    {
        $scopeTagProperty = "roleScopeTagIds"
    }
    elseif(($obj.PSObject.Properties | Where Name -eq "roleScopeTags"))
    {
        $scopeTagProperty = "roleScopeTags"
    }
    else { return }

    $scopesIds = @()
    if($global:chkReplaceDependencyIDs.IsChecked -eq $false -and $global:chkReplaceDependencyIDs.IsEnabled -eq $false)
    {
        if($global:chkImportScopes.IsChecked -eq $true) 
        {
            $scopesIds += $obj.$scopeTagProperty
        }
    }
    else
    {    
        $loadedScopeTags = $global:LoadedDependencyObjects["ScopeTags"]
        $usingDefault = (($obj."$scopeTagProperty" | measure).Count -eq 1 -and ($obj."$scopeTagProperty")[0] -eq "0")
        if($loadedScopeTags -and $global:chkImportScopes.IsChecked -eq $true -and $usingDefault -eq $false -and $loadedScopeTags)
        {        
            foreach($scopeId in $obj."$scopeTagProperty")
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
    }

    if($scopesIds.Count -eq 0)
    {
        $scopesIds += "0" # Import with Default ScopeTag as default.
    }
    $obj."$scopeTagProperty" = $scopesIds
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
                $objType -eq "#microsoft.graph.exclusionGroupAssignmentTarget" -or
                $objType -eq "#microsoft.graph.cloudPcManagementGroupAssignmentTarget")
            {
                Add-GraphMigrationObject $objInfo.groupid "/groups" "Group"
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

    $path = Get-GraphMigrationTableFile $script:ExportRoot

    if(-not $path) { return }

    $path = Expand-FileName $path

    # Check if group is already processed
    $groupObj = Get-GraphMigrationObject $groupId
    if(-not $groupObj)
    {
        # Get group info
        $groupObj = Invoke-GraphRequest "/groups/$groupId" -ODataMetadata "none" -NoError
    }

    if($groupObj)
    {
        # Add group to cache
        if($global:AADObjectCache.ContainsKey($groupId) -eq $false) { $global:AADObjectCache.Add($groupId, $groupObj) }

        # Add group to migration file
        if((Add-GraphMigrationObjectToFile $groupObj $path "Group"))
        {
            # Export group info to json file for possible import
            $grouspPath = Join-Path $path "Groups"
            if(-not (Test-Path $grouspPath)) { mkdir -Path $grouspPath -Force -ErrorAction SilentlyContinue | Out-Null }
            $fileName = "$grouspPath\$((Remove-InvalidFileNameChars $groupObj.displayName)).json"
            Save-GraphObjectToFile $groupObj $fileName
        }
    }
    else 
    {        
        Write-Log "No group found with ID $($groupId). It might be deleted." 2
    }
}

function Add-GraphMigrationObject
{
    param($objId, $grapAPI, $objTypeName)

    if(-not $objId) { return }

    $path = Get-GraphMigrationTableFile $script:ExportRoot

    if(-not $path) { return }

    $path = Expand-FileName $path

    # Check if object is already processed
    $graphObj = Get-GraphMigrationObject $objId
    if(-not $graphObj -and ($global:AADObjectCache.ContainsKey($objId) -eq $false))
    {
        # Get object info
        $graphObj = Invoke-GraphRequest "$($grapAPI)/$objId" -ODataMetadata "none" -NoError
    }

    if($graphObj)
    {
        # Add object to cache
        if($global:AADObjectCache.ContainsKey($objId) -eq $false) { $global:AADObjectCache.Add($objId, $graphObj ) }

        # Add object to migration file
        if((Add-GraphMigrationObjectToFile $graphObj $path $objTypeName))
        {
            if($objTypeName -eq "Group")
            {
                # Export group info to json file for possible import
                $grouspPath = Join-Path $path "Groups"
                if(-not (Test-Path $grouspPath)) { mkdir -Path $grouspPath -Force -ErrorAction SilentlyContinue | Out-Null }
                $fileName = "$grouspPath\$((Remove-InvalidFileNameChars $graphObj.displayName)).json"
                Save-GraphObjectToFile $graphObj $fileName
            }
        }
    }
    else
    {
        if($global:AADObjectCache.ContainsKey($objId) -eq $false) { $global:AADObjectCache.Add($objId, $null) }
        Write-Log "No $objTypeName found with ID $($objId). It might be deleted." 2
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
function Add-GraphMigrationObjectToFile
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
    ConvertTo-Json $global:migFileObj -Depth 50 | Out-File $migFileName -Force

    $true # New object was added
}

function Get-GraphMigrationTableForImport
{
    $global:GraphMigrationTable = $null
    # Migration table must be located in the root of the import path
    $path = $global:txtImportPath.Text
    $path = Expand-FileName $path
    
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
                $migTableGroupName = $migObj.DisplayName.Trim()
                $obj = (Invoke-GraphRequest "/groups?`$filter=displayName eq '$($migTableGroupName)'").Value
                if(-not $obj)
                {
                    $groupFi = $null
                    if($global:GraphMigrationTable)
                    {
                        $fi = [IO.FileInfo]$global:GraphMigrationTable
                        $groupFi = [IO.FileInfo]($fi.DirectoryName + "\Groups\$((Remove-InvalidFileNameChars $migTableGroupName)).json")
                    }

                    if($groupFi.Exists -eq $true)
                    {
                        # ToDo: Create group from Json (could be a dynamic group)
                        # Warn if synched group
                        $groupObj = Get-GraphObjectFromFile $groupFi.FullName

                        #isAssignableToRole - For Role assignment groupd.
                        $keepProps = @("displayName","description","mailEnabled","mailNickname","securityEnabled","membershipRule","groupTypes", "membershipRuleProcessingState")
                        foreach($prop in $groupObj.PSObject.Properties)
                        {
                            if($prop.Name -in $keepProps) { continue }
                            
                            Remove-Property $groupObj $prop.Name
                        }
                        $groupObj.displayName = $groupObj.displayName.Trim()
                        $groupJson = ConvertTo-Json $groupObj -Depth 50
                    }
                    else
                    {
                        $groupName = $migTableGroupName
                        Write-Log "No group object found for $groupName. Creating a cloud group with default settings" 2
                        $dateStr = ((Get-Date).ToString("yyMMddHHmmss"))
                        
                        if(($groupName.Length + $dateStr.Length) -gt 64)
                        {
                            $nickName = $groupName.Substring(0,(64-$dateStr.Length))
                        }
                        else
                        {
                            $nickName = $groupName
                        }
                        $nickName = $nickName + $dateStr
                        
                        $groupJson = @"
                        { 
                            "displayName": "$($groupName)",
                            "mailEnabled": false,
                            "mailNickname": "$($nickName)",
                            "securityEnabled": true         
                        }
"@
                    }
                    Write-Log "Create AAD Group $($migTableGroupName)"
                
                    $obj = Invoke-GraphRequest "/groups" -HttpMethod "POST" -Content $groupJson
                }

                if($obj)
                {
                    $global:MigrationTableCache += (New-Object PSObject -Property @{
                        OriginalId = $migObj.Id            
                        Id = $obj.Id
                        Type = $migObj.Type    
                    })
                }
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
    Add-GraphDependencyObjects @("ScopeTags","AssignmentFilters")
}

function Get-GraphDependencyObjects
{
    param($objectType)

    Get-GraphDependencyDefaultObjects

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

        $depObjectType = $global:viewObjects.ViewItems | Where Id -eq $Dep

        if(-not $depObjectType)
        {
            Write-Log "No ViewItem found with Id $dep" 2
            $global:LoadedDependencyObjects.Add($dep,$null)
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
            $global:LoadedDependencyObjects.Add($depObjectType.Id,$null)
            continue    
        }

        $depFiles = Get-GraphFileObjects $path -ObjectType $depObjectType
        
        $url = ($depObjectType.API + "?`$select=$((?? $depObjectType.IdProperty "Id")),$((?? $depObjectType.NameProperty "displayName"))")

        if($depObjectType.QUERYLIST)
        {
            $url = "$($url.Trim())&$($depObjectType.QUERYLIST.Trim())"
        }

        $depObjects = (Invoke-GraphRequest $url -ODataMetadata "none" -AllPages).Value
        $arrDepObjects = @()
        foreach($depObject in $depObjects)
        {
            $name = Get-GraphObjectName $depObject $depObjectType
            
            $fileObj = $depFiles | Where { (Get-GraphObjectName $_.Object $depObjectType) -eq $name }
            if(-not $fileObj)
            {
                Write-Log "Could not find an exported '$($depObjectType.Title)' object with name $name" 2
                $arrDepObjects += New-Object PSObject -Property @{
                    OriginalId = $null
                    Name = $name
                    Id = Get-GraphObjectId $depObject $depObjectType
                    Type = $depObjectType.Id
                }                
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
        else
        {
            $global:LoadedDependencyObjects.Add($depObjectType.Id,$null)
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

    $script:ExportRoot = (Get-XamlProperty $script:exportForm "txtExportPath" "Text")
    $folder = Get-GraphObjectFolder  $objectType $script:ExportRoot (Get-XamlProperty $script:exportForm "chkAddObjectType" "IsChecked") (Get-XamlProperty $script:exportForm "chkAddCompanyName" "IsChecked")
    
    $folder = Expand-FileName $folder

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
    Save-Setting "" "LastUsedRoot" $script:ExportRoot

    Write-Status ""
}

function Export-GraphObject
{
    param($objToExport, 
            $objectType, 
            $exportFolder,
            [switch]$IsFullObject,
            [switch]$SkipAddID,
            [switch]$PassThru)

    if(-not $exportFolder -or -not $objToExport -or -not $objectType) { return }

    Write-Status "Export $((Get-GraphObjectName $objToExport $objectType))"

    if($IsFullObject -eq $true)
    {
        $obj = $objToExport
    }
    else
    {
        $obj = Get-GraphExportObject $objToExport $objectType        
    }
    
    if(-not $obj)
    {
        Write-Log "No object to export" 3
        return
    }
    
    Add-GraphNavigationProperties $obj $objectType 

    try 
    {
        if([IO.Directory]::Exists($exportFolder) -eq $false)
        {
            [IO.Directory]::CreateDirectory($exportFolder) | Out-Null
        }

        if($global:chkExportAssignments.IsChecked -ne $true -and $obj.Assignments)
        {
            Remove-Property $obj "Assignments"
        }

        $fileName = (Get-GraphObjectName $obj $objectType).Trim('.')
        if($SkipAddID -ne $true -and (Get-SettingValue "AddIDToExportFile") -eq $true -and $obj.Id -and $objectType.SkipAddIDOnExport -ne $true)
        {
            $fileName = ($fileName + "_" + $obj.Id)
        }

        $fullPath = ([IO.Path]::Combine($exportFolder, (Remove-InvalidFileNameChars "$($fileName).json")))
        Save-GraphObjectToFile $obj $fullPath
        
        if($objectType.PostExportCommand)
        {
            & $objectType.PostExportCommand $obj $objectType $exportFolder
        }

        Add-GraphMigrationInfo $obj

        if($PassThru -eq $true)
        {
            $fullPath
        }
    }
    catch 
    {
        Write-LogError "Failed to export object" $_.Exception
    }
}

<#
    Update the navigation references for an object
#>
function Set-GraphNavigationPropertiesFromFile
{
    param($navPropObject)

    if(-not $navPropObject.File -or -not $navPropObject.ImportedObject)
    {
        return
    }

    # Reload data from file. Some object properties was removed before import...
    $objFileInfo = Get-GraphObjectFromFile $navPropObject.File.FileInfo.FullName

    if(-not ($objFileInfo.PSObject.Properties | Where { $_.Name -like "#CustomRef_*" })) { return }

    Set-GraphNavigationProperties $navPropObject.ImportedObject $objFileInfo $navPropObject.File.ObjectType
}

function Set-GraphNavigationProperties
{
    param($newObj, $oldObj, $objectType, [switch]$FromOldObject)

    if($objectType.NavigationProperties -ne $true) { return }

    if(-not $newObj -or -not $oldObj -or -not $objectType)
    {
        return
    }

    if((Get-SettingValue "ResolveReferenceInfo") -ne $true) { return }

    $entityName = $oldObj.'@odata.type'.Split('.')[-1]

    $nameProp = ?? $objectType.NameProperty "displayName"

    $props = Get-GraphEntityTypeProperties $entityName
    
    foreach($prop in ($props | Where LocalName -eq "NavigationProperty" ))
    {
        # Is this the correct way of filter out Assignments, summaries etc.?
        if($prop.ContainsTarget -eq $true) { continue }


        if(-not ($oldObj."$($prop.Name)@odata.associationLink")) { continue }

        $associationLink = $oldObj."$($prop.Name)@odata.associationLink" -replace $oldObj.Id,$newObj.Id
        $refBodyObjs = @()
        $refObjName = $null
        $refObjId = $null
        if($prop.Type -like "Collection(*")
        {
            $multiNavProperty = $true
            $method = "POST"
        }
        else
        {
            $multiNavProperty = $false
            $method = "PUT"
        }

        if($FromOldObject -eq $true)
        {
            $navProp = Invoke-GraphRequest -URL $oldObj."$($prop.Name)@odata.navigationLink" -ODataMetadata "minimal" -NoError

            if(-not $navProp) { continue }
            
            if($multiNavProperty)
            {
                $navProperties = $navProp.Value                
            }
            else
            {
                $navProperties = $navProp
            }

            foreach($navProp in $navProperties)
            {
                $refBodyObjs += [PSCustomObject]@{                    
                    RefObjName = $navProp.displayName ### NOT Correct. Migh be another property but we don't know the type
                    RefObjId = $navProp.Id
                    RefBody = ([PSCustomObject]@{
                        "@odata.id" = ("https://$global:MSALGraphEnvironment/beta/$($objectType.API)('$($navProp.Id)')")
                    })
                }
            }
        }
        else
        {
            if(-not ($oldObj."#CustomRef_$($prop.Name)")) { continue } # Not included in the export file

            $idx = $oldObj."#CustomRef_$($prop.Name)".IndexOf("|:|")
            if($idx -gt -1)
            {
                $refObjNames = $oldObj."#CustomRef_$($prop.Name)".SubString(0,$idx)
            }
            else
            {
                $refObjNames = $oldObj."#CustomRef_$($prop.Name)"
            }

            foreach($refObjName in $refObjNames.Split(","))
            {            
                $refObjects = Invoke-GraphRequest -URL "$($objectType.API)?`$filter=$($nameProp) eq '$($refObjName)'" -NoError

                $objectsFound = ($refObjects.value | measure).Count

                if($objectsFound -eq 1)
                {
                    # Are there any references that allows multiple ref objects?                
                    foreach($refObj in $refObjects.value)
                    {
                        $refBodyObjs += [PSCustomObject]@{
                            RefObjName = $refObjName
                            RefObjId = $refObj.Id
                            RefBody = ([PSCustomObject]@{
                                "@odata.id" = ("https://$global:MSALGraphEnvironment/beta/$($objectType.API)('$($refObj.Id)')")
                            })
                        }
                    }
                }
                elseif($objectsFound -gt 1)
                {
                    Write-Log "Multiple objects ($objectsFound) found with $nameProp $refObjName. Skipping reference." 2
                    continue
                }
                else
                {
                    Write-Log "No object found with $nameProp $refObjName" 2
                    continue
                }
            }
        }

        foreach($refObject in $refBodyObjs)
        {
            Write-Log "Add $($refObject.RefObjName) ($($refObject.RefObjId)) to navigation property $($prop.Name)"
            $body = $refObject.RefBody | ConvertTo-Json -Depth 50
            Invoke-GraphRequest -URL $associationLink -HttpMethod $method -Content $body | Out-Null
        }
    }    
}


<#
    Add Navigation Property data to the object so they are included in the exported json file
#>
function Add-GraphNavigationProperties
{
    param($obj, $objType)
    
    if($objectType.NavigationProperties -ne $true) { return }

    if(-not $obj.'@odata.type') { return }

    if((Get-SettingValue "ResolveReferenceInfo") -ne $true) { return }

    $entityName = $obj.'@odata.type'.Split('.')[-1]

    $props = Get-GraphEntityTypeProperties $entityName

    foreach($prop in ($props | Where LocalName -eq "NavigationProperty" ))
    {
        # Is this the correct way of filter out Assignments, summaries etc.?
        if($prop.ContainsTarget -eq $true) { continue }

        if(-not ($obj."$($prop.Name)@odata.navigationLink")) { continue }
        $navProp = Invoke-GraphRequest -URL $obj."$($prop.Name)@odata.navigationLink" -ODataMetadata "minimal" -NoError
        if($navProp)
        {
            $value = $null
            $refType = ""
            if($navProp.value -is [Object[]])
            {
                if($navProp.value.Count -gt 0 -and $navProp.value[0].'@odata.type') { $refType = $navProp.value[0].'@odata.type' }
                $refValues = @()
                $navProp.value | ForEach-Object { $refValues += (Get-GraphObjectName $_ $objType) }
                if($refValues.Count -gt 0)
                {
                    if(($refValues -join "") -like "*,*")
                    {
                        Write-Log "One or mor referenced objects has the comma (,) character in the name. Cannot add navigation property $($prop.Name)" 3
                    }
                    $value = ($refValues -join ",") 
                }
            }
            else
            {
                if($navProp.'@odata.type') { $refType = $navProp.'@odata.type' }
                $value = (Get-GraphObjectName $navProp $objType)
            }
            if($refType -and $value)
            {
                $value = ($value + "|:|" + $refType)
            }
            
            $obj | Add-Member -NotePropertyName "#CustomRef_$($prop.Name)" -NotePropertyValue $value
        }
    }
}

function Get-GraphBatchObjects
{
    param($objects, $txtNameFilter)

    $batchResults = @()
    $batchArr = @()
    $skipped = 0
    $objectType = $null

    foreach($obj in $objects)
    {        
        $objectType = $obj.ObjectType
        $objName = Get-GraphObjectName $obj.Object $obj.ObjectType

        if($objName -and $txtNameFilter -and $objName -notmatch [RegEx]::Escape($txtNameFilter))
        {
            $skipped++
        }
        else
        {
            $ometadata = ?? $obj.ObjectType.ODataMetadata "Full"
            $batchArr += [PSCustomObject]@{
                id = ($batchArr.Count + 1)
                method = "GET"
                url = (Get-GraphObject $obj.Object $obj.ObjectType -GetAPI) 
                headers = @{"Accept"="application/json;odata.metadata=$ometadata"}
            }
        }
    }
    
    if($batchArr.Count -eq 0) { return }

    $batchResults = @((Invoke-GraphBatchRequest $batchArr $objectType.Title).body)

    if(($batchResults | measure).Count -ne ($objects.Count - $skipped))
    {
        Write-Log "Not all batch objects returned. Expected $($objects.Count - $skipped) but only got $(($batchResults | measure).Count)"
    }    

    if($objectType -and ($batchResults | measure).Count -gt 0)
    {
        $batchResultsTmp = $batchResults
        $batchResults = Add-GraphObjectProperties $batchResultsTmp $objectType -property $objectType.ViewProperties
        
        $curObj = 1
        foreach($obj in $batchResults)
        {
            if($obj.Object -and $obj.ObjectType.PostGetCommand)
            {
                Write-Status "Run PostGetCommand - $((Get-GraphObjectName $obj.Object $obj.ObjectType)) ($($curObj)/$(@($batchResults).Count))" -Force
                & $obj.ObjectType.PostGetCommand $obj $obj.ObjectType
            }
            $curObj++
        }
    }
    $batchResults
}

function Invoke-GraphBatchRequest
{
    param($batchObjects, $batchType, [switch]$SkipWarnings, [switch]$IncludedFailed)

    $batchArr = @()
    $batchResults = @()
    $batchTotal = 0
    $curBatch = 1

    foreach($obj in $batchObjects)
    {
        $batchArr += $obj

        if($batchArr.Count -eq 20 -or (($batchTotal + $batchArr.Count) -eq $batchObjects.Count))
        {            
            $batchObj = [PSCustomObject]@{
                requests = @($batchArr)
            }

            Write-Status "Get batch $curBatch $batchType" -Force

            $batchTotal += $batchArr.Count
            $json = $batchObj | ConvertTo-Json -Depth 50
            $maxRetryCount = 10
            $curRetry = 0
            
            do
            {                
                $retry = $false
                $retryArr = @()
                $retryAfter = 0
                $tmpResults = Invoke-GraphRequest -Url "`$batch" -Body $json -Method "POST"

                foreach($batchResult in ($tmpResults.responses | Sort -Property Id))
                {
                    if($batchResult.Status -ge 300 -or -not $batchResult.body)
                    {
                        $reqObj = $batchObj.requests | where id -eq $batchResult.Id
                        if($batchResult.Status -eq 429 -and $reqObj)
                        {                 
                            if($batchResult.headers.'Retry-After' -and $batchResult.headers.'Retry-After' -gt $retryAfter)
                            {           
                                try
                                {
                                    $retryAfter = [int]$batchResult.headers.'Retry-After'
                                }
                                catch{}
                            }
                            $retryArr += $reqObj
                        }
                        else
                        {
                            if($SkipWarnings -ne $true)
                            {
                                Write-Log "Batch result $($batchResult.Status) for URL $($reqObj.URL). Skipping..." 2
                            }

                            if($IncludedFailed -eq $true)
                            {
                                $batchResults += $batchResult
                            }
                        }
                        continue
                    }
                    $batchResults += $batchResult
                }

                if($retryArr.Count -gt 0)
                {
                    $curRetry++
                    if($curRetry -gt $maxRetryCount)
                    {
                        Write-Log "Max retry reached for batch process. Aborting..." 3                        
                    }
                    else
                    {
                        if($retryAfter -lt 5) { $retryAfter = 5 }
                        Write-Log "Batch result returned 429 - 'Too many requests'. Retrying $($retryArr.Count). Wait for $($retryAfter) seconds." 2
                        $retry = $true
                        $tmpBatchObj = [PSCustomObject]@{
                            requests = $retryArr
                        }
                        $json = $tmpBatchObj | ConvertTo-Json -Depth 50
                        Start-Sleep -Seconds $retryAfter
                    }
                }

            }while($retry)        
            $curBatch++
            $batchArr = @()
        }
    }

    if($batchResults.Count -ne $batchObjects.Count -and $SkipWarnings -ne $true)
    {
        Write-Log "Not all batch objects returned. Expected $($batchObjects.Count) but only got $($batchResults.Count)" 2
    }    

    $batchResults
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
    $objClone = $obj | ConvertTo-Json -Depth 50 | ConvertFrom-Json

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

    $json = ConvertTo-Json $obj -Depth 50
    if($fromFile)
    {
        # Call Update-JsonForEnvironment before importing the object
        # E.g. PolicySets contains references, AppConfiguration policies reference apps etc.
        $json = Update-JsonForEnvironment $json
    }

    if($global:Organization.Id)
    {
        $json = $json -replace "%OrganizationId%",$global:Organization.Id
    }    

    $newObj = (Invoke-GraphRequest -Url $strAPI -Content $json -HttpMethod $method @params)

    if($newObj -is [Boolean] -and $newObj -and $method -eq "PATCH")
    {
        $newObj = (Get-GraphObject -obj $obj -objectType $objectType -SkipAssignments).Object
    }
    elseif($newObj -and $method -eq "POST")
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
    $script:copyForm = Initialize-Window ($global:AppRootFolder + "\Xaml\CopyDialog.xaml")
    if(-not $script:copyForm) { return }    

    $newName = "$((Get-GraphObjectName $dgObjects.SelectedItem $global:curObjectType)) - Copy"
    if($global:curObjectType.CopyDefaultName)
    {
        $newName = $global:curObjectType.CopyDefaultName
        $dgObjects.SelectedItem.PSObject.Properties | foreach { $newName =  $newName -replace "%$($_.Name)%", $dgObjects.SelectedItem."$($_.Name)" }
    }

    Set-XamlProperty $script:copyForm "txtObjectName" "Text" $newName
    $descriptionProperty = $global:dgObjects.SelectedItem.Object | gm | Where { $_.Name -eq "Description" }
    if($descriptionProperty)
    {
        Set-XamlProperty $script:copyForm "txtObjectDescription" "Text" $global:dgObjects.SelectedItem.Object.Description
    }
    else
    {
        Set-XamlProperty $script:copyForm "txtObjectDescription" "IsEnabled" $false
    }

    $script:copyForm.Add_ContentRendered({
        $txtName = $script:copyForm.FindName("txtObjectName")
        if($txtName)
        {
            $txtName.SelectAll();
        }
    })

    Add-XamlEvent $script:copyForm "btnOk" "Add_Click" -scriptBlock ([scriptblock]{
        $script:copyForm.DialogResult = $true;	
    })

    $script:copyForm.Owner = $global:window
    $script:copyForm.Icon = $global:Window.Icon     
    $ret = $script:copyForm.ShowDialog()

    if($ret)
    {
        $newName = Get-XamlProperty $script:copyForm "txtObjectName" "Text"
        if(-not $newName)
        {
            Write-Log "New name cannot be empty. Copy object skipped" 2
            Write-Status ""
            return 
        }

        # Export profile
        Write-Status "Export $((Get-GraphObjectName $dgObjects.SelectedItem $global:curObjectType))"

        $exportObj = (Get-GraphObject $dgObjects.SelectedItem.Object $global:curObjectType -SkipAssignments).Object

        if($global:curObjectType.PreCopyCommand)
        {
            if((& $global:curObjectType.PreCopyCommand $exportObj $global:curObjectType $newName))
            {
                if((Get-SettingValue "RefreshObjectsAfterCopy") -eq $true)
                {
                    Show-GraphObjects
                }
                Write-Status ""
                return
            }
        }

        # Convert to Json and back to clone the object
        $obj = ConvertTo-Json $exportObj -Depth 50 | ConvertFrom-Json
        if($obj)
        {
            # Import new profile
            Set-GraphObjectName $obj $global:curObjectType $newName
            if((Get-XamlProperty  $script:copyForm "txtObjectDescription" "IsEnabled" $false) -eq $true)
            {
                $obj.Description = Get-XamlProperty  $script:copyForm "txtObjectDescription" "Text"
            }

            $newObj = Import-GraphObject $obj $global:curObjectType
            if($newObj)
            {
                Set-GraphNavigationProperties $newObj $exportObj $global:curObjectType -FromOldObject

                if($global:curObjectType.PostCopyCommand)
                {
                    & $global:curObjectType.PostCopyCommand $exportObj $newObj $global:curObjectType
                }

                if((Get-SettingValue "RefreshObjectsAfterCopy") -eq $true)
                {
                    Show-GraphObjects
                }
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

    Add-ObjectColumnInfoClass

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

    Set-XamlProperty  $script:detailsForm "txtValue" "Text" (ConvertTo-Json $global:dgObjects.SelectedItem.Object -Depth 50)
    
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
            Set-XamlProperty  $script:detailsForm "txtValue" "Text" (ConvertTo-Json $obj.Object -Depth 50)
            Set-XamlProperty  $script:detailsForm "btnFull" "IsEnabled" $false
        }
        Write-Status ""
    })

    #Settings tab

    Set-XamlProperty  $script:detailsForm "txtObjectName" "Text" (Get-GraphObjectName  $global:dgObjects.SelectedItem.Object $global:dgObjects.SelectedItem.ObjectType)

    $descriptionProperty = $global:dgObjects.SelectedItem.Object | gm | Where { $_.Name -eq "Description" }
    if($descriptionProperty)
    {
        Set-XamlProperty  $script:detailsForm "txtObjectDescription" "Text" $global:dgObjects.SelectedItem.Object.Description       
    }
    else
    {
        Set-XamlProperty  $script:detailsForm "txtObjectDescription" "IsEnabled" $false
    }

    Add-XamlEvent $script:detailsForm "btnObjectSettingsSave" "Add_Click" -scriptBlock ([scriptblock]{
        
        $curObjectName = (Get-GraphObjectName  $global:dgObjects.SelectedItem.Object $global:dgObjects.SelectedItem.ObjectType)
        if(([System.Windows.MessageBox]::Show("Are you sure you want to upload object settings?`n`nCurrent name:`n$($curObjectName)", "Update object settings?", "YesNo", "Warning")) -ne "Yes")
        {
            return
        }

        Write-Status "Update object settings for $curObjectName"
        $nameValue = (Get-XamlProperty  $script:detailsForm "txtObjectName" "Text")
        if(-not $nameValue)
        {
            [System.Windows.MessageBox]::Show("Name property must not be empty!", "Error", "OK", "Error") 
            return
        }
        # Save settings here...
        $nameProp = (?? $global:dgObjects.SelectedItem.Object.NameProperty "displayName")
        $idProp = (?? $global:dgObjects.SelectedItem.Object.IDProperty "id")

        $updateObj = [PSCustomObject]@{
            $idProp = $global:dgObjects.SelectedItem.Object."$idProp"
            $nameProp = $nameValue
        }

        if(($global:dgObjects.SelectedItem.Object."@odata.type"))
        {
            $updateObj | Add-Member -NotePropertyName "@odata.type" -NotePropertyValue ($global:dgObjects.SelectedItem.Object."@odata.type")
        }

        if((Get-XamlProperty  $script:detailsForm "txtObjectDescription" "IsEnabled") -eq $true)
        {
            $updateObj | Add-Member -NotePropertyName "description" -NotePropertyValue (Get-XamlProperty  $script:detailsForm "txtObjectDescription" "Text")
        }
        
        $api = "$($global:dgObjects.SelectedItem.ObjectType.API)/$($global:dgObjects.SelectedItem.Object."$idProp")"

        $json = $updateObj | ConvertTo-Json -Depth 20 

        $ret = Invoke-GraphRequest $api -HttpMethod "PATCH" -Content $json
        Write-Status ""
        if($ret -ne $true)
        {
            [System.Windows.MessageBox]::Show("Object settings conld not be verified!`n`nCheck the log file", "Update arning", "OK", "Warning")
        }
    })

    #Columns tab
    $script:colObjectProperties = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    Set-XamlProperty  $script:detailsForm "lstObjectColumns" "ItemsSource" $script:colObjectProperties

    $objProps = @()
    $nameProp = (?? $global:dgObjects.SelectedItem.Object.DisplayProperty "displayName")
    foreach($prop in ($global:dgObjects.SelectedItem.Object | gm | Where MemberType -eq "NoteProperty"))
    {
        if($prop.Name.Contains('@') -or $prop.Name.Contains('#')) { continue }
        $objProps += ([PSCustomObject]@{ Name=$prop.Name;Value=$prop })        
    }
    $objProps = $objProps | sort -Property Name
    Set-XamlProperty  $script:detailsForm "lstObjectProperties" "ItemsSource" $objProps

    Add-XamlEvent $script:detailsForm "btnObjectColumnsReset" "Add_Click" -scriptBlock ([scriptblock]{
        
        $script:colObjectProperties.Clear()
            
        Show-ObjectDefaultColumnsSettings
    })

    Add-XamlEvent $script:detailsForm "lstObjectColumns" "Add_SelectionChanged" -scriptBlock ([scriptblock]{
        
        Set-XamlProperty  $script:detailsForm "grdObjectColumns" "DataContext" (Get-XamlProperty  $script:detailsForm "lstObjectColumns" "SelectedItem")    
    })

    Add-XamlEvent $script:detailsForm "btnObjectColumnsAdd" "Add_Click" -scriptBlock ([scriptblock]{
        
        $selectedItem = Get-XamlProperty  $script:detailsForm "lstObjectProperties" "SelectedItem"
        if($selectedItem)
        {
            $script:colObjectProperties.Add(([ObjectColumnInfo]::new($selectedItem.Name, "")))
        }        
    })

    Add-XamlEvent $script:detailsForm "btnObjectColumnsMoveUp" "Add_Click" -scriptBlock ([scriptblock]{
        
        $selectedIndex = Get-XamlProperty  $script:detailsForm "lstObjectColumns" "SelectedIndex"
        if($selectedIndex -gt 0)
        {
            $tmpObj = $script:colObjectProperties[$selectedIndex]
            $script:colObjectProperties.RemoveAt($selectedIndex)
            $tmpObj = $script:colObjectProperties.Insert(($selectedIndex-1),$tmpObj)
            Set-XamlProperty  $script:detailsForm "lstObjectColumns" "SelectedIndex" ($selectedIndex-1)
        }
    })

    Add-XamlEvent $script:detailsForm "btnObjectColumnsMoveDown" "Add_Click" -scriptBlock ([scriptblock]{
        
        $selectedIndex = Get-XamlProperty  $script:detailsForm "lstObjectColumns" "SelectedIndex"
        if($selectedIndex -ge 0 -and $selectedIndex -lt ($script:colObjectProperties.Count-1))
        {
            $tmpObj = $script:colObjectProperties[$selectedIndex]
            $script:colObjectProperties.RemoveAt($selectedIndex)
            $tmpObj = $script:colObjectProperties.Insert(($selectedIndex+1),$tmpObj)
            Set-XamlProperty  $script:detailsForm "lstObjectColumns" "SelectedIndex" ($selectedIndex+1)
        }
    })

    Add-XamlEvent $script:detailsForm "btnObjectColumnsDelete" "Add_Click" -scriptBlock ([scriptblock]{
        
        $selectedIndex = Get-XamlProperty $script:detailsForm "lstObjectColumns" "SelectedIndex"
        if($selectedIndex -ge 0)
        {
            if(([System.Windows.MessageBox]::Show("Are you sure you want to remove selected column?", "Remove Columns?", "YesNo", "Warning")) -ne "Yes")
            {
                return
            }
            $script:colObjectProperties.Remove($tmpObj)
        }
    })    

    Add-XamlEvent $script:detailsForm "btnObjectColumnsClear" "Add_Click" -scriptBlock ([scriptblock]{
        
            if(([System.Windows.MessageBox]::Show("Are you sure you want to clear custom column settings?", "Clear Custom Columns?", "YesNo", "Warning")) -ne "Yes")
            {
                return
            }
            $script:colObjectProperties.Clear()        
    }) 
    
    Add-XamlEvent $script:detailsForm "btnObjectColumnsSave" "Add_Click" -scriptBlock ([scriptblock]{
        
        if(([System.Windows.MessageBox]::Show("Are you sure you want to save custom column settings?", "Save Custom Columns?", "YesNo", "Warning")) -ne "Yes")
        {
            return
        }

        if($script:colObjectProperties.Count -gt 0)
        {
            $arrCols = @()
            if((Get-XamlProperty $script:detailsForm "chkObjectColumnOverride" "IsChecked") -eq $true)
            {
                $arrCols += "0"
            }
            
            foreach($colProp in $script:colObjectProperties)
            {
                $tmp = $colProp.Property
                if($colProp.Header -and $colProp.Header -cne $colProp.Property)
                {
                    $tmp = "$($tmp)=$($colProp.Header)"
                }

                $arrCols +=  $tmp
            }
            $strCols = $arrCols -join ","

            Save-Setting "EndpointManager\ObjectColumns" "$($global:curObjectType.Id)" $strCols
        }
        else
        {
            $strCols = $null
            Remove-Setting "EndpointManager\ObjectColumns" "$($global:curObjectType.Id)"
        }

        Show-ObjectDefaultColumnsSettings
    })     
    
    Show-ObjectDefaultColumnsSettings

    # Show dialog
    Show-ModalForm $FormTitle $detailsForm
}

function local:Add-ObjectColumnInfoClass
{ 
    if (("ObjectColumnInfo" -as [type]))
    {
        return
    }

    $classDef = @"
    using System.ComponentModel;

    public class ObjectColumnInfo : INotifyPropertyChanged
    {
        public string Property { get { return _property; } set { _property = value;  NotifyPropertyChanged("Property");  } }
        private string _property = null;

        public string Header { get { return _header; } set { _header = value;  NotifyPropertyChanged("Header");  } }
        private string _header = null;

        public ObjectColumnInfo(string Property, string Header)
        {
            _property = Property;
            _header = Header;
        }

        public event PropertyChangedEventHandler PropertyChanged;  

        // This method is called by the Set accessor of each property.  
        // The CallerMemberName attribute that is applied to the optional propertyName  
        // parameter causes the property name of the caller to be substituted as an argument.  
        private void NotifyPropertyChanged(string propertyName = "")  
        {  
            if(PropertyChanged != null) { PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName)); }
        }        
    }

"@
    [Reflection.Assembly]::LoadWithPartialName("System.ComponentModel") | Out-Null
    Add-Type -TypeDefinition $classDef -IgnoreWarnings -ReferencedAssemblies @('System.ComponentModel')
}

function Local:Show-ObjectDefaultColumnsSettings
{
    $strColSettings = ?? (Get-Setting "EndpointManager\ObjectColumns" "$($global:curObjectType.Id)") $global:curObjectType.DefaultColumns
    $script:colObjectProperties.Clear()
    $defaultColumns = (?? $global:curObjectType.ViewProperties (@("displayName","description","id")))
    if($strColSettings)
    {                
        $arrColSettings += $strColSettings.Split(@(',',';'))
    
        Set-XamlProperty  $script:detailsForm "chkObjectColumnOverride" "IsChecked" ($arrColSettings.Count -gt 0 -and $arrColSettings[0] -eq "0")
        Set-XamlProperty $script:detailsForm "lblObjectColumnsConfig" "Content" $strColSettings

        $start = 0
        if($arrColSettings.Count -gt 0 -and ($arrColSettings[0] -eq "0" -or $arrColSettings[0] -eq "1"))
        {
            $start++
        }

        $colArr = @()

        for($i = $start;$i -lt $arrColSettings.Count;$i++)
        {
            $colProp,$colHeader= $arrColSettings[$i].Split("=")
            if(-not $colHeader)
            {
                $colHeader = $colProp
            }
            $script:colObjectProperties.Add([ObjectColumnInfo]::new($colProp,$colHeader))
            $colArr += $colProp
        }

        if(($arrColSettings.Count -eq 0 -or $arrColSettings[0] -ne "0"))
        {
            $tmpArr = $defaultColumns
            $tmpArr += $colArr

            $colArr = $tmpArr
        }

        Set-XamlProperty $script:detailsForm "lblObjectColumnsConfig" "Content" ("$(($colArr-join ','))")
    }
    else
    {
        Set-XamlProperty $script:detailsForm "lblObjectColumnsConfig" "Content" "$(($defaultColumns -join ',')) (Default)"    
    }
}

function Get-GraphObjectName
{
    param($obj, $objectType)

    if($objectType.GetObjectName)
    {
        return (& $objectType.GetObjectName $obj $objectType)

    }

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


function Get-GraphAllEntityTypes
{
    param($entityType, $xml, $hashTable)

    if(-not $hashTable.ContainsKey($entityType))
    {
        $hashTable.Add($entityType, $xml.SelectSingleNode("//*[name()='EntityType' and @Name='$entityType']"))
    }

    $nodes = $xml.SelectNodes("//*[@BaseType='graph.$entityType']")

    foreach($node in $nodes)
    {
        if($node.Abstract -ne "true")
        {
            $hashTable.Add($node.Name, $node)
        }
        Get-GraphAllEntityTypes $node.Name $xml $hashTable
    }    
}

function Get-GraphEntityTypeObject
{
    param($entityType, $xml, $skipProperties = @())

    $props = Get-GraphEntityTypeProperties $entityType $xml

    if(-not $props) { return }

    $obj = [PSCustomObject]@{
        
    }

    foreach($prop in $props)
    {
        if($prop.Name -in $skipProperties) { continue }
        $obj | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $null
    }
    $obj
}

function Get-GraphEntityTypeProperties
{
    param($entityType, $xml)

    Get-GraphMetaData

    if(-not $xml) { $xml = $global:metaDataXML }
    if(-not $xml) { return }

    $tmpEntity = $xml.SelectSingleNode("//*[name()='EntityType' and @Name='$entityType']")
    if(-not $tmpEntity) { return }

    $entities = @()
    $entities += $tmpEntity
    
    while($tmpEntity.BaseType)
    {        
        $baseType = $tmpEntity.BaseType.Split('.')[-1]
        $tmpEntity = $xml.SelectSingleNode("//*[name()='EntityType' and @Name='$baseType']")
        if($tmpEntity) 
        {
            $entities += $tmpEntity
        }
    }
    $properties = @()
    [array]::Reverse($entities)
    foreach($enitiy in $entities)
    {
        $properties += $enitiy.SelectNodes("*[name()='Property' or name()='NavigationProperty']")
    }
    
    $properties 
}

function Get-GraphObjectFromFile
{
    param($fileName)

    if(-not $fileName) { return } 

    if([System.IO.File]::Exists($fileName) -eq $false)
    {
        Write-LogDebug "File $fileName not found" 2
        return
    }

    $json = Get-Content -LiteralPath $fileName -Raw
    if($global:Organization.Id)
    {
        $json = $json -replace "%OrganizationId%",$global:Organization.Id
    }        

    try 
    {
        $json | ConvertFrom-Json        
    }
    catch 
    {
        Write-LogError "Failed to convert json file $fileName" $_.Exception
    }
}

function Save-GraphObjectToFile
{
    param($obj, $fileName)

    $json = $obj | ConvertTo-Json -Depth 50

    if($global:Organization.Id)
    {
        $json = $json -replace $global:Organization.Id, "%OrganizationId%"
    }
    
    try
    {
        $json | Out-File -LiteralPath $fileName -Force -ErrorAction Stop
    }
    catch 
    {
        Write-LogError "Failed to save file $fileName" $_.Exception
    }
}

function Get-GraphObjectFile
{
    param($obj, $objectType, $path)
    $fileName = (Get-GraphObjectName $obj $objectType).Trim('.')

    if((Get-SettingValue "AddIDToExportFile") -eq $true -and $obj.Id)
    {
        $fileName = ($fileName + "_" + $obj.Id)
    }
    $fileName = "$((Remove-InvalidFileNameChars $fileName)).json"
    if($path)
    {
        $fileName = "$path\$fileName"
    }

    $fileName
}

function Confirm-GraphMatchFilter
{
    param($graphObj, [string]$filter)

    if(-not $filter.Trim()) { return $true }

    $filterScope = ""

    if($filter -like "scope:*" -or $filter -like "tag:*")
    {
        $filterScope = $filter.Split(':')[1]
    }

    $objName = Get-GraphObjectName $graphObj.Object $graphObj.ObjectType
    if($filterScope)
    {
        if(($graphObj.Object.PSObject.Properties | Where Name -eq "roleScopeTagIds"))
        {
            $scopeTagProperty = "roleScopeTagIds"
        }
        elseif(($graphObj.Object.PSObject.Properties | Where Name -eq "roleScopeTags"))
        {
            $scopeTagProperty = "roleScopeTags"
        }
        else
        {
            Write-Log "$objName excluded based on Scope(Tags) not supported on $($graphObj.ObjectType.GroupId) objects"
            continue
        }

        if(-not $script:scopeTags -and $script:offlineDocumentation -ne $true)
        {
            $script:scopeTags = (Invoke-GraphRequest -Url "/deviceManagement/roleScopeTags").Value
        }
        
        $found = $false
        foreach($scopeTagId in $graphObj.Object."$scopeTagProperty")
        {                    
            $scopeTagObj = $script:scopeTags | Where Id -eq $scopeTagId
            if($scopeTagObj -and $filterScope -and $scopeTagObj.displayName -match [RegEx]::Escape($filterScope))
            {
                return $true               
            }
        }

        if($found -eq $false)
        {
            Write-Log "$objName excluded based on no Scope(Tags) found that matches the filter"
            return $false
        }
    }
    else
    {                
        if($objName -and $filter -and $objName -notmatch [RegEx]::Escape($filter))
        {
            Write-Log "$objName excluded based on the name does not match the filter"
            return $false
        }
    }
    return $true
}