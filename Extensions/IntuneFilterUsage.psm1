<#
.SYNOPSIS
Module for listing Intune assignment filter usage

.DESCRIPTION

.NOTES
  Author:         Mikael Karlsson
#>
function Get-ModuleVersion
{
    '1.1.1'
}

function Invoke-InitializeModule
{
    Add-EMToolsViewItem (New-Object PSObject -Property @{
        Title = "Intune Filter Usage"
        Id = "IntuneFilterUsage"
        ViewID = "EMTools"
        Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
        Icon="DeviceConfiguration"
        ShowViewItem = { Show-IntuneToolsFilterUsage }
    })
}

function Show-IntuneToolsFilterUsage
{
    if(-not $script:frmIntuneFilterUsage)
    {
        $script:frmIntuneFilterUsage = Get-XamlObject ($global:AppRootFolder + "\Xaml\IntuneToolsFiterUsage.xaml") #-AddVariables

        if(-not $script:frmIntuneFilterUsage) { return }
        
        Add-XamlEvent $script:frmIntuneFilterUsage "btnGetIntuneFilterUsage" "add_click" ({
            Write-Status "Get Intune Filter Usage"
            Get-EMIntuneFilterUsage
            Write-Status ""
        })

        Add-XamlEvent $script:frmIntuneFilterUsage "btnIntuneFilterUsageCopy" "add_click" ({
            $dgValues = Get-DataGridValues ($script:frmIntuneFilterUsage.FindName("dgIntuneFilterUsage")) 
            $dgValues | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
        })

        Add-XamlEvent $script:frmIntuneFilterUsage "btnIntuneFilterUsagesSave" "add_click" ({
        
            $dlgSave = New-Object -Typename System.Windows.Forms.SaveFileDialog
            $dlgSave.FileName = $obj.FileName
            $dlgSave.DefaultExt = "*.csv"
            $dlgSave.Filter = "CSV (*.csv)|*.csv|All files (*.*)| *.*"            
            if($dlgSave.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $dlgSave.Filename)
            {
                $dgValues = Get-DataGridValues ($script:frmIntuneFilterUsage.FindName("dgIntuneFilterUsage")) 
                $dgValues | ConvertTo-Csv -NoTypeInformation | Out-File -LiteralPath $dlgSave.Filename -Encoding UTF8 -Force
            }            
        })
    }
    
    $global:grdToolsMain.Children.Clear()
    $global:grdToolsMain.Children.Add($frmIntuneFilterUsage)
}

function Get-DataGridValues_old
{
    param($dataGrid)
    
    $dgColumns = $dataGrid.Columns
    #$dgColumns = Get-XamlProperty $script:frmIntuneFilterUsage "dgIntuneFilterUsage" "Columns"

    $properties = @()

    foreach($tmpCol in $dgColumns)
    {
        $propName = $tmpCol.Binding.Path.Path
        $properties += @{n=$tmpCol.Header;e=([Scriptblock]::Create("`$_.$propName"))}
    }

    ($script:objFilterUsage | Select -Property $properties)
}

function Get-EMIntuneFilterUsage
{
    param($rootDir)

    Write-Status "Gather Intune Filter Information"

    Set-XamlProperty $script:frmIntuneFilterUsage "dgIntuneFilterUsage" "ItemsSource" $null

    $objectType = Get-GraphObjectType "AssignmentFilters"

    $loadedGroups = @{}
    $loadedGroups.Add("adadadad-808e-44e2-905a-0b7873a8a531","All Devices")
    $loadedGroups.Add("acacacac-9df4-4c7d-9d50-4ef0226f57a9","All Users")

    $script:objFilters = (Invoke-GraphRequest -Url $objectType.API).Value

    $script:objFilterUsage = @()
    $groupIDs = @()

    foreach($filter in $script:objFilters)
    {   
        Write-Status "Get payloads for filter $($filter.displayName)"

        $payloadsManual = @()

        $payloads = (Invoke-GraphRequest -Url "$($objectType.API)/$($filter.ID)/payloads").value

        $batchObjs = @()
        foreach($payload in $payloads)
        {
            $guid = [Guid]::NewGuid().Guid
            
            $payloadsObj = @{
                Payload = $payload
                ID = $guid
                Requests = @()                
            }

            if($groupIDs -notcontains $payload.groupId)
            {
                $groupIDs += $payload.groupId
            }

            $batchObjs += $payloadsObj
            
            if($payload.payloadType -eq "win32app")
            {
                $payloadsObj.Requests += [ordered]@{
                    id = "$($guid)_deviceHealthScripts"
                    method = "GET"
                    url = "/deviceManagement/deviceHealthScripts/$($payload.payloadId)/?`$select=displayName,isGlobalScript"
                    headers = @{"x-ms-command-name"="AssignmentFilterPayloadProxy_resolvePayloadNames_BatchItem"}
                }
            }
            elseif($payload.payloadType -eq "application")
            {
                $payloadsObj.Requests += [ordered]@{
                    id = "$($guid)_mobileApps"
                    method = "GET"
                    url = "/deviceAppManagement/mobileApps/$($payload.payloadId)/?`$select=displayName"
                    headers = @{"x-ms-command-name"="AssignmentFilterPayloadProxy_resolvePayloadNames_BatchItem"}
                }
            }
            elseif($payload.payloadType -eq "deviceManagmentConfigurationAndCompliancePolicy")
            {
                $payloadsObj.Requests += [ordered]@{
                    id = "$($guid)_configurationPolicies"
                    method = "GET"
                    url = "/deviceManagement/configurationPolicies/$($payload.payloadId)/?`$select=name,platforms,technologies,templateReference"
                    headers = @{"x-ms-command-name"="AssignmentFilterPayloadProxy_resolvePayloadNames_BatchItem"}
                }
            }
            elseif($payload.payloadType -eq "groupPolicyConfiguration")
            {
                $payloadsObj.Requests += [ordered]@{
                    id = "$($guid)_groupPolicyConfigurations"
                    method = "GET"
                    url = "/deviceManagement/groupPolicyConfigurations/$($payload.payloadId)/?`$select=displayName"
                    headers = @{"x-ms-command-name"="AssignmentFilterPayloadProxy_resolvePayloadNames_BatchItem"}
                }
            }
            elseif($payload.payloadType -eq "enrollmentConfiguration")
            {
                if(-not $script:enrolmentConfigurations)
                {
                    $script:enrolmentConfigurations = @()
                    $script:enrolmentConfigurations += (Invoke-GraphRequest -Url "/deviceManagement/deviceEnrollmentConfigurations?`$select=displayName,id,deviceEnrollmentConfigurationType").value
                    $script:enrolmentConfigurations += (Invoke-GraphRequest -Url "/deviceManagement/deviceEnrollmentConfigurations?`$select=displayName,id,deviceEnrollmentConfigurationType&`$filter=deviceEnrollmentConfigurationType eq 'EnrollmentNotificationsConfiguration'").value
                }

                $payloadsManual += $payload

                <#
                $payloadsObj.Requests += [ordered]@{
                    id = "$($guid)_enrollmentConfiguration"
                    method = "GET"
                    url = "/deviceManagement/deviceEnrollmentConfigurations/$($enrolmentConfig.Id)/?`$select=displayName"
                    headers = @{"x-ms-command-name"="AssignmentFilterPayloadProxy_resolvePayloadNames_BatchItem"}
                }
                #>
            }                                 
            else
            {
                $payloadsObj.Requests += [ordered]@{
                    id = "$($guid)_deviceCompliancePolicies"
                    method = "GET"
                    url = "/deviceManagement/deviceCompliancePolicies/$($payload.payloadId)/?`$select=displayName"
                    headers = @{"x-ms-command-name"="AssignmentFilterPayloadProxy_resolvePayloadNames_BatchItem"}
                }
        
                $payloadsObj.Requests += [ordered]@{
                    id = "$($guid)_deviceConfigurations"
                    method = "GET"
                    url = "/deviceManagement/deviceConfigurations/$($payload.payloadId)/?`$select=displayName"
                    headers = @{"x-ms-command-name"="AssignmentFilterPayloadProxy_resolvePayloadNames_BatchItem"}
                }
                
                $payloadsObj.Requests += [ordered]@{
                    id = "$($guid)_mobileAppConfigurations"
                    method = "GET"
                    url = "/deviceAppManagement/mobileAppConfigurations/$($payload.payloadId)/?`$select=displayName"
                    headers = @{"x-ms-command-name"="AssignmentFilterPayloadProxy_resolvePayloadNames_BatchItem"}
                }
            }
        }

        if($batchObjs.Count -gt 0) 
        {
            $objName = Get-GraphObjectName $filter $objectType
            $responses = Invoke-GraphBatchRequest @($batchObjs.Requests) $objName -SkipWarnings

            foreach($response in ($responses | Where Status -lt 300))
            {
                $payload = ($batchObjs | Where { $response.id -like "$($_.ID)*"}).Payload

                if($payload.assignmentFilterType -eq "Include")
                {
                    $filterType = "Include"
                }
                else
                {
                    $filterType = "Exclude"
                }

                $typeStr = $null
                if($payload.payloadType -eq "application")
                {
                    $typeStr = Get-LanguageString "AppType.windowsClassicApp"
                }
                elseif($payload.payloadType -eq "win32app")
                {
                    $typeStr = "Proactive Remediations"
                }
                elseif($payload.payloadType -eq "groupPolicyConfiguration")
                {
                    $typeStr = "Settings Catalog"
                }
                elseif($payload.payloadType -eq "deviceManagmentConfigurationAndCompliancePolicy")
                {
                    $typeStr = "Administrative Templates"
                }                
                else
                {
                    $typeStr = (Get-PolicyTypeName $response.body.'@odata.type' $payload.payloadType)
                }
 
                if(-not $typeStr) { $typeStr = $payload.payloadType}

                $script:objFilterUsage += [PSCustomObject]@{
                    FiterObject = $filter
                    PayloadObject = $payload
                    FilterName = $filter.displayName
                    PolicyName = ?? $response.body.Name $response.body.displayName
                    Type = $response.body.'@odata.type'
                    PayloadType = $typeStr
                    Mode = $filterType
                    GroupID = $payload.groupId
                    GroupName = $payload.groupId
                }
            }

            foreach($response in ($responses | Where Status -ge 300))
            {
                $payload = ($batchObjs | Where { $response.id -like "$($_.ID)*"}).Payload
                Write-Log "Failed to get info for payload with id $($payload.payloadId) of type $($payload.payloadType). Might be deleted or not supported." 2
            }
        }

        foreach($payload in $payloadsManual) 
        {
            $payloadPolicy = $script:enrolmentConfigurations | Where Id -like "$($payload.payloadId)*" | Select -First 1
    
            if($payloadPolicy)
            {
                if($payloadPolicy.deviceEnrollmentConfigurationType -eq "enrollmentNotificationsConfiguration")
                {
                    $typeStr = "Enrollment notifications"
                }
                elseif($payloadPolicy.deviceEnrollmentConfigurationType -eq "windows10EnrollmentCompletionPageConfiguration")
                {
                    $typeStr = "Enrollment Status Page"
                }
                else
                {
                    $typeStr = (Get-PolicyTypeName $payloadPolicy.body.'@odata.type' $payload.payloadType)                    
                }

                if($payload.assignmentFilterType -eq "Include")
                {
                    $filterType = "Include"
                }
                else
                {
                    $filterType = "Exclude"
                }

                $script:objFilterUsage += [PSCustomObject]@{
                    FiterObject = $filter
                    PayloadObject = $payload
                    FilterName = $filter.displayName
                    PolicyName = ?? $payloadPolicy.Name $payloadPolicy.displayName
                    Type = $payloadPolicy.'@odata.type'
                    PayloadType = $typeStr
                    Mode = $filterType
                    GroupID = $payload.groupId
                    GroupName = $payload.groupId
                }            
            }
        }
    }

    if($groupIDs.Count -gt 0)
    {
        $guid = [Guid]::NewGuid().Guid
        $groupObjs = @()
        $x = 1
        foreach($groupID in $groupIDs)
        {
            if($loadedGroups.ContainsKey($groupID)) { continue }
            $groupObjs += [ordered]@{
                id= "$($guid)_$x"
                method="GET"
                url="/groups/$($groupID)/?`$select=displayName,id"
                headers = @{"x-ms-command-name"="AssignmentFilterPayloadProxy_resolvePayloadGroupAssignments_BatchItem"}   
            }
            $x++
        }
        
        if($groupObjs.Count -gt 0)
        {
            $responses = Invoke-GraphBatchRequest $groupObjs "Groups"
            
            $batchObj = [ordered]@{
                requests = @($groupObjs)
                }
                
            $responses = (Invoke-GraphRequest -Url "`$batch" -Body ($batchObj | ConvertTo-Json -Depth 50 -Compress) -Method "POST").responses
            
            foreach($response in ($responses | Where Status -eq 200))
            {
                if($response.body.displayName -and $response.body.id -and $loadedGroups.ContainsKey($response.body.id) -eq $false) 
                {
                    $loadedGroups.Add($response.body.id, $response.body.displayName)
                }
            }
        }

        foreach($groupID in $loadedGroups.Keys)
        {
            $filterObjs = $script:objFilterUsage | WHere GroupID -eq $groupID
            if($filterObjs -and $loadedGroups[$groupID])
            {
                foreach($filterObj in $filterObjs) {
                    $filterObj.GroupName = $loadedGroups[$groupID]
                }
            }
        }
        $script:enrolmentConfigurations = $null
    }

    Add-XamlEvent $script:frmIntuneFilterUsage "txtIntuneFilterUsageFilter" "Add_LostFocus" ({
        Invoke-IntueFilterUsageBoxChanged $this
    })
    
    Add-XamlEvent $script:frmIntuneFilterUsage "txtIntuneFilterUsageFilter" "Add_GotFocus" ({
        if($this.Tag -eq "1" -and $this.Text -eq "Filter") { $this.Text = "" }
        Invoke-IntueFilterUsageBoxChanged $this ($script:frmIntuneFilterUsage.FindName("dgIntuneFilterUsage"))
    })    
    
    Add-XamlEvent $script:frmIntuneFilterUsage "txtIntuneFilterUsageFilter" "Add_TextChanged" ({
        Invoke-IntueFilterUsageBoxChanged $this ($script:frmIntuneFilterUsage.FindName("dgIntuneFilterUsage"))
    })

    Invoke-IntueFilterUsageBoxChanged ($script:frmIntuneFilterUsage.FindName("txtIntuneFilterUsageFilter")) ($script:frmIntuneFilterUsage.FindName("dgIntuneFilterUsage"))

    $ocList = [System.Collections.ObjectModel.ObservableCollection[object]]::new(@($script:objFilterUsage))

    Set-XamlProperty $script:frmIntuneFilterUsage "dgIntuneFilterUsage" "ItemsSource" ([System.Windows.Data.CollectionViewSource]::GetDefaultView($ocList))
}

function Invoke-IntueFilterUsageBoxChanged
{
    param($txtBox, $dgObject)

    $filter = $null
    
    if($txtBox.Text.Trim() -eq "" -and $txtBox.IsFocused -eq $false)
    {
        $txtBox.FontStyle = "Italic"
        $txtBox.Tag = 1
        $txtBox.Text = "Filter"
        $txtBox.Foreground="Lightgray"        
    }
    elseif($txtBox.Tag -eq "1" -and $txtBox.Text -eq "Filter" -and $txtBox.IsFocused -eq $false)
    {
        
    }
    else
    {
        $txtBox.FontStyle = "Normal"
        $txtBox.Tag = $null
        $txtBox.Foreground="Black"
        $txtBox.Background="White"

        if($txtBox.Text)
        {
            $filter = {
                param ($item)

                return ($item.FilterName -match [regex]::Escape($txtBox.Text) -or $item.PolicyName -match [regex]::Escape($txtBox.Text) -or $item.GroupName -match [regex]::Escape($txtBox.Text) )
            }
        }         
    }

    if($dgObject.ItemsSource -is [System.Windows.Data.ListCollectionView] -and $txtBox.IsFocused -eq $true)
    {
        # This causes odd behaviour with focus e.g. and item has to be clicked twice to be selected 
        $dgObject.ItemsSource.Filter = $filter
        #$dgObject.ItemsSource.Refresh()
    }
}
