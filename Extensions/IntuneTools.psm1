<#
.SYNOPSIS
Module for Intune tools

.DESCRIPTION
This module is for the Intune Tools View.

# Full ADMX reference can be found here (from 2007):
# http://download.microsoft.com/download/5/0/8/5081217f-4a2a-470e-a7fa-5976e40b0839/Group%20Policy%20ADMX%20Syntax%20Reference%20Guide.doc

# Schema documented 2017
# https://docs.microsoft.com/en-us/previous-versions/windows/desktop/policy/admx-schema

# ADMX schema reference can be found here
# https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-gpreg/6e10478a-e9e6-4fdc-a1f6-bdd9bd7f2209

.NOTES
  Author:         Mikael Karlsson
#>

$global:EMToolsViewObject = $null

function Get-ModuleVersion
{
    '1.0.5'
}

function Invoke-InitializeModule
{
    Add-ADMXRegClasses    

    Add-EMToolsViewItem

    # https://docs.microsoft.com/en-us/windows/client-management/mdm/win32-and-centennial-app-policy-configuration
    # ADMX ingestion cannot write to these paths:
    $script:unsupportedLocations = @('System','Software\Microsoft','Software\Policies\Microsoft')
    # With excemption for these paths:
    $script:unsupportedOverride = @('Software\Policies\Microsoft\Office','Software\Microsoft\Office','Software\Microsoft\Windows\CurrentVersion\Explorer','Software\Microsoft\Internet Explorer','software\policies\microsoft\shared tools\proofing tools','software\policies\microsoft\imejp','software\policies\microsoft\ime\shared','software\policies\microsoft\shared tools\graphics filters','software\policies\microsoft\windows\currentversion\explorer','software\policies\microsoft\softwareprotectionplatform','software\policies\microsoft\officesoftwareprotectionplatform','software\policies\microsoft\windows\windows search\preferences','software\policies\microsoft\exchange','software\microsoft\shared tools\proofing tools','software\microsoft\shared tools\graphics filters','software\microsoft\windows\windows search\preferences','software\microsoft\exchange','software\policies\microsoft\vba\security','software\microsoft\onedrive','software\Microsoft\Edge','Software\Microsoft\EdgeUpdate')
 
    $script:admxTemplate = @"
<policyDefinitions revision="1.0" schemaVersion="1.0">
    <categories>
        <category name="RegImport" />
    </categories>
    <policies>
        <policy name="" class="" displayName="" explainText="" presentation="" key="" valueName="">
            <parentCategory ref="RegImport" />
            <supportedOn ref="windows:SUPPORTED_Windows7" />
            <enabledValue>
                <decimal value="1" />
            </enabledValue>
            <disabledValue>
                <decimal value="0" />
            </disabledValue>

            <elements>
            </elements>
      </policy>
    </policies>
</policyDefinitions>      
"@    
}

function Add-EMToolsViewItem
{
    param($viewItem)

    if(-not $global:EMToolsViewObject)
    {
        $viewPanel = Get-XamlObject ($global:AppRootFolder + "\Xaml\EndpointManagerTools.xaml") -AddVariables
    
        if(-not $viewPanel) { return }

        #Add menu group and items
        $global:EMToolsViewObject = (New-Object PSObject -Property @{ 
            Title = "Intune Tools"
            Description = "Additional tools for managing Intune"
            ID = "EMTools"
            AuthenticationID = "MSAL"
            ViewPanel = $viewPanel 
            ItemChanged = { Show-EMTool }
            Activating = { Invoke-EMToolsActivatingView }
            Authentication = (Get-MSALAuthenticationObject)
            Authenticate = { Invoke-EMToolsAuthenticateToMSAL }
            AppInfo = (Get-GraphAppInfo "EMAzureApp" $global:DefaultAzureApp)
            SaveSettings = { Invoke-EMSaveSettings }
            Permissions = @()
        })

        Add-ViewObject $global:EMToolsViewObject

        Add-ViewItem (New-Object PSObject -Property @{
            Title = "ADMX Import"
            Id = "ADMXImport"
            ViewID = "EMTools"
            Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
            Icon="DeviceConfiguration"
            ShowViewItem = { Show-ADMXIngestion }
        })
        
        Add-ViewItem (New-Object PSObject -Property @{
            Title = "Reg Values"
            Id = "ADMXRegValues"
            ViewID = "EMTools"
            Permissons=@("DeviceManagementConfiguration.ReadWrite.All")
            Icon="DeviceConfiguration"
            ShowViewItem = { Show-ADMXRegValues }
        })
    }

    if($viewItem)
    {
        Add-ViewItem $viewItem
    }
}

function Invoke-EMToolsActivatingView
{

}

function Invoke-EMToolsAuthenticateToMSAL
{
    $global:EMToolsViewObject.AppInfo = Get-GraphAppInfo "EMAzureApp" $global:DefaultAzureApp
    Set-MSALCurrentApp $global:EMToolsViewObject.AppInfo
    $usr = (?? $global:MSALToken.Account.UserName (Get-Setting "" "LastLoggedOnUser"))
    if($usr)
    {
        & $global:msalAuthenticator.Login -Account $usr
    }
}

function Show-EMTool
{
    if($global:lstMenuItems.SelectedItem.ShowViewItem)
    {
        & $global:lstMenuItems.SelectedItem.ShowViewItem
    }
    else
    {
        $global:grdToolsMain.Children.Clear()
    }
}

#region ADMX functions
function Show-ADMXIngestion
{
    if(-not $script:admxPanel)
    {
        $script:admxPanel = Get-XamlObject ($global:AppRootFolder + "\Xaml\EndpointManagerToolsADMX.xaml") -AddVariables
    
        if(-not $script:admxPanel) { return }

        $global:btnADMXLoadADMX.Add_Click({        

            $of = [System.Windows.Forms.OpenFileDialog]::new()
            $of.Multiselect = $false
            $of.Filter = "ADMX Files (*.admx)|*.admx"
            $of.FileName = Get-Setting "Tools" "ADMXLastADMXFile"
            if($of.ShowDialog() -eq "OK")
            {            
                $script:currentADMXFile = [IO.FileInfo]$of.FileName
                Write-Status "Loading policy settings from $($script:currentADMXFile.Name)"
                Save-Setting "Tools" "ADMXLastADMXFile" $of.FileName
                Start-AdmxLoadFile $of.FileName            
                Write-Status ""
            }
        })

        $global:btnADMXLoadADML.Add_Click({
            $of = [System.Windows.Forms.OpenFileDialog]::new()
            $of.Multiselect = $false
            $of.Filter = "ADML Files (*.adml)|*.adml"
            $of.FileName = Get-Setting "Tools" "ADMXLastADMLFile"
            if($of.ShowDialog() -eq "OK")
            {
                Write-Status "Loading ADML policy $($of.FileName)"
                Save-Setting "Tools" "ADMXLastADMLFile" $of.FileName
                Invoke-LoadADMXSettings $of.FileName
                Write-Status ""
            }
        })

        $global:btnADMXImport.Add_Click({
            Write-Status "Import policy" 
            Import-ADMXPolicy
            Write-Status ""
        })
        
        $global:btnADMXPolicyNameRandom.Add_Click({
            $guid = [Guid]::NewGuid()
            if($global:txtADMXPolicyFileName.Text)
            {
                if($global:txtADMXPolicyFileName.Text.Split('_')[-1].Length -ne $guid.Guid.Length)
                {
                    $global:txtADMXPolicyFileName.Text = ($global:txtADMXPolicyFileName.Text + "_" + $guid.Guid)
                }
            }
            else
            {
                $global:txtADMXPolicyFileName.Text = $guid.Guid
            }
        })
        
        <#
        $global:txtADMXFilterSettings.Add_LostFocus({        
            Invoke-ADMXFilterPolicies $this
        })

        $global:txtADMXFilterSettings.Add_GotFocus({
            if($this.Tag -eq "1" -and $this.Text -eq "Filter") { $this.Text = "" }
            Invoke-ADMXFilterPolicies $this
        })
        
        $global:txtADMXFilterSettings.Add_TextChanged({
            Invoke-ADMXFilterPolicies $this
        })    
        Invoke-ADMXFilterPolicies $global:txtADMXFilterSettings

        #>

        $global:tvADMXCategories.Add_SelectedItemChanged({
            if($global:tvADMXCategories.SelectedItem.AllPolicies -eq $true)
            {
                $global:dgADMXCategoryPolicies.ColumnWidth = [System.Windows.Controls.DataGridLength]::Auto
                $global:dgADMXCategoryPolicies.Columns[0].Width = [System.Windows.Controls.DataGridLength]::Auto
                $global:dgADMXCategoryPolicies.Columns[1].Width = [System.Windows.Controls.DataGridLength]::Auto
                $global:dgADMXCategoryPolicies.Columns[2].Width = [System.Windows.Controls.DataGridLength]::Auto
                $global:dgADMXCategoryPolicies.Columns[2].Visibility = "Visible"
                $script:ocADMXSettingsList = [System.Collections.ObjectModel.ObservableCollection[object]]::new(@($global:tvADMXCategories.SelectedItem.Tag)) 
            }
            else
            {
                $global:dgADMXCategoryPolicies.ColumnWidth = [System.Windows.Controls.DataGridLength]"*"
                $global:dgADMXCategoryPolicies.Columns[0].Width = [System.Windows.Controls.DataGridLength]"10*"
                $global:dgADMXCategoryPolicies.Columns[1].Width = [System.Windows.Controls.DataGridLength]::Auto
                $global:dgADMXCategoryPolicies.Columns[2].Visibility = "Collapsed"
                $script:ocADMXSettingsList = [System.Collections.ObjectModel.ObservableCollection[object]]::new(@($script:admxPolicies | Where { $_.CategoryId -eq $global:tvADMXCategories.SelectedItem.Tag.CategoryName -and $_.SettingClass -eq $global:tvADMXCategories.SelectedItem.Tag.SettingClass}))
            }
            $global:dgADMXCategoryPolicies.ItemsSource = $script:ocADMXSettingsList 
        })

        $global:dgADMXCategoryPolicies.Add_MouseDoubleClick({        
            if(-not $global:dgADMXCategoryPolicies.SelectedItem) { return }
            Show-ADMXSettingProperties $global:dgADMXCategoryPolicies.SelectedItem
        })    

        $global:mnuADMXSettingsContextMenu.Add_Opened({
            $global:mnuADMXSettingEdit.IsEnabled = $null -ne $global:dgADMXCategoryPolicies.SelectedItem
        })
        
        $global:mnuADMXSettingEdit.Add_Click({
            if(-not $global:dgADMXCategoryPolicies.SelectedItem) { return }
            Show-ADMXSettingProperties $global:dgADMXCategoryPolicies.SelectedItem
        })

        $winADMLFile = "$($env:WinDir)\PolicyDefinitions\en-US\Windows.adml"
        if([IO.File]::Exists($winADMLFile))
        {
            [xml]$script:windowsADML = Get-Content $winADMLFile
        }
        else
        {
            Write-Log "Could not find Windows.adml. Support OS text might not be displayed correctly" 2    
        }
    }

    $global:grdToolsMain.Children.Clear()
    $global:grdToolsMain.Children.Add($script:admxPanel)
}

function Show-ADMXSettingProperties 
{
    param($settingObj)

    $script:settingsForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\EndpointManagerToolsADMXSettingProperties.xaml") -AddVariables

    if(-not $script:settingsForm) { return }

    class ListValue
    {
        [string]$Key
        [string]$Value
    }  

    Set-ADMXSettingProperties $settingObj

    Add-XamlEvent $script:settingsForm "btnADMXSettingsOK" "add_click" ({
        Save-ADMXSettings
        
        Show-ModalObject
    })
    
    
    Add-XamlEvent $script:settingsForm "btnADMXPreviousSetting" "add_click" ({
        if($global:dgADMXCategoryPolicies.SelectedIndex -eq 0) { return }
        Save-ADMXSettings
        $global:dgADMXCategoryPolicies.SelectedIndex = $global:dgADMXCategoryPolicies.SelectedIndex - 1
        Set-ADMXSettingButtonsStatus
        Set-ADMXSettingProperties $global:dgADMXCategoryPolicies.SelectedItem
    })
    
    Add-XamlEvent $script:settingsForm "btnADMXNextSetting" "add_click" ({
        if($global:dgADMXCategoryPolicies.SelectedIndex -eq ($global:dgADMXCategoryPolicies.ItemsSource.Count - 1)) { return }         
        Save-ADMXSettings
        $global:dgADMXCategoryPolicies.SelectedIndex = $global:dgADMXCategoryPolicies.SelectedIndex + 1
        Set-ADMXSettingButtonsStatus
        Set-ADMXSettingProperties $global:dgADMXCategoryPolicies.SelectedItem
    })     

    Add-XamlEvent $script:settingsForm "btnADMXSettingsCancel" "add_click" ({
        $global:grdADMXElements.Children.Clear()
        $script:settingsForm = $null
        Show-ModalObject
    })
    
    $global:rbADMXSettingEnabled.Add_Checked({
        $script:curItemSettingStatus = 1
        Set-ADMXControlStatus
    })
    
    $global:rbADMXSettingDisabled.Add_Checked({
        $script:curItemSettingStatus = 0
        Set-ADMXControlStatus
    })
    
    $global:rbADMXSettingNotConfigured.Add_Checked({
        $script:curItemSettingStatus = $null
        Set-ADMXControlStatus
    })    

    $global:chkADMXManualConfig.Add_Click({
        Set-ADMXControlStatus
    })

    $global:tcADMXPolicyConfig.Add_SelectionChanged({
        param($sender, $e)

        if($e.AddedItems[0] -eq $global:tabADMXSettings)
        {            
            if($global:dgADMXSettings.SelectedItem.ManualConfig -ne 1)
            {
                $global:txtADMXSettings.Text = Get-ADMXSettingsString $script:settingsForm.DataContext
            }
        }
    })    

    Show-ModalForm $settingObj.Name $script:settingsForm -HideButtons
}

function Save-ADMXSettings
{
    if($global:chkADMXManualConfig.IsChecked)
    {
        $script:settingsForm.DataContext.PolicySettings = $global:txtADMXSettings.Text
    }
    else
    {
        $script:settingsForm.DataContext.PolicySettings = Get-ADMXSettingsString $script:settingsForm.DataContext
    }
    $script:settingsForm.DataContext.ManualConfig = (?: ($global:chkADMXManualConfig.IsChecked) 1 0)
    $script:settingsForm.DataContext.SettingStatus = $script:curItemSettingStatus
    Set-ADMXSettingStatusText $script:settingsForm.DataContext
    [System.Windows.Data.CollectionViewSource]::GetDefaultView($global:dgADMXCategoryPolicies.ItemsSource).Refresh()
    $global:grdADMXElements.Children.Clear()
}

function Set-ADMXSettingButtonsStatus
{
    $global:btnADMXPreviousSetting.IsEnabled = $global:dgADMXCategoryPolicies.SelectedIndex -gt 0
    $global:btnADMXNextSetting.IsEnabled = $global:dgADMXCategoryPolicies.SelectedIndex -lt ($global:dgADMXCategoryPolicies.ItemsSource.Count - 1)
}

function Set-ADMXSettingProperties
{
    param($settingObj)

    $global:grdADMXElements.Children.Clear()

    $script:curItemSettingStatus = $settingObj.SettingStatus
    if($settingObj.SettingStatus -eq 0)
    {
        $global:rbADMXSettingDisabled.IsChecked = $true
    }
    elseif($settingObj.SettingStatus -eq 1)
    {
        $global:rbADMXSettingEnabled.IsChecked = $true
    }
    else
    {
        $global:rbADMXSettingNotConfigured.IsChecked = $true
    }

    $global:txtADMXSettings.Text = $settingObj.PolicySettings

    $global:chkADMXManualConfig.IsChecked = $settingObj.ManualConfig

    if(-not $settingObj.PolicyDefinition)
    {
        $settingObj.PolicyDefinition = Format-XML $settingObj.Definition.OuterXml
    }

    if(-not $settingObj.supportedOn -and $settingObj.Definition.supportedOn.ref)
    {
        $settingObj.supportedOn = ($script:supportedOn | Where Id -eq $settingObj.Definition.supportedOn.ref.Split(':')[-1]).DisplayName
    }

    Set-ADMXElementsPanel $settingObj
    
    $script:settingsForm.DataContext = $settingObj

    Set-ADMXControlStatus
}

function Set-ADMXControlStatus
{
    $global:grdADMXElements.IsEnabled = ($script:curItemSettingStatus -eq 1 -and $global:chkADMXManualConfig.IsChecked -eq $false)
    $global:txtADMXSettings.IsReadOnly = ($script:curItemSettingStatus -ne 1 -or $global:chkADMXManualConfig.IsChecked -eq $false)    
}

function Start-AdmxLoadFile
{
    param($fileName)

    $script:xmlNS = $null
    $script:xmlNSPrefix = $null

    $script:admx = $null
    $script:admxPolicies = @()
    $script:supportedOn = @()
    $script:admxPoliciesHT = @{}
    $script:categoryPaths = @{}

    $global:txtADMXProfileName.Text = ""
    $global:txtADMXProfileDescription.Text = ""
    $global:txtADMXPolicyFileName.Text = ""
    $global:btnADMXLoadADML.IsEnabled = $false
    $global:txtADMXPolicyIngestName.Text = ""
    $global:txtADMXPolicyAppName.Text = $null

    $admxFI = [IO.FileInfo]$fileName   
    if($admxFI.Exists -eq $false) { return }

    $admlFile = [IO.Path]::Combine($admxFi.DirectoryName, "en-US\$($admxFi.BaseName).adml")
    if([IO.File]::Exists($admlFile) -eq $false)
    {
        $admlFile = [IO.Path]::Combine($admxFi.DirectoryName, "$($admxFi.BaseName).adml")
    }
       
    if([IO.File]::Exists($admlFile) -eq $false)
    {
        Write-Log "Could not find an ADML file" 2
        $admlFile = $null
    }

    try 
    {
        Write-Log "Load ADMX file $fileName"
        [xml]$script:admxXML = Get-Content $fileName
        
        $namespace = $script:admxXML.DocumentElement.NamespaceURI
        if($namespace)
        {
            $script:xmlNS = New-Object System.Xml.XmlNamespaceManager($script:admxXML.NameTable)
            $script:xmlNS.AddNamespace("ns", $namespace)
            $script:xmlNSPrefix = "ns:"
        }
        else
        {
            $script:xmlNS = $null
            $script:xmlNSPrefix = ""
        }

        $prefix = $script:admxXML.policyDefinitions.policyNamespaces.SelectSingleNode("$($script:xmlNSPrefix)target[@prefix]",$script:xmlNS)
        if($prefix)
        {
            if($prefix.namespace)
            {
                $policyId = $prefix.namespace.Split('.')[-1]
            }
            else
            {
                $policyId = $prefix.prefix
            }
        }
        else
        {
            $policyId = $tmpFI.BaseName -replace " ",""
            Write-Log "Failed to get policy id from XML. Using file base name" 2
        }
        $global:txtADMXPolicyAppName.Text = $policyId
        
        $global:txtADMXPolicyFileName.Text = $policyId
    }
    catch 
    {
        Write-LogError "Failed to load ADMX file" $_.Exception
    }

    $global:btnADMXLoadADML.IsEnabled = $true

    Invoke-LoadADMXSettings $admlFile
}

function Invoke-LoadADMXSettings
{
    param($admlFile)

    $script:lngADML = $null

    if($admlFile -and [IO.File]::Exists($admlFile))
    {
        try 
        {
            Write-Log "Load ADML file $admlFile"
            [xml]$script:lngADML = Get-Content $admlFile
        }
        catch 
        {
            Write-LogError "Failed to load ADML file $admlFile" $_.Exception
        }
    }

    $script:stringTable = @{}
    foreach($strNode in $script:lngADML.policyDefinitionResources.resources.stringTable.string)
    {
        $script:stringTable.Add($strNode.id, $strNode.'#text')
    }

    $script:supportedOn = @()
    foreach($polObj in $script:admxXML.policyDefinitions.supportedOn.definitions.definition)
    {
        $script:supportedOn += [PSCustomObject]@{
            Id=$polObj.Name
            DisplayName=(Get-ADMXADMLString $polObj)
        }
    }
    
    if($script:windowsADML)
    {
        foreach($winString in ($script:windowsADML.policyDefinitionResources.resources.stringTable.string | Where Id -like "SUPPORTED_*"))
        {
            $script:supportedOn += [PSCustomObject]@{
                Id=$winString.id
                DisplayName=$winString.'#text'
            }
        }
    }

    $devicePolicies = @()
    $userPolicies = @()

    foreach($polObj in $script:admxXML.policyDefinitions.policies.policy)
    {
        $displayName = $null
        $description = $null
        $category = $null
        $curSetting = $null

        if($polObj.parentCategory.ref)
        {
            $category = Get-ADMXCategoryNamePath $polObj.parentCategory.ref "/"
        }     

        $displayName = Get-ADMXADMLString $polObj  

        $description = Get-ADMXADMLString $polObj "explainText"
    
        if($polObj.Class -eq "Both")
        {
            $classArr = @("Device","User")
        }
        elseif($polObj.Class -eq "Machine")
        {
            $classArr = "Device"
        }

        $settingExists = $false
        foreach($class in $classArr)
        {
            #This will happen when loading a new ADML file
            $tmpName = ($polObj.Name + "_" + $class)
            if($script:admxPoliciesHT.ContainsKey($tmpName))
            {
                $curSetting = $script:admxPoliciesHT[$tmpName]
                
                foreach($tmpSetting in $curSetting)
                {
                    $settingExists = $true
                    $curSetting.Name = $displayName
                    $curSetting.Description = $description
                }            
            }
        }

        if($settingExists -eq $false)
        {
            if($polObj.Class -eq "Both")
            {
                $classArr = @("Device","User")
            }
            elseif($polObj.Class -eq "Machine")
            {
                $classArr = "Device"
            }
            else
            {
                $classArr = $polObj.Class
            }

            foreach($class in $classArr)
            {
                $newSetting = [PSCustomObject]@{
                    Name = $displayName
                    Description = $description
                    OMAURIName = $null
                    OMAURIDescription = $null
                    Category = $category
                    CategoryId = $polObj.parentCategory.ref
                    Id = $polObj.Name
                    Definition = $polObj
                    SettingStatus = $null
                    SettingStatusText = $null
                    PolicySettings = $null
                    PolicyDefinition = $null #Format-XML $polObj.OuterXml
                    ElementsPanel = $null
                    ManualConfig = $false
                    SettingClass = $class
                    SupportedOn = $null
                }

                $script:admxPoliciesHT.Add(($polObj.Name + "_" + $class), $newSetting)
                $script:admxPolicies += $newSetting
                if($newSetting.SettingClass -eq "User")
                {
                    $userPolicies += $newSetting
                }
                else
                {
                    $devicePolicies += $newSetting
                }
            }
        }
    }

    $script:admxPolicies | foreach-object { Set-ADMXSettingStatusText $_ }

    $script:admxPolicies = $script:admxPolicies | Sort -Property Name

    $global:tvADMXCategories.Items.Clear()

    $treeItems = @()

    $tvItem = [PSCustomObject]@{
        Name = "Computer Configuration"
        Children = @()
    }

    if($script:admxXML.policyDefinitions.policies)
    {
        $policies = $script:admxXML.policyDefinitions.policies.SelectNodes("$($script:xmlNSPrefix)policy[@class = 'Both' or @class = 'Machine']", $script:xmlNS)
        if($policies)
        {
            $categories = $policies.parentCategory | Select ref -Unique
            Add-ADMXCategories $categories $tvItem "Device"
        }
    }

    $treeItems += $tvItem

    $tvItem = [PSCustomObject]@{
        Name = "User Configuration"
        Children = @()
    }

    if($script:admxXML.policyDefinitions.policies)
    {
        $policies = $script:admxXML.policyDefinitions.policies.SelectNodes("$($script:xmlNSPrefix)policy[@class = 'Both' or @class = 'User']", $script:xmlNS)
        if($policies)
        {
            $categories = $policies.parentCategory | Select ref -Unique
            Add-ADMXCategories $categories $tvItem "User"
        }
    }

    $treeItems += $tvItem
    
    $treeItems | foreach-object { Add-ADMXCategoryTreeNode $_ $global:tvADMXCategories }

    if($devicePolicies.Count -gt 0)
    {
        $tvItem = [System.Windows.Controls.TreeViewItem]::new()
        $tvItem.Header = "All Policies"
        $tvItem.Tag = $devicePolicies
        $tvItem | Add-Member -MemberType NoteProperty -Name "AllPolicies" -Value $true
        $global:tvADMXCategories.Items[0].Items.Add($tvItem) | Out-Null  
    }

    if($userPolicies.Count -gt 0)
    {
        $tvItem = [System.Windows.Controls.TreeViewItem]::new()
        $tvItem.Header = "All Policies"
        $tvItem.Tag = $userPolicies
        $tvItem | Add-Member -MemberType NoteProperty -Name "AllPolicies" -Value $true
    }

    $global:tvADMXCategories.Items[1].Items.Add($tvItem) | Out-Null        
}

function Set-ADMXSettingStatusText
{
    param($settingObj)

    if($settingObj.SettingStatus -eq 0)
    {
        $settingObj.SettingStatusText = "Disabled"
    }
    elseif($settingObj.SettingStatus -eq 1)
    {
        $settingObj.SettingStatusText = "Enabled"
    }
    else
    {
        $settingObj.SettingStatusText = "Not Configured"
    }    
}
function Add-ADMXCategories
{
    param($categories, $parent, $settingClass)

    foreach($cat in $categories.ref)
    {
        $catPath = Get-ADMXCategoryIdPath $cat

        $tvObj = $parent
        foreach($catName in $catPath.Split('/'))
        {
            $curParent = $tvObj.Children | Where { $_.CategoryNode.name -eq $catName }
            if(-not $curParent)
            {
                $cat = $script:admxXML.policyDefinitions.categories.selectSingleNode("$($script:xmlNSPrefix)category[@name='$($catName)']",$script:xmlNS)
                $curParent += [PSCustomObject]@{
                    Name = Get-ADMXADMLString $cat 
                    CategoryName = $catName
                    CategoryNode = $cat
                    SettingClass = $settingClass             
                    Children = @()
                }

                $tvObj.Children += $curParent 
            }
            $tvObj = $curParent
        }
    }
}

Function Add-ADMXCategoryTreeNode
{
    Param($obj, $parent)

    $tvItem = [System.Windows.Controls.TreeViewItem]::new()
    $tvItem.Header = $obj.Name
    $tvItem.Tag = $obj
    $parent.Items.Add($tvItem) | Out-Null

    $obj.Children | Sort -Property Name | foreach-object { Add-ADMXCategoryTreeNode $_ $tvItem }
}

function Get-ADMXADMLString
{
    param($xmlNode, $property = "displayName")

    $propValue = $xmlNode.$property

    if(-not $script:lngADML -or -not $xmlNode.$property) { return $propValue}

    $tmpNode = $null

    if($xmlNode.$property.StartsWith("`$(") )
    {
        $tmp = $xmlNode.$property.SubString(2, $xmlNode.$property.Length - 3)
        $type,$strId = $tmp.Split('.')        
    }
    else
    {
        $strId =  $propValue
    }
    
    if($script:stringTable.ContainsKey($strId))
    {
        # Way quicker to use a hash table over querying all items
        $propValue = $script:stringTable[$strId]
    }
    else
    {
        $propValue = $tmpNode."#text"
    }

    $propValue
}

function Get-ADMXADMLPresentationString
{
    param($presentationInfo, $xmlNode)

    $tmp = $null

    if(-not $script:lngADML) { return $tmp }

    $presentationNode = $presentationInfo.selectSingleNode("./*[@refId='$($xmlNode.id)']")
    if($presentationNode)
    {
        $tmp = ?? $presentationNode.Label.'#text' $presentationNode.'#text'     
    }
    else
    {
        $tmp = $xmlNode.id
    }
    $tmp
}

function Get-ADMXCategoryIdPath
{
    param($categoryId, $delimiter = "/")

    $catObj = $script:admxXML.policyDefinitions.categories.selectSingleNode("$($script:xmlNSPrefix)category[@name='$($categoryId)']",$script:xmlNS)
    
    $categories = @()
    while($catObj)
    {
        $categories += $catObj.name
        if($catObj.parentCategory.ref)
        {
            $catObj = $script:admxXML.policyDefinitions.categories.selectSingleNode("$($script:xmlNSPrefix)category[@name='$($catObj.parentCategory.ref)']",$script:xmlNS)
        }
        else 
        {
            break
        }
    }
    
    [array]::Reverse($categories)

    $categories -join $delimiter
}

function Get-ADMXCategoryNamePath
{
    param($categoryId, $delimiter = "/")

    if($script:categoryPaths.ContainsKey($categoryId))
    {
        return $script:categoryPaths[$categoryId]
    }

    $catObj = $script:admxXML.policyDefinitions.categories.selectSingleNode("$($script:xmlNSPrefix)category[@name='$($categoryId)']",$script:xmlNS)
    
    $categories = @()
    while($catObj)
    {
        $categories += Get-ADMXADMLString $catObj
        if($catObj.parentCategory.ref)
        {
            $catObj = $script:admxXML.policyDefinitions.categories.selectSingleNode("$($script:xmlNSPrefix)category[@name='$($catObj.parentCategory.ref)']", $script:xmlNS)
        }
        else 
        {
            break
        }
    }
    
    [array]::Reverse($categories)

    $catPath = $categories -join $delimiter

    $script:categoryPaths.Add($categoryId, $catPath)

    $catPath
}

function Invoke-ADMXFilterPolicies 
{ 
    param($txtBox)

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
                    return ($item.Name -match [regex]::Escape($txtBox.Text))
            }
        }         
    }

    if($global:dgADMXSettings.ItemsSource -is [System.Windows.Data.ListCollectionView]  -and $txtBox.IsFocused -eq $true)
    {
        # This causes odd behaviour with focus e.g. and item has to be clicked twice to be selected 
        $global:dgADMXSettings.ItemsSource.Filter = $filter
        #$global:dgADMXSettings.ItemsSource.Refresh()
    }
}

function Get-ADMXPresentationNode
{
    param($item)

    $presentation = $null

    if($item.Definition.presentation -and $script:lngADML)
    {
        if($item.Definition.presentation.StartsWith("`$(") )
        {
            $tmp = $item.Definition.presentation.SubString(2, $item.Definition.presentation.Length - 3)
            $type,$strId = $tmp.Split('.')
            $presentation = $script:lngADML.policyDefinitionResources.resources.presentationTable.presentation | Where Id -eq $strId #selectSingleNode("presentation[@id='$($strId)']")
        }
        else
        {
            $presentation = $script:lngADML.policyDefinitionResources.resources.presentationTable | Where Id -eq $item.Definition.presentation #.selectSingleNode("presentation[@id='$($item.Definition.presentation)']")
        }
    }
    $presentation
}

function Set-ADMXElementsPanel
{
    param($item)

    if(-not $item.Definition.elements) { return }

    if(-not $item.ElementsPanel)
    {
        $grd = [System.Windows.Controls.Grid]::new()
        $presentation = Get-ADMXPresentationNode $item
        $i = 0

        if($presentation)
        {
            # Policy node schema
            # https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-gpreg/81a89003-5121-4216-b788-fde8daa71c78
            foreach($presentationNode in $presentation.ChildNodes)
            {
                $ctrl = $null

                $elementNode = $item.Definition.elements.ChildNodes | Where id -eq $presentationNode.refId #selectSingleNode("/*[@id='$($presentationNode.refId)']")
                if($presentationNode.Label.'#text')
                { $stringLabel = $presentationNode.Label.'#text' }
                elseif($presentationNode.Label)
                { $stringLabel = $presentationNode.Label }
                else
                { $stringLabel = $presentationNode.'#text' }               

                if($stringLabel -and $presentationNode.LocalName -ne "CheckBox") #$presentationNode.LocalName -eq "text")
                {
                    $ctrl = [System.Windows.Controls.TextBlock]::new()
                    $ctrl.Text = $stringLabel #$presentationNode."#text"                   
                    Add-GridObject $grd $ctrl
                }
                
                if($presentationNode.LocalName -eq "text")
                {
                    continue
                }

                if($ctrl)
                {
                    $ctrl.Margin = "0,5,0,0"
                }
                
                if($presentationNode.LocalName -eq "textbox")
                {
                    $ctrl = [System.Windows.Controls.TextBox]::new()
                    if($presentationNode.defaultValue -ne $null)
                    {
                        $ctrl.Text = $presentationNode.defaultValue
                    }                    
                }
                elseif($presentationNode.LocalName -eq "DecimalTextBox" -or
                        $presentationNode.LocalName -eq "LongDecimalTextBox" ) 
                {
                    $ctrl = Get-NumericUpDownControl $presentationNode.refId (?? $elementNode.minValue 0) (?? $elementNode.maxValue 9999) (?? $elementNode.SpinStep 1)
                    if(-not $ctrl) { continue }                    
                    if($presentationNode.defaultValue -ne $null)
                    {
                        $ctrl.Children[0].Text = $presentationNode.defaultValue
                    }
                }
                elseif($presentationNode.LocalName -eq "multiText")
                {
                    $ctrl = [System.Windows.Controls.TextBox]::new()
                    $ctrl.Height = 100
                    $ctrl.AcceptsReturn = $true
                }
                elseif($presentationNode.LocalName -eq "CheckBox")
                {
                    $ctrl = [System.Windows.Controls.CheckBox]::new()
                    $ctrl.Content = $stringLabel
                    if($presentationNode.defaultChecked -eq $true)
                    {
                        $ctrl.IsChecked = $true
                    }                    
                }
                elseif($presentationNode.LocalName -eq "ComboBox" -or
                        $presentationNode.LocalName -eq "DropdownList" )
                {
                    $ctrl = [System.Windows.Controls.ComboBox]::new()
                    $ctrl.DisplayMemberPath = "Name"
                    $ctrl.SelectedValuePath = "Value"

                    $valItems = @()
                    foreach($valItem in $elementNode.ChildNodes)
                    {
                        $displayName = Get-ADMXADMLString $valItem                        

                        if($valItem.value.decimal.value)
                        {
                            $value = $valItem.value.decimal.value
                        }
                        elseif($valItem.value.longDecimal)
                        {
                            $value = $valItem.value.longDecimal.'#text'
                        }                        
                        elseif($valItem.value.string)
                        {
                            $value = $valItem.value.string.'#text'
                        }
                        else
                        {
                            Write-Log "Unsupported value type for $($elementNode.Id): $($valItem.value.InnerXml)" 2
                            $value = "<SET MANUALLY!!!>"
                        }
                        $valItems += [PSCustomObject]@{
                            Name = $displayName 
                            Value = $value
                            } 
                    }

                    if($presentationNode.defaultItem -ne $null)
                    {
                        try
                        {
                            $ctrl.SelectedIndex = $presentationNode.defaultItem
                        }
                        catch {}
                    }

                    if($presentationNode.NoSort -ne "true")
                    {
                        $valItems = $valItems | Sort -Property Name
                    }

                    $ctrl.ItemsSource = $valItems
                }
                elseif($presentationNode.LocalName -eq "listBox")
                {
                    $ctrl = [System.Windows.Controls.DataGrid]::new()
                    $ctrl.CanUserAddRows = $true
                    $ctrl.CanUserDeleteRows = $true
                    $ctrl.CanUserSortColumns = $false
                    $ctrl.CanUserResizeRows = $false
                    $ctrl.AutoGenerateColumns = $false
                    $ctrl.ColumnWidth = [System.Windows.Controls.DataGridLength]"*"
                    
                    $column = [System.Windows.Controls.DataGridTextColumn]::new()
                    $column.Header = "Value Name"
                    $column.Width = [System.Windows.Controls.DataGridLength]"1*"
                    $binding = [System.Windows.Data.Binding]::new("Key")
                    $column.Binding = $binding
                    if($elementNode.explicitValue -ne "true")
                    {
                        $column.Visibility = "Collapsed"
                    }
                    $ctrl.Columns.Add($column)

                    $column = [System.Windows.Controls.DataGridTextColumn]::new()
                    $column.Header = "Value"
                    $column.Width = [System.Windows.Controls.DataGridLength]"1*"
                    $binding = [System.Windows.Data.Binding]::new("Value")
                    $column.Binding = $binding
                    $ctrl.Columns.Add($column)

                    $ctrl.ItemsSource = [System.Collections.Generic.List[ListValue]]::new()
                }
                else
                {
                    Write-Log "Unsupported object type in presentation: $($presentationNode.LocalName). Control: $presentationNode.refId" 2
                    continue
                }

                Add-GridObject $grd $ctrl

                if($presentationNode.refId)
                {
                    $ctrl.Tag = $elementNode
                    $ctrl.Name = $presentationNode.refId
                }
            }            
        }
        elseif($item.Definition.elements.ChildNodes)
        {            
            # This should NOT be used. Presentation settings should be defined in the ADML file
            Write-Log "No presentation settings found for $($item.Definition.Name)" 2
            foreach($elementNode in $item.Definition.elements.ChildNodes)
            {
                try
                {
                    $ctrl = $null
                    if($elementNode.LocalName -eq "text")
                    {
                        $ctrl = [System.Windows.Controls.TextBox]::new()
                    }
                    elseif($elementNode.LocalName -eq "multiText")
                    {
                        $ctrl = [System.Windows.Controls.TextBox]::new()
                        $ctrl.Height = 100
                        $ctrl.AcceptsReturn = $true
                    }         
                    elseif($elementNode.LocalName -eq "enum")
                    {
                        $ctrl = [System.Windows.Controls.ComboBox]::new()
                        $ctrl.DisplayMemberPath = "Name"
                        $ctrl.SelectedValuePath = "Value"

                        $valItems = @()
                        foreach($valItem in $elementNode.ChildNodes)
                        {
                            $displayName = Get-ADMXADMLString $valItem                        

                            if($valItem.value.decimal.value)
                            {
                                $value = $valItem.value.decimal.value
                            }
                            elseif($valItem.value.string)
                            {
                                $value = $valItem.value.string.'#text'
                            }
                            else
                            {
                                Write-Log "Unsupported value type for $($elementNode.Id): $($valItem.value.InnerXml)" 2
                                $value = "<SET MANUALLY!!!>"
                            }
                            $valItems += [PSCustomObject]@{
                                Name = $displayName 
                                Value = $value
                                } 
                        }
                        $ctrl.ItemsSource = $valItems
                    }
                    elseif($elementNode.LocalName -eq "list")
                    {
                        $ctrl = [System.Windows.Controls.DataGrid]::new()
                        $ctrl.CanUserAddRows = $true
                        $ctrl.CanUserDeleteRows = $true
                        $ctrl.CanUserSortColumns = $false
                        $ctrl.CanUserResizeRows = $false
                        $ctrl.ColumnWidth = [System.Windows.Controls.DataGridLength]::Auto
                        $ctrl.AutoGenerateColumns = $false
                        $ctrl.ColumnWidth = [System.Windows.Controls.DataGridLength]"*"
                        
                        $column = [System.Windows.Controls.DataGridTextColumn]::new()
                        $column.Header = "Value Name"
                        $column.Width = [System.Windows.Controls.DataGridLength]"1*"
                        $binding = [System.Windows.Data.Binding]::new("Key")
                        $column.Binding = $binding
                        if($elementNode.explicitValue -ne "true")
                        {
                            $column.Visibility = "Collapsed"
                        }
                        $ctrl.Columns.Add($column)

                        $column = [System.Windows.Controls.DataGridTextColumn]::new()
                        $column.Header = "Value"
                        $column.Width = [System.Windows.Controls.DataGridLength]"1*"
                        $binding = [System.Windows.Data.Binding]::new("Value")
                        $column.Binding = $binding
                        $ctrl.Columns.Add($column)

                        $ctrl.ItemsSource = [System.Collections.Generic.List[ListValue]]::new()
                    }
                    elseif($elementNode.LocalName -eq "decimal")
                    {
                        $ctrl = [System.Windows.Controls.TextBox]::new()
                    }
                    elseif($elementNode.LocalName -eq "boolean")
                    {
                        $ctrl = [System.Windows.Controls.ComboBox]::new()
                    }
                    else
                    {
                        Write-Log "Element type not supported: $($elementNode.LocalName)" 2
                        continue
                    }

                    $displayName = Get-ADMXADMLPresentationString $presentation $elementNode
                                        
                    if($displayName)
                    {
                        $rd = [System.Windows.Controls.RowDefinition]::new()
                        $rd.Height = [double]::NaN #[System.Windows.GridLength]::Auto         
                        $grd.RowDefinitions.Add($rd)

                        $tb = [System.Windows.Controls.TextBlock]::new()
                        if($i -gt 0) { $tb.Margin = "0,5,0,0" }
                        $tb.Text = $displayName
                        $tb.SetValue([System.Windows.Controls.Grid]::RowProperty,$i)
                        $grd.Children.Add($tb)

                        $i++
                    }

                    $rd = [System.Windows.Controls.RowDefinition]::new()
                    $rd.Height = [double]::NaN 
                    $grd.RowDefinitions.Add($rd)

                    $ctrl.SetValue([System.Windows.Controls.Grid]::RowProperty,$i)
                    $ctrl.Tag = $elementNode
                    $ctrl.Name = $elementNode.Id
                    $grd.Children.Add($ctrl)

                    #$grd.RegisterName($ctrl.Name, $ctrl)

                    $i++
                }
                catch
                {
                    Write-LogError "Failed to add ADMX element $($elementNode.LocalName) with id $($elementNode.id)" $_.Exception
                }
            }
        }
        $rd = [System.Windows.Controls.RowDefinition]::new()
        #$rd.Height = [System.Windows.GridLength]::new(1, "Star")
        $grd.RowDefinitions.Add($rd)        
        $item.ElementsPanel = $grd
    }

    if($item.ElementsPanel)
    {
        $global:grdADMXElements.Children.Add($item.ElementsPanel)
        foreach($elementNode in $item.Definition.elements.ChildNodes)
        {            
            $ctrl = [System.Windows.LogicalTreeHelper]::FindLogicalNode($item.ElementsPanel, $elementNode.Id)
            if(-not $ctrl)
            {
                Write-Log "Could not find a control with id $($elementNode.Id)" 3
                continue
            }
            #$global:grdADMXElements.RegisterName($ctrl.Name, $ctrl)
        }
    }
}

function Get-ADMXSettingsString
{
    param($item)

    if(-not $item -or -not $item.Definition.elements) { return }

    if($script:curItemSettingStatus -ne 1)
    {
        $item.PolicySettings = $null
        return
    }

    $policySettings = @()
    foreach($elementNode in $item.Definition.elements.ChildNodes)
    {
        #$ctrl = $item.ElementsPanel.FindName($elementNode.Id)
        $ctrl = [System.Windows.LogicalTreeHelper]::FindLogicalNode($item.ElementsPanel, $elementNode.Id)
        if(-not $ctrl)
        {
            Write-Log "Could not find a control with id $($elementNode.Id)" 3
            continue
        }
        $ctrlValue = $null

        if($elementNode.LocalName -eq "text")
        {
            $ctrlValue = $ctrl.Text
        }
        elseif($elementNode.LocalName -eq "multiText")
        {
            $ctrlValue = $ctrl.Text -replace [Environment]::NewLine,"&#xF000;"
        }         
        elseif($elementNode.LocalName -eq "enum")
        {
            $ctrlValue = $ctrl.SelectedValue
        }
        elseif($elementNode.LocalName -eq "list")
        {
            $i = 1
            $keyValueArr = @()
            foreach($keyValue in $ctrl.ItemsSource)
            {
                if(-not $keyValue.Value) { continue }
                $keyValueArr += "$((?? $keyValue.Key $i))&#xF000;$($keyValue.Value)"
                $i++
            }
            $ctrlValue = $keyValueArr -join "&#xF000;"
        }
        elseif($elementNode.LocalName -eq "decimal" -or
            $ctrl.Tag.LocalName -eq "longDecimal")
        {            
            $ctrlValue = $ctrl.Children[0].Text
        }
        elseif($elementNode.LocalName -eq "boolean")
        {
            if($ctrl -is [System.Windows.Controls.CheckBox])
            {
                $ctrlValue = ?: $ctrl.IsChecked "1" "0"
                if($ctrl.IsChecked -eq $false)
                {
                    #continue # GPO setting skips unchecked checkbox. 
                }
            }
            elseif($ctrl -is [System.Windows.Controls.ComboBox])
            {
                $ctrlValue = $ctrl.SelectedValue
            }
            else
            {
                Write-Log "Boolean element type not supported: $($elementNode.LocalName)" 2
                continue
            }
        }
        else
        {
            Write-Log "Element type not supported: $($elementNode.LocalName)" 2
            continue
        }

        if(-not $ctrlValue)
        {
            if($elementNode.required -eq $true)
            {
                Write-Log "Required value is missing for $($elementNode.Id)" 3
            }
            else
            {
                Write-Log "Value not set for $($elementNode.Id). Value will not be added"
            }
            continue
        }

        $policySettings += "<data id=`"$($elementNode.Id)`" value=`"$($ctrlValue)`"/>"
    }
    $policySettings -join [Environment]::NewLine # Or ""?
}

function Get-ADMXCategoryOMAURIPath
{
    param($categoryId)

    $catObj = $script:admxXML.policyDefinitions.categories.selectSingleNode("$($script:xmlNSPrefix)category[@name='$($categoryId)']",$script:xmlNS)
    
    $categories = @()
    while($catObj)
    {
        $categories += $catObj.name
        $catObj = $script:admxXML.policyDefinitions.categories.selectSingleNode("$($script:xmlNSPrefix)category[@name='$($catObj.parentCategory.ref)']",$script:xmlNS)
    }
    
    [array]::Reverse($categories)

    $categories -join "~"
}

function Import-ADMXPolicy
{
    if(-not $global:txtADMXProfileName.Text.Trim())
    {
        [System.Windows.MessageBox]::Show("Profile Name name must be specified", "Error", "OK", "Error")
        return
    }

    if(-not $global:txtADMXPolicyFileName.Text.Trim())
    {
        [System.Windows.MessageBox]::Show("ADMX Policy Name name must be specified", "Error", "OK", "Error")
        return
    }

    $admxConfiguredSettings = @()

    foreach($admxPolicy in $($script:admxPolicies | Where { $_.SettingStatus -eq 0 -or $_.SettingStatus -eq 1 }))    
    {
        if($admxPolicy.SettingStatus -eq 0)
        {
            $policyValue = "<disabled />"            
        }
        else
        {
            $policyValue = "<enabled />"
            
            $strValue = $admxPolicy.PolicySettings

            if($strValue)
            {
                $policyValue += "`n`n$strValue"
            }            
        }

        $catPath = Get-ADMXCategoryOMAURIPath $admxPolicy.Definition.parentCategory.ref

        $omaUriPath = "./$($admxPolicy.SettingClass)/Vendor/MSFT/Policy/Config/$($global:txtADMXPolicyAppName.Text)~Policy~$catPath/$($admxPolicy.Definition.name)"
        
        if($admxPolicy.OMAURIDescription)
        {
            $desc = $admxPolicy.OMAURIDescription
        }
        elseif($admxPolicy.Description -and $admxPolicy.Description.Length -gt 1000)
        {
            $desc = $admxPolicy.Description.SubString(0,1000)
        }
        else
        {
            $desc = $admxPolicy.Description 
        }

        $admxConfiguredSettings += [PSCustomObject]@{
                        "@odata.type" = "#microsoft.graph.omaSettingString"
                        "displayName" = (?? $admxPolicy.OMAURIName $admxPolicy.Name)
                        "description" =  $desc
                        "omaUri" =  $omaUriPath
                        "value" = $policyValue
                        }        
    }

    $intuneObj = [PSCustomObject]@{
        "@odata.type" = "#microsoft.graph.windows10CustomConfiguration"
        "displayName" = $global:txtADMXProfileName.Text
        "omaSettings" =  @()
        "roleScopeTagIds" = @()
        "assignments" = @()
        }

    if($global:txtADMXProfileDescription.Text)
    {
        $intuneObj | Add-Member -MemberType NoteProperty -Name "description" -Value $global:txtADMXProfileDescription.Text
    }

    if($global:chkADMXPolicyIngest.IsChecked)
    {
        $intuneObj.omaSettings += [PSCustomObject]@{
            "@odata.type" = "#microsoft.graph.omaSettingString"
            "displayName" = (?? $global:txtADMXPolicyIngestName.Text ("$($script:currentADMXFile.Name) Ingestion"))
            "description" =  $null
            "omaUri" =  "./Device/Vendor/MSFT/Policy/ConfigOperations/ADMXInstall/$($global:txtADMXPolicyAppName.Text)/Policy/$($global:txtADMXPolicyFileName.Text)"
            "value" = (Format-XML $script:admxXML)
            }
    }

    $intuneObj.omaSettings += $admxConfiguredSettings

    $json = $intuneObj | ConvertTo-Json -Depth 20

    $obj = Invoke-GraphRequest -Url "/deviceManagement/deviceConfigurations" -Body $json -Method "POST"
    
    if($obj)
    {
        Write-Log "Device configuration profile '$($intuneObj.displayName)' created with id $($obj.Id)"
    }
    else
    {
        $text = "Failed to create device configuration profile '$($intuneObj.displayName)'"
        Write-Log $text 3
        [System.Windows.MessageBox]::Show(($text + [Environment]::NewLine + [Environment]::NewLine + 'Check log for errors'), "Error!", "OK", "Error")
    }    
}

#endregion

#region Reg Values
function Show-ADMXRegValues
{
    if(-not $script:frmADMXRegProfile)
    {
        $script:frmADMXRegProfile = Get-XamlObject ($global:AppRootFolder + "\Xaml\EndpointManagerToolsADMXRegValues.xaml") -AddVariables

        if(-not $script:frmADMXRegProfile) { return }

        $global:cbADMXRegPolicyType.ItemsSource = @(
            [PSCustomObject]@{
                Name = "Policy"
                Value = "Policy"
            },
            [PSCustomObject]@{
                Name = "Preference"
                Value = "Preference"
            }
        )
        
        $script:ADMXRegProfile = [ADMXRegProfile]::new()
        $script:frmADMXRegProfile.DataContext = $script:ADMXRegProfile

        $global:btnADMXRegClear.Add_Click({        
            $script:addedRegValues.Clear()
            $global:txtADMXRegProfileName.Text = ""
            $global:txtADMXRegProfileDescription.Text = ""
            $global:cbADMXRegPolicyType.SelectedValue = $null
        })

        $global:dgADMXRegAddedPolicies.Add_MouseDoubleClick({        
            if(-not $this.SelectedItem) { return }
            Show-ADMXRegSettings $this.SelectedItem
        })  

        $global:mnuADMXRegPoliciesContextMenu.Add_Opened({
            $global:mnuADMXRegPolicyEdit.IsEnabled = $null -ne $global:dgADMXRegAddedPolicies.SelectedItem
        })
        
        $global:mnuADMXRegPolicyEdit.Add_Click({
            if(-not $global:dgADMXRegAddedPolicies.SelectedItem) { return }
            Show-ADMXRegSettings $global:dgADMXRegAddedPolicies.SelectedItem
        })    

        $global:btnADMXAddRegValue.Add_Click({
            Show-ADMXRegSettings
        })
        
        $global:btnADMXRegImport.Add_Click({
            Write-Status "Import Reg Settings Policy"
            Import-ADMXRegProfile
            Write-Status ""
        })
    }
    
    $global:grdToolsMain.Children.Clear()
    $global:grdToolsMain.Children.Add($frmADMXRegProfile)
}

function Show-ADMXRegSettings 
{
    param($regProfile)
    
    $script:frmADMXRegPolicies = Get-XamlObject ($global:AppRootFolder + "\Xaml\EndpointManagerToolsADMXAddRegPolicy.xaml") -AddVariables

    if(-not $script:frmADMXRegPolicies) { return }

    $newValue = [ADMXRegPolicyElement]::new()

    if(-not $regProfile)
    {
        $regProfile = [ADMXRegPolicy]::new()
        $script:newRegPolicy = $true
    }
    else
    {
        $script:newRegPolicy = $false
    }

    $script:frmADMXRegPolicies.DataContext = [PSCustomObject]@{
        RegPolicy = $regProfile
        PolicyElement = $newValue
    }

    $regHives = @()

    # Should this be listed for preferences?
    $regHives += [PSCustomObject]@{
            Name = "HKEY_LOCAL_MACHINE"
            Value = "HKLM"
        }

    $regHives += [PSCustomObject]@{
        Name = "HKEY_CURRENT_USER"
        Value = "HKCU"
    }

    $global:cbADMXRegHive.ItemsSource = $regHives

    $global:cbADMXRegPolicyStatus.ItemsSource = @(
        [PSCustomObject]@{
            Name = "Enabled"
            Value = "Enabled"
        },
        [PSCustomObject]@{
            Name = "Disabled"
            Value = "Disabled"
        }      
    )

    $global:dgADMXRegAddedElements.Add_MouseDoubleClick({        
        if(-not $this.SelectedItem) { return }
        $global:btnADMXRegElementAdd.Visibility = "Collapsed"
        $global:btnADMXRegElementNew.Visibility = "Visible"            

        $selectedItem = $this.SelectedItem
        $tmp = $script:frmADMXRegPolicies.DataContext
        $script:frmADMXRegPolicies.DataContext = $null
        $tmp.PolicyElement = $selectedItem
        $script:frmADMXRegPolicies.DataContext = $tmp

        Set-ADMXRegAttributeControls

        [System.Windows.Forms.Application]::DoEvents()
    })  

    $global:cbADMXRegElementDataType.ItemsSource = @(
        [PSCustomObject]@{
            Name = "String"
            Value = "text"
        },
        [PSCustomObject]@{
            Name = "Multi-string"
            Value = "multiText"
        },
        [PSCustomObject]@{
            Name = "List"
            Value = "list"
        },
        [PSCustomObject]@{
            Name = "DWORD (32-bit)"
            Value = "decimal"
        }
        # Looks like longDecimal is not supported
        #,
        #[PSCustomObject]@{
        #    Name = "QWORD (64-bit)"
        #    Value = "longDecimal"
        #}        
    )

    Add-XamlEvent $script:frmADMXRegPolicies "cbADMXRegElementDataType" "Add_SelectionChanged"({
        Set-ADMXRegAttributeControls        
    })      

    Add-XamlEvent $script:frmADMXRegPolicies "btnADMXRegElementAdd" "add_click" ({        

        if($global:txtADMXRegElementKey.Text -and (Get-ADMXRegIsKeySupported $global:txtADMXRegElementKey.Text) -eq $false)
        {
            return
        }
        
        if(-not $script:frmADMXRegPolicies.DataContext.RegPolicy.Key)
        {
            [System.Windows.MessageBox]::Show("The Key value must be specified for the policy", "Error!","OK", "Error")
            return
        }
        
        $newValue = [ADMXRegPolicyElement]::new()
        $tmp = $script:frmADMXRegPolicies.DataContext
        $script:frmADMXRegPolicies.DataContext = $null
        $tmp.RegPolicy.PolicyElements.Add($tmp.PolicyElement)
        $tmp.PolicyElement = $newValue
        $script:frmADMXRegPolicies.DataContext = $tmp
        [System.Windows.Forms.Application]::DoEvents()
    })

    Add-XamlEvent $script:frmADMXRegPolicies "btnADMXRegElementNew" "add_click" ({
        $newValue = [ADMXRegPolicyElement]::new()
        $tmp = $script:frmADMXRegPolicies.DataContext
        $script:frmADMXRegPolicies.DataContext = $null
        #$tmp.RegPolicy.PolicyElements.Add($tmp.PolicyElement)
        $tmp.PolicyElement = $newValue
        $script:frmADMXRegPolicies.DataContext = $tmp
        [System.Windows.Forms.Application]::DoEvents()
        $global:btnADMXRegElementAdd.Visibility = "Visible"
        $this.Visibility = "Collapsed"
    })    

    Add-XamlEvent $script:frmADMXRegPolicies "btnADMXRegAddNew" "add_click" ({

        if((Get-ADMXRegIsKeySupported $global:txtADMXRegKey.Text) -eq $false)
        {
            return
        }

        if($script:newRegPolicy)
        {
            $script:frmADMXRegProfile.DataContext.ADMXPolicies.Add($script:frmADMXRegPolicies.DataContext.RegPolicy)
        }        

        Show-ModalObject
    })

    Add-XamlEvent $script:frmADMXRegPolicies "btnADMXRegCancel" "add_click" ({
        $script:frmADMXRegPolicies = $null
        Show-ModalObject
    })    
    Show-ModalForm "Add new reg policy" $script:frmADMXRegPolicies -HideButtons
}

function Set-ADMXRegAttributeControls
{
    if($global:cbADMXRegElementDataType.SelectedValue -eq "list" -or $global:cbADMXRegElementDataType.SelectedValue -eq "multiText")
    {
        $global:spADMXRegAttributeValueSeparator.Visibility = "Visible"
        $global:txtADMXRegAttributeValueSeparator.Visibility = "Visible"
    }
    else
    {
        $global:spADMXRegAttributeValueSeparator.Visibility = "Collapsed"
        $global:txtADMXRegAttributeValueSeparator.Visibility = "Collapsed"
    }


    if($global:cbADMXRegElementDataType.SelectedValue -eq "list")
    {
        $global:spADMXRegAttributeValuePrefix.Visibility = "Visible"
        $global:txtADMXRegAttributeValuePrefix.Visibility = "Visible"
    }
    else
    {
        $global:spADMXRegAttributeValuePrefix.Visibility = "Collapsed"
        $global:txtADMXRegAttributeValuePrefix.Visibility = "Collapsed"
    }

    $global:txtADMXRegElementValueName.IsReadOnly = ($global:cbADMXRegElementDataType.SelectedValue -eq "list")

    if($global:cbADMXRegElementDataType.SelectedValue -eq "text" -or 
        $global:cbADMXRegElementDataType.SelectedValue -eq "list")
    {
        $global:spADMXRegAttributeExpandable.Visibility = "Visible"
        $global:chkADMXRegAttributeExpandable.Visibility = "Visible"
    }
    else
    {
        $global:spADMXRegAttributeExpandable.Visibility = "Collapsed"
        $global:chkADMXRegAttributeExpandable.Visibility = "Collapsed"
    }

    if($global:cbADMXRegElementDataType.SelectedValue -eq "list")
    {
        $global:spADMXRegAttributeAdditive.Visibility = "Visible"
        $global:chkADMXRegAttributeAdditive.Visibility = "Visible"
    }
    else
    {
        $global:spADMXRegAttributeAdditive.Visibility = "Collapsed"
        $global:chkADMXRegAttributeAdditive.Visibility = "Collapsed"
    }    
}

function Get-ADMXRegIsKeySupported
{
    param($regKey)

    $tmpPath = $regKey.Trim('\')
    $tmpPath = $tmpPath + "\"

    $blockedSource = ""
    foreach($blockedPath in $script:unsupportedLocations)
    {
        if($tmpPath -like "$($blockedPath)\*")
        {
            $blockedSource = $blockedPath
            break
        }
    }

    if($blockedSource)
    {
        foreach($excemptionPath in $script:unsupportedOverride)
        {
            if($tmpPath -like "$($excemptionPath)\*")
            {
                $blockedSource = ""
                break
            }
        }
    }

    if($blockedSource)
    {
        [System.Windows.MessageBox]::Show("The registry key '$($global:txtADMXRegKey.Text)' is not supported" + [Environment]::NewLine + [Environment]::NewLine + "Blocked by root key: $blockedSource", "Unsupported reg ket", "OK", "Error")
        return $false
    }
    return $true
}

function Import-ADMXRegProfile
{    
    $xml = [xml]$script:admxTemplate

    $guidId = [Guid]::NewGuid()
    $rgPolicyFileName = ("RegPolicy_" + $guidId)

    $xml.policyDefinitions.categories.category.name = ($xml.policyDefinitions.categories.category.name + "_" + $guidId)

    $intuneObj = [PSCustomObject]@{
        "@odata.type" = "#microsoft.graph.windows10CustomConfiguration"
        "displayName" = $script:frmADMXRegProfile.DataContext.ProfileName
        "omaSettings" =  @()
        "roleScopeTagIds" = @()
        "assignments" = @()
        }

    $admxRegSettings = @()

    foreach($regPolicy in $script:frmADMXRegProfile.DataContext.ADMXPolicies)
    {
        if($regPolicy.PolicyStatus -eq "Enabled")
        {
            $OMAURIString = "<enabled />"
        }
        else
        {
            $OMAURIString = "<disabled />"
        }
        
        if($regPolicy.PolicyName)
        {
            # No space in name
            $policyName =  $regPolicy.PolicyName -replace " ","_"
        }
        else
        {
            $policyName = [Guid]::NewGuid().Guid
        }

        #build XML
        $newNode = $xml.policyDefinitions.policies.ChildNodes[0].CloneNode($true)
        $newNode.name = $policyName
        $newNode.class = (?: ($regPolicy.Hive -eq "HKLM") "Machine" "User")
        $newNode.displayName = "`$(string.$($policyName))"
        $newNode.presentation = "`$(presentation.$($policyName))"
        $newNode.key = $regPolicy.Key.Trim('\')
        $newNode.parentCategory.ref = $xml.policyDefinitions.categories.category.name        

        if(-not $regPolicy.StatusValueName)
        {
            $newNode.RemoveChild($newNode.SelectSingleNode("enabledValue"))
            $newNode.RemoveChild($newNode.SelectSingleNode("disabledValue"))
        }
        else
        {
            $newNode.valueName = $regPolicy.StatusValueName
        }

        if($regPolicy.PolicyElements -eq $null -or $regPolicy.PolicyElements.Count -eq 0)
        {
            $newNode.RemoveChild($newNode.SelectSingleNode("elements"))
        }
        else
        {
            $omaUriItems = @()
            foreach($element in $regPolicy.PolicyElements)
            {
                $child = $xml.CreateElement($element.DataType)
                if($element.DataType -eq "multitext" -or 
                        $element.DataType -eq "list")
                {
                    $splitter = ?? $global:txtADMXRegAttributeValueSeparator.Text ";"
                    $value = $element.Value -replace $splitter,"&#xF000;"
                }
                else
                {
                    $value = $element.Value
                }

                if($element.DataType -eq "list")
                {
                    # ToDo:
                    # Add support for additive in UI
                    Add-XMLAttribute $child "additive" "true"
                    Add-XMLAttribute $child "valuePrefix" $element.AttributePrefix
                }
                else
                {
                    Add-XMLAttribute $child "valueName" $element.ValueName
                }

                if(($element.DataType -eq "text" -or $element.DataType -eq "list") -and $element.AttributeExpandable)
                {
                    Add-XMLAttribute $child "expandable" "true"
                }

                if($element.DataType -eq "list" -and $element.AttributeAdditive)
                {
                    Add-XMLAttribute $child "additive" "true"
                }                

                if($element.AttributeSoft)
                {
                    Add-XMLAttribute $child "soft" "true"
                }

                if($element.DataType -eq "list")
                {
                    $keyStr = ?? $element.Key $regPolicy.Key
                    if($keyStr)
                    {                        
                        $idStr = $keyStr.Trim("\").Split('\')[-1]
                    }
                    else
                    {
                        $idStr = [Guid]::NewGuid().Guid
                    }
                    Add-XMLAttribute $child "id" ($idStr + "_Id")
                }
                else
                {
                    Add-XMLAttribute $child "id" ($element.ValueName + "_Id")
                }

                if($element.Key)
                {
                    Add-XMLAttribute $child "key" $element.Key.Trim('\')
                }

                $omaUriItems += "<data id=`"$(($element.ValueName + "_Id"))`" value=`"$($value)`"/>"
                
                $newNode.SelectSingleNode("elements").AppendChild($child) | Out-Null
            }
        }

        if($omaUriItems.Count -gt 0)
        {
            $OMAURIString = ($OMAURIString + [Environment]::NewLine + [Environment]::NewLine + ($omaUriItems -join [Environment]::NewLine)) 
        }

        $xml.policyDefinitions.SelectSingleNode("policies").AppendChild($newNode) | Out-Null

        $admxRegSettings += [PSCustomObject]@{
            "@odata.type" = "#microsoft.graph.omaSettingString"
            "displayName" = "Set $($policyName)"
            "omaUri" =  "./$((?: ($regPolicy.Hive -eq "HKLM") "Device" "User"))/Vendor/MSFT/Policy/Config/IntuneManagementReg~$(($script:frmADMXRegProfile.DataContext.PolicyType))~$($newNode.parentCategory.ref)/$($policyName)"
            "value" = $OMAURIString
        }    
    }

    $xml.policyDefinitions.SelectSingleNode("policies").RemoveChild($xml.policyDefinitions.policies.SelectSingleNode("policy")) | Out-Null    

    $intuneObj.omaSettings += [PSCustomObject]@{
        "@odata.type" = "#microsoft.graph.omaSettingString"
        "displayName" = "Reg ADMX Ingestion"
        "description" =  "This XML is generated by Intune Managemet tool"
        "omaUri" =  "./Device/Vendor/MSFT/Policy/ConfigOperations/ADMXInstall/IntuneManagementReg/$(($script:frmADMXRegProfile.DataContext.PolicyType))/$($rgPolicyFileName)"
        "value" = (Format-XML $xml)
        }    

    $intuneObj.omaSettings += $admxRegSettings

    $json = $intuneObj | ConvertTo-Json -Depth 20

    $obj = Invoke-GraphRequest -Url "/deviceManagement/deviceConfigurations" -Body $json -Method "POST"
    
    if($obj)
    {
        Write-Log "Custom profile '$($intuneObj.displayName)' created with id $($obj.Id)"
    }
    else
    {
        $text = "Failed to create device configuration profile '$($intuneObj.displayName)'"
        Write-Log $text 3
        [System.Windows.MessageBox]::Show(($text + [Environment]::NewLine + [Environment]::NewLine + 'Check log for errors'), "Error!", "OK", "Error")
    }
}

function Add-XMLAttribute
{
    param($xmlNode, $attribute, $Value)
    $xmlAttrib = $xmlNode.OwnerDocument.CreateAttribute($attribute)
    $xmlAttrib.Value = $Value
    $xmlNode.Attributes.Append($xmlAttrib)
}

function Add-ADMXRegClasses
{ 
    if (("ADMXRegPolicyElement" -as [type]))
    {
        return
    }
   
    $classDef = @"
    using System.ComponentModel;

    public class ADMXRegPolicyElement : INotifyPropertyChanged
    {
        public string DataType { get { return _dataType; } set { _dataType = value;  NotifyPropertyChanged("DataType"); NotifyPropertyChanged("DataTypeDisplayString");  } }
        private string _dataType = null;

        public string DataTypeDisplayString { get {
            if(DataType == "text")
                return "String";
            else if(DataType == "multiText")
                return "Multi-string";
            else if(DataType == "list")
                return "List";
            else if(DataType == "decimal")
                return "DWORD (32-bit)";
            else if(DataType == "longDecimal")
                return "QWORD (64-bit)";
            else
                return DataType;
        } }

        public string Key { get { return _key; } set { _key = value; NotifyPropertyChanged("Key"); } }
        private string _key;

        public string ValueName { get { return _valueName; } set { _valueName = value; NotifyPropertyChanged("ValueName"); } }
        private string _valueName;        

        public string Value { get { return _value; } set { _value = value; NotifyPropertyChanged("Value"); } }
        private string _value;

        public string AttributePrefix { get { return _attributePrefix; } set { _attributePrefix = value; NotifyPropertyChanged("AttributePrefix"); } }
        private string _attributePrefix;

        public bool AttributeSoft { get { return _attributeSoft; } set { _attributeSoft = value; NotifyPropertyChanged("AttributeSoft"); } }
        private bool _attributeSoft = false;        

        public bool AttributeExpandable { get { return _attributeExpandable; } set { _attributeExpandable = value; NotifyPropertyChanged("AttributeExpandable"); } }
        private bool _attributeExpandable = false;        

        public bool AttributeAdditive { get { return _attributeAdditive; } set { _attributeAdditive = value; NotifyPropertyChanged("AttributeAdditive"); } }
        private bool _attributeAdditive = false;        

        public event PropertyChangedEventHandler PropertyChanged;  

        // This method is called by the Set accessor of each property.  
        // The CallerMemberName attribute that is applied to the optional propertyName  
        // parameter causes the property name of the caller to be substituted as an argument.  
        private void NotifyPropertyChanged(string propertyName = "")  
        {  
            if(PropertyChanged != null) { PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName)); }
        }        
    }

    public class ADMXRegPolicy
    {
        public string PolicyName {get; set;}
        public string PolicyStatus {get; set;}
        public string Hive {get; set;}
        public string Key {get; set;}
        public bool StatusValueEnabled { get; set;}
        public string StatusValueName { get; set;}
        public System.Collections.ObjectModel.ObservableCollection<ADMXRegPolicyElement> PolicyElements {get; set;}

        public ADMXRegPolicy()
        {
            PolicyElements = new System.Collections.ObjectModel.ObservableCollection<ADMXRegPolicyElement>();
            Hive = "HKLM";
            PolicyStatus = "Enabled";
            StatusValueEnabled = true;
        }
    }

    public class ADMXRegProfile
    {
        public string ProfileName {get; set;}
        public string ProfileDescription {get; set;}
        public string PolicyType {get; set;}
        public System.Collections.ObjectModel.ObservableCollection<ADMXRegPolicy> ADMXPolicies {get; set;}
        public string XmlString {get; set;}

        public ADMXRegProfile()
        {
            ADMXPolicies = new System.Collections.ObjectModel.ObservableCollection<ADMXRegPolicy>();
            PolicyType = "Policy";
        }
    }
"@
    [Reflection.Assembly]::LoadWithPartialName("System.ComponentModel") | Out-Null
    Add-Type -TypeDefinition $classDef -IgnoreWarnings -ReferencedAssemblies @('System.ComponentModel')
}

#endregion