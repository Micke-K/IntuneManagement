<#
Documentation for Intune 

This module contains the base documentation function for document objects in Intune

This module will document Settings objects (Settings Catalog, Endpoint Security and Administrative Templates) and 
property objects like Configuration Profiles and Comliance Policies

Property objectes are documented based on the ObjectCategories.json file and the json files in the Documentation\ObjectInfo folder.
These json files contains the definition of each property of an object

Settings objects are documented with MS Graph APIs.

A basic Output provider is included that exports the objects to a CSV file

#>

$global:documentationOutput = @()
$global:documentationProviders = @()

function Get-ModuleVersion
{
    '2.3.0'
}

function Invoke-InitializeModule
{
    $script:alwaysUseMigTableForTranslation = $false

    # Make sure we add the default Output types
    Add-OutputType

    $script:columnHeaders = @{
        Name="Inputs.displayNameLabel"
        Value="TableHeaders.value"
        Description="TableHeaders.description"
        GroupMode="SettingDetails.modeTableHeader" #assignmentTypeSelectionLabel?
        Group="TableHeaders.assignedGroups"
        Groups="TableHeaders.groups"
        useDeviceContext="SettingDetails.installContextLabel"
        uninstallOnDeviceRemoval="SettingDetails.UninstallOnRemoval"
        isRemovable="SettingDetails.installAsRemovable"
        vpnConfigurationId="PolicyType.vpn"
        Action="SettingDetails.actionColumnName"
        Schedule="ScheduledAction.List.schedule"
        MessageTemplate="ScheduledAction.Notification.messageTemplate"
        EmailCC="ScheduledAction.Notification.additionalRecipients"
        Rule="ApplicabilityRules.GridLabel.Rule"
        ValueWithLabel="TableHeaders.value"
        Status="TableHeaders.status"
        CombinedValueWithLabel="TableHeaders.value"
        CombinedValue="TableHeaders.value"
        useDeviceLicensing="TableHeaders.licenseType"
        Filter="AppResources.AppSettingsUx.assignmentFilterColumnHeader"
        filterMode="AppResources.AppSettingsUx.assignmentFilterTypeColumnHeader"
        deliveryOptimizationPriority="AppResources.AppSettingsUx.deliveryOptimizationPriorityHeader"
        startTimeColumnLabel="AppResources.AppSettingsUx.startTimeColumnLabel"
        installTimeSettings="AppResources.AppSettingsUx.deadlineTimeColumnLabel"
        restartSettings="AppResources.AppSettingsUx.restartGracePeriodHeader"
        notifications="AppResources.AppSettingsUx.assignmentToast"
        Settings="TableHeaders.settings"
        returnCode='Win32ReturnCodes.Columns.returnCode'
        type='Win32ReturnCodes.Columns.codeType'
        RecommendedValue="AzureIAMCommon.Recommended"
    }    
}

function Invoke-ShowMainWindow
{
    $button = [System.Windows.Controls.Button]::new()
    $button.Content = "Document"
    $button.Name = "btnDocument"
    $button.MinWidth = 100
    $button.Margin = "0,0,5,0" 
    $button.IsEnabled = $false
    $button.ToolTip = "Document selected objects"

    $button.Add_Click({ 

        $objects = ?? ($global:dgObjects.ItemsSource | Where IsSelected -eq $true) $global:dgObjects.SelectedItem 

        Show-DocumentationForm -Objects $objects
    })    

    $global:spSubMenu.RegisterName($button.Name, $button)

    $global:spSubMenu.Children.Insert(0, $button)
}

function Invoke-EMSelectedItemsChanged
{
    $hasSelectedItems = ($global:dgObjects.ItemsSource | Where IsSelected -eq $true) -or ($null -ne $global:dgObjects.SelectedItem)
    Set-XamlProperty $global:dgObjects.Parent "btnDocument" "IsEnabled" $hasSelectedItems
}

function Invoke-GraphObjectsChanged
{
    $btnDocument = $global:spSubMenu.Children | Where-Object { $_.Name -eq "btnDocument" }
    $btnExport = $global:spSubMenu.Children | Where-Object { $_.Name -eq "btnExport" }
    if($btnDocument -and $btnExport)
    {
        $btnDocument.Visibility = $btnExport.Visibility 
    }
}

function Invoke-ViewActivated
{
    if($global:currentViewObject.ViewInfo.ID -ne "IntuneGraphAPI") { return }
    
    $tmp = $mnuMain.Items | Where Name -eq "EMBulk"
    if($tmp)
    {
        $tmp.AddChild(([System.Windows.Controls.Separator]::new())) | Out-Null
        $subItem = [System.Windows.Controls.MenuItem]::new()
        $subItem.Header = "_Document Types"
        $subItem.Add_Click({Invoke-DocumentObjectTypes})
        $tmp.AddChild($subItem)

        $subItem = [System.Windows.Controls.MenuItem]::new()
        $subItem.Header = "D_ocument Selected"
        $subItem.Add_Click({Invoke-DocumentSelectedObjects})
        $tmp.AddChild($subItem)       
    }
}

function Set-DocColumnHeaderLanguageId
{
    param($columnName, $lngId)

    if(-not $script:columnHeaders -or -not $lngId) { return }

    if($script:columnHeaders.ContainsKey($columnName))
    {
        $script:columnHeaders[$columnName] = $lngId
    }
    else
    {
        $script:columnHeaders.Add($columnName, $lngId)
    }
}

function Invoke-DocTranslateColumnHeader
{
    param($columnName)

    $lngText = ""
    if($script:columnHeaders.ContainsKey($columnName))
    {
        $lngText = Get-LanguageString $script:columnHeaders[$columnName]
    }

    (?? $lngText $columnName)
}

function Add-OutputType
{
    param($outputInfo)

    if(-not $global:documentationOutput)
    {
        $global:documentationOutput = @()
    }

    if($global:documentationOutput.Count -eq 0)
    {
        $global:documentationOutput += [PSCustomObject]@{
            Name = "None (Raw output only)"
            Value = "none"
        }

        $global:documentationOutput += [PSCustomObject]@{
            Name="CSV"
            Value="csv"
            OutputOptions = (Add-CSVOptionsControl)
            Activate = { Invoke-CSVActivate @args }
            PreProcess = { Invoke-CSVPreProcessItems @args }
            Process = { Invoke-CSVProcessItem @args }
        }        
    }

    if(!$outputInfo) { return }

    $global:documentationOutput += $outputInfo
}

function Add-DocumentationProvicer
{
    param($docProvider)

    if(-not $global:documentationProviders)
    {
        $global:documentationProviders = @()
    }

    $global:documentationProviders += $docProvider
}

function Get-ObjectDocumentation
{
    param($documentationObj)

    $additionalInfo = ""
    if($documentationObj.Object.'@ObjectFromFile' -eq $true)
    {
        $additionalInfo = " - From File"
    }

    Write-Status "Get documentation info for $((Get-GraphObjectName $documentationObj.Object $documentationObj.ObjectType)) ($($documentationObj.ObjectType.Title))$additionalInfo"

    $status = $null
    $inputType = "Settings"    

    if(-not $script:scopeTags -and $script:offlineDocumentation -ne $true)
    {
        $script:scopeTags = (Invoke-GraphRequest -Url "/deviceManagement/roleScopeTags").Value
    }   

    $obj = $documentationObj.Object
    $objectType = $documentationObj.ObjectType
    $script:currentObject = $obj

    $script:languageStrings = $null
 
    $script:CurrentSubCategory = $null
    $script:objectBasicInfo = @()
    $script:objectSettingsData = @()
    $script:objectComplianceActionData = @()
    $script:applicabilityRules = @()
    $script:objectAssignments = @()
    $script:objectScripts = @()
    $script:customTables = @()
    $script:admxCategories = $null

    $script:ObjectTypeFullTable = @{} # Hash table with objects that should be documented in a single table eg ScopeTags

    $updateFilteredObject = $true

    $type = $obj.'@OData.Type'

    $defaultDocumentationProperties =  @("Name","Value")

    $retObj = $null
    foreach($docProvider in ($global:documentationProviders | Sort -Property Priority))
    {
        if($docProvider.DocumentObject)
        {
            $retObj = & $docProvider.DocumentObject $documentationObj
            if($retObj -ne $null)
            {
                break
            }
        }
    }
    
    $propertyObjectProperties = @("Name","Value","Category","SubCategory","RawValue","RawJsonValue","DefaultValue","UnconfiguredValue","EntityKey","Description","Enabled")

    #region Custom documentation
    if($retObj)
    {
        $status =  $retObj.ErrorText
        $properties = $retObj.Properties
        $inputType = ?? $retObj.InputType "Property"
        $defaultDocumentationProperties = ?? $retObj.DefaultDocumentationProperties $defaultDocumentationProperties
    }
    #endregion
    #region Manually created file - ODataType
    elseif([IO.File]::Exists(($global:AppRootFolder + "\Documentation\ObjectInfo\$($obj.'@OData.Type').json")))
    {
        $inputType = "Property"
        # Process object based on OData type
        Invoke-TranslateCustomProfileObject $obj "$($obj.'@OData.Type')" | Out-Null
        $properties = $propertyObjectProperties
    }
    #endregion
    #region Manually created file - ObjectType id
    elseif($objectType -and [IO.File]::Exists(($global:AppRootFolder + "\Documentation\ObjectInfo\#$($objectType.Id).json")))
    {
        $inputType = "Property"
        # Process object based on Intune Object Type ($objectType)
        # '#' is added to front of name to distinguish manually created files from generated files
        Invoke-TranslateCustomProfileObject $obj "#$($objectType.Id)" | Out-Null
        $properties = $propertyObjectProperties
    }    
    #endregion
    #region Settings Catalog
    elseif($type -eq "#microsoft.graph.deviceManagementConfigurationPolicy")
    {  
        Invoke-TranslateSettingsObject $obj $objectType | Out-Null
        $properties = @("Name","Value","RootCategory","Category","RawValue","RawJsonValue","DefaultValue","Description")
    }
    #endregion
    #region Endpoint Security
    elseif($type -eq "#microsoft.graph.deviceManagementIntent")
    {
        Invoke-TranslateIntentObject $obj $objectType | Out-Null
        $properties = @("Name","Value","Category","FullValueTable","RawValue","RecommendedValue","SettingId","Description")
        $defaultDocumentationProperties = @("Name","Value","RecommendedValue")
    }
    #endregion
    #region Administrative Templates
    elseif($type -eq "#microsoft.graph.groupPolicyConfiguration")
    {
        Invoke-TranslateADMXObject $obj $objectType | Out-Null
        $properties = @("Name","Status","Value","Category","CategoryPath","RawValue","ValueWithLabel","Created","Modified", "Class", "DefinitionId")
        $defaultDocumentationProperties =  @("Name","Status",(?? $script:ValueOutputProperty "Value"))
        $updateFilteredObject = $false
        $inputType = "Property"
    }
    #endregion
    #region Profile Types e.g. DeviceConfiguration Policies etc
    elseif($type)
    {
        $inputType = "Property"
        $processed = $true

        
        $processed = Invoke-TranslateProfileObject $obj

        if($processed -eq $false) 
        {
            $errText = "No object file or object info found for $((Get-GraphObjectName $obj $objType)) ($($obj.'@OData.Type'))`n`nObject type not supported for documentation"
            # Object not processed
            $status = $errText
            Write-Log $errText 3             
        }
        $properties = $propertyObjectProperties
    }
    #endregion


    if($objectType.DocumentAll -eq $true) { return }

    if($script:objectBasicInfo.Count -gt 0)
    {
        Add-ScopeTagStrings $obj
    }

    if($obj.scheduledActionsForRule.scheduledActionConfigurations)
    {
        Invoke-TranslateScheduledActionType $obj.scheduledActionsForRule
    }    

    if(($obj.assignments | measure).Count -gt 0)
    {
        Invoke-TranslateAssignments $obj
    }

    $script:settingsProperties = $properties

    $objectDocumentationInfo = [PSCustomObject]@{
        BasicInfo = $script:objectBasicInfo
        Settings = $script:objectSettingsData
        ComplianceActions = $script:objectComplianceActionData
        ApplicabilityRules = $script:applicabilityRules
        CustomTables = $script:customTables
        Assignments = $script:objectAssignments
        Scripts = $script:objectScripts
        DisplayProperties = $properties
        DefaultDocumentationProperties = $defaultDocumentationProperties
        ErrorText = $status
        InputType = $inputType
        UpdateFilteredObject = $updateFilteredObject
        UnconfiguredProperties = $obj."@UnconfiguredProperties"
    }
    
    foreach($docProvider in ($global:documentationProviders | Sort -Property Priority))
    {
        if($docProvider.ObjectDocumented)
        {
            & $docProvider.ObjectDocumented $obj $objectType $objectDocumentationInfo
        }
    }    

    $objectDocumentationInfo
}

function Get-DocumentedSettings
{
    $script:objectSettingsData
}

function Invoke-ObjectDocumentation
{
    param($documentationObj)

    $global:intentCategories = $null
    $global:catRecommendedSettings = $null
    $global:intentCategoryDefs = $null
    $global:cfgCategories = $null
    $script:admxCategories = $null
    $script:migTable = $null

    $script:DocumentationLanguage = "en"        
    $script:objectSeparator = [System.Environment]::NewLine
    $script:propertySeparator = ","

    $loadExportedInfo = $false

    if($documentationObj.Object."@ObjectFileName") {
        $path = [IO.Path]::GetDirectoryName($documentationObj.Object."@ObjectFileName")
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
                    Write-Log "Load Migration table from $migFileName"
                    $script:migTable = ConvertFrom-Json (Get-Content $migFileName -Raw)        
                }
            }
            catch {}
        }
        if(-not $script:migTable)  {
            Write-Log "Migration table not found" 2
        }
    }

    Get-ObjectDocumentation $documentationObj
}

function Add-RawDataInfo
{
    param($obj, $objectType)

    $params = @{}

    if($txtDocumentationRawData.Text)
    {
        $txtDocumentationRawData.Text += "`n`n"    
    }

    $txtDocumentationRawData.Text += "#########################################################" 
    $txtDocumentationRawData.Text += "`n`n# Object: $((Get-GraphObjectName $obj $objectType))" 
    $txtDocumentationRawData.Text += "`n# Type: $($obj.'@OData.Type')" 
    $txtDocumentationRawData.Text += "`n`n#########################################################" 

    if($script:objectBasicInfo.Count -gt 0)
    {
        $txtDocumentationRawData.Text += "`n`n# Basic Info`n`n" 
        $txtDocumentationRawData.Text += (($script:objectBasicInfo | Select -Property Name,Value | ConvertTo-Csv -NoTypeInformation @params) -join ([System.Environment]::NewLine))
    }

    if($script:objectSettingsData)
    {
        $txtDocumentationRawData.Text += "`n`n# Object Settings`n`n" 
        $txtDocumentationRawData.Text += (($script:objectSettingsData | Select -Property $script:settingsProperties | ConvertTo-Csv -NoTypeInformation @params) -join ([System.Environment]::NewLine))
    }

    if(($documentedObj.ApplicabilityRules | measure).Count -gt 0)
    {
        $txtDocumentationRawData.Text += "`n`n# Applicability rules`n`n" 
        $txtDocumentationRawData.Text += (($script:applicabilityRules | Select Rule,Property,Value,Category | ConvertTo-Csv -NoTypeInformation) -join ([System.Environment]::NewLine))
    }    

    if($script:objectComplianceActionData.Count -gt 0)
    {
        $txtDocumentationRawData.Text += "`n`n# Compliance actions`n`n" 
        $txtDocumentationRawData.Text += (($script:objectComplianceActionData | Select Action,Schedule,MessageTemplate,EmailCC,Category | ConvertTo-Csv -NoTypeInformation) -join ([System.Environment]::NewLine))
    }

    if($script:objectAssignments.Count -gt 0)
    {
        if($script:objectAssignments.Count -gt 0 -and $script:objectAssignments[0].RawIntent)
        {
            $properties = @("GroupMode","Group","Filter","FilterMode","Category","SubCategory","RawIntent")
        }
        else
        {
            $properties = @("GroupMode","Group","Filter","FilterMode","Category")
        }

        $txtDocumentationRawData.Text += "`n`n# Assignments`n`n" 
        $txtDocumentationRawData.Text += (($script:objectAssignments | Select -Property $properties | ConvertTo-Csv -NoTypeInformation) -join ([System.Environment]::NewLine))
    }
}

function Get-ObjectTypeGroupName
{
    param($objectType)

    if($objectType.Id -eq "DeviceConfiguration")
    {
        return (Get-LanguageString "PolicySelection.Templates.title")
    }
    elseif($objectType.Id -eq "ConditionalAccess")
    {
        return (Get-LanguageString "TermsOfUse.Details.Tab.cAPolicies")
    }
    else
    {
        return $ObjectType.Title
    }

}

function Get-ObjectTypeString
{
    param($obj, $objectType)

    $objTypeId = ?? $objectType.GroupId $objectType
    
    if($objTypeId -eq "DeviceConfiguration")
    {
        return (Get-LanguageString "SettingDetails.deviceConfigurationTitle")
        #!!!return (Get-LanguageString "PolicySelection.Templates.title")
        
    }
    elseif($objTypeId -eq "CompliancePolicies")
    {
        return (Get-LanguageString "SettingDetails.deviceComplianceTitle")
    }
    elseif($objTypeId -eq "Apps")
    {
        return (Get-LanguageString "SettingDetails.clientAppsTitle")
    }
    elseif($objTypeId -eq "WinEnrollment")
    {
        return (Get-LanguageString "SettingDetails.windowsEnrollmentTitle")
    }
    elseif($objTypeId -eq "AppleEnrollment")
    {
        return (Get-LanguageString "SettingDetails.appleEnrollmentTitle")
    }
    elseif($objTypeId -eq "PolicySets")
    {
        return (Get-LanguageString "SettingDetails.policySetsTitle")
    }    
    elseif($objTypeId -eq "AppConfiguration")
    {
        return (Get-LanguageString "SettingDetails.appConfiguration")
    }    
    elseif($objTypeId -eq "AppProtection")
    {
        return (Get-LanguageString "SettingDetails.appProtectionPolicy")
    } 
    elseif($objTypeId -eq "CustomAttributes")
    {
        return (Get-LanguageString "Titles.customAttributes")
    }   
    elseif($objTypeId -eq "ConditionalAccess")
    {
        return (Get-LanguageString "SecurityTemplate.conditionalAccess")
    }
    elseif($objTypeId -eq "EndpointAnalytics")
    {
        return (Get-LanguageString "SettingDetails.healthMonScopeBootPerf")
    }
    elseif($objTypeId -eq "EndpointSecurity")
    {
        return (Get-LanguageString "PolicyType.EndpointSecurityTemplate.default")
    }
    elseif($objTypeId -eq "EnrollmentRestrictions")
    {
        return (Get-LanguageString "Titles.enrollmentRestrictions")
    }
    elseif($objTypeId -eq "Scripts")
    {
        return (Get-LanguageString "Titles.scriptManagement")
    }
    elseif($objTypeId -eq "WinUpdatePolicies")
    {
        return (Get-LanguageString "Titles.windows10UpdateRings")
    }
    elseif($objTypeId -eq "WinFeatureUpdates")
    {
        return (Get-LanguageString "Titles.featureUpdateDeployments")
    }
    elseif($objTypeId -eq "WinQualityUpdates")
    {
        return (Get-LanguageString "Titles.windows10QualityUpdate")
    }    
    elseif($objTypeId -eq "WinDriverUpdatePolicies")
    {
        return (Get-LanguageString "Titles.windows10DriverUpdate")
    }
    elseif($objTypeId -eq "TenantAdmin")
    {
        return (Get-LanguageString "Titles.tenantAdmin")
    }
    elseif($objTypeId -eq "Azure")
    {
        return "Azure"
    }    
}

function Add-BasicDefaultValues
{
    param($obj, $objectType)

    Add-BasicPropertyValue (Get-LanguageString "SettingDetails.nameName") (Get-GraphObjectName $obj $objectType)
    Add-BasicPropertyValue (Get-LanguageString "SettingDetails.descriptionName") $obj.description

    $objInfo = Get-TranslationFiles $obj.'@OData.Type'
    $appType = Get-GraphAppType $obj

    if($objInfo)
    {
        $platformType = Get-LanguageString "Platform.$($objInfo.PlatformLanguageId)"
        $profileType = Get-LanguageString "ConfigurationTypes.$($objInfo.PolicyType)"

        if($platformType) { Add-BasicPropertyValue (Get-LanguageString "SettingDetails.platformSupported") $platformType }
        if($profileType) { Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") $profileType }
    }

    if($appType)
    {
        $appTypeName = Get-LanguageString "AppType.$($appType.LanguageId)"
        if($appTypeName) { Add-BasicPropertyValue (Get-LanguageString "Inputs.installationSourceLabel") $appTypeName }
    }

    if($obj."@OData.Type" -eq "#microsoft.graph.deviceManagementIntent")
    {
        if(-not $script:baseLineTemplates)
        {
            $script:baseLineTemplates = (Invoke-GraphRequest -Url "/deviceManagement/templates").Value
        }

        $baseLineTemplate = $script:baseLineTemplates | Where Id -eq $obj.templateId
        if(-not $baseLineTemplate)
        {
            Write-Log "Could not find Baseline Template with Id $($obj.templateId)" 3            
        }
    
        $platformType = Get-LanguageString "Platform.$($baseLineTemplate.platformType)"

        if($platformType) { Add-BasicPropertyValue (Get-LanguageString "SettingDetails.platformSupported") $platformType }
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.Category") (Get-IntentCategory (?: ($baseLineTemplate.templateSubtype -eq "none") $baseLineTemplate.templateType $baseLineTemplate.templateSubtype))
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.policyType") $baseLineTemplate.displayName
    }

    if($obj.deviceManagementApplicabilityRuleDeviceMode)
    {
        # Future setting?
    }

    if($obj.deviceManagementApplicabilityRuleOsEdition)
    {
        $arrOSList = @()
        foreach($os in $obj.deviceManagementApplicabilityRuleOsEdition.osEditionTypes)
        {
            $arrOSList += Get-LanguageString "ApplicabilityRules.$os"
        }
        $script:applicabilityRules += [PSCustomObject]@{
            Rule = Get-LanguageString "ApplicabilityRules.$((?: ($obj.deviceManagementApplicabilityRuleOsEdition.ruleType -eq "exclude") "dontAssignIf" "assignIf"))"
            Property = Get-LanguageString "ApplicabilityRules.windows10OsEdition"
            Value = $arrOSList -join $script:objectSeparator
            Id = "deviceManagementApplicabilityRuleOsEdition"
            Category = Get-LanguageString "SettingDetails.applicabilityRules"
        }
    }
    if($obj.deviceManagementApplicabilityRuleOsVersion)
    {        
        $script:applicabilityRules += [PSCustomObject]@{
            Rule = Get-LanguageString "ApplicabilityRules.$((?: ($obj.deviceManagementApplicabilityRuleOsVersion.ruleType -eq "exclude") "dontAssignIf" "assignIf"))"
            Property = Get-LanguageString "ApplicabilityRules.windows10OsVersion"
            Value = "$($obj.deviceManagementApplicabilityRuleOsVersion.minOSVersion) $((Get-LanguageString "ApplicabilityRules.toText")) $($obj.deviceManagementApplicabilityRuleOsVersion.maxOSVersion)"
            Id = "deviceManagementApplicabilityRuleOsVersion"
            Category = Get-LanguageString "SettingDetails.applicabilityRules"
        }        
    }
}

function Add-CustomTable
{
    param($TableId, $Columns = @("Name", "Value"), $Values, [int]$Order = 100, $LanguageId = "")

    $script:customTables += [PSCustomObject]@{
        Id = $TableId
        Columns = $Columns
        Values = $Values
        LanguageId = $LanguageId
        Order = $Order
    }
}

function Add-BasicAdditionalValues
{
    param($obj, $objectType)

    if($obj.createdDateTime)
    {
        try
        {
            if($obj.createdDateTime -is [DateTime])
            {
                $tmpDate = $obj.createdDateTime
            }
            else
            {
                $tmpDate = ([DateTime]::Parse($obj.createdDateTime))
            }
            $tmpDateStr = "$($tmpDate.ToLongDateString()) $($tmpDate.ToLongTimeString())"
            
        }
        catch
        {
            Write-Log "Failed to parse date from $($obj.createdDateTime) (Object type: $($obj.createdDateTime.GetType().Name))" 2
            $tmpDateStr = $obj.createdDateTime
        }
        Add-BasicPropertyValue (Get-LanguageString "Inputs.createdDateTime") $tmpDateStr
    }

    if($obj.lastModifiedDateTime)
    {
        try
        {
            if($obj.lastModifiedDateTime -is [DateTime])
            {
                $tmpDate = $obj.lastModifiedDateTime
            }
            else
            {
                $tmpDate = ([DateTime]::Parse($obj.lastModifiedDateTime))
            }
            $tmpDateStr = "$($tmpDate.ToLongDateString()) $($tmpDate.ToLongTimeString())"
        }
        catch
        {
            Write-Log "Failed to parse date from $($obj.lastModifiedDateTime) (Object type: $($obj.lastModifiedDateTime.GetType().Name))" 2
            $tmpDateStr = $obj.lastModifiedDateTime
        }
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.lastModified") $tmpDateStr
    }
    elseif($obj.modifiedDateTime)
    {
        try
        {
            if($obj.modifiedDateTime -is [DateTime])
            {
                $tmpDate = $obj.modifiedDateTime
            }
            else
            {
                $tmpDate = ([DateTime]::Parse($obj.modifiedDateTime))
            }
            $tmpDateStr = "$($tmpDate.ToLongDateString()) $($tmpDate.ToLongTimeString())"
        }
        catch
        {
            Write-Log "Failed to parse date from $($obj.modifiedDateTime) (Object type: $($obj.modifiedDateTime.GetType().Name))" 2
            $tmpDateStr = $obj.modifiedDateTime
        }
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.lastModified") $tmpDateStr
    }

    if($obj.version)
    {
        Add-BasicPropertyValue (Get-LanguageString "SettingDetails.eDPPolicyAppsListVersionName") $obj.version
    }
    #Add-BasicPropertyValue (Get-LanguageString "TableHeaders.type") $obj.'@OData.Type'
}

function Add-BasicPropertyValue
{
    param($name, $value)

    $script:objectBasicInfo += [PSCustomObject]@{
        Name=$name
        Value=$value
        Category=(Get-LanguageString "SettingDetails.basics")
    }
}

function Add-ScopeTagStrings
{
    param($obj)

    $objScopeTags = @()
    if(($obj.PSObject.Properties | Where Name -eq "roleScopeTagIds"))
    {
        $scopeTagProperty = "roleScopeTagIds"
    }
    elseif(($obj.PSObject.Properties | Where Name -eq "roleScopeTags"))
    {
        $scopeTagProperty = "roleScopeTags"
    }
    else
    {
        return
    }
    
    if(($obj."$scopeTagProperty" | measure).Count -gt 0)
    {
        foreach($scopeTagId in $obj."$scopeTagProperty")
        {          
            $scopeTagName = $scopeTagId
            if($scopeTagId -eq "0")
            {
                $scopeTagName = (Get-LanguageString "SettingDetails.default")
            }
            elseif($script:scopeTags)
            {
                $scopeTagObj = $script:scopeTags | Where Id -eq $scopeTagId
                if($scopeTagObj.displayName)
                {
                    $scopeTagName = $scopeTagObj.displayName                
                }
            }
            $objScopeTags += $scopeTagName
        }
    }
    if($objScopeTags.Count -gt 0)
    {
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.scopeTags") ($objScopeTags -join $script:objectSeparator)
    }
}

function Add-ObjectScript
{
    param($Header, 
            $Caption, 
            $Script,
            [ValidateSet("Base64", "String")]
            [string]
            $ScriptType = "Base64")

    if($global:chkIncludeScripts.IsChecked -ne $true) { return }

    if($ScriptType -eq "Base64")
    {
        $scriptContent = Get-Base64ScriptContent $Script -RemoveSignature:$false
    }
    else
    {
        $scriptContent = $script
    }

    $RemoveSignature = $global:chkExcludeScriptSignature.IsChecked -eq $true
    if($RemoveSignature -eq $true)
    {
        $x = $scriptContent.IndexOf("# SIG # Begin signature block")
        if($x -gt 0)
        {
            $scriptContent = $scriptContent.SubString(0,$x)
            $scriptContent = $scriptContent + "# SIG # Begin signature block`nSignature data excluded..."
        }
    }    

    $script:objectScripts += [PSCustomObject]@{
        Header = $Header
        Caption = $Caption
        ScriptContent = $scriptContent
    }
}

function Get-ObjectPlatformFromType
{
    param($obj)

    $platform = $null

    $lowerAppType = $obj.'@OData.Type'.ToLower()
    if($lowerAppType.Contains("ios"))
    {
        $platform = "iOS"
    }
    elseif($lowerAppType.Contains("mac"))
    {
        $platform = "Mac"
    }
    elseif($lowerAppType.Contains("windowsphone"))
    {
        $platform = "WindowsPhone"
    }
    elseif($lowerAppType.Contains("windows") -or $lowerAppType.Contains("win32") -or $lowerAppType.Contains("mirosoftstore"))
    {
        $platform = "Windows10"
    }
    elseif($lowerAppType.Contains("androidForWork"))
    {
        $platform = "androidForWork"
    }
    elseif($lowerAppType.Contains("android"))
    {
        $platform = "Android"
    }

    $platform
}

#region Admin Templates

function Invoke-TranslateADMXObject
{
    param($obj, $objectType)

    Add-BasicDefaultValues $obj $objectType
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "Titles.groupPolicy")
    #Add-BasicPropertyValue (Get-LanguageString "SettingDetails.platformSupported") $platformType
    Add-BasicAdditionalValues $obj $objectType

    $params = @{}
    ## Set language
    if($script:DocumentationLanguage)
    {
        $params.Add("AdditionalHeaders", @{"Accept-Language"=$script:DocumentationLanguage})
    }

    if(-not $script:admxCategories)
    {
        $script:admxCategories = (Invoke-GraphRequest "deviceManagement/groupPolicyCategories?`$expand=parent(`$select=id, displayName, isRoot),definitions(`$select=id, displayName, categoryPath, classType, policyType)&`$select=id, displayName, isRoot" -ODataMetadata "skip" @params).value
    }

    if(!$obj.definitionValues)
    {
        $definitionValues = (Invoke-GraphRequest "deviceManagement/groupPolicyConfigurations('$($obj.Id)')/definitionValues?`$expand=definition(`$select=id,classType,displayName,policyType,groupPolicyCategoryId)" -ODataMetadata "minimal").value
    }
    else
    {
        # Documenting exported json
        $definitionValues = $obj.definitionValues
    }

    $enabledStr = Get-LanguageString "Inputs.enabled"
    $disabledStr = Get-LanguageString "Inputs.disabled"

    $propertyStr = Get-LanguageString "ApplicabilityRules.GridLabel.property"
    $valueStr = Get-LanguageString "ApplicabilityRules.GridLabel.value"


    foreach($definitionValue in $definitionValues)
    {
        if(-not $definitionValue.definition -and $definitionValues.'definition@odata.bind')
        {
            $url = $definitionValue.'definition@odata.bind' -replace $global:graphURL, ("https://$((?? $global:MSALGraphEnvironment "graph.microsoft.com"))/beta")
            $definition = Invoke-GraphRequest -Url $url -ODataMetadata "minimal" @params
            if($definition)
            {
                $definitionValue | Add-Member -MemberType NoteProperty -Name "definition" -Value $definition
            }            
        }

        $categoryObj = $script:admxCategories | Where { $definitionValue.definition.id -in ($_.definitions.id) }
        $category = $script:admxCategories.definitions | Where { $definitionValue.definition.id -in ($_.id) }
        $settingPresentationValues = $null

        # Get presentation values for the current settings (with presentation object included)
        if($definitionValue.presentationValues -or $obj.'@ObjectFromFile' -eq $true) #$definitionValue.'definition@odata.bind')
        {
            $settingPresentationValues = (Invoke-GraphRequest -Url "$($definitionValue.'definition@odata.bind')/presentations"  -ODataMetadata "minimal").value
            
            $presentationValues = @()
            if($settingPresentationValues)
            {
                # Do this to make sure they are documented in the correct order
                foreach($settingPresentationValue in $settingPresentationValues)
                {
                    $tmpPresentationVal = $definitionValue.presentationValues | Where 'presentation@odata.bind' -Like "*$($settingPresentationValue.Id)*"
                    if($tmpPresentationVal)
                    {
                        $presentationValues += $tmpPresentationVal
                    }
                    else 
                    {
                        $presentationValues = @()
                        break
                    }
                }
            }

            if($presentationValues.Count -eq 0)
            {
                Write-Log "Could not find definition for definition id '$($definitionValue.id)'. Values might be documented in the wrong order!" 2
                $presentationValues = $definitionValue.presentationValues
            }
        }
        elseif($definitionValue.id)
        {
            $presentationValues = (Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations/$($obj.id)/definitionValues/$($definitionValue.id)/presentationValues?`$expand=presentation"  -ODataMetadata "minimal" @params).value                
        }
        else
        {
            $presentationValues = $null    
        }

        $value = $null
        $rawValues = @()
        $values = @()
        $valuesWithLabel = @()
        $tableValue = @()

        foreach($presentationValue in $presentationValues)
        {
            if(-not $presentationValue.presentation -and $presentationValue.'presentation@odata.bind')
            {
                $presentation = Invoke-GraphRequest -Url $presentationValue.'presentation@odata.bind'
                if($presentation)
                {
                    $presentationValue | Add-Member -MemberType NoteProperty -Name "presentation" -Value $presentation
                }
            }
            
            $rawValue = $presentationValue.value
            $label = $presentationValue.presentation.label
            $value = $null
            if($presentationValue.presentation.'@odata.type' -eq '#microsoft.graph.groupPolicyPresentationDropdownList')
            {                
                $value = ($presentationValue.presentation.items | Where value -eq $rawValue).displayName
            }
            elseif($presentationValue.'@odata.type' -eq '#microsoft.graph.groupPolicyPresentationValueList')
            {
                $arrValues = @()
                foreach($tmpValue in $presentationValue.values)
                {
                    $arrValues += ($tmpValue.name + $script:propertySeparator + $tmpValue.value)
                }
                $value = $arrValues -join $script:objectSeparator
            }
            elseif($presentationValue.'@odata.type' -eq '#microsoft.graph.groupPolicyPresentationValueMultiText')
            {
                $value = $presentationValue.values -join $script:objectSeparator
            }
            else
            {
                #groupPolicyPresentationValueBoolean
                #groupPolicyPresentationValueDecimal
                #groupPolicyPresentationValueLongDecimal
                #groupPolicyPresentationValueText

                $value = $rawValue
            }
            $htFullValue = [ordered]@{}
            $htFullValue.Add($propertyStr,$label)
            $htFullValue.Add($valueStr,$value)

            $tableValue += [PSCustomObject]$htFullValue 

            $valuesWithLabel += "$label $value"
            $values += $value
            $rawValues += $rawValue
        }
        $status = (?: ($definitionValue.enabled -eq $true) $enabledStr $disabledStr)
        $script:objectSettingsData += New-Object PSObject -Property @{ 
            Name = $definitionValue.definition.displayName
            Description = $definitionValue.definition.explainText
            Status = $status
            Value = $values -join $script:objectSeparator
            CombinedValue = ($status + $script:objectSeparator + ($values -join $script:objectSeparator))
            ValueWithLabel = $valuesWithLabel -join $script:objectSeparator
            FullValueTable = $tableValue
            CombinedValueWithLabel = ($status + $script:objectSeparator + ($valuesWithLabel -join $script:objectSeparator))
            RawValue = $rawValues -join $script:propertySeparator
            Class = $definitionValue.definition.classType
            DefinitionId = $definitionValue.definition.id
            Created = $definitionValue.createdDateTime
            Modified = $definitionValue.lastModifiedDateTime
            #Category = $categoryObj.displayName
            Category = $category.categoryPath
            CategoryPath = $category.categoryPath
            EntityKey = $definitionValue.definition.id # Required for Compare
        }        
    }

    $script:objectSettingsData = $script:objectSettingsData | Sort -Property CategoryPath
}


#endregion


#region Settings Catalog

function Invoke-TranslateSettingsObject
{
    param($obj, $objectType)

    $platformType = Get-LanguageString "Platform.$($obj.platforms)"

    Add-BasicDefaultValues $obj $objectType
    Add-BasicPropertyValue (Get-LanguageString "TableHeaders.configurationType") (Get-LanguageString "ConfigurationTypes.settingsCatalog")
    if($obj.templateReference.templateId)
    {
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.Category") (Get-IntentCategory $obj.templateReference.templateFamily)
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.policyType") $obj.templateReference.templateDisplayName

        #Add-BasicPropertyValue (Get-LanguageString "PolicyType.EndpointSecurityTemplate.default") $obj.templateReference.templateDisplayName
    }
    Add-BasicPropertyValue (Get-LanguageString "SettingDetails.platformSupported") $platformType
    Add-BasicAdditionalValues $obj $objectType
    
    $params = @{}
    ## Set language
    if($script:DocumentationLanguage)
    {
        $params.Add("AdditionalHeaders", @{"Accept-Language"=$script:DocumentationLanguage})
    }

    <#
    if($obj.templateReference.templateId)
    {
        $cfgSettings = (Invoke-GraphRequest "/deviceManagement/configurationPolicyTemplates('$($obj.templateReference.templateId)')/settingTemplates?`$expand=settingDefinitions&top=1000" -ODataMetadata "minimal" @params).Value
    }
    else
    {
        $cfgSettings = (Invoke-GraphRequest "/deviceManagement/configurationPolicies('$($obj.Id)')/settings?`$expand=settingDefinitions&top=1000" -ODataMetadata "minimal" @params).Value
    }
    #>
    $cfgSettings = (Invoke-GraphRequest "/deviceManagement/configurationPolicies('$($obj.Id)')/settings?`$expand=settingDefinitions&top=1000" -ODataMetadata "minimal" @params).Value

    if($obj.'@ObjectFromFile')
    {
        $cfgSettings = $obj.Settings
    }

    if(-not $global:cfgCategories)
    {
        $global:cfgCategories = (Invoke-GraphRequest "/deviceManagement/configurationCategories?`$filter=platforms has 'windows10' and technologies has 'mdm'" -ODataMetadata "minimal" @params).Value
    }

    if(-not $global:cachedCfgSettings) 
    {
        $global:cachedCfgSettings = @{}
    }

    $script:settingCatalogasCategories = @{}
    foreach($cfgSetting in $cfgSettings)
    {
        if($obj.'@ObjectFromFile' -and -not $cfgSetting.settingDefinitions) 
        {
            if($global:cachedCfgSettings.ContainsKey($cfgSetting.settingInstance.settingDefinitionId) -eq $false) 
            {
                $defObj = Invoke-GraphRequest "/deviceManagement/configurationSettings/$($cfgSetting.settingInstance.settingDefinitionId)"
                $global:cachedCfgSettings.Add($defObj.Id, $defObj)
            }
        }
        else
        {
            $defObj = $cfgSetting.settingDefinitions | Where id -eq $cfgSetting.settingInstance.settingDefinitionId
            if($global:cachedCfgSettings.ContainsKey($cfgSetting.settingInstance.settingDefinitionId) -eq $false) 
            {                
                $global:cachedCfgSettings.Add($defObj.Id, $defObj)
            }
        }
        #$defObj = $cfgSetting.settingDefinitions | Where { $_.id -eq $cfgSetting.settingInstance.settingDefinitionId -or $_.id -eq $cfgSettings.settingInstanceTemplate.settingDefinitionId }
        if(-not $defObj -or $script:settingCatalogasCategories.ContainsKey($defObj.categoryId)) { continue }

        $catObj = $global:cfgCategories | Where Id -eq $defObj.categoryId 
        $rootCatObj = $global:cfgCategories | Where Id -eq $catObj.rootCategoryId

        #$catSettings = Invoke-GraphRequest "/deviceManagement/configurationSettings?`$filter=categoryId eq '$($defObj.categoryId)' and applicability/platform has 'windows10' and applicability/technologies has 'mdm'" -ODataMetadata "minimal" @params

        $script:settingCatalogasCategories.Add($defObj.categoryId, (New-Object PSObject -Property @{ 
            Category=$catObj
            #Settings=$catSettings
            RootCategory=$rootCatObj
        }))
    }

    $script:curSettingsCatologPolicy = @()

    $cfgSettings | % { Add-SettingsSetting $_.settingInstance $_.settingDefinitions } | Out-Null

    #$script:objectSettingsData = $script:curSettingsCatologPolicy
    
    foreach($item in ($script:curSettingsCatologPolicy | Select @{l="CategoryID";e={$_.CategoryDefinition.Id}}, @{l="SubCategoryID";e={$_.SubCategoryDefinition.Id}} -Unique))
    {
        $script:objectSettingsData += ($script:curSettingsCatologPolicy | Where { $_.CategoryDefinition.Id -eq $item.CategoryID -and $_.SubCategoryDefinition.Id -eq $item.SubCategoryID })   
    }
    
    if($docProvider.PostSettingsCatalog)
    {
        & $docProvider.PostSettingsCatalog $obj $objectType $script:objectSettingsData
    }
}

function Add-SettingsSetting
{
    param($settingInstance, $settingsDefs, $ItemLevel, [switch]$SkippAdd)

    $defaultValue = $null
    $tableValue = $null
    $value = $null
    $rawValue = $null
    $rawJsonValue = $null
    $show = $true
    $children = @()
    $childSettings = @()

    $settingsDef = $settingsDefs | Where id -eq $settingInstance.settingDefinitionId
    if(-not $settingsDef -and $settingInstance.settingDefinitionId) 
    {
        if($global:cachedCfgSettings.ContainsKey($settingInstance.settingDefinitionId) -eq $false) 
        {
            $settingsDef = Invoke-GraphRequest "/deviceManagement/configurationSettings/$($settingInstance.settingDefinitionId)"
            $global:cachedCfgSettings.Add($settingInstance.settingDefinitionId, $settingsDef)
        }
        else
        {
            $settingsDef = $global:cachedCfgSettings[$settingInstance.settingDefinitionId]
        }        
    }
    $categoryDef = $global:cfgCategories | Where Id -eq $settingsDef.categoryId #$script:settingCatalogasCategories[$settingsDef.categoryId]

    if($settingsDef.categoryId -ne $categoryDef.rootCategoryId)
    {
        $objCategory = $global:cfgCategories | Where Id -eq $categoryDef.rootCategoryId
        #$objCategory = $script:settingCatalogasCategories[$categoryDef.RootCategory.Id]
        $subCategory = $categoryDef
    }
    else
    {
        $subCategory = $null
        $objCategory = $categoryDef
    }    

    $settingInfo = [PSCustomObject]@{
        SettingId = $settingsDef.Id
        SettingKey = ""
        SettingName = $settingsDef.Name
        Name = $settingsDef.displayName
        Description=$settingsDef.description
        CategortyId = $objCategory.id
        Category=$objCategory.displayName
        CategoryDefinition=$objCategory
        SubCategory=$subCategory.displayName
        SubCategoryDefinition=$subCategory       
        Value = $null
        RawValue = $null
        RawJsonValue=$null
        TableValue = $null
        DefaultValue = $null
        Level = $ItemLevel
        Parent = $null
        Show = $show
        Type = $settingInstance.'@odata.type'
        PropertyIndex = 0
        RowIndex = 0
        ChildSettings = @() #($childSettings | Sort DisplayName)
    }

    $script:curSettingsCatologPolicy += $settingInfo

    if($settingInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance')
    { # Drop down and select one value
            
        $rawValue = $settingInstance.choiceSettingValue.value
        $value = ($settingsDef.Options | Where itemId -eq $rawValue).displayName

        if($settingsDef.defaultOptionId)
        {
            $defaultValue = ($settingsDef.Options | Where itemId -eq $settingsDef.defaultOptionId).displayName
        }

        foreach($childSetting in $settingInstance.choiceSettingValue.children)
        {
            $tmpSetting = Add-SettingsSetting $childSetting $settingsDefs ($ItemLevel + 1) -SkippAdd #-SkippAdd:$SkippAdd
            if($tmpSetting)
            {
                $tmpSetting.Parent = $settingInfo
                $settingInfo.ChildSettings += $tmpSetting
            }
        }
    }
    elseif($settingInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance')
    { # Simple value eg a string
        $value = $settingInstance.simpleSettingValue.value 
        $rawValue = $value
        if($settingsDef.defaultValue.value)
        {
            $defaultValue = $settingsDef.defaultValue.value
        }
    }
    elseif($settingInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationChoiceSettingCollectionInstance')
    { # Drop down and select one or more values
        $itemValues = @()
        $itemRawValues = @()
        $tableValue = [ordered]@{}

        foreach($colObj in $settingInstance.choiceSettingCollectionValue)
        {
            $itemRawValues += $colObj.value
            $itemValues += ($settingsDef.Options | Where itemId -eq $colObj.Value).displayName
        }

        $value = $itemValues -join $script:propertySeparator
        $rawValue = $itemRawValues -join $script:propertySeparator
        $rawJsonValue = $settingInstance.choiceSettingCollectionValue | ConvertTo-Json -Depth 50 -Compress

        if($settingsDef.defaultOptionId)
        {
            $defaultValue = ($settingsDef.Options | Where itemId -eq $settingsDef.defaultOptionId).displayName
        }                
    }
    elseif($settingInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance')
    { # Multiple settings under one group. Group will not be displayed

        $settingInfo.Show = $false
        $index=1
        foreach($groupSettingCollection in $settingInstance.groupSettingCollectionValue)
        {
            $childSettingsArr = @()
            # Not sure if this is the best way but it looks better for tested policies
            if($script:currentObject.templateReference.templateId -and $settingsDefs)
            {
                $childIDs = $settingsDefs.id # Endpoint Security objects
            }
            else
            {
                $childIDs = $settingsDef.childIds # Setings Catalog and from file documentation
            }
            #foreach($childId in $settingsDefs.id) #$settingsDef.childIds)
            foreach($childId in $childIDs)            
            {            
                $childSetting = ($groupSettingCollection.children | Where settingDefinitionId -eq $childId)
                if(-not $childSetting) { continue } #Not configured
                $tmpSetting = Add-SettingsSetting $childSetting $settingsDefs ($ItemLevel + 1) -SkippAdd #-SkippAdd:$SkippAdd
                if($tmpSetting)
                {
                    $tmpSetting.Parent = $childSettings
                    $tmpSetting.RowIndex = $index
                    $childSettings += $tmpSetting
                    $childSettingsArr += $tmpSetting
                    if($settingsDef.childIds.Count -gt 1)
                    {
                        $tmpSetting.PropertyIndex = $childSettingsArr.Count
                    }
                }
                $rowIndex++
            }

            $settingInfo.ChildSettings += [PSCustomObject]@{
                Id=$index++
                Type=$groupSettingCollection.'@odata.type'
                Settings=$childSettingsArr
            }
        }

        $rawJsonValue = $settingInstance.groupSettingCollectionValue | ConvertTo-Json -Depth 50 -Compress
    }
    elseif($settingInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationSimpleSettingCollectionInstance')
    { # Group of simple values
        $itemValues = @()
        foreach($colObj in $settingInstance.simpleSettingCollectionValue)
        {
            $itemValues += $colObj.value            
        }

        if($settingsDef.defaultValue.value)
        {
            $defaultValue = $settingsDef.defaultValue.value
        }

        $value = $itemValues -join $script:propertySeparator
        $rawValue = $itemValues -join $script:propertySeparator
        $rawJsonValue = $settingInstance.simpleSettingCollectionValue | ConvertTo-Json -Depth 50 -Compress
    }
    elseif($settingInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationGroupSettingInstance')
    {  # This will skip adding the group itself as enabled...
       # It will only add information about the children

        $show = $false
        foreach($groupSettingValue in $settingInstance.groupSettingValue)
        {
            #$children += $groupSettingValue #.children
            foreach($childSetting in $groupSettingValue.children)
            {
                $tmpSetting = Add-SettingsSetting $childSetting $settingsDefs ($ItemLevel + 1) -SkippAdd #-SkippAdd:$SkippAdd
                if($tmpSetting)
                {
                    $tmpSetting.Parent = $settingInfo
                    $settingInfo.ChildSettings += $tmpSetting
                }
            }            
        }
        $rawJsonValue = $settingInstance.groupSettingValue | ConvertTo-Json -Depth 50 -Compress
    }
    else
    {
        Write-Log "Unhandled object type: $($settingInstance.'@odata.type')" 2
        return
    }

    if(!$rawJsonValue -and $rawValue)
    {
        $rawJsonValue = $rawValue | ConvertTo-Json -Depth 50 -Compress
    }

    $settingInfo.Value = $value
    $settingInfo.RawValue = $rawValue
    $settingInfo.RawJsonValue=$rawJsonValue
    $settingInfo.DefaultValue = $defaultValue

    $settingInfo
}

#endregion

#region Intent Objects (Endpoint Security)

function Get-IntentCategory
{
    param($templateType)

    if(-not $templateType)
    {
        Write-Log "Get-IntentCategory called with empty Category" 2
        return
    }

    if($templateType.StartsWith("endpointSecurity"))
    {
        $templateType = $templateType.Substring(16)
    }

    if($templateType -eq "accountProtection")
    {
        return (Get-LanguageString "SecurityTemplate.accountProtection")
    }
    elseif($templateType -eq "antivirus")
    {
        return (Get-LanguageString "SecurityTemplate.antivirus")
    }
    elseif($templateType -eq "diskEncryption")
    {
        return (Get-LanguageString "SecurityTemplate.diskEncryption")
    }
    elseif($templateType -eq "endpointDetectionReponse")
    {
        return (Get-LanguageString "SecurityTemplate.eDR")
    }    
    elseif($templateType -eq "attackSurfaceReduction")
    {
        return (Get-LanguageString "SecurityTemplate.aSR")
    }
    elseif($templateType -eq "attackSurfaceReduction")
    {
        return (Get-LanguageString "SecurityTemplate.aSR")
    }
    elseif($templateType -eq "firewall")
    {
        return (Get-LanguageString "SecurityTemplate.firewall")
    }
    elseif($templateType -eq "securityBaseline" -or 
        $templateType -eq "baseline" -or
        $templateType -eq "advancedThreatProtectionSecurityBaseline" -or
        $templateType -eq "microsoftEdgeSecurityBaseline")
    {
        return (Get-LanguageString "Titles.securityBaselines")
    }
    else
    {
        Write-Log "Could not translate templateSubtype $templateType"
        return $templateType
    }
}

function Invoke-TranslateIntentObject
{
    param($obj, $objectType)

    $params = @{}
    if($script:DocumentationLanguage)
    {
        $params.Add("AdditionalHeaders", @{"Accept-Language"=$script:DocumentationLanguage})
    }

    Add-BasicDefaultValues $obj $objectType
    Add-BasicAdditionalValues $obj $objectType

    if($global:intentCategories -is [HashTable])
    {
        $categories = $global:intentCategories[$obj.TemplateId]
    }
    else
    {
        $global:intentCategories = @{}
        $global:catRecommendedSettings = @{}
    }

    if($global:intentCategoryDefs -isnot [HashTable])
    {
        $global:intentCategoryDefs = @{}
    }    

    if(-not $categories)
    {
        $categories = (Invoke-GraphRequest "/deviceManagement/templates/$($obj.TemplateId)/categories?`$expand=settingDefinitions" -ODataMetadata "minimal" @params).Value
        $global:intentCategories.Add($obj.TemplateId, $categories)
    }
    
    $script:objectSettings = @()

    foreach($category in ($categories | Sort -Property displayName))
    {
        # Get settings for the category. This will put them in the correct order...
        if($obj.'@ObjectFromFile' -ne $true)
        {
            $settings = (Invoke-GraphRequest "/deviceManagement/intents/$($obj.Id)/categories/$($category.Id)/settings?`$expand=Microsoft.Graph.DeviceManagementComplexSettingInstance/Value" -ODataMetadata "minimal" @params).Value
        }
        else 
        {
            $settings = $obj.settings
        }

        if($global:catRecommendedSettings.ContainsKey($category.Id))
        {
            $catRecommendedSettingsObj = $global:catRecommendedSettings[$obj.TemplateId]
        }
        else
        {
            # Get setting defenitions for the category
            #$catDefObjs = (Invoke-GraphRequest "/devicemanagement/templates/$($obj.TemplateId)/categories/$($category.Id)/settingDefinitions" -ODataMetadata "minimal").Value
            #$global:intentCategoryDefs.Add($category.Id,$catDefObjs)
            # Get recommended settings for the category
            $catRecommendedSettingsObj = (Invoke-GraphRequest "/deviceManagement/templates/$($obj.TemplateId)/categories/$($category.Id)/RecommendedSettings?" -ODataMetadata "minimal" @params).Value
            $global:catRecommendedSettings.Add($category.Id, $catRecommendedSettingsObj)
        }

        $catDefIds = ($category.settingDefinitions | Select Id).Id        

        foreach($settingObj in $settings) #($settings | Where { $_.definitionId -in $catDefIds }))
        {
            Get-IntentSettingInfo $settingObj $category $settingObj.definitionId $settings
        }
    }
    
    foreach($objSetting in ($script:objectSettings | Where { $_.ParentId -eq $null -and ($_.Dependencies | measure).Count -eq 0 }))
    {
        Add-IntentSettingObjectToList $objSetting
    }
}

function Add-IntentSettingObjectToList
{
    param($objSetting)

    if(($script:objectSettingsData | Where Id -eq $objSetting.Id)) { return }

    $passConstraint = $true
    $hasConstraint = $false
    foreach($dependencyObj in $objSetting.SettingDefinition.dependencies)
    {
        $dependencyItemObj = ($script:objectSettings | Where { $_.SettingDefinition.Id -eq $dependencyObj.definitionId })
        if($dependencyObj.constraints.Count -gt 0)
        {                    
            $hasConstraint = $true
            foreach($constraint in $dependencyObj.constraints)
            {
                if($constraint.'@odata.type' -eq "#microsoft.graph.deviceManagementSettingBooleanConstraint")
                {
                    if($dependencyItemObj.RawValue -eq $null -and $constraint.value -eq $false -or
                        ($dependencyItemObj.RawValue -and $dependencyItemObj.RawValue.ToString() -ne $constraint.value.ToString()))
                    {
                        $passConstraint = $false
                        break
                    }
                }
                elseif($constraint.'@odata.type' -eq "#microsoft.graph.deviceManagementEnumConstraint")
                {
                    if(-not ($constraint.values | Where Value -eq $dependencyItemObj.RawValue))
                    {
                        $passConstraint = $false
                        break
                    }
                }
                elseif($constraint.'@odata.type' -eq "#microsoft.graph.deviceManagementSettingIntegerConstraint")
                {
                    if($dependencyItemObj.RawValue -ge $constraint.minimumValue -and $dependencyItemObj.RawValue -le $constraint.maximumValue)
                    {
                        $passConstraint = $false
                        break
                    }
                }
            }
        }
        else
        {
            $passConstraint = ($dependencyItemObj.RawValue -ne $null -and $dependencyItemObj.RawValue.ToString() -ne "NotConfigured" -and $dependencyItemObj.RawValue.ToString() -ne "False")
        }
        if(-not $passConstraint) { break ]}
    }

    if(-not $passConstraint)
    {
        return
    }
    
    if($hasConstraint)
    {
        $objSetting.Level = $objSetting.Level + 1
    }

    $recommendedSetting = $global:catRecommendedSettings[$objSetting.CategoryObject.Id] | Where definitionId -eq  $objSetting.SettingId

    if($recommendedSetting.valueJson -and ($objSetting.ValueSet -eq $false -or $recommendedSetting.valueJson -ne ($objSetting.RawValue | ConvertTo-Json -Compress))) {
        $objSetting | Add-Member Noteproperty -Name "RecommendedValue" -Value ($recommendedSetting.valueJson | ConvertFrom-Json) -Force
    }    

    $script:objectSettingsData += $objSetting

    if($objSetting.ValueSet -eq $false) { return }

    foreach($depObj in ($script:objectSettings | Where { $_.Dependencies.definitionId -eq $objSetting.SettingDefinition.Id }))
    {
        Add-IntentSettingObjectToList $depObj
    }

    foreach($depObj in ($script:objectSettings | Where { $_.ParentId -eq $objSetting.Id -and ($_.Dependencies | measure).Count -eq 0} ))
    {
        Add-IntentSettingObjectToList $depObj    
    }
}

function Get-IntentSettingInfo
{
    param($valueObj, $category, $defId, $allSettings, [switch]$SkipConvertValue, [switch]$PassThru, $parentDef = $null)

    $defObj = $category.settingDefinitions | Where id -eq $defId    

    if(-not $defObj) { return } # Should never happen!

    $itemValue = $null
    $itemFullValue = $null

    if($SkipConvertValue -ne $true)
    {
        $rawValue = $valueObj.valueJson | ConvertFrom-Json
    }
    else
    {
        $rawValue = $valueObj
    }

    $valueSet = Get-IsIntentObjectConfigured $rawValue

    if($valueSet -eq $false)
    {
        ; # Skip child settings
    }
    elseif($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementCollectionSettingInstance' -or
            $defObj.'@odata.type' -eq '#microsoft.graph.deviceManagementComplexSettingDefinition' -or
            $defObj."valueType" -eq  "collection")
    {
        $valueArr = @()
        
        if($defObj.elementDefinitionId)
        {
            # Get the element defenition of the items in the collection
            $elementDefObj = $category.settingDefinitions | Where id -eq $defObj.elementDefinitionId
        }
        else
        {
            $elementDefObj = $defObj
        }

        if($elementDefObj.propertyDefinitionIds)
        {
            # Elements are Complex and based of one or more definitions
            $itemFullValue = @()
            foreach($tmpValue in $rawValue)
            {
                $htFullPropInfo = [ordered]@{}
                foreach($tmpPropDefId in $propDefObj.propertyDefinitionIds)
                {
                    $tmpDef = $category.settingDefinitions | Where id -eq $tmpPropDefId
                    $tmpProp = $tmpPropDefId.Split('_')[-1]
                    $htFullPropInfo.Add((?? $tmpDef.displayName $tmpProp), $tmpPropValue.$tmpProp)
                }                

                $arrValue = ""
                foreach($propertyDefinitionId in $elementDefObj.propertyDefinitionIds)
                {
                    $propDefObj = $category.settingDefinitions | Where id -eq $propertyDefinitionId
                    if($propDefObj.elementDefinitionId)
                    {
                        # The collection element can be based of multiple definitions with specific element types
                        $propDefObj = $category.settingDefinitions | Where id -eq $propDefObj.elementDefinitionId
                    }
                    if($arrValue) { $arrValue = ($arrValue + $script:propertySeparator)}
                    $propName = $propertyDefinitionId.Split('_')[-1]
                    $propValue = @()
                    foreach($childTmpVales in $tmpValue.$propName)
                    {
                        $propValue += Get-IntentObjectValue $propDefObj $childTmpVales
                    }

                    $htFullPropInfo.Add((?? $propDefObj.displayName $propName), $tmpValue.$propName)                    
<#
                    $propFullValue = @()
                    foreach($tmpPropValue in $propValue)
                    {
                        $htFullPropInfo = [ordered]@{}
                        foreach($tmpPropDefId in $propDefObj.propertyDefinitionIds)
                        {
                            $tmpDef = $category.settingDefinitions | Where id -eq $tmpPropDefId
                            $tmpProp = $tmpPropDefId.Split('_')[-1]
                            $htFullPropInfo.Add((?? $tmpDef.displayName $tmpProp), $tmpPropValue.$tmpProp)
                        }
                        $propFullValue += [PSCustomObject]$htFullPropInfo
                    }
#>
                    $arrValue = ($arrValue + ($propValue -join $script:propertySeparator))
                }
                $itemFullValue += [PSCustomObject]$htFullPropInfo
                $valueArr += $arrValue
            }                        
        }
        elseif($rawValue)
        {
            # Elements are Strings, Integers etc.
            foreach($tmpValue in $rawValue)
            {
                $valueArr += (Get-IntentObjectValue $elementDefObj $tmpValue)
            }
        }

        if($valueArr.Count -gt 0)
        {
            $itemValue = $valueArr -join $script:objectSeparator # Or `n or?
        }
        $valueSet = $valueArr.Count -gt 0
    }
    elseif($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementAbstractComplexSettingInstance' -or
        $defObj.'@odata.type' -eq '#microsoft.graph.deviceManagementAbstractComplexSettingDefinition')
    {
        $tmpDef = $category.settingDefinitions | Where { $_.id -eq $rawValue.implementationId -or $_.id -eq $rawValue.'$implementationId' }
        if($tmpDef)
        {
            $itemValue = $tmpDef.displayName
        }
        else
        {
            $valueSet = $false
        }
    }
    else
    {
        $itemValue = Get-IntentObjectValue $defObj $rawValue
        if(-not $itemValue) 
        {
            $valueSet = $false
        }
    }

    if($valueSet -eq $false)
    {
        $itemValue = Get-LanguageString "SettingDetails.notConfigured"
    }
    elseif(-not $itemValue)
    {
        $itemValue = $rawValue
    }

    $curObjectInfo = New-Object PSObject -Property @{ 
        Name=$defObj.displayName
        Description=$defObj.description
        Category=$category.displayName
        CategoryDescription=$category.description # Will not have a description
        CategoryObject=$category
        Value=$itemValue
        FullValueTable=$itemFullValue
        RawValue=$rawValue
        SettingDefinition = $defObj
        Dependencies = $defObj.dependencies
        ValueSet = $valueSet
        Id=[Guid]::NewGuid() #(([Guid]::NewGuid()).Guid)
        ParentId = $null
        SettingId = $defObj.Id # ToDo: Must have parent Id as well e.g. Firewall settings in Win10 baseline
        ParentSettingId = $parentDef.Id
        Level = 0
    }

    $script:objectSettings += $curObjectInfo    

    if($valueSet -eq $false)
    {
        ; # Skip children if value is not set...
    }
    elseif($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementComplexSettingInstance' -or 
        $defObj.'@odata.type' -eq '#microsoft.graph.deviceManagementComplexSettingDefinition')
    {        
        if($valueObj.Value)
        {
            $isValueSet = $false
            $sortByDefinitionIds = $true # !!!
            if($sortByDefinitionIds -eq $true -and $defObj.propertyDefinitionIds)
            {
                foreach($childDefId in $defObj.propertyDefinitionIds)
                {
                    $childSetting = $valueObj.Value | Where DefinitionId -eq $childDefId
                    if($childSetting)
                    {
                        $objValueInfo = Get-IntentSettingInfo $childSetting $category $childSetting.definitionId $allSettings -PassThru -parentDef $defObj
                        $objValueInfo.ParentId = $curObjectInfo.Id
                        if($objValueInfo.RawValue -is [Boolean] -and $objValueInfo.RawValue -eq $true)
                        {
                            $isValueSet = $true
                        }
                        elseif($objValueInfo.RawValue -is [String] -and -not [String]::IsNullOrEmpty($objValueInfo.RawValue) -and $objValueInfo.RawValue -ne "notConfigured" -and -not [String]::IsNullOrEmpty($objValueInfo.Value))
                        {
                            $isValueSet = $true
                        }
                        elseif($objValueInfo.RawValue -isnot [Boolean] -and $objValueInfo.RawValue -isnot [String])
                        {
                            $isValueSet = $true
                        }
                    }
                }
            }
            else
            {
                foreach($childSetting in $valueObj.Value)
                {
                    $objValueInfo = Get-IntentSettingInfo $childSetting $category $childSetting.definitionId $allSettings -PassThru -parentDef $defObj
                    $objValueInfo.ParentId = $curObjectInfo.Id
                }
            }
        }
        elseif($rawValue -and $defObj.propertyDefinitionIds)
        {
            $isValueSet = $false
            $isDefault = $true
            foreach($childDefId in $defObj.propertyDefinitionIds)
            {
                $propName = $childDefId.Split('_')[-1]                
                $objValueInfo = Get-IntentSettingInfo $rawValue.$propName $category $childDefId $allSettings -SkipConvertValue -PassThru -parentDef $defObj
                if($objValueInfo.ValueSet -eq $true) { $isValueSet = $true }
                
                if($objValueInfo.SettingDefinition.constraints -and  
                    $objValueInfo.SettingDefinition.constraints[0].'@odata.type' -eq "#microsoft.graph.deviceManagementEnumConstraint" -and 
                    ($objValueInfo.SettingDefinition.constraints[0].values | measure).Count -gt 0)
                {
                    if($objValueInfo.SettingDefinition.constraints[0].values[0].value -ne $rawValue.$propName)
                    {
                        $isDefault = $false
                    }
                }
                elseif($objValueInfo.SettingDefinition.valueType -eq "string")
                {
                    if($null -ne $rawValue.$propName)
                    {
                        $isDefault = $false
                    }
                }
                elseif($objValueInfo.SettingDefinition.valueType -eq "boolean")
                {
                    if($false -ne $rawValue.$propName)
                    {
                        $isDefault = $false
                    }
                }                
                $objValueInfo.ParentId = $curObjectInfo.Id                
            }
            if($isDefault)
            {
                # All child items are using default settings so set value not set
                # ToDo: Verify ALL cases...don't like this but it looks like this is how it's done
                $isValueSet = $false
            }
        }
        else
        {
            $isValueSet = $false
        }

        $curObjectInfo.Value = ?: $isValueSet "Configure" (Get-LanguageString "SettingDetails.notConfigured")
        $curObjectInfo.ValueSet = $isValueSet
        $curObjectInfo.FullValueTable = $null
    }
    elseif(($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementAbstractComplexSettingInstance' -or 
            $defObj.'@odata.type' -eq '#microsoft.graph.deviceManagementAbstractComplexSettingDefinition') -and
            $rawValue -and $tmpDef)
    {
        foreach($childDefId in $tmpDef.propertyDefinitionIds)
        {
            $propName = $childDefId.Split('_')[-1]

            $objValueInfo = Get-IntentSettingInfo $rawValue.$propName $category $childDefId $allSettings -SkipConvertValue -PassThru -parentDef $defObj
            $objValueInfo.ParentId = $curObjectInfo.Id
        }
    }

    if($PassThru)
    {
        $curObjectInfo
    }
}

function Get-IntentObjectValue
{
    param($defObj, $rawValue)

    $itemValue = $null

    if($defObj.constraints.'@odata.type' -eq "#microsoft.graph.deviceManagementEnumConstraint")
    {
        $tmpOption = $defObj.constraints.Values  | Where value -eq $rawValue
        if(-not $tmpOption -and $rawValue -eq $null)
        {
            # This is NOT ok. There is no defaultValue of the setting definitions
            # Ex when this is used:
            # deviceConfiguration--windows10EndpointProtectionConfiguration_defenderScanDirection
            $tmpOption =  $defObj.constraints.Values[0]
        }
        $itemValue = $tmpOption.displayName
    }
    elseif($defObj.valueType -eq "boolean")
    {
        if($rawValue -eq "True")
        {
            $itemValue = Get-LanguageString "SettingDetails.yes"
        }
    }
    else
    {
        $itemValue = $rawValue
    }

    $itemValue
}

function Get-IsIntentObjectConfigured
{
    param($obj)

    # Custom checks if needed

    return $true
}

#endregion

#region Profile Types e.g. DeviceConfiguration Profiles, Compliance Policies etc
function Invoke-TranslateProfileObject
{
    param($obj, $objectType)

    $objInfo = Get-TranslationFiles $obj.'@OData.Type'
    
    if(-not $objInfo)
    {        
        return $false
    }
    elseif(($objInfo | measure).Count -gt 1)
    {
        Write-Log "Multiple category types returned for $($obj.'@OData.Type'). Aborting..." 3
        return $false
    }

    Add-BasicDefaultValues $obj $objectType
    Add-BasicAdditionalValues $obj $objectType

    $allFiles = @()
    if($objInfo.Categories)
    {
        foreach($objCategory in ($objInfo.Categories))
        {
            $file = ($global:AppRootFolder + "\Documentation\ObjectInfo\$($objCategory)_$($objInfo.PolicyType).json")
            $fi = [IO.FileInfo]$file
            if($fi.Exists -eq $false)
            {
                Write-Log "Category file '$file' not found"
                continue
            }
            $allFiles += $fi 
        }
    }
    else
    {
        # Should only be one file. Compliance policies might have more
        $files = [IO.Directory]::EnumerateFiles($global:AppRootFolder + "\Documentation\ObjectInfo", "*_$($objInfo.PolicyType).json")
        if(($files | measure).Count -eq 0)
        {
            Write-Log "No category files returned for $($objInfo.PolicyType)" 2
        }
        elseif(($files | measure).Count -gt 1)
        {
            Write-Log "Multiple category files returned for $($objInfo.PolicyType)" 2
        }

        foreach($file in $files)
        {            
            $fi = [IO.FileInfo]$file
            $allFiles += $fi 
        }
    }

    Add-CustomProfileProperties $obj

    foreach($fi in $allFiles)
    {
        try 
        {
            $categoryObj = (Get-Content $fi.FullName -Encoding UTF8) | ConvertFrom-Json 
            $script:CurrentSubCategory = ""
            
            if($docProvider.TranslateSectionFile)
            {
                $retObj = & $docProvider.TranslateSectionFile $obj $objectType $fi $categoryObj
                if($retObj -is [Boolean] -and $retObj -eq $true)
                {
                    # Handled by custom function
                    continue
                }
            }            
            Invoke-TranslateSection $obj $categoryObj."$($fi.BaseName)" $objInfo            
        }
        catch 
        {
            Write-LogError "Failed to translate file $($fi.Name)" $_.Exception    
        }
    }

    return $true
}

function Get-LanguageString
{
    param($string, $defaultValue = $null, [switch]$IgnoreMissing)

    if(-not $script:languageStrings)
    {
        $lng = ?? $script:DocumentationLanguage "en"
        $fileContent = Get-Content ($global:AppRootFolder + "\Documentation\Strings-$($lng).json") -Encoding UTF8
        $script:languageStrings =  $fileContent | ConvertFrom-Json
    }

    if(!$string) { return }

    $arrParts = $string.Split('.')

    if([String]::IsNullOrEmpty($arrParts[-1])) { return }

    $obj = $script:languageStrings
    foreach($part in $arrParts.Split('.'))
    {
        if($obj.$part -or $part -eq "Empty")
        {
            $obj = $obj.$part
        }
        else
        {
            if($part -and $IgnoreMissing -ne $true)
            {
                Write-Host "Could not find string $string. Part '$part' was not found"
            }
            return $defaultValue
        }
    }
    return $obj.Trim("`n")
}

function Get-TranslationFiles
{
    param($type)

    if(-not $script:ObjectPolicyTypes)
    {
        $script:ObjectPolicyTypes = Get-Content ($global:AppRootFolder + "\Documentation\ObjectCategories.json") | ConvertFrom-Json    
    }

    return $script:ObjectPolicyTypes | Where ObjectType -eq $type
}

function Get-Category
{
    param($categoryId)

    if($categoryId -is [String])
    {
        return Get-LanguageString (?: $categoryId.Contains(".") $categoryId "Category.$($categoryId)")
    }

    if(-not $script:Categories)
    {
        $script:CategoryIds = Get-Content ($global:AppRootFolder + "\Documentation\CategoryId.json") -Encoding UTF8 | ConvertFrom-Json
    }
    
    Get-LanguageString "Category.$($script:CategoryIds."$($categoryId)")"
}

function Invoke-VerifyCondition
{
    param($obj, $prop, $objInfo)

    if(!$prop.Condition -or ($prop.Condition.Expressions | measure).Count -eq 0) { return $true}

    # Default condition type is or
    $type = ?: ($prop.Condition.type -eq "and") "and" "or"

    $defaultReturn = ?: ($type -eq "and") $true $false
    foreach($expression in $prop.Condition.Expressions)
    {
        if(!$expression.property) { continue }

        $tmpProp = $obj.PSObject.Properties | Where Name -eq $expression.property

        if(!$tmpProp)
        {
            if($expression.ignoreMissing -eq $true) { continue }
            return $false
        }

        if($expression.operator -eq "null")
        {            
            $tmpRet = $null -eq $tmpProp.Value
        }
        elseif($null -eq $expression.value)
        {
            # Value not specified. Check if the property is set
            $tmpRet = $null -ne $tmpProp.Value 
        }
        elseif($expression.operator -eq "ne")
        {
            $tmpRet = $obj."$($expression.property)" -ne $expression.value
        }
        elseif($expression.operator -eq "gt")
        {
            $tmpRet = $obj."$($expression.property)" -gt $expression.value
        }
        elseif($expression.operator -eq "ge")
        {
            $tmpRet = $obj."$($expression.property)" -ge $expression.value
        }
        elseif($expression.operator -eq "lt")
        {
            $tmpRet = $obj."$($expression.property)" -lt $expression.value
        }
        elseif($expression.operator -eq "le")
        {
            $tmpRet = $obj."$($expression.property)" -le $expression.value
        }
        elseif($expression.operator -eq "like")
        {
            $tmpRet = $obj."$($expression.property)" -like $expression.value
        }
        elseif($expression.operator -eq "notlike")
        {
            $tmpRet = $obj."$($expression.property)" -notlike $expression.value
        }
        else
        {
            # Default operator is eq
            $tmpRet = $obj."$($expression.property)" -eq $expression.value
        }
        
        if($tmpRet -eq $true -and $type -eq "or") { return $true }
        if($tmpRet -eq $false -and $type -eq "and") { return $false }
    }
    return $defaultReturn
}

function Invoke-TranslateSection
{
    param($obj, $sectionObject, $objInfo, $parent = $null)

    if($null -eq $parent -or $script:propLevel -lt 0)
    {
        $script:propLevel = 0
    }
    elseif($parent -ne $script:currentParent)# -and $parent.skipAddLevel -ne $true)
    {
        $script:propLevel++
    }

    foreach($prop in $sectionObject)
    {
        $valueData = $null
        $value = $null
        $valueSet = $false
        $useParentProp = $false
        
        #if($prop.enabled -eq $false -and $objInfo.ShowDisabled -ne $true) { continue }

        if((Invoke-VerifyCondition $obj $prop $objInfo) -eq $false) 
        {
            Write-LogDebug "Condition returned false: $(($prop.Condition | ConvertTo-Json -Depth 50 -Compress))" 2
            continue
        }

        $obj = Get-CustomPropertyObject $obj $prop

        $rawValue = $obj."$($prop.entityKey)"

        if($prop.dataType -eq 8)
        {
            if($prop.nameResourceKey -eq "LearnMore") { continue }
            elseif($prop.nameResourceKey -eq "Empty") { $script:CurrentSubCategory = $null }
            elseif($prop.nameResourceKey -in $script:categoriesToIgnore) { continue }
            elseif($prop.nameResourceKey)
            {
                $tmpStr = (Get-LanguageString (?: $prop.nameResourceKey.Contains(".") $prop.nameResourceKey "SettingDetails.$($prop.nameResourceKey)"))
                if($tmpStr -and $tmpStr.Length -lt 75) # !!! categoriesToIgnore will take care of some
                {
                    $script:CurrentSubCategory = $tmpStr
                }
                elseif($tmpStr)
                {
                    Write-LogDebug "SubCategpry ignored based on length: $tmpStr" 2
                }
            }
            $script:propLevel = -1
            Invoke-ChildSections $obj $prop
            continue 
        }
        elseif($prop.dataType -eq 5) # complex Options
        {
            # Skip if disabled
            if($prop.enabled -eq $false -and $objInfo.ShowDisabled -ne $true) { continue }

            # Set sub-category to nameResourceKey value if EntityKey is empty?
            if(-not $prop.EntityKey -and $prop.nameResourceKey)
            {
                $script:propLevel = -1
                $script:CurrentSubCategory = (Get-LanguageString (?: $prop.nameResourceKey.Contains(".") $prop.nameResourceKey "SettingDetails.$($prop.nameResourceKey)"))
            }
            else
            {
                $script:propLevel--
            }

            foreach($tmpObj in $obj)
            {
                Invoke-TranslateSection $tmpObj $prop.complexOptions $objInfo -Parent $prop
            }
            continue                    
        }
        elseif($prop.dataType -eq 6) # Complex option based on sub property
        {
            # Skip if disabled
            if($prop.enabled -eq $false -and $objInfo.ShowDisabled -ne $true) { continue }

            # Use sub property if EntityKey is specified
            $script:propLevel--
            $propObj = $null
            if($prop.entityKey)
            {                
                $propObj = $obj.PSObject.Properties | Where Name -eq $prop.entityKey
            }

            foreach($tmpObj in (?: ($propObj -ne $null) $rawValue $obj))
            {
                Invoke-TranslateSection $tmpObj $prop.complexOptions $objInfo -Parent $prop
            }
            continue
        }
        elseif($prop.dataType -eq 9) # Label. Skip the label but add children
        {
            $script:propLevel--
            Invoke-ChildSections $obj $prop
            continue
        }
        elseif($prop.dataType -eq 10) # Information box
        {            
            continue
        }        
        elseif($prop.dataType -eq 101) # Static label. Language string id in value property
        {
            if($prop.value)
            {
                $value = Get-LanguageString $prop.value

                Add-PropertyInfo $prop $value $rawValue $rawValue
            }
        }
        elseif($prop.dataType -eq 107) # Static value. String in value property
        {
            if($prop.value)
            {
                Add-PropertyInfo $prop $prop.value $prop.value $prop.value
            }
        }
        elseif([String]::IsNullOrEmpty($prop.entityKey) -eq $false)
        {            
            $valueSet = ($rawValue -ne $null)
            $skipChildren = $false
            if($rawValue -eq $null -and ![String]::IsNullOrEmpty($prop.unconfiguredValue) -and $global:chkSetUnconfiguredValue.IsChecked)
            {
                $propValue = $prop.unconfiguredValue
                Add-NotConfiguredProperty $prop
            }
            elseif($rawValue -eq $null -and ![String]::IsNullOrEmpty($prop.defaultValue) -and $global:chkSetDefaultValue.IsChecked)
            {
                $propValue = $prop.defaultValue
            }
            elseif($rawValue -eq $null -and ![String]::IsNullOrEmpty($prop.emptyValueResourceKey) -and $global:chkSetDefaultValue.IsChecked)
            {
                $propValue = Get-LanguageString $prop.emptyValueResourceKey
            }
            else
            {
                $propValue = $rawValue
            }      

            $addPropertyInfo = $true            
            $customValue = Get-CustomProfileValue $obj $prop
            if($customValue -is [Boolean] -and $customValue -eq $false)
            {
                continue # Property added by Custom info. Skip processing
            }
            elseif(-not $customValue)
            {
                # Some properties are added but remarked to let it fall through to not supported handler                
                
                # Do linked property BEFORE property detection
                # Linked properties are based on navigationLinks so they are not an actual property
                if($prop.dataType -eq 4) # Linked certificate 
                {                    
                    $useParentProp = $true
                    # Use $script:currentObject since $obj could be a property on the original object
                    # Cert links are always specified on the main object
                    $cert = $null
                    if($script:offlineDocumentation -ne $true)
                    {
                        $cert = Invoke-GraphRequest -URL $script:currentObject."$($prop.entityKey)@odata.navigationLink" -ODataMetadata "minimal" -NoError
                    }
                    if($cert)
                    {
                        if($cert.value -is [Object[]])
                        {
                            $certs = @()
                            $cert.value | ForEach-Object { $certs += $_.displayName }
                            if($certs.Count -gt 0)
                            {
                                $value = ($certs -join $script:objectSeparator)
                            }
                        }
                        elseif($cert.displayName)
                        {
                            $value = $cert.displayName
                        }
                        $rawValue = $value
                    }
                    elseif($script:currentObject.'@ObjectFromFile' -eq $true)
                    {
                        if($script:currentObject."#CustomRef_$($prop.entityKey)")
                        {
                            $idx = $script:currentObject."#CustomRef_$($prop.entityKey)".IndexOf("|:|")
                            if($idx -gt -1)
                            {
                                $value = $script:currentObject."#CustomRef_$($prop.entityKey)".SubString(0,$idx)
                            }
                            else
                            {
                                $value = $script:currentObject."#CustomRef_$($prop.entityKey)"
                            }
                        }
                        else
                        {
                            $value = "##TBD - Linked Certificate"
                        }
                        $rawValue = $value
                    }
                }
                elseif($prop.dataType -eq 200) # Multi option based on boolean value
                {
                    $value = Get-LanguageString $prop.entityKey
                }
                elseif(($prop.allowMissing -ne $true) -and
                    ($prop.entityKey -ne ".") -and 
                    (-not ($obj.PSObject.Properties | Where Name -eq $prop.entityKey)) -and 
                    (-not ($obj.PSObject.Properties | Where Name -eq "$($prop.entityKey)@odata.navigationLink")))
                {
                    if($prop.enabled -ne $false)
                    {
                        Write-Log "Property with EntityKey $($prop.entityKey) is missing. Property will not be added!" 2
                    }
                    else
                    {
                        Write-LogDebug "Disabled property with EntityKey $($prop.entityKey) is missing. Property will not be added!" 2
                    }
                    continue
                }
                elseif($prop.dataType -eq 0) # Boolean
                {
                    $value = Invoke-TranslateBoolean $obj $prop
                }
                elseif($prop.dataType -eq 1) # Base64 e.g. certificate data
                {                    
                    if($prop.filenameEntityKey -and $obj."$($prop.filenameEntityKey)")
                    {
                        $value = $obj."$($prop.filenameEntityKey)"
                    }
                    else
                    {
                        $value = $obj."$($prop.EntityKey)"
                        if($value)
                        {
                            try
                            {
                                # Is this always Base64 string?
                                $value = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($value))
                            }
                            catch { }
                        }                        
                    }                    
                }
                elseif($prop.dataType -eq 2) # Multiline string e.g. XML file 
                {
                    if($prop.filenameEntityKey -and $obj."$($prop.filenameEntityKey)")
                    {
                        $value = $obj."$($prop.filenameEntityKey)"
                    }
                    else
                    {
                        $value = $obj."$($prop.EntityKey)"
                        if($value)
                        {
                            try
                            {
                                # Is this always Base64 string?
                                $value = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($value))
                            }
                            catch { }
                        }   
                    }
                }
                elseif($prop.dataType -eq 3) # Image
                {
                    # Better text?
                    $valueData = $propValue 
                    $value = ?: $propValue "Image file" $null
                }                                
                elseif($prop.dataType -eq 7) # omaSettingDateTime
                {
                    $value = $propValue # Might require formatting 
                }
                elseif($prop.dataType -eq 9) # DataType
                {
                    continue # Ignore @odata.type labels
                }
                elseif($prop.dataType -eq 10) # Label
                {
                    continue # Ignore labels
                }
                elseif($prop.dataType -eq 11) # App Picker
                {
                    #$value = $propValue
                }
                elseif($prop.dataType -eq 12) # Multiline String? Sometimes and array and sometimes a binary
                {
                    if(($propValue | measure).Count -gt 0)
                    {
                        $value = $propValue -join $script:objectSeparator
                    }
                }
                elseif($prop.dataType -eq 13) # Multi option
                {
                    $value = Invoke-TranslateMultiOption $obj $prop
                }
                elseif($prop.dataType -eq 14) # Int32
                {
                    $value = $propValue
                }
                elseif($prop.dataType -eq 15) # Int64
                {
                    $value = $propValue
                }
                elseif($prop.dataType -eq 16) # Option
                {
                    Invoke-TranslateOption $obj $prop | Out-Null
                    continue
                }
                elseif($prop.dataType -eq 19) # Option
                {
                    Invoke-TranslateOption $obj $prop | Out-Null
                    continue
                }
                elseif($prop.dataType -eq 20) # String
                {
                    $value = $propValue
                }
                elseif($prop.dataType -eq 21) # Table
                {
                    $value = Invoke-TranslateTable $obj $prop
                    continue
                }
                elseif($prop.dataType -eq 22) # Scale value e.g. 4 Years (Certificate validity period)
                {
                    $value = $propValue
                    $scaleEntityKey = ?? $obj."$($prop.scaleEntityKey)" $prop.defaultScale
                    if($scaleEntityKey)
                    {
                        $scaleOption = $prop.scaleOptions | Where value -eq $scaleEntityKey
                        if($scaleOption.nameResourceKey)
                        {
                            $value = "{0} {1}" -f $propValue, (Get-LanguageString "SettingDetails.$($scaleOption.nameResourceKey)")
                        }
                    }
                }
                #elseif($prop.dataType -eq 25) {}# Home screen
                # All data types above 100 are custom
                elseif($prop.dataType -eq 100) # Custom Edm.Duration
                {
                    $value = Invoke-TranslateDuration $obj $prop                    
                }
                elseif($prop.dataType -eq 102) # Culture name
                {
                    $value = Get-CutureLanguageString (?? $propValue $prop.unconfiguredValue)
                }
                elseif($prop.dataType -eq 103) # Boolean action but hide children on false
                {
                    $value = Invoke-TranslateBoolean $obj $prop

                    $skipChildren = ($propValue -eq $false)
                }                
                elseif($prop.dataType -eq 104) # Multi option based on boolean value
                {
                    $value = Invoke-TranslateMultiOptionBoolean $obj $prop

                    $skipChildren = ($propValue -eq $false)
                }
                elseif($prop.dataType -eq 105) # Multi option based on boolean value but $false if the value is selected
                {
                    $value = Invoke-TranslateMultiOptionBoolean $obj $prop $false

                    $skipChildren = ($propValue -eq $false)
                }
                elseif($prop.dataType -eq 106) # Array of languages (Culture)
                {
                    $arrTmp = @()
                    foreach($tmpLng in $propValue)
                    {
                        $arrTmp += Get-CutureLanguageString $tmpLng
                    }
                    $value = $arrTmp -join $script:objectSeparator
                }                   
                elseif($prop.dataType -eq 108) # String with format
                {
                    $value = $propValue
                    if($prop.formatStringKey) {
                        $str = Get-LanguageString $prop.formatStringKey
                        if($str)
                        {
                            $value = $str -f $propValue
                        }
                    }
                }
                else
                {
                    Write-Log "Unsupported property '$((Get-LanguageString "SettingDetails.$($prop.nameResourceKey)"))' ($($prop.nameResourceKey)) for object property $($prop.entityKey). Type: $($prop.dataType)" 2
                    $value = $propValue
                }
            }
            else
            {
                $value = $customValue.Value
                $rawValue = $customValue.RawValue
                $valueSet = ($rawValue -ne $null)
                $addPropertyInfo = $customValue.AddPropertyInfo
            }

            if($addPropertyInfo)
            {
                Add-PropertyInfo (?: ($useParentProp -and $parent) $parent $prop) $value $rawValue 
            }
        }
        else
        {
            Write-Log "No property entity key: $($prop.dataType) ($($prop.nameResourceKey))"
        }

        if($valueSet -and $skipChildren -ne $true)
        {
            Invoke-ChildSections $obj $prop
        }
    }

    if($null -ne $parent -and $parent -ne $script:currentParent -and $script:propLevel -gt 0)
    {
        $script:propLevel--
    }
}

function Get-CutureLanguageString
{
    param($culture)

    if(!$culture)
    {
        return
    }

    try
    {
        if($culture -eq "os-default")
        {
            return Get-LanguageString "Autopilot.OOBE.useOSDefaultLanguage"
        }
        elseif($culture -eq "user-select")
        {
            return Get-LanguageString "Autopilot.OOBE.userSelect"
        }
        else
        {
            if(!$script:languageStrings) { Get-LanguageString } # This will load the language strings

            if($script:languageStrings.Languages.$culture)
            {
                return $script:languageStrings.Languages.$culture
            }
            $lngArr = $culture.Split('-')
            if($lngArr.Length -eq 3)
            {
                if($script:languageStrings.Languages."$(($lngArr[0]-$lngArr[1]))") 
                {
                    return $script:languageStrings.Languages."$(($lngArr[0]-$lngArr[1]))"
                }
            }

            if($lngArr.Length -gt 1 -and $script:languageStrings.Languages."$(($lngArr[0]))") 
            {
                return $script:languageStrings.Languages."$(($lngArr[0]))"
            }

            Write-Log "Translated language for $culture not found" 2
            ([cultureinfo]$culture).EnglishName
        }
    }
    catch{}
}

function Add-PropertyInfo
{
    param($prop, $value, $originalValue, $jsonValue, $tableValue)

    if($prop.Category -eq "1000")
    {
        Add-BasicPropertyValue (Get-LanguageString $prop.nameResourceKey) $value
        return
    }

    $script:objectSettingsData += Get-PropertyInfo $prop $value $originalValue $jsonValue $tableValue

    Invoke-CustomPostAddValue $prop
}

function Add-PropertyInfoObject
{
    param($propInfo)

    if($propInfo -eq $null) { return }

    if($prop.dataType -eq 99)
    {

    }

    $script:objectSettingsData += $propInfo
}

function Get-PropertyInfo
{
    param($prop,$value,$originalValue, $jsonValue, $tableValue)

    if($prop.nameResource)
    {
        $name = $prop.nameResource
    }
    else
    {
        $name = Get-LanguageString (?: $prop.nameResourceKey.Contains(".") $prop.nameResourceKey "SettingDetails.$($prop.nameResourceKey)")
    }

    $description = ""
    if($prop.descriptionResource)
    {
        $description = $prop.descriptionResource
    }
    elseif($prop.descriptionResourceKey)
    {
        $description = Get-LanguageString (?: $prop.descriptionResourceKey.Contains(".") $prop.descriptionResourceKey "SettingDetails.$($prop.descriptionResourceKey)")
    }

    $categoryStr = $null

    if($prop.category)
    {
        $categoryStr = Get-Category $prop.category
    }

    if(!$jsonValue -and $null -ne $rawValue -and "$($rawValue)" -ne "")
    {
        $jsonValue = $rawValue | ConvertTo-Json -Depth 50 -Compress
    }

    if($prop.emptyValueResourceKey)
    {
        $defValue = Get-LanguageString $prop.emptyValueResourceKey
    }
    else
    {    
        $defValue = $prop.defaultValue
    }

    return New-Object PSObject -Property @{ 
        Name=$name
        Description=$description
        Value=$value
        Category=$categoryStr
        SubCategory=$script:CurrentSubCategory
        Property=$prop.entityKey
        ValueSet=$valueSet
        DataType=$prop.dataType
        RawValue=$originalValue
        RawJsonValue=$jsonValue
        DefaultValue=$defValue
        FullValueTable = $tableValue 
        UnconfiguredValue = $prop.unconfiguredValue
        AlwaysAddValue = $prop.alwaysAddValue -eq $true
        Enabled=$prop.Enabled 
        EntityKey=$prop.EntityKey
        Level=$script:propLevel
    }
}

function Invoke-ChildSections
{
    param($obj, $sectionObject)

    $objTmp = Get-CustomChildObject $obj $sectionObject

    $tmpLevel = $script:propLevel
    Invoke-TranslateSection $objTmp $sectionObject.Children $objInfo -Parent $sectionObject
    
    $script:propLevel = $tmpLevel
    Invoke-TranslateSection $objTmp $sectionObject.ChildSettings $objInfo -Parent $sectionObject    
}

function Get-CustomProfileValue
{
    param($obj, $prop)

    foreach($docProvider in ($global:documentationProviders | Sort -Property Priority))
    {
        if($docProvider.GetCustomProfileValue)
        {
            $retObj = & $docProvider.GetCustomProfileValue $obj $prop $script:currentObject $script:propertySeparator $script:objectSeparator
            if($retObj -ne $null)
            {
                return $retObj
            }
        }
    }
}

function Invoke-CustomPostAddValue
{
    param($prop)

    foreach($docProvider in ($global:documentationProviders | Sort -Property Priority))
    {
        if($docProvider.PostAddValue)
        {
            $retObj = & $docProvider.PostAddValue $script:currentObject $prop $script:propertySeparator $script:objectSeparator
            if($retObj -ne $null)
            {
                return $retObj
            }
        }
    }
}

function Get-CustomPropertyObject
{
    param($obj, $prop)

    foreach($docProvider in ($global:documentationProviders | Sort -Property Priority))
    {
        if($docProvider.GetCustomPropertyObject)
        {
            # $script:currentObject is used to always send in the main object
            $retObj = & $docProvider.GetCustomPropertyObject $script:currentObject $prop
            if($retObj)
            {
                return $retObj
            }
        }
    }

    return $obj
}

function Get-CustomChildObject
{
    param($obj, $prop)

    foreach($docProvider in ($global:documentationProviders | Sort -Property Priority))
    {
        if($docProvider.GetCustomChildObject)
        {
            # $script:currentObject is used to always send in the main object
            $retObj = & $docProvider.GetCustomChildObject $script:currentObject $prop
            if($retObj)
            {
                return $retObj
            }
        }
    }

    return $obj
}

function Invoke-InitDocumentation
{
    foreach($docProvider in ($global:documentationProviders | Sort -Property Priority))
    {
        if($docProvider.InitializeDocumentation)
        {
            & $docProvider.InitializeDocumentation
        }
    }
}

function Add-CustomProfileProperties
{
    param($obj)

    foreach($docProvider in ($global:documentationProviders | Sort -Property Priority))
    {
        if($docProvider.AddCustomProfileProperty)
        {
            if((& $docProvider.AddCustomProfileProperty $obj $script:propertySeparator $script:objectSeparator))
            {
                return
            }
        }
    }
}

function Get-CustomIgnoredCategories 
{
    param($obj)

    $script:categoriesToIgnore = @(
            #microsoft.graph.windows10EndpointProtectionConfiguration
            "defenderSecurityCenterContactOptionsText",
            "globalConfigurationsDescription","generalNetworkSettingsHeader",
            "firewallCreateRules","exploitGuardCFHeadingText","exploitGuardNFTitle",
            "exploitGuardEPExplainationPart1","exploitGuardEPExplainationPart2","exploitGuardEPExplainationPart3","exploitGuardEPExplainationPart4",
            "defenderSecurityCenterSubHeaderText","defenderSecurityCenterITContactInformationSubHeaderText",
            "windows10EndpointProtectionDeviceGuardLearnMore",
            #microsoft.graph.windows10GeneralConfiguration
            "win10DefaultPrivacyHeader",
            "dfciBuiltinHeaderDescName"
            
            )
}

function Add-DefenderFirewallSettings
{
    param($fwSettings, $fwType)

    foreach($fwProp in ($fwSettings.PSObject.Properties | Where { $_.Name -like "*Blocked" -or $_.Name -like '*NotMerged' }).Name)
    {            
        if($fwProp -like "*Blocked")
        {
            $blockedValue = $fwSettings.$fwProp
            $propPre = $fwProp.SubString(0, $fwProp.Length - 7)
            if(($fwSettings.PSObject.Properties | Where { $_.Name -eq "$($propPre)Required" }))
            {
                $nonBlockedValue = $fwSettings."$($propPre)Required"
            }
            elseif(($fwSettings.PSObject.Properties | Where { $_.Name -eq "$($propPre)Allowed" }))
            {
                $nonBlockedValue = $fwSettings."$($propPre)Allowed"
            }
            else { continue }

            $fwPropName = "firewallSyntheticProfile$($fwType)$($propPre)"                
        }
        else
        {
            $blockedValue = $fwSettings.$fwProp
            $propPre = $fwProp.SubString(0, $fwProp.Length - 9)
            if(($fwSettings.PSObject.Properties | Where { $_.Name -eq "$($propPre)Merged" }))
            {
                $nonBlockedValue = $fwSettings."$($propPre)Merged"
            }
            else { continue }

            $fwPropName = "firewallSyntheticProfile$($fwType)$($propPre)Merge"
        }

        $fwValue =  "notConfigured"
        if($blockedValue -eq $true -and $nonBlockedValue -eq $false)
        { 
            $fwValue = "blocked"
        }  
        elseif($blockedValue -eq $false -and $nonBlockedValue -eq $true)
        { 
            $fwValue = "allowed"
        }
        $obj | Add-Member Noteproperty -Name $fwPropName -Value $fwValue -Force            
    }
}


function Get-ConditionalLaunchSetting
{
    param($obj, $lngId, $prop, $actionLngId, [switch]$SkipValue)

    if($null -eq $obj."$($prop)" -or ($obj."$($prop)" -is [String] -and $obj."$($prop)" -eq "notConfigured")) { return }
    elseif($obj."$($prop)" -is [String] -and $obj."$($prop)".StartsWith("P"))
    {
        $ts = Get-DurationValue $obj."$($prop)" -ReturnTimeSpan
        if($prop -eq "periodOfflineBeforeAccessCheck")
        {
            $settingValue = $ts.TotalMinutes
        }
        else
        {
            $settingValue = $ts.TotalDays
        }
    }
    elseif($SkipValue -ne $true)
    {
        $settingValue = $obj."$($prop)"
    }
    else
    {
        $settingValue = $null
    }

    [PSCustomObject]@{
        Setting=(Get-LanguageString (?: $lngId.Contains(".") $lngId "SettingDetails.$lngId"))
        Value=$settingValue
        Action=(Get-LanguageString (?: $actionLngId.Contains(".") $actionLngId "SettingDetails.$actionLngId"))
    }
}

function Invoke-TranslateBoolean
{
    param($obj, $prop)

    $propValue = (?? ($obj."$($prop.entityKey)") $prop.unconfiguredValue)

    if($propValue -is [String] -and ($propValue -eq "true" -or $propValue -eq "false"))
    {
        $propValue = [Boolean]::Parse($propValue)
    }

    if("$propValue" -eq "notConfigured")
    {
        Get-LanguageString "BooleanActions.notConfigured"
    }
    elseif($prop.booleanActions -eq 0 -and $propValue -eq $true)
    {
        Get-LanguageString "BooleanActions.allow"        
    }
    elseif($prop.booleanActions -eq 1 -and $propValue -eq $true)
    {
        Get-LanguageString "BooleanActions.require"        
    }
    elseif($prop.booleanActions -eq 2 -and $propValue -eq $true)
    {
        Get-LanguageString "BooleanActions.enable"        
    }
    elseif($prop.booleanActions -eq 3 -and $propValue -eq $true)
    {
        Get-LanguageString "BooleanActions.block"        
    }
    elseif($prop.booleanActions -eq 4 -and $propValue -eq $true)
    {
        Get-LanguageString "BooleanActions.configured"        
    }
    elseif($prop.booleanActions -eq 5 -and $propValue -eq $true)
    {
        Get-LanguageString "BooleanActions.disable"        
    }
    elseif($prop.booleanActions -eq 6 -and $propValue -eq $true)
    {
        Get-LanguageString "BooleanActions.limit"        
    }
    elseif($prop.booleanActions -eq 7 -and $propValue -eq $true)
    {
        Get-LanguageString "BooleanActions.show"        
    }
    elseif($prop.booleanActions -eq 8 -and $propValue -eq $true)
    {
        Get-LanguageString "BooleanActions.hide"        
    }
    elseif($prop.booleanActions -eq 9 -and $propValue -eq $true)
    {
        Get-LanguageString "BooleanActions.yes"        
    }
    # Custom 
    elseif($prop.booleanActions -eq 100)
    {
        Get-LanguageString (?: $propValue "BooleanActions.block" "BooleanActions.allow")
    }    
    elseif($prop.booleanActions -eq 101)
    {
        Get-LanguageString (?: $propValue "BooleanActions.require" "SettingDetails.notRequired")
    }
    elseif($prop.booleanActions -eq 102)
    {
        Get-LanguageString (?: $propValue "BooleanActions.enable" "BooleanActions.disable")
    }
    elseif($prop.booleanActions -eq 107 -and $propValue -eq $true)
    {
        Get-LanguageString (?: $propValue "BooleanActions.show" "BooleanActions.hide")
    }
    elseif($prop.booleanActions -eq 108 -and $propValue -eq $true)
    {
        Get-LanguageString (?: $propValue "BooleanActions.hide" "BooleanActions.show")
    }
    elseif($prop.booleanActions -eq 109)
    {
        Get-LanguageString (?: $propValue "BooleanActions.yes" "SettingDetails.no")
    }
    elseif($prop.booleanActions -eq 110)
    {
        Get-LanguageString (?: $propValue "SettingDetails.no" "BooleanActions.yes")
    }
    elseif($prop.booleanActions -eq 120)
    {
        Get-LanguageString (?: $propValue "SettingDetails.onOption" "SettingDetails.offOption")
    }
    elseif($prop.booleanActions -eq 200)
    {
        Get-LanguageString (?: $propValue "BooleanActions.allow" "BooleanActions.block")
    }    
    elseif($prop.booleanActions -eq 201)
    {
        Get-LanguageString (?: $propValue "SettingDetails.notRequired" "BooleanActions.require")
    }
    elseif($prop.booleanActions -eq 220)
    {
        Get-LanguageString (?: $propValue "SettingDetails.offOption" "SettingDetails.onOption")
    }
    else
    {
        Add-NotConfiguredProperty $prop
        Get-LanguageString "BooleanActions.notConfigured"
    }
}

function Add-NotConfiguredProperty
{
    param($prop)

    # Add not configured prop to the base object
    if(-not ($script:currentObject.PSObject.Properties | Where Name -eq "@UnconfiguredProperties"))
    {
        $script:currentObject | Add-Member Noteproperty -Name "@UnconfiguredProperties" -Value @() -Force 
    }

    $script:currentObject.'@UnconfiguredProperties' += $prop
}

function Invoke-TranslateOption
{
    param($obj, $prop, [Switch]$SkipOptionChildren, $propValue = $null)

    if(-not $propValue)
    {
        $propValue = $obj."$($prop.entityKey)"
    }

    if($obj.defenderSecurityCenterDisableRansomwareUI -eq $true)
    {
        $obj.defenderSecurityCenterDisableRansomwareUI = "blockOption"
    }    

    foreach($option in $prop.options)
    {
        if("$propValue" -eq "$($option.Value)") # Compare strings since some properties is boolean but option value is string
        {
            if($option.nameResource)
            {
                $optionValue = $prop.nameResource
            }
            elseif($option.displayText)
            {
                $optionValue = $option.displayText
            }
            elseif($option.nameResourceKey)
            {
                if($option.nameResourceKey -eq "notConfigured")
                {
                    Add-NotConfiguredProperty $prop
                }

                $optionValue = (Get-LanguageString (?: $option.nameResourceKey.Contains(".") $option.nameResourceKey "SettingDetails.$($option.nameResourceKey)"))
            }
            else
            {
                $optionValue = $option.Value
            }
            
            @{
                Option=$option
                Value=$optionValue
            }

            Add-PropertyInfo $prop $optionValue $propValue
            if($SkipOptionChildren -ne $true)
            {
                Invoke-ChildSections (Get-CustomChildObject $obj $prop) $option
            }
            break
        }
    }

    # This does not work. Sometime's boolean children are added on $true and sometimes on $false. Depends on the property type
    # This WILL cause "disabled properties" in the documentation
    # Log for now to identify how many properties it is 

    if($propValue -is [Boolean] -and $prop.ChildSettings.Count -gt 0)
    {
        Write-Log "Child properties for boolean $($prop.EntityKey) value=$($propValue) added. Disabled items might be included." 2
    }

    #if($propValue -isnot [Boolean] -or $propValue -eq $true)
    #{
    #    # Skip Child Settings if the option value is not configured
        Invoke-ChildSections $obj $prop
    #}
}

function Invoke-TranslateMultiOption
{
    param($obj, $prop)

    if(($obj.PSObject.Properties | Where Name -eq $prop.entityKey))
    {
        $propValues = $obj."$($prop.entityKey)"
        if($propValues -is [String])
        {
            $propValues = $propValues.Split(',')
        }
    }
    elseif( $prop.entityKey -like "*List")
    {   
        # Should NOT be used!     
        $tmpProp = $prop.entityKey.SubString(0, $prop.entityKey.Length - 4)
        $propValues = $obj."$($tmpProp)"
    }

    $selectedValues = @()
    foreach($propValue in $propValues)
    {
        $option = $prop.Options | Where Value -eq $propValue
        if($option)
        {
            if($option.nameResource)
            {
                $selectedValues += $option.nameResource
            }
            else
            {
                $selectedValues += (Get-LanguageString (?: $option.nameResourceKey.Contains(".") $option.nameResourceKey "SettingDetails.$($option.nameResourceKey)"))
            }
        }
    }

    if($selectedValues.Count -gt 0)
    {
        ($selectedValues -join $script:propertySeparator)
    }
    else
    {
        Add-NotConfiguredProperty $prop
        Get-LanguageString "BooleanActions.notConfigured"
    }
}

function Invoke-TranslateMultiOptionBoolean
{
    param($obj, $prop, $selectedValue = $true)

    $propObj = $obj."$($prop.entityKey)"

    $selectedValues = @()
    foreach($propValue in ($propObj.PSObject.Properties).Name)
    {
        if($propObj.$propValue -isnot [Boolean]) { continue }
        $option = $prop.options | Where value -eq $propValue
        if(-not $option) { continue }

        if($propObj.$propValue -eq $selectedValue)
        {
            if($option.nameResource)
            {
                $selectedValues += $option.nameResource
            }
            else
            {
                $selectedValues += (Get-LanguageString (?: $option.nameResourceKey.Contains(".") $option.nameResourceKey "SettingDetails.$($option.nameResourceKey)"))
            }
        }
    }
    
    if($selectedValues.Count -gt 0)
    {
        ($selectedValues -join $script:propertySeparator)
    }
    else
    {
        Get-LanguageString "BooleanActions.notConfigured"
    }
}

function Invoke-TranslateTable
{
    param($obj, $prop)

    if($prop.entityKey -eq ".")
    {
        $propValue = $obj
    }
    else
    {
        $propValue = $obj."$($prop.entityKey)"
    }

    $items = @()
    
    $itemFullValue = @()
    foreach($item in $propValue)
    {
        $itemValues = @()
        $htFullPropInfo = [ordered]@{}
        foreach($column in $prop.Columns)
        {
            if($column.metadata.entityKey -eq "unusedForSingleItems")
            {
                $itemValues += $item
            }
            elseif($column.metadata.entityKey -eq $prop.entityKey -and ($prop.Columns | measure).Count -eq 1)
            {
                # Some tables has the same EntityKey for the table and the column. That will generate the wrong value
                $itemValues += $item 
            }
            elseif(($prop.Columns | measure).Count -eq 1 -and $item."$($column.metadata.entityKey)" -eq $null -and $obj."$($column.metadata.entityKey)" -eq $null -and $item -is [String])
            {
                # Not sure how correct this is but some tables has one column with and EntityKey but the objects is a string list
                $itemValues += $item
            }
            else
            {
                if(($item.PSObject.Properties | Where Name -like $column.metadata.entityKey))
                {
                    $itemTmpVal = $item."$($column.metadata.entityKey)"
                }
                else
                {
                    $itemTmpVal = $obj."$($column.metadata.entityKey)"
                }
                $itemValues += $itemTmpVal
                if($prop.Columns.Count -gt 1)
                {
                    if($column.metadata.nameResourceKey)
                    {
                        $htFullPropInfo.Add((?? (Get-LanguageString (?: $column.metadata.nameResourceKey.Contains(".") $column.metadata.nameResourceKey "SettingDetails.$($column.metadata.nameResourceKey)")) $column.metadata.entityKey), ($itemTmpVal -join $script:propertySeparator))
                    }
                    else
                    {
                        Write-Log "Property $(?? $prop.nameResourceKey $prop.entityKey) does not have nameResourceKey on one of the columns" 2
                    }
                }
            }
        }
        if($htFullPropInfo.Count -gt 0)
        {
            $itemFullValue += [PSCustomObject]$htFullPropInfo
        }

        if($prop.separator)
        {
            $items += $itemValues -join $prop.separator
        }
        else
        {
            $items += $itemValues -join $script:propertySeparator
        }
    }

    if($items.Count -gt 0)
    {
        $params = @{}
        if($itemFullValue.Count -gt 0)
        {
            $params.Add("tableValue", $itemFullValue)
        }
        
        if((-not $prop.nameResourceKey -or $prop.nameResourceKey -eq "Empty") -and $prop.columns[0].metadata.nameResourceKey)
        {
            Add-PropertyInfo $prop.columns[0].metadata ($items -join $script:objectSeparator) $propValue @params
        }
        else
        {
            Add-PropertyInfo $prop ($items -join $script:objectSeparator) $propValue @params
        }        
    }
    else
    {
        if((-not $prop.nameResourceKey -or $prop.nameResourceKey -eq "Empty") -and $prop.Columns[0].metadata.nameResourceKey)
        {
            Add-PropertyInfo $prop.Columns[0].metadata $null
        }
        else
        {
            Add-PropertyInfo $prop $null
        }
    }

    Invoke-ChildSections $obj $prop
}

function Invoke-TranslateDuration
{
    param($obj, $prop)

    (Get-DurationValue  $obj."$($prop.entityKey)")
}

function Get-DurationValue
{   param($durationValue, [Switch]$ReturnTimeSpan)

    if(-not $durationValue -or !$durationValue.StartsWith("P")) { return "0" }


    $arr = @('P','T','Y','D','H','M','S')
    $values = $durationValue.Split($arr)

    $years = 0
    $days = 0
    $hours = 0
    $minutes = 0
    $seconds = 0

    $i = 0
    foreach($tmp in $arr)
    {
        if($durationValue.Contains($tmp))
        {
            if($ReturnTimeSpan -eq $true)
            {
                if($tmp -eq "Y") { $years = [int]$values[$i] }
                elseif($tmp -eq "D") { $days = [int]$values[$i] }
                elseif($tmp -eq "H") { $hours = [int]$values[$i] }
                elseif($tmp -eq "M") { $minutes = [int]$values[$i] }
                elseif($tmp -eq "S") { $seconds = [int]$values[$i] }
            }
            elseif(![String]::IsNullOrEmpty($values[$i])) { return $values[$i] }
            $i++
        }
    }

    if($ReturnTimeSpan -eq $true)
    {
        $days += ($years * 365) # Not really 100% true
        [timespan]::new($days, $hours, $minutes, $seconds)
    }
    else
    {
        return "0"
    }
}

function Add-Header
{
    param($sectionObj, [switch]$Category)

    if($sectionObj.nameResourceKey)
    {
        $script:CurrentSubCategory = (Get-LanguageString "SettingDetails.$($sectionObj.nameResourceKey)")
    }
}

#endregion

#region 
function Invoke-TranslateScheduledActionType
{
    param($scheduleActionObjects)

    $category = Get-LanguageString "Category.complianceActionsLabel"

    foreach($actionRule in $scheduleActionObjects)
    {
        foreach($actionConfig in $actionRule.scheduledActionConfigurations)
        {
            $notificationTemplate = $null
            $notificationTemplateId = $null
            $additionalNotifications = $null
            $additionalNotificationsList = $null

            if($actionConfig.actionType -eq "notification")
            {
                # Had to change the key of the language string. 
                # Json in 5.x does not support importing json data with two keys with the same name e.g.
                # notification and Notification 
                # Not a good workaround but -AsHashtable is only available in PowerShell 6.x + and JavaScriptSerializer is not good at encoding
                $notificationEnumString = "emailNotification"
            }
            else
            {
                $notificationEnumString = $actionConfig.actionType
            }

            $actionType = Get-LanguageString "ScheduledAction.$($notificationEnumString)"
            
            if($actionConfig.gracePeriodHours -eq 0)
            {
                $schedule = Get-LanguageString "ScheduledAction.List.immediately"
            }
            else
            {
                # Always sets the gracePeriodHours property but always uses days
                $schedule = ((Get-LanguageString "ScheduledAction.List.gracePeriodDays") -f ($actionConfig.gracePeriodHours/24))
            }

            if($actionConfig.actionType -eq "notification")
            {
                if($actionConfig.notificationTemplateId -ne [Guid]::Empty)
                {
                    $notificationTemplate = Get-LanguageString "ScheduledAction.Notification.selected"
                    $notificationTemplateId = $actionConfig.notificationTemplateId
                }
                else
                {
                    $notificationTemplate = Get-LanguageString "ScheduledAction.Notification.noneSelected"
                }

                if($actionConfig.notificationMessageCCList.Count -gt 0)
                {
                    $additionalNotifications = ((Get-LanguageString "ScheduledAction.Notification.numSelected") -f $actionConfig.notificationMessageCCList.Count)
                    $additionalNotificationsList = $actionConfig.notificationMessageCCList -join ","
                }
                else
                {
                    $additionalNotifications = Get-LanguageString "ScheduledAction.Notification.noneSelected"
                }
            }

            $objClone = $actionConfig | ConvertTo-Json -Depth 50 | ConvertFrom-Json

            Remove-Property $objClone "Id"
            foreach($prop in $objClone.PSObject.Properties)
            {
                if($prop.Name -like "*@odata*") 
                { 
                    Remove-Property $objClone $prop.Name    
                }
            }            

            # ToDo: Resolve MessageTemplateId and EmailCCIds to actual object names
            $script:objectComplianceActionData += New-Object PSObject -Property @{
                IdStr = ($objClone | ConvertTo-Json -Depth 50 -Compress)
                Action = $actionType
                Schedule = $schedule
                MessageTemplate = $notificationTemplate
                MessageTemplateId = $notificationTemplateId
                EmailCC = $additionalNotifications
                EmailCCIds = $additionalNotificationsList
                Category=$category
                RawJsonValue=($actionConfig | ConvertTo-Json -Depth 50 -Compress)
            }
        }
    }
}
#endregion

#region Assignments
function Invoke-TranslateAssignments
{
    param($obj)

    if($global:chkExcludeAssignments.IsChecked -eq $true) { return }

    $filtersInfo = $null

    $included = @()
    $excluded = @()
    $tmpAssignments = @()

    $groupIds = @()
    $filterIds = @()

    $category = Get-LanguageString "TableHeaders.assignments"

    $groupModeInclude = Get-LanguageString "TableHeaders.includedGroups"
    $groupModeExnclude = Get-LanguageString "TableHeaders.excludedGroups"

    $noFilter = Get-LanguageString "AssignmentFilters.noFilters"
    $filterInclude = Get-LanguageString "SettingDetails.include"
    $filterExclude = Get-LanguageString "SettingDetails.exclude"

    foreach($assignment in $obj.assignments)
    {
        if($assignment.target.groupId -and $assignment.target.groupId -notin $groupIds)
        {
            $groupIds += $assignment.target.groupId
        }

        if($assignment.target.deviceAndAppManagementAssignmentFilterId -and ($assignment.target.deviceAndAppManagementAssignmentFilterId -notin $filterIds))
        {
            $filterIds += $assignment.target.deviceAndAppManagementAssignmentFilterId
        }
    }

    $groupInfo = $null

    if($groupIds.Count -gt 0 -and $script:offlineDocumentation -ne $true)
    {
        $ht = @{}
        $ht.Add("ids", @($groupIds | Unique))

        $body = $ht | ConvertTo-Json

        $groupInfo = (Invoke-GraphRequest -Url "/directoryObjects/getByIds?`$select=displayName,id" -Content $body -Method "Post").Value
    }
    
    if(($null -eq $groupInfo -or ($groupInfo | measure).Count -eq 0) -and $obj."@ObjectFromFile" -eq $true -and $script:migTable)
    {
        ### Get group info from mig table when documenting from file if there's no access to the environment
        $groupInfo = $script:migTable.Objects | Where Type -eq "Group" 
    }    

    if($filterIds.Count -gt 0)
    {        
        if($script:offlineDocumentation -eq $true)
        {
            if($script:offlineObjects["AssignmentFilters"])
            {
                $filtersInfo = $script:offlineObjects["AssignmentFilters"] | Where { $_.Id -in $filterIds }
            }
            else
            {
                Write-Log "No assignment filters loaded for Offline documentation. Check export folder" 2
            }
        }
        else
        {
            $batchInfo = @{}
            $requests = @()
            #{"requests":[{"id":"<FilterID>","method":"GET","url":"deviceManagement/assignmentFilters/<FilterID>?$select=displayName"}]}
            foreach($filterId in $filterIds)
            {
                $requests += [PSCustomObject]@{
                    id = $filterId
                    method = "GET"
                    "url" = "deviceManagement/assignmentFilters/$($filterId)?`$select=displayName"
                }
            }
            $batchInfo = @{"requests"=$requests}
            $jsonBody = $batchInfo | ConvertTo-Json

            $filtersInfo = (Invoke-GraphRequest -Url "/`$batch" -Content $jsonBody -Method "Post").responses.body
        }
    }
    
    foreach($assignment in $obj.assignments)
    {
        $groupMode = $null
        $groupName = $null

        $filterName = $null
        $filterMode = $null
        $fullValueTable = $null

        # ToDo: Not an OK way of specifying this!
        if(($assignment.PSObject.Properties | Where Name -eq "intent") -and
            ($assignment.PSObject.Properties | Where Name -eq "settings"))
        {
            $assignmentType = "mobileAppAssignmentSettings"            
        }
        else
        {
            $assignmentType = "genericAssignment"    
        }

        if($assignment.target.'@odata.type' -eq "#microsoft.graph.exclusionGroupAssignmentTarget")
        {
            $groupMode = "exclude"
        }
        else
        {
            $groupMode = "include"
        }
    
        if($assignment.target.GroupId)
        {
            $groupName = ($groupInfo | Where id -eq $assignment.target.GroupId).displayName
            if(-not $groupName)
            {
                $groupName = $assignment.target.GroupId
            }
        }
        elseif($assignment.target.'@odata.type' -eq "#microsoft.graph.allDevicesAssignmentTarget")
        {
            $groupName = Get-LanguageString "SettingDetails.allDevices"
        }
        elseif($assignment.target.'@odata.type' -eq "#microsoft.graph.allLicensedUsersAssignmentTarget")
        {
            $groupName = Get-LanguageString "SettingDetails.allUsers"
        }
        else
        {
            $groupName = "unknown" # Should not get here!    
        }

        if(($assignment.target.PSObject.Properties | Where Name -eq "deviceAndAppManagementAssignmentFilterId"))
        {        
            $filterName = $noFilter
            $filterMode = $noFilter
    
            if($assignment.target.deviceAndAppManagementAssignmentFilterId -and $filtersInfo)
            {
                $filtersObj = $filtersInfo | Where Id -eq $assignment.target.deviceAndAppManagementAssignmentFilterId
                if($filtersObj.displayName)
                {
                    $filterName = $filtersObj.displayName
                }

                if($assignment.target.deviceAndAppManagementAssignmentFilterType -eq "include")
                {
                    $filterMode = $filterInclude
                }
                else
                {
                    $filterMode = $filterExclude
                }
            }
        }

        # ToDo: Add support if Direct or from Policy Set

        # Add additional app assignment settings
        # Only add settings for Included assignments
        $assignmentSettingProps = [ordered]@{}
        if($assignmentType -eq "mobileAppAssignmentSettings" -and $groupMode -eq "include" -and $assignment.settings -ne $null)
        {            
            foreach($settingProp in @("useDeviceContext",
                                        "vpnConfigurationId",
                                        "uninstallOnDeviceRemoval",
                                        "isRemovable",
                                        "useDeviceLicensing",
                                        "androidManagedStoreAppTrackIds",
                                        "deliveryOptimizationPriority",
                                        "installTimeSettings",
                                        "notifications",
                                        "restartSettings"))
            {
                if(($assignment.settings.PSObject.Properties | Where Name -eq $settingProp))
                {
                    if($settingProp -eq "useDeviceLicensing")
                    {
                        if($assignment.settings.$settingProp -eq $true)
                        {
                            $value = Get-LanguageString "SettingDetails.licenseTypeDevice"
                        }
                        else
                        {
                            $value = Get-LanguageString "SettingDetails.licenseTypeUser"
                        }
                    }
                    elseif($settingProp -eq "restartSettings")
                    {
                        if($null -eq $assignment.settings.$settingProp)
                        {
                            $value = Get-LanguageString "SettingDetails.disabledOption"
                        }
                        else
                        { 
                            $valueArr = @()    
                            #$valueArr += Get-LanguageString "SettingDetails.enabledOption"
                            $valueArr += "$((Get-LanguageString "Assignment.RestartGracePeriod.durationInMinutes"))=$($assignment.settings.restartSettings.gracePeriodInMinutes)"
                            $valueArr += "$((Get-LanguageString "Assignment.RestartGracePeriod.countdownDialog"))=$($assignment.settings.restartSettings.countdownDisplayBeforeRestartInMinutes)"
                            
                            if($null -eq $assignment.settings.restartSettings.restartNotificationSnoozeDurationInMinutes)
                            {
                                $valueArr += "$((Get-LanguageString "Assignment.RestartGracePeriod.allowSnooze"))=$((Get-LanguageString "SettingDetails.no"))"
                            }
                            else
                            {
                                $valueArr += "$((Get-LanguageString "Assignment.RestartGracePeriod.allowSnooze"))=$((Get-LanguageString "SettingDetails.yes"))"
                                $valueArr += "$((Get-LanguageString "Assignment.RestartGracePeriod.snoozeDurationInMinutes"))=$($assignment.settings.restartSettings.restartNotificationSnoozeDurationInMinutes)"
                            }
                            $value = $valueArr -join $script:objectSeparator
                        }
                    }
                    elseif($settingProp -eq "notifications")
                    {
                        $value = ?? (Get-LanguageString "AppResources.AssignmentToast.$($assignment.settings.$settingProp)") $assignment.settings.$settingProp
                    }
                    elseif($settingProp -eq "installTimeSettings")
                    {                        
                        $asap = Get-LanguageString "Assignment.SoftwareInstallationTime.defaultTime"
                        $startValue = $asap
                        $value = $asap

                        if($assignment.settings.installTimeSettings)
                        {
                            if($assignment.settings.installTimeSettings.startDateTime)
                            {
                                $instTime = Get-Date $assignment.settings.installTimeSettings.startDateTime
                                
                                if($assignment.settings.installTimeSettings.useLocalTime -eq $false)
                                {
                                    $hours = ($instTime.ToUniversalTime() - $instTime).Hours
                                    $instTime = $instTime.AddHours($hours)
                                }
                                $startValue = "$($instTime.ToShortDateString()) $($instTime.ToShortTimeString())" 
                            }

                            if($assignment.settings.installTimeSettings.deadlineDateTime)
                            {
                                $endTime = Get-Date $assignment.settings.installTimeSettings.deadlineDateTime
                                
                                if($assignment.settings.installTimeSettings.useLocalTime -eq $false)
                                {
                                    $hours = ($endTime.ToUniversalTime() - $endTime).Hours
                                    $endTime = $endTime.AddHours($hours)
                                }
                                $value = "$($endTime.ToShortDateString()) $($endTime.ToShortTimeString())" 
                            }
                        }

                        $assignmentSettingProps.Add("startTimeColumnLabel", $startValue)
                        if($assignment.Intent -eq "available")
                        {
                            continue # No install deadline on available assignments
                        }
                    }                    
                    elseif($settingProp -eq "deliveryOptimizationPriority")
                    {
                        $tmpStr = Get-LanguageString "AppResources.DeliveryOptimizationPriority.displayText"
                        if($assignment.settings.$settingProp -ne "foreground")
                        {
                            $tmpType = Get-LanguageString "AppResources.DeliveryOptimizationPriority.backgroundNormal"
                        }
                        else
                        {
                            $tmpType = Get-LanguageString "AppResources.DeliveryOptimizationPriority.foreground"
                        }
                        $value = $tmpStr -f $tmpType
                    }
                    elseif($assignment.settings.$settingProp -eq "notConfigured")
                    {
                        $value = Get-LanguageString "BooleanActions.notConfigured"
                    }                    
                    else
                    {
                        $value = $assignment.settings.$settingProp
                    }
                    $assignmentSettingProps.Add($settingProp, $value)
                }
            }
        }
        
        if($assignmentType -eq "genericAssignment")
        {            
            $assignObj = [PSCustomObject]@{
                GroupMode = (?: ($groupMode -eq "include") $groupModeInclude $groupModeExnclude)
                Group = $groupName
                Type = "GenericAssignment"
                Category = (?: ($groupMode -eq "include") $groupModeInclude $groupModeExnclude)
            }

            if($groupMode -eq "include")
            {
                $included += $assignObj
                if($filterMode -ne $null)
                {
                    $assignObj | Add-Member Noteproperty -Name "Filter" -Value $filterName -Force 
                    $assignObj | Add-Member Noteproperty -Name "FilterMode" -Value $filterMode -Force 
                }
            }
            else
            {
                $excluded += $assignObj
            }
        }
        else
        {
            $appAssignment = @{        
                Group =  $groupName
                GroupMode = Get-LanguageString "AssignmentAction.$($groupMode)"
                Category = Get-LanguageString "InstallIntent.$($assignment.Intent)"
                RawIntent = $assignment.Intent
                RawJsonValue = ($assignment | ConvertTo-Json -Depth 50 -Compress)
            }

            if($groupMode -eq "include")
            {
                $appAssignment.Add("Filter", $filterName)
                $appAssignment.Add("FilterMode", $filterMode)
            }            

            if($assignmentSettingProps.Count -gt 0)
            {
                $appAssignment.Add("Settings", $assignmentSettingProps)
            }
            
            $tmpAssignments += [PSCustomObject]$appAssignment
        }        
    }
    
    if($included.Count -gt 0)
    {
        $script:objectAssignments += $included
        <#
        $script:objectAssignments += [PSCustomObject]@{                
            GroupMode = Get-LanguageString "TableHeaders.includedGroups"
            Groups = ($included -join $script:objectSeparator)            
            Type = "GenericAssignment"
        }
        #>
    }

    if($excluded.Count -gt 0)
    {
        $script:objectAssignments += $excluded
        <#
        $script:objectAssignments += [PSCustomObject]@{
                
            GroupMode = Get-LanguageString "TableHeaders.excludedGroups"
            Groups = ($excluded -join $script:objectSeparator)
            Type = "GenericAssignment"
        }
        #>
    }

    if($tmpAssignments.Count -gt 0)
    {
        # Sort the items in the correct order
        foreach($intent in @("required","available","availableWithoutEnrollment","uninstall"))
        {
            $script:objectAssignments += $tmpAssignments | Where RawIntent -eq $intent
        }
    }
}
#endregion

#region Applications

function Get-GraphAppType
{
    param($obj)

    if(-not $script:allAppTypes)
    {
        $fi = [IO.FileInfo]($global:AppRootFolder + "\Documentation\AppTypes.json")
        if(!$fi.Exists)
        {        
            return $false
        }
        $script:allAppTypes = Get-Content ($global:AppRootFolder + "\Documentation\AppTypes.json") | ConvertFrom-Json        
    }

    foreach($appType in ($script:allAppTypes | Where ODataType -eq $obj.'@OData.Type'))
    {
        if($appType.Condition)
        {
            if($obj."$($appType.Condition.Property)" -eq $appType.Condition.Value)
            {
                return $appType
            }
        }
        else
        {
            return $appType
        }
    }
}
#endregion

#region Custom Policy objects - base on json file named after OData.Type
function Invoke-TranslateCustomProfileObject
{
    param($obj, $fileName)

    $fi = [IO.FileInfo]($global:AppRootFolder + "\Documentation\ObjectInfo\$($fileName).json")
    if(!$fi.Exists)
    {        
        return $false
    }
    $jsonObj = Get-Content ($global:AppRootFolder + "\Documentation\ObjectInfo\$($fileName).json") | ConvertFrom-Json

    $platformType = Get-ObjectPlatformFromType $obj 

    Add-BasicDefaultValues $obj $objectType 

    Add-CustomProfileProperties $obj

    Invoke-TranslateSection $obj $jsonObj

    Add-BasicAdditionalValues $obj $objectType

    return $true
}
#endregion

function Add-CustomSettingObject
{
    param($settingsObj)

    $script:objectSettingsData += $settingsObj
}

function Save-DocumentationFile
{
    param($content, $fileName, [switch]$OpenFile)

    try
    {
        $content | Out-File -LiteralPath $fileName -Force -Encoding utf8 -ErrorAction Stop
        Write-Log "$fileName saved successfully"
        if($OpenFile -eq $true)
        {
            Invoke-Item $fileName
        }        
    }
    catch
    {
        Write-LogError "Failed to save file $fileName." $_.Exception
    }
}

#region Initiate documentation
function Show-DocumentationForm
{
    param(
        [Parameter(Mandatory=$true,ParameterSetName = "Objects")]
        $objects, 
        [Parameter(Mandatory=$true,ParameterSetName = "ObjectTypes")]
        $objectTypes,
        [Switch]
        [Parameter(Mandatory=$false,ParameterSetName = "Objects")]
        $SelectedDocuments,    
        [Switch]
        $ShowFolderSource 
        )

    $objectList = @()
    
    if($objects)
    {
        if($SelectedDocuments -eq $true)
        {
            $objectList += $objects
            $objects | ForEach-Object { $_.IsSelected = $true }
        }
        else
        {
            $objects | ForEach-Object {                
                # This will create a new "root" object so it doesn't update properties like IsSelected ect on the original object
                $item = [PSCustomObject]@{
                    Title =  $null
                }

                foreach($prop in $_.PSObject.Properties)
                {
                    $item | Add-Member Noteproperty -Name $prop.Name -Value $_."$($prop.Name)" -Force
                }

                $item.Title = (Get-GraphObjectName $_.Object $_.ObjectType) 
                $item.IsSelected = $true 
                $objectList += $item                
            }
        }
        $sourceType = "Objects"
    }
    elseif($objectTypes)
    {
        foreach($groupId in ($objectTypes | Select GroupId -Unique).GroupId)
        {
            if(-not $groupId) { continue }

            #$script:DocumentationLanguage = ?? $global:cbDocumentationLanguage.SelectedValue "en"
            $script:DocumentationLanguage = "en"
            $groupName = Get-ObjectTypeString -ObjectType $groupId
            if(-not $groupName)
            {
                Write-Log "Group id $groupId could not be translated" 2
                $groupName = $groupId
            }
            $item = [PSCustomObject]@{
                IsSelected = $true
                Title =  $groupName
                GroupId = $groupId
            }
            $objectList += $item
        }        

        $objectList = $objectList | Sort -Property Title

        $sourceType = "ObjectTypes"
    }
    else
    {
        return
    }

    if($objectList.Count -eq 0) { Write-Log "No objects found/selected!";return }
    $ocList = [System.Collections.ObjectModel.ObservableCollection[object]]::new(@($objectList))

    $script:docForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\DocumentationForm.xaml") -AddVariables
    if(-not $script:docForm) { return }

    $global:grdDocumentObjects.Tag = $sourceType

    if($sourceType -eq "Objects")
    {
        $global:btnClearDocumentationList.Visibility = ?: ($SelectedDocuments -eq $true) "Visible" "Collapsed"
        $global:btnAddToDocumentationList.Visibility = ?: ($SelectedDocuments -ne $true) "Visible" "Collapsed"
    }
    else
    {
        $global:btnClearDocumentationList.Visibility = "Collapsed"
        $global:btnAddToDocumentationList.Visibility = "Collapsed"
    }

    if($ShowFolderSource -ne $true)
    {
        $global:grdDocumentFromFolder.Visibility = "Collapsed"
        $global:spDocumentFromFolder.Visibility = "Collapsed"

        $global:txtDocumentFilter.Visibility = "Collapsed"
        $global:spDocumentFilter.Visibility = "Collapsed"
    }
    else
    {
        Add-XamlEvent $script:docForm "browseDocumentFromFolder" "add_click" {
            $folder = Get-Folder (Get-XamlProperty $script:docForm "txtDocumentFromFolder" "Text") "Select root folder for export files"
            if($folder)
            {
                Set-XamlProperty $script:docForm "txtDocumentFromFolder" "Text" $folder
            }
        }        
    }

    $column = Get-GridCheckboxColumn "IsSelected"
    $global:grdDocumentObjects.Columns.Add($column)

    $column.Header.IsChecked = $true # All items are checked by default
    $column.Header.add_Click({
            foreach($item in $global:grdDocumentObjects.ItemsSource)
            {
                $item.IsSelected = $this.IsChecked
            }
            $global:grdDocumentObjects.Items.Refresh()
        }
    )    

    # Add title column
    $binding = [System.Windows.Data.Binding]::new("Title")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Title"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:grdDocumentObjects.Columns.Add($column)
    
    if($SelectedDocuments -eq $true)
    {
        $binding = [System.Windows.Data.Binding]::new("ObjectType.Id")
        $column = [System.Windows.Controls.DataGridTextColumn]::new()
        $column.Header = "Type"
        $column.IsReadOnly = $true
        $column.Binding = $binding
        $global:grdDocumentObjects.Columns.Add($column)  
    }

    $global:grdDocumentObjects.ItemsSource = [System.Windows.Data.CollectionViewSource]::GetDefaultView($ocList)    

    Set-FromValues
    
    Add-XamlEvent $script:docForm "btnClose" "add_click" {
        $script:docForm = $null
        Show-ModalObject 
    }

    Add-XamlEvent $script:docForm "btnAddToDocumentationList" "add_click" {
        Add-DocumentationObjects ($global:grdDocumentObjects.ItemsSource | Where IsSelected -eq $true)
    }

    Add-XamlEvent $script:docForm "btnClearDocumentationList" "add_click" {
        if([System.Windows.MessageBox]::Show("Are you sure you want to clear the items in the list?", "Documentation", "YesNo", "Question") -eq "Yes")
        {
            $global:grdDocumentObjects.ItemsSource = $null
            $global:btnStartDocumentation.IsEnabled = $false
        }
    }

    Add-XamlEvent $script:docForm "btnStartDocumentation" "add_click" {

        Invoke-StartDocumentatiom       
    }
    
    Add-XamlEvent $script:docForm "btnCopyBasic" "add_click" {
        if($script:objectBasicInfo.Count -gt 0)
        {
            $script:objectBasicInfo | Select -Property Name,Value | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
        }
    }

    Add-XamlEvent $script:docForm "btnCopySettings" "add_click" {
        if($script:objectSettingsData.Count -gt 0)
        {
            $script:objectSettingsData | Select -Property $script:settingsProperties | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
        }
    }

    Add-XamlEvent $script:docForm "btnCopyAll" "add_click" {
        $tmpArr = @()
        
        if($script:objectBasicInfo.Count -gt 0)
        {
            $tmpArr += $script:objectBasicInfo | Select -Property Name,Value | ConvertTo-Csv -NoTypeInformation 
        }

        if($script:objectSettingsData.Count -gt 0)
        {
            $tmpArr += $script:objectSettingsData | Select -Property $script:settingsProperties | ConvertTo-Csv -NoTypeInformation
        }

        if($script:applicabilityRules.Count -gt 0)
        {
            $tmpArr += $script:applicabilityRules | Select -Property Rule,Property,Value,Category | ConvertTo-Csv -NoTypeInformation
        }                

        if($script:objectComplianceActionData.Count -gt 0)
        {
            $tmpArr += $script:objectComplianceActionData | Select Action,Schedule,MessageTemplate,EmailCC,Category | ConvertTo-Csv -NoTypeInformation
        }
        $tmpArr | Set-Clipboard
    }

    $global:cbDocumentationType.Add_SelectionChanged({        
        Set-OutputOptionsTabStatus $this
    })
    
    if($global:grdDocumentObjects.Tag -eq "ObjectTypes")
    {
        $global:btnExportSettingsForSilentExport.Visibility = "Visible"

        Add-XamlEvent $script:docForm "btnExportSettingsForSilentExport" "add_click" ({
            $sf = [System.Windows.Forms.SaveFileDialog]::new()
            $sf.FileName = "BulkDocumentation.json"
            $sf.DefaultExt = "*.json"
            $sf.Filter = "Json (*.json)|*.json|All files (*.*)|*.*"
            if($sf.ShowDialog() -eq "OK")
            {
                $tmp = [PSCustomObject]@{
                    Name = "ObjectTypes"
                    Type = "Custom"
                    ObjectTypes = @()
                }

                foreach($tmpObj in ($global:grdDocumentObjects.ItemsSource | Where IsSelected -eq $true))
                {
                    $tmp.ObjectTypes += $tmpObj.GroupId
                }
                Export-GraphBatchSettings $sf.FileName $script:docForm "BulkDocumentation" @($tmp)
            }  
        })
    }

    Set-OutputOptionsTabStatus $global:cbDocumentationType

    Show-ModalForm "Intune Documentation" $script:docForm -HideButtons
}

function Invoke-SilentBatchJob
{
    param($settingsObj)

    if($settingsObj.BulkDocumentation)
    {
        $global:currentViewObject =  $global:viewObjects | Where { $_.ViewInfo.ID -eq "IntuneGraphAPI" }
        Start-DocSilentBulkDocumentation $settingsObj
    }
}

function Start-DocSilentBulkDocumentation
{
    param($settingsObj)

    $script:docForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\DocumentationForm.xaml") -AddVariables
    if(-not $script:docForm) { return }
        
    $global:grdDocumentObjects.Tag = "ObjectTypes"

    # Get all objects but not selected
    # This will allow dependencies
    $script:docObjectTypes = Get-GraphBatchObjectTypes $settingsObj.BulkDocumentation -NotSelected -All
    
    $objectList = @()
    $objTypes = $settingsObj.BulkDocumentation | Where Name -eq ObjectTypes
    if($objTypes)
    {        
        # Select object types from the batch file        
        foreach($objTypeId in $objTypes.ObjectTypes)
        {
            $objectList += [PSCustomObject]@{
                IsSelected = $true
                Title = $objTypeId
                GroupId = $objTypeId
            }
        }
    }    

    if($objectList.Count -eq 0) { Write-Log "No objects found/selected!";return }
    $ocList = [System.Collections.ObjectModel.ObservableCollection[object]]::new(@($objectList)) 

    $global:grdDocumentObjects.ItemsSource = [System.Windows.Data.CollectionViewSource]::GetDefaultView($ocList)    

    Set-FromValues

    $global:cbDocumentationType.SelectedValue = ($settingsObj.BulkDocumentation | Where Name -eq "cbDocumentationType").value

    Set-OutputOptionsTabStatus $global:cbDocumentationType    
    
    # Skip warning since it can't find custom tab controls in main from (eg Word settings)
    # Probably requires some RegisterName...
    # ToDo: Fix this in a proper way
    Set-BatchProperties $settingsObj.BulkDocumentation $script:docForm -SkipMissingControlWarning
    Set-BatchProperties $settingsObj.BulkDocumentation $global:ccOutputCustomOptions.Content -SkipMissingControlWarning

    Invoke-StartDocumentatiom
}

function local:Set-FromValues
{
    $global:cbDocumentationType.ItemsSource = $global:documentationOutput
    $global:cbDocumentationType.SelectedValue = (Get-Setting "Documentation" "OutputType" "none")

    if(-not $script:Languages)
    {
        $script:Languages = Get-Content ($global:AppRootFolder + "\Documentation\Languages.json") -Encoding UTF8 | ConvertFrom-Json
    }

    if($script:Languages)
    {
        $global:cbDocumentationLanguage.ItemsSource = $script:Languages
        $global:cbDocumentationLanguage.SelectedValue = (Get-Setting "Documentation" "Language" "en")
    }

    $global:cbDocumentationPropertySeparator.ItemsSource = @(",",";","-","|")
    try
    {
        $global:cbDocumentationPropertySeparator.SelectedIndex = $global:cbDocumentationPropertySeparator.ItemsSource.IndexOf((Get-Setting "Documentation" "PropertySeparator" ";"))
    }
    catch {}

    $objectSeparator = "[ { Name: `"New line`",Value: `"$([System.Environment]::NewLine)`" }, {Name: `";`",Value: `";`" }, {Name: `"|`",Value: `"|`" }]" | ConvertFrom-Json
    $global:cbDocumentationObjectSeparator.ItemsSource = $objectSeparator
    $global:cbDocumentationObjectSeparator.SelectedValue = (Get-Setting "Documentation" "ObjectSeparator" ([System.Environment]::NewLine)) #"$([System.Environment]::NewLine)")

    $valueOutputPropertiyTypes = "[ { Name: `"Value`",Value: `"value`" }, {Name: `"Value with label`", Value: `"valueWithLabel`" }]" | ConvertFrom-Json
    $global:cbDocumentationValueOutputProperty.ItemsSource = $valueOutputPropertiyTypes
    $global:cbDocumentationValueOutputProperty.SelectedValue = (Get-Setting "Documentation" "ValueOutputProperty" "value")

    $global:chkSetUnconfiguredValue.IsChecked = ((Get-Setting "Documentation" "SetUnconfiguredValue" "true") -ne "false")
    $global:chkSetDefaultValue.IsChecked = ((Get-Setting "Documentation" "SetDefaultValue" "false") -ne "false")

    $global:chkIncludeScripts.IsChecked = ((Get-Setting "Documentation" "IncludeScripts" "true") -ne "false")
    $global:chkExcludeScriptSignature.IsChecked = ((Get-Setting "Documentation" "ExcludeScriptSignature" "false") -ne "false")
    $global:chkExcludeAssignments.IsChecked = ((Get-Setting "Documentation" "ExcludeAssignments" "false") -eq "true")
    
    $notConfiguredItems = "[ { Name: `"Not configured (Localized)`",Value: `"notConfigured`" }, { Name: `"Empty`",Value: `"empty`" }, { Name: `"Don't change`",Value: `"asis`" }]" | ConvertFrom-Json
    $global:cbNotConifugredText.ItemsSource = $notConfiguredItems
    $global:cbNotConifugredText.SelectedValue = (Get-Setting "Documentation" "NotConfiguredText" "")

    $global:chkSkipNotConfigured.IsChecked = ((Get-Setting "Documentation" "SkipNotConfigured" "false") -ne "false")
    $global:chkSkipDefaultValues.IsChecked = ((Get-Setting "Documentation" "SkipDefaultValues" "false") -ne "false")
    $global:chkSkipDisabled.IsChecked = ((Get-Setting "Documentation" "SkipDisabled" "true") -ne "false")    
}
function local:Invoke-StartDocumentatiom
{
    $txtDocumentationRawData.Text = ""

    $script:offlineDocumentation = $false
    $script:offlineObjects = @{}

    $loadExportedInfo = $true
    $script:migTable = $null
    $script:scopeTags = $null

    $diSource = $nul
    $global:intentCategories = $null
    $global:catRecommendedSettings = $null
    $global:intentCategoryDefs = $null
    $global:cfgCategories = $null

    $script:DocumentationLanguage = ?? $global:cbDocumentationLanguage.SelectedValue "en"        
    $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
    $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","

    $script:ValueOutputProperty = ?? $global:cbDocumentationValueOutputProperty.SelectedValue "value"  

    Save-Setting "Documentation" "OutputType" $global:cbDocumentationType.SelectedValue        

    Save-Setting "Documentation" "Language" $script:DocumentationLanguage
    Save-Setting "Documentation" "ObjectSeparator" $script:objectSeparator
    Save-Setting "Documentation" "PropertySeparator" $script:propertySeparator

    Save-Setting "Documentation" "ValueOutputProperty" $script:ValueOutputProperty

    Save-Setting "Documentation" "SetUnconfiguredValue" $global:chkSetUnconfiguredValue.IsChecked
    Save-Setting "Documentation" "SetDefaultValue" $global:chkSetDefaultValue.IsChecked

    Save-Setting "Documentation" "IncludeScripts" $global:chkIncludeScripts.IsChecked
    Save-Setting "Documentation" "ExcludeScriptSignature" $global:chkExcludeScriptSignature.IsChecked
    Save-Setting "Documentation" "ExcludeAssignments" $global:chkExcludeAssignments.IsChecked    

    Save-Setting "Documentation" "SkipNotConfigured" $global:chkSkipNotConfigured.IsChecked
    Save-Setting "Documentation" "SkipDefaultValues" $global:chkSkipDefaultValues.IsChecked
    Save-Setting "Documentation" "SkipDisabled" $global:chkSkipDisabled.IsChecked

    Save-Setting "Documentation" "NotConfiguredText" $global:cbNotConifugredText.SelectedValue

    Get-CustomIgnoredCategories $obj

    Invoke-InitDocumentation

    if($global:grdDocumentObjects.Tag -eq "Objects")
    {
        $sourceList = @()
        $groupIds = $global:grdDocumentObjects.ItemsSource.ObjectType.GroupId | Select -Unique | Sort GroupId
        foreach($groupId in $groupIds)
        {
            $groupSourceList = @()
            $curObjectType = $null 
            foreach($tmpObj in ($global:grdDocumentObjects.ItemsSource | Where { $_.IsSelected -eq $true -and $_.ObjectType.GroupId -eq $groupId } ))
            {
                $groupSourceList += $tmpObj
                $curObjectType = $tmpObj.ObjectType
                if($curObjectType.GroupId -eq "EndpointSecurity")
                {
                    $catName = Get-IntentCategory $tmpObj.Category
                    $tmpObj | Add-Member Noteproperty -Name "CategoryName" -Value $catName -Force
                }
                else
                {
                    $tmpObj | Add-Member Noteproperty -Name "Category" -Value $tmpObj.ObjectType.Id -Force
                }

                $tmpObj | Add-Member Noteproperty -Name "GroupName" -Value (Get-ObjectTypeGroupName $tmpObj.ObjectType) -Force
            }

            if($groupSourceList.Count -eq 0) { contnue }

            if($curObjectType.GroupId -eq "EndpointSecurity")
            {
                $sortProps = @("CategoryName","Title") # "displayName")
            }
            else
            {
                #!!!###$sortProps =  @({$_.ObjectType.Title},{(Get-GraphObjectName $_.Object $_.ObjectType)}) #@("Title") #@((?? $objectType.NameProperty "displayName"))
                $sortProps = @({$_.GroupName},{(Get-GraphObjectName $_.Object $_.ObjectType)}) 
            }
            $sourceList += ($groupSourceList | Sort-Object -Property $sortProps)
        }
    }
    elseif($global:grdDocumentObjects.Tag -eq "ObjectTypes")
    {
        $fromExportFolder = $false
        if($global:txtDocumentFromFolder.Text)
        {
            $diSource = [IO.DirectoryInfo]$global:txtDocumentFromFolder.Text
            if($diSource.Exists -eq $false)
            {
                [System.Windows.MessageBox]::Show("Source folder not found:`n`n$($diSource.FullName)", "Documentation", "OK", "Error")
                Write-Status ""
                return
            }
            $fromExportFolder = $true
        }

        $sourceList = @()
        foreach($objGroup in ($global:grdDocumentObjects.ItemsSource | Where IsSelected -eq $true))
        {
            $groupSourceList = @()
            foreach($objectType in ($global:currentViewObject.ViewItems | Where GroupId -eq $objGroup.GroupId))
            {
                Write-Status "Get $($objectType.Title) objects"
            
                if($fromExportFolder -eq $false)
                {
                    [array]$graphObjects = Get-GraphObjects -property $objectType.ViewProperties -objectType $objectType
                }
                else
                {
                    $objectPath = [IO.Path]::Combine($diSource.FullName,$objectType.ID)
                    if([IO.Directory]::Exists($objectPath) -eq $false)
                    {
                        Write-Log "Object path for $($objectType.Title) ($($objectType.ID)) not found. Skipping object type" 2 
                        continue
                    }
                    $graphObjects = Get-GraphFileObjects $objectPath -ObjectType $objectType
                    $graphObjects | ForEach-Object { $_.Object | Add-Member Noteproperty -Name "@ObjectFromFile" -Value $true -Force }
                }
                $groupSourceList += $graphObjects
            }

            if($objGroup.GroupId -eq "EndpointSecurity")
            {
                foreach($tmpObj in $groupSourceList)
                {
                    $catName = Get-IntentCategory $tmpObj.Category
                    $tmpObj | Add-Member Noteproperty -Name "CategoryName" -Value $catName -Force
                }
                $sortProps = @("CategoryName","displayName")
            }
            else
            {
                foreach($tmpObj in $groupSourceList)
                {
                    $tmpObj | Add-Member Noteproperty -Name "Category" -Value $tmpObj.ObjectType.Id -Force
                    $tmpObj | Add-Member Noteproperty -Name "GroupName" -Value (Get-ObjectTypeGroupName $tmpObj.ObjectType) -Force
                }                
                #!!!###$sortProps = @({$_.ObjectType.Title},{(Get-GraphObjectName $_.Object $_.ObjectType)}) 
                $sortProps = @({$_.GroupName},{(Get-GraphObjectName $_.Object $_.ObjectType)}) 
            }
            $sourceList += ($groupSourceList | Sort-Object -Property $sortProps)
        }
    }
    else
    {
        return  
    }

    if($fromExportFolder -eq $true -and $diSource -and $loadExportedInfo -eq $true)
    {
        $loadExportedInfo = $false

        $migFileName = [IO.Path]::Combine($diSource.FullName,"MigrationTable.json")
        if([IO.File]::Exists($migFileName) -eq $false)
        {
            Write-Log "MigrationTable not found. Groups will be documented with GroupId" 2
        }
        else
        {
            Write-Log "Load Migration table from $migFileName"
            $script:migTable = ConvertFrom-Json (Get-Content $migFileName -Raw)
        }

        if($script:migTable.TenantId -and $script:migTable.TenantId -ne $global:organization.id)
        {
            $script:offlineDocumentation = $true
        }

        if($script:offlineDocumentation -eq $true)
        {
            Add-DocOfflineDependencies "ScopeTags" $diSource.FullName
            Add-DocOfflineDependencies "AssignmentFilters" $diSource.FullName
            Add-DocOfflineObjectTypeDependencies  $diSource.FullName
            if($script:offlineObjects.ContainsKey("ScopeTags"))
            {
                $script:scopeTags = @($script:offlineObjects["ScopeTags"])
            }                
        }
    }

    if($global:cbDocumentationType.SelectedItem.PreProcess)
    {
        Write-Status "Run PreProcess for $($global:cbDocumentationType.SelectedItem.Name)"
        & $global:cbDocumentationType.SelectedItem.PreProcess
    }

    $tmpCurObjectType = $null
    $tmpCurObjectGroup = $null
    $allObjectTypeObjects = @()

    $groupCategoryCount = 2 # Force group header. The above code is not working since it should be inside the actual group

    $filter = $global:txtDocumentFilter.Text.Trim()
    
    Write-Log "Filter: $filter" 

    $tmpList = @()
    
    # Remove objects that should be excluded based on filter
    # Get documentation data and remove objects not supporting documentation
    foreach($curObj in $sourceList)
    {
        # Filter out objects if they return scopeTags in a list
        if($curObj.ObjectType.ScopeTagsReturnedInList -ne $false -and (Confirm-GraphMatchFilter $curObj $filter) -eq $false)
        {
            continue
        }
        
        if($curObj.Object."@ObjectFromFile" -eq $true)
        {
            $obj = $curObj
        }
        else
        {
            $obj = Get-GraphObject $curObj.Object $curObj.ObjectType
            $curObj.Object = $obj.Object 
        }

        # Second check for object types that don't return scopeTag in list
        if($curObj.ObjectType.ScopeTagsReturnedInList -eq $false -and (Confirm-GraphMatchFilter $obj $filter) -eq $false)
        {        
            continue
        }

        $documentedObj = Get-ObjectDocumentation $curObj
        if($documentedObj.ErrorText)
        {
            if($txtDocumentationRawData.Text)
            {
                $txtDocumentationRawData.Text += "`n`n"    
            }

            $txtDocumentationRawData.Text += "#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" 
            $txtDocumentationRawData.Text += "`n#`n# Object: $((Get-GraphObjectName $curObj.Object $curObj.ObjectType))" 
            $txtDocumentationRawData.Text += "`n# Type: $($curObj.Object.'@OData.Type')" 
            $txtDocumentationRawData.Text += "`n#`n# Object not documented. Error:"
            $txtDocumentationRawData.Text += "`n# $(($documentedObj.ErrorText -replace "`n","`n# "))"
            $txtDocumentationRawData.Text += "`n#`n#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" 
            continue
        }
        elseif(-not $documentedObj)
        {
            continue
        }

        $curObj | Add-Member Noteproperty -Name "Documentation" -Value $documentedObj -Force 

        $tmpList += $curObj
    }
    $sourceList = $tmpList
    
    # Add each object to the documentation
    foreach($curGroupId in ($sourceList.ObjectType | Select GroupID -Unique).GroupID)
    {
        # New object group e.g. Script, Tenant, Device Configuration
        # A group matches a menu item in the protal but can contain multiple object types
        if($global:cbDocumentationType.SelectedItem.NewObjectGroup)
        {
            Write-Status "Run NewObjectGroup for $($global:cbDocumentationType.SelectedItem.Name)"
            $ret = & $global:cbDocumentationType.SelectedItem.NewObjectGroup $curGroupId
            if($ret -is [boolean] -and $ret -eq $true) { continue }
        }

        foreach($curCategory in ($sourceList | Where { $_.ObjectType.GroupID -eq $curGroupId } | Select Category -Unique).Category)
        {
            $obj = $sourceList | Where { $_.Category -eq $curCategory } | Select -First 1
            if(-not $obj) { continue }

            # New object type e.g Administrative Template, VPN profile etc.
            if($global:cbDocumentationType.SelectedItem.NewObjectType)
            {
                if($obj.ObjectType.GroupId -eq "EndpointSecurity")
                {
                    $objectTypeString = $obj.CategoryName
                }
                else
                {
                    $objectTypeString = $obj.GroupName #!!!###$obj.ObjectType.Title
                }
                $objectTypeString = (?? $objectTypeString $obj.ObjectType.Title)

                Write-Status "Run NewObjectType for $($global:cbDocumentationType.SelectedItem.Name)"
                $ret = & $global:cbDocumentationType.SelectedItem.NewObjectType $objectTypeString
                if($ret -is [boolean] -and $ret -eq $true) { continue }
            }

            #ObjectType
            foreach($tmpObj in ($sourceList | Where { $_.ObjectType.GroupID -eq $curGroupId -and $_.Category -eq $curCategory}))
            {                
                $obj = $tmpObj
                $documentedObj = $obj.Documentation
                
                if($global:cbDocumentationType.SelectedItem.CustomProcess)
                {
                    # The provider takes care of all the processing
                    Write-Status "Run CustomProcess for $($global:cbDocumentationType.SelectedItem.Name)"
                    $ret = & $global:cbDocumentationType.SelectedItem.CustomProcess $obj $documentedObj
                    if($ret -is [boolean] -and $ret -eq $true) { continue }
                }

                if($documentedObj) 
                {
                    Add-RawDataInfo $obj.Object $obj.ObjectType
    
                    $updateNotConfigured = $true
                    $notConfiguredLoc = Get-LanguageString "BooleanActions.notConfigured"
                    $notConfiguredText = ""
                    if($global:cbNotConifugredText.SelectedValue -eq "notConfigured")
                    {
                        $notConfiguredText = $notConfiguredLoc
                    }
                    elseif($global:cbNotConifugredText.SelectedValue -eq "asis")
                    {
                        $updateNotConfigured = $false 
                    }
                    
                    if($global:cbDocumentationType.SelectedItem.Process)
                    {
                        Write-Status "Process $((Get-GraphObjectName $tmpObj.Object $tmpObj.ObjectType)) ($($obj.ObjectType.Title)) - $($global:cbDocumentationType.SelectedItem.Name)"
    
                        $filteredSettings = @()
                        foreach($item in $documentedObj.Settings)
                        {
                            if(-not ($item.PSObject.Properties | Where Name -eq "RawValue") -or $documentedObj.UpdateFilteredObject -eq $false)
                            {
                                $filteredSettings = $documentedObj.Settings
                                break
                            }
                            
                            if($item.AlwaysAddValue -eq $true)
                            {
                                
                            }
                            elseif($global:chkSkipNotConfigured.IsChecked -and (($item.RawValue -isnot [array] -and ([String]::IsNullOrEmpty($item.RawValue) -or ("$($item.RawValue)" -eq "notConfigured"))) -or ($item.RawValue -is [array] -and ($item.RawValue | measure).Count -eq 0)))
                            {
                                # Skip unconfigured items e.g. properties with null values
                                # Note: This could removed configured properties if RawValue is not specified
                                continue                
                            }
                            elseif($global:chkSkipNotConfigured.IsChecked -and $documentedObj.UnconfiguredProperties -and ($documentedObj.UnconfiguredProperties | Where EntityKey -eq $item.EntityKey))
                            {
                                # Skip unconfigured items e.g. boolean with a value but Not Configured
                                continue                
                            }
    
                            if($global:chkSkipDefaultValues.IsChecked -and (($item.DefaultValue -and $item.RawValue -eq $item.DefaultValue) -or ($item.UnconfiguredValue -and $item.RawValue -eq $item.UnconfiguredValue)))
                            {
                                # Skip items that is using default or unconfiguered values                                                               
                                continue                
                            }                            
    
                            if($global:chkSkipDisabled.IsChecked -and ($item.Enabled -is [Boolean] -and ($item.Enabled -eq $false)))
                            {
                                # Skip Disabled items
                                continue
                            }
                            elseif($item.EntityKey -and ($item.Enabled -is [Boolean] -and ($item.Enabled -eq $false)))
                            {
                                if(($documentedObj.Settings | Where { $_.EntityKey -eq $item.EntityKey -and $_.Enabled -eq $true }))
                                {
                                    # Skip a disabled item if there is another item with the same property that is enabled
                                    continue
                                }
                            }
    
                            if($updateNotConfigured -and (($item.RawValue -isnot [array] -and ($null -eq $item.RawValue -or "$($item.RawValue)" -eq "" -or "$($item.RawValue)" -eq "notConfigured") -and [String]::IsNullOrEmpty($item.Value)) -or ($item.RawValue -is [array] -and ($item.RawValue | measure).Count -eq 0)))
                            {
                                $item.Value = $notConfiguredText
                            }
                            
                            if($global:chkSkipNotConfigured.IsChecked -and $item.Value -eq $notConfiguredLoc)
                            {
                                # Skip unconfigured items based on value e.g. value = Not Configured 
                                Write-Log "Skipping property $($item.Name) based on '$($notConfiguredLoc)' string value" 2
                                continue
                            }
    
                            $filteredSettings += $item
                        }
    
                        $documentedObj | Add-Member Noteproperty -Name "FilteredSettings" -Value $filteredSettings -Force 

                        & $global:cbDocumentationType.SelectedItem.Process $obj.Object $obj.ObjectType $documentedObj
                    }
                }
            }
        }
    }
    
    if($global:cbDocumentationType.SelectedItem.PostProcess)
    {
        Write-Status "Run PostProcess for $($global:cbDocumentationType.SelectedItem.Name)"
        & $global:cbDocumentationType.SelectedItem.PostProcess
    }

    if($script:offlineDocumentation -eq $true)
    {
        # Clear the dependency objects loaded for Offline documentation
        $global:LoadedDependencyObjects = $null
    }
    $script:offlineDocumentation = $false
    Write-Status ""         
}

function Confirm-MatchFilter
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
    
function Get-DocOfflineObjects
{
    param($objectName)

    if($script:offlineDocumentation -eq $false) { return }

    if($script:offlineObjects.ContainsKey($objectName))
    {
        $script:offlineObjects[$objectName]
    }
}

function Set-OutputOptionsTabStatus
{
    param($control)

    $global:tabOutputSettings.Visibility = (?: ($control.SelectedItem.Value -eq "none") "Collapsed" "Visible")

    $global:lblCustomOptions.Content = ?: $control.SelectedItem.OutputOptions ($control.SelectedItem.Name) ""
    $global:ccOutputCustomOptions.Content = $control.SelectedItem.OutputOptions

    if($control.SelectedItem.Activate)
    {
        & $control.SelectedItem.Activate
    }
}

function Invoke-DocumentObjectTypes
{
   Show-DocumentationForm -objectTypes $global:currentViewObject.ViewItems -ShowFolderSource
}

function Invoke-DocumentSelectedObjects
{
    if(-not $script:selectedObjects -or $script:selectedObjects.Count -eq 0)
    {
        [System.Windows.MessageBox]::Show("No objects added for documentation", "Documentation", "OK", "Info")
        return
    }

    Show-DocumentationForm -objects $script:selectedObjects -SelectedDocuments
}

function Add-DocOfflineObjectTypeDependencies 
{
    param($fromFolder)

    foreacH($viewItem in $global:currentViewObject.ViewItems)
    {
        foreach($dep in $viewItem.Dependencies)
        {
            Add-DocOfflineDependencies $dep $fromFolder
        }
    }
}

function Add-DocOfflineDependencies 
{
    param($objectTypeName, $fromFolder)

    if($script:offlineObjects.ContainsKey($objectTypeName)) { return }

    $tmpObjType = $global:currentViewObject.ViewItems | Where Id -eq $objectTypeName

    if($tmpObjType)
    {
        $objPath = [IO.Path]::Combine($fromFolder,$tmpObjType.Id)
        if([IO.Directory]::Exists($objPath) -eq $false)
        {
            Write-Log "Object path for $($tmpObjType.Title) ($($objPath)) not found" 2                 
        }
        else
        {                    
            $tmpObjects = Get-GraphFileObjects $objPath -ObjectType $tmpObjType
            $script:offlineObjects.Add($tmpObjType.Id, @(($tmpObjects | Select Object).Object))
        }
    }
}

function Add-DocumentationObjects
{
    param($objects)

    $itemCount = ($objects | measure).Count

    if($itemCount -eq 0) { return }

    if(-not $script:selectedObjects)
    {
        $script:selectedObjects = @()
    }
    $script:selectedObjects += $objects

    [System.Windows.MessageBox]::Show("$itemCount object(s) added to the documentation list", "Documentation", "OK", "Info")

}

function Clear-DocumentationObjects
{
    $script:selectedObjects = @()
}

#endregion

#region CSV Options

function Add-CSVOptionsControl
{
    $script:csvForm = Get-XamlObject ($global:AppRootFolder + "\Xaml\DocumentationCSVOptions.xaml") -AddVariables

    $global:cbCSVDocumentationProperties.ItemsSource = ("[ { Name: `"Simple (Name and Value)`",Value: `"simple`" }, { Name: `"Extended`",Value: `"extended`" }, { Name: `"Custom`",Value: `"custom`" }]" | ConvertFrom-Json)
    $global:cbCSVDocumentationProperties.SelectedValue = (Get-Setting "Documentation" "CSVExportProperties" "simple")
    $global:txtCSVCustomProperties.Text = (Get-Setting "Documentation" "CSVCustomDisplayProperties" "Name,Value,Category")

    $global:spCSVCustomProperties.Visibility = (?: ($global:cbCSVDocumentationProperties.SelectedValue -ne "custom") "Collapsed" "Visible")
    $global:txtCSVCustomProperties.Visibility = (?: ($global:cbCSVDocumentationProperties.SelectedValue -ne "custom") "Collapsed" "Visible")

    $global:cbCSVDelimiter.ItemsSource = @("", ",",";","-","|")
    try
    {
        $global:cbCSVDelimiter.SelectedIndex = $global:cbCSVDelimiter.ItemsSource.IndexOf((Get-Setting "Documentation" "CSVDelimiter"))
    }
    catch {}

    Add-XamlEvent $script:csvForm "browseCSVDocumentationPath" "add_click" {
        $folder = Get-Folder (Get-XamlProperty $script:csvForm "txtCSVDocumentationPath" "Text") "Select root folder for export"
        if($folder)
        {
            Set-XamlProperty $script:csvForm "txtCSVDocumentationPath" "Text" $folder
        }
    }
    
    Add-XamlEvent $script:csvForm "cbCSVDocumentationProperties" "add_selectionChanged" {
        $global:spCSVCustomProperties.Visibility = (?: ($this.SelectedValue -ne "custom") "Collapsed" "Visible")
        $global:txtCSVCustomProperties.Visibility = (?: ($this.SelectedValue -ne "custom") "Collapsed" "Visible")
    }

    $script:csvForm
}
function Invoke-CSVActivate
{
    $global:txtCSVDocumentationPath.Text = (?? (Get-Setting "" "LastUsedRoot") (Get-SettingValue "RootFolder"))
    $global:chkCSVAddObjectType.IsChecked = (Get-SettingValue "AddObjectType")
    $global:chkCSVAddCompanyName.IsChecked = (Get-SettingValue "AddCompanyName")
}

function Invoke-CSVPreProcessItems
{
    Save-Setting "Documentation" "CSVExportProperties" $global:cbCSVDocumentationProperties.SelectedValue
    Save-Setting "Documentation" "CSVCustomDisplayProperties" $global:txtCSVCustomProperties.Text
    Save-Setting "Documentation" "CSVDelimiter" $global:cbCSVDelimiter.Text
}

function Invoke-CSVProcessItem
{
    param($obj, $objectType, $documentedObj)

    if(!$documentedObj -or !$obj -or !$objectType) { return }

    $folder = Get-GraphObjectFolder $objectType (Get-XamlProperty $script:csvForm "txtCSVDocumentationPath" "Text") (Get-XamlProperty $script:csvForm "chkCSVAddObjectType" "IsChecked") (Get-XamlProperty $script:csvForm "chkCSVAddCompanyName" "IsChecked")

    try 
    {
        if([IO.Directory]::Exists($folder) -eq $false)
        {
            [IO.Directory]::CreateDirectory($folder)
        }

        $objName = Get-GraphObjectName $obj $objectType

        #BasicInfo
        #Settings
        #ComplianceActions
        #Assignments
        #DisplayProperties

        $itemsToExport = @()
        
        $params = @{}
        if($global:cbCSVDelimiter.Text)
        {
            $params.Add('Delimiter',$global:cbCSVDelimiter.Text)
        }

        if(($global:cbCSVDocumentationProperties.SelectedValue -eq 'extended' -and $documentedObj.DisplayProperties) -or 
            ($global:cbCSVDocumentationProperties.SelectedValue -eq 'custom' -and $global:txtCSVCustomProperties.Text))
        {
            if(($documentedObj.BasicInfo | measure).Count -gt 0)
            {
                $itemsToExport += ""
                $itemsToExport += "# Basic info"
                $itemsToExport += ""
                $itemsToExport += $documentedObj.BasicInfo | ConvertTo-Csv -NoTypeInformation
            }

            if(($documentedObj.FilteredSettings | measure).Count -gt 0)
            {
                $itemsToExport += ""
                $itemsToExport += "# Settings"
                $itemsToExport += ""
                if($global:cbCSVDocumentationProperties.SelectedValue -eq 'extended')
                {
                    $displayProperties = $documentedObj.DisplayProperties
                }
                else
                {
                    $displayProperties = $global:txtCSVCustomProperties.Text.Split(",")
                }
                $itemsToExport += $documentedObj.FilteredSettings | Select $displayProperties | ConvertTo-Csv -NoTypeInformation
            }

            if(($documentedObj.ApplicabilityRules | measure).Count -gt 0)
            {
                $itemsToExport += ""
                $itemsToExport += "# Applicability Rules"
                $itemsToExport += ""
                $itemsToExport += $script:applicabilityRules | Select Rule,Property,Value,Category | ConvertTo-Csv -NoTypeInformation
            }

            if(($documentedObj.ComplianceActions | measure).Count -gt 0)
            {
                $itemsToExport += ""
                $itemsToExport += "# Compliance Actions"
                $itemsToExport += ""
                $itemsToExport += $documentedObj.ComplianceActions | Select Action,Schedule,MessageTemplate,EmailCC,Category | ConvertTo-Csv -NoTypeInformation
            }

            if(($documentedObj.Assignments | measure).Count -gt 0)
            {
                if($documentedObj.Assignments[0].RawIntent)
                {
                    $properties = @("GroupMode","Group","Category","SubCategory")
                }
                elseif($documentedObj.Assignments[0].Group)
                {
                    $properties = @("GroupMode","Group","Category")
                }
                else
                {
                    $properties = @("GroupMode","Groups","Category")
                }

                $itemsToExport += ""
                $itemsToExport += "# Assignments"
                $itemsToExport += ""
                $itemsToExport += $documentedObj.Assignments  | Select $properties | ConvertTo-Csv -NoTypeInformation @params
            }
        }        
        else
        {
            $itemsToExport += $documentedObj.BasicInfo
            $itemsToExport += $documentedObj.FilteredSettings
            $itemsToExport = $itemsToExport | Select Name,Value | ConvertTo-Csv -NoTypeInformation @params
        }

        $fileName = $folder + "\$((Remove-InvalidFileNameChars $objName)).csv"
        Write-Log "Save documentation to $fileName"
        $itemsToExport | Out-File -LiteralPath $fileName -Encoding UTF8 -Force
    }
    catch 
    {
        Write-LogError "Failed to save CSV file $(($folder + "\$($objName).csv"))" $_.Exception
    }
}
#endregion

#region Invoke-ScopeTags
function Get-ScopeTagsItems
{
    param($documentationObjects, $objectType)

    if((($documentationObjects | Where { $_.Object.ObjectType.Id -eq "ScopeTags" }) | measure).Count -gt 0)
    {
        $items = @()

        $nameLabel = Get-LanguageString "Inputs.displayNameLabel"
        $descriptionLabel = Get-LanguageString "TableHeaders.description" 
        $assignmentsLabel = Get-LanguageString "TableHeaders.assignments"
        
        foreach($fullObj in ($documentationObjects | Where { $_.Object.ObjectType.Id -eq "ScopeTags" }))
        {
            $obj = $fullObj.Object

            $groupIDs, $groupInfo, $filterIds,$filtersInfo = Get-ObjectAssignments $obj

            $items += [PSCustomObject]@{
                $nameLabel = $obj.displayName
                ID = $obj.Id
                $descriptionLabel = $obj.Description
                $assignmentsLabel = ($groupInfo.displayName -join $script:objectSeparator)
            }
        }

        $items = $items | Sort -Property $nameLabel

        $documentationInfo = [PSCustomObject]@{
            TypeName = (Get-LanguageString "SettingDetails.scopeTags")
            ObjectType = $objectType
            Properties = @($nameLabel,"id", $descriptionLable, $assignmentsLabel)
            Items = $items
        }

        return $documentationInfo
    }
}
#endregion

function Get-ObjectAssignments
{
    param($obj)

    $groupIds = @()
    $groupInfo = @()
    $filterIds = @()

    foreach($assignment in $obj.assignments)
    {
        if($assignment.target.groupId -and $assignment.target.groupId -notin $groupIds)
        {
            $groupIds += $assignment.target.groupId
        }

        if($assignment.target.deviceAndAppManagementAssignmentFilterId -and ($assignment.target.deviceAndAppManagementAssignmentFilterId -notin $filterIds))
        {
            $filterIds += $assignment.target.deviceAndAppManagementAssignmentFilterId
        }
    }

    $groupInfo = $null

    if($groupIds.Count -gt 0 -and $script:offlineDocumentation -ne $true)
    {
        $ht = @{}
        $ht.Add("ids", @($groupIds | Unique))

        $body = $ht | ConvertTo-Json

        $groupInfo = (Invoke-GraphRequest -Url "/directoryObjects/getByIds?`$select=displayName,id" -Content $body -Method "Post").Value
    }
    
    if(($null -eq $groupInfo -or ($groupInfo | measure).Count -eq 0) -and $obj."@ObjectFromFile" -eq $true -and $script:migTable)
    {
        ### Get group info from mig table when documenting from file if there's no access to the environment
        $groupInfo = $script:migTable.Objects | Where Type -eq "Group" 
    }

    if($filterIds.Count -gt 0)
    {        
        if($script:offlineDocumentation -eq $true)
        {
            if($script:offlineObjects["AssignmentFilters"])
            {
                $filtersInfo = $script:offlineObjects["AssignmentFilters"] | Where { $_.Id -in $filterIds }
            }
            else
            {
                Write-Log "No assignment filters loaded for Offline documentation. Check export folder" 2
            }
        }
        else
        {
            $batchInfo = @{}
            $requests = @()
            #{"requests":[{"id":"<FilterID>","method":"GET","url":"deviceManagement/assignmentFilters/<FilterID>?$select=displayName"}]}
            foreach($filterId in $filterIds)
            {
                $requests += [PSCustomObject]@{
                    id = $filterId
                    method = "GET"
                    "url" = "deviceManagement/assignmentFilters/$($filterId)?`$select=displayName"
                }
            }
            $batchInfo = @{"requests"=$requests}
            $jsonBody = $batchInfo | ConvertTo-Json

            $filtersInfo = (Invoke-GraphRequest -Url "/`$batch" -Content $jsonBody -Method "Post").responses.body
        }
    }    
    
    @($groupIds, $groupInfo, $filterIds, $filtersInfo)
}

function Get-TableObjects
{
    param($objectTypeId)

    if(-not $objectTypeId -or $script:ObjectTypeFullTable -isnot [HashTable]) { return }

    if($script:ObjectTypeFullTable.ContainsKey($objectTypeId))
    {
        return $script:ObjectTypeFullTable[$objectTypeId]
    }
}

function Set-TableObjects
{
    param($objectInfo)

    if(-not $objectInfo.ObjectType -or $script:ObjectTypeFullTable -isnot [HashTable]) { return }

    if($script:ObjectTypeFullTable.ContainsKey($objectInfo.ObjectType.Id) -eq $false)
    {
        $script:ObjectTypeFullTable.Add($objectInfo.ObjectType.Id, $objectInfo)
    }
}

function Get-PolicyTypeName
{
    param($type, $default = $null)

    $categoryObj =  Get-TranslationFiles $type

    if($null -eq $categoryObj) { return $default }

    $lngStr = Get-LanguageString "PolicyType.$($categoryObj.PolicyTypeLanguageId)"

    if($lngStr) { return $lngStr }

    return $defult
}