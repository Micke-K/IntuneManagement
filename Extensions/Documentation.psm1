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
    '1.0.1'
}

function Invoke-InitializeModule
{
    # Make sure we add the default Output types
    Add-OutputType
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
    $global:dgObjects.add_selectionChanged({
        ##Set-XamlProperty $global:dgObjects.Parent "btnDocument" "IsEnabled" (?: ($global:dgObjects.SelectedItem -eq $null) $false $true)
        #$itemSelected = ($global:dgObjects.ItemsSource | Where IsSelected -eq $true).Count -ge 0 -or $global:dgObjects.SelectedItem

        Set-XamlProperty $global:dgObjects.Parent "btnDocument" "IsEnabled" (?: ($global:dgObjects.SelectedItem -eq $null) $false $true)
    })

    $button.Add_Click({ 

        $objects = ?? ($global:dgObjects.ItemsSource | Where IsSelected -eq $true) $global:dgObjects.SelectedItem 

        Show-DocumentationForm -Objects $objects
    })    

    $global:spSubMenu.RegisterName($button.Name, $button)

    $global:spSubMenu.Children.Insert(0, $button)
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

    Write-Status "Get documentation info for $((Get-GraphObjectName $documentationObj.Object $documentationObj.ObjectType)) ($($documentationObj.ObjectType.Title))"

    $status = $null
    $inputType = "Settings"    

    if(-not $script:scopeTags)
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
        $properties = @("Name","Value","Category","RawValue","SettingId","Description")
    }
    #endregion
    #region Administrative Templates
    elseif($type -eq "#microsoft.graph.groupPolicyConfiguration")
    {
        Invoke-TranslateADMXObject $obj $objectType | Out-Null
        $properties = @("Name","Status","Value","Category","CategoryPath","RawValue","ValueWithLabel","Created","Modified", "Class", "DefinitionId")
        $defaultDocumentationProperties =  @("Name","Status","Value")
        $updateFilteredObject = $false
        $inputType = "Property"
    }
    #endregion
    #region Profile Types e.g. DeviceConfiguration Policies etc
    elseif($type)
    {
        $inputType = "Property"
        $processed = $true
        <#
        if([IO.File]::Exists(($global:AppRootFolder + "\Documentation\ObjectInfo\$($obj.'@OData.Type').json")))
        {
            # Process object based on OData type
            $processed = Invoke-TranslateCustomProfileObject $obj "$($obj.'@OData.Type')"
        }
        elseif($objectType -and [IO.File]::Exists(($global:AppRootFolder + "\Documentation\ObjectInfo\#$($objectType.Id).json")))
        {
            # Process object based on Intune Object Type ($objectType)
            # '#' is added to front of name to distinguish manually created files from generated files
            $processed = Invoke-TranslateCustomProfileObject $obj "#$($objectType.Id)"
        }
        else
        {
            # Process objects based on generated Category Files and ObjectCategories.json
            $processed = Invoke-TranslateProfileObject $obj
        }
        #>
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

    [PSCustomObject]@{
        BasicInfo = $script:objectBasicInfo
        Settings = $script:objectSettingsData
        ComplianceActions = $script:objectComplianceActionData
        ApplicabilityRules = $script:applicabilityRules
        Assignments = $script:objectAssignments
        DisplayProperties = $properties
        DefaultDocumentationProperties = $defaultDocumentationProperties
        ErrorText = $status
        InputType = $inputType
        UpdateFilteredObject = $updateFilteredObject        
    }
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

    $script:DocumentationLanguage = "en"        
    $script:objectSeparator = [System.Environment]::NewLine
    $script:propertySeparator = ","

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

function Get-ObjectTypeString
{
    param($obj, $objectType)

    $objTypeId = ?? $objectType.GroupId $objectType
    
    if($objTypeId -eq "DeviceConfiguration")
    {
        return (Get-LanguageString "SettingDetails.deviceConfigurationTitle")
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
    elseif($objTypeId -eq "TenantAdmin")
    {
        return (Get-LanguageString "Titles.tenantAdmin")
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

function Add-BasicAdditionalValues
{
    param($obj, $objectType)

    if($obj.createdDateTime)
    {
        $tmpDate = ([DateTime]::Parse($obj.createdDateTime))
        Add-BasicPropertyValue (Get-LanguageString "Inputs.createdDateTime") "$($tmpDate.ToLongDateString()) $($tmpDate.ToLongTimeString())"
    }

    if($obj.lastModifiedDateTime)
    {
        $tmpDate = ([DateTime]::Parse($obj.lastModifiedDateTime))
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.lastModified") "$($tmpDate.ToLongDateString()) $($tmpDate.ToLongTimeString())"
    }
    elseif($obj.modifiedDateTime)
    {
        $tmpDate = ([DateTime]::Parse($obj.modifiedDateTime))
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.lastModified") "$($tmpDate.ToLongDateString()) $($tmpDate.ToLongTimeString())"
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
    if(($obj.roleScopeTagIds | measure).Count -gt 0)
    {
        foreach($scopeTagId in $obj.roleScopeTagIds)
        {
            $scopeTagObj = $script:scopeTags | Where Id -eq $scopeTagId
            if($scopeTagObj)
            {
                $objScopeTags += $scopeTagObj.displayName
            }
        }
    }
    if($objScopeTags.Count -gt 0)
    {
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.scopeTags") ($objScopeTags -join $script:objectSeparator)
    }
}

function Get-AllEntityTypes
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
        Get-AllEntityTypes $node.Name $xml $hashTable
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

    foreach($definitionValue in $definitionValues)
    {
        if(-not $definitionValue.definition -and $definitionValues.'definition@odata.bind')
        {
            $definition = Invoke-GraphRequest -Url $definitionValue.'definition@odata.bind' -ODataMetadata "minimal" @params
            if($definition)
            {
                $definitionValue | Add-Member -MemberType NoteProperty -Name "definition" -Value $definition
            }            
        }

        $categoryObj = $script:admxCategories | Where { $definitionValue.definition.id -in ($_.definitions.id) }
        $category = $script:admxCategories.definitions | Where { $definitionValue.definition.id -in ($_.id) }
        # Get presentation values for the current settings (with presentation object included)
        if($definitionValue.presentationValues -or $obj.'@CompareObject' -eq $true) #$definitionValue.'definition@odata.bind')
        {
            # Documenting exported json
            #$presentationValues = (Invoke-GraphRequest -Url "$($definitionValue.'definition@odata.bind')/presentations?`$expand=presentation"  -ODataMetadata "minimal").value
            $presentationValues = $definitionValue.presentationValues
        }
        else
        {
            $presentationValues = (Invoke-GraphRequest -Url "/deviceManagement/groupPolicyConfigurations/$($obj.id)/definitionValues/$($definitionValue.id)/presentationValues?`$expand=presentation"  -ODataMetadata "minimal" @params).value
        }

        $value = $null
        $rawValues = @()
        $values = @()
        $valuesWithLabel = @()

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
    Add-BasicPropertyValue (Get-LanguageString "SettingDetails.platformSupported") $platformType
    Add-BasicAdditionalValues $obj $objectType
    
    $params = @{}
    ## Set language
    if($script:DocumentationLanguage)
    {
        $params.Add("AdditionalHeaders", @{"Accept-Language"=$script:DocumentationLanguage})
    }

    $cfgSettings = (Invoke-GraphRequest "/deviceManagement/configurationPolicies('$($obj.Id)')/settings?`$expand=settingDefinitions&top=1000" -ODataMetadata "minimal" @params).Value   

    if(-not $global:cfgCategories)
    {
        $global:cfgCategories = (Invoke-GraphRequest "/deviceManagement/configurationCategories?`$filter=platforms has 'windows10' and technologies has 'mdm'" -ODataMetadata "minimal" @params).Value
    }

    $categories = @{}
    foreach($cfgSetting in $cfgSettings)
    {
        $defObj = $cfgSetting.settingDefinitions | Where id -eq $cfgSetting.settingInstance.settingDefinitionId
        if(-not $defObj -or $categories.ContainsKey($defObj.categoryId)) { continue }

        $catObj = $global:cfgCategories | Where Id -eq $defObj.categoryId 
        $rootCatObj = $global:cfgCategories | Where Id -eq $catObj.rootCategoryId

        $catSettings = Invoke-GraphRequest "/deviceManagement/configurationSettings?`$filter=categoryId eq '$($defObj.categoryId)' and applicability/platform has 'windows10' and applicability/technologies has 'mdm'" -ODataMetadata "minimal" @params

        $categories.Add($defObj.categoryId, (New-Object PSObject -Property @{ 
            Category=$catObj
            Settings=$catSettings
            RootCategory=$rootCatObj
         }))
    }

    Add-SettingsSetting $obj $objectType $categories $cfgSettings
}

function Add-SettingsSetting
{
    param($obj, $objectType, $categories, $cfgSettings, $settigsDefs = $null)

    foreach($cfgSetting in $cfgSettings)
    {
        $children = $null
        $skipAdd = $false

        $cfgInstance =  ?? $cfgSetting.settingInstance $cfgSetting
        if($cfgSetting.settingDefinitions)
        {
            $settigsDefs = $cfgSetting.settingDefinitions
        }
        $defaultValue = $null
        $rawValue=$null
        $rawJsonValue = $null
        $defObj = $settigsDefs | Where id -eq $cfgInstance.settingDefinitionId
        $catObj = $categories[$defObj.categoryId]

        if($cfgInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance')
        {
            $rawValue = $cfgInstance.choiceSettingValue.value            
            $itemValue = ($defObj.Options | Where itemId -eq $rawValue).displayName
            if($defObj.defaultOptionId)
            {
                $defaultValue = ($defObj.Options | Where itemId -eq $defObj.defaultOptionId).displayName
            }
            $children = $cfgInstance.choiceSettingValue.children
        }        
        elseif($cfgInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationChoiceSettingCollectionInstance')
        {
            $itemValues = @()
            $itemRawValues = @()
            foreach($colObj in $cfgInstance.simpleSettingCollectionValue)
            {
                $tmpValue = $colObj.value            
                $itemValues += ($defObj.Options | Where itemId -eq $tmpValue).displayName
                $itemRawValues += $tmpValue
            }

            $rawValue = $itemValues -join $script:propertySeparator
            $itemValue = $itemRawValues -join $script:propertySeparator

            if($defObj.defaultOptionId)
            {
                $defaultValue = ($defObj.Options | Where itemId -eq $defObj.defaultOptionId).displayName
            }                
        }
        elseif($cfgInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance')
        {
            $itemValue = $cfgInstance.simpleSettingValue.value 
            $rawValue = $itemValue
            if($defObj.defaultValue.value)
            {
                $defaultValue = $defObj.defaultValue.value
            }
        }
        elseif($cfgInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationSimpleSettingCollectionInstance')
        {
            $itemValues = @()
            $rawValue = $cfgInstance.simpleSettingCollectionValue
            foreach($colObj in $cfgInstance.simpleSettingCollectionValue)
            {
                $itemValues += $colObj.value            
            }

            if($defObj.defaultValue.value)
            {
                $defaultValue = $defObj.defaultValue.value
            }
            $itemValue = $itemValues -join $script:propertySeparator
        }
        elseif($cfgInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationGroupSettingInstance')
        {
            # This will skip adding the group itself as enabled...
            # It will only add information about the children
            Add-SettingsSetting $obj $objectType $categories $cfgInstance.groupSettingValue.children $settigsDefs
            continue
        }
        elseif($cfgInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance')
        {
            # ToDo: Fix support for other child types
            # This assumes that all children are deviceManagementConfigurationSimpleSettingInstance
            $objItems = @()
            foreach($childObj in $cfgInstance.groupSettingCollectionValue)
            {
                $objProps = @()
                foreach($childId in $defObj.childIds)
                {
                    $childSetting = $childObj.children | Where settingDefinitionId -eq $childId
                    if($childSetting)
                    {
                        $objProps += $childSetting.simpleSettingValue.value 
                    }
                }
                $objItems += $objProps -join $script:propertySeparator
            }
            $itemValue = $objItems -join $script:objectSeparator
            $rawValue = $itemValue
            $rawJsonValue = $cfgInstance.groupSettingCollectionValue | ConvertTo-Json -Depth 20 -Compress
        }
        else
        {
            Write-Log "Unsupported setting type: $($cfgInstance.'@odata.type')"    
        }

        if(!$rawJsonValue -and $rawValue)
        {
            $rawJsonValue = $rawValue | ConvertTo-Json -Depth 20 -Compress
        }

        $script:objectSettingsData += New-Object PSObject -Property @{ 
            Name=$defObj.displayName
            Description=$defObj.description
            Category=$catObj.Category.displayName
            CategoryDescription=$catObj.Category.description
            Value=$itemValue
            RawValue=$rawValue
            RawJsonValue=$rawJsonValue
            DefaultValue=$defaultValue
            RootCategory=$catObj.RootCategory.displayName
            CategoryObject=$catObj
            SettingId=$cfgInstance.settingDefinitionId
        }

        if($children)
        {
            Add-SettingsSetting $obj $objectType $categories $children $settigsDefs
        }
    }
}

#endregion

#region Intent Objects (Endpoint Security)

function Get-IntentCategory
{
    param($templateType)

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
    elseif($templateType -eq "???")
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
    elseif($templateType -eq "securityBaseline" -or 
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
        if($obj.'@CompareObject' -ne $true)
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
    
    foreach($objSetting in ($script:objectSettings | Where { $_.ParentId -eq $null -and ($_.Dependecies | measure).Count -eq 0 }))
    {
        #if($objSetting.Dependecies) { continue }
        Add-IntentSettingObjectToList $objSetting
    }    
}

function Add-IntentSettingObjectToList
{
    param($objSetting)

    if(($script:objectSettingsData | Where Id -eq $objSetting.Id)) { return }

    $passConstraint = $true
    foreach($dependencyObj in $objSetting.SettingDefinition.dependencies)
    {
        $dependencyItemObj = ($script:objectSettings | Where { $_.SettingDefinition.Id -eq $dependencyObj.definitionId })
        if($dependencyObj.constraints.Count -gt 0)
        {                    
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

    $script:objectSettingsData += $objSetting

    if($objSetting.ValueSet -eq $false) { return }

    foreach($depObj in ($script:objectSettings | Where { $_.Dependecies.definitionId -eq $objSetting.SettingDefinition.Id }))
    {
        Add-IntentSettingObjectToList $depObj
    }

    foreach($depObj in ($script:objectSettings | Where { $_.ParentId -eq $objSetting.Id -and ($_.Dependecies | measure).Count -eq 0} ))
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

            foreach($tmpValue in $rawValue)
            {
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
                    $arrValue = ($arrValue + ($propValue -join $script:propertySeparator))
                }
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
    elseif($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementAbstractComplexSettingInstance')
    {
        $tmpDef = $category.settingDefinitions | Where id -eq $rawValue.'$implementationId'
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
        Value=$itemValue
        RawValue=$rawValue
        CategoryObject=$category
        SettingDefinition = $defObj
        Dependecies = $defObj.dependencies
        ValueSet = $valueSet
        Id=[Guid]::NewGuid() #(([Guid]::NewGuid()).Guid)
        ParentId = $null
        SettingId = $defObj.Id # ToDo: Must have parent Id as well e.g. Firewall settings in Win10 baseline
        ParentSettingId = $parentDef.Id
    }

    $script:objectSettings += $curObjectInfo    

    if($valueSet -eq $false)
    {
        ; # Skip children if value is not set...
    }
    elseif($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementComplexSettingInstance')
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
    }
    elseif($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementAbstractComplexSettingInstance' -and $rawValue -and $tmpDef )
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
        # Shuld only be one file. Compliance policies might have more
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
            Invoke-TranslateSection $obj $categoryObj."$($fi.BaseName)" $objInfo            
        }
        catch 
        {
            Write-LogError "Failed tp translate file $($fi.Name)" $_.Exception    
        }
    }

    return $true
}

function Get-LanguageString
{
    param($string, $defaultValue = $null, [switch]$IgnoreMissing)

    if(-not $script:languageStrings)
    {
        $fileContent = Get-Content ($global:AppRootFolder + "\Documentation\Strings-$($script:DocumentationLanguage).json") -Encoding UTF8
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

        if($expression.value -eq $null)
        {
            # Value not specified. Check if the property is set
            $tmpRet = $tmpProp.Value -ne $null
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

    foreach($prop in $sectionObject)
    {
        $valueData = $null
        $value = $null
        $valueSet = $false
        $useParentProp = $false
        
        #if($prop.enabled -eq $false -and $objInfo.ShowDisabled -ne $true) { continue }

        if((Invoke-VerifyCondition $obj $prop $objInfo) -eq $false) 
        {
            Write-LogDebug "Condition returned false: $(($prop.Condition | ConvertTo-Json -Depth 10 -Compress))" 2
            continue
        }

        $obj = Get-CustomPropertyObject $obj $prop

        $rawValue = $obj."$($prop.entityKey)"

        if($prop.dataType -eq 8)
        {
            if($prop.nameResourceKey -eq "LearnMore") { continue }
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
                    Write-Log "SubCategpry ignored based on length: $tmpStr" 2
                }
            }
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
                $script:CurrentSubCategory = (Get-LanguageString (?: $prop.nameResourceKey.Contains(".") $prop.nameResourceKey "SettingDetails.$($prop.nameResourceKey)"))
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

                Add-PropertyInfo $prop $value $rawValue
            }
        }
        elseif([String]::IsNullOrEmpty($prop.entityKey) -eq $false)
        {            
            $valueSet = ($rawValue -ne $null)
            $skipChildren = $false
            if($rawValue -eq $null -and ![String]::IsNullOrEmpty($prop.unconfiguredValue) -and $global:chkSetUnconfiguredValue.IsChecked)
            {
                $propValue = $prop.unconfiguredValue
            }
            elseif($rawValue -eq $null -and ![String]::IsNullOrEmpty($prop.defaultValue) -and $global:chkSetDefaultValue.IsChecked)
            {
                $propValue = $prop.defaultValue
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
                    $cert = Invoke-GraphRequest -URL $script:currentObject."$($prop.entityKey)@odata.navigationLink" -ODataMetadata "minimal" -NoError
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
                    $value = $obj."$((?? $prop.filenameEntityKey $prop.EntityKey))"
                    $valueData = $obj."$((?? $prop.dataEntityKey $prop.EntityKey))"
                }
                elseif($prop.dataType -eq 2) # Multiline string e.g. XML file 
                {
                    $value = $obj."$((?? $prop.filenameEntityKey $prop.EntityKey))"
                    $valueData = $obj."$((?? $prop.dataEntityKey $prop.EntityKey))"
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
    param($prop, $value, $originalValue, $jsonValue)

    if($prop.Category -eq "1000")
    {
        Add-BasicPropertyValue (Get-LanguageString $prop.nameResourceKey) $value
        return
    }

    $script:objectSettingsData += Get-PropertyInfo $prop $value $originalValue $jsonValue
}

function Add-PropertyInfoObject
{
    param($propInfo)

    if($propInfo -eq $null) { return }

    $script:objectSettingsData += $propInfo
}

function Get-PropertyInfo
{
    param($prop,$value,$originalValue, $jsonValue)

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

    if(!$jsonValue -and $rawValue -ne $null -and "$($rawValue)" -ne "")
    {
        $jsonValue = $rawValue | ConvertTo-Json -Depth 10 -Compress
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
        DefaultValue=$prop.defaultValue
        UnconfiguredValue=$prop.unconfiguredValue
        Enabled=$prop.Enabled 
        EntityKey=$prop.EntityKey
    }
}

function Invoke-ChildSections
{
    param($obj, $sectionObject)
    
    $objTmp = Get-CustomChildObject $obj $sectionObject

    Invoke-TranslateSection $objTmp $sectionObject.Children $objInfo -Parent $sectionObject
    
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

    if($obj."$($prop)" -eq $null -or ($obj."$($prop)" -is [String] -and $obj."$($prop)" -eq "notConfigured")) { return }
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
        Get-LanguageString "BooleanActions.notConfigured"
    }
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

            Add-PropertyInfo $prop $optionValue -originalValue $propValue
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
    foreach($item in $propValue)
    {
        $itemValues = @()
        foreach($column in $prop.Columns)
        {
            if($column.metadata.entityKey -eq "unusedForSingleItems")
            {
                $itemValues += $item
            }
            elseif($column.metadata.entityKey -eq $prop.entityKey -and ($prop.Columns | measure).Count -eq 1)
            {
                # Some tables has the same EntityKey for the table and the columen. That will generate the wrong value
                $itemValues += $item 
            }
            elseif(($prop.Columns | measure).Count -eq 1 -and $item."$($column.metadata.entityKey)" -eq $null -and $obj."$($column.metadata.entityKey)" -eq $null -and $item -is [String])
            {
                # Not sure how correct this is but some tables has one column with and EntityKey but the objects is a string list
                $itemValues += $item
            }
            else
            {
                $itemValues += (?? $item."$($column.metadata.entityKey)" $obj."$($column.metadata.entityKey)")
            }
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
        if((-not $prop.nameResourceKey -or $prop.nameResourceKey -eq "Empty") -and $prop.columns[0].metadata.nameResourceKey)
        {
            Add-PropertyInfo $prop.columns[0].metadata ($items -join $script:objectSeparator) $propValue
        }
        else
        {
            Add-PropertyInfo $prop ($items -join $script:objectSeparator) $propValue
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
            $additionalNotifications = $null

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
                }
                else
                {
                    $notificationTemplate = Get-LanguageString "ScheduledAction.Notification.noneSelected"
                }

                if($actionConfig.notificationMessageCCList.Count -gt 0)
                {
                    $additionalNotifications = ((Get-LanguageString "ScheduledAction.Notification.numSelected") -f $actionConfig.notificationMessageCCList.Count)
                }
                else
                {
                    $additionalNotifications = Get-LanguageString "ScheduledAction.Notification.noneSelected"
                }
            }

            $script:objectComplianceActionData += New-Object PSObject -Property @{
                Action = $actionType
                Schedule = $schedule
                MessageTemplate = $notificationTemplate
                EmailCC = $additionalNotifications
                Category=$category
                RawJsonValue=($actionConfig | ConvertTo-Json -Depth 20 -Compress)
            }
        }
    }
}
#endregion

#region Assignments
function Invoke-TranslateAssignments
{
    param($obj)

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

    if($groupIds.Count -gt 0)
    {
        $ht = @{}
        $ht.Add("ids", @($groupIds | Unique))

        $body = $ht | ConvertTo-Json

        $groupInfo = (Invoke-GraphRequest -Url "/directoryObjects/getByIds?`$select=displayName,id" -Content $body -Method "Post").Value
    }

    if($filterIds.Count -gt 0)
    {
        $batchInfo = @{}
        $requests = @()
        #{"requests":[{"id":"<FilterID>","method":"GET","url":"deviceManagement/assignmentFilters/<FilterID>?$select=displayName"}]}
        foreach($filterId in $filterIds)
        {
            $requests += [PSCustomObject]@{
                id = $filterIds
                method = "GET"
                "url" = "deviceManagement/assignmentFilters/$($filterId)?`$select=displayName"
            }
        }
        $batchInfo = @{"requests"=$requests}
        $jsonBody = $batchInfo | ConvertTo-Json

        $filtersInfo = Invoke-GraphRequest -Url "/`$batch" -Content $jsonBody -Method "Post"
    }
    
    foreach($assignment in $obj.assignments)
    {
        $groupMode = $null
        $groupName = $null

        $filterName = $null
        $filterMode = $null

        if($assignment.Intent)
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
        }
        elseif($assignment.target.'@odata.type' -eq "#microsoft.graph.allDevicesAssignmentTarget")
        {
            $groupName = Get-LanguageString "SettingDetails.allUsers"
        }
        elseif($assignment.target.'@odata.type' -eq "#microsoft.graph.allLicensedUsersAssignmentTarget")
        {
            $groupName = Get-LanguageString "SettingDetails.allDevices"
        }
        else
        {
            $groupName = "unknown" # Should not get here!    
        }

        if(($assignment.target.PSObject.Properties | Where Name -eq "deviceAndAppManagementAssignmentFilterId"))
        {        
            $filterName = $noFilter
            $filterMode = $noFilter
    
            if($assignment.target.deviceAndAppManagementAssignmentFilterId -and $filtersInfo.responses)
            {
                $filtersObj = $filtersInfo.responses | Where Id -eq $assignment.target.deviceAndAppManagementAssignmentFilterId
                if($filtersObj.body.displayName)
                {
                    $filterName = $filtersObj.body.displayName
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
                    $assignmentSettingProps.Add($settingProp, $assignment.settings.$settingProp)
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
                RawJsonValue = ($assignment | ConvertTo-Json -Depth 20 -Compress)
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
        $SelectedDocuments)

    $objectList = @()
    
    if($objects)
    {
        if($SelectedDocuments -eq $true)
        {
            $objectList += $objects
        }
        else
        {
            $objects | ForEach-Object {
                $item = [PSCustomObject]@{
                    IsSelected = $true
                    Title =  (Get-GraphObjectName $_.Object $_.ObjectType) 
                    Object = $_.Object
                    ObjectType = $_.ObjectType
                }
                $objectList += $item
            }
        }
        $sourceType = "Objects"
    }
    elseif($objectTypes)
    {
        foreach($groupId in ($objectTypes | Select GroupId -Unique).GroupId)
        {
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

    $global:chkSetUnconfiguredValue.IsChecked = ((Get-Setting "Documentation" "SetUnconfiguredValue" "true") -ne "false")
    $global:chkSetDefaultValue.IsChecked = ((Get-Setting "Documentation" "SetDefaultValue" "false") -ne "false")

    $notConfiguredItems = "[ { Name: `"Not configured (Localized)`",Value: `"notConfigured`" }, { Name: `"Empty`",Value: `"empty`" }, { Name: `"Don't change`",Value: `"asis`" }]" | ConvertFrom-Json
    $global:cbNotConifugredText.ItemsSource = $notConfiguredItems
    $global:cbNotConifugredText.SelectedValue = (Get-Setting "Documentation" "NotConfiguredText" "")

    $global:chkSkipNotConfigured.IsChecked = ((Get-Setting "Documentation" "SkipNotConfigured" "false") -ne "false")
    $global:chkSkipDefaultValues.IsChecked = ((Get-Setting "Documentation" "SkipDefaultValues" "false") -ne "false")
    $global:chkSkipDisabled.IsChecked = ((Get-Setting "Documentation" "SkipDisabled" "true") -ne "false")

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

        $txtDocumentationRawData.Text = ""

        $global:intentCategories = $null
        $global:catRecommendedSettings = $null
        $global:intentCategoryDefs = $null
        $global:cfgCategories = $null

        $script:DocumentationLanguage = ?? $global:cbDocumentationLanguage.SelectedValue "en"        
        $script:objectSeparator = ?? $global:cbDocumentationObjectSeparator.SelectedValue ([System.Environment]::NewLine)
        $script:propertySeparator = ?? $global:cbDocumentationPropertySeparator.SelectedValue ","

        Save-Setting "Documentation" "OutputType" $global:cbDocumentationType.SelectedValue        

        Save-Setting "Documentation" "Language" $script:DocumentationLanguage
        Save-Setting "Documentation" "ObjectSeparator" $script:objectSeparator
        Save-Setting "Documentation" "PropertySeparator" $script:propertySeparator

        Save-Setting "Documentation" "SetUnconfiguredValue" $global:chkSetUnconfiguredValue.IsChecked
        Save-Setting "Documentation" "SetDefaultValue" $global:chkSetDefaultValue.IsChecked
    

        Save-Setting "Documentation" "SkipNotConfigured" $global:chkSkipNotConfigured.IsChecked
        Save-Setting "Documentation" "SkipDefaultValues" $global:chkSkipDefaultValues.IsChecked
        Save-Setting "Documentation" "SkipDisabled" $global:chkSkipDisabled.IsChecked

        Save-Setting "Documentation" "NotConfiguredText" $global:cbNotConifugredText.SelectedValue

        Get-CustomIgnoredCategories $obj

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
                }

                if($groupSourceList.Count -eq 0) { contnue }

                if($curObjectType -eq "EndpointSecurity")
                {
                    $catName = Get-IntentCategory $tmpObj.Category
                    $tmpObj | Add-Member Noteproperty -Name "CategoryName" -Value $catName -Force
                    $sortProps = @("CategoryName","displayName")
                }
                else
                {
                    $sortProps = @((?? $objectType.NameProperty "displayName"))
                }
                $sourceList += $groupSourceList | Sort-Object -Property $sortProps
            }
        }
        elseif($global:grdDocumentObjects.Tag -eq "ObjectTypes")
        {
            $sourceList = @()
            foreach($objGroup in ($global:grdDocumentObjects.ItemsSource | Where IsSelected -eq $true))
            {
                $groupSourceList = @()
                foreach($objectType in ($global:currentViewObject.ViewItems | Where GroupId -eq $objGroup.GroupId))
                {                    
                    Write-Status "Get $($objectType.Title) objects"

                    $url = $objectType.API
                    if($objectType.QUERYLIST)
                    {
                        $url = "$($url.Trim())?$($objectType.QUERYLIST.Trim())"
                    }
                
                    $graphObjects = @(Get-GraphObjects -Url $url -property $objectType.ViewProperties -objectType $objectType)
                
                    if($objectType.PostListCommand)
                    {
                        $graphObjects = & $objectType.PostListCommand $graphObjects $objectType
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
                    $sortProps = @((?? $objectType.NameProperty "displayName"))
                }
                $sourceList += $groupSourceList | Sort-Object -Property $sortProps
            }
        }
        else
        {
            return  
        }

        if($global:cbDocumentationType.SelectedItem.PreProcess)
        {
            Write-Status "Run PreProcess for $($global:cbDocumentationType.SelectedItem.Name)"
            & $global:cbDocumentationType.SelectedItem.PreProcess
        }

        $tmpCurObjectType = $null
        foreach($tmpObj in ($sourceList))
        {
            $obj = Get-GraphObject $tmpObj.Object $tmpObj.ObjectType

            if($obj)
            {
                $documentedObj = Get-ObjectDocumentation $obj

                if($documentedObj.ErrorText)
                {
                    $txtDocumentationRawData.Text += "#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" 
                    $txtDocumentationRawData.Text += "`n#`n# Object: $((Get-GraphObjectName $obj.Object $obj.ObjectType))" 
                    $txtDocumentationRawData.Text += "`n# Type: $($obj.Object.'@OData.Type')" 
                    $txtDocumentationRawData.Text += "`n#`n# Object not documented. Error:"
                    $txtDocumentationRawData.Text += "`n# $(($documentedObj.ErrorText -replace "`n","`n# "))"
                    $txtDocumentationRawData.Text += "`n#`n#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" 
                    continue
                }

                if($global:cbDocumentationType.SelectedItem.CustomProcess)
                {
                    # The provider takes care of all the processing
                    Write-Status "Run CustomProcess for $($global:cbDocumentationType.SelectedItem.Name)"
                    & $global:cbDocumentationType.SelectedItem.CustomProcess $obj $documentedObj
                    continue
                }

                if($tmpCurObjectType -ne $obj.ObjectType.GroupId)
                {
                    if($global:cbDocumentationType.SelectedItem.NewObjectType)
                    {
                        Write-Status "Run NewObjectType for $($global:cbDocumentationType.SelectedItem.Name)"
                        & $global:cbDocumentationType.SelectedItem.NewObjectType $obj $documentedObj
                    }
                    $tmpCurObjectType = $obj.ObjectType.GroupId
                }

                if($documentedObj) 
                {
                    Add-RawDataInfo $obj.Object $obj.ObjectType

                    $updateNotConfigured = $true
                    $notConfiguredText = ""
                    if($global:cbNotConifugredText.SelectedValue -eq "notConfigured")
                    {
                        $notConfiguredText = Get-LanguageString "BooleanActions.notConfigured"
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
                            if(-not ($item.PSObject.Properties | Where Name -eq RawValue) -or $documentedObj.UpdateFilteredObject -eq $false)
                            {
                                $filteredSettings = $documentedObj.Settings
                                break
                            }
                            
                            if($global:chkSkipNotConfigured.IsChecked -and (([String]::IsNullOrEmpty($item.RawValue) -or $item.RawValue -eq "notConfigured")))
                            {
                                # Skip unconfigured items
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

                            if($updateNotConfigured -and ($item.RawValue -eq $null -or "$($item.RawValue)" -eq "" -or "$($item.RawValue)" -eq "notConfigured") -and [String]::IsNullOrEmpty($item.Value))
                            {
                                $item.Value = $notConfiguredText
                            }                            

                            $filteredSettings += $item
                        }

                        $documentedObj | Add-Member Noteproperty -Name "FilteredSettings" -Value $filteredSettings -Force 

                        & $global:cbDocumentationType.SelectedItem.Process $obj.Object $obj.ObjectType $documentedObj
                    }
                }
            }
        }

        if($global:cbDocumentationType.SelectedItem.PostProcess)
        {
            Write-Status "Run PostProcess for $($global:cbDocumentationType.SelectedItem.Name)"
            & $global:cbDocumentationType.SelectedItem.PostProcess
        }
        
        Write-Status ""
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

        if($script:objectSettingsData.Count -gt 0)
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

    Set-OutputOptionsTabStatus $global:cbDocumentationType

    Show-ModalForm "Intune Documentation" $script:docForm -HideButtons
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
   Show-DocumentationForm -objectTypes $global:currentViewObject.ViewItems
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
                else
                {
                    $properties = @("GroupMode","Groups","Category")
                }

                $itemsToExport += ""
                $itemsToExport += "# Assignments"
                $itemsToExport += ""
                $itemsToExport += $documentedObj.Assignments  | Select $properties | ConvertTo-Csv -NoTypeInformation
            }
        }        
        else
        {
            $itemsToExport += $documentedObj.BasicInfo
            $itemsToExport += $documentedObj.FilteredSettings
            $itemsToExport = $itemsToExport | Select Name,Value | ConvertTo-Csv -NoTypeInformation
        }

        $itemsToExport | Out-File ($folder + "\$($objName).csv") -Encoding UTF8 -Force
    }
    catch 
    {
        Write-LogError "Failed to save CSV file $(($folder + "\$($objName).csv"))" $_.Exception
    }
}
#endregion