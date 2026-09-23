# Intent input provider — deviceManagementIntent (Endpoint Security baselines
# and templates).
#
# Ported from old Extensions/Documentation.psm1:1580 (Invoke-TranslateIntent-
# Object + helpers). Claims @odata.type='#microsoft.graph.deviceManagementIntent'.
#
# Intent settings live under /deviceManagement/templates/{templateId}/categories
# (with $expand=settingDefinitions) and the per-intent values come from
# /deviceManagement/intents/{intentId}/categories/{catId}/settings. Each setting
# may be Simple / Collection / Complex / AbstractComplex with recursive children
# and dependency constraints that hide settings whose parents aren't configured.
#
# Live Graph dependencies (resolved via Invoke-MSGraphAPI):
#   /deviceManagement/templates/{tid}/categories?$expand=settingDefinitions
#   /deviceManagement/intents/{iid}/categories/{cid}/settings?$expand=...
#   /deviceManagement/templates/{tid}/categories/{cid}/RecommendedSettings
#
# Batch-cached on the [DocumentationContext] ($ctx.IntentCategories,
# $ctx.IntentCatRecommendedSettings) so a bulk run of N intents against the
# same template only pays the round-trips once.

function Invoke-InitializeIntentInput {
    Add-DocumentationInputProvider ([PSCustomObject]@{
        Name      = 'Intent'
        Order     = 40
        Match     = { param($PolicyObject) $PolicyObject.JsonObject.'@odata.type' -eq '#microsoft.graph.deviceManagementIntent' }
        Translate = { param($PolicyObject, $Context) Invoke-TranslateIntentObject $PolicyObject $Context }
    })
}

function Invoke-TranslateIntentObject {
    param($PolicyObject, [DocumentationContext]$Context)

    $Context.DefaultDocumentationProperties = @('Name','Value','RecommendedValue')

    $obj = $PolicyObject.JsonObject

    Add-BasicDefaultValues $PolicyObject

    $baseLineTemplates = Get-CacheObject "BaseLineTemplates"
    if(-not $baseLineTemplates)
    {
        $baseLineTemplates = (Invoke-MSGraphAPI -Url "/deviceManagement/templates").Value
        Set-CacheObject "BaseLineTemplates" $baseLineTemplates -Persistent
    }

    $baseLineTemplate = $baseLineTemplates | Where-Object Id -eq $obj.templateId
    if(-not $baseLineTemplate)
    {
        Write-Log "Could not find Baseline Template with Id $($obj.templateId)" 3
    }
    else {
        $platformType = Get-LanguageString "Platform.$($baseLineTemplate.platformType)"

        if($platformType) { Add-BasicPropertyValue (Get-LanguageString "SettingDetails.platformSupported") $platformType 'platformSupported'}

        if ($baseLineTemplate.templateSubtype -eq "none")
        {
            $templateCategoory = $baseLineTemplate.templateType
        } else {
            $templateCategoory = $baseLineTemplate.templateSubtype
        }
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.Category") (Get-IntentCategoryFromTemplateType $templateCategoory) "basicCategory"
        Add-BasicPropertyValue (Get-LanguageString "TableHeaders.policyType") $baseLineTemplate.displayName "basicPolicyType"
    }

    Add-BasicAdditionalValues $PolicyObject
    if (-not $obj.templateId) {
        Write-Log "Intent: no templateId on '$($obj.displayName)' - cannot translate settings" 2
        return
    }

    # Built-in ES template schema is generic. Seed/share the session-persistent
    # caches by reference so per-templateId/per-category writes below warm the
    # cache automatically and survive across runs (and tenant switches).
    $Context.IntentCategories = Get-CacheObject "DocIntentCategories" $Context.IntentCategories
    Set-CacheObject "DocIntentCategories" $Context.IntentCategories -Persistent
    $Context.IntentCatRecommendedSettings = Get-CacheObject "DocIntentRecommendedSettings" $Context.IntentCatRecommendedSettings
    Set-CacheObject "DocIntentRecommendedSettings" $Context.IntentCatRecommendedSettings -Persistent

    # --- Template categories (batch-cached per templateId) ---
    $categories = $Context.IntentCategories[$obj.templateId]
    if (-not $categories) {
        # Built-in Endpoint Security template schema (by templateId) is generic -
        # same on every tenant - so resolved from any connected tenant, even when
        # the source tenant of the export is gone.
        if (-not (Test-DocumentationGraphAvailable)) {
            Write-Log "Intent: no tenant connected and no cached template categories for $($obj.templateId) - settings will not render" 2
            return
        }
        try {
            $headers = @{}
            if ($Context.Language -and $Context.Language -ne 'en') { $headers['Accept-Language'] = $Context.Language }
            $resp = Invoke-MSGraphAPI -Url "/deviceManagement/templates/$($obj.templateId)/categories?`$expand=settingDefinitions" -AdditionalHeaders $headers
            $categories = @($resp.Value)
            $Context.IntentCategories[$obj.templateId] = $categories
        }
        catch {
            Write-LogError "Intent: failed to fetch template categories for $($obj.templateId)" $_.Exception
            return
        }
    }

    # Per-object setting buffer (drained at the end into Context.SettingsData
    # in dependency-respecting order via Add-IntentSettingObjectToList).
    $script:_intentObjectSettings = [System.Collections.Generic.List[object]]::new()
    $script:_intentEmittedIds     = @{}

    foreach ($category in ($categories | Sort-Object -Property displayName)) {
        # Per-intent settings for this category (skipped when the input is an
        # offline file with .settings inlined).
        $settings = $null
        if ($obj.'@ObjectFromFile' -eq $true) {
            $settings = $obj.settings
        }
        elseif (-not $Context.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
            # Source-tenant-specific: this intent's configured values by id (404s elsewhere).
            # Export path is the @ObjectFromFile branch above.
            try {
                $headers = @{}
                if ($Context.Language -and $Context.Language -ne 'en') { $headers['Accept-Language'] = $Context.Language }
                $resp = Invoke-MSGraphAPI -Url "/deviceManagement/intents/$($obj.Id)/categories/$($category.Id)/settings?`$expand=Microsoft.Graph.DeviceManagementComplexSettingInstance/Value" -AdditionalHeaders $headers
                $settings = $resp.Value
            }
            catch {
                Write-LogError "Intent: failed to fetch settings for intent=$($obj.Id) category=$($category.Id)" $_.Exception
                continue
            }
        }
        if (-not $settings) { continue }

        # Recommended settings (template-level, also batch-cached per categoryId)
        if (-not $Context.IntentCatRecommendedSettings.ContainsKey($category.Id)) {
            # Template-level recommended settings (by templateId) are generic schema.
            if (Test-DocumentationGraphAvailable) {
                try {
                    $headers = @{}
                    if ($Context.Language -and $Context.Language -ne 'en') { $headers['Accept-Language'] = $Context.Language }
                    $resp = Invoke-MSGraphAPI -Url "/deviceManagement/templates/$($obj.templateId)/categories/$($category.Id)/RecommendedSettings" -AdditionalHeaders $headers
                    $Context.IntentCatRecommendedSettings[$category.Id] = @($resp.Value)
                }
                catch {
                    Write-LogError "Intent: failed to fetch recommended settings for template=$($obj.templateId) category=$($category.Id)" $_.Exception
                    $Context.IntentCatRecommendedSettings[$category.Id] = @()
                }
            }
            else {
                $Context.IntentCatRecommendedSettings[$category.Id] = @()
            }
        }

        foreach ($settingObj in $settings) {
            Get-IntentSettingInfo $settingObj $category $settingObj.definitionId $settings $Context | Out-Null
        }
    }

    # Drain top-level settings (those with no parent and no dependencies).
    # Children/dependents get visited recursively by Add-IntentSettingObjectToList.
    $tops = $script:_intentObjectSettings | Where-Object {
        $null -eq $_.ParentId -and (($_.Dependencies | Measure-Object).Count -eq 0)
    }
    foreach ($s in $tops) {
        Add-IntentSettingObjectToList $s $Context
    }
}

# Ordered emit: respects dependency constraints (parents resolve to permitted
# values before dependent children are added) and recurses to children of any
# emitted setting.
function Add-IntentSettingObjectToList {
    param($objSetting, [DocumentationContext]$Context)

    if ($script:_intentEmittedIds.ContainsKey([string]$objSetting.Id)) { return }

    $passConstraint = $true
    $hasConstraint  = $false
    foreach ($dependencyObj in $objSetting.SettingDefinition.dependencies) {
        $dependencyItemObj = $script:_intentObjectSettings | Where-Object { $_.SettingDefinition.Id -eq $dependencyObj.definitionId } | Select-Object -First 1
        if ($dependencyObj.constraints.Count -gt 0) {
            $hasConstraint = $true
            foreach ($constraint in $dependencyObj.constraints) {
                switch ($constraint.'@odata.type') {
                    '#microsoft.graph.deviceManagementSettingBooleanConstraint' {
                        if (($null -eq $dependencyItemObj.RawValue -and $constraint.value -eq $false) -or
                            ($dependencyItemObj.RawValue -and "$($dependencyItemObj.RawValue)" -ne "$($constraint.value)")) {
                            $passConstraint = $false
                        }
                    }
                    '#microsoft.graph.deviceManagementEnumConstraint' {
                        if (-not ($constraint.values | Where-Object Value -EQ $dependencyItemObj.RawValue)) {
                            $passConstraint = $false
                        }
                    }
                    '#microsoft.graph.deviceManagementSettingIntegerConstraint' {
                        # Old code inverts the comparison — passes when value is OUT of range.
                        # Preserving the (buggy?) behavior for golden parity.
                        if ($dependencyItemObj.RawValue -ge $constraint.minimumValue -and
                            $dependencyItemObj.RawValue -le $constraint.maximumValue) {
                            $passConstraint = $false
                        }
                    }
                }
                if (-not $passConstraint) { break }
            }
        }
        else {
            # No explicit constraint — dependency just has to be "set"
            $passConstraint = ($null -ne $dependencyItemObj.RawValue -and
                               "$($dependencyItemObj.RawValue)" -ne 'NotConfigured' -and
                               "$($dependencyItemObj.RawValue)" -ne 'False')
        }
        if (-not $passConstraint) { break }
    }

    if (-not $passConstraint) { return }

    if ($hasConstraint) { $objSetting.Level = $objSetting.Level + 1 }

    # Attach recommended-value comparison (purely informational on the emitted row)
    $recommendedSetting = $Context.IntentCatRecommendedSettings[$objSetting.CategoryObject.Id] |
                          Where-Object definitionId -EQ $objSetting.SettingId | Select-Object -First 1
    if ($recommendedSetting.valueJson -and ($objSetting.ValueSet -eq $false -or
        $recommendedSetting.valueJson -ne ($objSetting.RawValue | ConvertTo-Json -Depth 50 -Compress))) {
        $objSetting | Add-Member -MemberType NoteProperty -Name 'RecommendedValue' `
                                 -Value ($recommendedSetting.valueJson | ConvertFrom-Json) -Force
    }

    $Context.AddSetting($objSetting)
    $script:_intentEmittedIds[[string]$objSetting.Id] = $true

    if ($objSetting.ValueSet -eq $false) { return }

    # Recurse: dependents (settings whose dependencies include this one)
    foreach ($depObj in ($script:_intentObjectSettings | Where-Object {
        $_.Dependencies.definitionId -eq $objSetting.SettingDefinition.Id
    })) {
        Add-IntentSettingObjectToList $depObj $Context
    }

    # Recurse: children (settings with ParentId pointing at this one and no deps)
    foreach ($depObj in ($script:_intentObjectSettings | Where-Object {
        $_.ParentId -eq $objSetting.Id -and (($_.Dependencies | Measure-Object).Count -eq 0)
    })) {
        Add-IntentSettingObjectToList $depObj $Context
    }
}

# Recursive setting parser. Builds a per-setting PSCustomObject with all the
# metadata the emit step needs, pushes it onto $script:_intentObjectSettings,
# and recurses into Complex / AbstractComplex / Collection children.
function Get-IntentSettingInfo {
    param(
        $valueObj, $category, $defId, $allSettings, [DocumentationContext]$Context,
        [switch]$SkipConvertValue, [switch]$PassThru, $parentDef = $null
    )

    $defObj = $category.settingDefinitions | Where-Object id -EQ $defId | Select-Object -First 1
    if (-not $defObj) { return }

    $itemValue     = $null
    $itemFullValue = $null

    $rawValue = if ($SkipConvertValue) { $valueObj } else { $valueObj.valueJson | ConvertFrom-Json }

    $valueSet = Get-IsIntentObjectConfigured $rawValue

    if ($valueSet -eq $false) {
        # Skip child settings
    }
    elseif ($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementCollectionSettingInstance' -or
            $defObj.'@odata.type' -eq '#microsoft.graph.deviceManagementComplexSettingDefinition' -or
            $defObj.valueType        -eq 'collection') {
        $valueArr = @()
        $elementDefObj = if ($defObj.elementDefinitionId) {
            $category.settingDefinitions | Where-Object id -EQ $defObj.elementDefinitionId | Select-Object -First 1
        } else { $defObj }

        if ($elementDefObj.propertyDefinitionIds) {
            # Each element is itself a record of N properties — emit the
            # FullValueTable so output providers can render it as a table.
            $itemFullValue = @()
            foreach ($tmpValue in $rawValue) {
                $htFullPropInfo = [ordered]@{}
                $arrValue = ''
                foreach ($propertyDefinitionId in $elementDefObj.propertyDefinitionIds) {
                    $propDefObj = $category.settingDefinitions | Where-Object id -EQ $propertyDefinitionId | Select-Object -First 1
                    if ($propDefObj.elementDefinitionId) {
                        $propDefObj = $category.settingDefinitions | Where-Object id -EQ $propDefObj.elementDefinitionId | Select-Object -First 1
                    }
                    if ($arrValue) { $arrValue = $arrValue + $Context.PropertySeparator }
                    $propName  = $propertyDefinitionId.Split('_')[-1]
                    $propValue = @()
                    foreach ($childTmpValue in $tmpValue.$propName) {
                        $propValue += Get-IntentObjectValue $propDefObj $childTmpValue
                    }
                    $colName = if ($propDefObj.displayName) { $propDefObj.displayName } else { $propName }
                    $htFullPropInfo.Add($colName, $tmpValue.$propName)
                    $arrValue = $arrValue + ($propValue -join $Context.PropertySeparator)
                }
                $itemFullValue += [PSCustomObject]$htFullPropInfo
                $valueArr += $arrValue
            }
        }
        elseif ($rawValue) {
            foreach ($tmpValue in $rawValue) {
                $valueArr += (Get-IntentObjectValue $elementDefObj $tmpValue)
            }
        }

        if ($valueArr.Count -gt 0) {
            $itemValue = $valueArr -join $Context.ObjectSeparator
        }
        $valueSet = $valueArr.Count -gt 0
    }
    elseif ($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementAbstractComplexSettingInstance' -or
            $defObj.'@odata.type'   -eq '#microsoft.graph.deviceManagementAbstractComplexSettingDefinition') {
        $tmpDef = $category.settingDefinitions | Where-Object {
            $_.id -eq $rawValue.implementationId -or $_.id -eq $rawValue.'$implementationId'
        } | Select-Object -First 1
        if ($tmpDef) {
            $itemValue = $tmpDef.displayName
        }
        else {
            $valueSet = $false
        }
    }
    else {
        $itemValue = Get-IntentObjectValue $defObj $rawValue
        if (-not $itemValue) { $valueSet = $false }
    }

    if ($valueSet -eq $false) {
        $itemValue = Get-LanguageString 'SettingDetails.notConfigured'
        $rawValue = $null
    }
    elseif (-not $itemValue) {
        $itemValue = $rawValue
    }

    $curObjectInfo = [PSCustomObject]@{
        Name                = $defObj.displayName
        Description         = $defObj.description
        Category            = $category.displayName
        CategoryDescription = $category.description
        CategoryObject      = $category
        Value               = $itemValue
        FullValueTable      = $itemFullValue
        RawValue            = $rawValue
        SettingDefinition   = $defObj
        Dependencies        = $defObj.dependencies
        ValueSet            = $valueSet
        Id                  = [Guid]::NewGuid()
        ParentId            = $null
        SettingId           = $defObj.Id
        ParentSettingId     = $parentDef.Id
        Level               = 0
    }
    $script:_intentObjectSettings.Add($curObjectInfo)

    if ($valueSet -eq $false) {
        # Skip children if value not set
    }
    elseif ($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementComplexSettingInstance' -or
            $defObj.'@odata.type'   -eq '#microsoft.graph.deviceManagementComplexSettingDefinition') {
        if ($valueObj.Value) {
            $isValueSet = $false
            if ($defObj.propertyDefinitionIds) {
                foreach ($childDefId in $defObj.propertyDefinitionIds) {
                    $childSetting = $valueObj.Value | Where-Object DefinitionId -EQ $childDefId | Select-Object -First 1
                    if ($childSetting) {
                        $objValueInfo = Get-IntentSettingInfo $childSetting $category $childSetting.definitionId $allSettings $Context -PassThru -parentDef $defObj
                        $objValueInfo.ParentId = $curObjectInfo.Id
                        if (($objValueInfo.RawValue -is [bool] -and $objValueInfo.RawValue -eq $true) -or
                            ($objValueInfo.RawValue -is [string] -and -not [string]::IsNullOrEmpty($objValueInfo.RawValue) -and
                                $objValueInfo.RawValue -ne 'notConfigured' -and -not [string]::IsNullOrEmpty($objValueInfo.Value)) -or
                            ($objValueInfo.RawValue -isnot [bool] -and $objValueInfo.RawValue -isnot [string])) {
                            $isValueSet = $true
                        }
                    }
                }
            }
            else {
                foreach ($childSetting in $valueObj.Value) {
                    $objValueInfo = Get-IntentSettingInfo $childSetting $category $childSetting.definitionId $allSettings $Context -PassThru -parentDef $defObj
                    $objValueInfo.ParentId = $curObjectInfo.Id
                }
                $isValueSet = $true
            }
        }
        elseif ($rawValue -and $defObj.propertyDefinitionIds) {
            $isValueSet = $false
            $isDefault  = $true
            foreach ($childDefId in $defObj.propertyDefinitionIds) {
                $propName = $childDefId.Split('_')[-1]
                $objValueInfo = Get-IntentSettingInfo $rawValue.$propName $category $childDefId $allSettings $Context -SkipConvertValue -PassThru -parentDef $defObj
                if ($objValueInfo.ValueSet -eq $true) { $isValueSet = $true }
                if ($objValueInfo.SettingDefinition.constraints -and
                    $objValueInfo.SettingDefinition.constraints[0].'@odata.type' -eq '#microsoft.graph.deviceManagementEnumConstraint' -and
                    ($objValueInfo.SettingDefinition.constraints[0].values | Measure-Object).Count -gt 0) {
                    if ($objValueInfo.SettingDefinition.constraints[0].values[0].value -ne $rawValue.$propName) {
                        $isDefault = $false
                    }
                }
                elseif ($objValueInfo.SettingDefinition.valueType -eq 'string') {
                    if ($null -ne $rawValue.$propName) { $isDefault = $false }
                }
                elseif ($objValueInfo.SettingDefinition.valueType -eq 'boolean') {
                    if ($false -ne $rawValue.$propName) { $isDefault = $false }
                }
                $objValueInfo.ParentId = $curObjectInfo.Id
            }
            if ($isDefault) { $isValueSet = $false }
        }
        else {
            $isValueSet = $false
        }

        $curObjectInfo.Value = if ($isValueSet) { 'Configure' } else { Get-LanguageString 'SettingDetails.notConfigured' }
        $curObjectInfo.ValueSet = $isValueSet
        $curObjectInfo.FullValueTable = $null
    }
    elseif (($valueObj.'@odata.type' -eq '#microsoft.graph.deviceManagementAbstractComplexSettingInstance' -or
             $defObj.'@odata.type'   -eq '#microsoft.graph.deviceManagementAbstractComplexSettingDefinition') -and
            $rawValue -and $tmpDef) {
        foreach ($childDefId in $tmpDef.propertyDefinitionIds) {
            $propName = $childDefId.Split('_')[-1]
            $objValueInfo = Get-IntentSettingInfo $rawValue.$propName $category $childDefId $allSettings $Context -SkipConvertValue -PassThru -parentDef $defObj
            $objValueInfo.ParentId = $curObjectInfo.Id
        }
    }

    if ($PassThru) { $curObjectInfo }
}

# Translates a raw setting value via its definition (enum / boolean / raw passthrough).
function Get-IntentObjectValue {
    param($defObj, $rawValue)

    if ($defObj.constraints.'@odata.type' -eq '#microsoft.graph.deviceManagementEnumConstraint') {
        $tmpOption = $defObj.constraints.Values | Where-Object value -EQ $rawValue | Select-Object -First 1
        if (-not $tmpOption -and $null -eq $rawValue) {
            # No defaultValue on the setting definition — fall back to first option.
            # Old-code wart preserved for golden parity.
            $tmpOption = $defObj.constraints.Values[0]
        }
        return $tmpOption.displayName
    }
    elseif ($defObj.valueType -eq 'boolean') {
        if ($rawValue -eq 'True') { return (Get-LanguageString 'SettingDetails.yes') }
        return $null
    }
    return $rawValue
}

# Hook for custom "is configured?" rules. Old code always returns true; kept as
# a function so type-specific overrides can be wired in later.
function Get-IsIntentObjectConfigured {
    param($obj)
    return $true
}

# Template-type to friendly category-name lookup. Used by BasicInfo "Type"
# row when the input provider lands templateType resolution in v2; for now
# only exported so handlers can reuse the mapping.
function Get-IntentCategoryFromTemplateType {
    param([string]$TemplateType)

    if (-not $TemplateType) {
        Write-Log 'Get-IntentCategoryFromTemplateType called with empty TemplateType' 2
        return $null
    }

    # Captured before the prefix is stripped: whether the family was security-shaped
    # is what decides if failing to map it is worth reporting (see the default arm).
    $isSecurityFamily = $TemplateType.StartsWith('endpointSecurity') -or $TemplateType -match 'baseline'

    if ($TemplateType.StartsWith('endpointSecurity')) {
        $TemplateType = $TemplateType.Substring(16)
    }

    switch ($TemplateType) {
        'accountProtection'      { return (Get-LanguageString 'SecurityTemplate.accountProtection') }
        'antivirus'              { return (Get-LanguageString 'SecurityTemplate.antivirus') }
        'diskEncryption'         { return (Get-LanguageString 'SecurityTemplate.diskEncryption') }
        'endpointDetectionReponse' { return (Get-LanguageString 'SecurityTemplate.eDR') }
        'attackSurfaceReduction' { return (Get-LanguageString 'SecurityTemplate.aSR') }
        'firewall'               { return (Get-LanguageString 'SecurityTemplate.firewall') }
        { $_ -in @('securityBaseline','baseline','advancedThreatProtectionSecurityBaseline','microsoftEdgeSecurityBaseline') } {
            return (Get-LanguageString 'Titles.securityBaselines')
        }
        # Not a security template, but it reaches this mapper the same way: the
        # Settings Catalog provider asks for a category name for every family it
        # documents, and the Apple ADE enrollment policies are this one. Without an
        # arm here the row read 'enrollmentConfiguration'. PolicySet.deviceEnrollment
        # is an existing key, so the label localizes with everything else.
        'enrollmentConfiguration' { return (Get-LanguageString 'PolicySet.deviceEnrollment') }
        default {
            # Only a security-shaped family is expected to resolve here. The Settings
            # Catalog provider (Get-IntentCategoryName) calls this for EVERY
            # templateFamily and documents the raw value when it does not map, so a
            # family like 'enrollmentConfiguration' is a normal outcome rather than a
            # problem - warning about it once per policy put a wall of yellow in the
            # log of any tenant with Apple ADE policies and buried the real signal.
            if ($isSecurityFamily) {
                Write-Log "Could not translate Intent Template type $TemplateType" 2
            }
            else {
                Write-LogDebug "No Intent category mapping for template family '$TemplateType'; documented as-is"
            }
            return $TemplateType
        }
    }
}

Invoke-InitializeIntentInput
