# Settings Catalog walker.
#
# Ported from old Extensions/Documentation.psm1:1210 (Add-SettingsSetting,
# ~230 LOC). Recursive walker over the deviceManagementConfigurationSetting
# tree — handles 6 settingInstance variants:
#   - SimpleSettingInstance              (string/int value)
#   - ChoiceSettingInstance              (single dropdown, may have child settings)
#   - ChoiceSettingCollectionInstance    (multi-select dropdown)
#   - GroupSettingCollectionInstance     (table-like rows of grouped sub-settings)
#   - SimpleSettingCollectionInstance    (list of simple values)
#   - GroupSettingInstance               (single group container — emits only children)
#
# Settings catalog state on the context:
#   $ctx.CachedCfgSettings  - settingDefinitionId -> full definition object
#   $ctx.CfgCategories      - flat list of category objects
#   $script:_curSettingsCatologPolicy  - per-policy buffer of settingInfo rows
#     (drained by the input provider into $ctx.SettingsData in category order)

$script:_curSettingsCatologPolicy = @()

function Reset-SettingsCatalogPolicyBuffer {
    $script:_curSettingsCatologPolicy = @()
}

function Get-SettingsCatalogPolicyBuffer {
    return $script:_curSettingsCatologPolicy
}

function Add-SettingsSetting {
    param(
        $SettingInstance,
        $SettingsDefs,
        [int]$ItemLevel = 0,
        [switch]$SkipAdd
    )

    if (-not $SettingInstance) { return }

    $ctx = Get-CurrentDocumentationContext

    $defaultValue = $null
    $tableValue   = $null
    $value        = $null
    $rawValue     = $null
    $rawJsonValue = $null
    $show         = $true
    $childSettings = @()

    # Look up the settings definition: prefer inline ($expand=settingDefinitions
    # exports), then context cache, then live Graph as last resort. The live
    # endpoint (configurationSettings/{id}) is GENERIC schema - identical on every
    # tenant - so it is gated only on connectivity (Test-DocumentationGraphAvailable),
    # NOT on SourceTenantUnavailable: documenting an export while signed into a
    # different tenant must still resolve setting names.
    $settingsDef = $null
    if ($SettingsDefs) {
        $settingsDef = $SettingsDefs | Where-Object id -EQ $SettingInstance.settingDefinitionId | Select-Object -First 1
    }
    if (-not $settingsDef -and $SettingInstance.settingDefinitionId) {
        if ($ctx.CachedCfgSettings.ContainsKey($SettingInstance.settingDefinitionId)) {
            $settingsDef = $ctx.CachedCfgSettings[$SettingInstance.settingDefinitionId]
        }
        elseif (Test-DocumentationGraphAvailable) {
            try {
                $settingsDef = Invoke-MSGraphAPI -Url "/deviceManagement/configurationSettings/$($SettingInstance.settingDefinitionId)" -AdditionalHeaders (Get-DocAcceptLanguageHeaders $ctx)
                if ($settingsDef) {
                    $ctx.CachedCfgSettings[$SettingInstance.settingDefinitionId] = $settingsDef
                }
            }
            catch {
                Write-LogError "Failed to fetch settings catalog definition for $($SettingInstance.settingDefinitionId)" $_.Exception
            }
        }
    }

    # Category lookup: root category becomes Category, leaf becomes SubCategory
    $categoryDef = $null
    $objCategory = $null
    $subCategory = $null
    if ($settingsDef.categoryId) {
        $categoryDef = $ctx.CfgCategories | Where-Object Id -EQ $settingsDef.categoryId | Select-Object -First 1
        if ($categoryDef -and $settingsDef.categoryId -ne $categoryDef.rootCategoryId) {
            $objCategory = $ctx.CfgCategories | Where-Object Id -EQ $categoryDef.rootCategoryId | Select-Object -First 1
            $subCategory = $categoryDef
        }
        else {
            $objCategory = $categoryDef
        }
    }

    $settingName = ''
    $settingDescription = ''
    if ($settingsDef.displayName) {
        $settingName = $settingsDef.displayName.Trim([Environment]::NewLine).Trim("`n")
    }
    if ($settingsDef.description) {
        $settingDescription = $settingsDef.description.Trim([Environment]::NewLine).Trim("`n")
    }

    $settingInfo = [PSCustomObject]@{
        SettingId             = $settingsDef.Id
        SettingKey            = ''
        SettingName           = $settingsDef.Name
        Name                  = $settingName
        Description           = $settingDescription
        CategoryId            = $objCategory.id
        Category              = $objCategory.displayName
        CategoryDefinition    = $objCategory
        SubCategory           = $subCategory.displayName
        SubCategoryDefinition = $subCategory
        Value                 = $null
        RawValue              = $null
        RawJsonValue          = $null
        TableValue            = $null
        DefaultValue          = $null
        Level                 = $ItemLevel
        Parent                = $null
        Show                  = $show
        Type                  = $SettingInstance.'@odata.type'
        PropertyIndex         = 0
        RowIndex              = 0
        ChildSettings         = @()
    }

    if (-not $SkipAdd) {
        $script:_curSettingsCatologPolicy += $settingInfo
    }

    switch ($SettingInstance.'@odata.type') {

        '#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance' {
            # Single dropdown
            $rawValue = $SettingInstance.choiceSettingValue.value
            $opt = $settingsDef.Options | Where-Object itemId -EQ $rawValue | Select-Object -First 1
            $value = $opt.displayName
            if ($settingsDef.defaultOptionId) {
                $defaultValue = ($settingsDef.Options | Where-Object itemId -EQ $settingsDef.defaultOptionId).displayName
            }
            # Children added to the buffer (NOT -SkipAdd) so the HTML output's
            # flat row iterator emits them with `Level` padding under the
            # parent. Old code at Documentation.psm1:1300 declared the
            # -SkippAdd switch but never honored it, so children were always
            # added — matching that behavior here. See [[group-setting-collection-children]].
            foreach ($childSetting in $SettingInstance.choiceSettingValue.children) {
                $tmp = Add-SettingsSetting $childSetting $SettingsDefs ($ItemLevel + 1)
                if ($tmp) { $tmp.Parent = $settingInfo; $settingInfo.ChildSettings += $tmp }
            }
        }

        '#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance' {
            # Single primitive value
            $value = $SettingInstance.simpleSettingValue.value
            $rawValue = $value
            if ($settingsDef.defaultValue.value) {
                $defaultValue = $settingsDef.defaultValue.value
            }
        }

        '#microsoft.graph.deviceManagementConfigurationChoiceSettingCollectionInstance' {
            # Multi-select dropdown
            $itemValues = @()
            $itemRawValues = @()
            foreach ($colObj in $SettingInstance.choiceSettingCollectionValue) {
                $itemRawValues += $colObj.value
                $opt = $settingsDef.Options | Where-Object itemId -EQ $colObj.Value | Select-Object -First 1
                $itemValues += $opt.displayName
            }
            $value        = $itemValues -join $ctx.PropertySeparator
            $rawValue     = $itemRawValues -join $ctx.PropertySeparator
            $rawJsonValue = $SettingInstance.choiceSettingCollectionValue | ConvertTo-Json -Depth 50 -Compress
            if ($settingsDef.defaultOptionId) {
                $defaultValue = ($settingsDef.Options | Where-Object itemId -EQ $settingsDef.defaultOptionId).displayName
            }
        }

        '#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance' {
            # Table-like rows of grouped sub-settings — group row itself isn't shown
            $settingInfo.Show = $false
            $rowIndex = 1
            foreach ($groupSettingCollection in $SettingInstance.groupSettingCollectionValue) {
                $childArr = @()
                # Endpoint Security templates supply $settingsDefs.id; pure Settings
                # Catalog uses $settingsDef.childIds. Old code at L1347-1354.
                $childIds = if ($ctx.CurrentObject.templateReference.templateId -and $SettingsDefs) {
                    $SettingsDefs.id
                } else {
                    $settingsDef.childIds
                }
                foreach ($childId in $childIds) {
                    $childSetting = $groupSettingCollection.children | Where-Object settingDefinitionId -EQ $childId | Select-Object -First 1
                    if (-not $childSetting) { continue }
                    # Children added to buffer (no -SkipAdd) so the HTML output's
                    # flat-row iterator can render each one with `Level` padding —
                    # the parent itself has Show=false above, so only the
                    # children are visible. Without this, the entire group
                    # vanishes from output (the Linux 'Allowed Distros' regression).
                    $tmp = Add-SettingsSetting $childSetting $SettingsDefs ($ItemLevel + 1)
                    if ($tmp) {
                        $tmp.Parent   = $childSettings
                        $tmp.RowIndex = $rowIndex
                        $childSettings += $tmp
                        $childArr += $tmp
                        if (($settingsDef.childIds | Measure-Object).Count -gt 1) {
                            $tmp.PropertyIndex = $childArr.Count
                        }
                    }
                }
                $settingInfo.ChildSettings += [PSCustomObject]@{
                    Id       = $rowIndex++
                    Type     = $groupSettingCollection.'@odata.type'
                    Settings = $childArr
                }
            }
            $rawJsonValue = $SettingInstance.groupSettingCollectionValue | ConvertTo-Json -Depth 50 -Compress
        }

        '#microsoft.graph.deviceManagementConfigurationSimpleSettingCollectionInstance' {
            # List of primitive values
            $itemValues = @()
            foreach ($colObj in $SettingInstance.simpleSettingCollectionValue) {
                $itemValues += $colObj.value
            }
            if ($settingsDef.defaultValue.value) { $defaultValue = $settingsDef.defaultValue.value }
            $value        = $itemValues -join $ctx.PropertySeparator
            $rawValue     = $itemValues -join $ctx.PropertySeparator
            $rawJsonValue = $SettingInstance.simpleSettingCollectionValue | ConvertTo-Json -Depth 50 -Compress
        }

        '#microsoft.graph.deviceManagementConfigurationGroupSettingInstance' {
            # Single group container — group itself isn't emitted, only children
            $settingInfo.Show = $false
            foreach ($groupSettingValue in $SettingInstance.groupSettingValue) {
                foreach ($childSetting in $groupSettingValue.children) {
                    # Same rationale as the GroupSettingCollection case above —
                    # children must reach the buffer (no -SkipAdd) so they
                    # render in HTML output once the Show=false parent is dropped.
                    $tmp = Add-SettingsSetting $childSetting $SettingsDefs ($ItemLevel + 1)
                    if ($tmp) { $tmp.Parent = $settingInfo; $settingInfo.ChildSettings += $tmp }
                }
            }
            $rawJsonValue = $SettingInstance.groupSettingValue | ConvertTo-Json -Depth 50 -Compress
        }

        default {
            Write-Log "Unhandled settings catalog instance type: $($SettingInstance.'@odata.type')" 2
            return
        }
    }

    if (-not $rawJsonValue -and $rawValue) {
        $rawJsonValue = $rawValue | ConvertTo-Json -Depth 50 -Compress
    }

    $settingInfo.Value        = $value
    $settingInfo.RawValue     = $rawValue
    $settingInfo.RawJsonValue = $rawJsonValue
    $settingInfo.DefaultValue = $defaultValue

    return $settingInfo
}

# Resolve a Settings Catalog payload - Collection(deviceManagementConfigurationSetting) -
# into documentation rows, ordered by (Category, SubCategory).
#
# THE single implementation. Two payload shapes carry settings-catalog settings:
#   deviceManagement/configurationPolicies              (Settings Catalog policies)
#   deviceAppManagement/targetedManagedAppConfigurations (the "Settings catalog"
#                                                         step of a MAM app config)
# The MAM handler used to keep its own copy of this block, and it had drifted:
# it omitted the configurationCategories fetch below, so category/subcategory
# grouping silently collapsed on any run that had not already documented a
# Settings Catalog policy (making the output order-dependent). Both callers now
# go through here.
#
# Rows are returned rather than pushed onto the context, so the caller decides
# whether they belong in the main settings table or in a table of their own.
function Get-SettingsCatalogDocumentationRows
{
    param(
        $Settings,
        [DocumentationContext]$Context
    )

    $cfgSettings = @($Settings)
    if ($cfgSettings.Count -eq 0) { return @() }

    # Generic schema caches (session-persistent, shared by reference so later
    # writes by the walker warm the cache automatically). Definitions are generic
    # Intune schema, so they persist across runs and tenant switches.
    $Context.CachedCfgSettings = Get-CacheObject "DocCfgSettingDefinitions" $Context.CachedCfgSettings
    Set-CacheObject "DocCfgSettingDefinitions" $Context.CachedCfgSettings -Persistent

    $Context.CfgCategories = Get-CacheObject "CfgCategories" (@())

    # Generic schema (configurationCategories) - same on every tenant - so gated
    # only on connectivity, not on SourceTenantUnavailable. Without this the
    # walker cannot resolve a row's category and the nesting disappears.
    if (-not ($Context.CfgCategories | Where-Object { $_.settingUsage -eq 'configuration' }) -and (Test-DocumentationGraphAvailable)) {
        try {
            Write-Log "Cache Settings Catalog configurationCategories"
            $resp = Invoke-MSGraphAPI -Url "/deviceManagement/configurationCategories" -ODataMetadata 'minimal' -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context)
            $Context.CfgCategories += @($resp.Value)
            Set-CacheObject "CfgCategories" $Context.CfgCategories -Persistent
        }
        catch {
            Write-LogError 'Failed to fetch configuration categories' $_.Exception
        }
    }

    # Seed the definition cache from inline settingDefinitions on each setting
    foreach ($cfgSetting in $cfgSettings) {
        if (-not $cfgSetting.settingDefinitions) { continue }
        $defObj = $cfgSetting.settingDefinitions | Where-Object id -EQ $cfgSetting.settingInstance.settingDefinitionId | Select-Object -First 1
        if ($defObj -and -not $Context.CachedCfgSettings.ContainsKey($defObj.Id)) {
            $Context.CachedCfgSettings[$defObj.Id] = $defObj
        }
    }

    # Walk each top-level setting into the shared buffer
    Reset-SettingsCatalogPolicyBuffer
    foreach ($cfgSetting in $cfgSettings) {
        Add-SettingsSetting $cfgSetting.settingInstance $cfgSetting.settingDefinitions | Out-Null
    }

    # Drain the buffer in (Category, SubCategory) order - this grouping is what
    # produces the portal's nesting in the rendered table.
    $buffer = Get-SettingsCatalogPolicyBuffer
    $unique = $buffer |
              Select-Object @{ l='CategoryID';    e={ $_.CategoryDefinition.Id    } },
                            @{ l='SubCategoryID'; e={ $_.SubCategoryDefinition.Id } } -Unique

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($pair in $unique) {
        $matching = $buffer | Where-Object {
            $_.CategoryDefinition.Id    -eq $pair.CategoryID -and
            $_.SubCategoryDefinition.Id -eq $pair.SubCategoryID
        }
        foreach ($row in $matching) {
            if ($row.Show -eq $false) { continue }
            [void]$rows.Add($row)
        }
    }

    return $rows.ToArray()
}
