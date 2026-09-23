
$script:CoreObjectProperties = @(
    'displayName', 'name', 'description', 'id', 'version',
    'createdDateTime', 'lastModifiedDateTime', 'modifiedDateTime',
    'roleScopeTagIds', 'roleScopeTags'
)

Add-AppEvent "CompareColumnVisibilityChanged"

# Thin registration wrappers so other subsystems can contribute providers/types
# without touching [CompareRegistry] directly (mirrors Add-Documentation* /
# Register-AuthProvider).
function Register-CompareProvider {
    param([object]$Provider)
    [CompareRegistry]::RegisterProvider($Provider)
}
function Register-CompareOutputProvider {
    param([object]$Provider)
    [CompareRegistry]::RegisterOutputProvider($Provider)
}
function Register-CompareComparisonType {
    param([PSCustomObject]$Type)
    [CompareRegistry]::RegisterComparisonType($Type)
}

# ---- Compare subsystem registration (runs once at module load) ----
# Registration used to live in Initialize-CompareModule, wired to the WPF-only
# "AppInitialized" event - so the Avalonia backend never populated the compare
# catalog (its combos came up empty). Doing it here at module load means BOTH
# backends get it. [CompareRegistry] is the source of truth; the $script:compare*
# variables below are kept as LIVE VIEWS onto the registry lists (same List
# references) so the existing UI combos keep working unchanged and automatically
# pick up late registrations - notably the documentation subsystem's "doc" type,
# which self-registers after this file loads (see Internal/Documentation.ps1),
# replacing the old Get-Command "Get-GraphDocumentation" reach-in that lived here.
[CompareRegistry]::Reset()

[CompareRegistry]::RegisterProvider([CompareExportFilesProvider]::new())
[CompareRegistry]::RegisterProvider([CompareIntuneWithExportProvider]::new())
[CompareRegistry]::RegisterProvider([CompareNamedObjectsProvider]::new())
[CompareRegistry]::RegisterProvider([CompareExportedFoldersProvider]::new())

[CompareRegistry]::RegisterOutputProvider([CompareCSVOutputProvider]::new())
[CompareRegistry]::RegisterOutputProvider([CompareJsonOutputProvider]::new())

# "Property" is the built-in comparison type. No Compare scriptblock on the "doc"
# type (contributed by documentation) because documentation compare needs the
# POLICY WRAPPERS (JsonObject + PolicyType), not the raw JSON generic dispatch
# passes; Compare-PolicyObjects routes Value "doc" straight to
# Compare-ObjectsBasedonDocumentation with the wrappers.
[CompareRegistry]::RegisterComparisonType([PSCustomObject]@{
    Name             = "Property"
    Value            = "property"
    Compare          = { param($O1, $O2) Compare-ObjectsBasedonProperty $O1 $O2 }
    RemoveProperties = @('Category', 'SubCategory')
})

# Live views onto the registry lists for the UI combos (see comment above).
$script:compareProviders       = [CompareRegistry]::Providers
$script:comparisonTypes        = [CompareRegistry]::ComparisonTypes
$script:compareOutputProviders = [CompareRegistry]::OutputProviders

# Non-registry compare state (fixed UI enum + per-run caches).
$script:CompareProviderOptionsCache = $null
$script:compareRuntimeOptions       = @{}
$script:defaultCompareProps = [Collections.Generic.List[String]]@('ObjectName', 'Id', 'Type', 'Category', 'SubCategory', 'Property', 'Value1', 'Value2', 'Match')
$script:compareOutputTypes = @(
    [PSCustomObject]@{ Name = "One file for each object type"; Value = "objectType" },
    [PSCustomObject]@{ Name = "One file for all objects";      Value = "all" }
)

function Set-CompareRuntimeOptions
{
    param(
        [string]$CompareType,
        $CompareDefinition,
        $IgnoreCoreProperties,
        [string]$SaveType,
        [CompareOutputProviderBase]$OutputProvider,
        [string]$CsvDelimiter,
        [string]$ObjectSeparator,
        $SkipAssignments
    )

    if($null -eq $script:compareRuntimeOptions) { $script:compareRuntimeOptions = @{} }

    foreach($key in $PSBoundParameters.Keys)
    {
        $script:compareRuntimeOptions[$key] = $PSBoundParameters[$key]
    }
}

function Get-CompareRuntimeOption
{
    param([string]$Name, $Default = $null)

    if($script:compareRuntimeOptions -and $script:compareRuntimeOptions.ContainsKey($Name))
    {
        return $script:compareRuntimeOptions[$Name]
    }
    return $Default
}

function Get-CompareTypeDefinition
{
    param([string]$CompareType, $CompareDefinition)

    if($CompareDefinition) { return $CompareDefinition }

    $runtimeDefinition = Get-CompareRuntimeOption "CompareDefinition"
    if($runtimeDefinition) { return $runtimeDefinition }

    if(-not $CompareType) { $CompareType = Get-CompareRuntimeOption "CompareType" "property" }
    $definition = [CompareRegistry]::FindComparisonType($CompareType)
    if($definition) { return $definition }

    return [PSCustomObject]@{
        Name             = "Property"
        Value            = "property"
        Compare          = { param($O1, $O2) Compare-ObjectsBasedonProperty $O1 $O2 }
        RemoveProperties = @('Category', 'SubCategory')
    }
}

function Get-CompareTypeValue
{
    param([string]$CompareType, $CompareDefinition)

    $definition = Get-CompareTypeDefinition $CompareType $CompareDefinition
    if($definition -and $definition.Value) { return $definition.Value }
    return "property"
}

function Get-CompareTypeRemoveProperties
{
    param([string]$CompareType, $CompareDefinition)

    $definition = Get-CompareTypeDefinition $CompareType $CompareDefinition
    if($definition -and $definition.RemoveProperties) { return @($definition.RemoveProperties) }
    return @()
}

function ConvertTo-CompareJson
{
    param($Value)
    if($null -eq $Value) { return "" }
    return (Remove-ComparePropertyNoise $Value | ConvertTo-Json -Depth 10 -Compress)
}

# True for keys that are server-assigned noise inside a property VALUE:
# nested entity `id` keys (row ids on sub-objects like
# scheduledActionConfigurations) and OData link/annotation keys. `@odata.type`
# (including 'prop@odata.type' hints) is real configuration and is kept.
# The policy's own top-level id is a core property handled separately.
function Test-ComparePropertyNoiseKey
{
    param([string]$Key)

    if($Key -eq 'id') { return $true }
    # Server-stamped timestamps on NESTED sub-objects (e.g. agreement file
    # localizations) differ between otherwise identical configurations. The
    # policy's own top-level timestamps are core properties handled separately.
    if($Key -in @('createdDateTime', 'lastModifiedDateTime', 'modifiedDateTime')) { return $true }
    if($Key -notlike '*@odata.*') { return $false }
    if($Key -like '*@odata.type') { return $false }
    return $true
}

# Deep-clone a property value dropping server-assigned noise, so two policies
# with identical CONFIGURATION compare equal even when Graph embedded
# object-specific ids/links (@odata.id, editLink, navigationLink,
# associationLink, context) in the nested payload - e.g. an original vs its
# copy, or an exported file vs the live object.
function Remove-ComparePropertyNoise
{
    param($Value)

    if($null -eq $Value -or $Value -is [string] -or $Value.GetType().IsValueType) { return $Value }

    if($Value -is [System.Collections.IDictionary]) {
        $out = [ordered]@{}
        foreach($key in $Value.Keys) {
            if(Test-ComparePropertyNoiseKey ([string]$key)) { continue }
            $out[[string]$key] = Remove-ComparePropertyNoise $Value[$key]
        }
        return [PSCustomObject]$out
    }

    if($Value -is [System.Collections.IEnumerable] -and $Value -isnot [PSCustomObject]) {
        $items = [System.Collections.Generic.List[object]]::new()
        foreach($item in $Value) {
            [void]$items.Add((Remove-ComparePropertyNoise $item))
        }
        return $items.ToArray()
    }

    if($Value -is [PSCustomObject]) {
        $out = [ordered]@{}
        foreach($prop in $Value.PSObject.Properties) {
            if(Test-ComparePropertyNoiseKey $prop.Name) { continue }
            $out[$prop.Name] = Remove-ComparePropertyNoise $prop.Value
        }
        return [PSCustomObject]$out
    }

    return $Value
}

function Add-CompareProperty
{
    param($Name, $Value1, $Value2, $Category = $null, $SubCategory = $null, $Match = $null, [switch]$Skip)

    $v1 = if($null -eq $Value1) { "" } else { $Value1.ToString().Trim('"') }
    $v2 = if($null -eq $Value2) { "" } else { $Value2.ToString().Trim('"') }

    $matchValue = if($null -ne $Match) { [bool]$Match } else { [bool]($v1 -ceq $v2) }
    if($Skip -eq $true) { $matchValue = $null }

    $compare = [PSCustomObject]@{
        PropertyName = $Name
        Object1Value = $v1
        Object2Value = $v2
        Category     = $Category
        SubCategory  = $SubCategory
        Match        = $matchValue
    }

    $script:compareProperties += $compare
}

function Get-PropertyByPath
{
    param($Obj, [string]$Path)
    $parts   = $Path -split '\.'
    $current = $Obj
    foreach($part in $parts)
    {
        if($null -eq $current) { return $null }
        $current = $current.$part
    }
    return $current
}

function Get-SettingsCatalogValueText
{
    param($Instance)
    if(-not $Instance) { return "" }
    $odataType = "$($Instance.'@odata.type')"
    switch -Wildcard ($odataType)
    {
        "*SimpleSettingInstance" {
            if($Instance.simpleSettingValue) { return "$($Instance.simpleSettingValue.value)" }
        }
        "*SimpleSettingCollectionInstance" {
            if($Instance.simpleSettingCollectionValue)
            {
                return (($Instance.simpleSettingCollectionValue | ForEach-Object { "$($_.value)" }) -join '; ')
            }
        }
        "*ChoiceSettingInstance" {
            if($Instance.choiceSettingValue)
            {
                $val = "$($Instance.choiceSettingValue.value)"
                if($Instance.choiceSettingValue.children)
                {
                    $childVals = @()
                    foreach($ch in $Instance.choiceSettingValue.children)
                    {
                        $cv = Get-SettingsCatalogValueText $ch
                        if($cv) { $childVals += $cv }
                    }
                    if($childVals) { return "$val [$($childVals -join '; ')]" }
                }
                return $val
            }
        }
        "*ChoiceSettingCollectionInstance" {
            if($Instance.choiceSettingCollectionValue)
            {
                return (($Instance.choiceSettingCollectionValue | ForEach-Object { "$($_.value)" }) -join '; ')
            }
        }
        "*GroupSettingCollectionInstance" {
            if($Instance.groupSettingCollectionValue)
            {
                $parts = @()
                foreach($grp in $Instance.groupSettingCollectionValue)
                {
                    if($grp.children)
                    {
                        foreach($ch in $grp.children)
                        {
                            $cv = Get-SettingsCatalogValueText $ch
                            if($cv) { $parts += $cv }
                        }
                    }
                }
                return ($parts -join '; ')
            }
        }
    }
    return ($Instance | ConvertTo-Json -Depth 10 -Compress)
}

function Get-SettingsCatalogSettingValue
{
    param($Setting)
    if(-not $Setting) { return "" }
    $instance = if($Setting.settingInstance) { $Setting.settingInstance } else { $Setting }
    Get-SettingsCatalogValueText $instance
}

function Get-SettingsCatalogSettingKey
{
    param($Setting)
    if(-not $Setting) { return $null }
    $defId = Get-PropertyByPath $Setting "settingInstance.settingDefinitionId"
    if($defId) { return "$defId" }
    return $null
}

function Get-SettingsCatalogSettingCategory
{
    param($Setting)
    if(-not $Setting) { return "" }
    $defId = Get-PropertyByPath $Setting "settingInstance.settingDefinitionId"
    if(-not $defId) { return "" }
    # settingDefinitionId is shaped like "device_vendor_msft_policy_config_<area>_<setting>"
    # Use the segment after "policy_config" if present, otherwise the segment after "vendor_msft".
    $segments = "$defId" -split '_'
    $configIdx = [Array]::IndexOf($segments, "config")
    if($configIdx -ge 0 -and $configIdx + 1 -lt $segments.Length)
    {
        return $segments[$configIdx + 1]
    }
    if($segments.Length -ge 1) { return $segments[0] }
    return ""
}

function Get-IntentSettingValue
{
    param($Setting)
    if(-not $Setting) { return "" }
    # valueJson is the authoritative serialization for intent settings (handles collections + complex types).
    if($Setting.PSObject.Properties['valueJson'] -and -not [string]::IsNullOrWhiteSpace("$($Setting.valueJson)"))
    {
        return "$($Setting.valueJson)"
    }
    if($Setting.PSObject.Properties['value'] -and $null -ne $Setting.value)
    {
        $v = $Setting.value
        if($v -is [string] -or $v -is [bool] -or $v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal])
        {
            return "$v"
        }
        return ($v | ConvertTo-Json -Depth 10 -Compress)
    }
    return ($Setting | ConvertTo-Json -Depth 10 -Compress)
}

function Get-IntentSettingCategory
{
    param($Setting)
    if(-not $Setting) { return "" }
    $defId = "$($Setting.definitionId)"
    if(-not $defId) { return "" }
    # Intent definitionId is shaped like "deviceConfiguration--baselineProtectionSettings_<group>_<name>".
    # Use the segment between "--" and the first "_".
    if($defId -match '--([^_]+)')
    {
        return $matches[1]
    }
    return ""
}

function Get-DefinitionValueKey
{
    param($DefValue)
    if(-not $DefValue) { return $null }
    if($DefValue.PSObject.Properties['#Definition_displayName'] -and $DefValue.'#Definition_displayName')
    {
        return "$($DefValue.'#Definition_displayName')"
    }
    if($DefValue.definition -and $DefValue.definition.displayName)
    {
        return "$($DefValue.definition.displayName)"
    }
    $bind = $DefValue.'definition@odata.bind'
    if($bind -and $bind -match "groupPolicyDefinitions\('([^']+)'\)")
    {
        return $matches[1]
    }
    if($DefValue.id) { return "$($DefValue.id)" }
    return $null
}

function Get-DefinitionValueDisplayValue
{
    param($DefValue)
    if(-not $DefValue) { return "" }
    $parts = @()
    $parts += "Enabled: $($DefValue.enabled)"
    if($DefValue.presentationValues)
    {
        foreach($pv in $DefValue.presentationValues)
        {
            $label = if($pv.PSObject.Properties['#Presentation_Label']) { "$($pv.'#Presentation_Label')" } else { "" }
            $val = $null
            if($pv.PSObject.Properties['value'] -and $null -ne $pv.value) { $val = "$($pv.value)" }
            elseif($pv.PSObject.Properties['values'] -and $pv.values) { $val = ($pv.values -join ', ') }
            if($null -ne $val)
            {
                if($label) { $parts += "$label = $val" } else { $parts += $val }
            }
        }
    }
    return ($parts -join '; ')
}

function Resolve-FullPolicyForCompare
{
    param($Policy)
    if(-not $Policy) { return }

    if($Policy.PSObject.Properties['_TokenId'] -and $Policy._TokenId -and $Policy.PolicyType)
    {
        $Policy.PolicyType.GetFullObject($Policy) | Out-Null
    }

    if($Policy.PolicyType)
    {
        $config = $Policy.PolicyType.GetCompareConfig()
        if($config -and $config.Hydrate)
        {
            & $config.Hydrate $Policy
        }
    }
}

function Compare-PolicyObjects
{
    param(
        [object[]]$Policies,
        [string]$CompareType,
        $CompareDefinition
    )

    if(-not $Policies -or $Policies.Count -lt 2)
    {
        Write-Log "Compare-PolicyObjects requires at least 2 policies" 3
        return @()
    }

    $p1 = $Policies[0]
    $p2 = $Policies[1]

    $Obj1 = if($p1) { $p1.Object } else { $null }
    $Obj2 = if($p2) { $p2.Object } else { $null }

    if(-not $Obj1 -or -not $Obj2)
    {
        # One side missing entirely (e.g. exported but deleted, or never
        # exported). That is a mismatch, not an ignored row.
        $n1 = if($p1) { $p1.Name } else { "" }
        $n2 = if($p2) { $p2.Name } else { "" }
        return @([PSCustomObject]@{ PropertyName="Object"; Object1Value=$n1; Object2Value=$n2; Category=$null; SubCategory=$null; Match=$false })
    }

    $policyType   = if($p1.PolicyType) { $p1.PolicyType } else { $p2.PolicyType }
    $compareValue = Get-CompareTypeValue $CompareType $CompareDefinition
    $compareDefinition = Get-CompareTypeDefinition $compareValue $CompareDefinition

    $docAvailable = [bool](Get-Command "Get-GraphDocumentation" -ErrorAction SilentlyContinue)

    # Forced documentation routing (parity with the original project): these
    # types keep their settings behind separate endpoints, so a raw property
    # comparison is meaningless regardless of the selected comparison type.
    $docForcedTypes = @(
        '#microsoft.graph.deviceManagementConfigurationPolicy',
        '#microsoft.graph.deviceManagementIntent',
        '#microsoft.graph.groupPolicyConfiguration')
    if($docAvailable -and $Obj1.'@OData.Type' -in $docForcedTypes)
    {
        return (Compare-ObjectsBasedonDocumentation $p1 $p2 $policyType)
    }

    if($compareValue -eq "doc" -and $docAvailable)
    {
        return (Compare-ObjectsBasedonDocumentation $p1 $p2 $policyType)
    }

    $settingsConfig = if($policyType) { $policyType.GetCompareConfig() } else { $null }

    if($settingsConfig)
    {
        if($compareValue -eq "property" -or -not $docAvailable)
        {
            Compare-ObjectsBasedonSettings $Obj1 $Obj2 $settingsConfig $policyType
        }
        else
        {
            Compare-ObjectsBasedonDocumentation $p1 $p2 $policyType
        }
    }
    elseif($compareValue -eq "property")
    {
        Compare-ObjectsBasedonProperty $Obj1 $Obj2 $policyType
    }
    elseif($compareDefinition -and $compareDefinition.Compare)
    {
        & $compareDefinition.Compare $Obj1 $Obj2
    }
    else
    {
        Compare-ObjectsBasedonProperty $Obj1 $Obj2 $policyType
    }
}

function Set-CompareColumnVisibility
{
    param($ShowCategory = $false, $ShowSubCategory = $false)

    Invoke-AppEvent "CompareColumnVisibilityChanged" @{
        ShowCategory    = [bool]$ShowCategory
        ShowSubCategory = [bool]$ShowSubCategory
    }
}

function Test-IgnoreCoreProperties
{
    return [bool](Get-CompareRuntimeOption "IgnoreCoreProperties" $false)
}

function Compare-ObjectsBasedonProperty
{
    param($Obj1, $Obj2, $PolicyType = $null)

    Write-Status "Compare objects based on property values"
    Set-CompareColumnVisibility $false

    $nameProp = if($PolicyType -and $PolicyType.NameProperty) { $PolicyType.NameProperty } else { "displayName" }

    $coreProps  = @($nameProp) + @($script:CoreObjectProperties | Where-Object { $_ -cne $nameProp })
    # 'Advertisements' is the original project's name for the embedded
    # assignment list; the new exports carry 'assignments'. Both honour the
    # SkipAssignments compare option.
    $postProps  = @("Advertisements", "assignments")
    # isAssigned is a lazily-recomputed STATUS flag, not configuration - it can
    # stay true long after the last assignment was removed.
    $skipProps  = @("@ObjectFromFile", "@ObjectFileName", "isAssigned")
    # Per-type server-state properties (e.g. policySet status/errorCode/items
    # async processing state) - see _ComparePropertiesToSkip on the class.
    if($PolicyType) { $skipProps += @($PolicyType.ComparePropertiesToSkip) }
    $ignoreCore = Test-IgnoreCoreProperties

    $script:compareProperties = @()

    foreach($propName in $coreProps)
    {
        if(-not ($Obj1.PSObject.Properties | Where-Object Name -eq $propName)) { continue }
        Add-CompareProperty $propName (ConvertTo-CompareJson $Obj1.$propName) (ConvertTo-CompareJson $Obj2.$propName) -Skip:$ignoreCore
    }

    $addedProps = @()
    foreach($propName in ($Obj1.PSObject.Properties.Name))
    {
        if($propName -in $coreProps -or $propName -in $postProps -or $propName -in $skipProps) { continue }
        if($propName -like "*@OData*" -or $propName -like "#microsoft.graph*") { continue }
        $addedProps += $propName
        Add-CompareProperty $propName (ConvertTo-CompareJson $Obj1.$propName) (ConvertTo-CompareJson $Obj2.$propName)
    }

    foreach($propName in ($Obj2.PSObject.Properties.Name))
    {
        if($propName -in $coreProps -or $propName -in $postProps -or $propName -in $skipProps -or $propName -in $addedProps) { continue }
        if($propName -like "*@OData*" -or $propName -like "#microsoft.graph*") { continue }
        Add-CompareProperty $propName (ConvertTo-CompareJson $Obj1.$propName) (ConvertTo-CompareJson $Obj2.$propName)
    }

    $skipAssignments = [bool](Get-CompareRuntimeOption "SkipAssignments" $false)
    foreach($propName in $postProps)
    {
        if(-not ($Obj1.PSObject.Properties | Where-Object Name -eq $propName)) { continue }
        Add-CompareProperty $propName (ConvertTo-CompareJson $Obj1.$propName) (ConvertTo-CompareJson $Obj2.$propName) -Skip:$skipAssignments
    }

    $script:compareProperties
}

function Compare-ObjectsBasedonSettings
{
    param($Obj1, $Obj2, [hashtable]$Config, $PolicyType = $null)

    Write-Status "Compare objects based on settings values"

    $getCategory = $Config.GetCategory
    Set-CompareColumnVisibility ([bool]$getCategory) $false

    $script:compareProperties = @()

    $nameProp = if($PolicyType -and $PolicyType.NameProperty) { $PolicyType.NameProperty } else { "displayName" }

    $ignoreCore = Test-IgnoreCoreProperties

    foreach($propName in (@($nameProp) + @($script:CoreObjectProperties | Where-Object { $_ -cne $nameProp })))
    {
        if(-not ($Obj1.PSObject.Properties | Where-Object Name -eq $propName)) { continue }
        Add-CompareProperty $propName (ConvertTo-CompareJson $Obj1.$propName) (ConvertTo-CompareJson $Obj2.$propName) -Skip:$ignoreCore
    }

    $settingsProp = $Config.Prop
    $getKey       = $Config.GetKey
    $getValue     = $Config.GetValue

    $settings1 = @(if($Obj1.$settingsProp) { $Obj1.$settingsProp } else { @() })
    $settings2 = @(if($Obj2.$settingsProp) { $Obj2.$settingsProp } else { @() })

    $map2 = @{}
    for($i = 0; $i -lt $settings2.Count; $i++)
    {
        $s = $settings2[$i]
        $key = & $getKey $s
        $keyStr = if($null -ne $key -and "$key" -ne "") { "$key" } else { "$($settingsProp)[$i]" }
        if(-not $map2.ContainsKey($keyStr)) { $map2[$keyStr] = $s }
    }

    $processedKeys = [System.Collections.Generic.HashSet[string]]::new()
    for($i = 0; $i -lt $settings1.Count; $i++)
    {
        $s = $settings1[$i]
        $key    = & $getKey $s
        $keyStr = if($null -ne $key -and "$key" -ne "") { "$key" } else { "$($settingsProp)[$i]" }
        $s2     = if($map2.ContainsKey($keyStr)) { $map2[$keyStr] } else { $null }
        $v1     = & $getValue $s
        $v2     = if($s2) { & $getValue $s2 } else { "" }
        $cat    = if($getCategory) { & $getCategory $s } else { $null }
        Add-CompareProperty $keyStr $v1 $v2 $cat
        [void]$processedKeys.Add($keyStr)
    }

    for($i = 0; $i -lt $settings2.Count; $i++)
    {
        $s = $settings2[$i]
        $key    = & $getKey $s
        $keyStr = if($null -ne $key -and "$key" -ne "") { "$key" } else { "$($settingsProp)[$i]" }
        if($processedKeys.Contains($keyStr)) { continue }
        $cat    = if($getCategory) { & $getCategory $s } else { $null }
        Add-CompareProperty $keyStr "" (& $getValue $s) $cat
    }

    $script:compareProperties
}

# Stable join key for one documentation row. EntityKey is THE identity
# (language-independent: schema entityKey, settingDefinitionId-derived, or a
# BasicInfo source field name). Settings rows qualify it with the category
# context because schema entityKeys (e.g. 'enabled') legally repeat across
# categories within one object. Rows without an EntityKey (logged by
# DocumentationContext.AddSetting) fall back to the localized Name.
function Get-CompareDocRowKey
{
    param($Row, [switch]$Basic)

    $entityKey = [string]$Row.EntityKey
    if([string]::IsNullOrEmpty($entityKey)) { return "name:$($Row.Name)" }
    if($Basic) { return $entityKey }
    return "$entityKey|$($Row.Category)|$($Row.SubCategory)"
}

function Compare-ObjectsBasedonDocumentation
{
    # Takes the POLICY WRAPPERS (JsonObject/Object + PolicyType), not raw JSON
    # — Get-GraphDocumentation needs the wrapper to dispatch input providers.
    param($Policy1, $Policy2, $PolicyType = $null)

    if(-not (Get-Command "Get-GraphDocumentation" -ErrorAction SilentlyContinue))
    {
        Write-Log "Documentation engine not available. Falling back to property comparison." 2
        Compare-ObjectsBasedonProperty $Policy1.Object $Policy2.Object $PolicyType
        return
    }

    Write-Status "Compare objects based on documentation values"

    if(-not $PolicyType) { $PolicyType = $Policy1.PolicyType }
    if(-not $PolicyType -and $script:compareSource) { $PolicyType = $script:compareSource.PolicyType }

    # Multi-value rows are flattened by the doc engine using this separator.
    $docOptions = @{
        ObjectSeparator = [string](Get-CompareRuntimeOption "ObjectSeparator" ([System.Environment]::NewLine))
    }

    $doc1 = Get-GraphDocumentation -PolicyObject $Policy1 -Options $docOptions
    $doc2 = Get-GraphDocumentation -PolicyObject $Policy2 -Options $docOptions

    $settingsValue = if($PolicyType -and $PolicyType.CompareValue) { $PolicyType.CompareValue } else { "Value" }
    $ignoreCore    = Test-IgnoreCoreProperties
    $obj1 = $Policy1.Object
    $obj2 = $Policy2.Object

    $hasSubCategory = ($null -ne ($doc1.FilteredSettings | Where-Object SubCategory)) -or
                      ($null -ne ($doc2.FilteredSettings | Where-Object SubCategory))
    Set-CompareColumnVisibility $true $hasSubCategory

    $script:compareProperties = @()

    # --- BasicInfo: outer join on EntityKey ---
    if($doc1.BasicInfo -and -not ($doc1.BasicInfo | Where-Object EntityKey -eq 'id'))
    {
        Add-CompareProperty "Id" $obj1.Id $obj2.Id ($doc1.BasicInfo[0].Category) -Skip:$ignoreCore
    }

    $basicIndex2 = @{}
    foreach($row in $doc2.BasicInfo)
    {
        $key = Get-CompareDocRowKey $row -Basic
        if(-not $basicIndex2.ContainsKey($key)) { $basicIndex2[$key] = $row }
    }
    $addedBasic = [System.Collections.Generic.HashSet[string]]::new()
    foreach($row in $doc1.BasicInfo)
    {
        $key = Get-CompareDocRowKey $row -Basic
        if(-not $addedBasic.Add($key)) { continue }
        $row2 = $basicIndex2[$key]
        Add-CompareProperty $row.Name $row.Value $row2.Value $row.Category -Skip:$ignoreCore
    }
    foreach($row in $doc2.BasicInfo)
    {
        $key = Get-CompareDocRowKey $row -Basic
        if(-not $addedBasic.Add($key)) { continue }
        Add-CompareProperty $row.Name $null $row.Value $row.Category -Skip:$ignoreCore
    }

    # --- Settings: outer join on EntityKey (category-qualified) ---
    $index2 = @{}
    foreach($row in $doc2.FilteredSettings)
    {
        $key = Get-CompareDocRowKey $row
        if($index2.ContainsKey($key))
        {
            Write-LogDebug "Duplicate documentation compare key '$key' on '$($Policy2.Name)' - first row wins"
            continue
        }
        $index2[$key] = $row
    }

    $added = [System.Collections.Generic.HashSet[string]]::new()
    foreach($row in $doc1.FilteredSettings)
    {
        $key = Get-CompareDocRowKey $row
        if(-not $added.Add($key))
        {
            Write-LogDebug "Duplicate documentation compare key '$key' on '$($Policy1.Name)' - first row wins"
            continue
        }
        $row2 = $index2[$key]
        $val1 = $row.$settingsValue
        $val2 = if($row2) { $row2.$settingsValue } else { $null }
        if($val1 -isnot [array] -and $val2 -is [array] -and $val2.Count -gt 1) { $val2 = $val2[0] }
        Add-CompareProperty $row.Name $val1 $val2 $row.Category $row.SubCategory
    }

    # Rows defined only on object 2 are appended last, in object 2's order.
    foreach($row in $doc2.FilteredSettings)
    {
        $key = Get-CompareDocRowKey $row
        if(-not $added.Add($key)) { continue }
        Add-CompareProperty $row.Name $null $row.$settingsValue $row.Category $row.SubCategory
    }

    # --- Applicability rules: joined on rule Id (stable), like the original ---
    $addedRules = [System.Collections.Generic.HashSet[string]]::new()
    foreach($rule in $doc1.ApplicabilityRules)
    {
        [void]$addedRules.Add([string]$rule.Id)
        $rule2 = $doc2.ApplicabilityRules | Where-Object Id -eq $rule.Id | Select-Object -First 1
        $val1  = "$($rule.Rule)$($docOptions.ObjectSeparator)$($rule.Value)"
        $val2  = if($rule2) { "$($rule2.Rule)$($docOptions.ObjectSeparator)$($rule2.Value)" } else { $null }
        Add-CompareProperty $rule.Property $val1 $val2 $rule.Category
    }
    foreach($rule in $doc2.ApplicabilityRules)
    {
        if(-not $addedRules.Add([string]$rule.Id)) { continue }
        Add-CompareProperty $rule.Property $null "$($rule.Rule)$($docOptions.ObjectSeparator)$($rule.Value)" $rule.Category
    }

    # --- Assignments: joined on Group + GroupMode + RawIntent ---
    # Equality is decided on the full row (group, filter, filter mode, app
    # intent settings) but the rendered values show only the group name, so
    # a changed filter shows as a mismatch on the same group — the original
    # project's 'simpleFullCompare' mode.
    $skipAssignments = [bool](Get-CompareRuntimeOption "SkipAssignments" $false)
    $assignmentLabel = $null
    if(Get-Command "Get-LanguageString" -ErrorAction SilentlyContinue)
    {
        $assignmentLabel = Get-LanguageString "TableHeaders.assignment"
    }
    if([string]::IsNullOrEmpty($assignmentLabel)) { $assignmentLabel = "Assignment" }

    $assignmentJson = {
        param($AssignmentRow)
        if($null -eq $AssignmentRow) { return $null }
        ($AssignmentRow |
            Select-Object -Property * -ExcludeProperty RawJsonValue, RawIntent, GroupMode, Category |
            ConvertTo-Json -Depth 10 -Compress)
    }

    $addedAssignments = [System.Collections.Generic.HashSet[string]]::new()
    foreach($assignment in $doc1.Assignments)
    {
        $key = "$($assignment.Group)|$($assignment.GroupMode)|$($assignment.RawIntent)"
        if(-not $addedAssignments.Add($key)) { continue }
        $assignment2 = $doc2.Assignments | Where-Object {
            "$($_.Group)|$($_.GroupMode)|$($_.RawIntent)" -eq $key
        } | Select-Object -First 1

        $rowName = if($assignment.RawIntent) { $assignment.Category } else { $assignmentLabel }
        $fullMatch = ((& $assignmentJson $assignment) -eq (& $assignmentJson $assignment2))
        $group2 = if($assignment2) { $assignment2.Group } else { $null }
        Add-CompareProperty $rowName $assignment.Group $group2 $assignment.GroupMode -Match $fullMatch -Skip:$skipAssignments
    }
    foreach($assignment in $doc2.Assignments)
    {
        $key = "$($assignment.Group)|$($assignment.GroupMode)|$($assignment.RawIntent)"
        if(-not $addedAssignments.Add($key)) { continue }
        $rowName = if($assignment.RawIntent) { $assignment.Category } else { $assignmentLabel }
        Add-CompareProperty $rowName $null $assignment.Group $assignment.GroupMode -Match $false -Skip:$skipAssignments
    }

    $script:compareProperties
}

function Get-CompareCsvInfo
{
    param($CompareInfo, $Policy, [string]$CompareType, $CompareDefinition)

    $compareProps = [Collections.Generic.List[String]]($script:defaultCompareProps)
    foreach($p in (Get-CompareTypeRemoveProperties $CompareType $CompareDefinition)) { $compareProps.Remove($p) | Out-Null }

    $CompResultValues = @()
    foreach($CompValue in $CompareInfo)
    {
        $CompResultValues += [PSCustomObject]@{
            ObjectName  = $Policy.Name
            Id          = $Policy.ID
            Type        = $Policy.PolicyType.Title
            ODataType   = $Policy.Object.'@OData.Type'
            Property    = $CompValue.PropertyName
            Value1      = $CompValue.Object1Value
            Value2      = $CompValue.Object2Value
            Category    = $CompValue.Category
            SubCategory = $CompValue.SubCategory
            Match       = $CompValue.Match
        }
    }

    $CompResultValues | Select-Object -Property $compareProps | ConvertTo-Csv -NoTypeInformation
}

function Get-CompareOutputProps
{
    param([CompareProviderBase]$Provider, [string]$CompareType, $CompareDefinition)

    $Props = [Collections.Generic.List[String]]($script:defaultCompareProps)

    if($Provider -and $Provider.RemoveProperties)
    {
        foreach($p in $Provider.RemoveProperties) { $Props.Remove($p) | Out-Null }
    }
    foreach($p in (Get-CompareTypeRemoveProperties $CompareType $CompareDefinition)) { $Props.Remove($p) | Out-Null }
    return $Props
}

function Get-CompareOutputType
{
    $type = Get-CompareRuntimeOption "SaveType" "objectType"
    Save-SettingStoreValue "Compare" "SaveType" $type
    return $type
}

function Get-BulkCompareOutputProvider
{
    $Provider = Get-CompareRuntimeOption "OutputProvider"
    if(-not $Provider)
    {
        $Provider = [CompareRegistry]::FindOutputProvider("csv")
    }
    if(-not $Provider)
    {
        $Provider = [CompareCSVOutputProvider]::new()
    }

    $delimiter = Get-CompareRuntimeOption "CsvDelimiter"
    if($Provider -is [CompareCSVOutputProvider] -and $delimiter)
    {
        $Provider.Delimiter = $delimiter
    }
    return $Provider
}

function Get-CompareOutputProviderByExtension
{
    param([string]$Extension)

    $Provider = [CompareRegistry]::FindOutputProviderByExtension($Extension)
    if(-not $Provider) { $Provider = [CompareCSVOutputProvider]::new() }
    return $Provider
}

function Build-BulkCompareRow
{
    param($Pair, $CompValue)
    [PSCustomObject]@{
        ObjectName  = $Pair.Name
        Id          = $Pair.Id
        Type        = $Pair.PolicyType.Title
        ODataType   = if($Pair.Policy1) { $Pair.Policy1.Object.'@OData.Type' } else { $Pair.Policy2.Object.'@OData.Type' }
        Property    = $CompValue.PropertyName
        Value1      = $CompValue.Object1Value
        Value2      = $CompValue.Object2Value
        Category    = $CompValue.Category
        SubCategory = $CompValue.SubCategory
        Match       = $CompValue.Match
    }
}

function Save-BulkCompareResults
{
    param($CompResultValues, $File, $Props, [CompareOutputProviderBase]$OutputProvider = $null)

    if($CompResultValues.Count -eq 0) { return }

    if(-not $OutputProvider) { $OutputProvider = Get-BulkCompareOutputProvider }

    Write-Log "Save bulk compare results to $File"
    $OutputProvider.FormatRows($CompResultValues, $Props) | Out-File -LiteralPath $File -Force -Encoding UTF8
}

function Start-BulkCompare
{
    param(
        [CompareProviderBase]$Provider,
        [object[]]$SelectedGroups = @(),
        [CompareOutputProviderBase]$OutputProvider = $null,
        [string]$OutputType,
        [string]$CompareType,
        $CompareDefinition
    )

    Write-Log "****************************************************************"
    Write-Log "Start bulk compare: $($Provider.Name)"
    Write-Log "****************************************************************"

    if(-not $Provider.Validate()) { return }

    $Provider.SaveSettings()

    if(-not $Provider.IgnoreGroups)
    {
        $SelectedGroups = @($SelectedGroups)
        if(-not $SelectedGroups)
        {
            throw "No object types selected"
        }
    }

    if(-not $OutputProvider) { $OutputProvider = Get-BulkCompareOutputProvider }
    $compareProps   = Get-CompareOutputProps $Provider $CompareType $CompareDefinition
    if(-not $OutputType) { $OutputType = Get-CompareOutputType }
    $allRows        = @()
    $pairCount      = 0
    $ext            = $OutputProvider.Extension

    $pairs = $Provider.GetComparePairs($SelectedGroups)

    $perTypeRows = @{}

    foreach($Pair in $pairs)
    {
        $compareProperties = Compare-PolicyObjects @($Pair.Policy1, $Pair.Policy2) -CompareType $CompareType -CompareDefinition $CompareDefinition
        $pairCount++

        $rows = @()
        foreach($v in $compareProperties) { $rows += Build-BulkCompareRow $Pair $v }

        if($OutputType -eq "objectType")
        {
            $typeKey = $Pair.PolicyType.ID
            if(-not $perTypeRows.ContainsKey($typeKey))
            {
                $perTypeRows[$typeKey] = @{ Rows = @(); SaveFolder = $Pair.SaveFolder }
            }
            $perTypeRows[$typeKey].Rows += $rows
        }
        else
        {
            $allRows += $rows
        }
    }

    if($OutputType -eq "objectType")
    {
        foreach($typeKey in $perTypeRows.Keys)
        {
            $entry = $perTypeRows[$typeKey]
            if($entry.Rows.Count -gt 0)
            {
                $File = Join-Path $entry.SaveFolder "Compare_$((Get-Date).ToString('yyyyMMdd-HHmm')).$ext"
                Save-BulkCompareResults $entry.Rows $File $compareProps $OutputProvider
            }
        }
    }
    elseif($allRows.Count -gt 0)
    {
        $rootFolder = if($Provider.PSObject.Properties['SourcePath']) { $Provider.SourcePath } `
                      elseif($Provider.PSObject.Properties['ExportPath']) { $Provider.ExportPath } `
                      elseif($Provider.PSObject.Properties['SavePath'] -and $Provider.SavePath) { $Provider.SavePath } `
                      else { [Environment]::GetFolderPath("MyDocuments") }
        $File = Join-Path $rootFolder "Compare_$((Get-Date).ToString('yyyyMMdd-HHmm')).$ext"
        Save-BulkCompareResults $allRows $File $compareProps $OutputProvider
    }

    Write-Log "****************************************************************"
    Write-Log "Bulk compare finished: $pairCount pairs compared"
    Write-Log "****************************************************************"
    Write-Status ""

    if($pairCount -eq 0)
    {
        throw "No objects were compared. Verify settings and exported files."
    }
}
