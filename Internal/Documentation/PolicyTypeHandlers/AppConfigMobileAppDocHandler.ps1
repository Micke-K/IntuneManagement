# iOS Mobile App Configuration documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:2065. Claims
# @odata.type='#microsoft.graph.iosMobileAppConfiguration'.
#
# Two main paths:
#   1. iOS plist (base64'd encodedSettingXml) — parsed offline, key/value/type
#      rows emitted directly
#   2. settings collection (Outlook-specific or generic appConfig key/value)
#      The Outlook ObjectInfo translation needs the ObjectInfo JSON walker
#      (deferred); offline we fall through to raw key=value rows under the
#      generic "Additional configuration" subcategory.

class AppConfigMobileAppDocHandler : DocumentationHandlerBase {
    AppConfigMobileAppDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.iosMobileAppConfiguration')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        Add-BasicDefaultValues $PolicyObject
        Add-BasicAdditionalValues $PolicyObject
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'SettingDetails.appConfiguration') '@odata.type'
        Add-BasicPropertyValue (Get-LanguageString 'Inputs.enrollmentTypeLabel') (Get-LanguageString 'EnrollmentType.devicesWithEnrollment') 'enrollmentType'

        $platformId = Get-ObjectPlatformFromType $obj
        if ($platformId) {
            Add-BasicPropertyValue (Get-LanguageString 'Inputs.platformLabel') (Get-LanguageString "Platform.$platformId") 'platform'
        }

        # Targeted apps — resolve IDs to displayNames when the tenant catalog is
        # available, otherwise emit raw IDs.
        $allApps = Get-CDAllTenantApps
        $appsList = @()
        foreach ($id in $obj.targetedMobileApps) {
            $app = $allApps | Where-Object Id -EQ $id | Select-Object -First 1
            $appsList += if ($app -and $app.displayName) { $app.displayName } else { $id }
        }
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.targetedAppLabel') ($appsList -join $Context.ObjectSeparator) 'targetedMobileApps'

        $category = Get-LanguageString 'TableHeaders.settings'

        if ($obj.encodedSettingXml) {
            # iOS plist. The portal emits a bare <dict> root but Graph also accepts a
            # <plist> wrapper, and the portal's validator explicitly allows nested
            # <dict>/<array> values - so the walk has to recurse.
            $xml = $null
            try {
                $xml = [xml]([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.encodedSettingXml)))
            }
            catch {
                Write-LogError 'Failed to convert iOS encodedSettingXml to XML' $_.Exception
                return
            }

            $rootDict = if ($xml.dict) { $xml.dict } elseif ($xml.plist.dict) { $xml.plist.dict } else { $null }
            if (-not $rootDict) {
                Write-Log 'iOS app config: encodedSettingXml has no <dict> root; no settings documented' 2
                return
            }

            $plistRows = @(Expand-IosPlistDictionary $rootDict)

            # ValueType is not part of the default documentation properties, so also
            # emit the portal's 3-column grid (key / value type / value).
            if ($plistRows.Count -gt 0) {
                $typeRows = foreach ($row in $plistRows) {
                    [PSCustomObject]@{
                        ConfigurationKey   = $row.Key
                        ValueType          = Get-AppConfigValueTypeName $row.ValueType
                        ConfigurationValue = $row.Value
                        EntityKey          = $row.Key
                    }
                }
                Add-CustomTable 'AppConfigSettings' @('ConfigurationKey','ValueType','ConfigurationValue') $typeRows -Order 110 -LanguageId 'TableHeaders.settings'
            }
            return
        }

        # Outlook gets schema-driven translation (port of old DocumentationCustom.psm1:2141-2176).
        $isOutlook = $false
        foreach ($id in $obj.targetedMobileApps) {
            $app = $allApps | Where-Object Id -EQ $id | Select-Object -First 1
            if ($app.displayName -eq 'Microsoft Outlook') { $isOutlook = $true; break }
        }
        if (-not $isOutlook -and @($obj.settings | Where-Object { $_.appConfigKey -like 'com.microsoft.outlook*' })) { $isOutlook = $true }
        if ($isOutlook) {
            $hasAccountType = @($obj.settings | Where-Object { $_.appConfigKey -eq 'com.microsoft.outlook.EmailProfile.AccountType' })
            $outlookSettings = [PSCustomObject]@{ configureEmail = [bool]$hasAccountType }
            foreach ($setting in @($obj.settings)) {
                $val = if ($setting.appConfigKeyType -eq 'booleanType') { $setting.appConfigKeyValue -eq 'true' } else { $setting.appConfigKeyValue }
                $outlookSettings | Add-Member -MemberType NoteProperty -Name $setting.appConfigKey -Value $val -Force
            }
            Invoke-DocAppConfigManifest $outlookSettings (Join-Path (Join-Path $script:AppRootFolder 'Config\ObjectInfo') '#AppConfigOutlookDevice.json') $Context
        }

        # Remaining settings fall through to raw key=value rows under the
        # "Additional configuration" subcategory.
        $addedSettings = Get-DocumentedSettings
        $languageTitleId = 'TableHeaders.settings'

        $typeRows = @()
        foreach ($setting in $obj.settings) {
            if ($addedSettings | Where-Object EntityKey -EQ $setting.appConfigKey) {
                $languageTitleId = 'SettingDetails.additionalConfiguration'
                continue
            }

            # The portal grid shows the value TYPE as its own column; keep that
            # (a tokenType value like {{userprincipalname}} is not a literal string).
            $typeRows += [PSCustomObject]@{
                ConfigurationKey   = $setting.appConfigKey
                ValueType          = Get-AppConfigValueTypeName $setting.appConfigKeyType
                ConfigurationValue = $setting.appConfigKeyValue
                EntityKey          = $setting.appConfigKey
            }
        }
        if ($typeRows.Count -gt 0) {
            Add-CustomTable 'AppConfigSettings' @('ConfigurationKey','ValueType','ConfigurationValue') $typeRows -Order 110 -LanguageId $languageTitleId
        }
    }
}

# Localized name for an app-config value type. Accepts both the Graph
# mdmAppConfigKeyType members (stringType/integerType/realType/booleanType/
# tokenType) and the bare plist element names (string/integer/real/boolean/...).
# tokenType has no language string - the portal shows it blank, so the raw value
# is a strict improvement.
function Get-AppConfigValueTypeName {
    param([string]$ValueType)

    if (-not $ValueType) { return $null }
    $key = switch -Regex ($ValueType) {
        '^string'  { 'SettingDetails.string' }
        '^integer' { 'SettingDetails.integer' }
        '^real'    { 'SettingDetails.real' }
        '^boolean' { 'SettingDetails.boolean' }
        default    { $null }
    }
    if (-not $key) { return $ValueType }
    $value = Get-LanguageString $key -IgnoreMissing
    if ([string]::IsNullOrEmpty($value)) { $ValueType } else { $value }
}

# Flatten an iOS plist <dict> into one row per LEAF value. Nested containers are
# addressed the way the plist itself addresses them:
#   <dict>  -> "parent.child"
#   <array> -> "parent[0]"
# Without this, a nested container produced an EMPTY value (.'#text' on an element
# with element children returns nothing) and the payload was silently lost.
function Expand-IosPlistDictionary {
    param($DictNode, [string]$Prefix = '', [int]$Depth = 0)

    if ($Depth -gt 10) {
        Write-Log 'iOS app config: plist nesting deeper than 10 levels; remaining levels not documented' 2
        return
    }

    $children = @($DictNode.ChildNodes)
    for ($i = 0; $i -lt $children.Count; $i++) {
        if ($children[$i].Name -ne 'key') { continue }
        $name = $children[$i].'#text'
        $i++
        if ($i -ge $children.Count) { break }
        $valueNode = $children[$i]
        $key = if ($Prefix) { "$Prefix$name" } else { [string]$name }

        switch ($valueNode.Name) {
            'true'  { [PSCustomObject]@{ Key = $key; ValueType = 'boolean'; Value = 'true' } }
            'false' { [PSCustomObject]@{ Key = $key; ValueType = 'boolean'; Value = 'false' } }
            'dict'  { Expand-IosPlistDictionary $valueNode "$key." ($Depth + 1) }
            'array' {
                $idx = 0
                foreach ($item in @($valueNode.ChildNodes)) {
                    $itemKey = "$key[$idx]"
                    if ($item.Name -eq 'dict') { Expand-IosPlistDictionary $item "$itemKey." ($Depth + 1) }
                    elseif ($item.Name -eq 'true' -or $item.Name -eq 'false') {
                        [PSCustomObject]@{ Key = $itemKey; ValueType = 'boolean'; Value = $item.Name }
                    }
                    else {
                        [PSCustomObject]@{ Key = $itemKey; ValueType = $item.Name; Value = $item.'#text' }
                    }
                    $idx++
                }
            }
            default { [PSCustomObject]@{ Key = $key; ValueType = $valueNode.Name; Value = $valueNode.'#text' } }
        }
    }
}

[DocumentationRegistry]::RegisterHandler([AppConfigMobileAppDocHandler]::new())
