# Managed App Configuration documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:2196. Claims
# @odata.type='#microsoft.graph.targetedManagedAppConfiguration'.
#
# Outlook + Edge ObjectInfo translations (and the Edge bookmark/AllowList/
# BlockList delimiter rewrites) deferred until the walker is ported. Offline
# falls through to raw customSettings under TACSettings.generalSettings.

class ManagedAppConfigDocHandler : DocumentationHandlerBase {
    ManagedAppConfigDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.targetedManagedAppConfiguration')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        Add-BasicDefaultValues $PolicyObject
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'SettingDetails.appConfiguration') '@odata.type'

        $customApps, $publishedApps = Get-CDMobileApps $obj.Apps

        Add-BasicPropertyValue (Get-LanguageString 'Inputs.enrollmentTypeLabel') (Get-LanguageString 'EnrollmentType.devicesWithoutEnrollment') 'enrollmentType'
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.publicApps') ($publishedApps -join $Context.ObjectSeparator) 'publishedApps'
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.customApps') ($customApps    -join $Context.ObjectSeparator) 'customApps'

        # appGroupType - "Target policy to". The all* variants don't list individual
        # apps, so surfacing the mode is what tells the reader the scope. Graph's
        # enum member is allCoreMicrosoftApps; the portal's string is coreMicrosoftApps.
        $appGroupTypeKeys = @{
            'selectedPublicApps'   = 'AppGroupType.selectedPublicApps'
            'allApps'              = 'AppGroupType.allApps'
            'allMicrosoftApps'     = 'AppGroupType.allMicrosoftApps'
            'allCoreMicrosoftApps' = 'AppGroupType.coreMicrosoftApps'
        }
        $agtRaw = "$($obj.appGroupType)"
        if ($agtRaw) {
            $agtKey   = $appGroupTypeKeys[$agtRaw]
            $agtValue = if ($agtKey) { Get-LanguageString $agtKey -IgnoreMissing } else { $null }
            if ([string]::IsNullOrEmpty($agtValue)) { $agtValue = $agtRaw }
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.targetPolicyTo') $agtValue 'appGroupType'
        }

        # targetedAppManagementLevels - flags enum, returned as a comma-separated
        # string (e.g. "mdm, androidEnterprise"). Map each flag to its label,
        # falling back to the raw flag value when no string exists.
        $mgmtLevelKeys = @{
            'unspecified'                                            = 'AppProtection.allAppTypes'
            'unmanaged'                                              = 'AppProtection.appsOnUnmanagedDevices'
            'mdm'                                                    = 'AppProtection.appsOnIntuneManagedDevices'
            'androidEnterprise'                                      = 'AppProtection.appsInAndroidWorkProfile'
            'androidEnterpriseDedicatedDevicesWithAzureAdSharedMode' = 'AppProtection.appsOnAndroidEnterpriseDedicatedDevicesWithAzureAdSharedMode'
            'androidOpenSourceProjectUserAssociated'                 = 'AppProtection.appsOnAndroidOpenSourceProjectUserAssociated'
            'androidOpenSourceProjectUserless'                       = 'AppProtection.appsOnAndroidOpenSourceProjectUserless'
        }
        $mgmtRaw = "$($obj.targetedAppManagementLevels)"
        if ($mgmtRaw) {
            $mgmtParts = @()
            foreach ($lvl in ($mgmtRaw -split ',')) {
                $lvlTrim = $lvl.Trim()
                if (-not $lvlTrim) { continue }
                $lvlKey   = $mgmtLevelKeys[$lvlTrim]
                $lvlValue = if ($lvlKey) { Get-LanguageString $lvlKey -IgnoreMissing } else { $null }
                if ([string]::IsNullOrEmpty($lvlValue)) { $lvlValue = $lvlTrim }
                $mgmtParts += $lvlValue
            }
            if ($mgmtParts.Count -gt 0) {
                Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.managementType') ($mgmtParts -join $Context.ObjectSeparator) 'targetedAppManagementLevels'
            }
        }

        Add-BasicAdditionalValues $PolicyObject

        # Outlook / Edge get schema-driven translation via their ObjectInfo files
        # (port of old DocumentationCustom.psm1:2229-2260). Build a flat settings
        # object keyed by customSetting name, then walk the matching manifest.
        $appSettings = [PSCustomObject]@{}
        foreach ($setting in @($obj.customSettings)) {
            $appSettings | Add-Member -MemberType NoteProperty -Name $setting.name -Value $setting.value -Force
        }
        $objInfoDir = Join-Path $script:AppRootFolder 'Config\ObjectInfo'

        # Unpack every packed/delimited value BEFORE the manifests read them, so the
        # rewrite also benefits the raw fall-through rows below.
        # NB Where-Object: on a property-less object .PSObject.Properties.Name is
        # $null, and @($null) is a one-element array containing $null
        foreach ($name in @($appSettings.PSObject.Properties.Name | Where-Object { $_ })) {
            $sep = $script:_mamPackedSettingSeparators[$name]
            if (-not $sep -or -not $appSettings.$name -or $appSettings.$name -isnot [string]) { continue }
            $unpacked = $appSettings.$name
            # Record separator first - doing it the other way round destroys it
            if ($sep.RecordSep) { $unpacked = $unpacked.Replace($sep.RecordSep, $Context.ObjectSeparator) }
            if ($sep.FieldSep)  { $unpacked = $unpacked.Replace($sep.FieldSep, $Context.PropertySeparator) }
            $appSettings.$name = $unpacked
        }

        # App identities differ per platform: Outlook/Edge are packageId on Android,
        # bundleId on iOS and windowsAppId on Windows. Matching only one of them
        # silently skipped the whole manifest for the other platforms.
        if (Test-MamAppTargeted $obj.Apps @('com.microsoft.office.outlook')) {
            Invoke-DocAppConfigManifest $appSettings (Join-Path $objInfoDir '#AppConfigOutlookApp.json') $Context
        }
        if (Test-MamAppTargeted $obj.Apps @('com.microsoft.msedge', 'com.microsoft.emmx', 'com.microsoft.edge')) {
            Invoke-DocAppConfigManifest $appSettings (Join-Path $objInfoDir '#AppConfigEdgeApp.json') $Context
        }

        # Settings-catalog settings (the "Settings catalog" wizard step, used by the
        # Windows MAM flavour). Without this a policy whose entire payload lives in
        # `settings` documented as a header and nothing else.
        Invoke-DocMamSettingsCatalog $obj $Context

        # Remaining customSettings fall through to raw key=value rows.
        $addedSettings = Get-DocumentedSettings
        $category = Get-LanguageString 'TACSettings.generalSettings'

        foreach ($setting in $obj.customSettings) {
            if ($addedSettings | Where-Object EntityKey -EQ $setting.name) { continue }
            # Use the unpacked value when one was produced above
            $value = if ($null -ne $appSettings.PSObject.Properties[$setting.name]) { $appSettings."$($setting.name)" } else { $setting.value }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = $setting.name
                Value = $value
                EntityKey = $setting.name
                Category = $category
            })
        }
    }
}

# Document the settings-catalog part of a MAM app configuration.
#
# The Managed apps wizard has a "Settings catalog" step whose values land in the
# `settings` navigation property (Collection(deviceManagementConfigurationSetting))
# rather than in customSettings - Windows MAM policies are entirely settings-catalog.
# The portal renders it as its own blade ABOVE the classic Settings blade, so these
# rows go into a table of their own (negative Order = before the settings table)
# instead of being merged into it.
#
# Resolution is the SettingsCatalog provider's own code - see
# Get-SettingsCatalogDocumentationRows. This used to be a copy of it that had
# drifted, losing the category grouping.
function Invoke-DocMamSettingsCatalog {
    param($Obj, [DocumentationContext]$Context)

    $cfgSettings = @($Obj.settings)

    $hasDefs = $false
    foreach ($s in $cfgSettings) {
        if ($s.settingDefinitions -and ($s.settingDefinitions | Measure-Object).Count -gt 0) { $hasDefs = $true; break }
    }

    # Source-tenant-specific fetch (by policy id) - same gating as the Settings
    # Catalog provider. Exports carrying settings inline still work offline via the
    # walker's generic per-setting definition fallback.
    if (-not $hasDefs -and $Obj.Id -and -not $Context.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
        try {
            $resp = Invoke-MSGraphAPI -Url "/deviceAppManagement/targetedManagedAppConfigurations('$($Obj.Id)')/settings?`$expand=settingDefinitions&`$top=1000" -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context) -ODataMetadata 'minimal' -NoError
            if ($resp -and $resp.Value) { $cfgSettings = @($resp.Value) }
        }
        catch {
            Write-LogError "Failed to fetch settings catalog settings for app configuration $($Obj.Id)" $_.Exception
        }
    }

    if (@($cfgSettings).Count -eq 0) { return }

    $rows = @(Get-SettingsCatalogDocumentationRows $cfgSettings $Context)
    if ($rows.Count -eq 0) { return }

    Add-CustomTable 'SettingsCatalog' @('Name','Value') $rows -Order -100 -LanguageId 'SettingDetails.settingsCatalog'
}

# MAM settings whose value packs multiple records/fields into one string.
# RecordSep splits repeated records, FieldSep splits fields inside a record.
$script:_mamPackedSettingSeparators = @{
    # title|url pairs, records separated by ||
    'com.microsoft.intune.mam.managedbrowser.bookmarks'                = @{ RecordSep = '||'; FieldSep = '|' }
    'com.microsoft.intune.mam.managedbrowser.managedTopSites'          = @{ RecordSep = '||'; FieldSep = '|' }
    # plain pipe-separated lists
    'com.microsoft.intune.mam.managedbrowser.AllowListURLs'            = @{ RecordSep = '|' }
    'com.microsoft.intune.mam.managedbrowser.BlockListURLs'            = @{ RecordSep = '|' }
    'com.microsoft.intune.mam.managedbrowser.disabledFeatures'         = @{ RecordSep = '|' }
    'com.microsoft.intune.mam.managedbrowser.InternalPagesBlockList'   = @{ RecordSep = '|' }
    'com.microsoft.intune.mam.managedbrowser.PopupsAllowedForUrls'     = @{ RecordSep = '|' }
    'com.microsoft.intune.mam.managedbrowser.PopupsBlockedForUrls'     = @{ RecordSep = '|' }
    'com.microsoft.intune.mam.managedbrowser.FileUploadAllowedForUrls' = @{ RecordSep = '|' }
    'com.microsoft.intune.mam.managedbrowser.FileUploadBlockedForUrls' = @{ RecordSep = '|' }
    'com.microsoft.intune.mam.managedbrowser.NewTabPageLayout.Custom'  = @{ RecordSep = '|' }
}

# True when any targeted app matches one of the given app identifiers on ANY
# platform identity (Android packageId / iOS bundleId / Windows windowsAppId).
function Test-MamAppTargeted {
    param($Apps, [string[]]$Identifiers)

    foreach ($app in @($Apps)) {
        $id = $app.mobileAppIdentifier
        if (-not $id) { continue }
        foreach ($candidate in @($id.packageId, $id.bundleId, $id.windowsAppId)) {
            if ($candidate -and $candidate -in $Identifiers) { return $true }
        }
    }
    return $false
}

[DocumentationRegistry]::RegisterHandler([ManagedAppConfigDocHandler]::new())
