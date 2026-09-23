# Device Enrollment Platform Restriction documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:4345. Claims the two
# platform-restriction @odata.types:
#   ...deviceEnrollmentPlatformRestrictionConfiguration   (single platform)
#   ...deviceEnrollmentPlatformRestrictionsConfiguration  (all platforms — the
#                                                          aggregate that emits
#                                                          one row block per
#                                                          platform sub-restriction)
#
# Doesn't handle deviceEnrollmentLimitConfiguration — that has a different
# shape (single "Device limit" setting) and lives behind the generic Profile
# input provider in the old engine.

class EnrollmentPlatformRestrictionDocHandler : DocumentationHandlerBase {
    EnrollmentPlatformRestrictionDocHandler() {
        $this.ODataTypes = @(
            '#microsoft.graph.deviceEnrollmentPlatformRestrictionConfiguration',
            '#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration'
        )
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        Add-BasicDefaultValues $PolicyObject
        Add-BasicAdditionalValues $PolicyObject
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'Titles.deviceTypeEnrollmentRestrictions') '@odata.type'

        # platformType (single variant) -> Platform.* language id
        $singlePlatformLngId = switch ($obj.platformType) {
            'androidForWork' { 'androidWorkProfile' }
            'mac'            { 'macOS' }
            'ios'            { 'iOS' }
            'android'        { 'android' }
            'windows'        { 'windows' }
            'tvos'           { 'tvOS' }
            'visionOS'       { 'visionOS' }
            default          { $obj.platformType }
        }

        $isAggregate = $obj.'@odata.type' -eq '#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration'
        if ($isAggregate) {
            # The default "All users and all devices" config carries one sub-restriction
            # per platform. Graph exposes 10, including the legacy macRestriction (a dupe
            # of the version-capable macOSRestriction) and the deprecated
            # windowsMobileRestriction. Render the current platforms; prefer
            # macOSRestriction over macRestriction. $platformMap: property -> name key.
            $platform    = Get-LanguageString 'AzureCA.classicPolicyAllPlatforms'
            $platformMap = [ordered]@{
                'androidForWorkRestriction' = 'Platform.androidWorkProfile'
                'androidRestriction'        = 'Platform.android'
                'iosRestriction'            = 'Platform.iOS'
                'macOSRestriction'          = 'Platform.macOS'
                'tvosRestriction'           = 'Platform.tvOS'
                'visionOSRestriction'       = 'Platform.visionOS'
                'windowsRestriction'        = 'Platform.windows'
                'windowsHomeSkuRestriction' = 'Devices.windowsHomeSku'
            }
        }
        else {
            $platform    = Get-LanguageString "Platform.$singlePlatformLngId"
            $platformMap = [ordered]@{ 'platformRestriction' = "Platform.$singlePlatformLngId" }
        }

        Add-BasicPropertyValue (Get-LanguageString 'Inputs.platformLabel') $platform 'platformType'

        $allowStr        = Get-LanguageString 'BooleanActions.allow'
        $blockStr        = Get-LanguageString 'BooleanActions.block'
        $category        = Get-LanguageString 'EnrollmentRestrictions.DeviceType.platformSettings'
        $cantRestrictStr = Get-LanguageString 'EnrollmentRestrictions.DeviceType.cannotRestrict'

        foreach ($prop in $platformMap.Keys) {
            $restrict = $obj.$prop
            if (-not $restrict) { continue }

            $nameKey = $platformMap[$prop]
            $typeStr = Get-LanguageString $nameKey

            # OS version range, blank when unset. macOSRestriction is version-capable,
            # so (unlike the legacy macRestriction the old handler forced to "cannot
            # restrict") every platform now reports its actual osMin/osMax range.
            $version = if ($restrict.osMinimumVersion -or $restrict.osMaximumVersion) {
                           "$($restrict.osMinimumVersion)-$($restrict.osMaximumVersion)"
                       } else { '' }

            # Manufacturer blocking: old code has a typo (`'andriod'` instead of
            # `'android'`) which means only 'androidWorkProfile' actually emits
            # the blockedManufacturers list. Everything else — including the
            # correctly-spelled 'android' (device administrator) — falls
            # through to "Restriction not supported". Preserving the behavior
            # because golden fixtures match it; the typo is a known wart in
            # old code that fixing would silently change output.
            $blockedManufacturers = if ($nameKey -eq 'Platform.androidWorkProfile') {
                @($restrict.blockedManufacturers) -join $Context.PropertySeparator
            } else {
                $cantRestrictStr
            }

            # Aggregate variant uses platform name as SubCategory; single variant uses $null
            $subCategory = if ($isAggregate) { $typeStr } else { $null }

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'EnrollmentRestrictions.DeviceType.type')
                Value = $typeStr
                EntityKey = 'platformType'
                Category = $category; SubCategory = $subCategory
            })

            $platformAccess = if ($restrict.platformBlocked) { $blockStr } else { $allowStr }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'EnrollmentRestrictions.DeviceType.platform')
                Value = $platformAccess
                EntityKey = 'platformBlocked'
                Category = $category; SubCategory = $subCategory
            })

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'EnrollmentRestrictions.DeviceType.versions')
                Value = $version
                EntityKey = 'versions'
                Category = $category; SubCategory = $subCategory
            })

            $personalAccess = if ($restrict.personalDeviceEnrollmentBlocked) { $blockStr } else { $allowStr }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'EnrollmentRestrictions.DeviceType.personal')
                Value = $personalAccess
                EntityKey = 'personalDeviceEnrollmentBlocked'
                Category = $category; SubCategory = $subCategory
            })

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'EnrollmentRestrictions.DeviceType.deviceManufacturer')
                Value = $blockedManufacturers
                EntityKey = 'blockedManufacturers'
                Category = $category; SubCategory = $subCategory
            })
        }
    }
}

[DocumentationRegistry]::RegisterHandler([EnrollmentPlatformRestrictionDocHandler]::new())
