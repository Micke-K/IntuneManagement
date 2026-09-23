# Android Managed Store App Configuration documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:1928. Claims
# @odata.type='#microsoft.graph.androidManagedStoreAppConfiguration' and
# (per the plan's batching) also the legacy androidForWorkMobileAppConfig
# variant since both have the same shape.
#
# Profile applicability translates to one of three workProfile/deviceOwner
# variants which becomes the "Profile type" basic-info value.
# Outlook ObjectInfo translation deferred until the walker is ported.

class AppConfigAndroidStoreDocHandler : DocumentationHandlerBase {
    AppConfigAndroidStoreDocHandler() {
        $this.ODataTypes = @(
            '#microsoft.graph.androidManagedStoreAppConfiguration',
            '#microsoft.graph.androidForWorkMobileAppConfiguration'
        )
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        # ProfileType maps to a localized "App configuration" suffix
        $profileString = switch ($obj.profileApplicability) {
            'default'             { Get-LanguageString 'ProfileType.workProfileAndDeviceOwner' }
            'androidWorkProfile'  { Get-LanguageString 'ProfileType.workProfileOnly' }
            'androidDeviceOwner'  { Get-LanguageString 'ProfileType.deviceOwnerOnly' }
            default               { $null }
        }
        # Pass profileString as the Profile-type override so Add-BasicDefaultValues
        # emits one (and only one) Profile type row matching old engine's pattern
        # at DocumentationCustom.psm1:1955.
        Add-BasicDefaultValues $PolicyObject $profileString
        Add-BasicAdditionalValues $PolicyObject
        # Targeted apps — resolved to displayNames when catalog available
        $allApps  = Get-CDAllTenantApps
        $appsList = @()
        foreach ($id in $obj.targetedMobileApps) {
            $app = $allApps | Where-Object Id -EQ $id | Select-Object -First 1
            $appsList += if ($app -and $app.displayName) { $app.displayName } else { $id }
        }
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.targetedAppLabel') ($appsList -join $Context.ObjectSeparator) 'targetedMobileApps'

        if ($obj.packageId) {
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.packageId') $obj.packageId 'packageId'
        }

        # appSupportsOemConfig is the discriminator the portal uses to split OEMConfig
        # policies into their own blade - surface it so an OEMConfig policy isn't
        # documented as an ordinary app configuration. NB: TableHeaders.configurationType
        # renders as "Profile type" and would collide with the row above.
        if ($obj.appSupportsOemConfig -eq $true) {
            Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.policyType') (Get-LanguageString 'ConfigurationTypes.androidForWorkOemConfig') 'appSupportsOemConfig'
        }

        # connectedAppsEnabled - "Connected apps" toggle. The portal offers
        # Enabled / Not configured (not Enabled/Disabled).
        $connKey = if ($obj.connectedAppsEnabled -eq $true) { 'Inputs.enabled' } else { 'Inputs.notConfigured' }
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.connectedApps') (Get-LanguageString $connKey) 'connectedAppsEnabled'

        # credentialProviderRoleState - androidAppCredentialProviderRoleState enum
        # (notConfigured / allowed only). The portal renders the same two options as
        # the connected-apps toggle: Enabled / Not configured.
        $credKeys = @{
            'notConfigured' = 'Inputs.notConfigured'
            'allowed'       = 'Inputs.enabled'
        }
        $credRaw = "$($obj.credentialProviderRoleState)"
        if ($credRaw) {
            $credKey   = $credKeys[$credRaw]
            $credValue = if ($credKey) { Get-LanguageString $credKey -IgnoreMissing } else { $null }
            if ([string]::IsNullOrEmpty($credValue)) { $credValue = $credRaw }
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.credentialProvider') $credValue 'credentialProviderRoleState'
        }

        if (-not $obj.payloadJson) { return }

        $payloadData = $null
        try {
            $payloadData = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($obj.payloadJson)) | ConvertFrom-Json
        }
        catch {
            Write-LogError 'Failed to parse Android managed-store payloadJson' $_.Exception
            return
        }

        # Outlook gets schema-driven translation (port of old DocumentationCustom.psm1:1987-2005).
        if ($obj.packageId -eq 'com.microsoft.office.outlook') {
            $hasAccountType = @($payloadData.managedProperty | Where-Object { $_.key -eq 'com.microsoft.outlook.EmailProfile.AccountType' })
            $outlookSettings = [PSCustomObject]@{ configureEmail = [bool]$hasAccountType }
            foreach ($mp in @($payloadData.managedProperty)) {
                $valueProp = $mp.PSObject.Properties | Where-Object Name -Like 'value*' | Select-Object -First 1
                $outlookSettings | Add-Member -MemberType NoteProperty -Name $mp.key -Value $valueProp.Value -Force
            }
            Invoke-DocAppConfigManifest $outlookSettings (Join-Path (Join-Path $script:AppRootFolder 'Config\ObjectInfo') '#AppConfigOutlookDevice.json') $Context
        }

        # Outlook translation applied above; remaining managedProperty entries
        # fall through to the additional-settings table.
        $addedSettings = Get-DocumentedSettings

        # Friendly names / descriptions / enum labels come from the app's own
        # managed-configuration schema when a tenant is reachable.
        $schema = Get-CDAndroidAppConfigSchema $obj.packageId

        $additionalSettings = @()
        $hasDescription = $false
        foreach ($row in (Expand-AndroidManagedProperties $payloadData.managedProperty '' 0 $schema)) {
            if ($addedSettings | Where-Object EntityKey -EQ $row.Key) { continue }
            if ($row.Description) { $hasDescription = $true }

            $additionalSettings += [PSCustomObject]@{
                Name        = $row.Name
                Key         = $row.Key
                ValueType   = $row.ValueType
                Value       = $row.Value
                Description = $row.Description
                EntityKey   = $row.Key
                Category    = Get-LanguageString 'TACSettings.generalSettings'
                SubCategory = Get-LanguageString 'SettingDetails.additionalConfiguration'
            }
        }
        if ($additionalSettings.Count -gt 0) {
            # Keep the raw key visible next to the friendly name, and only add the
            # description column when the schema actually supplied any.
            $columns = if ($hasDescription) { @('Name','Key','ValueType','Value','Description') } else { @('Name','Key','ValueType','Value') }
            Add-CustomTable 'AdditionalSettings' $columns $additionalSettings -Order 110
        }

        # Permissions table. Portal grid is 4 columns: friendly name, permission
        # state, raw permission name (prefix stripped) and permission group.
        $permissions = @()
        foreach ($p in $obj.permissionActions) {
            $tail = $p.permission.Split('.')[-1]
            $permissionStr = if ($tail) {
                # Language ids drop the underscores (READ_CALENDAR -> readCalendar);
                # PowerShell member lookup is case-insensitive so the raw upper-case
                # form resolves too.
                $lngId = $tail -replace '_',''
                $resolved = Get-LanguageString "AndroidForWorkAppPermissions.Permissions.$lngId" -IgnoreMissing
                if ($resolved) { $resolved } else { $tail }
            } else { $p.permission }

            $actionStr = $p.action
            $resolvedAction = Get-LanguageString "AndroidForWorkAppPermissions.Action.$($p.action)" -IgnoreMissing
            if ($resolvedAction) { $actionStr = $resolvedAction }

            $permissions += [PSCustomObject]@{
                Permission      = $permissionStr
                PermissionState = $actionStr
                PermissionName  = $tail
                PermissionGroup = Get-AndroidPermissionGroup $tail
                EntityKey       = $p.permission
            }
        }
        if ($permissions.Count -gt 0) {
            Add-CustomTable 'Permissions' @('Permission','PermissionState','PermissionName','PermissionGroup') $permissions -Order 115 -LanguageId 'AndroidForWorkAppPermissions.permissionsTitle'
        }
    }
}

# Managed-configuration schema for one Managed Google Play app. The portal uses
# this to show friendly names instead of raw keys, a description column, the real
# data type (choice/multiselect/bundle...) and enum labels via `selections`.
#
#   GET /deviceManagement/androidManagedStoreAppConfigurationSchemas('app:<packageId>')
#
# Returns a hashtable keyed by schemaItemKey. `nestedSchemaItems` carries the
# members of bundles/bundle arrays (linked to their parent by index/parentIndex),
# so nested leaves get friendly names too. Cached per run and per package; offline
# / source-unavailable returns an empty map and every caller degrades to raw keys.
function Get-CDAndroidAppConfigSchema {
    param([string]$PackageId)

    if (-not $PackageId) { return @{} }

    $ctx = Get-CurrentDocumentationContext
    if (-not $ctx.PSObject.Properties['_AndroidAppConfigSchemas']) {
        $ctx | Add-Member -MemberType NoteProperty -Name '_AndroidAppConfigSchemas' -Value (@{}) -Force
    }
    if ($ctx._AndroidAppConfigSchemas.ContainsKey($PackageId)) { return $ctx._AndroidAppConfigSchemas[$PackageId] }

    $map = @{}
    # The schema is generic app metadata (same on every tenant), so this is gated on
    # connectivity only - not on SourceTenantUnavailable.
    if (Test-DocumentationGraphAvailable) {
        try {
            $url = "/deviceManagement/androidManagedStoreAppConfigurationSchemas('app:$PackageId')"
            $resp = Invoke-MSGraphAPI -Url $url -ODataMetadata 'minimal' -NoError
            foreach ($item in @($resp.schemaItems) + @($resp.nestedSchemaItems)) {
                if ($item.schemaItemKey -and -not $map.ContainsKey($item.schemaItemKey)) {
                    $map[$item.schemaItemKey] = $item
                }
            }
        }
        catch {
            Write-LogError "Failed to load Android app configuration schema for $PackageId" $_.Exception
        }
    }
    $ctx._AndroidAppConfigSchemas[$PackageId] = $map
    return $map
}

# Android permission -> permission group. Groups are Android platform constants
# (not localized - the portal renders them verbatim in its 4th grid column).
$script:_androidPermissionGroups = @{
    'READ_CALENDAR' = 'CALENDAR'; 'WRITE_CALENDAR' = 'CALENDAR'
    'CAMERA' = 'CAMERA'
    'READ_CONTACTS' = 'CONTACTS'; 'WRITE_CONTACTS' = 'CONTACTS'; 'GET_ACCOUNTS' = 'CONTACTS'
    'ACCESS_FINE_LOCATION' = 'LOCATION'; 'ACCESS_COARSE_LOCATION' = 'LOCATION'; 'ACCESS_BACKGROUND_LOCATION' = 'LOCATION'
    'RECORD_AUDIO' = 'MICROPHONE'
    'READ_PHONE_STATE' = 'PHONE'; 'CALL_PHONE' = 'PHONE'; 'READ_CALL_LOG' = 'PHONE'; 'WRITE_CALL_LOG' = 'PHONE'
    'ADD_VOICEMAIL' = 'PHONE'; 'USE_SIP' = 'PHONE'; 'PROCESS_OUTGOING_CALLS' = 'PHONE'
    'BODY_SENSORS' = 'SENSORS'; 'BODY_SENSORS_BACKGROUND' = 'SENSORS'
    'SEND_SMS' = 'SMS'; 'RECEIVE_SMS' = 'SMS'; 'READ_SMS' = 'SMS'; 'RECEIVE_WAP_PUSH' = 'SMS'; 'RECEIVE_MMS' = 'SMS'
    'READ_EXTERNAL_STORAGE' = 'STORAGE'; 'WRITE_EXTERNAL_STORAGE' = 'STORAGE'
    'POST_NOTIFICATIONS' = 'NOTIFICATIONS'
    'READ_MEDIA_VIDEO' = 'MEDIA'; 'READ_MEDIA_IMAGES' = 'MEDIA'; 'READ_MEDIA_AUDIO' = 'MEDIA'
    'BLUETOOTH_CONNECT' = 'DEVICES'; 'NEARBY_WIFI_DEVICES' = 'DEVICES'; 'NEARBY_DEVICES' = 'DEVICES'
}

function Get-AndroidPermissionGroup {
    param([string]$PermissionName)
    if (-not $PermissionName) { return $null }
    $script:_androidPermissionGroups[$PermissionName]
}

# Flatten a Google managed-configuration `managedProperty` array into one row per
# LEAF value. The payload supports six value shapes, two of which nest without
# bound (portal JSON-editor schema: valueBool / valueInteger / valueString /
# valueStringArray / valueBundle / valueBundleArray):
#
#   valueBundle       -> { managedProperty: [ ... ] }          rendered as "parent.child"
#   valueBundleArray  -> [ { managedProperty: [...] }, ... ]   rendered as "parent[0].child"
#
# Without this, a bundle rendered as the PowerShell object stringification and a
# bundle array rendered as a bare "," (the -join of an array of objects).
function Expand-AndroidManagedProperties {
    param($ManagedProperties, [string]$Prefix = '', [int]$Depth = 0, $Schema = @{})

    if ($Depth -gt 10) {
        Write-Log 'Android app config: managedProperty nesting deeper than 10 levels; remaining levels not documented' 2
        return
    }

    $ctx = Get-CurrentDocumentationContext

    foreach ($mp in @($ManagedProperties)) {
        if (-not $mp) { continue }
        $key = if ($Prefix) { "$Prefix$($mp.key)" } else { [string]$mp.key }

        $valueProp = $mp.PSObject.Properties | Where-Object Name -Like 'value*' | Select-Object -First 1
        if (-not $valueProp) { continue }

        # Schema is keyed by the app's own schemaItemKey, not by our dotted path
        $schemaItem = $Schema[[string]$mp.key]

        switch ($valueProp.Name) {
            'valueBundle' {
                Expand-AndroidManagedProperties $valueProp.Value.managedProperty "$key." ($Depth + 1) $Schema
            }
            'valueBundleArray' {
                $idx = 0
                foreach ($bundle in @($valueProp.Value)) {
                    Expand-AndroidManagedProperties $bundle.managedProperty "$key[$idx]." ($Depth + 1) $Schema
                    $idx++
                }
            }
            default {
                $val = $valueProp.Value
                # choice / multiselect store the selection VALUE; the schema carries
                # the friendly name for each in `selections`
                if ($schemaItem.selections) {
                    $val = @($val | ForEach-Object {
                        $raw = $_
                        $sel = $schemaItem.selections | Where-Object { "$($_.value)" -eq "$raw" } | Select-Object -First 1
                        if ($sel.name) { $sel.name } else { $raw }
                    })
                }
                if ($val -is [array]) { $val = $val -join $ctx.ObjectSeparator }

                [PSCustomObject]@{
                    Key         = $key
                    Name        = ?? $schemaItem.displayName $key
                    ValueType   = if ($schemaItem.dataType) { $schemaItem.dataType } else { $valueProp.Name.Substring(5) }
                    Value       = $val
                    Description = $schemaItem.description
                }
            }
        }
    }
}

[DocumentationRegistry]::RegisterHandler([AppConfigAndroidStoreDocHandler]::new())
