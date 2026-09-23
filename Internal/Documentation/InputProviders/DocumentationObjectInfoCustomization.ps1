# Documentation customizations - the port of the old Extensions/DocumentationCustom.psm1
# provider. Registered once as a MatchAll ObjectInfo customizer, so its hooks run for
# EVERY object flowing through the Manifest/Profile (ObjectInfo) walk, exactly like the
# old single provider; each hook branches internally on @odata.type and no-ops for
# types it doesn't handle. These augment schema-driven translation and deliberately do
# not claim whole-object handler dispatch (that is what DocumentationHandlerBase is for).

function Set-DocumentationSyntheticProperty {
    param($Obj, [string]$Name, $Value)
    if ($null -eq $Obj) { return }
    $Obj | Add-Member -MemberType NoteProperty -Name $Name -Value $Value -Force
}

function ConvertFrom-DocumentationBase64 {
    param($Value)
    if (-not $Value) { return $Value }
    try { return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]$Value)) }
    catch { return $Value }
}

function Add-DocumentationEncodedScript {
    param([DocumentationContext]$Context, [string]$Header, [string]$Caption, $Content)
    if ($Content) { $Context.AddScript($Header, $Caption, (ConvertFrom-DocumentationBase64 $Content)) }
}

function Initialize-DocumentationCustomObject {
    param($Obj, [DocumentationContext]$Context)
    $type = [string]$Obj.'@odata.type'

    switch ($type) {
        '#microsoft.graph.deviceEnrollmentNotificationConfiguration' {
            # Not a v3 port - there was no v3 support for this type.
            #
            # Almost nothing the portal shows for this policy is on the policy.
            # The two properties that look like they hold the notification channel
            # and its template do not: measured against three policies in a live
            # tenant, templateType comes back as the string "0" (a value the
            # published enum does not even define) and notificationMessageTemplateId
            # is an all-zero GUID. What the policy really stores is
            # notificationTemplates, whose entries are shaped "<Channel>_<templateId>"
            # (e.g. "Push_49d3fb35-..."), one per enabled channel. The subject and
            # body live on that template as its localized messages, and every
            # template carries the same internal display name
            # ("EnrollmentNotificationInternalMEO"), so the name is worth nothing
            # to a reader and the message text is worth everything.
            $templates = Get-CDNotificationMessageTemplates
            foreach ($channel in 'push', 'email') {
                $prefix = "$($channel)_"
                $entry = @($Obj.notificationTemplates) |
                            Where-Object { "$_".StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) } |
                            Select-Object -First 1
                # A channel is on exactly when the policy carries a template for it.
                Set-DocumentationSyntheticProperty $Obj "$($channel)NotificationEnabled" ([bool]$entry)

                $subject = $null
                $body    = $null
                if ($entry) {
                    $templateId = ("$entry" -split '_', 2)[1]
                    $template = $templates | Where-Object { $_.id -eq $templateId } | Select-Object -First 1
                    $messages = @($template.localizedNotificationMessages)
                    # The locale the policy asks for, else the template's own
                    # default, else whatever single message the template has.
                    $message = $null
                    if ($Obj.defaultLocale) { $message = $messages | Where-Object { $_.locale -eq $Obj.defaultLocale } | Select-Object -First 1 }
                    if (-not $message) { $message = $messages | Where-Object { $_.isDefault } | Select-Object -First 1 }
                    if (-not $message) { $message = $messages | Select-Object -First 1 }
                    $subject = $message.subject
                    $body    = $message.messageTemplate
                }
                Set-DocumentationSyntheticProperty $Obj "$($channel)NotificationSubject" $subject
                Set-DocumentationSyntheticProperty $Obj "$($channel)NotificationMessage" $body
            }

            # The portal renders brandingOptions as one On/Off toggle per flag under
            # Email Header / Email Footer rather than as a single combined value, so
            # the document does too. These come from the POLICY: the templates in a
            # live tenant carry brandingOptions of their own and the portal ignores
            # them (a policy reading "none" shows every toggle off even where its
            # email template sets all five).
            $branding = @([string]$Obj.brandingOptions -split ',' | ForEach-Object { $_.Trim() })
            $brandingFlags = [ordered]@{
                brandingIncludeCompanyLogo        = 'includeCompanyLogo'
                brandingIncludeCompanyName        = 'includeCompanyName'
                brandingIncludeContactInformation = 'includeContactInformation'
                brandingIncludeCompanyPortalLink  = 'includeCompanyPortalLink'
                brandingIncludeDeviceDetails      = 'includeDeviceDetails'
            }
            foreach ($propertyName in @($brandingFlags.Keys)) {
                Set-DocumentationSyntheticProperty $Obj $propertyName ($branding -contains $brandingFlags[$propertyName])
            }
        }
        '#microsoft.graph.androidWorkProfileGeneralDeviceConfiguration' {
            $package = $Obj.vpnAlwaysOnPackageIdentifier
            $known = @('com.cisco.anyconnect.vpn.android.avf','com.f5.edge.client_ics','com.paloaltonetworks.globalprotect','net.pulsesecure.pulsesecure')
            Set-DocumentationSyntheticProperty $Obj 'vpnAlwaysOnPackageIdentifierSelector' $(if (-not $package) { $null } elseif ($package -in $known) { $package } else { 'custom' })
            Set-DocumentationSyntheticProperty $Obj 'vpnAlwaysOnEnabled' (-not [string]::IsNullOrEmpty($package))
        }
        '#microsoft.graph.androidDeviceOwnerGeneralDeviceConfiguration' {
            $package = $Obj.vpnAlwaysOnPackageIdentifier
            $known = @('com.cisco.anyconnect.vpn.android.avf','com.f5.edge.client_ics','com.paloaltonetworks.globalprotect','net.pulsesecure.pulsesecure')
            Set-DocumentationSyntheticProperty $Obj 'vpnAlwaysOnPackageIdentifierSelector' $(if (-not $package) { $null } elseif ($package -in $known) { $package } else { 'custom' })
            Set-DocumentationSyntheticProperty $Obj 'vpnAlwaysOnEnabled' (-not [string]::IsNullOrEmpty($package))
            Set-DocumentationSyntheticProperty $Obj 'globalProxyEnabled' ($null -ne $Obj.globalProxy)
            if ($Obj.globalProxy.proxyAutoConfigURL) {
                Set-DocumentationSyntheticProperty $Obj 'globalProxyTypeSelector' 'proxyAutoConfig'
                Set-DocumentationSyntheticProperty $Obj 'globalProxyProxyAutoConfigURL' $Obj.globalProxy.proxyAutoConfigURL
            }
            elseif ($Obj.globalProxy.host) {
                Set-DocumentationSyntheticProperty $Obj 'globalProxyTypeSelector' 'direct'
                Set-DocumentationSyntheticProperty $Obj 'globalProxyHost' $Obj.globalProxy.host
                Set-DocumentationSyntheticProperty $Obj 'globalProxyPort' $Obj.globalProxy.port
                Set-DocumentationSyntheticProperty $Obj 'globalProxyExcludedHosts' $Obj.globalProxy.excludedHosts
            }
            if ($Obj.PSObject.Properties['factoryResetDeviceAdministratorEmails']) {
                Set-DocumentationSyntheticProperty $Obj 'factoryResetProtections' $(if (@($Obj.factoryResetDeviceAdministratorEmails).Count) { 'factoryResetProtectionEnabled' } else { 'factoryResetProtectionDisabled' })
                Set-DocumentationSyntheticProperty $Obj 'googleAccountEmailAddressesList' (@($Obj.factoryResetDeviceAdministratorEmails) -join $Context.ObjectSeparator)
            }
            if ($Obj.PSObject.Properties['passwordBlockKeyguardFeatures']) { Set-DocumentationSyntheticProperty $Obj 'passwordBlockKeyguardFeaturesList' $Obj.passwordBlockKeyguardFeatures }
            if ($Obj.PSObject.Properties['stayOnModes']) { Set-DocumentationSyntheticProperty $Obj 'stayOnModesList' $Obj.stayOnModes }
            if ($Obj.PSObject.Properties['playStoreMode']) { Set-DocumentationSyntheticProperty $Obj 'publicPlayStoreEnabled' ($Obj.playStoreMode -eq 'blockList') }
        }
        '#microsoft.graph.androidEasEmailProfileConfiguration' {
            Set-DocumentationSyntheticProperty $Obj 'domainNameSourceType' $(if ($null -ne $Obj.customDomainName) { 'CustomDomainName' } else { 'AAD' })
        }
        '#microsoft.graph.windowsDeliveryOptimizationConfiguration' {
            Set-DocumentationSyntheticProperty $Obj 'groupIdSourceSelector' $(if ($Obj.groupIdSource.groupIdSourceOption) { $Obj.groupIdSource.groupIdSourceOption } else { 'notConfigured' })
        }
        '#microsoft.graph.androidManagedAppProtection' { Initialize-ManagedAppProtectionSyntheticProperties $Obj }
        '#microsoft.graph.iosManagedAppProtection' { Initialize-ManagedAppProtectionSyntheticProperties $Obj }
        '#microsoft.graph.defaultManagedAppProtection' { Initialize-ManagedAppProtectionSyntheticProperties $Obj }
        '#microsoft.graph.windowsUpdateForBusinessConfiguration' {
            Set-DocumentationSyntheticProperty $Obj 'useDeadLineSettings' ($null -ne $Obj.deadlineForFeatureUpdatesInDays -or $null -ne $Obj.deadlineForQualityUpdatesInDays -or $null -ne $Obj.deadlineGracePeriodInDays -or $null -ne $Obj.postponeRebootUntilAfterDeadline)
        }
        '#microsoft.graph.azureADWindowsAutopilotDeploymentProfile' { Initialize-AutopilotSyntheticProperties $Obj 'azureAD' }
        '#microsoft.graph.activeDirectoryWindowsAutopilotDeploymentProfile' { Initialize-AutopilotSyntheticProperties $Obj 'hybrid' }
        '#microsoft.graph.officeSuiteApp' {
            Set-DocumentationSyntheticProperty $Obj 'VersionToInstall' $(if ([string]::IsNullOrEmpty($Obj.targetVersion)) { Get-LanguageString 'SettingDetails.latest' } else { $Obj.targetVersion })
            Set-DocumentationSyntheticProperty $Obj 'useMicrosoftSearchAsDefault' ($Obj.excludedApps.bing -eq $false)
            if ($Obj.officeConfigurationXml) { Set-DocumentationSyntheticProperty $Obj 'MSAppsConfigXml' (ConvertFrom-DocumentationBase64 $Obj.officeConfigurationXml) }
        }
        '#microsoft.graph.win32LobApp' { Initialize-Win32LobAppSyntheticProperties $Obj $Context }
        '#microsoft.graph.iosGeneralDeviceConfiguration' { Initialize-IosGeneralSyntheticProperties $Obj $Context }
        '#microsoft.graph.iosUpdateConfiguration' { Initialize-IosUpdateSyntheticProperties $Obj $Context }
        '#microsoft.graph.windowsWifiEnterpriseEAPConfiguration' { Initialize-WifiEnterpriseEapSyntheticProperties $Obj }
        '#microsoft.graph.windowsKioskConfiguration' { Initialize-WindowsKioskSyntheticProperties $Obj }
        '#microsoft.graph.windows10VpnConfiguration' {
            Set-DocumentationSyntheticProperty $Obj 'syntheticWipOrApps' $(if ($Obj.windowsInformationProtectionDomain) { 1 } elseif ($Obj.onlyAssociatedAppsCanUseConnection) { 2 } else { 0 })
            if ($null -eq $Obj.profileTarget) { $Obj.profileTarget = 'user' }
        }
        '#microsoft.graph.androidDeviceOwnerVpnConfiguration' { Initialize-AndroidVpnSyntheticProperties $Obj $Context }
        '#microsoft.graph.androidForWorkVpnConfiguration'     { Initialize-AndroidVpnSyntheticProperties $Obj $Context }
        '#microsoft.graph.androidWorkProfileVpnConfiguration' { Initialize-AndroidVpnSyntheticProperties $Obj $Context }
        '#microsoft.graph.iosDeviceFeaturesConfiguration' {
            foreach ($configuration in @($Obj.iosSingleSignOnExtension.configurations)) {
                Set-DocumentationSyntheticProperty $configuration 'configurationKey' $configuration.key
                Set-DocumentationSyntheticProperty $configuration 'configurationValue' $configuration.value
                $configurationType = switch ($configuration.'@odata.type') {
                    '#microsoft.graph.keyStringValuePair'  { Get-LanguageString 'SettingDetails.singleSignOnExtensionConfigurationsTypeColumnOptionString' }
                    '#microsoft.graph.keyBooleanValuePair' { Get-LanguageString 'SettingDetails.singleSignOnExtensionConfigurationsTypeColumnOptionBoolean' }
                    '#microsoft.graph.keyIntegerValuePair' { Get-LanguageString 'SettingDetails.singleSignOnExtensionConfigurationsTypeColumnOptionInteger' }
                    default { 'notConfigured' }
                }
                Set-DocumentationSyntheticProperty $configuration 'configurationType' $configurationType
            }
            Set-DocumentationSyntheticProperty $Obj 'kerberosPrincipalName' $(if ($Obj.singleSignOnSettings.kerberosPrincipalName) { $Obj.singleSignOnSettings.kerberosPrincipalName } else { 'notConfigured' })
            Set-DocumentationSyntheticProperty $Obj 'singleSignOnExtensionType' $(if ($Obj.iosSingleSignOnExtension.'@odata.type') { $Obj.iosSingleSignOnExtension.'@odata.type' } else { 'notConfigured' })
        }
        '#microsoft.graph.macOSDeviceFeaturesConfiguration' {
            Set-DocumentationSyntheticProperty $Obj 'singleSignOnExtensionType' $(if ($Obj.macOSSingleSignOnExtension.'@odata.type') { $Obj.macOSSingleSignOnExtension.'@odata.type' } else { 'notConfigured' })
        }
        '#microsoft.graph.deviceHealthScript' {
            Set-DocumentationSyntheticProperty $Obj 'detectionScriptAdded' (-not [string]::IsNullOrEmpty($Obj.detectionScriptContent))
            Set-DocumentationSyntheticProperty $Obj 'remediationScriptAdded' (-not [string]::IsNullOrEmpty($Obj.remediationScriptContent))
            Set-DocumentationSyntheticProperty $Obj 'useLoggedOnCredentials' ($Obj.runAsAccount -ne 'system')
            if ($Obj.detectionScriptContent) {
                $decoded = ConvertFrom-DocumentationBase64 $Obj.detectionScriptContent
                Set-DocumentationSyntheticProperty $Obj 'detectionScriptContentString' $decoded
                $header = Get-LanguageString 'ProactiveRemediations.Create.Settings.DetectionScriptMultiLineTextBox.label'
                $Context.AddScript($header, "$header - $($Obj.displayName)", $decoded)
            }
            if ($Obj.remediationScriptContent) {
                $decoded = ConvertFrom-DocumentationBase64 $Obj.remediationScriptContent
                Set-DocumentationSyntheticProperty $Obj 'remediationScriptContentString' $decoded
                $header = Get-LanguageString 'ProactiveRemediations.Create.Settings.RemediationScriptMultiLineTextBox.label'
                $Context.AddScript($header, "$header - $($Obj.displayName)", $decoded)
            }
        }
        '#microsoft.graph.hardwareConfiguration' {
            $vendor = Get-LanguageString "HardwareConfig.Settings.Tab.HardwareDropDown.$($Obj.hardwareConfigurationFormat)" $Obj.hardwareConfigurationFormat -IgnoreMissing
            Set-DocumentationSyntheticProperty $Obj 'vendor' $vendor
            if ($Obj.configurationFileContent) {
                $decoded = ConvertFrom-DocumentationBase64 $Obj.configurationFileContent
                Set-DocumentationSyntheticProperty $Obj 'configurationFileContentString' $decoded
                $header = Get-LanguageString 'HardwareConfig.Settings.Tab.Configurations.perDevicePassword'
                $Context.AddScript($header, "$header - $($Obj.displayName)", $decoded)
            }
        }
        '#microsoft.graph.deviceManagementScript' { Add-DocumentationEncodedScript $Context $Obj.fileName "$($Obj.displayName) - $(Get-LanguageString 'WindowsManagement.powerShellScriptObjectName')" $Obj.scriptContent }
        '#microsoft.graph.deviceShellScript' { Add-DocumentationEncodedScript $Context $Obj.fileName "$($Obj.displayName) - $(Get-LanguageString 'WindowsManagement.shellScriptObjectName')" $Obj.scriptContent }
        '#microsoft.graph.deviceCustomAttributeShellScript' { Add-DocumentationEncodedScript $Context $Obj.fileName "$($Obj.displayName) - $(Get-LanguageString 'WindowsManagement.customAttributeObjectName')" $Obj.scriptContent }
        '#microsoft.graph.windows10EndpointProtectionConfiguration' { Initialize-EndpointProtectionSyntheticProperties $Obj }
        '#microsoft.graph.windows10GeneralConfiguration' { Initialize-Windows10GeneralSyntheticProperties $Obj $Context }
        '#microsoft.graph.windowsFeatureUpdateProfile' { Initialize-FeatureUpdateSyntheticProperties $Obj $Context }
        '#microsoft.graph.windows10EnrollmentCompletionPageConfiguration' { Initialize-EnrollmentStatusPageSyntheticProperties $Obj $Context }
        '#microsoft.graph.macOSEndpointProtectionConfiguration' {
            Set-DocumentationSyntheticProperty $Obj 'firewallAllowedApps' @($Obj.firewallApplications | Where-Object allowsIncomingConnections -EQ $true)
            Set-DocumentationSyntheticProperty $Obj 'firewallBlockedApps' @($Obj.firewallApplications | Where-Object allowsIncomingConnections -EQ $false)
        }
        '#microsoft.graph.windows10TeamGeneralConfiguration' {
            Set-DocumentationSyntheticProperty $Obj 'syntheticAzureOperationalInsightsEnabled' ($Obj.azureOperationalInsightsBlockTelemetry -eq $false)
            Set-DocumentationSyntheticProperty $Obj 'syntheticMaintenanceWindowEnabled' ($Obj.maintenanceWindowBlocked -eq $false)
        }
        '#microsoft.graph.windowsWifiConfiguration' {
            if ($Obj.wifiSecurityType -eq 'wpa2Personal') { $Obj.preSharedKey = '********' }
        }
    }

    if ($Obj.PSObject.Properties['securityRequireSafetyNetAttestationBasicIntegrity'] -and $Obj.PSObject.Properties['securityRequireSafetyNetAttestationCertifiedDevice']) {
        $value = if ($Obj.securityRequireSafetyNetAttestationBasicIntegrity -and $Obj.securityRequireSafetyNetAttestationCertifiedDevice) { 'basicIntegrityAndCertified' }
                 elseif ($Obj.securityRequireSafetyNetAttestationBasicIntegrity) { 'basicIntegrity' } else { 'notConfigured' }
        Set-DocumentationSyntheticProperty $Obj 'androidSafetyNetAttestationOptions' $value
    }
    Initialize-ConditionalLaunchSyntheticProperties $Obj

    # Default enrollment policies: title the document by which policy it is.
    #
    # Five enrollment types ship a tenant default named "All users and all
    # devices" - device limit, platform restrictions, enrollment status page,
    # Windows Hello for Business and Windows Restore - so a document headed by the
    # display name does not say which policy it is. The name is still documented, as
    # the Name row inside the table.
    #
    # All five opt in, not just the two documented from a hand-written manifest:
    # they now share one second-level heading
    # (Core/DocumentationEnrollmentGrouping.ps1), which turns the identical titles
    # from a cosmetic problem into three indistinguishable siblings under the same
    # heading. The title, the sort order and the "is this the tenant default" test
    # come from that one file so the heading and its children cannot drift apart.
    if ($Context) {
        $enrollmentName = Get-DocumentationEnrollmentDefaultName $Context.PolicyObject
        if ($enrollmentName) { $Context.DocumentName = $enrollmentName }
    }

    # MergeBasicInfo stays limited to the two types whose settings come from a
    # hand-written manifest. It changes the table layout, which the retitling above
    # does not, and the other three render their own tables correctly already.
    $mergeBasicInfoTypes = @(
        '#microsoft.graph.deviceEnrollmentWindowsHelloForBusinessConfiguration'
        '#microsoft.graph.windowsRestoreDeviceEnrollmentConfiguration'
    )
    if ($type -in $mergeBasicInfoTypes -and $Context) {
        $Context.MergeBasicInfo = $true
    }
}

function Initialize-AndroidVpnSyntheticProperties {
    param($Obj, $Context)

    Initialize-AndroidVpnCustomSettings $Obj
    Initialize-AndroidVpnTunnelSiteName $Obj $Context
}

function Initialize-AndroidVpnCustomSettings {
    param($Obj)

    # The portal's Microsoft Defender "Custom settings" table (Configuration key /
    # Value type / Configuration value) has no matching Graph property. Graph stores it
    # inside customData: each customData entry's value is itself a JSON array of
    # {key,type,value} rows, where type is "int"/"string"/"bool". Parse those rows into a
    # synthetic sharedVpnCustomData collection matching the dataType-21 table declared in
    # the androiddeviceownervpn / androidforworkvpn ObjectInfo files. The type column is
    # translated here (int->Integer, ...) because the table walker does not translate
    # option columns - same reason the iOS SSO configurationType is pre-translated above.
    # (customData is also consumed raw by the netMotionMobility key/value table, but that
    # renders under a different connection-type option, so there is no conflict.)
    if (-not $Obj.customData) { return }

    $typeStrings = @{
        'int'    = Get-LanguageString 'SettingDetails.integer'
        'string' = Get-LanguageString 'SettingDetails.string'
        'bool'   = Get-LanguageString 'SettingDetails.boolean'
    }

    $rows = @()
    foreach ($entry in @($Obj.customData)) {
        if ([string]::IsNullOrEmpty($entry.value)) { continue }
        $parsed = $null
        try { $parsed = $entry.value | ConvertFrom-Json -ErrorAction Stop } catch { continue }
        foreach ($item in @($parsed)) {
            if ($null -eq $item -or -not $item.PSObject.Properties['key']) { continue }
            $typeValue = if ($item.type -and $typeStrings.ContainsKey([string]$item.type)) { $typeStrings[[string]$item.type] } else { $item.type }
            $rows += [PSCustomObject]@{
                key   = $item.key
                type  = $typeValue
                value = $item.value
            }
        }
    }

    if ($rows.Count -gt 0) {
        Set-DocumentationSyntheticProperty $Obj 'sharedVpnCustomData' $rows
    }
}

function Initialize-AndroidVpnTunnelSiteName {
    param($Obj, $Context)

    # "Change the site" (microsoftTunnelSiteId) is stored as a bare GUID. The portal shows
    # the tunnel site's friendly name instead, so resolve the id against the tenant's
    # microsoftTunnelSites and overwrite the property with the display name. The site list
    # is fetched once and cached on the context for the run (same lazy pattern as the
    # Feature Update catalog / iOS update versions). Offline (no Graph) the raw id is kept.
    if (-not $Obj.microsoftTunnelSiteId -or -not $Context) { return }
    if (-not (Test-DocumentationGraphAvailable)) { return }

    if (-not $Context.PSObject.Properties['_MicrosoftTunnelSites']) {
        $Context | Add-Member -MemberType NoteProperty -Name '_MicrosoftTunnelSites' -Value $null -Force
    }
    if ($null -eq $Context._MicrosoftTunnelSites) {
        try { $Context._MicrosoftTunnelSites = @((Invoke-MSGraphAPI -Url '/deviceManagement/microsoftTunnelSites').value) }
        catch { $Context._MicrosoftTunnelSites = @(); Write-LogError 'Failed to load microsoftTunnelSites' $_.Exception }
    }

    $site = $Context._MicrosoftTunnelSites | Where-Object id -EQ $Obj.microsoftTunnelSiteId | Select-Object -First 1
    if ($site.displayName) {
        Set-DocumentationSyntheticProperty $Obj 'microsoftTunnelSiteId' $site.displayName
    }
}

function Initialize-ManagedAppProtectionSyntheticProperties {
    param($Obj)
    Set-DocumentationSyntheticProperty $Obj 'overrideFingerprint' ($null -ne $Obj.pinRequiredInsteadOfBiometricTimeout -and $Obj.pinRequiredInsteadOfBiometricTimeout -ne 'PT0S')
    Set-DocumentationSyntheticProperty $Obj 'pinReset' ($null -ne $Obj.periodBeforePinReset -and $Obj.periodBeforePinReset -notin @('PT0S', 'P0D'))
    # Portal infers "unmanaged browser" from a configured custom browser: package id on
    # Android, protocol on iOS (managedBrowser itself stays notConfigured in that case).
    Set-DocumentationSyntheticProperty $Obj 'managedBrowserSelection' $(if ($Obj.customBrowserPackageId -or $Obj.customBrowserProtocol) { 'unmanagedBrowser' } else { $Obj.managedBrowser })
    Set-DocumentationSyntheticProperty $Obj 'encryptOrgData' ($Obj.appDataEncryptionType -ne 'useDeviceSettings')
    # Portal Purview toggle: On when purviewContentEvaluationRequired is set and not notRequired
    # (null and notRequired both render as Off in the portal).
    Set-DocumentationSyntheticProperty $Obj 'purviewEvaluationEnabled' ($null -ne $Obj.purviewContentEvaluationRequired -and $Obj.purviewContentEvaluationRequired -ne 'notRequired')
    # Portal "Send org data to other apps": managedApps refines into the two Open-In variants
    # via the iOS-only booleans (absent on Android, so the raw value passes through there).
    $sendData = $Obj.allowedOutboundDataTransferDestinations
    if ($sendData -eq 'managedApps') {
        if ($Obj.disableProtectionOfManagedOutboundOpenInData -eq $false -and $Obj.filterOpenInToOnlyManagedApps -eq $true) { $sendData = 'managedAppsWithOpenInSharing' }
        elseif ($Obj.disableProtectionOfManagedOutboundOpenInData -eq $true -and $Obj.filterOpenInToOnlyManagedApps -eq $false) { $sendData = 'managedAppsWithOSSharing' }
    }
    Set-DocumentationSyntheticProperty $Obj 'sendDataSelector' $sendData
    # Portal "Receive data from other apps": the 4th option "All apps with incoming org data"
    # is encoded as allowedInboundDataTransferSources=allApps + protectInboundDataFromUnknownSources=true.
    $receiveData = $Obj.allowedInboundDataTransferSources
    if ($receiveData -eq 'allApps' -and $Obj.protectInboundDataFromUnknownSources -eq $true) { $receiveData = 'allAppsWithIncomingOrgData' }
    Set-DocumentationSyntheticProperty $Obj 'receiveDataSelector' $receiveData
    # Portal "Biometrics instead of PIN for access" (Android + Default): fingerprintAndBiometricEnabled
    # is authoritative when non-null; otherwise derived from the two legacy blocked flags
    # (enabled unless BOTH block). The portal only ever writes fingerprintAndBiometricEnabled.
    $biometric = $Obj.fingerprintAndBiometricEnabled
    if ($null -eq $biometric) { $biometric = (-not $Obj.biometricAuthenticationBlocked) -or (-not $Obj.fingerprintBlocked) }
    Set-DocumentationSyntheticProperty $Obj 'fingerprintAndBiometricEnabledResolved' $biometric
}

function Initialize-AutopilotSyntheticProperties {
    param($Obj, [string]$JoinType)
    Set-DocumentationSyntheticProperty $Obj 'applyDeviceNameTemplate' (-not [string]::IsNullOrEmpty($Obj.deviceNameTemplate))
    if (-not $Obj.outOfBoxExperienceSettings) { return }
    Set-DocumentationSyntheticProperty $Obj.outOfBoxExperienceSettings 'azureADJoinType' $JoinType
    Set-DocumentationSyntheticProperty $Obj.outOfBoxExperienceSettings 'isLanguageSet' (-not [string]::IsNullOrEmpty($Obj.language))
    if ([string]::IsNullOrEmpty($Obj.language)) { $Obj.language = 'user-select' }
}

function Initialize-Windows10GeneralSyntheticProperties {
    param($Obj, [DocumentationContext]$Context)
    Set-DocumentationSyntheticProperty $Obj 'syntheticDefenderDetectedMalwareActionsEnabled' ($null -ne $Obj.defenderDetectedMalwareActions)
    Set-DocumentationSyntheticProperty $Obj 'networkProxyUseScriptUrlName' (-not [string]::IsNullOrEmpty($Obj.networkProxyAutomaticConfigurationUrl))
    Set-DocumentationSyntheticProperty $Obj 'networkProxyUseManualServerName' ($null -ne $Obj.networkProxyServer.address)
    if ($Obj.networkProxyServer.address) {
        $parts = $Obj.networkProxyServer.address -split ':', 2
        Set-DocumentationSyntheticProperty $Obj 'networkProxyServerName' $parts[0]
        Set-DocumentationSyntheticProperty $Obj 'networkProxyServerPort' $(if ($parts.Count -gt 1) { $parts[1] } else { '' })
    }
    Set-DocumentationSyntheticProperty $Obj 'networkProxyExceptionsTextString' (@($Obj.networkProxyServer.exceptions) -join $Context.PropertySeparator)
    Set-DocumentationSyntheticProperty $Obj 'useForLocalAddresses' ($Obj.networkProxyServer.useForLocalAddresses -eq $true)

    # NOTE: reproduced exactly from the old code (edgeDisplayHomeButton is sourced
    # from networkProxyServer.useForLocalAddresses there - preserved for output parity).
    Set-DocumentationSyntheticProperty $Obj 'edgeDisplayHomeButton' ($Obj.networkProxyServer.useForLocalAddresses -eq $true)

    $searchEngineValue = 0
    switch ($Obj.edgeSearchEngine.edgeSearchEngineOpenSearchXmlUrl) {
        'default' { $searchEngineValue = 1 }
        'bing'    { $searchEngineValue = 2 }
        'https://go.microsoft.com/fwlink/?linkid=842596' { $searchEngineValue = 3 }
        'https://go.microsoft.com/fwlink/?linkid=842600' { $searchEngineValue = 4 }
        default   { if ($Obj.edgeSearchEngine.edgeSearchEngineOpenSearchXmlUrl) { $searchEngineValue = 5 } }
    }
    Set-DocumentationSyntheticProperty $Obj 'edgeSearchEngineDropDown' $searchEngineValue

    $curApp = $null
    $perAppPrivacy = @()
    foreach ($appItem in @($Obj.privacyAccessControls | Where-Object { $null -ne $_.appDisplayName })) {
        if ($curApp -ne $appItem.appDisplayName) {
            $perAppPrivacy += [PSCustomObject]@{ appPackageName = $appItem.appPackageFamilyName; appName = $appItem.appDisplayName }
            $curApp = $appItem.appDisplayName
        }
    }
    Set-DocumentationSyntheticProperty $Obj 'perAppPrivacy' $perAppPrivacy
}

function Initialize-EndpointProtectionSyntheticProperties {
    param($Obj)
    $printProps = @($Obj.PSObject.Properties | Where-Object Name -Like 'applicationGuardAllowPrint*')
    Set-DocumentationSyntheticProperty $Obj 'applicationGuardAllowPrinting' (@($printProps | Where-Object Value -EQ $true).Count -gt 0)
    Set-DocumentationSyntheticProperty $Obj 'applicationGuardPrintSettings' @($printProps | Where-Object Value -EQ $true | ForEach-Object Name)

    $fwProps = @($Obj.PSObject.Properties | Where-Object Name -Like 'firewallIPSecExemptionsAllow*')
    Set-DocumentationSyntheticProperty $Obj 'firewallSyntheticPresharedKeyEncodingMethod' (@($fwProps | Where-Object Value -EQ $true).Count -gt 0)
    Set-DocumentationSyntheticProperty $Obj 'firewallSyntheticIPsecExemptions' @($fwProps | Where-Object Value -EQ $true | ForEach-Object Name)

    Set-DocumentationSyntheticProperty $Obj 'firewallSyntheticProfileDomainfirewallEnabled'  ($null -ne $Obj.firewallProfileDomain)
    Set-DocumentationSyntheticProperty $Obj 'firewallSyntheticProfilePrivatefirewallEnabled' ($null -ne $Obj.firewallProfilePrivate)
    Set-DocumentationSyntheticProperty $Obj 'firewallSyntheticProfilePublicfirewallEnabled'  ($null -ne $Obj.firewallProfilePublic)
    Add-DocumentationDefenderFirewallSettings $Obj $Obj.firewallProfileDomain  'Domain'
    Add-DocumentationDefenderFirewallSettings $Obj $Obj.firewallProfilePrivate 'Private'
    Add-DocumentationDefenderFirewallSettings $Obj $Obj.firewallProfilePublic  'Public'

    Set-DocumentationSyntheticProperty $Obj 'bitLockerBaseConfigureEncryptionMethods' $(if ($null -ne $Obj.bitLockerSystemDrivePolicy.encryptionMethod) { $true } else { $null })
    Set-DocumentationSyntheticProperty $Obj 'bitLockerSystemDriveEncryptionMethod'    $Obj.bitLockerSystemDrivePolicy.encryptionMethod
    Set-DocumentationSyntheticProperty $Obj 'bitLockerFixedDriveEncryptionMethod'     $Obj.bitLockerFixedDrivePolicy.encryptionMethod
    Set-DocumentationSyntheticProperty $Obj 'bitLockerRemovableDriveEncryptionMethod' $Obj.bitLockerRemovableDrivePolicy.encryptionMethod

    $sysPol = $Obj.bitLockerSystemDrivePolicy
    Set-DocumentationSyntheticProperty $sysPol 'bitLockerMinimumPinLength' $(if ($null -ne $sysPol.minimumPinLength) { $true } else { $null })
    Set-DocumentationSyntheticProperty $sysPol 'bitLockerSyntheticSystemDrivePolicybitLockerDriveRecovery' $(if ($null -ne $sysPol.recoveryOptions) { $true } else { $null })

    $prebootOption = $null
    if ($null -eq $sysPol.prebootRecoveryUrl -and $null -eq $sysPol.prebootRecoveryEnableMessageAndUrl) { $prebootOption = 'default' }
    elseif ($sysPol.prebootRecoveryUrl -eq '' -and $sysPol.prebootRecoveryEnableMessageAndUrl -eq '') { $prebootOption = 'empty' }
    elseif ($sysPol.prebootRecoveryUrl) { $prebootOption = 'customURL' }
    elseif ($sysPol.prebootRecoveryEnableMessageAndUrl) { $prebootOption = 'customMessage' }
    Set-DocumentationSyntheticProperty $sysPol 'bitLockerPrebootRecoveryMsgURLOption' $prebootOption

    if ($sysPol.recoveryOptions) {
        foreach ($tmpProp in @($sysPol.recoveryOptions.PSObject.Properties.Name)) {
            Set-DocumentationSyntheticProperty $sysPol "bitLockerSyntheticSystemDrivePolicy$tmpProp" $sysPol.recoveryOptions.$tmpProp
        }
    }

    $fixedPol = $Obj.bitLockerFixedDrivePolicy
    Set-DocumentationSyntheticProperty $fixedPol 'bitLockerSyntheticFixedDrivePolicybitLockerDriveRecovery' $(if ($null -ne $fixedPol.recoveryOptions) { $true } else { $null })
    if ($fixedPol.recoveryOptions) {
        foreach ($tmpProp in @($fixedPol.recoveryOptions.PSObject.Properties.Name)) {
            Set-DocumentationSyntheticProperty $fixedPol "bitLockerSyntheticFixedDrivePolicy$tmpProp" $fixedPol.recoveryOptions.$tmpProp
        }
    }

    Set-DocumentationSyntheticProperty $fixedPol 'bitLockerSyntheticFixedDrivePolicyrequireEncryptionForWriteAccess' $fixedPol.requireEncryptionForWriteAccess
    Set-DocumentationSyntheticProperty $Obj.bitLockerRemovableDrivePolicy 'bitLockerSyntheticRemovableDrivePolicyrequireEncryptionForWriteAccess' $Obj.bitLockerRemovableDrivePolicy.requireEncryptionForWriteAccess

    $appLockerType = 'notConfigured'
    if ($Obj.appLockerApplicationControl -eq 'enforceComponentsStoreAppsAndSmartlocker') { $appLockerType = 'allow' }
    if ($Obj.appLockerApplicationControl -eq 'auditComponentsAndStoreApps') { $appLockerType = 'audit' }
    Set-DocumentationSyntheticProperty $Obj 'appLockerApplicationControlType' $appLockerType
}

# Port of the old Add-DefenderFirewallSettings - projects a firewall profile's
# *Blocked/*NotMerged property pairs into firewallSyntheticProfile<Type>* values
# (blocked/allowed/notConfigured) on the top object.
function Add-DocumentationDefenderFirewallSettings {
    param($Obj, $FwSettings, [string]$FwType)
    if (-not $FwSettings) { return }
    foreach ($fwProp in @($FwSettings.PSObject.Properties | Where-Object { $_.Name -like '*Blocked' -or $_.Name -like '*NotMerged' }).Name) {
        if ($fwProp -like '*Blocked') {
            $blockedValue = $FwSettings.$fwProp
            $propPre = $fwProp.Substring(0, $fwProp.Length - 7)
            if ($FwSettings.PSObject.Properties | Where-Object Name -EQ "$($propPre)Required") { $nonBlockedValue = $FwSettings."$($propPre)Required" }
            elseif ($FwSettings.PSObject.Properties | Where-Object Name -EQ "$($propPre)Allowed") { $nonBlockedValue = $FwSettings."$($propPre)Allowed" }
            else { continue }
            $fwPropName = "firewallSyntheticProfile$FwType$propPre"
        }
        else {
            $blockedValue = $FwSettings.$fwProp
            $propPre = $fwProp.Substring(0, $fwProp.Length - 9)
            if ($FwSettings.PSObject.Properties | Where-Object Name -EQ "$($propPre)Merged") { $nonBlockedValue = $FwSettings."$($propPre)Merged" }
            else { continue }
            $fwPropName = "firewallSyntheticProfile$FwType${propPre}Merge"
        }
        $fwValue = 'notConfigured'
        if ($blockedValue -eq $true -and $nonBlockedValue -eq $false) { $fwValue = 'blocked' }
        elseif ($blockedValue -eq $false -and $nonBlockedValue -eq $true) { $fwValue = 'allowed' }
        Set-DocumentationSyntheticProperty $Obj $fwPropName $fwValue
    }
}

function Initialize-FeatureUpdateSyntheticProperties {
    param($Obj, [DocumentationContext]$Context)

    # Resolve the friendly feature-update name from the (generic, tenant-independent)
    # update catalog, cached on the context for the run. Falls back to the version +
    # "not supported" label when offline or the version isn't in the catalog.
    $verInfoTxt = $null
    if (Test-DocumentationGraphAvailable) {
        if (-not $Context.PSObject.Properties['_Win10FeatureUpdates']) {
            $Context | Add-Member -MemberType NoteProperty -Name '_Win10FeatureUpdates' -Value $null -Force
        }
        if ($null -eq $Context._Win10FeatureUpdates) {
            try { $Context._Win10FeatureUpdates = @((Invoke-MSGraphAPI -Url '/deviceManagement/windowsUpdateCatalogItems/microsoft.graph.windowsFeatureUpdateCatalogItem').value) }
            catch { $Context._Win10FeatureUpdates = @(); Write-LogError 'Failed to load windowsFeatureUpdateCatalogItems' $_.Exception }
        }
        $verInfo = $Context._Win10FeatureUpdates | Where-Object version -EQ $Obj.featureUpdateVersion | Select-Object -First 1
        if ($verInfo) { $verInfoTxt = $verInfo.displayName }
    }
    if (-not $verInfoTxt) {
        $verInfoTxt = '{0} ({1})' -f $Obj.featureUpdateVersion, (Get-LanguageString 'WindowsFeatureUpdate.EndOFSupportStatus.notSupported')
    }
    Set-DocumentationSyntheticProperty $Obj 'featureUpdateDisplayName' $verInfoTxt

    if ($Obj.rolloutSettings.offerStartDateTimeInUTC -and $Obj.rolloutSettings.offerEndDateTimeInUTC) {
        Set-DocumentationSyntheticProperty $Obj 'featureUpdateRolloutOption' 'gradualRollout'
        Set-DocumentationSyntheticProperty $Obj 'featureUpdateRolloutStartDate' ([datetime]$Obj.rolloutSettings.offerStartDateTimeInUTC).ToLongDateString()
        Set-DocumentationSyntheticProperty $Obj 'featureUpdateRolloutEndDate' ([datetime]$Obj.rolloutSettings.offerEndDateTimeInUTC).ToLongDateString()
        if ($null -ne $Obj.rolloutSettings.offerIntervalInDays) { Set-DocumentationSyntheticProperty $Obj 'featureUpdateRolloutInterval' $Obj.rolloutSettings.offerIntervalInDays }
    }
    elseif ($Obj.rolloutSettings.offerStartDateTimeInUTC) {
        Set-DocumentationSyntheticProperty $Obj 'featureUpdateRolloutOption' 'startDateOnly'
        Set-DocumentationSyntheticProperty $Obj 'featureUpdateRolloutStartDate' ([datetime]$Obj.rolloutSettings.offerStartDateTimeInUTC).ToLongDateString()
    }
    else {
        Set-DocumentationSyntheticProperty $Obj 'featureUpdateRolloutOption' 'immediateStart'
    }
}

function Initialize-EnrollmentStatusPageSyntheticProperties {
    param($Obj, [DocumentationContext]$Context)
    $timeout = if ([int]$Obj.installProgressTimeoutInMinutes -eq 0) { 60 } else { $Obj.installProgressTimeoutInMinutes }
    Set-DocumentationSyntheticProperty $Obj 'InstallProgressTimeout' $timeout
    Set-DocumentationSyntheticProperty $Obj 'showCustomErrorMessage' (-not [string]::IsNullOrEmpty($Obj.customErrorMessage))
    if (@($Obj.selectedMobileAppIds).Count -eq 0) {
        Set-DocumentationSyntheticProperty $Obj 'waitForApps' (Get-LanguageString 'EnrollmentStatusScreen.Apps.useSelectedAppsAll')
        return
    }
    $apps = Get-CDAllTenantApps
    $names = foreach ($appId in @($Obj.selectedMobileAppIds)) {
        $app = $apps | Where-Object Id -EQ $appId | Select-Object -First 1
        if ($app) { $app.displayName } else { Write-Log "No app found with id $appId" 3 }
    }
    Set-DocumentationSyntheticProperty $Obj 'waitForApps' (@($names) -join $Context.ObjectSeparator)
}

# Map a managedAppRemediationAction enum value (block/wipe/warn/blockWhenSettingIsSupported)
# to its SettingDetails action language key. Returns $null for an unknown value.
function Get-ConditionalLaunchActionId {
    param([string]$Action)
    switch ($Action) {
        'block'                       { 'blockAccess' }
        'wipe'                        { 'wipeData' }
        'warn'                        { 'warn' }
        'blockWhenSettingIsSupported' { 'knoxConditionalLaunchSettingBlockWhenSettingIsSupported' } # "Block access on supported devices"
        default                       { $null }
    }
}

# Device manufacturer / model conditions use the "Allow specified (...)" action labels.
function Get-ConditionalLaunchAllowSpecifiedActionId {
    param([string]$Action)
    switch ($Action) {
        'wipe'  { 'allowSpecifiedWipe' }
        default { 'allowSpecifiedBlock' }
    }
}

function Get-DocumentationConditionalLaunchSetting {
    param($Obj, [string]$SettingId, [string]$Property, [string]$ActionId, [switch]$SkipValue)
    if (-not $ActionId) { return $null }
    $value = $Obj."$Property"
    # Treat every "unconfigured" sentinel as no value: null, empty, notConfigured,
    # none (SafetyNet types), and the empty patch-version placeholder.
    if ($null -eq $value -or "$value" -in @('', 'notConfigured', 'none', '0000-00-00')) { return $null }
    if ($value -is [string] -and $value -match '^P(T|\d)') {
        try {
            $duration = [Xml.XmlConvert]::ToTimeSpan($value)
            $value = if ($Property -eq 'periodOfflineBeforeAccessCheck') { $duration.TotalMinutes }
                     elseif ($Property -eq 'gracePeriodToBlockAppsDuringOffClockHours') { $duration.TotalSeconds }
                     else { $duration.TotalDays }
        }
        catch { }
    }
    elseif ($SkipValue) { $value = $null }
    return [PSCustomObject]@{
        Setting = Get-LanguageString "SettingDetails.$SettingId"
        Value   = $value
        Action  = Get-LanguageString "SettingDetails.$ActionId"
    }
}

function Initialize-ConditionalLaunchSyntheticProperties {
    param($Obj)
    if (-not $Obj.PSObject.Properties['periodOfflineBeforeWipeIsEnforced']) { return }
    $rows = @()

    # ---- Shared device / app conditions (iOS + Android + WIP) ----
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'maxPinAttempts' 'maximumPinRetries' $(if ($Obj.appActionIfMaximumPinRetriesExceeded -eq 'block') { 'resetPin' } else { 'wipeData' })
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'offlineGracePeriod' 'periodOfflineBeforeAccessCheck' 'blockMinutes'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'offlineGracePeriod' 'periodOfflineBeforeWipeIsEnforced' 'wipeDays'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minAppVersion' 'minimumWipeAppVersion' 'wipeData'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minAppVersion' 'minimumRequiredAppVersion' 'blockAccess'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minAppVersion' 'minimumWarningAppVersion' 'warn'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minOSVersion' 'minimumWipeOsVersion' 'wipeData'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minOSVersion' 'minimumRequiredOsVersion' 'blockAccess'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minOSVersion' 'minimumWarningOsVersion' 'warn'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'maxOSVersion' 'maximumWipeOsVersion' 'wipeData'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'maxOSVersion' 'maximumRequiredOsVersion' 'blockAccess'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'maxOSVersion' 'maximumWarningOsVersion' 'warn'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'onlineButUnableToCheckin' 'appActionIfUnableToAuthenticateUser' $(if ($Obj.appActionIfUnableToAuthenticateUser -eq 'block') { 'blockAccess' } else { 'wipeData' }) -SkipValue
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'jailbrokenRootedDevices' 'appActionIfDeviceComplianceRequired' $(if ($Obj.appActionIfDeviceComplianceRequired -eq 'block') { 'blockAccess' } else { 'wipeData' }) -SkipValue
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'maximumAllowedDeviceThreatLevel' 'maximumAllowedDeviceThreatLevel' $(if ($Obj.mobileThreatDefenseRemediationAction -eq 'wipe') { 'wipeData' } else { 'blockAccess' })
    # Primary MTD service - value-only row (no action column in the portal grid)
    $mtdPriority = [string]$Obj.mobileThreatDefensePartnerPriority
    if ($mtdPriority -in @('defenderOverThirdPartyPartner', 'thirdPartyPartnerOverDefender')) {
        $rows += [PSCustomObject]@{
            Setting = Get-LanguageString 'SettingDetails.primaryMtdService' 'Primary MTD service' -IgnoreMissing
            Value   = if ($mtdPriority -eq 'defenderOverThirdPartyPartner') { Get-LanguageString 'SettingDetails.microsoftDefenderForEndpoint' 'Microsoft Defender for Endpoint' -IgnoreMissing }
                      else { Get-LanguageString 'SettingDetails.mobileThreatDefenseNonMicrosoft' }
            Action  = $null
        }
    }
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'clockedOutAccount' 'appActionIfAccountIsClockedOut' (Get-ConditionalLaunchActionId ([string]$Obj.appActionIfAccountIsClockedOut)) -SkipValue
    # "User Clock Status grace period" - duration rendered in seconds, fixed Block action
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'gracePeriodToBlockAppsDuringOffClockHours' 'gracePeriodToBlockAppsDuringOffClockHours' 'blockAccess'

    # ---- iOS specific ----
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minSdkVersion' 'minimumWipeSdkVersion' 'wipeData'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minSdkVersion' 'minimumRequiredSdkVersion' 'blockAccess'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minSdkVersion' 'minimumWarningSdkVersion' 'warn'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'deviceModels' 'allowedIosDeviceModels' (Get-ConditionalLaunchAllowSpecifiedActionId ([string]$Obj.appActionIfIosDeviceModelNotAllowed))

    # ---- Android specific ----
    # Security patch level (Android)
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minPatchVersion' 'minimumWipePatchVersion' 'wipeData'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minPatchVersion' 'minimumRequiredPatchVersion' 'blockAccess'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minPatchVersion' 'minimumWarningPatchVersion' 'warn'
    # Company Portal version (Android)
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minimumCompanyPortalVersion' 'minimumWipeCompanyPortalVersion' 'wipeData'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minimumCompanyPortalVersion' 'minimumRequiredCompanyPortalVersion' 'blockAccess'
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'minimumCompanyPortalVersion' 'minimumWarningCompanyPortalVersion' 'warn'
    # Company Portal update deferral -> "Max Company Portal version age (days)" (Android).
    # One portal setting, up to three rows; each action has its own day-count property. 0 = unset.
    foreach ($cp in @(
            @{ Prop = 'warnAfterCompanyPortalUpdateDeferralInDays';  Action = 'warn' },
            @{ Prop = 'blockAfterCompanyPortalUpdateDeferralInDays'; Action = 'blockAccess' },
            @{ Prop = 'wipeAfterCompanyPortalUpdateDeferralInDays';  Action = 'wipeData' })) {
        if ($Obj."$($cp.Prop)" -gt 0) {
            $rows += Get-DocumentationConditionalLaunchSetting $Obj 'maximumCompanyPortalVersionAge' $cp.Prop $cp.Action
        }
    }
    # Device manufacturers / models (Android)
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'deviceManufacturers' 'allowedAndroidDeviceManufacturers' (Get-ConditionalLaunchAllowSpecifiedActionId ([string]$Obj.appActionIfAndroidDeviceManufacturerNotAllowed))
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'deviceModels' 'allowedAndroidDeviceModels' (Get-ConditionalLaunchAllowSpecifiedActionId ([string]$Obj.appActionIfAndroidDeviceModelNotAllowed))
    # Samsung Knox device attestation (Android)
    $rows += Get-DocumentationConditionalLaunchSetting $Obj 'samsungKnoxAttestationRequired' 'appActionIfSamsungKnoxAttestationRequired' (Get-ConditionalLaunchActionId ([string]$Obj.appActionIfSamsungKnoxAttestationRequired)) -SkipValue

    # SafetyNet / Play Integrity - guard on the required*Type, not the always-set action.
    # Portal rows: "Require threat scan on apps" (apps verification, no value),
    # "Play integrity verdict" (device attestation, value = integrity level) and
    # "Play Integrity verdict evaluation type" (hardware-backed, value-only row).
    if (("$($Obj.requiredAndroidSafetyNetAppsVerificationType)") -notin @('', 'none')) {
        $actId = Get-ConditionalLaunchActionId ([string]$Obj.appActionIfAndroidSafetyNetAppsVerificationFailed)
        if ($actId) {
            $rows += [PSCustomObject]@{
                Setting = Get-LanguageString 'SettingDetails.requireThreatScanOnApps'
                Value   = $null
                Action  = Get-LanguageString "SettingDetails.$actId"
            }
        }
    }
    $attType = [string]$Obj.requiredAndroidSafetyNetDeviceAttestationType
    if ($attType -notin @('', 'none')) {
        $attKey = switch ($attType) {
            'basicIntegrity'                       { 'androidPlayIntegrityVerdictBasicIntegrity' }
            'basicIntegrityAndDeviceCertification' { 'androidPlayIntegrityVerdictBasicAndDeviceIntegrity' }
            default                                { $null }
        }
        $actId = Get-ConditionalLaunchActionId ([string]$Obj.appActionIfAndroidSafetyNetDeviceAttestationFailed)
        if ($actId) {
            $rows += [PSCustomObject]@{
                Setting = Get-LanguageString 'SettingDetails.playIntegrityVerdict'
                Value   = if ($attKey) { Get-LanguageString "SettingDetails.$attKey" } else { $attType }
                Action  = Get-LanguageString "SettingDetails.$actId"
            }
        }
        # Evaluation type only surfaces in the portal while an attestation row exists
        if ([string]$Obj.requiredAndroidSafetyNetEvaluationType -eq 'hardwareBacked') {
            $rows += [PSCustomObject]@{
                Setting = Get-LanguageString 'SettingDetails.playIntegrityVerdictEvaluationType'
                Value   = Get-LanguageString 'SettingDetails.requiredAndroidPlayIntegrityVerdictEvaluationTypeHardwareBacked'
                Action  = $null
            }
        }
    }

    # Require device lock / minimum passcode complexity (Android). Only the
    # configured complexity threshold carries an action; appActionIfDeviceLockNotSet
    # defaults to 'block' even when unconfigured, so it is not emitted on its own.
    foreach ($lvl in @(
            @{ Prop = 'appActionIfDevicePasscodeComplexityLessThanLow';    Level = 'low' },
            @{ Prop = 'appActionIfDevicePasscodeComplexityLessThanMedium'; Level = 'medium' },
            @{ Prop = 'appActionIfDevicePasscodeComplexityLessThanHigh';   Level = 'high' })) {
        $act = [string]$Obj."$($lvl.Prop)"
        if ([string]::IsNullOrEmpty($act)) { continue }
        $actId = Get-ConditionalLaunchActionId $act
        if (-not $actId) { continue }
        $rows += [PSCustomObject]@{
            Setting = Get-LanguageString 'SettingDetails.requireDeviceLockComplexityOnApps'
            Value   = Get-LanguageString "SettingDetails.$($lvl.Level)"
            Action  = Get-LanguageString "SettingDetails.$actId"
        }
    }

    $rows = @($rows | Where-Object { $null -ne $_ })
    if ($rows.Count) { Set-DocumentationSyntheticProperty $Obj 'ConditionalLaunchSettings' $rows }
}

function Initialize-IosGeneralSyntheticProperties {
    param($Obj, [DocumentationContext]$Context)

    if ([string]::IsNullOrEmpty($Obj.KioskModeAppTypeDropDown)) {
        $kioskMode = $null
        if ($Obj.kioskModeAppStoreUrl) { $kioskMode = 0 }
        elseif ($Obj.kioskModeManagedAppId) { $kioskMode = 1 }
        elseif ($Obj.kioskModeBuiltInAppId) { $kioskMode = 2 }
        if ($null -ne $kioskMode) { Set-DocumentationSyntheticProperty $Obj 'KioskModeAppTypeDropDown' $kioskMode }
    }

    $mediaRegion = 'notConfigured'
    foreach ($mediaRatingProp in @($Obj.PSObject.Properties | Where-Object { $_.Name -like 'mediaContentRating*' -and $_.Name -notlike '*@odata.type' -and $_.Name -ne 'mediaContentRatingApps' }).Name) {
        if ($null -ne $Obj.$mediaRatingProp) { $mediaRegion = $mediaRatingProp; break }
    }
    Set-DocumentationSyntheticProperty $Obj 'MediaContentRatingRegionSelectorDropDown' $mediaRegion

    $cellularBlock = 'none'
    $roamingBlock  = 'none'
    $tmpRule = $Obj.networkUsageRules | Where-Object cellularDataBlocked -EQ $true
    if ($tmpRule) {
        $cellularBlock = if ($tmpRule.managedApps) { 'choose' } else { 'all' }
        Set-DocumentationSyntheticProperty $Obj 'networkUsageRulesCellularDataList' ($tmpRule.managedApps -join $Context.ObjectSeparator)
    }
    $tmpRule = $Obj.networkUsageRules | Where-Object cellularDataBlockWhenRoaming -EQ $true
    if ($tmpRule) {
        $roamingBlock = if ($tmpRule.managedApps) { 'choose' } else { 'all' }
        Set-DocumentationSyntheticProperty $Obj 'networkUsageRulesCellularRoamingDataList' $tmpRule.managedApps
    }
    Set-DocumentationSyntheticProperty $Obj 'networkUsageRulesCellularDataBlockType' $cellularBlock
    Set-DocumentationSyntheticProperty $Obj 'networkUsageRulesCellularRoamingDataBlockType' $roamingBlock
}

function Get-DocumentationIosUpdateTimeLabel {
    param([string]$Time, [string[]]$HourWords)
    if ([string]::IsNullOrEmpty($Time)) { return '' }
    $hour = [int]($Time.Split(':')[0])
    $when = 'AM'
    if ($hour -gt 12) { $when = 'PM'; $hour = $hour - 12 }
    $hourStr = if ($hour -ge 0 -and $hour -le 11) { $HourWords[$hour] } else { '' }
    return (Get-LanguageString "SettingDetails.$($hourStr)$($when)Option")
}

function Initialize-IosUpdateSyntheticProperties {
    param($Obj, [DocumentationContext]$Context)

    # iOS available-update versions are generic schema (same on every tenant), cached
    # on the context for the run and gated on connectivity only.
    $verInfo = $null
    if (Test-DocumentationGraphAvailable) {
        if (-not $Context.PSObject.Properties['_IosAvailableUpdateVersions']) {
            $Context | Add-Member -MemberType NoteProperty -Name '_IosAvailableUpdateVersions' -Value $null -Force
        }
        if ($null -eq $Context._IosAvailableUpdateVersions) {
            try {
                $vers = @((Invoke-MSGraphAPI -Url '/deviceManagement/deviceConfigurations/getIosAvailableUpdateVersions').value)
                $Context._IosAvailableUpdateVersions = @($vers | Sort-Object -Property productVersion -Descending)
            } catch { $Context._IosAvailableUpdateVersions = @(); Write-LogError 'Failed to load getIosAvailableUpdateVersions' $_.Exception }
        }
        $verInfo = @($Context._IosAvailableUpdateVersions | Where-Object productVersion -EQ $Obj.desiredOsVersion)
    }

    $versionText = '{0} {1}' -f (Get-LanguageString 'SoftwareUpdates.IosUpdatePolicy.Settings.IOSVersion.prefix'), $Obj.desiredOsVersion
    if (-not $verInfo) {
        $versionText = "$versionText ($(Get-LanguageString 'SoftwareUpdates.IosUpdatePolicy.Settings.IOSVersion.noLongerSupported'))"
    }
    elseif ($verInfo[0].productVersion -eq $Obj.desiredOsVersion) {
        $versionText = "$versionText ($(Get-LanguageString 'SoftwareUpdates.IosUpdatePolicy.Settings.IOSVersion.latestUpdate'))"
    }
    Set-DocumentationSyntheticProperty $Obj 'versionInfo' $versionText

    $hourWords = @('twelve','one','two','three','four','five','six','seven','eight','nine','ten','eleven')
    $timeWindows = @()
    foreach ($tw in @($Obj.customUpdateTimeWindows)) {
        $startDay  = Get-LanguageString "SettingDetails.$($tw.startDay)"
        $endDay    = Get-LanguageString "SettingDetails.$($tw.endDay)"
        $startTime = Get-DocumentationIosUpdateTimeLabel $tw.startTime $hourWords
        $endTime   = Get-DocumentationIosUpdateTimeLabel $tw.endTime $hourWords
        $timeWindows += ($startDay + $Context.PropertySeparator + $startTime + $Context.PropertySeparator + $endDay + $Context.PropertySeparator + $endTime)
    }
    # Property name intentionally 'timeWidows' - the ObjectInfo JSON reads that
    # (mis)spelling; preserved for parity.
    Set-DocumentationSyntheticProperty $Obj 'timeWidows' ($timeWindows -join $Context.ObjectSeparator)
}

function Initialize-WifiEnterpriseEapSyntheticProperties {
    param($Obj)
    if ($Obj.authenticationMethod -eq 'derivedCredential') { return }

    $idCertType = $null
    if ($Obj.'#CustomRef_identityCertificateForClientAuthentication' -and $Obj.'@ObjectFromFile' -eq $true) {
        $idCert = $Obj.'#CustomRef_identityCertificateForClientAuthentication'
        $idx = $idCert.IndexOf('|:|')
        if ($idx -gt -1) { $idCertType = $idCert.Substring($idx + 3) }
    }
    elseif (Test-DocumentationGraphAvailable) {
        try {
            $idCert = Invoke-MSGraphAPI -Url $Obj.'identityCertificateForClientAuthentication@odata.navigationLink' -ODataMetadata 'minimal' -NoError
            $idCertType = $idCert.'@odata.type'
        } catch {}
    }

    $clientCertType = $null
    if ($idCertType -like '*Pkcs*') { $clientCertType = 'PKCS certificate' }
    elseif ($idCertType -like '*SCEP*') { $clientCertType = 'SCEP certificate' }
    if ($clientCertType) { $Obj.authenticationMethod = $clientCertType }
}

function Initialize-WindowsKioskSyntheticProperties {
    param($Obj)
    $appCfg = $Obj.kioskProfiles[0].appConfiguration
    if (-not $appCfg) { return }

    $uwpAppType = $null
    if ($appCfg.'@odata.type' -eq '#microsoft.graph.windowsKioskSingleWin32App') {
        $uwpAppType = 'win32App'
        $appCfg.'@odata.type' = '#microsoft.graph.windowsKioskSingleUWPApp'
    }
    elseif ($appCfg.uwpApp.appUserModelId -like 'Microsoft.MicrosoftEdge*') { $uwpAppType = 'edge' }
    elseif ($appCfg.uwpApp.appUserModelId -like 'Microsoft.KioskBrowser*') { $uwpAppType = 'kioskBrowser' }
    elseif ($appCfg.uwpApp.appUserModelId) { $uwpAppType = 'managed' }

    Set-DocumentationSyntheticProperty $appCfg 'uwpAppType' $uwpAppType
    if ($Obj.windowsKioskForceUpdateSchedule) { Set-DocumentationSyntheticProperty $Obj 'hasForceRestart' $true }
}

# Win32 LOB app documenter - port of the old DocumentationCustom.psm1 win32LobApp
# branch. Synthesizes the fields #Applications.json reads (return codes, requirement
# + detection rule tables, dependency/supersedence lists, install/uninstall types,
# win10Release) and emits install/uninstall/requirement/detection scripts.
function Initialize-Win32LobAppSyntheticProperties {
    param($Obj, [DocumentationContext]$Context)

    $sep = $Context.ObjectSeparator

    $returnCodes = @()
    foreach ($rc in @($Obj.returnCodes)) {
        $returnCodes += [PSCustomObject]@{
            returnCode = $rc.returnCode
            type       = (Get-LanguageString "Win32ReturnCodes.CodeTypes.$($rc.type)")
        }
    }

    # Dependencies / supersedence resolved from the relationships API. This is
    # source-tenant-specific (the related app ids only exist on the origin tenant),
    # so gate it like the other source-tenant lookups.
    $dependencyApps = @()
    $supersededApps = @()
    if (($Obj.dependentAppCount -gt 0 -or $Obj.supersededAppCount -gt 0) -and
        -not $Context.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
        try {
            $url = "/deviceAppManagement/mobileApps/$($Obj.Id)/relationships?`$filter=targetType%20eq%20microsoft.graph.mobileAppRelationshipType%27child%27"
            $relationships = (Invoke-MSGraphAPI -Url $url).value
            foreach ($rel in @($relationships)) {
                if ($rel.'@odata.type' -eq '#microsoft.graph.mobileAppDependency') {
                    $key = if ($rel.dependencyType -eq 'autoInstall') { 'win32DependenciesAutoInstall' } else { 'win32DependenciesDetect' }
                    $dependencyApps += ('{0} {1}' -f $rel.targetDisplayName, (Get-LanguageString "SettingDetails.$key"))
                }
                elseif ($rel.'@odata.type' -eq '#microsoft.graph.mobileAppSupersedence') {
                    $key = if ($rel.supersedenceType -eq 'update') { 'win32SupersedenceUpdate' } else { 'win32SupersedenceReplace' }
                    $supersededApps += ('{0} {1}' -f $rel.targetDisplayName, (Get-LanguageString "SettingDetails.$key"))
                }
            }
        }
        catch { Write-LogError "Failed to load Win32 app relationships for $($Obj.Id)" $_.Exception }
    }

    # Install script (present only when the app content was decrypted on export -
    # the '#ScriptInfo' synthetic). Falls back to the command-line installer type.
    if ($Obj.activeInstallScript.'#ScriptInfo') {
        $Obj.installCommandLine = $null
        $installType = 'Win32Program.installScript'
        Set-DocumentationSyntheticProperty $Obj 'activeInstallScript' ([PSCustomObject]@{
            scriptType            = 'install'
            scriptName            = $Obj.activeInstallScript.'#ScriptInfo'.displayName
            scriptContent         = $Obj.activeInstallScript.'#ScriptInfo'.ScriptContent
            enforceSignatureCheck = $Obj.activeInstallScript.'#ScriptInfo'.enforceSignatureCheck
            runAs32Bit            = $Obj.activeInstallScript.'#ScriptInfo'.runAs32Bit
        })
        Add-DocumentationEncodedScript $Context $Obj.activeInstallScript.'#ScriptInfo'.displayName "$($Obj.displayName) - $($Obj.activeInstallScript.'#ScriptInfo'.displayName)" $Obj.activeInstallScript.'#ScriptInfo'.content
    }
    else {
        $installType = 'Win32Program.installCommand'
        Set-DocumentationSyntheticProperty $Obj 'activeInstallScript' $null
    }

    # Uninstall script
    if ($Obj.activeUninstallScript.'#ScriptInfo') {
        $Obj.uninstallCommandLine = $null
        $uninstallType = 'Win32Program.uninstallScript'
        Set-DocumentationSyntheticProperty $Obj 'activeUninstallScript' ([PSCustomObject]@{
            scriptType            = 'uninstall'
            scriptName            = $Obj.activeUninstallScript.'#ScriptInfo'.displayName
            scriptContent         = $Obj.activeUninstallScript.'#ScriptInfo'.ScriptContent
            enforceSignatureCheck = $Obj.activeUninstallScript.'#ScriptInfo'.enforceSignatureCheck
            runAs32Bit            = $Obj.activeUninstallScript.'#ScriptInfo'.runAs32Bit
        })
        Add-DocumentationEncodedScript $Context $Obj.activeUninstallScript.'#ScriptInfo'.displayName "$($Obj.displayName) - $($Obj.activeUninstallScript.'#ScriptInfo'.displayName)" $Obj.activeUninstallScript.'#ScriptInfo'.content
    }
    else {
        $uninstallType = 'Win32Program.uninstallCommand'
        Set-DocumentationSyntheticProperty $Obj 'activeUninstallScript' $null
    }

    # Requirement rules -> summary strings + a flat property/value row set (matches
    # the old behavior: rows from all rules accumulate into one table).
    $requirementRulesSummary = @()
    $requirementRules = @()
    foreach ($rule in @($Obj.requirementRules)) {
        if ($rule.'@odata.type' -eq '#microsoft.graph.win32LobAppFileSystemRequirement') {
            $lngId = 'fileType'; $textValue = $rule.path
        }
        elseif ($rule.'@odata.type' -eq '#microsoft.graph.win32LobAppRegistryRequirement') {
            $lngId = 'registry'; $textValue = $rule.keyPath
        }
        else {
            $lngId = 'script'; $textValue = $rule.displayName
            Add-DocumentationEncodedScript $Context $rule.displayName "$($Obj.displayName) - Requirement script" $rule.scriptContent
        }
        $requirementRulesSummary += ('{0} {1}' -f (Get-LanguageString "Win32Requirements.AdditionalRequirements.RequirementTypeOptions.$lngId"), $textValue)
        $requirementRules += Add-CDDocumentRequirementRule $rule
    }

    # Detection rules
    $detectionRulesSummary = @()
    $detectionRules = @()
    if (@($Obj.detectionRules) | Where-Object '@odata.type' -EQ '#microsoft.graph.win32LobAppPowerShellScriptDetection') {
        $detectionRulesType = Get-LanguageString 'DetectionRules.RuleConfigurationOptions.customScript'
        foreach ($rule in @($Obj.detectionRules)) {
            $header = Get-LanguageString 'ProactiveRemediations.Create.Settings.DetectionScriptMultiLineTextBox.label'
            Add-DocumentationEncodedScript $Context $header "$($Obj.displayName) - $header" $rule.scriptContent
        }
    }
    else {
        $detectionRulesType = Get-LanguageString 'DetectionRules.RuleConfigurationOptions.manual'
        foreach ($rule in @($Obj.detectionRules)) {
            if ($rule.'@odata.type' -eq '#microsoft.graph.win32LobAppFileSystemDetection') {
                $lngId = 'file'; $textValue = $rule.path
            }
            elseif ($rule.'@odata.type' -eq '#microsoft.graph.win32LobAppRegistryDetection') {
                $lngId = 'registry'; $textValue = $rule.keyPath
            }
            else {
                $lngId = 'mSI'; $textValue = $rule.productCode
            }
            $detectionRulesSummary += ('{0} {1}' -f (Get-LanguageString "DetectionRules.Manual.RuleTypeOptions.$lngId"), $textValue)
            $detectionRules += Add-CDDocumentDetectionRule $rule
        }
    }

    Set-DocumentationSyntheticProperty $Obj 'requirementRulesSummary'    ($requirementRulesSummary -join $sep)
    Set-DocumentationSyntheticProperty $Obj 'detectionRulesSummary'      ($detectionRulesSummary -join $sep)
    Set-DocumentationSyntheticProperty $Obj 'dependencyApps'             ($dependencyApps -join $sep)
    Set-DocumentationSyntheticProperty $Obj 'supersededApps'             ($supersededApps -join $sep)
    Set-DocumentationSyntheticProperty $Obj 'detectionRulesType'         $detectionRulesType
    Set-DocumentationSyntheticProperty $Obj 'requirementRulesTranslated' $requirementRules
    Set-DocumentationSyntheticProperty $Obj 'detectionRulesTranslated'   $detectionRules
    Set-DocumentationSyntheticProperty $Obj 'returnCodes'                $returnCodes
    if($Obj.minimumSupportedWindowsRelease -is [string] -and $Obj.minimumSupportedWindowsRelease.length -eq 4) {
        Set-DocumentationSyntheticProperty $Obj 'win10Release' (Get-LanguageString "MinimumOperatingSystem.Windows.V10Release.release$($obj.minimumSupportedWindowsRelease)")
    }
    elseif($Obj.minimumSupportedWindowsRelease -like "Windows10_*") {
        Set-DocumentationSyntheticProperty $Obj 'win10Release' (Get-LanguageString "MinimumOperatingSystem.Windows.V10Release.release$(($Obj.minimumSupportedWindowsRelease.Split('_')[-1]))")
    }
    elseif($Obj.minimumSupportedWindowsRelease -like "Windows11_*") {
        Set-DocumentationSyntheticProperty $Obj 'win10Release' (Get-LanguageString "MinimumOperatingSystem.Windows.V11Release.release$(($Obj.minimumSupportedWindowsRelease.Split('_')[-1]))")
    }
    Set-DocumentationSyntheticProperty $Obj 'installerType'              (Get-LanguageString $installType)
    Set-DocumentationSyntheticProperty $Obj 'uninstallerType'            (Get-LanguageString $uninstallType)

    if($Obj.allowedArchitectures) {
        $archs = @()
        if ($Obj.allowedArchitectures -like "*x86*") { $archs += (Get-LanguageString 'ArchitectureOptions.thirtyTwoBitInstructionSet') }
        if ($Obj.allowedArchitectures -like "*x64*") { $archs += (Get-LanguageString 'ArchitectureOptions.sixtyFourBitInstructionSet') }
        if ($Obj.allowedArchitectures -like "*arm64*") { $archs += (Get-LanguageString 'ArchitectureOptions.arm64InstructionSet') }
        Set-DocumentationSyntheticProperty $Obj 'allowedArchitecturesTranslated' ($archs -join $sep)
    }
    else {
        Set-DocumentationSyntheticProperty $Obj 'allowedArchitecturesTranslated' (Get-LanguageString 'Win32Requirements.allowedArchitecturesNoRadioButton')
    }
}

# Win32 requirement rule -> array of {property, value} rows. Port of the old
# Add-CDDocumentRequirementRule.
function Add-CDDocumentRequirementRule {
    param($Rule)

    $strYes = Get-LanguageString 'SettingDetails.yes'
    $strNo  = Get-LanguageString 'SettingDetails.no'
    $ruleInfo = @()

    if ($Rule.'@odata.type' -eq '#microsoft.graph.win32LobAppFileSystemRequirement') {
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.requirementType'); value = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.RequirementTypeOptions.fileType') }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.FileRule.path'); value = $Rule.path }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.FileRule.fileOrFolder'); value = $Rule.fileOrFolderName }
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.File.property')
            value = switch ($Rule.detectionType) {
                'createdDate'  { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.dateCreated' }
                'modifiedDate' { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.dateModified' }
                'doesNotExist' { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.doesNotExist' }
                'exists'       { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.fileOrFolderExists' }
                'sizeInMB'     { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.sizeInMB' }
                'version'      { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.version' }
                default        { Get-LanguageString 'BooleanActions.notConfigured' }
            }
        }
        if ($Rule.detectionValue -and $Rule.operator) {
            $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.FileRule.operator'); value = (Get-CDDocumentOperatorString $Rule.operator) }
            $detectionValue = $Rule.detectionValue
            if ($Rule.detectionType -eq 'createdDate' -or $Rule.detectionType -eq 'modifiedDate') {
                try { $detectionValue = (Get-Date $Rule.detectionValue).ToString() } catch {}
            }
            $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.FileRule.value'); value = $detectionValue }
        }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.RegistryRule.associatedWith32Bit'); value = $(if ($Rule.check32BitOn64System -eq $true) { $strYes } else { $strNo }) }
    }
    elseif ($Rule.'@odata.type' -eq '#microsoft.graph.win32LobAppRegistryRequirement') {
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.requirementType'); value = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.RequirementTypeOptions.registry') }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.Registry.keyPath'); value = $Rule.keyPath }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.Registry.valueName'); value = $Rule.valueName }
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.Registry.registryRequirement')
            value = switch ($Rule.detectionType) {
                'doesNotExist' { if ($Rule.valueName) { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.valueDoesNotExist' } else { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.keyDoesNotExist' } }
                'exists'       { if ($Rule.valueName) { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.valueExists' } else { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.keyExists' } }
                'integer'      { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.integerComparison' }
                'string'       { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.stringComparison' }
                'version'      { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.versionComparison' }
                default        { Get-LanguageString 'BooleanActions.notConfigured' }
            }
        }
        if ($Rule.detectionValue -and $Rule.operator) {
            $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.Registry.operator'); value = (Get-CDDocumentOperatorString $Rule.operator) }
            $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.RegistryRule.value'); value = $Rule.detectionValue }
        }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.RegistryRule.associatedWith32Bit'); value = $(if ($Rule.check32BitOn64System -eq $true) { $strYes } else { $strNo }) }
    }
    elseif ($Rule.'@odata.type' -eq '#microsoft.graph.win32LobAppPowerShellScriptRequirement') {
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.requirementType'); value = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.RequirementTypeOptions.script') }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.scriptName'); value = $Rule.displayName }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.CustomScript.runAs32Bit'); value = $(if ($Rule.runAs32Bit -eq $true) { $strYes } else { $strNo }) }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.loggedOnCredentials'); value = $(if ($Rule.runAsAccount -ne 'system') { $strYes } else { $strNo }) }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.enforceSignatureCheck'); value = $(if ($Rule.enforceSignatureCheck -eq $true) { $strYes } else { $strNo }) }
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.requirementMethod')
            value = switch ($Rule.detectionType) {
                'string'   { Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.string' }
                'dateTime' { Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.dateTime' }
                'integer'  { Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.integer' }
                'float'    { Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.float' }
                'version'  { Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.version' }
                'boolean'  { Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.RequirementMethodOptions.boolean' }
                default    { Get-LanguageString 'BooleanActions.notConfigured' }
            }
        }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.Registry.operator'); value = (Get-CDDocumentOperatorString $Rule.operator) }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'Win32Requirements.AdditionalRequirements.Script.value'); value = $Rule.detectionValue }
    }
    return $ruleInfo
}

# Win32 detection rule -> array of {property, value} rows. Port of the old
# Add-CDDocumentDetectionRule.
function Add-CDDocumentDetectionRule {
    param($Rule)

    $strYes = Get-LanguageString 'SettingDetails.yes'
    $strNo  = Get-LanguageString 'SettingDetails.no'
    $ruleInfo = @()

    if ($Rule.'@odata.type' -eq '#microsoft.graph.win32LobAppFileSystemDetection') {
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.ruleType'); value = (Get-LanguageString 'DetectionRules.Manual.RuleTypeOptions.file') }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.FileRule.path'); value = $Rule.path }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.FileRule.fileOrFolder'); value = $Rule.fileOrFolderName }
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString 'DetectionRules.Manual.FileRule.detectionMethod')
            value = switch ($Rule.detectionType) {
                'createdDate'  { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.dateCreated' }
                'modifiedDate' { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.dateModified' }
                'doesNotExist' { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.doesNotExist' }
                'exists'       { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.fileOrFolderExists' }
                'sizeInMB'     { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.sizeInMB' }
                'version'      { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.version' }
                default        { Get-LanguageString 'BooleanActions.notConfigured' }
            }
        }
        if ($Rule.detectionValue -and $Rule.operator) {
            $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.FileRule.operator'); value = (Get-CDDocumentOperatorString $Rule.operator) }
            $detectionValue = $Rule.detectionValue
            if ($Rule.detectionType -eq 'createdDate' -or $Rule.detectionType -eq 'modifiedDate') {
                try { $detectionValue = (Get-Date $Rule.detectionValue).ToString() } catch {}
            }
            $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.FileRule.value'); value = $detectionValue }
        }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.RegistryRule.associatedWith32Bit'); value = $(if ($Rule.check32BitOn64System -eq $true) { $strYes } else { $strNo }) }
    }
    elseif ($Rule.'@odata.type' -eq '#microsoft.graph.win32LobAppRegistryDetection') {
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.ruleType'); value = (Get-LanguageString 'DetectionRules.Manual.RuleTypeOptions.registry') }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.RegistryRule.keyPath'); value = $Rule.keyPath }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.RegistryRule.valueName'); value = $Rule.valueName }
        $ruleInfo += [PSCustomObject]@{
            property = (Get-LanguageString 'DetectionRules.Manual.RegistryRule.detectionMethod')
            value = switch ($Rule.detectionType) {
                'doesNotExist' { if ($Rule.valueName) { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.valueDoesNotExist' } else { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.keyDoesNotExist' } }
                'exists'       { if ($Rule.valueName) { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.valueExists' } else { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.keyExists' } }
                'integer'      { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.integerComparison' }
                'string'       { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.stringComparison' }
                'version'      { Get-LanguageString 'DetectionRules.Manual.RegistryRule.DetectionMethodOptions.versionComparison' }
                default        { Get-LanguageString 'BooleanActions.notConfigured' }
            }
        }
        if ($Rule.detectionValue -and $Rule.operator) {
            $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.RegistryRule.operator'); value = (Get-CDDocumentOperatorString $Rule.operator) }
            $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.RegistryRule.value'); value = $Rule.detectionValue }
        }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.RegistryRule.associatedWith32Bit'); value = $(if ($Rule.check32BitOn64System -eq $true) { $strYes } else { $strNo }) }
    }
    else {
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.ruleType'); value = (Get-LanguageString 'DetectionRules.Manual.RuleTypeOptions.mSI') }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.MsiRule.productCode'); value = $Rule.productCode }
        $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.MsiRule.productVersionCheck'); value = $(if ($null -ne $Rule.productVersion) { $strYes } else { $strNo }) }
        if ($null -ne $Rule.productVersion) {
            $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.MsiRule.operator'); value = (Get-CDDocumentOperatorString $Rule.productVersionOperator) }
            $ruleInfo += [PSCustomObject]@{ property = (Get-LanguageString 'DetectionRules.Manual.MsiRule.productVersion'); value = (Get-CDDocumentOperatorString $Rule.productVersion) }
        }
    }
    return $ruleInfo
}

function Get-CDDocumentOperatorString {
    param($Operator)
    switch ($Operator) {
        'notConfigured'      { Get-LanguageString 'BooleanActions.notConfigured' }
        'equal'              { Get-LanguageString 'DetectionRules.ComparisonOperators.equals' }
        'notEqual'           { Get-LanguageString 'DetectionRules.ComparisonOperators.notEqualTo' }
        'greaterThan'        { Get-LanguageString 'DetectionRules.ComparisonOperators.greaterThan' }
        'greaterThanOrEqual' { Get-LanguageString 'DetectionRules.ComparisonOperators.greaterThanOrEqualTo' }
        'lessThan'           { Get-LanguageString 'DetectionRules.ComparisonOperators.lessThan' }
        'lessThanOrEqual'    { Get-LanguageString 'DetectionRules.ComparisonOperators.lessThanOrEqualTo' }
        'exists'             { Get-LanguageString 'DetectionRules.Manual.FileRule.DetectionMethodOptions.fileOrFolderExists' }
        default              { $Operator }
    }
}

function Get-DocumentationCustomPropertyObject {
    param($TopObj, $Obj, $Prop)
    if ($TopObj.'@odata.type' -ne '#microsoft.graph.windows10EndpointProtectionConfiguration') { return $null }
    switch ($Prop.entityKey) {
        'startupAuthenticationRequired' { return $TopObj.bitLockerSystemDrivePolicy }
        'bitLockerSyntheticFixedDrivePolicyrequireEncryptionForWriteAccess' { return $TopObj.bitLockerFixedDrivePolicy }
        'bitLockerSyntheticRemovableDrivePolicyrequireEncryptionForWriteAccess' { return $TopObj.bitLockerRemovableDrivePolicy }
    }
    return $null
}

function Get-DocumentationCustomChildObject {
    param($TopObj, $Obj, $Prop)
    switch ($TopObj.'@odata.type') {
        '#microsoft.graph.windows10GeneralConfiguration' { if ($Prop.entityKey -eq 'syntheticDefenderDetectedMalwareActionsEnabled') { return $TopObj.defenderDetectedMalwareActions } }
        '#microsoft.graph.iosDeviceFeaturesConfiguration' {
            if ($Prop.entityKey -eq 'kerberosPrincipalName') { return $TopObj.singleSignOnSettings }
            if ($Prop.entityKey -eq 'singleSignOnExtensionType') { return $TopObj.iosSingleSignOnExtension }
        }
        '#microsoft.graph.macOSDeviceFeaturesConfiguration' { if ($Prop.entityKey -eq 'singleSignOnExtensionType') { return $TopObj.macOSSingleSignOnExtension } }
        '#microsoft.graph.windows10EndpointProtectionConfiguration' {
            if ($Prop.entityKey -eq 'applicationGuardPrintSettings') { return $TopObj.applicationGuardPrintSettings }
            if ($Prop.entityKey -eq 'firewallSyntheticIPsecExemptions') { return $TopObj.firewallSyntheticIPsecExemptions }
        }
    }
    return $null
}

function Get-DocumentationCustomProfileValue {
    param($TopObj, $Obj, $Prop, [DocumentationContext]$Context)
    $type = [string]$TopObj.'@odata.type'
    if ($type -eq '#microsoft.graph.windowsDeliveryOptimizationConfiguration' -and $Prop.entityKey -eq 'groupIdSourceSelector') {
        Invoke-TranslateOption $Obj $Prop -SkipOptionChildren | Out-Null
        return $false
    }
    if ($type -in @('#microsoft.graph.androidManagedAppProtection','#microsoft.graph.iosManagedAppProtection') -and $Prop.entityKey -eq 'apps') {
        $customApps, $publishedApps = Get-CDMobileApps $Obj.apps
        Add-PropertyInfo $Prop ($publishedApps -join $Context.ObjectSeparator) ($publishedApps -join $Context.PropertySeparator)
        $info = Get-PropertyInfo $Prop ($customApps -join $Context.ObjectSeparator) ($customApps -join $Context.PropertySeparator)
        $info.Name = Get-LanguageString 'SettingDetails.customApps'
        $info.Description = ''
        Add-PropertyInfoObject $info
        return $false
    }
    if ($type -in @('#microsoft.graph.windowsInformationProtectionPolicy','#microsoft.graph.mdmWindowsInformationProtectionPolicy')) {
        if ($Prop.entityKey -eq 'enterpriseIPRanges') {
            $ranges = foreach ($rangeGroup in @($Obj.enterpriseIPRanges)) {
                $values = foreach ($range in @($rangeGroup.ranges)) { "$($range.lowerAddress)-$($range.upperAddress)" }
                if ($values) { "$($rangeGroup.displayName)$($Context.PropertySeparator)$($values -join $Context.PropertySeparator)" }
            }
            $ipv4 = @($ranges | Where-Object { $_ -match '\.' })
            if ($ipv4.Count) { foreach ($range in $ipv4) { Add-PropertyInfo $Prop $range $range } }
            else { Add-PropertyInfo $Prop $null $null }
            $ipv6 = @($ranges | Where-Object { $_ -match ':' })
            if (-not $ipv6.Count) {
                $info = Get-PropertyInfo $Prop $null $null
                $info.Name = Get-LanguageString 'WipPolicySettings.iPv6Ranges'
                Add-PropertyInfoObject $info
            }
            foreach ($range in $ipv6) {
                $info = Get-PropertyInfo $Prop $range $range
                $info.Name = Get-LanguageString 'WipPolicySettings.iPv6Ranges'
                Add-PropertyInfoObject $info
            }
            return $false
        }
        if ($Prop.entityKey -eq 'enterpriseProxiedDomains') {
            foreach ($domain in @($Obj.enterpriseProxiedDomains)) {
                $value = "$($domain.displayName)$($Context.PropertySeparator)$(@($domain.proxiedDomains.ipAddressOrFQDN) -join $Context.PropertySeparator)"
                Add-PropertyInfo $Prop $value $value
            }
            return $false
        }
    }
    if ($type -like '#microsoft.graph.windows*SCEPCertificateProfile' -and $Prop.entityKey -in @('subjectNameFormat','subjectAlternativeNameType')) { return $false }
    if ($type -eq '#microsoft.graph.windows10GeneralConfiguration') {
        if ($Prop.entityKey -eq 'startMenuAppListVisibility') {
            $value = ([string]$Obj.startMenuAppListVisibility) -replace ',(?! )', ', '
            Invoke-TranslateOption $Obj $Prop -PropValue $value | Out-Null
            return $false
        }
        $privacy = $Obj.privacyAccessControls | Where-Object { $_.dataCategory -eq $Prop.entityKey -and $null -eq $_.appDisplayName } | Select-Object -First 1
        if ($privacy) {
            Invoke-TranslateOption $privacy $Prop -PropValue $privacy.accessLevel | Out-Null
            return $false
        }
    }
    if ($type -eq '#microsoft.graph.windows10EndpointProtectionConfiguration') {
        if ($Prop.entityKey -eq 'applicationGuardEnabled') { return $false }
        if ($Prop.entityKey -eq 'bitLockerRecoveryPasswordRotation') {
            Invoke-TranslateOption $TopObj $Prop | Out-Null
            return $false
        }
    }
    if ($type -eq '#microsoft.graph.windowsHealthMonitoringConfiguration' -and $Prop.entityKey -eq 'configDeviceHealthMonitoringScope' -and ($Prop.options | Where-Object value -EQ 'healthMonitoring')) { return $false }
    if ($type -eq '#microsoft.graph.windows10VpnConfiguration') {
        if ($Prop.entityKey -eq 'enableSplitTunneling' -and $Prop.enabled -eq $false) { return $false }
        if ($Prop.entityKey -eq 'eapXml' -and $Obj.eapXml) {
            $value = ConvertFrom-DocumentationBase64 $Obj.eapXml
            Add-PropertyInfo $Prop $value $value
            return $false
        }
    }
    if ($type -eq '#microsoft.graph.windowsUpdateForBusinessConfiguration' -and $Prop.entityKey -in @('businessReadyUpdatesOnly','autoRestartNotificationDismissal','scheduleRestartWarningInHours','scheduleImminentRestartWarningInMinutes','deliveryOptimizationMode')) { return $false }
    return $null
}

function Finalize-DocumentationCustomObject {
    param($Obj, [DocumentationContext]$Context)
    if ($Obj.'@odata.type' -ne '#microsoft.graph.windows10EndpointProtectionConfiguration') { return }
    foreach ($categoryId in @('bitLocker','xboxServices')) {
        $category = Get-LanguageString "Category.$categoryId"
        if (-not $category) { continue }
        $rows = @($Context.SettingsData | Where-Object Category -EQ $category)
        $configured = $rows | Where-Object {
            if ($categoryId -eq 'bitLocker' -and $_.EntityKey -like 'startupAuthenticationTpm*') { return $false }
            return $null -ne $_.RawValue -and $_.RawValue -ne $_.DefaultValue
        } | Select-Object -First 1
        if ($configured) { continue }
        foreach ($row in $rows) { [void]$Context.SettingsData.Remove($row) }
    }
}

function Invoke-DocumentationCustomPostAddValue {
    param($TopObj, $Prop)
    if ($TopObj.'@odata.type' -ne '#microsoft.graph.windowsUpdateForBusinessConfiguration') { return }
    if ($Prop.entityKey -eq 'featureUpdatesDeferralPeriodInDays') {
        $tmp = [PSCustomObject]@{ nameResourceKey='allowWindows11UpgradeName'; descriptionResourceKey='allowWindows11UpgradeDescription'; entityKey='allowWindows11Upgrade'; dataType=0; booleanActions=109; category=$Prop.category }
        Add-PropertyInfo $tmp (Invoke-TranslateBoolean $TopObj $tmp) $TopObj.allowWindows11Upgrade
    }
    if ($Prop.entityKey -eq 'featureUpdatesRollbackWindowInDays') {
        $configured = $TopObj.businessReadyUpdatesOnly -notin @('businessReadyOnly','all','userDefined')
        $tmp = [PSCustomObject]@{ nameResourceKey='preReleaseBuilds'; descriptionResourceKey='preReleaseBuildsDescription'; entityKey='preReleaseEnabled'; dataType=0; booleanActions=2; category=$Prop.category }
        Add-PropertyInfo $tmp $(if ($configured) { Get-LanguageString 'BooleanActions.enable' } else { Get-LanguageString 'BooleanActions.notConfigured' }) $TopObj.businessReadyUpdatesOnly
        if ($configured) {
            $tmp = [PSCustomObject]@{ nameResourceKey='preReleaseChannel'; descriptionResourceKey='preReleaseBuildsDescription'; entityKey='businessReadyUpdatesOnly'; dataType=0; booleanActions=2; category=$Prop.category }
            Add-PropertyInfo $tmp (Get-LanguageString "SettingDetails.$($TopObj.businessReadyUpdatesOnly)Option") $TopObj.businessReadyUpdatesOnly
        }
    }
}

Add-DocumentationObjectInfoCustomizer ([PSCustomObject]@{
    Name               = 'DocumentationCustomizations'
    # Runs for every object; each hook branches on @odata.type internally (and no-ops
    # for types it doesn't handle), matching the old DocumentationCustom.psm1 provider.
    MatchAll           = $true
    InitializeObject   = { param($Obj, $Context) Initialize-DocumentationCustomObject $Obj $Context }
    GetPropertyObject  = { param($TopObj, $Obj, $Prop, $Context) Get-DocumentationCustomPropertyObject $TopObj $Obj $Prop }
    GetChildObject     = { param($TopObj, $Obj, $Prop, $Context) Get-DocumentationCustomChildObject $TopObj $Obj $Prop }
    GetProfileValue    = { param($TopObj, $Obj, $Prop, $Context) Get-DocumentationCustomProfileValue $TopObj $Obj $Prop $Context }
    PostAddValue       = { param($TopObj, $Prop, $Context) Invoke-DocumentationCustomPostAddValue $TopObj $Prop }
    FinalizeObject     = { param($Obj, $Context) Finalize-DocumentationCustomObject $Obj $Context }
})
