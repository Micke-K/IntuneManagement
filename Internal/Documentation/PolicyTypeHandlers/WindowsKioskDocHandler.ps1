# Windows Kiosk Configuration documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:3923 (Invoke-
# CDDocumentWindowsKioskConfiguration). Claims @odata.type=
# '#microsoft.graph.windowsKioskConfiguration'.
#
# Generic Profile/walker can't handle this type because the actual settings
# live under nested $obj.kioskProfiles[0].appConfiguration /
# .userAccountsConfiguration with discriminated @odata.type subtypes
# (windowsKioskSingleWin32App, windowsKioskSingleUWPApp, windowsKioskMultipleApps,
# windowsKioskAutologon, windowsKioskAzureADGroup/User, etc.). A handler with
# explicit subtype dispatch is required.
#
# Preserves an old-engine quirk: when userAccountsConfiguration is an array of
# mixed AAD User + AAD Group entries, PS evaluates the switch on the implicit
# array of @odata.types and `-eq 'kioskAADUserAndGroup'` on the resulting array
# returns the matching items (truthy in if-test). The "User logon type" row
# then shows empty because `"SettingDetails.$($logonTypeLngId)"` interpolates
# the array as space-joined.

class WindowsKioskDocHandler : DocumentationHandlerBase {
    WindowsKioskDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.windowsKioskConfiguration')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        # ---- Basic info ----
        Add-BasicDefaultValues $PolicyObject
        Add-BasicAdditionalValues $PolicyObject
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'Category.kioskConfigurationV2') '@odata.type'
        # Old engine emits a Platform row from $obj.platform. The raw payload
        # doesn't carry one for this type, so the lookup resolves to empty —
        # golden fixtures still contain the empty row, so emit it for parity.
        Add-BasicPropertyValue (Get-LanguageString 'Inputs.platformLabel') (Get-LanguageString "Platform.$($obj.platform)") 'platform'

        # ---- Settings ----
        $category = Get-LanguageString 'Category.kiosk'

        $appConfig = $obj.kioskProfiles[0].appConfiguration
        $userConfig = $obj.kioskProfiles[0].userAccountsConfiguration

        # kioskMode dispatch
        if ($appConfig.'@odata.type' -eq '#microsoft.graph.windowsKioskSingleWin32App' -or
            $appConfig.'@odata.type' -eq '#microsoft.graph.windowsKioskSingleUWPApp') {
            $kioskModeType = 'single'
            $kioskMode     = Get-LanguageString 'SettingDetails.kioskSelectionSingleMode'
        }
        else {
            $kioskModeType = 'multi'
            $kioskMode     = Get-LanguageString 'SettingDetails.kioskSelectionMultiMode'
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name      = Get-LanguageString 'SettingDetails.kioskSelectionName'
            Value     = $kioskMode
            EntityKey = 'kioskMode'
            Category  = $category
            SubCategory = $null
        })

        # User logon type
        $logonTypeLngId = switch ($userConfig.'@odata.type') {
            '#microsoft.graph.windowsKioskAutologon'    { 'kioskUserLogonTypeAutologon' }
            '#microsoft.graph.windowsKioskAzureADUser'  { 'kioskAADUserAndGroup' }
            '#microsoft.graph.windowsKioskAzureADGroup' { 'kioskAADUserAndGroup' }
            '#microsoft.graph.windowsKioskLocalUser'    { 'kioskAppTypeStore' }
            '#microsoft.graph.windowsKioskVisitor'      { 'kioskVisitor' }
        }
        $logonType = if ($logonTypeLngId) {
            Get-LanguageString "SettingDetails.$logonTypeLngId"
        } else {
            Write-Log "Unknown kiosk user logon type. $($userConfig.'@odata.type')" 2
            $null
        }

        Add-CustomSettingObject ([PSCustomObject]@{
            Name      = Get-LanguageString 'SettingDetails.kioskSelectionUsers'
            Value     = $logonType
            EntityKey = 'userAccountsConfigurationType'
            Category  = $category
            SubCategory = $null
        })

        # User logon name(s)
        if ($logonTypeLngId -eq 'kioskAADUserAndGroup') {
            $aadUser  = Get-LanguageString 'SettingDetails.kioskAADUser'
            $aadGroup = Get-LanguageString 'SettingDetails.kioskAADGroup'
            $users = @()
            foreach ($u in $userConfig) {
                $sep = $Context.PropertySeparator
                if ($u.'@odata.type' -eq '#microsoft.graph.windowsKioskAzureADUser') {
                    $users += "$($u.userPrincipalName)$sep$aadUser"
                }
                else {
                    $users += "$($u.displayName)$sep$aadGroup"
                }
            }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name      = Get-LanguageString 'SettingDetails.kioskUserAccountName'
                Value     = $users -join $Context.ObjectSeparator
                EntityKey = 'userAccounts'
                Category  = $category
                SubCategory = $null
            })
        }
        elseif ($userConfig.'@odata.type' -eq '#microsoft.graph.windowsKioskLocalUser') {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name      = Get-LanguageString 'SettingDetails.kioskUserAccountName'
                Value     = $userConfig.userName
                EntityKey = 'userName'
                Category  = $category
                SubCategory = $null
            })
        }

        # Single-app: detect underlying app type and emit type-specific rows
        if ($kioskModeType -eq 'single') {
            $uwpAppType = $null
            $appType    = $null
            if ($appConfig.'@odata.type' -eq '#microsoft.graph.windowsKioskSingleWin32App') {
                $uwpAppType = 'win32App'
                $appType    = Get-LanguageString 'SettingDetails.selectWin32AppForEdge86'
            }
            elseif ($appConfig.'@odata.type' -eq '#microsoft.graph.windowsKioskSingleUWPApp') {
                if ($appConfig.uwpApp.appUserModelId -like 'Microsoft.MicrosoftEdge*') {
                    $uwpAppType = 'edge'
                    $appType    = Get-LanguageString 'SettingDetails.selectMicrosoftEdgeApp'
                }
                elseif ($appConfig.uwpApp.appUserModelId -like 'Microsoft.KioskBrowser*') {
                    $uwpAppType = 'kioskBrowser'
                    $appType    = Get-LanguageString 'SettingDetails.selectKioskBrowserApp'
                }
                else {
                    $uwpAppType = 'storeApp'
                    $appType    = Get-LanguageString 'SettingDetails.selectStoreApp'
                }
            }

            Add-CustomSettingObject ([PSCustomObject]@{
                Name      = Get-LanguageString 'SettingDetails.kioskApplicationType'
                Value     = $appType
                EntityKey = 'kioskApplicationType'
                Category  = $category
                SubCategory = $null
            })

            $edgeKioskModeType = if ($appConfig.win32App.edgeKioskType -eq 'publicBrowsing') {
                Get-LanguageString 'SettingDetails.edgeKioskModeTypePublicBrowsingInPrivate'
            } else {
                Get-LanguageString 'SettingDetails.edgeKioskModeTypeDigitalSignage'
            }

            if ($uwpAppType -eq 'win32App') {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.win32EdgeKioskUrl'
                    Value     = $appConfig.win32App.edgeKiosk
                    EntityKey = 'edgeKiosk'
                    Category  = $category
                    SubCategory = $null
                })
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.edgeKioskModeType'
                    Value     = $edgeKioskModeType
                    EntityKey = 'edgeKioskType'
                    Category  = $category
                    SubCategory = $null
                })
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.edgeKioskResetAfterIdleTimeInMinutesName'
                    Value     = $appConfig.win32App.edgeKioskIdleTimeoutMinutes
                    EntityKey = 'edgeKioskIdleTimeoutMinutes'
                    Category  = $category
                    SubCategory = $null
                })
            }
            elseif ($uwpAppType -eq 'edge') {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.edgeKioskModeType'
                    Value     = $edgeKioskModeType
                    EntityKey = 'edgeKioskType'
                    Category  = $category
                    SubCategory = $null
                })
            }
            elseif ($uwpAppType -eq 'kioskBrowser') {
                $show = Get-LanguageString 'BooleanActions.show'
                $hide = Get-LanguageString 'BooleanActions.hide'

                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.win10KioskBrowserDefaultWebsiteName'
                    Value     = $obj.kioskBrowserDefaultUrl
                    EntityKey = 'kioskBrowserDefaultUrl'
                    Category  = $category
                    SubCategory = $null
                })
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.win10KioskBrowserHomeButtonName'
                    Value     = if ($obj.kioskBrowserEnableHomeButton) { $show } else { $hide }
                    EntityKey = 'kioskBrowserEnableHomeButton'
                    Category  = $category
                    SubCategory = $null
                })
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.win10KioskBrowserNavigationButtonName'
                    Value     = if ($obj.kioskBrowserEnableNavigationButtons) { $show } else { $hide }
                    EntityKey = 'kioskBrowserEnableNavigationButtons'
                    Category  = $category
                    SubCategory = $null
                })
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.win10KioskBrowserEnableEndSessionButtonName'
                    Value     = if ($obj.kioskBrowserEnableEndSessionButton) { $show } else { $hide }
                    EntityKey = 'kioskBrowserEnableEndSessionButton'
                    Category  = $category
                    SubCategory = $null
                })
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.edgeKioskResetAfterIdleTimeInMinutesName'
                    Value     = $obj.kioskBrowserRestartOnIdleTimeInMinutes
                    EntityKey = 'kioskBrowserRestartOnIdleTimeInMinutes'
                    Category  = $category
                    SubCategory = $null
                })
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.win10BlockedWebsitesName'
                    Value     = $obj.kioskBrowserBlockedURLs -join $Context.ObjectSeparator
                    EntityKey = 'kioskBrowserBlockedURLs'
                    Category  = $category
                    SubCategory = $null
                })
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.win10AllowedWebsitesName'
                    Value     = $obj.kioskBrowserBlockedUrlExceptions -join $Context.ObjectSeparator
                    EntityKey = 'kioskBrowserBlockedUrlExceptions'
                    Category  = $category
                    SubCategory = $null
                })
            }
            elseif ($uwpAppType -eq 'storeApp') {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.kioskModeAppStoreUrlOrManagedAppIdName'
                    Value     = $appConfig.uwpApp.name
                    EntityKey = 'edgeKioskType'
                    Category  = $category
                    SubCategory = $null
                })
            }
        }

        # Multi-app: app table + start-layout / taskbar / downloads rows
        if ($kioskModeType -eq 'multi') {
            $apps = @()
            foreach ($app in $appConfig.apps) {
                $kioskTypeLngId = switch ($app.appType) {
                    'aumId'   { 'kioskAppTypeAUMID' }
                    'desktop' { 'kioskAppTypeDesktop' }
                    'store'   { 'kioskAppTypeStore' }
                    default   { 'kioskAppTypeUnknown' }
                }
                $kioskTileLngId = switch ($app.startLayoutTileSize) {
                    'medium' { 'kioskTileMedium' }
                    'small'  { 'kioskTileSmall' }
                    'wide'   { 'kioskTileWide' }
                    'large'  { 'kioskTileLarge' }
                }
                $sep = $Context.PropertySeparator
                $autoLaunchStr = if ($app.autoLaunch -eq $true) { Get-LanguageString 'SettingDetails.yes' } else { Get-LanguageString 'SettingDetails.no' }
                $apps += '{0}{1}{2}{3}{4}{5}{6}' -f $app.Name, $sep, (Get-LanguageString "SettingDetails.$kioskTypeLngId"), $sep, $autoLaunchStr, $sep, (Get-LanguageString "SettingDetails.$kioskTileLngId")
            }

            if ($apps.Count -gt 0) {
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.kioskAppTableName'
                    Value     = $apps -join $Context.ObjectSeparator
                    EntityKey = 'kioskApps'
                    Category  = $category
                    SubCategory = $null
                })
            }

            $altLayout = if ($null -ne $appConfig.startMenuLayoutXml) { Get-LanguageString 'SettingDetails.yes' } else { Get-LanguageString 'SettingDetails.no' }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name      = Get-LanguageString 'SettingDetails.alternativeStartLayoutName'
                Value     = $altLayout
                EntityKey = 'alternativeStartLayout'
                Category  = $category
                SubCategory = $null
            })

            if ($null -ne $appConfig.startMenuLayoutXml) {
                $xmlStr = try {
                    [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($appConfig.startMenuLayoutXml))
                } catch { $appConfig.startMenuLayoutXml }
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.kioskStartMenuLayoutXmlName'
                    Value     = $xmlStr
                    EntityKey = 'startMenuLayoutXml'
                    Category  = $category
                    SubCategory = $null
                })
            }

            $taskBar = if ($appConfig.showTaskBar) { Get-LanguageString 'BooleanActions.show' } else { Get-LanguageString 'BooleanActions.hide' }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name      = Get-LanguageString 'SettingDetails.kioskShowTaskbarName'
                Value     = $taskBar
                EntityKey = 'showTaskBar'
                Category  = $category
                SubCategory = $null
            })

            $downloads = if ($appConfig.allowAccessToDownloadsFolder) { Get-LanguageString 'SettingDetails.yes' } else { Get-LanguageString 'SettingDetails.no' }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name      = Get-LanguageString 'SettingDetails.win10KioskAccessDownloadsFolderName'
                Value     = $downloads
                EntityKey = 'allowAccessToDownloadsFolder'
                Category  = $category
                SubCategory = $null
            })

            # disallowDesktopApps blocks classic Win32/desktop apps on the multi-app
            # kiosk. No scraped label exists for this toggle, so use an ASCII literal
            # (the Yes/No value is still localized).
            $disallowDesktopApps = if ($appConfig.disallowDesktopApps) { Get-LanguageString 'SettingDetails.yes' } else { Get-LanguageString 'SettingDetails.no' }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name      = 'Block desktop (Win32) apps'
                Value     = $disallowDesktopApps
                EntityKey = 'disallowDesktopApps'
                Category  = $category
                SubCategory = $null
            })
        }

        # Force-restart maintenance window
        $forceUpdateLng = if ($obj.windowsKioskForceUpdateSchedule) { 'BooleanActions.require' } else { 'BooleanActions.notConfigured' }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name      = Get-LanguageString 'SettingDetails.kioskForceRestart'
            Value     = Get-LanguageString $forceUpdateLng
            EntityKey = 'windowsKioskForceUpdateSchedule'
            Category  = $category
            SubCategory = $null
        })

        if ($obj.windowsKioskForceUpdateSchedule) {
            try {
                $startDateObj = if ($obj.windowsKioskForceUpdateSchedule.startDateTime -is [DateTime]) {
                    $tmp = $obj.windowsKioskForceUpdateSchedule.startDateTime
                    if ($tmp.Kind -eq [DateTimeKind]::Utc) { $tmp.ToLocalTime() } else { $tmp }
                } else {
                    Get-Date $obj.windowsKioskForceUpdateSchedule.startDateTime -ErrorAction Stop
                }

                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.kioskStartDateTime'
                    Value     = ($startDateObj.ToShortDateString() + $Context.ObjectSeparator + $startDateObj.ToShortTimeString())
                    EntityKey = 'startDateTime'
                    Category  = $category
                    SubCategory = $null
                })

                $recurrenceType = switch ($obj.windowsKioskForceUpdateSchedule.recurrence) {
                    'weekly'  { 'kioskWeekly' }
                    'monthly' { 'kioskMonthly' }
                    default   { 'kioskDaily' }
                }
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name      = Get-LanguageString 'SettingDetails.kioskRecurrence'
                    Value     = Get-LanguageString "SettingDetails.$recurrenceType"
                    EntityKey = 'recurrence'
                    Category  = $category
                    SubCategory = $null
                })

                if ($obj.windowsKioskForceUpdateSchedule.recurrence -eq 'weekly') {
                    Add-CustomSettingObject ([PSCustomObject]@{
                        Name      = Get-LanguageString 'SettingDetails.dayOfWeek'
                        Value     = Get-LanguageString "SettingDetails.$($obj.windowsKioskForceUpdateSchedule.dayofWeek)"
                        EntityKey = 'dayofWeek'
                        Category  = $category
                        SubCategory = $null
                    })
                }
                elseif ($obj.windowsKioskForceUpdateSchedule.recurrence -eq 'monthly') {
                    Add-CustomSettingObject ([PSCustomObject]@{
                        Name      = Get-LanguageString 'SettingDetails.dayOfMonth'
                        Value     = $obj.windowsKioskForceUpdateSchedule.dayofMonth
                        EntityKey = 'dayofMonth'
                        Category  = $category
                        SubCategory = $null
                    })
                }
            }
            catch { Write-Log "Failed to format kiosk force-update schedule: $($_.Exception.Message)" 2 }
        }
    }
}

[DocumentationRegistry]::RegisterHandler([WindowsKioskDocHandler]::new())
