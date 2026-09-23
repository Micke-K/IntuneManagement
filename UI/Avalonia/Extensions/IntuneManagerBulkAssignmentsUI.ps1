# Avalonia port of the Bulk Assignments dialog from
# UI/WPF/Extensions/IntuneManagerUI.ps1 (Show-GraphBulkAssignmentsForm + helpers).
# New file because Bulk Assignments is a self-contained subsystem and the WPF
# original sits in the giant IntuneManagerUI file we're trying not to grow
# further (architecture rule R9).
#
# Slice 4e: Bulk Assignments. Combines a Group/API DataGrid (Slice 4a-style)
# with an assignment-list DataGrid + row-level group picker sub-modal +
# 13-tab schema-driven settings sub-modal. Driver Set-GraphBulkAssignments
# lives in Public/, so the click handler just preflight-validates, prompts,
# and forwards to the cmdlet.
#
# Test-BulkAssignmentSupported lives in Internal/BulkAssignments.ps1
# (loaded in Avalonia mode), so no duplication is needed here.

function Update-BulkAssignObjectList
{
    if (-not $script:dgBulkAssignObjects) { return }

    $rows = [System.Collections.Generic.List[BulkAssignRowItem]]::new()

    if ($script:bulkAssignMode -eq "Type") {
        $sortedTypes = $script:bulkAssignEligibleTypes | Sort-Object Title
        foreach ($pt in $sortedTypes) {
            $rows.Add([BulkAssignRowItem]@{
                Title       = [string]$pt.Title
                Selected    = $true
                ObjectGroup = $null
                ObjectType  = $pt
            })
        }
    } else {
        $sortedGroups = $script:bulkAssignEligibleGroups | Sort-Object Title
        foreach ($grp in $sortedGroups) {
            $rows.Add([BulkAssignRowItem]@{
                Title       = [string]$grp.Title
                Selected    = $true
                ObjectGroup = $grp
                ObjectType  = $null
            })
        }
    }

    $script:bulkAssignObjects = @($rows)
    $script:dgBulkAssignObjects.ItemsSource = $script:bulkAssignObjects
}

function Get-BulkAssignSelectedObjectIds
{
    if ($null -eq $script:bulkAssignObjects) { return @() }
    if ($script:bulkAssignMode -eq "Type") {
        return @($script:bulkAssignObjects |
            Where-Object { $_.Selected -and $_.ObjectType -and $_.ObjectType.Id } |
            ForEach-Object { $_.ObjectType.Id })
    }
    return @($script:bulkAssignObjects |
        Where-Object { $_.Selected -and $_.ObjectGroup -and $_.ObjectGroup.Id } |
        ForEach-Object { $_.ObjectGroup.Id })
}

# Append a row to the assignments DataGrid. Used by All Devices / All Users
# buttons and by the group picker sub-modal after the user clicks Add.
function Add-BulkAssignmentRow
{
    param(
        [Parameter(Mandatory)][string]$TargetType,
        [string]$GroupId,
        [string]$GroupName,
        [string]$FilterId,
        [string]$FilterName,
        [string]$FilterType = "include",
        [string]$Intent = "required"
    )

    $ui = $script:UIProvider
    $typeDisplay = switch ($TargetType) {
        "groupAssignmentTarget"            { "Include group" }
        "exclusionGroupAssignmentTarget"   { "Exclude group" }
        "allDevicesAssignmentTarget"       { "All devices" }
        "allLicensedUsersAssignmentTarget" { "All users" }
        default                            { $TargetType }
    }
    $filterDisplay = if ($FilterId) { "$FilterName ($FilterType)" } else { "" }

    $row = [BulkAssignmentRowItem]@{
        TargetType        = $TargetType
        TargetTypeDisplay = $typeDisplay
        GroupId           = $GroupId
        GroupName         = if ($GroupName) { $GroupName } else { $GroupId }
        FilterId          = $FilterId
        FilterName        = $FilterName
        FilterType        = $FilterType
        FilterDisplay     = $filterDisplay
        Intent            = $Intent
        Settings          = @{}
        SettingsDisplay   = "(none)"
    }

    $existing = @($script:colBulkAssignList | Where-Object {
        $_.TargetType -eq $row.TargetType -and
        [string]$_.GroupId -eq [string]$row.GroupId -and
        [string]$_.FilterId -eq [string]$row.FilterId -and
        ([string]::IsNullOrEmpty($row.FilterId) -or $_.FilterType -eq $row.FilterType) -and
        [string]$_.Intent -eq [string]$row.Intent
    })
    if ($existing.Count -gt 0) {
        $ui.ShowMessageBox("That target is already in the list (with the same intent).", "Bulk Assignments", "OK", "Information") | Out-Null
        return
    }

    [void]$script:colBulkAssignList.Add($row)
}

# ─── Bulk Assignments — per-platform settings ────────────────────────────────
#
# Schema for each settings type — one entry per TabItem in
# BulkAssignmentsSettings.axaml. Drives both load (Row.Settings → controls)
# and save (controls → Row.Settings) passes from one piece of code per
# direction. Same shape as the WPF original.
$script:_bulkAssignSettingsSchema = @(
    @{ Apply = "chkApplyIosLob"; Type = "iosLobAppAssignmentSettings";
       Fields = @(
           @{ Control="chkIosLob_IsRemovable";    Key="isRemovable";              Kind="bool" }
           @{ Control="chkIosLob_PreventBackup";  Key="preventManagedAppBackup";  Kind="bool" }
           @{ Control="chkIosLob_Uninstall";      Key="uninstallOnDeviceRemoval"; Kind="bool" }
           @{ Control="txtIosLob_VpnId";          Key="vpnConfigurationId";       Kind="string" }
       )}
    @{ Apply = "chkApplyIosStore"; Type = "iosStoreAppAssignmentSettings";
       Fields = @(
           @{ Control="chkIosStore_IsRemovable";    Key="isRemovable";              Kind="bool" }
           @{ Control="chkIosStore_PreventBackup";  Key="preventManagedAppBackup";  Kind="bool" }
           @{ Control="chkIosStore_Uninstall";      Key="uninstallOnDeviceRemoval"; Kind="bool" }
           @{ Control="txtIosStore_VpnId";          Key="vpnConfigurationId";       Kind="string" }
       )}
    @{ Apply = "chkApplyIosVpp"; Type = "iosVppAppAssignmentSettings";
       Fields = @(
           @{ Control="chkIosVpp_DeviceLicensing";    Key="useDeviceLicensing";       Kind="bool" }
           @{ Control="chkIosVpp_IsRemovable";        Key="isRemovable";              Kind="bool" }
           @{ Control="chkIosVpp_PreventBackup";      Key="preventManagedAppBackup";  Kind="bool" }
           @{ Control="chkIosVpp_PreventAutoUpdate";  Key="preventAutoAppUpdate";     Kind="bool" }
           @{ Control="chkIosVpp_Uninstall";          Key="uninstallOnDeviceRemoval"; Kind="bool" }
           @{ Control="txtIosVpp_VpnId";              Key="vpnConfigurationId";       Kind="string" }
       )}
    @{ Apply = "chkApplyIosDdm"; Type = "iosDdmLobAppAssignmentSettings";
       Fields = @(
           @{ Control="txtIosDdm_Domains";         Key="associatedDomains";                      Kind="stringList" }
           @{ Control="chkIosDdm_DirectDownload";  Key="associatedDomainsDirectDownloadAllowed"; Kind="bool" }
           @{ Control="chkIosDdm_PreventBackup";   Key="preventManagedAppBackup";                Kind="bool" }
           @{ Control="chkIosDdm_TapToPay";        Key="tapToPayScreenLockEnabled";              Kind="bool" }
           @{ Control="txtIosDdm_VpnId";           Key="vpnConfigurationId";                     Kind="string" }
       )}
    @{ Apply = "chkApplyAndroidStore"; Type = "androidManagedStoreAppAssignmentSettings";
       Fields = @(
           @{ Control="cbAndroidStore_AutoUpdate"; Key="autoUpdateMode";                 Kind="enum" }
           @{ Control="txtAndroidStore_Tracks";    Key="androidManagedStoreAppTrackIds"; Kind="stringList" }
       )}
    @{ Apply = "chkApplyMacLob"; Type = "macOsLobAppAssignmentSettings";
       Fields = @(
           @{ Control="chkMacLob_Uninstall"; Key="uninstallOnDeviceRemoval"; Kind="bool" }
       )}
    @{ Apply = "chkApplyMacVpp"; Type = "macOsVppAppAssignmentSettings";
       Fields = @(
           @{ Control="chkMacVpp_DeviceLicensing";   Key="useDeviceLicensing";       Kind="bool" }
           @{ Control="chkMacVpp_PreventBackup";     Key="preventManagedAppBackup";  Kind="bool" }
           @{ Control="chkMacVpp_PreventAutoUpdate"; Key="preventAutoAppUpdate";     Kind="bool" }
           @{ Control="chkMacVpp_Uninstall";         Key="uninstallOnDeviceRemoval"; Kind="bool" }
       )}
    @{ Apply = "chkApplyMsStore"; Type = "microsoftStoreForBusinessAppAssignmentSettings";
       Fields = @(
           @{ Control="chkMsStore_DeviceContext"; Key="useDeviceContext"; Kind="bool" }
       )}
    @{ Apply = "chkApplyWinAppx"; Type = "windowsAppXAppAssignmentSettings";
       Fields = @(
           @{ Control="chkWinAppx_DeviceContext"; Key="useDeviceContext"; Kind="bool" }
       )}
    @{ Apply = "chkApplyWinUap"; Type = "windowsUniversalAppXAppAssignmentSettings";
       Fields = @(
           @{ Control="chkWinUap_DeviceContext"; Key="useDeviceContext"; Kind="bool" }
       )}
    @{ Apply = "chkApplyWin32"; Type = "win32LobAppAssignmentSettings";
       Fields = @(
           @{ Control="cbWin32_DeliveryOpt";    Key="deliveryOptimizationPriority"; Kind="enum" }
           @{ Control="cbWin32_Notifications";  Key="notifications";                Kind="enum" }
       )
       Nested = @(
           @{ Key="autoUpdateSettings"; Type="win32LobAppAutoUpdateSettings";
              Fields=@(
                @{ Control="cbWin32_Superseded"; Key="autoUpdateSupersededAppsState"; Kind="enum" }
              )}
           @{ Key="installTimeSettings"; Type="mobileAppInstallTimeSettings";
              Fields=@(
                @{ Control="chkWin32_UseLocalTime";   Key="useLocalTime";     Kind="bool" }
                @{ Control="txtWin32_StartTime";      Key="startDateTime";    Kind="datetime" }
                @{ Control="txtWin32_DeadlineTime";   Key="deadlineDateTime"; Kind="datetime" }
              )}
           @{ Key="restartSettings"; Type="win32LobAppRestartSettings";
              Fields=@(
                @{ Control="txtWin32_GracePeriod"; Key="gracePeriodInMinutes";                   Kind="int" }
                @{ Control="txtWin32_Countdown";   Key="countdownDisplayBeforeRestartInMinutes"; Kind="int" }
              )}
       )}
    @{ Apply = "chkApplyWinGet"; Type = "winGetAppAssignmentSettings";
       Fields = @(
           @{ Control="cbWinGet_Notifications"; Key="notifications"; Kind="enum" }
       )
       Nested = @(
           @{ Key="installTimeSettings"; Type="winGetAppInstallTimeSettings";
              Fields=@(
                @{ Control="chkWinGet_UseLocalTime"; Key="useLocalTime";     Kind="bool" }
                @{ Control="txtWinGet_DeadlineTime"; Key="deadlineDateTime"; Kind="datetime" }
              )}
           @{ Key="restartSettings"; Type="winGetAppRestartSettings";
              Fields=@(
                @{ Control="txtWinGet_GracePeriod"; Key="gracePeriodInMinutes";                   Kind="int" }
                @{ Control="txtWinGet_Countdown";   Key="countdownDisplayBeforeRestartInMinutes"; Kind="int" }
              )}
       )}
    @{ Apply = "chkApplyWinAutoUpdate"; Type = "windowsAutoUpdateCatalogAppAssignmentSettings";
       Fields = @(
           @{ Control="cbWinAutoUpdate_DeliveryOpt";   Key="deliveryOptimizationPriority"; Kind="enum" }
           @{ Control="cbWinAutoUpdate_Notifications"; Key="notificationType";             Kind="enum" }
       )
       Nested = @(
           @{ Key="installTimeSettings"; Type="windowsAutoUpdateCatalogAppInstallTimeSettings";
              Fields=@(
                @{ Control="chkWinAutoUpdate_UseLocalTime"; Key="useLocalTime";     Kind="bool" }
                @{ Control="txtWinAutoUpdate_StartTime";    Key="startDateTime";    Kind="datetime" }
                @{ Control="txtWinAutoUpdate_DeadlineTime"; Key="deadlineDateTime"; Kind="datetime" }
              )}
           @{ Key="restartSettings"; Type="windowsAutoUpdateCatalogAppRestartSettings";
              Fields=@(
                @{ Control="txtWinAutoUpdate_GracePeriod"; Key="gracePeriodInMinutes"; Kind="int" }
              )}
       )}
)

function Get-BulkAssignmentSettingValue
{
    param($Form, $Field)

    $hostType = Get-AvaloniaHost
    $ctl = $hostType::FindByName($Form, $Field.Control)
    if (-not $ctl) { return $null }

    switch ($Field.Kind) {
        "bool"   { return [bool]$ctl.IsChecked }
        "string" {
            $v = [string]$ctl.Text
            if ([string]::IsNullOrWhiteSpace($v)) { return $null }
            return $v
        }
        "int" {
            $v = [string]$ctl.Text
            if ([string]::IsNullOrWhiteSpace($v)) { return $null }
            $parsed = 0
            if ([int]::TryParse($v, [ref]$parsed)) { return $parsed }
            return $null
        }
        "enum" {
            $v = [string]$ctl.SelectedItem
            if ([string]::IsNullOrWhiteSpace($v)) { return $null }
            return $v
        }
        "datetime" {
            $v = [string]$ctl.Text
            if ([string]::IsNullOrWhiteSpace($v)) { return $null }
            return $v
        }
        "stringList" {
            $v = [string]$ctl.Text
            if ([string]::IsNullOrWhiteSpace($v)) { return $null }
            $items = @($v -split "[`r`n;,]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if ($items.Count -eq 0) { return $null }
            return ,$items
        }
    }
    return $null
}

function Set-BulkAssignmentSettingValue
{
    param($Form, $Field, $Value)

    $hostType = Get-AvaloniaHost
    $ctl = $hostType::FindByName($Form, $Field.Control)
    if (-not $ctl) { return }

    switch ($Field.Kind) {
        "bool"       { if ($null -ne $Value) { $ctl.IsChecked = [bool]$Value } }
        "string"     { if ($null -ne $Value) { $ctl.Text = [string]$Value } }
        "int"        { if ($null -ne $Value) { $ctl.Text = [string]$Value } }
        "enum"       { if ($null -ne $Value) { $ctl.SelectedItem = [string]$Value } }
        "datetime"   { if ($null -ne $Value) { $ctl.Text = [string]$Value } }
        "stringList" {
            if ($null -ne $Value) {
                $ctl.Text = (@($Value) -join [Environment]::NewLine)
            }
        }
    }
}

function Show-BulkAssignmentSettingsDialog
{
    param([Parameter(Mandatory)]$Row)

    $ui = $script:UIProvider
    $form = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/BulkAssignmentsSettings.axaml'))
    if (-not $form) { return }
    $script:_bulkAssignSettingsForm = $form
    $script:_bulkAssignSettingsRow  = $Row

    $hostType = Get-AvaloniaHost

    # Populate enum combos once. Lists come from the Graph CSDL.
    $enumLists = @{
        cbAndroidStore_AutoUpdate     = @("default","postponed","priority")
        cbWin32_DeliveryOpt           = @("notConfigured","foreground")
        cbWin32_Notifications         = @("showAll","showReboot","hideAll")
        cbWin32_Superseded            = @("notConfigured","enabled")
        cbWinGet_Notifications        = @("showAll","showReboot","hideAll")
        cbWinAutoUpdate_DeliveryOpt   = @("notConfigured","foreground")
        cbWinAutoUpdate_Notifications = @("showAll","showReboot","hideAll")
        cbScript_ScheduleType         = @("Hourly","Daily","Once")
    }
    foreach ($name in $enumLists.Keys) {
        $ctl = $hostType::FindByName($form, $name)
        if ($ctl) { $ctl.ItemsSource = $enumLists[$name] }
    }

    # Custom load: Health Script tab uses a polymorphic runSchedule whose
    # @odata.type depends on ScheduleType. Stored under the synthetic
    # "deviceHealthScriptAssignment" key in Row.Settings.
    if ($Row.Settings -and $Row.Settings.ContainsKey('deviceHealthScriptAssignment')) {
        $hs = $Row.Settings['deviceHealthScriptAssignment']
        if ($hs) {
            ($hostType::FindByName($form, 'chkApplyHealthScript')).IsChecked = $true
            if ($hs.ContainsKey('runRemediationScript')) {
                ($hostType::FindByName($form, 'chkScript_RunRemediation')).IsChecked = [bool]$hs['runRemediationScript']
            }
            if ($hs.ContainsKey('scheduleType')) { ($hostType::FindByName($form, 'cbScript_ScheduleType')).SelectedItem = [string]$hs['scheduleType'] }
            if ($hs.ContainsKey('interval'))     { ($hostType::FindByName($form, 'txtScript_Interval')).Text           = [string]$hs['interval'] }
            if ($hs.ContainsKey('time'))         { ($hostType::FindByName($form, 'txtScript_Time')).Text               = [string]$hs['time'] }
            if ($hs.ContainsKey('useUtc'))       { ($hostType::FindByName($form, 'chkScript_UseUtc')).IsChecked         = [bool]$hs['useUtc'] }
            if ($hs.ContainsKey('date'))         { ($hostType::FindByName($form, 'txtScript_Date')).Text                = [string]$hs['date'] }
        }
    }

    # Load existing values from Row.Settings into the controls.
    foreach ($tab in $script:_bulkAssignSettingsSchema) {
        $existing = $null
        if ($Row.Settings -and $Row.Settings.ContainsKey($tab.Type)) {
            $existing = $Row.Settings[$tab.Type]
        }
        if (-not $existing) { continue }

        $applyCtl = $hostType::FindByName($form, $tab.Apply)
        if ($applyCtl) { $applyCtl.IsChecked = $true }

        foreach ($field in $tab.Fields) {
            if ($existing.ContainsKey($field.Key)) {
                Set-BulkAssignmentSettingValue $form $field $existing[$field.Key]
            }
        }
        if ($tab.Nested) {
            foreach ($nested in $tab.Nested) {
                if (-not $existing.ContainsKey($nested.Key)) { continue }
                $nestedHash = $existing[$nested.Key]
                if (-not $nestedHash) { continue }
                foreach ($field in $nested.Fields) {
                    if ($nestedHash.ContainsKey($field.Key)) {
                        Set-BulkAssignmentSettingValue $form $field $nestedHash[$field.Key]
                    }
                }
            }
        }
    }

    $btnOK     = $hostType::FindByName($form, 'btnBulkAssignSettingsOK')
    $btnCancel = $hostType::FindByName($form, 'btnBulkAssignSettingsCancel')

    if ($btnOK) {
        $btnOK.add_Click({
            $r = $script:_bulkAssignSettingsRow
            $f = $script:_bulkAssignSettingsForm
            $h = Get-AvaloniaHost
            $new = @{}

            foreach ($tab in $script:_bulkAssignSettingsSchema) {
                $applyCtl = $h::FindByName($f, $tab.Apply)
                if (-not $applyCtl -or -not $applyCtl.IsChecked) { continue }

                $hash = @{}
                foreach ($field in $tab.Fields) {
                    $v = Get-BulkAssignmentSettingValue $f $field
                    if ($null -ne $v) { $hash[$field.Key] = $v }
                }
                if ($tab.Nested) {
                    foreach ($nested in $tab.Nested) {
                        $nestedHash = @{}
                        foreach ($field in $nested.Fields) {
                            $v = Get-BulkAssignmentSettingValue $f $field
                            if ($null -ne $v) { $nestedHash[$field.Key] = $v }
                        }
                        if ($nestedHash.Count -gt 0) {
                            $nestedHash["@odata.type"] = "#microsoft.graph.$($nested.Type)"
                            $hash[$nested.Key] = $nestedHash
                        }
                    }
                }

                if ($hash.Count -gt 0) { $new[$tab.Type] = $hash }
            }

            # Custom save: Health Script tab. Stored as a flat hashtable; the
            # public command picks the schedule @odata.type at POST time from
            # the scheduleType value here.
            $applyHs = $h::FindByName($f, 'chkApplyHealthScript')
            if ($applyHs -and $applyHs.IsChecked) {
                $hsHash = @{
                    runRemediationScript = [bool]($h::FindByName($f, 'chkScript_RunRemediation')).IsChecked
                }
                $st = [string]($h::FindByName($f, 'cbScript_ScheduleType')).SelectedItem
                if ($st) {
                    $hsHash['scheduleType'] = $st
                    $intervalText = [string]($h::FindByName($f, 'txtScript_Interval')).Text
                    $parsedInterval = 0
                    if (-not [string]::IsNullOrWhiteSpace($intervalText) -and [int]::TryParse($intervalText, [ref]$parsedInterval)) {
                        $hsHash['interval'] = $parsedInterval
                    }
                    if ($st -in @('Daily','Once')) {
                        $timeText = [string]($h::FindByName($f, 'txtScript_Time')).Text
                        if (-not [string]::IsNullOrWhiteSpace($timeText)) { $hsHash['time'] = $timeText.Trim() }
                        $hsHash['useUtc'] = [bool]($h::FindByName($f, 'chkScript_UseUtc')).IsChecked
                    }
                    if ($st -eq 'Once') {
                        $dateText = [string]($h::FindByName($f, 'txtScript_Date')).Text
                        if (-not [string]::IsNullOrWhiteSpace($dateText)) { $hsHash['date'] = $dateText.Trim() }
                    }
                }
                $new['deviceHealthScriptAssignment'] = $hsHash
            }

            $r.Settings = $new
            $r.SettingsDisplay = if ($new.Count -gt 0) { "$($new.Count) platform(s)" } else { "(none)" }

            # Avalonia DataGrid has no Items.Refresh; rebind to surface the
            # SettingsDisplay change (BulkAssignmentRowItem has no INPC).
            try {
                if ($script:dgBulkAssignList) {
                    $current = $script:dgBulkAssignList.ItemsSource
                    $script:dgBulkAssignList.ItemsSource = $null
                    $script:dgBulkAssignList.ItemsSource = $current
                }
            } catch { }

            $script:_bulkAssignSettingsForm = $null
            $script:_bulkAssignSettingsRow  = $null
            Close-TopModalObject
        })
    }

    if ($btnCancel) {
        $btnCancel.add_Click({
            $script:_bulkAssignSettingsForm = $null
            $script:_bulkAssignSettingsRow  = $null
            Close-TopModalObject
        })
    }

    $ui.ShowModalForm("App settings - $($Row.GroupName)", $form, $true)
}

# Group search for the picker. Reads module-scope state so it is callable from
# event handlers (no captured locals / GetNewClosure — see [[avalonia-closure-dynamic-module]]).
function Invoke-BulkAssignGroupSearch
{
    $st = $script:_bulkAssignPickerState
    if (-not $st) { return }
    try {
        $term = [string]$st.TxtSearch.Text
        if ($term.Length -lt 1) {
            $st.LstResults.ItemsSource = $null
            $script:_bulkAssignGroupResultsMap = @{}
            return
        }
        $escaped = $term.Replace("'", "''")
        $url = "groups?`$top=25&`$select=id,displayName&`$filter=startswith(displayName,'$escaped')"
        $resp = Invoke-MSGraphAPI -Url $url -TokenId (Get-DefaultTokenId)
        $display = [System.Collections.Generic.List[string]]::new()
        $map = @{}
        if ($resp -and $resp.value) {
            foreach ($g in @($resp.value | Sort-Object displayName)) {
                $name = [string]$g.displayName
                $key = $name
                $i = 2
                while ($map.ContainsKey($key)) { $key = "$name ($i)"; $i++ }
                [void]$display.Add($key)
                $map[$key] = [PSCustomObject]@{ Id = [string]$g.id; Name = $name }
            }
        }
        $st.LstResults.ItemsSource = @($display)
        $script:_bulkAssignGroupResultsMap = $map
    } catch {
        Write-LogError "Bulk Assignments: group search failed" $_.Exception
        $st.LstResults.ItemsSource = @()
        $script:_bulkAssignGroupResultsMap = @{}
    }
}

# Group-picker sub-modal. Searches AAD groups via Graph (startsWith), lets
# the user pick include/exclude direction and optionally attach a
# deviceAndAppManagementAssignmentFilter. Stacks over the Bulk Assignments
# form via the modal-container Grid.
function Show-BulkAssignmentGroupPicker
{
    $ui = $script:UIProvider
    $picker = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/BulkAssignmentsGroupPicker.axaml'))
    if (-not $picker) { return }

    $script:_bulkAssignGroupPicker = $picker
    $hostType = Get-AvaloniaHost

    $txtSearch  = $hostType::FindByName($picker, 'txtGroupSearch')
    $btnSearch  = $hostType::FindByName($picker, 'btnGroupSearch')
    $lstResults = $hostType::FindByName($picker, 'lstGroupResults')
    $cbFilter   = $hostType::FindByName($picker, 'cbGroupFilter')
    $cbIntent   = $hostType::FindByName($picker, 'cbGroupIntent')
    $btnOK      = $hostType::FindByName($picker, 'btnGroupPickerOK')
    $btnCancel  = $hostType::FindByName($picker, 'btnGroupPickerCancel')

    if (-not $txtSearch -or -not $btnSearch -or -not $lstResults -or -not $cbFilter -or -not $cbIntent) {
        Write-Log "Bulk Assignments group picker XAML loaded, but required controls were not found" 3
        return
    }

    # Avalonia ListBox has no DisplayMemberPath; we project results to plain
    # strings for display, but keep a parallel id map keyed by displayName so
    # OK can resolve the selection back to its id. Search results may have
    # duplicate displayNames; suffix duplicates with " (id)" for uniqueness.
    $script:_bulkAssignGroupResultsMap = @{}

    # Populate assignment filters. Avalonia ComboBox can't bind PSCustomObject
    # by DisplayMemberPath either — use parallel string list + id map.
    $tokenId = Get-DefaultTokenId
    $filterDisplay = [System.Collections.Generic.List[string]]::new()
    $filterIdMap = @{}
    [void]$filterDisplay.Add("(no filter)")
    $filterIdMap["(no filter)"] = $null
    try {
        $resp = Invoke-MSGraphAPI -Url "deviceManagement/assignmentFilters" -TokenId $tokenId -AllPages
        if ($resp -and $resp.value) {
            foreach ($f in @($resp.value | Sort-Object displayName)) {
                $name = [string]$f.displayName
                $key = $name
                $i = 2
                while ($filterIdMap.ContainsKey($key)) { $key = "$name ($i)"; $i++ }
                [void]$filterDisplay.Add($key)
                $filterIdMap[$key] = [string]$f.id
            }
        }
    } catch {
        Write-LogDebug "Bulk Assignments: failed to load assignment filters: $($_.Exception.Message)"
    }
    $cbFilter.ItemsSource = @($filterDisplay)
    $cbFilter.SelectedIndex = 0
    $script:_bulkAssignFilterIdMap = $filterIdMap

    # installIntent values from Graph's mobileAppAssignment schema. Display
    # text in the ComboBox; map back to value on OK.
    $intentDisplayMap = [ordered]@{
        'Required'                     = 'required'
        'Available'                    = 'available'
        'Uninstall'                    = 'uninstall'
        'Available without enrollment' = 'availableWithoutEnrollment'
        'Not available (hidden)'       = 'notAvailable'
    }
    $cbIntent.ItemsSource = @($intentDisplayMap.Keys)
    $cbIntent.SelectedIndex = 0
    $script:_bulkAssignIntentMap = $intentDisplayMap

    # Module-scope state so handlers resolve controls at click time (no captured
    # locals / GetNewClosure / $script:UIProvider — see [[avalonia-closure-dynamic-module]]).
    $script:_bulkAssignPickerState = @{
        Picker     = $picker
        TxtSearch  = $txtSearch
        LstResults = $lstResults
        CbFilter   = $cbFilter
        CbIntent   = $cbIntent
    }

    $btnSearch.add_Click({ Invoke-BulkAssignGroupSearch })
    $txtSearch.add_KeyDown({
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Return) { Invoke-BulkAssignGroupSearch }
    })

    if ($btnOK) {
        $btnOK.add_Click({
            $st = $script:_bulkAssignPickerState
            if (-not $st) { return }
            try {
                $h = Get-AvaloniaHost
                $sel = $st.LstResults.SelectedItem
                if (-not $sel) {
                    Show-MessageBox "Pick a group from the search results first." "Bulk Assignments" "OK" "Warning" | Out-Null
                    return
                }
                $resolved = $script:_bulkAssignGroupResultsMap[[string]$sel]
                if (-not $resolved) { return }

                $rbExclude = $h::FindByName($st.Picker, 'rbGroupExclude')
                $isExclude = [bool]$rbExclude.IsChecked
                $targetType = if ($isExclude) { "exclusionGroupAssignmentTarget" } else { "groupAssignmentTarget" }

                $filterKey = [string]$st.CbFilter.SelectedItem
                $filterId  = $script:_bulkAssignFilterIdMap[$filterKey]
                $filterName = if ($filterId) { $filterKey } else { $null }
                $filterType = "include"
                $rbFilterExclude = $h::FindByName($st.Picker, 'rbGroupFilterExclude')
                if ($filterId -and [bool]$rbFilterExclude.IsChecked) {
                    $filterType = "exclude"
                }

                $intentKey = [string]$st.CbIntent.SelectedItem
                $intent = if ($intentKey -and $script:_bulkAssignIntentMap.Contains($intentKey)) {
                    $script:_bulkAssignIntentMap[$intentKey]
                } else { "required" }

                Add-BulkAssignmentRow `
                    -TargetType $targetType `
                    -GroupId    $resolved.Id `
                    -GroupName  $resolved.Name `
                    -FilterId   $filterId `
                    -FilterName $filterName `
                    -FilterType $filterType `
                    -Intent     $intent

                $script:_bulkAssignGroupPicker = $null
                $script:_bulkAssignGroupResultsMap = @{}
                $script:_bulkAssignPickerState = $null
                Close-TopModalObject
            } catch {
                Write-LogError "Bulk Assignments group picker OK failed" $_.Exception
            }
        })
    }

    if ($btnCancel) {
        $btnCancel.add_Click({
            $script:_bulkAssignGroupPicker = $null
            $script:_bulkAssignGroupResultsMap = @{}
            $script:_bulkAssignPickerState = $null
            Close-TopModalObject
        })
    }

    $ui.ShowModalForm("Pick a group", $picker, $true)
}

function Show-GraphBulkAssignmentsForm
{
    $ui = $script:UIProvider
    $form = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/BulkAssignmentsForm.axaml'))
    if (-not $form) { return }

    $hostType = Get-AvaloniaHost
    $script:bulkAssignForm        = $form
    $script:dgBulkAssignObjects   = $hostType::FindByName($form, 'dgBulkAssignObjects')
    $script:dgBulkAssignList      = $hostType::FindByName($form, 'dgBulkAssignList')
    $script:txtBulkAssignStatus   = $hostType::FindByName($form, 'txtBulkAssignStatus')

    $rbAdd      = $hostType::FindByName($form, 'rbBulkAssignAdd')
    $rbReplace  = $hostType::FindByName($form, 'rbBulkAssignReplace')
    $rbRemove   = $hostType::FindByName($form, 'rbBulkAssignRemove')
    $txtFilter  = $hostType::FindByName($form, 'txtBulkAssignNameFilter')
    $rbGroup    = $hostType::FindByName($form, 'rbBulkAssignViewGroup')
    $rbType     = $hostType::FindByName($form, 'rbBulkAssignViewType')
    $btnApply   = $hostType::FindByName($form, 'btnBulkAssignApply')
    $btnClose   = $hostType::FindByName($form, 'btnBulkAssignClose')

    $btnAddGroup    = $hostType::FindByName($form, 'btnBulkAssignAddGroup')
    $btnAddAllDev   = $hostType::FindByName($form, 'btnBulkAssignAddAllDev')
    $btnAddAllUsers = $hostType::FindByName($form, 'btnBulkAssignAddAllUsers')
    $btnSettings    = $hostType::FindByName($form, 'btnBulkAssignSettings')
    $btnRemoveSel   = $hostType::FindByName($form, 'btnBulkAssignRemoveSel')

    $assignmentSettings = [IntuneManagerAssignmentSettings]::new()

    # Eligible types: same gate as Set-GraphBulkAssignments so the UI does
    # not offer groups/APIs the command will skip.
    $script:bulkAssignEligibleTypes = @($script:IntuneTypes | Where-Object {
        Test-BulkAssignmentSupported $_
    })
    $eligibleTypeIds = @($script:bulkAssignEligibleTypes | ForEach-Object { $_.Id })
    $script:bulkAssignEligibleGroups = @($script:IntuneGroups | Where-Object {
        $_.Title -and ($_.PolicyTypes | Where-Object { $_.Id -in $eligibleTypeIds })
    })

    $script:bulkAssignMode = "Group"
    Update-BulkAssignObjectList

    $headerCb = Initialize-AvaloniaGridSelectAllHeader -Grid $script:dgBulkAssignObjects -BindingProperty 'Selected' -InitiallyChecked $true

    # Module-scope state so event handlers resolve it at click time (no captured
    # locals / GetNewClosure / $script:UIProvider — see [[avalonia-closure-dynamic-module]]).
    $script:_bulkAssignState = @{
        HeaderCb  = $headerCb
        Settings  = $assignmentSettings
        TxtFilter = $txtFilter
        BtnApply  = $btnApply
        BtnClose  = $btnClose
    }

    if ($rbGroup) {
        $rbGroup.add_IsCheckedChanged({
            param($s, $e)
            if (-not $s.IsChecked) { return }
            $script:bulkAssignMode = "Group"
            Update-BulkAssignObjectList
            $st = $script:_bulkAssignState
            if ($st -and $st.HeaderCb) { $st.HeaderCb.IsChecked = $true }
        })
    }
    if ($rbType) {
        $rbType.add_IsCheckedChanged({
            param($s, $e)
            if (-not $s.IsChecked) { return }
            $script:bulkAssignMode = "Type"
            Update-BulkAssignObjectList
            $st = $script:_bulkAssignState
            if ($st -and $st.HeaderCb) { $st.HeaderCb.IsChecked = $true }
        })
    }

    # Action radio → settings.Action
    if ($rbAdd) {
        $rbAdd.add_IsCheckedChanged({
            param($s, $e)
            if ($s.IsChecked -and $script:_bulkAssignState) { $script:_bulkAssignState.Settings.Action = "Add" }
        })
    }
    if ($rbReplace) {
        $rbReplace.add_IsCheckedChanged({
            param($s, $e)
            if ($s.IsChecked -and $script:_bulkAssignState) { $script:_bulkAssignState.Settings.Action = "Replace" }
        })
    }
    if ($rbRemove) {
        $rbRemove.add_IsCheckedChanged({
            param($s, $e)
            if ($s.IsChecked -and $script:_bulkAssignState) { $script:_bulkAssignState.Settings.Action = "Remove" }
        })
    }

    # Select-all behaviour moved into the DataGrid column header via
    # Initialize-AvaloniaGridSelectAllHeader (called earlier in this function).

    # Assignments list — backing ObservableCollection so adds/removes show
    # immediately. Each row is a [BulkAssignmentRowItem] (CLR-typed per
    # [[avalonia-binding-needs-clr-types]]).
    $script:colBulkAssignList = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    $script:dgBulkAssignList.ItemsSource = $script:colBulkAssignList

    if ($btnAddGroup) {
        $btnAddGroup.add_Click({ Show-BulkAssignmentGroupPicker })
    }
    if ($btnAddAllDev) {
        $btnAddAllDev.add_Click({
            Add-BulkAssignmentRow -TargetType "allDevicesAssignmentTarget" -GroupName "All Devices"
        })
    }
    if ($btnAddAllUsers) {
        $btnAddAllUsers.add_Click({
            Add-BulkAssignmentRow -TargetType "allLicensedUsersAssignmentTarget" -GroupName "All Users"
        })
    }
    if ($btnSettings) {
        $btnSettings.add_Click({
            $sel = $script:dgBulkAssignList.SelectedItem
            if (-not $sel) {
                Show-MessageBox "Select an assignment row first." "Bulk Assignments" "OK" "Information" | Out-Null
                return
            }
            Show-BulkAssignmentSettingsDialog -Row $sel
        })
    }
    if ($btnRemoveSel) {
        $btnRemoveSel.add_Click({
            $sel = $script:dgBulkAssignList.SelectedItem
            if ($sel) { [void]$script:colBulkAssignList.Remove($sel) }
        })
    }

    # Double-tap on a row opens the settings dialog (less destructive than remove).
    $script:dgBulkAssignList.add_DoubleTapped({
        $sel = $script:dgBulkAssignList.SelectedItem
        if ($sel) { Show-BulkAssignmentSettingsDialog -Row $sel }
    })

    if ($btnApply) {
        $btnApply.add_Click({
            $st = $script:_bulkAssignState
            if (-not $st) { return }
            try {
                $assignmentSettings = $st.Settings

                $selectionIds = Get-BulkAssignSelectedObjectIds
                $unit = if ($script:bulkAssignMode -eq "Type") { "policy type" } else { "object group" }
                if ($selectionIds.Count -eq 0) {
                    Show-MessageBox "Select at least one $unit to update." "Bulk Assignments" "OK" "Warning" | Out-Null
                    return
                }

                $assignmentSettings.Filter      = if ($st.TxtFilter) { [string]$st.TxtFilter.Text } else { '' }
                $assignmentSettings.Assignments = @($script:colBulkAssignList)

                if ($assignmentSettings.Assignments.Count -eq 0) {
                    Show-MessageBox "Add at least one assignment target before applying." "Bulk Assignments" "OK" "Warning" | Out-Null
                    return
                }

                # Mass-wipe guard for Replace: warn before overwriting
                # everything with potentially fewer rows than originally set.
                if ($assignmentSettings.Action -eq "Replace") {
                    $proceed = Show-MessageBox "Replace will OVERWRITE every existing assignment on every matched policy with only the rows above ($($assignmentSettings.Assignments.Count) target(s)).`n`nContinue?" "Confirm replace" "YesNo" "Warning"
                    if ($proceed -ne "Yes") { return }
                }

                $assignmentSummary = ($assignmentSettings.Assignments | ForEach-Object {
                    $parts = @($_.TargetTypeDisplay, $_.GroupName)
                    if ($_.FilterId) { $parts += "filter: $($_.FilterName) ($($_.FilterType))" }
                    ($parts -join ' / ')
                }) -join "`n  "

                $filterText = if ($assignmentSettings.Filter) { $assignmentSettings.Filter } else { "(none)" }
                $confirmMsg = @"
About to update assignments on every policy that matches:

  Action          : $($assignmentSettings.Action)
  Assignments     :
  $assignmentSummary
  Name filter     : $filterText
  $unit count     : $($selectionIds.Count)

Continue?
"@
                $confirm = Show-MessageBox $confirmMsg "Confirm Bulk Assignment update" "YesNo" "Warning"
                if ($confirm -ne "Yes") { return }

                if ($st.BtnApply) { $st.BtnApply.IsEnabled = $false }
                if ($st.BtnClose) { $st.BtnClose.IsEnabled = $false }

                try {
                    $startParams = @{ AssignmentSettings = $assignmentSettings }
                    if ($script:bulkAssignMode -eq "Type") { $startParams.PolicyType = $selectionIds }
                    else                                   { $startParams.PolicyGroup = $selectionIds }

                    $summary = Set-GraphBulkAssignments @startParams
                    Write-Status ""

                    $unsupportedNote = if ($summary.UnsupportedTypes -and $summary.UnsupportedTypes.Count -gt 0) {
                        "`nUnsupported types skipped: $($summary.UnsupportedTypes -join ', ')"
                    } else { "" }

                    $msg = ("Scanned: {0}`nMatched filter: {1}`nUpdated: {2}`nNo change: {3}`nFailed: {4}`nDuration: {5:hh\:mm\:ss}{6}" -f `
                        $summary.PoliciesScanned, $summary.PoliciesMatched, $summary.PoliciesUpdated, $summary.PoliciesSkipped, $summary.PoliciesFailed, $summary.Duration, $unsupportedNote)
                    if ($script:txtBulkAssignStatus) {
                        $script:txtBulkAssignStatus.Text = ("Last run: updated {0}, no change {1}, failed {2}" -f $summary.PoliciesUpdated, $summary.PoliciesSkipped, $summary.PoliciesFailed)
                    }
                    $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
                    if ($unknownNote) { $msg += "`n`n$unknownNote" }
                    Show-MessageBox $msg "Bulk Assignments" "OK" "Information" | Out-Null

                    if ($script:IntuneManagerSelectedObject) {
                        Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject | Out-Null
                    }
                } catch {
                    Write-LogError "Bulk Assignments run failed" $_.Exception
                    Show-MessageBox "Bulk Assignments run failed: $($_.Exception.Message)" "Bulk Assignments" "OK" "Error" | Out-Null
                } finally {
                    if ($st.BtnApply) { $st.BtnApply.IsEnabled = $true }
                    if ($st.BtnClose) { $st.BtnClose.IsEnabled = $true }
                }
            } catch {
                Write-LogError "Bulk Assignments Apply handler failed" $_.Exception
            }
        })
    }

    if ($btnClose) {
        $btnClose.add_Click({
            $script:bulkAssignForm      = $null
            $script:dgBulkAssignObjects = $null
            $script:dgBulkAssignList    = $null
            $script:bulkAssignObjects   = $null
            $script:colBulkAssignList   = $null
            $script:_bulkAssignState    = $null
            Show-ModalObject
        })
    }

    $form.add_KeyDown({
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            $script:bulkAssignForm      = $null
            $script:dgBulkAssignObjects = $null
            $script:dgBulkAssignList    = $null
            $script:bulkAssignObjects   = $null
            $script:colBulkAssignList   = $null
            $script:_bulkAssignState    = $null
            Show-ModalObject
            $e.Handled = $true
        }
    })

    $ui.ShowModalForm("Bulk Assignments", $form, $true)
}
