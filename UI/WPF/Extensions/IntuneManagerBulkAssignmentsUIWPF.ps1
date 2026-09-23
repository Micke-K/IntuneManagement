# Bulk Assignments form + assignment-row helpers.
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Show-GraphBulkAssignmentsForm
{
    $script:bulkAssignForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\BulkAssignmentsForm.xaml"))
    if(-not $script:bulkAssignForm) { return }

    $script:dgBulkAssignObjects = $script:bulkAssignForm.FindName("dgBulkAssignObjects")
    $script:dgBulkAssignList    = $script:bulkAssignForm.FindName("dgBulkAssignList")
    $script:txtBulkAssignStatus = $script:bulkAssignForm.FindName("txtBulkAssignStatus")

    $assignmentSettings = [IntuneManagerAssignmentSettings]::new()
    $script:bulkAssignForm.DataContext = $assignmentSettings

    # ── Eligible types: use the same support gate as Set-GraphBulkAssignments
    # so the UI does not offer groups/APIs that the command will skip.
    $script:bulkAssignEligibleTypes = @($script:IntuneTypes | Where-Object {
        Test-BulkAssignmentSupported $_
    })
    $eligibleTypeIds = @($script:bulkAssignEligibleTypes | ForEach-Object { $_.Id })
    $script:bulkAssignEligibleGroups = @($script:IntuneGroups | Where-Object {
        $_.Title -and ($_.PolicyTypes | Where-Object { $_.Id -in $eligibleTypeIds })
    })

    # ── Object list (mirrors Bulk Scope Tag DataGrid setup) ──────────────────
    $column = Get-GridCheckboxColumn "Selected"
    $script:dgBulkAssignObjects.Columns.Add($column)
    $script:bulkAssignSelectAllHeader = $column.Header
    $script:bulkAssignSelectAllHeader.IsChecked = $true
    $script:bulkAssignSelectAllHeader.add_Click({
        foreach($item in $script:dgBulkAssignObjects.ItemsSource) {
            $item.Selected = $this.IsChecked
        }
        $script:dgBulkAssignObjects.Items.Refresh()
    })

    $col = [System.Windows.Controls.DataGridTextColumn]::new()
    $col.Header     = "Object type"
    $col.IsReadOnly = $true
    $col.Binding    = [System.Windows.Data.Binding]::new("Title")
    $script:dgBulkAssignObjects.Columns.Add($col)

    $script:bulkAssignMode = "Group"
    Update-BulkAssignObjectList

    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "rbBulkAssignViewGroup", "add_Checked", {
        $script:bulkAssignMode = "Group"; Update-BulkAssignObjectList
        if($script:bulkAssignSelectAllHeader) { $script:bulkAssignSelectAllHeader.IsChecked = $true }
    })
    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "rbBulkAssignViewType", "add_Checked", {
        $script:bulkAssignMode = "Type"; Update-BulkAssignObjectList
        if($script:bulkAssignSelectAllHeader) { $script:bulkAssignSelectAllHeader.IsChecked = $true }
    })

    # ── Action radio → settings.Action sync ───────────────────────────────────
    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "rbBulkAssignAdd",     "add_Checked", { $script:bulkAssignForm.DataContext.Action = "Add" })
    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "rbBulkAssignReplace", "add_Checked", { $script:bulkAssignForm.DataContext.Action = "Replace" })
    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "rbBulkAssignRemove",  "add_Checked", { $script:bulkAssignForm.DataContext.Action = "Remove" })

    # ── Assignments list — backing ObservableCollection so adds/removes show
    # immediately. Each row carries display strings + the raw fields the
    # public command consumes.
    $script:colBulkAssignList = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    $script:dgBulkAssignList.ItemsSource = $script:colBulkAssignList

    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "btnBulkAssignAddGroup", "add_click", {
        Show-BulkAssignmentGroupPicker
    })
    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "btnBulkAssignAddAllDev", "add_click", {
        Add-BulkAssignmentRow -TargetType "allDevicesAssignmentTarget" -GroupName "All Devices"
    })
    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "btnBulkAssignAddAllUsers", "add_click", {
        Add-BulkAssignmentRow -TargetType "allLicensedUsersAssignmentTarget" -GroupName "All Users"
    })
    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "btnBulkAssignSettings", "add_click", {
        $sel = $script:dgBulkAssignList.SelectedItem
        if(-not $sel) {
            $script:UIProvider.ShowMessageBox("Select an assignment row first.", "Bulk Assignments", "OK", "Information") | Out-Null
            return
        }
        Show-BulkAssignmentSettingsDialog -Row $sel
    })
    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "btnBulkAssignRemoveSel", "add_click", {
        $sel = $script:dgBulkAssignList.SelectedItem
        if($sel) { [void]$script:colBulkAssignList.Remove($sel) }
    })
    # Double-click opens the settings dialog (less destructive than remove).
    $script:dgBulkAssignList.Add_MouseDoubleClick({
        $sel = $script:dgBulkAssignList.SelectedItem
        if($sel) { Show-BulkAssignmentSettingsDialog -Row $sel }
    })

    # ── Apply ─────────────────────────────────────────────────────────────────
    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "btnBulkAssignApply", "add_click", {
        $btnApply = $script:bulkAssignForm.FindName("btnBulkAssignApply")
        $btnClose = $script:bulkAssignForm.FindName("btnBulkAssignClose")

        $selectionIds = Get-BulkAssignSelectedObjectIds
        $unit = if($script:bulkAssignMode -eq "Type") { "policy type" } else { "object group" }
        if($selectionIds.Count -eq 0) {
            $script:UIProvider.ShowMessageBox("Select at least one $unit to update.", "Bulk Assignments", "OK", "Warning") | Out-Null
            return
        }

        $settings = $script:bulkAssignForm.DataContext
        $settings.Assignments = @($script:colBulkAssignList)

        if($settings.Assignments.Count -eq 0) {
            $script:UIProvider.ShowMessageBox("Add at least one assignment target before applying.", "Bulk Assignments", "OK", "Warning") | Out-Null
            return
        }

        # Mass-wipe guard: Replace with one row that isn't the actual desired
        # final state still wipes everything else. Warn before continuing.
        if($settings.Action -eq "Replace") {
            $proceed = $script:UIProvider.ShowMessageBox(
                "Replace will OVERWRITE every existing assignment on every matched policy with only the rows above ($($settings.Assignments.Count) target(s)).`n`nContinue?",
                "Confirm replace", "YesNo", "Warning")
            if($proceed -ne "Yes") { return }
        }

        $assignmentSummary = ($settings.Assignments | ForEach-Object {
            $parts = @($_.TargetTypeDisplay, $_.GroupName)
            if($_.FilterId) { $parts += "filter: $($_.FilterName) ($($_.FilterType))" }
            ($parts -join ' / ')
        }) -join "`n  "

        $filterText = if($settings.Filter) { $settings.Filter } else { "(none)" }
        $confirm = $script:UIProvider.ShowMessageBox(@"
About to update assignments on every policy that matches:

  Action          : $($settings.Action)
  Assignments     :
  $assignmentSummary
  Name filter     : $filterText
  $unit count     : $($selectionIds.Count)

Continue?
"@, "Confirm Bulk Assignment update", "YesNo", "Warning")
        if($confirm -ne "Yes") { return }

        if($btnApply) { $btnApply.IsEnabled = $false }
        if($btnClose) { $btnClose.IsEnabled = $false }
        # No pre-call Write-Status: Set-GraphBulkAssignments now sets its own two-line
        # status ("Bulk assignments — <Action>" + sub-step) as soon as it starts.

        $startParams = @{ AssignmentSettings = $settings }
        if($script:bulkAssignMode -eq "Type") { $startParams.PolicyType = $selectionIds }
        else                                  { $startParams.PolicyGroup = $selectionIds }

        try {
            $summary = Set-GraphBulkAssignments @startParams
            Write-Status $null

            $unsupportedNote = if($summary.UnsupportedTypes -and $summary.UnsupportedTypes.Count -gt 0) {
                "`nUnsupported types skipped: $($summary.UnsupportedTypes -join ', ')"
            } else { "" }

            $msg = ("Scanned: {0}`nMatched filter: {1}`nUpdated: {2}`nNo change: {3}`nFailed: {4}`nDuration: {5:hh\:mm\:ss}{6}" -f `
                $summary.PoliciesScanned, $summary.PoliciesMatched, $summary.PoliciesUpdated, $summary.PoliciesSkipped, $summary.PoliciesFailed, $summary.Duration, $unsupportedNote)
            $script:txtBulkAssignStatus.Text = ("Last run: updated {0}, no change {1}, failed {2}" -f $summary.PoliciesUpdated, $summary.PoliciesSkipped, $summary.PoliciesFailed)
            $unknownNote = Get-IntuneUnknownSelectorSummary $summary.UnknownSelectors
            if($unknownNote) { $msg += "`n`n$unknownNote" }
            $script:UIProvider.ShowMessageBox($msg, "Bulk Assignments", "OK", "Information") | Out-Null

            try { $script:dgIntuneManagerObjects.Items.Refresh() }
            catch { Write-LogDebug "Main grid refresh failed after bulk assignment run: $($_.Exception.Message)" }
        }
        catch {
            Write-LogError "Bulk Assignments run failed" $_.Exception
            $script:UIProvider.ShowMessageBox("Bulk Assignments run failed: $($_.Exception.Message)", "Bulk Assignments", "OK", "Error") | Out-Null
        }
        finally {
            if($btnApply) { $btnApply.IsEnabled = $true }
            if($btnClose) { $btnClose.IsEnabled = $true }
        }
    })

    $script:UIProvider.AddXamlEvent($script:bulkAssignForm, "btnBulkAssignClose", "add_click", {
        $script:bulkAssignForm = $null
        $script:bulkAssignSelectAllHeader = $null
        Show-ModalObject
    })

    $script:UIProvider.ShowModalForm("Bulk Assignments", $script:bulkAssignForm, $true)
}

# Rebuild the bulk-assign object DataGrid for the current view mode.
function Update-BulkAssignObjectList
{
    if(-not $script:dgBulkAssignObjects) { return }

    $script:bulkAssignObjects = @()
    if($script:bulkAssignMode -eq "Type") {
        $sorted = $script:bulkAssignEligibleTypes | Sort-Object Title
        foreach($pt in $sorted) {
            $script:bulkAssignObjects += New-Object PSObject -Property @{
                Title       = $pt.Title
                Selected    = $true
                ObjectGroup = $null
                ObjectType  = $pt
            }
        }
    }
    else {
        $sorted = $script:bulkAssignEligibleGroups | Sort-Object Title
        foreach($grp in $sorted) {
            $script:bulkAssignObjects += New-Object PSObject -Property @{
                Title       = $grp.Title
                Selected    = $true
                ObjectGroup = $grp
                ObjectType  = $null
            }
        }
    }
    $script:dgBulkAssignObjects.ItemsSource = $script:bulkAssignObjects
}

function Get-BulkAssignSelectedObjectIds
{
    if($script:bulkAssignMode -eq "Type") {
        return @($script:bulkAssignObjects |
            Where-Object { $_.Selected -eq $true -and $_.ObjectType -and $_.ObjectType.Id } |
            ForEach-Object { $_.ObjectType.Id })
    }
    return @($script:bulkAssignObjects |
        Where-Object { $_.Selected -eq $true -and $_.ObjectGroup -and $_.ObjectGroup.Id } |
        ForEach-Object { $_.ObjectGroup.Id })
}

# Append a row to the assignments DataGrid. Used by All Devices / All Users
# buttons and by the group picker dialog after the user picks OK.
function Add-BulkAssignmentRow
{
    param(
        [Parameter(Mandatory)][string]$TargetType,
        [string]$GroupId,
        [string]$GroupName,
        [string]$FilterId,
        [string]$FilterName,
        [string]$FilterType = "include",
        # installIntent for mobileAppAssignment. Ignored by the public command
        # for non-app types, so safe to set on every row.
        [string]$Intent = "required"
    )

    # Display strings — keep the data layer plain (TargetType is the raw
    # @odata.type fragment the public command uses; *Display columns are
    # what the DataGrid binds to).
    $typeDisplay = switch ($TargetType) {
        "groupAssignmentTarget"          { "Include group" }
        "exclusionGroupAssignmentTarget" { "Exclude group" }
        "allDevicesAssignmentTarget"     { "All devices" }
        "allLicensedUsersAssignmentTarget" { "All users" }
        default                          { $TargetType }
    }
    $filterDisplay = if($FilterId) { "$FilterName ($FilterType)" } else { "" }

    $row = [PSCustomObject]@{
        TargetType        = $TargetType
        TargetTypeDisplay = $typeDisplay
        GroupId           = $GroupId
        GroupName         = if($GroupName) { $GroupName } else { $GroupId }
        FilterId          = $FilterId
        FilterName        = $FilterName
        FilterType        = $FilterType
        FilterDisplay     = $filterDisplay
        Intent            = $Intent
        # Per-platform app assignment settings, keyed by Graph settings type
        # name (e.g. "win32LobAppAssignmentSettings"). Absent key = no override
        # for that type → Graph defaults are used on POST. Populated via the
        # App settings... dialog (Show-BulkAssignmentSettingsDialog).
        Settings          = @{}
        SettingsDisplay   = "(none)"
    }

    # Dedupe — don't add the same target+intent twice. Two rows with the
    # same target but different intents are intentionally allowed (Required
    # vs Available is meaningful for apps).
    $existing = @($script:colBulkAssignList | Where-Object {
        $_.TargetType -eq $row.TargetType -and
        [string]$_.GroupId -eq [string]$row.GroupId -and
        [string]$_.FilterId -eq [string]$row.FilterId -and
        ([string]::IsNullOrEmpty($row.FilterId) -or $_.FilterType -eq $row.FilterType) -and
        [string]$_.Intent -eq [string]$row.Intent
    })
    if($existing.Count -gt 0) {
        $script:UIProvider.ShowMessageBox("That target is already in the list (with the same intent).", "Bulk Assignments", "OK", "Information") | Out-Null
        return
    }

    [void]$script:colBulkAssignList.Add($row)
}

# ─── Bulk Assignments — per-platform settings ────────────────────────────────
#
# Schema for each settings type lives in $script:_bulkAssignSettingsSchema —
# one entry per TabItem in BulkAssignmentsSettings.xaml. Drives both the load
# (Row.Settings → UI controls) and save (UI controls → Row.Settings) passes
# from one piece of code per direction.
#
# Field types:
#   bool        — CheckBox.IsChecked
#   string      — TextBox.Text (empty skipped)
#   int         — TextBox.Text → int (empty skipped, non-numeric skipped)
#   enum        — ComboBox.SelectedItem (string)
#   datetime    — TextBox.Text passed through as-is (Graph accepts ISO8601)
#   stringList  — TextBox.Text split on newline / semicolon / comma (empties skipped)
#
# Nested complex types (win32LobAppRestartSettings, etc.) are encoded under a
# Nested array on the tab descriptor; the saved hashtable gets a nested
# hashtable + @odata.type stamp.
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
           @{ Control="cbAndroidStore_AutoUpdate"; Key="autoUpdateMode";                Kind="enum" }
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
                @{ Control="txtWin32_GracePeriod"; Key="gracePeriodInMinutes";                       Kind="int" }
                @{ Control="txtWin32_Countdown";   Key="countdownDisplayBeforeRestartInMinutes";     Kind="int" }
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
                @{ Control="txtWinGet_GracePeriod"; Key="gracePeriodInMinutes";                  Kind="int" }
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

# Pull a field value out of a control and convert to the right Graph type.
# Returns $null when the field has no usable value (empty string, no
# selection, parse failure) — caller skips the key entirely so Graph keeps
# its default rather than receiving an empty/zero value.
function Get-BulkAssignmentSettingValue
{
    param($Form, $Field)

    $ctl = $Form.FindName($Field.Control)
    if(-not $ctl) { return $null }

    switch ($Field.Kind) {
        "bool"   { return [bool]$ctl.IsChecked }
        "string" {
            $v = [string]$ctl.Text
            if([string]::IsNullOrWhiteSpace($v)) { return $null }
            return $v
        }
        "int" {
            $v = [string]$ctl.Text
            if([string]::IsNullOrWhiteSpace($v)) { return $null }
            $parsed = 0
            if([int]::TryParse($v, [ref]$parsed)) { return $parsed }
            return $null
        }
        "enum" {
            $v = [string]$ctl.SelectedItem
            if([string]::IsNullOrWhiteSpace($v)) { return $null }
            return $v
        }
        "datetime" {
            $v = [string]$ctl.Text
            if([string]::IsNullOrWhiteSpace($v)) { return $null }
            return $v  # Graph accepts ISO8601 as a string in the JSON body
        }
        "stringList" {
            $v = [string]$ctl.Text
            if([string]::IsNullOrWhiteSpace($v)) { return $null }
            $items = @($v -split "[`r`n;,]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if($items.Count -eq 0) { return $null }
            return ,$items
        }
    }
    return $null
}

# Push a stored hashtable value back into the UI control. Inverse of
# Get-BulkAssignmentSettingValue. Tolerates missing keys (control left at
# its XAML-default state).
function Set-BulkAssignmentSettingValue
{
    param($Form, $Field, $Value)

    $ctl = $Form.FindName($Field.Control)
    if(-not $ctl) { return }

    switch ($Field.Kind) {
        "bool"       { if($null -ne $Value) { $ctl.IsChecked = [bool]$Value } }
        "string"     { if($null -ne $Value) { $ctl.Text = [string]$Value } }
        "int"        { if($null -ne $Value) { $ctl.Text = [string]$Value } }
        "enum"       { if($null -ne $Value) { $ctl.SelectedItem = [string]$Value } }
        "datetime"   { if($null -ne $Value) { $ctl.Text = [string]$Value } }
        "stringList" {
            if($null -ne $Value) {
                $ctl.Text = (@($Value) -join [Environment]::NewLine)
            }
        }
    }
}

function Show-BulkAssignmentSettingsDialog
{
    param([Parameter(Mandatory)]$Row)

    $form = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\BulkAssignmentsSettings.xaml"))
    if(-not $form) { return }
    $script:_bulkAssignSettingsForm = $form
    $script:_bulkAssignSettingsRow  = $Row

    # Populate enum combos once. Lists come from the Graph CSDL — keep them
    # in lock-step with the metadata when Microsoft adds new enum members.
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
    foreach($name in $enumLists.Keys) {
        $ctl = $form.FindName($name)
        if($ctl) { $ctl.ItemsSource = $enumLists[$name] }
    }

    # Custom load: Health Script tab uses a polymorphic runSchedule whose
    # @odata.type depends on ScheduleType. Stored under the synthetic
    # "deviceHealthScriptAssignment" key in Row.Settings (not a Graph type
    # name — just an internal marker the public command consumes).
    if($Row.Settings -and $Row.Settings.ContainsKey('deviceHealthScriptAssignment')) {
        $hs = $Row.Settings['deviceHealthScriptAssignment']
        if($hs) {
            $form.FindName('chkApplyHealthScript').IsChecked = $true
            if($hs.ContainsKey('runRemediationScript')) {
                $form.FindName('chkScript_RunRemediation').IsChecked = [bool]$hs['runRemediationScript']
            }
            if($hs.ContainsKey('scheduleType'))      { $form.FindName('cbScript_ScheduleType').SelectedItem = [string]$hs['scheduleType'] }
            if($hs.ContainsKey('interval'))          { $form.FindName('txtScript_Interval').Text           = [string]$hs['interval'] }
            if($hs.ContainsKey('time'))              { $form.FindName('txtScript_Time').Text               = [string]$hs['time'] }
            if($hs.ContainsKey('useUtc'))            { $form.FindName('chkScript_UseUtc').IsChecked         = [bool]$hs['useUtc'] }
            if($hs.ContainsKey('date'))              { $form.FindName('txtScript_Date').Text                = [string]$hs['date'] }
        }
    }

    # Load existing values from Row.Settings into the controls.
    foreach($tab in $script:_bulkAssignSettingsSchema) {
        $existing = $null
        if($Row.Settings -and $Row.Settings.ContainsKey($tab.Type)) {
            $existing = $Row.Settings[$tab.Type]
        }
        if(-not $existing) { continue }

        $applyCtl = $form.FindName($tab.Apply)
        if($applyCtl) { $applyCtl.IsChecked = $true }

        foreach($field in $tab.Fields) {
            if($existing.ContainsKey($field.Key)) {
                Set-BulkAssignmentSettingValue $form $field $existing[$field.Key]
            }
        }
        if($tab.Nested) {
            foreach($nested in $tab.Nested) {
                if(-not $existing.ContainsKey($nested.Key)) { continue }
                $nestedHash = $existing[$nested.Key]
                if(-not $nestedHash) { continue }
                foreach($field in $nested.Fields) {
                    if($nestedHash.ContainsKey($field.Key)) {
                        Set-BulkAssignmentSettingValue $form $field $nestedHash[$field.Key]
                    }
                }
            }
        }
    }

    $script:UIProvider.AddXamlEvent($form, "btnBulkAssignSettingsOK", "add_click", {
        $r = $script:_bulkAssignSettingsRow
        $f = $script:_bulkAssignSettingsForm
        $new = @{}

        foreach($tab in $script:_bulkAssignSettingsSchema) {
            $applyCtl = $f.FindName($tab.Apply)
            if(-not $applyCtl -or -not $applyCtl.IsChecked) { continue }

            $hash = @{}
            foreach($field in $tab.Fields) {
                $v = Get-BulkAssignmentSettingValue $f $field
                if($null -ne $v) { $hash[$field.Key] = $v }
            }
            if($tab.Nested) {
                foreach($nested in $tab.Nested) {
                    $nestedHash = @{}
                    foreach($field in $nested.Fields) {
                        $v = Get-BulkAssignmentSettingValue $f $field
                        if($null -ne $v) { $nestedHash[$field.Key] = $v }
                    }
                    if($nestedHash.Count -gt 0) {
                        $nestedHash["@odata.type"] = "#microsoft.graph.$($nested.Type)"
                        $hash[$nested.Key] = $nestedHash
                    }
                }
            }

            # Only include the tab if SOMETHING is set. Otherwise it's just
            # the apply-checkbox ticked with all defaults — no value to send.
            if($hash.Count -gt 0) { $new[$tab.Type] = $hash }
        }

        # Custom save: Health Script tab. Stored as a flat hashtable; the
        # public command picks the schedule @odata.type at POST time from
        # the scheduleType value here.
        $applyHs = $f.FindName('chkApplyHealthScript')
        if($applyHs -and $applyHs.IsChecked) {
            $hsHash = @{
                runRemediationScript = [bool]$f.FindName('chkScript_RunRemediation').IsChecked
            }
            $st = [string]$f.FindName('cbScript_ScheduleType').SelectedItem
            if($st) {
                $hsHash['scheduleType'] = $st
                $intervalText = [string]$f.FindName('txtScript_Interval').Text
                $parsedInterval = 0
                if(-not [string]::IsNullOrWhiteSpace($intervalText) -and [int]::TryParse($intervalText, [ref]$parsedInterval)) {
                    $hsHash['interval'] = $parsedInterval
                }
                if($st -in @('Daily','Once')) {
                    $timeText = [string]$f.FindName('txtScript_Time').Text
                    if(-not [string]::IsNullOrWhiteSpace($timeText)) { $hsHash['time'] = $timeText.Trim() }
                    $hsHash['useUtc'] = [bool]$f.FindName('chkScript_UseUtc').IsChecked
                }
                if($st -eq 'Once') {
                    $dateText = [string]$f.FindName('txtScript_Date').Text
                    if(-not [string]::IsNullOrWhiteSpace($dateText)) { $hsHash['date'] = $dateText.Trim() }
                }
            }
            $new['deviceHealthScriptAssignment'] = $hsHash
        }

        $r.Settings = $new
        $r.SettingsDisplay = if($new.Count -gt 0) { "$($new.Count) platform(s)" } else { "(none)" }
        try { $script:dgBulkAssignList.Items.Refresh() } catch { }

        $script:_bulkAssignSettingsForm = $null
        $script:_bulkAssignSettingsRow  = $null
        Close-TopModalObject
    })

    $script:UIProvider.AddXamlEvent($form, "btnBulkAssignSettingsCancel", "add_click", {
        $script:_bulkAssignSettingsForm = $null
        $script:_bulkAssignSettingsRow  = $null
        Close-TopModalObject
    })

    $script:UIProvider.ShowModalForm("App settings - $($Row.GroupName)", $form, $true)
}

# Group-picker dialog. Searches AAD groups via Graph (startsWith), lets the
# user pick include/exclude direction and optionally attach a deviceAnd
# AppManagementAssignmentFilter. Reuses the global modal scaffolding so the
# parent Bulk Assignments form stays put underneath.
function Show-BulkAssignmentGroupPicker
{
    $picker = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\BulkAssignmentsGroupPicker.xaml"))
    if(-not $picker) { return }

    $script:_bulkAssignGroupPicker = $picker

    $txtSearch = $picker.FindName("txtGroupSearch")
    $btnSearch = $picker.FindName("btnGroupSearch")
    $lstResults = $picker.FindName("lstGroupResults")
    $cbFilter   = $picker.FindName("cbGroupFilter")
    $cbIntent   = $picker.FindName("cbGroupIntent")

    if(-not $txtSearch -or -not $btnSearch -or -not $lstResults -or -not $cbFilter -or -not $cbIntent) {
        Write-Log "Bulk Assignments group picker XAML loaded, but required controls were not found" 3
        return
    }

    # Populate assignment filters once. Empty entry = no filter.
    $tokenId = Get-DefaultTokenId
    $filters = @([PSCustomObject]@{ Id = $null; Name = "(no filter)" })
    try {
        $resp = Invoke-MSGraphAPI -Url "deviceManagement/assignmentFilters" -TokenId $tokenId -AllPages
        if($resp -and $resp.value) {
            foreach($f in @($resp.value | Sort-Object displayName)) {
                $filters += [PSCustomObject]@{ Id = [string]$f.id; Name = [string]$f.displayName }
            }
        }
    }
    catch { Write-LogDebug "Bulk Assignments: failed to load assignment filters: $($_.Exception.Message)" }
    $cbFilter.ItemsSource = $filters
    $cbFilter.SelectedIndex = 0

    # installIntent values from Graph's mobileAppAssignment schema. Required
    # is the most common bulk action so make it the default.
    $intents = @(
        [PSCustomObject]@{ Name = "Required";                       Value = "required" }
        [PSCustomObject]@{ Name = "Available";                      Value = "available" }
        [PSCustomObject]@{ Name = "Uninstall";                      Value = "uninstall" }
        [PSCustomObject]@{ Name = "Available without enrollment";   Value = "availableWithoutEnrollment" }
        [PSCustomObject]@{ Name = "Not available (hidden)";         Value = "notAvailable" }
    )
    $cbIntent.ItemsSource    = $intents
    $cbIntent.SelectedValue  = "required"

    $searchAction = {
        $activePicker = $script:_bulkAssignGroupPicker
        $searchBox = if($activePicker) { $activePicker.FindName("txtGroupSearch") } else { $null }
        $resultsList = if($activePicker) { $activePicker.FindName("lstGroupResults") } else { $null }
        if(-not $searchBox -or -not $resultsList -or -not ($resultsList -is [System.Windows.Controls.ItemsControl])) {
            Write-Log "Bulk Assignments: group search controls were not available" 3
            return
        }

        $term = [string]$searchBox.Text
        if($term.Length -lt 1) {
            $resultsList.ItemsSource = $null
            return
        }
        $escaped = $term.Replace("'", "''")
        $url = "groups?`$top=25&`$select=id,displayName&`$filter=startswith(displayName,'$escaped')"
        try {
            $resp = Invoke-MSGraphAPI -Url $url -TokenId (Get-DefaultTokenId)
            if($resp -and $resp.value) {
                $resultsList.ItemsSource = @($resp.value | Sort-Object displayName)
            }
            else {
                $resultsList.ItemsSource = @()
            }
        }
        catch {
            Write-LogError "Bulk Assignments: group search failed" $_.Exception
            $resultsList.ItemsSource = @()
        }
    }

    $btnSearch.Add_Click($searchAction)
    $txtSearch.Add_KeyDown({
        param($src, $e)
        if($e.Key -eq "Return") { & $searchAction }
    })

    $script:UIProvider.AddXamlEvent($picker, "btnGroupPickerOK", "add_click", {
        $sel = $script:_bulkAssignGroupPicker.FindName("lstGroupResults").SelectedItem
        if(-not $sel) {
            $script:UIProvider.ShowMessageBox("Pick a group from the search results first.", "Bulk Assignments", "OK", "Warning") | Out-Null
            return
        }
        $isExclude = [bool]$script:_bulkAssignGroupPicker.FindName("rbGroupExclude").IsChecked
        $targetType = if($isExclude) { "exclusionGroupAssignmentTarget" } else { "groupAssignmentTarget" }

        $filterItem = $script:_bulkAssignGroupPicker.FindName("cbGroupFilter").SelectedItem
        $filterId   = if($filterItem) { [string]$filterItem.Id } else { $null }
        $filterName = if($filterItem -and $filterId) { [string]$filterItem.Name } else { $null }
        $filterType = "include"
        if($filterId -and [bool]$script:_bulkAssignGroupPicker.FindName("rbGroupFilterExclude").IsChecked) {
            $filterType = "exclude"
        }

        $intent = [string]$script:_bulkAssignGroupPicker.FindName("cbGroupIntent").SelectedValue
        if(-not $intent) { $intent = "required" }

        Add-BulkAssignmentRow `
            -TargetType $targetType `
            -GroupId    ([string]$sel.id) `
            -GroupName  ([string]$sel.displayName) `
            -FilterId   $filterId `
            -FilterName $filterName `
            -FilterType $filterType `
            -Intent     $intent

        $script:_bulkAssignGroupPicker = $null
        Close-TopModalObject
    })

    $script:UIProvider.AddXamlEvent($picker, "btnGroupPickerCancel", "add_click", {
        $script:_bulkAssignGroupPicker = $null
        Close-TopModalObject
    })

    $script:UIProvider.ShowModalForm("Pick a group", $picker, $true)
}

