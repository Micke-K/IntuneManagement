# Avalonia parallel of UI/WPF/ClassExtensions/IntuneScriptClassUIExtension.ps1.
# Same ScriptMethod names (AddUIExportExtensions / AddUIDetailsExtension) so
# the Add-IntuneManagerExportUIExtensions wrapper from
# Extensions/IntuneManagerExtensionHooks.ps1 invokes them transparently.
#
# Two cross-cutting concerns:
#
#   * [[scriptmethod-loses-module-scope]] — Add-Member ScriptMethod scriptblocks
#     execute outside the defining module's session state and can't see
#     module-scoped functions by name. Every helper used inside the SBs is
#     captured up front via `${function:X}` and invoked through `&`. The outer
#     SB calls .GetNewClosure() so the captures travel with it; nested OnClick
#     handlers do the same so their inner-scope copies survive too.
#
#   * The WPF original used [System.Windows.Forms.SaveFileDialog] (R13 forbids
#     WinForms in UI/Avalonia) and built [System.Windows.Controls.*] objects
#     directly. Both are routed through the Avalonia helpers in
#     IntuneManagerExtensionHooks.ps1 (Add-AvaloniaExportPropertyCheckbox,
#     Add-AvaloniaDetailsButton, Show-AvaloniaScriptEditor) which use
#     [Avalonia.Controls.*] + Host::SaveFilePicker.

##############################################
# Helpers (file-private functions invoked from inside the ScriptMethods).
# Defined first so the `${function:X}` captures below see live refs before
# the SBs build their closures.
##############################################

function Invoke-DownloadScriptAvalonia
{
    param($ScriptPolicy)

    if (-not $ScriptPolicy) { return }
    if ($ScriptPolicy.IsFullObject -eq $false) { [void]$ScriptPolicy.Get() }
    if (-not $ScriptPolicy.JsonObject.scriptContent) { return }

    Write-Status ""

    $picked = (Get-AvaloniaHost)::SaveFilePicker(
        $script:Window,
        'Save script',
        [string]$ScriptPolicy.JsonObject.FileName,
        $null,
        'All files',
        '*.*')
    if (-not $picked) { return }

    Save-IntuneScriptContent $ScriptPolicy $picked | Out-Null
}

function Invoke-EditScriptAvalonia
{
    param($ScriptPolicy)

    if (-not $ScriptPolicy) { return }
    if ($ScriptPolicy.IsFullObject -eq $false) { [void]$ScriptPolicy.Get() }
    if (-not $ScriptPolicy.JsonObject.scriptContent) { return }

    $scriptText = [System.Text.Encoding]::UTF8.GetString(
        [System.Convert]::FromBase64String($ScriptPolicy.JsonObject.scriptContent))

    Show-AvaloniaScriptEditor `
        -Title       "Edit: $($ScriptPolicy.JsonObject.displayName)" `
        -InitialText $scriptText `
        -OnSave {
            param([string]$NewText)

            $pre    = [System.Text.Encoding]::UTF8.GetPreamble()
            $utfBom = [System.Text.Encoding]::UTF8.GetString($pre)
            if ($NewText.StartsWith($utfBom)) { $NewText = $NewText.Remove(0, $utfBom.Length) }
            $bytes       = [System.Text.Encoding]::UTF8.GetBytes($NewText)
            $encodedText = [Convert]::ToBase64String($bytes)

            if ($ScriptPolicy.JsonObject.scriptContent -eq $encodedText) { return }

            $confirm = $script:UIProvider.ShowMessageBox("Are you sure you want to update the script?`n`nObject:`n$($ScriptPolicy.Name)", "Update script?", "YesNo", "Warning")
            if ($confirm -ne "Yes") { return }

            Write-Status "Update $($ScriptPolicy.Name)"
            $cloned = $ScriptPolicy.Clone()
            $cloned.JsonObject.scriptContent = $encodedText
            Remove-GraphPropertiesForImport $cloned $cloned.JsonObject
            foreach ($prop in $ScriptPolicy.PolicyType.PropertiesToRemoveForUpdate) {
                Remove-Property $cloned.JsonObject $prop
            }
            Remove-Property $cloned.JsonObject "Assignments"
            Remove-Property $cloned.JsonObject "isAssigned"

            if ($ScriptPolicy.Set($null, $cloned.JsonString, "PATCH", $null)) {
                Write-Log "Script saved successfully"
            } else {
                Write-Log "Failed to save script" 3
            }
            Write-Status ""
        }.GetNewClosure()
}

##############################################
# Captures (see [[scriptmethod-loses-module-scope]])
##############################################

$sbAddCheckbox    = ${function:Add-AvaloniaExportPropertyCheckbox}
$sbAddDetailsBtn  = ${function:Add-AvaloniaDetailsButton}
$sbSetCache       = ${function:Set-CacheObject}
$sbWriteLogError  = ${function:Write-LogError}
$sbInvokeDownloadScript = ${function:Invoke-DownloadScriptAvalonia}
$sbInvokeEditScript    = ${function:Invoke-EditScriptAvalonia}

##############################################
# ScriptBlocks (built AFTER captures are in scope so GetNewClosure picks them up)
##############################################

$SBAddScriptUIExportExtension = {
    param($Form)

    if (-not $Form) { return }

    $chk = & $sbAddCheckbox -Form $Form `
        -CheckboxName 'chkExportScript' `
        -LabelText    'Export script' `
        -ToolTip      'Export the script associated with selected profiles' `
        -InitialChecked $true
    if (-not $chk) { return }

    & $sbSetCache "ExportScripts" $chk.IsChecked

    $chk.add_IsCheckedChanged({
        param($s, $e)
        & $sbSetCache "ExportScripts" $s.IsChecked
    }.GetNewClosure())
}.GetNewClosure()

$SBAddScriptUIDetailsExtension = {
    param($Form)

    if (-not $Form) { return }

    $onDownload = {
        param($s, $e)
        $row = $script:dgIntuneManagerObjects.SelectedItem
        $sel = if ($row) { $row.Source } else { $null }
        if (-not $sel) { return }
        try { & $sbInvokeDownloadScript $sel }
        catch { & $sbWriteLogError "Download script failed" $_.Exception }
    }.GetNewClosure()

    $onEdit = {
        param($s, $e)
        $row = $script:dgIntuneManagerObjects.SelectedItem
        $sel = if ($row) { $row.Source } else { $null }
        if (-not $sel) { return }
        try { & $sbInvokeEditScript $sel }
        catch { & $sbWriteLogError "Edit script failed" $_.Exception }
    }.GetNewClosure()

    & $sbAddDetailsBtn -Form $Form -Name 'btnDownload' -Content 'Download' `
        -InsertIndex 0 -OnClick $onDownload | Out-Null
    & $sbAddDetailsBtn -Form $Form -Name 'btnEdit' -Content 'Edit' `
        -InsertIndex 1 -OnClick $onEdit | Out-Null
}.GetNewClosure()

##############################################
# Init
##############################################

function Invoke-InitializeScriptUIExtensions
{
    $typeClasses = @()

    Get-SubClasses "ScriptTypeBase" | ForEach-Object {
        try {
            $tmpClass = Get-SingletonObject $_.Name
            if ($null -ne $tmpClass.PolicyGroup) {
                $typeClasses += $tmpClass
            }
        } catch {}
    }

    foreach ($typeClass in $typeClasses) {
        Add-ObjectMethod $typeClass "AddUIExportExtensions" $SBAddScriptUIExportExtension
        Add-ObjectMethod $typeClass "AddUIDetailsExtension"  $SBAddScriptUIDetailsExtension
    }
}

Invoke-InitializeScriptUIExtensions
