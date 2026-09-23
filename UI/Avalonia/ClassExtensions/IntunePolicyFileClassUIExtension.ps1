# Avalonia parallel of UI/WPF/ClassExtensions/IntunePolicyFileClassUIExtension.ps1.
# Same shape as IntuneScriptClassUIExtension.ps1; differences:
#
#   * The "JSON property to edit" is dynamic — `_PolicyFileAttributes[0]` on
#     the type class points at the property holding base64 file bytes (e.g.
#     "fileContents" for custom attribute scripts). The TODO note about
#     "Support multiple attributes" carries over from the WPF original.
#   * Download writes raw bytes (no UTF-8 BOM stripping) to preserve binary
#     payloads. Filename defaults to JsonObject.FileName, falling back to
#     "<attribute>.json".
#
# Same scope-loss + R13 constraints as IntuneScriptClassUIExtension.ps1 —
# helpers captured up front and OnClick / OnSave SBs use .GetNewClosure().

##############################################
# Helpers
##############################################

function Invoke-DownloadPolicyFileAvalonia
{
    param($Policy)

    if (-not $Policy) { return }
    if ($Policy.IsFullObject -eq $false) { [void]$Policy.Get() }

    $attribute = $Policy.PolicyType._PolicyFileAttributes[0]
    if (-not $Policy.JsonObject.$attribute) { return }

    $fileName = ?? $Policy.JsonObject.FileName "$($attribute).json"
    Write-Log "Download PowerShell file '$fileName' from $($Policy.Name)"

    $picked = (Get-AvaloniaHost)::SaveFilePicker(
        $script:Window,
        'Save file',
        [string]$fileName,
        $null,
        'All files',
        '*.*')
    if (-not $picked) { return }

    [IO.File]::WriteAllBytes($picked, ([System.Convert]::FromBase64String($Policy.JsonObject.$attribute)))
}

function Invoke-EditPolicyFileAvalonia
{
    param($Policy)

    if (-not $Policy) { return }
    if ($Policy.IsFullObject -eq $false) { [void]$Policy.Get() }

    $attribute = $Policy.PolicyType._PolicyFileAttributes[0]  # TODO: Support multiple attributes
    if (-not $Policy.JsonObject.$attribute) { return }

    $policyFileText = [System.Text.Encoding]::UTF8.GetString(
        [System.Convert]::FromBase64String($Policy.JsonObject.$attribute))

    Show-AvaloniaScriptEditor `
        -Title       "Edit: $($Policy.JsonObject.displayName)" `
        -InitialText $policyFileText `
        -OnSave {
            param([string]$NewText)

            $pre    = [System.Text.Encoding]::UTF8.GetPreamble()
            $utfBom = [System.Text.Encoding]::UTF8.GetString($pre)
            if ($NewText.StartsWith($utfBom)) { $NewText = $NewText.Remove(0, $utfBom.Length) }
            $bytes       = [System.Text.Encoding]::UTF8.GetBytes($NewText)
            $encodedText = [Convert]::ToBase64String($bytes)
            $attr        = $Policy.PolicyType._PolicyFileAttributes[0]

            if ($Policy.JsonObject.$attr -eq $encodedText) { return }

            $confirm = $script:UIProvider.ShowMessageBox("Are you sure you want to update the file?`n`nObject:`n$($Policy.Name)", "Update file?", "YesNo", "Warning")
            if ($confirm -ne "Yes") { return }

            Write-Status "Update $($Policy.Name)"
            $cloned = $Policy.Clone()
            $cloned.JsonObject.$attr = $encodedText
            Remove-GraphPropertiesForImport $cloned $cloned.JsonObject
            foreach ($prop in $Policy.PolicyType.PropertiesToRemoveForUpdate) {
                Remove-Property $cloned.JsonObject $prop
            }
            Remove-Property $cloned.JsonObject "Assignments"
            Remove-Property $cloned.JsonObject "isAssigned"

            if ($Policy.Set($null, $cloned.JsonString, "PATCH", $null)) {
                Write-Log "File saved successfully"
            } else {
                Write-Log "Failed to save script" 3
            }
            Write-Status ""
        }.GetNewClosure()
}

##############################################
# Captures
##############################################

$sbAddCheckbox    = ${function:Add-AvaloniaExportPropertyCheckbox}
$sbAddDetailsBtn  = ${function:Add-AvaloniaDetailsButton}
$sbSetCache       = ${function:Set-CacheObject}
$sbWriteLogError  = ${function:Write-LogError}
$sbInvokeDownloadPolicyFile = ${function:Invoke-DownloadPolicyFileAvalonia}
$sbInvokeEditPolicyFile    = ${function:Invoke-EditPolicyFileAvalonia}

##############################################
# ScriptBlocks
##############################################

$SBAddPolicyFileUIExportExtension = {
    param($Form)

    if (-not $Form) { return }

    $chk = & $sbAddCheckbox -Form $Form `
        -CheckboxName 'chkExportPolicyFiles' `
        -LabelText    'Export file' `
        -ToolTip      'Export the files associated with selected profiles' `
        -InitialChecked $true
    if (-not $chk) { return }

    & $sbSetCache "ExportPolicyFiles" $chk.IsChecked

    $chk.add_IsCheckedChanged({
        param($s, $e)
        & $sbSetCache "ExportPolicyFiles" $s.IsChecked
    }.GetNewClosure())
}.GetNewClosure()

$SBAddPolicyFileUIDetailsExtension = {
    param($Form)

    if (-not $Form) { return }

    $onDownload = {
        param($s, $e)
        $row = $script:dgIntuneManagerObjects.SelectedItem
        $sel = if ($row) { $row.Source } else { $null }
        if (-not $sel) { return }
        try { & $sbInvokeDownloadPolicyFile $sel }
        catch { & $sbWriteLogError "Download policy file failed" $_.Exception }
    }.GetNewClosure()

    $onEdit = {
        param($s, $e)
        $row = $script:dgIntuneManagerObjects.SelectedItem
        $sel = if ($row) { $row.Source } else { $null }
        if (-not $sel) { return }
        try { & $sbInvokeEditPolicyFile $sel }
        catch { & $sbWriteLogError "Edit policy file failed" $_.Exception }
    }.GetNewClosure()

    & $sbAddDetailsBtn -Form $Form -Name 'btnDownload' -Content 'Download' `
        -InsertIndex 0 -OnClick $onDownload | Out-Null
    & $sbAddDetailsBtn -Form $Form -Name 'btnEdit' -Content 'Edit' `
        -InsertIndex 1 -OnClick $onEdit | Out-Null
}.GetNewClosure()

##############################################
# Init
##############################################

function Invoke-InitializePolicyFileUIExtensions
{
    $typeClasses = @()

    Get-SubClasses "IntunePolicyTypeBase" | ForEach-Object {
        if (Test-ClassIsAbstract $_) { return }
        try {
            $tmpClass = Get-SingletonObject $_.Name
            if ($null -ne $tmpClass.PolicyGroup -and $tmpClass._PolicyFileAttributes.Count -gt 0) {
                $typeClasses += $tmpClass
            }
        } catch {}
    }

    foreach ($typeClass in $typeClasses) {
        Add-ObjectMethod $typeClass "AddUIExportExtensions" $SBAddPolicyFileUIExportExtension
        Add-ObjectMethod $typeClass "AddUIDetailsExtension"  $SBAddPolicyFileUIDetailsExtension
    }
}

Invoke-InitializePolicyFileUIExtensions
