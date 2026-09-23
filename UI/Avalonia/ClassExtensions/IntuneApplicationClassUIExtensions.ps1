# Avalonia parallel of UI/WPF/ClassExtensions/IntuneApplicationClassUIExtensions.ps1.
# Attaches AddUIExportExtensions / AddUIDetailsExtension to ApplicationType
# only. Same scope-loss + R13 constraints documented in
# IntuneScriptClassUIExtension.ps1 — helpers captured up front, OpenFileDialog
# / SaveFileDialog routed through Host::OpenFilePicker / Host::SaveFilePicker.
#
# Two non-trivial differences from the script flow:
#
#   1. Export adds TWO checkboxes — "Export script" (shared with the script
#      path; chains through Add-AvaloniaExportPropertyCheckbox under the
#      `chkExportScript` name) and an Application-specific
#      "Export application file" (`chkExportAppContent`). The latter's
#      initial state comes from a persisted setting, not a hard-coded `$true`.
#
#   2. Download is a multi-step pipeline: SaveFilePicker for the encrypted
#      content, Find-AppEncryptionFile to locate the matching .json,
#      OpenFilePicker fallback when the json is missing, Start-DecryptFile to
#      produce the unencrypted .zip, then delete the .encrypted intermediate.

##############################################
# Helpers
##############################################

function Invoke-UploadAppFileAvalonia
{
    param($Row)

    if (-not $Row -or -not $Row.Source) { return }
    $appItem = $Row.Source
    $obj     = $appItem.Object

    if ($obj.publishingState -ne "notPublished") {
        $confirm = $script:UIProvider.ShowMessageBox("Are you sure you want to upload a new file for the app?`n`nApplication:`n$($appItem.Name)", "Update app file?", "YesNo", "Warning")
        if ($confirm -ne "Yes") { return }
    }

    $pkgPath = Get-SettingValue "IntuneAppPackagesFolder"
    $picked = (Get-AvaloniaHost)::OpenFilePicker(
        $script:Window,
        'Select Intune Win32 file',
        $pkgPath,
        'Intune Win32',
        '*.intunewin')
    if (-not $picked) { return }

    Write-Status "Import $($obj.displayName) file"
    Start-ApplicationImportFile $appItem $picked
    Write-Status ""
}

function Invoke-DownloadAppFileAvalonia
{
    param($Row)

    if (-not $Row -or -not $Row.Source) { return }
    $appItem = $Row.Source
    $obj     = $appItem.Object

    # Prefer the package fileName; fall back to the app display name so the save
    # dialog never defaults to a bare ".encrypted" (win32 list rows may not carry
    # a fileName).
    $baseName = if ($obj.FileName) { [string]$obj.FileName } elseif ($obj.displayName) { [string]$obj.displayName } else { 'app' }

    Write-Status "Download file"
    try {
        $pkgPath = Get-SettingValue "IntuneAppDownloadFolder" (Get-SettingValue "IntuneAppPackagesFolder")
        $picked = (Get-AvaloniaHost)::SaveFilePicker(
            $script:Window,
            'Save encrypted intunewin',
            ($baseName + ".encrypted"),
            'encrypted',
            'Encrypted intunewin',
            '*.encrypted')
        if (-not $picked) { return }

        $contentFileObj = Start-DownloadAppContent $appItem $picked

        if (-not [IO.File]::Exists($picked)) { return }

        $fullPath = Find-AppEncryptionFile $obj $contentFileObj $pkgPath
        if (-not [IO.File]::Exists($fullPath)) {
            $browse = $script:UIProvider.ShowMessageBox("Could not find decryption file for $($obj.displayName)`nApp Id: $($obj.id)`nContent version $($obj.committedContentVersion)`n`nDo you want to browse for the file?", "Encryption file not found", "YesNo", "Warning")
            if ($browse -eq "Yes") {
                $alt = (Get-AvaloniaHost)::OpenFilePicker(
                    $script:Window,
                    'Select decryption json',
                    $pkgPath,
                    'JSON files',
                    '*.json')
                if ($alt) { $fullPath = $alt }
            }
        }

        if ([IO.File]::Exists($fullPath)) {
            Write-Status "Decrypting file"
            $encryptionInfo = ConvertFrom-Json ([IO.File]::ReadAllText($fullPath))
            if ($encryptionInfo.fileEncryptionInfo) { $encryptionInfo = $encryptionInfo.fileEncryptionInfo }
            $zipName = ($baseName -replace 'intunewin$', 'zip')
            if ($zipName -eq $baseName) { $zipName = "$baseName.zip" }
            $destination = [IO.Path]::Combine($pkgPath, $zipName)
            Start-DecryptFile $picked $destination $encryptionInfo.encryptionKey $encryptionInfo.initializationVector
            try { [IO.File]::Delete($picked) }
            catch { Write-LogError "Failed to delete exported encrypted file" $_.Exception }
        } else {
            Write-Log "Decryption file for $($obj.displayName) not found. Skipping decryption" 2
        }
    }
    finally {
        # Always clear the status overlay - an error mid-download previously left
        # it stuck on screen and looked like a hang.
        Write-Status ""
    }
}

##############################################
# Captures
##############################################

$sbAddCheckbox    = ${function:Add-AvaloniaExportPropertyCheckbox}
$sbAddDetailsBtn  = ${function:Add-AvaloniaDetailsButton}
$sbSetCache       = ${function:Set-CacheObject}
$sbWriteLogError  = ${function:Write-LogError}
$sbGetSetting     = ${function:Get-SettingStoreValue}
$sbInvokeUploadApp   = ${function:Invoke-UploadAppFileAvalonia}
$sbInvokeDownloadApp = ${function:Invoke-DownloadAppFileAvalonia}

##############################################
# ScriptBlocks
##############################################

$SBAddApplicationUIExportExtension = {
    param($Form)

    if (-not $Form) { return }

    # "Export script" checkbox (shared with script-typed exports).
    $chkScript = & $sbAddCheckbox -Form $Form `
        -CheckboxName 'chkExportScript' `
        -LabelText    'Export script' `
        -ToolTip      'Export the script associated with selected profiles' `
        -InitialChecked $true
    if ($chkScript) {
        & $sbSetCache "ExportScripts" $chkScript.IsChecked
        $chkScript.add_IsCheckedChanged({
            param($s, $e)
            & $sbSetCache "ExportScripts" $s.IsChecked
        }.GetNewClosure())
    }

    # "Export application file" checkbox — initial state from setting.
    $initial = ((& $sbGetSetting "Intune" "ExportAppFile" "false") -eq "true")
    $chkApp = & $sbAddCheckbox -Form $Form `
        -CheckboxName 'chkExportAppContent' `
        -LabelText    'Export application file' `
        -ToolTip      'Export the application file. Note: Application file will only be exported if encryption file is found.' `
        -InitialChecked $initial
    if ($chkApp) {
        & $sbSetCache "ExportAppContent" $chkApp.IsChecked
        $chkApp.add_IsCheckedChanged({
            param($s, $e)
            & $sbSetCache "ExportAppContent" $s.IsChecked
        }.GetNewClosure())
    }
}.GetNewClosure()

$SBAddApplicationUIDetailsExtension = {
    param($Form)

    if (-not $Form) { return }

    $onUpload = {
        param($s, $e)
        $row = $script:dgIntuneManagerObjects.SelectedItem
        if (-not $row) { return }
        try { & $sbInvokeUploadApp $row }
        catch { & $sbWriteLogError "Upload application file failed" $_.Exception }
    }.GetNewClosure()

    $onDownload = {
        param($s, $e)
        $row = $script:dgIntuneManagerObjects.SelectedItem
        if (-not $row) { return }
        try { & $sbInvokeDownloadApp $row }
        catch { & $sbWriteLogError "Download application file failed" $_.Exception }
    }.GetNewClosure()

    # Insert order matches the WPF original: Upload at 0 first, then Download
    # at 0 — net layout left-to-right is Download, Upload, [existing buttons].
    & $sbAddDetailsBtn -Form $Form -Name 'btnUploadAppfile' -Content 'Upload' `
        -InsertIndex 0 -OnClick $onUpload | Out-Null
    & $sbAddDetailsBtn -Form $Form -Name 'btnDownloadAppfile' -Content 'Download' `
        -InsertIndex 0 -OnClick $onDownload | Out-Null
}.GetNewClosure()

##############################################
# Init
##############################################

function Invoke-InitializeApplicationUIExtensions
{
    $typeClasses = @()
    try { $typeClasses += Get-SingletonObject "ApplicationType" } catch {}

    foreach ($typeClass in $typeClasses) {
        if (-not $typeClass) { continue }
        Add-ObjectMethod $typeClass "AddUIExportExtensions" $SBAddApplicationUIExportExtension
        Add-ObjectMethod $typeClass "AddUIDetailsExtension"  $SBAddApplicationUIDetailsExtension
    }
}

Invoke-InitializeApplicationUIExtensions
