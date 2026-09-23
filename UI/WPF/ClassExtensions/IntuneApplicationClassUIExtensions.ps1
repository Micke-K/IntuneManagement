##############################################
# ScriptBlocks
##############################################

$SBAddApplicationUIExportExtension = [scriptblock] {
    param($Form)

    if (-not $Form) { return }

    Add-UIScriptExportExtensions $Form

    $grdExportProperties = $null
    try {
        $grdExportProperties = $Form.FindName("grdExportProperties")
    }
    catch {}
    if (-not $grdExportProperties) { return }

    if ($grdExportProperties.FindName("chkExportAppContent")) {
        return # Already added
    }

    $xaml = @"
<StackPanel $($script:wpfNS) Orientation="Horizontal" Margin="0,0,5,0">
<Label Content="Export application file" />
<Rectangle Style="{DynamicResource InfoIcon}" ToolTip="Export the application file. Note: Application file will only be exported if ecryption file is found." />
</StackPanel>
"@
    try {
        $label = [Windows.Markup.XamlReader]::Parse($xaml)

        $chkExportAppContent = [System.Windows.Controls.CheckBox]::new()
        $chkExportAppContent.IsChecked = ((Get-SettingStoreValue "Intune" "ExportAppFile" "false") -eq "true")
        $chkExportAppContent.VerticalAlignment = "Center" 
        $chkExportAppContent.Name = "chkExportAppContent" 

        Set-CacheObject "ExportAppContent" $chkExportAppContent.IsChecked

        $rd = [System.Windows.Controls.RowDefinition]::new()
        $rd.Height = [double]::NaN            
        $grdExportProperties.RowDefinitions.Add($rd)

        $label.SetValue([System.Windows.Controls.Grid]::RowProperty, $grdExportProperties.RowDefinitions.Count - 1)
        $chkExportAppContent.SetValue([System.Windows.Controls.Grid]::RowProperty, $grdExportProperties.RowDefinitions.Count - 1)
        $chkExportAppContent.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)

        $grdExportProperties.Children.Add($label)
        $grdExportProperties.Children.Add($chkExportAppContent)

        $Form.RegisterName($chkExportAppContent.Name, $chkExportAppContent)

        $changed = [scriptblock] {
            param($S, $E)
            Set-CacheObject "ExportAppContent" $this.IsChecked
        }

        $chkExportAppContent.Add_Checked($changed)
        $chkExportAppContent.Add_Unchecked($changed)
    }
    catch {}
}

$SBAddApplicationUIDetailsExtension = [scriptblock] {
    param($Form)

    $buttonPanel = $Form.FindName("pnlButtons")
    if (-not $buttonPanel) { return }

    $btnUpload = New-Object System.Windows.Controls.Button    
    $btnUpload.Content = 'Upload'
    $btnUpload.Name = 'btnUploadAppfile'
    $btnUpload.Margin = "0,0,5,0"  
    $btnUpload.Width = "100"
    
    $btnUpload.Add_Click({
            if ($script:dgIntuneManagerObjects.SelectedItem.Object.publishingState -ne "notPublished") {
                # Only allow upload of not published apps
                # Use portal to replace app file...
                if (([System.Windows.MessageBox]::Show("Are you sure you want to upload a new file for the app?`n`nApplication:`n$($script:dgIntuneManagerObjects.SelectedItem.Name)", "Update app file?", "YesNo", "Warning")) -ne "Yes") {
                    return
                }
            }
    
            $pkgPath = Get-SettingValue "IntuneAppPackagesFolder"

            $of = [System.Windows.Forms.OpenFileDialog]::new()
            $of.FileName = $script:dgIntuneManagerObjects.SelectedItem.Object.fileName
            $of.DefaultExt = "*.intunewin"
            $of.Filter = "Intune Win32 (*.intunewin)|*.*"
            $of.Multiselect = $false

            if ($pkgPath -and [IO.Directory]::Exists($pkgPath)) {
                $of.InitialDirectory = $pkgPath
            }
        
            if ($of.ShowDialog() -eq "OK") {
                Write-Status "Import $($script:dgIntuneManagerObjects.SelectedItem.Object.displayName) file"
                Start-ApplicationImportFile $script:dgIntuneManagerObjects.SelectedItem $of.FileName
                Write-Status ""
            }
        })

    $buttonPanel.Children.Insert(0, $btnUpload)

    $btnDownload = New-Object System.Windows.Controls.Button    
    $btnDownload.Content = 'Download'
    $btnDownload.Name = 'btnDownloadAppfile'
    $btnDownload.Margin = "0,0,5,0"  
    $btnDownload.Width = "100"
    
    $btnDownload.Add_Click({
            Write-Status "Download file"
            $obj = $script:dgIntuneManagerObjects.SelectedItem.Object
            #$obj = Invoke-MSGraphAPI -Url "/deviceAppManagement/mobileApps/$($obj.id)"

            $pkgPath = Get-SettingValue "IntuneAppDownloadFolder" (Get-SettingValue "IntuneAppPackagesFolder")

            $dlgSave = [System.Windows.Forms.SaveFileDialog]::new()
            $dlgSave.InitialDirectory = $pkgPath
            $dlgSave.FileName = ($obj.FileName + ".encrypted")
            $dlgSave.DefaultExt = "*.encrypted"
            $dlgSave.Filter = "Encrypted intunewin (*.encrypted)|*.encrypted|All files (*.*)|*.*"

            if ($dlgSave.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $dlgSave.Filename) {
                $contentFileObj = Start-DownloadAppContent $script:dgIntuneManagerObjects.SelectedItem $dlgSave.FileName

                if ([IO.File]::Exists($dlgSave.FileName)) {
                    $fullPath = Find-AppEncryptionFile $obj $contentFileObj $pkgPath
                    if ([IO.File]::Exists($fullPath) -eq $false) {
                        if (([System.Windows.MessageBox]::Show("Could not find decryption file for $($obj.displayName)`nApp Id: $($obj.id)`nContent version $($obj.committedContentVersion)`n`nDo you want to browse for the file?", "Encryption file not found", "YesNo", "Warning")) -eq "Yes") {
                            $of = [System.Windows.Forms.OpenFileDialog]::new()
                            $of.InitialDirectory = $pkgPath
                            $of.DefaultExt = "*.json"
                            $of.Filter = "Json (*.json)|*.json"
                            $of.Multiselect = $false
                        
                            if ($of.ShowDialog() -eq "OK") {
                                $fullPath = $of.FileName
                            }                    
                        }
                    }

                    if ([IO.File]::Exists($fullPath)) {
                        Write-Status "Decrypting file"
                        $encryptionInfo = ConvertFrom-Json ([IO.File]::ReadAllText($fullPath))
                        if ($encryptionInfo.fileEncryptionInfo) {
                            $encryptionInfo = $encryptionInfo.fileEncryptionInfo
                        }
                        $destination = $pkgPath + ("\$($obj.FileName)" -replace 'intunewin$', 'zip')
                        Start-DecryptFile $dlgSave.Filename $destination $encryptionInfo.encryptionKey $encryptionInfo.initializationVector
                        try { [IO.File]::Delete($dlgSave.Filename) }
                        catch {
                            Write-LogError "Failed to delete exported encrypted file" $_.Exception
                        }                
                    }
                    else {
                        Write-Log "Decryption file for $($obj.displayName) not found. Skipping decryption" 2
                    }
                }
            }

            Write-Status ""
        })

    $buttonPanel.Children.Insert(0, $btnDownload)
}

##############################################
# Functions
##############################################

function Invoke-InitializeApplicationUIExtensions {
    $typeClasses = @()

    $typeClasses += Get-SingletonObject "ApplicationType"

    foreach ($typeClass in $typeClasses) {
        Add-ObjectMethod $typeClass "AddUIExportExtensions" $SBAddApplicationUIExportExtension
        Add-ObjectMethod $typeClass "AddUIDetailsExtension" $SBAddApplicationUIDetailsExtension
    }
}

##############################################
# Initialize
##############################################

Invoke-InitializeApplicationUIExtensions