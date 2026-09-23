##############################################
# ScriptBlocks
##############################################

$SBAddScriptUIExportExtension = [scriptblock] {
    param($Form)

    if (-not $Form) { return }

    Add-UIScriptExportExtensions $Form
}

$SBAddScriptUIDetailsExtension = [scriptblock] {
    param($Form)

    $buttonPanel = $Form.FindName("pnlButtons")
    if (-not $buttonPanel) { return }

    $btn = New-Object System.Windows.Controls.Button    
    $btn.Content = 'Download'
    $btn.Name = 'btnDownload'
    $btn.Margin = "0,0,5,0"  
    $btn.Width = "100"
    
    $btn.Add_Click({
            Invoke-DownloadScript $script:dgIntuneManagerObjects.SelectedItem
        })

    $buttonPanel.Children.Insert(0, $btn)

    $btn = New-Object System.Windows.Controls.Button    
    $btn.Content = 'Edit'
    $btn.Name = 'btnEdit'
    $btn.Margin = "0,0,5,0"  
    $btn.Width = "100"
    
    $btn.Add_Click({
            Invoke-EditScript $script:dgIntuneManagerObjects.SelectedItem
        })

    $buttonPanel.Children.Insert(1, $btn)
}

##############################################
# Functions
##############################################

function Invoke-InitializeScriptUIExtensions {
    $typeClasses = @()

    Get-SubClasses "ScriptTypeBase" | ForEach-Object { 
        try {
            $tmpClass = Get-SingletonObject $_.Name
            if ($null -ne $tmpClass.PolicyGroup) {
                $typeClasses += $tmpClass
            }
        }
        catch {}
    }

    foreach ($typeClass in $typeClasses) {
        Add-ObjectMethod $typeClass "AddUIExportExtensions" $SBAddScriptUIExportExtension
        Add-ObjectMethod $typeClass "AddUIDetailsExtension" $SBAddScriptUIDetailsExtension
    }
}

function Add-UIScriptExportExtensions {
    param(
        $Form,
        $Description = "Export the script associated with selected profiles"
        )
    
    $grdExportProperties = $null
    try {
        $grdExportProperties = $Form.FindName("grdExportProperties")
    }
    catch {}
    if (-not $grdExportProperties) { return }

    if ($grdExportProperties.FindName("chkExportScript")) {
        return # Already added
    }

    $xaml = @"
<StackPanel $($script:wpfNS) Orientation="Horizontal" Margin="0,0,5,0">
<Label Content="Export script" />
<Rectangle Style="{DynamicResource InfoIcon}" ToolTip="$Description" />
</StackPanel>
"@
    try {
        $label = [Windows.Markup.XamlReader]::Parse($xaml)

        $chkExportScript = [System.Windows.Controls.CheckBox]::new()
        $chkExportScript.IsChecked = $true
        $chkExportScript.VerticalAlignment = "Center" 
        $chkExportScript.Name = "chkExportScript" 

        Set-CacheObject "ExportScripts" $chkExportScript.IsChecked

        $rd = [System.Windows.Controls.RowDefinition]::new()
        $rd.Height = [double]::NaN            
        $grdExportProperties.RowDefinitions.Add($rd)

        $label.SetValue([System.Windows.Controls.Grid]::RowProperty, $grdExportProperties.RowDefinitions.Count - 1)
        $chkExportScript.SetValue([System.Windows.Controls.Grid]::RowProperty, $grdExportProperties.RowDefinitions.Count - 1)
        $chkExportScript.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)

        $grdExportProperties.Children.Add($label)
        $grdExportProperties.Children.Add($chkExportScript)

        $Form.RegisterName($chkExportScript.Name, $chkExportScript)

        $changed = [scriptblock] {
            param($S, $E)
            Set-CacheObject "ExportScripts" $this.IsChecked
        }

        $chkExportScript.Add_Checked($changed)
        $chkExportScript.Add_Unchecked($changed)
    }
    catch {}
}

function Invoke-DownloadScript {
    param($ScriptPolicy)

    if (-not $ScriptPolicy) { return }

    if ($ScriptPolicy.IsFullObject -eq $false) {
        [void]$ScriptPolicy.Get()
    }

    if (-not $ScriptPolicy.JsonObject.scriptContent) { return }

    Write-Status ""

    $dlgSave = [System.Windows.Forms.SaveFileDialog]::new()
    $dlgSave.InitialDirectory = Get-SettingValue "RootFolder" $env:Temp
    $dlgSave.FileName = $ScriptPolicy.JsonObject.FileName
    if($dlgSave.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $dlgSave.Filename)
    {
        Save-IntuneScriptContent $ScriptPolicy $dlgSave.FileName | Out-Null
    }
}

function Invoke-EditScript {
    param($ScriptPolicy)

    if (-not $ScriptPolicy) { return }

    if ($ScriptPolicy.IsFullObject -eq $false) {
        [void]$ScriptPolicy.Get()
    }    

    if (-not $ScriptPolicy.JsonObject.scriptContent) { return }
    $script:currentScriptObject = $ScriptPolicy

    $script:editForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\EditScriptDialog.xaml"))

    if (-not $script:editForm) { return }

    $script:UIProvider.SetXamlProperty($script:editForm, "txtEditScriptTitle", "Text", "Edit: $($ScriptPolicy.JsonObject.displayName)")

    $scriptText = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ScriptPolicy.JsonObject.scriptContent))
    $script:UIProvider.SetXamlProperty($script:editForm, "txtScriptText", "Text", $scriptText)

    $script:currentModal = $null
    if ($script:grdModal.Children.Count -gt 0) {
        $script:currentModal = $script:grdModal.Children[0]
    }

    $script:UIProvider.AddXamlEvent($script:editForm, "btnSaveScriptEdit", "add_click", {
            $scriptText = $script:UIProvider.GetXamlProperty($script:editForm, "txtScriptText", "Text")
            $pre = [System.Text.Encoding]::UTF8.GetPreamble()
            $utfBOM = [System.Text.Encoding]::UTF8.GetString($pre)
            if ($scriptText.startsWith($utfBOM)) {
                # Remove UTF8 BOM bytes
                $scriptText = $scriptText.Remove(0, $utfBOM.Length)
            }
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($scriptText)
            $encodedText = [Convert]::ToBase64String($bytes)

            if ($script:currentScriptObject.JsonObject.scriptContent -ne $encodedText) {
                # Save script
                if (($script:UIProvider.ShowMessageBox("Are you sure you want to update the script?`n`nObject:`n$($script:currentScriptObject.Name)", "Update script?", "YesNo", "Warning")) -eq "Yes") {
                    Write-Status "Update $($script:currentScriptObject.Name)"
                    $clonedObject = $script:currentScriptObject.Clone()
                    $clonedObject.JsonObject.scriptContent = $encodedText                
                    Remove-GraphPropertiesForImport $clonedObject $clonedObject.JsonObject
                    foreach ($prop in $script:currentScriptObject.PolicyType.PropertiesToRemoveForUpdate) {
                        Remove-Property $clonedObject.JsonObject $prop
                    }                
                    Remove-Property $clonedObject.JsonObject "Assignments"
                    Remove-Property $clonedObject.JsonObject "isAssigned"

                    if ($script:currentScriptObject.Set($null, $clonedObject.JsonString, "PATCH", $null)) {
                        Write-Log "Script saved successfully"
                    }
                    else {
                        Write-Log "Failed to save script" 3
                    }
                    Write-Status ""
                }
            }

            $script:grdModal.Children.Clear()
            if ($script:currentModal) {
                $script:grdModal.Children.Add($script:currentModal)
            }
            [System.Windows.Forms.Application]::DoEvents()
        })

    $script:UIProvider.AddXamlEvent($script:editForm, "btnCancelScriptEdit", "add_click", {
            $script:grdModal.Children.Clear()
            if ($script:currentModal) {
                $script:grdModal.Children.Add($script:currentModal)
            }
            [System.Windows.Forms.Application]::DoEvents()
        })
    
    $script:grdModal.Children.Clear()
    $script:editForm.SetValue([System.Windows.Controls.Grid]::RowProperty, 1)
    $script:editForm.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)
    $script:grdModal.Children.Add($script:editForm) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
}

##############################################
# Initialize
##############################################

Invoke-InitializeScriptUIExtensions
