##############################################
# ScriptBlocks
##############################################

$SBAddPolicyFileUIExportExtension = [scriptblock] {
    param($Form)

    if (-not $Form) { return }

    Add-UIPolicyFileExportExtensions $Form
}

$SBAddPolicyFileUIDetailsExtension = [scriptblock] {
    param($Form)

    $buttonPanel = $Form.FindName("pnlButtons")
    if (-not $buttonPanel) { return }

    $btn = New-Object System.Windows.Controls.Button    
    $btn.Content = 'Download'
    $btn.Name = 'btnDownload'
    $btn.Margin = "0,0,5,0"  
    $btn.Width = "100"
    
    $btn.Add_Click({
            Invoke-DownloadPolicyFile $script:dgIntuneManagerObjects.SelectedItem
        })

    $buttonPanel.Children.Insert(0, $btn)

    $btn = New-Object System.Windows.Controls.Button    
    $btn.Content = 'Edit'
    $btn.Name = 'btnEdit'
    $btn.Margin = "0,0,5,0"  
    $btn.Width = "100"
    
    $btn.Add_Click({
            Invoke-EditPolicyFile $script:dgIntuneManagerObjects.SelectedItem
        })

    $buttonPanel.Children.Insert(1, $btn)
}

##############################################
# Functions
##############################################

function Invoke-InitializePolicyFileUIExtensions {
    $typeClasses = @()

    Get-SubClasses "IntunePolicyTypeBase" | ForEach-Object {
        if(Test-ClassIsAbstract $_) { return }
        try {
            $tmpClass = Get-SingletonObject $_.Name
            if ($null -ne $tmpClass.PolicyGroup -and $tmpClass._PolicyFileAttributes.Count -gt 0) {
                $typeClasses += $tmpClass
            }
        }
        catch {}
    }

    foreach ($typeClass in $typeClasses) {
        Add-ObjectMethod $typeClass "AddUIExportExtensions" $SBAddPolicyFileUIExportExtension
        Add-ObjectMethod $typeClass "AddUIDetailsExtension" $SBAddPolicyFileUIDetailsExtension
    }
}

function Add-UIPolicyFileExportExtensions {
    param(
        $Form,
        $Description = "Export the files associated with selected profiles"
        )
    
    $grdExportProperties = $null
    try {
        $grdExportProperties = $Form.FindName("grdExportProperties")
    }
    catch {}
    if (-not $grdExportProperties) { return }

    if ($grdExportProperties.FindName("chkExportPolicyFile")) {
        return # Already added
    }

    $xaml = @"
<StackPanel $($script:wpfNS) Orientation="Horizontal" Margin="0,0,5,0">
<Label Content="Export file" />
<Rectangle Style="{DynamicResource InfoIcon}" ToolTip="$Description" />
</StackPanel>
"@
    try {
        $label = [Windows.Markup.XamlReader]::Parse($xaml)

        $chkExportPolicyFiles = [System.Windows.Controls.CheckBox]::new()
        $chkExportPolicyFiles.IsChecked = $true
        $chkExportPolicyFiles.VerticalAlignment = "Center" 
        $chkExportPolicyFiles.Name = "chkExportPolicyFiles" 

        Set-CacheObject "ExportPolicyFiles" $chkExportPolicyFiles.IsChecked

        $rd = [System.Windows.Controls.RowDefinition]::new()
        $rd.Height = [double]::NaN            
        $grdExportProperties.RowDefinitions.Add($rd)

        $label.SetValue([System.Windows.Controls.Grid]::RowProperty, $grdExportProperties.RowDefinitions.Count - 1)
        $chkExportPolicyFiles.SetValue([System.Windows.Controls.Grid]::RowProperty, $grdExportProperties.RowDefinitions.Count - 1)
        $chkExportPolicyFiles.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)

        $grdExportProperties.Children.Add($label)
        $grdExportProperties.Children.Add($chkExportPolicyFiles)

        $Form.RegisterName($chkExportPolicyFiles.Name, $chkExportPolicyFiles)

        $changed = [scriptblock] {
            param($S, $E)
            Set-CacheObject "ExportPolicyFiles" $this.IsChecked
        }

        $chkExportPolicyFiles.Add_Checked($changed)
        $chkExportPolicyFiles.Add_Unchecked($changed)
    }
    catch {}
}

function Invoke-EditPolicyFile {
    param($Policy)

    if (-not $Policy) { return }

    if ($Policy.IsFullObject -eq $false) {
        [void]$Policy.Get()
    }

    $attribute = $Policy.PolicyType._PolicyFileAttributes[0] # TODO: Support multiple attributes

    if (-not $Policy.JsonObject.$attribute) { return }
    $script:currentPolicyFileObject = $Policy

    $script:editForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\EditScriptDialog.xaml"))

    if (-not $script:editForm) { return }

    $script:UIProvider.SetXamlProperty($script:editForm, "txtEditScriptTitle", "Text", "Edit: $($Policy.JsonObject.displayName)")

    $policyFileText = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($Policy.JsonObject.$attribute))
    $script:UIProvider.SetXamlProperty($script:editForm, "txtScriptText", "Text", $policyFileText)

    $script:currentModal = $null
    if ($script:grdModal.Children.Count -gt 0) {
        $script:currentModal = $script:grdModal.Children[0]
    }

    $script:UIProvider.AddXamlEvent($script:editForm, "btnSaveScriptEdit", "add_click", {
            $policyFileText = $script:UIProvider.GetXamlProperty($script:editForm, "txtScriptText", "Text")
            $pre = [System.Text.Encoding]::UTF8.GetPreamble()
            $utfBOM = [System.Text.Encoding]::UTF8.GetString($pre)
            if ($policyFileText.startsWith($utfBOM)) {
                # Remove UTF8 BOM bytes
                $policyFileText = $policyFileText.Remove(0, $utfBOM.Length)
            }
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($policyFileText)
            $encodedText = [Convert]::ToBase64String($bytes)
            $attribute = $script:currentPolicyFileObject.PolicyType._PolicyFileAttributes[0]

            if ($script:currentPolicyFileObject.JsonObject.$attribute -ne $encodedText) {
                # Save script
                if (($script:UIProvider.ShowMessageBox("Are you sure you want to update the file?`n`nObject:`n$($script:currentPolicyFileObject.Name)", "Update file?", "YesNo", "Warning")) -eq "Yes") {
                    Write-Status "Update $($script:currentPolicyFileObject.Name)"
                    $clonedObject = $script:currentPolicyFileObject.Clone()
                    $clonedObject.JsonObject.$attribute = $encodedText
                    Remove-GraphPropertiesForImport $clonedObject $clonedObject.JsonObject
                    foreach ($prop in $script:currentPolicyFileObject.PolicyType.PropertiesToRemoveForUpdate) {
                        Remove-Property $clonedObject.JsonObject $prop
                    }                
                    Remove-Property $clonedObject.JsonObject "Assignments"
                    Remove-Property $clonedObject.JsonObject "isAssigned"

                    if ($script:currentPolicyFileObject.Set($null, $clonedObject.JsonString, "PATCH", $null)) {
                        Write-Log "File saved successfully"
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

function Invoke-DownloadPolicyFile
{
    param($Policy)

    if(-not $Policy) { return }

    if($Policy.IsFullObject -eq $false) {
        [void]$Policy.Get()
    }    
    Write-Status ""

    $attribute = $Policy.PolicyType._PolicyFileAttributes[0]

    if($Policy.JsonObject.$attribute)
    {
        $fileName = ?? $Policy.JsonObject.FileName "$($attribute).json"
        Write-Log "Download PowerShell file '$($fileName)' from $($Policy.Name)"
        
        $dlgSave = [System.Windows.Forms.SaveFileDialog]::new()
        $dlgSave.InitialDirectory = Get-SettingValue "RootFolder" $env:Temp
        $dlgSave.FileName = $fileName
        if($dlgSave.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $dlgSave.Filename)
        {
            # Changed to WriteAllBytes to get rid of BOM characters from Custom Attribute file 
            [IO.File]::WriteAllBytes($dlgSave.FileName, ([System.Convert]::FromBase64String($Policy.JsonObject.$attribute)))
        }
    }    
}

##############################################
# Initialize
##############################################

Invoke-InitializePolicyFileUIExtensions