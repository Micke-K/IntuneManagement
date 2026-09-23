# Avalonia parallel of UI/WPF/Extensions/IntuneManagerUI.ps1's
# Add-IntuneManagerExportUIExtensions / Add-IntuneManagerImportUIExtensions
# wrappers, plus shared helpers used by ClassExtensions/*UIExtension.ps1.
#
# The wrappers iterate the selected policy types and call their attached
# AddUIExportExtensions / AddUIImportExtensions ScriptMethods. The
# Avalonia versions of those ScriptMethods live in
# UI/Avalonia/ClassExtensions/ — same method names so the wrappers stay
# WPF-compatible. New file per architecture R9 (this is a small new
# subsystem; folding it into IntuneManagerUI.ps1 would push that file
# further away from "one public function per file").

function Add-IntuneManagerExportUIExtensions
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyTypeBase[]]
        $InputObject,
        [Parameter(Mandatory = $true)]
        [Object]
        $Form
    )

    Begin { Write-Log "Add Export UI Extensions - Start" }

    Process {
        foreach ($policyType in $InputObject) {
            if ($policyType.AddUIExportExtensions) {
                try { $policyType.AddUIExportExtensions($Form) }
                catch { Write-LogError "AddUIExportExtensions failed for $($policyType.Name)" $_.Exception }
            }
        }
    }

    End { Write-Log "Add Export UI Extensions - Done" }
}

function Add-IntuneManagerImportUIExtensions
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyTypeBase[]]
        $InputObject,
        [Parameter(Mandatory = $true)]
        [Object]
        $Form
    )

    Begin { Write-Log "Add Import UI Extensions - Start" }

    Process {
        foreach ($policyType in $InputObject) {
            if ($policyType.AddUIImportExtensions) {
                try { $policyType.AddUIImportExtensions($Form) }
                catch { Write-LogError "AddUIImportExtensions failed for $($policyType.Name)" $_.Exception }
            }
        }
    }

    End { Write-Log "Add Import UI Extensions - Done" }
}

function Add-AvaloniaExportPropertyCheckbox
{
    # Shared helper used by every per-type Export extension to graft a
    # labeled CheckBox onto grdExportProperties. Returns the CheckBox so
    # callers can wire add_IsCheckedChanged. Returns $null if the form
    # doesn't have the grid (e.g. wrong dialog), or if a checkbox with
    # the same name was already added (prevents double-add on re-open).
    param(
        [Parameter(Mandatory)] $Form,
        [Parameter(Mandatory)] [string] $CheckboxName,
        [Parameter(Mandatory)] [string] $LabelText,
        [string] $ToolTip,
        [bool]   $InitialChecked = $true
    )

    $hostType = Get-AvaloniaHost
    $grdExportProperties = $hostType::FindByName($Form, 'grdExportProperties')
    if (-not $grdExportProperties) { return $null }

    # Idempotency guard — same as the WPF helper.
    foreach ($child in @($grdExportProperties.Children)) {
        if ($child -is [Avalonia.Controls.CheckBox] -and $child.Name -eq $CheckboxName) {
            return $child
        }
    }

    $rd = [Avalonia.Controls.RowDefinition]::new()
    $rd.Height = [Avalonia.Controls.GridLength]::Auto
    $grdExportProperties.RowDefinitions.Add($rd)
    $rowIndex = $grdExportProperties.RowDefinitions.Count - 1

    $label = [Avalonia.Controls.Label]::new()
    $label.Content = $LabelText
    $label.Margin  = [Avalonia.Thickness]::new(0, 5, 5, 0)
    $label.VerticalAlignment = [Avalonia.Layout.VerticalAlignment]::Center
    if ($ToolTip) { [Avalonia.Controls.ToolTip]::SetTip($label, $ToolTip) }

    $chk = [Avalonia.Controls.CheckBox]::new()
    $chk.Name = $CheckboxName
    $chk.IsChecked = $InitialChecked
    $chk.Margin  = [Avalonia.Thickness]::new(0, 5, 0, 0)
    $chk.VerticalAlignment = [Avalonia.Layout.VerticalAlignment]::Center

    [Avalonia.Controls.Grid]::SetRow($label, $rowIndex)
    [Avalonia.Controls.Grid]::SetColumn($label, 0)
    [Avalonia.Controls.Grid]::SetRow($chk, $rowIndex)
    [Avalonia.Controls.Grid]::SetColumn($chk, 1)

    $grdExportProperties.Children.Add($label) | Out-Null
    $grdExportProperties.Children.Add($chk)   | Out-Null

    return $chk
}

function Add-AvaloniaDetailsButton
{
    # Shared helper used by per-type Details extensions to insert a
    # button into the ObjectDetails JSON tab's pnlButtons WrapPanel.
    # The Avalonia WrapPanel exposes Children.Insert(int, control), so
    # we mirror the WPF Insert(0, ...) pattern that puts new buttons
    # at the left of "Load full" / "Copy".
    param(
        [Parameter(Mandatory)] $Form,
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [string] $Content,
        [Parameter(Mandatory)] [scriptblock] $OnClick,
        [int] $InsertIndex = 0
    )

    $hostType    = Get-AvaloniaHost
    $buttonPanel = $hostType::FindByName($Form, 'pnlButtons')
    if (-not $buttonPanel) { return $null }

    $btn = [Avalonia.Controls.Button]::new()
    $btn.Name    = $Name
    $btn.Content = $Content
    $btn.MinWidth = 100
    $btn.Margin  = [Avalonia.Thickness]::new(0, 0, 5, 0)
    # Module-bind the handler: an unbound scriptblock runs against the caller's
    # session state at click time, where the module's private functions
    # (Save-IntuneScriptContent, Start-DownloadAppContent, Write-LogError, ...)
    # do not resolve - every per-type Details button failed silently.
    $btn.add_Click((ConvertTo-AvaloniaEventScriptBlock $OnClick))

    if ($InsertIndex -lt 0 -or $InsertIndex -gt $buttonPanel.Children.Count) {
        $buttonPanel.Children.Add($btn) | Out-Null
    } else {
        $buttonPanel.Children.Insert($InsertIndex, $btn)
    }

    return $btn
}

function Show-AvaloniaScriptEditor
{
    # param() MUST be the first statement in the body - only comments may precede
    # it. With a statement in front, PowerShell parses `param(...)` as a COMMAND
    # named "param" instead of a parameter block: the file still parses, the
    # function silently ends up with NO parameters, and every call fails on
    # "A parameter cannot be found that matches parameter name 'Title'". So
    # $ui is assigned below the block, not above it.
    #
    # Avalonia equivalent of UI/WPF/XAML/EditScriptDialog.xaml + the
    # Invoke-EditScript / Invoke-EditPolicyFile WPF flows. Hosted as a
    # nested grdModal level via Show-ModalForm so it stacks on top of
    # the Object Details modal; dismissed via Close-TopModalObject
    # (Show-ModalObject with no args would tear down the parent Details
    # modal too).
    #
    # Accepts an OnSave callback rather than returning text — Show-ModalForm
    # is non-blocking on the overlay path, so the save work has to execute
    # inside the Save click handler. The callback receives the post-edit
    # text as its only argument.
    param(
        [Parameter(Mandatory)] [string]      $Title,
        [Parameter(Mandatory)] [string]      $InitialText,
        [Parameter(Mandatory)] [scriptblock] $OnSave
    )

    $ui = $script:UIProvider

    $editForm = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/EditScriptDialog.axaml'))
    if (-not $editForm) { return }

    $hostType = Get-AvaloniaHost
    $txtTitle  = $hostType::FindByName($editForm, 'txtEditScriptTitle')
    $txtScript = $hostType::FindByName($editForm, 'txtScriptText')
    $btnSave   = $hostType::FindByName($editForm, 'btnSaveScriptEdit')
    $btnCancel = $hostType::FindByName($editForm, 'btnCancelScriptEdit')

    if ($txtTitle)  { $txtTitle.Text  = $Title }
    if ($txtScript) { $txtScript.Text = $InitialText }

    # The editor's text box and save callback live in module scope: captured
    # locals are stripped from the handler, so Save used to close the dialog
    # having silently discarded every edit.
    $script:_scriptEditorState = @{ TxtScript = $txtScript; OnSave = $OnSave }

    if ($btnSave) {
        $btnSave.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            $st = $script:_scriptEditorState
            if (-not $st) { Close-TopModalObject; return }
            $newText = if ($st.TxtScript) { [string]$st.TxtScript.Text } else { '' }
            try { if ($st.OnSave) { & $st.OnSave $newText } }
            catch { Write-LogError "Script editor OnSave callback failed" $_.Exception }
            $script:_scriptEditorState = $null
            Close-TopModalObject
        }))
    }
    if ($btnCancel) {
        $btnCancel.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            Close-TopModalObject
        }.GetNewClosure()))
    }

    $editForm.add_KeyDown((ConvertTo-AvaloniaEventScriptBlock {
        param($s, $e)
        if ($e.Key -eq [Avalonia.Input.Key]::Escape) {
            Close-TopModalObject
            $e.Handled = $true
        }
    }.GetNewClosure()))

    $ui.ShowModalForm($Title, $editForm, $true)
}
