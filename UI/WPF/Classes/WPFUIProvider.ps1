#ImportOrder 60
#
# WPF concrete subclass of UIProvider. Each method captures the
# corresponding free function from CoreUIWPF.ps1 (and friends) at construction
# time via ${function:X}, then invokes the captured scriptblock with `&`.
# This is defensive: PS native class methods *may* see module-scoped free
# functions, but the documented workaround for ScriptMethod scriptblocks
# (see [[scriptmethod-loses-module-scope]]) is the capture pattern. Until a
# smoke test confirms direct calls work, we capture explicitly.
#
# Folder picking delegates to the existing WPF Get-Folder helper. Open/save
# file dialogs are intentionally not part of the base provider until both
# backends have an implemented contract.

class WPFUIProvider : UIProvider {

    #region Modal / dialog

    [void] ShowMessageBox([string]$Text) {
        $f = ${function:Show-MessageBox}
        & $f $Text | Out-Null
    }

    [void] ShowMessageBox([string]$Text, [string]$Caption) {
        $f = ${function:Show-MessageBox}
        & $f $Text $Caption | Out-Null
    }

    [object] ShowMessageBox([string]$Text, [string]$Caption, [string]$Button, [string]$Icon) {
        $f = ${function:Show-MessageBox}
        return (& $f $Text $Caption $Button $Icon)
    }

    [void] ShowModalForm([string]$FormTitle, [object]$FormObject) {
        $f = ${function:Show-ModalForm}
        & $f $FormTitle $FormObject
    }

    [void] ShowModalForm([string]$FormTitle, [object]$FormObject, [bool]$HideButtons) {
        $f = ${function:Show-ModalForm}
        if ($HideButtons) { & $f $FormTitle $FormObject -HideButtons }
        else              { & $f $FormTitle $FormObject }
    }

    [void] ShowModalObject() {
        $f = ${function:Show-ModalObject}
        & $f
    }

    [void] ShowModalObject([object]$Obj) {
        $f = ${function:Show-ModalObject}
        & $f $Obj
    }

    [void] CloseTopModalObject() {
        $f = ${function:Close-TopModalObject}
        & $f
    }

    [string] ShowInputDialog([string]$FormTitle, [string]$FormText, [string]$DefaultValue) {
        $f = ${function:Show-InputDialog}
        return (& $f $FormTitle $FormText $DefaultValue)
    }

    [void] ShowAboutDialog() {
        $f = ${function:Show-AboutDialog}
        & $f
    }

    [bool] RequestUIConfirmation([string]$Message, [string]$Caption) {
        $f = ${function:Request-UIConfirmation}
        return [bool](& $f $Message $Caption)
    }

    [void] ShowPopup([object]$Popup) {
        $f = ${function:Show-Popup}
        & $f $Popup
    }

    [void] HidePopup() {
        $f = ${function:Hide-Popup}
        & $f
    }

    #endregion

    #region File / folder pickers

    [string] ShowFolderPicker([string]$Description) {
        $f = ${function:Get-Folder}
        return (& $f -Title $Description)
    }

    [string] ShowFolderPicker([string]$Path, [string]$Description) {
        $f = ${function:Get-Folder}
        return (& $f -Path $Path -Title $Description)
    }

    #endregion

    #region XAML helpers

    [object] GetXamlObject([string]$FileName) {
        $f = ${function:Get-XamlObject}
        return (& $f $FileName)
    }

    [object] GetXamlObject([string]$FileName, [bool]$AddVariables) {
        $f = ${function:Get-XamlObject}
        if ($AddVariables) { return (& $f $FileName -AddVariables) }
        else               { return (& $f $FileName) }
    }

    [object] GetXamlObject([string]$FileName, [bool]$AddVariables, [bool]$AddStyles) {
        $f = ${function:Get-XamlObject}
        $h = @{}
        if ($AddVariables) { $h.AddVariables = $true }
        if ($AddStyles)    { $h.AddStyles = $true }
        return (& $f $FileName @h)
    }

    [void] AddXamlEvent([object]$XamlObj, [string]$ControlName, [string]$EventName, [scriptblock]$Handler) {
        $f = ${function:Add-XamlEvent}
        & $f $XamlObj $ControlName $EventName $Handler
    }

    [void] AddXamlVariables([object]$XamlObj, [object]$Scope) {
        $f = ${function:Add-XamlVariables}
        & $f $XamlObj $Scope
    }

    [void] SetXamlProperty([object]$XamlObj, [string]$ControlName, [string]$PropertyName, [object]$Value) {
        $f = ${function:Set-XamlProperty}
        & $f $XamlObj $ControlName $PropertyName $Value
    }

    [object] GetXamlProperty([object]$XamlObj, [string]$ControlName, [string]$PropertyName) {
        $f = ${function:Get-XamlProperty}
        return (& $f $XamlObj $ControlName $PropertyName)
    }

    [void] SetControlVisible([object]$XamlObj, [string]$ControlName, [bool]$Visible) {
        $value = if ($Visible) { "Visible" } else { "Collapsed" }
        $this.SetXamlProperty($XamlObj, $ControlName, "Visibility", $value)
    }

    #endregion

    #region Main window / chrome

    [void] ShowMainWindow([object]$View) {
        $f = ${function:Show-MainWindow}
        & $f $View
    }

    [void] SetMainTitle() {
        $f = ${function:Set-MainTitle}
        & $f
    }

    [object] GetMainWindow() {
        $f = ${function:Get-MainWindow}
        return (& $f)
    }

    [void] SetAppTheme([string]$ThemeName) {
        $f = ${function:Set-AppTheme}
        & $f $ThemeName
    }

    #endregion

    #region Threading / status

    [void] UpdateUIStatus([string]$Text) {
        $f = ${function:Update-UIStatus}
        & $f -Text $Text | Out-Null
    }

    [bool] UpdateUIStatus([System.Collections.IDictionary]$Params) {
        $f = ${function:Update-UIStatus}
        $h = @{}
        foreach ($k in $Params.Keys) { $h[$k] = $Params[$k] }
        return [bool](& $f @h)
    }

    [void] InvokeUIMessagePump() {
        $f = ${function:Invoke-UIMessagePump}
        & $f
    }

    [void] SetClipboardText([string]$Text) {
        # Clipboard.SetText throws on null/empty; Clear() is the empty equivalent.
        if([string]::IsNullOrEmpty($Text)) { [System.Windows.Clipboard]::Clear(); return }
        [System.Windows.Clipboard]::SetText($Text)
    }

    #endregion

    #region Auth chrome

    [void] ShowAuthenticationInfo() {
        $f = ${function:Show-AuthenticationInfo}
        & $f
    }

    [void] SetEnvironmentInfo() {
        $f = ${function:Set-EnvironmentInfo}
        & $f
    }

    [void] SetEnvironmentInfo([string]$TenantName) {
        $f = ${function:Set-EnvironmentInfo}
        & $f $TenantName
    }

    #endregion
}
