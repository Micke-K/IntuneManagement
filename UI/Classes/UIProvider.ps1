#ImportOrder 30
#
# Abstract base class for the per-backend UI provider. Concrete subclasses
# live in UI/WPF/Classes/WPFUIProvider.ps1 and UI/Avalonia/Classes/AvaloniaUIProvider.ps1.
# The loader instantiates exactly one of them into $script:UIProvider after
# all classes are loaded but before AppInitialized fires (see
# IntuneManagement.psm1).
#
# Methods are virtual-by-throw: every method on this base raises
# "<name> not implemented by <typename>" so a missing override surfaces at
# the call site rather than silently no-op'ing. Subclasses override every
# method they support; cross-boundary callers (Internal/Public) only invoke
# methods that both backends implement.
#
# Parameter types are deliberately loose ([object]) for surfaces where the
# concrete shape differs between backends — e.g. ShowModalForm receives a
# WPF FrameworkElement under WPF and an Avalonia Control under Avalonia.

class UIProvider {

    #region Modal / dialog

    [void] ShowMessageBox([string]$Text) {
        throw "ShowMessageBox(text) not implemented by $($this.GetType().Name)"
    }

    [void] ShowMessageBox([string]$Text, [string]$Caption) {
        throw "ShowMessageBox not implemented by $($this.GetType().Name)"
    }

    [object] ShowMessageBox([string]$Text, [string]$Caption, [string]$Button, [string]$Icon) {
        throw "ShowMessageBox(button,icon) not implemented by $($this.GetType().Name)"
    }

    [void] ShowModalForm([string]$FormTitle, [object]$FormObject) {
        throw "ShowModalForm not implemented by $($this.GetType().Name)"
    }

    [void] ShowModalForm([string]$FormTitle, [object]$FormObject, [bool]$HideButtons) {
        throw "ShowModalForm(hideButtons) not implemented by $($this.GetType().Name)"
    }

    [void] ShowModalObject() {
        throw "ShowModalObject() not implemented by $($this.GetType().Name)"
    }

    [void] ShowModalObject([object]$Obj) {
        throw "ShowModalObject not implemented by $($this.GetType().Name)"
    }

    [void] CloseTopModalObject() {
        throw "CloseTopModalObject not implemented by $($this.GetType().Name)"
    }

    [string] ShowInputDialog([string]$FormTitle, [string]$FormText, [string]$DefaultValue) {
        throw "ShowInputDialog not implemented by $($this.GetType().Name)"
    }

    [void] ShowAboutDialog() {
        throw "ShowAboutDialog not implemented by $($this.GetType().Name)"
    }

    [bool] RequestUIConfirmation([string]$Message, [string]$Caption) {
        throw "RequestUIConfirmation not implemented by $($this.GetType().Name)"
    }

    [void] ShowPopup([object]$Popup) {
        throw "ShowPopup not implemented by $($this.GetType().Name)"
    }

    [void] HidePopup() {
        throw "HidePopup not implemented by $($this.GetType().Name)"
    }

    #endregion

    #region File / folder pickers

    [string] ShowFolderPicker([string]$Description) {
        throw "ShowFolderPicker not implemented by $($this.GetType().Name)"
    }

    [string] ShowFolderPicker([string]$Path, [string]$Description) {
        throw "ShowFolderPicker(path,description) not implemented by $($this.GetType().Name)"
    }

    #endregion

    #region XAML helpers

    [object] GetXamlObject([string]$FileName) {
        throw "GetXamlObject not implemented by $($this.GetType().Name)"
    }

    [object] GetXamlObject([string]$FileName, [bool]$AddVariables) {
        throw "GetXamlObject(addVariables) not implemented by $($this.GetType().Name)"
    }

    [object] GetXamlObject([string]$FileName, [bool]$AddVariables, [bool]$AddStyles) {
        throw "GetXamlObject(addVariables,addStyles) not implemented by $($this.GetType().Name)"
    }

    [void] AddXamlEvent([object]$XamlObj, [string]$ControlName, [string]$EventName, [scriptblock]$Handler) {
        throw "AddXamlEvent not implemented by $($this.GetType().Name)"
    }

    [void] AddXamlVariables([object]$XamlObj, [object]$Scope) {
        throw "AddXamlVariables not implemented by $($this.GetType().Name)"
    }

    [void] SetXamlProperty([object]$XamlObj, [string]$ControlName, [string]$PropertyName, [object]$Value) {
        throw "SetXamlProperty not implemented by $($this.GetType().Name)"
    }

    [object] GetXamlProperty([object]$XamlObj, [string]$ControlName, [string]$PropertyName) {
        throw "GetXamlProperty not implemented by $($this.GetType().Name)"
    }

    # Show or hide a named control. $Visible is the LOGICAL intent (shown/hidden);
    # each backend maps it to its own framework property: WPF sets Visibility to
    # "Visible"/"Collapsed", Avalonia sets the IsVisible bool. Callers must use this
    # method (not the underlying property) so code stays backend-agnostic.
    [void] SetControlVisible([object]$XamlObj, [string]$ControlName, [bool]$Visible) {
        throw "SetControlVisible not implemented by $($this.GetType().Name)"
    }

    #endregion

    #region Main window / chrome

    [void] ShowMainWindow([object]$View) {
        throw "ShowMainWindow not implemented by $($this.GetType().Name)"
    }

    [void] SetMainTitle() {
        throw "SetMainTitle not implemented by $($this.GetType().Name)"
    }

    [object] GetMainWindow() {
        throw "GetMainWindow not implemented by $($this.GetType().Name)"
    }

    [void] SetAppTheme([string]$ThemeName) {
        throw "SetAppTheme not implemented by $($this.GetType().Name)"
    }

    #endregion

    #region Threading / status

    [void] UpdateUIStatus([string]$Text) {
        throw "UpdateUIStatus not implemented by $($this.GetType().Name)"
    }

    [bool] UpdateUIStatus([System.Collections.IDictionary]$Params) {
        throw "UpdateUIStatus(IDictionary) not implemented by $($this.GetType().Name)"
    }

    [void] InvokeUIMessagePump() {
        throw "InvokeUIMessagePump not implemented by $($this.GetType().Name)"
    }

    # Copy text to the system clipboard. Backends implement this over their own
    # framework clipboard (WPF System.Windows.Clipboard / Avalonia TopLevel.Clipboard)
    # rather than PowerShell's Set-Clipboard, which is a silent no-op on a bare
    # Linux box without xclip/xsel/wl-copy.
    [void] SetClipboardText([string]$Text) {
        throw "SetClipboardText not implemented by $($this.GetType().Name)"
    }

    #endregion

    #region Auth chrome

    [void] ShowAuthenticationInfo() {
        throw "ShowAuthenticationInfo not implemented by $($this.GetType().Name)"
    }

    [void] SetEnvironmentInfo() {
        throw "SetEnvironmentInfo() not implemented by $($this.GetType().Name)"
    }

    [void] SetEnvironmentInfo([string]$TenantName) {
        throw "SetEnvironmentInfo not implemented by $($this.GetType().Name)"
    }

    #endregion
}
