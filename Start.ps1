param([switch]$ShowUI)

# Linux: xrdp/XFCE sessions often leave XDG_RUNTIME_DIR (and sometimes
# DBUS_SESSION_BUS_ADDRESS) unset. Without them the Avalonia UI can't reach
# xdg-desktop-portal, so file dialogs fall back to the oversized managed picker
# and "open in browser" fails. Point them at the standard runtime dir / bus when
# present, before the module (and Avalonia) initialize.
if (-not $IsWindows) {
    try {
        $uid = (& id -u).Trim()
        if ([string]::IsNullOrEmpty($env:XDG_RUNTIME_DIR) -and (Test-Path "/run/user/$uid")) {
            $env:XDG_RUNTIME_DIR = "/run/user/$uid"
        }
        if ([string]::IsNullOrEmpty($env:DBUS_SESSION_BUS_ADDRESS) -and $env:XDG_RUNTIME_DIR -and (Test-Path "$($env:XDG_RUNTIME_DIR)/bus")) {
            $env:DBUS_SESSION_BUS_ADDRESS = "unix:path=$($env:XDG_RUNTIME_DIR)/bus"
        }
    }
    catch { }
}

Import-Module ($PSScriptRoot + "\IntuneManagement.psd1") -Force

if($ShowUI -eq $true) {
    Show-IMMainWindow -View "IntuneManagement"
}