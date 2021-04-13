[CmdletBinding(SupportsShouldProcess=$True)]
param(
    [switch]
    $ShowConsoleWindow
)
Import-Module ($PSScriptRoot + "\CloudAPIPowerShellManagement.psd1") -Force
Initialize-CloudAPIManagement -View "IntuneGraphAPI" -ShowConsoleWindow:($ShowConsoleWindow)