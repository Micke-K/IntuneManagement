[CmdletBinding(SupportsShouldProcess=$True)]
param(
    [switch]
    $ShowConsoleWindow,
    [switch]
    $JSonSettings,
    [string]
    $JSonFile
)
Import-Module ($PSScriptRoot + "\CloudAPIPowerShellManagement.psd1") -Force
Initialize-CloudAPIManagement -View "IntuneGraphAPI" -ShowConsoleWindow:($ShowConsoleWindow) -JSonSettings:($JSonSettings) -JSonFile $JSonFile
