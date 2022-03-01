[CmdletBinding(SupportsShouldProcess=$True)]
param(
    [switch]
    $ShowConsoleWindow,
    [switch]
    $JSonSettings,
    [string]
    $JSonFile,
    [switch]
    $Silent,
    [string]
    $SilentBatchFile = "",
    [string]
    $TenantId,
    [string]
    $AppId,
    [string]
    $Secret,
    [string]
    $Certificate
)
Import-Module ($PSScriptRoot + "\CloudAPIPowerShellManagement.psd1") -Force
$param = $PSBoundParameters
Initialize-CloudAPIManagement -View "IntuneGraphAPI" @param
