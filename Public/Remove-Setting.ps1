<#
.SYNOPSIS
    Remove a stored IntuneManagement setting so it falls back to the next level.

.DESCRIPTION
    Exported as Remove-IMSetting.

    Removing a value is not the same as setting it to nothing: a removed tenant value
    reverts to the global value, and a removed global value reverts to the default
    the setting was registered with. That is how the settings dialog's per-tenant
    checkbox works, and this is the same operation.

    Like Set-IMSetting this changes the persistent store unless the session is using
    an in-memory one.

.PARAMETER Key
    The setting key.

.PARAMETER Scope
    Global (default) removes the value every tenant sees, so the setting falls back
    to its registered default. Tenant removes the tenant-specific value only,
    leaving the global one in place.

.PARAMETER TenantID
    The tenant to remove for. Defaults to the connected tenant.

.PARAMETER SubPath
    Storage path for a key that is not a registered setting. Required in that case,
    reported and ignored for a registered one (its registration decides the path).

.EXAMPLE
    Remove-IMSetting ExportFolder -Scope Tenant

    Stop overriding the export folder for this tenant; the global value applies again.

.EXAMPLE
    Get-IMSetting UseBatchAPI -Detailed
    Remove-IMSetting UseBatchAPI
    Get-IMSetting UseBatchAPI -Detailed

    Source goes from Global back to Default.

.LINK
    Set-IMSetting
.LINK
    Get-IMSetting
#>
function Remove-Setting
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Key,
        [ValidateSet("Global", "Tenant")]
        [string]$Scope = "Global",
        [string]$TenantID,
        $SubPath = $null
    )

    $store = Get-SettingsStoreInfo
    $target = if($store.Path) { "$($store.Mode) store at $($store.Path)" } else { "$($store.Mode) store" }
    $scopeText = if($Scope -eq "Tenant") { " for tenant scope" } else { "" }

    if(-not $PSCmdlet.ShouldProcess($target, "Remove setting '$Key'$scopeText")) { return }

    Write-Log "Remove setting '$Key' from the $target$scopeText"

    Remove-SettingValue -Key $Key -Tenant:($Scope -eq "Tenant") -TenantID $TenantID -SubPath $SubPath
}
