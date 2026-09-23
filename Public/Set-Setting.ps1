<#
.SYNOPSIS
    Write an IntuneManagement setting by key.

.DESCRIPTION
    Exported as Set-IMSetting.

    A key and a value is all that is needed. The storage path comes from the
    setting's registration, so the value always lands where Get-IMSetting (and the
    application) will look for it - which hand-written paths repeatedly did not.

    THIS WRITES TO THE PERSISTENT STORE by default: the registry on Windows, the
    settings file elsewhere. That is deliberate - configuring the tool from a script
    should not need an extra opt-in switch - but it means a script run on a shared
    machine changes that machine's configuration. Two ways to avoid it:

      Use-IMSettingsStore -Memory     nothing is written to disk for this session
      Get-IMSettingsStore             assert which store is active before writing

    Every write is logged with the store and the path it went to.

    Values are stored in exactly the shape the settings dialog stores them in, so a
    value written here is indistinguishable from one set in the UI.

.PARAMETER Key
    The setting key. Use Get-IMSettingDefinition to discover the available keys.

.PARAMETER Value
    The value to store. $null removes the value (see also Remove-IMSetting).
    Booleans accept $true/$false or 'true'/'false' in any case.

.PARAMETER Scope
    Global (default) writes the value every tenant sees. Tenant writes it for one
    tenant only, which takes precedence over the global value when that tenant is
    connected; it fails rather than quietly falling back to a global write when no
    tenant id can be resolved.

.PARAMETER TenantID
    The tenant to write for. Defaults to the connected tenant. Implies nothing on
    its own - pass -Scope Tenant to select the tenant scope.

.PARAMETER SubPath
    Storage path for a key that is not a registered setting (the hidden keys, and
    per-feature state). Required in that case. For a registered setting it is
    reported and ignored - the registration decides where the value is stored, so a
    write cannot be aimed at a path the application never reads.

.PARAMETER PassThru
    Return the resolved setting after writing it.

.EXAMPLE
    Set-IMSetting ExportFolder 'C:\Intune\Export'

.EXAMPLE
    Use-IMSettingsStore -Memory -Seed
    Set-IMSetting UseBatchAPI $true
    Set-IMSetting GraphPageSize 500

    Configure a run without touching the machine's stored settings - the pattern for
    a runbook on a shared worker.

.EXAMPLE
    Set-IMSetting ExportFolder '\\server\intune\contoso' -Scope Tenant

    Set the export folder for the connected tenant only.

.EXAMPLE
    Set-IMSetting ExportReplaceTokens 'TenantId' -SubPath 'IntuneManager'

    A key with no entry in the settings dialog needs its path spelled out.

.LINK
    Get-IMSetting
.LINK
    Remove-IMSetting
.LINK
    Use-IMSettingsStore
#>
function Set-Setting
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Key,
        [Parameter(Mandatory = $true, Position = 1)]
        [AllowNull()]
        [AllowEmptyString()]
        $Value,
        # Same vocabulary as Get-IMSetting -Scope, minus Effective: there is no such
        # thing as writing the effective value.
        [ValidateSet("Global", "Tenant")]
        [string]$Scope = "Global",
        [string]$TenantID,
        $SubPath = $null,
        [switch]$PassThru
    )

    $store = Get-SettingsStoreInfo
    $target = if($store.Path) { "$($store.Mode) store at $($store.Path)" } else { "$($store.Mode) store" }

    $scopeText = if($Scope -eq "Tenant") { " for tenant scope" } else { "" }

    if(-not $PSCmdlet.ShouldProcess($target, "Set setting '$Key' to '$Value'$scopeText")) { return }

    # Logged, not just debug-logged: a write that persisted to a machine's registry
    # or settings file should be visible in the log afterwards.
    Write-Log "Set setting '$Key' in the $target$scopeText"

    Set-SettingValue -Key $Key -Value $Value -Tenant:($Scope -eq "Tenant") -TenantID $TenantID -SubPath $SubPath -PassThru:$PassThru
}
