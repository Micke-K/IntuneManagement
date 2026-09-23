<#
.SYNOPSIS
    Read an IntuneManagement setting by key.

.DESCRIPTION
    Exported as Get-IMSetting.

    The public read side of the settings system. A key is all that is needed - the
    storage path is taken from the setting's registration, so automation never has
    to know that "GraphPageSize" lives under "IntuneManager" or that the store is a
    registry key on Windows and a JSON file everywhere else.

    Resolution order is the same one the application itself uses: the value for the
    connected tenant, then the global value, then the value the setting was
    registered with. -Scope pins it to one level.

    With no -Key, every registered setting is returned.

.PARAMETER Key
    The setting key. Accepts pipeline input. Omit to return all registered settings.

.PARAMETER Scope
    Effective (default) resolves tenant, then global, then the registered default.
    Global reads the global value only. Tenant reads the tenant value only.

    A scoped read returns $null when nothing is stored at that scope - it does NOT
    fall back to the registered default, so "not overridden here" stays
    distinguishable from "overridden to the same value as the default". With
    -Detailed the Source of such a read is 'NotSet'. Tenant scope needs a tenant:
    it reports an error rather than a value when none is connected and no -TenantID
    is given.

.PARAMETER TenantID
    Which tenant to resolve against. Defaults to the connected tenant.

.PARAMETER SubPath
    Storage path for a key that is not a registered setting (the hidden keys such as
    ExportReplaceTokens, and per-feature state). Required in that case - an
    unregistered key has no registration to take a path from. Reported and ignored
    for a registered one, exactly as on Set-IMSetting.

    An unregistered key has no registered default, so its Source is Tenant, Global
    or NotSet - never Default - and no type normalization is applied: the value
    comes back as the string the store holds.

.PARAMETER Detailed
    Return the value together with where it came from (Tenant, Global, Default or
    NotSet), its type, its storage path and its registered default, instead of just
    the value.

.EXAMPLE
    Get-IMSetting ExportFolder

.EXAMPLE
    Get-IMSetting UseBatchAPI -Detailed

    Shows whether the value in effect was set for this tenant, set globally, or is
    just the default - which is what to check before changing it.

.EXAMPLE
    'AddCompanyName','AddObjectType' | Get-IMSetting

.EXAMPLE
    Get-IMSetting | Where-Object Source -ne 'Default'

    Every setting that has actually been configured.

.EXAMPLE
    Set-IMSetting ExportReplaceTokens 'TenantId' -SubPath 'IntuneManager'
    Get-IMSetting ExportReplaceTokens -SubPath 'IntuneManager'

    Reading back a hidden key needs the same path the write used.

.LINK
    Set-IMSetting
.LINK
    Get-IMSettingDefinition
.LINK
    Get-IMSettingsStore
#>
function Get-Setting
{
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$Key,
        [ValidateSet("Effective", "Global", "Tenant")]
        [string]$Scope = "Effective",
        [string]$TenantID,
        # $null, not "": "" is the root of the store, so "not supplied" and "the
        # root" have to stay distinguishable. Same convention as Set-IMSetting.
        $SubPath = $null,
        [switch]$Detailed
    )

    process
    {
        # No key = every registered setting. A bare list of values would be
        # meaningless without the keys beside them, so that form is always detailed.
        $keys = @($Key)
        if(-not $Key)
        {
            $keys = @(Get-SettingsSections | ForEach-Object { $_.Values } | ForEach-Object Key)
            $Detailed = $true

            # Enumerating the registered settings takes every path from its own
            # registration; there is nothing for a single -SubPath to apply to.
            if($null -ne $SubPath) { Write-Log "Ignoring -SubPath '$SubPath': it applies to a single unregistered -Key, not to a listing of registered settings" 2 }
            $SubPath = $null
        }

        foreach($settingKey in $keys)
        {
            $resolved = Resolve-SettingValue -Key $settingKey -TenantID $TenantID -SubPath $SubPath `
                            -GlobalOnly:($Scope -eq "Global") -TenantOnly:($Scope -eq "Tenant")

            if(-not $resolved) { continue }

            if($Detailed) { $resolved }
            else { $resolved.Value }
        }
    }
}
