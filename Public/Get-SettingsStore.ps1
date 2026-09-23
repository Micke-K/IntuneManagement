<#
.SYNOPSIS
    Report which settings store is active, and whether it persists.

.DESCRIPTION
    Exported as Get-IMSettingsStore.

    Set-IMSetting writes to the persistent store by default, so a script that must
    not change the machine it runs on can assert on this first:

        if((Get-IMSettingsStore).Persisted) { throw 'Refusing to write to a persistent store' }

    Mode is the store in effect, which is not always the store that was asked for:
    a settings file that cannot be read falls back, and off Windows there is no
    registry to fall back to. RequestedMode is what was asked for.

.PARAMETER IncludeValues
    Also return every value currently in the store, as SubPath/Key/Value rows. This
    reads the whole store, including a recursive registry walk in registry mode.

.EXAMPLE
    Get-IMSettingsStore

.EXAMPLE
    (Get-IMSettingsStore -IncludeValues).Values | Format-Table

.EXAMPLE
    Get-IMSettingsStore | Select-Object Mode, Persisted, Path

.LINK
    Use-IMSettingsStore
.LINK
    Export-IMSettingsStore
#>
function Get-SettingsStore
{
    [CmdletBinding()]
    param([switch]$IncludeValues)

    $info = Get-SettingsStoreInfo

    if($IncludeValues)
    {
        $info | Add-Member -MemberType NoteProperty -Name "Values" -Value @(Get-SettingsStoreEntries) -PassThru
    }
    else
    {
        $info
    }
}
