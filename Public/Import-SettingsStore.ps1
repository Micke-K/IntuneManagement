<#
.SYNOPSIS
    Load a JSON settings file into the active settings store.

.DESCRIPTION
    Exported as Import-IMSettingsStore.

    Values are written one at a time through the normal write path, so the import
    behaves like the store it is going into: values persist when the store is a file
    or the registry, and stay in memory when it is not.

    The file MERGES into what is already there rather than replacing it. For a clean
    slate, start from an empty in-memory store:

        Use-IMSettingsStore -Memory        # no -Seed, so nothing is inherited
        Import-IMSettingsStore -Path .\runbook-settings.json

    Keys in the file that are not registered settings are imported and reported -
    some are legitimate (hidden keys, per-feature state) and a typo looks the same.

.PARAMETER Path
    The JSON settings file to load, in the shape Export-IMSettingsStore writes.

.EXAMPLE
    Use-IMSettingsStore -Memory
    Import-IMSettingsStore -Path .\runbook-settings.json
    Get-IMSettingsStore -IncludeValues

.EXAMPLE
    Import-IMSettingsStore -Path .\baseline.json -WhatIf

    Check what store the import would land in before doing it.

.LINK
    Export-IMSettingsStore
.LINK
    Use-IMSettingsStore
#>
function Import-SettingsStore
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path
    )

    $store = Get-SettingsStoreInfo
    $target = if($store.Path) { "$($store.Mode) store at $($store.Path)" } else { "$($store.Mode) store" }

    if(-not $PSCmdlet.ShouldProcess($target, "Import settings from $Path")) { return }

    Import-SettingsStoreFromFile -Path $Path | Out-Null
}
