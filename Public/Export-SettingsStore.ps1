<#
.SYNOPSIS
    Write the whole active settings store to a JSON file.

.DESCRIPTION
    Exported as Export-IMSettingsStore.

    Works whatever the store is: a registry store is walked and written out as the
    same nested JSON a settings file uses, so a machine that has been configured
    through the UI can hand its configuration to a file that automation loads with
    Import-IMSettingsStore.

    The file includes every value in the store, tenant-specific values (nested under
    the tenant id) and unregistered keys alike.

.PARAMETER Path
    The file to write. Its folder is created if needed. An existing file is replaced.

.EXAMPLE
    Export-IMSettingsStore -Path .\intune-settings.json

.EXAMPLE
    Export-IMSettingsStore -Path \\share\config\prod.json

    Capture a working configuration once, commit it, and have every run import it
    instead of relying on whatever is on the machine.

.LINK
    Import-IMSettingsStore
.LINK
    Get-IMSettingsStore
#>
function Export-SettingsStore
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path
    )

    if(-not $PSCmdlet.ShouldProcess($Path, "Export the settings store")) { return }

    Export-SettingsStoreToFile -Path $Path | Out-Null
}
