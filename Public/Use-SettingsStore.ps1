<#
.SYNOPSIS
    Choose where settings are read from and written to for the rest of the session.

.DESCRIPTION
    Exported as Use-IMSettingsStore.

    Three stores:

      -Memory        settings live in this session only. Reads and writes work
                     normally and nothing is read from or written to disk. This is
                     what automation on a shared machine wants: a runbook on a
                     Hybrid Worker must not rewrite that worker's configuration, and
                     must not inherit whatever the last run left behind.
      -Path <file>   a JSON settings file. Created if it does not exist.
      -Registry      HKCU:\Software\IntuneManagement. Windows only.

    The default without calling this is the registry on Windows and a settings file
    under LocalApplicationData elsewhere. It can also be set before the module is
    imported, which is the only way to keep the very first log lines from resolving
    against the machine's store:

      $env:IM_SETTINGS_STORE = 'Memory'
      $env:IM_SETTINGS_FILE  = 'C:\config\intune-settings.json'

.PARAMETER Memory
    Use an in-memory store. Nothing is written to disk.

.PARAMETER Path
    Use this JSON settings file.

.PARAMETER Registry
    Use the Windows registry.

.PARAMETER Seed
    Copy the machine's currently persisted settings into the in-memory store first,
    so the session starts from the real configuration and then diverges without
    writing back. Without it the store starts empty and every setting resolves to
    its registered default.

.PARAMETER PassThru
    Return the resulting store, as Get-IMSettingsStore would.

.EXAMPLE
    Use-IMSettingsStore -Memory -PassThru

.EXAMPLE
    Use-IMSettingsStore -Memory
    Import-IMSettingsStore -Path .\runbook-settings.json

    The full automation pattern: an empty store, then the configuration the run is
    supposed to use, from a file under source control. Nothing on the worker is read
    or changed.

.EXAMPLE
    Use-IMSettingsStore -Path 'D:\shared\IntuneManagement.json'

.LINK
    Get-IMSettingsStore
.LINK
    Import-IMSettingsStore
.LINK
    Export-IMSettingsStore
#>
function Use-SettingsStore
{
    [CmdletBinding(DefaultParameterSetName = "Memory", SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = "Memory")]
        [switch]$Memory,
        [Parameter(Mandatory = $true, ParameterSetName = "Json", Position = 0)]
        [string]$Path,
        [Parameter(Mandatory = $true, ParameterSetName = "Registry")]
        [switch]$Registry,
        [switch]$Seed,
        [switch]$PassThru
    )

    $mode = $PSCmdlet.ParameterSetName

    if(-not $PSCmdlet.ShouldProcess("the settings store", "Switch to the $mode store")) { return }

    Set-SettingsStoreMode -Mode $mode -Path $Path -Seed:$Seed

    if($PassThru) { Get-SettingsStoreInfo }
}
