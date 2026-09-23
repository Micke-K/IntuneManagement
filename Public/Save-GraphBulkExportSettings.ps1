function Save-GraphBulkExportSettings {
    <#
    .SYNOPSIS
        Save bulk-export form state to a JSON file usable by Start-GraphBulkExport.
    .DESCRIPTION
        Round-trips an [IntuneManagerExportSettings] instance plus the selected
        PolicyGroup / PolicyType IDs to disk. Loaded back via
        Start-GraphBulkExport -SettingsFile <path>.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [IntuneManagerExportSettings]
        $ExportSettings,

        [string[]]$PolicyGroup,
        [string[]]$PolicyType
    )

    $payload = [ordered]@{
        ExportFolder            = $ExportSettings.ExportFolder
        Filter                  = $ExportSettings.Filter
        ExportAssignments       = $ExportSettings.ExportAssignments
        AddCompanyName          = $ExportSettings.AddCompanyName
        AddObjectType           = $ExportSettings.AddObjectType
        ExportNestedGroupLevels = $ExportSettings.ExportNestedGroupLevels
        PolicyGroup             = @($PolicyGroup)
        PolicyType              = @($PolicyType)
    }

    $dir = [IO.Path]::GetDirectoryName($Path)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    ConvertTo-Json $payload -Depth 5 | Out-File -LiteralPath $Path -Encoding utf8 -Force
    Write-Log "Bulk-export settings saved to $Path"
}
