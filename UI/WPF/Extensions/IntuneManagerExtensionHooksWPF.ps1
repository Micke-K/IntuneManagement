# Add-IntuneManagerExportUIExtensions / Add-IntuneManagerImportUIExtensions iterators.
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Add-IntuneManagerExportUIExtensions
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyTypeBase[]]
        $InputObject,
        [Parameter(Mandatory = $true)]
        [Object]
        $Form
    )

    Begin { Write-Log "Add Export UI Extensions - Start" }

    Process {
        foreach($policyType in $InputObject) {
            if($policyType.AddUIExportExtensions) {
                $policyType.AddUIExportExtensions($Form)
            }
        }
    }

    End { Write-Log "Add Export UI Extensions - Done" }
}

function Add-IntuneManagerImportUIExtensions
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyTypeBase[]]
        $InputObject,
        [Parameter(Mandatory = $true)]
        [Object]
        $Form
    )

    Begin { Write-Log "Add Import UI Extensions - Start" }

    Process {
        foreach($policyType in $InputObject) {
            if($policyType.AddUIImportExtensions) {
                $policyType.AddUIImportExtensions($Form)
            }
        }
    }

    End { Write-Log "Add Import UI Extensions - Done" }
}

