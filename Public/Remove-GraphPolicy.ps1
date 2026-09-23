function Remove-GraphPolicy {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    [OutputType([IntunePolicyBase[]])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyBase[]]
        $InputObject,

        # Specifies destination environment. Default is current logged on environment.
        [int]
        $TokenId = 0
    )

    Begin {
        Write-LogDebug "Start deleting policies"
        $deleted = @()
        $bulkDelete = $null
        if (Test-GraphBatchEnabled) {
            $bulkDelete = @{}
        }
    }

    Process {
        foreach ($policyObject in $InputObject) {
            if ($PSCmdlet.ShouldProcess($policyObject.Name, 'DELETE')) {
                if ($policyObject.Delete($bulkDelete) -and -not $bulkDelete) {
                    $deleted += $policyObject
                }
            }
        }
    }

    End {        
        if($bulkDelete.Count -gt 0) {
            $batchObjects = [System.Collections.Generic.List[PSCustomObject]]::new()
            $bulkDelete.Keys | ForEach-Object { $batchObjects.Add($_) }
            $batchResults = Invoke-GraphBatchRequest -BatchObject $batchObjects -TokenId $TokenId -BatchType "Delete"
            $batchResults | ForEach-Object {
                $result = $_
                $policy = $bulkDelete.Values | Where-Object Id -eq $_.Id
                if ($result.Status -ge 200 -and $result.Status -lt 300) {
                    $deleted += $policy
                    Write-Log "Policy $($policy.Name) ($($policy.Id)) deleted successfully"
                }
                else {
                    Write-LogError "Failed to delete policy $($policy.Name) ($($policy.Id)). Status code: $($result.Status) $($result.ErrorMessage)"
                }
            }
        }            
        
        Write-LogDebug "Delete policy finished"
        return $deleted
    }
}
