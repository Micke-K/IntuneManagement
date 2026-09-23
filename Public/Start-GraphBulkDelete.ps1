function Start-GraphBulkDelete {
    <#
    .SYNOPSIS
        Bulk-delete Intune policy objects.

    .DESCRIPTION
        Public, UI-independent driver for bulk delete. The WPF and Avalonia
        bulk-delete forms are thin callers of this function (they own the
        confirmation prompt); the same function can be invoked from a
        scheduled task or pipeline.

        DESTRUCTIVE: every object of the selected groups (matching the
        optional name filter) is deleted from the signed-in tenant. There is
        no confirmation in this driver - callers confirm before invoking.

        Groups are processed in DESCENDING ImportOrder so dependents are
        deleted before their dependencies.

    .PARAMETER Filter
        Name filter (literal substring, case-insensitive). Empty = delete
        every object of the selected groups.

    .PARAMETER PolicyGroup
        Restrict the delete to these PolicyGroup IDs. Mandatory - bulk
        deleting every group implicitly is too dangerous for a default.

    .PARAMETER TokenId
        Authentication token id. Defaults to the current default token.

    .EXAMPLE
        Start-GraphBulkDelete -PolicyGroup DeviceConfiguration -Filter "[Test]"

    .OUTPUTS
        PSCustomObject with summary statistics (Groups, Deleted, Duration).
    #>
    [CmdletBinding()]
    param(
        [string]
        $Filter,

        [Parameter(Mandatory = $true)]
        [string[]]
        $PolicyGroup,

        [Int]
        $TokenId = (Get-DefaultTokenId)
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $selection = Resolve-IntuneTargetSelectors -PolicyGroup $PolicyGroup -Caller 'Start-GraphBulkDelete'
    $unknownSelectors = $selection.Unknown   # ids from -PolicyGroup that matched nothing
    $targetGroups = [System.Collections.Generic.List[object]]::new()
    $targetGroups.AddRange([object[]]$selection.Groups)
    if ($targetGroups.Count -eq 0) {
        Write-Log "Bulk delete: no policy groups selected" 2
        return [PSCustomObject]@{ Groups = 0; Deleted = 0; UnknownSelectors = $unknownSelectors; Duration = $stopwatch.Elapsed }
    }

    Write-Log "****************************************************************"
    Write-Log "Start bulk delete"
    Write-Log "****************************************************************"

    $totalDeleted = 0
    # Reverse import-order so dependents are deleted before their dependencies.
    $orderedGroups = @($targetGroups | Sort-Object { ($_.PolicyTypes | Measure-Object ImportOrder -Minimum).Minimum } -Descending)
    $groupTotal    = $orderedGroups.Count
    $groupIndex    = 0

    foreach ($grp in $orderedGroups) {
        $groupIndex++
        Write-Log "----------------------------------------------------------------"
        Write-Log "Delete $($grp.Title) objects"
        Write-Log "----------------------------------------------------------------"

        try {
            Write-Status `
                -Text   ("Bulk delete - {0} ({1} of {2})" -f $grp.Title, $groupIndex, $groupTotal) `
                -Detail "Listing policies" `
                -Force
            $policies = @(Get-GraphPolicies -PolicyGroup $grp.ID -TokenId $TokenId)

            if ($Filter) {
                $policies = @($policies | Where-Object { $_.Name -match [RegEx]::Escape($Filter) })
            }

            if ($policies.Count -eq 0) {
                Write-Log "No $($grp.Title) objects found"
                continue
            }

            Write-Log "Deleting $($policies.Count) $($grp.Title) object(s)"
            Write-Status -Detail ("Deleting {0} object(s)" -f $policies.Count) -SkipLog -Force
            $policies | Remove-GraphPolicy -Confirm:$false | Out-Null
            $totalDeleted += $policies.Count
        }
        catch {
            Write-LogError "Failed when deleting $($grp.Title) objects" $_.Exception
        }
    }

    Write-Status $null
    Write-Log "****************************************************************"
    Write-Log "Bulk delete finished"
    Write-Log "****************************************************************"

    return [PSCustomObject]@{
        Groups           = $orderedGroups.Count
        Deleted          = $totalDeleted
        UnknownSelectors = $unknownSelectors
        Duration         = $stopwatch.Elapsed
    }
}
