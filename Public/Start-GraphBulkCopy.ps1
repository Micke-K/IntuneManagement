function Start-GraphBulkCopy {
    <#
    .SYNOPSIS
        Bulk-copy Intune policy objects by name pattern.

    .DESCRIPTION
        Public, UI-independent driver for bulk copy (ported from the original
        project's Copy extension). For every selected policy type, each object
        whose name contains CopyFromPattern is copied with the pattern replaced
        by CopyToPattern (e.g. "Test - Baseline" -> "Prod - Baseline").

        Objects are skipped when an object of the same @odata.type already has
        the target name, so the operation is safe to re-run.

        The WPF and Avalonia UIs are thin callers of this function; the same
        function can be invoked from a scheduled task or pipeline.

    .PARAMETER CopyFromPattern
        Substring identifying the source objects (matched case-insensitively
        against the policy name).

    .PARAMETER CopyToPattern
        Replacement text for CopyFromPattern in the copied object's name.

    .PARAMETER PolicyType
        Restrict the copy to these PolicyType IDs. If both PolicyType and
        PolicyGroup are omitted, every type whose group allows Copy is processed.

    .PARAMETER PolicyGroup
        Restrict the copy to these PolicyGroup IDs (expanded to their member types).

    .PARAMETER TokenId
        Authentication token id. Defaults to the current default token.

    .EXAMPLE
        Start-GraphBulkCopy -CopyFromPattern "Test - " -CopyToPattern "Prod - "

    .EXAMPLE
        Start-GraphBulkCopy -CopyFromPattern "Pilot" -CopyToPattern "Rollout" -PolicyGroup DeviceConfiguration

    .OUTPUTS
        PSCustomObject with summary statistics (Types, Copied, Skipped, FailedTypes, Duration).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $CopyFromPattern,

        [Parameter(Mandatory = $true)]
        [string]
        $CopyToPattern,

        [string[]]
        $PolicyType,

        [string[]]
        $PolicyGroup,

        [Int]
        $TokenId = (Get-DefaultTokenId)
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # ---- 1. Determine target policy types (same shape as Start-GraphBulkExport) ----
    $selection = Resolve-IntuneTargetSelectors -PolicyType $PolicyType -PolicyGroup $PolicyGroup -Caller 'Start-GraphBulkCopy'
    $unknownSelectors = $selection.Unknown   # ids from -PolicyType/-PolicyGroup that matched nothing
    $targetTypes = [System.Collections.Generic.List[object]]::new()
    $targetTypes.AddRange([object[]]$selection.Types)

    if (-not $PolicyType -and -not $PolicyGroup) {
        foreach ($grp in $script:IntuneGroups) {
            if (-not $grp.Title) { continue }
            if ($grp.ShowButtons -is [Object[]] -and $grp.ShowButtons -notcontains "Copy") { continue }
            foreach ($pt in $grp.PolicyTypes) { [void]$targetTypes.Add($pt) }
        }
    }

    $targetTypes = @($targetTypes | Sort-Object -Property Id -Unique)

    if ($targetTypes.Count -eq 0) {
        Write-Log "Bulk copy: no policy types selected" 2
        return [PSCustomObject]@{ Types = 0; Copied = 0; Skipped = 0; FailedTypes = 0; UnknownSelectors = $unknownSelectors; Duration = $stopwatch.Elapsed }
    }

    Write-Log "****************************************************************"
    Write-Log "Start bulk copy ('$CopyFromPattern' -> '$CopyToPattern', $($targetTypes.Count) policy type(s))"
    Write-Log "****************************************************************"

    $escapedFrom  = [regex]::Escape($CopyFromPattern)
    $totalCopied  = 0
    $totalSkipped = 0
    $failedTypes  = 0
    $typeIndex    = 0

    foreach ($pt in $targetTypes) {
        $typeIndex++
        Write-Status -Text ("Bulk copy - {0} ({1} of {2})" -f $pt.Title, $typeIndex, $targetTypes.Count) -Detail "Listing policies" -SkipLog -Force

        try {
            $policies = @(Get-GraphPolicies -PolicyType $pt.Id -TokenId $TokenId -ErrorAction Stop)
        }
        catch {
            Write-LogError "Bulk copy: failed to list $($pt.Title) objects" $_.Exception
            $failedTypes++
            continue
        }
        if ($policies.Count -eq 0) { continue }

        $sources = @($policies | Where-Object { $_.Name -imatch $escapedFrom })
        if ($sources.Count -eq 0) { continue }

        Write-Log "----------------------------------------------------------------"
        Write-Log "Copy $($pt.Title) objects"
        Write-Log "----------------------------------------------------------------"

        # Existing-name index so re-runs don't create duplicates. Keyed on
        # @odata.type + name, matching the original implementation.
        $existing = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($p in $policies) {
            [void]$existing.Add("$($p.JsonObject.'@odata.type')|$($p.Name)")
        }

        $toCopy = @()
        foreach ($src in $sources) {
            $targetName = $src.Name -ireplace $escapedFrom, $CopyToPattern
            if ($existing.Contains("$($src.JsonObject.'@odata.type')|$targetName")) {
                Write-Log "Object with name '$targetName' already exists. '$($src.Name)' will not be copied" 2
                $totalSkipped++
                continue
            }
            $toCopy += $src
        }
        if ($toCopy.Count -eq 0) { continue }

        Write-Status -Detail ("Copying {0} object(s)" -f $toCopy.Count) -SkipLog -Force
        try {
            $copied = @($toCopy | Copy-GraphPolicy -CopyFromPatternName $CopyFromPattern -Name $CopyToPattern -TokenId $TokenId)
            $totalCopied += $copied.Count
            if ($copied.Count -lt $toCopy.Count) {
                Write-Log "Bulk copy: $($toCopy.Count - $copied.Count) $($pt.Title) object(s) failed to copy" 2
            }
        }
        catch {
            Write-LogError "Bulk copy: failed while copying $($pt.Title) objects" $_.Exception
            $failedTypes++
        }
    }

    Write-Status $null
    Write-Log "****************************************************************"
    Write-Log "Bulk copy finished. Copied: $totalCopied, skipped (name exists): $totalSkipped"
    Write-Log "****************************************************************"

    return [PSCustomObject]@{
        Types            = $targetTypes.Count
        Copied           = $totalCopied
        Skipped          = $totalSkipped
        FailedTypes      = $failedTypes
        UnknownSelectors = $unknownSelectors
        Duration         = $stopwatch.Elapsed
    }
}
