function Compare-GraphPolicy {
    [CmdletBinding(DefaultParameterSetName = 'Direct')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'Direct', Position = 0, ValueFromPipeline = $true)]
        [object[]]
        $Policies,

        [Parameter(Mandatory = $true, ParameterSetName = 'ExportFiles')]
        [CompareExportFilesProvider]
        $ExportFiles,

        [Parameter(Mandatory = $true, ParameterSetName = 'IntuneWithExport')]
        [CompareIntuneWithExportProvider]
        $IntuneWithExport,

        [Parameter(Mandatory = $true, ParameterSetName = 'NamedObjects')]
        [CompareNamedObjectsProvider]
        $NamedObjects,

        [Parameter(Mandatory = $true, ParameterSetName = 'ExportedFolders')]
        [CompareExportedFoldersProvider]
        $ExportedFolders,

        [Parameter(ParameterSetName = 'ExportFiles')]
        [Parameter(ParameterSetName = 'IntuneWithExport')]
        [Parameter(ParameterSetName = 'NamedObjects')]
        [Parameter(ParameterSetName = 'ExportedFolders')]
        [string[]]
        $PolicyGroupIds = @()
    )

    Process {
        switch($PSCmdlet.ParameterSetName)
        {
            'Direct' {
                if($Policies.Count -lt 2)
                {
                    if($Policies.Count -eq 1)
                    {
                        Show-GraphCompareForm $Policies[0]
                    }
                    else
                    {
                        Write-Error "Provide at least one policy. With a single policy the compare UI will open."
                    }
                    return
                }

                # Batch-hydrate via the unified orchestrator instead of N
                # sequential per-policy Get() calls. N>=2 here (single-policy
                # branch above returned early), so this always takes the
                # parallel $batch path inside Invoke-PolicyHydrate.
                $needFull = @($Policies | Where-Object {
                    $_ -and $_.PSObject.Properties['_IsFullObject'] -and -not $_._IsFullObject -and $_.Id -and $_.PolicyType
                })
                if($needFull.Count -gt 0) {
                    Invoke-PolicyHydrate -Policies $needFull
                }
                Compare-PolicyObjects $Policies
            }

            default {
                $provider = switch($PSCmdlet.ParameterSetName)
                {
                    'ExportFiles'     { $ExportFiles }
                    'IntuneWithExport' { $IntuneWithExport }
                    'NamedObjects'    { $NamedObjects }
                    'ExportedFolders' { $ExportedFolders }
                }

                $groups = if($PolicyGroupIds.Count -gt 0)
                {
                    # Unknown ids are logged and raised as non-terminating errors by
                    # the shared resolver instead of vanishing from the -in filter.
                    @((Resolve-IntuneTargetSelectors -PolicyGroup $PolicyGroupIds -Caller 'Compare-GraphPolicy').Groups)
                }
                else
                {
                    @($script:IntuneGroups)
                }

                if(-not $groups)
                {
                    Write-Error "No policy groups found. Ensure the module is initialized and group IDs are correct."
                    return
                }

                $pairs = $provider.GetComparePairs($groups)
                foreach($pair in $pairs)
                {
                    $result = Compare-PolicyObjects @($pair.Policy1, $pair.Policy2)
                    [PSCustomObject]@{
                        Name       = $pair.Name
                        Id         = $pair.Id
                        PolicyType = $pair.PolicyType.Title
                        Result     = $result
                    }
                }
            }
        }
    }
}
