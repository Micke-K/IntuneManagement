function Export-GraphPolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyBase[]]
        $InputObject,
        [Parameter(Mandatory = $true)]
        [IntuneManagerExportSettings]
        $ExportSettings,
        # Emit the full path of each exported file to the pipeline. Opt-in; the
        # per-policy branch in Process only outputs when this is set.
        [switch]
        $PassThru
    )

    Begin {
        Write-Log "Start Export"

        # Honour the ClearCacheBeforeExportImport setting (bit 2 = manual
        # export). Begin runs once per pipeline invocation so a multi-policy
        # export clears at most once. Bulk export clears via its own driver —
        # the defer-flag it sets doubles as the "called from bulk" marker so
        # the per-type pipelines inside a bulk run don't re-clear.
        if (-not $script:_bulkExportDeferMigFlush) {
            Invoke-GraphCacheClearBeforeOperation -Operation ManualExport
        }

        $exportFolderRoot = $ExportSettings.ExportFolder
        if ($ExportSettings.AddCompanyName) {
            # The organisation display name is tenant-controlled and can legally
            # contain path separators ("Contoso A/S"), which would silently split
            # the export into a nested folder. Treat it as a single path segment.
            $companyFolder = Remove-InvalidFileNameChars (Get-CurrentOrganizationName)
            if (-not [String]::IsNullOrWhiteSpace($companyFolder)) {
                $exportFolderRoot = [IO.Path]::Combine($exportFolderRoot, $companyFolder)
            }
        }

        # Create the export root up-front so the user sees the folder even when no
        # policies are written (empty selection / all-zero matches / a downstream
        # error). Previously the directory only got created inside Process, so a
        # silent-skip in Process meant no folder appeared anywhere.
        try {
            if (-not [IO.Directory]::Exists($exportFolderRoot)) {
                [IO.Directory]::CreateDirectory($exportFolderRoot) | Out-Null
                Write-Log "Export: created folder $exportFolderRoot"
            }
        }
        catch {
            Write-LogError "Export: failed to create export root '$exportFolderRoot'" $_.Exception
        }

        $script:_exportSkippedNoTenantID = 0
    }

    Process {
        foreach ($policyObject in $InputObject) {
            if (-not $policyObject.TenantID) {
                # Was silent — explicit log so "Exported N, on disk 0" doesn't repeat.
                # `continue` (not `return`) so the foreach moves to the next item: in
                # PowerShell, `return` inside a Process foreach exits that whole Process
                # invocation, dropping remaining pipeline items the caller passed in.
                $script:_exportSkippedNoTenantID++
                Write-Log "Export: skipped '$($policyObject.Name)' - TenantID not set on object" 2
                continue
            }

            Write-Log "Export $($policyObject.Name) - $($policyObject.PolicyName)"

            $exportFolder = $exportFolderRoot

            if ($ExportSettings.AddObjectType) {
                $exportFolder = [IO.Path]::Combine($exportFolder, $policyObject.PolicyType.Folder)
            }

            if ($policyObject.IsFullObject -eq $false) {
                # Bulk export pre-hydrates via Invoke-PolicyHydrate; this
                # safety-net Get() covers any policy that arrived un-hydrated
                # (single-policy export path, or a row that bulk skipped).
                # [void]: Get() is [Boolean], so a bare call emits True/False
                # onto this cmdlet's output stream and pollutes stdout.
                [void]$policyObject.Get()
            }

            Add-GraphNavigationProperties $policyObject

            try {
                if ([IO.Directory]::Exists($exportFolder) -eq $false) {
                    [IO.Directory]::CreateDirectory($exportFolder) | Out-Null
                }

                if ($ExportSettings.ExportAssignments -ne $true -and $policyObject.Assignments) {
                    Remove-Property $policyObject "Assignments"
                }
        
                $fullPath = $policyObject.ExportToFile($exportFolder)

                if ($fullPath) {
                    Set-CacheObject "CurrentExportAssignments" $ExportSettings.ExportAssignments
                    $policyObject.PolicyType.PostExportCommand($policyObject, $fullPath)

                    # Pass both $exportFolder (per-policy directory; some callers
                    # need it for sidecar logic) AND $exportFolderRoot as the
                    # explicit MigrationRoot. The latter eliminates the
                    # ".Parent.FullName" guesswork in Add-GraphMigrationObject
                    # which lands MigrationTable.json + Groups/ one level too high
                    # whenever $Folder is already the export root — i.e. when
                    # AddObjectType=false (per-policy files written straight to
                    # the org folder, no per-type subfolder).
                    Add-GraphMigrationInfo $policyObject -Folder $exportFolder -MigrationRoot $exportFolderRoot -MaxGroupDepth $ExportSettings.ExportNestedGroupLevels

                    if ($PassThru -eq $true) {
                        $fullPath
                    }
                }
            }
            catch {
                Write-LogError "Failed to export object" $_.Exception
            }            
        }
    }
    
    End {
        if ($script:_exportSkippedNoTenantID -gt 0) {
            Write-Log ("Export finished - $($script:_exportSkippedNoTenantID) policy/policies were skipped because TenantID was not set. Folder: $exportFolderRoot") 2
        }
        else {
            Write-Log "Export finished. Folder: $exportFolderRoot"
        }

        # Flush any deferred MigrationTable.json writes — but only when not running
        # inside a bulk-export pipeline (Start-GraphBulkExport sets the guard so it
        # can flush ONCE across all types instead of once per type, which would
        # rewrite the growing migration table N times).
        if (-not $script:_bulkExportDeferMigFlush -and
            (Get-Command Save-GraphMigrationFilesPending -ErrorAction SilentlyContinue)) {
            try { Save-GraphMigrationFilesPending } catch {
                Write-LogError "Export: failed to flush migration table(s)" $_.Exception
            }
        }
    }

}
