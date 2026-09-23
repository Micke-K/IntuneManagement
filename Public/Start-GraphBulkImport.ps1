function Start-GraphBulkImport {
    <#
    .SYNOPSIS
        Bulk-import exported Intune policy objects from disk.

    .DESCRIPTION
        Public, UI-independent driver for bulk import. The WPF and Avalonia
        bulk-import forms are thin callers of this function; the same function
        can be invoked from a scheduled task or pipeline.

        Groups are processed in ascending ImportOrder so dependencies (Scope
        Tags, etc.) import before the objects that reference them.

    .PARAMETER ImportFolder
        Root folder of a previous export (the folder that contains the
        per-policy-type sub-folders).

    .PARAMETER Filter
        Name filter (literal substring, case-insensitive) applied to the file
        objects before import. Empty = import everything found.

    .PARAMETER PolicyGroup
        Restrict the import to these PolicyGroup IDs. If omitted, every group
        that allows Import is processed.

    .PARAMETER ImportAssignments
        Import object assignments. Persisted to the ImportAssignments setting
        (the import pipeline reads it from there) when supplied.

    .PARAMETER ImportScopeTags
        Import scope tags. Persisted to the ImportScopeTags setting when supplied.

    .PARAMETER ReplaceDependencyIDs
        Translate dependency object IDs via the migration table. Persisted to
        the ResolveReferenceInfo setting when supplied.

    .PARAMETER ImportType
        How files are imported: alwaysImport (default), skipIfExist, update,
        replace, or replace_with_assignments. Persisted to the ImportType
        setting when supplied. Anything other than alwaysImport resolves each
        file against the existing objects (Resolve-IntuneImportUpdateTarget)
        and skips / updates / replaces accordingly.

    .PARAMETER TokenId
        Authentication token id. Defaults to the current default token.

    .EXAMPLE
        Start-GraphBulkImport -ImportFolder C:\IntuneExport

    .EXAMPLE
        Start-GraphBulkImport -ImportFolder C:\IntuneExport -PolicyGroup DeviceConfiguration -Filter "Baseline"

    .OUTPUTS
        PSCustomObject with summary statistics (Groups, Imported, Duration).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $ImportFolder,

        [string]
        $Filter,

        [string[]]
        $PolicyGroup,

        [Nullable[bool]]
        $ImportAssignments,

        [Nullable[bool]]
        $ImportScopeTags,

        [Nullable[bool]]
        $ReplaceDependencyIDs,

        [string]
        $ImportType,

        [Int]
        $TokenId = (Get-DefaultTokenId)
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $ImportFolder = Expand-FileName $ImportFolder
    if (-not [IO.Directory]::Exists($ImportFolder)) {
        throw "Import folder not found: $ImportFolder"
    }

    # The import pipeline (ImportObject / assignment + scope-tag handling) reads
    # these via Get-SettingValue, which resolves them from their registered
    # SubPath ("IntuneManager"). Persist to the SAME SubPath so a supplied
    # per-run value is actually read back (saving to root "" left them orphaned).
    if ($null -ne $ImportAssignments)    { Save-SettingStoreValue "IntuneManager" "ImportAssignments"    $ImportAssignments }
    if ($null -ne $ImportScopeTags)      { Save-SettingStoreValue "IntuneManager" "ImportScopeTags"      $ImportScopeTags }
    if ($null -ne $ReplaceDependencyIDs) { Save-SettingStoreValue "IntuneManager" "ResolveReferenceInfo" $ReplaceDependencyIDs }
    if ($ImportType)                     { Save-SettingStoreValue "IntuneManager" "ImportType"           $ImportType }

    # ---- Resolve target groups ----
    $selection = Resolve-IntuneTargetSelectors -PolicyGroup $PolicyGroup -Caller 'Start-GraphBulkImport'
    $unknownSelectors = $selection.Unknown   # ids from -PolicyGroup that matched nothing
    $targetGroups = [System.Collections.Generic.List[object]]::new()
    $targetGroups.AddRange([object[]]$selection.Groups)
    if (-not $PolicyGroup) {
        foreach ($grp in $script:IntuneGroups) {
            if (-not $grp.Title) { continue }
            if ($grp.ShowButtons -is [Object[]] -and $grp.ShowButtons -notcontains "Import") { continue }
            [void]$targetGroups.Add($grp)
        }
    }
    if ($targetGroups.Count -eq 0) {
        Write-Log "Bulk import: no policy groups selected" 2
        return [PSCustomObject]@{ Groups = 0; Imported = 0; UnknownSelectors = $unknownSelectors; Duration = $stopwatch.Elapsed }
    }

    # Honour the ClearCacheBeforeExportImport setting (bit 4 = bulk import).
    Invoke-GraphCacheClearBeforeOperation -Operation BulkImport -TokenId $TokenId

    Write-Log "****************************************************************"
    Write-Log "Start bulk import from $ImportFolder"
    Write-Log "****************************************************************"

    # Effective import type: parameter wins, then the persisted setting.
    $importTypeEffective = if ($ImportType) { $ImportType } else { [string](Get-SettingValue "ImportType" "alwaysImport") }
    if (-not $importTypeEffective) { $importTypeEffective = "alwaysImport" }

    $sameTenant = $false
    if ($importTypeEffective -ne "alwaysImport") {
        try {
            $null, $sameTenant = Get-MigrationTableInfo $ImportFolder (Get-CurrentTenantId)
        } catch { }
    }

    $totalImported = 0
    $totalUpdated  = 0
    $totalReplaced = 0
    $totalSkipped  = 0
    # Ascending ImportOrder so dependencies import before their dependents.
    $orderedGroups = @($targetGroups | Sort-Object { ($_.PolicyTypes | Measure-Object ImportOrder -Minimum).Minimum })
    $groupTotal    = $orderedGroups.Count
    $groupIndex    = 0

    foreach ($grp in $orderedGroups) {
        $groupIndex++
        $policyTypes = $grp.PolicyTypes
        $subFolders  = @($policyTypes | ForEach-Object { $_.Folder } | Where-Object { $_ })

        Write-Log "----------------------------------------------------------------"
        Write-Log "Import $($grp.Title)"
        Write-Log "----------------------------------------------------------------"

        try {
            Write-Status `
                -Text   ("Bulk import - {0} ({1} of {2})" -f $grp.Title, $groupIndex, $groupTotal) `
                -Detail "Loading objects from folder" `
                -Force

            $params = @{ Path = $ImportFolder; PolicyTypes = $policyTypes }
            if ($subFolders.Count -gt 0) { $params["SubFolders"] = $subFolders }

            $policiesToImport = @(Get-PoliciesFromFolder @params)

            if ($Filter) {
                $policiesToImport = @($policiesToImport | Where-Object { $_.Name -match [RegEx]::Escape($Filter) })
            }

            if ($policiesToImport.Count -gt 0) {
                # Any mode other than alwaysImport resolves each file against
                # the existing objects first (skip / update / replace).
                if ($importTypeEffective -ne "alwaysImport") {
                    $existing = @(Get-GraphPolicies -PolicyGroup $grp.ID -TokenId $TokenId)
                    $toCreate = @()

                    foreach ($importPolicy in $policiesToImport) {
                        $match = Resolve-IntuneImportUpdateTarget -ImportPolicy $importPolicy -ExistingPolicies $existing -SameTenant $sameTenant -ImportType $importTypeEffective

                        switch ($match.Action) {
                            "Update" {
                                try {
                                    if ($importPolicy.UpdateObject($match.Target, $TokenId)) { $totalUpdated++ }
                                } catch { Write-LogError "UpdateObject failed for $($importPolicy.Name)" $_.Exception }
                            }
                            "Replace" {
                                try {
                                    if (Invoke-IntuneImportReplace -ImportPolicy $importPolicy -Target $match.Target -ImportType $importTypeEffective -TokenId $TokenId) { $totalReplaced++ }
                                } catch { Write-LogError "Replace failed for $($importPolicy.Name)" $_.Exception }
                            }
                            "Ambiguous" {
                                $totalSkipped++
                                Write-Log "Skip import for $($importPolicy.Name): $($match.Message)" 2
                            }
                            "Skip" {
                                $totalSkipped++
                                Write-Log "Skip import for $($importPolicy.Name): $($match.Message)"
                            }
                            default { $toCreate += $importPolicy }
                        }
                    }
                    $policiesToImport = $toCreate
                }

                if ($policiesToImport.Count -gt 0) {
                    Write-Status -Detail ("Importing {0} object(s)" -f $policiesToImport.Count) -SkipLog -Force
                    $imported = @($policiesToImport | Import-GraphPolicy)
                    $totalImported += $imported.Count
                    Write-Log "Imported $($imported.Count) $($grp.Title) object(s)"
                }
            }
            else {
                Write-Log "No $($grp.Title) files found in $ImportFolder"
            }
        }
        catch {
            Write-LogError "Failed when importing $($grp.Title)" $_.Exception
        }
    }

    Write-Status $null
    Write-Log "****************************************************************"
    Write-Log "Bulk import finished. $totalImported imported, $totalUpdated updated, $totalReplaced replaced, $totalSkipped skipped"
    Write-Log "****************************************************************"

    return [PSCustomObject]@{
        Groups           = $orderedGroups.Count
        Imported         = $totalImported
        Updated          = $totalUpdated
        Replaced         = $totalReplaced
        Skipped          = $totalSkipped
        UnknownSelectors = $unknownSelectors
        Duration         = $stopwatch.Elapsed
    }
}
