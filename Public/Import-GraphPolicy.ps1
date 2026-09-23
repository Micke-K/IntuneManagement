function Import-GraphPolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyBase[]]
        $InputObject,

        [int]
        $TokenId = (Get-DefaultTokenId)
    )

    Begin {
        Write-Log "Start Import of $(($InputObject | Measure-Object).Count) object(s)"
        $importedPolicies = @()
        $navigationPropObjects = @()

        $policyObjectList = [System.Collections.Generic.List[object]]::new()

        $bulkImport = $null
        if (Test-GraphBatchEnabled) {
            $bulkImport = @{}
        }

        ### !!! ToDo: Fix support for sorting based on priority + PreFilesImportCommand 
    }

    Process {
        foreach ($policyObject in $InputObject) {            
            $policyObjectList.Add($policyObject)
        }
    }

    End {
        # Group first, THEN sort the groups. Group-Object orders its output by
        # key, so a Sort-Object placed before it is thrown away and every import
        # ran alphabetically by type Id - ImportOrder had no effect at all. That
        # put App Config ahead of the Applications it targets and Policy Sets
        # ahead of the Settings Catalog they bundle; it only ever went unnoticed
        # because a same-tenant import finds every reference already in place.
        $policyTypeGroups = $policyObjectList |
            Group-Object -Property { $_.PolicyType.Id } |
            Sort-Object { [int]$_.Group[0].PolicyType.ImportOrder }
        foreach ($policyTypeGroup in $policyTypeGroups) {
            $policyType = Get-PolicyTypeFromID $policyTypeGroup.Name
            Write-Log "Import $($policyTypeGroup.Count) $($policyType.Title) policy object(s)" 
            
            $policyTypeObjects = $policyType.PreImportPolicies($policyTypeGroup.Group)

            $bulkImport = $null
            if (Test-GraphBatchEnabled) {
                $bulkImport = @{}
            }

            foreach ($policyObject in $policyTypeObjects) {
                # One failing policy must not abort the whole batch - log it
                # and keep importing the rest.
                $importedPolicy = $null
                try {
                    $importedPolicy = $policyObject.ImportObject($TokenId, $bulkImport)
                }
                catch {
                    Write-LogError "Import failed for $($policyType.Title) object '$($policyObject.Name)'" $_.Exception
                }
                if ($importedPolicy) {
                    $importedPolicies += [PSCustomObject]@{
                        ImportedObject = $importedPolicy
                        FromObject     = $policyObject
                    }

                    if ($importedPolicy.PolicyType.NavigationProperties -eq $true) {
                        $navigationPropObjects += [PSCustomObject]@{
                            ImportedObject = $importedPolicy
                            FromObject     = $policyObject
                        }
                    }
                }
            }

            if($null -ne $bulkImport) {
                # ToDo: Fix support for dependencies to make sure dependency objkects are imported first

                $batchObjects = [System.Collections.Generic.List[PSCustomObject]]::new()
                $bulkImport.Keys | ForEach-Object { $batchObjects.Add($_) }
                $batchResults = Invoke-GraphBatchRequest -BatchObject $batchObjects -TokenId $TokenId -BatchType "Import"

                $batchResults | ForEach-Object {
                    $result = $_
                    $result | Add-Member -MemberType NoteProperty -Name "Success" -Value ($result.Status -lt 300) -Force
                    $result | Add-Member -MemberType NoteProperty -Name "Content" -Value ($result.body | ConvertTo-Json -Depth 50) -Force
                    $key = $bulkImport.Keys | Where-Object id -eq $result.id
                    $bulkEntry = $bulkImport[$key]
                    if($bulkEntry) {
                        $policy = $bulkEntry.ImportObject
                        $sourcePolicy = $bulkEntry.FromObject
                        $importedPolicy = $policy.ProcessImportResponse($TokenId, $result, $key.Method)
                        if($importedPolicy) {
                            $importedPolicies += [PSCustomObject]@{
                                ImportedObject = $importedPolicy
                                FromObject     = $sourcePolicy
                            }

                            if ($importedPolicy.PolicyType.NavigationProperties -eq $true) {
                                $navigationPropObjects += [PSCustomObject]@{
                                    ImportedObject = $importedPolicy
                                    FromObject     = $sourcePolicy
                                }
                            }
                        }
                    }
                }
            }
        }

        if ($importedPolicies.Count -gt 0) {
            foreach ($importedPolicy in $importedPolicies) {
                $importedPolicy.ImportedObject.PolicyType.PostBulkImportCommand($importedPolicy.ImportedObject, $importedPolicy.FromObject)
            }

            foreach ($navPropObj in $navigationPropObjects) {
                Set-GraphNavigationProperties $navPropObj.ImportedObject $navPropObj.FromObject
            }
        }

        Write-Log "Import finished. $($importedPolicies.Count) policies imported"
        return $importedPolicies
    }
}
