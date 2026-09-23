function Get-GraphPolicyFromFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Alias("FileInfo")]
        [IO.FileInfo[]]
        $InputObject, 
        [IntunePolicyTypeBase[]]
        $FromPolicyTypes,
        [String]
        $TenantId
    )

    Begin {
        Write-LogDebug "Get Policies from file(s)"

        $policyObjects = @()
    }

    Process {
        foreach ($fi in $InputObject) {
            if ($fi.Exists -eq $false) {
                Write-Log "File $($fi.FullName) not found. Cannot load policy" 3
                continue
            }

            try {
                Write-LogDebug "Get policy from file $($fi.FullName)"
                $jsonObj = ConvertFrom-Json ([IO.File]::ReadAllText($fi.FullName)) -ErrorAction Stop
                $policyType = Get-PoliciesTypeFromObject $jsonObj $FromPolicyTypes
                if ($policyType) {
                    $policyObject = $policyType.GetObject($fi)
                    if ($policyObject) {
                        if ($TenantId) {
                            $policyObject.TenantId = $TenantId
                        }
                        $policyObjects += $policyObject
                    }
                }
                else {
                    # Expected not to find some policies if policy types is passed in
                    Write-LogDebug "Could not get policy type from file $($fi.FullName)" 3
                }
            }
            catch {
                Write-LogDebug "Could get policy from file $($fi.FullName)" $_.Exception
            }
        }
    }
    
    End { 
        Write-LogDebug "Get policy from file finished"
        if ($policyObjects.Count -gt 0) {
            return $policyObjects
        }
    }
}