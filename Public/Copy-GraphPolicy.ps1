function Copy-GraphPolicy {
    [CmdletBinding()]
    [OutputType([IntunePolicyBase[]])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyBase[]]
        $InputObject,
        
        [Parameter(Mandatory = $true)]
        # Name of new policy
        [String]
        $Name,

        # Description of new policy
        [String]
        $Description,

        # Specifies a name pattern to replace. Used when copying multiple policies. The Name parameter then replaces the pattern in each policy name.
        [String]
        $CopyFromPatternName,

        # Specifies a description pattern to replace. Used when copying multiple policies. The Description parameter then replaces the pattern in each policy description.
        [String]
        $CopyFromPatternDescription,

        # Specifies destination environment. Default is current logged on environment.
        [int]
        $TokenId = (Get-DefaultTokenId),

        # Override the scope tag IDs the new copy will have. When omitted the copy
        # inherits whatever scope tags the source had; when provided (even as an
        # empty array) this list wins. Passed to CopyObject which writes them into
        # the cloned JSON's ScopeTagProperty (typically roleScopeTagIds) before POST.
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]]
        $ScopeTagIds
    )

    Begin {
        Write-Log "Start Copy"
        $copiedPolicies = @()
        # Collect across Process invocations — `$policies | Copy-GraphPolicy`
        # triggers Process once per pipeline item, so per-Process batching
        # would still be per-item single GETs. Hydrate + copy run in End.
        $allInput = [System.Collections.Generic.List[object]]::new()
    }

    Process {
        foreach ($p in $InputObject) {
            if ($p) { [void]$allInput.Add($p) }
        }
    }

    End {
        # Pre-hydrate the source policies in one Invoke-PolicyHydrate call —
        # N>1 takes the parallel $batch path. CopyObject internally calls
        # $this.Get() with the IsFullObject guard, so already-hydrated rows
        # are a no-op there.
        $hydrateTargets = @($allInput | Where-Object {
            $_.PSObject.Properties['_IsFullObject'] -and -not $_._IsFullObject -and
            $_.Id -and $_.PolicyType -and $_.IsFromFile -ne $true
        })
        if ($hydrateTargets.Count -gt 0) {
            Invoke-PolicyHydrate -Policies $hydrateTargets
        }

        foreach ($policyObject in $allInput) {
            $newName = $null
            $newDescription = $null

            if ($CopyFromPatternName -and $policyObject.Name -imatch [regex]::Escape($CopyFromPatternName)) {
                $newName = $policyObject.Name -ireplace [regex]::Escape($CopyFromPatternName), $Name
            }
            elseif ($CopyFromPatternName) {
                Write-Verbose "$CopyFromPatternName did not match the pattern of the policy name '$($policyObject.Name)'. Skipping policy"
                continue
            }
            else {
                $newName = $Name
            }

            if ($Description -and $CopyFromPatternDescription -and $policyObject.Description -imatch [regex]::Escape($CopyFromPatternDescription)) {
                $newDescription = $policyObject.Description -ireplace [regex]::Escape($CopyFromPatternDescription), $Description
            }
            elseif ($Description -and -not $CopyFromPatternDescription) {
                $newDescription = $Description
            }

            Write-Verbose "Copy policy '$($policyObject.Name)' to '$($newName)'."

            if($PSBoundParameters.ContainsKey('ScopeTagIds')) {
                $newPolicy = $policyObject.CopyObject($newName, $newDescription, $TokenId, $ScopeTagIds)
            }
            else {
                $newPolicy = $policyObject.CopyObject($newName, $newDescription, $TokenId)
            }
            if ($newPolicy) {
                $copiedPolicies += $newPolicy
            }
        }

        Write-Log "Copy finished. $($copiedPolicies.Count) policies copied"
        return $copiedPolicies
    }
}
