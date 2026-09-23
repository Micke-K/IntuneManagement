<#
.SYNOPSIS
    Bulk-document Intune/Entra policies through one or more output providers.

.DESCRIPTION
    Public, UI-independent driver for bulk documentation. The WPF UI is a
    thin caller of this function; the same function can be invoked from a
    scheduled task or CI/CD pipeline without loading any XAML.

    Iterates the supplied PolicyObjects (or, if -PolicyType/-PolicyGroup are
    used, the policies fetched via Get-GraphPolicies), runs each through
    Invoke-DocumentationForObject to build a per-object result, and drives
    each selected output provider's lifecycle hooks:
        Activate -> PreProcess
            -> (NewObjectGroup -> NewObjectType -> Process* -> ProcessAllObjects)*
        -> PostProcess

.PARAMETER OutputFormat
    Comma-separated registered output Values (e.g. 'json,md'). Each one
    must be in [DocumentationRegistry]::Outputs (run Get-DocumentationOutput
    to list).

.PARAMETER PolicyObject
    Pre-fetched policy objects. If omitted, the function fetches via
    Get-GraphPolicies using -PolicyType / -PolicyGroup.

.PARAMETER PolicyType
    Restrict iteration to these PolicyType IDs.

.PARAMETER PolicyGroup
    Restrict iteration to these PolicyGroup IDs.

.PARAMETER Language
    Language code for translatable strings. Defaults to 'en'.

.PARAMETER SourceFolder
    Load policies from an exported folder instead of querying Graph.

.PARAMETER Options
    Hashtable of engine-wide flags. See Get-GraphDocumentation.

.EXAMPLE
    Start-GraphBulkDocumentation -OutputFormat json -PolicyType ConditionalAccessType

.EXAMPLE
    $policies | Start-GraphBulkDocumentation -OutputFormat 'md'
#>
function Start-GraphBulkDocumentation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        # Completion instead of [ValidateSet([DocumentationOutputProvider])] for PS5.1
        # import compatibility (the generator is PS7-only). Also lets comma-separated
        # values ('json,md') through, which a single-value ValidateSet would reject.
        # Unknown formats are ignored at runtime (matched against the output registry).
        [ArgumentCompleter({ param($commandName, $parameterName, $wordToComplete) @(& (Get-Module IntuneManagement) { Get-DocumentationOutputValues }) | Where-Object { $_ -like "$wordToComplete*" } })]
        [string]
        $OutputFormat,
        [Parameter(ValueFromPipeline, Mandatory = $true, ParameterSetName = 'PolicyObject', Position = 1)]
        $PolicyObject,
        [Parameter(Mandatory = $true, ParameterSetName = 'PolicyType', Position = 1)]
        [ArgumentCompleter({ param($commandName, $parameterName, $wordToComplete) @(& (Get-Module IntuneManagement) { Get-IntunePolicyTypeValues }) | Where-Object { $_ -like "$wordToComplete*" } })]
        [string[]]
        $PolicyType,
        [Parameter(Mandatory = $true, ParameterSetName = 'PolicyGroup', Position = 1)]
        [ArgumentCompleter({ param($commandName, $parameterName, $wordToComplete) @(& (Get-Module IntuneManagement) { Get-IntunePolicyGroupValues }) | Where-Object { $_ -like "$wordToComplete*" } })]
        [string[]]
        $PolicyGroup,
        [Parameter(Mandatory = $true, ParameterSetName = 'Folder', Position = 1)]
        [string]
        $SourceFolder,
        [string]
        $Language = 'en',
        [hashtable]
        $Options
    )

    begin {
        $collected = [System.Collections.Generic.List[object]]::new()
    }

    process {
        if ($PolicyObject) {
            foreach ($p in $PolicyObject) { $collected.Add($p) }
        }
    }

    end {
        if ($SourceFolder) {
            if (-not $Options) { $Options = @{} }
            $Options.SourceTenantUnavailable = $true
            Initialize-DocumentationSourceTenantContext -SourceFolder $SourceFolder
            if ($collected.Count -eq 0) {
                $fetched = Get-DocumentationPoliciesFromSourceFolder -SourceFolder $SourceFolder -PolicyType $PolicyType -PolicyGroup $PolicyGroup
                foreach ($p in $fetched) { $collected.Add($p) }
            }
        }
        elseif ($collected.Count -eq 0 -and ($PolicyType -or $PolicyGroup)) {
            $params = @{}
            if ($PolicyType)  { $params['PolicyType']  = $PolicyType }
            if ($PolicyGroup) { $params['PolicyGroup'] = $PolicyGroup }
            $fetched = Get-GraphPolicies @params
            foreach ($p in $fetched) { $collected.Add($p) }
        }

        if ($collected.Count -eq 0) {
            Write-Log "Start-GraphBulkDocumentation: no policies to document" 2
            return
        }
        Write-Log "$($collected.Count) policies found"

        $items = $collected.ToArray()
        # Save/restore module-wide language: generation points Get-LanguageString at
        # the requested language; other consumers must not inherit it afterwards.
        $prevLang = $script:CurrentLanguage
        try {
        Set-DocumentationContextRunOptions -Context (Get-DocContextSingleton) -Options $Options -Language $Language

        # NameFilter can be a legacy string (plain substring / scope: / tag:) and/or
        # scriptblock stages: LIST{ } runs pre-hydrate (cheap, drops items before they
        # are hydrated); ITEM{ } runs post-hydrate (has full body incl. scope tags —
        # required for types whose list endpoint omits roleScopeTagIds, e.g. Applications).
        if ($Options -and $Options.NameFilter) {
            $parsed = ConvertFrom-DocumentationNameFilter ([string]$Options.NameFilter)

            # LIST stage (pre-hydrate): legacy match + LIST scriptblock.
            if ($parsed.Legacy) {
                # A scope:/tag: filter needs scope tags. Some types (e.g. Applications,
                # _ScopeTagsReturnedInList = $false) don't return roleScopeTagIds in the
                # list response - only on the full per-object GET - so filtering them
                # pre-hydrate would wrongly drop every one. Filter the types that DO
                # expose tags in the list now (cheap), and defer the rest to a post-
                # hydrate pass. A plain-substring name filter matches .Name (always
                # present pre-hydrate), so it keeps the single cheap pass.
                if ($parsed.Legacy.Trim() -match '^(?i:scope|tag):') {
                    # NOTE: ScopeTagsReturnedInList is a property of the PolicyType
                    # (IntunePolicyTypeBase), not the policy instance - so it must be
                    # read via .PolicyType. Reading it off the instance returns $null
                    # and would route every item to the ready bucket.
                    $deferred = @($items | Where-Object { $_.PolicyType -and $_.PolicyType.ScopeTagsReturnedInList -eq $false })
                    $ready    = @($items | Where-Object { -not ($_.PolicyType -and $_.PolicyType.ScopeTagsReturnedInList -eq $false) })

                    $ready = @($ready | Where-Object { Test-DocumentationPolicyFilter -PolicyObject $_ -Filter $parsed.Legacy })

                    if ($deferred.Count -gt 0) {
                        # Hydrate the deferred survivors so their scope tags populate.
                        # Skipped offline (SourceFolder): those file objects carry
                        # roleScopeTags already and have no token to hydrate against.
                        if (-not $Options.SourceTenantUnavailable) {
                            $hydrateTargets = @($deferred | Where-Object {
                                $_ -and $_.PSObject.Properties['_IsFullObject'] -and -not $_._IsFullObject -and
                                $_.Id -and $_.PolicyType -and $_.IsFromFile -ne $true
                            })
                            if ($hydrateTargets.Count -gt 0) {
                                Write-Status "Documentation - hydrating policies for scope filter" -SkipLog -Force
                                Invoke-PolicyHydrate -Policies $hydrateTargets
                            }
                        }
                        $deferred = @($deferred | Where-Object { Test-DocumentationPolicyFilter -PolicyObject $_ -Filter $parsed.Legacy })
                    }

                    $items = @($ready) + @($deferred)
                }
                else {
                    $items = @($items | Where-Object { Test-DocumentationPolicyFilter -PolicyObject $_ -Filter $parsed.Legacy })
                }
            }
            if ($parsed.List) {
                $items = @($items | Where-Object { Test-DocumentationFilterScriptBlock -PolicyObject $_ -ScriptBlock $parsed.List })
            }

            # ITEM stage (post-hydrate): hydrate the survivors, then filter. Skipped
            # for SourceFolder (offline) runs — those objects have no token to hydrate
            # against and rely on whatever the export already carries. Invoke-PolicyHydrate
            # is idempotent, so the engine's own later hydrate call is a no-op for these.
            if ($parsed.Item -and -not ($Options.SourceTenantUnavailable)) {
                $hydrateTargets = @($items | Where-Object {
                    $_ -and $_.PSObject.Properties['_IsFullObject'] -and -not $_._IsFullObject -and
                    $_.Id -and $_.PolicyType -and $_.IsFromFile -ne $true
                })
                if ($hydrateTargets.Count -gt 0) {
                    Write-Status "Documentation - hydrating policies for ITEM filter" -SkipLog -Force
                    Invoke-PolicyHydrate -Policies $hydrateTargets
                }
                $items = @($items | Where-Object { Test-DocumentationFilterScriptBlock -PolicyObject $_ -ScriptBlock $parsed.Item })
            }
        }
        Write-Log "$($items.Count) policies to document"
        Invoke-DocumentationOutputs -OutputValue $OutputFormat -PolicyObjects $items -Options $Options -Language $Language
        }
        finally {
            $script:CurrentLanguage = $prevLang
        }
    }
}
