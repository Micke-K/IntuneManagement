function Get-GraphPolicies {
    [CmdletBinding(DefaultParameterSetName = 'PolicyType')]
    [OutputType([IntunePolicyBase[]])]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'PolicyType', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        # Tab-completion for registered PolicyType IDs. Deliberately NOT a
        # [ValidateSet([IntunePolicyTypeValues])]: a generator-backed ValidateSet
        # (1) makes the cmdlet uncallable when no types are loaded yet (the
        # generator returns an empty set and binding throws "validValues out of
        # range"), and (2) cannot be mocked by Pester 3.4. Unknown ids are already
        # filtered out below (Where-Object Id -eq), so completion is enough.
        [ArgumentCompleter({ param($commandName, $parameterName, $wordToComplete) @(& (Get-Module IntuneManagement) { Get-IntunePolicyTypeValues }) | Where-Object { $_ -like "$wordToComplete*" } })]
        [string[]]$PolicyType,
        [Parameter(Mandatory = $true, ParameterSetName = 'PolicyGroup', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ArgumentCompleter({ param($commandName, $parameterName, $wordToComplete) @(& (Get-Module IntuneManagement) { Get-IntunePolicyGroupValues }) | Where-Object { $_ -like "$wordToComplete*" } })]
        [string[]]$PolicyGroup,
        [Parameter(Mandatory = $false, ParameterSetName = 'PolicyType')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PolicyGroup')]
        [switch]
        $SinglePage,
        [string]
        [ValidateSet("NextPage", "AllRemainingPages")]
        [Parameter(Mandatory = $true, ParameterSetName = 'ObjectsPaging')]
        $Paging,

        # When set, the returned policies will have their .Object.assignments populated.
        # Policy types that don't support assignments (SupportsAssignments = $false) are filtered out.
        # For types where assignments cannot be expanded inline, /assignments is batch-fetched after listing.
        [Parameter(Mandatory = $false, ParameterSetName = 'PolicyType')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PolicyGroup')]
        [switch]
        $IncludeAssignments,

        # Name prefix to search for. Sent to Graph as startswith() on the type's
        # name property when the endpoint supports it (PolicyType.SupportsNameFilter);
        # the result set is always re-checked client-side, so types whose endpoint
        # rejects the filter still return only matching policies.
        [Parameter(Mandatory = $false, ParameterSetName = 'PolicyType')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PolicyGroup')]
        [string]
        $NameFilter,

        # Specifies environment to get policies from. Default is current logged on environment.
        [Int]
        $TokenId = (Get-DefaultTokenId)
    )

    $params         = @{}
    $graphPolicies  = [System.Collections.Generic.List[IntunePolicyBase]]::new()
    $newPageObject  = $false

    # URL-keyed coalescing: when N policy types resolve to the same listing URL
    # (e.g. all 6 deviceEnrollmentConfigurations subtypes), we issue ONE list request
    # and run each row through every consumer's CheckPolicy in registration order —
    # first non-null match wins. Without this, group bulk-exports were issuing 5+
    # identical sub-requests in the $batch and discarding 80% of each response.
    #
    # Entry shapes:
    #   batch entry : { BatchObject = <{id,method,url,headers}>, Types = [PolicyType...] }
    #   api   entry : { ListURL = <string>,                     Types = [PolicyType...] }
    # The Types list preserves registration order so CheckPolicy ties are deterministic.
    $batchEntries = [System.Collections.Generic.List[PSCustomObject]]::new()
    $apiEntries   = [System.Collections.Generic.List[PSCustomObject]]::new()

    if ($PSCmdlet.ParameterSetName -ne "ObjectsPaging") {

        $policyTypes = [System.Collections.Generic.List[object]]::new()

        # -PolicyType and -PolicyGroup are separate parameter sets, so only one is
        # ever populated. An unknown id is logged and raised as a non-terminating
        # error by the resolver - the same contract as the bulk drivers, which this
        # cmdlet used to fall short of by dropping it in silence.
        $selection = Resolve-IntuneTargetSelectors -PolicyType $PolicyType -PolicyGroup $PolicyGroup -Caller 'Get-GraphPolicies'
        $policyTypes.AddRange([object[]]$selection.Types)

        # IMPORTANT: -IncludeAssignments controls whether assignments get fetched
        # alongside the policies, NOT which types get listed. Types with
        # SupportsAssignments=$false (Conditional Access, Named Locations, Terms of
        # Use, Filters, Role Definitions, ADMX Files, Reusable Settings, several
        # Tenant-Admin extras, etc.) still need to be listed — the assignment-fetch
        # in Add-GraphPolicyAssignments already skips them at the per-policy level.
        # Filtering them out here used to drop the whole type silently from bulk
        # export, which is why those folders were empty.

        # Every type that can describe its list call as a batch sub-request does so,
        # whatever the batching setting says: Invoke-GraphBatchRequest decides whether
        # those go out as one $batch or as direct calls, and either way the sub-
        # request carries the type's own Accept header. Choosing the plain-URL branch
        # here when batching was off sent the list call with the wrapper's default
        # metadata instead, so an Applications export with batching off carried
        # @odata annotations and navigation links the batched export never had.
        $useBatchAPI = $true

        # Coalesce by URL only for types that verify each returned row. Types that
        # accept every row must keep their own request; otherwise the first broad
        # type on a shared endpoint can absorb siblings and silently misclassify
        # policies.
        $batchByUrl = @{}
        $apiByUrl   = @{}

        foreach ($policyTypeObj in $policyTypes) {
            $batchObject = $null
            if ($useBatchAPI) { $batchObject = $policyTypeObj.GetListBatchObject($NameFilter) }
            if ($batchObject) {
                $key = if($policyTypeObj.VerifyObject) { $batchObject.url } else { "$($batchObject.url)|$($policyTypeObj.Id)" }
                if ($batchByUrl.ContainsKey($key)) {
                    [void]$batchByUrl[$key].Types.Add($policyTypeObj)
                }
                else {
                    $entry = [PSCustomObject]@{
                        BatchObject = $batchObject
                        Types       = [System.Collections.Generic.List[object]]@($policyTypeObj)
                    }
                    $batchByUrl[$key] = $entry
                    [void]$batchEntries.Add($entry)
                }
                continue
            }

            $listURL = $policyTypeObj.GetListURL($NameFilter)
            if (-not $listURL) { continue }
            $listKey = if($policyTypeObj.VerifyObject) { $listURL } else { "$listURL|$($policyTypeObj.Id)" }
            if ($apiByUrl.ContainsKey($listKey)) {
                [void]$apiByUrl[$listKey].Types.Add($policyTypeObj)
            }
            else {
                $entry = [PSCustomObject]@{
                    ListURL = $listURL
                    Types   = [System.Collections.Generic.List[object]]@($policyTypeObj)
                }
                $apiByUrl[$listKey] = $entry
                [void]$apiEntries.Add($entry)
            }
        }

        if ($SinglePage -eq $true) { $newPageObject = $true }
        else                       { $params.Add("AllPages", $true) }
    }
    elseif ($script:GraphPagingCache) {
        # Resume paging: cache stores the same {BatchObject,Types} / {ListURL,Types}
        # entries we built above, with BatchObject.url already set to the next-page link.
        foreach ($e in $script:GraphPagingCache.BatchTypes) { [void]$batchEntries.Add($e) }
        foreach ($e in $script:GraphPagingCache.APITypes)   { [void]$apiEntries.Add($e)   }
        # The nextLink carries any server-side filter, but the client-side
        # re-check below needs the original search text too.
        $NameFilter = $script:GraphPagingCache.NameFilter
        if ($Paging -eq "AllRemainingPages") { $params.Add("AllPages", $true) }
        $newPageObject = $true
    }
    else {
        Write-LogDebug "No more pages"
        return
    }

    $params.Add("TokenId", $TokenId)

    if ($newPageObject -eq $true) {
        $script:GraphPagingCache = [PSCustomObject]@{
            BatchObjects = [System.Collections.Generic.List[PSCustomObject]]::new()
            BatchTypes   = [System.Collections.Generic.List[object]]::new()
            APITypes     = [System.Collections.Generic.List[PSCustomObject]]::new()
            NameFilter   = $NameFilter
        }
    }
    else {
        $script:GraphPagingCache = $null
    }

    # A name filter can be rejected by an endpoint we have not probed (several
    # Intune endpoints answer 400 or even 500 to any $filter). Retrying that
    # request unfiltered turns "the search silently found nothing" into "the
    # search worked, just slower" - the client-side re-check at the end still
    # narrows the result. Only the first attempt retries, and never while
    # resuming paging (those URLs are nextLinks, not ones we can rebuild).
    $canRetryUnfiltered = ($NameFilter -and $PSCmdlet.ParameterSetName -ne "ObjectsPaging")

    $pendingBatchEntries = $batchEntries
    $batchAttempt = 0
    while ($pendingBatchEntries.Count -gt 0) {
        # Flatten entries into the batch-objects list passed to Invoke-GraphBatchRequest,
        # and build a sub-request-id -> entry lookup so we can fan rows back out.
        $batchObjects = [System.Collections.Generic.List[PSCustomObject]]::new()
        $entryById    = @{}
        foreach ($entry in $pendingBatchEntries) {
            [void]$batchObjects.Add($entry.BatchObject)
            $entryById["$($entry.BatchObject.id)"] = $entry
        }

        # -IncludedFailed: a rejected name filter comes back as a failed sub-result
        # (a 400 body from $batch, or status 0 with no body from a direct call) and
        # the retry below needs to see it. Without it the dispatcher dropped failed
        # results before this loop, so the unfiltered retry never fired.
        $batchResults = Invoke-GraphBatchRequest $batchObjects "Policy objects" @params -IncludedFailed
        $retryBatchEntries = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($batchResult in $batchResults) {
            $entry = $entryById["$($batchResult.ID)"]
            if (-not $entry) { continue }

            # Failed = no body at all (a direct call that returned nothing), an error body
            # (a $batch sub-result), or an explicit non-2xx status. A body with no status
            # property is a success - some result shapes never carried one.
            $listFailed = ((-not $batchResult.body) -or
                           ($batchResult.body.PSObject.Properties['error']) -or
                           ($batchResult.PSObject.Properties['Status'] -and [int]$batchResult.Status -ge 300))
            if ($canRetryUnfiltered -and $batchAttempt -eq 0 -and $listFailed) {
                $unfilteredUrl = $entry.Types[0].GetListBatchObject().url
                if ($unfilteredUrl -ne $entry.BatchObject.url) {
                    Write-Log "$($entry.Types[0].ID): endpoint rejected the name filter - retrying without it and filtering locally" 2
                    $entry.BatchObject.url = $unfilteredUrl
                    [void]$retryBatchEntries.Add($entry)
                    continue
                }
            }
            if ($listFailed) {
                Write-Log "$($entry.Types[0].ID): list request failed (status $($batchResult.Status)) - no objects for this type" 2
                continue
            }

            if ($batchResult.body.value) {
                foreach ($v in $batchResult.body.value) {
                    # Walk consumers in registration order; first CheckPolicy
                    # that accepts the row (GetObject returns non-null) wins.
                    foreach ($pt in $entry.Types) {
                        $tmpPolicy = $pt.GetObject($v)
                        if ($tmpPolicy) {
                            [void]$graphPolicies.Add($tmpPolicy)
                            break
                        }
                    }
                }
            }
            elseif ($entry.Types[0].SingleObject -and $batchResult.body -and -not $batchResult.body.PSObject.Properties['error']) {
                # Single-object endpoint: the body IS the object — no value array.
                foreach ($pt in $entry.Types) {
                    $tmpPolicy = $pt.GetObject($batchResult.body)
                    if ($tmpPolicy) {
                        [void]$graphPolicies.Add($tmpPolicy)
                        break
                    }
                }
            }

            if ($batchResult.body.'@odata.nextLink' -and -not $params.Contains("AllPages")) {
                $nextLink = $batchResult.body.'@odata.nextLink'
                # Use the first sibling's API for the offset slice — siblings share _API by definition.
                $apiPrefix = $entry.Types[0].API
                $entry.BatchObject.url = $nextLink.Substring($nextLink.IndexOf($apiPrefix))
                [void]$script:GraphPagingCache.BatchObjects.Add($entry.BatchObject)
                [void]$script:GraphPagingCache.BatchTypes.Add($entry)
            }
        }

        $pendingBatchEntries = $retryBatchEntries
        $batchAttempt++
    }

    # For APIs not supported in batch requests
    foreach ($entry in $apiEntries) {
        $responseContent = Invoke-MSGraphAPI -Url $entry.ListURL @params

        if ($canRetryUnfiltered -and -not $responseContent) {
            $unfilteredUrl = $entry.Types[0].GetListURL()
            if ($unfilteredUrl -ne $entry.ListURL) {
                Write-Log "$($entry.Types[0].ID): endpoint rejected the name filter - retrying without it and filtering locally" 2
                $responseContent = Invoke-MSGraphAPI -Url $unfilteredUrl @params
            }
        }

        $primaryId = $entry.Types[0].ID
        $extra = if ($entry.Types.Count -gt 1) { " (+$($entry.Types.Count - 1) sibling type(s))" } else { "" }
        Write-Log "Value return count for $primaryId$extra $(($responseContent.value | Measure-Object).Count)"
        foreach ($listObject in $responseContent.value) {
            foreach ($pt in $entry.Types) {
                $tmpPolicy = $pt.GetObject($listObject)
                if ($tmpPolicy) {
                    [void]$graphPolicies.Add($tmpPolicy)
                    break
                }
            }
        }
        if (-not $responseContent.value -and $entry.Types[0].SingleObject -and $responseContent) {
            # Single-object endpoint: the response IS the object — no value array.
            foreach ($pt in $entry.Types) {
                $tmpPolicy = $pt.GetObject($responseContent)
                if ($tmpPolicy) {
                    [void]$graphPolicies.Add($tmpPolicy)
                    break
                }
            }
        }
        if ($responseContent.'@odata.nextLink' -and -not $params.Contains("AllPages")) {
            $entry.ListURL = $responseContent.'@odata.nextLink'
            [void]$script:GraphPagingCache.APITypes.Add($entry)
        }
    }

    if ($script:GraphPagingCache -and $script:GraphPagingCache.BatchObjects.Count -eq 0 -and $script:GraphPagingCache.APITypes.Count -eq 0) {
        $script:GraphPagingCache = $null
    }

    if ($NameFilter) {
        # Re-check every row client-side. Not every endpoint honours the
        # server-side clause (PolicyType.SupportsNameFilter is $false for
        # deviceCompliancePolicies v1 and assignmentFilters), and a type whose
        # CheckPolicy claims a row from a coalesced sibling request may not have
        # been filtered at all.
        $matching = [System.Collections.Generic.List[IntunePolicyBase]]::new()
        foreach ($p in $graphPolicies) {
            # Same semantics as the server-side clause: case-insensitive substring.
            if ("$($p.Name)".IndexOf($NameFilter, [System.StringComparison]::InvariantCultureIgnoreCase) -ge 0) {
                [void]$matching.Add($p)
            }
        }
        $graphPolicies = $matching
    }

    if ($graphPolicies.Count -gt 0) {
        # Resolve the tenant id provider-agnostically. Get-TokenInfo only sees the
        # MSAL token registry, so on the MgGraph provider $tokenInfo is $null and
        # every policy ended up with TenantID = $null — which Export-GraphPolicy
        # then silently skipped at the `-not $policyObject.TenantID` guard, leading
        # to "Exported N policies" telemetry with zero files on disk.
        #
        # Get-OperationTokenInfo, not Get-TokenInfo: Graph runs this call against ONE
        # token, and 0 (the default parameter value, and the caller's spelling for
        # "the default token") means "no filter, every token" to Get-TokenInfo. With a
        # second tenant signed in, every listed policy was stamped with an ARRAY of
        # tenant ids and an array _TokenID, so the export masked the wrong tenant and
        # the per-tenant settings lookups keyed off a joined string.
        $tokenInfo  = Get-OperationTokenInfo $TokenId
        $tenantId   = if ($tokenInfo) { $tokenInfo.TenantID } else { $null }
        $tokenIdVal = if ($tokenInfo) { $tokenInfo.Id }       else { 0 }
        if (-not $tenantId) {
            try {
                $provider = Get-AuthProvider
                if ($provider) {
                    $userInfo = $provider.GetUserInfo($TokenId)
                    if ($userInfo -and $userInfo.TenantId) { $tenantId = $userInfo.TenantId }
                }
            } catch { }
        }
        if (-not $tenantId) {
            Write-Log "Get-GraphPolicies: could not resolve TenantID for export - policies will be missing this property" 2
        }
        foreach ($p in $graphPolicies) {
            $p.TenantID = $tenantId
            $p._TokenID = $tokenIdVal
        }

        if ($IncludeAssignments -eq $true) {
            Add-GraphPolicyAssignments -Policies $graphPolicies -TokenId $TokenId
        }
    }

    return $graphPolicies.ToArray()
}
