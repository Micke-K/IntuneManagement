function Initialize-IntuneAssignmentsModule
{
    $script:intuneAssignmentsProviders = @(
        [IntuneAssignmentsFolderProvider]::new()
        [IntuneAssignmentsIntuneProvider]::new()
    )
}

function Test-IntuneAssignmentsSupported
{
    <#
    .SYNOPSIS
    True when a policy type can carry assignments at all.

    .DESCRIPTION
    Types whose Graph entity has no assignments navigation property (Conditional
    Access, Terms of Use, Named Locations, Filters, Role Definitions, ...) declare
    _SupportsAssignments = $false. Listing them in the Assignments tool produced a
    row per policy with an empty assignment set, which is noise.

    Only an explicit $false excludes a type. PolicyType is duck-typed in places
    (tests pass a PSCustomObject), and a missing property must not silently hide a
    type that really does support assignments - same rule as
    Add-GraphPolicyAssignments.
    #>
    param($PolicyType)

    if(-not $PolicyType) { return $false }
    return ($PolicyType.SupportsAssignments -ne $false)
}

function Get-IntuneAssignmentPolicyTypes
{
    <#
    .SYNOPSIS
    The policy types the Assignments tool should look at.
    #>
    param($PolicyTypes = $script:IntuneTypes)

    return @($PolicyTypes | Where-Object { Test-IntuneAssignmentsSupported $_ })
}

function Get-IntuneAssignmentViewOptions
{
    <#
    .SYNOPSIS
    The View choices offered above the Assignments list.

    .DESCRIPTION
    Returned as plain Name/Value objects so both UI backends can adapt them:
    WPF binds them directly (DisplayMemberPath/SelectedValuePath), Avalonia wraps
    them in [NameValueComboItem] because its binder needs a real CLR type.
    The first entry is the default.
    #>

    return @(
        [PSCustomObject]@{ Name = "Show all";                          Value = "All" }
        [PSCustomObject]@{ Name = "Show assigned policies only";       Value = "Assigned" }
        [PSCustomObject]@{ Name = "Show policies with no assignments"; Value = "Unassigned" }
    )
}

function Test-IntuneAssignmentViewMatch
{
    <#
    .SYNOPSIS
    True when a row belongs in the currently selected View.

    .DESCRIPTION
    An unknown or empty view shows everything, so a UI that has not populated its
    combo yet (or a saved value from a later version) never blanks the list.
    #>
    param($Row, [string]$View)

    if(-not $Row) { return $false }

    $count = 0
    if($null -ne $Row.AssignmentCount) { $count = [int]$Row.AssignmentCount }

    switch($View)
    {
        "Assigned"   { return ($count -gt 0) }
        "Unassigned" { return ($count -le 0) }
        default      { return $true }
    }

    return $true
}

function Get-IntuneAssignmentRow
{
    param(
        $Policy,
        [hashtable]$GroupsById = $null,
        [hashtable]$FiltersById = $null,
        [hashtable]$GroupCache = $null,
                   $TokenId    = $null
    )

    if(-not $Policy -or -not $Policy.Object) { return $null }

    $nameProp = if($Policy.PolicyType -and $Policy.PolicyType.NameProperty) { $Policy.PolicyType.NameProperty } else { "displayName" }
    $name     = $Policy.Object.$nameProp
    if([string]::IsNullOrWhiteSpace("$name") -and $Policy.PSObject.Properties['Name']) { $name = $Policy.Name }

    $type = "$($Policy.Object.'@OData.Type')"
    if($type) { $type = $type.Split('.')[-1] }
    elseif($Policy.PolicyType -and $Policy.PolicyType.Title) { $type = $Policy.PolicyType.Title }
    else { $type = "" }

    $included = @()
    $excluded = @()
    $includedFilters = @()
    $excludedFilters = @()

    foreach($assignment in @($Policy.Object.assignments))
    {
        if(-not $assignment -or -not $assignment.target) { continue }

        $targetType = "$($assignment.target.'@odata.type')"
        $resolved   = $null
        $isExcluded = $false

        switch($targetType)
        {
            "#microsoft.graph.groupAssignmentTarget"          { $resolved = Resolve-IntuneAssignmentGroup $assignment.target.groupId $GroupsById $GroupCache $TokenId }
            "#microsoft.graph.exclusionGroupAssignmentTarget" { $resolved = Resolve-IntuneAssignmentGroup $assignment.target.groupId $GroupsById $GroupCache $TokenId; $isExcluded = $true }
            "#microsoft.graph.allDevicesAssignmentTarget"     { $resolved = "All Devices" }
            "#microsoft.graph.allLicensedUsersAssignmentTarget" { $resolved = "All Users" }
            default
            {
                if($assignment.target.groupId) { $resolved = Resolve-IntuneAssignmentGroup $assignment.target.groupId $GroupsById $GroupCache $TokenId }
            }
        }

        if($null -eq $resolved) { continue }

        if($isExcluded) { $excluded += $resolved } else { $included += $resolved }

        $filterId = [string]$assignment.target.deviceAndAppManagementAssignmentFilterId
        if(Test-AssignmentFilterDefined $filterId) {
            $filterName = Resolve-IntuneAssignmentFilter $filterId $FiltersById
            $filterType = [string]$assignment.target.deviceAndAppManagementAssignmentFilterType
            $filterLabel = if($filterType) { "$filterName ($filterType)" } else { $filterName }

            if($filterType -eq "exclude") { $excludedFilters += $filterLabel }
            else { $includedFilters += $filterLabel }
        }
    }

    $included = @($included | Where-Object { $_ } | Select-Object -Unique)
    $excluded = @($excluded | Where-Object { $_ } | Select-Object -Unique)
    $includedFilters = @($includedFilters | Where-Object { $_ } | Select-Object -Unique)
    $excludedFilters = @($excludedFilters | Where-Object { $_ } | Select-Object -Unique)

    return [PSCustomObject]@{
        Object               = $Policy.Object
        Id                   = $Policy.Object.id
        Name                 = $name
        Type                 = $type
        AssignmentCount      = @($Policy.Object.assignments).Count
        HasFilters           = ($includedFilters.Count -gt 0 -or $excludedFilters.Count -gt 0)
        Included             = $included
        Excluded             = $excluded
        IncludedFilters      = $includedFilters
        ExcludedFilters      = $excludedFilters
        IncludedString       = ($included -join "; ")
        ExcludedString       = ($excluded -join "; ")
        IncludedFilterString = ($includedFilters -join "; ")
        ExcludedFilterString = ($excludedFilters -join "; ")
    }
}

function Resolve-IntuneAssignmentGroup
{
    param(
        [string]   $GroupId,
        [hashtable]$GroupsById,
        [hashtable]$GroupCache,
                   $TokenId
    )

    if([string]::IsNullOrWhiteSpace($GroupId)) { return $null }

    if($GroupsById -and $GroupsById.ContainsKey($GroupId)) { return $GroupsById[$GroupId] }

    if($null -ne $GroupCache -and $GroupCache.ContainsKey($GroupId)) { return $GroupCache[$GroupId] }

    if($TokenId)
    {
        try
        {
            $group = Invoke-MSGraphAPI -Url "groups/$($GroupId)?`$select=id,displayName" -ODataMetadata "skip" -TokenId $TokenId
            if($group -and $group.displayName)
            {
                if($GroupCache -ne $null) { $GroupCache[$GroupId] = $group.displayName }
                return $group.displayName
            }
        }
        catch
        {
            Write-Log "Could not resolve group with ID $GroupId" 2
        }
    }
    else
    {
        Write-Log "Could not find a group with ID $GroupId" 2
    }

    return $GroupId
}

# Intune's "no assignment filter" sentinel.
#
# Graph writes an all-zeros filter id on assignment targets, including on types the
# portal offers no filter UI for at all (enrollment notifications, for one), and it is
# usually paired with filterType 'exclude'. Taken at face value that reads as a real
# filter, so documentation reported a filter literally named
# 00000000-0000-0000-0000-000000000000 in Exclude mode against an assignment that has
# no filter. It is never a real filter id.
$script:NoAssignmentFilterId = '00000000-0000-0000-0000-000000000000'

# $true only for a filter id that actually identifies a filter. The sentinel and an
# absent/blank value are the same thing - "not defined" - so every caller can ask this
# one question instead of re-deriving it (the answer was already open-coded in the
# documentation batch lookup and the migration walk, and missing everywhere else).
function Test-AssignmentFilterDefined
{
    [OutputType([bool])]
    param($FilterId)

    $id = [string]$FilterId
    if([string]::IsNullOrWhiteSpace($id)) { return $false }
    return $id.Trim() -ne $script:NoAssignmentFilterId
}

function Resolve-IntuneAssignmentFilter
{
    param(
        [string]$FilterId,
        [hashtable]$FiltersById
    )

    if([string]::IsNullOrWhiteSpace($FilterId)) { return $null }
    if($FiltersById -and $FiltersById.ContainsKey($FilterId)) { return $FiltersById[$FilterId] }
    return "<unresolved: $FilterId>"
}

function Get-IntuneAssignmentFiltersFromFolder
{
    param([string]$Path)

    $filtersById = @{}
    if(-not $Path -or -not [IO.Directory]::Exists($Path)) { return $filtersById }

    $filterType = $null
    if($script:IntuneTypes) {
        $filterType = $script:IntuneTypes | Where-Object Id -eq 'AssignmentFilters' | Select-Object -First 1
    }

    $searchFolders = @()
    if($filterType -and $filterType.Folder) {
        $typeFolder = [IO.Path]::Combine($Path, $filterType.Folder)
        if([IO.Directory]::Exists($typeFolder)) { $searchFolders += $typeFolder }
    }
    if($searchFolders.Count -eq 0) {
        $searchFolders += $Path
    }

    foreach($folder in $searchFolders) {
        foreach($file in [IO.Directory]::EnumerateFiles($folder, "*.json", [IO.SearchOption]::AllDirectories))
        {
            try {
                $obj = ConvertFrom-Json ([IO.File]::ReadAllText($file))
                $odataType = "$($obj.'@odata.type')$($obj.'@OData.Type')"
                if($obj.id -and $obj.displayName -and (
                    $odataType -match 'deviceAndAppManagementAssignmentFilter' -or
                    $obj.PSObject.Properties['assignmentFilterManagementType'])) {
                    $filtersById[$obj.id] = $obj.displayName
                }
            }
            catch {
                Write-Log "Failed to parse assignment filter file $file" 2
            }
        }
    }

    return $filtersById
}

function Get-IntuneAssignmentFiltersFromIntune
{
    param([int]$TokenId)

    $filtersById = @{}
    try {
        $resp = Invoke-MSGraphAPI -Url "deviceManagement/assignmentFilters?`$select=id,displayName" -TokenId $TokenId -AllPages
        $filters = @()
        if($resp -and $resp.value) { $filters = @($resp.value) }
        elseif($resp -is [Array]) { $filters = @($resp) }

        foreach($filter in $filters) {
            if($filter.id -and $filter.displayName) { $filtersById[$filter.id] = $filter.displayName }
        }
    }
    catch {
        Write-Log "Could not load assignment filters" 2
    }

    return $filtersById
}

function Get-IntuneAssignmentGroupIds
{
    param($Policies)

    $ids = [System.Collections.Generic.HashSet[string]]::new()
    foreach($policy in @($Policies)) {
        foreach($assignment in @($policy.Object.assignments)) {
            $groupId = [string]$assignment.target.groupId
            if(-not [string]::IsNullOrWhiteSpace($groupId)) { [void]$ids.Add($groupId) }
        }
    }
    return @($ids)
}

function Get-IntuneAssignmentGroupsFromIntune
{
    param(
        [string[]]$GroupIds,
        [int]$TokenId
    )

    $groupsById = @{
        "adadadad-808e-44e2-905a-0b7873a8a531" = "All Devices"
        "acacacac-9df4-4c7d-9d50-4ef0226f57a9" = "All Users"
    }

    $toLookup = @($GroupIds | Where-Object { $_ -and -not $groupsById.ContainsKey($_) } | Select-Object -Unique)
    if($toLookup.Count -eq 0) { return $groupsById }

    $batch = [System.Collections.Generic.List[PSCustomObject]]::new()
    $idx = 0
    foreach($groupId in $toLookup) {
        $idx++
        [void]$batch.Add([PSCustomObject]@{
            id      = "grp_$idx"
            method  = "GET"
            url     = "groups/$groupId/?`$select=displayName,id"
            headers = @{ "Accept" = "application/json" }
        })
    }

    $idByRequest = @{}
    $idx = 0
    foreach($groupId in $toLookup) {
        $idx++
        $idByRequest["grp_$idx"] = $groupId
    }

    try {
        $results = @(Invoke-GraphBatchRequest -BatchObjects $batch -BatchType "AssignmentGroupNames" -TokenId $TokenId -SkipWarnings -IncludedFailed)
        foreach($result in $results) {
            $groupId = $idByRequest["$($result.Id)"]
            if(-not $groupId) { continue }
            if($result.Status -ge 200 -and $result.Status -lt 300 -and $result.body -and $result.body.displayName) {
                $groupsById[$groupId] = $result.body.displayName
            }
            else {
                $groupsById[$groupId] = "<unresolved: $groupId>"
            }
        }
    }
    catch {
        Write-Log "Could not batch resolve assignment groups" 2
    }

    return $groupsById
}

function Get-IntuneAssignmentsFromFolder
{
    param([string]$Path)

    Write-Status "Loading assignments from folder"

    $rows = @()
    if(-not $Path -or -not [IO.Directory]::Exists($Path)) { return $rows }

    Save-SettingStoreValue "IntuneAssignments" "ExportPath" $Path
    Save-SettingStoreValue "" "LastUsedFullPath" $Path

    $groupsById  = @{}
    $filtersById = Get-IntuneAssignmentFiltersFromFolder $Path
    $groupsPath  = [IO.Path]::Combine($Path, "Groups")
    if([IO.Directory]::Exists($groupsPath))
    {
        foreach($file in [IO.Directory]::EnumerateFiles($groupsPath, "*.json"))
        {
            try
            {
                $g = ConvertFrom-Json ([IO.File]::ReadAllText($file))
                if($g.id -and $g.displayName) { $groupsById[$g.id] = $g.displayName }
            }
            catch
            {
                Write-Log "Failed to parse group file $file" 2
            }
        }
    }

    if(-not $script:IntuneTypes)
    {
        Write-Log "Intune types are not initialized. Cannot load policies from folder." 3
        return $rows
    }

    foreach($policyType in (Get-IntuneAssignmentPolicyTypes))
    {
        $typeFolder = [IO.Path]::Combine($Path, $policyType.Folder)
        if(-not [IO.Directory]::Exists($typeFolder)) { continue }

        Write-Status "Read $($policyType.Title)" -Force -SkipLog

        $policies = @(Get-PoliciesFromFolder -Path $typeFolder -PolicyTypes @($policyType))
        foreach($policy in $policies)
        {
            $row = Get-IntuneAssignmentRow -Policy $policy -GroupsById $groupsById -FiltersById $filtersById
            if($row) { $rows += $row }
        }
    }

    return $rows
}

function Get-IntuneAssignmentsFromIntune
{
    Write-Status "Loading assignments from Intune"

    $rows = @()
    if(-not $script:IntuneTypes)
    {
        Write-Log "Intune types are not initialized. Cannot load policies from Intune." 3
        return $rows
    }

    $tokenId = Get-DefaultTokenId

    # Ask Graph only for the types that can carry assignments. Every policy group
    # also contains unassignable types (Conditional Access, Terms of Use, ...), so
    # -PolicyGroup would list those too and we would discard the rows afterwards:
    # one wasted call each, and some of those endpoints answer 4xx/5xx, which then
    # shows up as batch errors in the Graph Calls view
    # (identityGovernance/termsOfUse/agreements returns 500).
    # Equivalent to the old group expansion: every registered type belongs to a
    # group, so this is that same set minus the unassignable types.
    $typeIds = @(Get-IntuneAssignmentPolicyTypes | Where-Object { $_.Id } | ForEach-Object { $_.Id })

    if($typeIds.Count -eq 0) { return $rows }

    $policies = @(Get-GraphPolicies -PolicyType $typeIds -TokenId $tokenId -IncludeAssignments)
    $groupsById = Get-IntuneAssignmentGroupsFromIntune -GroupIds (Get-IntuneAssignmentGroupIds -Policies $policies) -TokenId $tokenId
    $filtersById = Get-IntuneAssignmentFiltersFromIntune -TokenId $tokenId

    foreach($policy in $policies)
    {
        # Get-GraphPolicies deliberately lists every type (bulk export needs the
        # unassignable ones too - see Public/Get-GraphPolicies.ps1), so filter here
        # rather than at the query.
        if(-not (Test-IntuneAssignmentsSupported $policy.PolicyType)) { continue }

        $row = Get-IntuneAssignmentRow -Policy $policy -GroupsById $groupsById -FiltersById $filtersById -TokenId $tokenId
        if($row) { $rows += $row }
    }

    return $rows
}

# Moved from Public/Get-GraphPolicies.ps1 (architecture R11): assignment-
# expansion helper for the policy list path. Module-internal, not exported.
function Add-GraphPolicyAssignments
{
    param(
        $Policies,
        [int]$TokenId
    )

    if (-not $Policies -or $Policies.Count -eq 0) { return }

    # Build batch for policies whose type didn't get assignments expanded inline.
    $assignmentBatch = [System.Collections.Generic.List[PSCustomObject]]::new()
    $policyById      = @{}
    $idCounter       = 0

    foreach ($policy in $Policies) {
        if (-not $policy.PolicyType -or $policy.PolicyType.SupportsAssignments -eq $false) { continue }
        if (-not $policy.Id) { continue }
        # Already populated by inline $expand=assignments; skip.
        # An empty array (@()) is also considered populated -- don't re-fetch.
        if ($null -ne $policy.Object.assignments) { continue }

        # Polymorphic list collections need a per-object base (App protection routes
        # to the concrete platform collection); $null means this object can't carry
        # assignments at all. PolicyType is duck-typed in places (tests pass a
        # PSCustomObject), so fall back to the plain API when the hook is absent.
        $assignmentBase = if ($policy.PolicyType.PSObject.Methods['GetAssignmentsBaseURL']) {
            $policy.PolicyType.GetAssignmentsBaseURL($policy)
        } else {
            $policy.PolicyType.API
        }
        if (-not $assignmentBase) { continue }

        $idCounter++
        $reqId  = "asn_$idCounter"
        # Some types (policySets) return 400 on the /assignments navigation
        # GET; only the single-object $expand variant works for them.
        if ($policy.PolicyType.AssignmentsViaExpand -eq $true) {
            $apiUrl = "$assignmentBase/$($policy.Id)?`$expand=assignments".TrimStart('/')
        }
        else {
            $apiUrl = "$assignmentBase/$($policy.Id)/assignments".TrimStart('/')
        }
        [void]$assignmentBatch.Add([PSCustomObject]@{
            id      = $reqId
            method  = "GET"
            url     = $apiUrl
            headers = @{ "Accept" = "application/json;odata.metadata=minimal" }
        })
        $policyById[$reqId] = $policy
    }

    if ($assignmentBatch.Count -eq 0) { return }

    Write-Log "Loading assignments for $($assignmentBatch.Count) policies"
    $batchResults = Invoke-GraphBatchRequest $assignmentBatch "Policy assignments" -TokenId $TokenId -AllPages

    foreach ($result in $batchResults) {
        $policy = $policyById["$($result.Id)"]
        if (-not $policy -or -not $policy.Object) { continue }

        # Expand-variant responses carry the assignments on the object body;
        # navigation GETs return them as the value collection.
        $value = if ($policy.PolicyType.AssignmentsViaExpand -eq $true) { $result.body.assignments } else { $result.body.value }
        if ($null -eq $value) { $value = @() } else { $value = @($value) }

        # PSCustomObject from ConvertFrom-Json: assigning a property creates it if missing.
        $policy.Object | Add-Member -MemberType NoteProperty -Name "assignments" -Value $value -Force
    }
}

# Moved from Internal/MSGraph.ps1 (architecture R4 - MSGraph.ps1 should hold
# only generic Graph helpers; these are assignment-feature functions).
#region Assignments functions
function Import-GraphObjectAssignment
{
    param($PolicyObject, $SourceObject, [switch]$CopyAssignments)

    # Honour the ImportAssignments setting for normal imports/updates. Explicit
    # copy paths (Replace) pass -CopyAssignments and always import assignments.
    if(-not $CopyAssignments -and (Get-SettingValue "ImportAssignments") -ne $true) { return }

    $assignments = $SourceObject.JsonObject.assignments

    if(($assignments | Measure-Object).Count -eq 0) { return }

    $preConfig = $null
    $clonedAssignments = $assignments | ConvertTo-Json -Depth 50 | ConvertFrom-Json

    $preConfig = $PolicyObject.PolicyType.PreImportAssignmentsCommand($PolicyObject, $SourceObject)

    if($preConfig -isnot [Hashtable]) { $preConfig = @{} }

    if($preConfig["Import"] -eq $false) { return } # Assignment managed manually so skip further processing

    $api = ?? $preConfig["API"] "$($PolicyObject.PolicyType.API)/$($PolicyObject.Id)/$($PolicyObject.PolicyType.AssignAction)"

    $method = ?? $preConfig["Method"] "POST"

    $clonedAssignments = ?? $preConfig["Assignments"] $clonedAssignments

    $keepProperties = ?? $PolicyObject.PolicyType.AssignmentPropertiesToKeep @("target")
    $keepTargetProperties = ?? $PolicyObject.PolicyType.AssignmentTargetPropertiesToKeep @("@odata.type","groupId","deviceAndAppManagementAssignmentFilterId","deviceAndAppManagementAssignmentFilterType")
    
    $ObjectAssignments = @()
    foreach($assignment in $clonedAssignments)
    {
        if(($assignment.target.UserId -and $CopyAssignments -ne $true) -or ($assignment.Source -and $assignment.Source -ne "direct"))
        {
            # E.g. Source could be PolicySet...so should not be added here
            continue 
        }

        # Only blank the id when it still exists. Remove-GraphPropertiesForImport
        # recurses into assignment child objects during import/copy prep and strips
        # 'id' (it is in the default remove list); the same source object is then
        # reused here, so 'id' may already be gone. Setting a missing property on a
        # PSCustomObject throws ("The property 'Id' cannot be found..."), which broke
        # copy/import of any policy that has assignments when ImportAssignments is on.
        # (The loop below strips id anyway unless a type keeps it, so this is only to
        # normalise a kept id to empty.)
        if($assignment.PSObject.Properties['Id']) { $assignment.Id = "" }
        foreach($prop in $assignment.PSObject.Properties)
        {
            if($prop.Name -in $keepProperties) { continue }
            Remove-Property $assignment $prop.Name
        }

        foreach($prop in $assignment.target.PSObject.Properties)
        {
            if($prop.Name -in $keepTargetProperties) { continue }
            Remove-Property $assignment.target $prop.Name
        }
        
        $ObjectAssignments += $assignment
    }

    if($ObjectAssignments.Count -eq 0) { return } # No "Direct" assignments

    $htAssignments = @{}
    $htAssignments.Add($PolicyObject.PolicyType.AssignmentsType, @($ObjectAssignments))

    $json = $htAssignments | ConvertTo-Json -Depth 50
    if($CopyAssignments -ne $true)
    {
        # Translation context must be the SOURCE object: $PolicyObject here is the
        # freshly-created target object (TenantId = TARGET tenant, no FileInfo, no
        # _ClonedFromObject), so Update-JsonForEnvironment could never locate the
        # export root or detect cross-tenant - group/dependency translation silently
        # no-oped on the assignment path. $SourceObject is the file-loaded original
        # with the migration-table TenantID and the export FileInfo.
        $json = Update-JsonForEnvironment $json $SourceObject $PolicyObject.TokenId
    }

    $importedAssignments = Invoke-MSGraphAPI $api -HttpMethod $method -Content $json -TokenId $PolicyObject.TokenId

    $PolicyObject.PolicyType.PostImportAssignmentsCommand($PolicyObject, $SourceObject, $importedAssignments)
}

function Add-GraphAssignmentsToExportFile
{
    param($PolicyObject, $FileName)

    $exportAssignments = Get-CacheObject "CurrentExportAssignments" (Get-SettingValue "ExportAssignments")
    if($exportAssignments -ne $true) { return }

    if([IO.File]::Exists($FileName) -eq $false)
    {
        Write-Log "File not found: $FileName. Could not add assignments to file" 3
        return
    }
    
    $tmpObj = Get-GraphObjectFromFile $FileName

    # See Add-GraphPolicyAssignments: polymorphic list collections need a per-object
    # base URL, and $null means the object cannot carry assignments.
    $assignmentBase = if($PolicyObject.PolicyType.PSObject.Methods['GetAssignmentsBaseURL']) {
        $PolicyObject.PolicyType.GetAssignmentsBaseURL($PolicyObject)
    } else {
        $PolicyObject.PolicyType.API
    }
    if(-not $assignmentBase) { return }

    if($PolicyObject.PolicyType.AssignmentsViaExpand -eq $true) {
        # Types whose /assignments navigation GET 400s (policySets) - fetch
        # via the single-object $expand variant instead.
        $url = "$assignmentBase/$($PolicyObject.id)?`$expand=assignments"
        $assignments = (Invoke-MSGraphAPI -Url $url -ODataMetadata "Minimal" -TokenId $PolicyObject.TokenId).assignments
    }
    else {
        $url = "$assignmentBase/$($PolicyObject.id)/assignments"
        $assignments = (Invoke-MSGraphAPI -Url $url -ODataMetadata "Minimal" -TokenId $PolicyObject.TokenId).Value
    }
    if($assignments)
    {
        if(-not ($tmpObj.PSObject.Properties | Where-Object Name -eq "assignments"))
        {
            $tmpObj | Add-Member -MemberType NoteProperty -Name "assignments" -Value $assignments
        }
        else
        {
            $tmpObj.Assignments = $assignments
        }
        Save-GraphObjectToFile $tmpObj $FileName
    }
}
