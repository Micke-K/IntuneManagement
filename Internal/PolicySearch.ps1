# Shared plumbing behind the policy-picker dialogs in both UI backends.
#
# The picker used to work off a full-tenant sweep: every Add Pair click ran
# Get-GraphPolicies once per policy group, which took minutes on a real tenant.
# Instead the user now picks a scope (one policy group or one policy type) and
# types a name; the search goes to Graph as a server-side prefix filter and
# comes back with just the matches.
#
# Lives in Internal/ rather than UI/ so WPF and Avalonia agree on the scope
# list and the search call, and so the UI stays a thin caller (R9/R10).

# Minimum characters before a search is sent. Below this the prefix matches
# most of the tenant, which is the slow full sweep this exists to avoid. Users
# who really want everything in the scope click "Load all in scope" instead.
$script:PolicySearchMinLength = 2

function Get-PolicySearchMinLength
{
    return $script:PolicySearchMinLength
}

# The scope list for the picker's "Search in" dropdown: every policy group,
# each followed by its own policy types. Groups are listed first so picking
# "everything in Configuration" is one click; the indented type rows narrow it
# to a single API.
#
# Key is what the UI binds SelectedValue to - Id alone is not unique because a
# group and a type can carry the same Id (e.g. Applications).
function Get-PolicySearchScopes
{
    param(
        # Marks the scope for this policy type as the default selection. The
        # compare flows pass the source policy's type, which is the scope the
        # user wants almost every time.
        [string]$DefaultPolicyTypeId
    )

    $scopes = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach($group in @($script:IntuneGroups | Where-Object { $_.Title } | Sort-Object Title))
    {
        [void]$scopes.Add([PSCustomObject]@{
            Key       = "Group|$($group.Id)"
            Id        = $group.Id
            Kind      = "Group"
            Title     = "$($group.Title) (all types)"
            GroupId   = $group.Id
            IsDefault = $false
        })

        foreach($type in @($group.PolicyTypes | Where-Object { $_.Title -and $_.API } | Sort-Object Title))
        {
            [void]$scopes.Add([PSCustomObject]@{
                Key       = "Type|$($type.Id)"
                Id        = $type.Id
                Kind      = "Type"
                # Leading spaces indent the type under its group row. ASCII
                # only - see the string-literal gate in Static.Tests.ps1.
                Title     = "    $($type.Title)"
                GroupId   = $group.Id
                IsDefault = ($DefaultPolicyTypeId -and $type.Id -eq $DefaultPolicyTypeId)
            })
        }
    }

    return $scopes.ToArray()
}

# The tenants the picker can search, one entry per live token - same source and
# labelling as the Copy dialog's destination-tenant combo (Get-TokenInfo, which
# is provider-agnostic, so MSAL / OAuth / MgGraph tokens all appear). Returns an
# empty array when only one tenant is signed in; callers hide the combo then.
#
# Searching another tenant works because Get-GraphPolicies stamps each policy
# with the TokenId it came from, and GetFullObject / Invoke-PolicyHydrate resolve
# against that stamp - so a policy picked from tenant B still hydrates and
# compares against tenant B.
function Get-PolicySearchTenants
{
    $tokens = @(Get-TokenInfo)
    if($tokens.Count -le 1) { return @() }

    $tenants = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach($token in $tokens)
    {
        if(-not $token) { continue }

        $title = Get-StringOrDefault $token.TenantName $token.TenantID
        if([String]::IsNullOrWhiteSpace($title)) { $title = "Tenant $($token.Id)" }
        if($token.IsDefault -eq $true) { $title = "$title (Current)" }

        [void]$tenants.Add([PSCustomObject]@{
            Key       = "Token|$($token.Id)"
            TokenId   = [int]$token.Id
            TenantId  = $token.TenantID
            Title     = $title
            IsDefault = ($token.IsDefault -eq $true)
        })
    }

    return $tenants.ToArray()
}

# Key of the tenant to preselect: the current one, else the first.
function Get-PolicySearchDefaultTenantKey
{
    param($Tenants)

    $all = @($Tenants)
    if($all.Count -eq 0) { return $null }

    $current = $all | Where-Object { $_.IsDefault } | Select-Object -First 1
    if($current) { return $current.Key }

    return $all[0].Key
}

# Default scope key for a policy type id: the type's own row when it is in the
# list, otherwise its group, otherwise the first scope. Returns $null when
# there are no scopes at all.
function Get-PolicySearchDefaultScopeKey
{
    param(
        $Scopes,
        [string]$PolicyTypeId
    )

    $all = @($Scopes)
    if($all.Count -eq 0) { return $null }

    if($PolicyTypeId)
    {
        $typeScope = $all | Where-Object { $_.Kind -eq "Type" -and $_.Id -eq $PolicyTypeId } | Select-Object -First 1
        if($typeScope) { return $typeScope.Key }
    }

    return $all[0].Key
}

# Run one scoped search. $Scope is an entry from Get-PolicySearchScopes; an
# empty $SearchText returns everything in the scope (the "Load all in scope"
# path). Types whose endpoint rejects a server-side filter are handled inside
# Get-GraphPolicies, which always re-checks names client-side.
function Search-IntunePolicies
{
    param(
        [Parameter(Mandatory = $true)]
        $Scope,
        [string]$SearchText,
        # Ids to drop from the result - used to keep a policy from being
        # compared against itself.
        [string[]]$ExcludeIds,
        [Int]$TokenId = (Get-DefaultTokenId)
    )

    if(-not $Scope) { return @() }

    $params = @{ TokenId = $TokenId }
    if($Scope.Kind -eq "Group") { $params.Add("PolicyGroup", $Scope.Id) }
    else                        { $params.Add("PolicyType",  $Scope.Id) }
    if($SearchText) { $params.Add("NameFilter", $SearchText) }

    $found = @()
    try {
        $found = @(Get-GraphPolicies @params)
    }
    catch {
        Write-LogError "Search-IntunePolicies failed for scope $($Scope.Key)" $_.Exception
        return @()
    }

    if($ExcludeIds -and $ExcludeIds.Count -gt 0) {
        $found = @($found | Where-Object { $ExcludeIds -notcontains $_.Id })
    }

    return $found
}
