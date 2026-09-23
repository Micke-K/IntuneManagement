# Default grid columns for the Intune Manager view, derived from what a policy
# type or group supports instead of a hand-typed list per class.
#
# Every group and type used to carry its own _ViewProperties string list. With no
# rule behind them the lists drifted: "Policy type" repeated itself on type views,
# Platform sat blank on types that have none and was missing on Settings Catalog
# where it matters most, two views bound properties that do not exist on a row,
# and Description - the second thing anyone searches for - was nowhere. The grid
# filter matches only the bound columns, so a column that is not shown is a field
# that cannot be searched.
#
# A class now declares capabilities (see IntunePolicyTypeBase / IntunePolicyGroupBase
# in Classes/IntuneBaseClasses.ps1) and the ViewProperties accessor calls
# Get-IntuneDefaultColumns. An explicit _ViewProperties still wins if a class sets
# one, so the escape hatch exists - nothing in the tree uses it, and
# Tests/ViewColumns.Tests.ps1 pins that.
#
# Column entry syntax is unchanged: "<bindingPath>[=<Header>]".

function Get-IntuneDefaultColumns
{
    param([Parameter(Mandatory = $true)]$Target)

    if($Target -is [IntunePolicyGroupBase]) { return ,(Get-IntuneGroupDefaultColumns $Target) }
    return ,(Get-IntuneTypeDefaultColumns $Target)
}

# Type view - every row is one type, so the generic Policy type column would
# repeat itself; the type's _SubTypeColumn takes its place when rows differ.
function Get-IntuneTypeDefaultColumns
{
    param([Parameter(Mandatory = $true)]$PolicyType)

    $cols = [System.Collections.Generic.List[string]]::new()
    $cols.Add("Name")
    if($PolicyType._SubTypeColumn) { $cols.Add($PolicyType._SubTypeColumn) }
    foreach($c in @($PolicyType._ExtraColumns)) { if($c) { $cols.Add($c) } }
    if($PolicyType._HasPlatform) { $cols.Add("Platform") }
    $cols.Add("Description")
    if($PolicyType._HasModified) { $cols.Add("LastModified=Modified") }
    $cols.Add("ID")
    return ,$cols.ToArray()
}

# Group view - rows of several types. A column earns its width only if most rows
# can fill it, so the decisions read the member types' flags.
function Get-IntuneGroupDefaultColumns
{
    param([Parameter(Mandatory = $true)]$PolicyGroup)

    $members = @($PolicyGroup._PolicyTypes | Where-Object { $null -ne $_ })

    # One member type: the group view IS that type's view.
    if($members.Count -eq 1) { return ,@($members[0].ViewProperties) }

    $cols = [System.Collections.Generic.List[string]]::new()
    $cols.Add("Name")

    # Policy type is the row's PolicyName: per row where the object class sets one
    # (Device Configuration "iOS Wi-Fi", Endpoint Security "Antivirus"), else the
    # member type's own name. Shown when there is more than one member, unless the
    # group says otherwise (Applications: Type carries it).
    $showType = if($null -ne $PolicyGroup._ShowPolicyTypeColumn) { [bool]$PolicyGroup._ShowPolicyTypeColumn } else { $members.Count -gt 1 }
    if($showType) { $cols.Add("PolicyName=Policy type") }

    foreach($c in @($PolicyGroup._ExtraColumns)) { if($c) { $cols.Add($c) } }

    # Platform: the group can force or hide it; otherwise at least half the members
    # resolve one, and it is not constant - every member carrying the same
    # type-level platform would read "Windows" on every row.
    if($null -ne $PolicyGroup._ShowPlatformColumn)
    {
        if([bool]$PolicyGroup._ShowPlatformColumn) { $cols.Add("Platform") }
    }
    elseif($members.Count -gt 0)
    {
        $withPlatform = @($members | Where-Object { $_._HasPlatform }).Count
        if(($withPlatform * 2) -ge $members.Count)
        {
            $typeLevel = @($members | ForEach-Object { [string]$_._PlatformName })
            $allTyped  = (@($typeLevel | Where-Object { -not $_ }).Count -eq 0)
            $constant  = $allTyped -and (@($typeLevel | Select-Object -Unique).Count -eq 1)
            if(-not $constant) { $cols.Add("Platform") }
        }
    }

    $cols.Add("Description")

    if(@($members | Where-Object { $_._HasModified }).Count -gt 0) { $cols.Add("LastModified=Modified") }

    $cols.Add("ID")
    return ,$cols.ToArray()
}

# Graph enum values as grid text: "boundAndValidated" -> "Bound and validated",
# "superseded" -> "Superseded". Used by the row properties that normalise
# template lifecycle / token state across types so a group column reads evenly.
function ConvertTo-DisplayWords
{
    param([string]$Value)

    if([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $spaced = [regex]::Replace($Value.Trim(), '(?<=[a-z0-9])([A-Z])', ' $1')
    return $spaced.Substring(0, 1).ToUpperInvariant() + $spaced.Substring(1).ToLowerInvariant()
}
