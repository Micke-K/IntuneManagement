# Documentation-only grouping for the tenant-default enrollment policies.
#
# deviceManagement/deviceEnrollmentConfigurations carries several policies that read
# as one enrollment-restrictions area, but each is its own policy TYPE here
# (EnrollmentLimit, EnrollmentRestrictions, EnrollmentStatusPage,
# WindowsHelloForBusiness, WindowsRestore). The document builds its second level from
# PolicyType.Title, so every one of those types got a heading of its own holding a
# single child - and because all five ship a tenant default named "All users and all
# devices", three of those children were indistinguishable from each other while the
# remaining two repeated their own heading word for word:
#
#   Windows Hello for Business        <- heading, from PolicyType.Title
#     Windows Hello for Business      <- the only child, same text
#
# Group the five under one heading, and title each child by its type's PolicyName
# ("Device limit restrictions", "Device platform restrictions", ...) so each entry
# says which policy it is.
#
# Deliberately documentation-only. The heading cannot come from _APITitle: that same
# property titles the app's navigation menu, so changing it there would rename the
# nav and the type's identity for every other consumer.

$script:_docEnrollmentGroup = [PSCustomObject]@{
    # Grouping key, not a real policy-type id. The engine groups the second level on
    # this value, so it must not collide with any PolicyType.Id or an unrelated type
    # would be merged into this heading.
    Id      = 'DocEnrollmentRestrictions'
    Title   = 'Enrollment Restrictions'
    TypeIds = @(
        'EnrollmentLimit'
        'EnrollmentRestrictions'
        'EnrollmentStatusPage'
        'WindowsHelloForBusiness'
        'WindowsRestore'
    )
}

# The documentation grouping a policy belongs to, or $null for everything else -
# which is every other policy type, so the engine keeps its normal per-type heading.
function Get-DocumentationTypeGroup {
    param($PolicyObject)

    $typeId = $null
    if ($PolicyObject -and $PolicyObject.PolicyType) { $typeId = [string]$PolicyObject.PolicyType.Id }
    if ($typeId -and $script:_docEnrollmentGroup.TypeIds -contains $typeId) {
        return $script:_docEnrollmentGroup
    }
    return $null
}

# The title a grouped tenant-default policy is documented under, or $null to keep the
# policy's own display name.
#
# Only the tenant default is retitled. A custom policy of the same type - a second
# enrollment status page, a per-platform "Block Android Device Administrator
# Enrollment" restriction - has a real name of its own and must keep it, otherwise
# several of them would collapse onto the same title.
#
# Default-ness is read from two signals for the same reason
# DeviceEnrollmentObject.GetFileName (Classes/IntuneEnrollmentClasses.ps1) uses two:
# priority is not guaranteed to be present on every payload, and the id form is only
# reliable once the id has been populated. A policy that is default in either sense
# is treated as the default.
function Get-DocumentationEnrollmentDefaultName {
    param($PolicyObject)

    if (-not (Get-DocumentationTypeGroup $PolicyObject)) { return $null }

    $obj = if ($PolicyObject.PSObject.Properties['JsonObject'] -and $PolicyObject.JsonObject) {
        $PolicyObject.JsonObject
    } else {
        $PolicyObject
    }

    $isDefault = ($obj.priority -eq 0) -or ([string]$PolicyObject.Id -match '_Default')
    if (-not $isDefault) { return $null }

    # PolicyName first - "Device limit restrictions" says more than the heading-shaped
    # "Enrollment Limit" - then Title, for a type that declares no _PolicyName.
    $policyName = [string]$PolicyObject.PolicyType.PolicyName
    if ([string]::IsNullOrWhiteSpace($policyName)) {
        $policyName = [string]$PolicyObject.PolicyType.Title
    }
    if ([string]::IsNullOrWhiteSpace($policyName)) { return $null }
    return $policyName
}

# The text an object will be titled with, resolvable BEFORE documentation runs.
#
# The engine sorts policies up front, but the retitling above happens per object
# while it is being documented, so sorting on .Name alone ordered the document by
# text the reader never sees - three "All users and all devices" siblings in
# arbitrary order. Sorting on this keeps the document and its table of contents in
# the same, stable order.
function Get-DocumentationSortName {
    param($PolicyObject)

    $name = Get-DocumentationEnrollmentDefaultName $PolicyObject
    if ($name) { return $name }
    return [string]$PolicyObject.Name
}
