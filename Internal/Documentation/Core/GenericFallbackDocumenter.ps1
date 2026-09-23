# Generic fallback documenter.
#
# Documents policy objects that match NO handler and NO input provider (the
# engine's 'NoProvider' path) instead of producing an empty stub. Output is a
# deliberately simple, schema-less dump:
#   - standard basic-info rows (Name / Description / Platform / Profile type /
#     Created / Modified / Version) via the shared Add-Basic* helpers
#   - one settings row per non-internal property + value
#   - object-valued properties recurse into named sub-levels: the parent name
#     becomes the Category (depth 1) then SubCategory (depth 2). Deeper objects,
#     and any array-of-objects, are emitted as a single compact-JSON value (the
#     output model only has two named levels; HTML/MD/Word still indent by Level).
#   - arrays of scalars are joined; empty arrays / null / empty values are skipped
#   - secret-looking properties are redacted; very long strings are truncated
#
# Scope tags + assignments are NOT added here - the engine runs its existing
# Add-ScopeTagsBasicInfoIfApplicable / Add-AssignmentsForObjectIfApplicable
# post-steps after this returns (BasicInfo is populated, so they are not no-ops).
#
# Opt-in: the engine only calls this when Options.FallbackDocumentation is $true.

$script:_fallbackSecretRegex   = '(?i)(password|secret|privatekey|private_key|pfxblob|clientsecret|encryptionkey|\bpfx\b)'
$script:_fallbackExcludedNames = @(
    'id', 'createdDateTime', 'lastModifiedDateTime', 'modifiedDateTime', 'version',
    'roleScopeTagIds', 'roleScopeTags', 'assignments', 'supportsScopeTags',
    # already emitted as basic-info rows by Add-BasicDefaultValues
    'displayName', 'name', 'description'
)
$script:_fallbackMaxNamedDepth = 2      # Category + SubCategory; deeper -> compact JSON
$script:_fallbackMaxStringLen  = 2000   # truncate longer string values

function Get-FallbackDisplayName {
    param([string]$PropName)
    if (-not $PropName) { return $PropName }
    $s = [regex]::Replace($PropName, '([a-z0-9])([A-Z])', '$1 $2')        # camelCase -> camel Case
    $s = [regex]::Replace($s, '([A-Z]+)([A-Z][a-z])', '$1 $2')            # ABCWord  -> ABC Word
    $s = ($s -replace '[_\-]', ' ').Trim()
    if ($s.Length -gt 0) { $s = $s.Substring(0, 1).ToUpper() + $s.Substring(1) }
    return $s
}

function Test-FallbackExcluded {
    param([string]$Name, $RemoveList)
    if (-not $Name) { return $true }
    if ($Name -like '*@odata*') { return $true }
    if ($Name.StartsWith('#')) { return $true }
    if ($script:_fallbackExcludedNames -contains $Name) { return $true }
    if ($RemoveList -and ($RemoveList -contains $Name)) { return $true }
    return $false
}

function Test-FallbackIsObject {
    param($Value)
    return ($Value -is [System.Management.Automation.PSCustomObject] -or $Value -is [hashtable])
}

function Test-FallbackIsArray {
    param($Value)
    return ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string])
}

function Format-FallbackScalar {
    param($Value, [string]$Name)
    if ($null -eq $Value) { return $null }
    if ($Name -match $script:_fallbackSecretRegex) { return '*** redacted ***' }
    if ($Value -is [bool])     { return ([bool]$Value).ToString() }
    if ($Value -is [datetime]) { return (Format-BasicDateValue $Value) }
    $s = [string]$Value
    if ($s.Length -gt $script:_fallbackMaxStringLen) {
        $s = $s.Substring(0, $script:_fallbackMaxStringLen) + ' ... (truncated)'
    }
    return $s
}

function Add-FallbackRow {
    param([string]$Name, $Value, $RawValue, [string]$Category, [string]$SubCategory, [int]$Level, [string]$EntityKey)
    $ctx = Get-CurrentDocumentationContext
    $ctx.AddSetting([PSCustomObject]@{
        Name = $Name; Value = $Value; Category = $Category; SubCategory = $SubCategory
        Level = $Level; RawValue = $RawValue; EntityKey = $EntityKey
    })
}

# Recursively walk an object's properties, emitting one row per non-internal
# property. Scalars are emitted before nested objects at each level so flat
# properties group above their sub-sections.
function Add-FallbackProperties {
    param($Obj, [int]$Level, [string]$Category, [string]$SubCategory, [string]$PathPrefix, $RemoveList)
    if ($null -eq $Obj) { return }

    $candidates = @()
    foreach ($p in $Obj.PSObject.Properties) {
        if (Test-FallbackExcluded $p.Name $RemoveList) { continue }
        if ($null -eq $p.Value) { continue }
        $candidates += $p
    }

    # Pass 1: scalars + arrays (leaf rows). Pass 2: nested objects (sub-levels).
    $leaves  = @($candidates | Where-Object { -not (Test-FallbackIsObject $_.Value) })
    $nested  = @($candidates | Where-Object {      (Test-FallbackIsObject $_.Value) })

    foreach ($p in $leaves) {
        $name      = $p.Name
        $val       = $p.Value
        $display   = Get-FallbackDisplayName $name
        $entityKey = if ($PathPrefix) { "$PathPrefix.$name" } else { $name }

        if ($name -match $script:_fallbackSecretRegex) {
            Add-FallbackRow $display '*** redacted ***' $null $Category $SubCategory $Level $entityKey
            continue
        }
        if (Test-FallbackIsArray $val) {
            $items = @($val)
            if ($items.Count -eq 0) { continue }
            $hasComplex = $false
            foreach ($it in $items) { if ((Test-FallbackIsObject $it) -or (Test-FallbackIsArray $it)) { $hasComplex = $true; break } }
            if ($hasComplex) {
                Add-FallbackRow $display ($val | ConvertTo-Json -Depth 20 -Compress) $val $Category $SubCategory $Level $entityKey
            }
            else {
                $joined = ($items | ForEach-Object { Format-FallbackScalar $_ $name }) -join ([Environment]::NewLine)
                if ($joined) { Add-FallbackRow $display $joined $val $Category $SubCategory $Level $entityKey }
            }
            continue
        }
        $fv = Format-FallbackScalar $val $name
        if ($null -eq $fv -or "$fv" -eq '') { continue }
        Add-FallbackRow $display $fv $val $Category $SubCategory $Level $entityKey
    }

    foreach ($p in $nested) {
        $name      = $p.Name
        $val       = $p.Value
        $display   = Get-FallbackDisplayName $name
        $entityKey = if ($PathPrefix) { "$PathPrefix.$name" } else { $name }

        if ($Level -lt $script:_fallbackMaxNamedDepth) {
            $childCat = if ($Level -eq 0) { $display } else { $Category }
            $childSub = if ($Level -eq 1) { $display } else { $SubCategory }
            Add-FallbackProperties $val ($Level + 1) $childCat $childSub $entityKey $RemoveList
        }
        else {
            Add-FallbackRow $display ($val | ConvertTo-Json -Depth 20 -Compress) $val $Category $SubCategory $Level $entityKey
        }
    }
}

function Invoke-GenericFallbackDocumentation {
    param($PolicyObject, [DocumentationContext]$Context)

    Set-CurrentDocumentationContext $Context

    # Standard header rows (Name / Description / Platform / Profile type, then
    # Created / Modified / Version). Populating BasicInfo also un-gates the
    # engine's scope-tag and assignment post-steps.
    Add-BasicDefaultValues $PolicyObject
    Add-BasicAdditionalValues $PolicyObject

    $obj = if ($PolicyObject.PSObject.Properties['JsonObject'] -and $PolicyObject.JsonObject) {
        $PolicyObject.JsonObject
    } else {
        $PolicyObject
    }

    $removeList = $null
    if ($PolicyObject.PolicyType -and $PolicyObject.PolicyType.PSObject.Properties['_PropertiesToRemove']) {
        $removeList = $PolicyObject.PolicyType._PropertiesToRemove
    }

    Add-FallbackProperties $obj 0 $null $null '' $removeList

    $Context.InputType = 'GenericFallback'
}

# Register the generic fallback as the LAST input provider (Order = MaxValue) so
# it only ever sees objects that no specific provider claimed. Its Match honors
# the opt-in Options.FallbackDocumentation flag - when off, it declines and the
# object falls through to the engine's NoProvider stub, exactly as before.
function Invoke-InitializeGenericFallbackInput {
    Add-DocumentationInputProvider ([PSCustomObject]@{
        Name      = 'GenericFallback'
        Order     = [int]::MaxValue
        Match     = {
            param($PolicyObject, $Context)
            return [bool]($Context -and $Context.Options -and $Context.Options.FallbackDocumentation -eq $true)
        }
        Translate = { param($PolicyObject, $Context) Invoke-GenericFallbackDocumentation $PolicyObject $Context }
    })
}

Invoke-InitializeGenericFallbackInput
