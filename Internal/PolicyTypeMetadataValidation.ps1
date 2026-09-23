<#
.SYNOPSIS
    Startup-time sanity check for IntunePolicyType classes against Graph metadata.

.DESCRIPTION
    Cross-references each registered IntunePolicyTypeBase subclass against the
    cached Graph Beta CSDL (%LOCALAPPDATA%\IntuneManagement\GraphMetaData.xml).
    Catches three classes of registration bug that the framework cannot detect
    at runtime:

      1. _API points to an endpoint that does not exist in the public schema
         (typo, portal-only API, removed endpoint).
      2. _QueryList references an isof('microsoft.graph.X') type that does not
         exist as an EntityType in the schema (misspelled type name).
      3. Two PolicyTypes share the same _API + _QueryList — a sibling-type
         filter swap (e.g. an iOS type accidentally filtering on a macOS type).

    Issues are written to the log at warning level. The validation is best-effort
    and gracefully skipped if metadata is unavailable.

.NOTES
    Sibling of Tools\Audit-PolicyTypeFlags.ps1, but runs in-process at startup
    against the live $script:IntuneTypes registry rather than scraping source.
#>

# PolicyType IDs whose _API is known to be absent from the public Graph metadata
# but is still in active use (portal-only first-party endpoints, etc.).
# Keep this list small and document the reason next to each entry.
$script:PolicyTypeMetadataApiAllowlist = @(
    # Portal-only endpoint per project_known_bugs_from_review.md (item 5).
    # User decision 2026-05-18: leave registration as-is, don't re-flag.
    "InventoryPolicies"
)

function Test-PolicyTypeMetadata
{
    [CmdletBinding()]
    param(
        [object[]]$PolicyTypes,
        [xml]$MetadataXml
    )

    if(-not $PolicyTypes) { $PolicyTypes = $script:IntuneTypes }
    if(-not $PolicyTypes -or $PolicyTypes.Count -eq 0) { return @() }

    if(-not $MetadataXml)
    {
        # -NoDownload: this is a diagnostic sanity check that runs on AppInitialized,
        # i.e. inside Import-Module. It must never be the reason a 7-8 MB metadata
        # download happens at load. Cached metadata validates; no cache just skips
        # (handled immediately below).
        Get-GraphMetaData -NoDownload
        $MetadataXml = $script:GraphMetaDataXML
    }

    if(-not $MetadataXml)
    {
        Write-Log "Test-PolicyTypeMetadata: Graph metadata not available; skipping validation" 2
        return @()
    }

    $nsm = New-Object System.Xml.XmlNamespaceManager $MetadataXml.NameTable
    $nsm.AddNamespace("e", "http://docs.oasis-open.org/odata/ns/edm")

    $entityNames = @{}
    foreach($et in $MetadataXml.SelectNodes("//e:EntityType", $nsm))
    {
        $entityNames[$et.Name] = $true
    }

    $navPropNames = @{}
    foreach($np in $MetadataXml.SelectNodes("//e:NavigationProperty", $nsm))
    {
        $navPropNames[$np.Name] = $true
    }

    $issues = @()
    $apiQueryGroups = @{}

    foreach($pt in $PolicyTypes)
    {
        $api    = $pt._API
        $id     = $pt._ID
        $qList  = $pt._QueryList

        if(-not $api) { continue }

        # --- Rule 1: API last segment must exist as a NavigationProperty somewhere ---
        $apiPath     = $api -replace '%[^%]+%', 'x'   # strip placeholders like %OrganizationId%
        $segments    = @($apiPath -split '/' | Where-Object { $_ })
        $lastSegment = if($segments.Count) { $segments[-1] } else { $null }

        if($lastSegment -and -not $navPropNames.ContainsKey($lastSegment) -and $id -notin $script:PolicyTypeMetadataApiAllowlist)
        {
            $issues += [PSCustomObject]@{
                Severity     = "Warning"
                PolicyTypeId = $id
                Issue        = "API endpoint not in Graph metadata"
                Detail       = "_API='$api' - last segment '$lastSegment' is not a navigation property in GraphMetaData.xml. Likely portal-only, typo, or removed."
            }
        }

        # --- Rule 2: every isof('microsoft.graph.X') must reference a known EntityType ---
        if($qList)
        {
            $decoded = [uri]::UnescapeDataString($qList)
            foreach($m in [regex]::Matches($decoded, "isof\(\s*'(?:microsoft\.graph\.|graph\.)?([A-Za-z0-9_]+)'\s*\)"))
            {
                $typeName = $m.Groups[1].Value
                if(-not $entityNames.ContainsKey($typeName))
                {
                    $issues += [PSCustomObject]@{
                        Severity     = "Warning"
                        PolicyTypeId = $id
                        Issue        = "isof() type not in Graph metadata"
                        Detail       = "_QueryList references microsoft.graph.$typeName which is not an EntityType in GraphMetaData.xml."
                    }
                }
            }
        }

        # --- Rule 3: collect (API, QueryList) pairs for duplicate detection ---
        if($qList)
        {
            $key = "$api||$qList"
            if(-not $apiQueryGroups.ContainsKey($key))
            {
                $apiQueryGroups[$key] = @()
            }
            $apiQueryGroups[$key] += $id
        }
    }

    foreach($key in $apiQueryGroups.Keys)
    {
        $ids = $apiQueryGroups[$key]
        if($ids.Count -lt 2) { continue }

        $sepIdx    = $key.IndexOf("||")
        $apiPart   = $key.Substring(0, $sepIdx)
        $queryPart = $key.Substring($sepIdx + 2)

        $issues += [PSCustomObject]@{
            Severity     = "Warning"
            PolicyTypeId = ($ids -join ", ")
            Issue        = "Duplicate API+QueryList across PolicyTypes"
            Detail       = "PolicyTypes [$($ids -join ', ')] all use _API='$apiPart' with identical _QueryList='$queryPart'. One is likely mis-filtered (sibling type swap)."
        }
    }

    return $issues
}

function Invoke-PolicyTypeMetadataValidation
{
    [CmdletBinding()]
    param()

    try
    {
        $issues = Test-PolicyTypeMetadata
        if(-not $issues -or $issues.Count -eq 0)
        {
            Write-LogDebug "PolicyType metadata validation: no issues found"
            return
        }

        Write-Log "PolicyType metadata validation: $($issues.Count) issue(s) found" 2
        foreach($i in $issues)
        {
            Write-Log "  [$($i.PolicyTypeId)] $($i.Issue) - $($i.Detail)" 2
        }
    }
    catch
    {
        Write-LogError "PolicyType metadata validation failed" $_.Exception
    }
}
