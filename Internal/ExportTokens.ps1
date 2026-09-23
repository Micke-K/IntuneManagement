# Organization tokenization for exported JSON.
#
# Exported files replace tenant-specific values with placeholders so the same
# file can be imported into another tenant: the tenant id becomes
# %OrganizationId%, and the organization (company) name can become
# %OrganizationName% - but does not by default, see $script:GraphExportTokensDefault.
# Before this file the replace was open-coded at four call sites (two export, two
# import), was unconditional, and used the raw value as a REGEX - so an
# organization name containing regex metacharacters, e.g. "Contoso (AU)", was
# silently never replaced.
#
# This file owns the whole concept:
#   * which tokens exist, how they are spelled, and whether import restores them
#     ($script:GraphExportTokens - one row per token, both directions declared);
#   * the two settings that control it (below);
#   * the replace itself. Convert-GraphExportLiteralReplace is the ONLY place in
#     the project that turns an organization value into a token or back.
#
# Direction of the two converters:
#   Convert-GraphOrganizationValueToToken   export: "Contoso" -> %OrganizationName%
#   Convert-GraphOrganizationTokenToValue   import: %OrganizationName% -> "Fabrikam"
#
# Every token the export can write, the import puts back - to the TARGET tenant's
# values, which is what makes an exported file tenant-neutral rather than merely
# tenant-anonymised.
#
# Only the export direction is gated by the settings. Restoring is unconditional
# on purpose: a file exported last year carries tokens no matter how the setting
# is configured today, and leaving them literal would import broken values.
#
# See Docs/ExportImportAndCopy.md ("Organization Tokens").

# This file sorts before Internal/IntuneManager.ps1, which is where the
# Import/Export section is normally declared. Add-SettingsSection is idempotent,
# so declaring it here too makes the registration below load-order independent.
Add-SettingsSection -Title "Import/Export" -Id "ImportExport" -Order 10

Add-SettingsObject -Title "Replace organization values in export files" -Key "ExportReplaceOrganizationValues" -Type "Boolean" `
-Description "Replace tenant-specific values with placeholders in exported JSON. By default the tenant id becomes %OrganizationId%; the organization name is left as written unless you opt in. Turn this off to export the raw values - readable and diff-friendly, but the file is then tied to the tenant it came from. Importing a file that already contains placeholders always resolves them, whatever this is set to. Which values are replaced can be changed, see Docs/ExportImportAndCopy.md." -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

# One row per token. Order matters and matches the original open-coded order:
# the id is replaced before the name, which is what keeps a tenant whose
# organization name is unset (the authentication providers then fall back to
# using the tenant GUID as the name) producing %OrganizationId%, not
# %OrganizationName%, for that GUID.
#
# RestoreOnImport is $true for BOTH tokens. Anything the export writes, the
# import has to put back - a token is a placeholder for a value, not a value.
#
# It used to be $false for the name, inherited from before this file existed:
# the open-coded export replaced both values but the open-coded import resolved
# %OrganizationId% only. So a tokenized name was never put back and the import
# CREATED the object with the placeholder still in it - a policy literally named
# "%OrganizationName% - Baseline", and every description, rule or OMA-URI value
# that mentioned the organization carrying "%OrganizationName%" verbatim. Even a
# same-tenant export/re-import round trip lost the name that way.
#
# The justification recorded for it was that restoring would defeat cross-tenant
# name matching, since Normalize-IntuneImportPolicyName strips the token so
# "%OrganizationName% - Baseline" matches "Contoso - Baseline" in the target.
# That does not follow: match resolution runs in the driver/UI layer on the
# object loaded from the FILE (Resolve-IntuneImportUpdateTarget, called before
# ImportObject/UpdateObject), and the restore runs inside those. The matcher
# therefore sees the token either way, and restoring changes only what is
# written to Graph. Repeat imports of the same file keep matching too: the file
# still holds the token, and the normalizer also strips the literal organization
# name of the tenant being imported INTO.
#
# If the target organization name cannot be resolved, Get-GraphExportTokenValue
# returns nothing and the token is left literal rather than replaced with a
# guess. Convert-GraphOrganizationTokenToValue logs a warning when that happens
# to a token the document actually contains, because the object it then creates
# carries a placeholder in a user-visible field.
#
# Name is the canonical value for the ExportReplaceTokens setting; Aliases are
# the other spellings accepted there (Name itself included, so matching is one
# lookup). Adding a token is one row plus one case in Get-GraphExportTokenValue.
$script:GraphExportTokens = @(
    [PSCustomObject]@{
        Name            = "OrganizationId"
        Token           = "%OrganizationId%"
        Aliases         = @("OrganizationId", "TenantId", "OrgId")
        RestoreOnImport = $true
    }
    [PSCustomObject]@{
        Name            = "OrganizationName"
        Token           = "%OrganizationName%"
        Aliases         = @("OrganizationName", "CompanyName", "OrgName", "TenantName")
        RestoreOnImport = $true
    }
)

# Used when the hidden ExportReplaceTokens setting is missing or empty.
#
# The tenant id only. The id is machine-readable and unsafe to leave literal - it
# is what ties a file to one tenant, and nothing reads it as prose. The
# organization name is the opposite on both counts: it appears inside policy
# names, descriptions, rules and OMA-URI values, where it is text a human wrote,
# and the boundary guard cannot make a short name ("IT") safe in every document.
# A false match there does not just mask a value, it rewrites unrelated prose -
# and on import rewrites it again, to the target organization's name. Opting in
# is a deliberate choice about your own naming conventions, so it is not made for
# you. Add OrganizationName to ExportReplaceTokens to enable it - see the setting
# comment in Get-GraphExportTokenNames.
$script:GraphExportTokensDefault = "OrganizationId"

# Every token literal, for callers that need the spellings but not the replace
# (import name normalization).
function Get-GraphExportTokenStrings
{
    [OutputType([string[]])]
    param()

    return @($script:GraphExportTokens | ForEach-Object { $_.Token })
}

# Case-insensitive LITERAL replace - the single replace in the token pipeline.
#
# Both sides need escaping. -replace treats its pattern as a regex, so an
# organization name like "Contoso (AU)" used to be read as a capture group and
# never matched; and it treats '$' in the replacement as a group reference, so a
# value containing '$' would be mangled ('$$' is how .NET spells a literal '$').
# IgnoreCase keeps the pre-existing behaviour of -replace, which matters for a
# hand-edited file that spells the token in another case.
#
# -RequireValueBoundary rejects a match that is glued to an alphanumeric on
# either side, so the value is only replaced where it stands as a value of its
# own. It is what stops a short organization name from being found inside
# unrelated words: with an organization literally named "IT", an unguarded
# replace turned every "Security" in the file into "Secur%OrganizationName%y".
# The exported file is corrupt past recovery either way - importing it now
# resolves the placeholder to the target organization, so "Security" comes back
# as "SecurFabrikamy" rather than as itself. The boundary is spelled
# [\p{L}\p{N}] rather than \w on purpose - '_' has to count as a boundary, or a
# tenant id inside a compound Graph id ("<guid>_<guid>_0") would stop being
# masked - and rather than [0-9A-Za-z], which gave a non-ASCII name no guard at all.
#
# The cost is that a value glued to a word is no longer replaced:
# "ContosoBaseline" keeps the literal name where it used to become
# "%OrganizationName%Baseline". That is the right trade - a missed placeholder
# leaves a readable, tenant-specific name, a false one silently rewrites
# unrelated data - but it does mean cross-tenant name matching only works for
# names where the organization name is a separate word.
#
# Only the export direction passes the switch. On import the search string is a
# token literal ("%OrganizationId%"), which is delimited by '%' already.
function Convert-GraphExportLiteralReplace
{
    [OutputType([string])]
    param([string]$Text, [string]$Find, [string]$ReplaceWith, [switch]$RequireValueBoundary)

    if(-not $Text -or -not $Find) { return $Text }

    $pattern = [Regex]::Escape($Find)

    if($RequireValueBoundary)
    {
        # [\p{L}\p{N}] - any Unicode letter or number - rather than [0-9A-Za-z].
        # Organization names are not ASCII: with an ASCII-only class, a name
        # starting with 'A' or 'E' got no left-hand guard at all (neither is an
        # ASCII alphanumeric), so the guard silently did nothing on that side and
        # the name was still found inside unrelated words. \p also keeps '_' a
        # boundary, which \w would not - a tenant id inside a compound Graph id
        # ("<guid>_<guid>_0") has to stay maskable.
        #
        # Anchored per side: a value that already starts or ends with punctuation
        # ("Contoso (AU)") needs no guard on that side, and adding one there would
        # never match - there is no boundary between two non-alphanumerics.
        if($Find -match '^[\p{L}\p{N}]')  { $pattern = "(?<![\p{L}\p{N}])$pattern" }
        if($Find -match '[\p{L}\p{N}]$')  { $pattern = "$pattern(?![\p{L}\p{N}])" }
    }

    return [Regex]::Replace($Text,
                            $pattern,
                            ($ReplaceWith -replace '\$', '$$$$'),
                            [Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

# The value a token stands for. Callers can override per call: a cross-tenant
# export has to mask the exported object's OWN tenant id, not the id of the
# tenant that happens to be the default one.
function Get-GraphExportTokenValue
{
    [OutputType([string])]
    param([string]$Name, [string]$OrganizationId, [string]$OrganizationName)

    switch($Name)
    {
        "OrganizationId"   { if($OrganizationId)   { return $OrganizationId }   return (Get-CurrentTenantId) }
        "OrganizationName" {
            if($OrganizationName) { return $OrganizationName }

            # Only fall back to the connected tenant's name when the call IS about the
            # connected tenant. An -OrganizationId for another tenant with no name meant
            # tenant A's data was searched for tenant B's organization name: a false hit
            # rewrites unrelated data with a token import never restores, and a miss
            # leaves the file unchanged. Returning nothing leaves the name literal,
            # which is the recoverable outcome.
            if($OrganizationId -and $OrganizationId -ne [string](Get-CurrentTenantId)) { return $null }

            return (Get-CurrentOrganizationName)
        }
    }

    return $null
}

# The organization an object's DATA belongs to, which on a cross-tenant export is
# not the organization the default token points at. Exporting an object fetched
# from tenant A while tenant B is the default used to mask B's id in A's data,
# leaving A's real id in the file - a mixed-tenant export that then imported A's
# tenant id verbatim into the target.
#
# Preference order: an explicit TenantId on the object, then the token it was
# fetched with (a listed policy carries _TokenId, not a tenant id), then the
# current tenant.
#
# The id and the name are ONE pair and always describe the same tenant. They used
# to be resolved independently, so an object carrying an explicit TenantId with no
# usable token - which is exactly what Get-GraphPolicyFromFile -TenantId produces -
# got that tenant's id alongside the CONNECTED tenant's organization name.
function Get-GraphObjectOrganizationInfo
{
    [OutputType([PSCustomObject])]
    param($GraphObject)

    $id = $null
    $name = $null

    if($GraphObject)
    {
        # PSObject.Properties, not a bare $obj.Prop: these objects are PowerShell
        # class instances as well as PSCustomObjects, and a missing property on a
        # class instance is not a silent $null under StrictMode.
        foreach($prop in @("TenantId", "_TenantId"))
        {
            if($GraphObject.PSObject.Properties[$prop] -and $GraphObject.$prop)
            {
                $id = [string]$GraphObject.$prop
                break
            }
        }

        # -gt 0, not -ne $null: _TokenId is declared [int] on IntunePolicyBase, so it
        # is 0 - never $null - on an object that carries no token, and Get-TokenInfo
        # does not filter on 0. A null check therefore handed back EVERY registered
        # token, making $tokenInfo an array and .TenantID an array of tenant ids the
        # moment a second tenant was signed in. 0 means "the default token", which is
        # what the fallback below already resolves.
        if($GraphObject.PSObject.Properties['_TokenId'] -and [int]$GraphObject._TokenId -gt 0)
        {
            $tokenInfo = Get-OperationTokenInfo ([int]$GraphObject._TokenId)
            if($tokenInfo)
            {
                if(-not $id -and $tokenInfo.TenantID) { $id = [string]$tokenInfo.TenantID }

                # Only when the token describes the tenant the id came from. An object
                # can carry an explicit TenantId (A) and a token for another tenant (B),
                # and A's id paired with B's name is the same mixed-tenant export the
                # id resolution above exists to prevent.
                if($tokenInfo.TenantName -and $id -eq [string]$tokenInfo.TenantID)
                {
                    $name = [string]$tokenInfo.TenantName
                }
            }
        }
    }

    if(-not $id) { $id = Get-CurrentTenantId }

    # Name resolved FOR the id, not independently of it. The connected tenant's name is
    # only the right answer when the id is the connected tenant; for any other tenant
    # the name has to come from that tenant's own token.
    if(-not $name -and $id)
    {
        if($id -eq [string](Get-CurrentTenantId))
        {
            $name = Get-CurrentOrganizationName
        }
        else
        {
            $tenantToken = Get-TokenInfoForTenant $id
            if($tenantToken -and $tenantToken.TenantName) { $name = [string]$tenantToken.TenantName }
        }
    }

    # The name can legitimately come back empty: a tenant nobody is signed in to has no
    # discoverable organization name. Empty is the correct answer - it leaves the name
    # literal in the exported file, which stays readable and importable. The old
    # behaviour, substituting whichever organization happened to be connected, silently
    # tokenized the WRONG name, so import would resolve it to the target organization
    # and the original text would be unrecoverable from the file.
    return [PSCustomObject]@{ OrganizationId = $id; OrganizationName = $name }
}

# Which tokens the export direction should write, as canonical names. Empty array
# = replace nothing (export raw values).
#
# -OrganizationId selects whose per-tenant settings are consulted. It is the tenant
# being EXPORTED, which on a cross-tenant export is not the connected one - so a
# per-tenant "also replace the organization name" would otherwise be read from the
# wrong tenant and silently not apply.
function Get-GraphExportTokenNames
{
    [OutputType([string[]])]
    param([string]$OrganizationId)

    # Resolved before the master switch is read, because BOTH settings are
    # per-tenant and both have to answer for the tenant being exported. Reading the
    # switch against the connected tenant while reading the selector against the
    # exported one was worse than reading either consistently: a source tenant that
    # had turned replacement ON exported raw values because the DEFAULT tenant had
    # it off, which is the case tokenization exists to prevent.
    $tenantId = if($OrganizationId) { $OrganizationId } else { Get-CurrentTenantId }

    # Master switch. Off = the pre-tokenization behaviour: raw values in the file.
    if((Get-SettingValue "ExportReplaceOrganizationValues" -TenantID $tenantId) -ne $true) { return @() }

    # HIDDEN setting - deliberately not registered with Add-SettingsObject, so it
    # never appears in the settings UI. Set it by hand to narrow what is
    # replaced:
    #
    #   Windows:     HKCU:\Software\IntuneManagement\IntuneManager
    #                (or HKCU:\Software\IntuneManagement\<TenantId>\IntuneManager
    #                for one tenant only)
    #                value name ExportReplaceTokens, type String
    #   non-Windows: the "IntuneManager" object in the settings JSON file
    #
    # e.g. "OrganizationId,OrganizationName" -> the organization name is tokenized
    # too, which is what the default used to be.
    #
    # Missing, empty or blank falls back to $script:GraphExportTokensDefault - the
    # tenant id, which is also what v3.x replaced. Get-SettingStoreValue already
    # treats "" as missing; the IsNullOrWhiteSpace checks extend that to a value
    # of nothing but spaces, which is a blank value by any reading and must not
    # mean "replace nothing".
    $configured = $null
    if($tenantId) { $configured = Get-SettingStoreValue "$tenantId\IntuneManager" "ExportReplaceTokens" }
    if([string]::IsNullOrWhiteSpace($configured)) { $configured = Get-SettingStoreValue "IntuneManager" "ExportReplaceTokens" $script:GraphExportTokensDefault }
    if([string]::IsNullOrWhiteSpace($configured)) { $configured = $script:GraphExportTokensDefault }

    # Same separators as ImportMatchOrganizationTokens: comma, semicolon, newline.
    $selected = @()
    foreach($part in ([regex]::Split([string]$configured, '[,;\r\n]+')))
    {
        if([string]::IsNullOrWhiteSpace($part)) { continue }

        $trimmed = $part.Trim()
        $definition = $script:GraphExportTokens | Where-Object { $_.Aliases -contains $trimmed }
        if(-not $definition)
        {
            # Warn rather than fall back to the default: silently replacing more
            # than was asked for is worse than replacing nothing, and the export
            # is valid either way.
            Write-Log "ExportReplaceTokens: unknown token '$trimmed' ignored. Valid values: $(($script:GraphExportTokens | ForEach-Object { $_.Name }) -join ', ')" 2
            continue
        }
        $selected += $definition.Name
    }

    # Table order, not the order they were listed in, so the id is always
    # replaced before the name (see the table comment) whatever the setting says.
    # Dedup comes free.
    return @($script:GraphExportTokens |
                Where-Object { $selected -contains $_.Name } |
                ForEach-Object { $_.Name })
}

# Export direction. -Tokens restricts the call site to a subset (the settings can
# only narrow it further); omit it for "whatever is enabled".
function Convert-GraphOrganizationValueToToken
{
    [OutputType([string])]
    param(
        [string]$Json,
        [string]$OrganizationId,
        [string]$OrganizationName,
        [string[]]$Tokens
    )

    if(-not $Json) { return $Json }

    foreach($name in (Get-GraphExportTokenNames -OrganizationId $OrganizationId))
    {
        if($Tokens -and $Tokens -notcontains $name) { continue }

        $definition = $script:GraphExportTokens | Where-Object Name -eq $name
        if(-not $definition) { continue }

        $value = Get-GraphExportTokenValue -Name $name -OrganizationId $OrganizationId -OrganizationName $OrganizationName
        if(-not $value) { continue }

        $Json = Convert-GraphExportLiteralReplace $Json $value $definition.Token -RequireValueBoundary
    }

    return $Json
}

# Import direction. NOT gated by the settings - see the header comment.
function Convert-GraphOrganizationTokenToValue
{
    [OutputType([string])]
    param(
        [string]$Json,
        [string]$OrganizationId,
        [string]$OrganizationName
    )

    if(-not $Json) { return $Json }

    foreach($definition in $script:GraphExportTokens)
    {
        if($definition.RestoreOnImport -ne $true) { continue }

        $value = Get-GraphExportTokenValue -Name $definition.Name -OrganizationId $OrganizationId -OrganizationName $OrganizationName
        if(-not $value)
        {
            # Left literal rather than replaced with a guess - but say so. The
            # object about to be created carries the placeholder in whatever field
            # held it, which for a display name or description is user-visible.
            # Silence here is what made the old never-restore behaviour hard to
            # spot: the file looked fine and the imported object was wrong.
            if($Json.IndexOf($definition.Token, [StringComparison]::OrdinalIgnoreCase) -ge 0)
            {
                Write-Log "Import: cannot resolve $($definition.Token) for the target tenant - it stays literal in the imported object. Sign in to the target tenant, or edit the file." 2
            }
            continue
        }

        $Json = Convert-GraphExportLiteralReplace $Json $definition.Token $value
    }

    return $Json
}
