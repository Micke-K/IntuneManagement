#Add Setting setion
Add-SettingsSection -Title "Intune" -Id "IntuneManager" -Order 10

Add-SettingsSection -Title "Import/Export" -Id "ImportExport" -Order 10
Add-SettingsSection -Title "MS Graph General" -Id "GraphGeneral" -Order 10
Add-SettingsSection -Title "Intune Tools" -Id "IntuneTools" -Order 10

$script:IntuneManagerImportOptions = @()
$script:IntuneManagerImportOptions += [NameValueObject]::new("Always import", "alwaysImport")
$script:IntuneManagerImportOptions += [NameValueObject]::new("Skip if object exists", "skipIfExist")
$script:IntuneManagerImportOptions += [NameValueObject]::new("Replace", "replace")
$script:IntuneManagerImportOptions += [NameValueObject]::new("Replace with assignments", "replace_with_assignments")
$script:IntuneManagerImportOptions += [NameValueObject]::new("Update", "update")

#Add Setting values
Add-SettingsObject -Title "App packages folder" -Key "IntuneAppPackagesFolder" -Type "Folder" `
    -Description "Root folder where intune app packages are located" `
    -SubPath "IntuneManager" -Section "IntuneManager"

Add-SettingsObject -Title "Save Encryption File" -Key "IntuneSaveEncryptionFile" -Type "Boolean" `
    -Description "Save encryption file when uploading an app. This can then be used to when downloading the app file." `
    -SubPath "IntuneManager" -Section "IntuneManager"    

Add-SettingsObject -Title "App download folder" -Key "IntuneAppDownloadFolder" -Type "Folder" `
    -Description "Folder where app packages will be downloaded and where encryption files will be saved" `
    -SubPath "IntuneManager" -Section "IntuneManager"

# GraphPageSize is consumed by the ENGINE (IntunePolicyTypeBase list URLs), so it must
# be registered here, not in the UI extensions - a UI-only registration left headless
# (IM_UI_BACKEND=None) sessions resolving the wrong storage path and no default.
Add-SettingsObject -Title "Graph Page Size" -Key "GraphPageSize" -Type "List" `
    -ItemsSource @(
        [PSCustomObject]@{ Name = "Graph Default"; Value = "0"    },
        [PSCustomObject]@{ Name = "5";             Value = "5"    },
        [PSCustomObject]@{ Name = "20";            Value = "20"   },
        [PSCustomObject]@{ Name = "50";            Value = "50"   },
        [PSCustomObject]@{ Name = "100";           Value = "100"  },
        [PSCustomObject]@{ Name = "1000";          Value = "1000" },
        [PSCustomObject]@{ Name = "All";           Value = "All"  }
    ) `
    -DefaultValue "0" -Description "How many items load at a time" `
    -SubPath "IntuneManager" -Section "IntuneManager"

Add-SettingsObject -Title "Root folder" -Key "RootFolder" -Type "Folder" `
-Description "Root folder for exporting/importing objects" `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Add object type" -Key "AddObjectType" -Type "Boolean" `
-Description "Default setting for adding object type to the export folder" -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Add company name" -Key "AddCompanyName" -Type "Boolean" `
-Description "Default setting for adding company name to the export folder" -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Export Assignments" -Key "ExportAssignments" -Type "Boolean" `
-Description "Default setting for exporting assignments" -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Export nested group levels" -Key "ExportNestedGroupLevels" -Type "String" `
-Description "Depth of group-membership recursion when exporting groups. 1 (default) only exports the group assigned to the policy. 2 also exports groups that are direct members of the assigned group. 3 goes one level deeper, and so on. Higher values mean more Graph calls and may hit throttling on large group trees." -DefaultValue "1" `
-SubPath "IntuneManager" -Section "ImportExport"

# Per-policy group migration on import (Internal/IntuneGroupMigration.ps1). Unlike 3.x
# (which bulk-created every migration-table group and never actually read these
# settings), only groups referenced by the imported policy are resolved/created, and
# both toggles are wired for real.
Add-SettingsObject -Title "Create groups and filters" -Key "CreateGroupOnImport" -Type "Boolean" `
-Description "Create Entra groups and assignment filters referenced by an imported policy when they do not exist in the target tenant. Groups are created from the export's Groups sidecar (dynamic groups keep their membership rule) or as a default cloud security group; filters from the AssignmentFilters sidecar (platform + rule preserved)." -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Convert synced groups" -Key "ConvertSyncedGroupOnImport" -Type "Boolean" `
-Description "When a group referenced by an imported policy was AD-synced in the source tenant and does not exist in the target, recreate it as a cloud Entra group. When off, the group is skipped and the reference is left untranslated." -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Import type" -Key "ImportType" -Type "List" -ItemsSource $script:IntuneManagerImportOptions `
-Description "How files are imported. Always import: no detection of existing objects. Skip if object exists: skip when a matching object is found. Replace: import the file, copy the existing object's assignments to it, then delete the existing object. Replace with assignments: same but assignments come from the import file. Update: settings on the existing object are replaced from the file." -DefaultValue "alwaysImport" `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Import match policy reference" -Key "ImportMatchEnablePolicyToken" -Type "Boolean" `
-Description "Allow update and skip-if-exists import matching by policy reference token in the object name, e.g. [SEC-1053]. Exact name and same-tenant ID matching are always available." -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Import match policy reference regex" -Key "ImportMatchPolicyReferenceRegex" -Type "String" `
-Description "Regex used to find policy reference tokens in names. It should include a named capture group called ref. Default matches values like [SEC-1053] or SEC-1053." -DefaultValue "(?i)(?:\[(?<ref>[A-Z][A-Z0-9]{1,15}-\d{2,10})\])|(?<ref>\b[A-Z][A-Z0-9]{1,15}-\d{2,10}\b)" `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Import match normalized name" -Key "ImportMatchEnableNormalizedName" -Type "Boolean" `
-Description "Allow update and skip-if-exists import matching after removing organization-specific prefixes or variables from names." -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Import match organization tokens" -Key "ImportMatchOrganizationTokens" -Type "String" `
-Description "Extra organization-specific words or prefixes to remove before normalized-name matching. Separate values with comma, semicolon, or new lines." -DefaultValue "" `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Import Assignments" -Key "ImportAssignments" -Type "Boolean" `
-Description "Default value for Import assignments when importing objects" -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Import Scope (Tags)" -Key "ImportScopeTags" -Type "Boolean" `
-Description "Default value for Import Scope (Tags) when importing objects" -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Add ID to export file" -Key "AddIDToExportFile" -Type "Boolean" `
-Description "This will add object ID to the export file to support objects with the same name e.g. ObjectName_ObjectId.json" -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Combine requests into batch calls" -Key "UseBatchAPI" -Type "Boolean" `
-Description "Combine Graph requests into `$batch calls, up to 20 per call, on every path that can batch: listing, policy bodies, sub-resources, assignments, import and delete. Turn off to send every request on its own - slower, but each call shows individually in the Graph log. See Docs/GraphBatching.md." -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Send batch calls in parallel (experimental)" -Key "UseParallelBatchAPI" -Type "Boolean" `
-Description "Send `$batch calls concurrently instead of one at a time. Controls concurrency only - whether requests are batched at all is 'Combine requests into batch calls'. Requires PowerShell 7+. Speeds up large queries but raises the chance of HTTP 429 throttling; queues of 20 requests or fewer still go out one at a time. Leave off unless you've tested it in your tenant." -DefaultValue $false `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Parallel batch throttle limit" -Key "ParallelBatchThrottle" -Type "String" `
-Description "Maximum number of concurrent `$batch POST requests when 'Parallel batch dispatch' is enabled. Higher values are faster but more likely to trigger HTTP 429 throttling. Recommended range: 2-8." -DefaultValue "4" `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Pace one-per-second Graph endpoints" -Key "GraphPaceIdentityEndpoints" -Type "Boolean" `
-Description "Graph allows one request per second per tenant, across all applications, on Conditional Access policies, named locations, authentication strengths and identity protection - and sends no Retry-After when it throttles them. When enabled, requests to those endpoints are sent one at a time, one second apart, and are kept out of parallel batch dispatch. Turn off only if Microsoft has raised the limit for your tenant." -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Graph request timeout (seconds)" -Key "MSGraphRequestTimeoutSec" -Type "Int" `
-Description "Maximum time in seconds to wait for a single Graph request before it is aborted. Bounds how long a stalled request can block the UI. Default 100." -DefaultValue 100 `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Resolve reference info" -Key "ResolveReferenceInfo" -Type "Boolean" `
-Description "This will export/import info for referenced/navigation properties eg certificates in VPN profiles etc." -DefaultValue $true `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Sort Json Properties" -Key "SortJsonProperties" -Type "Boolean" `
-Description "Sort JSON properties alphabetically when exporting to improve file readability and consistency" -DefaultValue $false `
-SubPath "IntuneManager" -Section "ImportExport"

# Exported files used to go through Out-File with no -Encoding, which resolves
# to UTF-16LE on PowerShell 5.1 and UTF-8 on 7 - so the same policy exported
# from two hosts produced byte-incompatible files and tools consuming the
# export rejected one of them. UTF-8 without BOM is the default; the other
# options exist for anyone whose pipeline depended on the old behaviour.
$script:ExportEncodingOptions = @()
$script:ExportEncodingOptions += [NameValueObject]::new("UTF-8", "utf8")
$script:ExportEncodingOptions += [NameValueObject]::new("UTF-8 with BOM", "utf8bom")
$script:ExportEncodingOptions += [NameValueObject]::new("Unicode (UTF-16 LE)", "unicode")

Add-SettingsObject -Title "Export file encoding" -Key "ExportFileEncoding" -Type "List" -ItemsSource $script:ExportEncodingOptions `
-Description "Character encoding for exported JSON files. UTF-8 is recommended - it is what Intune, git and other tools expect. Change it only if an existing pipeline depends on another encoding." -DefaultValue "utf8" `
-SubPath "IntuneManager" -Section "ImportExport"

# File names are sanitised with the RUNNING platform's rules, and those differ:
# Windows rejects 41 characters, Linux and macOS only '/' and NUL. So a policy
# named "Microsoft Defender: Antivirus" exports as-is on Linux and the resulting
# file cannot be opened or imported on Windows. Enable this to sanitise against
# every supported platform's rules instead, so an export moves between machines
# unchanged. Off by default because turning it on changes exported file names.
Add-SettingsObject -Title "Portable file names" -Key "PortableFileNames" -Type "Boolean" `
-Description "Remove characters that are invalid on ANY supported platform from exported file names, not just the ones invalid on the current one. Windows rejects `" < > | : * ? \ / while Linux and macOS reject only /, so an export made on Linux can contain file names Windows cannot open. Enable this when exports are shared between platforms. Changes the file names of new exports." -DefaultValue $false `
-SubPath "IntuneManager" -Section "ImportExport"

# PowerShell's own JSON writer formats differently per host: 5.1 puts two spaces
# after the colon and aligns nested values under their key, 7 uses plain
# two-space indentation. So an indented export from v3 (5.1) and from v4 (7) are
# different text for identical data. Compact output is byte-identical on both,
# which is what anyone diffing or hashing exports across hosts needs.
$script:ExportJsonFormatOptions = @()
$script:ExportJsonFormatOptions += [NameValueObject]::new("Indented (readable)", "indented")
$script:ExportJsonFormatOptions += [NameValueObject]::new("Compact (single line)", "compact")

Add-SettingsObject -Title "Export Json format" -Key "ExportJsonFormat" -Type "List" -ItemsSource $script:ExportJsonFormatOptions `
-Description "Layout of exported JSON. Indented is readable but PowerShell 5.1 and 7 indent differently, so files exported from different hosts differ even when the data is identical. Compact is identical on both - use it with 'Sort Json Properties' if the export is stored in git or compared by automation." -DefaultValue "indented" `
-SubPath "IntuneManager" -Section "ImportExport"

# Multi Admin Approval. Since June 2026 MAA also gates app-auth (client credentials)
# writes, not just interactive admin actions. Graph rejects a protected write that
# carries no justification, so without this the operation fails with an opaque
# BadRequest. Empty by default: it only matters on tenants that have an MAA access
# policy, and sending a justification on a tenant without one is harmless but noisy.
Add-SettingsObject -Title "Multi Admin Approval justification" -Key "MultiAdminApprovalJustification" -Type "String" `
-Description "Reason sent with every create/update/delete when the tenant requires Multi Admin Approval. Leave empty unless Tenant administration > Multi Admin Approval has an access policy covering what you are changing. With a justification set, protected changes are queued for a second administrator to approve instead of failing." -DefaultValue "" `
-SubPath "IntuneManager" -Section "ImportExport"

$script:CacheClearOptions = @()
$script:CacheClearOptions += [NameValueObject]::new("Bulk export only", "1")
$script:CacheClearOptions += [NameValueObject]::new("Bulk and manual export", "3")
$script:CacheClearOptions += [NameValueObject]::new("Bulk export and import", "7")
$script:CacheClearOptions += [NameValueObject]::new("All operations", "15")

Add-SettingsObject -Title "Clear Cache Before Export/Import" -Key "ClearCacheBeforeExportImport" -Type "List" -ItemsSource $script:CacheClearOptions `
-Description "Automatically clear object cache before export/import operations (useful when re-downloading profiles to get latest data)" -DefaultValue "1" `
-SubPath "IntuneManager" -Section "ImportExport"

Add-SettingsObject -Title "Clear All Objects From Cache" -Key "ClearAllObjectsFromCache" -Type "Boolean" `
-Description "Whether to clear all cached objects (default is to keep assignments and scopes cached). Use this option if you experience stale cache issues." -DefaultValue $false `
-SubPath "IntuneManager" -Section "ImportExport"

# Tenant-lockout guard ported from the original project: controls the state
# imported Conditional Access policies get. Default 'disabled' so importing a
# foreign tenant's "block everything" policy can't lock the target tenant out.
$script:CAImportStateOptions = @()
$script:CAImportStateOptions += [NameValueObject]::new("As Exported - Change On to Report-only", "AsExportedReportOnly")
$script:CAImportStateOptions += [NameValueObject]::new("As Exported", "AsExported")
$script:CAImportStateOptions += [NameValueObject]::new("Report-only", "enabledForReportingButNotEnforced")
$script:CAImportStateOptions += [NameValueObject]::new("Off", "disabled")

Add-SettingsObject -Title "Default Conditional Access Policy State" -Key "ConditionalAccessState" -Type "List" -ItemsSource $script:CAImportStateOptions `
-Description "Define the state imported Conditional Access policies get. It is recommended to keep this Off (disabled) to avoid accidental tenant lock out." -DefaultValue "disabled" `
-SubPath "IntuneManager" -Section "ImportExport"

# Applies the ConditionalAccessState setting to a CA policy about to be
# imported. Lives here (not in the class) so the rewrite is unit-testable -
# class methods cache their first command binding, which defeats mocking.
# Called from ConditionalAccessType.PreImportCommand, so it covers manual,
# bulk and headless (Start-GraphBulkImport) imports alike.
function Set-CAPolicyImportState
{
    param($PolicyObject)

    if(-not $PolicyObject -or -not $PolicyObject.JsonObject) { return }

    $caState = Get-SettingValue "ConditionalAccessState" "disabled"
    if(-not $caState -or $caState -eq "AsExported") { return }

    if($caState -eq "AsExportedReportOnly")
    {
        if($PolicyObject.JsonObject.state -eq "enabled")
        {
            Write-Log "Conditional Access import: changing Enabled policy '$($PolicyObject.Name)' to Report-only"
            $PolicyObject.JsonObject.state = "enabledForReportingButNotEnforced"
        }
    }
    else
    {
        if($PolicyObject.JsonObject.state -ne $caState)
        {
            Write-Log "Conditional Access import: setting policy '$($PolicyObject.Name)' state to $caState"
        }
        $PolicyObject.JsonObject.state = $caState
    }
}

# Batched per-type object counts for the left-nav "Show item counts" gear
# option. Shared by both UI backends. One $batch round of
# <type API>?$top=1&$count=true per call; results cached on the module scope
# so a view-switch back doesn't refetch - only -Force (gear Refresh) does.
function Get-IntuneTypeCounts
{
    param([switch]$Force, [Int]$TokenId = (Get-DefaultTokenId))

    if(-not $Force -and $script:_intuneTypeCounts) { return $script:_intuneTypeCounts }

    $counts = @{}
    $batch = [System.Collections.Generic.List[PSCustomObject]]::new()
    $byId  = @{}
    foreach($t in $script:IntuneTypes) {
        if(-not $t._API) { continue }
        # Hit just _API + $top=1 + $count=true. Counts reflect the endpoint
        # total, not any _QueryList-filtered view - good enough for a glance;
        # users wanting filtered counts open the type.
        $url = "$($t._API)?`$top=1&`$count=true"
        $bid = [string]$batch.Count
        [void]$batch.Add([PSCustomObject]@{
            id      = $bid
            method  = "GET"
            url     = $url
            # ConsistencyLevel: eventual is required by Graph for $count on
            # several endpoints. Harmless where unnecessary.
            headers = @{ ConsistencyLevel = "eventual"; Accept = "application/json" }
        })
        $byId[$bid] = $t.Id
    }
    if($batch.Count -gt 0) {
        try {
            $results = @(Invoke-GraphBatchRequest -BatchObjects $batch -BatchType "MenuCounts" -TokenId $TokenId -IncludedFailed -SkipWarnings)
            foreach($r in $results) {
                $tid = $byId["$($r.Id)"]
                if(-not $tid) { continue }
                $count = $null
                if($r.body -and $r.body.PSObject.Properties['@odata.count']) {
                    try { $count = [int]$r.body.'@odata.count' } catch { }
                }
                if($null -ne $count) { $counts[$tid] = $count }
            }
        }
        catch { Write-LogError "Get-IntuneTypeCounts: count batch failed" $_.Exception }
    }

    $script:_intuneTypeCounts = $counts
    return $counts
}

# Honour the two cache-clear settings above before an export/import operation.
# Bitmask matches the original project: 1 = bulk export, 2 = manual export,
# 4 = bulk import, 8 = manual import. Called by Start-GraphBulkExport /
# Export-GraphPolicy and the import entry points in both UI backends.
function Invoke-GraphCacheClearBeforeOperation
{
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('BulkExport','ManualExport','BulkImport','ManualImport')]
        [string]$Operation,

        [Int]$TokenId = (Get-DefaultTokenId)
    )

    $bit = switch($Operation) {
        'BulkExport'   { 1 }
        'ManualExport' { 2 }
        'BulkImport'   { 4 }
        'ManualImport' { 8 }
    }

    $value = (Get-SettingValue "ClearCacheBeforeExportImport" "1") -as [int]
    if($null -eq $value) { $value = 1 }
    if(($value -band $bit) -eq 0) { return }

    if((Get-SettingValue "ClearAllObjectsFromCache") -eq $true) {
        # Clear every tenant-tagged cache entry (dependency objects, AAD
        # objects, baseline templates, ADMX definitions) for every tenant seen
        # this session. The cache store lives in Core.ps1 but shares this
        # module's scope; enumerate it to find the per-tenant tags.
        Write-Log "Clearing all cached tenant objects before $Operation"
        $tenantTags = @($script:cacheObjects.Values |
            ForEach-Object { $_.Tags } |
            Where-Object { $_ -like 'TenantCache_*' } |
            Sort-Object -Unique)
        foreach($tag in $tenantTags) {
            Clear-CacheObject -Tags $tag -Force
        }
        if($null -ne $script:GraphGetDependencyPoliciesByTenant) {
            $script:GraphGetDependencyPoliciesByTenant.Clear()
        }
    }
    else {
        # Default: only the current tenant's dependency objects (Entra groups,
        # scope tags, assignment filters) so re-runs pick up fresh data.
        # Get-OperationTokenInfo, not Get-TokenInfo: this resolves the ONE tenant whose
        # cache is being cleared, and Get-TokenInfo's 0 means "every token", which would
        # hand Clear-TenantCache an ARRAY of tenant ids and clear nothing. Not reachable
        # today - the -TokenId default and both public callers resolve Get-DefaultTokenId
        # to a real id, and a 0 there means nothing is signed in - so this is the
        # accessor matching the intent, not a live bug.
        $tokenInfo = Get-OperationTokenInfo $TokenId
        if($tokenInfo -and $tokenInfo.TenantId) {
            Write-Log "Clearing cached dependency objects for tenant $($tokenInfo.TenantId) before $Operation"
            Clear-TenantCache -TenantId $tokenInfo.TenantId
        }
    }
}

# The "Silent/Batch Job" (GraphSilent) settings were removed 2026-08-14: nothing in
# this codebase ever read them, batch/automation callers authenticate explicitly via
# Connect-IntuneManagement before starting work, and credentials (GraphAzureAppSecret
# was stored PLAINTEXT) do not belong in settings. One-shot cleanup of any previously
# saved values - remove this block after the next release.
foreach($legacyKey in 'GraphAzureAppId','GraphAzureAppSecret','GraphAzureAppCert','GraphAzureAppLogin','GraphAzureTenantId') {
    try { Remove-SettingStoreValue "IntuneManager" $legacyKey } catch { }
}

# RefreshObjectsAfterCopy / AllowDelete / AllowBulkDelete are UI-only settings
# (button gating and post-copy refresh) - registered in
# UI/Classes/UICommonSettings.ps1, shared by both backends.

Add-SettingsObject -Title "Expand assignments" -Key "ExpandAssignments" -Type "Boolean" `
-Description "Expand assignments when listing objects. This can be used in custom columns based on assignment info" -DefaultValue $true  `
-SubPath "IntuneManager" -Section "GraphGeneral"

Add-SettingsObject -Title "Use Graph 1.0 (Not Recommended)" -Key "UseGraphV1" -Type "Boolean" `
-Description "This will use production verionof graph, v1.0. Note: Thot officially supported since this can have unpredicted results. Some parts will require Beta version of Graph." -DefaultValue $false  `
-SubPath "IntuneManager" -Section "GraphGeneral"

# FormatOMAURI is a UI-only setting (ADMX tools display) - registered in
# UI/Classes/UICommonSettings.ps1, shared by both backends.

$script:IntuneGroups = @()
$script:IntuneTypes = @()

# Accessors backing the -PolicyType / -PolicyGroup ArgumentCompleters on
# Get-GraphPolicies and Start-GraphBulkDocumentation (invoked via
# & (Get-Module IntuneManagement) { Get-IntunePolicyTypeValues }). ArgumentCompleter
# is used instead of a [ValidateSet([IValidateSetValuesGenerator])] so the module
# still imports on Windows PowerShell 5.1 (that interface is PS7-only).
function Get-IntunePolicyTypeValues {
    if (-not $script:IntuneTypes) { return @() }
    return @($script:IntuneTypes | ForEach-Object { $_.Id } | Sort-Object)
}

function Get-IntunePolicyGroupValues {
    if (-not $script:IntuneGroups) { return @() }
    return @($script:IntuneGroups | ForEach-Object { $_.Id } | Sort-Object)
}

# Add Event Handlers
Add-AppEventHandler "AppInitialized" "Invoke-IntuneEventAppInitialized"
Add-AppEventHandler "SettingsUpdated" "Invoke-IntuneEventSettingsUpdated"

#region Event functions

function Invoke-IntuneEventSettingsUpdated
{
    [CmdLetbinding()]
    param()
}

function Invoke-IntuneEventAppInitialized
{
    [CmdLetbinding()]
    param()

    # Load Intune Policy Group objects
    $script:IntuneGroups += Get-SubClasses "IntunePolicyGroupBase" | ForEach-Object { 
        try {
            Get-SingletonObject $_.Name
        }
        catch {}
    }

    # Load Intune Policy Type objects
    Get-SubClasses "IntunePolicyTypeBase" | ForEach-Object {
        if(Test-ClassIsAbstract $_) { return }
        try {
            $tmpClass = Get-SingletonObject $_.Name
            if($null -ne $tmpClass.PolicyGroup) {
                $script:IntuneTypes += $tmpClass
            }
            else {
                # A concrete IntunePolicyTypeBase with no PolicyGroup never joins
                # $script:IntuneTypes, so it silently won't appear in the UI, be
                # documentable, or be reachable by -PolicyType. That is almost always a
                # bug: the type forgot to set _PolicyGroup + call AddPolicyType in Init().
                # Surface it instead of dropping it silently. (If the class is meant to
                # be a base, mark it with 'static [bool] $IsAbstract = $true'.)
                Write-Log "Policy type '$($_.Name)' has no PolicyGroup and was skipped - it will not appear in the UI or be documentable. Set _PolicyGroup and call AddPolicyType in Init(), or mark the class 'static [bool] `$IsAbstract = `$true' if it is a base class." 2
            }
        }
        catch {}
    }

    $script:intuneManagerTypeAPIHashTable = @{}

    foreach($intuneType in $script:IntuneTypes) {

        if($script:intuneManagerTypeAPIHashTable.ContainsKey($intuneType.API) -eq $false) {
            $script:intuneManagerTypeAPIHashTable.Add($intuneType.API, @())
        }
        $script:intuneManagerTypeAPIHashTable[$intuneType.API] += $intuneType
    }

    # Tab-completion for -PolicyType / -PolicyGroup comes from ArgumentCompleters on
    # the parameters (Get-GraphPolicies / Start-GraphBulkDocumentation), backed by
    # Get-IntunePolicyTypeValues / Get-IntunePolicyGroupValues above.

    # -NoDownload: this runs on AppInitialized, i.e. during Import-Module. Warming
    # the cache is worth doing when the file is already there, but importing the
    # module must never block on fetching 7-8 MB of Graph beta $metadata - the
    # consumers that actually need it (Get-GraphObjectClassName and friends) fetch
    # on first use.
    Get-GraphMetaData -NoDownload

    Invoke-PolicyTypeMetadataValidation
}

#endregion

#region Import / Export functions
function Add-IntuneManagerExportProperties
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyTypeBase[]]
        $InputObject,
        [Parameter(Mandatory = $true)]
        [IntuneManagerExportSettings]
        $ExportSettings
    )

    Begin { Write-Log "Add Export Object Properties - Start" }

    Process {
        foreach($policyType in $InputObject) {
            $policyType.AddExportProperties($ExportSettings)
        }
    }
    
    End { Write-Log "Add Export Object Properties - Done" }
}

function Add-IntuneManagerImportProperties
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IntunePolicyTypeBase[]]
        $InputObject,
        [Parameter(Mandatory = $true)]
        [IntuneManagerImportSettings]
        $ImportSettings
    )

    Begin { Write-Log "Add Import Object Properties - Start" }

    Process {
        foreach($policyType in $InputObject) {
            $policyType.AddImportProperties($ImportSettings)
        }
    }
    
    End { Write-Log "Add Import Object Properties - Done" }

}

#endregion

function Get-PolicyTypeFromURL
{
    param([URI]$URI)

    $policyType = $null

    $API = $uri.Segments[2..4] -join ""

    if(-not $API) { return $null }

    $endIndex = $API.IndexOf('(')
    if($endIndex -gt 0) {
        $API = $API.SubString(0,$endIndex)
    }

    $policyType = $script:IntuneTypes | Where-Object API -eq "$($API.Trim("/"))"

    if($policyType) {
        Write-LogDebug "Found PolicyType $($policyType.ID)"
    }
    else {
        Write-Log "Could not find PolicyType from API $API" 2
    }

    return $policyType
}

function Get-PolicyTypeFromID
{
    param([String]$ID)

    # Internal lookup - Import-GraphPolicy groups objects by their OWN type id, so a
    # miss here is an inconsistency worth a warning, not a user-supplied selector.
    # The selector contract (collect + Write-Error) is Resolve-IntuneTargetSelector.
    $policyType = Get-IntunePolicyTypeById $ID

    if($policyType) {
        Write-LogDebug "Found PolicyType $($policyType.Title) ($($policyType.ID))"
    }
    else {
        Write-Log "Could not find PolicyType from ID $ID" 2
    }

    return $policyType
}

# The two registry lookups every id-based accessor shares. No logging, no errors:
# the caller decides what a miss means.
function Get-IntunePolicyTypeById
{
    param([string]$Id)
    return ($script:IntuneTypes | Where-Object Id -eq $Id)
}

function Get-IntunePolicyGroupById
{
    param([string]$Id)
    return ($script:IntuneGroups | Where-Object Id -eq $Id)
}

#region Target selectors (-PolicyType / -PolicyGroup)
# Shared resolution of the -PolicyType / -PolicyGroup ids that the bulk drivers,
# Get-GraphPolicies, Compare-GraphPolicy and the documentation engine all accept.
# Each used to carry its own copy of the lookup and each dealt with an id that
# matched nothing by writing a level-2 log line and moving on - so a saved settings
# file carrying a stale id produced a success-shaped result that quietly did less.
#
# Now: every unknown id is logged, collected (surfaced in the UI summaries), and
# raised as a NON-TERMINATING error so a script gets fail-fast via -ErrorAction Stop.
# Lives here rather than in its own file because this is where the registry
# ($script:IntuneTypes / $script:IntuneGroups) and its other accessors already are.

function Resolve-IntuneTargetSelector
{
    param(
        [string]$Id,

        # Mandatory on purpose: an omitted -Kind coerces to '' and would silently
        # take the group branch, reporting a real type id as unknown.
        [Parameter(Mandatory = $true)]
        [ValidateSet('PolicyType', 'PolicyGroup')]
        [string]$Kind,

        # Collects "<Kind> '<Id>'" for anything that does not resolve. Optional -
        # pass $null when the caller does not report unresolved selectors.
        $Unknown,

        # Prefixes the log line so a shared message still says who was asking.
        [string]$Caller
    )

    # No blank-id short-circuit: a '' from a malformed settings file must be
    # reported like any other unknown, not dropped silently.
    $found = if($Kind -eq 'PolicyType') { Get-IntunePolicyTypeById $Id } else { Get-IntunePolicyGroupById $Id }
    if($found) { return $found }

    if($null -ne $Unknown) { [void]$Unknown.Add("$Kind '$Id'") }

    $prefix = if($Caller) { "$($Caller): " } else { "" }
    Write-Log "$($prefix)Unknown $Kind '$Id' - skipped. No $Kind with that id is registered in this session." 2

    # Non-terminating, so the caller's loop still reaches every id and reports
    # them all; -ErrorAction Stop on the public cmdlet makes it terminate.
    Write-Error -Message "$($prefix)Unknown $Kind '$Id' - no $Kind with that id is registered." `
        -ErrorId 'IntuneManagement.UnknownSelector' -Category ObjectNotFound -TargetObject $Id

    return $null
}

# Resolves both selector lists at once and returns { Types; Groups; Unknown }.
#   Types   - every selected type, plus every type of every selected group, unique
#             by Id in first-seen order (a type picked directly AND via its group
#             appears once - the drivers used to process it twice).
#   Groups  - the resolved group objects, unique by Id.
#   Unknown - "<Kind> '<Id>'" for every id that matched nothing.
# All three are always arrays, never $null, so callers can AddRange / index freely.
function Resolve-IntuneTargetSelectors
{
    param(
        [string[]]$PolicyType,
        [string[]]$PolicyGroup,
        [string]$Caller
    )

    $unknown      = [System.Collections.Generic.List[string]]::new()
    $types        = [System.Collections.Generic.List[object]]::new()
    $groups       = [System.Collections.Generic.List[object]]::new()
    $seenTypeIds  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $seenGroupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    # `$null -ne $_` and not @($PolicyType): an omitted parameter is $null, and
    # @($null) is a one-element array that would report a phantom '' selector.
    # A genuine '' inside the list still flows through and is reported.
    foreach($id in @($PolicyType | Where-Object { $null -ne $_ }))
    {
        $found = Resolve-IntuneTargetSelector -Id $id -Kind PolicyType -Unknown $unknown -Caller $Caller
        if($found -and $seenTypeIds.Add([string]$found.Id)) { [void]$types.Add($found) }
    }
    foreach($gid in @($PolicyGroup | Where-Object { $null -ne $_ }))
    {
        $grp = Resolve-IntuneTargetSelector -Id $gid -Kind PolicyGroup -Unknown $unknown -Caller $Caller
        if(-not $grp) { continue }
        if($seenGroupIds.Add([string]$grp.Id)) { [void]$groups.Add($grp) }
        foreach($pt in @($grp.PolicyTypes))
        {
            if($pt -and $seenTypeIds.Add([string]$pt.Id)) { [void]$types.Add($pt) }
        }
    }

    return [PSCustomObject]@{
        Types   = $types.ToArray()
        Groups  = $groups.ToArray()
        Unknown = $unknown.ToArray()
    }
}

# Formats the collected list for a result summary / message box. Returns $null when
# nothing was skipped so callers can test it directly.
function Get-IntuneUnknownSelectorSummary
{
    param($Unknown)

    if($null -eq $Unknown) { return $null }
    $items = @($Unknown)
    if($items.Count -eq 0) { return $null }

    return "$($items.Count) selector(s) matched nothing and were skipped: $($items -join ', ')"
}
#endregion

function Get-PoliciesTypeFromObject
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$PolicyObject, 
        [IntunePolicyTypeBase[]]$FromPolicyTypes)

    if(-not $PolicyObject.'@odata.type' -and -not $PolicyObject.'@odata.id') {
        Write-Log "No @odata.type or @odata.id found on object. Cannot identify Policy Type" 3
        return $null
    }

    $checkPolicies = @()
    if($PolicyObject.'@odata.id') {
        $metadata = '$metadata#'
        $policyType = $PolicyObject.'@odata.id'

        if(($index = $policyType.IndexOf($metadata)) -and $index -gt 0) { #     
            $subString = $policyType.Substring($index + $metadata.Length)
        }
        else {
            $subString = $policyType
        }            
        $endIndex = $subString.IndexOf('(')
        if($endIndex -lt 1) {
            $endIndex = $subString.IndexOf('/$entity')
        }

        $FromPolicyTypeIds = @()
        if($FromPolicyTypes) {
            $FromPolicyTypeIds += $FromPolicyTypes.ID
        }

        if($endIndex -gt 0) {
            $subString = $subString.Substring(0, $endIndex)
            if($script:intuneManagerTypeAPIHashTable.ContainsKey($subString)) {
                if($FromPolicyTypeIds.Count -eq 0 -or $script:intuneManagerTypeAPIHashTable[$subString].Id -in $FromPolicyTypeIds) {
                    $checkPolicies += $script:intuneManagerTypeAPIHashTable[$subString]
                }
            }
        }
    }

    if($checkPolicies.Count -eq 0) { 
        if($FromPolicyTypes.Count -gt 0 ) {
            $checkPolicies = $FromPolicyTypes
        }
        else {
            $checkPolicies = $script:IntuneGroups.PolicyTypes
        }
    }

    foreach($policyType in ($checkPolicies | Sort-Object -Property _PolicyTypeOrder -Descending)) {
        if($policyType.CheckPolicy($PolicyObject)) {
            return $policyType
        }
    }

    # File objects carry only @odata.type (no top-level @odata.id), so the base
    # @odata.id-based CheckPolicy cannot identify them - and most policy types do not
    # override CheckPolicy at all. When the caller has already narrowed the candidates
    # to a single type (Get-PoliciesFromFolder passes only the types that own the file's
    # export subfolder) and the object is identifiable, trust that single candidate
    # instead of dropping it. Siblings that share a folder keep more than one candidate
    # and are disambiguated by their own @odata.type CheckPolicy in the loop above.
    # Excluded: a StrictODataTypeCheck type (App Protection / App Config) whose CheckPolicy
    # is a complete @odata.type matcher - if it rejected the object above, that is an
    # authoritative rejection (e.g. a foreign file misplaced in its folder), not
    # file-blindness, so it must NOT be rescued here.
    if($PSBoundParameters.ContainsKey("FromPolicyTypes") -and @($checkPolicies).Count -eq 1 -and
       -not $checkPolicies[0].StrictODataTypeCheck -and
       -not $PolicyObject.'@odata.id' -and $PolicyObject.'@odata.type') {
        Write-Log "Resolved policy type '$($checkPolicies[0].Id)' for object with OData.type = $($PolicyObject.'@odata.type') from single narrowed candidate (no CheckPolicy match)" 4
        return $checkPolicies[0]
    }

    if($null -eq $PSBoundParameters["FromPolicyTypes"]) {
        # Expected not to find some policies if policy types is passed in
        Write-Log "Could not find policy type based on object with OData.type = $($PolicyObject.'@odata.type') and OData.id = $($PolicyObject.'@odata.id')" 3
    }
    return $null
}

function Get-PoliciesFromFolder
{
    [CmdletBinding()]
    param(
        [String]
        $Path, 
        [String[]]
        $Exclude = @("*_settings.json","*_assignments.json", "MigrationTable*.json"), 
        [String[]]
        $SubFolders = @(), 
        [Alias("Recurse")]
        [switch]
        $SearchSubFolders, 
        [int]
        $Depth = 0,
        [IntunePolicyTypeBase[]]
        $PolicyTypes
    )

    if(-not $Path -or (Test-Path $Path -PathType Container) -eq $false) { 
        Write-Log "Path '$path' not found. Cannot get policies." 3
        return
    }

    $params = @{}
    if($Exclude)
    {
        $params.Add("Exclude", $Exclude)
    }

    if($SearchSubFolders -eq $true) {
        $params.Add("Recurse", $true)
    }

    if($Depth -gt 0) {
        $params.Add("Depth", $Depth)
    }

    $policyObjects = @()

    $paths = @()

    if($SubFolders.Count -gt 0) {
        Get-ChildItem -path $Path -Directory | Where-Object { $_.Name -in $SubFolders -and $_ -is [IO.DirectoryInfo] } | ForEach-Object { $paths += $_.FullName }
    }
    else {
        $paths += $path
    }

    $migFileObj = $null
    $lookForMigFile = $true

    # Map each candidate PolicyType by the export subfolder name it writes to (Folder,
    # with Id as a fallback). Exports store objects in per-type subfolders, so the
    # file's parent folder name narrows the candidate set passed to the resolver.
    # Without this, every type competes on CheckPolicy - and most types identify only
    # by @odata.type (or only by @odata.id, which files lack), so they misresolve.
    # Shared folders (e.g. ReusableSettings, EnrollmentRestrictions) keep more than one
    # candidate, disambiguated by each type's own @odata.type CheckPolicy.
    $typesByFolder = @{}
    if($PolicyTypes) {
        foreach($pt in $PolicyTypes) {
            foreach($key in @($pt.Folder, $pt.Id)) {
                if(-not $key) { continue }
                $k = ([string]$key).ToLowerInvariant()
                if(-not $typesByFolder.ContainsKey($k)) { $typesByFolder[$k] = @() }
                if($pt -notin $typesByFolder[$k]) { $typesByFolder[$k] += $pt }
            }
        }
    }

    foreach($policyPath in $paths) {
        foreach($file in (Get-ChildItem -path "$policyPath\*.json" @params))
        {
            if($null -eq $migFileObject -and $lookForMigFile) {
                $migFilePath = Get-GraphMigrationTableFromPath $file.DirectoryName
                if($migFilePath) {
                    $migFileObj = ConvertFrom-Json ([IO.File]::ReadAllText($migFilePath))
                }
                $lookForMigFile = $false
            }

            $fileParams = @{}
            if($PolicyTypes) {
                $folderName = ([string]$file.Directory.Name).ToLowerInvariant()
                if($typesByFolder.ContainsKey($folderName)) {
                    $fileParams["FromPolicyTypes"] = $typesByFolder[$folderName]
                }
                else {
                    $fileParams["FromPolicyTypes"] = $PolicyTypes
                }
            }

            $graphObj = Get-GraphPolicyFromFile $file @fileParams -TenantId $migFileObj.TenantID
            if(-not $graphObj) { continue }

            $policyObjects += $graphObj
        }
    }

    return @($policyObjects)
}

#region Import match-resolution helpers
# Lifted out of UI/Extensions/IntuneManagerUI.ps1 + the Avalonia copy in
# UI/Avalonia/Extensions/IntuneManagerImportUI.ps1 — pure data logic with no
# control dependencies. Both UI trees call Resolve-IntuneImportUpdateTarget
# from their import dialogs to decide Create / Update / Skip / Ambiguous; the
# helpers also feed Update-IntuneImportMatchInfo (still per-tree because it
# pokes UI controls).

function Get-IntuneImportPolicyReferenceTokens
{
    param(
        [string]$Name,
        [string]$PolicyReferenceRegex = $null
    )

    if([string]::IsNullOrWhiteSpace($Name)) { return @() }

    $tokens = [System.Collections.Generic.List[string]]::new()
    $defaultRegex = '(?i)(?:\[(?<ref>[A-Z][A-Z0-9]{1,15}-\d{2,10})\])|(?<ref>\b[A-Z][A-Z0-9]{1,15}-\d{2,10}\b)'
    $regex = $PolicyReferenceRegex
    if([string]::IsNullOrWhiteSpace($regex)) { $regex = Get-SettingValue "ImportMatchPolicyReferenceRegex" }
    if([string]::IsNullOrWhiteSpace($regex)) { $regex = $defaultRegex }

    try {
        $matchResults = [regex]::Matches($Name, $regex)
    }
    catch {
        Write-Log "Invalid ImportMatchPolicyReferenceRegex setting. Using default policy reference regex. Error: $($_.Exception.Message)" 2
        $matchResults = [regex]::Matches($Name, $defaultRegex)
    }

    foreach($match in $matchResults) {
        if(-not $match.Groups['ref'] -or -not $match.Groups['ref'].Success) { continue }

        $token = $match.Groups['ref'].Value
        if($token -and -not $tokens.Contains($token.ToUpperInvariant())) {
            [void]$tokens.Add($token.ToUpperInvariant())
        }
    }

    return $tokens.ToArray()
}

function Normalize-IntuneImportPolicyName
{
    param(
        [string]$Name,
        [string]$OrganizationName = (Get-CurrentOrganizationName),
        [string]$OrganizationTokens = $null
    )

    if([string]::IsNullOrWhiteSpace($Name)) { return "" }

    $normalized = $Name
    $orgTokens = [System.Collections.Generic.List[string]]::new()
    # Token spellings come from Internal/ExportTokens.ps1 so they are declared in
    # one place. Every token is stripped regardless of what the export setting is
    # set to - the name being matched can come from any older export.
    foreach($orgToken in @((Get-GraphExportTokenStrings) + @($OrganizationName, (Get-CurrentTenantId)))) {
        if([string]::IsNullOrWhiteSpace($orgToken) -or $orgTokens.Contains([string]$orgToken)) { continue }
        [void]$orgTokens.Add([string]$orgToken)
    }

    $extraOrgTokens = $OrganizationTokens
    if($null -eq $extraOrgTokens) { $extraOrgTokens = Get-SettingValue "ImportMatchOrganizationTokens" "" }
    if(-not [string]::IsNullOrWhiteSpace($extraOrgTokens)) {
        foreach($orgToken in ([regex]::Split([string]$extraOrgTokens, '[,;\r\n]+'))) {
            if([string]::IsNullOrWhiteSpace($orgToken)) { continue }

            $trimmed = $orgToken.Trim()
            if(-not $orgTokens.Contains($trimmed)) {
                [void]$orgTokens.Add($trimmed)
            }
        }
    }

    foreach($orgToken in $orgTokens) {
        if([string]::IsNullOrWhiteSpace($orgToken)) { continue }

        $escaped = [regex]::Escape([string]$orgToken)
        $normalized = [regex]::Replace($normalized, "(?i)\[\s*$escaped\s*\]", " ")
        $normalized = [regex]::Replace($normalized, "(?i)\(\s*$escaped\s*\)", " ")
        $normalized = [regex]::Replace($normalized, "(?i)(?<![\p{L}\p{N}])$escaped(?![\p{L}\p{N}])", " ")
    }

    # –/— = en/em dash. Spelled as regex escapes so the source stays
    # ASCII - raw dash chars in the pattern would be misread when PS5.1 loads
    # this BOM-less file as ANSI.
    $normalized = [regex]::Replace($normalized, '^[\s\-_\u2013\u2014:|\\/]+', '')
    $normalized = [regex]::Replace($normalized, '[\s\-_\u2013\u2014:|\\/]+$', '')
    $normalized = [regex]::Replace($normalized, '\s+', ' ')
    return $normalized.Trim().ToUpperInvariant()
}

function New-IntuneImportMatchResult
{
    param(
        [string]$Action,
        [string]$Strategy = "",
        $Target = $null,
        [string]$Message = ""
    )

    [PSCustomObject]@{
        Action   = $Action
        Strategy = $Strategy
        Target   = $Target
        Message  = $Message
    }
}

function Resolve-IntuneImportUpdateTarget
{
    param(
        [Parameter(Mandatory)]$ImportPolicy,
        $ExistingPolicies,
        [bool]$SameTenant,
        [string]$ImportType = "alwaysImport",
        [string]$OrganizationName = (Get-CurrentOrganizationName),
        $EnablePolicyToken = $null,
        [string]$PolicyReferenceRegex = $null,
        $EnableNormalizedName = $null,
        [string]$OrganizationTokens = $null
    )

    $existing = @($ExistingPolicies | Where-Object { $_ -and $_.PolicyType -and $_.PolicyType.Id -eq $ImportPolicy.PolicyType.Id })

    if($ImportType -eq "alwaysImport") {
        return New-IntuneImportMatchResult -Action "Create" -Message "Create new object"
    }

    # Action when a single existing object matched: update and the two replace
    # variants target it; skipIfExist skips it.
    $matchedAction = switch($ImportType) {
        "update"                   { "Update" }
        "replace"                  { "Replace" }
        "replace_with_assignments" { "Replace" }
        default                    { "Skip" }
    }
    $matchedVerb = switch($matchedAction) {
        "Update"  { "Update by" }
        "Replace" { "Replace by" }
        default   { "Existing object matched by" }
    }

    if($SameTenant -and $ImportPolicy.Id) {
        $idMatches = @($existing | Where-Object { $_.Id -eq $ImportPolicy.Id })
        if($idMatches.Count -eq 1) {
            return New-IntuneImportMatchResult -Action $matchedAction -Strategy "Id" -Target $idMatches[0] -Message "$matchedVerb ID"
        }
        elseif($idMatches.Count -gt 1) {
            return New-IntuneImportMatchResult -Action "Ambiguous" -Strategy "Id" -Message "Multiple existing objects matched ID $($ImportPolicy.Id)"
        }
    }

    $nameMatches = @($existing | Where-Object { $_.Name -eq $ImportPolicy.Name })
    if($nameMatches.Count -eq 1) {
        return New-IntuneImportMatchResult -Action $matchedAction -Strategy "ExactName" -Target $nameMatches[0] -Message "$matchedVerb exact name"
    }
    elseif($nameMatches.Count -gt 1) {
        return New-IntuneImportMatchResult -Action "Ambiguous" -Strategy "ExactName" -Message "Multiple existing objects matched name '$($ImportPolicy.Name)'"
    }

    $policyTokenEnabled = if($null -ne $EnablePolicyToken) { [bool]$EnablePolicyToken } else { (Get-SettingValue "ImportMatchEnablePolicyToken") -ne $false }
    if($policyTokenEnabled) {
        $tokens = @(Get-IntuneImportPolicyReferenceTokens -Name $ImportPolicy.Name -PolicyReferenceRegex $PolicyReferenceRegex)
        if($tokens.Count -gt 0) {
            $tokenMatches = @($existing | Where-Object {
                $existingTokens = @(Get-IntuneImportPolicyReferenceTokens -Name $_.Name -PolicyReferenceRegex $PolicyReferenceRegex)
                @($existingTokens | Where-Object { $_ -in $tokens }).Count -gt 0
            })
            if($tokenMatches.Count -eq 1) {
                return New-IntuneImportMatchResult -Action $matchedAction -Strategy "PolicyToken" -Target $tokenMatches[0] -Message "$matchedVerb policy token $($tokens -join ', ')"
            }
            elseif($tokenMatches.Count -gt 1) {
                return New-IntuneImportMatchResult -Action "Ambiguous" -Strategy "PolicyToken" -Message "Multiple existing objects matched policy token $($tokens -join ', ')"
            }
        }
    }

    $normalizedNameEnabled = if($null -ne $EnableNormalizedName) { [bool]$EnableNormalizedName } else { (Get-SettingValue "ImportMatchEnableNormalizedName") -ne $false }
    if($normalizedNameEnabled) {
        $normalizedImportName = Normalize-IntuneImportPolicyName -Name $ImportPolicy.Name -OrganizationName $OrganizationName -OrganizationTokens $OrganizationTokens
        if($normalizedImportName) {
            $normalizedMatches = @($existing | Where-Object {
                (Normalize-IntuneImportPolicyName -Name $_.Name -OrganizationName $OrganizationName -OrganizationTokens $OrganizationTokens) -eq $normalizedImportName
            })
            if($normalizedMatches.Count -eq 1) {
                return New-IntuneImportMatchResult -Action $matchedAction -Strategy "NormalizedName" -Target $normalizedMatches[0] -Message "$matchedVerb normalized name"
            }
            elseif($normalizedMatches.Count -gt 1) {
                return New-IntuneImportMatchResult -Action "Ambiguous" -Strategy "NormalizedName" -Message "Multiple existing objects matched normalized name '$normalizedImportName'"
            }
        }
    }

    if($ImportType -eq "update") {
        return New-IntuneImportMatchResult -Action "Skip" -Message "No matching object found to update"
    }

    # replace / replace_with_assignments with no existing match fall through
    # to a plain import - same as the original project.
    return New-IntuneImportMatchResult -Action "Create" -Message "Create new object"
}

function Invoke-IntuneImportReplace
{
    # Replace-mode import: import the file as a NEW object, copy assignments,
    # then delete the existing object.
    #   replace                  - assignments are copied from the EXISTING object
    #                              (live assignments survive the swap).
    #   replace_with_assignments - assignments come from the import FILE.
    param(
        [Parameter(Mandatory)]$ImportPolicy,
        [Parameter(Mandatory)]$Target,
        [ValidateSet("replace", "replace_with_assignments")]
        [string]$ImportType = "replace",
        [int]$TokenId = (Get-DefaultTokenId)
    )

    Write-Log "Replace $($Target.PolicyType.Title) object '$($Target.Name)' ($($Target.Id)) with import of '$($ImportPolicy.Name)' ($ImportType)"

    if($ImportType -eq "replace")
    {
        # The standard import path applies file assignments via
        # Import-GraphObjectAssignment; drop them so only the existing
        # object's live assignments end up on the new object.
        Remove-Property $ImportPolicy.JsonObject "assignments"
    }

    $imported = @($ImportPolicy | Import-GraphPolicy -TokenId $TokenId)
    $newObj = if($imported.Count -gt 0) { $imported[0].ImportedObject } else { $null }
    if(-not $newObj)
    {
        Write-Log "Replace aborted for '$($Target.Name)' ($($Target.Id)): import of '$($ImportPolicy.Name)' failed - existing object kept" 3
        return $null
    }

    # Replace creates a NEW object with a NEW Id, so any Policy Set that bundled
    # the old object now references a deleted payload. Find those Policy Sets by
    # inspecting their items directly - NOT via the policy's assignments. A
    # policy only carries a source='policySets' assignment when the set is
    # assigned to groups, so assignment-based detection (the original project's
    # approach) silently misses unassigned Policy Sets. Querying the sets is
    # correct for both. Done before the old object is deleted so both payloads
    # still resolve during the swap.
    $policySetIds = @(Find-IntunePolicySetsForPayload -PayloadId ([string]$Target.Id) -TokenId $TokenId)
    foreach($policySetId in $policySetIds)
    {
        try
        {
            Update-IntunePolicySetItemPayload -PolicySetId $policySetId -OldPayloadId ([string]$Target.Id) -NewPayloadId ([string]$newObj.Id) -TokenId $TokenId | Out-Null
        }
        catch
        {
            Write-LogError "Replace: failed to re-point Policy Set $policySetId from $($Target.Id) to $($newObj.Id)" $_.Exception
        }
    }

    if($ImportType -eq "replace")
    {
        # CopyAssignments reads the existing object's assignments, so load them.
        try { $Target.PolicyType.GetFullObject($Target) | Out-Null } catch { }
        # Out-Null: keep side-effect pipeline output out of this function's
        # return value - it must emit ONLY $newObj.
        Import-GraphObjectAssignment $newObj $Target -CopyAssignments | Out-Null
    }

    try
    {
        # Remove-GraphPolicy emits the deleted policy; suppress it so the
        # function returns only $newObj (callers do `$x = Invoke-...`).
        $Target | Remove-GraphPolicy -Confirm:$false | Out-Null
        Write-Log "Replace finished: '$($Target.Name)' ($($Target.Id)) deleted, replaced by $($newObj.Id)"
    }
    catch
    {
        Write-LogError "Replace: new object $($newObj.Id) imported but failed to delete previous object '$($Target.Name)' ($($Target.Id))" $_.Exception
    }

    return $newObj
}

function Update-IntunePolicySetItemPayload
{
    # Re-point a Policy Set item from one payload object to another. Used by
    # Replace import: a replaced policy gets a NEW Id, so any Policy Set that
    # bundled the old policy must be updated to reference the new one. Ported
    # from the original project's Update-EMPolicySetAssignment.
    #
    # The Policy Set /update contract takes added/updated/deleted item lists.
    # We clone the existing item, swap its payloadId to the new object, add the
    # clone, and delete the old item by its item Id - so priority/settings on
    # the item are preserved.
    param(
        [Parameter(Mandatory)][string]$PolicySetId,
        [Parameter(Mandatory)][string]$OldPayloadId,
        [Parameter(Mandatory)][string]$NewPayloadId,
        [int]$TokenId = (Get-DefaultTokenId)
    )

    $psObj = Invoke-MSGraphAPI -Url "deviceAppManagement/policySets/$PolicySetId`?`$expand=assignments,items" -ODataMetadata "minimal" -TokenId $TokenId
    if(-not $psObj)
    {
        Write-Log "Replace: Policy Set $PolicySetId not found; cannot re-point payload $OldPayloadId" 2
        return $false
    }

    $curItem = $psObj.items | Where-Object { $_.payloadId -eq $OldPayloadId } | Select-Object -First 1
    if(-not $curItem)
    {
        Write-Log "Replace: Policy Set '$($psObj.displayName)' ($PolicySetId) has no item for payload $OldPayloadId" 2
        return $false
    }

    $curItemClone = $curItem | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $newItem      = $curItem | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $newItem.payloadId = $NewPayloadId
    if($newItem.guidedDeploymentTags -is [String] -and [String]::IsNullOrEmpty($newItem.guidedDeploymentTags))
    {
        $newItem.guidedDeploymentTags = @()
    }

    # Keep only the properties the /update add contract accepts; everything
    # else (id, status, priority, OData read-only fields) must be stripped.
    # This matches the original project's proven set - adding priority/itemType
    # made the add contract reject the payload. Comparison is case-insensitive
    # so JSON casing of the live item doesn't matter.
    $keepProperties = @('@odata.type','payloadId','settings','guidedDeploymentTags')
    foreach($prop in @($newItem.PSObject.Properties | Where-Object { $_.Name -notin $keepProperties }))
    {
        Remove-Property $newItem $prop.Name
    }

    $update = [ordered]@{
        addedPolicySetItems   = @($newItem)
        updatedPolicySetItems = @()
        deletedPolicySetItems = @([string]$curItemClone.id)
    }
    $json = $update | ConvertTo-Json -Depth 20

    Write-Log "Replace: re-pointing Policy Set '$($psObj.displayName)' ($PolicySetId) item from payload $OldPayloadId to $NewPayloadId"
    $resp = Invoke-MSGraphAPI -Url "deviceAppManagement/policySets/$PolicySetId/update" -HttpMethod POST -Content $json -TokenId $TokenId -FullResponseObject
    if($resp -and $resp.Success)
    {
        return $true
    }

    $detail = if($resp) { "$($resp.StatusCode) $($resp.ErrorMessage)" } else { "no response" }
    Write-Log "Replace: Policy Set $PolicySetId /update failed for payload swap $OldPayloadId -> $NewPayloadId ($detail)" 3
    return $false
}

function Find-IntunePolicySetsForPayload
{
    # Return the Ids of every Policy Set whose items reference $PayloadId.
    # Inspects the sets' items directly so it finds membership regardless of
    # whether the Policy Set is assigned (an unassigned set produces no
    # source='policySets' assignment on the member policy).
    param(
        [Parameter(Mandatory)][string]$PayloadId,
        [int]$TokenId = (Get-DefaultTokenId)
    )

    $found = [System.Collections.Generic.List[string]]::new()

    # $expand is rejected on the policySets COLLECTION (Graph 400 "Expand is
    # not allowed" - same restriction as $expand=assignments), so list the ids
    # then GET each set's items individually (single-set $expand=items works).
    $list = Invoke-MSGraphAPI -Url "deviceAppManagement/policySets?`$select=id" -ODataMetadata minimal -TokenId $TokenId
    foreach($ps in @($list.value))
    {
        if(-not $ps.id) { continue }
        $full = $null
        try { $full = Invoke-MSGraphAPI -Url "deviceAppManagement/policySets/$($ps.id)?`$expand=items" -ODataMetadata minimal -TokenId $TokenId } catch { $full = $null }
        if($full -and @($full.items | Where-Object { $_.payloadId -eq $PayloadId }).Count -gt 0) { [void]$found.Add([string]$ps.id) }
    }
    return @($found | Select-Object -Unique)
}

#endregion Import match-resolution helpers
