function Add-DocumentationOutputProvider {
    param([PSCustomObject]$Provider)
    [DocumentationRegistry]::RegisterOutput($Provider)
}

# Public registration wrapper for phase-3 input providers (Settings Catalog, ADMX,
# Intent, ComplianceV2, generic Profile). Symmetric with Add-DocumentationOutputProvider.
# Provider hashtable shape:
#   Name      [string]      provider identifier (e.g. 'SettingsCatalog')
#   Match     [scriptblock] takes $PolicyObject, returns $true if this provider claims it
#   Translate [scriptblock] takes ($PolicyObject, $Context); fills the context
function Add-DocumentationInputProvider {
    param([PSCustomObject]$Provider)
    [DocumentationRegistry]::RegisterInputProvider($Provider)
}

# Public registration wrapper for phase-4 *DocHandler classes. The class is
# typically constructed at file load time; this just registers the instance.
function Add-DocumentationHandler {
    param([object]$Handler)
    [DocumentationRegistry]::RegisterHandler($Handler)
}

function Merge-DocumentationOptions {
    param([hashtable]$Options)

    $merged = [DocumentationContext]::DefaultOptions()

    # Seed the saved user preference BEFORE the caller's options, so an explicit
    # -Options value still wins. Only for keys the documentation dialog persists
    # and the engine reads on every object; without this a silent run would
    # ignore the checkbox the same user ticked in the UI.
    # Compared, not cast: the store persists the checkbox as text, and
    # [bool]"False" is $true in PowerShell - a cast here switched the fallback
    # back on for every silent run after the user had turned it off. Same
    # idiom Get-SettingValue uses for its registered Boolean settings.
    $stored = Get-DocumentationSetting 'FallbackDocumentation' $null
    if ($null -ne $stored) { $merged.FallbackDocumentation = ($stored -eq $true -or $stored -eq 'true') }

    if ($Options) {
        foreach ($key in $Options.Keys) { $merged[$key] = $Options[$key] }
    }
    if (-not $merged.Outputs) { $merged.Outputs = @{} }
    # Backward compat: the flag was renamed OfflineDocumentation -> SourceTenantUnavailable.
    # Honour the legacy key from any external -Options @{OfflineDocumentation=$true} caller,
    # but let an explicit new-key value win if both are supplied.
    if ($Options -and $Options.ContainsKey('OfflineDocumentation') -and -not ($Options.ContainsKey('SourceTenantUnavailable'))) {
        $merged.SourceTenantUnavailable = [bool]$Options['OfflineDocumentation']
    }
    $merged.Remove('OfflineDocumentation')
    return $merged
}

# True when ANY Intune tenant is reachable. Generic Intune schema (Settings
# Catalog setting definitions/categories, Compliance categories, ADMX
# definitions, built-in Endpoint Security intent templates, role
# resourceOperations) is identical on every tenant, so it can be resolved from
# whatever tenant is connected even when the SOURCE tenant of an export is gone
# (SourceTenantUnavailable). Source-tenant-specific lookups stay additionally
# gated on -not $ctx.SourceTenantUnavailable.
function Test-DocumentationGraphAvailable {
    if (-not (Get-Command Invoke-MSGraphAPI -ErrorAction SilentlyContinue)) { return $false }
    return -not [string]::IsNullOrEmpty((Get-CurrentTenantId))
}

# Storage layout for documentation settings:
#   generic / form-level keys   -> Documentation
#   per-output-provider keys    -> Documentation\<providerValue>
#
# Per-provider option keys are prefixed with the provider's registered Value
# (e.g. HTMLDocumentName for value=html, WordCoverPage for value=word). This
# function walks every registered output provider and picks the one whose
# Value is a case-insensitive prefix of the key. Longest match wins so a
# short-prefix provider can't swallow keys belonging to a longer-prefix one
# (e.g. hypothetical "H" would not steal "HTML*" keys from the html provider).
# Falls back to the generic Documentation subkey when no prefix matches, so
# form-level options (SkipDisabled, IncludeScripts, ...) route there.
#
# Adding a new output provider does not require an edit here — dropping a
# file under OutputProviders/ that calls Add-DocumentationOutputProvider
# registers the Value automatically and this lookup picks it up.
function Get-DocumentationSettingSubPath {
    param([string]$Key)
    if (-not $Key) { return 'Documentation' }

    $best = $null
    foreach ($provider in [DocumentationRegistry]::Outputs) {
        $val = [string]$provider.Value
        if (-not $val) { continue }
        if (-not $Key.StartsWith($val, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        if (-not $best -or $val.Length -gt ([string]$best.Value).Length) { $best = $provider }
    }
    if ($best) { return "Documentation\$(([string]$best.Value).ToLowerInvariant())" }
    return 'Documentation'
}

# Read a documentation setting, auto-routing to the right subpath by key prefix.
function Get-DocumentationSetting {
    param([Parameter(Mandatory)][string]$Key, $Default)
    Get-SettingStoreValue (Get-DocumentationSettingSubPath $Key) $Key $Default
}

function Get-DocumentationOutputOption {
    param(
        [Parameter(Mandatory)][string]$Output,
        [Parameter(Mandatory)][string]$Name,
        $Default
    )

    $options = $script:_docRunOptions
    if ($options -and $options.Outputs -and $options.Outputs[$Output]) {
        $providerOptions = $options.Outputs[$Output]
        if ($providerOptions -is [hashtable] -and $providerOptions.ContainsKey($Name)) {
            return $providerOptions[$Name]
        }
        if ($providerOptions.PSObject.Properties[$Name]) {
            return $providerOptions.$Name
        }
    }
    return Get-SettingStoreValue (Get-DocumentationSettingSubPath $Name) $Name $Default
}

# Read a generic (engine-wide, not per-output-provider) run option. Unlike
# Get-DocumentationOutputOption, this reads the TOP LEVEL of the merged run
# options ($script:_docRunOptions) - the same place ToResult reads
# SkipNotConfigured/SkipDisabled etc. - then falls back to the persisted
# generic Documentation setting. Use this for flags every output provider must
# honour identically (e.g. SkipDocumentInfo).
function Get-DocumentationOption {
    param(
        [Parameter(Mandatory)][string]$Name,
        $Default
    )

    $options = $script:_docRunOptions
    if ($options -is [hashtable] -and $options.ContainsKey($Name)) {
        return $options[$Name]
    }
    return Get-SettingStoreValue (Get-DocumentationSettingSubPath $Name) $Name $Default
}

function Set-DocumentationContextRunOptions {
    param(
        [DocumentationContext]$Context,
        [hashtable]$Options,
        [string]$Language
    )

    $wasOffline = $Context.SourceTenantUnavailable
    $merged = Merge-DocumentationOptions $Options
    if ($Language) { $merged.Language = $Language }
    $Context.Options = $merged
    $Context.Language = [string]$merged.Language
    $Context.PropertySeparator = [string]$merged.PropertySeparator
    $Context.ObjectSeparator = [string]$merged.ObjectSeparator
    $Context.SourceTenantUnavailable = [bool]$merged.SourceTenantUnavailable
    if ($wasOffline -ne $Context.SourceTenantUnavailable) {
        $Context.ScopeTags = @()
        # Same reason: what a live run cached must not answer an offline one.
        Reset-DocumentationTenantLookups -Context $Context
    }
}

# Drops the tenant lookups a run caches on the singleton context - the
# notification message templates and the two app catalogues - so the next run
# reads current tenant state. They are NoteProperties the lookup helpers in
# TranslatePrimitives.ps1 attach on first use, which is why this checks for each
# before touching it. Called at the start of every bulk run and whenever a run
# flips to offline; the single-object Get-GraphDocumentation path keeps them on
# purpose, so a script documenting policies one call at a time still fetches
# each catalogue once.
function Reset-DocumentationTenantLookups {
    param([DocumentationContext]$Context)

    if (-not $Context) { return }
    foreach ($name in '_NotificationMessageTemplates', '_AllTenantApps', '_AllManagedApps') {
        $property = $Context.PSObject.Properties[$name]
        if ($property) { $property.Value = $null }
    }
}

function Get-DocumentationPoliciesFromSourceFolder {
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [string[]]$PolicyType,
        [string[]]$PolicyGroup
    )

    if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
        throw "Documentation source folder not found: $SourceFolder"
    }

    # Same resolver as the bulk drivers, so Start-GraphBulkDocumentation reports an
    # unknown -PolicyType/-PolicyGroup the same way they do (log + non-terminating
    # error). Note the fallback: with NO selector every type is documented; with a
    # selector that resolves to nothing, NOTHING is. A typo used to widen the run
    # to the whole folder.
    # @( if ... ) and not `$types = if ...`: an empty array written from the block
    # enumerates to nothing and would leave $types as $null - which is exactly
    # how the old code widened an unknown selector to the whole folder.
    $types = @(if ($PolicyType -or $PolicyGroup) {
        (Resolve-IntuneTargetSelectors -PolicyType $PolicyType -PolicyGroup $PolicyGroup -Caller 'Start-GraphBulkDocumentation').Types
    } else {
        $script:IntuneTypes
    })
    # Scan subfolders named by either Id or Folder: most types write to a folder named
    # by their Id, but siblings that share a folder (e.g. the reusable-settings types)
    # export under a common .Folder that differs from their Id. Missing the .Folder name
    # here would silently skip those objects before Get-PoliciesFromFolder can resolve them.
    $subFolders = @($types | ForEach-Object { $_.Id; $_.Folder } | Where-Object { $_ } | Select-Object -Unique)
    $params = @{ Path = $SourceFolder; PolicyTypes = $types; SubFolders = $subFolders }
    @(Get-PoliciesFromFolder @params)
}

function Initialize-DocumentationSourceTenantContext {
    param([Parameter(Mandatory)][string]$SourceFolder)

    # Documenting an export folder: the source tenant may be unreachable, so
    # source-tenant-specific lookups are suppressed and scope-tag names come from
    # the export's ScopeTags/ sidecar instead. Generic schema is still resolved
    # from any connected tenant (see Test-DocumentationGraphAvailable).
    $ctx = Get-DocContextSingleton -Options @{ SourceTenantUnavailable = $true }
    $ctx.ScopeTags = @()
    $scopeFolder = Join-Path $SourceFolder 'ScopeTags'
    if (Test-Path -LiteralPath $scopeFolder -PathType Container) {
        $tags = foreach ($file in Get-ChildItem -LiteralPath $scopeFolder -Filter '*.json' -File) {
            try {
                $item = [IO.File]::ReadAllText($file.FullName) | ConvertFrom-Json
                if ($item.id) { $item }
            } catch {
                Write-Log "Failed to load source-tenant scope tag sidecar '$($file.FullName)'" 2
            }
        }
        if ($tags) { $ctx.ScopeTags = @($tags) }
    }
}

# Backward-compat alias for the former name (pre-rename of OfflineDocumentation).
function Initialize-DocumentationOfflineContext {
    param([Parameter(Mandatory)][string]$SourceFolder)
    Initialize-DocumentationSourceTenantContext -SourceFolder $SourceFolder
}

function Test-DocumentationPolicyFilter {
    param(
        [Parameter(Mandatory)]$PolicyObject,
        [string]$Filter
    )

    if ([string]::IsNullOrWhiteSpace($Filter)) { return $true }
    $filterText = $Filter.Trim()
    if ($filterText -notmatch '^(?i:scope|tag):') {
        return [bool]($PolicyObject.Name -match [regex]::Escape($filterText))
    }

    $scopeFilter = $filterText.Substring($filterText.IndexOf(':') + 1)
    $obj = if ($PolicyObject.PSObject.Properties['JsonObject'] -and $PolicyObject.JsonObject) {
        $PolicyObject.JsonObject
    } else {
        $PolicyObject
    }
    $prop = if ($obj.PSObject.Properties['roleScopeTagIds']) { 'roleScopeTagIds' }
            elseif ($obj.PSObject.Properties['roleScopeTags']) { 'roleScopeTags' }
            else { return $false }
    $ctx = Get-DocContextSingleton
    if ((-not $ctx.ScopeTags -or $ctx.ScopeTags.Count -eq 0) -and -not $ctx.SourceTenantUnavailable -and
        (Test-DocumentationGraphAvailable)) {
        try {
            $response = Invoke-MSGraphAPI -Url '/deviceManagement/roleScopeTags'
            if ($response.Value) { $ctx.ScopeTags = @($response.Value) }
        } catch {
            Write-LogError 'Failed to load scope tags for documentation filter' $_.Exception
        }
    }
    foreach ($id in @($obj.$prop)) {
        $name = if ($id -eq '0') { 'Default' } else { ($ctx.ScopeTags | Where-Object Id -EQ $id | Select-Object -First 1).displayName }
        if ($name -match [regex]::Escape($scopeFilter)) { return $true }
    }
    return $false
}

# ---- Scriptblock NameFilter support (LIST{ } / ITEM{ }) ----
#
# The NameFilter option accepts, in addition to the legacy plain-substring and
# scope:/tag: forms handled by Test-DocumentationPolicyFilter above, two
# scriptblock stages:
#   LIST{ <expr> }  evaluated on list-stage objects, BEFORE hydration
#   ITEM{ <expr> }  evaluated AFTER the surviving items are hydrated
# In both, $_ is the IntunePolicyBase wrapper (.Name, .Platform, .PolicyType,
# .ScopeTags, .ScopeTagNames, .JsonObject).
#
# SECURITY: a filter body is never blindly [scriptblock]::Create'd. It is parsed
# to an AST and rejected unless it is a pure read-only expression — no method
# calls, no assignments, no commands, no file redirections. See
# Assert-DocumentationFilterSafe.

# Node types that mutate state or execute code. Any of these in a filter body
# means the body is rejected. Kept as a script-scope list so it's defined once.
$script:_docFilterForbiddenAstTypes = @(
    [System.Management.Automation.Language.AssignmentStatementAst]    # $_.x = 1
    [System.Management.Automation.Language.InvokeMemberExpressionAst] # $_.Foo(), [Type]::Bar(), .Where({})
    [System.Management.Automation.Language.CommandAst]                # Remove-Item, iex, any cmdlet/function/exe
    [System.Management.Automation.Language.FileRedirectionAst]        # > out.txt
)

# Parse a filter body to an AST, reject it if it contains any forbidden node
# type, and return a validated [scriptblock]. Throws with a clear message on a
# syntax error or a disallowed operation. This is the single choke point that
# makes LIST{ } / ITEM{ } block creation safe.
function Assert-DocumentationFilterSafe {
    param([Parameter(Mandatory)][string]$Body)

    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($Body, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        throw "Documentation filter has a syntax error: $($errors[0].Message)"
    }

    $forbidden = $ast.FindAll({
            param($node)
            foreach ($t in $script:_docFilterForbiddenAstTypes) {
                if ($node -is $t) { return $true }
            }
            return $false
        }, $true)

    if ($forbidden -and $forbidden.Count -gt 0) {
        $first = $forbidden[0]
        throw ("Documentation filter contains a disallowed operation " +
               "($($first.GetType().Name)): '$($first.Extent.Text)'. " +
               "Filters may only read properties and compare values - no method calls, " +
               "assignments, commands, or redirections.")
    }

    return [scriptblock]::Create($Body)
}

# Parse a NameFilter string into its three possible components. Legacy forms
# (plain substring, scope:, tag:) go into .Legacy and are handled by
# Test-DocumentationPolicyFilter unchanged. LIST{ } / ITEM{ } bodies are
# validated (Assert-DocumentationFilterSafe) and returned as scriptblocks.
function ConvertFrom-DocumentationNameFilter {
    param([string]$Filter)

    $result = [PSCustomObject]@{ Legacy = $null; List = $null; Item = $null }
    if ([string]::IsNullOrWhiteSpace($Filter)) { return $result }

    $text = $Filter.Trim()

    # No LIST{ / ITEM{ marker → the whole string is a legacy filter.
    if ($text -notmatch '(?im)^\s*(LIST|ITEM)\s*\{') {
        $result.Legacy = $text
        return $result
    }

    # Walk the string extracting one or more LIST{...} / ITEM{...} segments,
    # balancing braces so a body containing { } parses correctly.
    $i = 0
    while ($i -lt $text.Length) {
        while ($i -lt $text.Length -and [char]::IsWhiteSpace($text[$i])) { $i++ }
        if ($i -ge $text.Length) { break }

        $m = [regex]::Match($text.Substring($i), '^(?i)(LIST|ITEM)\s*\{')
        if (-not $m.Success) {
            throw "Documentation filter: expected 'LIST{' or 'ITEM{' at position $i in '$text'."
        }
        $kind       = $m.Groups[1].Value.ToUpperInvariant()
        $braceStart = $i + $m.Length - 1   # index of the opening '{'

        $depth = 0
        $j = $braceStart
        for (; $j -lt $text.Length; $j++) {
            if ($text[$j] -eq '{') { $depth++ }
            elseif ($text[$j] -eq '}') { $depth--; if ($depth -eq 0) { break } }
        }
        if ($depth -ne 0) {
            throw "Documentation filter has unbalanced '{ }' in the $kind block."
        }

        $body = $text.Substring($braceStart + 1, $j - $braceStart - 1)
        $sb   = Assert-DocumentationFilterSafe -Body $body
        if ($kind -eq 'LIST') { $result.List = $sb } else { $result.Item = $sb }

        $i = $j + 1
    }

    return $result
}

# Evaluate a validated filter scriptblock against one policy object with $_
# bound to it. Fail-closed: a null block passes (no filtering); a block that
# throws excludes the item and logs. Result coerced to bool (last emitted value).
function Test-DocumentationFilterScriptBlock {
    param(
        [Parameter(Mandatory)]$PolicyObject,
        [scriptblock]$ScriptBlock
    )
    if (-not $ScriptBlock) { return $true }
    try {
        $out = $PolicyObject | ForEach-Object -Process $ScriptBlock
        return [bool]($out | Select-Object -Last 1)
    }
    catch {
        Write-LogError "Documentation filter scriptblock failed for '$($PolicyObject.Name)'" $_.Exception
        return $false
    }
}

function Get-DocumentationLanguages {
    $folder = Join-Path $script:AppRootFolder 'Config\LanguageStrings'
    foreach ($file in Get-ChildItem -LiteralPath $folder -Filter 'Strings-*.json' -ErrorAction SilentlyContinue) {
        $name = $file.BaseName.Substring('Strings-'.Length)
        $englishName = try { ([cultureinfo]$name).EnglishName } catch { $name }
        [PSCustomObject]@{ Name = $name; EnglishName = $englishName }
    }
}

function Get-DocumentationUISetting {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$LegacyName,
        $Default
    )

    $fallback = if ($LegacyName) { Get-SettingValue $LegacyName $Default } else { $Default }
    Get-SettingStoreValue 'Documentation' $Name $fallback
}

function Save-DocumentationOptionsDefaults {
    param(
        [hashtable]$Options,
        [switch]$CommonOnly
    )
    if (-not $Options) { return }
    foreach ($key in $Options.Keys) {
        if ($key -eq 'Outputs') { continue }
        $value = $Options[$key]
        # Skip $null: a missing UI control (e.g. an output XAML that doesn't
        # carry the field yet) returns $null from the form-options collector,
        # and writing that back would erase the user's actually-persisted
        # setting on every click-Export.
        if ($null -eq $value) { continue }
        # Generic keys -> Documentation; per-provider keys -> Documentation\<provider>.
        # Same routing as the loaders (Get-DocumentationSetting) so save/load stay symmetric.
        Save-SettingStoreValue (Get-DocumentationSettingSubPath $key) $key $value
    }
    if (-not $CommonOnly -and $Options.Outputs) {
        foreach ($provider in $Options.Outputs.Keys) {
            foreach ($key in $Options.Outputs[$provider].Keys) {
                $value = $Options.Outputs[$provider][$key]
                if ($null -eq $value) { continue }
                Save-SettingStoreValue (Get-DocumentationSettingSubPath $key) $key $value
            }
        }
    }
}

function Save-DocumentationFile {
    param(
        [string]$Content,
        [string]$FileName,
        [switch]$OpenFile
    )

    try {
        # Out-File doesn't auto-create the parent directory — when an output
        # provider's path template lands in a fresh date-stamped folder
        # (e.g. D:\Intune\Documentation\Test\YYYYMMDD\...html) the write blows
        # up with "Could not find a part of the path". Create the directory
        # first; idempotent if it already exists.
        $dir = [IO.Path]::GetDirectoryName($FileName)
        if ($dir -and -not [IO.Directory]::Exists($dir)) {
            [void][IO.Directory]::CreateDirectory($dir)
        }
        $Content | Out-File -LiteralPath $FileName -Force -Encoding utf8 -ErrorAction Stop
        Write-Log "$FileName saved successfully"
        if ($OpenFile) {
            # Invoke-Item opens the default app on Windows but does nothing useful
            # for an .html file on Linux/macOS. Open-ExternalUri uses ShellExecute
            # (-> xdg-open / open / default browser) on every host.
            Open-ExternalUri $FileName
        }
    }
    catch {
        Write-LogError "Failed to save file $FileName." $_.Exception
        throw
    }
}

# Shared $script:columnHeaders map — property-name → language-id mappings used by
# Invoke-DocTranslateColumnHeader to localize table column headers. Seeded with the
# old engine's table (Documentation.psm1:33). Lookup is case-insensitive, and
# Invoke-DocTranslateColumnHeader keys on the last '.'-segment, so 'Settings.<key>'
# columns resolve by <key>. Output providers/handlers can register more via
# Set-DocColumnHeaderLanguageId.
$script:columnHeaders = @{
    Name                         = 'Inputs.displayNameLabel'
    Value                        = 'TableHeaders.value'
    Description                  = 'TableHeaders.description'
    GroupMode                    = 'SettingDetails.modeTableHeader'
    Group                        = 'TableHeaders.assignedGroups'
    Groups                       = 'TableHeaders.groups'
    useDeviceContext             = 'SettingDetails.installContextLabel'
    uninstallOnDeviceRemoval     = 'SettingDetails.UninstallOnRemoval'
    isRemovable                  = 'SettingDetails.installAsRemovable'
    preventManagedAppBackup      = 'AppResources.AppSettingsUx.preventManagedAppBackup'
    preventAutoAppUpdate         = 'AppResources.AppSettingsUx.preventAutoAppUpdate'
    vpnConfigurationId           = 'PolicyType.vpn'
    Action                       = 'SettingDetails.actionColumnName'
    Schedule                     = 'ScheduledAction.List.schedule'
    MessageTemplate              = 'ScheduledAction.Notification.messageTemplate'
    EmailCC                      = 'ScheduledAction.Notification.additionalRecipients'
    Rule                         = 'ApplicabilityRules.GridLabel.Rule'
    ValueWithLabel               = 'TableHeaders.value'
    Status                       = 'TableHeaders.status'
    CombinedValueWithLabel       = 'TableHeaders.value'
    CombinedValue                = 'TableHeaders.value'
    useDeviceLicensing           = 'TableHeaders.licenseType'
    Filter                       = 'AppResources.AppSettingsUx.assignmentFilterColumnHeader'
    filterMode                   = 'AppResources.AppSettingsUx.assignmentFilterTypeColumnHeader'
    deliveryOptimizationPriority = 'AppResources.AppSettingsUx.deliveryOptimizationPriorityHeader'
    startTimeColumnLabel         = 'AppResources.AppSettingsUx.startTimeColumnLabel'
    installTimeSettings          = 'AppResources.AppSettingsUx.deadlineTimeColumnLabel'
    restartSettings              = 'AppResources.AppSettingsUx.restartGracePeriodHeader'
    notifications                = 'AppResources.AppSettingsUx.assignmentToast'
    Settings                     = 'TableHeaders.settings'
    returnCode                   = 'Win32ReturnCodes.Columns.returnCode'
    type                         = 'Win32ReturnCodes.Columns.codeType'
    RecommendedValue             = 'AzureIAMCommon.Recommended'
    ConfigurationKey             = 'SettingDetails.configurationKey'
    ValueType                    = 'SettingDetails.valueType'
    ConfigurationValue           = 'SettingDetails.configurationValue'
}

function Set-DocColumnHeaderLanguageId {
    param([string]$Property, [string]$LanguageId)
    $script:columnHeaders[$Property] = $LanguageId
}

function Invoke-DocTranslateColumnHeader {
    param([string]$ColumnName)

    if ($script:columnHeaders.ContainsKey($ColumnName)) {
        $lng = Get-LanguageString $script:columnHeaders[$ColumnName]
        if ($lng) { return $lng }
    }
    return $ColumnName
}

# Object-type group/category label. Phase 2 will populate the full GroupId →
# language-id mapping from the old engine. For now returns the input string
# as-is so headers render with the raw type/group id.
function Get-DocObjectTypeString {
    param($ObjectTypeOrGroupId)

    if ($null -eq $ObjectTypeOrGroupId) { return "" }
    if ($ObjectTypeOrGroupId -is [string]) { return $ObjectTypeOrGroupId }
    if ($ObjectTypeOrGroupId.PSObject.Properties.Name -contains 'Title') {
        return $ObjectTypeOrGroupId.Title
    }
    return "$ObjectTypeOrGroupId"
}

# Resolves the output folder for an object during per-file export (used by CSV).
# Optionally nests under organization name and/or policy-type id.
function Get-DocObjectFolder {
    param(
        [string]$RootFolder,
        $PolicyType,
        [switch]$AddOrganization,
        [switch]$AddObjectType
    )

    $path = $RootFolder
    $orgName = Get-CurrentOrganizationName
    if ($AddOrganization -and $orgName) {
        $path = Join-Path $path $orgName
    }
    if ($AddObjectType -and $PolicyType -and $PolicyType.Id) {
        $path = Join-Path $path $PolicyType.Id
    }
    $path
}

# Documents one PolicyObject. Looks up a registered *DocHandler first (by @odata.type);
# falls back to the schema-driven input-provider chain when no handler claims it.
# Returns the per-object result PSCustomObject (see [DocumentationContext]::ToResult).
#
# Phase 2 only wires the dispatch shape — handler registry and input providers
# are populated in phases 4 and 3 respectively. Until then this returns an empty
# result with InputType='NoProvider' so downstream code can detect uninstrumented
# types without throwing.
function Invoke-DocumentationForObject {
    param(
        [Parameter(Mandatory)] $PolicyObject,
        [DocumentationContext]$Context
    )

    if (-not $Context) {
        $Context = [DocumentationContext]::new($PolicyObject, 'en', $null)
    }
    else {
        $Context.ResetForObject($PolicyObject)
    }

    # Translate primitives (Invoke-TranslateBoolean / Option / ... and Add-PropertyInfo)
    # read state from a module-level pointer; engine sets it once per object so
    # input providers and handlers can use the old function signatures unchanged.
    Set-CurrentDocumentationContext $Context

    $odataType = $null
    if ($PolicyObject.PSObject.Properties['JsonObject'] -and $PolicyObject.JsonObject) {
        $odataType = $PolicyObject.JsonObject.'@odata.type'
    }
    if (-not $odataType -and $PolicyObject.PSObject.Properties['@odata.type']) {
        $odataType = $PolicyObject.'@odata.type'
    }

    $handler = if ($odataType) { [DocumentationRegistry]::FindHandler($odataType) } else { $null }
    if ($handler) {
        try {
            $handler.Document($PolicyObject, $Context)
            $Context.InputType = "Handler:$($handler.GetType().Name)"
        }
        catch {
            $Context.ErrorText = "Handler $($handler.GetType().Name) failed: $($_.Exception.Message)"
            Write-LogError "Documentation handler $($handler.GetType().Name) failed for $($PolicyObject.Name)" $_.Exception
        }
        Add-ScopeTagsBasicInfoIfApplicable $PolicyObject $Context
        Add-AssignmentsForObjectIfApplicable $PolicyObject $Context
        Add-ComplianceActionsForObjectIfApplicable $PolicyObject $Context
        Add-PolicyIdBasicInfoIfRequested $PolicyObject $Context
        return $Context.ToResult()
    }

    # Providers are evaluated in Order; the first whose Match claims the object
    # wins and the rest are skipped. The generic fallback registers last
    # (Order = MaxValue) and claims the object only when Options.FallbackDocumentation
    # is enabled, so an unsupported type with the option off falls through to the
    # NoProvider stub below.
    $inputProvider = [DocumentationRegistry]::FindInputProvider($PolicyObject, $Context)
    if ($inputProvider) {
        try {
            & $inputProvider.Translate $PolicyObject $Context
            $Context.InputType = "Input:$($inputProvider.Name)"
        }
        catch {
            $Context.ErrorText = "Input provider $($inputProvider.Name) failed: $($_.Exception.Message)"
            Write-LogError "Documentation input provider $($inputProvider.Name) failed for $($PolicyObject.Name)" $_.Exception
        }
        Finalize-DocumentationObjectInfoObject $Context.CurrentObject
        Add-ScopeTagsBasicInfoIfApplicable $PolicyObject $Context
        Add-AssignmentsForObjectIfApplicable $PolicyObject $Context
        Add-ComplianceActionsForObjectIfApplicable $PolicyObject $Context
        Add-PolicyIdBasicInfoIfRequested $PolicyObject $Context
        return $Context.ToResult()
    }

    # Nothing claimed the object and the generic fallback is switched off. Say so
    # once, clearly, and emit NOTHING - not even the policy-id row, which on its
    # own produced a document containing a single id and no other content. The
    # caller skips a NoProvider result rather than handing it to the output
    # providers, so the type is absent from the documentation instead of present
    # and blank.
    $Context.InputType = 'NoProvider'
    $typeTitle = if ($PolicyObject.PSObject.Properties['PolicyType'] -and $PolicyObject.PolicyType) { [string]$PolicyObject.PolicyType.Title } else { $null }
    if (-not $typeTitle) { $typeTitle = [string]$odataType }
    Write-Log "Documentation: '$($PolicyObject.Name)' was skipped - the type $typeTitle has no documentation support. Enable 'Document unsupported types' in Output Settings to document it from its raw properties." 2
    return $Context.ToResult()
}

# deviceComplianceActionType members whose Graph name differs from the portal's
# language id. Verified against the full enum 2026-08-30 - these two are the only
# ones; everything else resolves as ScheduledAction.<member>.
$script:_complianceActionLanguageIds = @{
    'notification'                 = 'notificationLabel'
    'removeResourceAccessProfiles' = 'removeSourceAccessProfile'
}

# Engine post-step: translate an object's scheduledActionsForRule into ComplianceActions
# rows (the "Actions for noncompliance" table shown for compliance policies, V1 and V2).
# Run centrally like the other post-steps, matching the old dispatcher tail
# (Documentation.psm1:347 -> Invoke-TranslateScheduledActionType, :3518). No-op for
# objects that don't carry scheduledActionsForRule.
function Add-ComplianceActionsForObjectIfApplicable {
    param([object]$PolicyObject, [DocumentationContext]$Context)

    $obj = if ($PolicyObject.PSObject.Properties['JsonObject'] -and $PolicyObject.JsonObject) { $PolicyObject.JsonObject } else { $PolicyObject }
    $rules = @($obj.scheduledActionsForRule)
    if ($rules.Count -eq 0) { return }

    foreach ($actionRule in $rules) {
        foreach ($actionConfig in @($actionRule.scheduledActionConfigurations)) {
            if (-not $actionConfig) { continue }

            # Action-type strings live directly under ScheduledAction, EXCEPT where
            # the Graph enum member and the portal's language id disagree (checked
            # against every deviceComplianceActionType member 2026-08-30):
            #
            #   notification    ScheduledAction.Notification is an OBJECT of
            #                   email-picker sub-strings and the lookup is
            #                   case-insensitive, so the direct key hit the
            #                   container and threw. The label leaf is
            #                   notificationLabel ("Send email to end user").
            #                   This is the collision the old engine dodged by
            #                   renaming to 'emailNotification' - the comment that
            #                   used to sit here claimed the sub-object no longer
            #                   existed, which was wrong.
            #   removeResourceAccessProfiles
            #                   the portal ships this as removeSourceAccessProfile
            #                   ("Remove source access profile"); without the alias
            #                   the raw enum member rendered.
            #
            # Every other member (noAction/block/retire/wipe/pushNotification/
            # remoteLock) resolves directly.
            $actionKey = [string]$actionConfig.actionType
            if ($script:_complianceActionLanguageIds.ContainsKey($actionKey)) {
                $actionKey = $script:_complianceActionLanguageIds[$actionKey]
            }

            $actionType = Get-LanguageString "ScheduledAction.$actionKey" -IgnoreMissing
            if ([string]::IsNullOrEmpty($actionType)) { $actionType = [string]$actionConfig.actionType }

            $schedule = if ($actionConfig.gracePeriodHours -eq 0) {
                Get-LanguageString 'ScheduledAction.List.immediately'
            }
            else {
                # gracePeriodHours is always stored in hours but displayed in days.
                (Get-LanguageString 'ScheduledAction.List.gracePeriodDays') -f ($actionConfig.gracePeriodHours / 24)
            }

            $notificationTemplate    = $null
            $additionalNotifications = $null
            if ($actionConfig.actionType -eq 'notification') {
                $notificationTemplate = if ($actionConfig.notificationTemplateId -ne [Guid]::Empty) {
                    Get-LanguageString 'ScheduledAction.Notification.selected'
                } else {
                    Get-LanguageString 'ScheduledAction.Notification.noneSelected'
                }
                $additionalNotifications = if (@($actionConfig.notificationMessageCCList).Count -gt 0) {
                    (Get-LanguageString 'ScheduledAction.Notification.numSelected') -f @($actionConfig.notificationMessageCCList).Count
                } else {
                    Get-LanguageString 'ScheduledAction.Notification.noneSelected'
                }
            }

            $Context.AddComplianceAction($actionType, $schedule, $notificationTemplate, $additionalNotifications)
        }
    }
}

# Engine post-step: append a Scope tags row to BasicInfo when the raw object
# has roleScopeTagIds or roleScopeTags. Doing this in the engine (rather than
# each handler/provider) means every documented type picks it up for free,
# matching what the old engine did from its central dispatcher tail (old code
# Documentation.psm1:344). No-op when basicInfo is empty (handler/provider
# didn't run successfully).
function Add-ScopeTagsBasicInfoIfApplicable {
    param([object]$PolicyObject, [DocumentationContext]$Context)

    if ($Context.BasicInfo.Count -eq 0) { return }

    $obj = if ($PolicyObject.PSObject.Properties['JsonObject'] -and $PolicyObject.JsonObject) {
        $PolicyObject.JsonObject
    } else {
        $PolicyObject
    }
    Add-ScopeTagStrings $obj
}

# Engine post-step: translate $obj.assignments into Assignment rows on the
# context (Generic + App). No-op if BasicInfo is empty (handler/provider failed)
# or if Context.Options.ExcludeAssignments is true. Delegates the actual
# translation to Add-AssignmentsForObject in TranslatePrimitives.ps1.
function Add-AssignmentsForObjectIfApplicable {
    param([object]$PolicyObject, [DocumentationContext]$Context)

    if ($Context.BasicInfo.Count -eq 0) { return }

    $obj = if ($PolicyObject.PSObject.Properties['JsonObject'] -and $PolicyObject.JsonObject) {
        $PolicyObject.JsonObject
    } else {
        $PolicyObject
    }
    Add-AssignmentsForObject $obj
}

# Engine post-step: when Context.Options.IncludePolicyId is set, append the
# raw object's Id to BasicInfo. Doing this in the engine (rather than each
# handler/provider) means every documented type picks it up for free, and
# downstream tools can rely on the row being present whenever the option is on.
function Add-PolicyIdBasicInfoIfRequested {
    param([object]$PolicyObject, [DocumentationContext]$Context)

    if (-not $Context.Options.IncludePolicyId) { return }

    $policyId = $null
    if ($PolicyObject.PSObject.Properties['JsonObject'] -and $PolicyObject.JsonObject -and $PolicyObject.JsonObject.id) {
        $policyId = $PolicyObject.JsonObject.id
    }
    elseif ($PolicyObject.PSObject.Properties['Id']) {
        $policyId = $PolicyObject.Id
    }
    if (-not $policyId) { return }

    # Label "Policy ID" — the existing SettingDetails.policyId language string
    # already resolves to "Policy ID" in the en strings file; falls back to a
    # literal if the key is ever missing.
    $label = Get-LanguageString 'SettingDetails.policyId'
    if (-not $label) { $label = 'Policy ID' }

    $idRow = [PSCustomObject]@{ Name = $label; Value = $policyId; EntityKey = 'id' }

    # Position: right after the conventional Name row (EntityKey='displayName').
    # Fall back to position 1 (just after the first row) if no displayName row
    # exists, or append if BasicInfo is empty.
    $insertAt = -1
    for ($i = 0; $i -lt $Context.BasicInfo.Count; $i++) {
        if ($Context.BasicInfo[$i].EntityKey -eq 'displayName') {
            $insertAt = $i + 1
            break
        }
    }
    if ($insertAt -ge 0) {
        $Context.BasicInfo.Insert($insertAt, $idRow)
    }
    elseif ($Context.BasicInfo.Count -gt 0) {
        $Context.BasicInfo.Insert(1, $idRow)
    }
    else {
        $Context.BasicInfo.Add($idRow)
    }
}

# Orchestrates one batch through one or more registered output providers.
# Drives the lifecycle hooks (Activate -> PreProcess -> {NewObjectGroup -> NewObjectType -> Process}* -> ProcessAllObjects -> PostProcess).
#
# $OutputValue is the registry Value field (e.g. 'json'); accepts comma-separated
# list for multi-output batches. $PolicyObjects is the iteration source.
function Get-DocumentationPolicyGroupId {
    param([object]$PolicyObject)

    $policyType = $PolicyObject.PolicyType
    if ($policyType -and $policyType.PolicyGroup -and $policyType.PolicyGroup.Id) {
        return [string]$policyType.PolicyGroup.Id
    }
    if ($policyType -and $policyType.PSObject.Properties['GroupId'] -and $policyType.GroupId) {
        return [string]$policyType.GroupId
    }
    return 'Other'
}

function Get-DocumentationPolicyGroupName {
    param([object]$PolicyObject)

    $policyType = $PolicyObject.PolicyType
    if ($policyType -and $policyType.PolicyGroup -and $policyType.PolicyGroup.Title) {
        return [string]$policyType.PolicyGroup.Title
    }
    return Get-DocObjectTypeString (Get-DocumentationPolicyGroupId $PolicyObject)
}

function Get-DocumentationPolicyTypeName {
    param([object]$PolicyObject)

    # Several policy types can share one second-level heading - see
    # Core/DocumentationEnrollmentGrouping.ps1. Nothing else is affected: the
    # resolver returns $null for every type that is not in a grouping.
    $group = Get-DocumentationTypeGroup $PolicyObject
    if ($group) { return [string]$group.Title }

    if ($PolicyObject.PolicyType -and $PolicyObject.PolicyType.Title) {
        return [string]$PolicyObject.PolicyType.Title
    }
    if ($PolicyObject.PolicyType -and $PolicyObject.PolicyType.Id) {
        return [string]$PolicyObject.PolicyType.Id
    }
    return 'Other'
}

# The key the second level is grouped on. Must stay in step with
# Get-DocumentationPolicyTypeName above: grouping on PolicyType.Id while naming from
# a shared grouping title would emit the same heading once per member type.
function Get-DocumentationPolicyTypeId {
    param([object]$PolicyObject)

    $group = Get-DocumentationTypeGroup $PolicyObject
    if ($group) { return [string]$group.Id }

    if ($PolicyObject.PolicyType -and $PolicyObject.PolicyType.Id) {
        return [string]$PolicyObject.PolicyType.Id
    }
    return 'Other'
}

# Run-level prefetch: resolve every tenant-wide lookup and per-policy
# sub-resource the providers would otherwise fetch one-at-a-time mid-run.
#
#   1. Seed ScopeTags + FilterNamesById from the login-time tenant dependency
#      cache (Initialize-TenantDependencyCache preloads both at authentication;
#      previously the doc pipeline ignored that cache and re-fetched).
#   2. Everything still missing goes into ONE Invoke-GraphBatchRequest call:
#      - deviceManagement/roleScopeTags                  (if ScopeTags empty)
#      - deviceManagement/assignmentFilters?$select=...  (if filters not loaded)
#      - deviceManagement/configurationCategories        (if run has Settings Catalog policies)
#      - deviceManagement/complianceCategories           (if run has Compliance V2 policies)
#      - configurationPolicies('<id>')/settings?$expand=settingDefinitions  per SC policy
#      - compliancePolicies('<id>')/settings?$expand=settingDefinitions     per CompV2 policy
#      The batcher chunks at 20 sub-requests and parallelizes chunks, so 91
#      Settings Catalog policies cost ~5 round-trips instead of 91 sequential GETs.
#
# Providers keep their per-policy lazy fallbacks for the single-policy
# Get-GraphDocumentation path; this prefetch just makes the bulk path batch.
function Initialize-DocumentationRunPrefetch {
    param(
        [object[]]$PolicyObjects,
        [DocumentationContext]$Context
    )

    if (-not $Context -or $Context.SourceTenantUnavailable) { return }
    if (-not (Get-Command Invoke-GraphBatchRequest -ErrorAction SilentlyContinue)) { return }

    # Per-run freshness: a policy re-documented in the same session must show
    # its current settings, so the per-policy prefetch never spans runs.
    $Context.PrefetchedPolicySettings = @{}

    $tokenId = 0
    foreach ($p in @($PolicyObjects)) {
        if ($p -and $null -ne $p._TokenId) { $tokenId = [int]$p._TokenId; break }
    }

    # ---- 1. Seed from the login-time tenant dependency cache (no Graph) ----
    try {
        # Get-OperationTokenInfo, not Get-TokenInfo: $tokenId is seeded 0 above and a
        # policy carrying no token has _TokenId 0 - never $null - so 0 survives here.
        # Get-TokenInfo reads that 0 as "no filter, every token", which would name
        # both tenants in the DependencyObjects_<tenant> key with a second one live.
        $tokenInfo = Get-OperationTokenInfo $tokenId
        if ($tokenInfo -and $tokenInfo.TenantId) {
            $dep = Get-CacheObject "DependencyObjects_$($tokenInfo.TenantId)"
            if ($dep -is [hashtable]) {
                if ((-not $Context.ScopeTags -or $Context.ScopeTags.Count -eq 0) -and $dep.ContainsKey('ScopeTags')) {
                    # Project the IntunePolicyBase wrappers (and the synthetic
                    # Default entry) to the raw {id, displayName} shape
                    # Add-ScopeTagStrings / Test-DocumentationPolicyFilter read.
                    $Context.ScopeTags = @($dep['ScopeTags'] | Where-Object { $_ } | ForEach-Object {
                        [PSCustomObject]@{ id = [string]$_.Id; displayName = [string]$_.Name }
                    })
                    Write-Log "Documentation prefetch: seeded $($Context.ScopeTags.Count) scope tag(s) from tenant dependency cache"
                }
                if (-not $Context.FiltersLoaded) {
                    $filters = @($dep.Values | Where-Object {
                        $_ -and $_.PSObject.Properties['PolicyType'] -and $_.PolicyType.Id -eq 'AssignmentFilters'
                    })
                    if ($filters.Count -gt 0) {
                        foreach ($f in $filters) { $Context.FilterNamesById[[string]$f.Id] = [string]$f.Name }
                        $Context.FiltersLoaded = $true
                        Write-Log "Documentation prefetch: seeded $($filters.Count) assignment filter(s) from tenant dependency cache"
                    }
                }
            }
        }
    }
    catch {
        Write-LogDebug "Documentation prefetch: dependency-cache seed failed: $($_.Exception.Message)"
    }

    # ---- 1b. Batch-resolve every assignment group name in one getByIds call ----
    # Front-loads $ctx.GroupNamesById so Add-AssignmentsForObject is a pure cache hit
    # for every policy instead of one getByIds per policy. Same groups either way
    # (only assignment-referenced ones) and same output - purely fewer round-trips.
    # Not gated on the batching setting: this is one getByIds POST per thousand
    # ids, not a $batch, and skipping it only made a run resolve one group per
    # policy instead. UseBatchAPI now means exactly "combine requests into $batch"
    # (see Test-GraphBatchEnabled) and this path never did that.
    if (-not $Context.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
        try {
            Sync-DocumentationGroupPreload -PolicyObjects $PolicyObjects -Context $Context
        }
        catch {
            Write-LogError 'Documentation prefetch: batch group resolution failed - falling back to per-policy resolution' $_.Exception
        }
    }

    # ---- 2. One $batch for everything still missing ----
    $minimalHeaders = @{ Accept = 'application/json;odata.metadata=minimal' }
    $settingsHeaders = @{ Accept = 'application/json;odata.metadata=minimal' }
    if ($Context.Language -and $Context.Language -ne 'en') {
        $settingsHeaders['Accept-Language'] = $Context.Language
    }

    $requests = [System.Collections.Generic.List[PSCustomObject]]::new()
    $routes   = @{}   # reqId -> [PSCustomObject]@{ Kind; PolicyId }

    if (-not $Context.ScopeTags -or $Context.ScopeTags.Count -eq 0) {
        [void]$requests.Add([PSCustomObject]@{ id = 'scopeTags'; method = 'GET'; url = 'deviceManagement/roleScopeTags'; headers = $minimalHeaders })
        $routes['scopeTags'] = [PSCustomObject]@{ Kind = 'scopeTags' }
    }
    if (-not $Context.FiltersLoaded) {
        [void]$requests.Add([PSCustomObject]@{ id = 'filters'; method = 'GET'; url = 'deviceManagement/assignmentFilters?$select=id,displayName&$top=999'; headers = $minimalHeaders })
        $routes['filters'] = [PSCustomObject]@{ Kind = 'filters' }
    }

    # Category lists, only when the run actually contains policies that use them.
    $Context.CfgCategories = Get-CacheObject "CfgCategories" (@())
    $hasSettingsCatalog = $false
    $hasComplianceV2    = $false
    $settingsIdx = 0
    foreach ($p in @($PolicyObjects)) {
        if (-not $p -or $p.IsFromFile -eq $true) { continue }
        $obj = $p.JsonObject
        if (-not $obj -or -not $p.Id) { continue }

        $odata = [string]$obj.'@odata.type'
        $needsSettings = $false
        $settingsUrl = $null

        if ($odata -eq '#microsoft.graph.deviceManagementConfigurationPolicy') {
            $hasSettingsCatalog = $true
            $settingsUrl = "deviceManagement/configurationPolicies('$($p.Id)')/settings?`$expand=settingDefinitions&`$top=1000"
        }
        elseif ($odata -eq '#microsoft.graph.deviceManagementCompliancePolicy') {
            $hasComplianceV2 = $true
            $settingsUrl = "deviceManagement/compliancePolicies('$($p.Id)')/settings?`$expand=settingDefinitions&`$top=1000"
        }
        else { continue }

        # Skip when the body already carries settings WITH inline definitions —
        # the providers use those directly. Hydrated bodies typically have
        # settings without definitions, so they still need the enrich fetch.
        $needsSettings = $true
        foreach ($s in @($obj.Settings)) {
            if ($s.settingDefinitions -and @($s.settingDefinitions).Count -gt 0) { $needsSettings = $false; break }
        }
        if (-not $needsSettings) { continue }

        $settingsIdx++
        $reqId = "polset_$settingsIdx"
        [void]$requests.Add([PSCustomObject]@{ id = $reqId; method = 'GET'; url = $settingsUrl; headers = $settingsHeaders })
        $routes[$reqId] = [PSCustomObject]@{ Kind = 'policySettings'; PolicyId = [string]$p.Id }
    }

    if ($hasSettingsCatalog -and -not ($Context.CfgCategories | Where-Object { $_.settingUsage -eq 'configuration' })) {
        [void]$requests.Add([PSCustomObject]@{ id = 'cfgCats'; method = 'GET'; url = 'deviceManagement/configurationCategories'; headers = $minimalHeaders })
        $routes['cfgCats'] = [PSCustomObject]@{ Kind = 'cfgCats' }
    }
    if ($hasComplianceV2 -and -not ($Context.CfgCategories | Where-Object { $_.settingUsage -eq 'compliance' })) {
        [void]$requests.Add([PSCustomObject]@{ id = 'compCats'; method = 'GET'; url = 'deviceManagement/complianceCategories'; headers = $minimalHeaders })
        $routes['compCats'] = [PSCustomObject]@{ Kind = 'compCats' }
    }

    if ($requests.Count -eq 0) { return }

    Write-Log "Documentation prefetch: $($requests.Count) request(s) in one batch (tenant-wide + per-policy settings)"
    $results = @(Invoke-GraphBatchRequest -BatchObjects $requests -BatchType 'Documentation:Prefetch' -TokenId $tokenId -SkipWarnings -IncludedFailed)

    $categoriesUpdated = $false
    foreach ($result in $results) {
        $route = $routes["$($result.Id)"]
        if (-not $route) { continue }
        if ($result.Status -lt 200 -or $result.Status -ge 300 -or -not $result.body) {
            Write-Log "Documentation prefetch: '$($result.Id)' returned HTTP $($result.Status) - provider lazy fallback will retry" 2
            continue
        }

        switch ($route.Kind) {
            'scopeTags' {
                if ($result.body.value) { $Context.ScopeTags = @($result.body.value) }
            }
            'filters' {
                foreach ($f in @($result.body.value)) {
                    if ($f.id) { $Context.FilterNamesById[[string]$f.id] = $f.displayName }
                }
                $Context.FiltersLoaded = $true
            }
            'cfgCats' {
                $Context.CfgCategories += @($result.body.value)
                $categoriesUpdated = $true
            }
            'compCats' {
                $Context.CfgCategories += @($result.body.value)
                $categoriesUpdated = $true
            }
            'policySettings' {
                $Context.PrefetchedPolicySettings[$route.PolicyId] = @($result.body.value)
            }
        }
    }

    if ($categoriesUpdated) {
        Set-CacheObject "CfgCategories" $Context.CfgCategories -Persistent
    }
}

# Batch group-name resolution (run from the prefetch when the "Use Batch API"
# setting is on). Collects every assignment groupId across all policies in the run
# and resolves them to display names in one getByIds batch (chunked at 1000),
# caching into $Context.GroupNamesById.
# Best-effort: assignments must already be hydrated onto the policy bodies (the
# run hydrate step in Invoke-DocumentationOutputs runs first); anything not found
# here still resolves lazily per-object in Add-AssignmentsForObject.
function Sync-DocumentationGroupPreload {
    param(
        [object[]]$PolicyObjects,
        [DocumentationContext]$Context
    )
    if (-not $Context) { return }

    $ids = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($p in @($PolicyObjects)) {
        if (-not $p) { continue }
        $obj = if ($p.PSObject.Properties['JsonObject']) { $p.JsonObject } else { $p }
        if (-not $obj) { continue }
        foreach ($a in @($obj.assignments)) {
            if (-not $a -or -not $a.target) { continue }
            $gid = [string]$a.target.groupId
            if ($gid) { [void]$ids.Add($gid) }
        }
    }
    if ($ids.Count -eq 0) { return }

    Write-Status "Documentation - preloading assignment groups" -SkipLog -Force
    Write-Log "Documentation prefetch: preloading $($ids.Count) assignment group name(s) in one batch"
    [void](Resolve-DocumentationGroupNames -GroupIds @($ids) -Context $Context)
}

# What an output provider should call this object in a heading or a caption.
#
# Normally the policy's display name, but a documented object may ask to be titled
# something else through DocumentName - the default enrollment policies do, because
# five different types all ship under the name "All users and all devices" and the
# document would not say which one it is. The display name is not lost: it still
# appears as the Name row inside the table.
#
# Headings and captions only. The file name, the CSV rows and the JSON object keep
# the real display name, because those identify the object rather than describe it.
function Get-DocumentationDisplayName {
    param($PolicyObject, $DocumentedObject)

    if ($DocumentedObject -and
        $DocumentedObject.PSObject.Properties['DocumentName'] -and
        -not [string]::IsNullOrWhiteSpace([string]$DocumentedObject.DocumentName)) {
        return [string]$DocumentedObject.DocumentName
    }

    # A grouped enrollment default is titled by its policy type even when nothing set
    # DocumentName. These five types reach the document by different routes - a
    # dedicated Handler for the platform restrictions, the ObjectInfo manifest for
    # Windows Hello and Windows Restore, the generic fallback for the rest - and a
    # Handler wins over the input provider that would otherwise have set the name, so
    # setting it there alone left the platform restrictions headed "All users and all
    # devices". Resolving it here covers every route once instead of each handler
    # having to remember. See Core/DocumentationEnrollmentGrouping.ps1.
    $enrollmentName = Get-DocumentationEnrollmentDefaultName $PolicyObject
    if ($enrollmentName) { return $enrollmentName }

    return [string]$PolicyObject.Name
}

# The name the object currently being documented is titled under, for table
# captions. The per-object Process function of each output provider stamps
# $script:docDisplayName from Get-DocumentationDisplayName above; this reads it
# back where only $PolicyObject is in scope. Falls back to the display name so a
# caption built outside that path still reads correctly.
#
# One copy, in the engine: the four document providers share a module scope, so
# four identical definitions would silently collapse into whichever file loaded
# last.
function Get-DocCaptionName {
    param($PolicyObject)

    if (-not [string]::IsNullOrWhiteSpace($script:docDisplayName)) { return $script:docDisplayName }
    return [string]$PolicyObject.Name
}

function Invoke-DocumentationOutputs {
    param(
        [string]$OutputValue,
        [object[]]$PolicyObjects,
        [hashtable]$Options,
        [string]$Language = 'en'
    )

    if (-not $OutputValue) { throw "OutputValue is required" }
    $Options = Merge-DocumentationOptions $Options
    if ($Language) { $Options.Language = $Language }
    $script:_docRunOptions = $Options
    Set-DocumentationContextRunOptions -Context (Get-DocContextSingleton) -Options $Options -Language $Language
    # A run reads current tenant state: whatever the previous run cached of the
    # message templates and app catalogues is dropped here, not carried across.
    Reset-DocumentationTenantLookups -Context (Get-DocContextSingleton)

    # Hydrate every input policy through the unified orchestrator BEFORE the
    # output loop. Handlers and input providers read $PolicyObject.JsonObject.*
    # and assume a fully hydrated body + sub-resources (e.g. SettingsCatalog
    # assignments, AdminTemplate definition/presentation values, branding
    # images, RoleDefinition assignment expansion). Without this, online doc
    # runs that source policies from Get-GraphPolicies silently emit incomplete
    # output for any type whose list endpoint doesn't include sub-resources.
    #
    # Skip when the source tenant is unavailable (SourceFolder docs) and per-policy
    # for file-loaded objects — they have no token and Invoke-PolicyHydrate
    # would issue Graph calls under the default token, hitting either an auth
    # error or (worse) the wrong tenant. (Hydrate is source-tenant-specific:
    # by-id GETs that 404 on any other tenant — distinct from generic schema.)
    if (-not $Options.SourceTenantUnavailable) {
        $hydrateTargets = @($PolicyObjects | Where-Object {
            $_ -and $_.PSObject.Properties['_IsFullObject'] -and -not $_._IsFullObject -and
            $_.Id -and $_.PolicyType -and $_.IsFromFile -ne $true
        })
        if ($hydrateTargets.Count -gt 0) {
            Write-Status "Documentation - hydrating policies" -SkipLog -Force
            Invoke-PolicyHydrate -Policies $hydrateTargets
        }

        # Batch-prefetch tenant-wide lookups + per-policy settings so the
        # output loop below never pays a sequential mid-run Graph GET. Failure
        # is non-fatal: providers keep their lazy per-policy fallbacks.
        try {
            Write-Status "Documentation - pre-fetching settings" -SkipLog -Force
            Initialize-DocumentationRunPrefetch -PolicyObjects $PolicyObjects -Context (Get-DocContextSingleton)
        }
        catch {
            Write-LogError 'Documentation prefetch failed - falling back to per-policy fetches' $_.Exception
        }
    }

    $failures = [System.Collections.Generic.List[object]]::new()
    $recordFailure = {
        param([string]$Stage, $Output, [string]$PolicyName, [string]$Message, $Exception)
        Write-LogError $Message $Exception
        [void]$failures.Add([PSCustomObject]@{
            Stage      = $Stage
            Output     = if ($Output) { [string]$Output.Value } else { "" }
            PolicyName = $PolicyName
            Message    = $Message
        })
    }

    $values = $OutputValue.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $selected = foreach ($v in $values) {
        $o = [DocumentationRegistry]::FindOutput($v)
        if (-not $o) {
            & $recordFailure "Lookup" $null "" "No registered output for value '$v'" $null
            continue
        }
        $o
    }
    if (-not $selected) {
        $details = @($failures | ForEach-Object Message | Select-Object -Unique)
        throw "Documentation output failed in $($failures.Count) hook invocation(s): $($details -join '; ')"
    }

    # Keep the document hierarchy deterministic at each level:
    # object group display name -> object type display name -> policy name.
    # Third key is the title the object will actually be documented under, not its
    # raw display name - see Get-DocumentationSortName.
    $orderedPolicies = @($PolicyObjects | Sort-Object `
        @{ Expression = { Get-DocumentationPolicyGroupName $_ } }, `
        @{ Expression = { Get-DocumentationPolicyTypeName $_ } }, `
        @{ Expression = { Get-DocumentationSortName $_ } })
    $script:_docRunProcessed = 0
    # Types that had no documentation support and were skipped rather than
    # written out empty. Surfaced in the summary so a short document is
    # explained by a number instead of looking like a silent failure.
    $script:_docRunSkipped = 0
    $groups = @($orderedPolicies | Group-Object { Get-DocumentationPolicyGroupId $_ } |
        Sort-Object { Get-DocumentationPolicyGroupName $_.Group[0] })

    # Output lifecycle contract - providers rely on this fixed ordering:
    #   Activate -> PreProcess        (once per provider, this loop)
    #   -> { NewObjectGroup -> { NewObjectType -> Process* -> ProcessAllObjects } }
    #   -> PostProcess                (once per provider, below)
    # PreProcess is the single place each provider (re)initializes its $script:
    # accumulators, and it always runs here before any Process call, so Process
    # may assume they exist. (#39: ordering is guaranteed by this loop - some
    # accumulators are shared across providers, so per-provider re-entry guards
    # would be fragile and are deliberately not used.)
    foreach ($output in $selected) {
        try { if ($output.Activate)   { & $output.Activate } } catch { & $recordFailure "Activate" $output "" "Activate failed for $($output.Name): $($_.Exception.Message)" $_.Exception }
        try { if ($output.PreProcess) { & $output.PreProcess } } catch { & $recordFailure "PreProcess" $output "" "PreProcess failed for $($output.Name): $($_.Exception.Message)" $_.Exception }
    }

    foreach ($groupBucket in $groups) {
        $groupName = Get-DocumentationPolicyGroupName $groupBucket.Group[0]
        foreach ($output in $selected) {
            try { if ($output.NewObjectGroup) { & $output.NewObjectGroup $groupName } } catch { & $recordFailure "NewObjectGroup" $output "" "NewObjectGroup failed for $($output.Name): $($_.Exception.Message)" $_.Exception }
        }

        $typeBuckets = @($groupBucket.Group | Group-Object { Get-DocumentationPolicyTypeId $_ } |
            Sort-Object { Get-DocumentationPolicyTypeName $_.Group[0] })
        foreach ($typeBucket in $typeBuckets) {
            $typeName = Get-DocumentationPolicyTypeName $typeBucket.Group[0]
            foreach ($output in $selected) {
                try { if ($output.NewObjectType) { & $output.NewObjectType $typeName } } catch { & $recordFailure "NewObjectType" $output "" "NewObjectType failed for $($output.Name): $($_.Exception.Message)" $_.Exception }
            }

            $typeObjects = @($typeBucket.Group | Sort-Object { Get-DocumentationSortName $_ })

            foreach ($policyObj in $typeObjects) {
                $script:_docRunProcessed++
                Write-Status `
                    -Text   ("Documenting {0} ({1} of {2})" -f $typeName, $script:_docRunProcessed, $orderedPolicies.Count) `
                    -Detail ([string]$policyObj.Name) `
                    -SkipLog -Force
                $result = Invoke-DocumentationForObject -PolicyObject $policyObj -Context (Get-DocContextSingleton -Options $Options)

                # An unsupported type with the fallback off produces no content at
                # all. Handing that to the output providers wrote an empty page -
                # or, with Include policy ID on, a page holding a single id row.
                # Skip it: the engine has already logged why, by name and type.
                if ($result.InputType -eq 'NoProvider') {
                    $script:_docRunSkipped++
                    continue
                }

                foreach ($output in $selected) {
                    try { if ($output.Process) { & $output.Process $policyObj $result } } catch { & $recordFailure "Process" $output ([string]$policyObj.Name) "Process failed for $($output.Name) on $($policyObj.Name): $($_.Exception.Message)" $_.Exception }
                }
            }

            foreach ($output in $selected) {
                try { if ($output.ProcessAllObjects) { & $output.ProcessAllObjects $typeObjects } } catch { & $recordFailure "ProcessAllObjects" $output "" "ProcessAllObjects failed for $($output.Name): $($_.Exception.Message)" $_.Exception }
            }
        }
    }

    Write-Status "Documentation - saving output files" -SkipLog -Force
    foreach ($output in $selected) {
        try { if ($output.PostProcess) { & $output.PostProcess } } catch { & $recordFailure "PostProcess" $output "" "PostProcess failed for $($output.Name): $($_.Exception.Message)" $_.Exception }
    }
    Write-Status $null

    if ($failures.Count -gt 0) {
        $details = @($failures | ForEach-Object Message | Select-Object -Unique)
        throw "Documentation output failed in $($failures.Count) hook invocation(s): $($details -join '; ')"
    }

    if ($script:_docRunSkipped -gt 0) {
        Write-Log "Documentation: $($script:_docRunSkipped) object(s) skipped - their type has no documentation support and 'Document unsupported types' is off" 2
    }

    [PSCustomObject]@{
        Outputs      = @($selected | ForEach-Object Value)
        PolicyCount  = @($PolicyObjects).Count
        SkippedCount = $script:_docRunSkipped
        FailureCount = 0
    }
}

# Singleton context per-batch so Add* state doesn't leak between objects but
# we don't pay constructor cost per object either. ResetForObject clears the
# accumulators between objects.
$script:_docContextSingleton = $null
function Get-DocContextSingleton {
    param([hashtable]$Options)
    if (-not $script:_docContextSingleton) {
        $script:_docContextSingleton = [DocumentationContext]::new($null, 'en', (Merge-DocumentationOptions $Options))
        Set-DocumentationContextRunOptions -Context $script:_docContextSingleton -Options $Options
    }
    elseif ($Options) {
        Set-DocumentationContextRunOptions -Context $script:_docContextSingleton -Options $Options
    }
    return $script:_docContextSingleton
}
