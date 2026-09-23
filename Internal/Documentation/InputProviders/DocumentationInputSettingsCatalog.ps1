# Settings Catalog input provider.
#
# Ported from old Extensions/Documentation.psm1:1107 (Invoke-TranslateSettings-
# Object, ~100 LOC). Claims @odata.type='#microsoft.graph.deviceManagement
# ConfigurationPolicy' and translates the policy's settings via the recursive
# walker (Add-SettingsSetting in SettingsCatalogWalker.ps1).
#
# Live Graph dependencies (resolved through Invoke-MSGraphAPI):
#   /deviceManagement/configurationPolicies/{id}/settings?$expand=settingDefinitions
#   /deviceManagement/configurationCategories?$filter=platforms has 'windows10' and technologies has 'mdm'
#   /deviceManagement/configurationSettings/{id}  (per-setting fallback when defs aren't expanded)
#
# These are batch-cached on the [DocumentationContext] ($ctx.CfgCategories,
# $ctx.CachedCfgSettings) so a bulk run pays the cost once. The per-policy
# settings fetch (by id) is source-tenant-specific and skipped when
# $ctx.SourceTenantUnavailable; the GENERIC schema (setting definitions via
# the walker's configurationSettings/{id} fallback, and configurationCategories)
# is still resolved from any connected tenant (Test-DocumentationGraphAvailable).
# With no tenant at all the provider still runs, producing raw IDs.
#
# OFFLINE SMOKE TEST DEFERRED: golden-file validation against the provided
# fixture (C:/Intune/OldDocumentation/SettingsCatalog/[Testing] Windows 11
# Settings.json) needs the policy re-exported with $expand=settings($expand=
# settingDefinitions) + a sidecar fixture for scope tags. Until then this
# provider is exercised live against a tenant; its structure mirrors the old
# code's so trust-the-port applies.

function Invoke-InitializeSettingsCatalogInput {
    Add-DocumentationInputProvider ([PSCustomObject]@{
        Name      = 'SettingsCatalog'
        Order     = 20
        Match     = { param($PolicyObject) $PolicyObject.JsonObject.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationPolicy' }
        Translate = { param($PolicyObject, $Context) Invoke-TranslateSettingsCatalogObject $PolicyObject $Context }
    })
}

function Invoke-TranslateSettingsCatalogObject {
    param($PolicyObject, [DocumentationContext]$Context)

    $obj = $PolicyObject.JsonObject

    # --- BasicInfo header rows ---
    Add-BasicDefaultValues $PolicyObject
    Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'ConfigurationTypes.settingsCatalog') '@odata.type'

    if ($obj.templateReference.templateId) {
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.Category') (Get-IntentCategoryName $obj.templateReference.templateFamily) 'templateFamily'
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.policyType') $obj.templateReference.templateDisplayName 'templateDisplayName'
    }

    if ($obj.platforms) {
        $platformType = Get-LanguageString "Platform.$($obj.platforms)"
        if ($platformType) {
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.platformSupported') $platformType 'platforms'
        }
    }

    Add-BasicAdditionalValues $PolicyObject
    # --- Settings ---
    # Prefer in-policy settings (export / hydrate with $expand=settings has them
    # inline). When settingDefinitions are not also inline — the hydrate body URL
    # only does `?$expand=Settings`, NOT `?$expand=Settings($expand=settingDefinitions)`
    # — the SettingsCatalog walker falls back to a sequential per-setting
    # /configurationSettings/{id} GET (one round-trip per settingInstance),
    # which scales linearly with setting count and crushes bulk-doc runs.
    # One enrich call per policy collapses that N+1 to a single per-policy call.
    $cfgSettings = @()
    if ($obj.Settings -and ($obj.Settings | Measure-Object).Count -gt 0) {
        $cfgSettings = @($obj.Settings)
    }

    $hasDefs = $false
    foreach ($s in $cfgSettings) {
        if ($s.settingDefinitions -and ($s.settingDefinitions | Measure-Object).Count -gt 0) {
            $hasDefs = $true
            break
        }
    }

    # Bulk runs: Initialize-DocumentationRunPrefetch already fetched these in
    # one Graph $batch — consume from the per-run cache (authoritative for this
    # run, even when empty, so an empty-settings policy doesn't trigger a
    # redundant live GET). The live GET below is the lazy fallback for the
    # single-policy Get-GraphDocumentation path.
    if (-not $hasDefs -and $Context.PrefetchedPolicySettings.ContainsKey([string]$obj.Id)) {
        $cfgSettings = @($Context.PrefetchedPolicySettings[[string]$obj.Id])
        $hasDefs = $true
    }

    # Source-tenant-specific: fetches THIS policy's settings by id, which 404s on
    # any other tenant. Stays gated on -not SourceTenantUnavailable. When the
    # source is gone but the export carries settings inline (no defs), the walker's
    # generic per-setting configurationSettings/{id} fallback resolves the schema.
    if (-not $hasDefs -and -not $Context.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
        try {
            $headers = @{}
            if ($Context.Language -and $Context.Language -ne 'en') {
                $headers['Accept-Language'] = $Context.Language
            }
            $resp = Invoke-MSGraphAPI -Url "/deviceManagement/configurationPolicies('$($obj.Id)')/settings?`$expand=settingDefinitions&`$top=1000" -AdditionalHeaders $headers -ODataMetadata 'minimal'
            if ($resp -and $resp.Value) {
                $cfgSettings = @($resp.Value)
            }
        }
        catch {
            Write-LogError "Failed to fetch settings for policy $($obj.Id)" $_.Exception
        }
    }

    if ($cfgSettings.Count -eq 0) {
        Write-Log "SettingsCatalog: no settings to document for $($obj.name)" 2
        return
    }

    # Schema caching, the walk and the (Category, SubCategory) grouping are shared
    # with the MAM app-configuration handler - see
    # Get-SettingsCatalogDocumentationRows in Core/SettingsCatalogWalker.ps1.
    foreach ($row in (Get-SettingsCatalogDocumentationRows $cfgSettings $Context)) {
        $Context.AddSetting($row)
    }

    Invoke-DocumentationSettingsCatalogPostProcess $obj $Context
}

# Settings Catalog uses an intent-style category mapping that's distinct from
# Get-DocObjectTypeString (which is for group/category headers in the OUTPUT,
# not for BasicInfo rows). Delegates to the Intent provider's
# Get-IntentCategoryFromTemplateType (the port of old Documentation.psm1:1523
# Get-IntentCategory), so endpoint-security-family catalogs show the localized
# category name instead of the raw templateFamily (e.g. endpointSecurityAntivirus).
function Get-IntentCategoryName {
    param($TemplateType)
    if (-not $TemplateType) { return '' }
    if (Get-Command Get-IntentCategoryFromTemplateType -ErrorAction SilentlyContinue) {
        $mapped = Get-IntentCategoryFromTemplateType $TemplateType
        if ($mapped) { return $mapped }
    }
    if ($TemplateType -is [string]) { return $TemplateType }
    return "$TemplateType"
}

Invoke-InitializeSettingsCatalogInput
