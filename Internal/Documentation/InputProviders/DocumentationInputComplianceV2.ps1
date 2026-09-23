# Compliance V2 input provider — schema-driven compliance policies on the
# /deviceManagement/compliancePolicies endpoint.
#
# Ported from old Extensions/Documentation.psm1:1444 (Invoke-TranslateComplianceV2-
# Object). Claims @odata.type='#microsoft.graph.deviceManagementCompliancePolicy'.
#
# Mirrors DocumentationInputSettingsCatalog.ps1 — same recursive setting walker
# (Add-SettingsSetting), same batch caches on the [DocumentationContext]
# ($ctx.CfgCategories, $ctx.CachedCfgSettings), same Category/SubCategory
# grouping at the end. Differences:
#   - Settings endpoint: /deviceManagement/compliancePolicies/<id>/settings
#   - Categories endpoint: /deviceManagement/complianceCategories with the
#     linux/linuxMdm template filter (matches old code at L1471)
#   - platformSupported row uses $obj.platforms directly (compliance policies
#     are single-platform, no templateReference indirection)

function Invoke-InitializeComplianceV2Input {
    Add-DocumentationInputProvider ([PSCustomObject]@{
        Name      = 'ComplianceV2'
        Order     = 30
        Match     = { param($PolicyObject) $PolicyObject.JsonObject.'@odata.type' -eq '#microsoft.graph.deviceManagementCompliancePolicy' }
        Translate = { param($PolicyObject, $Context) Invoke-TranslateComplianceV2Object $PolicyObject $Context }
    })
}

function Invoke-TranslateComplianceV2Object {
    param($PolicyObject, [DocumentationContext]$Context)

    $obj = $PolicyObject.JsonObject

    # --- BasicInfo header rows ---
    Add-BasicDefaultValues $PolicyObject
    Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'ConfigurationTypes.settingsCatalog') '@odata.type'

    if ($obj.platforms) {
        $platformType = Get-LanguageString "Platform.$($obj.platforms)"
        if ($platformType) {
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.platformSupported') $platformType 'platforms'
        }
    }

    Add-BasicAdditionalValues $PolicyObject
    # --- Settings ---
    # Prefer in-policy settings WITH inline settingDefinitions. Inline settings
    # WITHOUT definitions (the hydrate body only does ?$expand=Settings) would
    # send the walker into a per-setting /configurationSettings/{id} N+1 — the
    # same trap the Settings Catalog provider fixed; enrich instead.
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

    # Source-tenant-specific: fetches THIS policy's settings by id (404s elsewhere).
    # Stays gated on -not SourceTenantUnavailable; the walker's generic per-setting
    # configurationSettings/{id} fallback resolves schema when source is gone.
    if (-not $hasDefs -and -not $Context.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
        try {
            $headers = @{}
            if ($Context.Language -and $Context.Language -ne 'en') {
                $headers['Accept-Language'] = $Context.Language
            }
            $resp = Invoke-MSGraphAPI -Url "/deviceManagement/compliancePolicies('$($obj.Id)')/settings?`$expand=settingDefinitions&`$top=1000" -AdditionalHeaders $headers
            if ($resp -and $resp.Value) {
                $cfgSettings = @($resp.Value)
            }
        }
        catch {
            Write-LogError "Failed to fetch settings for compliance policy $($obj.Id)" $_.Exception
        }
    }

    if ($cfgSettings.Count -eq 0) {
        Write-Log "ComplianceV2: no settings to document for $($obj.name)" 2
        return
    }

    # --- Generic schema caches (session-persistent, shared by reference with the
    #     Settings Catalog provider so the walker's per-setting definition fetches
    #     warm one shared cache). ---
    $Context.CachedCfgSettings = Get-CacheObject "DocCfgSettingDefinitions" $Context.CachedCfgSettings
    Set-CacheObject "DocCfgSettingDefinitions" $Context.CachedCfgSettings -Persistent

    # --- Categories (batch-cached). Old code unions linux/linuxMdm template
    # categories into the same $global:cfgCategories the Settings Catalog uses;
    # we mirror that by appending to $ctx.CfgCategories rather than replacing.


    $Context.CfgCategories = Get-CacheObject "CfgCategories" (@())

    # Generic schema (complianceCategories) - same on every tenant - gated only on connectivity.
    if (-not ($Context.CfgCategories | Where-Object { $_.settingUsage -eq 'compliance' }) -and (Test-DocumentationGraphAvailable)) {
        try {
            $resp = Invoke-MSGraphAPI -Url "/deviceManagement/complianceCategories" -ODataMetadata 'minimal' -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context)
            #$resp = Invoke-MSGraphAPI -Url "/deviceManagement/complianceCategories?`$templateCategory=True&`$filter=platforms has 'linux' and technologies has 'linuxMdm'"
            $Context.CfgCategories += @($resp.Value)
            Set-CacheObject "CfgCategories" $Context.CfgCategories -Persistent
        }
        catch {
            Write-LogError 'Failed to fetch compliance categories' $_.Exception
        }
    }

    # --- Seed definition cache from inline settingDefinitions ---
    foreach ($cfgSetting in $cfgSettings) {
        if (-not $cfgSetting.settingDefinitions) { continue }
        $defObj = $cfgSetting.settingDefinitions | Where-Object id -EQ $cfgSetting.settingInstance.settingDefinitionId | Select-Object -First 1
        if ($defObj -and -not $Context.CachedCfgSettings.ContainsKey($defObj.Id)) {
            $Context.CachedCfgSettings[$defObj.Id] = $defObj
        }
    }

    # --- Walk each top-level setting via the shared SettingsCatalog walker ---
    Reset-SettingsCatalogPolicyBuffer
    foreach ($cfgSetting in $cfgSettings) {
        Add-SettingsSetting $cfgSetting.settingInstance $cfgSetting.settingDefinitions | Out-Null
    }

    # --- Drain buffer into SettingsData grouped by (Category, SubCategory) ---
    $buffer = Get-SettingsCatalogPolicyBuffer
    $unique = $buffer |
              Select-Object @{ l='CategoryID';    e={ $_.CategoryDefinition.Id    } },
                            @{ l='SubCategoryID'; e={ $_.SubCategoryDefinition.Id } } -Unique

    foreach ($pair in $unique) {
        $rows = $buffer | Where-Object {
            $_.CategoryDefinition.Id    -eq $pair.CategoryID -and
            $_.SubCategoryDefinition.Id -eq $pair.SubCategoryID
        }
        foreach ($row in $rows) {
            if ($row.Show -eq $false) { continue }
            $Context.AddSetting($row)
        }
    }
}

Invoke-InitializeComplianceV2Input
