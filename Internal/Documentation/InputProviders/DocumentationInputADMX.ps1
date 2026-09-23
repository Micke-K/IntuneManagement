# Administrative Templates (ADMX / Group Policy) input provider.
#
# Ported from old Extensions/Documentation.psm1:916 (Invoke-TranslateADMXObject,
# ~190 LOC). Claims @odata.type='#microsoft.graph.groupPolicyConfiguration'
# and translates each definitionValue into a documented setting row.
#
# Definition resolution falls back through three sources, in order:
#   1. inline $definitionValue.definition  (live $expand=definition export)
#   2. #Definition_* flat fields the new project's exporter promotes for
#      offline use (#Definition_displayName / categoryPath / classType / Id)
#   3. live Graph fetch via Invoke-MSGraphAPI
# Rows whose displayName can't be resolved (no inline, no embedded, no Graph)
# are skipped — matches the golden fixture's offline behavior.
#
# Presentation values (the configured values for each ADMX setting) translate
# differently per presentation type:
#   DropdownList   -> map raw value to item.displayName
#   ValueList      -> name=value pairs joined
#   MultiText      -> values joined
#   Boolean/Decimal/LongDecimal/Text -> raw value
#
# Those joined strings stay in Value / ValueWithLabel / RawValue, which Compare,
# CSV, Word and JSON all read. The RENDERED table (FullValueTable, used by the
# HTML / Markdown / Atlassian outputs) is built separately by
# ConvertTo-ADMXValueTable so a list or multi-text setting gets one row per item
# instead of one cell holding everything joined, and an explicit-value list gets
# its own Key column.
#
# Settings sorted by CategoryPath at end (matches old code's tail sort).

function Invoke-InitializeADMXInput {
    Add-DocumentationInputProvider ([PSCustomObject]@{
        Name      = 'ADMX'
        Order     = 50
        Match     = { param($PolicyObject) $PolicyObject.JsonObject.'@odata.type' -eq '#microsoft.graph.groupPolicyConfiguration' }
        Translate = { param($PolicyObject, $Context) Invoke-TranslateADMXPolicyObject $PolicyObject $Context }
    })
}

function Invoke-TranslateADMXPolicyObject {
    param($PolicyObject, [DocumentationContext]$Context)

    $valueProperty = if ($Context.Options.ValueOutputProperty -eq 'valueWithLabel') { 'ValueWithLabel' } else { 'Value' }
    $Context.DisplayProperties = @('Name','Status','Value','Category','CategoryPath','RawValue','ValueWithLabel','Created','Modified','Class','DefinitionId')
    $Context.DefaultDocumentationProperties = @('Name','Status',$valueProperty)

    $obj = $PolicyObject.JsonObject

    # --- BasicInfo header ---
    Add-BasicDefaultValues $PolicyObject -SkipProperties @('')
    Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'Titles.groupPolicy') '@odata.type'
    # (Platform supported deliberately omitted — old code at L922 has it commented out;
    #  groupPolicyConfiguration is Windows-only by definition.)
    Add-BasicAdditionalValues $PolicyObject
    # --- Categories cache (batch-scoped, lazy) ---
    # Generic ADMX category/definition catalog - same on every tenant. Seed from the
    # session-persistent cache (like CfgCategories); on a miss, fetch from any
    # connected tenant and warm the cache so later runs in the session skip the GET.
    if (-not $Context.ADMXCategories -or $Context.ADMXCategories.Count -eq 0) {
        $Context.ADMXCategories = Get-CacheObject "DocADMXCategories" (@())
    }
    if ((-not $Context.ADMXCategories -or $Context.ADMXCategories.Count -eq 0) -and
        (Test-DocumentationGraphAvailable)) {
        try {
            $url = "deviceManagement/groupPolicyCategories?`$expand=parent(`$select=id,displayName,isRoot),definitions(`$select=id,displayName,categoryPath,classType,policyType)&`$select=id,displayName,isRoot"
            $resp = Invoke-MSGraphAPI -Url $url -ODataMetadata 'skip' -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context)
            if ($resp.Value) {
                $Context.ADMXCategories = @($resp.Value)
                Set-CacheObject "DocADMXCategories" $Context.ADMXCategories -Persistent
            }
        }
        catch {
            Write-LogError 'Failed to load ADMX group policy categories' $_.Exception
        }
    }

    # --- definitionValues ---
    $definitionValues = @()
    if ($obj.definitionValues) {
        $definitionValues = @($obj.definitionValues)
    }
    elseif (-not $Context.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
        # Source-tenant-specific: THIS policy's definitionValues by id (404s elsewhere).
        try {
            $url = "deviceManagement/groupPolicyConfigurations('$($obj.Id)')/definitionValues?`$expand=definition(`$select=id,classType,displayName,policyType,groupPolicyCategoryId)"
            $resp = Invoke-MSGraphAPI -Url $url -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context)
            $definitionValues = @($resp.Value)
        }
        catch {
            Write-LogError "Failed to load definitionValues for ADMX policy $($obj.Id)" $_.Exception
        }
    }
    if ($definitionValues.Count -eq 0) { return }

    $enabledStr  = Get-LanguageString 'Inputs.enabled'
    $disabledStr = Get-LanguageString 'Inputs.disabled'
    $propertyStr = Get-LanguageString 'ApplicabilityRules.GridLabel.property'
    $valueStr    = Get-LanguageString 'ApplicabilityRules.GridLabel.value'
    $keyStr      = Get-LanguageString 'SettingDetails.keyColumn'

    ## ToDo: Preload the presentation Definitions for all definitionValues with presentationValues defined in one batch
    # e.g. $definitionValues | Where presentationValues -ne $null -> Add to batch and fetch all in one call.

    $rows = @()
    foreach ($defValue in $definitionValues) {
        $definition = Resolve-ADMXDefinition $defValue $Context
        if (-not $definition -or -not $definition.displayName) {
            # Unresolvable in current mode — skip (matches golden's offline behavior)
            continue
        }

        # Category path: prefer the definition's own field; fall back to the cached
        # categories lookup when only an id is available.
        $categoryPath = $definition.categoryPath
        if (-not $categoryPath -and $Context.ADMXCategories.Count -gt 0) {
            $matched = $Context.ADMXCategories.definitions | Where-Object { $_.id -eq $definition.id } | Select-Object -First 1
            if ($matched) { $categoryPath = $matched.categoryPath }
        }

        # Presentation values — only present when the policy carries configured values
        $presentationValues = Resolve-ADMXPresentationValues $defValue $obj $Context

        $values          = @()
        $valuesWithLabel = @()
        $rawValues       = @()
        # One entry per presentation - its label plus the rows it contributes to
        # the rendered table. Multi-valued presentations contribute one row per
        # item. The flat $values / $valuesWithLabel / $rawValues below are
        # deliberately built exactly as before: Compare joins on them.
        $presEntries     = @()
        # Presentation values present: map each to its label + value (per type)
        foreach ($pv in $presentationValues) {
            # Generic presentation metadata (label/dropdown items) resolved from any connected tenant.
            if (-not $pv.presentation -and $pv.'presentation@odata.bind' -and (Test-DocumentationGraphAvailable)) {
                try {
                    $pres = Invoke-MSGraphAPI -Url $pv.'presentation@odata.bind' -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context)
                    if ($pres) { $pv | Add-Member -MemberType NoteProperty -Name 'presentation' -Value $pres -Force }
                }
                catch { }
            }

            $rawValue = $pv.value
            $label    = $pv.presentation.label
            $value    = $null
            $valueRows     = @()

            switch ($pv.presentation.'@odata.type') {
                '#microsoft.graph.groupPolicyPresentationDropdownList' {
                    $value = ($pv.presentation.items | Where-Object value -EQ $rawValue).displayName
                    $valueRows  = @([PSCustomObject]@{ Key = $null; Value = $value })
                }
                default {
                    switch ($pv.'@odata.type') {
                        '#microsoft.graph.groupPolicyPresentationValueList' {
                            $arr = @()
                            foreach ($v in $pv.values) {
                                $arr  += "$($v.name)$($Context.PropertySeparator)$($v.value)"
                                $valueRows += [PSCustomObject]@{ Key = $v.name; Value = $v.value }
                            }
                            $value = $arr -join $Context.ObjectSeparator
                            # A plain <list> (no explicitValue) stores each item in
                            # 'name' and leaves 'value' empty. Those are single-column
                            # items, not key/value pairs - fold name into the value.
                            if (@($valueRows | Where-Object { "$($_.Value)" -ne '' }).Count -eq 0) {
                                $valueRows = @($valueRows | ForEach-Object { [PSCustomObject]@{ Key = $null; Value = $_.Key } })
                            }
                        }
                        '#microsoft.graph.groupPolicyPresentationValueMultiText' {
                            $value = $pv.values -join $Context.ObjectSeparator
                            $valueRows  = @(foreach ($v in $pv.values) { [PSCustomObject]@{ Key = $null; Value = $v } })
                        }
                        default {
                            # Boolean / Decimal / LongDecimal / Text — value is the raw scalar
                            $value = $rawValue
                            $valueRows  = @([PSCustomObject]@{ Key = $null; Value = $rawValue })
                        }
                    }
                }
            }

            $presEntries     += [PSCustomObject]@{ Label = $label; Rows = @($valueRows) }
            $valuesWithLabel += "$label $value"
            $values    += $value
            $rawValues += $rawValue
        }

        $tableValue = ConvertTo-ADMXValueTable -Entries $presEntries `
            -PropertyHeader $propertyStr -KeyHeader $keyStr -ValueHeader $valueStr

        $status = if ($defValue.enabled -eq $true) { $enabledStr } else { $disabledStr }

        $combinedValue = $status
        if ($values) {
            $combinedValue += $Context.ObjectSeparator + ($values -join $Context.ObjectSeparator)
        }

        $combinedValueWithLabel = $status
        if ($valuesWithLabel) {
            $combinedValueWithLabel += $Context.ObjectSeparator + ($valuesWithLabel -join $Context.ObjectSeparator)
        }

        $rows += [PSCustomObject]@{
            Name                   = $definition.displayName
            Description            = $definition.explainText
            Status                 = $status
            Value                  = $values -join $Context.ObjectSeparator
            CombinedValue          = $combinedValue
            ValueWithLabel         = $valuesWithLabel -join $Context.ObjectSeparator
            FullValueTable         = $tableValue
            CombinedValueWithLabel = $combinedValueWithLabel
            RawValue               = $rawValues -join $Context.PropertySeparator
            Class                  = $definition.classType
            DefinitionId           = $definition.id
            Created                = $defValue.createdDateTime
            Modified               = $defValue.lastModifiedDateTime
            Category               = $categoryPath
            CategoryPath           = $categoryPath
            EntityKey              = $definition.id   # Required for Compare
            AlwaysAddValue         = if($null -ne $defValue.enabled) { $true } else { $false }  # Always include Value if defined
        }
    }

    foreach ($row in ($rows | Sort-Object -Property CategoryPath, Name)) {
        $Context.AddSetting($row)
    }
}

# Builds the rendered value table for one ADMX setting.
#
# The column shape is decided once for the whole setting, because the HTML,
# Markdown and Atlassian renderers read their headers from the FIRST row and then
# fetch every later row by those same property names - mixing shapes inside one
# setting would silently blank cells. So: two columns (Property | Value) unless
# some presentation contributed real key/value pairs, in which case three
# (Property | Key | Value) for every row of that setting.
#
# The property label is written on the first row of each presentation only, so a
# multi-item list reads as one labelled block instead of repeating the label.
#
# $Entries is one object per presentation: { Label; Rows = @({ Key; Value }) }.
function ConvertTo-ADMXValueTable {
    param(
        $Entries,
        [string]$PropertyHeader,
        [string]$KeyHeader,
        [string]$ValueHeader
    )

    $entryArr = @($Entries)
    if ($entryArr.Count -eq 0) { return @() }

    $hasKeys = $false
    foreach ($entry in $entryArr) {
        foreach ($row in @($entry.Rows)) {
            if ("$($row.Key)" -ne '') { $hasKeys = $true; break }
        }
        if ($hasKeys) { break }
    }

    # A translation that collides with another header would throw on the ordered
    # hashtable below, so fall back to the untranslated column name.
    if ($hasKeys -and ($KeyHeader -eq $PropertyHeader -or $KeyHeader -eq $ValueHeader -or -not $KeyHeader)) {
        $KeyHeader = 'Key'
    }

    $table = @()
    foreach ($entry in $entryArr) {
        $rows = @($entry.Rows)
        # A presentation with nothing configured still shows its label.
        if ($rows.Count -eq 0) { $rows = @([PSCustomObject]@{ Key = $null; Value = $null }) }

        $first = $true
        foreach ($row in $rows) {
            $out = [ordered]@{}
            $out[$PropertyHeader] = if ($first) { $entry.Label } else { '' }
            if ($hasKeys) { $out[$KeyHeader] = $row.Key }
            $out[$ValueHeader] = $row.Value
            $table += [PSCustomObject]$out
            $first = $false
        }
    }

    # Comma so a single-row table doesn't unroll to a bare object on return.
    return ,$table
}

# Three-tier definition resolution: inline -> embedded #Definition_* flat
# fields -> live Graph fetch. Returns the synthesized/fetched definition or
# $null if nothing resolved.
function Resolve-ADMXDefinition {
    param($DefinitionValue, [DocumentationContext]$Context)

    # 1. Inline (live $expand=definition export already populated it)
    if ($DefinitionValue.definition -and $DefinitionValue.definition.displayName) {
        return $DefinitionValue.definition
    }

    # 2. Embedded #Definition_* flat fields (new project's exporter prefix)
    $embeddedDisplayName = $DefinitionValue.'#Definition_displayName'
    if ($embeddedDisplayName) {
        $syn = [PSCustomObject]@{
            id           = $DefinitionValue.'#Definition_Id'
            displayName  = $embeddedDisplayName
            classType    = $DefinitionValue.'#Definition_classType'
            categoryPath = $DefinitionValue.'#Definition_categoryPath'
            explainText  = $null
            policyType   = $null
        }
        # Attach for future calls
        $DefinitionValue | Add-Member -MemberType NoteProperty -Name 'definition' -Value $syn -Force
        return $syn
    }

    # 3. Live Graph fetch via definition@odata.bind URL. The definition is GENERIC
    #    schema (groupPolicyDefinitions) - same on every tenant - so gated only on
    #    connectivity, not on SourceTenantUnavailable.
    if ($DefinitionValue.'definition@odata.bind' -and (Test-DocumentationGraphAvailable)) {
        try {
            $url = $DefinitionValue.'definition@odata.bind'
            $def = Invoke-MSGraphAPI -Url $url -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context)
            if ($def) {
                $DefinitionValue | Add-Member -MemberType NoteProperty -Name 'definition' -Value $def -Force
                return $def
            }
        }
        catch {
            Write-LogError "Failed to fetch ADMX definition from $($DefinitionValue.'definition@odata.bind')" $_.Exception
        }
    }

    return $null
}

# Resolves presentation values + their presentation metadata. Returns array,
# possibly empty for definitionValues with no configured presentationValues
# (i.e. ADMX settings that are simply Enabled/Disabled with no inputs).
function Resolve-ADMXPresentationValues {
    param($DefinitionValue, $PolicyObject, [DocumentationContext]$Context)

    # Already inline? Order them by the canonical presentation order if we can.
    if ($DefinitionValue.presentationValues -and $DefinitionValue.presentationValues.Count -gt 0) {
        # The canonical presentation list is generic schema; reorder only needs a
        # connected tenant. Without one, keep the inline order.
        if (-not (Test-DocumentationGraphAvailable)) {
            return @($DefinitionValue.presentationValues)
        }
        # Live: pull the canonical presentation list so we can reorder
        try {
            $url = "$($DefinitionValue.'definition@odata.bind')/presentations"
            $resp = Invoke-MSGraphAPI -Url $url -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context)
            $canon = @($resp.Value)
            if ($canon.Count -gt 0) {
                $ordered = @()
                foreach ($p in $canon) {
                    $match = $DefinitionValue.presentationValues | Where-Object 'presentation@odata.bind' -Like "*$($p.Id)*" | Select-Object -First 1
                    if ($match) { $ordered += $match } else { $ordered = @(); break }
                }
                if ($ordered.Count -gt 0) { return $ordered }
            }
        } catch { }
        return @($DefinitionValue.presentationValues)
    }

    # Live fetch (when fixture exported without presentationValues inline). These are
    # the policy's CONFIGURED values, fetched by policy id - source-tenant-specific
    # (404s elsewhere) - so gated on -not SourceTenantUnavailable.
    if ($DefinitionValue.id -and -not $Context.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
        try {
            # Should never get here - $DefinitionValue.id will be empty.
            $url = "/deviceManagement/groupPolicyConfigurations/$($PolicyObject.id)/definitionValues/$($DefinitionValue.id)/presentationValues?`$expand=presentation"
            $resp = Invoke-MSGraphAPI -Url $url -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context)
            return @($resp.Value)
        }
        catch {
            Write-LogError "Failed to fetch ADMX presentationValues for $($DefinitionValue.id)" $_.Exception
        }
    }

    return @()
}

Invoke-InitializeADMXInput
