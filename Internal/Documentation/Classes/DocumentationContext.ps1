# Per-object documentation context. Replaces all of the $script:* accumulators
# the old Documentation.psm1 reset at the top of every Get-ObjectDocumentation
# call. Input providers and *DocHandler classes mutate the lists via Add*()
# methods; the engine calls ToResult() to produce the per-object PSCustomObject
# the output providers consume.
#
# Substitution map (from DocumentationMigration.md):
#   $global:chkIncludeScripts.IsChecked        -> Options.IncludeScripts
#   $global:chkExcludeScriptSignature.IsChecked -> Options.ExcludeScriptSignature
#   $global:documentationLanguage              -> Language
#   $script:objectBasicInfo                    -> BasicInfo
#   $script:objectSettingsData                 -> SettingsData
#   $script:objectComplianceActionData         -> ComplianceActions
#   $script:applicabilityRules                 -> ApplicabilityRules
#   $script:objectAssignments                  -> Assignments
#   $script:objectScripts                      -> Scripts
#   $script:customTables                       -> CustomTables
#   $script:admxCategories                     -> ADMXCategories
#   $script:ObjectTypeFullTable                -> ObjectTypeFullTable
#   $script:scopeTags                          -> ScopeTags
#   $script:languageStrings                    -> LanguageStrings
#   $script:offlineDocumentation               -> SourceTenantUnavailable (formerly OfflineDocumentation)
#   $script:settingsProperties, CurrentSubCategory, ValueOutputProperty -> per-call
#     fields used during input-provider walks; added incrementally as providers land

class DocumentationContext {
    # ---- Inputs ----
    [object]   $PolicyObject
    [string]   $Language
    [hashtable]$Options

    # ---- Cross-batch caches (lazily populated; survive ResetForObject so they
    #      amortize across all policies in one bulk run, matching old code's
    #      $global:* caches) ----
    [hashtable]$LanguageStrings   = @{}
    [object[]] $ScopeTags         = @()                                # /deviceManagement/roleScopeTags
    [hashtable]$CachedCfgSettings = @{}                                # SettingsCatalog: settingDefinitionId -> definition
    [object[]] $CfgCategories     = @()                                # /deviceManagement/configurationCategories + /complianceCategories
    # Resolved assignment-filter display names (filterId -> displayName, or $null
    # for IDs the directory didn't return on lookup so we don't keep retrying).
    # Survives ResetForObject so a bulk-doc run resolves each filter exactly
    # once across all policies.
    [hashtable]$FilterNamesById   = @{}
    # True once the tenant-wide assignment-filter list has been loaded into
    # FilterNamesById (from the login-time dependency cache, the run prefetch
    # batch, or the lazy fallback in Add-AssignmentsForObject). Acts as a
    # negative-cache stamp too, so a failed tenant-wide fetch isn't retried
    # per policy.
    [bool]     $FiltersLoaded     = $false
    # Resolved assignment-group display names (groupId -> displayName, or $null
    # for IDs the directory didn't return on lookup so we don't keep retrying).
    # Survives ResetForObject so one bulk-doc run resolves each group at most
    # once across all policies. Front-loaded in one getByIds batch by the prefetch
    # (when the "Use Batch API" setting is on); otherwise filled lazily per-object
    # by Add-AssignmentsForObject. Same negative-cache semantics as FilterNamesById.
    [hashtable]$GroupNamesById    = @{}
    # Per-run prefetch of policy sub-resources, populated by
    # Initialize-DocumentationRunPrefetch in one Graph $batch before the output
    # loop: policyId -> settings[] (with settingDefinitions expanded) for
    # Settings Catalog / Compliance V2 policies. Cleared at the start of every
    # run so a re-documented policy reflects current tenant state.
    [hashtable]$PrefetchedPolicySettings = @{}
    # When true, the SOURCE tenant the export came from is not reachable, so
    # source-tenant-specific lookups (assignments->groups, scope-tag/filter/app
    # names, named locations, ToU, linked certs, reusable settings, per-policy
    # settings-by-id) are skipped. It does NOT mean "no tenant at all": generic
    # Intune schema (setting definitions/categories, ADMX, intent templates,
    # resourceOperations) is identical on every tenant and is still fetched from
    # whatever tenant is connected, gated by Test-DocumentationGraphAvailable.
    [bool]     $SourceTenantUnavailable

    # Title the document by something other than the policy's display name, and
    # fold the Basics rows into the settings table instead of emitting a second
    # table. Both exist for the default enrollment policies: Windows Hello for
    # Business, Windows Restore, the device limit, the platform restrictions and
    # the enrollment status page all ship under the SAME display name, "All users
    # and all devices", so a document headed by that name does not say which
    # policy it is, and the two tables between them hold only a handful of rows.
    # Set per object by the ObjectInfo customizer; both reset between objects.
    [string]   $DocumentName   = $null
    [bool]     $MergeBasicInfo = $false

    # ---- Per-object accumulators (drained by ToResult) ----
    [System.Collections.Generic.List[object]] $BasicInfo          = [System.Collections.Generic.List[object]]::new()
    [System.Collections.Generic.List[object]] $SettingsData       = [System.Collections.Generic.List[object]]::new()
    [System.Collections.Generic.List[object]] $ComplianceActions  = [System.Collections.Generic.List[object]]::new()
    [System.Collections.Generic.List[object]] $ApplicabilityRules = [System.Collections.Generic.List[object]]::new()
    [System.Collections.Generic.List[object]] $Assignments        = [System.Collections.Generic.List[object]]::new()
    [System.Collections.Generic.List[object]] $Scripts            = [System.Collections.Generic.List[object]]::new()
    [System.Collections.Generic.List[object]] $CustomTables       = [System.Collections.Generic.List[object]]::new()
    [object[]]  $ADMXCategories = @()                                  # /deviceManagement/groupPolicyCategories?$expand=parent,definitions — batch cache, survives ResetForObject

    # Intent template caches (deviceManagementIntent input provider). Keyed by
    # templateId / categoryId; survive ResetForObject so a bulk run of N intents
    # against the same template only pays one round-trip per category.
    [hashtable] $IntentCategories            = @{}                     # templateId -> categories[] (with $expand=settingDefinitions)
    [hashtable] $IntentCatRecommendedSettings = @{}                    # categoryId -> recommendedSettings[]

    # ---- Cross-type accumulator (for ScopeTags consolidated table at end of batch) ----
    [hashtable] $ObjectTypeFullTable = @{}

    # ---- Result fields populated by input providers / handlers ----
    [string[]] $DefaultDocumentationProperties = @('Name','Value')
    [object[]] $DisplayProperties              = @()
    [string]   $ErrorText
    [string]   $InputType
    [bool]     $UpdateFilteredObject
    [object[]] $UnconfiguredProperties         = @()

    # ---- Walker / translate-primitive state (replaces $script:currentObject,
    #      $script:CurrentSubCategory, $script:propLevel, $script:propertySeparator,
    #      $script:objectSeparator) ----
    [object]   $CurrentObject
    [string]   $CurrentSubCategory
    [int]      $PropLevel = 0
    [string]   $PropertySeparator = ','
    [string]   $ObjectSeparator   = [System.Environment]::NewLine

    DocumentationContext() {
        $this.Options  = [DocumentationContext]::DefaultOptions()
        $this.Language = 'en'
    }

    DocumentationContext([object]$PolicyObject, [string]$Language, [hashtable]$Options) {
        $this.PolicyObject = $PolicyObject
        $this.Language     = if ($Language) { $Language } else { 'en' }
        $this.Options      = if ($Options)  { $Options }  else { [DocumentationContext]::DefaultOptions() }
    }

    # Default Options hashtable. Keys recognised by the engine:
    #   IncludeScripts          [bool]  emit Scripts collection (default true)
    #   ExcludeScriptSignature  [bool]  strip script signing blocks (default false)
    #   IncludePolicyId         [bool]  engine appends an 'Id' BasicInfo row after
    #                                   handler/input dispatch when true (default false)
    static [hashtable] DefaultOptions() {
        return @{
            IncludeScripts         = $true
            ExcludeScriptSignature = $false
            IncludePolicyId        = $false
            Language               = 'en'
            PropertySeparator      = ';'
            ObjectSeparator        = [System.Environment]::NewLine
            # When true, the engine post-step that translates $obj.assignments
            # into Assignment rows is skipped. Old engine gated this on
            # $global:chkExcludeAssignments (default off — assignments included).
            ExcludeAssignments     = $false
            # ObjectInfo walker: when a property is null on the input, the walker
            # can either skip emitting a row OR substitute the prop's declared
            # unconfiguredValue / defaultValue / emptyValueResourceKey. Old engine
            # gates these on UI checkboxes. Legacy defaults substitute an explicit
            # unconfigured value but do not substitute a declared default value.
            SetUnconfiguredValue   = $true
            SetDefaultValue        = $false
            SkipNotConfigured      = $false
            SkipDefaultValues      = $false
            SkipDisabled           = $true
            NotConfiguredText      = 'notConfigured'
            ValueOutputProperty    = 'value'
            SourceTenantUnavailable = $false
            # When true, output providers skip the document-info header they emit
            # at the top of a Full document (Organization / Generated by / Generated
            # date). Generic because every provider writes the same block; read via
            # Get-DocumentationOption so all providers honour it consistently.
            SkipDocumentInfo       = $false
            # A policy type that no handler and no input provider claims is
            # documented by the generic fallback - basic info plus one row per
            # property - instead of producing a page with nothing on it. On by
            # default: a rough table beats a blank page, and the alternative
            # (Enrollment Notifications, Windows Hello for Business and friends
            # documenting to nothing) reads as a bug to everyone who hits it.
            # Off means those types are logged and skipped entirely, not emitted
            # as an empty document. Overridden per run via -Options, and per
            # user by the Output Settings checkbox, which saves to this key.
            FallbackDocumentation  = $true
            Outputs                = @{}
        }
    }

    # Reset per-object accumulators so the same context instance can be reused
    # across objects in a bulk run (cheaper than constructing a new one).
    [void] ResetForObject([object]$PolicyObject) {
        $this.PolicyObject = $PolicyObject
        $this.CurrentObject = if ($PolicyObject -and $PolicyObject.PSObject.Properties['JsonObject']) { $PolicyObject.JsonObject } else { $PolicyObject }
        $this.BasicInfo.Clear()
        $this.SettingsData.Clear()
        $this.ComplianceActions.Clear()
        $this.ApplicabilityRules.Clear()
        $this.Assignments.Clear()
        $this.Scripts.Clear()
        $this.CustomTables.Clear()
        # ADMXCategories deliberately not cleared — batch-scoped cache (matches old $script:admxCategories)
        $this.DefaultDocumentationProperties = @('Name','Value')
        $this.DisplayProperties = @()
        $this.ErrorText         = $null
        $this.InputType         = $null
        $this.UpdateFilteredObject = $false
        $this.UnconfiguredProperties = @()
        $this.CurrentSubCategory = $null
        $this.PropLevel = 0
        $this.DocumentName   = $null
        $this.MergeBasicInfo = $false
    }

    # ---- Add* methods (used by input providers / handlers) ----

    [void] AddBasic([string]$Name, [object]$Value) {
        $this.AddBasic($Name, $Value, $null)
    }

    # Preferred overload — pass the source field name (e.g. 'displayName',
    # 'state', 'createdDateTime') as EntityKey. Compare logic and other
    # downstream tools key rows by EntityKey rather than the localized Name.
    [void] AddBasic([string]$Name, [object]$Value, [string]$EntityKey) {
        $this.BasicInfo.Add([PSCustomObject]@{
            Name = $Name; Value = $Value; EntityKey = $EntityKey
        })
    }

    [void] AddProperty([string]$Name, [object]$Value) {
        $this.SettingsData.Add([PSCustomObject]@{ Name = $Name; Value = $Value })
    }

    [void] AddProperty([string]$Name, [object]$Value, [string]$Category) {
        $this.SettingsData.Add([PSCustomObject]@{ Name = $Name; Value = $Value; Category = $Category })
    }

    [void] AddProperty([string]$Name, [object]$Value, [string]$Category, [string]$SubCategory) {
        $this.SettingsData.Add([PSCustomObject]@{ Name = $Name; Value = $Value; Category = $Category; SubCategory = $SubCategory })
    }

    [void] AddSetting([object]$Setting) {
        # Every settings row must carry a stable, language-independent
        # EntityKey — it is the join key for documentation-based compare (and
        # any other downstream tooling). Manifest/Profile/handler rows set it
        # themselves; Settings Catalog / Intent walker rows carry the identity
        # as SettingId (+ ParentSettingId / RowIndex), so derive it here:
        #   [<ParentSettingId>/]<SettingId>[#<RowIndex>]
        # RowIndex > 0 disambiguates repeated group-collection rows (e.g.
        # firewall rule lists) — positional matching, same as the old project.
        if ($Setting -and (-not $Setting.PSObject.Properties['EntityKey'] -or
                [string]::IsNullOrEmpty([string]$Setting.EntityKey))) {
            $key = $null
            if ($Setting.PSObject.Properties['SettingId'] -and $Setting.SettingId) {
                $key = [string]$Setting.SettingId
                if ($Setting.PSObject.Properties['ParentSettingId'] -and $Setting.ParentSettingId) {
                    $key = "$($Setting.ParentSettingId)/$key"
                }
                if ($Setting.PSObject.Properties['RowIndex'] -and [int]$Setting.RowIndex -gt 0) {
                    $key = "$key#$($Setting.RowIndex)"
                }
            }
            if ($key) {
                $Setting | Add-Member -MemberType NoteProperty -Name 'EntityKey' -Value $key -Force
            }
            else {
                Write-LogDebug "Documentation row without EntityKey: '$($Setting.Name)' ($($this.InputType))"
            }
        }
        $this.SettingsData.Add($Setting)
    }

    [void] AddComplianceAction([string]$Action, [string]$Schedule, [string]$MessageTemplate, [string]$EmailCC) {
        $this.ComplianceActions.Add([PSCustomObject]@{
            Action          = $Action
            Schedule        = $Schedule
            MessageTemplate = $MessageTemplate
            EmailCC         = $EmailCC
        })
    }

    [void] AddApplicabilityRule([string]$Rule, [string]$Property, [object]$Value) {
        $this.ApplicabilityRules.Add([PSCustomObject]@{
            Rule = $Rule; Property = $Property; Value = $Value
        })
    }

    [void] AddAssignment([object]$Assignment) {
        $this.Assignments.Add($Assignment)
    }

    [void] AddCustomTable([object]$Table) {
        $this.CustomTables.Add($Table)
    }

    # AddScript honors Options.IncludeScripts so output providers don't need to gate
    # on $global:chkIncludeScripts anymore.
    [void] AddScript([string]$Header, [string]$Caption, [string]$Content) {
        if (-not $this.Options.IncludeScripts) { return }
        if (-not $Content) { return }
        if ($this.Options.ExcludeScriptSignature) {
            $Content = [regex]::Replace(
                $Content,
                '(?ms)^\s*# SIG # Begin signature block.*?# SIG # End signature block\s*$',
                ''
            ).TrimEnd()
        }
        $this.Scripts.Add([PSCustomObject]@{
            Header        = $Header
            Caption       = $Caption
            ScriptContent = $Content
        })
    }

    # Produce the per-object result PSCustomObject that output providers consume.
    # Output-provider contract: this is the exact set of fields outputs may read.
    [PSCustomObject] ToResult() {
        $updateNotConfigured = $true
        $notConfiguredLoc = Get-LanguageString 'SettingDetails.notConfigured'
        $notConfiguredText = ''
        if($this.Options.NotConfiguredText -eq 'notConfigured') {
            $notConfiguredText = $notConfiguredLoc
        }
        elseif($this.Options.NotConfiguredText -eq 'asis') {
            $updateNotConfigured = $false
        }

        $settings = @($this.SettingsData | Where-Object {
            if((-not ($_.PSObject.Properties | Where-Object Name -eq "RawValue")) -or
                ($_.AlwaysAddValue -eq $true))
            { return $true }

            if ($this.Options.SkipDisabled -and $_.Enabled -is [bool] -and -not $_.Enabled) { return $false }
            if ($this.Options.SkipDefaultValues -and
                $null -ne $_.RawValue -and
                (($null -ne $_.DefaultValue -and $_.RawValue -eq $_.DefaultValue) -or
                 ($null -ne $_.UnconfiguredValue -and $_.RawValue -eq $_.UnconfiguredValue))) { return $false }
            if ($this.Options.SkipNotConfigured -and
                (($_.RawValue -isnot [array] -and (
                        $null -eq $_.RawValue -or
                        "$($_.RawValue)" -eq "" -or
                        "$($_.RawValue)" -eq "notConfigured")) -or
                        #($null -ne $_.UnconfiguredValue -and $_.RawValue -eq $_.UnconfiguredValue)) -or
                 ($_.RawValue -is [array] -and $_.RawValue.Count -eq 0) -or
                 ($this.UnconfiguredProperties | Where-Object EntityKey -eq $_.EntityKey))) { return $false }


            if($updateNotConfigured -and (($_.RawValue -isnot [array] -and ($null -eq $_.RawValue -or "$($_.RawValue)" -eq "" -or "$($_.RawValue)" -eq "notConfigured") -and [String]::IsNullOrEmpty($_.Value)) -or ($_.RawValue -is [array] -and ($_.RawValue | Measure-Object).Count -eq 0)))
            {
                $_.Value = $notConfiguredText
            }

            if ($this.Options.SkipNotConfigured -and $_.Value -eq $notConfiguredLoc) { 
                Write-Log "Skipping property $($_.Name) based on '$($notConfiguredLoc)' string value" 2
                return $false
            }
            return $true
        })

        if ($this.Options.NotConfiguredText -eq 'empty') {
            foreach ($setting in $settings) {
                if ($setting.Value -eq 'notConfigured') { $setting.Value = '' }
            }
        }

        # Basic-info rows carry only a localized display Value (no RawValue), so
        # apply SkipNotConfigured the same way settings do at their final check
        # above: drop any row whose value is a "not configured" string. Handlers
        # emit that value from more than one language key (SettingDetails.notConfigured
        # for the settings path, Inputs.notConfigured for basic toggles like
        # connectedAppsEnabled / credentialProviderRoleState), so match against both
        # so this stays correct if a locale ever diverges the two.
        $basicRows = @($this.BasicInfo)
        if ($this.Options.SkipNotConfigured) {
            $notConfiguredValues = @(
                $notConfiguredLoc
                Get-LanguageString 'Inputs.notConfigured'
            ) | Where-Object { $_ } | Select-Object -Unique
            $basicRows = @($basicRows | Where-Object {
                if ($_.Value -in $notConfiguredValues) {
                    Write-Log "Skipping basic property $($_.Name) based on '$($_.Value)' string value" 2
                    return $false
                }
                return $true
            })
        }

        # One table instead of two: the basic rows lead, then the settings, in the
        # order a reader meets them on the portal blade. Done here rather than in
        # each output provider so every format gets the same shape, and done AFTER
        # the engine's post-steps so the scope-tag and assignment rows they append
        # to BasicInfo come along too.
        if ($this.MergeBasicInfo) {
            $settings  = @($basicRows) + @($settings)
            $basicRows = @()
        }

        return [PSCustomObject]@{
            BasicInfo                       = $basicRows
            FilteredSettings                = $settings
            DocumentName                    = $this.DocumentName
            ComplianceActions               = @($this.ComplianceActions)
            ApplicabilityRules              = @($this.ApplicabilityRules)
            Assignments                     = @($this.Assignments)
            Scripts                         = @($this.Scripts)
            CustomTables                    = @($this.CustomTables)
            DisplayProperties               = $this.DisplayProperties
            DefaultDocumentationProperties  = $this.DefaultDocumentationProperties
            ErrorText                       = $this.ErrorText
            InputType                       = $this.InputType
            UpdateFilteredObject            = $this.UpdateFilteredObject
            UnconfiguredProperties          = $this.UnconfiguredProperties
        }
    }
}
