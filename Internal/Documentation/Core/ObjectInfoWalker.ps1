# ObjectInfo JSON walker — the core of the generic Profile input provider.
#
# Ported from old Extensions/Documentation.psm1:2276 (Invoke-TranslateSection,
# ~440 LOC) plus the Invoke-VerifyCondition helper (~70 LOC) and
# Get-CultureLanguageString (~50 LOC).
#
# Drives schema-driven translation for ~140 policy types catalogued in
# Config/ObjectCategories.json. Each ObjectInfo JSON file under
# Config/ObjectInfo/<category>_<policyType>.json describes the per-property
# metadata (dataType, entityKey, nameResourceKey, child layout) and the
# walker dispatches each property to the appropriate translate primitive
# based on dataType.

# ---- Section walker ----

# Tracks the parent prop being walked so propLevel adjusts correctly when
# recursing into children. Module-scope (replaces old $script:currentParent).
$script:_currentSectionParent = $null

# Defaults: properties whose nameResourceKey shows up in this list are skipped
# entirely (purely visual elements in the old portal that don't translate
# to documentation content). Old code at Documentation.psm1:2977.
# ToDo: Review if these should be implemented or actually ignored
$script:_categoriesToIgnore = @(
    'defenderSecurityCenterContactOptionsText'
    'globalConfigurationsDescription','generalNetworkSettingsHeader'
    'firewallCreateRules','exploitGuardCFHeadingText','exploitGuardNFTitle'
    'exploitGuardEPExplainationPart1','exploitGuardEPExplainationPart2'
    'exploitGuardEPExplainationPart3','exploitGuardEPExplainationPart4'
    'defenderSecurityCenterSubHeaderText','defenderSecurityCenterITContactInformationSubHeaderText'
    'windows10EndpointProtectionDeviceGuardLearnMore'
    'win10DefaultPrivacyHeader','dfciBuiltinHeaderDescName'
)

# Resource keys that upstream renamed while the blade metadata kept referencing
# the old name. Microsoft's own portal cannot render these tooltips either, so
# there is nothing to wait for - map them to the current name.
# Confirmed 2026-08-22 by the IntuneLanuageAndObjects generator, which extracts
# the portal's ClientResources verbatim.
$script:_resourceKeyAliases = @{
    'autoInstallAndRebootAtScheduledTime'      = 'autoInstallAndRebootAtScheduledTimeOption'
    # Only connecteddevices_iosgeneral.json references this, so the iOS variant
    # is the correct target; a MacOS variant also exists upstream.
    'blockAirPrintiBeaconDiscoveryDescription' = 'blockAirPrintiBeaconDiscoveryDescriptionIOS'
}

# Resolve an ObjectInfo resource key to its display string, or $null.
#
# Keys without a namespace live under SettingDetails. Two upstream quirks are
# handled here so callers do not each reimplement them:
#
#  - Purely numeric keys are portal metadata artifacts, not string ids. The
#    AndroidDeviceOwner and AOSP PKCS files carry emptyValueResourceKey:"1"
#    verbatim from the blade metadata, which resolves to nothing and logs a
#    "Could not find string" warning on every documented policy. Skipped the way
#    the existing 'Empty' / 'LearnMore' sentinels are.
#  - Renamed keys are redirected via $script:_resourceKeyAliases.
function Get-ObjectInfoResourceString {
    param(
        [string]$Key,
        # emptyValueResourceKey values are already fully qualified upstream and
        # must not get the SettingDetails prefix.
        [switch]$NoPrefix
    )

    if ([string]::IsNullOrWhiteSpace($Key)) { return $null }
    if ($Key -match '^\d+$') { return $null }

    if ($script:_resourceKeyAliases.ContainsKey($Key)) { $Key = $script:_resourceKeyAliases[$Key] }

    $full = if ($NoPrefix -or $Key.Contains('.')) { $Key } else { "SettingDetails.$Key" }
    try { return Get-LanguageString $full }
    catch {
        Write-Log "Get-LanguageString '$full' failed: $($_.Exception.Message)" 2
        return $null
    }
}

# Walk a flat settings object through an ObjectInfo manifest file. Used by the
# AppConfig handlers to give Outlook/Edge their schema-driven rows (the old code
# called Invoke-TranslateSection directly against #AppConfig*.json). $ManifestPath
# is a full path under Config\ObjectInfo\.
function Invoke-DocAppConfigManifest {
    param($SettingsObject, [string]$ManifestPath, [DocumentationContext]$Context)
    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) { return }
    try {
        $jsonObj = [IO.File]::ReadAllText($ManifestPath) | ConvertFrom-Json
    }
    catch {
        Write-LogError "Failed to read AppConfig manifest $ManifestPath" $_.Exception
        return
    }
    if (-not $jsonObj) { return }
    $prev = $Context.CurrentObject
    $Context.CurrentObject = $SettingsObject
    try { Invoke-TranslateSection $SettingsObject $jsonObj $null }
    catch { Write-LogError "Failed to translate AppConfig manifest $(Split-Path -Leaf $ManifestPath)" $_.Exception }
    finally { $Context.CurrentObject = $prev }
}

function Invoke-TranslateSection {
    param($Obj, $SectionObject, $ObjInfo, $Parent = $null)

    $ctx = Get-CurrentDocumentationContext

    # Reset/adjust propLevel based on whether we're a new walk or recursing
    if ($null -eq $Parent -or $ctx.PropLevel -lt 0) {
        $ctx.PropLevel = 0
    }
    elseif ($Parent -ne $script:_currentSectionParent) {
        $ctx.PropLevel++
    }

    foreach ($prop in $SectionObject) {
        $value        = $null
        $valueSet     = $false
        $useParentProp = $false
        $payloadFile  = $false
        $skipChildren = $false

        if (-not (Invoke-VerifyCondition $Obj $prop $ObjInfo)) {
            Write-LogDebug "Condition returned false: $($prop.Condition | ConvertTo-Json -Depth 50 -Compress)"
            continue
        }

        $Obj = Get-CustomPropertyObject $Obj $prop
        $rawValue = $Obj."$($prop.entityKey)"

        # ---- Section/category headers (dataType 8) ----
        if ($prop.dataType -eq 8) {
            if ($prop.nameResourceKey -eq 'LearnMore') { continue }
            elseif ($prop.nameResourceKey -eq 'Empty') { $ctx.CurrentSubCategory = $null }
            elseif ($prop.nameResourceKey -in $script:_categoriesToIgnore) { continue }
            elseif ($prop.nameResourceKey) {
                $key = if ($prop.nameResourceKey.Contains('.')) { $prop.nameResourceKey } else { "SettingDetails.$($prop.nameResourceKey)" }
                $tmpStr = Get-LanguageString $key
                if ($tmpStr -and $tmpStr.Length -lt 75) {
                    $ctx.CurrentSubCategory = $tmpStr
                }
                elseif ($tmpStr) {
                    Write-LogDebug "SubCategory ignored based on length: $tmpStr"
                }
            }
            $ctx.PropLevel = -1
            Invoke-ChildSections $Obj $prop
            # A header without child sections leaves the -1 reset sentinel dangling;
            # the next property's childSettings recursion would then reset to level 0
            # instead of indenting one level under its parent row. Normalize here so
            # only the header's own children get the flat-level reset.
            if ($ctx.PropLevel -lt 0) { $ctx.PropLevel = 0 }
            continue
        }

        # ---- Complex options (dataType 5) ----
        if ($prop.dataType -eq 5) {
            if ($prop.enabled -eq $false -and $ObjInfo.ShowDisabled -ne $true) { continue }
            if (-not $prop.EntityKey -and $prop.nameResourceKey) {
                $ctx.PropLevel = -1
                $key = if ($prop.nameResourceKey.Contains('.')) { $prop.nameResourceKey } else { "SettingDetails.$($prop.nameResourceKey)" }
                $ctx.CurrentSubCategory = Get-LanguageString $key
            }
            else {
                $ctx.PropLevel--
            }
            foreach ($tmpObj in $Obj) {
                Invoke-TranslateSection $tmpObj $prop.complexOptions $ObjInfo -Parent $prop
            }
            continue
        }

        # ---- Complex option based on sub-property (dataType 6) ----
        if ($prop.dataType -eq 6) {
            if ($prop.enabled -eq $false -and $ObjInfo.ShowDisabled -ne $true) { continue }
            $ctx.PropLevel--
            $propObj = $null
            if ($prop.entityKey) { $propObj = $Obj.PSObject.Properties | Where-Object Name -EQ $prop.entityKey }
            $iter = if ($null -ne $propObj) { $rawValue } else { $Obj }
            foreach ($tmpObj in $iter) {
                Invoke-TranslateSection $tmpObj $prop.complexOptions $ObjInfo -Parent $prop
            }
            continue
        }

        # ---- Skip-but-add-children label (dataType 9) ----
        if ($prop.dataType -eq 9) {
            $ctx.PropLevel--
            Invoke-ChildSections $Obj $prop
            continue
        }

        # ---- Information box: ignore (dataType 10) ----
        if ($prop.dataType -eq 10) { continue }

        # ---- Static-string label (dataType 101): language-id lookup ----
        if ($prop.dataType -eq 101) {
            if ($prop.value) {
                $value = Get-LanguageString $prop.value
                Add-PropertyInfo $prop $value $rawValue $rawValue
            }
            continue
        }

        # ---- Static value (dataType 107) ----
        if ($prop.dataType -eq 107) {
            if ($prop.value) {
                Add-PropertyInfo $prop $prop.value $prop.value $prop.value
            }
            continue
        }

        # ---- Generic property path (dataType varies, requires entityKey) ----
        if (-not [string]::IsNullOrEmpty($prop.entityKey)) {
            $valueSet = ($null -ne $rawValue)

            # Determine propValue (with defaults fallback). Old engine gates the
            # unconfigured/default substitutions on $global:chk* UI checkboxes
            # which default to UNCHECKED — meaning when a property is null on
            # the input, the walker just uses null (and most translate primitives
            # then either skip the row or emit "Not configured" via their own
            # logic). My port honors that by gating on $ctx.Options.SetUnconfigured
            # Value / SetDefaultValue (also default false).
            $propValue = if ($null -ne $rawValue) { $rawValue }
                         elseif (-not [string]::IsNullOrEmpty($prop.unconfiguredValue) -and $ctx.Options.SetUnconfiguredValue) {
                             Add-NotConfiguredProperty $prop
                             $prop.unconfiguredValue
                         }
                         elseif (-not [string]::IsNullOrEmpty($prop.defaultValue) -and $ctx.Options.SetDefaultValue) {
                             $prop.defaultValue
                         }
                         elseif (-not [string]::IsNullOrEmpty($prop.emptyValueResourceKey) -and $ctx.Options.SetDefaultValue) {
                             Get-ObjectInfoResourceString $prop.emptyValueResourceKey -NoPrefix
                         }
                         else { $rawValue }

            $addPropertyInfo = $true
            $customValue = Get-CustomProfileValue $Obj $prop

            if ($customValue -is [bool] -and $customValue -eq $false) {
                continue
            }
            elseif (-not $customValue) {

                # Linked certificate (dataType 4): live Graph navigationLink
                # Stub offline — uses #CustomRef_ embedded data when present
                if ($prop.dataType -eq 4) {
                    $useParentProp = $true
                    $cert = $null
                    if (-not $ctx.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
                        try {
                            $url = $ctx.CurrentObject."$($prop.entityKey)@odata.navigationLink"
                            if ($url) {
                                # Most policies advertise the navigationLink even when no
                                # certificate is associated; the GET 404s. Cache the
                                # outcome on $ctx so the second walk of the same policy
                                # (the schema lists the same entityKey twice for the
                                # SCEP+PKCS+derived flavours) and any later policies in
                                # the bulk run skip a known-empty fetch.
                                if (-not $ctx.PSObject.Properties['_LinkedCertCache']) {
                                    $ctx | Add-Member -MemberType NoteProperty -Name '_LinkedCertCache' -Value (@{}) -Force
                                }
                                if ($ctx._LinkedCertCache.ContainsKey($url)) {
                                    $cert = $ctx._LinkedCertCache[$url]
                                }
                                else {
                                    try {
                                        $cert = Invoke-MSGraphAPI -Url $url -ODataMetadata 'minimal' -NoError
                                    } catch { $cert = $null }
                                    $ctx._LinkedCertCache[$url] = $cert
                                }
                            }
                        } catch { }
                    }
                    if ($cert) {
                        if ($cert.value -is [object[]]) {
                            $certs = @($cert.value | ForEach-Object { $_.displayName }) | Where-Object { $_ }
                            if ($certs.Count -gt 0) { $value = $certs -join $ctx.ObjectSeparator }
                        }
                        elseif ($cert.displayName) {
                            $value = $cert.displayName
                        }
                        $rawValue = $value
                    }
                    elseif ($ctx.CurrentObject.'@ObjectFromFile' -eq $true -or $ctx.SourceTenantUnavailable) {
                        $refKey = "#CustomRef_$($prop.entityKey)"
                        if ($ctx.CurrentObject.$refKey) {
                            $sep = $ctx.CurrentObject.$refKey.IndexOf('|:|')
                            $value = if ($sep -gt -1) { $ctx.CurrentObject.$refKey.Substring(0, $sep) } else { $ctx.CurrentObject.$refKey }
                        }
                        $rawValue = $value
                    }
                }
                # Multi-option based on boolean value where the property name IS the key (dataType 200)
                elseif ($prop.dataType -eq 200) {
                    $value = Get-LanguageString $prop.entityKey
                }
                # Property missing on the object (and not "allowMissing")
                elseif (-not $prop.allowMissing -and
                        $prop.entityKey -ne '.' -and
                        -not ($Obj.PSObject.Properties | Where-Object Name -EQ $prop.entityKey) -and
                        -not ($Obj.PSObject.Properties | Where-Object Name -EQ "$($prop.entityKey)@odata.navigationLink")) {
                    if ($prop.enabled -ne $false) {
                        Write-Log "Property with EntityKey $($prop.entityKey) is missing. Property will not be added!" 2
                    }
                    else {
                        Write-LogDebug "Disabled property with EntityKey $($prop.entityKey) is missing. Property will not be added!"
                    }
                    continue
                }
                else {
                    # NOTE: `continue` inside `switch` only goes to the next
                    # switch case match in PowerShell — it does NOT skip code
                    # after the switch. Cases that handle their own row emission
                    # (Option / Table) must set $addPropertyInfo = $false so the
                    # Add-PropertyInfo call below is skipped. (Earlier port used
                    # `continue` here and produced duplicate rows.)
                    switch ([int]$prop.dataType) {
                        0  { $value = Invoke-TranslateBoolean $Obj $prop }
                        1  {
                            # Base64 e.g. certificate data
                            $value = if ($prop.filenameEntityKey -and $Obj."$($prop.filenameEntityKey)") {
                                $Obj."$($prop.filenameEntityKey)"
                            } else {
                                $v = $Obj."$($prop.EntityKey)"
                                if ($v) { try { [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($v)) } catch { $v } } else { $v }
                            }
                        }
                        2  {
                            # Multiline string (often a base64-wrapped XML payload file)
                            if ($prop.filenameEntityKey -and $Obj."$($prop.filenameEntityKey)") {
                                $value = $Obj."$($prop.filenameEntityKey)"
                                $payloadFile = $true
                            }
                            else {
                                $v = $Obj."$($prop.EntityKey)"
                                $value = if ($v) { try { [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($v)) } catch { $v } } else { $v }
                            }
                        }
                        3  {
                            # Image — placeholder label; raw image data dropped (no consumer yet).
                            $value = if ($propValue) { 'Image file' } else { $null }
                        }
                        7  { $value = $propValue }   # omaSettingDateTime — formatting deferred
                        11 { }                       # App picker — value left $null
                        12 {
                            # Multiline string / array
                            if (($propValue | Measure-Object).Count -gt 0) {
                                $value = $propValue -join $ctx.ObjectSeparator
                            }
                        }
                        13 { $value = Invoke-TranslateMultiOption $Obj $prop }
                        14 { $value = $propValue }   # Int32
                        15 { $value = $propValue }   # Int64
                        16 { Invoke-TranslateOption $Obj $prop | Out-Null; $addPropertyInfo = $false; $skipChildren = $true }
                        19 { Invoke-TranslateOption $Obj $prop | Out-Null; $addPropertyInfo = $false; $skipChildren = $true }
                        20 { $value = $propValue }   # String
                        21 { Invoke-TranslateTable $Obj $prop; $addPropertyInfo = $false; $skipChildren = $true }
                        22 {
                            # Scale value e.g. "4 Years"
                            $value = $propValue
                            $scaleEntityKey = if ($Obj."$($prop.scaleEntityKey)") { $Obj."$($prop.scaleEntityKey)" } else { $prop.defaultScale }
                            if ($scaleEntityKey) {
                                $scaleOption = $prop.scaleOptions | Where-Object value -EQ $scaleEntityKey | Select-Object -First 1
                                if ($scaleOption.nameResourceKey) {
                                    $value = '{0} {1}' -f $propValue, (Get-LanguageString "SettingDetails.$($scaleOption.nameResourceKey)")
                                }
                            }
                        }
                        100 { $value = Invoke-TranslateDuration $Obj $prop }
                        102 {
                            $culture = if ($propValue) { $propValue } else { $prop.unconfiguredValue }
                            $value = Get-CultureLanguageString $culture
                        }
                        103 {
                            # Boolean action but hide children on false
                            $value = Invoke-TranslateBoolean $Obj $prop
                            $skipChildren = ($propValue -eq $false)
                        }
                        104 {
                            $value = Invoke-TranslateMultiOptionBoolean $Obj $prop
                            $skipChildren = ($propValue -eq $false)
                        }
                        105 {
                            $value = Invoke-TranslateMultiOptionBoolean $Obj $prop $false
                            $skipChildren = ($propValue -eq $false)
                        }
                        106 {
                            # Array of cultures
                            $tmp = @()
                            foreach ($lng in $propValue) { $tmp += Get-CultureLanguageString $lng }
                            $value = $tmp -join $ctx.ObjectSeparator
                        }
                        108 {
                            # String with format
                            $value = $propValue
                            if ($prop.formatStringKey) {
                                $fmt = Get-LanguageString $prop.formatStringKey
                                if ($fmt) { $value = $fmt -f $propValue }
                            }
                        }
                        default {
                            $nameForLog = if ($prop.nameResourceKey) { Get-LanguageString "SettingDetails.$($prop.nameResourceKey)" } else { '' }
                            Write-Log "Unsupported property '$nameForLog' ($($prop.nameResourceKey)) for object property $($prop.entityKey). Type: $($prop.dataType)" 2
                            $value = $propValue
                        }
                    }
                }
            }
            else {
                $value           = $customValue.Value
                $rawValue        = $customValue.RawValue
                $valueSet        = ($null -ne $rawValue)
                $addPropertyInfo = $customValue.AddPropertyInfo
            }

            if ($addPropertyInfo) {
                $propForAdd = if ($useParentProp -and $Parent) { $Parent } else { $prop }
                Add-PropertyInfo $propForAdd $value $rawValue

                if ($payloadFile -and $Obj.payload) {
                    $tmpProp = [PSCustomObject]@{
                        nameResourceKey        = 'uploadResult'
                        descriptionResourceKey = ''
                        entityKey              = 'payloadData'
                        dataType               = 20
                        booleanActions         = 0
                        category               = $prop.Category
                    }
                    $payloadValue = try { [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($Obj.payload)) } catch { $Obj.payload }
                    Add-PropertyInfo $tmpProp $payloadValue $Obj.payload
                }
            }
        }
        else {
            Write-Log "No property entity key: $($prop.dataType) ($($prop.nameResourceKey))" 2
        }

        if ($valueSet -and -not $skipChildren) {
            Invoke-ChildSections $Obj $prop
        }
    }

    if ($null -ne $Parent -and $Parent -ne $script:_currentSectionParent -and $ctx.PropLevel -gt 0) {
        $ctx.PropLevel--
    }
}

# ---- Condition verifier (dataType-independent property gate) ----
function Invoke-VerifyCondition {
    param($Obj, $Prop, $ObjInfo)

    if (-not $Prop.Condition -or ($Prop.Condition.Expressions | Measure-Object).Count -eq 0) { return $true }

    $type = if ($Prop.Condition.type -eq 'and') { 'and' } else { 'or' }
    $defaultReturn = ($type -eq 'and')

    foreach ($expression in $Prop.Condition.Expressions) {
        if (-not $expression.property) { continue }
        $tmpProp = $Obj.PSObject.Properties | Where-Object Name -EQ $expression.property
        if (-not $tmpProp) {
            if ($expression.ignoreMissing -eq $true) { continue }
            return $false
        }

        $tmpRet = switch ($expression.operator) {
            'null'    { $null -eq $tmpProp.Value }
            'ne'      { $Obj."$($expression.property)" -ne $expression.value }
            'gt'      { $Obj."$($expression.property)" -gt $expression.value }
            'ge'      { $Obj."$($expression.property)" -ge $expression.value }
            'lt'      { $Obj."$($expression.property)" -lt $expression.value }
            'le'      { $Obj."$($expression.property)" -le $expression.value }
            'like'    { $Obj."$($expression.property)" -like    $expression.value }
            'notlike' { $Obj."$($expression.property)" -notlike $expression.value }
            default {
                if ($null -eq $expression.value) {
                    $null -ne $tmpProp.Value
                }
                else {
                    $Obj."$($expression.property)" -eq $expression.value
                }
            }
        }

        if ($tmpRet -eq $true  -and $type -eq 'or')  { return $true }
        if ($tmpRet -eq $false -and $type -eq 'and') { return $false }
    }
    return $defaultReturn
}

# ---- Culture-code -> language name ----
# Used by dataType 102 (Culture name) and 106 (Array of languages).
# Looks up Languages.<culture> in the loaded Strings-en.json; falls back to
# the OS culture's EnglishName.
function Get-CultureLanguageString {
    param($Culture)

    if (-not $Culture) { return $null }
    try {
        if ($Culture -eq 'os-default') { return Get-LanguageString 'Autopilot.OOBE.useOSDefaultLanguage' }
        if ($Culture -eq 'user-select') { return Get-LanguageString 'Autopilot.OOBE.userSelect' }

        # Force language strings to load by calling Get-LanguageString once
        Get-LanguageString $null | Out-Null

        $cache = Get-CacheObject "LanguageStrings_$($Culture)"
        if (-not $cache) { $cache = Get-CacheObject 'LanguageStrings_en' }
        if ($cache.Languages.$Culture) { return $cache.Languages.$Culture }

        $parts = $Culture.Split('-')
        if ($parts.Length -eq 3) {
            $tri = "$($parts[0])-$($parts[1])"
            if ($cache.Languages.$tri) { return $cache.Languages.$tri }
        }
        if ($parts.Length -gt 1 -and $cache.Languages."$($parts[0])") {
            return $cache.Languages."$($parts[0])"
        }

        Write-Log "Translated language for $Culture not found" 2
        return ([cultureinfo]$Culture).EnglishName
    }
    catch { return $null }
}
