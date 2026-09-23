# Generic Profile input provider.
#
# Ported from old Extensions/Documentation.psm1:2057 (Invoke-TranslateProfile-
# Object). Claims any @odata.type catalogued in Config/ObjectCategories.json
# that isn't already handled by a more-specific input provider or handler
# (first-match-wins dispatch in [DocumentationRegistry]).
#
# For each catalogued type:
#   1. Emit BasicInfo via Add-BasicDefaultValues (which itself reads
#      ObjectCategories.json for Platform-supported + Profile-type rows)
#   2. Emit Created/Modified/Version via Add-BasicAdditionalValues
#   3. Find category files: either the explicit Categories list, or every
#      file matching *_<policyType>.json under Config/ObjectInfo/
#   4. Load each as JSON, dispatch to Invoke-TranslateSection walker
#
# The walker handles all the per-prop dataType dispatching to translate
# primitives (Boolean/Option/MultiOption/Table/Duration etc.).

function Invoke-InitializeProfileInput {
    Add-DocumentationInputProvider ([PSCustomObject]@{
        Name      = 'Profile'
        Order     = 60
        Match     = {
            param($PolicyObject)
            $odata = $PolicyObject.JsonObject.'@odata.type'
            if (-not $odata) { return $false }
            if (-not (Get-Command Get-PolicyObjectCategoryInfo -ErrorAction SilentlyContinue)) { return $false }
            $info = Get-PolicyObjectCategoryInfo $odata
            return ($null -ne $info -and $null -ne $info.PolicyType)
        }
        Translate = { param($PolicyObject, $Context) Invoke-TranslateProfilePolicyObject $PolicyObject $Context }
    })
}

function Invoke-TranslateProfilePolicyObject {
    param($PolicyObject, [DocumentationContext]$Context)

    $obj = $PolicyObject.JsonObject

    $objInfo = Get-PolicyObjectCategoryInfo $obj.'@odata.type'
    if (-not $objInfo) { return }

    # Header rows
    Add-BasicDefaultValues $PolicyObject
    Add-BasicAdditionalValues $PolicyObject
    # Pin '@ObjectFromFile' so the walker's source-unavailable branches (linked
    # certificates etc.) treat the input as a file-based object even when
    # SourceTenantUnavailable isn't set.
    if (-not $obj.PSObject.Properties['@ObjectFromFile']) {
        $obj | Add-Member -MemberType NoteProperty -Name '@ObjectFromFile' -Value $true -Force
    }
    $Context.CurrentObject = $obj
    Initialize-DocumentationObjectInfoObject $obj

    # Resolve the list of ObjectInfo JSON files to walk for this PolicyType
    $objectInfoDir = Join-Path $script:AppRootFolder 'Config\ObjectInfo'
    $allFiles = @()

    if ($objInfo.Categories -and $objInfo.Categories.Count -gt 0) {
        foreach ($cat in $objInfo.Categories) {
            $path = Join-Path $objectInfoDir "$($cat.ToLower())_$($objInfo.PolicyType.ToLower()).json"
            if (Test-Path -LiteralPath $path) {
                $allFiles += [IO.FileInfo]$path
            }
            else {
                Write-Log "ObjectInfo file '$path' not found for $($objInfo.PolicyType)" 2
            }
        }
    }
    else {
        # Single-file path — find any *_<policyType>.json
        $pattern = "*_$($objInfo.PolicyType.ToLower()).json"
        if (Test-Path $objectInfoDir) {
            $files = Get-ChildItem -Path $objectInfoDir -Filter $pattern -ErrorAction SilentlyContinue
            if (-not $files) {
                Write-Log "No ObjectInfo files matching '$pattern' for $($objInfo.PolicyType)" 2
            }
            foreach ($f in $files) { $allFiles += $f }
        }
    }

    foreach ($fi in $allFiles) {
        try {
            $categoryObj = [IO.File]::ReadAllText($fi.FullName) | ConvertFrom-Json
            $Context.CurrentSubCategory = ''
            # Per-file custom handlers override the generic walker (old code:
            # Invoke-CDDocumentTranslateSectionFile, called via docProvider.
            # TranslateSectionFile hook from Invoke-TranslateProfileObject).
            # Returns $true if the custom handler emitted rows; $false to fall
            # through to Invoke-TranslateSection.
            if (Invoke-DocCustomSectionFileTranslator -Obj $obj -FileInfo $fi -CategoryObj $categoryObj -ObjInfo $objInfo) {
                continue
            }
            # Each ObjectInfo file wraps its section array under a key matching the file's basename
            $sections = $categoryObj."$($fi.BaseName)"
            if ($sections) {
                Invoke-TranslateSection $obj $sections $objInfo
            }
        }
        catch {
            Write-LogError "Failed to translate ObjectInfo file $($fi.Name)" $_.Exception
        }
    }
}

# Custom per-(odata.type, fileBaseName) section-file translators. Mirrors old
# Extensions/DocumentationCustom.psm1's Invoke-CDDocumentTranslateSectionFile.
# Each block returns $true if it emitted rows and the generic walker should be
# skipped for this file, $false to fall through.
function Invoke-DocCustomSectionFileTranslator {
    param($Obj, [IO.FileInfo]$FileInfo, $CategoryObj, $ObjInfo)

    # --- Compliance: Custom Compliance category (Windows 10) ----------------
    # Generic walker can't emit useful rows for `customcompliance_compliancewindows10`
    # because the manifest's dataType=25 ("home screen", unused) is the child
    # carrying the actual content, and the rules live on a separate
    # $obj.deviceCompliancePolicyScript navigation property rather than on the
    # boolean entityKey the parent points to. Three rows are emitted by hand:
    #   - Custom compliance (Require / Not configured)
    #   - Select your discovery script (resolved displayName)
    #   - Upload and validate the JSON file (base64-decoded rulesContent)
    if ($Obj.'@odata.type' -eq '#microsoft.graph.windows10CompliancePolicy' -and
        $FileInfo.BaseName -eq 'customcompliance_compliancewindows10') {

        $category = Get-PolicyObjectCategoryString ($CategoryObj."$($FileInfo.BaseName)".category)

        if ($null -eq $Obj.deviceCompliancePolicyScript) {
            $propValue = Get-LanguageString 'BooleanActions.notConfigured'
            $rawValue  = 'notConfigured'
        } else {
            $propValue = Get-LanguageString 'BooleanActions.require'
            $rawValue  = 'require'
        }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name        = Get-LanguageString 'SettingDetails.adminConfiguredComplianceSettingName'
            Value       = $propValue
            EntityKey   = 'deviceCompliancePolicyScript'
            RawValue    = $rawValue
            Category    = $category
            SubCategory = $null
        })

        if ($Obj.deviceCompliancePolicyScript) {
            # Resolve script displayName via shared cache. Offline runs / cache
            # miss fall back to the script id so the row isn't blank.
            $scriptId = [string]$Obj.deviceCompliancePolicyScript.deviceComplianceScriptId
            $scriptName = $scriptId
            if (-not [string]::IsNullOrEmpty($scriptId)) {
                $cache = Get-CacheObject 'DocAllCustomCompliancePolicies'
                # Custom compliance scripts are authored in the source tenant (not
                # generic schema), so this is gated on source-tenant availability.
                if (-not $cache -and -not (Get-CurrentDocumentationContext).SourceTenantUnavailable -and (Test-DocumentationGraphAvailable)) {
                    try {
                        $cache = @((Invoke-MSGraphAPI -Url "/deviceManagement/deviceComplianceScripts?`$select=displayName,id" -ODataMetadata 'minimal').value)
                        Set-CacheObject 'DocAllCustomCompliancePolicies' $cache
                    } catch {
                        Write-Log "Failed to fetch deviceComplianceScripts for resolution: $($_.Exception.Message)" 2
                    }
                }
                if ($cache) {
                    $match = $cache | Where-Object Id -EQ $scriptId | Select-Object -First 1
                    if ($match.displayName) { $scriptName = $match.displayName }
                }
            }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name        = Get-LanguageString 'CustomCompliance.FilePicker.scriptFileLabel'
                Value       = $scriptName
                EntityKey   = 'deviceComplianceScriptName'
                Category    = $category
                SubCategory = $null
            })

            if ($Obj.deviceCompliancePolicyScript.rulesContent) {
                $rules = try {
                    [System.Text.Encoding]::UTF8.GetString(
                        [System.Convert]::FromBase64String($Obj.deviceCompliancePolicyScript.rulesContent))
                } catch {
                    [string]$Obj.deviceCompliancePolicyScript.rulesContent
                }
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name        = Get-LanguageString 'CustomCompliance.UploadFile.jsonFileLabel'
                    Value       = $rules
                    EntityKey   = 'jsonFileContent'
                    Category    = $category
                    SubCategory = $null
                })
            }
        }
        return $true
    }

    return $false
}

Invoke-InitializeProfileInput
