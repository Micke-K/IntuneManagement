# Manifest input provider.
#
# Bridges the old Documentation\ObjectInfo\ "manifest" JSON files (a flat
# array of property descriptors describing how to translate a policy
# object). Different from the Profile provider's category files, which:
#   - Live alongside as <category>_<policyType>.json
#   - Wrap their section array under a key matching the file basename
#   - Are looked up via ObjectCategories.json (Get-PolicyObjectCategoryInfo)
#
# Manifest files instead are:
#   - Flat top-level arrays
#   - Named directly by @odata.type:  <odata.type>.json
#         e.g. #microsoft.graph.hardwareConfiguration.json
#   - Or named by PolicyType Id:      #<typeId>.json
#         e.g. #Applications.json, #Autopilot.json
#
# These files exist for ~21 PolicyTypes that aren't catalogued in
# ObjectCategories.json (Applications, AppProtection, BIOS hardware
# configs, EnrollmentLimit/Notification/StatusPage, WindowsUpdate
# profiles, MacScripts, PowerShell/HealthScripts, etc.) so without this
# provider every one of those types renders an empty HTML stub.
#
# Match order: this file's basename ("Manifest") sorts before "Profile"
# so it gets first shot at types that aren't already claimed by a
# DocHandler or one of the specific schema providers (ADMX/Compliance V2
# /Intent/SettingsCatalog).
#
# Old code reference: Extensions/Documentation.psm1:268-284 (the dispatcher
# branches that test File.Exists on the two filename forms) +
# Extensions/Documentation.psm1:4016 (Invoke-TranslateCustomProfileObject —
# the helper that loaded a flat array and called Invoke-TranslateSection).

function Invoke-InitializeManifestInput {
    Add-DocumentationInputProvider ([PSCustomObject]@{
        Name      = 'Manifest'
        Order     = 10
        Match     = {
            param($PolicyObject)
            $path = Get-DocumentationManifestPath $PolicyObject
            return [bool]$path
        }
        Translate = { param($PolicyObject, $Context) Invoke-TranslateManifestPolicyObject $PolicyObject $Context }
    })
}

# Looks up a manifest file for $PolicyObject. Returns the resolved path or
# $null. Tries @odata.type first, then PolicyType.Id with '#' prefix.
function Get-DocumentationManifestPath {
    param($PolicyObject)

    $obj = $PolicyObject.JsonObject

    $dir = Join-Path $script:AppRootFolder 'Config\ObjectInfo'

    $odata = [string]$obj.'@odata.type'
    if ($odata) {
        $path = Join-Path $dir "$odata.json"
        if (Test-Path -LiteralPath $path -PathType Leaf) { 
            Write-Log "Manifest input provider: Found file based on OData type: $path"
            return $path
        }
    }

    $typeId = $null
    if ($PolicyObject.PSObject.Properties['PolicyType'] -and $PolicyObject.PolicyType) {
        $typeId = [string]$PolicyObject.PolicyType.Id
    }
    if ($typeId) {
        $path = Join-Path $dir "#$typeId.json"
        if (Test-Path -LiteralPath $path -PathType Leaf) { 
            Write-Log "Manifest input provider: Found file based on PolicyType.Id: $path"
            return $path
        }
    }

    return $null
}

function Invoke-TranslateManifestPolicyObject {
    param($PolicyObject, [DocumentationContext]$Context)

    $obj = $PolicyObject.JsonObject

    $path = Get-DocumentationManifestPath $PolicyObject
    if (-not $path) { return }

    # Header rows (matches Profile provider so output looks identical for
    # both code paths — manifest vs category-driven).
    Add-BasicDefaultValues $PolicyObject

    # Add app name for apps
    $appType = Get-GraphApplicationType $PolicyObject
    if($appType)
    {
        $appTypeName = Get-LanguageString "AppType.$($appType.LanguageId)"
        if($appTypeName) { Add-BasicPropertyValue (Get-LanguageString "Inputs.installationSourceLabel") $appTypeName }
    }

    $Context.CurrentObject = $obj
    Initialize-DocumentationObjectInfoObject $obj

    try {
        $manifest = [IO.File]::ReadAllText($path) | ConvertFrom-Json
    } catch {
        Write-LogError "Failed to read manifest $path" $_.Exception
        return
    }

    if (-not $manifest) { return }

    # Manifest is a flat array (no per-file wrapper key), so pass it directly
    # to the walker. No $ObjInfo - the manifest doesn't come from
    # ObjectCategories.json.
    try {
        $Context.CurrentSubCategory = ''
        Invoke-TranslateSection $obj $manifest $null
    } catch {
        Write-LogError "Failed to translate manifest $(Split-Path -Leaf $path)" $_.Exception
    }

    Add-BasicAdditionalValues $PolicyObject

}

Invoke-InitializeManifestInput
