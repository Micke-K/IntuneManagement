function Get-LanguageString
{
    param($String, $DefaultValue = $null, [switch]$IgnoreMissing)

    $lng = ?? $script:CurrentLanguage "en"

    $languageStrings = Get-CacheObject "LanguageStrings_$lng"

    if(-not $languageStrings)
    {        
        try {
            # ReadAllText, not Get-Content: it defaults to UTF8 with BOM detection on
            # both PS5.1 and PS7, so there is no -Encoding flag to forget. Bare
            # Get-Content reads these as the ANSI codepage on 5.1 and mangles every
            # non-Latin string. It is also ~3x faster on these 1-2 MB files.
            $languageStrings = [IO.File]::ReadAllText(([IO.Path]::Combine($script:AppRootFolder, "Config", "LanguageStrings", "Strings-$($lng).json"))) | ConvertFrom-Json
            Set-CacheObject "LanguageStrings_$lng" $languageStrings
        }
        catch {
            Write-LogError "Failed to load language string for $lng" $_.Exception
        }
    }

    if(!$String) { return }

    $arrParts = $String.Split('.')

    if([String]::IsNullOrEmpty($arrParts[-1])) { return }

    $languageStringObject = $languageStrings
    foreach($part in $arrParts.Split('.'))
    {
        if($languageStringObject.$part -or $part -eq "Empty")
        {
            $languageStringObject = $languageStringObject.$part
        }
        else
        {
            if($part -and $IgnoreMissing -ne $true)
            {
                # A segment starting with an uppercase letter is expected noise, not
                # a defect. The language pipeline forbids sibling keys that differ
                # only by case and resolves the clash by deleting the uppercase-
                # leading one, subtree included. The lookup above is case-insensitive,
                # so any uppercase key whose lowercase twin survived has already
                # matched - what reaches here is a key Microsoft declares in setting
                # metadata but never publishes in a resource file. Nothing we or the
                # language project can fix, so keep it out of the normal log.
                #
                # -cmatch, not -match: PowerShell's default comparison is
                # case-INsensitive, so -match '^[A-Z]' would silence every miss.
                if($part -cmatch '^[A-Z]')
                {
                    Write-LogDebug "Could not find string $String. Part '$part' was not found"
                }
                else
                {
                    Write-Log "Could not find string $String. Part '$part' was not found"
                }
            }
            return $DefaultValue
        }
    }

    # A key can land on a CONTAINER rather than a leaf string - the portal ships
    # both 'ScheduledAction.notification' (the action label) and a
    # 'ScheduledAction.Notification' object of email-picker sub-strings, and the
    # lookup above is case-insensitive. Calling .Trim() on that object threw and
    # took the whole policy's documentation down with it. Callers that build a
    # key from Graph data (an enum member, an @odata.type) cannot know in advance
    # which they will hit, so treat a container as "not found".
    if($languageStringObject -isnot [String])
    {
        if($IgnoreMissing -ne $true)
        {
            Write-Log "String $String resolves to a container, not a value. Using the default" 2
        }
        return $DefaultValue
    }

    return $languageStringObject.Trim("`n")
}

# Renamed from Get-TranslationFiles
function Get-PolicyObjectCategoryInfo
{
    param($ODataType)

    $policyObjectCategories = Get-CacheObject "PolicyObjectCategories"

    if(-not $policyObjectCategories)
    {
        $policyObjectCategories = [IO.File]::ReadAllText(([IO.Path]::Combine($script:AppRootFolder, "Config", "ObjectCategories.json"))) | ConvertFrom-Json
        Set-CacheObject "PolicyObjectCategories" $policyObjectCategories
    }

    return $policyObjectCategories | Where-Object ObjectType -eq $ODataType
}

function Get-PolicyObjectCategoryString
{
    param($CategoryId)

    if($CategoryId -is [String])
    {
        return Get-LanguageString (?: $CategoryId.Contains(".") $CategoryId "Category.$($CategoryId)")
    }

    $policyObjectCategoryIds = Get-CacheObject "PolicyObjectCategoryIds"

    if(-not $policyObjectCategoryIds)
    {
        $policyObjectCategoryIds = [IO.File]::ReadAllText(([IO.Path]::Combine($script:AppRootFolder, "Config", "CategoryId.json"))) | ConvertFrom-Json
        # Was never cached, so the file was re-read on every category lookup - once
        # per documented row. The other three Config readers here all cache.
        Set-CacheObject "PolicyObjectCategoryIds" $policyObjectCategoryIds
    }
    
    Get-LanguageString "Category.$($policyObjectCategoryIds."$($CategoryId)")"
}

function Get-ApplicationType
{
    param($PolicyObject)

    $appTypes = Get-CacheObject "ApplicationTypes"

    if(-not $appTypes)
    {
        $fi = [IO.FileInfo]([IO.Path]::Combine($script:AppRootFolder, "Config", "AppTypes.json"))
        if(!$fi.Exists)
        {
            return $false
        }
        $appTypes = [IO.File]::ReadAllText($fi.FullName) | ConvertFrom-Json
        Set-CacheObject "ApplicationTypes" $appTypes
    }

    foreach($appType in ($appTypes | Where-Object ODataType -eq $PolicyObject.JsonObject.'@OData.Type'))
    {
        if($appType.Condition)
        {
            if($PolicyObject.JsonObject."$($appType.Condition.Property)" -eq $appType.Condition.Value)
            {
                return $appType
            }
        }
        else
        {
            return $appType
        }
    }
}

function Get-TemplatePolicyTypeName
{
    param($ODataType, $Default = $null)

    $categoryObject =  Get-PolicyObjectCategoryInfo $ODataType

    if($null -eq $categoryObject) { return $Default }

    $policyName = Get-LanguageString "PolicyType.$($categoryObject.PolicyTypeLanguageId)"

    if($policyName) { return $policyName }

    return $Default
}

# Graph enum members and the portal's language ids do not always agree for the
# same platform. Where they differ, the portal ships the string under ITS id and
# will never publish one under the Graph name, so the mapping belongs here rather
# than in a request to the IntuneLanuageAndObjects pipeline.
#
#   aosp   deviceManagementConfigurationPlatforms member (Settings Catalog).
#          The portal files it as androidAOSP - "Android (AOSP)" - and uses the
#          same wording elsewhere (PolicyType.aospDeviceOwnerCompliancePolicy is
#          "Android (AOSP) compliance policy").
#
# Only add an entry when the target key genuinely describes the SAME platform.
# Members with no equivalent string at all (windowsPhone81, the
# *MobileApplicationManagement pseudo-platforms, none / unknownFutureValue) are
# left to render blank and are reported by Tools/Audit-PolicyPlatforms.ps1.
$script:PlatformLanguageIdAliases = @{
    'aosp' = 'androidAOSP'
}

# Platform label for an @odata.type that Config/ObjectCategories.json has no row
# for. Consulted ONLY as a fallback, so a row shipped by the language pipeline
# later automatically wins and this table quietly becomes redundant.
#
# The app / app-protection entries use the generic AppResources.AppTypePlatform
# namespace rather than Platform.*, on purpose: Platform.* renders enrollment
# flavours ("Android device administrator", "Android Enterprise"), which is the
# wording the portal uses on the CREATE blade. When LISTING these policies the
# portal says plain "Android" - and a list is what this feeds.
$script:PolicyPlatformOverrides = @{
    # Apps. Get-GraphApplicationPlatform matches a platform token inside the
    # @odata.type; these two carry none.
    '#microsoft.graph.officeSuiteApp'                        = 'AppResources.AppTypePlatform.windows'
    '#microsoft.graph.microsoftStoreForBusinessApp'          = 'AppResources.AppTypePlatform.windows'

    # App configuration for managed devices. androidManagedStoreAppConfiguration
    # is deliberately absent - it HAS an ObjectCategories row and already
    # resolves (as "Android Enterprise", from the generated data).
    '#microsoft.graph.iosMobileAppConfiguration'             = 'AppResources.AppTypePlatform.ios'
    '#microsoft.graph.androidForWorkMobileAppConfiguration'  = 'AppResources.AppTypePlatform.android'

    # App protection. managedAppPolicies is polymorphic - one PolicyType serving
    # several platforms - so this cannot be expressed as a type-level default.
    '#microsoft.graph.androidManagedAppProtection'           = 'AppResources.AppTypePlatform.android'
    '#microsoft.graph.iosManagedAppProtection'               = 'AppProtection.iOSPlatformLabel'
    '#microsoft.graph.windowsManagedAppProtection'           = 'AppResources.AppTypePlatform.windows'
    '#microsoft.graph.mdmWindowsInformationProtectionPolicy' = 'AppResources.AppTypePlatform.windows'

    # MAM app configuration targets iOS and Android apps from one policy, so no
    # single platform applies.
    '#microsoft.graph.targetedManagedAppConfiguration'       = 'Platform.multiplePlatforms'
}

# Keys the portal does not publish. The Platform column is UI, and this project
# keeps UI text in English (documentation is what gets localized), so an English
# literal is the convention - same shape as the iOS/iPadOS fallback in
# IntuneApplicationClasses. If the pipeline ever ships the key, it wins.
$script:PlatformStringFallbacks = @{
    'Platform.multiplePlatforms' = 'Multiple platforms'
}

function Get-PlatformStringByKey
{
    param([string]$Key)

    if([String]::IsNullOrWhiteSpace($Key)) { return $null }

    $name = Get-LanguageString $Key -IgnoreMissing
    if($name) { return $name }

    if($script:PlatformStringFallbacks.ContainsKey($Key)) { return $script:PlatformStringFallbacks[$Key] }

    return $null
}

# The override for one @odata.type, or $null when there is none.
function Get-PolicyPlatformOverride
{
    param([string]$ODataType)

    if([String]::IsNullOrWhiteSpace($ODataType)) { return $null }
    if(-not $script:PolicyPlatformOverrides.ContainsKey($ODataType)) { return $null }

    return Get-PlatformStringByKey $script:PolicyPlatformOverrides[$ODataType]
}

# Translate a platform enum value returned by Graph.
#
# Several endpoints return a comma-separated flags value rather than a single
# member - configurationPolicies hands back "android,iOS" for a policy that
# targets both - and there is no language key for the combined string, so the
# naive lookup returned nothing and the Platform column rendered blank.
# Translate each member and join.
#
# Returns $null unless EVERY member translates: a partially resolved value would
# render as a misleadingly narrow platform (showing just "iOS/iPadOS" for an
# android+iOS policy), and a blank is honest. Tools/Audit-PolicyPlatforms.ps1
# reports what is still blank and why.
function Get-PlatformDisplayName
{
    param([string]$Value)

    if([String]::IsNullOrWhiteSpace($Value)) { return $null }

    $parts = @($Value.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if($parts.Count -eq 0) { return $null }

    $names = @()
    foreach($part in $parts)
    {
        # Hashtable lookup is case-insensitive, which is what we want - the same
        # platform appears as both 'aosp' and 'AOSP' across Graph payloads.
        $key = if($script:PlatformLanguageIdAliases.ContainsKey($part)) { $script:PlatformLanguageIdAliases[$part] } else { $part }

        $name = Get-LanguageString "Platform.$key" -IgnoreMissing
        if(-not $name) { return $null }
        $names += $name
    }

    return ($names -join ", ")
}

function Get-PolicyPlatformName
{
    param($ODataType, $Default = $null)

    $categoryObject =  Get-PolicyObjectCategoryInfo $ODataType

    if($null -eq $categoryObject)
    {
        # No generated row - fall back to the hand-maintained overrides.
        $override = Get-PolicyPlatformOverride $ODataType
        if($override) { return $override }

        return $Default
    }

    $platformName = Get-LanguageString "Platform.$($categoryObject.PlatformLanguageId)"

    if($platformName) { return $platformName }

    # Row exists but its PlatformLanguageId does not translate.
    $override = Get-PolicyPlatformOverride $ODataType
    if($override) { return $override }

    return $Default
}