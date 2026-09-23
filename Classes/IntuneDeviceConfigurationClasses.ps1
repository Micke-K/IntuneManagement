#ImportOrder 200

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class DeviceConfigurationGroup : IntunePolicyGroupBase
{
    DeviceConfigurationGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "DeviceConfiguration"
        $this._Name = "Configuration"
        $this._Icon = "DeviceConfiguration"

    }
}

#########################################################################################
#
# Device Configuration
#
#########################################################################################

#region Device Configuration

class DeviceConfigurationType : IntunePolicyTypeBase
{
    DeviceConfigurationType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceConfigurationGroup")
        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyName = "Device Configuration"

        $this._ID = "DeviceConfiguration"
        $this._API = "deviceManagement/deviceConfigurations"
        $this._QueryList = "`?`$filter=not%20isof(%27microsoft.graph.windowsUpdateForBusinessConfiguration%27)%20and%20not%20isof(%27microsoft.graph.iosUpdateConfiguration%27)%20and%20not%20isof(%27microsoft.graph.macOSSoftwareUpdateConfiguration%27)"
        $this._PolicyBaseName = "Device Configuration"
        $this._PropertiesToRemove = @('privacyAccessControls')
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        # The object class sets a per-row PolicyName ("iOS Wi-Fi") - the kind the
        # portal calls Profile type. PolicyType._PolicyBaseName was a per-type constant.
        $this._SubTypeColumn = "PolicyName=Type"
        $this._Icon = "DeviceConfiguration"
        $this._ObjectClass = "DeviceConfigurationObject"
        $this._NavigationProperties = $true
        $this._ExpandAssignmentsList = $false

    }

    PostImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        if(($PolicyObject.JsonObject.'@OData.Type' -like "#microsoft.graph.windows10GeneralConfiguration" -or 
            $SourceObject.JsonObject.'@OData.Type' -like "#microsoft.graph.windows10GeneralConfiguration") -and 
            ($SourceObject.privacyAccessControls | Measure-Object).Count -gt 0)
        {
                $privacyObject = [PSCustomObject]@{
                    windowsPrivacyAccessControls = $SourceObject.JsonObject.privacyAccessControls
                }
                $json = $privacyObject | ConvertTo-Json -Depth 20

                $url = (?? $this.PolicyType.APIPOST $this.PolicyType.API) + "/$($this.Id)/windowsPrivacyAccessControls"
                $PolicyObject.Set($url, $json) | Out-Null
        }
    }

    # OMA decrypt for windows10CustomConfiguration moved onto DeviceConfigurationObject
    # via the sub-resource batch contract; this type no longer overrides GetFullObject.
}

class DeviceConfigurationObject : IntunePolicyBase
{
    DeviceConfigurationObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    DeviceConfigurationObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "DeviceConfigurationType")
        $this._PolicyName = Get-TemplatePolicyTypeName $this.JsonObject.'@odata.type' $this._PolicyType.PolicyBaseType

        # windows10CustomConfiguration is the only odata.type with secret-referenced
        # OMA settings. Opt into the sub-resource batching contract only for those —
        # other DeviceConfiguration subtypes have nothing extra to fetch.
        if($this.JsonObject.'@odata.type' -eq '#microsoft.graph.windows10CustomConfiguration') {
            $this._HasSubResourceBatch = $true
        }
    }

    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        if($this.JsonObject.'@odata.type' -ne '#microsoft.graph.windows10CustomConfiguration') { return @() }

        $reqs = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach($omaSetting in @($this.JsonObject.omaSettings)) {
            if(-not $omaSetting -or $omaSetting.isEncrypted -ne $true) { continue }
            if(-not $omaSetting.secretReferenceValueId) { continue }
            $secretId = $omaSetting.secretReferenceValueId
            [void]$reqs.Add([PSCustomObject]@{
                Key = "oma:$secretId"
                # Function-call endpoint; Graph supports these inside $batch GETs.
                Url = "deviceManagement/deviceConfigurations/$($this.Id)/getOmaSettingPlainTextValue(secretReferenceValueId='$secretId')"
            })
        }
        return $reqs.ToArray()
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -ne 1 -or -not $Key.StartsWith('oma:')) { return @() }
        if(-not $Body -or -not $Body.Value) { return @() }

        $secretId = $Key.Substring(4)
        # secretReferenceValueId uniquely identifies the OMA setting on this policy;
        # walk omaSettings to find the row whose plaintext we just resolved.
        foreach($omaSetting in @($this.JsonObject.omaSettings)) {
            if($omaSetting -and $omaSetting.secretReferenceValueId -eq $secretId) {
                $omaSetting.isEncrypted = $false
                $omaSetting.secretReferenceValueId = $null

                if($omaSetting.'@odata.type' -eq '#microsoft.graph.omaSettingStringXml' -or
                    $omaSetting.'value@odata.type' -eq '#Binary') {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body.Value)
                    $omaSetting.value = [Convert]::ToBase64String($bytes)
                }
                else {
                    $omaSetting.value = $Body.Value
                }
                break
            }
        }

        return @()
    }
}

#endregion

#########################################################################################
#
# Device Configuration
#
#########################################################################################

# region Device Configuration

class HardwareConfigurationsType : IntunePolicyTypeBase
{
    HardwareConfigurationsType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceConfigurationGroup")
        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyFileAttributes += "configurationFileContent"
        $this._PolicyName = "BIOS configurations and other settings"

        $this._ID = "hardwareConfigurations"
        $this._API = "deviceManagement/hardwareConfigurations"
        # The /assign action's parameter is type-specific here - the default
        # 'assignments' body key is rejected with 400.
        $this._AssignmentsType = "hardwareConfigurationAssignments"
        $this._PropertiesToRemoveForUpdate = @('version')
        $this._PolicyBaseName = "Device Configuration"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        # The object class sets a per-row PolicyName ("iOS Wi-Fi") - the kind the
        # portal calls Profile type. PolicyType._PolicyBaseName was a per-type constant.
        $this._SubTypeColumn = "PolicyName=Type"
        $this._Icon = "DeviceConfiguration"
        $this._ObjectClass = "HardwareConfigurationObject"
        $this._NavigationProperties = $true

    }
}

class HardwareConfigurationObject : IntunePolicyBase
{
    HardwareConfigurationObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    HardwareConfigurationObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "HardwareConfigurationsType")
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
        $this._PolicyName = Get-TemplatePolicyTypeName $this.JsonObject.'@odata.type' $this._PolicyType.PolicyBaseType
    }
}

#endregion

#########################################################################################
#
# Settings Catalog
#
#########################################################################################

# region Settings Catalog

class SettingsCatalogType : SettingsCatalogTypeBase
{
    SettingsCatalogType() : Base()
    {
        ([SettingsCatalogType]$this).Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceConfigurationGroup")
        $this._ID = "SettingsCatalog"
        $this._QueryList = "?`$filter=templateReference/templateFamily eq 'none'"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyTypeOrder = 150
    }

    # ToDo: Verify that it works
    #PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    #{
    #    Add-GraphAssignmentsToExportFile $PolicyObject $PathToFile
    #}
}

#endregion

#########################################################################################
#
# Android OEM Config
#
#########################################################################################

# region Android OEM Config

class AndroidOEMConfigType : IntunePolicyTypeBase
{
    AndroidOEMConfigType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceConfigurationGroup")
        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyName = "Android OEM Config"

        $this._ID = "AndroidOEMConfig"
        # Platform default: the endpoint serves exactly one platform and the
        # objects carry no platforms/platformType field, so the column would
        # otherwise be blank (OEMConfig requires Android Enterprise).
        $this._PlatformName = Get-LanguageString "Platform.androidEnterprise" -IgnoreMissing
        $this._API = "deviceAppManagement/mobileAppConfigurations"
        $this._QueryList = "?`$filter=microsoft.graph.androidManagedStoreAppConfiguration/appSupportsOemConfig eq true"
        $this._Dependencies = @("Applications")
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._Icon = "DeviceConfiguration"
        $this._ObjectClass = "AndroidOEMConfigObject"
        $this._ExpandAssignmentsList = $false

    }

    [Hashtable]PreImportAssignmentsCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        return (@{ "API" = "$($this.API)/$($PolicyObject.Id)/microsoft.graph.managedDeviceMobileAppConfiguration/assign" })
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        Import-AppConfigurationTargetedApps $PolicyObject
        return $null
    }
}

class AndroidOEMConfigObject : IntunePolicyBase
{
    AndroidOEMConfigObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AndroidOEMConfigObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AndroidOEMConfigType")
    }
}

#endregion

#########################################################################################
#
# Administrative Templates
#
#########################################################################################

# region Administrative Templates

class AdminTemplateType : IntunePolicyTypeBase
{
    AdminTemplateType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceConfigurationGroup")
        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyName = "Administrative template"
        $this._PolicyBaseName = "Administrative templates"

        $this._ID = "AdministrativeTemplates"
        $this._API = "deviceManagement/groupPolicyConfigurations"                
        $this._PropertiesToRemove = @("definitionValues","policyConfigurationIngestionType")
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._Icon = "DeviceConfiguration"
        $this._ObjectClass = "AdminTemplateObject"
        $this._Dependencies = @("ADMXFiles")

    }

    PostImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        if($SourceObject.Object.definitionValues)
        {
            Import-GPOSetting $PolicyObject $SourceObject.Object.definitionValues
        }
    }

    [Hashtable]GetCompareConfig()
    {
        return @{
            Prop        = "definitionValues"
            GetKey      = { param($s) Get-DefinitionValueKey $s }
            GetValue    = { param($s) Get-DefinitionValueDisplayValue $s }
            GetCategory = { param($s) if($s.definition) { "$($s.definition.categoryPath)" } else { "$($s.'#Definition_categoryPath')" } }
            Hydrate     = {
                param($policy)
                if(-not $policy -or -not $policy.Object) { return }
                if($policy.Object.definitionValues) { return }
                if(-not $policy._TokenId) { return }
                $url = "deviceManagement/groupPolicyConfigurations/$($policy.Id)/definitionValues?`$expand=definition,presentationValues"
                $result = Invoke-MSGraphAPI -Url $url -ODataMetadata "skip" -TokenId $policy._TokenId
                if($result -and $result.value)
                {
                    $policy.Object | Add-Member NoteProperty -Name "definitionValues" -Value $result.value -Force
                }
            }
        }
    }
}

class AdminTemplateObject : IntunePolicyBase
{
    Hidden [Hashtable]$_SubResourceState = $null

    AdminTemplateObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AdminTemplateObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AdminTemplateType")
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing

        # Opt into the sub-resource batching contract — Invoke-PolicyHydrate
        # drives the definitionValues + presentationValues cascade via $batch.
        $this._HasSubResourceBatch = $true
    }

    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        $reqs = @([PSCustomObject]@{
            Key = 'definitionValues'
            Url = "deviceManagement/groupPolicyConfigurations/$($this.Id)/definitionValues?`$expand=definition"
        })
        # When the per-id GET returns policyConfigurationIngestionType='unknown',
        # the list endpoint filtered by id returns the resolved value
        # (mixed/custom/builtIn/...). Add the list-filter call to Phase 1 so the
        # apply method can patch the JsonObject before serialisation. Matches
        # what the OLD project's Get-GPOObjectSettings did at fetch time.
        if($this.JsonObject.policyConfigurationIngestionType -eq 'unknown') {
            $reqs += [PSCustomObject]@{
                Key = 'ingestionType'
                Url = "deviceManagement/groupPolicyConfigurations?`$filter=id eq '$($this.Id)'"
            }
        }
        return $reqs
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -eq 1 -and $Key -eq 'ingestionType') {
            $row = if($Body -and $Body.value) { @($Body.value)[0] } else { $null }
            if($row -and $row.policyConfigurationIngestionType) {
                $this.JsonObject.policyConfigurationIngestionType = $row.policyConfigurationIngestionType
            }
            return @()
        }

        if($Phase -eq 1 -and $Key -eq 'definitionValues') {
            $rawValues = if($Body -and $Body.value) { @($Body.value) } else { @() }

            # Reshape into the legacy on-disk format: drop the embedded
            # `definition` object, replace it with a `definition@odata.bind`
            # URL plus four `#Definition_*` flat fields used downstream by
            # compare/documentation/import. Mirrors the original
            # Get-GPOObjectSettings transform.
            $domain     = Get-GraphDomain $this._TokenId
            $apiVersion = $this._PolicyType.APIVersion
            $this._SubResourceState = @{}
            $flatList = @()

            foreach($defValue in $rawValues) {
                if(-not $defValue) { continue }
                $defId = $defValue.definition.id
                $flat = [ordered]@{
                    "enabled"               = $defValue.enabled
                    "definition@odata.bind" = "https://$domain/$apiVersion/deviceManagement/groupPolicyDefinitions('$defId')"
                }
                if($defValue.definition.categoryPath) {
                    $flat["#Definition_Id"]           = $defId
                    $flat["#Definition_displayName"]  = $defValue.definition.displayName
                    $flat["#Definition_classType"]    = $defValue.definition.classType
                    $flat["#Definition_categoryPath"] = $defValue.definition.categoryPath
                }
                $flatObj = [PSCustomObject]$flat
                $flatList += $flatObj

                if($defValue.id) {
                    $this._SubResourceState[$defValue.id] = [PSCustomObject]@{
                        Flat         = $flatObj
                        DefinitionId = $defId
                        HasCategory  = [bool]$defValue.definition.categoryPath
                    }
                }
            }

            Add-Member -InputObject $this.JsonObject -MemberType NoteProperty -Name 'definitionValues' -Value $flatList -Force

            $followups = @()
            foreach($defValue in $rawValues) {
                if(-not $defValue.id) { continue }
                $followups += [PSCustomObject]@{
                    Key = "presentationValues:$($defValue.id)"
                    Url = "deviceManagement/groupPolicyConfigurations/$($this.Id)/definitionValues/$($defValue.id)/presentationValues?`$expand=presentation"
                }
            }
            return $followups
        }

        if($Phase -eq 2 -and $Key -like 'presentationValues:*') {
            $defValueId = $Key.Substring('presentationValues:'.Length)
            $state = $null
            if($null -ne $this._SubResourceState) {
                $state = $this._SubResourceState[$defValueId]
            }
            if(-not $state) { return @() }

            $rawPres = if($Body -and $Body.value) { @($Body.value) } else { @() }
            if($rawPres.Count -eq 0) { return @() }

            $domain      = Get-GraphDomain $this._TokenId
            $apiVersion  = $this._PolicyType.APIVersion
            $defId       = $state.DefinitionId
            $hasCategory = $state.HasCategory
            $shaped = @()

            foreach($pv in $rawPres) {
                if(-not $pv) { continue }
                $presId = $pv.presentation.id
                Add-Member -InputObject $pv -MemberType NoteProperty -Name 'presentation@odata.bind' `
                    -Value "https://$domain/$apiVersion/deviceManagement/groupPolicyDefinitions('$defId')/presentations('$presId')" -Force
                if($hasCategory) {
                    Add-Member -InputObject $pv -MemberType NoteProperty -Name '#Presentation_Id'    -Value $presId -Force
                    Add-Member -InputObject $pv -MemberType NoteProperty -Name '#Presentation_Label' -Value $pv.presentation.label -Force
                }
                Remove-ObjectProperty $pv 'presentation'
                Remove-ObjectProperty $pv 'id'
                Remove-ObjectProperty $pv 'lastModifiedDateTime'
                Remove-ObjectProperty $pv 'createdDateTime'
                $shaped += $pv
            }

            Add-Member -InputObject $state.Flat -MemberType NoteProperty -Name 'presentationValues' -Value $shaped -Force
        }

        return @()
    }
}

#########################################################################################
#
# ADMX File
#
#########################################################################################

# region ADMX File

class ADMXFileType : IntunePolicyTypeBase
{
    ADMXFileType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceConfigurationGroup")
        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyName = "ADMX File"
        $this._ID = "ADMXFiles"
        $this._APITitle = "ADMX Files" 
        $this._API = "deviceManagement/groupPolicyUploadedDefinitionFiles"
        $this._NameProperty = "fileName"
        $this._PropertiesToRemove = @("languageCodes","targetPrefix","targetNamespace","policyType","revision","status","uploadDateTime")
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._Icon = "DeviceConfiguration"
        $this._ObjectClass = "ADMXFileObject"
        $this._ImportOrder = 45
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false

        # Restricted button set: ADMX Files support only View / Import / Export
        # (no Compare/Copy/Document/Delete). Graph returns `content` as null on
        # GET, so an export is the object's metadata only - the ADMX/ADML bytes
        # are not in it. Import therefore reads the source files from disk; see
        # PreImportCommand.
        $this._ShowButtons = @("View","Import","Export")
    }

    [IntunePolicyBase[]]PreImportPolicies([IntunePolicyBase[]]$PolicyObjects)
    {
        return @($PolicyObjects | sort-object -property @{e={$_.Object.lastModifiedDateTime}} )
    }

    # Rebuild the upload payload from the ADMX/ADML files on disk, because the
    # exported json cannot carry them (Graph never returns `content`).
    #
    # Both layouts these files ship in are accepted, searched in this order:
    #
    #   <folder>/<name>.admx + <folder>/<name>.adml           flat (the v3 layout)
    #   <folder>/<name>.admx + <folder>/<lang>/<name>.adml    per-language subfolders
    #
    # <folder> is the exported json's own folder, falling back to the "App
    # packages folder" setting - the same two-location search v3 did. Every
    # language subfolder holding a matching .adml is uploaded, so a multi-language
    # ADMX keeps all of its languages rather than only the default one.
    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        $obj = $PolicyObject.Object

        if(-not $obj.fileName)
        {
            Write-Log "ADMX File object has no fileName. It will not be imported." 2
            return @{ "Import" = $false }
        }

        $baseName = [IO.Path]::GetFileNameWithoutExtension($obj.fileName)

        $searchFolders = @()
        if($PolicyObject.FileInfo -and $PolicyObject.FileInfo.Directory.FullName)
        {
            $searchFolders += $PolicyObject.FileInfo.Directory.FullName
        }
        $pkgPath = Get-SettingValue "IntuneAppPackagesFolder"
        if($pkgPath) { $searchFolders += $pkgPath }

        $admxFile = $null
        $sourceFolder = $null
        foreach($folder in $searchFolders)
        {
            $candidate = [IO.Path]::Combine($folder, $obj.fileName)
            if(Test-Path -LiteralPath $candidate -PathType Leaf)
            {
                $admxFile = $candidate
                $sourceFolder = $folder
                break
            }
        }

        if(-not $admxFile)
        {
            Write-Log "ADMX file $($obj.fileName) not found in the export folder or the app packages folder. The ADMX File object will not be imported." 2
            return @{ "Import" = $false }
        }

        # defaultLanguageCode names the language of a flat .adml, which carries no
        # language in its path. Read it before it is stripped below.
        $defaultLanguage = ?? $obj.defaultLanguageCode "en-US"

        $languageFiles = @()

        $flatADML = [IO.Path]::Combine($sourceFolder, "$baseName.adml")
        if(Test-Path -LiteralPath $flatADML -PathType Leaf)
        {
            $languageFiles += [PSCustomObject]@{ Path = $flatADML; LanguageCode = $defaultLanguage }
        }

        foreach($dir in @(Get-ChildItem -LiteralPath $sourceFolder -Directory -ErrorAction SilentlyContinue))
        {
            $langADML = [IO.Path]::Combine($dir.FullName, "$baseName.adml")
            if(Test-Path -LiteralPath $langADML -PathType Leaf)
            {
                # A flat .adml already claimed this language - keep the first.
                if(@($languageFiles | Where-Object { $_.LanguageCode -eq $dir.Name }).Count -gt 0) { continue }
                $languageFiles += [PSCustomObject]@{ Path = $langADML; LanguageCode = $dir.Name }
            }
        }

        if($languageFiles.Count -eq 0)
        {
            Write-Log "No ADML file found for $($obj.fileName). Looked for $baseName.adml in $sourceFolder and in its language sub-folders. The ADMX File object will not be imported." 2
            return @{ "Import" = $false }
        }

        # ReadAllText detects the source encoding; re-encode as UTF-8. Vendor ADMX
        # often ships as UTF-16 with a BOM, which Intune rejects with "Data at the
        # root level is invalid". v3 wrote ASCII bytes here, which silently mangles
        # any non-ASCII character in a policy name or description.
        $obj.content = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([IO.File]::ReadAllText($admxFile)))

        $obj.groupPolicyUploadedLanguageFiles = @(
            foreach($languageFile in $languageFiles)
            {
                [PSCustomObject]@{
                    "@odata.type" = "#microsoft.graph.groupPolicyUploadedLanguageFile"
                    fileName      = [IO.Path]::GetFileName($languageFile.Path)
                    languageCode  = $languageFile.LanguageCode
                    content       = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([IO.File]::ReadAllText($languageFile.Path)))
                }
            })

        # Graph rejects the property on create ("needs to be null, taken from the
        # ADML file"), so remove it rather than blanking it - a null would still be
        # serialized into the body.
        if($obj.PSObject.Properties['defaultLanguageCode'])
        {
            $obj.PSObject.Properties.Remove('defaultLanguageCode')
        }

        Write-Log "Import ADMX $($obj.fileName) with $($languageFiles.Count) language file(s): $(($languageFiles | ForEach-Object { $_.LanguageCode }) -join ', ')"

        return $null
    }

    # Ingesting an ADMX creates new definition ids, so any cached definition map
    # for this tenant now points at ids that no longer exist. Import-GPOSetting
    # builds that cache to resolve custom ADMX settings on admin-template import;
    # leaving it stale is what made an admin template imported right after its
    # ADMX silently resolve nothing. v3 cleared the same cache here.
    PostImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        $tokenInfo = Get-OperationTokenInfo $PolicyObject.TokenId
        if($tokenInfo) { Clear-CacheObject "ADMXDefinitions_$($tokenInfo.TenantId)" }
    }
}

class ADMXFileObject : IntunePolicyBase
{
    ADMXFileObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    ADMXFileObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "ADMXFileType")
        $this._PolicyName = Get-TemplatePolicyTypeName $this.JsonObject.'@odata.type' $this._PolicyType.PolicyBaseType
        $this._PlatformName = Get-LanguageString "Platform.windows" -IgnoreMissing
    }
}

#endregion

# Get-GPOObjectSettings was the pre-refactor GPO definitionValues +
# presentationValues fetcher. AdminTemplateObject.GetSubResourceBatchRequests /
# ApplySubResourceBatchResult now drives the same flow through the unified
# sub-resource batch contract, so the original helper is dead code. Deleted
# along with its `headers = @{ "Accept" = "application/json"; "odata.metadata" = "minimal" }`
# line — that form splits the OData metadata directive into a separate header
# key Graph ignores, leaving the response at the default `full` metadata, the
# opposite of the intent.

function Import-GPOSetting
{
    param($PolicyObject, $Settings)
    
    if($PolicyObject)
    {
        Write-Status "Import settings for $($PolicyObject.Name)"

        $hasCustomADMX = $null -ne ($Settings | Where-Object { $null -ne $_.'#Definition_categoryPath' })

        # Get-OperationTokenInfo, not Get-TokenInfo: an imported object that carries no
        # token of its own has TokenId 0 - never $null - and 0 means "no filter, every
        # token" to Get-TokenInfo. With a second tenant signed in the key would name
        # both at once, so both tenants share one partition and one tenant's custom
        # ADMX definitions could be used to resolve the other tenant's import.
        $tokenInfo = Get-OperationTokenInfo $PolicyObject.TokenId
        $cacheId = "ADMXDefinitions_$($tokenInfo.TenantId)"

        $customADMXDefinitions = $null

        if($hasCustomADMX)
        {
            $customADMXDefinitions = Get-CacheObject $cacheId
            if(-not $customADMXDefinitions)
            {
                Write-Status "Import custom ADMX settings"
                $tmpCustomCategories = Invoke-MSGraphAPI -Url "deviceManagement/groupPolicyCategories?`$expand=definitions(`$select=id, displayName, categoryPath, classType)&`$select=id, displayName&`$filter=ingestionSource eq 'custom'" -ODataMetadata "Minimal" -TokenId $PolicyObject.TokenId
                if($tmpCustomCategories.Value)
                {
                    $customADMXDefinitions = @{}
                    foreach($tmpCat in $tmpCustomCategories.Value)
                    {
                        foreach($tmpDef in $tmpCat.definitions)
                        {
                            $key = ($tmpDef.displayName + $tmpDef.categoryPath + $tmpDef.classType).ToLower()
                            $val = [PSCustomObject]@{
                                Definition = $tmpDef
                                Category = $tmpCat
                                Presentations = $null
                            }
                            try {
                                $customADMXDefinitions.Add($key, $val)
                            }
                            catch {
                                Write-Log "Failed to add '$($tmpDef.displayName)' in category '$($tmpDef.categoryPath)' of class $($tmpDef.classType)" 3
                            }
                        }
                    }
                }
                Set-CacheObject $cacheId $customADMXDefinitions "TenantCache_$($tokenInfo.TenantId)"
            }
        }        

        # Batch the per-setting POSTs into a single Graph $batch (chunked at 20 by the
        # batcher) instead of one HTTP call per setting. The rewrite/strip logic below
        # mutates each setting in place; we collect the prepared sub-requests in the
        # loop and dispatch them once at the end.
        $importBatch = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach($setting in $Settings)
        {
            if($setting.'#Definition_categoryPath' -and $customADMXDefinitions -is [HashTable] -and $customADMXDefinitions.Count -gt 0)
            {
                $defVal = $null
                $key = ($setting.'#Definition_displayName' + $setting.'#Definition_categoryPath' + $setting.'#Definition_classType').ToLower()
                if($key -and $customADMXDefinitions.ContainsKey($key))
                {
                    $defVal = $customADMXDefinitions[$key]
                }
                elseif($key)
                {
                    Write-Log "No custom ADMX definitiona found for setting $($setting.'#Definition_displayName')" 2                    
                }
                else
                {
                    Write-Log "Setting $($setting.'#Definition_displayName') does not have information to be imported in the environment"
                }

                if($defVal)
                {
                    $setting.'definition@odata.bind' = $setting.'definition@odata.bind' -replace $setting.'#Definition_Id', $defVal.Definition.Id
                    if(($setting.presentationValues | Measure-Object).Count -gt 0)
                    {
                        if(-not $defVal.Presentations)
                        {
                            $tmpPresentation = Invoke-MSGraphAPI -Url "deviceManagement/groupPolicyDefinitions/$($defVal.Definition.Id)/presentations" -ODataMetadata "Minimal" -TokenId $PolicyObject.TokenId
                            if($tmpPresentation.value)
                            {
                                foreach($settingPresentation in $setting.presentationValues)
                                {
                                    $tmpPresentationVal = $tmpPresentation.value | Where-Object label -eq $settingPresentation.'#Presentation_Label'
                                    if($tmpPresentationVal)
                                    {
                                        $settingPresentation.'presentation@odata.bind' = $settingPresentation.'presentation@odata.bind' -replace $setting.'#Definition_Id', $defVal.Definition.Id
                                        $settingPresentation.'presentation@odata.bind' = $settingPresentation.'presentation@odata.bind' -replace $settingPresentation.'#Presentation_Id', $tmpPresentationVal.Id
                                    }
                                    else
                                    {
                                        Write-Log "Could not find a presentation value with label $($settingPresentation.'#Presentation_Label'). Setting will not be configured" 2
                                        continue
                                    }
                                }
                            }
                            else
                            {
                                Write-Log "Could not find presentation for setting $($settingPresentation.'#Presentation_Label'). Setting will not be configured." 2
                                continue
                            }
                        }
                    }
                }
                else
                {
                    Write-Log "Settings might not be available if imported in another environment" 3
                }
            }
            elseif($setting.'#Definition_categoryPath')
            {
                Write-Log "Custom AMDX settings cannot be imported without ADMX file imported. Definitions not found" 2
                continue
            }

            Remove-GraphPropertiesForImport $PolicyObject $setting

            if($true) 
            {
                foreach($tmpProp in (($setting.PSObject.Properties | Where-Object Name -like "#*").Name))
                {
                    Remove-Property $setting $tmpProp
                }
                
                foreach($settingPresentation in $setting.presentationValues)
                {
                    foreach($tmpProp in (($settingPresentation.PSObject.Properties | Where-Object Name -like "#*").Name))
                    {
                        Remove-Property $settingPresentation $tmpProp
                    }
                }
            }

            # Queue the setting POST; actual dispatch happens after the loop in a single batch.
            [void]$importBatch.Add([PSCustomObject]@{
                id      = [Guid]::NewGuid().Guid
                method  = "POST"
                url     = "deviceManagement/groupPolicyConfigurations/$($PolicyObject.id)/definitionValues"
                headers = @{ "Content-Type" = "application/json" }
                body    = ($setting | ConvertTo-Json -Depth 20 | ConvertFrom-Json)
            })
        }

        if($importBatch.Count -gt 0)
        {
            Invoke-GraphBatchRequest -BatchObjects $importBatch -BatchType "GPOSettingsImport" -TokenId $PolicyObject.TokenId | Out-Null
        }
    }
}

#########################################################################################
#
# Inventory Policies
#
#########################################################################################

class InventoryPoliciesType : IntunePolicyTypeBase
{
    InventoryPoliciesType() : Base() { $this.Init() }
    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceConfigurationGroup")
        $this._PolicyName = "Inventory Policies"
        $this._ID = "InventoryPolicies"
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "deviceManagement/inventoryPolicies"
        $this._QueryList = "?`$expand=settings"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._PropertiesToRemove = @('settingCount')
        $this._NameProperty = "name"
        $this._ObjectClass = "InventoryPoliciesObject"
        $this._Icon = "DeviceConfiguration"
        if($null -ne $this._PolicyGroup) { $this._PolicyGroup.AddPolicyType($this) }
    }
}

class InventoryPoliciesObject : IntunePolicyBase
{
    InventoryPoliciesObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }
    InventoryPoliciesObject() : Base() { $this.Init() }
    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "InventoryPoliciesType")
    }
}


#########################################################################################
#
# Policy Sets
#
#########################################################################################
#
# Sits under the Applications group because Intune's portal puts PolicySets
# under Apps and the Graph endpoint is rooted at deviceAppManagement. Each
# PolicySet bundles references to other policies (apps, configs, scripts,
# etc.) and assigns them as one unit.
#
# Cross-tenant import: items[].payloadId references are re-pointed by name
# (or dropped when unresolvable) via Resolve-IntunePolicySetItems - see
# Internal/IntunePolicySetImport.ps1.

#region PolicySetsObject

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class PolicySetObject : IntunePolicyBase
{
    PolicySetObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    PolicySetObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "PolicySetsType")
        # No _PlatformName — PolicySets are cross-platform bundles.
    }
}

#endregion

#region PolicySetsType
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class PolicySetsType : IntunePolicyTypeBase
{
    PolicySetsType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "DeviceConfigurationGroup")
        $this._PolicyName  = "Policy Sets"
        $this._ID          = "PolicySets"
        $this._API         = "deviceAppManagement/policySets"
        # items + assignments come back via $expand. Without items the round-
        # trip on export loses the bundled references; assignments load with
        # the policy so the assignments tool can see them.
        $this._Expand      = "items,assignments"
        $this._Permissions = @("DeviceManagementApps.ReadWrite.All")
        # status / errorCode / lastModifiedDateTime are server-set; strip on
        # both create and update to avoid Graph rejecting writes.
        $this._PropertiesToRemove          = @('status','errorCode')
        # items + assignments are navigation properties Graph rejects on
        # PATCH ("Cannot apply PATCH to navigation property ..."); items go
        # through the /update action in PreUpdateCommand instead.
        $this._PropertiesToRemoveForUpdate = @('status','errorCode','items','assignments')
        # PolicySet uses `roleScopeTags` (string array of tag IDs, NOT the
        # more common `roleScopeTagIds`). Same Microsoft Graph quirk as
        # AssignmentFilters — see reference_assignment_filter_scope_tag_property.
        $this._ScopeTagProperty = "roleScopeTags"
        # Policy sets reference other objects, so they must import LAST. The
        # default order is 1000; 200 put them ahead of Compliance, Settings
        # Catalog, App Protection and App Config, which are exactly what a set
        # bundles. It went unnoticed because same-tenant imports resolve items
        # by payloadId before the referenced objects are re-created; a portable
        # import resolves by name and dropped every item. v3 used 2000 for the
        # same reason.
        $this._ImportOrder = 2000
        $this._ObjectClass = "PolicySetObject"
        # Graph returns HTTP 400 on `deviceAppManagement/policySets?$expand=assignments`,
        # so suppress the list-URL expand. Per-policy GETs still use _Expand to
        # carry items + assignments — that variant Graph accepts.
        $this._ExpandAssignmentsList = $false
        # The per-policy /assignments navigation GET 400s too - only the
        # single-object $expand variant works.
        $this._AssignmentsViaExpand = $true
        # No /assign segment on policySets ("Resource not found for the
        # segment 'assign'") - the update action takes the full `assignments`
        # replacement list with the same semantics.
        $this._AssignAction = "update"
        $this._AssignmentObjectType = "#microsoft.graph.policySetAssignment"
        # status / errorCode track the ASYNC processing state of the set
        # ('notAssigned' -> 'success'), and every item embeds the same
        # per-item state - so identically-configured sets compare unequal on
        # raw properties. Item MEMBERSHIP drift is still caught by the
        # documentation compare (PolicySetDocHandler documents each item).
        $this._ComparePropertiesToSkip = @('status','errorCode','items')
        # PolicySets aren't platform-specific; the bundled items each have
        # their own platforms but the set itself spans them.
        $this._HasPlatform = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        # Re-point (or drop) cross-tenant item references and strip items to
        # the POST-safe shape - see Internal/IntunePolicySetImport.ps1.
        Resolve-IntunePolicySetItems -PolicyObject $PolicyObject -TokenId ([int]$PolicyObject.TokenId)
        return (@{})
    }

    [Hashtable]PreUpdateCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$ExistingObject)
    {
        # items cannot be PATCHed - push membership deltas through the
        # /update action first; the PATCH then carries metadata only
        # (_PropertiesToRemoveForUpdate strips items + assignments).
        Update-IntunePolicySetItems -PolicyObject $PolicyObject -ExistingObject $ExistingObject -TokenId ([int]$PolicyObject.TokenId)
        return (@{})
    }
}
#endregion