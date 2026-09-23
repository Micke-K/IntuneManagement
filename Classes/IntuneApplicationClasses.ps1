#ImportOrder 220

#########################################################################################
#
# Applications Group
#
#########################################################################################

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class ApplicationsGroup : IntunePolicyGroupBase
{
    ApplicationsGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "Applications"
        $this._Name = "Applications"
        $this._Icon = "Applications"
        # Type (Win32, iOS store, ...) is the discriminator here; Policy type would read
        # "Application" on every row but the iOS provisioning profiles.
        $this._ExtraColumns = @("ApplicationType=Type")
        $this._ShowPolicyTypeColumn = $false
    }
}

#########################################################################################
#
# Application Type
#
#########################################################################################

# region Application Type
class ApplicationType : IntunePolicyTypeBase
{
    ApplicationType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ApplicationsGroup")
        $this._APITitle = "Applications"
        $this._PolicyName = "Applications"
        $this._ID = "Applications"
        $this._API = "deviceAppManagement/mobileApps"
        $this._QueryList = "?`$filter=(microsoft.graph.managedApp/appAvailability eq null or microsoft.graph.managedApp/appAvailability eq 'lineOfBusiness' or isAssigned eq true)&`$orderby=displayName"
        $this._QuerySearch = $true
        $this._Expand = "categories,assignments" # ODataMetadata is set to minimal so assignments can't be autodetected
        $this._ODataMetadata = "minimal" # categories property not supported with ODataMetadata full
        $this._Permissions = @("DeviceManagementApps.ReadWrite.All")
        $this._PropertiesToRemove = @('uploadState','publishingState','isAssigned','dependentAppCount','supersedingAppCount','supersededAppCount','committedContentVersion','isFeatured','size','categories') #,'minimumSupportedWindowsRelease'
        $this._AssignmentsType = "mobileAppAssignments"
        $this._AssignmentPropertiesToKeep = @("@odata.type","target","settings","intent")
        $this._AssignmentTargetPropertiesToKeep = @("@odata.type","groupId","deviceAndAppManagementAssignmentFilterId","deviceAndAppManagementAssignmentFilterType")
        $this._ScopeTagsReturnedInList = $false
        $this._ExpandAssignmentsList = $false
        $this._ImportOrder = 60
        $this._ObjectClass = "ApplicationObject"
        $this._SubTypeColumn = "ApplicationType=Type"
        $this._ExtraColumns  = @("ApplicationTypeGroup=App type")

        # appUrl: Graph rejects a PATCH that carries it ("The property 'AppUrl'
        # cannot be patched") - a web app's URL is fixed at creation, as in the
        # portal. Only webApp has the property, so stripping it is safe for all.
        $this._PropertiesToRemoveForUpdate = @('platform', 'appUrl')

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.JsonObject.'@OData.Type' -in @('#microsoft.graph.microsoftStoreForBusinessApp','#microsoft.graph.androidStoreApp'))
        {
            Write-Log "App type '$($PolicyObject.JsonObject.'@OData.Type')' not supported for import" 2
            return @{ "Import" = $false }
        }

        if($PolicyObject.JsonObject.'@OData.Type' -eq '#microsoft.graph.officeSuiteApp')
        {
            if($PolicyObject.JsonObject.officeSuiteAppDefaultFileFormat -eq "notConfigured")
            {
                $PolicyObject.JsonObject.officeSuiteAppDefaultFileFormat = "officeOpenXMLFormat"
            }
        }

        return $null
    }

    PostImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        $tmpFilName = $null

        if($SourceObject.IsFromFile) {
            if(-not ($PolicyObject.JsonObject.PSObject.Properties | Where-Object Name -eq '@odata.type'))
            {
                # Add @odata.type property if it is missing. Required by app package import
                $PolicyObject.JsonObject | Add-Member -MemberType NoteProperty -Name '@odata.type' -Value $SourceObject.JsonObject.'@odata.type'
            }
        
            $fi = $SourceObject.FileInfo
            $tmpFilName = [IO.Path]::Combine($fi.DirectoryName, [string]$SourceObject.JsonObject.FileName)
        
            if([IO.File]::Exists($tmpFilName) -eq $false)
            {
                Write-LogDebug "App content file not found in Json folder: '$tmpFilName'"
                $tmpFilName = $null
            }
        }
        else {
        }
        
        Start-ApplicationImportFile $PolicyObject $tmpFilName
        Start-ApplicationAddInstallScripts $PolicyObject $SourceObject
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        $fi = [IO.FileInfo]"$PathToFile"

        if((Get-CacheObject "ExportScripts") -eq $true) {
            try
            {
                foreach($rule in ($PolicyObject.JsonObject.detectionRules | Where-Object '@OData.Type' -eq "#microsoft.graph.win32LobAppPowerShellScriptDetection"))
                {
                    if($rule.ScriptContent)
                    {
                        [IO.File]::WriteAllBytes(([IO.Path]::Combine($fi.DirectoryName, "$($fi.BaseName)_DetectionScript.ps1")), ([System.Convert]::FromBase64String($rule.ScriptContent)))
                    }
                }

                foreach($rule in $PolicyObject.JsonObject.requirementRules)
                {
                    if($rule.'@OData.Type' -eq "#microsoft.graph.win32LobAppPowerShellScriptRequirement")
                    {
                        if($rule.ScriptContent)
                        {
                            [IO.File]::WriteAllBytes(([IO.Path]::Combine($fi.DirectoryName, "$($fi.BaseName)_RequirementScript.ps1")), ([System.Convert]::FromBase64String($rule.ScriptContent)))
                        }
                    }
                }

                if($PolicyObject.JsonObject.activeInstallScript.'#ScriptInfo'.displayName)
                {
                    [IO.File]::WriteAllBytes(([IO.Path]::Combine($fi.DirectoryName, "$($fi.BaseName)_$($PolicyObject.JsonObject.activeInstallScript.'#ScriptInfo'.displayName)")), ([System.Convert]::FromBase64String($PolicyObject.JsonObject.activeInstallScript.'#ScriptInfo'.content)))
                }

                if($PolicyObject.JsonObject.activeUninstallScript.'#ScriptInfo'.displayName)
                {
                    [IO.File]::WriteAllBytes(([IO.Path]::Combine($fi.DirectoryName, "$($fi.BaseName)_$($PolicyObject.JsonObject.activeUninstallScript.'#ScriptInfo'.displayName)")), ([System.Convert]::FromBase64String($PolicyObject.JsonObject.activeUninstallScript.'#ScriptInfo'.content)))
                }
            }
            catch
            {
                Write-LogError "Failed to export application scripts" $_.Exception
            }
        }

        Save-SettingStoreValue "Intune" "ExportAppFile" (Get-CacheObject "ExportAppContent")
        if((Get-CacheObject "ExportAppContent") -eq $true) {
            if($script:_skipDirectGet -eq $true) {
                Write-Log "Bulk export: app content download is skipped because direct Graph GET calls are disabled" 2
                return
            }
            $encryptionSource = Get-SettingValue "IntuneAppDownloadFolder" (Get-SettingValue "IntuneAppPackagesFolder")
            $pkgPath = $fi.DirectoryName 

            if($pkgPath)
            {
                Write-Log "Download file $($PolicyObject.JsonObject.FileName)"

                $exportFile = [IO.Path]::Combine($pkgPath, "$($PolicyObject.JsonObject.FileName).encrypted")
                $contentFileObj = Start-DownloadAppContent $PolicyObject $exportFile -GetContentFileInfoOnly
                $encryptionFile = Find-AppEncryptionFile $PolicyObject $contentFileObj $encryptionSource            
                if($encryptionFile -and [IO.File]::Exists($encryptionFile))
                {
                    Start-DownloadFile $contentFileObj.azureStorageUri $exportFile
    
                    if([IO.File]::Exists($exportFile))
                    {
                        Write-Log "Decrypt file"
                        $encryptionInfo = ConvertFrom-Json ([IO.File]::ReadAllText($encryptionFile))
                        if($encryptionInfo.fileEncryptionInfo)
                        {
                            $encryptionInfo = $encryptionInfo.fileEncryptionInfo
                        }                    
                        $destination = $pkgPath + ("\$($PolicyObject.JsonObject.FileName)" -replace 'intunewin$', 'zip')
                        Start-DecryptFile $exportFile $destination $encryptionInfo.encryptionKey $encryptionInfo.initializationVector
                    }

                    try { [IO.File]::Delete($exportFile) }
                    catch {
                        Write-LogError "Failed to delete exported encrypted file" $_.Exception
                    }
                }
                else
                {
                    Write-Log "Could not find encryption file"
                }
            }
        }
    }

    [Hashtable]PreImportAssignmentsCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        if($PolicyObject.JsonObject.'@odata.type' -eq "#microsoft.graph.windowsMicrosoftEdgeApp")
        {
            $assignments = $SourceObject.JsonObject.assignments
            foreach($assignment in $assignments)
            {
                Remove-Property $assignment.target "deviceAndAppManagementAssignmentFilterId"
                Remove-Property $assignment.target "deviceAndAppManagementAssignmentFilterType"
            }
            return (@{"Assignments" = $assignments})
        }
        elseif($PolicyObject.JsonObject.'@odata.type' -eq "#microsoft.graph.winGetApp")
        {
            Write-LogDebug "Wait for '$($PolicyObject.Name)' to be published"
            $i = 2
            Start-Sleep -s ($i)
            $x = 0
            while($x -lt 10)
            {
                $appInfo = Invoke-MSGraphAPI -Url "$($PolicyObject.PolicyType.API)/$($PolicyObject.id)" -ODataMetadata "skip" -TokenId $PolicyObject.TokenId
                if($appInfo.publishingState -eq "Published")
                {
                    Write-LogDebug "Application '$($PolicyObject.Name)' is published"
                    return $null
                }
                Start-Sleep -s ($i)
                $x++
                if($x -ge 5) { $i++ }
            }
    
            Write-Log "Application '$($PolicyObject.Name)' is not published. Skipping assignments" 2
            return (@{"Import" = $false})
        }
        return $null
    }

    PostBulkImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        Add-ApplicationReferences $PolicyObject $SourceObject
    }    

    [Hashtable]PreUpdateCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        if($PolicyObject.JsonObject.'@OData.type' -eq "#microsoft.graph.windowsMobileMSI")
        {
            Remove-Property $PolicyObject.JsonObject "useDeviceContext"
        }
        elseif($PolicyObject.JsonObject.'@OData.type' -eq "#microsoft.graph.officeSuiteApp")
        {
            Remove-Property $PolicyObject.JsonObject "officeConfigurationXml"
            Remove-Property $PolicyObject.JsonObject "officePlatformArchitecture"
            Remove-Property $PolicyObject.JsonObject "developer"
            Remove-Property $PolicyObject.JsonObject "owner"
            Remove-Property $PolicyObject.JsonObject "publisher"
        }
        elseif($PolicyObject.JsonObject.'@OData.type' -eq "#microsoft.graph.winGetApp")
        {
            # Immutable after creation - Graph: "The property
            # 'InstallExperience' cannot be patched."
            Remove-Property $PolicyObject.JsonObject "installExperience"
            Remove-Property $PolicyObject.JsonObject "packageIdentifier"
            Remove-Property $PolicyObject.JsonObject "manifestHash"
        }
    
        Remove-Property $PolicyObject.JsonObject "appStoreUrl"

        # assignments is a navigation property (managed via /assign), not
        # inline-PATCHable: "Cannot apply PATCH to navigation property
        # 'assignments' on entity type mobileApp".
        Remove-Property $PolicyObject.JsonObject "assignments"
        Remove-Property $PolicyObject.JsonObject "assignments@odata.context"

        return $null
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        # Exported app json carries a top-level @odata.id (and, when assignments
        # were expanded, assignments@odata.context) - both identify the
        # mobileApps entity set, which only hosts apps.
        if($PolicyObject.'@odata.id' -like '*deviceAppManagement/mobileApps(*') {
            return $true
        }

        if($PolicyObject.'assignments@odata.context' -like '*#deviceAppManagement/mobileApps(*') {
            return $true
        }

        return $false
    }
}

Class ApplicationObject : IntunePolicyBase
{
    Hidden [String]$_AppTypeName = $null
    Hidden [String]$_AppTypeGroup = $null
    Hidden [String]$_InstallerType = $null
    # Carries phase-1 → phase-2 routing state (script id → { Script; Target })
    # so the orchestrator's phase-2 ApplyResult can find which install/uninstall
    # target object to attach #ScriptInfo to without re-walking the script list.
    Hidden [Hashtable]$_SubResourceState = $null

    ApplicationObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    ApplicationObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "ApplicationType")
        $this._PlatformName = Get-GraphApplicationPlatform $this
        $this._AppTypeGroup = Get-GraphApplicationTypeGroup $this
        $this._AppTypeName = (Get-GraphApplicationName $this)
        $this._HasSubResourceBatch = $true

        if($this.JsonObject."@OData.Type" -eq "#microsoft.graph.winGetApp") {
            if($this.JsonObject.packageIdentifier -like "9*")
            {
                $this._InstallerType = "UWP"
            }
            elseif($this.JsonObject.packageIdentifier -like "X*")
            {
                $this._InstallerType = "Win32"
            }
            else
            {
                Write-Log "Unknown package identifier for app $($this.Name): $($this.JsonObject.packageIdentifier)" 2
                $this._InstallerType = "Unknown"
            }
        }

        Add-ObjectProperty $this "ApplicationType" { $this._AppTypeName }
        Add-ObjectProperty $this "ApplicationTypeGroup" { $this._AppTypeGroup }
        Add-ObjectProperty $this "InstallerType" { $this._InstallerType }

    }

    # Phase 1: relationships (when there are dependencies/supersedences) and the
    # win32 script list (when an active install/uninstall script is referenced).
    # Phase 2 follow-ups are produced by the phase-1 ApplyResult below.
    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }

        $reqs = [System.Collections.Generic.List[PSCustomObject]]::new()

        if(([int64]($this.Object.dependentAppCount) -gt 0) -or ([int64]($this.Object.supersededAppCount) -gt 0)) {
            [void]$reqs.Add([PSCustomObject]@{
                Key = 'rel'
                Url = "deviceAppManagement/mobileApps/$($this.Id)/relationships?`$filter=targetType%20eq%20microsoft.graph.mobileAppRelationshipType%27child%27"
            })
        }

        if($this.Object.'@odata.type' -eq '#microsoft.graph.win32LobApp' -and
           ($this.Object.activeInstallScript.targetId -or $this.Object.activeUninstallScript.targetId)) {
            [void]$reqs.Add([PSCustomObject]@{
                Key = 'scriptlist'
                Url = "deviceAppManagement/mobileApps/$($this.Id)/microsoft.graph.win32LobApp/contentVersions/$($this.Object.committedContentVersion)/scripts/"
            })
        }

        return $reqs.ToArray()
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -eq 1 -and $Key -eq 'rel') {
            $deps = @()
            $sups = @()
            foreach($rel in @($Body.value)) {
                if($rel.'@odata.type' -eq '#microsoft.graph.mobileAppDependency') {
                    $deps += "$($rel.targetDisplayName)|!|$($rel.targetDisplayVersion)|!|$($rel.targetId)|!|$($rel.dependencyType)"
                }
                elseif($rel.'@odata.type' -eq '#microsoft.graph.mobileAppSupersedence') {
                    $sups += "$($rel.targetDisplayName)|!|$($rel.targetDisplayVersion)|!|$($rel.targetId)|!|$($rel.supersedenceType)"
                }
            }
            if($deps.Count -gt 0) {
                $this.Object | Add-Member -MemberType NoteProperty -Name '#CustomRefDependency' -Value ($deps -join '|*|') -Force
            }
            if($sups.Count -gt 0) {
                $this.Object | Add-Member -MemberType NoteProperty -Name '#CustomRefSupersedence' -Value ($sups -join '|*|') -Force
            }
            return @()
        }

        if($Phase -eq 1 -and $Key -eq 'scriptlist') {
            if($null -eq $this._SubResourceState) { $this._SubResourceState = @{} }
            $followups = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach($script in @($Body.value)) {
                $target = $null
                if($this.Object.activeInstallScript.targetId -eq $script.id) {
                    $target = $this.Object.activeInstallScript
                }
                elseif($this.Object.activeUninstallScript.targetId -eq $script.id) {
                    $target = $this.Object.activeUninstallScript
                }
                else {
                    Write-Log "Script with id $($script.id) is not referenced by active install or uninstall script. Skipping." 2
                    continue
                }
                $stateKey = "scriptcontent_$($script.id)"
                $this._SubResourceState[$stateKey] = [PSCustomObject]@{ Script = $script; Target = $target }
                [void]$followups.Add([PSCustomObject]@{
                    Key = $stateKey
                    Url = "deviceAppManagement/mobileApps/$($this.Id)/microsoft.graph.win32LobApp/contentVersions/$($this.Object.committedContentVersion)/scripts/$($script.id)?`$select=Id,Content"
                })
            }
            return $followups.ToArray()
        }

        if($Phase -eq 2 -and $Key -like 'scriptcontent_*' -and $null -ne $this._SubResourceState) {
            $state = $this._SubResourceState[$Key]
            if($state) {
                if($Body -and $Body.Content) {
                    $state.Script | Add-Member -MemberType NoteProperty -Name 'content' -Value $Body.Content -Force
                }
                $state.Target | Add-Member -MemberType NoteProperty -Name '#ScriptInfo' -Value $state.Script -Force
                $this._SubResourceState.Remove($Key) | Out-Null
            }
            return @()
        }

        return @()
    }
}

#endregion

function Start-ApplicationImportFile
{
    param($PolicyObject, $PackageFile = $null)
    
    if(-not $PolicyObject.JsonObject.'@odata.type') { return }

    if($null -eq $PackageFile)
    {
        $pkgPath = Get-SettingValue "IntuneAppPackagesFolder"

        if(-not $pkgPath -or [IO.Directory]::Exists($pkgPath) -eq $false) 
        {
            Write-LogDebug "Package source directory in Settings is not specified" 2
            return 
        }
        elseif([IO.Directory]::Exists($pkgPath) -eq $false) 
        {
            Write-LogDebug "Package source directory '$($pkgPath)' does not exist" 2
            return 
        }        

        $PackageFile = [IO.Path]::Combine($pkgPath, "$($PolicyObject.JsonObject.fileName)")
        $packageFile2 = [IO.Path]::Combine($pkgPath, $PolicyObject.Name, "$($PolicyObject.JsonObject.fileName)")
        if([IO.File]::Exists($PackageFile) -eq $false -and [IO.File]::Exists($packageFile2)) {
            $PackageFile = $packageFile2
        }
    }
    $fi = [IO.FileInfo]$PackageFile

    if($fi.Exists -eq $false) 
    {
        Write-LogDebug "Package source file $($fi.FullName) not found" 2
        return
    }

    Write-Status "Import application package file $($fi.FullName)"
    Write-Log "Import application file '$($($fi.FullName))' for $($PolicyObject.Name)"

    $appType = $PolicyObject.JsonObject.'@odata.type'.Trim('#')

    if($appType -eq "microsoft.graph.win32LobApp")
    {
        $fileEncryptionInfo = Copy-Win32LOBPackage $PackageFile $PolicyObject
    }
    elseif($appType -eq "microsoft.graph.windowsMobileMSI")
    {
        $fileEncryptionInfo = Copy-MSILOB $PackageFile $PolicyObject
    }
    elseif($appType -eq "microsoft.graph.windowsUniversalAppX")
    {
        $fileEncryptionInfo = Copy-MSIXLOB $PackageFile $PolicyObject
    }    
    elseif($appType -eq "microsoft.graph.iosLOBApp")
    {
        $fileEncryptionInfo = Copy-iOSLOB $PackageFile $PolicyObject
    }
    elseif($appType -eq "microsoft.graph.androidLOBApp")
    {
        $fileEncryptionInfo = Copy-AndroidLOB $PackageFile $PolicyObject
    }
    else
    {
        Write-Log "Unsupported application type $appType. File will not be uploaded" 2    
    }

    if((Get-SettingValue "IntuneSaveEncryptionFile") -eq $true)
    {
        if($fileEncryptionInfo)
        {
            $jsonEncryptionInfo = $fileEncryptionInfo | ConvertTo-Json -Depth 10
            
            $pkgPath = Get-SettingValue "IntuneAppDownloadFolder" (Get-SettingValue "IntuneAppPackagesFolder")
            if($pkgPath -and [IO.Directory]::Exists($pkgPath))
            {
                $obj = Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$($PolicyObject.id)" -ODataMetadata "Minimal" -TokenId $PolicyObject._TokenID
                $fullPath = [IO.Path]::Combine($pkgPath, "$($obj.displayName)_$($obj.id)_$($obj.committedContentVersion).json")
                $jsonEncryptionInfo | Out-File -FilePath $fullPath -Force -Encoding utf8
            }
        }
    }
}

function Start-ApplicationAddInstallScripts
{
    param($PolicyObject, $FromAppObj)

    if($FromAppObj -and ($FromAppObj.activeInstallScript."#ScriptInfo" -or $FromAppObj.activeUninstallScript."#ScriptInfo"))
    {
        Write-Log "Importing scripts for $($PolicyObject.displayName)"

        $scriptsAdded = $false        
        $jsonData = @{}
        $jsonData."@odata.type"             = "#microsoft.graph.win32LobApp"
        $jsonData."committedContentVersion" = "1"

        foreach ($scriptType in @('activeInstallScript','activeUninstallScript')) {
            $scriptInfo = $FromAppObj.$scriptType.'#ScriptInfo'
            if (-not $scriptInfo) { continue }

            Write-Log "Add $($scriptType -replace '^active','') script: $($scriptInfo.displayName)"

            $json = [ordered]@{
                '@odata.type'          = $scriptInfo.'@odata.type'
                displayName            = $scriptInfo.displayName
                enforceSignatureCheck  = $scriptInfo.enforceSignatureCheck
                runAs32Bit             = $scriptInfo.runAs32Bit
                content                = $scriptInfo.content
            } | ConvertTo-Json -Depth 10 -Compress

            $scriptObject = Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$($PolicyObject.id)/microsoft.graph.win32LobApp/contentVersions/1/scripts" -Method POST -Content $json -TokenId $PolicyObject._TokenID

            if ($scriptObject) {
                $jsonData.$scriptType = @{ targetId = $scriptObject.Id }
                $scriptsAdded = $true
            }
        }

        $i = 0
        while($true)
        {
            $scripts = Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$($PolicyObject.id)/microsoft.graph.win32LobApp/contentVersions/1/scripts" -TokenId $PolicyObject._TokenID
            if(-not $scripts)
            {
                Write-Log "Failed to retrieve scripts for app after adding. Skipping Install/Uninstall script config." 2
                return
            }

            if(($scripts.value.state | Select -Unique) -eq "commitSuccess")
            {
                Write-Log "Scripts added successfully"
                break
            }
            if($i -ge 12)
            {
                Write-Log "Install/Uninstall scripts are still not in pending state after waiting for 1 minute." 3
                return
            }

            Write-Log "Waiting for scripts to be added..."
            Start-Sleep -Seconds 5
            $i++
        }        

        if($scriptsAdded)
        {
            Write-Log "Add script info to app"
            $json = ConvertTo-Json $jsonData -Depth 10
            $status = Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$($PolicyObject.id)" -Method PATCH -Body $json -TokenId $PolicyObject._TokenID -FullResponseObject
            if($status.Success)
            {
                Write-Log "Install/Uninstall script info updated successfully"
            }
            else
            {
                Write-Log "Failed to update Install/Uninstall script info" 2
            }
        }
    }
}

function Add-ApplicationReferences
{
    param($PolicyObject, $SourceObject)

    if($SourceObject.JsonObject."#CustomRefDependency" -or $SourceObject.JsonObject."#CustomRefSupersedence")
    {
        Write-Log "Adding app references for $($PolicyObject.displayName)"

        $depAppsInfo = $SourceObject.JsonObject."#CustomRefDependency"
        $supAppsInfo = $SourceObject.JsonObject."#CustomRefSupersedence"

        $releationShips = [PSCustomObject]@{
            relationships = @()
        }

        if($depAppsInfo)
        {
            foreach($depApp in ($depAppsInfo -split "[|][*][|]"))
            {
                $appName, $appVer, $appId, $appType = $depApp -split "[|][!][|]"
                if(-not $appName -or -not $appVer)
                {
                    Write-Log "Could not get Name and Version from string: $($PolicyObject.displayName)" 2
                    continue
                }
                $tmpApps = (Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps?`$filter=displayName eq '$appName'" -TokenId $PolicyObject.TokenId).value
                if(-not $tmpApps)
                {
                    Write-Log "No application found with name $appName" 2
                    continue
                }
                $tmpApp = $tmpApps | Where-Object displayVersion -eq $appVer
                if(-not $tmpApp)
                {
                    Write-Log "No $appName application found with version $appVer" 2
                    continue
                }
                elseif(($tmpApp | Measure-Object).Count -gt 1)
                {
                    Write-Log "Multiple $appName application found with version $appVer" 2
                    continue
                }
                Write-Log "Add $appName ($appVer) to Dependency list"
                $releationShips.relationships += [PSCustomObject]@{
                    "@odata.type" = "#microsoft.graph.mobileAppDependency"
                    targetId = $tmpApp.Id
                    dependencyType = $appType
                }
            }
        }

        if($supAppsInfo)
        {
            foreach($suppApp in ($supAppsInfo -split "[|][*][|]"))
            {
                $appName, $appVer, $appId, $appType = $suppApp -split "[|][!][|]"
                if(-not $appName -or -not $appVer)
                {
                    Write-Log "Could not get Name and Version from string: $($PolicyObject.displayName)" 2
                    continue
                }
                $tmpApps = (Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps?`$filter=displayName eq '$appName'" -TokenId $PolicyObject.TokenId).value
                if(-not $tmpApps)
                {
                    Write-Log "No application found with name $appName" 2
                    continue
                }
                $tmpApp = $tmpApps | Where-Object displayVersion -eq $appVer
                if(-not $tmpApp)
                {
                    Write-Log "No $appName application found with version $appVer" 2
                    continue
                }
                elseif(($tmpApp | Measure-Object).Count -gt 1)
                {
                    Write-Log "Multiple $appName application found with version $appVer" 2
                    continue
                }
                Write-Log "Add $appName ($appVer) to Supersedence list"
                $releationShips.relationships += [PSCustomObject]@{
                    "@odata.type" = "#microsoft.graph.mobileAppSupersedence"
                    targetId = $tmpApp.Id
                    supersedenceType = $appType
                }
            }
        }

        if($releationShips.relationships.Count -gt 0)
        {
            $json = Update-JsonForEnvironment (ConvertTo-Json $releationShips -Depth 20) $PolicyObject $PolicyObject.TokenId

            Write-Log "Update app references"
            Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$($PolicyObject.Id)/updateRelationships" -Method "POST" -Body $json
        }
    }
}

#########################################################################################
#
# iOS LOB App Provisioning Configurations
#
#########################################################################################
#
# Apple-issued provisioning profiles (.mobileprovision files) that travel
# alongside iOS LOB apps. Without these, signed LOB apps stop launching when
# the embedded profile expires. Endpoint at
# /deviceAppManagement/iosLobAppProvisioningConfigurations.
#
# `payload` (Edm.Binary) carries the base64-encoded .mobileprovision file;
# it round-trips through JSON export/import as the payload string.

# region IosLobAppProvisioningConfigurationsType
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class IosLobAppProvisioningConfigurationsType : IntunePolicyTypeBase
{
    IosLobAppProvisioningConfigurationsType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ApplicationsGroup")
        $this._PolicyName  = "iOS app provisioning profiles"
        $this._ID          = "IosLobAppProvisioningConfigurations"
        # Platform default: the endpoint serves exactly one platform and the
        # objects carry no platforms/platformType field, so the column would
        # otherwise be blank (iosLobAppProvisioningConfigurations is iOS-only).
        $this._PlatformName = Get-LanguageString "Platform.iOS" -IgnoreMissing
        $this._API         = "deviceAppManagement/iosLobAppProvisioningConfigurations"
        # No dedicated icon yet — fall back to Applications. Tracked in TODO
        # under the icons-for-new-APIs entry.
        $this._Icon        = "Applications"
        $this._Permissions = @("DeviceManagementApps.ReadWrite.All")
        # version + expirationDateTime are derived from the embedded
        # .mobileprovision; createdDateTime / lastModifiedDateTime are
        # server-set. Strip on POST/PATCH.
        $this._PropertiesToRemove          = @('version','expirationDateTime')
        $this._PropertiesToRemoveForUpdate = @('version','expirationDateTime','payload','payloadFileName')
        # Default `assignments` shape with simple {target} — no overrides
        # needed beyond the inherited defaults.
        $this._ImportOrder = 90
        $this._ObjectClass = "IosLobAppProvisioningConfigurationObject"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class IosLobAppProvisioningConfigurationObject : IntunePolicyBase
{
    IosLobAppProvisioningConfigurationObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    IosLobAppProvisioningConfigurationObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "IosLobAppProvisioningConfigurationsType")
        $this._PlatformName = Get-LanguageString "Platform.iOS" -IgnoreMissing
        if(-not $this._PlatformName) { $this._PlatformName = "iOS/iPadOS" }
    }
}
