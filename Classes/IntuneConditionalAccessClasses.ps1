#ImportOrder 220

#########################################################################################
#
# Apple Update Group
#
#########################################################################################

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class ConditionalAccessGroup : IntunePolicyGroupBase
{
    ConditionalAccessGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "ConditionalAccess"
        $this._Name = "Conditional Access"
        $this._Icon = "ConditionalAccess"
    }
}

#########################################################################################
#
# Conditional Access Policies
#
#########################################################################################

# region Conditional Access Policies
class ConditionalAccessType : IntunePolicyTypeBase
{
    ConditionalAccessType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ConditionalAccessGroup")
        $this._PolicyName = "Conditional Access"
        $this._ID = "ConditionalAccess"
        $this._HasPlatform = $false
        $this._API = "identity/conditionalAccess/policies"
        # Entra object - no roleScopeTagIds in the Graph schema (a PATCH
        # no-ops), so no scope-tag support.
        $this._ScopeTagProperty = ""
        $this._Dependencies = @("NamedLocations","Applications","TermsOfUse","AuthenticationStrengths","AssignmentFilters")
        $this._Permissions = @("Policy.ReadWrite.ConditionalAccess")
        $this._ObjectClass = "ConditionalAccessObject"
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false
        $this._HasPageSizeSupport = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        # Tenant-lockout guard: rewrite the imported policy's state per the
        # ConditionalAccessState setting (default: disabled). Logic lives in
        # Internal/IntuneManager.ps1 so it is unit-testable.
        Set-CAPolicyImportState $PolicyObject

        if($PolicyObject.grantControls.authenticationStrength)
        {
            $PolicyObject.JsonObject.grantControls.operator = "AND"
            #$tmpObj = Get-GraphObjectFromFile $file

            #$authSetting = [PSCustomObject]@{
            #    id = $tmpObj.grantControls.authenticationStrength.id
            #}
            #$PolicyObject.JsonObject.grantControls.authenticationStrength = $authSetting
        }

        if($PolicyObject.JsonObject.sessionControls.disableResilienceDefaults -eq $false)
        {
            $PolicyObject.JsonObject.sessionControls.disableResilienceDefaults = $null
        }

        return $null
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        $ids = @()
        foreach($id in ($PolicyObject.JsonObject.conditions.users.includeGroups + $PolicyObject.JsonObject.conditions.users.excludeGroups))
        {
            if($id -in $ids) { continue }
            elseif($id -eq "GuestsOrExternalUsers") { continue }
            elseif($id -eq "All") { continue }
            elseif($id -eq "None") { continue }
            
            $ids += $id            
            Add-GraphMigrationObject $id "groups" "Group" ([IO.Path]::GetDirectoryName($PathToFile)) $PolicyObject._TokenId
        }
        
        foreach($id in ($PolicyObject.JsonObject.conditions.users.includeUsers + $PolicyObject.JsonObject.conditions.users.excludeUsers))
        {
            if($id -in $ids) { continue }
            elseif($id -eq "GuestsOrExternalUsers") { continue }
            elseif($id -eq "All") { continue }
            elseif($id -eq "None") { continue }
            
            $ids += $id
            Add-GraphMigrationObject $id "users" "User" ([IO.Path]::GetDirectoryName($PathToFile)) $PolicyObject._TokenId
        }  
    }
}

Class ConditionalAccessObject : IntunePolicyBase
{
    ConditionalAccessObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    ConditionalAccessObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "ConditionalAccessType")
    }
}

#########################################################################################
#
# Authentication Strengths
#
#########################################################################################

# region Authentication Strengths
class AuthenticationStrengthsType : IntunePolicyTypeBase
{
    AuthenticationStrengthsType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ConditionalAccessGroup")
        $this._PolicyName = "Authentication Strengths"
        $this._ID = "AuthenticationStrengths"
        $this._HasPlatform = $false
        $this._API = "identity/conditionalAccess/authenticationStrengths/policies"
        $this._ImportOrder = 45        
        $this._Permissions = @("Policy.ReadWrite.ConditionalAccess")
        $this._ObjectClass = "AuthenticationStrengthObject"
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false
        $this._TopItems = 0
        $this._Icon = "ConditionalAccess"
        $this._HasPageSizeSupport = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        if($PolicyObject.Object.policyType -ne "custom")
        {
            Write-Log "Built-in Authentication Strength objects cannot be imported" 2
            @{ "Import" = $false }
        }

        return $null
    }
}

Class AuthenticationStrengthObject : IntunePolicyBase
{
    AuthenticationStrengthObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AuthenticationStrengthObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AuthenticationStrengthsType")
    }
}

#########################################################################################
#
# Authentication Context
#
#########################################################################################

# region Authentication Context
class AuthenticationContextType : IntunePolicyTypeBase
{
    AuthenticationContextType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ConditionalAccessGroup")
        $this._PolicyName = "Authentication Context"
        $this._ID = "AuthenticationContext"
        $this._HasPlatform = $false
        $this._HasModified = $false
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "identity/conditionalAccess/authenticationContextClassReferences"
        $this._PropertiesToRemove = @("@odata.type")
        $this._SkipRemoveProperties = @('Id')
        $this._ImportOrder = 46
        $this._Permissions = @("Policy.ReadWrite.ConditionalAccess")
        $this._ObjectClass = "AuthenticationContextObject"
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false
        $this._TopItems = 0
        $this._Icon = "ConditionalAccess"
        $this._HasPageSizeSupport = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class AuthenticationContextObject : IntunePolicyBase
{
    AuthenticationContextObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    AuthenticationContextObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AuthenticationContextType")
    }
}

#########################################################################################
#
# Authentication Context
#
#########################################################################################

# region Authentication Context
class NamedLocationType : IntunePolicyTypeBase
{
    NamedLocationType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ConditionalAccessGroup")
        $this._PolicyName = "Named Locations"
        $this._ID = "NamedLocations"
        $this._HasPlatform = $false
        $this._API = "identity/conditionalAccess/namedLocations"
        # Entra object - no roleScopeTagIds in the Graph schema.
        $this._ScopeTagProperty = ""
        $this._ImportOrder = 50
        $this._Permissions = @("Policy.ReadWrite.ConditionalAccess")
        $this._ObjectClass = "NamedLocationObject"
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false
        $this._HasPageSizeSupport = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class NamedLocationObject : IntunePolicyBase
{
    NamedLocationObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    NamedLocationObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "NamedLocationType")
    }
}


#########################################################################################
#
# Terms of use
#
#########################################################################################

# region Terms of use
class TermsOfUseType : IntunePolicyTypeBase
{
    TermsOfUseType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ConditionalAccessGroup")
        $this._PolicyName = "Terms of use"
        $this._ID = "TermsOfUse"
        $this._HasPlatform = $false
        $this._HasModified = $false
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "identityGovernance/termsOfUse/agreements"
        # Entra object - no roleScopeTagIds in the Graph schema.
        $this._ScopeTagProperty = ""
        $this._ImportOrder = 75
        $this._Expand = "files"
        $this._QueryList = "?`$expand=files"
        $this._Permissions = @("Agreement.ReadWrite.All")
        $this._ObjectClass = "TermsOfUseObject"
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        $pkgPath = Get-SettingValue "IntuneAppPackagesFolder"

        if(-not $pkgPath -or [IO.Directory]::Exists($pkgPath) -eq $false) 
        {
            Write-Log "Intune app directory is either missing or does not exist" 2        
        }

        # The agreement document is resolved in this order, per localization:
        #   1. fileData.data already on the object - an export carries the PDF
        #      inline, fetched per localization by the sub-resource contract.
        #   2. <fileName> next to the exported json (FileInfo), then in the app
        #      packages folder - for a json that was exported without the data.
        #   3. The source object, for an in-memory COPY (below).
        # Refusing rather than proceeding matters: an agreement created without
        # its document lists fine but /file, /files and /file/localizations all
        # 404, and a later list with $expand=files returns 500 for the WHOLE
        # collection. Three of those were found in the test tenant on 2026-09-06
        # and had to be deleted by hand.
        #
        # An earlier version of this block required the PDF on disk whenever
        # FileInfo was set and ignored the embedded data. That only held together
        # because Clone() used to drop FileInfo, so an import from disk never
        # reached it with FileInfo populated. Once Clone() kept FileInfo, every
        # import of an export with an inline PDF was refused.
        $hasData = { param($f) ($f.PSObject.Properties['fileData'] -and $f.fileData -and
                                $f.fileData.PSObject.Properties['data'] -and $f.fileData.data) }

        if($PolicyObject.FileInfo) {
            foreach($file in $PolicyObject.Object.Files)
            {
                if(& $hasData $file) { continue }

                $pdfFile = $null
                if($PolicyObject.FileInfo.Directory.FullName)
                {
                    $pdfFile = [IO.Path]::Combine($PolicyObject.FileInfo.Directory.FullName, "$($file.fileName)")
                }
                if(($null -eq $pdfFile -or [IO.File]::Exists($pdfFile) -eq $false) -and $pkgPath)
                {
                    $pdfFile = [IO.Path]::Combine($pkgPath, "$($file.fileName)")
                }
                if($pdfFile -and [IO.File]::Exists($pdfFile))
                {
                    Write-Log "Add file data: $pdfFile"
                    $bytes = [IO.File]::ReadAllBytes($pdfFile)
                    $file | Add-Member -MemberType NoteProperty -Name 'fileData' -Value ([PSCustomObject]@{ data = [Convert]::ToBase64String($bytes) }) -Force
                }
                else
                {
                    Write-Log "Terms of use file $($file.fileName) not found next to the export or in the app packages folder" 2
                }
            }
        }

        $touFiles = @($PolicyObject.Object.Files)
        $touMissing = @($touFiles | Where-Object { -not (& $hasData $_) })

        if($touFiles.Count -gt 0 -and $touMissing.Count -gt 0)
        {
            $touSource = $PolicyObject._ClonedFromObject
            if($touSource -and $touSource.Id)
            {
                Write-Log "Terms of use '$($PolicyObject.Name)': agreement file(s) not loaded - fetching them from the source object"
                $touTokenId = 0
                if($touSource.PSObject.Properties['_TokenId'] -and $null -ne $touSource._TokenId) {
                    $touTokenId = [int]$touSource._TokenId
                }
                elseif($PolicyObject.PSObject.Properties['_TokenId'] -and $null -ne $PolicyObject._TokenId) {
                    $touTokenId = [int]$PolicyObject._TokenId
                }
                try { Invoke-PolicySubresourceFetch -Policies @($touSource) -TokenId $touTokenId | Out-Null }
                catch { Write-LogError "Failed to fetch terms of use file data from the source object" $_.Exception }

                foreach($touFile in $touMissing)
                {
                    $srcFile = @($touSource.Object.Files) | Where-Object { $_.id -eq $touFile.id } | Select-Object -First 1
                    if($srcFile -and $srcFile.PSObject.Properties['fileData'] -and $srcFile.fileData -and $srcFile.fileData.data)
                    {
                        $touFile | Add-Member -MemberType NoteProperty -Name 'fileData' -Value ([PSCustomObject]@{ data = $srcFile.fileData.data }) -Force
                    }
                }

                $touMissing = @($touFiles | Where-Object { -not (& $hasData $_) })
            }
        }

        if($touFiles.Count -eq 0 -or $touMissing.Count -gt 0)
        {
            Write-Log "Terms of use '$($PolicyObject.Name)': the agreement document is missing and could not be loaded from the source. The object will not be imported - creating it would leave an agreement with no document, which breaks the whole Terms of Use list." 2
            return @{"Import" = $false}
        }
        return $null
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        if(-not $PathToFile) { return }
        $fi = [IO.FileInfo]$PathToFile
        # File binary was fetched into fileData.data by TermsOfUseObject's
        # sub-resource contract during hydration. Write each to disk; no re-fetch.
        foreach($file in @($PolicyObject.Object.Files))
        {
            $data = $null
            if($file.PSObject.Properties['fileData'] -and $file.fileData -and $file.fileData.PSObject.Properties['data']) {
                $data = $file.fileData.data
            }
            if($data)
            {
                Write-Log "Save file $($file.FileName)"
                $fileName = [IO.Path]::Combine($fi.DirectoryName, "$($file.FileName)")
                [IO.File]::WriteAllBytes($fileName, [System.Convert]::FromBase64String($data))
            }
        }
    }
}

Class TermsOfUseObject : IntunePolicyBase
{
    # Maps each in-flight file-data request key back to its file object so
    # ApplySubResourceBatchResult can attach the fetched binary.
    Hidden [Hashtable]$_SubResourceState = $null

    TermsOfUseObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    TermsOfUseObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "TermsOfUseType")

        # Agreement file binaries aren't in the body; fetch each localization's
        # fileData via the sub-resource contract.
        $this._HasSubResourceBatch = $true
    }

    # The agreements/<id>/file/localizations('<fid>')/fileData/data API lives
    # ONLY here — was previously duplicated in Sync-BulkExportTermsOfUseFiles
    # (Internal/PolicyHydrateExtras.ps1) and TermsOfUseType.PostExportCommand.
    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        $this._SubResourceState = @{}
        $reqs = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach($file in @($this.Object.Files)) {
            if(-not $file.id) { continue }
            $existing = $null
            if($file.PSObject.Properties['fileData'] -and $file.fileData -and $file.fileData.PSObject.Properties['data']) {
                $existing = $file.fileData.data
            }
            if($existing) { continue }
            $key = "toufile_$($file.id)"
            $this._SubResourceState[$key] = $file
            [void]$reqs.Add([PSCustomObject]@{
                Key = $key
                Url = "agreements/$($this.Id)/file/localizations('$($file.id)')/fileData/data"
            })
        }
        return $reqs.ToArray()
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -ne 1 -or $null -eq $this._SubResourceState -or -not $this._SubResourceState.ContainsKey($Key)) { return @() }
        $file = $this._SubResourceState[$Key]

        $data = $null
        if($Body -is [string]) { $data = $Body }
        elseif($Body -and $Body.PSObject.Properties['value']) { $data = $Body.value }
        elseif($Body -and $Body.PSObject.Properties['data'])  { $data = $Body.data }

        if($data) {
            if($file.PSObject.Properties['fileData'] -and $file.fileData) {
                if($file.fileData.PSObject.Properties['data']) { $file.fileData.data = $data }
                else { $file.fileData | Add-Member -MemberType NoteProperty -Name 'data' -Value $data -Force }
            }
            else {
                Add-Member -InputObject $file -MemberType NoteProperty -Name 'fileData' -Value ([PSCustomObject]@{ data = $data }) -Force
            }
        }
        return @()
    }
}
