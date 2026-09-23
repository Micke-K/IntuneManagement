#ImportOrder 210

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class TenantAdminGroup : IntunePolicyGroupBase
{
    TenantAdminGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "TenantAdmin"
        # Assignment filters live here and are what people come to this group for:
        # their kind and platform are worth the blanks on tags, roles and categories.
        $this._ExtraColumns = @("FilterType=Filter Type")
        $this._ShowPlatformColumn = $true
        $this._Name = "Tenant administration"
        $this._Icon = "TenantSettings"
    }
}

#########################################################################################
#
# Scope Tags
#
#########################################################################################

# region Scope Tags
class ScopeTagsType : IntunePolicyTypeBase
{
    ScopeTagsType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "TenantAdminGroup")
        $this._PolicyName = "Scope Tags"
        $this._ID = "ScopeTags"
        $this._HasPlatform = $false
        $this._HasModified = $false
        $this._API = "deviceManagement/roleScopeTags"
        $this._QueryList = "?`$filter=isBuiltIn%20eq%20false"
        $this._Permissions = @("DeviceManagementRBAC.ReadWrite.All")
        $this._ImportOrder = 10
        $this._ExpandAssignmentsList = $false
        $this._ObjectClass = "ScopeTagObject"
        $this._Icon = "TenantSettings"
        #!!! ToDo: DocumentAll = $true

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        # Hydration already fanned out /assignments via the
        # _HasSubResourceBatch contract on ScopeTagObject — skip the
        # per-policy round-trip the helper would otherwise make.
        if($script:_skipDirectGet -eq $true) { return }
        Add-GraphAssignmentsToExportFile $PolicyObject $PathToFile
    }
}

Class ScopeTagObject : IntunePolicyBase
{
    ScopeTagObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    ScopeTagObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "ScopeTagsType")

        # Opt into the sub-resource batching contract so Invoke-PolicyHydrate
        # fans out the assignments fetch via $batch.
        $this._HasSubResourceBatch = $true
    }

    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        return @([PSCustomObject]@{
            Key = 'assignments'
            Url = "$($this._PolicyType.API)/$($this.Id)/assignments"
        })
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -eq 1 -and $Key -eq 'assignments') {
            # Comma-prefix forces an array reference through the if-expression —
            # without it, PowerShell unwraps a single-element @() on assignment
            # and ConvertTo-Json then emits a bare object instead of [{...}].
            $assignments = if($Body -and $Body.value) { ,@($Body.value) } else { ,@() }
            Add-Member -InputObject $this.JsonObject -MemberType NoteProperty -Name 'assignments' -Value $assignments -Force
        }
        return @()
    }
}


#########################################################################################
#
# Filters
#
#########################################################################################

# region Filters
class FiltersType : IntunePolicyTypeBase
{
    FiltersType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "TenantAdminGroup")
        $this._PolicyName = "Filters"
        $this._ID = "AssignmentFilters"
        $this._SubTypeColumn = "FilterType=Filter Type"   # same header as the group view
        $this._API = "deviceManagement/assignmentFilters"
        # This endpoint answers HTTP 400 to any $filter (verified 2026-08-27),
        # so name searches filter client-side instead.
        $this._SupportsNameFilter = $false
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._ImportOrder = 15
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false
        $this._PropertiesToRemoveForUpdate = @('platform')
        $this._PropertiesToRemove = @("payloads")
        $this._ScopeTagProperty = "roleScopeTags"
        $this._ObjectClass = "FilterObject"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class FilterObject : IntunePolicyBase
{
    Hidden [String]$_FilterType = $null

    FilterObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    FilterObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        if($this.Object.platform -like "*MobileApplicationManagement") {
            $this._FilterType = "Managed apps"
        }
        else {
            $this._FilterType = "Managed devices"
        }

        if($this.Object.platform -like "windows*") {
            $platformLng = "windows10"
        }
        else {
            if($this.Object.platform -like "*MobileApplicationManagement") {
                $platformLng = $this.Object.platform -replace "MobileApplicationManagement", ""
            }
            else {
                $platformLng = $this.Object.platform
            }
            
        }

        Add-ObjectProperty $this "FilterType" { $this._FilterType }

        $this._PlatformName = $this._PlatformName = Get-LanguageString "Platform.$($platformLng )"
        $this._PolicyType = (Get-SingletonObject "FiltersType")
    }
}

#########################################################################################
#
# Role Definitions
#
#########################################################################################

# region Role Definitions
class RoleDefinitionType : IntunePolicyTypeBase
{
    RoleDefinitionType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "TenantAdminGroup")
        $this._PolicyName = "Role Definitions"
        $this._ID = "RoleDefinitions"
        $this._HasPlatform = $false
        $this._HasModified = $false
        $this._API = "deviceManagement/roleDefinitions"
        $this._QueryList = "?`$filter=isBuiltIn%20eq%20false"
        $this._Permissions = @("DeviceManagementRBAC.ReadWrite.All")
        $this._ImportOrder = 20
        $this._ExpandAssignments = $false
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false
        $this._ObjectClass = "RoleDefinitionObject"
        # 'permissions' is the legacy mirror of 'rolePermissions' - PATCHing
        # both makes Graph union them, duplicating every action in the list.
        $this._PropertiesToRemoveForUpdate = @('isBuiltInRoleDefinition','isBuiltIn','roleAssignments','permissions') ### !!! ToDo: Add support for roleAssignments

        #!!! ToDo: DocumentAll = $true

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        if(-not $PathToFile) { return }
        if($script:_skipDirectGet -eq $true) { return }

        $folder = [IO.Path]::GetDirectoryName($PathToFile)

        # roleAssignments were enriched in place by RoleDefinitionObject's
        # sub-resource contract during hydration (and already written to the file
        # by ExportToFile). Resolve their referenced groups into the migration
        # table for cross-tenant import — no re-fetch, no re-save.
        foreach($roleAssignment in @($PolicyObject.Object.roleAssignments))
        {
            # _TokenId, not _TenantId: the parameter is a token id, and no policy object
            # has a _TenantId (it is TenantId, without the underscore). The typo passed
            # $null, which bound to the parameter default of 0 - "the default token" -
            # so a role definition exported from a non-default tenant resolved its group
            # references against the wrong directory.
            foreach($groupId in @($roleAssignment.resourceScopes)) { Add-GraphMigrationObject $groupId "groups" "Group" $folder $PolicyObject._TokenId }
            foreach($groupId in @($roleAssignment.members)) { Add-GraphMigrationObject $groupId "groups" "Group" $folder $PolicyObject._TokenId }
        }
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        Remove-Property $PolicyObject.Object "RoleAssignments"
        Remove-Property $PolicyObject.Object "RoleAssignments@odata.context"
        return $null
    }

    PostImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        if($PolicyObject.TenantId -eq $SourceObject.TenantId) { return }

        $dependencyObjects = Get-GraphDependencySourceObjects $PolicyObject
        $loadedScopeTags = $dependencyObjects["ScopeTags"]
        if(($SourceObject.Object.RoleAssignments | Measure-Object).Count -gt 0 -and ($loadedScopeTags | Measure-Object).Count -gt 0)
        {
            # Documentation way did not work so use the same way as the portal
            # Should be created with /deviceManagement/roleDefinitions/{roleDefinitionId}/roleAssignments
            foreach($roleAssignment in $SourceObject.Object.RoleAssignments)
            {
                $roleAssignmentObj = New-object PSObject @{ 
                    "description" = $roleAssignment.Description
                    "displayName"= $roleAssignment.DisplayName
                    "members" = $roleAssignment.members
                    "resourceScopes" = $roleAssignment.resourceScopes
                    "roleDefinition@odata.bind" = "https://$(Get-GraphDomain $PolicyObject._TokenId)/beta/deviceManagement/roleDefinitions('$($PolicyObject.Id)')"
                    "roleScopeTags@odata.bind" = @()
                }
    
                foreach($scopeTag in $roleAssignment.roleScopeTags)
                {
                    Get-GraphTranslatedDependencyObject $scopeTag.Id $SourceObject $PolicyObject
                    $scopeMigObj = $loadedScopeTags | Where-Object OriginalId -eq $scopeTag.Id
                    if(-not $scopeMigObj.Id) { continue }                
                    $roleAssignmentObj."roleScopeTags@odata.bind" += "https://$(Get-GraphDomain $PolicyObject._TokenId)/beta/deviceManagement/roleScopeTags('$($scopeMigObj.Id)')"
                }
    
                # This will update GroupIds
                $json = Update-JsonForEnvironment (ConvertTo-Json $roleAssignmentObj -Depth 20) $PolicyObject $PolicyObject._TokenId
    
                Write-Log "Import Role Assignments"
                Invoke-MSGraphAPI -Url "deviceManagement/roleAssignments" -Body $json -Method "POST"
            }
        }  
    }

}

Class RoleDefinitionObject : IntunePolicyBase
{
    # Collects enriched assignments across the (unordered) batch responses so
    # FinalizeSubResources can replace the id-only refs in one shot.
    Hidden [Hashtable]$_SubResourceState = $null

    RoleDefinitionObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    RoleDefinitionObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "RoleDefinitionType")

        # Opt into the sub-resource batching contract so Invoke-PolicyHydrate
        # fans out the per-assignment fetch via $batch.
        $this._HasSubResourceBatch = $true
    }

    # The roleAssignments $expand returns id-only refs; enrich each with the full
    # assignment + its roleScopeTags. The deviceManagement/roleAssignments/<id>
    # API lives ONLY here — was previously duplicated in
    # Sync-BulkExportRoleAssignmentDetails (Internal/PolicyHydrateExtras.ps1),
    # RoleDefinitionType.PostExportCommand, and RoleDefinitionDocHandler.
    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        $reqs = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach($roleAssignment in @($this.Object.roleAssignments)) {
            if(-not $roleAssignment.Id) { continue }
            [void]$reqs.Add([PSCustomObject]@{
                Key = "roleasn_$($roleAssignment.Id)"
                Url = "deviceManagement/roleAssignments/$($roleAssignment.Id)?`$expand=microsoft.graph.deviceAndAppManagementRoleAssignment/roleScopeTags"
            })
        }
        return $reqs.ToArray()
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -eq 1 -and $Key -like 'roleasn_*' -and $Body) {
            if($null -eq $this._SubResourceState) { $this._SubResourceState = @{} }
            if(-not $this._SubResourceState.ContainsKey('enriched')) {
                $this._SubResourceState['enriched'] = [System.Collections.Generic.List[object]]::new()
            }
            [void]$this._SubResourceState['enriched'].Add($Body)
        }
        return @()
    }

    # Replace the id-only refs with the enriched assignments, keeping the
    # lowercase `roleAssignments` property the export file, PostExportCommand,
    # PostImportCommand, and the doc handler all read (case-insensitively).
    [void] FinalizeSubResources()
    {
        if($null -ne $this._SubResourceState -and $this._SubResourceState.ContainsKey('enriched')) {
            Add-Member -InputObject $this.JsonObject -MemberType NoteProperty -Name 'roleAssignments' -Value $this._SubResourceState['enriched'].ToArray() -Force
        }
    }
}


#########################################################################################
#
# Intune Branding
#
#########################################################################################

# region Intune Branding
class IntuneBrandingType : IntunePolicyTypeBase
{
    IntuneBrandingType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "TenantAdminGroup")
        $this._PolicyName = "Intune Branding"
        $this._ObjectClass = "IntuneBrandingObject"
        $this._ID = "IntuneBranding"
        $this._HasPlatform = $false
        $this._HasModified = $false
        $this._API = "deviceManagement/intuneBrandingProfiles"
        $this._Permissions = @("DeviceManagementApps.ReadWrite.All")
        $this._NameProperty = "profileName"
        # Name is profileName (see _NameProperty); the flag is a raw JSON property.
        $this._ExtraColumns = @("Object.isDefaultProfile=Default")
        $this._SkipRemoveProperties = @('Id')
        $this._PropertiesToRemoveForUpdate = @('isDefaultProfile','disableClientTelemetry')
        $this._Icon = "Branding"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        $ret = @{}
    
        if($PolicyObject.JsonObject.isDefaultProfile)
        {
            $ret.Add("API",($this.API + "/" + $PolicyObject.Id))
            $ret.Add("Method","PATCH") # Default profile always exists so update it
    
            foreach($prop in @("profileName","isDefaultProfile","disableClientTelemetry","profileDescription"))
            {
                Remove-Property $PolicyObject.JsonObject $prop
            }
    
            $ret
        }
        else
        {
            # Create new Branding profile does not support images data in the json 
            # Workaround: (as done by the portal)
            # Create a new profile with basic info
            # Patch the profile with all the info
    
            foreach($prop in ($PolicyObject.JsonObject.PSObject.Properties | Where-Object {$_.Name -notin @("profileName","profileDescription","roleScopeTagIds")})) #"customPrivacyMessage"
            {
                Remove-Property $PolicyObject.JsonObject $prop.Name
            }
        }
        Remove-Property $PolicyObject.JsonObject "Id"

        return $null
    }

    PostImportCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        if($PolicyObject.JsonObject.isDefaultProfile) { return }

        foreach($prop in @("Id","isDefaultProfile","customPrivacyMessage","disableClientTelemetry")) #"isDefaultProfile","disableClientTelemetry"
        {
            Remove-Property $SourceObject.JsonObject $prop
        }
        $json = ($SourceObject.JsonObject | ConvertTo-Json -Depth 20)
        Invoke-MSGraphAPI -Url "$($this.API)/$($PolicyObject.Id)" -Body $json -Method "PATCH" | Out-Null
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        $fi = [IO.FileInfo]$PathToFile
        foreach($imgType in @("themeColorLogo","lightBackgroundLogo","landingPageCustomizedImage"))
        {
            if($PolicyObject.JsonObject.$imgType.Value)
            {
                $fileName = [IO.Path]::Combine($fi.DirectoryName, "$($PolicyObject.Name)_$imgType.jpg") 
                [IO.File]::WriteAllBytes($fileName, [System.Convert]::FromBase64String($PolicyObject.JsonObject.$imgType.Value))
            }
        }
    }

    [Hashtable]PreDeleteCommand([IntunePolicyBase]$PolicyObject)
    {
        if($PolicyObject.JsonObject.isDefaultProfile -eq $true)
        {
            return @{ "Delete" = $false }
        }

        return $null
    }

    [Hashtable]PreUpdateCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        if($SourceObject.Object.isDefaultProfile)
        {
            foreach($prop in @("profileName","isDefaultProfile","disableClientTelemetry","profileDescription"))
            {
                Remove-Property $PolicyObject.JsonObject $prop
            }
        }

        return $null
    }
}

Class IntuneBrandingObject : IntunePolicyBase
{
    IntuneBrandingObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }
    IntuneBrandingObject() : Base()
    {
        $this.Init()
    }

    Hidden static [String[]]$_ImageProperties = @('themeColorLogo','lightBackgroundLogo','landingPageCustomizedImage')

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "IntuneBrandingType")

        # The body endpoint returns image refs without data; hydration fetches
        # each image via the sub-resource contract (coalesced into the hydrate
        # $batch). Owns the API in one place — was previously duplicated in
        # Sync-BulkExportBrandingImages (Internal/PolicyHydrateExtras.ps1).
        $this._HasSubResourceBatch = $true
    }

    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        $reqs = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach($imageType in [IntuneBrandingObject]::_ImageProperties) {
            [void]$reqs.Add([PSCustomObject]@{
                Key     = $imageType
                Url     = "$($this._PolicyType.API)/$($this.Id)/$imageType"
                Headers = @{ Accept = 'application/json;odata.metadata=none' }
            })
        }
        return $reqs.ToArray()
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -eq 1 -and $Body -and $Body.Value -and $Key -in [IntuneBrandingObject]::_ImageProperties) {
            Add-Member -InputObject $this.JsonObject -MemberType NoteProperty -Name $Key -Value $Body -Force
        }
        return @()
    }
}

#########################################################################################
#
# Device Categories
#
#########################################################################################

class DeviceCategoriesType : IntunePolicyTypeBase
{
    DeviceCategoriesType() : Base() { $this.Init() }
    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "TenantAdminGroup")
        $this._PolicyName = "Device Categories"
        $this._ID = "DeviceCategories"
        $this._HasPlatform = $false
        $this._HasModified = $false
        $this._API = "deviceManagement/deviceCategories"
        $this._QueryList = "?`$top=500"
        $this._Permissions = @("DeviceManagementConfiguration.ReadWrite.All")
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false
        $this._ObjectClass = "DeviceCategoriesObject"
        $this._Icon = "TenantSettings"
        if($null -ne $this._PolicyGroup) { $this._PolicyGroup.AddPolicyType($this) }
    }
}

class DeviceCategoriesObject : IntunePolicyBase
{
    DeviceCategoriesObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }
    DeviceCategoriesObject() : Base() { $this.Init() }
    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "DeviceCategoriesType")
    }
}

#endregion
