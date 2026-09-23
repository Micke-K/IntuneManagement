function Get-GraphBulkScopeTagPatchUrl
{
    param($Policy)

    $objectClass = $null
    try { $objectClass = [string]$Policy._objectClass } catch { }

    if(-not [string]::IsNullOrWhiteSpace($objectClass)) {
        return "deviceAppManagement/$objectClass/$($Policy.Id)"
    }

    return "$($Policy.PolicyType.API)/$($Policy.Id)"
}

# Adjust an imported object's scope-tag IDs based on the ImportScopeTags setting
# and same-/cross-tenant detection (ported from the old project's Set-ScopeTags):
#   - ImportScopeTags off          -> reset to the Default scope tag ("0")
#   - on, same tenant              -> keep the source IDs as-is
#   - on, cross tenant             -> remap each ID via the loaded ScopeTags
#                                     dependency objects (OriginalId -> Id); keep
#                                     "0"; reset to Default if nothing resolves
# Uses the per-type scope-tag property (PolicyType.ScopeTagProperty:
# roleScopeTagIds / roleScopeTags / $null when the type has no scope tags).
# Called from IntuneBaseClasses.ImportObject before the body is serialized.
function Set-GraphImportScopeTags
{
    param($PolicyObject, [int]$TokenId)

    $scopeTagProperty = [string]$PolicyObject.PolicyType.ScopeTagProperty
    if(-not $scopeTagProperty) { return }

    $obj = $PolicyObject.JsonObject
    if(-not ($obj.PSObject.Properties | Where-Object Name -eq $scopeTagProperty)) { return }

    $scopeIds = @()
    if((Get-SettingValue "ImportScopeTags") -eq $true)
    {
        # Get-OperationTokenInfo, not Get-TokenInfo: ImportObject passes its own
        # TokenId straight through, and 0 - the caller's spelling for "the default
        # token" - means "no filter, every token" to Get-TokenInfo. With a second
        # tenant signed in, $tokenInfo.TenantId was an ARRAY, which stringifies to
        # "tenant-a tenant-b" and never equals the source tenant: a SAME-tenant
        # import was classified as cross-tenant, took the remap branch, found no
        # destination scope tags to map to and reset the policy to the Default scope
        # tag. Silent loss of the scope tags the policy was imported with.
        $tokenInfo      = Get-OperationTokenInfo $TokenId
        $sourceTenantId = $PolicyObject._ClonedFromObject.TenantId
        $crossTenant    = ($tokenInfo -and $sourceTenantId -and $sourceTenantId -ne $tokenInfo.TenantId)

        if(-not $crossTenant)
        {
            # Same tenant: keep the source scope-tag IDs.
            $scopeIds += $obj.$scopeTagProperty
        }
        else
        {
            # Cross tenant: remap via the loaded ScopeTags dependency objects.
            $usingDefault = (@($obj.$scopeTagProperty).Count -eq 1 -and "$(@($obj.$scopeTagProperty)[0])" -eq "0")
            if(-not $usingDefault)
            {
                $loadedScopeTags = (Get-GraphDependencySourceObjects $PolicyObject)["ScopeTags"]
                foreach($scopeId in $obj.$scopeTagProperty)
                {
                    if("$scopeId" -eq "0") { $scopeIds += "0"; continue }
                    $scopeMigObj = $loadedScopeTags | Where-Object OriginalId -eq $scopeId
                    if($scopeMigObj -and $scopeMigObj.Id) { $scopeIds += "$($scopeMigObj.Id)" }
                    elseif($scopeMigObj) { Write-Log "Could not find a destination ScopeTag for '$($scopeMigObj.Name)'. Make sure all ScopeTags are imported into the environment" 2 }
                }
            }
        }
    }

    # Default scope tag when nothing else applies (off / cross-tenant unresolved / already-default).
    if(@($scopeIds).Count -eq 0) { $scopeIds += "0" }

    $obj.$scopeTagProperty = @($scopeIds)
}

# Loads + normalizes the tenant scope-tag catalog. Cache lookup first
# (DependencyObjects_<TenantId>) with a live Get-GraphPolicies fallback.
# Always seeds Default (Id 0) when the live path is taken since the API
# does not return it. Returns a sorted list of [PSCustomObject]@{ Id; Name }
# de-duped on Id. Shared by both the WPF and Avalonia bulk-scope-tag dialogs.
function Get-BulkScopeTagCatalog
{
    # $LoadError (optional [ref]): when the live fetch fails this function still
    # returns the Default-only fallback and writes the failure message into
    # $LoadError so the UI layer can present it. Keeps this data function UI-free
    # (architecture R2) - it no longer pops a dialog of its own.
    param([int]$TokenId = (Get-DefaultTokenId), [ref]$LoadError)

    $allTags = @()
    $tokenInfo = Get-TokenInfo $TokenId
    if($tokenInfo -and $tokenInfo.TenantId) {
        $dependencyObjects = Get-CacheObject "DependencyObjects_$($tokenInfo.TenantId)" @{}
        if($dependencyObjects -and $dependencyObjects.ContainsKey("ScopeTags")) {
            $allTags = @($dependencyObjects["ScopeTags"])
        }
    }

    if($allTags.Count -eq 0) {
        $allTags = @([PSCustomObject]@{ Id = 0; Name = "Default" })
        try {
            $allTags += @(Get-GraphPolicies -PolicyType "ScopeTags" -TokenId $TokenId -ErrorAction Stop)
        }
        catch {
            Write-LogError "Bulk Scope Tags: failed to load scope tag catalogue" $_.Exception
            if($LoadError) { $LoadError.Value = $_.Exception.Message }
        }
    }

    $normalizedTags = @()
    $seenIds = [System.Collections.Generic.HashSet[string]]::new()
    foreach($tag in $allTags) {
        $id = if($null -ne $tag.Id) { [string]$tag.Id } elseif($null -ne $tag.ID) { [string]$tag.ID } else { $null }
        if([string]::IsNullOrWhiteSpace($id)) { continue }
        if(-not $seenIds.Add($id)) { continue }
        $normalizedTags += [PSCustomObject]@{ Id = $id; Name = [string]$tag.Name }
    }

    return @($normalizedTags | Sort-Object Name)
}
