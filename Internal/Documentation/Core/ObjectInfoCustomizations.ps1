# Shared dispatch for custom behavior used during schema-driven ObjectInfo walks.
# Whole-object handlers are intentionally separate: registering a handler prevents
# Manifest/Profile fallback, while these customizers augment that fallback.

function Add-DocumentationObjectInfoCustomizer {
    param([Parameter(Mandatory)][PSCustomObject]$Customizer)
    [DocumentationRegistry]::RegisterObjectInfoCustomizer($Customizer)
}

function Get-DocumentationObjectInfoType {
    param($Obj)
    if (-not $Obj) { return $null }
    return [string]$Obj.'@odata.type'
}

function Get-DocumentationObjectInfoCustomizers {
    param($Obj)
    $ctx = Get-CurrentDocumentationContext
    $topObj = if ($ctx.CurrentObject) { $ctx.CurrentObject } else { $Obj }
    return @([DocumentationRegistry]::FindObjectInfoCustomizers((Get-DocumentationObjectInfoType $topObj)))
}

function Initialize-DocumentationObjectInfoObject {
    param($Obj)
    $ctx = Get-CurrentDocumentationContext
    foreach ($customizer in (Get-DocumentationObjectInfoCustomizers $Obj)) {
        if ($customizer.InitializeObject) {
            & $customizer.InitializeObject $Obj $ctx
        }
    }
}

function Invoke-DocumentationObjectInfoGetPropertyObject {
    param($Obj, $Prop)
    $ctx = Get-CurrentDocumentationContext
    $topObj = $ctx.CurrentObject
    foreach ($customizer in (Get-DocumentationObjectInfoCustomizers $Obj)) {
        if (-not $customizer.GetPropertyObject) { continue }
        $ret = & $customizer.GetPropertyObject $topObj $Obj $Prop $ctx
        if ($ret) { return $ret }
    }
    return $Obj
}

function Invoke-DocumentationObjectInfoGetChildObject {
    param($Obj, $Prop)
    $ctx = Get-CurrentDocumentationContext
    $topObj = $ctx.CurrentObject
    foreach ($customizer in (Get-DocumentationObjectInfoCustomizers $Obj)) {
        if (-not $customizer.GetChildObject) { continue }
        $ret = & $customizer.GetChildObject $topObj $Obj $Prop $ctx
        if ($ret) { return $ret }
    }
    return $Obj
}

function Invoke-DocumentationObjectInfoGetProfileValue {
    param($Obj, $Prop)
    $ctx = Get-CurrentDocumentationContext
    $topObj = $ctx.CurrentObject
    foreach ($customizer in (Get-DocumentationObjectInfoCustomizers $Obj)) {
        if (-not $customizer.GetProfileValue) { continue }
        $ret = & $customizer.GetProfileValue $topObj $Obj $Prop $ctx
        if ($null -ne $ret) { return $ret }
    }
    return $null
}

function Invoke-DocumentationObjectInfoPostAddValue {
    param($Prop)
    $ctx = Get-CurrentDocumentationContext
    $topObj = $ctx.CurrentObject
    foreach ($customizer in (Get-DocumentationObjectInfoCustomizers $topObj)) {
        if ($customizer.PostAddValue) {
            & $customizer.PostAddValue $topObj $Prop $ctx
        }
    }
}

function Finalize-DocumentationObjectInfoObject {
    param($Obj)
    $ctx = Get-CurrentDocumentationContext
    foreach ($customizer in (Get-DocumentationObjectInfoCustomizers $Obj)) {
        if ($customizer.FinalizeObject) {
            & $customizer.FinalizeObject $Obj $ctx
        }
    }
}

# Translates the id of Reusable Settings in Settings Catalog policies to their display name
function Invoke-DocumentationSettingsCatalogPostProcess {
    param($Obj, [DocumentationContext]$Context)
    if (-not $Obj.templateReference.templateId -or
        -not ([string]$Obj.templateReference.templateId).StartsWith('19c8aa67-f286-4861-9aa0-f23541d31680')) {
        return
    }
    # Reusable settings are resolved by id from the source tenant (source-specific).
    if ($Context.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) {
        return
    }
    foreach ($setting in @($Context.SettingsData | Where-Object SettingId -EQ 'vendor_msft_firewall_mdmstore_firewallrules_{firewallrulename}_remoteaddressdynamickeywords')) {
        if (-not $setting.RawValue) { continue }
        try {
            $reusable = Invoke-MSGraphAPI -Url "/deviceManagement/reusablePolicySettings/$($setting.RawValue)"
            if ($reusable.displayName) { $setting.Value = $reusable.displayName }
            else { Write-Log "No Reusable Settings object found with ID $($setting.RawValue)" 2 }
        }
        catch {
            Write-LogError "Failed to resolve reusable setting $($setting.RawValue)" $_.Exception
        }
    }
}
