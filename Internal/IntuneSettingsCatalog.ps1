function Get-IntuneSettingsCatalogClasses
{
    [CmdLetbinding()]
    param()

    $settingCatalogClasses = Get-CacheObject "IntuneSettingsCatalogClasses"
    if(-not $settingCatalogClasses) {
        $settingCatalogClasses = @()
        $settingCatalogClasses += Get-SubClasses "SettingsCatalogTypeBase" | ForEach-Object {
            try { Get-SingletonObject $_.Name } catch {}
        }
        Set-CacheObject "IntuneSettingsCatalogClasses" -Value $settingCatalogClasses
    }

    return $settingCatalogClasses
} 

function Set-SettingsCatalogQueryFilter
{
    param($Class, [string[]]$Families)

    # Dedupe the MERGED list, not just the input: a second run (a re-login, or a
    # test re-seeding the template cache) otherwise appends every family again
    # and the filter grows an 'x or x' clause per pass.
    $Class._FamilyTypes = @(@($Class._FamilyTypes) + @($Families) | Where-Object { $_ } | Select-Object -Unique)
    if($Class._FamilyTypes.Count -gt 0) {
        $clauses = $Class._FamilyTypes | ForEach-Object { "templateReference/templateFamily eq '$_'" }
        $filter  = "?`$filter=" + ($clauses -join ' or ')
        Write-Log "$($Class.GetType().Name) settings catalog query filter: $filter"
        $Class._QueryList = $filter
    }
    else {
        # The tenant reports no template for any family this class claims. Its
        # startup filter names that family, and the service rejects a
        # templateFamily value it does not know yet with 400 (live-verified for
        # 'maintenanceWindows' before Microsoft rolled it out) - so every listing
        # would log a batch failure. Replace it with a filter that is always valid
        # and matches nothing.
        $Class._QueryList = "?`$filter=id eq '00000000-0000-0000-0000-000000000000'"
        Write-Log "$($Class.GetType().Name): tenant has no template for its family - listing disabled until one appears"
    }
}

function Invoke-IntuneSettingsCatalogAuthenticated
{
    [CmdLetbinding()]
    param($TokenInfo)

    if($script:_intuneSettingsCatalogInitialized) { return }

    $templates = Get-CacheObject "IntuneSettingsCatalogTemplates"
    if(-not $templates) {
        $templates = (Invoke-MSGraphAPI "deviceManagement/configurationPolicyTemplates").value
        Set-CacheObject "IntuneSettingsCatalogTemplates" -Value $templates
    }

    $settingCatalogClasses = Get-IntuneSettingsCatalogClasses

    # One descriptor per catalog class, evaluated IN ORDER. The last entry
    # (SettingsCatalogType) is the catch-all: its selector claims every family not
    # already taken, so it reads $familyTypes and must stay last. Selectors run via
    # Where-Object and resolve $familyTypes from this function's scope at match time.
    $familyTypes = @()
    $specs = @(
        [PSCustomObject]@{ Type = [EnrollmentSettingsCatalogType];       Extra = @();       Select = { $_.templateFamily -eq 'enrollmentConfiguration' } }
        [PSCustomObject]@{ Type = [EndpointSecuritySettingsCatalogType]; Extra = @();       Select = { $_.templateFamily -like 'endpointSecurity*' -or $_.templateFamily -eq 'baseline' } }
        [PSCustomObject]@{ Type = [ScriptSettingsCatalogType];           Extra = @();       Select = { $_.templateFamily -eq 'deviceConfigurationScripts' } }
        [PSCustomObject]@{ Type = [WindowsUpdateSettingsCatalogType];    Extra = @();       Select = { $_.templateFamily -eq 'maintenanceWindows' } }
        # This must be the last in the list, as it is the catch-all for any remaining families not already claimed by the other types.
        [PSCustomObject]@{ Type = [SettingsCatalogType];                 Extra = @('none'); Select = { $_.templateFamily -notin $familyTypes } }
    )

    foreach($spec in $specs) {
        $class = $settingCatalogClasses | Where-Object { $_ -is $spec.Type }
        if(-not $class) { continue }

        $selected = @($templates | Where-Object $spec.Select | ForEach-Object { $_.templateFamily } | Select-Object -Unique)
        Set-SettingsCatalogQueryFilter -Class $class -Families (@($spec.Extra) + $selected)

        $familyTypes += $class._FamilyTypes
    }

    $script:_intuneSettingsCatalogInitialized = $true
}

Add-AppEventHandler "AuthenticatedNewToken" "Invoke-IntuneSettingsCatalogAuthenticated"
#Add-AppEventHandler "AuthenticationTokenRefresh" "Invoke-IntuneSettingsCatalogAuthenticated"