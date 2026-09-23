# Named Location documentation handler (Country + IP variants).
#
# Ported from old Extensions/DocumentationCustom.psm1:2280 (country) + 2323 (IP).
# Single class claims both @odata.types since they share the same BasicInfo
# header shape and one varies only the settings.

class NamedLocationDocHandler : DocumentationHandlerBase {
    NamedLocationDocHandler() {
        $this.ODataTypes = @(
            '#microsoft.graph.countryNamedLocation',
            '#microsoft.graph.ipNamedLocation',
            '#microsoft.graph.compliantNetworkNamedLocation'
        )
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        Add-BasicDefaultValues $PolicyObject
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'AzureCA.menuItemNamedNetworks') '@odata.type'
        Add-BasicAdditionalValues $PolicyObject
        switch ($obj.'@odata.type') {
            '#microsoft.graph.countryNamedLocation'          { Invoke-NamedLocationCountrySettings          $obj $Context }
            '#microsoft.graph.ipNamedLocation'               { Invoke-NamedLocationIPSettings               $obj $Context }
            '#microsoft.graph.compliantNetworkNamedLocation' { Invoke-NamedLocationCompliantNetworkSettings $obj $Context }
        }
    }
}

function Invoke-NamedLocationCountrySettings {
    param($obj, [DocumentationContext]$Context)

    $lookupSuffix = if ($obj.countryLookupMethod -eq 'clientIpAddress') { 'ip' } else { 'gps' }
    Add-CustomSettingObject ([PSCustomObject]@{
        Name      = (Get-LanguageString 'AzureCA.NamedLocation.Form.CountryLookup.ariaLabel')
        Value     = (Get-LanguageString "AzureCA.NamedLocation.Form.CountryLookup.$lookupSuffix")
        EntityKey = 'countryLookupMethod'
    })

    $includeKey = if ($obj.includeUnknownCountriesAndRegions -eq $true) { 'Inputs.enabled' } else { 'Inputs.disabled' }
    Add-CustomSettingObject ([PSCustomObject]@{
        Name      = (Get-LanguageString 'AzureCA.NamedLocation.Form.Include.label')
        Value     = (Get-LanguageString $includeKey)
        EntityKey = 'includeUnknownCountriesAndRegions'
    })

    $countryNames = @()
    foreach ($code in $obj.countriesAndRegions) {
        $countryNames += Get-LanguageString "AzureIAMCommon.CountryNames.countryName$($code.ToUpper())"
    }
    Add-CustomSettingObject ([PSCustomObject]@{
        Name      = (Get-LanguageString 'AzureCA.NamedLocation.Type.countries')
        Value     = ($countryNames -join $Context.ObjectSeparator)
        EntityKey = 'countriesAndRegions'
    })
}

function Invoke-NamedLocationIPSettings {
    param($obj, [DocumentationContext]$Context)

    $trustedKey = if ($obj.isTrusted -eq $true) { 'Inputs.enabled' } else { 'Inputs.disabled' }
    Add-CustomSettingObject ([PSCustomObject]@{
        Name      = (Get-LanguageString 'AzureCA.NamedLocation.Form.Trusted.label')
        Value     = (Get-LanguageString $trustedKey)
        EntityKey = 'isTrusted'
    })

    $ipList = @()
    foreach ($range in $obj.ipRanges) { $ipList += $range.cidrAddress }
    Add-CustomSettingObject ([PSCustomObject]@{
        Name      = (Get-LanguageString 'AzureCA.namedNetworkIpRangesTab')
        Value     = ($ipList -join $Context.ObjectSeparator)
        EntityKey = 'ipRanges'
    })
}

function Invoke-NamedLocationCompliantNetworkSettings {
    param($obj, [DocumentationContext]$Context)

    # Built-in read-only location; its only meaningful setting is the trusted flag.
    $trustedKey = if ($obj.isTrusted -eq $true) { 'Inputs.enabled' } else { 'Inputs.disabled' }
    Add-CustomSettingObject ([PSCustomObject]@{
        Name      = (Get-LanguageString 'AzureCA.NamedLocation.Form.Trusted.label')
        Value     = (Get-LanguageString $trustedKey)
        EntityKey = 'isTrusted'
    })
}

[DocumentationRegistry]::RegisterHandler([NamedLocationDocHandler]::new())
