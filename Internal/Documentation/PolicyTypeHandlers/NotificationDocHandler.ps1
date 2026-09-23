# Notification message template documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:3660. Emits 4 branding-
# option rows ("Show company logo/name/contact/portal link" enable/disable)
# plus one row per localized message template with the locale name as label
# and the subject+body as value.

class NotificationDocHandler : DocumentationHandlerBase {
    NotificationDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.notificationMessageTemplate')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        Add-BasicDefaultValues $PolicyObject
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'Titles.notifications') '@odata.type'

        $category = Get-LanguageString 'TableHeaders.settings'

        # brandingOptions is a comma-separated string like "includeCompanyLogo,includeCompanyName"
        # or "none". Split into a hash for membership tests.
        $brandingFlags = @{}
        if ($obj.brandingOptions) {
            foreach ($flag in $obj.brandingOptions.Split(',')) {
                $brandingFlags[$flag.Trim()] = $true
            }
        }

        $brandingLabelMap = [ordered]@{
            'includeCompanyLogo'        = 'NotificationMessage.companyLogo'
            'includeCompanyName'        = 'NotificationMessage.companyName'
            'includeContactInformation' = 'NotificationMessage.companyContact'
            'includeCompanyPortalLink'  = 'NotificationMessage.iwLink'
            'includeDeviceDetails'      = 'NotificationMessage.deviceDetails'
        }

        foreach ($flag in $brandingLabelMap.Keys) {
            $valueKey = if ($brandingFlags.ContainsKey($flag)) { 'BooleanActions.enable' } else { 'BooleanActions.disable' }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name        = (Get-LanguageString $brandingLabelMap[$flag])
                Value       = (Get-LanguageString $valueKey)
                EntityKey   = $flag
                Category    = $category
                SubCategory = $null
            })
        }

        # Localized message templates
        $subCategory = Get-LanguageString 'NotificationMessage.listTitle'
        foreach ($template in $obj.localizedNotificationMessages) {
            $label = Get-NotificationLocaleLabel $template.locale
            if (-not $label) { continue }

            $value = $template.subject
            if ($template.isDefault) {
                $value = $value + $Context.ObjectSeparator + (Get-LanguageString 'NotificationMessage.isDefaultLocale') + ': ' + (Get-LanguageString 'SettingDetails.trueOption')
            }
            $fullValue = $value + $Context.ObjectSeparator + $template.messageTemplate

            Add-CustomSettingObject ([PSCustomObject]@{
                Name        = $label
                Value       = $fullValue
                EntityKey   = $template.locale
                Category    = $category
                SubCategory = $subCategory
            })
        }
    }
}

# Localized message templates are labeled by language name. Most languages
# pass through [cultureinfo].EnglishName.ToLower(); a handful with regional
# splits (en-US/UK, es-ES/MX, fr-CA/FR, pt-PT/BR, zh-TW/CN, nb-* -> norwegian)
# get a suffix. Old code at DocumentationCustom.psm1:3735-3796.
function Get-NotificationLocaleLabel {
    param([string]$Locale)
    if (-not $Locale) { return $null }

    $first, $second = $Locale.Split('-')
    try { $lng = ([cultureinfo]$first).EnglishName.ToLower() } catch { return $null }

    switch ($first) {
        'en' { switch ($second) { 'US' { $lng += 'US' }; 'GB' { $lng += 'UK' } } }
        'es' { switch ($second) { 'es' { $lng += 'Spain' }; 'mx' { $lng += 'Mexico' } } }
        'fr' { switch ($second) { 'ca' { $lng += 'Canada' }; 'fr' { $lng += 'France' } } }
        'pt' { switch ($second) { 'pt' { $lng += 'Portugal' }; 'br' { $lng += 'Brazil' } } }
        'zh' { switch ($second) { 'tw' { $lng += 'Traditional' }; 'cn' { $lng += 'Simplified' } } }
        'nb' { $lng = 'norwegian' }
    }

    return (Get-LanguageString "NotificationMessage.NotificationMessageTemplatesTab.$lng")
}

[DocumentationRegistry]::RegisterHandler([NotificationDocHandler]::new())
