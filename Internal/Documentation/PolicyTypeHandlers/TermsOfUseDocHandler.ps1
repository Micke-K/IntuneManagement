# Terms of Use (agreement) documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:2361. Note this handler
# does NOT use Add-BasicDefaultValues — the old engine only emits Name +
# Profile type for agreements (no Description, no Created/Modified).

class TermsOfUseDocHandler : DocumentationHandlerBase {
    TermsOfUseDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.agreement')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        $offLabel = Get-LanguageString 'SettingDetails.offOption'
        $onLabel  = Get-LanguageString 'SettingDetails.onOption'

        # BasicInfo: just Name + Profile type
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.nameName') $obj.displayName 'displayName'
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'AzureCA.menuItemTermsOfUse') '@odata.type'

        $viewingValue = if ($obj.isViewingBeforeAcceptanceRequired) { $onLabel } else { $offLabel }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString 'TermsOfUse.Wizard.agreementIsViewingBeforeAcceptanceRequiredLabel')
            Value = $viewingValue; Category = $null; SubCategory = $null
            EntityKey = 'isViewingBeforeAcceptanceRequired'
        })

        $perDeviceValue = if ($obj.isPerDeviceAcceptanceRequired) { $onLabel } else { $offLabel }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString 'TermsOfUse.Wizard.agreementIsPerDeviceAcceptanceRequiredLabel')
            Value = $perDeviceValue; Category = $null; SubCategory = $null
            EntityKey = 'isPerDeviceAcceptanceRequired'
        })

        $expirationValue = if ($obj.termsExpiration) { $onLabel } else { $offLabel }
        Add-CustomSettingObject ([PSCustomObject]@{
            Name = (Get-LanguageString 'TermsOfUse.Wizard.isAcceptanceExpirationEnabledLabel')
            Value = $expirationValue; Category = $null; SubCategory = $null
            EntityKey = 'isAcceptanceExpirationEnabledLabel'
        })

        # Expiration details (only when termsExpiration is set)
        if ($obj.termsExpiration.startDateTime) {
            try {
                if ($obj.termsExpiration.startDateTime -is [datetime]) {
                    $tmp = if ($obj.termsExpiration.startDateTime.Kind -eq 'Utc') { $obj.termsExpiration.startDateTime.ToLocalTime() } else { $obj.termsExpiration.startDateTime }
                }
                else {
                    $tmp = ([datetime]::Parse($obj.termsExpiration.startDateTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)).ToLocalTime()
                }
                $startStr = $tmp.ToShortDateString()
            }
            catch {
                Write-Log "Failed to parse date from string $($obj.termsExpiration.startDateTime)" 2
                $startStr = $obj.termsExpiration.startDateTime
            }

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'TermsOfUse.Wizard.acceptanceExpirationStartDateTimeLabel')
                Value = $startStr; Category = $null; SubCategory = $null
                EntityKey = 'startDateTime'
            })

            $freqValue = switch ($obj.termsExpiration.frequency) {
                'P365D' { Get-LanguageString 'TermsOfUse.AcceptanceExpirationFrequency.annually' }
                'P180D' { Get-LanguageString 'TermsOfUse.AcceptanceExpirationFrequency.biannually' }
                'P30D'  { Get-LanguageString 'TermsOfUse.AcceptanceExpirationFrequency.monthly' }
                'P90D'  { Get-LanguageString 'TermsOfUse.AcceptanceExpirationFrequency.quarterly' }
                default { $null }
            }
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'TermsOfUse.Wizard.acceptanceExpirationFrequencyLabel')
                Value = $freqValue; Category = $null; SubCategory = $null
                EntityKey = 'frequency'
            })
        }

        if ($null -ne $obj.userReacceptRequiredFrequency) {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'TermsOfUse.Wizard.acceptanceDurationLabel')
                Value = (Get-DurationValue $obj.userReacceptRequiredFrequency)
                Category = $null; SubCategory = $null
                EntityKey = 'userReacceptRequiredFrequency'
            })
        }
    }
}

[DocumentationRegistry]::RegisterHandler([TermsOfUseDocHandler]::new())
