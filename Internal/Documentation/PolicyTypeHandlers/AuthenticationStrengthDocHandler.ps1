# Authentication Strength documentation handler.
#
# Claims @odata.type='#microsoft.graph.authenticationStrengthPolicy'. Authentication
# strengths were previously only referenced as a Conditional Access grant control
# (ID -> displayName, see ConditionalAccessDocHandler); this handler documents the
# standalone policy object.
#
# The policy's substance is allowedCombinations: an OR-list of method combinations,
# where each combination is an AND-set of authentication methods (comma-joined in
# the raw value, e.g. "password,microsoftAuthenticatorPush"). Each method maps to
# AzureCA.AuthenticationStrength.Mode.<method>. A few methods (federatedMultiFactor,
# federatedSingleFactor) have no Mode string yet, so fall back to a humanised token
# rather than emit a raw enum value.

class AuthenticationStrengthDocHandler : DocumentationHandlerBase {
    AuthenticationStrengthDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.authenticationStrengthPolicy')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        # BasicInfo: Name + Profile type + Description
        $nameValue = if ($obj.displayName) { $obj.displayName } else { $PolicyObject.Name }
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.nameName') $nameValue 'displayName'
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.configurationType') (Get-LanguageString 'AzureCA.WhatIfBlade.authenticationStrength') '@odata.type'
        if ($obj.description) {
            Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.descriptionName') $obj.description 'description'
        }

        # allowedCombinations: OR-list of AND-combinations. Render each combination
        # as "Method A + Method B" on its own line.
        $comboLines = @()
        foreach ($combo in @($obj.allowedCombinations)) {
            if ([string]::IsNullOrWhiteSpace($combo)) { continue }
            $methodNames = @()
            foreach ($method in ($combo -split ',')) {
                $m = $method.Trim()
                if (-not $m) { continue }
                $methodNames += (Get-AuthenticationMethodLabel $m)
            }
            if ($methodNames.Count -gt 0) {
                $comboLines += ($methodNames -join ' + ')
            }
        }

        if ($comboLines.Count -gt 0) {
            Add-CustomSettingObject ([PSCustomObject]@{
                Name        = (Get-LanguageString 'AzureCA.policyControlAuthenticationStrengthDisplayedName')
                Value       = ($comboLines -join $Context.ObjectSeparator)
                Category    = $null
                SubCategory = $null
                EntityKey   = 'allowedCombinations'
            })
        }
    }
}

# Map an authentication-method enum token to its localised label, falling back to
# a humanised form (federatedMultiFactor -> "Federated Multi Factor") for tokens
# that have no Mode string. -IgnoreMissing keeps the log clean for known gaps.
function Get-AuthenticationMethodLabel {
    param([string]$Method)

    if ([string]::IsNullOrEmpty($Method)) { return $Method }

    $label = Get-LanguageString "AzureCA.AuthenticationStrength.Mode.$Method" -IgnoreMissing
    if (-not [string]::IsNullOrEmpty($label)) { return $label }

    # No Mode string: split camelCase into Title-cased words.
    $spaced = [regex]::Replace($Method, '(?<=[a-z0-9])(?=[A-Z])', ' ')
    $ci = [System.Globalization.CultureInfo]::InvariantCulture
    return (($spaced -split ' ' | Where-Object { $_ } | ForEach-Object { $ci.TextInfo.ToTitleCase($_.ToLower()) }) -join ' ')
}

[DocumentationRegistry]::RegisterHandler([AuthenticationStrengthDocHandler]::new())
