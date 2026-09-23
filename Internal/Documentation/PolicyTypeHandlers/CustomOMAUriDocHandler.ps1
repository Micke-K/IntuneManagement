# Custom OMA-URI documentation handler.
#
# Ported from old Extensions/DocumentationCustom.psm1:3519. Emits Name +
# Platform basic info, then 4-5 rows per OMA-URI setting (Name, Description,
# OMA-URI path, Data type, Value). Encrypted values are skipped offline;
# live runs fetch via /deviceConfigurations/.../getOmaSettingPlainTextValue.
#
# Claims all 4 CustomConfiguration variants in one handler.

class CustomOMAUriDocHandler : DocumentationHandlerBase {
    CustomOMAUriDocHandler() {
        $this.ODataTypes = @(
            '#microsoft.graph.windows10CustomConfiguration',
            '#microsoft.graph.androidForWorkCustomConfiguration',
            '#microsoft.graph.androidWorkProfileCustomConfiguration',
            '#microsoft.graph.androidCustomConfiguration'
        )
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        $obj = $PolicyObject.JsonObject

        Add-BasicDefaultValues $PolicyObject
        # Note: old code at L3534 has the configurationType BasicPropertyValue
        # commented out. Faithful port — skipping.

        $platformId = Get-ObjectPlatformFromType $obj
        if ($platformId) {
            Add-BasicPropertyValue (Get-LanguageString 'Inputs.platformLabel') (Get-LanguageString "Platform.$platformId") 'platform'
        }

        $category = Get-LanguageString 'SettingDetails.customPolicyOMAURISettingsName'

        $typeLabelMap = @{
            '#microsoft.graph.omaSettingString'        = 'SettingDetails.stringName'
            '#microsoft.graph.omaSettingBase64'        = 'SettingDetails.base64Name'
            '#microsoft.graph.omaSettingBoolean'       = 'SettingDetails.booleanName'
            '#microsoft.graph.omaSettingDateTime'      = 'SettingDetails.dateTimeName'
            '#microsoft.graph.omaSettingFloatingPoint' = 'SettingDetails.floatingPointName'
            '#microsoft.graph.omaSettingInteger'       = 'SettingDetails.integerName'
            '#microsoft.graph.omaSettingStringXml'     = 'SettingDetails.stringXMLName'
        }

        foreach ($setting in $obj.omaSettings) {
            $sub  = $setting.displayName
            $oma  = $setting.omaUri
            $type = if ($setting.PSObject.Properties['@OData.Type']) { $setting.'@OData.Type' } else { $setting.'@odata.type' }

            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'SettingDetails.nameName')
                Value = $setting.displayName
                EntityKey = "displayName_$oma"
                Category = $category; SubCategory = $sub
            })
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'TableHeaders.description')
                Value = $setting.description
                EntityKey = "description_$oma"
                Category = $category; SubCategory = $sub
            })
            Add-CustomSettingObject ([PSCustomObject]@{
                Name = (Get-LanguageString 'SettingDetails.oMAURIName')
                Value = $oma
                EntityKey = "omaUri_$oma"
                Category = $category; SubCategory = $sub
            })

            $typeKey = $typeLabelMap[$type]
            if ($typeKey) {
                $typeValue = Get-LanguageString $typeKey
                if ($typeValue) {
                    Add-CustomSettingObject ([PSCustomObject]@{
                        Name = (Get-LanguageString 'SettingDetails.dataTypeName')
                        Value = $typeValue
                        EntityKey = "type_$oma"
                        Category = $category; SubCategory = $sub
                    })
                }
            }

            # Value row — skip when encrypted unless we can resolve via Graph
            if ($setting.isEncrypted -ne $true) {
                $value = $setting.value
                if ($type -eq '#microsoft.graph.omaSettingStringXml' -and $value) {
                    try { $value = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($value)) } catch { }
                }
                Add-CustomSettingObject ([PSCustomObject]@{
                    Name = (Get-LanguageString 'SettingDetails.valueName')
                    Value = $value
                    EntityKey = "value_$oma"
                    Category = $category; SubCategory = $sub
                })
            }
            elseif (-not $Context.SourceTenantUnavailable -and (Test-DocumentationGraphAvailable) -and $setting.secretReferenceValueId) {
                try {
                    $url = "/deviceManagement/deviceConfigurations/$($obj.id)/getOmaSettingPlainTextValue(secretReferenceValueId='$($setting.secretReferenceValueId)')"
                    $resp = Invoke-MSGraphAPI -Url $url
                    if ($resp.Value) {
                        Add-CustomSettingObject ([PSCustomObject]@{
                            Name = (Get-LanguageString 'SettingDetails.valueName')
                            Value = $resp.Value
                            EntityKey = "value_$oma"
                            Category = $category; SubCategory = $sub
                        })
                    }
                }
                catch { Write-LogError "Failed to resolve encrypted OMA-URI value for $oma" $_.Exception }
            }
        }
    }
}

[DocumentationRegistry]::RegisterHandler([CustomOMAUriDocHandler]::new())
