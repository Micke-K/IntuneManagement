# CSV output provider.
#
# Consumes the per-object documentation result (see DocumentationOutputJson.ps1
# header for the field list). Writes one CSV file per object under
# <RootFolder>[/<Organization>][/<ObjectType.Id>]/<ObjectName>.csv.
#
# Headless: explicit public options override persisted Documentation settings.
# UI form construction belongs to the active UI backend; this file no longer
# imports XAML at module load. See [[architecture-rules]] R1/R2.
#
# Settings consumed:
#   CSVExportProperties         simple | extended | custom        default 'simple'
#   CSVCustomDisplayProperties  comma-separated property names    default 'Name,Value,Category'
#   CSVDelimiter                CSV delimiter character           default '' (auto)
#   CSVDocumentationPath        root folder for export            default ''
#   CSVAddObjectType            $true to nest under ObjectType    default $true
#   CSVAddCompanyName           $true to nest under Organization  default $false

function Invoke-InitializeCSVOutput {
    Add-DocumentationOutputProvider ([PSCustomObject]@{
        Name              = "CSV"
        Value             = "csv"
        # UI hints — where does this provider store its primary output path, and
        # is that path a folder (CSV writes many files) or a file?
        PrimaryPathOption = "CSVDocumentationPath"
        PathIsFolder      = $true
        PreProcess        = { Invoke-CSVPreProcessItems @args }
        Process    = { Invoke-CSVProcessItem @args }
    })
}

function Invoke-CSVPreProcessItems {
    # No-op. Settings are the source of truth; the UI form (when wired) writes
    # them via Save-SettingStoreValue directly. Hook retained for symmetry / future use.
}

function Invoke-CSVProcessItem {
    param($PolicyObject, $documentedObj)

    if (-not $documentedObj -or -not $PolicyObject) { return }

    $rootFolder     = Get-DocumentationOutputOption csv "CSVDocumentationPath" ""
    $addObjectType  = (Get-DocumentationOutputOption csv "CSVAddObjectType"  $true)  -eq $true
    $addCompanyName = (Get-DocumentationOutputOption csv "CSVAddCompanyName" $false) -eq $true
    $folder = Get-DocObjectFolder -RootFolder $rootFolder -PolicyType $PolicyObject.PolicyType -AddObjectType:$addObjectType -AddOrganization:$addCompanyName

    $objName = $PolicyObject.Name

    try {
        if (-not [IO.Directory]::Exists($folder)) {
            [IO.Directory]::CreateDirectory($folder) | Out-Null
        }

        $mode        = Get-DocumentationOutputOption csv "CSVExportProperties"        "simple"
        $customProps = Get-DocumentationOutputOption csv "CSVCustomDisplayProperties" "Name,Value,Category"
        $delimiter   = Get-DocumentationOutputOption csv "CSVDelimiter"               ""

        $csvParams = @{}
        if ($delimiter) { $csvParams['Delimiter'] = $delimiter }

        $itemsToExport = @()

        $useSectioned = ($mode -eq 'extended' -and $documentedObj.DisplayProperties) -or
                        ($mode -eq 'custom'   -and $customProps)

        if ($useSectioned) {
            if (($documentedObj.BasicInfo | Measure-Object).Count -gt 0) {
                $itemsToExport += ""
                $itemsToExport += "# Basic info"
                $itemsToExport += ""
                $itemsToExport += $documentedObj.BasicInfo | ConvertTo-Csv -NoTypeInformation @csvParams
            }

            if (($documentedObj.FilteredSettings | Measure-Object).Count -gt 0) {
                $itemsToExport += ""
                $itemsToExport += "# Settings"
                $itemsToExport += ""
                if ($mode -eq 'extended') {
                    $displayProperties = $documentedObj.DisplayProperties
                }
                else {
                    $displayProperties = $customProps.Split(",")
                }
                $itemsToExport += $documentedObj.FilteredSettings | Select-Object $displayProperties | ConvertTo-Csv -NoTypeInformation @csvParams
            }

            if (($documentedObj.ApplicabilityRules | Measure-Object).Count -gt 0) {
                $itemsToExport += ""
                $itemsToExport += "# Applicability Rules"
                $itemsToExport += ""
                $itemsToExport += $documentedObj.ApplicabilityRules | Select-Object Rule, Property, Value, Category | ConvertTo-Csv -NoTypeInformation @csvParams
            }

            if (($documentedObj.ComplianceActions | Measure-Object).Count -gt 0) {
                $itemsToExport += ""
                $itemsToExport += "# Compliance Actions"
                $itemsToExport += ""
                $itemsToExport += $documentedObj.ComplianceActions | Select-Object Action, Schedule, MessageTemplate, EmailCC, Category | ConvertTo-Csv -NoTypeInformation @csvParams
            }

            if (($documentedObj.Assignments | Measure-Object).Count -gt 0) {
                if     ($documentedObj.Assignments[0].RawIntent) { $properties = @("GroupMode","Group","Category","SubCategory") }
                elseif ($documentedObj.Assignments[0].Group)     { $properties = @("GroupMode","Group","Category") }
                else                                              { $properties = @("GroupMode","Groups","Category") }

                $itemsToExport += ""
                $itemsToExport += "# Assignments"
                $itemsToExport += ""
                $itemsToExport += $documentedObj.Assignments | Select-Object $properties | ConvertTo-Csv -NoTypeInformation @csvParams
            }
        }
        else {
            $rows = @()
            $rows += $documentedObj.BasicInfo
            $rows += $documentedObj.FilteredSettings
            $itemsToExport = $rows | Select-Object Name, Value | ConvertTo-Csv -NoTypeInformation @csvParams
        }

        $safeName = Remove-InvalidFileNameChars $objName
        $fileName = Join-Path $folder "$safeName.csv"
        Write-Log "Save documentation to $fileName"
        $itemsToExport | Out-File -LiteralPath $fileName -Encoding utf8 -Force
    }
    catch {
        Write-LogError "Failed to save CSV file for $objName in $folder" $_.Exception
    }
}

Invoke-InitializeCSVOutput
