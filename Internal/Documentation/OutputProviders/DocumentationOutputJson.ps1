# JSON output provider.
#
# Consumes the per-object documentation result:
#   BasicInfo          [{Name, Value}]
#   FilteredSettings   [{Name, Value, Category?, SubCategory?}]
#   ComplianceActions  [{Action, Schedule, MessageTemplate, EmailCC}]
#   ApplicabilityRules [{Rule, Property, Value}]
#   CustomTables       [{Values[], Columns[], LanguageId?, Order}]
#   Assignments        [{Group, GroupMode?, Filter?, FilterMode?, Settings?, RawIntent?}]
#   Scripts            [{ScriptContent, Caption}]   (pre-filtered by engine per options)

function Invoke-InitializeJsonOutput {
    Add-DocumentationOutputProvider ([PSCustomObject]@{
        Name              = "Json"
        Value             = "json"
        # Path metadata (see DocumentationOutputHTML.ps1 header comment).
        # Json exposes no file-browse buttons on the bulk-doc form today;
        # add UI/<backend>/ClassExtensions/DocumentationOutputJsonUIExtension.ps1
        # if that changes.
        PrimaryPathOption = "JSONDocumentName"
        PathIsFolder      = $false
        PreProcess        = { Invoke-JsonPreProcessItems @args }
        NewObjectGroup = { Invoke-JsonNewObjectGroup @args }
        NewObjectType  = { Invoke-JsonNewObjectType @args }
        Process        = { Invoke-JsonProcessItem @args }
        PostProcess    = { Invoke-JsonPostProcessItems @args }
    })
}

function Invoke-JsonPreProcessItems {
    $script:jsonAllObjects         = [System.Collections.Generic.List[object]]::new()
    $script:jsonCurrentTypeObjects = [System.Collections.Generic.List[object]]::new()
    $script:jsonCurrentTypeName    = $null

    $script:jsonOutputType = Get-DocumentationOutputOption json "JSONOutputFileType" "Full"

    $jsonFileName = Get-DocumentationOutputOption json "JSONDocumentName" ""
    if (-not $jsonFileName) { $jsonFileName = "%MyDocuments%\%Organization%-%Date%.json" }

    $script:jsonOutFile      = Expand-FileName $jsonFileName
    $script:jsonDocumentPath = [IO.Path]::GetDirectoryName($script:jsonOutFile)
}

function Invoke-JsonNewObjectGroup {
    param($groupId)
    # Groups are not used in flat JSON output
}

function Invoke-JsonNewObjectType {
    param($objectTypeName)

    if ($script:jsonOutputType -eq "ObjectType" -and
        $script:jsonCurrentTypeName -and
        $script:jsonCurrentTypeObjects.Count -gt 0) {
        Save-JsonTypeFile $script:jsonCurrentTypeName $script:jsonCurrentTypeObjects
    }

    $script:jsonCurrentTypeName    = $objectTypeName
    $script:jsonCurrentTypeObjects = [System.Collections.Generic.List[object]]::new()
}

function Invoke-JsonProcessItem {
    param($PolicyObject, $documentedObj)

    if (-not $documentedObj -or -not $PolicyObject) { return }

    $objName = $PolicyObject.Name
    $typeTitle = $PolicyObject.PolicyType.Title

    try {
        $jsonObj = [ordered]@{
            objectType = $typeTitle
            name       = $objName
        }

        if (($documentedObj.BasicInfo | Measure-Object).Count -gt 0) {
            $basicInfo = [ordered]@{}
            foreach ($item in $documentedObj.BasicInfo) {
                if ($item.Name) { $basicInfo[$item.Name] = $item.Value }
            }
            $jsonObj.basicInfo = $basicInfo
        }

        if (($documentedObj.FilteredSettings | Measure-Object).Count -gt 0) {
            $settings = [System.Collections.Generic.List[object]]::new()
            foreach ($item in $documentedObj.FilteredSettings) {
                $setting = [ordered]@{ name = $item.Name; value = $item.Value }
                if ($item.Category)    { $setting.category    = $item.Category }
                if ($item.SubCategory) { $setting.subCategory = $item.SubCategory }
                if ($item.PSObject.Properties['Level'] -and $item.Level) { $setting.level = $item.Level }
                $settings.Add($setting)
            }
            $jsonObj.settings = $settings
        }

        if (($documentedObj.ComplianceActions | Measure-Object).Count -gt 0) {
            $actions = [System.Collections.Generic.List[object]]::new()
            foreach ($item in $documentedObj.ComplianceActions) {
                $actions.Add([ordered]@{
                    action          = $item.Action
                    schedule        = $item.Schedule
                    messageTemplate = $item.MessageTemplate
                    emailCC         = $item.EmailCC
                })
            }
            $jsonObj.complianceActions = $actions
        }

        if (($documentedObj.ApplicabilityRules | Measure-Object).Count -gt 0) {
            $rules = [System.Collections.Generic.List[object]]::new()
            foreach ($item in $documentedObj.ApplicabilityRules) {
                $rules.Add([ordered]@{
                    rule     = $item.Rule
                    property = $item.Property
                    value    = $item.Value
                })
            }
            $jsonObj.applicabilityRules = $rules
        }

        foreach ($customTable in ($documentedObj.CustomTables | Sort-Object -Property Order)) {
            if (-not $customTable.Values -or ($customTable.Values | Measure-Object).Count -eq 0) { continue }

            $tableKey = if ($customTable.LanguageId) { $customTable.LanguageId.Split('.')[-1] } else { "customTable" }
            $tableArr = [System.Collections.Generic.List[object]]::new()

            foreach ($item in $customTable.Values) {
                $tableObj = [ordered]@{}
                foreach ($col in $customTable.Columns) {
                    $colName = $col.Split('.')[-1]
                    $tableObj[$colName] = "$($item.$colName)"
                }
                $tableArr.Add($tableObj)
            }
            $jsonObj[$tableKey] = $tableArr
        }

        if (($documentedObj.Assignments | Measure-Object).Count -gt 0) {
            $assignments = [System.Collections.Generic.List[object]]::new()
            $hasRawIntent = $null -ne $documentedObj.Assignments[0].RawIntent

            foreach ($item in $documentedObj.Assignments) {
                if ($hasRawIntent) {
                    $assignObj = [ordered]@{
                        groupMode = $item.GroupMode
                        group     = $item.Group
                    }
                    if ($null -ne $item.Filter)     { $assignObj.filter     = $item.Filter }
                    if ($null -ne $item.FilterMode) { $assignObj.filterMode = $item.FilterMode }
                    if ($item.Settings) {
                        $settingsObj = [ordered]@{}
                        foreach ($key in $item.Settings.Keys) {
                            if ($key -in @("Category","RawIntent")) { continue }
                            $settingsObj[$key] = $item.Settings[$key]
                        }
                        $assignObj.settings = $settingsObj
                    }
                }
                else {
                    $assignObj = [ordered]@{ group = $item.Group }
                    if ($item.PSObject.Properties.Name -contains "Filter")     { $assignObj.filter     = $item.Filter }
                    if ($item.PSObject.Properties.Name -contains "FilterMode") { $assignObj.filterMode = $item.FilterMode }
                }
                $assignments.Add($assignObj)
            }
            $jsonObj.assignments = $assignments
        }

        if (($documentedObj.Scripts | Measure-Object).Count -gt 0) {
            $scripts = [System.Collections.Generic.List[object]]::new()
            foreach ($scriptItem in $documentedObj.Scripts) {
                if (-not $scriptItem.ScriptContent) { continue }
                $scripts.Add([ordered]@{
                    caption = $scriptItem.Caption
                    content = $scriptItem.ScriptContent
                })
            }
            if ($scripts.Count -gt 0) { $jsonObj.scripts = $scripts }
        }

        $script:jsonCurrentTypeObjects.Add($jsonObj)
        if ($script:jsonOutputType -ne "ObjectType") { $script:jsonAllObjects.Add($jsonObj) }
    }
    catch {
        Write-LogError "Failed to process object $objName" $_.Exception
    }
}

function Invoke-JsonPostProcessItems {
    $openFile = (Get-DocumentationOutputOption json "JSONOpenFile" $true) -eq $true

    if ($script:jsonOutputType -eq "ObjectType") {
        if ($script:jsonCurrentTypeName -and $script:jsonCurrentTypeObjects.Count -gt 0) {
            Save-JsonTypeFile $script:jsonCurrentTypeName $script:jsonCurrentTypeObjects
        }
        Write-Log "Json documentation saved to folder: $($script:jsonDocumentPath)"
    }
    else {
        $jsonContent = ConvertTo-Json -InputObject @($script:jsonAllObjects) -Depth 20
        Save-DocumentationFile $jsonContent $script:jsonOutFile -OpenFile:$openFile
    }
}

function Save-JsonTypeFile {
    param($typeName, $objects)

    $safeTypeName = Remove-InvalidFileNameChars ($typeName.Replace(" ", "_"))
    $typeFileName = [IO.Path]::Combine($script:jsonDocumentPath, "$safeTypeName.json")
    $jsonContent  = ConvertTo-Json -InputObject @($objects) -Depth 20
    Save-DocumentationFile $jsonContent $typeFileName
    Write-Log "Saved $($objects.Count) objects to $typeFileName"
}

Invoke-InitializeJsonOutput
