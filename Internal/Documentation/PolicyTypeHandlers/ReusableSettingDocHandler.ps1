# Reusable setting (deviceManagementReusablePolicySetting) documentation
# handler.
#
# Claims @odata.type='#microsoft.graph.deviceManagementReusablePolicySetting'.
# A reusable setting is a name + description around ONE settings-catalog
# setting instance (settingInstance). The portal wizard
# (ReusableSettingsWizard.ReactView.js) reads that instance against the schema
# from deviceManagement/reusableSettings, which is where the display names of
# every child ("Keyword", "Auto Resolve", "Match type", "Device properties", ...)
# and of every choice option ("True", "Match all") live. The generic
# configurationSettings/{id} endpoint answers nothing for these ids (they carry
# {id} / {groupid} placeholders), so the schema is loaded from that list once
# per session and handed to the settings-catalog walker inline.
#
# Rendering: a list of items (GroupSettingCollection whose items hold only
# simple / choice values - firewall dynamic keyword addresses, Device Control
# device properties, printer devices) becomes one row with a table value
# (FullValueTable: one column per named child) plus a flat Value for the
# outputs that cannot show a table. Groups that nest further lists (the Device
# Control group itself) are flattened: their simple values become rows and
# their lists tables. Anything else is a plain Name = Value row, the Linux
# discovery script decoded.
#
# No schema reachable (fully offline): rows and choice values are named from
# the definition ids instead of showing nothing.
#
# The list endpoint omits settingInstance; the single GET with $select carries
# it (same contract as the export hydration in ReusableSettingsObjectBase).
# Documenting straight from the tenant therefore fetches it when the source
# tenant is available, and attaches it to the object so an export of the same
# run is complete.

$script:_docReusableSettingDefinitions = $null

function Reset-ReusableSettingDefinitionCache {
    $script:_docReusableSettingDefinitions = $null
}

# All reusable-setting definitions (roots and children) as Graph returns them,
# loaded once per session. Generic schema, so gated on connectivity only, not
# on SourceTenantUnavailable - documenting an export while signed into another
# tenant still resolves names.
function Get-ReusableSettingDefinitions {
    param([DocumentationContext]$Context)

    if ($null -ne $script:_docReusableSettingDefinitions) { return $script:_docReusableSettingDefinitions }
    if (-not (Test-DocumentationGraphAvailable)) { return @() }

    try {
        $resp = Invoke-MSGraphAPI -Url "/deviceManagement/reusableSettings" -ODataMetadata 'minimal' -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context)
        $script:_docReusableSettingDefinitions = @($resp.value)
    }
    catch {
        Write-LogError "Failed to fetch the reusable setting definitions" $_.Exception
        $script:_docReusableSettingDefinitions = @()
    }
    return $script:_docReusableSettingDefinitions
}

# A readable name from a definition id when the schema did not give one:
# the part after the parent id ("..._{id}_autoresolve" -> "Autoresolve"), else
# the last segment that is not a {placeholder}.
function Get-ReusableSettingFallbackName {
    param([string]$Id, [string]$ParentId)

    if ([string]::IsNullOrEmpty($Id)) { return '' }
    $tail = $Id
    if ($ParentId -and $Id.StartsWith("$ParentId`_", [StringComparison]::OrdinalIgnoreCase) -and $Id.Length -gt $ParentId.Length + 1) {
        $tail = $Id.Substring($ParentId.Length + 1)
    }
    else {
        $segments = @($Id.Split('_') | Where-Object { $_ -and $_ -notmatch '^\{.*\}$' })
        if ($segments.Count -gt 0) { $tail = $segments[-1] }
    }
    $tail = $tail.Replace('_', ' ').Trim()
    if ($tail.Length -eq 0) { return '' }
    return $tail.Substring(0, 1).ToUpperInvariant() + $tail.Substring(1)
}

# Row name: the schema display name; derived from the id only when no schema
# resolved the row at all. A schema row with a BLANK name is one Microsoft
# hides (the GUID id of a firewall address) and stays blank, so callers skip it.
function Get-ReusableSettingRowName {
    param($Row, [string]$ParentId, [object[]]$Definitions)
    $name = [string]$Row.Name
    if (-not [string]::IsNullOrWhiteSpace($name)) { return $name.Trim() }
    if ($Definitions -and $Row.SettingId -and ($Definitions | Where-Object id -EQ $Row.SettingId | Select-Object -First 1)) { return '' }
    return Get-ReusableSettingFallbackName $Row.SettingId $ParentId
}

# Value of a simple / choice row: what the walker resolved, else - for a choice
# whose option the schema did not name - the option suffix ("..._true" -> "True").
function Get-ReusableSettingRowValue {
    param($Row)

    $value = $Row.Value
    if ($null -eq $value -or ([string]$value).Length -eq 0) {
        $raw = [string]$Row.RawValue
        if ($raw -and $Row.SettingId -and $Row.Type -like '*ChoiceSettingInstance') {
            $value = Get-ReusableSettingFallbackName $raw $Row.SettingId
        }
        elseif ($raw) { $value = $raw }
    }
    if ($value -is [string] -and [string]$Row.SettingId -like '*discoveryscript*') {
        # The discovery-script definition stores the script base64-encoded.
        try {
            $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($value))
            if ($decoded) { $value = $decoded }
        }
        catch { }
    }
    return $value
}

# True when a walker node is a list or group container.
function Test-ReusableSettingGroupNode {
    param($Row)
    return ([string]$Row.Type -like '*GroupSettingCollectionInstance' -or [string]$Row.Type -like '*GroupSettingInstance')
}

# Emit the documentation rows for one walker node (and what it contains).
function ConvertTo-ReusableSettingDocRows {
    param(
        $Node,
        [DocumentationContext]$Context,
        [string]$ParentId,
        [int]$Level = 0,
        [object[]]$Definitions = @()
    )

    if (-not $Node) { return }
    $type = [string]$Node.Type

    if ($type -like '*GroupSettingCollectionInstance') {
        $items = @($Node.ChildSettings)
        $nested = $false
        foreach ($item in $items) {
            foreach ($child in @($item.Settings)) {
                if (Test-ReusableSettingGroupNode $child) { $nested = $true; break }
            }
            if ($nested) { break }
        }

        if ($nested -or $items.Count -eq 0) {
            # A group of groups: its simple values become rows and its lists tables.
            foreach ($item in $items) {
                foreach ($child in @($item.Settings)) {
                    ConvertTo-ReusableSettingDocRows -Node $child -Context $Context -ParentId $Node.SettingId -Level $Level -Definitions $Definitions
                }
            }
            return
        }

        # A list of flat items: one table, a column per named child (schema
        # order, first appearance), rows in item order.
        $columns = [System.Collections.Generic.List[string]]::new()
        $tableRows = @()
        $flatItems = @()
        foreach ($item in $items) {
            $cells = [ordered]@{}
            $parts = @()
            foreach ($child in @($item.Settings)) {
                $colName = Get-ReusableSettingRowName $child $Node.SettingId $Definitions
                if ([string]::IsNullOrEmpty($colName)) { continue }   # the unnamed GUID id of a firewall address
                if (-not $columns.Contains($colName)) { [void]$columns.Add($colName) }
                $cellValue = [string](Get-ReusableSettingRowValue $child)
                $cells[$colName] = $cellValue
                if ($cellValue.Length -gt 0) { $parts += "${colName}: $cellValue" }
            }
            $tableRows += ,$cells
            $flatItems += ($parts -join $Context.PropertySeparator)
        }
        $table = foreach ($cells in $tableRows) {
            $rowObj = [ordered]@{}
            foreach ($col in $columns) { $rowObj[$col] = if ($cells.Contains($col)) { $cells[$col] } else { '' } }
            [PSCustomObject]$rowObj
        }

        $label = Get-ReusableSettingRowName $Node $ParentId $Definitions
        # Microsoft named the firewall list root " Id"; the category (Firewall)
        # is the readable label there.
        if ([string]::IsNullOrWhiteSpace($label) -or $label -eq 'Id') { $label = [string]$Node.Category }
        if ([string]::IsNullOrWhiteSpace($label)) { $label = Get-ReusableSettingFallbackName $Node.SettingId $null }

        $Context.AddSetting([PSCustomObject]@{
            Name           = $label
            Value          = ($flatItems -join $Context.ObjectSeparator)
            Description    = $Node.Description
            Category       = $Node.Category
            SubCategory    = $Node.SubCategory
            Level          = $Level
            SettingId      = $Node.SettingId
            RowIndex       = $Node.RowIndex
            RawValue       = $Node.RawJsonValue
            RawJsonValue   = $Node.RawJsonValue
            FullValueTable = @($table)
        })
        return
    }

    if ($type -like '*GroupSettingInstance') {
        foreach ($child in @($Node.ChildSettings)) {
            ConvertTo-ReusableSettingDocRows -Node $child -Context $Context -ParentId $Node.SettingId -Level $Level -Definitions $Definitions
        }
        return
    }

    $rowName = Get-ReusableSettingRowName $Node $ParentId $Definitions
    if ([string]::IsNullOrEmpty($rowName)) { return }   # hidden by the schema
    $Context.AddSetting([PSCustomObject]@{
        Name         = $rowName
        Value        = (Get-ReusableSettingRowValue $Node)
        Description  = $Node.Description
        Category     = $Node.Category
        SubCategory  = $Node.SubCategory
        Level        = $Level
        SettingId    = $Node.SettingId
        RowIndex     = $Node.RowIndex
        RawValue     = $Node.RawValue
        RawJsonValue = $Node.RawJsonValue
        DefaultValue = $Node.DefaultValue
    })

    # A choice can carry dependent children.
    foreach ($child in @($Node.ChildSettings)) {
        if ($child -and $child.PSObject.Properties['SettingId']) {
            ConvertTo-ReusableSettingDocRows -Node $child -Context $Context -ParentId $Node.SettingId -Level ($Level + 1) -Definitions $Definitions
        }
    }
}

# The setting instance: from the object, else from the source tenant (the list
# endpoint leaves it out). Attached to the object so an export of the same run
# carries it.
function Get-ReusableSettingInstance {
    param($PolicyObject, [DocumentationContext]$Context)

    $obj = $PolicyObject.JsonObject
    if ($obj.settingInstance) { return $obj.settingInstance }
    if (-not $obj.id -or $Context.SourceTenantUnavailable -or -not (Test-DocumentationGraphAvailable)) { return $null }

    try {
        # minimal, not none: with none Graph drops the @odata.type of every
        # setting instance, and the walker cannot tell a list from a value.
        $live = Invoke-MSGraphAPI -Url "/deviceManagement/reusablePolicySettings/$($obj.id)?`$select=settinginstance,displayname,description" -ODataMetadata 'minimal' -NoError
        if ($live -and $live.settingInstance) {
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'settingInstance' -Value $live.settingInstance -Force
            return $live.settingInstance
        }
        Write-Log "Reusable setting '$($PolicyObject.Name)' ($($obj.id)) returned no settingInstance" 2
    }
    catch {
        Write-LogError "Failed to fetch the setting instance of reusable setting '$($PolicyObject.Name)'" $_.Exception
    }
    return $null
}

# The configuration-categories list carries only ordinary categories; one that
# exists for reusable settings alone (Device Control) is reachable by id only.
# Fetch it once and add it to the session cache, so both the Type row and the
# walker's row categories resolve.
function Resolve-ReusableSettingCategory {
    param([string]$CategoryId, [DocumentationContext]$Context)

    if ([string]::IsNullOrEmpty($CategoryId)) { return }
    if ($Context.CfgCategories | Where-Object id -EQ $CategoryId | Select-Object -First 1) { return }
    if (-not (Test-DocumentationGraphAvailable)) { return }
    try {
        $cat = Invoke-MSGraphAPI -Url "/deviceManagement/configurationCategories/$CategoryId" -ODataMetadata 'minimal' -AdditionalHeaders (Get-DocAcceptLanguageHeaders $Context) -NoError
        if ($cat -and $cat.id) {
            $Context.CfgCategories += @($cat)
            Set-CacheObject "CfgCategories" $Context.CfgCategories -Persistent
        }
    }
    catch {
        Write-LogError "Failed to fetch configuration category $CategoryId" $_.Exception
    }
}

# What kind of reusable setting this is, for the basic info: the category of
# the root definition (Firewall, Device Control), else the definition's own
# name (Script), else derived from the id.
function Get-ReusableSettingTypeName {
    param($RootDefinition, [string]$DefinitionId, [DocumentationContext]$Context)

    if ($RootDefinition -and $RootDefinition.categoryId) {
        $cat = $Context.CfgCategories | Where-Object id -EQ $RootDefinition.categoryId | Select-Object -First 1
        if ($cat -and -not [string]::IsNullOrWhiteSpace([string]$cat.displayName)) { return [string]$cat.displayName }
    }
    $name = if ($RootDefinition) { ([string]$RootDefinition.displayName).Trim() } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($name) -and $name -ne 'Id') { return $name }
    return Get-ReusableSettingFallbackName $DefinitionId $null
}

function Invoke-DocReusableSetting {
    param($PolicyObject, [DocumentationContext]$Context)

    $obj = $PolicyObject.JsonObject

    if ($PolicyObject.Name) {
        Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.nameName') $PolicyObject.Name 'displayName'
    }
    $descValue = if ($obj.description) { $obj.description } else { '' }
    Add-BasicPropertyValue (Get-LanguageString 'SettingDetails.descriptionName') $descValue 'description'

    Initialize-SettingsCatalogSchemaCache $Context
    $definitions = @(Get-ReusableSettingDefinitions $Context)
    $rootDef = $null
    if ($obj.settingDefinitionId) {
        $rootDef = $definitions | Where-Object id -EQ $obj.settingDefinitionId | Select-Object -First 1
        if ($rootDef) { Resolve-ReusableSettingCategory ([string]$rootDef.categoryId) $Context }
        Add-BasicPropertyValue (Get-LanguageString 'TableHeaders.type') (Get-ReusableSettingTypeName $rootDef ([string]$obj.settingDefinitionId) $Context) 'settingDefinitionId'
    }
    Add-BasicAdditionalValues $PolicyObject   # created / modified / version

    $instance = Get-ReusableSettingInstance $PolicyObject $Context
    if (-not $instance) { return }

    if ($Context.Options -and $Context.Options['IncludeScripts'] -eq $false -and [string]$obj.settingDefinitionId -like '*script*') { return }

    # The walker resolves names, options and categories from the inline
    # definitions; its buffer is not used, the returned tree is.
    Reset-SettingsCatalogPolicyBuffer
    $root = Add-SettingsSetting $instance $definitions
    Reset-SettingsCatalogPolicyBuffer
    if (-not $root) { return }

    ConvertTo-ReusableSettingDocRows -Node $root -Context $Context -ParentId $null -Level 0 -Definitions $definitions
}

class ReusableSettingDocHandler : DocumentationHandlerBase {
    ReusableSettingDocHandler() {
        $this.ODataTypes = @('#microsoft.graph.deviceManagementReusablePolicySetting')
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        Invoke-DocReusableSetting $PolicyObject $Context
    }
}

[DocumentationRegistry]::RegisterHandler([ReusableSettingDocHandler]::new())
