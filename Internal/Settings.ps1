# Settings store management.
#
# Settings are three layers, and this file is the middle one:
#
#   storage    Internal/Core.ps1 - Get/Save/Remove-SettingStoreValue. PATH
#              addressed: the caller supplies the SubPath, and whether that path
#              means a registry key, a node in a JSON tree or a node in an
#              in-memory tree is decided here.
#   resolver   this file - KEY addressed. Get-SettingValue (in Core.ps1, because
#              Write-Log needs it during preload) and its write-side counterparts
#              look the key up in the registered definitions and derive the
#              SubPath from it, so no caller has to know a path.
#   public     Public/*-IMSetting*.ps1 - the exported api.
#
# The store has three modes, resolved once at the top of Internal/Core.ps1:
#
#   Registry   HKCU:\Software\IntuneManagement. Windows only.
#   Json       a settings file: IM_SETTINGS_FILE if set, else
#              LocalApplicationData/IntuneManagement/Settings.json. The default
#              off Windows.
#   Memory     a settings tree with no file behind it. Reads and writes work
#              normally and nothing touches disk. For automation (a runbook on a
#              shared worker must not write to that worker's HKCU) and for tests.
#
# Memory mode is not a fourth code path: the storage primitives take their JSON
# branch whenever a settings object exists, and persist only when a settings FILE
# exists as well. Memory mode is an object with no file - see the condition in
# Save-SettingStoreValue.
#
# See Docs/Settings.md.

# IM_SETTINGS_STORE was set to something unrecognized - Core.ps1 fell back to the
# platform default, silently, because logging does not work that early. Say so now
# rather than have a runbook believe it opted out of disk writes.
if($script:SettingsStoreModeEnvRequest -and
   $script:SettingsStoreModeEnvRequest -notin @("Memory", "Json", "Registry"))
{
    Write-Log "IM_SETTINGS_STORE='$script:SettingsStoreModeEnvRequest' is not a valid store (Memory, Json, Registry). Using '$script:SettingsStoreMode'." 2
}

# What the store actually is right now. Derived from the state variables rather
# than reported from $script:SettingsStoreMode alone, because the mode is a
# request and the state is the outcome: Initialize-JsonSettings falls back to the
# registry when the file cannot be read, and off Windows that leaves no store at
# all. Callers (and Get-IMSettingsStore) need the outcome.
function Get-SettingsStoreInfo
{
    $persisted = $true
    $path = $null

    if($script:JsonSettingsObj -and $script:JSonSettingFile)
    {
        $mode = "Json"
        $path = $script:JSonSettingFile
    }
    elseif($script:JsonSettingsObj)
    {
        $mode = "Memory"
        $persisted = $false
    }
    elseif($script:IsWindowsOS)
    {
        $mode = "Registry"
        $path = Get-RegPath
    }
    else
    {
        # No settings file and no registry provider: reads fall back to the
        # registered defaults and writes go nowhere. Only reachable when the JSON
        # file failed to load off Windows.
        $mode = "None"
        $persisted = $false
    }

    [PSCustomObject]@{
        Mode        = $mode
        Path        = $path
        Persisted   = $persisted
        ValueCount  = @(Get-SettingsStoreEntries).Count
        # The store that was asked for, NOT $script:SettingsStoreMode: the JSON
        # fallback rewrites that variable to "Registry", so reporting it here made
        # Mode and RequestedMode agree in exactly the case a caller needs them to
        # differ - the one where the requested store failed to load.
        RequestedMode = $script:SettingsStoreModeRequested
    }
}

# Every value in the store as flat SubPath/Key/Value rows. Reading the registry
# recursively is the same walk Export-Settings already does, so both directions
# reuse Add-RegKeyToSettings and this only has to flatten one shape.
function Get-SettingsStoreEntries
{
    $tree = $script:JsonSettingsObj
    if(-not $tree) { $tree = Get-PersistedSettingsTree }
    if(-not $tree) { return @() }

    Get-SettingsTreeEntries $tree ""
}

# Recursive half of Get-SettingsStoreEntries. A nested object is a SubPath level;
# anything else is a value.
function Get-SettingsTreeEntries
{
    param($Node, [string]$SubPath)

    $entries = @()
    foreach($prop in $Node.PSObject.Properties)
    {
        # Exact type name, not `-is [PSCustomObject]`: that accelerator resolves to
        # PSObject, which almost everything satisfies once PowerShell has wrapped
        # it, so every leaf value would be walked as a subpath.
        if($null -ne $prop.Value -and $prop.Value.GetType().FullName -eq "System.Management.Automation.PSCustomObject")
        {
            $child = if($SubPath) { "$SubPath\$($prop.Name)" } else { $prop.Name }
            $entries += Get-SettingsTreeEntries $prop.Value $child
        }
        else
        {
            $entries += [PSCustomObject]@{
                SubPath = $SubPath
                Key     = $prop.Name
                Value   = $prop.Value
            }
        }
    }

    return $entries
}

# The persisted store as a nested PSCustomObject tree, whatever the platform.
# ConvertTo-Json/ConvertFrom-Json is the conversion, not a shortcut: it is exactly
# how Initialize-JsonSettings builds its object, so a seeded memory tree and a
# loaded file tree have the same shape and the same walk works on both.
function Get-PersistedSettingsTree
{
    if($script:JSonSettingFile -and [IO.File]::Exists($script:JSonSettingFile))
    {
        try
        {
            return (ConvertFrom-Json ([IO.File]::ReadAllText($script:JSonSettingFile)))
        }
        catch
        {
            Write-LogError "Failed to read settings file $script:JSonSettingFile" $_.Exception
            return $null
        }
    }

    if($script:IsWindowsOS)
    {
        $settingObj = [ordered]@{}
        Add-RegKeyToSettings $settingObj (Get-RegPath)
        return ($settingObj | ConvertTo-Json -Depth 20 | ConvertFrom-Json)
    }

    return $null
}

# Switch to memory mode. -Seed copies the persisted values in first, so a session
# starts from the user's real configuration and then diverges without writing
# back; without it the store starts empty and every unset key resolves to its
# registered default.
function Initialize-MemorySettings
{
    [CmdletBinding()]
    param([switch]$Seed)

    $tree = $null
    if($Seed)
    {
        $tree = Get-PersistedSettingsTree
        if(-not $tree) { Write-Log "No persisted settings to seed the in-memory store from" 2 }
    }

    if(-not $tree) { $tree = [PSCustomObject]@{} }

    # Order matters: null the file first so a write between the two assignments
    # cannot persist into the store we are leaving.
    $script:JSonSettingFile = $null
    $script:JsonSettingsObj = $tree
    $script:SettingsStoreMode = "Memory"
    $script:SettingsStoreModeRequested = "Memory"

    Write-Log "Settings store is now in memory$(if($Seed) { " (seeded with $(@(Get-SettingsStoreEntries).Count) values)" }). Nothing will be written to disk."
}

# Change the store at runtime. The public Use-IMSettingsStore is a thin wrapper.
function Set-SettingsStoreMode
{
    [CmdletBinding()]
    param(
        [ValidateSet("Memory", "Json", "Registry")]
        [string]$Mode,
        [string]$Path,
        [switch]$Seed
    )

    if($Mode -eq "Memory")
    {
        Initialize-MemorySettings -Seed:$Seed
        return
    }

    if($Mode -eq "Registry")
    {
        if(-not $script:IsWindowsOS)
        {
            Write-LogError "The registry settings store is only available on Windows"
            return
        }
        $script:JsonSettingsObj = $null
        $script:JSonSettingFile = $null
        $script:SettingsStoreMode = "Registry"
        $script:SettingsStoreModeRequested = "Registry"
        Write-Log "Settings store is now the registry: $(Get-RegPath)"
        return
    }

    # Json. A caller-supplied path replaces whatever file is loaded;
    # Initialize-JsonSettings creates it when it does not exist.
    $script:JsonSettingsObj = $null
    $script:JSonSettingFile = $Path
    $script:SettingsStoreMode = "Json"
    # Recorded BEFORE Initialize-JsonSettings, so a load failure - which calls
    # Clear-JsonSettingsValues and reverts the mode - still leaves "Json" as the
    # answer to "what did the caller ask for?".
    $script:SettingsStoreModeRequested = "Json"
    Initialize-JsonSettings

    if(-not $script:JsonSettingsObj)
    {
        # Initialize-JsonSettings already logged, and Clear-JsonSettingsValues has
        # reverted the mode - do not claim success.
        return
    }

    if($Seed -and $script:IsWindowsOS)
    {
        Write-Log "Seeding a settings file from the registry is what Export-Settings does; -Seed is ignored for the Json store" 2
    }
}

#region Resolver - address settings by KEY, never by path

# Does a value EXIST at this path? Deliberately not Get-SettingStoreValue: that
# reports an empty string as missing, so that clearing a text box in the settings
# form restores the registered default. "Configured" has to mean "written", which
# is a different question. Generalizes Get-IsTenantSettingConfigured, which could
# only ever ask about the current tenant's path.
function Test-SettingStoreValue
{
    param($SubPath = "", $Key = "")

    if(-not $Key) { return $false }

    if($script:JsonSettingsObj)
    {
        $node = Get-SettingsTreeNode $SubPath
        if(-not $node) { return $false }
        return ($null -ne ($node.PSObject.Properties | Where-Object Name -eq $Key))
    }

    if($script:IsWindowsOS)
    {
        try
        {
            $item = Get-Item -LiteralPath (Get-RegPath $SubPath) -ErrorAction Stop
            return ($item.GetValueNames() -contains $Key)
        }
        catch
        {
            # Missing key. Not an error - the value simply is not configured.
        }
    }

    return $false
}

# The store path a key lives at. This function is the whole point of the resolver
# layer: no caller outside it should ever spell a SubPath.
#
# A registered key takes its SubPath from its definition, which is what makes the
# write side agree with Get-SettingValue by construction. Every SubPath-mismatch
# bug so far (GraphPageSize, the bulk-export round-trip, DefaultCloud) was a hand-
# written path on one side only.
#
# $SubPath defaults to $null rather than "" because "" is a real path - the root of
# the store - so "not supplied" and "the root" have to stay distinguishable.
function Resolve-SettingStorePath
{
    [CmdletBinding()]
    param(
        [string]$Key,
        $Definition,
        [switch]$Tenant,
        [string]$TenantID,
        $SubPath = $null
    )

    if($Definition)
    {
        # A REGISTERED setting is stored where its registration says, always - a
        # caller-supplied -SubPath is reported and ignored rather than honoured.
        # Honouring it wrote the value to a path the application never reads, which
        # is precisely the path drift this layer exists to remove:
        # `Set-IMSetting GraphPageSize 100 -SubPath Wrong` looked like it worked and
        # changed nothing. -SubPath is for UNREGISTERED keys, where it is the only
        # way to name the path - both to write and (Resolve-SettingValue) to read.
        if($null -ne $SubPath -and "$SubPath" -ne "$($Definition.SubPath)")
        {
            Write-Log "Ignoring -SubPath '$SubPath' for '$Key': it is a registered setting and is always stored at '$($Definition.SubPath)'" 2
        }
        $SubPath = $Definition.SubPath
    }
    elseif($null -eq $SubPath)
    {
        Write-LogError "'$Key' is not a registered setting, so its storage path cannot be resolved. Pass -SubPath to read or write it anyway."
        return $null
    }

    # A registered setting with no -SubPath (EnvironmentText, AppTheme, every
    # General entry) stores at the ROOT of the tree, and its definition carries
    # $null rather than "". Returning that $null straight through would be read as
    # this function's failure sentinel by every caller, so a root-stored key could
    # never be written - Set-SettingValue returned quietly and the value vanished.
    if($null -eq $SubPath) { $SubPath = "" }

    if(-not $Tenant) { return $SubPath }

    if(-not $TenantID) { $TenantID = $script:OrganizationId }
    if(-not $TenantID)
    {
        # Silently writing the global value instead would be the worst outcome: the
        # caller asked for one tenant and would have changed every tenant.
        Write-LogError "Cannot resolve a tenant-specific path for '$Key': no tenant id was supplied and no tenant is connected."
        return $null
    }

    if($SubPath) { return "$TenantID\$SubPath" }
    return $TenantID
}

# A value in the shape the store has always held it in.
#
# Everything on disk is a string: the settings form calls Save-SettingStoreValue
# with no -Type, so it takes the "String" branch and writes $Value.ToString().
# Reproducing that exactly is a hard requirement - a value written through this
# layer has to be indistinguishable from one written by the form, or the two
# disagree about the same key.
function Format-SettingStoreValue
{
    param($Value, $Definition)

    # $null means "remove the value" to Save-SettingStoreValue. Pass it through.
    if($null -eq $Value) { return $null }

    if($Definition.Type -eq "Boolean")
    {
        # PascalCase "True"/"False" - that is what $true.ToString() produces and
        # what every existing store contains. Coerce first so $true, "true" and
        # "TRUE" all land in the one canonical shape.
        #
        # A string is compared to "true" rather than cast: [bool]"False" is $true
        # in PowerShell (any non-empty string is), which would turn every attempt
        # to store False into True. This is the same test Get-SettingValue uses to
        # read the value back, so writer and reader cannot disagree.
        $bool = if($Value -is [string]) { $Value -eq "true" } else { [bool]$Value }
        return $bool.ToString()
    }

    return $Value.ToString()
}

# Write a setting by key. The counterpart to Get-SettingValue, and the reason no
# caller needs to know that "GraphPageSize" lives under "IntuneManager".
#
# -Tenant writes the value for one tenant only, which is the same precedence
# Get-SettingValue reads with (tenant first, then global).
function Set-SettingValue
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [AllowEmptyString()]
        $Value,
        [switch]$Tenant,
        [string]$TenantID,
        $SubPath = $null,
        [switch]$PassThru
    )

    $definition = Get-SettingDefinitionByKey $Key
    $path = Resolve-SettingStorePath -Key $Key -Definition $definition -Tenant:$Tenant -TenantID $TenantID -SubPath $SubPath
    if($null -eq $path) { return }

    $stored = Format-SettingStoreValue $Value $definition

    # Always "String": see Format-SettingStoreValue. The registered Types
    # (Boolean, Int, File, List, ...) are UI editor hints, not storage types, and
    # are not valid RegistryValueKind names - passing one through would throw on
    # Windows.
    Save-SettingStoreValue $path $Key $stored "String"
    Write-LogDebug "Setting '$Key' set to '$stored' at '$path'"

    # -SubPath forwarded, or -PassThru on an unregistered key writes the value and
    # then fails to read back the very path it just wrote to.
    if($PassThru) { Resolve-SettingValue -Key $Key -TenantID $TenantID -SubPath $SubPath }
}

# Remove a setting by key so it falls back to the next level: a tenant value
# reverts to the global one, a global value to the registered default.
function Remove-SettingValue
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,
        [switch]$Tenant,
        [string]$TenantID,
        $SubPath = $null
    )

    $definition = Get-SettingDefinitionByKey $Key
    $path = Resolve-SettingStorePath -Key $Key -Definition $definition -Tenant:$Tenant -TenantID $TenantID -SubPath $SubPath
    if($null -eq $path) { return }

    Remove-SettingStoreValue $path $Key
    Write-LogDebug "Setting '$Key' removed from '$path'"
}

# Is this key written at the level asked about? Not "does it have a value" - an
# unconfigured key always has a value, its registered default.
function Test-SettingValueConfigured
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,
        [switch]$Tenant,
        [string]$TenantID,
        $SubPath = $null
    )

    $definition = Get-SettingDefinitionByKey $Key
    $path = Resolve-SettingStorePath -Key $Key -Definition $definition -Tenant:$Tenant -TenantID $TenantID -SubPath $SubPath
    if($null -eq $path) { return $false }

    Test-SettingStoreValue $path $Key
}

# The value AND where it came from. Get-SettingValue answers "what is it"; this
# answers "and why", which is what a runbook needs before it decides to change
# something and what the settings form needs to show a tenant override.
#
# Computed fresh from the store every call. Get-SettingValue caches its last read
# on the definition object as .Value, which is a per-session artifact of whoever
# read it last (and with which -TenantID) - not a fact about the store.
function Resolve-SettingValue
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,
        [string]$TenantID,
        [switch]$GlobalOnly,
        [switch]$TenantOnly,
        $SubPath = $null
    )

    $definition = Get-SettingDefinitionByKey $Key
    $registered = $null -ne $definition
    if(-not $registered)
    {
        # No registration means no SubPath, no Type and no default. The caller can
        # supply the path instead, which is what makes a hidden key (ExportReplaceTokens,
        # per-feature state) readable through the same resolver - and therefore through
        # Get-IMSetting - rather than only writable. Without this the write side accepted
        # -SubPath and the read side had no way to address what it had just written.
        if($null -eq $SubPath)
        {
            Write-LogError "'$Key' is not a registered setting, so it has no definition to resolve against. Pass -SubPath to read it anyway."
            return
        }

        # Stands in for a registration so the rest of this function, and
        # Resolve-SettingStorePath, need no unregistered-key branch. Type and
        # DefaultValue stay $null: an unregistered key has no declared type to
        # normalize to and no default to fall back on.
        $definition = [PSCustomObject]@{
            Key          = $Key
            Title        = $null
            Type         = $null
            DefaultValue = $null
            SubPath      = $SubPath
            Section      = $null
        }
    }

    if(-not $TenantID -and $script:OrganizationId) { $TenantID = $script:OrganizationId }

    # A tenant-scoped read with no tenant to read against cannot be answered.
    # Falling through would have reported the registered default as though it were
    # the tenant's value, so a caller checking "does this tenant override the
    # setting?" got a confident answer while signed out. The write side already
    # refuses this case (Resolve-SettingStorePath), so refusing it here keeps the
    # two consistent.
    if($TenantOnly -eq $true -and -not $TenantID)
    {
        Write-LogError "Cannot resolve the tenant value of '$Key': no tenant id was supplied and no tenant is connected."
        return
    }

    $value = $null
    $source = "Default"

    # Both paths come from Resolve-SettingStorePath rather than being spelled here,
    # so this reader cannot drift from the writer - including for a root-stored key,
    # where composing the tenant path by hand leaves a trailing separator.
    if($GlobalOnly -ne $true -and $TenantID)
    {
        $tenantPath = Resolve-SettingStorePath -Key $Key -Definition $definition -Tenant -TenantID $TenantID -SubPath $SubPath
        if($null -ne $tenantPath)
        {
            $value = Get-SettingStoreValue $tenantPath $definition.Key
            if($null -ne $value) { $source = "Tenant" }
        }
    }

    if($null -eq $value -and $TenantOnly -ne $true)
    {
        $value = Get-SettingStoreValue (Resolve-SettingStorePath -Key $Key -Definition $definition -SubPath $SubPath) $definition.Key
        if($null -ne $value) { $source = "Global" }
    }

    # A scoped read reports the ABSENCE of a value at that scope as $null/"NotSet",
    # never as the registered default: "this tenant does not override the setting"
    # and "this tenant overrides it to the same value as the default" are different
    # facts, and substituting the default made them indistinguishable. Only an
    # unscoped (Effective) read falls back to the default, which is what the
    # application itself resolves.
    #
    # An unregistered key reports NotSet at every scope, Effective included: it has
    # no registered default, so "Default" there would name a fallback that does not
    # exist and report $null as though the value had been resolved.
    if($null -eq $value -and (-not $registered -or $TenantOnly -eq $true -or $GlobalOnly -eq $true))
    {
        $source = "NotSet"
    }
    elseif($null -eq $value)
    {
        $value = $definition.DefaultValue
    }

    # Same normalization as Get-SettingValue, so both agree on a stored "False".
    if($definition.Type -eq "Boolean" -and $null -ne $value)
    {
        $value = $value -eq $true -or $value -eq "true"
    }

    [PSCustomObject]@{
        Key        = $definition.Key
        Value      = $value
        Source     = $source
        Type       = $definition.Type
        SubPath    = $definition.SubPath
        Section    = $definition.Section
        Default    = $definition.DefaultValue
        TenantID   = $TenantID
        Title      = $definition.Title
    }
}

#endregion

#region Portable settings - a store as a file, independent of where it lives

# Write the whole current store to a JSON file, whatever store it is. This is the
# other half of the runbook story: check a settings file into source control, then
# load it into an in-memory store at the start of a run so the worker's own
# registry and settings file are neither read nor written.
function Export-SettingsStoreToFile
{
    [CmdletBinding()]
    param([string]$Path)

    $tree = $script:JsonSettingsObj
    if(-not $tree) { $tree = Get-PersistedSettingsTree }
    if(-not $tree) { $tree = [PSCustomObject]@{} }

    try
    {
        $fi = [IO.FileInfo]$Path
        if($fi.Directory -and -not $fi.Directory.Exists) { $fi.Directory.Create() }

        $tree | ConvertTo-Json -Depth 20 | Out-File -LiteralPath $Path -Force -Encoding utf8
        $count = @(Get-SettingsTreeEntries $tree "").Count
        Write-Log "Exported $count settings to $Path"
        return $true
    }
    catch
    {
        Write-LogError "Failed to export settings to $Path" $_.Exception
        return $false
    }
}

# Load a settings file into the ACTIVE store, value by value through
# Save-SettingStoreValue rather than by swapping the tree. That is what makes one
# implementation correct for all three modes: the values persist in Json mode, land
# in the registry in Registry mode, and stay in memory in Memory mode - and an
# imported file merges into what is already there instead of replacing it.
#
# For a clean slate, switch to an unseeded memory store first
# (Use-IMSettingsStore -Memory) and import into that.
function Import-SettingsStoreFromFile
{
    [CmdletBinding()]
    param([string]$Path)

    if(-not [IO.File]::Exists($Path))
    {
        Write-LogError "Settings file '$Path' does not exist"
        return
    }

    try
    {
        $tree = ConvertFrom-Json ([IO.File]::ReadAllText($Path))
    }
    catch
    {
        Write-LogError "Failed to read settings file '$Path'" $_.Exception
        return
    }

    $entries = @(Get-SettingsTreeEntries $tree "")
    $unknown = @()

    foreach($entry in $entries)
    {
        # A key with no registration is not an error - the hidden keys
        # (ExportReplaceTokens) and the per-feature state namespaces are real and
        # deliberately unregistered. A typo looks exactly the same though, and
        # would silently do nothing, so say which ones they were.
        if(-not (Get-SettingDefinitionByKey $entry.Key)) { $unknown += $entry.Key }

        Save-SettingStoreValue $entry.SubPath $entry.Key $entry.Value "String"
    }

    if($unknown.Count -gt 0)
    {
        Write-Log "Imported $($unknown.Count) value(s) for keys that are not registered settings: $(($unknown | Sort-Object -Unique) -join ', ')" 2
    }

    Write-Log "Imported $($entries.Count) settings from $Path into the $((Get-SettingsStoreInfo).Mode) store"
    return $entries.Count
}

#endregion

