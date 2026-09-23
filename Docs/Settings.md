# Settings

Settings are the tool's configuration: everything the Settings dialog shows, plus a
handful of hidden keys and per-feature state. This page is about how they are stored,
how they are addressed, and how to drive them from automation without touching the
machine they run on.

For cache and logging, see [Settings Cache Logging](SettingsCacheLogging.md).


Every registered key, with type, default and description, is in
[SettingsReference.md](SettingsReference.md) - generated from the code, so it is always current.

## Three layers

Addressing a setting used to mean knowing where it lives. It does not any more:

| Layer | Where | Addressed by | Functions |
| --- | --- | --- | --- |
| storage | `Internal/Core.ps1` | **path** (`SubPath` + key) | `Get-SettingStoreValue`, `Save-SettingStoreValue`, `Remove-SettingStoreValue` |
| resolver | `Internal/Settings.ps1` | **key** | `Get-SettingValue`, `Set-SettingValue`, `Remove-SettingValue`, `Test-SettingValueConfigured`, `Resolve-SettingValue` |
| public | `Public/*Setting*.ps1` | **key** | `Get-IMSetting`, `Set-IMSetting`, `Remove-IMSetting`, `Get-IMSettingDefinition`, plus the store cmdlets below |

The dependency direction is one-way: public calls the resolver, the resolver calls
storage. Nothing goes the other way.

**Prefer the resolver.** It looks the key up in the registered definitions and derives
the `SubPath` from the registration, so a write cannot land somewhere a read never
looks. Every SubPath bug the project has had - `GraphPageSize` resolving differently in
headless sessions, the bulk-export remember-last-used round trip being silently dead,
`DefaultCloud` - was a hand-written path on one side only. `Tools/Audit-Settings.ps1`
reports remaining path-addressed access to a registered key as a warning.

`Get-SettingValue` lives in `Internal/Core.ps1` rather than with the rest of the
resolver because `Write-Log` reads `LogFile` through it during preload, long before
`Internal/` is dot-sourced.

### Registration

A setting exists because some file called `Add-SettingsObject`:

```powershell
Add-SettingsObject -Key "GraphPageSize" -Section "IntuneManager" -SubPath "IntuneManager" `
                   -Title "Page size" -Type "List" -DefaultValue "0" -ItemsSource $sizes
```

`-SubPath` is the storage path. With no `-SubPath` the `-Section` is used, except
`General`, which stores at the root of the store. `-Type` is a **UI editor hint**
(`Boolean`, `Int`, `File`, `List`, `Text`), not a storage type - everything on disk is
a string.

`Get-IMSettingDefinition` lists the registrations; `Get-IMSettingDefinition -Section
IntuneManager` narrows to one section.

## Three store modes

The mode is decided at the very top of `Internal/Core.ps1`, before anything can log:

| Mode | Backing | Default when |
| --- | --- | --- |
| `Registry` | `HKCU:\Software\IntuneManagement` | on Windows |
| `Json` | `IM_SETTINGS_FILE`, else `LocalApplicationData/IntuneManagement/Settings.json` | off Windows, or `IM_SETTINGS_FILE` is set |
| `Memory` | a tree with no file behind it - reads and writes work, nothing touches disk | never; opt in explicitly |

Environment variables, read once at load:

| Variable | Effect |
| --- | --- |
| `IM_SETTINGS_STORE` | `Memory`, `Json` or `Registry`. `Registry` off Windows falls back to `Json`. An unrecognized value falls back to the platform default and is warned about once logging works. |
| `IM_SETTINGS_FILE` | the `Json` store's file. Created if missing. Implies `Json`. |

Memory mode is not a fourth code path: the storage primitives take their JSON branch
whenever a settings **object** exists, and persist only when a settings **file** exists
too. Memory mode is an object with no file.

`Get-IMSettingsStore` reports what the store actually is - which is not always what was
requested, because a `Json` store whose file cannot be read falls back to the registry,
and off Windows that leaves no store at all (`Mode = "None"`: reads return registered
defaults, writes go nowhere).

```powershell
Get-IMSettingsStore                 # Mode / Path / Persisted / ValueCount / RequestedMode
Get-IMSettingsStore -IncludeValues   # ... plus every SubPath/Key/Value row
```

## Public API

```powershell
Get-IMSetting GraphPageSize                     # the effective value
Get-IMSetting GraphPageSize -Detailed           # value + Source (Default/Global/Tenant) + Type + Default + ...
Get-IMSetting ExportFolder -Scope Global        # ignore any tenant override
Get-IMSetting ExportFolder -Scope Tenant -TenantID <guid>

Set-IMSetting GraphPageSize 999
Set-IMSetting ExportFolder '\\server\intune\contoso' -Scope Tenant
Set-IMSetting UseBatchAPI $false -PassThru      # returns the resolved value afterwards

Remove-IMSetting ExportFolder -Scope Tenant     # revert to the global value
Remove-IMSetting ExportFolder                   # revert to the registered default
```

`Set-IMSetting` and `Remove-IMSetting` **persist by default** - to the registry or the
settings file, whichever the active store is. There is no opt-in switch. They support
`-WhatIf`, they log the store and path they wrote to, and `Get-IMSettingsStore` lets a
script assert on the store before writing anything. If a script must not write to the
machine it runs on, switch to a memory store first.

`Get-IMSetting` accepts a key from the pipeline, so `Get-IMSettingDefinition -Section
IntuneManager | Get-IMSetting -Detailed` dumps a whole section with provenance.

### Scope and precedence

A setting can be written globally or for one tenant. Reads take the tenant value first:

```text
tenant value  ->  global value  ->  registered DefaultValue
```

`-Scope Tenant` with no `-TenantID` uses the connected tenant, and **fails loudly** if
there is none - writing the global value instead would change every tenant when the
caller asked for one.

`Get-IMSetting -Detailed` reports which level answered in `Source`. That is computed
fresh from the store on every call; the `.Value` cached on a definition object by
`Get-SettingValue` is a per-session artifact of whoever read it last.

A **scoped** read (`-Scope Global` / `-Scope Tenant`) reports the absence of a value at
that level as `$null` with `Source = 'NotSet'`. It does not fall back to the registered
default, because "this tenant does not override the setting" and "this tenant overrides
it to the same value as the default" are different facts. Only the default `-Scope
Effective` falls back, which is what the application itself resolves.

### Keys with no registration

The hidden keys (`ExportReplaceTokens`) and the per-feature state namespaces are not
registered with `Add-SettingsObject`, so they have no path to take from a registration.
`-SubPath` supplies it, and is accepted on **all three** cmdlets - a key that can be
written this way can be read and removed the same way:

```powershell
Set-IMSetting    ExportReplaceTokens 'TenantId' -SubPath 'IntuneManager'
Get-IMSetting    ExportReplaceTokens -SubPath 'IntuneManager'
Remove-IMSetting ExportReplaceTokens -SubPath 'IntuneManager'
```

Such a key has no registered default and no declared type, so its `Source` is `Tenant`,
`Global` or `NotSet` - never `Default` - and the value comes back as the string the
store holds. Tenant precedence works as usual (`-Scope Tenant`, `-TenantID`), which is
how a per-tenant `ExportReplaceTokens` is set.

For a **registered** key `-SubPath` is reported in the log and ignored: the registration
decides the path, so a write cannot be aimed somewhere the application never reads.

## Automation: a run that reads and writes nothing on the worker

The problem: a runbook on a shared Hybrid Worker has no business reading, let alone
writing, that worker's `HKCU` hive or settings file. A memory store solves it, and a
checked-in settings file makes the configuration reviewable:

```powershell
Import-Module .\IntuneManagement.psd1

# Nothing from here on touches the worker's registry or settings file.
Use-IMSettingsStore -Memory
Import-IMSettingsStore .\config\documentation-run.json

# Whatever the file did not cover.
Set-IMSetting GraphPageSize 999
Set-IMSetting UseBatchAPI $true

Connect-IMIntuneManagement -Provider OAuth -TenantId $tid -AppId $appId -Certificate $cert
Start-IMGraphBulkDocumentation -OutputFormat html -PolicyGroup DeviceConfiguration
```

`IM_SETTINGS_STORE=Memory` in the environment does the same thing from the first line
of the module load, which matters if anything the module logs during load would
otherwise resolve against the real store.

| Cmdlet | Use |
| --- | --- |
| `Use-IMSettingsStore -Memory [-Seed]` | switch to memory. `-Seed` copies the persisted values in first, so the session starts from the real configuration and then diverges without writing back. |
| `Use-IMSettingsStore <path>` | switch to a `Json` store at that file, creating it if missing. |
| `Use-IMSettingsStore -Registry` | switch back to `HKCU`. Windows only. |
| `Export-IMSettingsStore <path>` | write the whole current store to a JSON file, whatever mode it is in - including a registry store. |
| `Import-IMSettingsStore <path>` | load a settings file into the **active** store, value by value. |

Two things about `Import-IMSettingsStore` worth knowing:

- It **merges** into what is already there rather than replacing it. For a clean slate,
  `Use-IMSettingsStore -Memory` (without `-Seed`) first, then import into that.
- It goes through the storage layer, so one implementation is correct for all three
  modes: values persist in `Json` mode, land in the registry in `Registry` mode, stay in
  memory in `Memory` mode. Which also means importing into a registry store **writes to
  the registry**.
- Keys with no registration are imported anyway - the hidden keys and per-feature state
  namespaces are real - but they are listed in a warning, because a typo looks
  identical and would otherwise silently do nothing.

## Stored shape

Everything is a string. Booleans are PascalCase `"True"`/`"False"`, which is what
`$true.ToString()` produces and what every existing store contains.
`Set-SettingValue` coerces through `Format-SettingStoreValue`, so `$true`, `"true"` and
`"TRUE"` all land in the one canonical shape, and a value written through the resolver
is indistinguishable from one written by the Settings dialog.

Watch out for `[bool]"False"` - it is `$true` in PowerShell, as any non-empty string is.
Stored Booleans are compared to `"true"`, never cast. Both the reader and the writer do
this the same way on purpose.

An **empty string** counts as "not set" on read, so clearing a text box in the Settings
dialog restores the registered default rather than persisting `""`. "Is it configured"
is therefore a different question from "does it have a value", and
`Test-SettingValueConfigured` (`Get-IMSetting` does not expose it) answers it by asking
whether the value exists at that path at all.

## Legacy layout note

A `SubPath` with more than one level (`<tenantid>\IntuneManager`) used to be stored by
the JSON/memory store as a **single flat property** literally named
`"<tenantid>\IntuneManager"`, while the registry stored the same path as nested keys.
The cause was `"a\b".Split(@('/','\'))` not splitting at all - PowerShell binds a string
array to a different `String.Split` overload and hands back the whole path as one
element. It is now `[char[]]`-cast and splits properly.

Existing settings files keep working: `Get-SettingsTreeNode` resolves a flat property
first and writes back to it when it finds one, so no tenant-specific value is orphaned
by the fix. New values nest.

## Testing

Settings leak from the developer machine into tests, and tests leak into the developer
machine. `Tests/TestBootstrap.ps1` has the two helpers:

```powershell
Describe 'x' {
    BeforeEach { $script:saved = Use-TestSettingsStore }
    AfterEach  { Restore-TestSettingsStore -State $script:saved }

    InModuleScope IntuneManagement { It 'y' { ... } }
}
```

`Use-TestSettingsStore` swaps in an empty memory store (so every key resolves to its
registered default) and returns the previous state for `Restore-TestSettingsStore`.

**The hooks must be declared at `Describe` level, wrapping `InModuleScope`.** Pester 4
silently never runs a `BeforeEach`/`AfterEach` declared *inside* an `InModuleScope`
block - nothing fails, the tests just run against the developer's real store and write
to it. `Static.Tests.ps1` has a gate for this with a shrink-only allowlist of the files
that still get it wrong.

`Tests/SettingsStore.Tests.ps1` covers the store modes, path resolution, value shape,
provenance, portability and the public surface.

## Auditing

`Tools/Audit-Settings.ps1` cross-references every registration against every read and
write. It is wired into `Static.Tests.ps1` as a gate; see its help for the full
taxonomy. Errors: orphan registrations, engine-consumed-but-UI-registered (breaks
headless sessions), SubPath-mismatched access, duplicate registrations. Warnings:
UI-only-but-engine-registered, and registered keys addressed by path instead of through
the resolver.

```powershell
.\Tools\Audit-Settings.ps1
.\Tools\Audit-Settings.ps1 -IncludeUndefined   # also keys read but never registered
```
