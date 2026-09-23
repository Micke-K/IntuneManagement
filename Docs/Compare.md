# Compare

## Purpose

The compare feature compares policies from Intune, export folders, selected objects, or named providers. It is used by the UI to show differences and by bulk compare workflows to generate structured output.

## Main Files

| File | Responsibility |
| --- | --- |
| `Public/Compare-GraphPolicy.ps1` | public entry point into compare UI. |
| `Internal/Compare.ps1` | compare engine and output providers. |
| `Classes/CompareClasses.ps1` | compare provider classes. |
| `UI/Extensions/CompareUI.ps1` | compare forms and UI orchestration. |

## Compare Providers

Compare providers abstract the source of objects:

| Provider | Source |
| --- | --- |
| `CompareExportFilesProvider` | exported JSON files. |
| `CompareIntuneWithExportProvider` | live Intune versus export folder. |
| `CompareNamedObjectsProvider` | named live objects. |
| `CompareExportedFoldersProvider` | two export folders. |

## Object Normalization

Before comparing, policies are converted to a compare-friendly JSON/object form:

1. full policy object is resolved when needed;
2. ignored core properties are removed;
3. settings catalog values can be flattened into readable keys;
4. documentation plugin output can be used when available;
5. output rows are generated with property paths and differences.

## Compare Providers - options

Each provider is a class whose properties are the options. Construct it, set the
properties, pass it to `Compare-IMGraphPolicy` through the matching parameter (or select
it in the Bulk > Compare form, which shows the same fields).

| Provider | Parameter | Properties |
|---|---|---|
| `CompareExportFilesProvider` | `-ExportFiles` | `ExportPath` - the export root; `NameFilter` - only policies whose name contains this. |
| `CompareIntuneWithExportProvider` | `-IntuneWithExport` | `ExportPath`, `NameFilter` - as above; each live policy is compared with its exported file. |
| `CompareNamedObjectsProvider` | `-NamedObjects` | `SourcePattern`, `ComparePattern` - name patterns that pair objects (`Pilot - X` with `Prod - X`); `SavePath` - where the result is written; `RemoveProperties` - properties dropped before comparing, default `@('Id')`. |
| `CompareExportedFoldersProvider` | `-ExportedFolders` | `SourcePath`, `ComparePath` - two export roots; `NameFilter`. |

All four take `-PolicyGroupIds <string[]>` on the command to limit the groups compared.

```powershell
$p = [CompareNamedObjectsProvider]::new()
$p.SourcePattern  = 'Pilot - '
$p.ComparePattern = 'Prod - '
$p.SavePath       = 'C:\Reports\pilot-vs-prod.csv'
Compare-IMGraphPolicy -NamedObjects $p -PolicyGroupIds DeviceConfiguration, Compliance
```

## Run options

Set once per run with `Set-CompareRuntimeOptions` (the Bulk > Compare form does this for
you); the compare functions read them.

| Option | Values | Effect |
|---|---|---|
| `CompareType` | `Property` (default), `Documentation` | Which comparison type - see the strategies below. `Documentation` compares what the documentation engine renders, so two policies that document the same are equal even if raw json differs. |
| `IgnoreCoreProperties` | bool | Skip id, timestamps, version and the other server-side properties. |
| `SaveType` | `objectType`, `all` | One output file per object type, or one file for everything. |
| `OutputProvider` | CSV or Json provider instance | Chosen by the output file's extension when saving from the form (`.csv` / `.json`). |
| `CsvDelimiter` | string | CSV delimiter; default is the culture's list separator. |
| `ObjectSeparator` | string | Separator between items of a multi-value property in the output. |
| `SkipAssignments` | bool | Leave assignments out of the comparison. |

The single-object form adds a result filter (All / Mismatch / Match) that only affects the
grid, not the saved file.

## Compare Strategies

| Strategy | Function | Notes |
| --- | --- | --- |
| property compare | `Compare-ObjectsBasedonProperty` | compares raw property paths. |
| settings compare | `Compare-ObjectsBasedonSettings` | extracts settings catalog and intent settings. |
| documentation compare | `Compare-ObjectsBasedonDocumentation` | uses `Invoke-ObjectDocumentation` when present. |

## Output

Bulk compare writes through an output provider. Two ship: **CSV** (`CompareCSVOutputProvider`,
honours `CsvDelimiter`) and **JSON** (`CompareJsonOutputProvider`). The provider is picked
from the output file's extension, or selected in the form.

Typical output fields include:

| Field | Meaning |
| --- | --- |
| object name/type | source policy identity. |
| property path | location of difference. |
| source value | value in left/source object. |
| target value | value in right/target object. |
| result type | same, different, missing, extra, etc. |

## Extension Notes

Add new compare behavior in `Internal/Compare.ps1` when the comparison semantics change. Add new source behavior by creating a provider class in `Classes/CompareClasses.ps1`. Keep UI code limited to collecting options and displaying results.

