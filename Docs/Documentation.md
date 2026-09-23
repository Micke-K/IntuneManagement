# Documenting policies

How to produce documentation from a script or a pipeline: the outputs, how to select
what gets documented, every option, and what each output provider expects. The UI's
Bulk > Document form drives exactly the same engine, so anything set there can be set
from `-Options`.

For the engine's internals - handlers, input providers, the customizer hooks and how to
add a renderer for a new type - read the source under `Internal/Documentation/`.

## Contents

- [The two commands](#the-two-commands)
- [Outputs](#outputs)
- [Selecting what to document](#selecting-what-to-document)
- [Documenting from an export folder](#documenting-from-an-export-folder)
- [Options](#options)
- [Per-output options](#per-output-options)
- [Where files land and how they are named](#where-files-land-and-how-they-are-named)
- [Languages](#languages)
- [Which types document, and how](#which-types-document-and-how)
- [Pipeline examples](#pipeline-examples)

## The two commands

| Command | Use when |
|---|---|
| `Start-IMGraphBulkDocumentation -OutputFormat <formats> ...` | You want files. Selects by type, group, object or export folder, runs every selected policy through the chosen outputs. |
| `Get-IMGraphDocumentation -PolicyObject <policy>` | You want the data. Returns one policy's documentation as an object - `BasicInfo`, `FilteredSettings`, `Assignments`, `Scripts`, `CustomTables` - to build your own report from. |

Both take `-Language` and `-Options`. Full parameter tables are in
[CommandReference.md](CommandReference.md#documentation).

## Outputs

`-OutputFormat` names one or more providers, comma-separated. `Get-IMDocumentationOutput`
lists them.

| Value | Produces | Notes |
|---|---|---|
| `html` | one `.html`, or one per policy | Self-contained; the CSS is inlined. |
| `md` | one `.md`, or one per policy | Optionally with the CSS inlined for renderers that honour it. |
| `word` | one `.docx` | **Windows only** - needs Word installed (COM automation). Silently degrades elsewhere. |
| `json` | one `.json`, or one per policy | The raw documentation objects - the same data `Get-IMGraphDocumentation` returns. |
| `csv` | one `.csv` per policy type | Flat rows; good for spreadsheets and diffing. |
| `atlassian` | one file, or one per policy, in Confluence storage format | Paste into a Confluence page, or push through the REST API. Headings carry stable anchors that downstream tooling can link to - see [DocumentationAtlassianOutput.md](DocumentationAtlassianOutput.md). |

```powershell
Start-IMGraphBulkDocumentation -OutputFormat html -PolicyGroup Compliance
Start-IMGraphBulkDocumentation -OutputFormat 'html,md,json' -PolicyGroup Compliance
```

## Selecting what to document

One parameter set per way of choosing:

```powershell
# Every policy of one or more types
Start-IMGraphBulkDocumentation -OutputFormat html -PolicyType CompliancePolicies, ConditionalAccess

# Every type in one or more groups
Start-IMGraphBulkDocumentation -OutputFormat html -PolicyGroup DeviceConfiguration, EndpointSecurity

# Specific policies, from the pipeline
Get-IMGraphPolicies -PolicyType SettingsCatalog -NameFilter 'Baseline' |
    Start-IMGraphBulkDocumentation -OutputFormat html

# Everything in an export folder (see the next section)
Start-IMGraphBulkDocumentation -OutputFormat html -SourceFolder C:\IntuneExport
```

Type and group ids are the ones in the README's policy table. Nothing is documented that
the signed-in identity cannot read; a type it has no access to is skipped with a log line.

## Documenting from an export folder

`-SourceFolder` documents the json files of an export instead of live policies. Two
things to know:

- **It still needs a signed-in tenant - but any tenant, not the source.** Setting
  definitions, templates and category names are generic Intune data and are looked up
  live from whatever tenant is connected. The policies themselves come from the files.
- **Source-tenant names are not looked up.** The mode sets
  `SourceTenantUnavailable = $true`, which skips every lookup only the source tenant could
  answer: assignment group names, scope tag names, filter names, app names. Those are
  resolved from the export's migration table when the export was made with one, and
  shown as ids otherwise.

This is the mode for documenting a tenant you no longer have access to, or for
generating documentation in a pipeline from an export artifact rather than from a live
sign-in with broad read rights.

## Options

`-Options` is a hashtable. Anything not given falls back to the value saved by the UI's
documentation form, then to the default below. All keys are optional.

| Key | Default | Effect |
|---|---|---|
| `IncludeScripts` | `$true` | Include script bodies (PowerShell, shell, remediation) in the output. |
| `ExcludeScriptSignature` | `$false` | Strip Authenticode signature blocks from included scripts. |
| `IncludePolicyId` | `$false` | Add the policy's id to its basic information. |
| `ExcludeAssignments` | `$false` | Leave assignments out. |
| `SkipNotConfigured` | `$false` | Omit settings that are not configured. |
| `SkipDefaultValues` | `$false` | Omit settings still at their default. |
| `SkipDisabled` | `$true` | Omit settings that are disabled. |
| `SetUnconfiguredValue` | `$true` | Render unconfigured settings with the text below instead of blank. |
| `SetDefaultValue` | `$false` | Render settings at default with their default value. |
| `NotConfiguredText` | `notConfigured` | Text for an unconfigured setting: `notConfigured` (the language string), `empty`, or `asis`. |
| `ValueOutputProperty` | `value` | For ADMX settings: `value`, or `valueWithLabel` to include the setting's label. |
| `PropertySeparator` | `;` | Separator between values of a multi-value property. |
| `ObjectSeparator` | newline | Separator between items of a collection. |
| `SkipDocumentInfo` | `$false` | Omit the "documented by / on" block. |
| `FallbackDocumentation` | `$true` | Types with no dedicated renderer are documented as basic information plus one row per property. Off, they are skipped with a warning. (The UI calls this *Document unsupported types*.) |
| `SourceTenantUnavailable` | `$false` | Set automatically by `-SourceFolder`; see above. The old name `OfflineDocumentation` still works. |
| `Outputs` | `@{}` | Per-output options - next section. |

```powershell
$o = @{
    SkipNotConfigured = $true
    SkipDefaultValues = $true
    IncludePolicyId   = $true
    ExcludeAssignments = $true
}
Start-IMGraphBulkDocumentation -OutputFormat html -PolicyGroup Compliance -Options $o
```

## Per-output options

Under `Outputs`, one hashtable per provider value. A key given here wins over the saved
UI setting, which wins over the default.

```powershell
$o = @{
    Outputs = @{
        html = @{ HTMLDocumentName = 'C:\Reports\Intune-%Date%.html'; HTMLDocumentFileType = 'Object' }
        md   = @{ MDDocumentName = 'C:\Reports\Intune.md'; MDIncludeCSS = $false; MDDocumentSkipDate = $true }
    }
}
Start-IMGraphBulkDocumentation -OutputFormat 'html,md' -PolicyGroup Compliance -Options $o
```

### `html`

| Key | Default | Effect |
|---|---|---|
| `HTMLDocumentName` | `%MyDocuments%\%Organization%-%Date%.html` | Output file. Placeholders below. |
| `HTMLDocumentFileType` | `Full` | `Full` - one file; `Object` - one file per policy, named after it, in the same folder. |
| `HTMLTitleProperty` | `Intune documentation` | Page title. |
| `HTMLCSSFile` | the shipped `DefaultHTMLStyle.css` | Your own stylesheet, inlined. |
| `HTMLOpenFile` | `$true` | Open the result when done - set `$false` in a pipeline. |

### `md`

| Key | Default | Effect |
|---|---|---|
| `MDDocumentName` | `%MyDocuments%\%Organization%-%Date%.md` | Output file. |
| `MDDocumentFileType` | `Full` | `Full` or `Object`, as for html. |
| `MDTitleProperty` | `Intune documentation` | Document title. |
| `MDCSSFile` | the shipped `DefaultMDStyle.css` | Stylesheet for `MDIncludeCSS`. |
| `MDIncludeCSS` | `$true` | Inline the CSS (renderers that honour `<style>` in Markdown). `$false` for plain Markdown. |
| `MDDocumentSkipDate` | `$false` | Leave the date out - keeps a committed file from changing on every run. |
| `MDOpenFile` | `$true` | Open when done. |

### `json`

| Key | Default | Effect |
|---|---|---|
| `JSONDocumentName` | `%MyDocuments%\%Organization%-%Date%.json` | Output file. |
| `JSONOutputFileType` | `Full` | `Full` or `Object`. |
| `JSONOpenFile` | `$true` | Open when done. |

### `csv`

| Key | Default | Effect |
|---|---|---|
| `CSVDocumentationPath` | *(none)* | Root folder for the files. Always set it - with no value the files are written relative to the current directory. One file per policy type, in a subfolder named after the type when `CSVAddObjectType` is on. |
| `CSVExportProperties` | `simple` | `simple` - the standard columns; `custom` - the list below. |
| `CSVCustomDisplayProperties` | `Name,Value,Category` | Columns for `custom`. |
| `CSVDelimiter` | *(culture list separator)* | Override, e.g. `;`. |
| `CSVAddObjectType` | `$true` | Object type column. |
| `CSVAddCompanyName` | `$false` | Tenant name column. |

### `atlassian`

| Key | Default | Effect |
|---|---|---|
| `AtlassianDocumentName` | `%MyDocuments%\%Organization%-%Date%.html` | Output file (Confluence storage format, despite the extension). |
| `AtlassianDocumentFileType` | `Full` | `Full` or `Object`. |
| `AtlassianTitleProperty` | `Intune documentation` | Page title. |
| `AtlassianOpenFile` | `$true` | Open when done. |

### `word`

Windows only.

| Key | Default | Effect |
|---|---|---|
| `WordDocumentName` | `%MyDocuments%\%Organization%-%Date%.docx` | Output file. |
| `WordDocumentTemplate` | *(none)* | A `.dotx` to build on. |
| `WordDocumentFormat` | `wdFormatDocumentDefault` | Any `WdSaveFormat` name, e.g. `wdFormatPDF`. |
| `WordDocumentationLevel` | `full` | How much to include. |
| `WordExportProperties` | `simple` | `simple` or `custom`; `WordCustomDisplayProperties` lists the columns for `custom`. |
| `WordAddCategories` / `WordAddSubCategories` | `$true` / `$true` | Category and sub-category headings. |
| `WordTitleProperty` / `WordSubjectProperty` | `Intune documentation` | Document properties. |
| `WordCoverPage` | `Ion (Dark)` | A cover page from the template's gallery. |
| `WordContentControls` | *(none)* | Content controls to fill. |
| `WordHeader1Style`, `WordHeader2Style`, `WordHeader3Style` | template defaults | Heading style names. |
| `WordTableStyle` | `Grid table 4 - Accent 3` | Style for settings tables. |
| `WordTableHeaderStyle`, `WordTableTextStyle` | template defaults | Styles for table header row and cell text. |
| `WordCategoryHeaderStyle`, `WordSubCategoryHeaderStyle` | template defaults | Styles for the category and sub-category headings. |
| `WordScriptTableStyle`, `WordScriptStyle` | template defaults | Styles for the table around an included script and the script text. |
| `WordTableCaptionPosition` | `below` | `above` or `below`. |
| `WordDocumentationLimitMaxLength` / `WordDocumentationLimitTruncateLength` | *(none)* | Truncate very long values; `WordDocumentationLimitAttach` (`$false`) attaches the full value instead. |
| `WordAttachJsonFile` | `$false` | Embed the policy json. |
| `WordOpenDocument` | `$true` | Open when done. |

## Where files land and how they are named

Document names accept placeholders, expanded at run time:

| Placeholder | Value |
|---|---|
| `%MyDocuments%` | The user's Documents folder. |
| `%Organization%` | The connected tenant's display name. |
| `%Date%` | `yyyy-MM-dd`. |
| `%DateTime%` | `yyyyMMdd-HHmm`. |

Placeholders are ordinary environment-variable expansion, so any `%VARIABLE%` in the
process environment works as well - `%BUILD_ARTIFACTSTAGINGDIRECTORY%` on an Azure DevOps
agent, for example.

With `*DocumentFileType = 'Object'` the name's folder is used and one file is written per
policy, named after the policy (invalid filename characters removed).

In a pipeline set every `*OpenFile` / `WordOpenDocument` to `$false`, and give absolute
paths - `%MyDocuments%` on a build agent is rarely where you want the artifact.

## Languages

`-Language` picks the strings file: 23 languages ship under `Config/LanguageStrings/`
(`Strings-<code>.json`): `cs de en es fr hu id it ja ko nl pl pt ru sv tr zh zh-chs zh-cht zh-hans zh-hant`. Setting names,
categories and values are localized; UI text is always English. The files are generated
and must not be edited by hand.

```powershell
Start-IMGraphBulkDocumentation -OutputFormat html -PolicyGroup Compliance -Language sv
```

## Which types document, and how

A policy is documented by the first of these that claims it:

1. A **handler** for its exact `@odata.type` (Conditional Access, Named Locations,
   Scope Tags, Role Definitions, Policy Sets, Kiosk, custom OMA-URI, ...).
2. An **input provider**: Settings Catalog policies walk their setting definitions;
   administrative templates their definition values; endpoint security intents their
   templates; Linux compliance its settings; and the profile providers use the
   per-type manifest files under `Config/ObjectInfo/`.
3. The **generic fallback** (when `FallbackDocumentation` is on): basic information plus
   one row per property.

The README's policy table says which of these each type gets (**yes** = 1 or 2,
**generic** = 3). Read-only types under *Intune Info*, ADMX Files and Multi Admin
Approval policies are not offered for documentation at all.

## Pipeline examples

Nightly HTML and Markdown from a service principal, no interactive state on the agent:

```powershell
Import-Module .\IntuneManagement.psd1
Use-IMSettingsStore -Memory
Connect-IMIntuneManagement -TenantId $env:TENANT_ID -AppId $env:APP_ID -Secret $env:APP_SECRET

$o = @{
    SkipNotConfigured = $true
    Outputs = @{
        html = @{ HTMLDocumentName = "$env:BUILD_ARTIFACTSTAGINGDIRECTORY\Intune-%Date%.html"; HTMLOpenFile = $false }
        md   = @{ MDDocumentName   = "$env:BUILD_ARTIFACTSTAGINGDIRECTORY\Intune.md"; MDDocumentSkipDate = $true; MDIncludeCSS = $false; MDOpenFile = $false }
    }
}
Start-IMGraphBulkDocumentation -OutputFormat 'html,md' -PolicyGroup DeviceConfiguration, Compliance, EndpointSecurity -Options $o
```

Document an export artifact produced by an earlier stage, one Markdown file per policy,
without the source tenant:

```powershell
Connect-IMIntuneManagement -Provider OAuth -ManagedIdentity     # any tenant the identity can read
$o = @{ Outputs = @{ md = @{ MDDocumentName = 'D:\docs\intune\index.md'; MDDocumentFileType = 'Object'; MDOpenFile = $false } } }
Start-IMGraphBulkDocumentation -OutputFormat md -SourceFolder D:\artifacts\IntuneExport -Options $o
```

Your own report from the data:

```powershell
Get-IMGraphPolicies -PolicyType CompliancePolicies |
    ForEach-Object {
        $d = Get-IMGraphDocumentation -PolicyObject $_ -Options @{ ExcludeAssignments = $true }
        [PSCustomObject]@{ Policy = $_.Name; Settings = $d.FilteredSettings.Count }
    } | Sort-Object Settings -Descending
```

Confluence, one page per policy, pushed with the REST API:

```powershell
$o = @{ Outputs = @{ atlassian = @{ AtlassianDocumentName = 'D:\out\intune.html'; AtlassianDocumentFileType = 'Object'; AtlassianOpenFile = $false } } }
Start-IMGraphBulkDocumentation -OutputFormat atlassian -PolicyGroup Compliance -Options $o
Get-ChildItem D:\out\*.html | ForEach-Object {
    $body = @{ type = 'page'; title = $_.BaseName; space = @{ key = 'INTUNE' }
               body = @{ storage = @{ value = (Get-Content $_ -Raw); representation = 'storage' } } } | ConvertTo-Json -Depth 6
    Invoke-RestMethod -Uri "$confluence/rest/api/content" -Method Post -Headers $auth -ContentType 'application/json' -Body $body
}
```
