# Automation Examples

One page of copy-paste recipes for the exported cmdlets. Everything the UI
does runs through these same commands, so anything shown here can be
scheduled, piped, or scripted.

```powershell
Import-Module .\IntuneManagement.psd1
```

All exported commands carry the `IM` prefix (`Connect-IMIntuneManagement`,
`Start-IMGraphBulkExport`, ...). Most commands take `-TokenId`; when omitted
they use the default token from the last `-DefaultToken` connect.

## Connect

```powershell
# Interactive sign-in (browser / WAM). Cached sessions resume silently.
Connect-IMIntuneManagement -DefaultToken

# Specific user / force a fresh prompt
Connect-IMIntuneManagement -Interactive -User admin@contoso.com -ForceInteractive

# Unattended: client secret (app registration)
Connect-IMIntuneManagement -TenantId $tid -AppId $appId -Secret $secret -DefaultToken

# Unattended: certificate
Connect-IMIntuneManagement -TenantId $tid -AppId $appId -CertificatePath .\auth.pfx

# Device code (headless box, sign in from another device)
Connect-IMIntuneManagement -DeviceCode

# Sovereign clouds
Connect-IMIntuneManagement -Cloud USGov -GCCType High -DefaultToken
```

Several tenants can be connected at once; each connect returns a token whose
`Id` you pass as `-TokenId` to target that tenant.

## Settings

```powershell
# Read and write by KEY - no storage paths, no internal functions
Get-IMSetting GraphPageSize
Get-IMSetting ExportFolder -Detailed        # value + where it came from
Set-IMSetting GraphPageSize 999
Remove-IMSetting ExportFolder               # back to the registered default

# What is configurable at all
Get-IMSettingDefinition -Section IntuneManager

# Per-tenant override (reads prefer it over the global value)
Set-IMSetting ExportFolder '\\server\intune\contoso' -Scope Tenant -TenantID $tid
```

Writes persist to the registry (Windows) or the settings file. On a shared
worker, where neither should be touched, run against an in-memory store:

```powershell
Use-IMSettingsStore -Memory                 # nothing from here on hits disk
Import-IMSettingsStore .\config\run.json    # a checked-in configuration
Set-IMSetting UseBatchAPI $true
Get-IMSettingsStore                         # assert the store before trusting it
```

`IM_SETTINGS_STORE=Memory` does the same from the first line of the module
load. See [Settings](Settings.md) for the store modes and precedence rules.

## Export

```powershell
# Everything, to one folder, assignments included
Start-IMGraphBulkExport -ExportFolder C:\IntuneExport -ExportAssignments $true

# Only some types / groups, name filter
Start-IMGraphBulkExport -ExportFolder C:\IntuneExport -PolicyType SettingsCatalog,CompliancePolicies
Start-IMGraphBulkExport -ExportFolder C:\IntuneExport -PolicyGroup DeviceConfiguration -Filter 'PROD-*'

# Reusable settings: build once, save, schedule the file
$s = [IntuneManagerExportSettings]::new()
$s.ExportFolder      = 'C:\IntuneExport'
$s.ExportAssignments = $true
$s.AddCompanyName    = $false
Save-IMGraphBulkExportSettings -Path C:\Jobs\nightly-export.json -ExportSettings $s
Start-IMGraphBulkExport -SettingsFile C:\Jobs\nightly-export.json

# Single policies via the pipeline
Get-IMGraphPolicies -PolicyType SettingsCatalog | Where-Object Name -like 'Win11*' |
    Export-IMGraphPolicy -ExportSettings $s
```

## Import

```powershell
# Import a full export folder (assignments + scope tags translated)
Start-IMGraphBulkImport -ImportFolder C:\IntuneExport -ImportAssignments $true -ImportScopeTags $true

# Only matching files / only some groups
Start-IMGraphBulkImport -ImportFolder C:\IntuneExport -Filter 'PROD-*' -PolicyGroup DeviceConfiguration

# Import behavior when the object already exists (-ImportType):
#   alwaysImport (default) | skipIfExist | update | replace | replace_with_assignments
Start-IMGraphBulkImport -ImportFolder C:\IntuneExport -ImportType update
```

Cross-tenant: the export folder's `MigrationTable.json` translates group and
dependency references automatically. Groups that do not exist in the target
tenant are created during import (settings `CreateGroupOnImport` and
`ConvertSyncedGroupOnImport`, both on by default). Policy sets re-point their
member references by display name.

## Documentation

```powershell
# One object -> result object (inspect or serialize yourself)
$p = Get-IMGraphPolicies -PolicyType CompliancePolicies | Select-Object -First 1
Get-IMGraphDocumentation -PolicyObject $p -Language en

# Bulk: whole policy groups to a single HTML file
Start-IMGraphBulkDocumentation -OutputFormat html -PolicyGroup DeviceConfiguration -Options @{
    Outputs = @{ html = @{ HTMLDocumentName = 'C:\Docs\Intune.html'; HTMLOpenFile = $false } }
}

# Several formats in one run, each with its own settings
Start-IMGraphBulkDocumentation -OutputFormat 'html,word,csv' -PolicyType SettingsCatalog -Options @{
    Outputs = @{
        html = @{ HTMLDocumentName = 'C:\Docs\Intune.html'; HTMLOpenFile = $false }
        word = @{ WordDocumentName = 'C:\Docs\Intune.docx'; WordOpenDocument = 'false' }
        csv  = @{ CSVDocumentationPath = 'C:\Docs\CSV' }
    }
}

# Localized output
Start-IMGraphBulkDocumentation -OutputFormat html -PolicyGroup CompliancePolicies -Language sv

# Document an EXPORT folder instead of the live tenant
Start-IMGraphBulkDocumentation -OutputFormat md -SourceFolder C:\IntuneExport

# Pipeline: document exactly the policies you select
Get-IMGraphPolicies -PolicyType ConditionalAccess | Start-IMGraphBulkDocumentation -OutputFormat json
```

Formats: `html`, `word`, `csv`, `md`, `json`, `atlassian`. Frequently used
per-format option keys (full list: `Get-IMDocumentationOutput`):

| Format | Keys |
| --- | --- |
| html | `HTMLDocumentName`, `HTMLDocumentFileType` (`Full`/`Object`), `HTMLCSSFile`, `HTMLOpenFile` |
| word | `WordDocumentName`, `WordDocumentTemplate`, `WordCoverPage`, `WordOpenDocument` |
| csv | `CSVDocumentationPath`, `CSVDelimiter`, `CSVAddObjectType` |
| md | `MDDocumentName`, `MDDocumentFileType`, `MDIncludeCSS` |
| json | `JSONDocumentName`, `JSONOutputFileType` |

`*DocumentFileType = 'Object'` writes one file per policy instead of one big
document.

## Compare

```powershell
# Live tenant vs an export folder (drift detection)
$provider = [CompareIntuneWithExportProvider]::new()
$provider.ExportPath = 'C:\IntuneExport'
Compare-IMGraphPolicy -IntuneWithExport $provider -PolicyGroupIds DeviceConfiguration

# Two selected policies against each other
$two = Get-IMGraphPolicies -PolicyType SettingsCatalog | Where-Object Name -in 'Baseline v1','Baseline v2'
Compare-IMGraphPolicy -Policies $two
```

## Bulk assignments

```powershell
# Add a group assignment (with an assignment filter) across a policy group
$s = [IntuneManagerAssignmentSettings]::new()
$s.Action      = 'Add'      # Add | Replace | Remove
$s.Assignments = @([PSCustomObject]@{
    TargetType = 'groupAssignmentTarget'   # or exclusionGroupAssignmentTarget,
    GroupId    = '<entra-group-id>'        #    allDevicesAssignmentTarget,
    FilterId   = '<filter-id>'             #    allLicensedUsersAssignmentTarget
    FilterType = 'include'
})
Set-IMGraphBulkAssignments -AssignmentSettings $s -PolicyGroup DeviceConfiguration -Filter 'PROD-*'

# Remove the same assignment again
Set-IMGraphBulkAssignments -AssignmentSettings $s -Action Remove -PolicyGroup DeviceConfiguration
```

## Bulk scope tags, copy, delete

```powershell
# Tag everything matching a name pattern
Set-IMGraphBulkScopeTags -Action Add -ScopeTagIds $tagId -PolicyType SettingsCatalog -Filter 'PROD-*'

# Copy every policy whose name contains the pattern, replacing it in the copy
# ("Test - Baseline" -> "Prod - Baseline"); re-run safe (existing names skipped)
Start-IMGraphBulkCopy -CopyFromPattern 'Test - ' -CopyToPattern 'Prod - ' -PolicyGroup DeviceConfiguration

# Delete by filter - test-prefix your filter, this is destructive
Start-IMGraphBulkDelete -Filter '[Test]*' -PolicyGroup DeviceConfiguration
```

## Lower-level building blocks

```powershell
# List objects (optionally with assignments)
Get-IMGraphPolicies -PolicyType SettingsCatalog -IncludeAssignments

# Load exported files back into policy objects (for selective import)
Get-ChildItem C:\IntuneExport\SettingsCatalog\*.json |
    Get-IMGraphPolicyFromFile | Import-IMGraphPolicy

# Raw Graph, with the module's auth, throttling, paging and batching
Invoke-IMMSGraphAPI -Url 'deviceManagement/managedDevices?$top=5' -AllPages
```
