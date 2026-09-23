# Command reference

Every command the module exports, with parameters, examples and what it returns.
Recipes that combine them are in [Examples.md](Examples.md).

## Conventions

- **Prefix.** The module exports with `DefaultCommandPrefix = 'IM'`, so `Get-GraphPolicies`
  is called as `Get-IMGraphPolicies`. The prefix goes after the verb: `Invoke-MSGraphAPI`
  becomes **`Invoke-IMMSGraphAPI`** (double M). Override with `Import-Module -Prefix`.
- **`-TokenId`.** Several tenants can be signed in at once. Commands that talk to Graph take
  `-TokenId` (from `Get-IMAuthToken`) to say which; omitted, they use the default token -
  the first sign-in, or the one connected with `-DefaultToken`.
- **Policy type and group ids** (`-PolicyType`, `-PolicyGroup`) are tab-completed, not
  validated: an unknown id is reported, not rejected at bind time. The ids are the ones in
  the policy table in [README.md](../README.md#supported-policy-types) - `DeviceConfiguration`,
  `SettingsCatalog`, `CompliancePolicies`, `DeviceEnrollments`, and so on.
- **`-WhatIf` / `-Confirm`** are supported on `Remove-IMGraphPolicy`, `Set-IMSetting`,
  `Remove-IMSetting`, `Use-IMSettingsStore`, `Export-IMSettingsStore` and
  `Import-IMSettingsStore`. **`Start-IMGraphBulkDelete` has neither** - the caller owns the
  confirmation.
- **Read-only** commands, safe for reporting and CI: `Get-IMAuthToken`, `Get-IMAccessibleTenant`,
  `Get-IMGraphEffectivePermissions`, `Get-IMGraphPolicies`, `Get-IMGraphPolicyFromFile`,
  `Compare-IMGraphPolicy`, `Get-IMGraphDocumentation`, `Get-IMDocumentationOutput`,
  `Get-IMSetting`, `Get-IMSettingDefinition`, `Get-IMSettingsStore`.

## Contents

| Area | Commands |
|---|---|
| [Signing in and tokens](#signing-in-and-tokens) | `Connect-IMIntuneManagement`, `Get-IMAuthToken`, `Get-IMAccessibleTenant`, `Get-IMGraphEffectivePermissions`, `Invoke-IMMSGraphAPI` |
| [Working with policies](#working-with-policies) | `Get-IMGraphPolicies`, `Get-IMGraphPolicyFromFile`, `Import-IMGraphPolicy`, `Export-IMGraphPolicy`, `Copy-IMGraphPolicy`, `Remove-IMGraphPolicy`, `Compare-IMGraphPolicy` |
| [Bulk operations](#bulk-operations) | `Start-IMGraphBulkExport`, `Start-IMGraphBulkImport`, `Start-IMGraphBulkCopy`, `Start-IMGraphBulkDelete`, `Set-IMGraphBulkAssignments`, `Set-IMGraphBulkScopeTags`, `Save-IMGraphBulkExportSettings` |
| [Documentation](#documentation) | `Get-IMGraphDocumentation`, `Start-IMGraphBulkDocumentation`, `Get-IMDocumentationOutput` |
| [Settings](#settings) | `Get-IMSetting`, `Set-IMSetting`, `Remove-IMSetting`, `Get-IMSettingDefinition`, `Use-IMSettingsStore`, `Get-IMSettingsStore`, `Export-IMSettingsStore`, `Import-IMSettingsStore` |
| [UI](#ui) | `Show-IMMainWindow` |

---

## Signing in and tokens

### Connect-IMIntuneManagement

Authenticate to Microsoft Graph. Each call registers a token; several tenants can be
live at once. Returns the token as a `PSCustomObject` (`TokenId`, `TenantId`,
`TenantName`, `Provider`, `Account`, `ExpiresOn`).

The parameter set chooses the credential. `-Provider` chooses the implementation:
`MSAL` (default) or `OAuth`; the saved *Active authentication provider* setting is the
fallback.

| Set | Parameters | Notes |
|---|---|---|
| Interactive (default) | `-Interactive` `[-TenantId]` `[-User]` `[-ForceInteractive]` `[-AuthenticationBroker]` `[-Browser]` | Silent from cache first, then browser. `-AuthenticationBroker` = WAM, Windows + PS 7 only. With `-Provider OAuth` this routes to device code. |
| DeviceCode | `-DeviceCode` `[-TenantId]` `[-AppId]` | Code shown here, browser step on any device. MFA / FIDO2 capable. |
| Secret | `-TenantId` `-AppId` `-Secret` | App registration with client secret. |
| Certificate | `-TenantId` `-AppId` `-Certificate <thumbprint or X509Certificate2>` | Looked up in `Cert:\CurrentUser\My`, then `Cert:\LocalMachine\My`. |
| CertificatePath | `-TenantId` `-AppId` `-CertificatePath` `[-CertificatePassword <SecureString>]` | `.pfx` file. |
| Token | `-Token` | Bring your own Graph bearer token. Cannot be refreshed. |
| ManagedIdentity | `-ManagedIdentity` `[-AppId]` | System-assigned, or user-assigned via `-AppId`. `-Provider OAuth`. |
| OAuthFederated | `-TenantId` `-AppId` `-FederatedTokenFile` or `-FederatedToken` | Workload identity federation (AKS, GitHub Actions OIDC). `-Provider OAuth`. |
| OAuthCredential | `-TenantId` `-AppId` `-Credential <PSCredential>` | ROPC. Non-MFA accounts only. `-Provider OAuth`. |

Common to every set: `-Cloud Public|USGov|USGovDOD|China` (default from the
*DefaultCloud* setting), `-DefaultToken` (make this the default token), `-Provider`.

```powershell
# Interactive, resumes silently from the cache when it can
Connect-IMIntuneManagement -Interactive

# A specific tenant, and force a fresh prompt
Connect-IMIntuneManagement -Interactive -TenantId contoso.onmicrosoft.com -ForceInteractive

# Unattended: client secret
Connect-IMIntuneManagement -TenantId contoso.onmicrosoft.com -AppId $appId -Secret $secret

# Unattended: certificate thumbprint
Connect-IMIntuneManagement -TenantId contoso.onmicrosoft.com -AppId $appId -Certificate 'A1B2C3...'

# Unattended: .pfx
Connect-IMIntuneManagement -TenantId $t -AppId $a -CertificatePath C:\certs\app.pfx `
    -CertificatePassword (ConvertTo-SecureString $pw -AsPlainText -Force)

# Device code - headless box, sign in from a phone
Connect-IMIntuneManagement -DeviceCode

# OAuth provider: managed identity on an Azure VM / Function / Automation account
Connect-IMIntuneManagement -Provider OAuth -ManagedIdentity

# OAuth provider: GitHub Actions OIDC / AKS workload identity
Connect-IMIntuneManagement -Provider OAuth -TenantId $t -AppId $a -FederatedTokenFile $env:AZURE_FEDERATED_TOKEN_FILE

# Bring your own token (the only way to reach APIs closed to public clients)
Connect-IMIntuneManagement -Token $bearer

# Sovereign cloud
Connect-IMIntuneManagement -Interactive -Cloud USGov
```

### Get-IMAuthToken

List the tokens currently held, across providers. Returns `IMAuthToken[]` -
`TokenId`, `TenantId`, `TenantName`, `Provider`, `Account`/`UPN`, `AppId`, `ExpiresOn`,
`IsDefault`.

| Parameter | Type | Notes |
|---|---|---|
| `-TokenId` | int | One token. |
| `-Provider` | string | Filter by provider. |
| `-TenantId` | string | Filter by tenant. |

```powershell
Get-IMAuthToken

# Pick a tenant's token for a later command
$prod = Get-IMAuthToken | Where-Object TenantName -eq 'Contoso Prod'
Get-IMGraphPolicies -PolicyType CompliancePolicies -TokenId $prod.TokenId
```

### Get-IMAccessibleTenant

List the tenants the signed-in account can reach - its home tenant and every tenant it
is a guest in. Graph cannot answer this; the list comes from Azure Resource Manager, so
the app registration needs the delegated permission *Azure Service Management /
user_impersonation*, and only the MSAL provider implements it. Without either, the
command warns and returns nothing.

| Parameter | Type | Notes |
|---|---|---|
| `-TokenId` | int | Ask the provider that owns this token. |

```powershell
Get-IMAccessibleTenant

# Confirm a guest tenant is reachable, then connect to it silently
$guest = Get-IMAccessibleTenant | Where-Object displayName -eq 'Fabrikam'
Connect-IMIntuneManagement -Interactive -TenantId $guest.tenantId
```

### Get-IMGraphEffectivePermissions

What the signed-in identity can actually do, per policy type: the app's token scopes
combined with the user's Intune RBAC or Entra directory roles. Lets a script learn it is
read-only for a type *before* a bulk import, instead of collecting 403s halfway through.
One row per policy type: `Id`, `Required`, `TokenAccess`, `RoleAccess`,
`EffectiveAccess`, `Result` (Match / Read-only / No access), plus the raw
`TokenLevel` / `RbacLevel` / `EffectiveLevel` (Full / Limited / None) and `Reason`.

| Parameter | Type | Notes |
|---|---|---|
| `-TokenId` | int | Evaluate this token. |
| `-PolicyType` | string[] | Only these type ids. |
| `-Raw` | switch | Return the RBAC context itself (allowed resource actions, the catalog, the raw response) instead of the table. |

App-only tokens bypass Intune RBAC, so `RbacLevel` is `$null` and the token level is the
answer. Scope tags are not modelled. The answer is cached per token and refreshed when
the token is re-issued.

```powershell
# What can this user not change, and why?
Get-IMGraphEffectivePermissions | Where-Object EffectiveLevel -ne Full |
    Format-Table Id, TokenLevel, RbacLevel, EffectiveLevel, Reason

# Gate a bulk import on write access to the types it touches
$blocked = Get-IMGraphEffectivePermissions -PolicyType DeviceConfiguration, SettingsCatalog |
    Where-Object EffectiveLevel -ne Full
if ($blocked) { throw "Read-only for: $($blocked.Id -join ', ')" }
```

### Invoke-IMMSGraphAPI

The low-level Graph call every other command uses: resolves the token, adds headers,
handles throttling, paging, batching and claims challenges. Use it for anything the
policy commands do not cover.

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-Url` | string | required | Relative (`deviceManagement/deviceCategories`) or absolute. |
| `-HttpMethod` (`-Method`) | string | `GET` | `GET`, `POST`, `PATCH`, `PUT`, `DELETE`, `OPTIONS`. |
| `-Content` (`-Body`) | string | | Request body, JSON. |
| `-Headers`, `-AdditionalHeaders` | hashtable | | |
| `-GraphVersion` | string | `beta` | `beta` or `v1.0`. The *UseGraphV1* setting flips the default. |
| `-ODataMetadata` | string | `full` | `full`, `minimal`, `none`, `skip`. |
| `-AllPages` | switch | | Follow `@odata.nextLink` to the end. |
| `-PageSize` | int | | `$top` for the first page. |
| `-Batch` | switch | | Queue into the current `$batch` instead of sending. |
| `-Outfile` | string | | Save the response body to a file. |
| `-FullResponseObject` | switch | | Return status code and headers as well as the body. |
| `-NoError` | switch | | Return `$null` on failure instead of logging an error. |
| `-SkipAuthentication` | switch | | |
| `-TokenId` | int | default token | |

Returns the parsed body (collections under `.value`). A Multi Admin Approval 412 comes
back with `ApprovalPending = $true` and the `ApprovalCode` on the result.

```powershell
# List with paging
(Invoke-IMMSGraphAPI -Url 'deviceManagement/deviceCategories' -AllPages).value

# Create
Invoke-IMMSGraphAPI -Url 'deviceManagement/deviceCategories' -HttpMethod POST `
    -Content '{"displayName":"Kiosks","description":"Shared kiosk devices"}'

# Status code as well as body
$r = Invoke-IMMSGraphAPI -Url "deviceManagement/deviceCategories/$id" -HttpMethod DELETE -FullResponseObject
$r.StatusCode
```

---

## Working with policies

The policy commands pass **policy objects** (`IntunePolicyBase`) down the pipeline. Get
them from a tenant with `Get-IMGraphPolicies` or from files with
`Get-IMGraphPolicyFromFile`; `Import`, `Export`, `Copy`, `Remove` and `Compare` consume
them.

### Get-IMGraphPolicies

List policies of one or more types or groups. Shared list URLs are coalesced and
batched. Returns `IntunePolicyBase[]` - each with `Name`, `Id`, `PolicyType`, `Object`
(the Graph JSON) and, with `-IncludeAssignments`, `Object.assignments`.

| Set | Parameters | Notes |
|---|---|---|
| PolicyType | `-PolicyType <string[]>` (position 0, pipeline) | |
| PolicyGroup | `-PolicyGroup <string[]>` (position 0, pipeline) | Every type in the group. |
| Paging | `-Paging NextPage\|AllRemainingPages` | Continue a paged listing. |

Common: `-NameFilter <string>` (server-side where the endpoint allows, always re-checked
client-side), `-IncludeAssignments`, `-SinglePage`, `-TokenId`.

```powershell
Get-IMGraphPolicies -PolicyType CompliancePolicies
Get-IMGraphPolicies -PolicyGroup DeviceConfiguration -IncludeAssignments
Get-IMGraphPolicies -PolicyType SettingsCatalog -NameFilter 'Baseline'
'CompliancePolicies', 'ConditionalAccess' | Get-IMGraphPolicies
```

### Get-IMGraphPolicyFromFile

Load exported json files as policy objects, resolving each file's policy type from its
`@odata.type` and folder. The result is what `Import-IMGraphPolicy` and
`Compare-IMGraphPolicy` accept.

| Parameter | Type | Notes |
|---|---|---|
| `-InputObject` (`-FileInfo`) | `IO.FileInfo[]`, pipeline | The files. |
| `-FromPolicyTypes` | `IntunePolicyTypeBase[]` | Narrow the candidate types (e.g. when the folder name does not match). |
| `-TenantId` | string | Record the source tenant on the objects. |

```powershell
Get-ChildItem C:\IntuneExport\CompliancePolicies\*.json | Get-IMGraphPolicyFromFile
```

### Import-IMGraphPolicy

Create the piped policies in the tenant. Types are imported in dependency order (scope
tags and filters first, policy sets last), references are translated where the export
carries the information to do so - assignment groups, scope tags, targeted apps, ADMX
setting ids - and per-type hooks run (Win32 content upload, ADMX/ADML files, Terms of
Use PDF). Returns one result per policy with `ImportedObject` (the created policy),
`SourceObject` and `Success`.

| Parameter | Type | Notes |
|---|---|---|
| `-InputObject` | `IntunePolicyBase[]`, pipeline | From `Get-IMGraphPolicyFromFile` or another tenant's `Get-IMGraphPolicies`. |
| `-TokenId` | int | Destination tenant. |

```powershell
# Everything under a folder
Get-ChildItem C:\IntuneExport -Recurse -Filter *.json | Get-IMGraphPolicyFromFile | Import-IMGraphPolicy

# Tenant to tenant without touching disk
$src = (Get-IMAuthToken | Where-Object TenantName -eq 'Lab').TokenId
$dst = (Get-IMAuthToken | Where-Object TenantName -eq 'Prod').TokenId
Get-IMGraphPolicies -PolicyType CompliancePolicies -TokenId $src | Import-IMGraphPolicy -TokenId $dst
```

### Export-IMGraphPolicy

Write the piped policies to json under the export folder, one file per policy, in the
type's subfolder. Assignments, scope-tag names and organization tokens follow the
`IntuneManagerExportSettings` passed in.

| Parameter | Type | Notes |
|---|---|---|
| `-InputObject` | `IntunePolicyBase[]`, pipeline | |
| `-ExportSettings` | `IntuneManagerExportSettings` | Required. `[IntuneManagerExportSettings]::new()` starts from the saved settings. |
| `-PassThru` | switch | Emit the full path of each written file. |

```powershell
$s = [IntuneManagerExportSettings]::new()
$s.ExportFolder      = 'C:\IntuneExport'
$s.ExportAssignments = $true
Get-IMGraphPolicies -PolicyType CompliancePolicies | Export-IMGraphPolicy -ExportSettings $s -PassThru
```

### Copy-IMGraphPolicy

Create a copy of each piped policy - in the same tenant, or in another with `-TokenId`.
Returns the new policies.

| Parameter | Type | Notes |
|---|---|---|
| `-InputObject` | `IntunePolicyBase[]`, pipeline | |
| `-Name` | string | Required. Name of the copy. With several inputs, use the patterns instead. |
| `-Description` | string | |
| `-CopyFromPatternName` / `-CopyFromPatternDescription` | string | Substring in the source name/description replaced by `-Name` / `-Description` - for copying many at once. |
| `-ScopeTagIds` | string[] | Scope tags for the copy. Omitted, the copy inherits the source's; an empty array clears them. |
| `-TokenId` | int | Destination tenant. |

The Copy dialog in the UI pre-fills the name from the type's `CopyDefaultName` template
where one is set (`%Name% Copy`); from a script the name is always what you pass.

```powershell
Get-IMGraphPolicies -PolicyType CompliancePolicies -NameFilter 'Pilot - W11' |
    Copy-IMGraphPolicy -Name 'Prod - W11'

# Many at once: "Pilot - X" -> "Prod - X"
Get-IMGraphPolicies -PolicyGroup DeviceConfiguration -NameFilter 'Pilot - ' |
    Copy-IMGraphPolicy -CopyFromPatternName 'Pilot - ' -Name 'Prod - '
```

### Remove-IMGraphPolicy

Delete the piped policies. Supports `-WhatIf` and `-Confirm` (ConfirmImpact Medium).
Batched when batching is on. Returns the deleted policies.

| Parameter | Type | Notes |
|---|---|---|
| `-InputObject` | `IntunePolicyBase[]`, pipeline | |
| `-TokenId` | int | |

```powershell
Get-IMGraphPolicies -PolicyType DeviceCategories -NameFilter '[Test]' | Remove-IMGraphPolicy -WhatIf
Get-IMGraphPolicies -PolicyType DeviceCategories -NameFilter '[Test]' | Remove-IMGraphPolicy -Confirm:$false
```

### Compare-IMGraphPolicy

Compare two or more policies property by property, or run one of the compare providers
over pairs. Returns rows of `Property`, `Value1`, `Value2`, `Match`.

| Set | Parameters | Notes |
|---|---|---|
| Direct | `-Policies <object[]>` (position 0, pipeline) | Two or more policy objects. |
| ExportFiles | `-ExportFiles <CompareExportFilesProvider>` | Each tenant policy vs its exported file. |
| IntuneWithExport | `-IntuneWithExport <CompareIntuneWithExportProvider>` | |
| NamedObjects | `-NamedObjects <CompareNamedObjectsProvider>` | Pairs matched by name pattern. |
| ExportedFolders | `-ExportedFolders <CompareExportedFoldersProvider>` | Two export folders. |

The provider sets take `-PolicyGroupIds <string[]>` to limit the groups compared.
Provider classes are in `Classes/CompareClasses.ps1`; see [Compare.md](Compare.md).

```powershell
$a, $b = Get-IMGraphPolicies -PolicyType CompliancePolicies -NameFilter 'W11' | Select-Object -First 2
Compare-IMGraphPolicy -Policies @($a, $b) | Where-Object Match -eq $false

# A tenant against last night's export
$p = [CompareIntuneWithExportProvider]::new()
$p.ExportPath = 'C:\IntuneExport'
Compare-IMGraphPolicy -IntuneWithExport $p -PolicyGroupIds DeviceConfiguration
```

---

## Bulk operations

The bulk commands are the headless form of the Bulk menu. Each returns a summary object
with counts and `Duration`. `-PolicyType` and `-PolicyGroup` select what to operate on;
`-Filter` is a name filter: a literal, case-insensitive substring of the policy name, the
same rule as 3.x. `-Filter '[Test]'` selects names containing the text `[Test]`; there is no
regex or wildcard syntax.

### Start-IMGraphBulkExport

Export whole policy groups or types to disk. Precedence: `-SettingsFile` < `-ExportSettings`
< explicit parameters. Returns `Types`, `Policies`, `Failed`, `Duration`.

| Parameter | Type | Notes |
|---|---|---|
| `-ExportFolder` | string | Root folder. Required unless a settings source provides it. |
| `-PolicyType` / `-PolicyGroup` | string[] | Default: every type whose group allows export. |
| `-Filter` | string | Name filter. |
| `-ExportAssignments` | bool | |
| `-AddCompanyName` | bool | Add a tenant-name folder level. |
| `-ExportSettings` | `IntuneManagerExportSettings` | A settings instance. |
| `-SettingsFile` | string | A file written by `Save-IMGraphBulkExportSettings`. |
| `-TokenId` | int | |

```powershell
Start-IMGraphBulkExport -ExportFolder C:\IntuneExport -PolicyGroup DeviceConfiguration
Start-IMGraphBulkExport -ExportFolder C:\IntuneExport -ExportAssignments $true -Filter 'Baseline'
Start-IMGraphBulkExport -SettingsFile .\nightly-export.json
```

### Start-IMGraphBulkImport

Import an export folder. Groups are processed in dependency order. Returns `Groups`,
`Imported`, `Duration`.

| Parameter | Type | Notes |
|---|---|---|
| `-ImportFolder` | string | Required. |
| `-PolicyGroup` | string[] | Default: every group that allows import. |
| `-Filter` | string | |
| `-ImportType` | string | `alwaysImport` (default), `skipIfExist`, `update`, `replace`, `replace_with_assignments`. |
| `-ImportAssignments`, `-ImportScopeTags`, `-ReplaceDependencyIDs` | bool | Persisted as the like-named settings for the session. |
| `-TokenId` | int | |

```powershell
Start-IMGraphBulkImport -ImportFolder C:\IntuneExport
Start-IMGraphBulkImport -ImportFolder C:\IntuneExport -PolicyGroup DeviceConfiguration -Filter 'Baseline' -ImportType update
```

### Start-IMGraphBulkCopy

Copy every policy whose name contains a pattern, to the same name with the pattern
replaced. Returns `Types`, `Copied`, `Skipped`, `FailedTypes`, `UnknownSelectors`, `Duration`.

| Parameter | Type | Notes |
|---|---|---|
| `-CopyFromPattern` | string | Required. |
| `-CopyToPattern` | string | Required. |
| `-PolicyType` / `-PolicyGroup` | string[] | Default: every type whose group allows copy. |
| `-TokenId` | int | |

```powershell
Start-IMGraphBulkCopy -CopyFromPattern 'Pilot - ' -CopyToPattern 'Prod - ' -PolicyGroup DeviceConfiguration
```

### Start-IMGraphBulkDelete

Delete every policy in the selected groups that matches the filter. **No `-WhatIf`, no
confirmation** - list first with `Get-IMGraphPolicies` using the same filter. Groups are
deleted in reverse dependency order. Returns `Groups`, `Deleted`, `Duration`.

| Parameter | Type | Notes |
|---|---|---|
| `-PolicyGroup` | string[] | Required. |
| `-Filter` | string | Empty means every object in the group. |
| `-TokenId` | int | |

```powershell
Get-IMGraphPolicies -PolicyGroup DeviceConfiguration -NameFilter '[Test]' | Select-Object Name   # look first
Start-IMGraphBulkDelete -PolicyGroup DeviceConfiguration -Filter '[Test]'
```

### Set-IMGraphBulkAssignments

Add, replace or remove assignments across types or groups. Handles the three assignment
shapes - simple targets, app assignments with intent and per-platform settings, health
scripts with schedules. App types that cannot take a filter (web apps) are assigned
without it, with a log line. Returns `Types`, `PoliciesScanned`, `PoliciesMatched`,
`PoliciesUpdated`, `PoliciesSkipped`, `PoliciesFailed`, `PoliciesUnsupported`, `Duration`.

| Parameter | Type | Notes |
|---|---|---|
| `-Action` | string | `Add`, `Replace`, `Remove`. |
| `-Assignments` | `PSCustomObject[]` | Descriptors: `TargetType` (`groupAssignmentTarget`, `exclusionGroupAssignmentTarget`, `allDevicesAssignmentTarget`, `allLicensedUsersAssignmentTarget`), `GroupId`, `GroupName`, `FilterId`, `FilterType` (`include`/`exclude`), `Intent` (apps), `Settings` (hashtable per platform). |
| `-Filter` | string | Name filter. |
| `-PolicyType` / `-PolicyGroup` | string[] | |
| `-AssignmentSettings` | `IntuneManagerAssignmentSettings` | Alternative to the three above. |
| `-TokenId` | int | |

```powershell
$a = [PSCustomObject]@{
    TargetType = 'groupAssignmentTarget'
    GroupId    = '<entra-group-id>'
    GroupName  = 'All Helpdesk Devices'
    FilterId   = '<assignment-filter-id>'
    FilterType = 'include'
}
Set-IMGraphBulkAssignments -Action Add -Assignments @($a) -PolicyGroup DeviceConfiguration -Filter 'Baseline'

# Apps: required install to a group
$app = [PSCustomObject]@{ TargetType = 'groupAssignmentTarget'; GroupId = $gid; GroupName = 'Pilot'; Intent = 'required' }
Set-IMGraphBulkAssignments -Action Add -Assignments @($app) -PolicyType Applications -Filter 'Office'

Set-IMGraphBulkAssignments -Action Remove -Assignments @($a) -PolicyGroup DeviceConfiguration
```

See [BulkAssignments.md](BulkAssignments.md) for the settings hashtables.

### Set-IMGraphBulkScopeTags

Add, replace or remove scope tags across types or groups; `-CleanupOrphans` removes
references to tags that no longer exist. Returns `Types`, `PoliciesScanned`,
`PoliciesMatched`, `PoliciesUpdated`, `PoliciesSkipped`, `PoliciesFailed`, `Duration`.

| Parameter | Type | Notes |
|---|---|---|
| `-Action` | string | `Add`, `Replace`, `Remove`. |
| `-ScopeTagIds` | string[] | Tag ids (`0` is the default tag). |
| `-Filter` | string | |
| `-CleanupOrphans` | bool | |
| `-PolicyType` / `-PolicyGroup` | string[] | |
| `-ScopeTagSettings` | `IntuneManagerScopeTagSettings` | Alternative to the above. |
| `-TokenId` | int | |

```powershell
Set-IMGraphBulkScopeTags -Action Add -ScopeTagIds 3, 4 -PolicyGroup DeviceConfiguration
Set-IMGraphBulkScopeTags -CleanupOrphans $true
```

### Save-IMGraphBulkExportSettings

Write an export configuration to a json file that `Start-IMGraphBulkExport -SettingsFile`
reads - the way to schedule the same export nightly.

| Parameter | Type | Notes |
|---|---|---|
| `-Path` | string | Required. |
| `-ExportSettings` | `IntuneManagerExportSettings` | Required. |
| `-PolicyGroup` / `-PolicyType` | string[] | What the file selects. |

```powershell
$s = [IntuneManagerExportSettings]::new()
$s.ExportFolder = '\\server\intune\exports'; $s.ExportAssignments = $true
Save-IMGraphBulkExportSettings -Path .\nightly-export.json -ExportSettings $s -PolicyGroup DeviceConfiguration, Compliance
Start-IMGraphBulkExport -SettingsFile .\nightly-export.json
```

---

## Documentation

### Get-IMGraphDocumentation

Document one policy and return the result object - `BasicInfo`, `FilteredSettings`,
`Assignments`, `Scripts`, `CustomTables` and the rest - without writing a file. Useful
for building your own report.

| Parameter | Type | Notes |
|---|---|---|
| `-PolicyObject` | pipeline | A policy from `Get-IMGraphPolicies` or `Get-IMGraphPolicyFromFile`. |
| `-Language` | string | `en` default; any language the strings ship in. |
| `-Options` | hashtable | See [Documentation.md](Documentation.md#options) for every key. |

```powershell
$p   = Get-IMGraphPolicies -PolicyType ConditionalAccess | Select-Object -First 1
$doc = Get-IMGraphDocumentation -PolicyObject $p
$doc.FilteredSettings | Format-Table Name, Value
```

### Start-IMGraphBulkDocumentation

Document many policies through one or more output providers. Selects by object, type,
group, or an export folder. Returns the run summary; files land where each output's
options say.

| Set | Parameters |
|---|---|
| PolicyObject | `-PolicyObject` (pipeline) |
| PolicyType | `-PolicyType <string[]>` |
| PolicyGroup | `-PolicyGroup <string[]>` |
| Folder | `-SourceFolder <string>` - an export folder. Still needs a signed-in tenant (any tenant) for setting definitions and templates; source-tenant names come from the export's migration table or stay as ids. |

Common: `-OutputFormat <string>` (required, position 0; comma-separated: `html`, `md`,
`word`, `json`, `csv`, `atlassian`), `-Language`, `-Options`.

```powershell
Start-IMGraphBulkDocumentation -OutputFormat html -PolicyGroup Compliance
Start-IMGraphBulkDocumentation -OutputFormat 'md,json' -SourceFolder C:\IntuneExport
Get-IMGraphPolicies -PolicyType SettingsCatalog -NameFilter 'Baseline' | Start-IMGraphBulkDocumentation -OutputFormat word
```

### Get-IMDocumentationOutput

List the registered output providers - `Name` and `Value` (what `-OutputFormat` matches).
No parameters.

```powershell
Get-IMDocumentationOutput | Format-Table Name, Value
```

---

## Settings

Settings are read by key. Three stores: the registry (Windows default), a json file, or
memory. Effective value = tenant override, else global, else the registered default.
See [Settings.md](Settings.md).

### Get-IMSetting

| Parameter | Type | Notes |
|---|---|---|
| `-Key` | string, position 0, pipeline | Omit for every registered setting (implies `-Detailed`). |
| `-Scope` | string | `Effective` (default), `Global`, `Tenant`. |
| `-TenantID` | string | Default: the connected tenant. |
| `-SubPath` | string | For keys stored under a sub-path (none for most). |
| `-Detailed` | switch | Include `Source` - Tenant, Global or Default. |

```powershell
Get-IMSetting ExportFolder
Get-IMSetting UseBatchAPI -Detailed
Get-IMSetting | Where-Object Source -ne 'Default'      # everything actually configured
```

### Set-IMSetting

Supports `-WhatIf`.

| Parameter | Type | Notes |
|---|---|---|
| `-Key` | string, position 0 | Required. |
| `-Value` | position 1 | Required; `$null` removes the value. |
| `-Scope` | string | `Global` (default) or `Tenant`. |
| `-TenantID`, `-SubPath`, `-PassThru` | | |

```powershell
Set-IMSetting ExportFolder 'C:\Intune\Export'
Set-IMSetting ExportFolder '\\server\intune\contoso' -Scope Tenant
Set-IMSetting UseParallelBatchAPI $true
```

### Remove-IMSetting

Remove a stored value so the next level applies. Supports `-WhatIf`.

| Parameter | Type | Notes |
|---|---|---|
| `-Key` | string, position 0 | Required. |
| `-Scope` | string | `Global` (default) or `Tenant`. |
| `-TenantID`, `-SubPath` | | |

```powershell
Remove-IMSetting ExportFolder -Scope Tenant
```

### Get-IMSettingDefinition

What the module knows about its settings: `Key`, `Title`, `Section`, `Type`,
`DefaultValue`, `Description`. Wildcards on both parameters.

| Parameter | Type |
|---|---|
| `-Key` | string, position 0 |
| `-Section` | string |

```powershell
Get-IMSettingDefinition | Format-Table Key, Section, Type, DefaultValue
Get-IMSettingDefinition -Key *Export*
```

### Use-IMSettingsStore

Choose the store for the rest of the session. Supports `-WhatIf`.

| Set | Parameters | Notes |
|---|---|---|
| Memory (default) | `-Memory` | Nothing is read from or written to the machine. |
| Json | `-Path <file>` (position 0) | Created if missing. |
| Registry | `-Registry` | Windows only. |

`-Seed` copies the persisted settings into the new store; `-PassThru` returns the store.

```powershell
# The runbook pattern: an empty store, then the run's configuration from source control
Use-IMSettingsStore -Memory
Import-IMSettingsStore -Path .\runbook-settings.json

Use-IMSettingsStore -Path 'D:\shared\IntuneManagement.json'
```

### Get-IMSettingsStore

Which store is active and whether it persists: `Mode`, `Persisted`, `Path`.
`-IncludeValues` adds every value.

```powershell
Get-IMSettingsStore
(Get-IMSettingsStore -IncludeValues).Values | Format-Table
```

### Export-IMSettingsStore

Write the whole active store to a json file - capture a working configuration once and
commit it. Supports `-WhatIf`. Takes `-Path` (position 0, required); the folder is
created and an existing file replaced.

```powershell
Export-IMSettingsStore -Path .\intune-settings.json
```

### Import-IMSettingsStore

Merge a json settings file into the active store (existing keys are overwritten, others
kept). Supports `-WhatIf`. Takes `-Path` (position 0, required). A missing file is
reported, not silently ignored.

```powershell
Use-IMSettingsStore -Memory
Import-IMSettingsStore -Path .\runbook-settings.json
Import-IMSettingsStore -Path .\baseline.json -WhatIf      # which store would this land in?
```

---

## UI

### Show-IMMainWindow

Show the application window for the active UI backend (WPF on Windows, Avalonia
elsewhere). Takes an optional `-View` to open on. Not available when the `UI` folder has
been removed from the deployment, and not for use from an ordinary `pwsh` session on
macOS - launch with `Start-Avalonia.command` there.

```powershell
Import-Module .\IntuneManagement.psd1
Show-IMMainWindow
```
