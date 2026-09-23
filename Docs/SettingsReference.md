# Settings reference

Every setting the module registers, by the section it appears under in the Settings
dialog. **Setting** is the name shown in the dialog; **Key** is what `Get-IMSetting` /
`Set-IMSetting` and a settings file use. Where values are stored and how tenant and global
values combine is in [Settings.md](Settings.md).

**Generated** by `Tools/Export-SettingsReference.ps1` from `Get-IMSettingDefinition` - do not edit by hand. 83 settings.

A runbook settings file for `Import-IMSettingsStore` is a json object keyed by these names:

```json
{ "ExportFolder": "D:\\exports", "ExportAssignments": true, "UseBatchAPI": true, "UseParallelBatchAPI": true }
```

## Contents

- [Authentication](#authentication) (8)
- [General](#general) (13)
- [Import/Export](#importexport) (30)
- [Intune](#intune) (8)
- [Intune Tools](#intune-tools) (1)
- [MS Graph General](#ms-graph-general) (5)
- [MSAL](#msal) (14)
- [OAuth](#oauth) (4)

## Authentication

| Setting | Key | Type | Default | Description |
|---|---|---|---|---|
| Active authentication provider | `ActiveAuthProvider` | List | `MSAL` | Which authentication backend to use. MSAL is the built-in default. MgGraph requires the Microsoft.Graph.Authentication PowerShell module to be installed. OAuth is a pure-PowerShell provider for automation (CI / scheduled tasks / managed identity / workload identity federation) - no SDK required. Values: `Microsoft Authentication Library`, `Microsoft Graph PowerShell SDK`, `Direct OAuth (no SDK)`. |
| Application Id | `EntraCustomAppId` | String |  | Custom Entra application (client) id used by MSAL and OAuth sign-in when no built-in application is selected above. The app registration needs the http://localhost redirect URI for browser-based login. |
| Authority | `EntraCustomAuthority` | String |  |  |
| Default cloud | `DefaultCloud` | List | `Public` | Microsoft cloud this tool signs in to by default. Public covers commercial + GCC commercial; USGov is GCC High; USGovDOD is GCC DoD; China is the Vianet cloud. Values: `Public (Global)`, `US Government (GCC High)`, `US Government (DoD)`, `China (Vianet)`. |
| Entra application | `EntraApp` | List |  | Built-in Entra application used for interactive sign-in by both the MSAL and OAuth providers. Leave empty to use the custom Application Id below, or the default Microsoft Graph PowerShell app. Values: `*** Do NOT use *** Microsoft Intune PowerShell`, `Microsoft Graph PowerShell`. |
| Interactive login timeout (seconds) | `MSGraphInteractiveTimeoutSec` | Int | `600` | Maximum time to wait for an interactive login to complete (MSAL embedded/broker and OAuth browser flows). Only applies to flows a user is waiting in front of; silent and app-secret token requests have their own much shorter cap. The default was 180 s, which killed legitimate sign-ins that needed an account picker plus an MFA approval - and the sign-in status has a Cancel button, so an abandoned window does not depend on this timeout. |
| Redirect URL | `EntraCustomAppRedirect` | String |  |  |
| Tenant Id | `EntraCustomTenantId` | String |  |  |

## General

| Setting | Key | Type | Default | Description |
|---|---|---|---|---|
| Add errors to PowerShell output | `LogOutputError` | Boolean | `true` | Write errors to the Error Output of the PS Host. If disabled, errors will be written as a Warning. Eg. disable this if automation should skip logging PowerShell errors. |
| Check for updates | `CheckForUpdates` | Boolean | `true` | Check GitHub if there is a later version available |
| Debug | `Debug` | Boolean | `false` |  |
| Environment color | `EnvironmentColor` | List |  | Background color of the environment badge. Text color is auto-calculated for contrast. |
| Environment name | `EnvironmentText` | Text |  | Label shown as a badge in the toolbar (e.g. Production, Lab). Leave empty to hide. |
| Hide experimental platform notice | `HideExperimentalPlatformNotice` | Boolean | `false` | Stop showing the startup notice that macOS and Linux support is experimental. Clear this to see it again. |
| Hide No-access items | `HideNoAccess` | Boolean | `false` | Remove items from the menu if object permissions is missing. Default is to mark them with red |
| Log file | `LogFile` | File |  |  |
| Max log file size | `LogFileSize` | Int | `1024` |  |
| Proxy URI | `ProxyURI` |  |  | Specify the URI for the proxy eg http://&lt;server&gt;:&lt;port&gt; |
| Show tenant name | `MenuShowOrganizationName` | Boolean | `true` | Adds the organization name next to the login info on the menu bar |
| Theme | `AppTheme` | List | `Default` | Application color theme. Default follows the Windows app theme. Values: `Default (follow Windows)`, `Light`, `Dark`. |
| Use Intune role permissions for access marking | `UseRbacAccessMarking` | Boolean | `true` | Also ask Intune which resource actions the signed-in user's role allows, and mark menu items the user cannot change (orange) or read (red). Off = mark from the app's token scopes only. Refresh the token from the Profile popup after a role change. |

## Import/Export

| Setting | Key | Type | Default | Description |
|---|---|---|---|---|
| Add company name | `AddCompanyName` | Boolean | `true` | Default setting for adding company name to the export folder |
| Add ID to export file | `AddIDToExportFile` | Boolean | `true` | This will add object ID to the export file to support objects with the same name e.g. ObjectName_ObjectId.json |
| Add object type | `AddObjectType` | Boolean | `true` | Default setting for adding object type to the export folder |
| Clear All Objects From Cache | `ClearAllObjectsFromCache` | Boolean | `false` | Whether to clear all cached objects (default is to keep assignments and scopes cached). Use this option if you experience stale cache issues. |
| Clear Cache Before Export/Import | `ClearCacheBeforeExportImport` | List | `1` | Automatically clear object cache before export/import operations (useful when re-downloading profiles to get latest data) Values: `Bulk export only`, `Bulk and manual export`, `Bulk export and import`, `All operations`. |
| Combine requests into batch calls | `UseBatchAPI` | Boolean | `true` | Combine Graph requests into $batch calls, up to 20 per call, on every path that can batch: listing, policy bodies, sub-resources, assignments, import and delete. Turn off to send every request on its own - slower, but each call shows individually in the Graph log. See Docs/GraphBatching.md. |
| Convert synced groups | `ConvertSyncedGroupOnImport` | Boolean | `true` | When a group referenced by an imported policy was AD-synced in the source tenant and does not exist in the target, recreate it as a cloud Entra group. When off, the group is skipped and the reference is left untranslated. |
| Create groups and filters | `CreateGroupOnImport` | Boolean | `true` | Create Entra groups and assignment filters referenced by an imported policy when they do not exist in the target tenant. Groups are created from the export's Groups sidecar (dynamic groups keep their membership rule) or as a default cloud security group; filters from the AssignmentFilters sidecar (platform + rule preserved). |
| Default Conditional Access Policy State | `ConditionalAccessState` | List | `disabled` | Define the state imported Conditional Access policies get. It is recommended to keep this Off (disabled) to avoid accidental tenant lock out. Values: `As Exported - Change On to Report-only`, `As Exported`, `Report-only`, `Off`. |
| Export Assignments | `ExportAssignments` | Boolean | `true` | Default setting for exporting assignments |
| Export file encoding | `ExportFileEncoding` | List | `utf8` | Character encoding for exported JSON files. UTF-8 is recommended - it is what Intune, git and other tools expect. Change it only if an existing pipeline depends on another encoding. Values: `UTF-8`, `UTF-8 with BOM`, `Unicode (UTF-16 LE)`. |
| Export Json format | `ExportJsonFormat` | List | `indented` | Layout of exported JSON. Indented is readable but PowerShell 5.1 and 7 indent differently, so files exported from different hosts differ even when the data is identical. Compact is identical on both - use it with 'Sort Json Properties' if the export is stored in git or compared by automation. Values: `Indented (readable)`, `Compact (single line)`. |
| Export nested group levels | `ExportNestedGroupLevels` | String | `1` | Depth of group-membership recursion when exporting groups. 1 (default) only exports the group assigned to the policy. 2 also exports groups that are direct members of the assigned group. 3 goes one level deeper, and so on. Higher values mean more Graph calls and may hit throttling on large group trees. |
| Graph request timeout (seconds) | `MSGraphRequestTimeoutSec` | Int | `100` | Maximum time in seconds to wait for a single Graph request before it is aborted. Bounds how long a stalled request can block the UI. Default 100. |
| Import Assignments | `ImportAssignments` | Boolean | `true` | Default value for Import assignments when importing objects |
| Import match normalized name | `ImportMatchEnableNormalizedName` | Boolean | `true` | Allow update and skip-if-exists import matching after removing organization-specific prefixes or variables from names. |
| Import match organization tokens | `ImportMatchOrganizationTokens` | String |  | Extra organization-specific words or prefixes to remove before normalized-name matching. Separate values with comma, semicolon, or new lines. |
| Import match policy reference | `ImportMatchEnablePolicyToken` | Boolean | `true` | Allow update and skip-if-exists import matching by policy reference token in the object name, e.g. [SEC-1053]. Exact name and same-tenant ID matching are always available. |
| Import match policy reference regex | `ImportMatchPolicyReferenceRegex` | String | `(?i)(?:\[(?<ref>[A-Z][A-Z0-9]{1,15}-\d{2,10})\])|(?<ref>\b[A-Z][A-Z0-9]{1,15}-\d{2,10}\b)` | Regex used to find policy reference tokens in names. It should include a named capture group called ref. Default matches values like [SEC-1053] or SEC-1053. |
| Import Scope (Tags) | `ImportScopeTags` | Boolean | `true` | Default value for Import Scope (Tags) when importing objects |
| Import type | `ImportType` | List | `alwaysImport` | How files are imported. Always import: no detection of existing objects. Skip if object exists: skip when a matching object is found. Replace: import the file, copy the existing object's assignments to it, then delete the existing object. Replace with assignments: same but assignments come from the import file. Update: settings on the existing object are replaced from the file. Values: `Always import`, `Skip if object exists`, `Replace`, `Replace with assignments`, `Update`. |
| Multi Admin Approval justification | `MultiAdminApprovalJustification` | String |  | Reason sent with every create/update/delete when the tenant requires Multi Admin Approval. Leave empty unless Tenant administration > Multi Admin Approval has an access policy covering what you are changing. With a justification set, protected changes are queued for a second administrator to approve instead of failing. |
| Pace one-per-second Graph endpoints | `GraphPaceIdentityEndpoints` | Boolean | `true` | Graph allows one request per second per tenant, across all applications, on Conditional Access policies, named locations, authentication strengths and identity protection - and sends no Retry-After when it throttles them. When enabled, requests to those endpoints are sent one at a time, one second apart, and are kept out of parallel batch dispatch. Turn off only if Microsoft has raised the limit for your tenant. |
| Parallel batch throttle limit | `ParallelBatchThrottle` | String | `4` | Maximum number of concurrent $batch POST requests when 'Parallel batch dispatch' is enabled. Higher values are faster but more likely to trigger HTTP 429 throttling. Recommended range: 2-8. |
| Portable file names | `PortableFileNames` | Boolean | `false` | Remove characters that are invalid on ANY supported platform from exported file names, not just the ones invalid on the current one. Windows rejects " < > \| : * ? \ / while Linux and macOS reject only /, so an export made on Linux can contain file names Windows cannot open. Enable this when exports are shared between platforms. Changes the file names of new exports. |
| Replace organization values in export files | `ExportReplaceOrganizationValues` | Boolean | `true` | Replace tenant-specific values with placeholders in exported JSON. By default the tenant id becomes %OrganizationId%; the organization name is left as written unless you opt in. Turn this off to export the raw values - readable and diff-friendly, but the file is then tied to the tenant it came from. Importing a file that already contains placeholders always resolves them, whatever this is set to. Which values are replaced can be changed, see Docs/ExportImportAndCopy.md. |
| Resolve reference info | `ResolveReferenceInfo` | Boolean | `true` | This will export/import info for referenced/navigation properties eg certificates in VPN profiles etc. |
| Root folder | `RootFolder` | Folder |  | Root folder for exporting/importing objects |
| Send batch calls in parallel (experimental) | `UseParallelBatchAPI` | Boolean | `false` | Send $batch calls concurrently instead of one at a time. Controls concurrency only - whether requests are batched at all is 'Combine requests into batch calls'. Requires PowerShell 7+. Speeds up large queries but raises the chance of HTTP 429 throttling; queues of 20 requests or fewer still go out one at a time. Leave off unless you've tested it in your tenant. |
| Sort Json Properties | `SortJsonProperties` | Boolean | `false` | Sort JSON properties alphabetically when exporting to improve file readability and consistency |

## Intune

| Setting | Key | Type | Default | Description |
|---|---|---|---|---|
| App download folder | `IntuneAppDownloadFolder` | Folder |  | Folder where app packages will be downloaded and where encryption files will be saved |
| App packages folder | `IntuneAppPackagesFolder` | Folder |  | Root folder where intune app packages are located |
| Get all pages | `GetAllPages` | Boolean | `true` | Get all pages when getting items in the UI. Note: This can take long time in environments with lots of policies and apps. |
| Graph Page Size | `GraphPageSize` | List | `0` | How many items load at a time Values: `Graph Default`, `5`, `20`, `50`, `100`, `1000`, `All`. |
| Max characters per cell | `ObjectListMaxCellLength` | Int | `50` | With single-line values on, cut a cell's text at this many characters so one long description cannot push the other columns out of view. Hover the cell for the full text; sorting and the filter still use the whole value. 0 = no limit. |
| Menu Object Type | `ObjectViewType` | List | `Group` | Specify object type for the menu. Group: Groups items together like the portal. Type - Single item based on API. Some APIs are split into multiple menu items. Values: `Group`, `Type (API)`. |
| Save Encryption File | `IntuneSaveEncryptionFile` | Boolean |  | Save encryption file when uploading an app. This can then be used to when downloading the app file. |
| Single-line values in object list | `ObjectListFirstLineOnly` | Boolean | `true` | Show only the first line of multi-line values (e.g. store app descriptions) in the object list, so every row is one line high. Hover a cell for the full text. Takes effect when the list is next loaded. |

## Intune Tools

| Setting | Key | Type | Default | Description |
|---|---|---|---|---|
| Format OMA-URI Settings | `FormatOMAURI` | Boolean | `false` | Automatically clean up XML formatting in OMA-URI and ADMX registry policies for consistent output |

## MS Graph General

| Setting | Key | Type | Default | Description |
|---|---|---|---|---|
| Expand assignments | `ExpandAssignments` | Boolean | `true` | Expand assignments when listing objects. This can be used in custom columns based on assignment info |
| Refresh Objects after copy | `RefreshObjectsAfterCopy` | Boolean | `true` | Reload the object list after copying an object |
| Show Bulk Delete | `AllowBulkDelete` | Boolean | `true` | Allow using bulk delete to delete all objects of selected types |
| Show Delete button | `AllowDelete` | Boolean | `false` | Allow deleting individual objectes |
| Use Graph 1.0 (Not Recommended) | `UseGraphV1` | Boolean | `false` | This will use production verionof graph, v1.0. Note: Thot officially supported since this can have unpredicted results. Some parts will require Beta version of Graph. |

## MSAL

| Setting | Key | Type | Default | Description |
|---|---|---|---|---|
| Enable Continuous Access Evaluation (CAE) | `EnableCAE` | Boolean | `true` | Adds the 'cp1' client capability so Entra ID can revoke this session in near-real-time when admin policies change. Takes effect after an app restart / fresh login. |
| Get Tenant List | `GetTenantList` | Boolean | `false` | Get a list of all tenants the current user has access to. Only used when the user has access to multiple tenants. This may cause duplicate login/consent prompts first time |
| Log MSAL correlation IDs | `MSALCorrelationLogging` | Boolean | `false` | Generate and log a correlation Guid per AcquireToken call. Useful when working with Microsoft support; off by default to keep logs quiet. |
| MSAL log level | `MSALLogLevel` | List | `Info` | Minimum severity of MSAL log messages to capture. Verbose is noisy. Values: `Error`, `Warning`, `Info`, `Verbose`. |
| MSAL logging | `MSALEnableLogging` | Boolean | `true` | Write MSAL's internal log messages to msal.log under %LOCALAPPDATA%\IntuneManagement. Useful for diagnosing authentication failures. |
| MSAL logging: include personal data (PII) | `MSALEnablePiiLogging` | Boolean | `false` | Include MSAL's PII log messages in msal.log and unredact broker/WAM error text (logged as '(pii)' otherwise). The log will then contain user names, tenant and object ids and token details - review it before sharing. Requires app restart. |
| MSAL region (optional) | `MSALRegion` | String |  | Force a regional ESTS endpoint (e.g. 'westeurope'). Leave empty for automatic. Applied via MSAL_FORCE_REGION; requires app restart. |
| Remember Login | `CacheMSALToken` | Boolean | `true` | Store the MSAL token in an encrypted file and automatically log on when the script starts. The token is stored in the users profile and can only be decrypted by the user that created it. Note: Requires restart |
| Sort Account List | `SortAccountList` | Boolean | `false` | Sort the list of cached accounts based on user name. Updated at restart or account change |
| Sort Tenant List | `SortTenantList` | Boolean | `false` | Sort the list of available tenants based on Tenant name. Updated at restart or account change |
| Use MsalCacheHelper library | `MSALUseCacheHelperLib` | Boolean | `true` | Use Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper for the token cache. Provides cross-process file locking and is the supported library. Disable to fall back to the legacy TokenCacheHelperEx (DPAPI, process-local lock). |
| Use system browser for login | `UseSystemBrowser` | Boolean | `true` | Sign in using the default web browser instead of the embedded view or WAM. Enables passkey/FIDO2 login and browser extensions. Takes precedence over WAM. Custom app registrations need the http://localhost redirect URI (Mobile and desktop applications). Turn off to use the embedded window. Requires app restart. |
| Use Web Account Manager (WAM) for login | `UseWAM` | Boolean | `false` | Use the Windows Web Account Manager broker (Windows Hello, device compliance claims). Turn this on only when the account you sign in with is your Windows account or one added under Settings > Accounts - for any other account WAM cannot keep the session alive and you are prompted again about every hour. Requires PowerShell 7 and an app restart. |
| WAM: list OS accounts | `WAMListOSAccounts` | Boolean | `true` | When WAM is enabled, also surface machine-joined Entra accounts in the account picker (PS7+). Off-by-default on PS5. |

## OAuth

| Setting | Key | Type | Default | Description |
|---|---|---|---|---|
| OAuth browser prompt | `OAuthPrompt` | List | `select_account` | OAuth /authorize prompt behaviour. 'Force login' re-authenticates even with an active browser session (equivalent to force-interactive); 'None' fails if interaction would be required. Values: `Select account`, `Force login`, `Consent`, `None (silent)`. |
| OAuth login hint (UPN) | `OAuthLoginHint` | String |  | Optional UPN to pre-fill on the sign-in page (login_hint). |
| OAuth redirect port | `OAuthRedirectPort` | Int | `0` | Fixed loopback port for the browser redirect (http://localhost:<port>). 0 = pick a free port automatically. Set a fixed port only if your app registration requires a specific http://localhost:<port> redirect. |
| Remember login (cache token) | `OAuthCacheToken` | Boolean | `false` | Persist the OAuth refresh token (DPAPI-encrypted, current user) so the app silently resumes the browser session after a restart. When off, you sign in again after each restart (usually a quick browser redirect via existing SSO). |
