# Bulk Export

## Purpose

`Start-GraphBulkExport` is the UI-independent bulk export driver. It can be called from the WPF UI, a scheduled task, or automation. It wraps listing, hydration, extra-data synchronization, file writing, and migration-table generation.

## Main Files

| File | Responsibility |
| --- | --- |
| `Public/Start-GraphBulkExport.ps1` | bulk export orchestration and helper functions. |
| `Public/Export-GraphPolicy.ps1` | per-policy file writing and migration hooks. |
| `Public/Get-GraphPolicies.ps1` | listing and assignment loading. |
| `Internal/MSGraph.ps1` | batch execution, migration table functions, navigation properties. |
| `UI/Extensions/IntuneManagerUI.ps1` | bulk export form and settings save UI. |

## Settings Resolution

Parameter precedence:

```text
IntuneManagerExportSettings defaults
  <- SettingsFile JSON
  <- explicit parameters
```

This allows saved scheduled-task settings while still making command-line overrides possible.

`ExportFullMembershipPrefixes` ("Get full membership") was removed: every group
path is batched now, and its only output was `#DirectMembers` /
`#DirectMemberCount` on the group sidecars, which nothing in the module read.
`Clear-BulkExportLegacyMembershipSetting` (`Internal/BulkExport.ps1`) runs at the
start of every export so a value that outlived the removal is never dropped in
silence: a non-empty value in a settings file is logged as a warning each run
(the file belongs to the caller and is not rewritten), and a value in the
settings store is warned about once and then cleared. The store sweep follows the
same precedence `Get-SettingValue` used, so it covers the per-tenant paths
(`<tenantId>\IntuneManager`, for the connected organization and for the export
token's tenant) as well as the global `IntuneManager` one.

## Target Resolution

Targets are resolved from:

| Input | Behavior |
| --- | --- |
| `-PolicyType` | exact type IDs. |
| `-PolicyGroup` | group IDs expanded to member policy types. |
| neither | all exportable groups and types. |

Unknown types/groups are logged and skipped.

## Parallel Pipeline

When `UseParallelBatchAPI` is enabled and PowerShell 7+ is running, bulk export uses a three-phase pipeline:

```text
Phase 1: list all selected types through Get-GraphPolicies
Phase 2: batch-fetch full policy bodies via Invoke-PolicyHydrate
Phase 2.5: sync extra data and prefetch assignment groups
Phase 3: write policy files type by type
```

When parallel mode is disabled, it processes each type sequentially:

```text
list type -> hydrate full objects -> sync extra data -> write type
```

## Extra Data Synchronization

Some Graph list/detail endpoints do not include all export data. `Invoke-PolicyExtraData` (a Phase-A wrapper that will shrink to nothing as helpers migrate to the per-class `_HasSubResourceBatch` contract) fills the remaining gaps before file writing.

| Helper | Data added | Status |
| --- | --- | --- |
| `Sync-BulkExportReusableSettings` | reusable setting instances. | Phase-A wrapper; migrates to class contract in Phase B. |
| `Sync-BulkExportBrandingImages` | branding image payloads. | Phase-A wrapper; migrates to `IntuneBrandingObject` contract. |
| `Sync-BulkExportAppConfigurationTargetApps` | targeted mobile app reference info. | Phase-A wrapper; migrates to AppConfig object contracts. |
| `Sync-BulkExportRoleAssignmentDetails` | role assignment expanded details. | Phase-A wrapper; migrates to `RoleDefinitionObject` Phase 2. |
| `Sync-BulkExportTermsOfUseFiles` | terms of use file data. | Phase-A wrapper; migrates to `TermsOfUseObject` contract. |
| `Sync-BulkExportMigrationGroups` | migration-table group resolution. | Cross-cutting; will be renamed `Invoke-MigrationGroupResolution`. |
| `Sync-BulkExportNestedGroupHierarchy` | nested-group hierarchy expansion. | Cross-cutting; will be renamed `Invoke-NestedGroupResolution`. |

ADMX definition values, presentation values, app dependencies, supersedence, and Win32 scripts are no longer handled by `Sync-Bulk*` helpers — those moved to the per-class `GetSubResourceBatchRequests` / `ApplySubResourceBatchResult` contract on `AdminTemplateObject` and `ApplicationObject`.

## Migration Performance

Bulk export resets migration caches at the start:

| Cache | Purpose |
| --- | --- |
| `_migFileCache` | in-memory `MigrationTable.json` objects. |
| `_migFileObjectsIndex` | duplicate prevention. |
| `_migFileDirty` | list of migration files to flush once. |
| `_migFilePathCache` | avoid repeated path resolution. |
| `_appConfigTargetAppCache` | avoid repeated app target lookups. |

`Sync-BulkExportMigrationGroups` prefetches assigned groups in one batch. `Add-GraphMigrationObject` can still fetch non-prefetched references such as Conditional Access users/groups and nested groups.

## Output

Returns a summary object:

| Field | Meaning |
| --- | --- |
| `Types` | number of policy types processed. |
| `Policies` | number of policies exported. |
| `Failed` | number of failed type operations. |
| `Duration` | elapsed time. |

## Extension Points

To make a policy type export fully:

1. ensure the list endpoint returns enough data or full hydration works;
2. add full-object URL handling when a type uses polymorphic endpoints;
3. add extra-data sync if the data is not part of the normal object body;
4. add `PostExportCommand` for references that must be included in migration data;
5. set properties-to-remove for import/update round trips.

