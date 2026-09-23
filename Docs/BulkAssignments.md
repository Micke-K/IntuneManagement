# Bulk Assignments

## Purpose

`Set-GraphBulkAssignments` adds, replaces, or removes assignment targets across many policies. It is public and UI-independent; the WPF bulk assignment form builds an `IntuneManagerAssignmentSettings` object and calls it.

## Main Files

| File | Responsibility |
| --- | --- |
| `Public/Set-GraphBulkAssignments.ps1` | assignment orchestration and body generation. |
| `Classes/IntuneBaseClasses.ps1` | `IntuneManagerAssignmentSettings` and policy type assignment metadata. |
| `Public/Get-GraphPolicies.ps1` | list policies and include current assignments. |
| `UI/Extensions/IntuneManagerUI.ps1` | bulk assignment UI, group picker, app settings dialog. |
| `UI/XAML/BulkAssignments*.xaml` | forms for assignment targets and settings. |

## Supported Actions

| Action | Behavior |
| --- | --- |
| `Add` | union current assignments with selected assignment rows. |
| `Replace` | replace all current assignments with selected rows. |
| `Remove` | remove selected rows from current assignments. |

Graph `/assign` endpoints replace the full assignment collection, so `Add` and `Remove` must first load current assignments and compute the final collection.

## Target Types

Supported target descriptors:

| Target type | Required fields |
| --- | --- |
| `groupAssignmentTarget` | `GroupId` |
| `exclusionGroupAssignmentTarget` | `GroupId` |
| `allDevicesAssignmentTarget` | none |
| `allLicensedUsersAssignmentTarget` | none |

Group targets can include assignment filters:

| Field | Meaning |
| --- | --- |
| `FilterId` | assignment filter ID. |
| `FilterType` | `include` or `exclude`. |

## Assignment Shapes

Different Graph APIs use different assignment body shapes.

| Shape | Detection | Body |
| --- | --- | --- |
| simple | default | `{ target }` |
| app | `AssignmentsType = mobileAppAssignments` | `{ target, intent, settings? }` |
| script | `AssignmentsType = deviceHealthScriptAssignments` | `{ target, runRemediationScript?, runSchedule? }` |

The function rejects policy types that cannot be represented by one of these shapes.

## Policy Type Gating

`Test-BulkAssignmentSupported` checks:

1. `SupportsAssignments = true`;
2. `AssignmentsType` exists;
3. the type is not a known non-policy assignment API;
4. `Get-BulkAssignmentShape` returns a supported shape.

`Get-BulkAssignmentObjectType` determines assignment entry `@odata.type`. It first checks `PolicyType.AssignmentObjectType`, then uses built-in heuristics for known types.

## No-Op Detection

The command builds tuple signatures for assignment comparison.

| Signature | Used for | Fields |
| --- | --- | --- |
| target signature | Add/Remove identity and dedupe | target type, group ID, filter ID/type, app intent |
| full signature | no-op detection | target signature plus app settings, health-script schedule, remediation flag |

This distinction matters because `Add` should not create duplicates, but `Replace` must still update settings for an existing target.

## App Settings

The UI stores app settings by settings type name, for example `win32LobAppAssignmentSettings`. At execution time:

1. the policy object's `@odata.type` is mapped to the assignment settings type;
2. the selected row's matching settings hashtable is converted to Graph-shaped objects;
3. nested hashtables and arrays are converted recursively for JSON serialization;
4. settings are omitted when not configured so Graph defaults apply.

## Health Script Schedule

Health script settings are stored under the synthetic key `deviceHealthScriptAssignment`. The command turns those values into:

| UI value | Graph type |
| --- | --- |
| `Hourly` | `deviceHealthScriptHourlySchedule` |
| `Daily` | `deviceHealthScriptDailySchedule` |
| `Once` | `deviceHealthScriptRunOnceSchedule` |

## Summary Output

The command returns:

| Field | Meaning |
| --- | --- |
| `Types` | eligible policy types. |
| `PoliciesScanned` | policies loaded. |
| `PoliciesMatched` | policies matching name filter. |
| `PoliciesUpdated` | successful assignment POSTs. |
| `PoliciesSkipped` | no changes required. |
| `PoliciesFailed` | failed POSTs or missing batch responses. |
| `UnsupportedTypes` | selected types skipped by support gate. |
| `Duration` | elapsed time. |

