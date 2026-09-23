# Settings Cache Logging

## Settings

**[Settings](Settings.md) is the reference** - the three layers, the three store modes
(registry / JSON file / in-memory), scope and precedence, the public `*-IMSetting`
cmdlets, portable settings files, and the testing helpers. In short:

```text
Add-SettingsObject
  -> setting metadata registered in a section
  -> Get-SettingValue / Get-IMSetting
      -> tenant-specific persisted value
      -> global persisted value
      -> setting DefaultValue
```

Important setting groups include:

| Area | Examples |
| --- | --- |
| Graph/auth | `ActiveAuthProvider`, app ID/secret/cert settings, cloud settings. |
| performance | `UseBatchAPI`, `UseParallelBatchAPI`, `ParallelBatchThrottle`, `GraphPaceIdentityEndpoints`. |
| import/export | `ImportType`, assignment/scope tag import defaults, matching settings. |
| UI | delete/bulk delete visibility, expand assignments, formatting. |
| cache | clear cache before import/export, clear all cache. |

## Cache

Core cache functions:

| Function | Purpose |
| --- | --- |
| `Set-CacheObject` | writes in-memory or persistent cache values. |
| `Get-CacheObject` | reads cache values with optional default. |
| `Clear-CacheObject` | removes cache values. |
| `Get-CacheStats` | diagnostics. |

Common cache keys:

| Key pattern | Meaning |
| --- | --- |
| `AADObjectCache_<tenant>` | resolved groups/users/service principals for migration/import. |
| `TenantCache_<tenant>` | persistent backing file for tenant cache values. |
| `_migFileCache` script variables | per-bulk-export migration table buffers. |
| `_appConfigTargetAppCache` | app config target app lookup cache. |

## Logging

Main functions:

| Function | Use |
| --- | --- |
| `Write-Log` | normal and warning/error messages. |
| `Write-LogDebug` | debug-only diagnostics. |
| `Write-LogError` | exception-aware logging. |
| `Write-Status` | UI status line updates for long operations. |

Log behavior is influenced by settings such as `Debug`, `LogFile`, `LogFileSize`, and `LogOutputError`.

## Graph Call Telemetry

`Invoke-MSGraphAPI` and batch helpers populate `$script:AllGraphCalls`. The Graph Calls UI uses it to show:

| Field | Meaning |
| --- | --- |
| request URL/method | what was called. |
| provider | MSAL, MgGraph, etc. |
| status/error | result and parsed Graph error. |
| duration/KB/object count/page count | performance and response shape. |
| batch sub-requests | item-level status inside `$batch`. |

## Testing Notes

Settings can leak from the developer machine into tests, and tests can write into the
developer's real store. Use `Use-TestSettingsStore` / `Restore-TestSettingsStore` from
`Tests/TestBootstrap.ps1` - and declare the hooks at `Describe` level, not inside
`InModuleScope`. See [Settings: Testing](Settings.md#testing).

