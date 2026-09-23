# Graph batching and parallelism

Two settings control how the module talks to Microsoft Graph, and since the
2026-09-07 rework each one means exactly what its name says:

| setting key | title in Settings | meaning |
| --- | --- | --- |
| `UseBatchAPI` | Combine requests into batch calls | Combine logical requests into `POST /$batch`, twenty per call, on **every** path that can batch. Off means `$batch` is never used: each request is its own HTTP call. Default on. |
| `UseParallelBatchAPI` | Send batch calls in parallel (experimental) | Send those `$batch` POSTs concurrently instead of one at a time. Concurrency and nothing else. PowerShell 7 only. Default off. |

Working definitions used on this page:

- **Batched** - several logical requests combined into one `POST /$batch`.
- **Parallel** - more than one HTTP request in flight at once. In this module
  that is always a set of `$batch` POSTs; there is no parallel direct-call path.

Both keys are read in exactly one place each, the two predicates at the top of
`Internal/MSGraph.ps1`:

```powershell
Test-GraphBatchEnabled      # (Get-SettingValue "UseBatchAPI") -eq $true
Test-GraphParallelEnabled   # (Get-SettingValue "UseParallelBatchAPI") -eq $true -and PS 7+
```

`Tests/Static.Tests.ps1` fails the build if either key is consulted anywhere
else. That rule is what keeps the two concerns from re-tangling: before it,
`UseBatchAPI` was checked at four call sites and ignored by the dispatcher
itself, and `UseParallelBatchAPI` silently selected an entire export pipeline.

## Where the decisions are made

`Invoke-GraphBatchRequest` (`Internal/MSGraph.ps1`) is the only place that
decides between batched, direct, serial and parallel dispatch. Every caller
hands it a queue of sub-requests and gets back the batch response shape
(`id`, `status`, `headers`, `body`) whichever way the requests went out.

1. Paced sub-requests (see below) are split into their own queue first.
2. **Batching off, or exactly one non-paced request**: every sub-request goes
   out as its own `Invoke-MSGraphAPI` call through
   `Invoke-GraphBatchRequestDirect`. The sub-request's `Accept` header becomes
   `-ODataMetadata`, other headers ride on `-AdditionalHeaders`, `-AllPages`
   is forwarded for GETs, and the status comes from the telemetry row the
   wrapper records for every call. A lone request takes this path even with
   batching on, because a `$batch` envelope around one GET is a whole extra
   round-trip for nothing.
3. **Batching on**: the paced queue drains first, one sub-request per POST,
   each gated on the tenant clock. Then the normal queue goes twenty per POST.
4. **Parallel on** and the normal queue has more than
   `$script:GraphParallelBatchMinQueue` (20) sub-requests: the chunks are
   dispatched concurrently through `Invoke-ParallelGraphBatchPosts`, up to
   `ParallelBatchThrottle` at a time. At or below the threshold the queue goes
   out serially and the log says so once:
   `Batch <type>: parallel requested but the queue is N (threshold 20) - dispatching serially`.

## Every Graph call path

| path | entry point | batches? | can be parallel? | controlled by |
| --- | --- | --- | --- | --- |
| Policy listing | `Get-GraphPolicies` | yes, list requests coalesced by URL; types without a batch object list directly | yes | both settings |
| Policy bodies (hydrate) | `Invoke-PolicyHydrateBodyBatch` | yes, chunks of `ParallelBatchThrottle` x 20; a single policy is a direct GET | yes | both |
| Sub-resources (branding images, app relationships, ToU files, ...) | `Invoke-PolicySubresourceFetch` | yes | yes | both |
| Assignments | `Add-GraphPolicyAssignments` | yes | yes | both |
| Bulk export group prefetch | `Sync-BulkExportMigrationGroups` | no `$batch`: one `directoryObjects/getByIds` POST per thousand ids | no | always runs; not a batching decision |
| Nested group hierarchy | `Sync-BulkExportNestedGroupHierarchy` | yes | yes | both |
| Migration objects queued as cache misses | `Resolve-GraphMigrationObjectsPending` | getByIds for groups/users/devices/service principals; `$batch` for assignment filters and anything else | for the `$batch` part | `UseBatchAPI` for the non-directory part |
| Documentation prefetch | `Initialize-DocumentationRunPrefetch` | scope tags, filters and category lists in one `$batch`; assignment groups through one getByIds preload | yes | both for the `$batch`; the preload always runs |
| Import | `Import-GraphPolicy` | bulk: POSTs queued then batched; direct: each object's own `ImportObject` | yes (bulk) | `UseBatchAPI` chooses bulk versus direct - they are different code, the direct path runs class overrides and carries `PreImportCommand` headers |
| Delete | `Remove-GraphPolicy` | same shape as import | yes (bulk) | `UseBatchAPI` chooses |
| Everything single-request | `Invoke-MSGraphAPI` | no | no | not a batching path |

### The two overrides

**Paced endpoints.** Conditional Access policies, named locations,
authentication strengths, authentication context references, risk detections
and risky users are limited by Graph to one request per second per tenant,
counted across every application, with no `Retry-After` on throttle. Requests
to them are always one sub-request per POST, never parallel, each gated on a
per-tenant clock, whatever the two settings say. The rules live in
`Internal/GraphRateLimits.ps1` and the `GraphPaceIdentityEndpoints` setting
switches only the pacing off. Pacing legitimately wins over both settings.

**Single requests.** A queue of one non-paced request goes direct regardless
of `UseBatchAPI`. See step 2 above.

## Bulk export specifically

`Start-GraphBulkExport` runs one pipeline whatever the settings say: list every
type in one call, hydrate every body across types, prefetch every referenced
group, walk the nested hierarchy when nested export is on, then write per
type. Concurrency takes effect only inside the dispatcher. The old per-type
interleaved loop that ran with parallel off is gone; it skipped the group
prefetch, which on a 963-object tenant cost 22 of a measured 26m 35s (one
direct GET per assignment group at ~1.6 s, against ~2 s for twenty of them in
one `$batch`).

`Add-GraphMigrationObject` never calls Graph. A cache miss is queued and
resolved in bulk when the migration tables are flushed, so a single-policy
export and a bulk export both pay one getByIds POST for all their groups, not
one GET each.

## Benchmark protocol

Export to a fresh scratch folder, never over a previous run, then compare
against a reference export:

| configuration | expectation |
| --- | --- |
| batch on, parallel off (the default) | close to the parallel time, not 6x slower |
| batch on, parallel on | fastest; no regression |
| batch off, parallel off | slowest; zero `Invoke batch` log lines; identical output |

Output equivalence means the same file list, content differences limited to
`lastModifiedDateTime` on objects genuinely edited between runs, and the same
`(Id, Type)` set in `MigrationTable.json` (its ordering may differ).

## Reading the log

| log line | meaning |
| --- | --- |
| `Invoke batch N <type> (M requests)` | one `$batch` POST, serial path |
| `<type>: dispatching N request(s) in parallel batches` | the parallel dispatcher took the queue |
| `Batch <type>: parallel requested but the queue is N (threshold 20) - dispatching serially` | parallel is on but the queue was too small |
| `Direct dispatch <type> (N requests, batching off)` | `UseBatchAPI` is off; each request is its own call |
| `Bulk export: pre-fetching N AAD group(s) ...` / `... prefetch complete - R resolved, M missing/deleted` | the getByIds group prefetch |
| `Migration objects: resolving N queued <kind> id(s) for tenant ...` | the deferred migration-object queue draining |
| `<type>: N of M requests - Conditional Access allows one per second` | the paced queue |

Design record: `Docs/superpowers/specs/2026-09-07-graph-batching-parallelism-control-design.md`.
Tests: `Tests/GraphBatching.Tests.ps1`, `Tests/GraphRateLimits.Tests.ps1`.
