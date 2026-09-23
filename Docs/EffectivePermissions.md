# Effective Permissions

What the signed-in identity can actually do per policy type, and how the left-nav
access marking, the Profile popup's **Permissions** dialog and
`Get-IMGraphEffectivePermissions` derive it. This page is the "how it works now".

## The two halves of a delegated login

| Half | Where it lives | What it says |
| --- | --- | --- |
| App consent | token `scp` (delegated) / `roles` (app-only) | which Graph scopes the *application* was granted |
| User authorization | Intune RBAC role assignments (+ scope tags) and Entra directory roles | what the *user* may do |

The token carries only the first half (plus `wids`, the directory role template
ids). A user holding just the built-in *Read Only Operator* Intune role signs in with
a token that says `DeviceManagementConfiguration.ReadWrite.All`, and every write is
refused with 403. Effective access is the intersection, so the marking has two layers:

```
Update-IntuneAccessLevels                       Internal/AccessLevel.ps1
  Layer 1  scp/roles vs _Permissions          -> Full / Limited / None
  Layer 2  Get-IntuneRbacContext              -> Full / Limited / None / Unknown   Internal/EffectivePermissions.ps1
  stamp    AccessType = worst(L1, L2), AccessInfo = both reasons
```

Layer 2 can only make a type worse. Unknown never colours: Layer 1 stands.

## Layer 2 in detail

1. **Applies only to delegated tokens.** `idtyp = app` (or no `scp`) means application
   permissions, which bypass Intune RBAC - the `roles` claim is the whole answer.
2. **Directory-role short-circuit.** Intune Administrator
   (`3a2c62db-5318-420d-8d74-23affee5d9d5`) or Global Administrator
   (`62e90394-69f5-4237-9190-012177145e10`) in `wids` grants complete Intune RBAC, so
   every Intune category is Full and Graph is not asked. Global Reader and the partial
   roles are *not* shortcuts - the user may also hold an Intune role.
3. **Otherwise two Graph calls:**
   - `GET /beta/deviceManagement/getEffectivePermissions(scope='*')` - the same function
     the Intune portal uses to enable its buttons. It answers with the **allowed**
     actions only: `notAllowedResourceActions` comes back empty on every response, so
     on its own it cannot tell "the role denies this action" from "no such action
     exists for this category".
   - `GET /beta/deviceManagement/resourceOperations` - the **catalogue**: every resource
     action Intune defines, one entry per action, its `id` being exactly the
     `Microsoft.Intune_<Category>_<Action>` name the first call uses. That is the
     existence answer the first call cannot give (260 actions over 53 resources on the
     tenant this was built against). It describes the service, not the user, so it is
     cached per tenant and survives a token refresh. When it cannot be read, existence
     is *unknown* and the verdicts degrade as shown below.
4. **Per type:** the `_API` path is mapped to an Intune resource category
   (`deviceManagement/deviceConfigurations` -> `DeviceConfigurations`, ...; table in
   `Internal/EffectivePermissions.ps1`, overridable per type with `_ResourceCategory`).
   Only actions that exist for the category are required, so a category with no
   `_Assign` (Roles) or one that spells it `Modify` (ManagedGooglePlay) is not marked
   down for the actions it never had.

   | For the category | Level |
   | --- | --- |
   | `_Read` is not in the catalogue | Unknown - no such category exists; say nothing |
   | `_Read` exists, not allowed | None |
   | `_Read` allowed, type declares only Read scopes | Full (read-only by design) |
   | `_Read` allowed, an existing write action (`Create/Update/Delete/Assign/Modify`) not allowed | Limited, tooltip lists the missing actions |
   | every existing action is allowed | Full |

   The Limited tooltip distinguishes the two shapes of it: "read-only for
   `<Category>`" when no write action at all is allowed, and "partial write access to
   `<Category>`" when some are (a role that can create and update but not delete or
   assign is still Limited, but it is not read-only). The Permissions popup's Role,
   Effective and Result columns show "Partial write" for that second shape rather
   than "Read" / "Read-only" - Effective and Result only when the token can write
   too, since a read-only token over a partial-write role really is read-only.

   With no catalogue, the same type yields Unknown when `_Read` is not allowed (it may
   not exist), Limited when no write action at all is allowed, and Full otherwise - a
   coarser answer that never invents an action name.

   APIs that Intune RBAC does not govern (`identity/*`, `identityGovernance/*`,
   `organization/*`) and the few listed in `$script:RbacUnmappedApis` are Unknown. A
   test fails if a new `deviceManagement/` or `deviceAppManagement/` type is in
   neither table.

## Caching and refresh

The user's answer is cached per tenant under the **token fingerprint** (`tid|oid|iat`)
with no TTL; the action catalogue is cached per tenant only, so a refresh re-asks
`getEffectivePermissions` and reuses the catalogue. Both are dropped on disconnect. Any newly minted token - routine renewal, a new sign-in, or **Refresh in the
Profile popup** - has a new `iat`, misses the cache and re-asks. That single user
action therefore covers both kinds of change:

| Change | Visible in the token? | Caught by |
| --- | --- | --- |
| Directory role via PIM (Intune Administrator, Global Admin) | only in a newly issued token; the routine silent acquire returns the cached one until near expiry | Refresh mints a token with the current `wids` |
| Intune RBAC assignment (a role added, PIM for Groups) | never - the token is identical | Refresh still produces a new `iat`, so Layer 2 re-asks Graph |

A failed lookup is remembered for the same fingerprint (no retry on every menu
rebuild) and retried after a refresh. Refresh is provider-agnostic: MSAL
(`WithForceRefresh`), delegated OAuth (`refresh_token` grant) and MgGraph
(`Connect-MgGraph` re-run) all mint a new token; BYO bearer tokens cannot refresh and
the button is already disabled for them.

**Scope tags are not modelled.** `getEffectivePermissions` is the global answer; a
user limited to some tags can still be refused on individual objects.

## Surfaces

- **Left nav** - orange (Limited) / red (None) with the reason in the tooltip; the
  existing `HideNoAccess` setting hides red rows. A summary line is logged at sign-in.
- **Profile popup -> Permissions** - one row per policy type: token level, Intune-role
  level, effective level, reason, plus a header saying where the Intune half came from
  and when the token was issued. WPF: `UI/WPF/Extensions/EffectivePermissionsUIWPF.ps1`;
  Avalonia: `UI/Avalonia/Extensions/EffectivePermissionsUIAvalonia.ps1` with
  `UI/Avalonia/Classes/EffectivePermissionRowItem.ps1`.
- **`Get-IMGraphEffectivePermissions`** - the same rows for scripts
  (`-PolicyType`, `-TokenId`), or `-Raw` for the context (source, the allowed action
  set, the catalogue of actions that exist, raw response). Runs even when the setting
  below is off.

## Setting

`UseRbacAccessMarking` (General, default on). Off restores the token-only marking
exactly. Registered in `Internal/EffectivePermissions.ps1` because engine code reads
it.

## Verifying against a tenant

Both endpoints, the response shapes and every **category name** in the table have been
run against a live tenant, and `Tests/EffectivePermissions.Tests.ps1` asserts the names
against the catalogue fixture, so a typo fails the suite. What is still unverified is
the other half of each entry - that a given `_API` really is governed by the category it
is mapped to. Those are marked `UNVERIFIED`, and **an unverified mapping is not a safe
mapping**: only a category name that does not exist at all degrades to Unknown. A mapping
to a name that exists but governs a *different* resource produces a confident verdict
derived from the wrong role permissions - a type marked Full because the user's role
covers the category it was mistakenly mapped to, or None because it does not. Both are
wrong, and neither shows as Unknown. The typo test cannot catch this: the wrong name is a
real one. Only a role that grants exactly one category can, so to check on a lab tenant:

```powershell
Import-Module .\IntuneManagement.psd1 -Force
Connect-IMIntuneManagement ...                      # delegated, as a NON-admin test user
Get-IMGraphEffectivePermissions -Raw | Select-Object Source, TenantId, AsOf
(Get-IMGraphEffectivePermissions -Raw).Allowed | Sort-Object       # what the user's role grants
(Get-IMGraphEffectivePermissions -Raw).Catalog | Sort-Object       # every action Intune defines
Get-IMGraphEffectivePermissions | Where-Object RbacLevel | Format-Table Id, ResourceCategory, TokenLevel, RbacLevel, EffectiveLevel
```

Assign the test user a custom role granting exactly one category, refresh the token, and
check that the types mapped to it are the ones that move. Move confirmed entries out of
`UNVERIFIED`, and add `_ResourceCategory` overrides where a type maps elsewhere.

## Tests

`Tests/EffectivePermissions.Tests.ps1` (offline, fixtures under
`Tests/Fixtures/EffectivePermissions/`): category resolution, completeness and every
category name against the catalogue, response parsing, per-type verdicts for a full
admin / read-only operator / custom role and for a missing catalogue, the never-upgrade
property, app-only skip, directory-role short-circuit, fingerprint caching, catalogue
caching across a refresh, negative caching, the setting gate, disconnect clearing,
`Update-IntuneAccessLevels` integration and the cmdlet's rows.

The fixtures are recorded responses, not hand-written: `ReadOnlyOperator.json` is the
built-in role's 55 actions, `ResourceOperations.json` the whole catalogue, and every
`notAllowedResourceActions` is empty because that is what Graph returns. A verdict test
that fails is a bug in the code or the mapping - do not "fix" it by inventing
not-allowed data the API never sends.
