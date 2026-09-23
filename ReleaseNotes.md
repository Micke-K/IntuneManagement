# Release Notes

## 4.0.0-beta1 - 2026-09-23

**This is a beta release.** 4.0 is a full rewrite of IntuneManagement. The feature set is
broadly familiar, but nearly every internal has changed. Test it against a lab tenant before
pointing it at production.

**Version 3 remains the supported release.** It stays on the default branch and keeps getting
fixes until 4.0 leaves beta. 4.0 lives on the `v4` branch until then, and the in-application
update check will not offer it to a 3.x installation while it is a pre-release. Run 4.0
alongside 3.x if you like: they keep their settings in different places and neither reads the
other's.

<br />

**BREAKING CHANGES**

These are the things that stop working if 4.0 simply replaces 3.x in an existing setup.

-  **Every pipeline and scheduled job has to be rewritten.** The 3.x silent mode -
   `Start-IntuneManagement.ps1 -Silent -SilentBatchFile <file>` driving a `BulkExport.json`
   or `BulkImport.json` exported from the forms, with `%Date%`, `%DateTime%` and
   `%Organization%` in paths - is gone, and 4.0 does not read those files. The launcher
   switches (`-TenantId`, `-AppId`, `-Secret`, `-Certificate`, `-JSonSettings`) are gone
   with it. The replacement is the public commands: `Connect-IMIntuneManagement` for the
   credentials, `Start-IMGraphBulkExport` / `-Import` / `-Documentation` for the work, and
   `Import-IMSettingsStore` for a settings file kept in source control. A 3.x nightly export
   becomes:

   ```powershell
   Import-Module .\IntuneManagement.psd1
   Connect-IMIntuneManagement -TenantId $tenant -AppId $app -Secret $secret
   Start-IMGraphBulkExport -ExportFolder "C:\Export\$(Get-Date -Format yyyy-MM-dd)" -ExportAssignments $true
   ```

   See [Docs/Examples.md](Docs/Examples.md) for the other operations.
-  **Settings do not carry over.** 3.x keeps its settings under
   `HKCU\Software\CloudAPIPowerShellManagement`; 4.0 uses `HKCU\Software\IntuneManagement`,
   and it does not read the 3.x JSON settings file either. The first start of 4.0 has a
   blank configuration: custom app id, export folders, documentation options and every
   other setting are entered again. That is also why the two versions can run side by side
   without touching each other.
-  **The cached sign-in does not carry over.** The token cache lives in a per-version data
   folder (`%LOCALAPPDATA%\IntuneManagement` in 4.0), so the first start of 4.0 is a fresh
   sign-in, on every machine.
-  **Two export folders changed case.** `AutoPilot` is now `Autopilot` and
   `HardwareConfigurations` is now `hardwareConfigurations`. Import matches folder names
   case-insensitively, so a 3.x export still imports; but a script that builds a path to
   one of those folders, or a Linux file system, sees two different names. Every other
   export folder keeps its 3.x name, and the JSON inside is compatible in both directions -
   4.0 exports can carry properties 3.x did not write, which 3.x ignores on import.
-  **Two policy types are not in 4.0**, both already disabled in the 3.x code:
   -  Intune *Locations* (`managementConditions`). Microsoft removed the object from
      Intune; it only ever served Android device administrator compliance policies.
   -  The iOS DEP enrollment profile, which 3.x carried commented out and never listed.
   Neither has been exportable from a current 3.x, so no export folder is orphaned.
-  **Scripts that imported the 3.x extension files will not run.** 4.0 is a single module,
   `IntuneManagement.psd1`, with a documented public command surface; the `Extensions\*.psm1`
   files are gone. Only commands listed in `FunctionsToExport` are supported, and they are
   prefixed `IM`. Anything else is internal and may change without notice.

Not breaking, but different: interactive sign-in opens the system browser by default. See
**Browser sign-in on every platform** below for the reason, the one requirement it puts on a
custom app registration, and how to turn it off.

<br />

**New features**

-  **Runs on Windows, macOS and Linux.**
   A second UI backend built on Avalonia runs the full application outside Windows. The
   original WPF backend remains the default on Windows and is unchanged in behaviour.
   Select a backend with the `IM_UI_BACKEND` environment variable (`WPF`, `Avalonia`, or
   `None` for headless automation). PowerShell 7.4 or later is required for the
   cross-platform UI.
   See [Docs/CrossPlatform.md](Docs/CrossPlatform.md).

-  **A supported automation API.**
   Every bulk operation available in the UI is also a public command that runs headless with
   no UI loaded: bulk export, import, copy, delete, assignments, scope tags and
   documentation. This is the intended path for DevOps pipelines and scheduled runs.

-  **Two authentication providers.**
   Selected per connection with `-Provider`:
   -  **MSAL** (default) - interactive browser sign-in, the WAM broker on Windows (off by
      default; turn it on only when the app account is your Windows account), client
      secret, certificate, and device code.
   -  **OAuth** - a pure PowerShell implementation with no SDK and no DLL dependency.
      Supports client secret, certificate, managed identity, workload identity federation
      (AKS, GitHub Actions OIDC), device code and ROPC.

   Multiple tenants can be signed in at once; each Graph call resolves its own token.

   A token acquired elsewhere can also be supplied directly with `-Token` (a BYO token).
   That is the only way to reach APIs Microsoft does not expose to public client
   applications. A BYO token cannot be refreshed - when it expires, supply a new one.

-  **Browser sign-in on every platform.**
   Interactive sign-in opens the system browser on Windows, macOS and Linux, by default.
   It is the default because it is the one place every sign-in method works: passkeys, FIDO2
   security keys, Windows Hello and phishing-resistant policies, none of which the embedded
   window can complete, and it reuses the browser session you already have. The requirement
   is the `http://localhost` redirect URI on the app registration; the default Microsoft
   application has it, a custom one may need it added. The setting **Use system browser for
   login** turns it off, in which case Windows uses the embedded window, or WAM when that is
   on. When no browser can be launched - a server session, a container, an SSH shell - the
   sign-in falls back to device code automatically rather than hanging. Device code can also
   be requested directly with `-DeviceCode`, which supports MFA, FIDO2 and security keys
   because the browser step happens on another device.

-  **Light and dark themes that follow the OS.**
   Both UI backends ship light and dark themes, and by default the whole application follows
   the desktop theme - the Windows app theme, the macOS appearance setting, and the GNOME
   colour scheme on Linux. It switches as soon as the OS setting
   changes, with no restart. Light or Dark can also be selected explicitly with the
   `AppTheme` setting. Linux desktops other than GNOME have no common way to report this, so
   they default to Light unless a theme is chosen.

-  **Access marking uses the user's Intune role, not only the app's scopes.**
   The menu already marked policy types the app's token could not use. It now also asks
   Intune what the signed-in user's role allows (`getEffectivePermissions`), so a user with a
   read-only Intune role sees the affected types in orange even when the app holds
   ReadWrite scopes. The Profile popup gains a **Permissions** button listing the token,
   role and effective level per type, and `Get-IMGraphEffectivePermissions` returns the same
   for scripts. Intune Administrator / Global Administrator skip the lookup. After a PIM
   activation or a new role assignment, click **Refresh** in the Profile popup - the answer
   follows the token. Setting `UseRbacAccessMarking` (default on) turns it off.
   See [Docs/EffectivePermissions.md](Docs/EffectivePermissions.md).

-  **Graph API call log.**
   A **Graph Calls** view lists every Graph request the session has made - method, URL,
   status code, duration and size - with filtering and refresh. It makes a slow export or an
   unexpected permission error diagnosable without turning on verbose logging or reading the
   log file.

-  **Settings and automation.**
   Settings can be read from and written to the registry, a JSON file, or an in-memory
   store that touches nothing on the machine. A run can import a settings file from source
   control, execute, and leave no trace. Settings resolve per tenant or globally.
   See [Docs/Settings.md](Docs/Settings.md).

-  **Documentation engine.**
   Policies can be documented to HTML, Markdown, Word, JSON, CSV, or Confluence storage
   format. Output providers, per-@odata.type handlers and input providers are all
   registries, so a new output format or a new policy type is an added file rather than an
   edit to a central switch. Enrollment notifications, Windows Hello for Business, Windows
   Restore and both Android enrollment profile types document with their portal labels for
   the first time, and the tenant-default enrollment policies - five of them share the name
   "All users and all devices" - are titled by their type.

-  **Cross-tenant migration only creates what the policy actually uses.**
   Import creates referenced Entra groups and assignment filters in the target tenant and
   rewrites assignment ids to match, and AD-synced source groups are recreated as cloud
   groups. The change from 3.x is scope: only the groups a policy actually references are
   created. 3.x imported every exported group, even when a single policy referenced one of
   them.

-  **Bulk scope tags.**
   Add, replace or remove scope tags across policy types or whole policy groups, with an
   orphan-cleanup mode that removes references to scope tags that no longer exist.

-  **Bulk assignments.**
   Add, replace or remove assignments across policy types or groups, including include and
   exclude targets, assignment filters, app install intents and health-script schedules.

-  **Graph batching and parallelism.**
   Requests are batched, and on PowerShell 7 batches can run in parallel. Both are settings.
   On a large tenant this is the difference between a 26-minute and a 4-minute full export.
   See [Docs/GraphBatching.md](Docs/GraphBatching.md).

-  **Multi Admin Approval support.**
   When a tenant requires approval for a Graph write, the pending-approval response is
   detected and the approval code is returned to the caller rather than surfacing as an
   opaque HTTP 403.

-  **Continuous Access Evaluation.**
   Claims challenges are handled and the request is retried transparently.

<br />

**Behind the scenes**

-  **PowerShell classes throughout.**
   Policy types, policy objects, authentication providers, compare providers and
   documentation handlers are now PowerShell classes with real inheritance. A policy type
   inherits its export, import, copy, compare and documentation behaviour from a base class
   and overrides only what differs, instead of each feature carrying a switch statement over
   every object type. Adding a policy type is a new subclass file - there is no central list
   to edit.

-  **Self-registering extension points.**
   Authentication providers, documentation output formats, per-object-type documentation
   handlers and compare providers all register themselves when they load. Adding one is
   adding a file.

-  **One engine, two user interfaces.**
   The WPF and Avalonia backends implement the same contract and contain no feature logic of
   their own, so both call exactly the same code underneath. That is also what makes the
   headless automation API possible.

-  **An automated test suite.**
   The engine is covered by an offline suite, and by an online release gate that exercises
   the full import, export, update, scope tag, assignment, documentation, copy, compare and
   delete lifecycle against a live tenant before a release is published. Both run in the
   development repository; this package is the application, without the test assets.

<br />

**Known limitations in this beta**

-  The Avalonia backend is new. It has been exercised on Windows, macOS and Linux, but far
   less than the WPF backend has been over the life of 3.x.
-  Application content cannot be exported unless the encryption information is available.
   Graph has no API for downloading decrypted app content, so an app whose encryption info
   the tenant no longer exposes can be exported as a policy but not as an installable
   package.
-  Inventory Policies will fail to list. Microsoft does not allow public client applications
   to call that API. It works with a BYO token - see **Two authentication providers**.
-  APIs that require additional licensing will fail for tenants without the licence. The
   object type is still listed; the request returns an error rather than an empty result.
-  A few policy types have no dedicated documentation renderer yet: Apple enrollment types,
   iOS app provisioning profiles, Inventory Policies, Entra and Intune branding, Terms and
   Conditions, Windows quality update policies and Windows 365. They document as their basic
   information plus one row per property. The **Document unsupported types** option turns
   that off, in which case they are skipped with a log entry rather than written out empty.
-  Copying a built-in administrative template can drop presentation values for settings that
   use them. Custom (ingested) ADMX is unaffected.
-  The `ADMX Files` policy type supports View, Import and Export only.
-  Types under `Intune Info` are read-only by design - Export and View only.
