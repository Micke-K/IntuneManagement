# Using your own Entra app registration

By default, interactive sign-in uses Microsoft's public **Microsoft Graph PowerShell**
application (`14d82eec-204b-4c2f-b7e8-296a70dab67e`), so the tool works with no setup.
This page explains how to create your own app registration and point the tool at it.

## When you need your own app

- **Unattended runs.** Scheduled tasks, pipelines and runbooks sign in with a client
  secret, a certificate or a federated credential. Those need an app you own.
- **The OAuth provider.** It always signs in as an app you configure.
- **Your tenant blocks Microsoft Graph PowerShell.** Some tenants require user
  assignment on that app or allow only approved apps. Sign-in then fails with
  `AADSTS65001` (no consent) or `AADSTS50105` (user not assigned).
- **Narrower permissions or a clear audit trail.** Your own app carries exactly the
  permissions you grant, and sign-in logs show its name instead of Microsoft's.

## 1. Create the registration

In the [Microsoft Entra admin center](https://entra.microsoft.com):

1. Go to **Entra ID > App registrations > New registration**.
2. Give it a name, for example *Intune Management*.
3. Under **Supported account types**, pick *Accounts in this organizational directory
   only* unless you manage several tenants with one app.
4. Under **Redirect URI**, choose **Public client/native (mobile & desktop)** and enter
   `http://localhost`.
5. Select **Register**, then copy the **Application (client) ID** and the
   **Directory (tenant) ID** from the overview page.

Then open **Authentication** on the new app:

- Make sure `http://localhost` is listed under **Mobile and desktop applications**.
  Interactive sign-in runs in the system browser, and that is the only redirect a
  browser can return to. On macOS and Linux it is the only option at all.
- Set **Allow public client flows** to **Yes**. Device code sign-in (`-DeviceCode`)
  and the OAuth provider's username/password flow need it.
- Only if you turn on **Use Web Account Manager (WAM) for login** in Settings: also add
  `ms-appx-web://microsoft.aad.brokerplugin/<application id>` as a redirect URI under
  **Mobile and desktop applications**.

## 2. Add API permissions

Open **API permissions > Add a permission > Microsoft Graph**.

- Pick **Delegated permissions** for people signing in interactively.
- Pick **Application permissions** for unattended runs.

Add the permissions for the parts of Intune you manage. With all of them, every
policy type can be read and changed:

```
DeviceManagementConfiguration.ReadWrite.All   DeviceManagementApps.ReadWrite.All
DeviceManagementServiceConfig.ReadWrite.All   DeviceManagementScripts.ReadWrite.All
DeviceManagementRBAC.ReadWrite.All            Policy.ReadWrite.ConditionalAccess
Agreement.ReadWrite.All                       Organization.ReadWrite.All
CloudPC.ReadWrite.All                         Group.ReadWrite.All
```

- `Group.ReadWrite.All` is what lets a cross-tenant import create the groups a policy
  is assigned to. `Group.Read.All` is enough if you never import into another tenant.
- For a read-only identity, grant the `.Read.All` variant of each permission instead.
  The left-hand menu then marks the types that identity cannot change.
- For delegated sign-in, keep the `User.Read` permission that a new registration
  already has. The tool reads the signed-in user's name and photo.

Then select **Grant admin consent for \<tenant\>**. The tool asks Entra for every
permission the app has been granted (the `.default` scope), so a permission that is
added but not consented is simply missing from the token.

For delegated sign-in, the user also needs an Intune role that allows the action. The
app's permissions are the upper limit; the user's role decides within it. See
[Effective permissions](EffectivePermissions.md).

## 3. Add a credential (unattended runs only)

Under **Certificates & secrets**, either:

- **Upload a certificate** (recommended). Keep the private key in the certificate store
  of the machine that runs the job, or in a `.pfx` file.
- **Add a client secret.** Copy the value when it is shown; Entra never shows it again.
  Secrets expire, so note the date.

Interactive and device code sign-in need no credential.

## 4. Use the app in the tool

### In the app's Settings

Open **Settings > Authentication** and fill in:

| Setting | Value |
|---|---|
| Entra application | Leave **empty**. A built-in app selected here wins over your own. |
| Application Id | The **Application (client) ID** from step 1. |
| Tenant Id | The **Directory (tenant) ID**. Required for a single-tenant app: without it, sign-in goes to the multi-tenant endpoint and Entra refuses a single-tenant app there. |
| Redirect URL | Leave empty. The system browser always uses `http://localhost`. |
| Authority | Leave empty, unless you have been told to use a specific authority URL. |

Sign out, then sign in again. The new app is used from the next sign-in.

### From PowerShell

The same settings, saved:

```powershell
Import-Module .\IntuneManagement.psd1
Set-IMSetting -Key EntraCustomAppId    -Value '<application id>'
Set-IMSetting -Key EntraCustomTenantId -Value '<tenant id>'
Connect-IMIntuneManagement -Interactive
```

Or for one sign-in only, without saving anything:

```powershell
Connect-IMIntuneManagement -Interactive -TenantId '<tenant id>' -AppId '<application id>'
Connect-IMIntuneManagement -DeviceCode  -TenantId '<tenant id>' -AppId '<application id>'
```

Unattended, with the credential from step 3:

```powershell
# Certificate by thumbprint (CurrentUser\My, then LocalMachine\My)
Connect-IMIntuneManagement -TenantId '<tenant id>' -AppId '<application id>' -Certificate '<thumbprint>'

# Certificate from a .pfx file
Connect-IMIntuneManagement -TenantId '<tenant id>' -AppId '<application id>' `
    -CertificatePath C:\certs\intune.pfx -CertificatePassword (Read-Host -AsSecureString)

# Client secret - read it from a vault or a pipeline secret, never from the script
Connect-IMIntuneManagement -TenantId '<tenant id>' -AppId '<application id>' -Secret $env:INTUNE_APP_SECRET

# Workload identity federation (GitHub Actions, AKS), OAuth provider
Connect-IMIntuneManagement -Provider OAuth -TenantId '<tenant id>' -AppId '<application id>' `
    -FederatedTokenFile $env:AZURE_FEDERATED_TOKEN_FILE
```

For workload identity federation, add the matching federated credential to the app
under **Certificates & secrets > Federated credentials** instead of a secret.

To open the window signed in as the app instead of a user, connect first and then run
`Show-IMMainWindow`. The full recipe, with the secret read from a vault, is in
[Examples](Examples.md#open-the-window-signed-in-as-an-app).

### Government clouds

Register the app in the government tenant itself: the Microsoft Entra admin center for
US Government is [entra.microsoft.us](https://entra.microsoft.us). Then add `-Cloud` to
every connect - `USGov` for GCC High, `USGovDOD` for DoD - or set **Settings >
Authentication > Default cloud** for interactive sign-in. GCC (commercial) tenants use
the public cloud and need nothing extra.

Without the right cloud, an app sign-in to a GCC High tenant fails with
`AADSTS900382: Confidential Client is not supported in Cross Cloud request`.

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `AADSTS50011` redirect URI mismatch | `http://localhost` is missing from the registration | Add it under **Authentication > Mobile and desktop applications** |
| `AADSTS65001` user or administrator has not consented | Permissions were added but not consented, or the tenant does not allow user consent | Select **Grant admin consent** on the app's API permissions |
| `AADSTS50194` application is not configured as multi-tenant | A single-tenant app was used without a tenant | Fill in **Tenant Id** in Settings, or pass `-TenantId` |
| `AADSTS7000218` request body must contain client_assertion or client_secret | Device code or username/password sign-in on an app that does not allow public client flows | Set **Allow public client flows** to **Yes** |
| `AADSTS700016` application not found in the directory | Wrong application id, or the app lives in another tenant | Check the id; for a single-tenant app, sign in to the tenant that owns it |
| Menu items marked as read-only, or 403 errors | The token lacks a write permission | Add the `.ReadWrite.All` permission and consent again, then sign out and back in |

When something else fails, the log names the cause. Turn on the **Debug** setting
for more detail, and the MSAL log (`msal.log` in the app data folder) records the full
sign-in exchange.
