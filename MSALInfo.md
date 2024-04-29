# Microsoft Authentication Library (MSAL)

Microsoft Authentication Library (MSAL) is used for authenticating to Microsoft Graph and other APIs used by the script. **Microsoft.Identity.Client.dll** is the library file that contains the authentication functions. This script is depending on the file to be able to authenticate to Microsoft Graph. 

See the GitHub repository [MSAL.PS](https://github.com/AzureAD/MSAL.PS) for more information on using MSAL with PowerShell. This script does not use the module directly for various reasons but the MSAL.PS repository is a great source of information and the best place to start for using MSAL with PowerShell. 

See GitHub repository [MSAL for .Net](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet) for additional information and source code. Latest version of MSAL can be downloaded [here](https://www.nuget.org/packages/Microsoft.Identity.Client/).

## Applications 

MSAL uses OAuth2 to authenticate to an Application in Azure. This script has two applications pre-defined in Settings:

- Microsoft Intune PowerShell (d1ddf0e4-d672-4dae-b554-9d5bdfd93547)
- Microsoft Graph PowerShell (14d82eec-204b-4c2f-b7e8-296a70dab67e)

Microsoft Intune PowerShell is the same application as the Intune PowerShell module uses and this is the default application for the script. 

The script will detect if the selected app is missing permissions and prompt for Consent if additional permissions are required. The script will also prompt for Consent if the application is not approved in the Tenant.

**Note:** Permissions in an app is just a filter and decides what can be done via the app. If a user only has Read permission but the app has ReadWrite, the user will not be able to update anything. 

A Custom application can be specified in Settings. This could theoretically be a custom created app in Azure. However, if an App is registered in Azure, it will be single tenant only. This can be used to backup an existing environment but the Azure App cannot be used to import data in another  tenant.  Use Enterprise Apps and Common/Organizations authority to allow access to multiple tenants.

## Token 

The MSAL authentication will create a token that is used when calling APIs in Microsoft Graph. This token is cached in an encrypted **msalcahce.bin3** file in the **%LOCALAPPDATA%\CloudAPIPowerShellManagement** folder. The file can only be decrypted by the same user. Caching the token can be disabled in Settings.

**Note:** Only the token is cached. Credentials are never stored.

The token expire after 1 hour. The script will do a "login" every time it calls an API. MSAL will manage the refresh of the token and only refresh it when it is about to expire or after it has expired. If the token needs to be refreshed e.g. the user was added to a new role, a forced refresh can be triggered in the Profile Info popup.  

The Token info will show information like role memberships, expiry time, scope etc. There are three tokens with information available:

* MSAL Token - Token created when authenticating with MSAL. Contains the Access and ID tokes
* Access token - Token contains permissions information and used when calling Microsoft Graph APIs 
* ID Token - Token contains user information and used when logging in 

Access and ID Tokens are encoded (**NOT** encrypted) Java Web Token (JWT) tokens. JWT tokens are based on three Base64Url strings separated with a dot (.) e.g. Header.Payload.VerifySignature.

See [JWT.IO](https://jwt.io/) for more information about JWT tokens.

## Multi Tenant Support 

Support for switching to other tenants can be enabled in Settings. This can be used if the user has a guest account in one or more tenants. This is disabled by default for the following reasons:

* Reduce login time - Getting the list of accessible tenants takes a few seconds extra 
* Reduce prompts - There is no API in Microsoft Graph that returns a list of tenants the current user has access to. Instead, a Azure management API is used. This will require permissions to Azure management which might cause an additional prompt for Consent when logging in.

**Note:** This is only used when a user has access to multiple tenants. Users from other tenants can always be used without enabling the 'Get Tenant List' setting.
