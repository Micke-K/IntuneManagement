# IntuneManagement with PowerShell and WPF UI

<p align="center">
  <a href="https://twitter.com/Micke_K_72">
    <img src="https://img.shields.io/twitter/follow/Micke_K_72.svg?style=social" target="_blank" />
  </a>
</p>

These PowerShell scripts are using Microsoft Authentication Library (MSAL), Microsoft Graph APIs and Azure Management APIs to manage objects in Intune and Azure. The scripts have a simple WPF UI and it supports operations like Export, Import, Copy, Download, Compare etc.

This makes it easy to backup or clone a complete Intune environment. The scripts can export and import objects including assignments and support import/export between tenants. The scripts will create a migration table during export and use that for importing assignments in other environments. It will create missing groups in the target environment during import. Group information like name, description and type will be imported based on the exported group e.g. dynamic groups are supported. There will be one json file for each group in the export folder.

The script also support dependencies e.g. an App Protection is depending on an App, Policy Sets are depending on Compliance Policies, objects has Scope Tags etc. Dependency support requires exported json files and that the dependency objects are imported in the environment. The script uses the exported json files to get the Id and names of the exported object and uses that information and updates Id's before import an object from a json file. The Bulk Import form shows the import order of the objects. The objects with the lowest order number will be imported first.

<p align="center">
  <img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/IntuneManagement.PNG" height="80%" width="80%"/>
</p>

This PowerShell application is based on the foundation modules CloudAPIPowerShellManagement and Core. These modules manages UI, settings, logging etc. The functionality for the application is located in the extension modules. This makes it easy to add/remove features, views etc. Additional features will be added...

**Security note:** Since the scripts are not signed, a warning might be display when running it and files might be blocked. The script will unblock all files. This is to avoid issues that it fails to load the MSAL library etc. If there are any security concerns, the PowerShell code can be reviewed and the DLL files can be downloaded manually from Microsoft repositories, see links below. The DLL files *CAN* be removed but MSAL is a pre-requisite for authentication. The script will try to find the DLL in the Az or MSAL.PS module if not found in the script root directory. DLL files are included to reduce dependencies.

## Starting the App

Before starting the app:

* The CMD files needs to be unblocked before the app can be started. The app can be started without it but Windows will prompt with a security warning.
* The script will unblock all other files

Before logging on:

* The app will use the Intune PowerShell Azure Enterprise Application by default and only use the permissions granted to that appliction. Disable **Use Default Permissions** in Settings to request additional permissions. This will cause a consent prompt if one or more permissions are missing for the app.
  **Note:** If the app has not been approved for the organization, a consent prompt will be displayed.  
* Enable **Get Tenant List** in Settings if accessing multiple environments with the same account e.g. a guest account in other tenants. This might cause a Consent prompt

Start the script by running **Start.cmd**, **Start-WithJson.cmd**, **Start-WithConsole.cmd** or **Start-IntuneManagement.ps1**. **Start-WithConsole.cmd** will leave the command prompt window open so you can see the log while running the app. 

## Silent Batch Job (Beta)

The script can be executed without UI. This is to support DevOps environment.

Silent batch job is supported by the following features:

* Export
* Import
* Documentation
* Compare


The Silent Batch Job feature requires an Azure App to be configured with a secret or a certificate. The app must also be delegated with the required permissions to Graph objects used by the tool.

The silent feature uses an exported json file with settings for the specified operation. This file can be generated in the Bulk Export/Import forms. Required settings for the silent job is configured in the form and then exported to a file. The path to the file is then passed on the command line. The file can be used for multiple environments if the **Add company name to the path** option is selected. Note that this requires that the app must have at least read permission on the Organizations API.

The following variables can be used in the path and filename:

* %Date% - This will be translated to yyyy-MM-dd format (eg 2020-02-27)
* %DateTime% - This will be translated to yyyyMMdd-HHmm format (eg 20200227-1750)
* %Organization% - Name of the tenant
* Any environment variable 

The tool will by default generate the files; `BulkExport.json` and `BulkImport.json`. These files can be merged into one file but that must be done manually. These files can also be edited manually. Each setting represents a control in the UI form. When the script is triggered silently, it will create the form in the background, populate it with the values from the file and then trigger the bulk function.

**Note** The Silent Batch feature use settings configured in the UI. If this is triggered in a DevOps envionment, it is recommended to generate a settings JSON file with the desired settings and then use that in the DevOps environment

The app authentication can either be passed on the command line or stored in the settings. Tennant Settings is required for multiple environments.

**Command line example:**

Start-IntuneManagement.ps1 -Silent -TenantId "<*TenantID*>" -SilentBatchFile <*PathToFile*> [-AppId <*AppId*>] [-Secret <*Secret*> | -Certificate <*CertThumb*>]

Start-IntuneManagement.ps1 -Silent -SilentBatchFile "C:\Temp\BatchImport.json" -TenantId "00000000-0000-0000-0000-000000000000" -AppId "00000000-0000-0000-0000-000000000001" -Secret "KJ76P~B9###9-.8I####-_MySecret"

**Setting example:**

Start-IntuneManagement.ps1 -Silent -SilentBatchFile "C:\Temp\BatchImport.json" -TenantId "00000000-0000-0000-0000-000000000000"

## Documentation

This script has an extension that can document profiles and policies in Intune. The output is using the same language strings as the Intune portal.

See [Documentation](Documentation.md) for more information

## Import

The script can import the exported json files in multiple ways.

* **Always import:** The script will try to import the file. It will not check if it exists. 
  This is the default behavior
  
* **Skip if object exists:** The script will look if there is an existing object with the same name and type. It will not import the file if existing object is detected

* **Replace (Preview):** If an existing object is detected, the script will 
  
  * Import the file without assignments
  * Copy assignments from the existing object
  * Run PostReplace commands - Priority will be set for Enrollment Restrictions etc.
  * Update PolicySets object(s) to use the new imported object (detected by policySet assignments)
  * Delete the original object 
  
* **Update (Preview):** This will update the existing object.
  
  The update APIs do not support all the properties that the import APIs do and object types behaves differently during update e.g. Settings for **Endpoint Security** objects will not be cleared. There is no API for removing settings, only adding. If a setting does not exist in the import file, the existing setting will be set to *Not Configured*. **Settings Catalog** replaces the whole settings property during update. 
  
  This has been tested with all supported object types *except* Import Scripts (Shell), Android OEM Config and Apple Enrollment Types.
  
  
  Each application type works differently. Update functionality has been tested on Win32, Windows MSI LoB, iOS Store, Microsoft Store and Microsoft 365 (Windows and MacOS).

**WARNING:** Use Replace with caution! Replace will delete the existing object after a new object is imported and the assignments are copied, but it could cause issues in the environment if something in the process goes wrong. Replacing single objects can break references e.g. replacing an Application can break AutoPilot profiles, App Protection and App Configuration policies. Verify the process in a test environment before using this!  

**Recommendation:** Backup all policies before running Replace/Update.

The Replace/Update feature can be used in a scenario where all profiles/policies are managed in a separate reference (Dev/Test) and then implemented in one or more destination environments. The existing objects will then be reset to have the same settings as the reference environment

## Comparison

This script has an extension that can compare objects in Intune with exported json files. It will display a data grid with the values and highlight updated values with red.

Objects can be compared based on property values or documented values. 

The property value method is a quick way to compare objects but it will only show the names and values of the native Intune object. This is not a good comparison method for Settings objects since they have all the settings in one property.

The documentation method is a bit slower but will show the values as they are stated in the Intune portal. This is the recommended way to compare objects but note that this is only supported on object types that supports documentation. 

Bulk compare is supported. This can be performed in two ways:

* **Intune Objects with Exported Files** - This will read each exported file and compare it with the existing object

  The result file will be stored in the exported folder structure. Either in the Object Type folder or the parent folder depending on the Save as setting. 

  **Note:** This cannot be used with files exported from a different environment since it used the Id as identifier

* **Named Objects in Intune** - Compare Intune objects based on patterns

  This can be used in where a pattern is used separate objects between different environments e.g. [Test] Policy 1 vs [Prod] Policy 1.

* **Files in Exported Folders** - Compare two exported folders

  This can be used in where a pattern is used separate objects between different environments e.g. [Test] Policy 1 vs [Prod] Policy 1.

Output files are by default stored in the My Documents folder.      

The output CSV can either be one file for ALL objects or one file for each Object Type.   

## Bulk Copy

Bulk copy can be used to clone objects based on a name pattern. This can be used in the same scenario as Bulk Compare where the object names includes an environment identifier. The application will identify all objects matching the source pattern and copy each object with a new name matching the 'Copy object name pattern'. The object will not be copied if it detects that an object already exists with the new name.

**Note:** Assignments will NOT be copied. 

## Intune Tools

Additional Intune Tools is included in the script.

* ADMX Import
  * Import 3rd part ADMX settings like Chrome etc. by using a UI similar to GPMC
  * Support for any language if there is an ADML file for it 
* Reg Values
  * Create registry settings in HKLM or HKCU. Supported value types:
    * String
    * Multi-String
    * Expandable String
    * DWORD
    *  List (Key/Value pair) 
  * This tool creates a custom ADMX file based on the specified registry keys.

* Intune Assignments
  * Quickly gather all assignments from an export
  * Easily identify all profiles/policies a group is assigned to
  * Export as CSV to analyse in Excel 

See [ADMX Import](ADMXImport.md) for more information about the ADMX tools

## Columns

Custom columns is supported. The script will by default add id, displayName and description with exception for some object types. These are configured in the EndpointManager.ps1 and  EndpointManagerInfo.psm1 files.

Custom columns can be added for each Object Type in the detailed view of an Object. 

Custom Columns settings are stored in the regitry, in HKCU\SOFTWARE\CloudAPIPowerShellManagement\EndpointManager\ObjectColumns\\*ObjectType*.

The *ObjectType* value represents the ViewItem ID specified in the EndpointManager.ps1 and EndpointManagerInfo.psm1 files. This is also the same as the parent folder when items are exported.

The REG_SZ value has the following syntax:

[0|1],Property1[=Header1],Property2[=Header2],...

0|1 - Optional value specifying if columns are added or replacing default columns

0 = Replace default columns with columns specified in the registry value
1 = Add columns specified in the registry to the default columns. This is default setting.

PropertyX - Property to display. View an object to see available properties 

HeaderX - Optional value for the column header. Property name will be used if this is not specified.

**Note:** Some object types returns multiple object types, @OData.Type. If a custom column is added but the property does not exist on all the object types that were returned, the column value will be empty. The code will **not** break if the property is missing on one or more returned objects. 

**Example 1:** 
Reg key: HKCU\SOFTWARE\CloudAPIPowerShellManagement\EndpointManager\ObjectColumns\DeviceConfiguration

REG_SZ value: **lastModifiedDateTime**

This will add the **lastModifiedDateTime** to the existing columns.

**Example 2:**

Reg key: HKCU\SOFTWARE\CloudAPIPowerShellManagement\EndpointManager\ObjectColumns\TermsOfUse

REG_SZ value: 0,id=Id,displayName=Name,files[0].displayName=Display Name

This will replace the default columns and add new columns with specific header. 

**Note:** Sub-properties are supported e.g. **files[0].displayName**. This will add a column based on the first object in the files property.

## Change log

See [Change Log](ReleaseNotes.md) for more information

## Authentication
See [MSAL Info](MSALInfo.md) for more information about authentication

## Settings

Settings for the script is default stored in the registry. However, the script supports settings to be stored in a json file so it can be copied between computers. Registry settings can be exported in the Settings dialog. 

To use settings based on a json file:

```
Start-IntuneManagement.ps1 -JSonSettings [-JSonFile <PathToFile>]
```

If only -JSonSettings is used the script will use the default json setting file:

```
%LOCALAPPDATA%\CloudAPIPowerShellManagement\Settings.json
```

Use -JSonFile for custom location of the file

Start-WithJson.cmd is included as an example on how to start the script with json settings.

**Note:** If the file can't be created, the script will revert back registry. Make sure that the script can write to the file. It is not recommended to store the file in a folder that requires UAC to get write permissions.  

## Supported Intune objects
* ADMX Files [^1]
* App Configurations (App and Device)
* App Protection
* Applications [^2]
* Apple Enrolment Types - NOT fully tested
* Autopilot profiles
* Baseline Security profiles
* Compliance policies
* Compliance policies v2
* Compliance scripts
* Conditional Access
* Custom Attributes
* Device Configuration (Administrative Templates, Configuration Policies, Android OEM Config, Settings Catalog)
* Endpoint Security (Account Protection, Disk Encryption, Firewall, Security Baselines etc.) [^3]
* Enrollment Restrictions
* Enrollment Status Page profiles
* Feature Updates
* Filters
* Intune Branding (Company Portal)
* Locations
* Named Locations
* Notifications
* Policy Sets
* Quality Updates
* Reusable Settings
* Role Definitions
* Scope Tags
* Scripts (PowerShell and Shell scripts, supports download of script)
* Terms and Conditions
* Terms of Use
* Update Policies
* Co-management Settings

[^1]: ADMX file must be located in the export folder or in the folder defined in Settings. Administrative Template policies based on ADMX files must be imported after the ADMX file is imported.
[^2]: Export is support but .intunewin files will not be exported. The .intunewin file must exist in the exported folder or in the folder defined in Settings -> App packages folder.
[^3]: Some Endpoint Security policies will be exported to the SettingsCatalg since they are based the Settings Catalog API.

**Note:** The scripts are using the BETA version of the Graph API which might change at any time.

## Azure Management APIs
* Tenants for the current user

**Note:** Azure Management APIs are undocumented APIs which might not be supported and they might change at any time.

## Pre-requisites
* .Net 4.7
* PowerShell 5.1
* MSAL
  * Microsoft.Identity.Client.dll version 4.42.1.0 is included in this version
* License and permissions in Azure to manage objects in Intune and Azure 

## References
* [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/api/overview?toc=./ref/toc.json&view=graph-rest-beta) 
* [Microsoft.Identity.Client](https://www.nuget.org/packages/Microsoft.Identity.Client/) (MSAL download) 
* [MSAL.PS Module](https://github.com/AzureAD/MSAL.PS)
* [Az PowerShellModule](https://docs.microsoft.com/en-us/powershell/azure/new-azureps-module-az)
* [Microsoft Intune PowerShell Module](https://github.com/microsoft/Intune-PowerShell-SDK)
* [Microsoft.WindowsAPICodePack](https://www.nuget.org/packages/Microsoft-WindowsAPICodePack-Core) and [Microsoft.WindowsAPICodePack.Shell](https://www.nuget.org/packages/Microsoft-WindowsAPICodePack-Shell) for Browse Folder dialogs

## Acknowledgments
The app encryption and upload is based on [Graph PowerShell Intune Examples](https://github.com/microsoftgraph/powershell-intune-samples)<br />
Some MSAL functionalities are based on [MSAL.PS Module](https://github.com/AzureAD/MSAL.PS)

## Known Issues

Device Configuration and App Configuration objects are split up in different object types. They are using different Graph APIs and each object type in the menu uses one API. This is also why all Endpoint Security objects are of the same object type. They use the same API but are separated based on the Baseline Template Id they use.

Android Store Apps are **not** imported. The Create API is documented in Microsoft Graph but it's not working. Looks like these apps must be synched from Google Play.

Using multiple tenants support causes multiple logins/consent prompts the first time if 'Microsoft Graph PowerShell' is used. Querying the API for tenant list uses a different scope that is not included by default in the 'Microsoft Graph PowerShell' app. 

~~Using multiple tenants support *might* cause and endless loop in the login screen and cause duplicate accounts in token cache. Actual cause is not found yet but it happens on rare occasions and it looks like it happens when a guest account is used. Workaround: Cancel the login, restart the script, logout and restart the script again.~~ - Not seen this in a long time. Please create an issue if this happens

When multi tenant settings is Enabled/Disabled, the Profile Info is not updated until the account is changed or app is restarted. Profile Info popup is built after logon.

The *List Applications* API might not list an imported app immediately after the import. Click *Refresh* to reload the application objects.

~~When using the filter box to search for items, the checkbox must be clicked twice to select an item.~~<br />
Issue fixed in 3.3.2 

Logout will only clear the token from cache and not from the browser e.g. if login is triggered after a logout, the user will still be listed in the 'Select user' dialog.

~~Referenced settings will NOT be imported/copied. There is no value stored in a property on the object for these settings. Example: A VPN profile has certificates as referenced properties. The certificates must be added manually after import/copy.~~
Issue fixed in 3.5.0 

Terms of Use requires that the pdf file is available. This must be manually copied to either the Export folder for Terms of Use or to the Intune Application folder specified in Settings. It is currently not possible to export the pdf file with Graph API. 

The ADMX/ADML files must be manually copied to either the Export folder for Administrative Templates or to the Intune Application folder specified in Settings. It is currently not possible to export the pdf file with Graph API.

See [Documentation](Documentation.md) for issues regarding the documentation process.

## TIP

Check the log file for errors. The UI might not show errors why login failed etc. The log uses the Endpoint Configuration Manager (SCCM) format and it is best viewed with CMTrace or OneTrace. An old version of CMTrace can be downloaded [here](https://www.microsoft.com/en-us/download/confirmation.aspx?id=50012).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.