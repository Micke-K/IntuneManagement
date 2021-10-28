# Release Notes

## 3.3.1 (Beta) - 2021-10-28

This is a **BETA** release. It contains core changes for Authentication and Settings management. Please report any issues [here](https://github.com/Micke-K/IntuneManagement/issues).

**New features**

- Added support for selecting GCC when using US Government Cloud

- Tenant Specific Setting
  
  The script now supports tenant specific settings. This can be used in scenarios like: only allow delete on you test environments, tenant specific Intune app folders etc.
  Login settings like Cloud and GCC is only used if logging on with a cached token. It will otherwise use the current tenant settings. 

  **Test feedback request:** If there are any users accessing multiple cloud environments like US Government with different GCC levels, please report any issues, working or not. Please report it to [Issue 26](https://github.com/Micke-K/IntuneManagement/issues/26)

  Note: Not all settings have be tested and verified and only Setting Values are supported e.g. last Bulk Compare strings are global. Cached settings might not be updated when connecting to another tenant.   
  
- Log View 
  
  View the log of the current session in the app
  
- Added support for documenting scripts for Word
  This is based on [Issue 34](https://github.com/Micke-K/IntuneManagement/issues/34)

  - New Script options in the Output option tab e.g. enable/disable script documentation, remove PowerShell signature block and documentation styles
  - Supports PowerShell/Shell scripts, Proactive Remediations and Win32 Apps (Requirement/Detection scripts)
  - Scripts will be documented in a separate table with style *HTLM Code* by default. Spell check is disabled for the script text.

- Permission detection if **Use Default Permissions** is enabled
  
  Default permissions will only use the permission consented to the selected Azure App. The script will check the required permissions with the Access Token. If permissions are missing for one or more objects, they will be marked as red in the menu or they can be excluded from the menu by enable **Hide No-access items** in Settings 

**Default Settings Value Changes**

* **Use Default Permissions** is now set to Disabled by default. With the Tenant Specific Settings feature, this can now be enabled globally or per tenant. Consultants accessing multiple environments might not have permissions to grant consent requests so this could be enabled on a global level and then disabled for tenants where the permissions can be added.  

**Fixes**

* Fixed an issue when using Json settings where it could not add child settings  

## 3.3.0 (Beta) - 2021-10-17

This is a **BETA** release. It contains core changes for Authentication and Settings management. Please report any issues [here](https://github.com/Micke-K/IntuneManagement/issues).

**New features**

- Support for Settings in Json files
  Settings can now be stored in json files and copied between devices.  

  See [Readme](README.md#Settings) on how to use this feature
  This is based on [Issue 33](https://github.com/Micke-K/IntuneManagement/issues/33)

- Bulk Compare for exported folders

  The tool can now compare two exported folders 
  This is based on [Issue 32](https://github.com/Micke-K/IntuneManagement/issues/32)

- Support for Azure AD US Government cloud and Azure AD China cloud. Default is Azure AD Public cloud. 

  Change cloud in Settings
  **Note:** This is a major change to the authentication. This may have an impact if a custom configured Azure app is used.
  This is based on [Issue 26](https://github.com/Micke-K/IntuneManagement/issues/26). Please report any problem, progress or testing with US Government/China cloud or if there are any issues when a custom configured Azure app is used. 

- Export can now add Id to the name of the backup file

  This can be used if there are multiple objects with the same name.

  This can be enabled in Settings. Backup file name will be <Name>_<Id>.json. 

- Export/Import/Compare/Delete now supports name filter
  Objects are filtered based on escaped RegEx -nomatch expression so wildcards are not supported. 

- IntuneAssignments report will now include the id of deleted groups 

**Fixes**

* Fixed an issue in Export. Groups were not exported if exporting multiple times and multiple folders during the same session.
* Fixed an issue in Compare where the csv file was not stored in the correct folder
* Fixed an issue in Compare where the comparing object may return System[]. This can happen if the generated files has multiple documentation items for a property. First result will be used.  

## 3.2.3 - 2021-10-07

**New features**

- Added support for Terms of Use Export/Import.
  This requires that the pdf file is available during import, in either the export folder or the Intune App folder. This is added as a Known Issue in [Readme](README.md).

  **Note:** This is in preview and it requires that the Preview option in Settings is enabled and then a script restart. This will most likely generate a new consent prompt.
  This is based on [Issue 27](https://github.com/Micke-K/IntuneManagement/issues/27)

- All objects are returned

  This might take long time in huge environments. 
  Please report feedback on how this works in environments with 1000+ objects e.g. does it take too long time, memory issues etc.
  This is based on [Issue 29](https://github.com/Micke-K/IntuneManagement/issues/29)

- Added support for custom columns
  This must be manually added to the registry.

  See [Readme](README.md#Columns) on how to use this
  This is based on [Issue 30](https://github.com/Micke-K/IntuneManagement/issues/30)

- Object count will be displayed

**Fixes**

* Fixed minor bugs in IntuneAssignments - Support Name for objects that don't use displayName
* Regenerated documentation and language files - New properties for the iOS Device Restriction profile is now supported

## 3.2.2 - 2021-09-23

**New features**

- Added support for setting Conditional Access policy state during import. The default setting is to import Conditional Access policies with the same sate as they were exported.
  This is based on feature request [Issue 25](https://github.com/Micke-K/IntuneManagement/issues/25)
  Note: Security defaults must be disabled before Conditional Access policies can be imported as Enabled.

**Fixes**

* Fixed bugs when using the ImportExtension command

## 3.2.1 - 2021-09-04

**New features**

- PowerShell Scripts can now be viewed and edited in the tool 
- Intune Tools
  - Added Intune Assignment - Simple tool to quickly gather all assignments from exported objects
- Documentation
  - Added documentation support for
    - Scope (Tags) 
      Note: This will generate one section for all Scopes in the word document
    - Health Scripts (Remediation Scripts) 

**Fixes**

* General

  * Custom Device Configuration profiles will convert encrypted OMA URI values when the full object is loaded instead of only during Copy and Export.
  * All file exports are now saved in UTF8   

* Compare

  * Fixed issue where the wrong name was specified if the compare object was missing
  * Administrative Templates, Settings Catalog and Endpoint Security will always compare based on documentation.
  * Encrypted OMA URI values are now supported

* Documentation

  * Minor updates to support documenting all objects of a specific object type in one section instead of one section per object
  * Fixed "Not Configured" value issues for empty arrays 
  * Fixed documentation of Microsoft 365 Apps when XML is used
  * Minor updates on VPN profile documentation. EAP XML will be in XML format and removed duplicate SplitTunneling values.
    Note: The EAP XML will require manual update of the column sizes in Word 
  
  

## 3.2.0 - 2021-08-15

**New features**

- Intune Tools (New View)

  - **ADMX Import** - Configure settings for 3rd party ADMX files with a UI similar to GPMC and create a Custom Profile based on the configured settings

  - **Reg Values** - Add registry values to HKLM or HKCU. This will create and ADMX based on the configured settings and create a Custom Profile in Intune

    See [ADMX Import](ADMXImport.md) for more information on how this works

    **Note:** There is only Import functionality in this version. It does not support updating an existing Custom Profile ADMX policies.
    
  - **Important!** Consider this tool to be in preview at this moment. It has only been tested on Cloud only joined devices. It looks like there are different functionality in the Policy CSP between hybrid and cloud only joined devices. 
    It would be great if anyone testing this on hybrid (or cloud only joined) devices could create an issue and report back the findings, even if it works as intended.

  There are indications that Microsoft is implementing this into the portal UI. The [groupPolicyUploadedDefinitionFile](https://docs.microsoft.com/en-us/graph/api/resources/intune-grouppolicy-grouppolicyuploadeddefinitionfile?view=graph-rest-beta) API suggests that the portal will support this in the future. It would be good if this could be integrated with the Settings Catalog.

- Documentation:

  - Create cover page and table of contents when no template is selected
  - Select CSV delimiter when documenting to a CSV file

- Compare:

  - Select CSV delimiter for bulk compare

**Fixes**

* Authentication

  * The script will start even if it failed to add type TokenCacheHelperEx. 

    This is based on [Issue 21](https://github.com/Micke-K/IntuneManagement/issues/21)

    **Note:** The token will not support caching if this fails. This could be caused by not having write access to the \CS folder or by restrictive ASR policies

* Export/Import

  * Added support for exporting OMA-URI values that are stored encrypted. 

    **Note:** OMA-URI strings and XML Files are stored encrypted. These values will be decrypted and stored in clear text. Be careful if sensitive data is stored e.g. passwords. 

  * Fix for updating existing Autopilot profiles during import. A new property was added that broke the functionality.
    This is based on the feature request in [Issue 17](https://github.com/Micke-K/IntuneManagement/issues/17)

* Documentation

  * New handling of Not Configured properties. Skipping unconfigured properties will now skip all these properties during documentation 
  * Minor fixes to avoid duplicate documentation of properties  

* Compare

  * Fixed bugs when comparing Intent objects (Endpoint Security) policies in Documentation mode.

* Copy

  * Copy Custom Profiles with encrypted values

**Additional Changes:**

* Documentation files has be re-generated to support new\updated properties on Property based objects. 

## 3.1.8 - 2021-07-18

**New features**

- Forget cached users - Forget a user by clicking on the bin icon in the user information. This will remove the user from the cached file. It will not remove it from the browser cache.
- Update existing profiles during import is moved to preview. 
  **Important:** See the Import section in the [Readme](README.md#Import) file for more information
  This is based on the feature request in [Issue 17](https://github.com/Micke-K/IntuneManagement/issues/17)

**Fixes**

* Fixed a bug when exporting Settings Catalog. When exporting settings based on key/value pairs, some parts were not converted to json objects. Import worked but not the update. Depth parameter was increased in the ConvertTo-Json functions. 

## 3.1.7 - 2021-07-12

**New features**

- Support for documenting Notifications
- **PREVIEW/EXPERIMENTAL** - Support for Replace/Update existing profiles/policies during import. 
  See the Import section in the [Readme](README.md#Import) file for more information
  This is based on the feature request in [Issue 17](https://github.com/Micke-K/IntuneManagement/issues/17)

**Fixes**

* Fixed bug that caused an exception when listing App Protection objects and only one object existed in the environment. 

  See [Issue 15](https://github.com/Micke-K/IntuneManagement/issues/15) for more info

* Import Priority based objects in the priority order specified in the files (Enrolment Restrictions and Autopilot profiles) 

* Set default settings for the options in the Import forms (Based on Settings)

* Delete Autopilot profiles with assignments

* Moved the assignments import to a separate function 

## 3.1.6 - 2021-07-07

**Fixes**

* Fixed invalid file name characters - [Issue 19](https://github.com/Micke-K/IntuneManagement/issues/19)

  * Added -LiteralPath to Get and Set-Content
  * Save CSV in document
  * Import/Export Administrative Template and Role Definitions 
  * Saving the PowerShell script file
  * Export with assignments for multiple profiles  

* Added support for [ and ] in file names

  **Note:** This can cause duplicate files if exporting to the same location as pre 3.1.6 export and the profile name contains [ or ]

* Changed to custom documentation for Custom OMA-URI profiles  

* Administrative Template now includes definitionValues in detailed view and export

* Fixed exporting PowerShell script in Bulk export. Option was only available if PowerShell was active type.

* Fixed issue with MigrationTable when exporting from two different environments without restart. The Group information was save to the same MigrationTable. 

## 3.1.5 - 2021-07-06

**Fixes**

* Fixed rushed update for [Issue 18](https://github.com/Micke-K/IntuneManagement/issues/18) 
* Fixed bug in Compare module

## 3.1.4 - 2021-07-06

**Fixes**

* Fixed issue importing Administrative Templates

  See [Issue 18](https://github.com/Micke-K/IntuneManagement/issues/18) for more info

## 3.1.3 - 2021-07-05

**New features**

- Bulk Compare
  - Compare with exported files
  - Compare with existing objects based on name patterns
- Bulk Copy 
  - Copy existing objects based on name patterns
- Support for documenting PolicySets
- Release Notes check - Check if there are any updates by comparing the local version of ReleaseNotes.md with the GitHub version 

**Fixes**

* Fixed bug that caused an exception when exporting objects with an assignment and the 'Export Assignment' option disabled. 

  See [Issue 16](https://github.com/Micke-K/IntuneManagement/issues/16) for more info

* Export Assignments in Bulk Export and Object Export did not get default value from Settings 

* Fixed issue where the required permissions were not passed during authentication  

## 3.1.2 - 2021-06-20

**New features**

- Delete and Bulk Delete - Delete selected items or delete ALL items of selected object types

  **Note:** This must be enabled in the settings. They are not visible by default.

  **WARNING:** Use this carefully! It will delete profiles and policies in Intune. 

- Support for new object Health Scripts

- Object permissions is now handled by ViewObject and authentication provider. This is to support future view extensions.

**Fixes**

* Azure Role Read permission can be disabled in settings
* Minor UI changes e.g. List Boxes for bulk Import/Export changed to DataGrid
* Minor bulk export fixes

## 3.1.1 - 2021-06-16

**New features**

- Download script for Custom Attribute
- Documentation
  - Added support for additional objects (Enrollment restrictions)

**Fixes**

- Failed to get user information during logon. Something was changed in Graph that caused calling ME with full ODATA to fail.

- Added RoleManagement.Read.Directory as a default required permission.

  **NOTE:** This will most likely cause a consent prompt 

- Some additional minor fixes

## 3.1.0 - 2021-06-08

**Breaking Changes**

* Export folder for PowerShell is changed to PowerShellScripts

**New features**

- Documentation
  - Document Intune objects to CSV or Word
  
  - Object will be documented with text from Intune
  
  - Document in any supported language
  
  - See [Documentation](Documentation.md) for more information
  
    **Note:** This is the first version of the documentation. It does not support ALL object types and might be missing some properties 
- Compare
  - Compare objects with exported files
  - Property comparison
  - Documentation comparison 
- Support for additional objects
  - Quality Updates, Filters, Mac Scripts and Custom Attributes

* Azure AD role displayed in token info
* Fixed sending additional headers in the Graph request 

**Fixes**

- Support for Privacy Access Controls in Windows 10 Device Restriction profiles
- Support for AppLocker files in WIP policy
- Select All checkboxes moved from below data grids to the header of the column

## 3.0.0 Beta 1 - 2021-04-01

**Breaking changes**

- Dropped support for Azure Branding and MAM/MDM settings...for now
- Import might not work for items exported with previous versions. Some folders are renamed, import is depending on additional information.

**New features**

- Authentication managed by Microsoft Authentication Library (MSAL)
  - Support for switching user 
  - Support for switching tenant. Multi tenant support must be enabled in Settings
  - Token info, Profile picture info support etc.
  - See [MSAL info](MSALInfo.md) for more information
- Support for multiple Views - Intune Management and Intune Info for now...
  - Intune Management - Export/Import/Copy objects in Intune
  - Intune Info - Show information about some objects in Intune
- Improved UI experience
  - Support for resizing the Window
  - Support for searching for objects
  - Refresh objects in the list
  - Scaled popup dialogs
- API management redeveloped from scratch to simplify support for new object types in the future
- Support for new object types (Settings Catalog, Named Locations, Scope Tags, Policy Sets etc.)
- Better support for migrating objects between environments  
  - Group migrations e.g. support for Dynamic Groups, different group types etc.
  - Support for dependency objects e.g. Policy Sets reference other objects like Compliance Settings etc. The import of an object uses exported json files to identify dependent items and map old Id to the new Id in the target environment  
  - Support for migrating Scope Tags (Uses the dependency functionallity so Scope Tags must be Exported/Imported)
  - Better support for migrating Assignments

**Dependencies**

- MSAL - **Microsoft.Identity.Client.dll**. This is included in Az / MSAL.PS modules or it can be installed separately. This release was developed and tested with MSAL version 4.21.0.0. 

## 2.0.0 - 2021-02-01

**Breaking changes**

- Removed support for AzureRM

**New features**

- Support for Az module

**Fixes**

- Allow more than 9 Conditional Access policies. Issue [#5](https://github.com/Micke-K/IntuneManagement/issues/5)
- Include WIP policies. Issue [#7](https://github.com/Micke-K/IntuneManagement/issues/7)
- Import is not working. Issue #6 and [#4](https://github.com/Micke-K/IntuneManagement/issues/4)
- Intune module can now be install with scope user. Issue [#8](https://github.com/Micke-K/IntuneManagement/issues/8)

## 1.0.0

- Intune Management with PowerShell
- Dependencies: Intune and AzureRM PowerShell modules
