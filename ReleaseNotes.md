# Release Notes
## 3.9.8 - 2024-10-12

**New features**

- **Intune Info**<br />
  - Added 'Baseline Templates - Settings Catalog'<br />
  This list templates for Settings Catalog policies eg. Security Baseline for Windows 10 and later<br />

**Fixes**

- **Import/Export**<br />
  - Fixed support for export/import App Configurations (Device) - Android between environments<br />
  Based on [Issue 255](https://github.com/Micke-K/IntuneManagement/issues/255)<br />
  Thank you @jimmywinberg for all the testing!<br />
  - Fixed support for export/import App Configurations (Device) - iOS (VPP) between environments<br />
  Based on [Issue 260](https://github.com/Micke-K/IntuneManagement/issues/260)<br />
  Thank you @Arne-RFA for all the testing!<br />
  - Added support for exporting Groups targeted in W365 assignments<br />
  Based on [Issue 261](https://github.com/Micke-K/IntuneManagement/issues/261)<br />
- Added tooltip that variables are supported in the Export folder path<br />
  Based on [Discussions 269](https://github.com/Micke-K/IntuneManagement/discussions/269)<br />

- **Documentation**<br />
  - App Configuration (Device) documentation updated<br />
  Added support for value type for Android policies<br />
  Please continue discussion on the Issue below if this is still not working<br />
  Based on [Issue 231](https://github.com/Micke-K/IntuneManagement/issues/231)<br />
  This required some rewriting of the core documentation and an update to all output providers<br />
  This will make it easier to add additional tables to the documentation in the future<br />
  - Fixed issue with missing group name when exporting CSV<br />
  Based on [Issue 274](https://github.com/Micke-K/IntuneManagement/issues/274)<br />
  - Fixed issue with Authentication Strength when documenting Conditional Access policies<br />
  - Language files re-generated<br />
  - ObjectInfo files re-generated. Some Android updates<br />
  - ObjectCategory file re-generated<br />

- **Compare**<br />
  - Fixed issue with assignments on exported files when doing a Documentation compare<br />
  The group name was not resolved from migration table file<br />
  Based on [Issue 274](https://github.com/Micke-K/IntuneManagement/issues/274)<br />

- **Authentication**<br />
  - Added setting to allow Sort Tenant List<br />
  Based on [Issue 265](https://github.com/Micke-K/IntuneManagement/issues/265)<br />

  <br />

## 3.9.7 - 2024-06-27

**New features**

- **Compare**<br />
  - Added support for automation with batch job<br />
  - Added a new Compare provider - Intune Objects with Exported Files (Name)<br />
  This will support comparison exported policies between environments<br />
  - Added support for skipping missing source policies<br />
  - Added support for skipping missing destination policies<br />
  Based on [Issue 203](https://github.com/Micke-K/IntuneManagement/issues/203) and [Issue 128](https://github.com/Micke-K/IntuneManagement/issues/128)<br />

- **Compliance**<br />
  - Added support for Compliance v2 policies eg Linx policies<br />

**Fixes**
- **Compare**<br />
  - Renamed default provider to "Exported Files with Intune Objects (Id)" from "Intune Objects with Exported Files"<br />
- **Generic**<br />
  - Fixed issue with domain names with special characters in Profile info<br />
  Based on [Issue 237](https://github.com/Micke-K/IntuneManagement/issues/237)<br />
  - Lots of spelling and languag fixes  in documentation, script and UI<br />
  A huge thank you to **@ee61r1** for doing all this!<br />

- **Import/Export**<br />
  - Added support for exporting script for MacOS Custom attribute<br />
  Based on [Issue 244](https://github.com/Micke-K/IntuneManagement/issues/244)<br />

- **Documentation**<br />
  - App Configuration (Device) documentation updated<br />
  Initial support for Android<br />
  Please continue discussion on the Issue below if this is still not working<br />
  Based on [Issue 231](https://github.com/Micke-K/IntuneManagement/issues/231)<br />
  - Added support for documenting MacOS Custom attribute<br />
  Based on [Issue 244](https://github.com/Micke-K/IntuneManagement/issues/244)<br />
  - Fixed issed when documenting Shell script. Code was not included<br />
  - Language files re-generated<br />
  - AppTypes file re-generated. Some apps were not documented with proper name<br />

  <br />

## 3.9.6 - 2024-04-22

**BREAKING CHANGE**<br />
Microsoft are decommissioning the Intune PowerShell App with id d1ddf0e4-d672-4dae-b554-9d5bdfd93547, mentioned [here](https://learn.microsoft.com/en-us/mem/intune/fundamentals/whats-new#plan-for-change-update-your-powershell-scripts-with-a-microsoft-entra-id-registered-app-id-by-april-2024)<br />
This was the default app in IntuneManagement. The default app is now changed to Microsoft Graph PowerShell app with id 14d82eec-204b-4c2f-b7e8-296a70dab67e<br />
The script will automatically use that app for new installationsbr<br />
A warning to change will be displayed if d1ddf0e4-d672-4dae-b554-9d5bdfd93547 is used<br />
You can also register a new app, documented [here](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app) and then configure that app in Settings<br />
<br />
*Note*: This might require consent for the required permissions<br />
<br />
There is no change if you are currently using a custom app or already changed to Microsoft Graph PowerShell in Settings<br />
<br />
Also note that changing application will reset cached accounts<br />

**New features**

- **Compare**<br />
  - Added support for ignoring Basic properties and Assignments<br />
  Based on [Issue 203](https://github.com/Micke-K/IntuneManagement/issues/203) and [Issue 128](https://github.com/Micke-K/IntuneManagement/issues/128)<br />
  **NOTE:** Properties will be logged but with empty value for Match<br />

**Fixes**
- **Compare**<br />
  - Fixed issue when comparing Settings Catalog settings with child settings eg Hardened UNC Paths in Security Baseline<br />
- **Import/Export**<br />
  - Added support for import of MSIX app content<br />
  Based on [Discussion 191](https://github.com/Micke-K/IntuneManagement/discussions/191)<br />
  - Disable autoload of modules to prevent loading MSGraph module if found<br />
  Based on [Issue 208](https://github.com/Micke-K/IntuneManagement/issues/208)<br />

- **Documentation**<br />
  - Language files re-generated.<br />
  - AppTypes file re-generated. Some apps were not documented with proper name.<br />

  <br />

## 3.9.5 - 2024-01-20

**Fixes**
- **Import/Export**<br />
  - Assignments were not exported for some policies with trailing . in the name<br />
  Based on [Issue 184](https://github.com/Micke-K/IntuneManagement/issues/184)<br />
  **NOTE:** Policy will not export if full path is over 260 characters<br /> 
  - Fixed issue with policies not being exported when Batch was enabled in Settings<br />
  and there was only one policy for the specified object type<br />
  - Failed to get App Protection policies when Proxy was configured<br />
  - Fixed issue with importing policies with dependency in tenants with 100+ policies for a single policy type<br />
  Dependency only imported first page. All pages will be imported now to resolve dependencies<br /> 
  Based on [Issue 183](https://github.com/Micke-K/IntuneManagement/issues/183)<br />
 - Fixed issue with multiple export folders when using %DateTime% in path<br />
  Based on [Issue 189](https://github.com/Micke-K/IntuneManagement/issues/189)<br />
 
- **Get Assignment Filter usage**<br />
  - Filters not returned if only assigned to one policy<br />
  Based on [Issue 141](https://github.com/Micke-K/IntuneManagement/issues/141)<br />
  **NOTE:** Start the tool from: Views -> Intune Tools -> Intune Filter Usage<br /> 

- **Compare**<br />
  - Comparing Settings Catalog objects with exported objects failed<br />
  Issue cause by offline documentation was not working<br />
  Based on [Issue 183](https://github.com/Micke-K/IntuneManagement/issues/183)<br />

- **Documentation**<br />
  - Offline documentation of Settings Catalog was not working.<br />
  Values were always documented from online object<br />
  - Conditional Access documentation updates for Android and iOS<br />
  - App Protection documentation updates for Android and iOS<br />
  - Language files re-generated. Azure shou now be Entra for some documentations.<br />

  <br />

## 3.9.4 - 2023-12-18

**Fixes**
- **Get Assignment Filter usage**<br />
  - All policies that supports filter should now be collected<br />
  Please create an issue if not all expected filters are listed<br />
  Based on [Issue 141](https://github.com/Micke-K/IntuneManagement/issues/141)<br />
  **NOTE:** Start the tool from: Views -> Intune Tools -> Intune Filter Usage<br />

- **Documentation**<br />
  - Added support for documenting Conditional Access policies based on Workloads<br />
  Not 100% tested. Please report if not documented correctly<br />
  <br />

## 3.9.3  - 2023-12-11

**New features**

- **New tool - Get Assignment Filter usage**<br />
  - List all policies and assignments with a Filter defined<br />
  Based on [Issue 141](https://github.com/Micke-K/IntuneManagement/issues/141)<br />
  **NOTE:** Start the tool from: Views -> Intune Tools -> Intune Filter Usage<br />
  
- **Batch Export of App Content Encryption Key from Intunewin files**<br />
  This script can export encryption keys from existing intunewin files<br />
  Example:<br />
  Export-EncrytionKeys -RootFolder C:\Intune\Packages -ExportFolder C:\Intune\Download<br />
  This will export the encryption key information for each .intunewinfiles under C:\Intune\Packages<br />
  One json file will be created (for each .intunwinfile) in the C:\Intune\Download folder<br />
  File name will be **<*IntunewinFileBaseName*>_<*UnencryptedFileSize*>.json**<br />
  Do **NOT** rename this file since the script will search for that file when downloading or exporting App content<br />
  The script will not require authentication and it will have no knowledge of apps in Intune<br />
  Filename and unencrypted file size is used as the identifier to match app content in Intune with encryption file<br />
  **Important notes:**<br /> 
  Exported and decrypted .intunewin files are not supported to use for import at the moment.<br />
  These files are just the "zip" version of the source and can be unzipped with any zip extraction tool<br />
  The .intunewin file used for import has the "zip" version of the file and an xml with the encryption information +<br />
  additional file information eg. msi properties, file size etc.<br />
  Use the exported unencrypted "zip" version to restore the original files. Re-run the packaging tool if it should be re-used as applications content<br />
  <br />
  Please report any issues or create a discussion if there are any questions<br />
  Script is located: **<*RootFolder*>\Scripts\Export-EncrytionKeys.ps1**<br />

<br />

**Fixes**
- **Export**<br />
  - Fixed issue where Assignments were included in export even if 'Export Assignments' was unchecked<br />
  Based on [Issue 171](https://github.com/Micke-K/IntuneManagement/issues/171)<br />

- **Documentation**<br />
  - Fixed issue where filter was not documented on some policies<br />
  - Fixed issue with Word Output provider if a policy only had one settings<br />
  
- **Custom ADMX Files**<br />
  - Fixed bug with migrating custom policies between environments. Cache was not cleared when swapping tenants or imported additional ADMX files<br />
  - Fixed documentention issue with Administrative template policies in GCC environment. Name and Category was missing<br />
  Based on [Issue 174](https://github.com/Micke-K/IntuneManagement/issues/174)<br />
  - Custom ADMX based policies was missing properties when swapping tenant<br />
  Based on [Issue 124](https://github.com/Micke-K/IntuneManagement/issues/124)<br />

- **Generic**<br />
  - Fixed logging issues when processing objects with a group that was deleted. ID was not reported<br />
  - Generic Batch request function created to support other batch requests eg Groups<br />
  <br />

## 3.9.2  - 2023-10-17

**New features**

- **Application Content Export - Experimental**<br />
  - Added support for Exporting Appliction with decrypted content<br />
  App file can be downloaded during export or from the detail view of the Application<br />
  Enable "Save Encryption File" and specify "App download folder" in Settings<br />
  "App download folder" is used for encryption file and manual download<br />
  File content will be downloaded to the export foler during export<br />
  Files will be downloaded with .encrypted extension and then decrypted to original file name<br />
  Please report any issue or any suggestions<br />
  **NOTE:** This will ONLY work if the encryption file is exported and available<br />
  
- **Authentication**<br />
  - Login with application<br />
  This will login with specified Azure App ID and Secret/Certificate that is used for Batch processes<br />
  NOTE: This will require a restart of the app<br />
  Start with app **must** use -TenantID on command line. AppID and Secret/Certificate can be specified in Settings or command line<br />
  Example: Start-IntuneManagement.ps1 -tenantId \"&lt;TenantID&gt;\" -appid \"&lt;AppID&gt;\" -secret \"&lt;Secret&gt;\"<br />
  See *Start-WithApp.cmd* for samle file<br />
  Based on [Issue 122](https://github.com/Micke-K/IntuneManagement/issues/122) and [Issue 134](https://github.com/Micke-K/IntuneManagement/issues/134)<br />

- **Support for new Settings**<br />
  - Save encryption file - Saves a json file with encryption data when an application file is uploaded eg created or uploaded in details view<br />
  - App download folder - Folder where application files should be downloaded and decrypted<br />
  - Login with App in UI (Preview) - Use app batch login in UI<br />
  - Use Graph 1.0 (Not Recommended) - Use Graph v1.0 instead of Beta. **Note:** Some features will NOT work in v1.0<br />
  Based on [Issue 170](https://github.com/Micke-K/IntuneManagement/issues/170)<br />

**Fixes**
- **Documentation**<br />
  - Language files re-generated eg Supersedence (preview) -> Supersedence<br />
  - Added support for documenting "Filter for devices" info for Conditional Access policies<br />
  Based on [Issue 168](https://github.com/Micke-K/IntuneManagement/issues/168)<br />

- **Custom ADMX Files**<br />
  - Fixed issues with migrating custom policies between environments (3rd time)<br />
  Based on [Issue 124](https://github.com/Micke-K/IntuneManagement/issues/124)<br />
  - Fixed issue when importing ADMX files - Encoding issue eg ADMX/ADML file was UTF8<br />
  Based on [Issue 169](https://github.com/Micke-K/IntuneManagement/issues/169)<br />

- **Importing Windows LoB Apps**<br />
  - Fixed issue when importing LoB Apps that was only targeted to System context<br />
  Available Assignment option was missing after import<br />
  Based on [Discussion 164](https://github.com/Micke-K/IntuneManagement/discussions/164)<br />
  - Added support for Depnedency and Supersedence reations at import<br />
  Application will need to be re-exported since additinal data is added to the export file<br />
  Based on [Discussion 159](https://github.com/Micke-K/IntuneManagement/discussions/159)<br />

- **Generic**<br />
  - Fixed issue when compiling Procxy CS file<br />
  - Tls 1.2 is now enforced.<br />
  Based on [Discussion 166](https://github.com/Micke-K/IntuneManagement/discussions/166)<br />
  <br />

## 3.9.1  - 2023-08-30

**New features**

- **Added support for Windows Update Driver Policies**<br />

- **Support for new Settings**<br />
  - Proxy configuration - If configured, Proxy will be used for authentication, APIs and upload<br />
  - Disable Write-Error output - Skip PowerShell errors in output<br />

**Default Settings Value Changes**
  - Conditional Access policies will now be imported as Disabled by default<br />
  - New import option added: As Exported - Change On to Report-only<br />
  - This is to avoid being locked out from the tenant when importing Conditional Access policies<br />
  - Based on [Discussion 139](https://github.com/Micke-K/IntuneManagement/discussions/139)<br />

**Fixes**
- **Documentation**<br />
  - Fixed issues with some Feature Updates properties<br />
  - Added missing strings on Windows Update polices<br />
  - Regenerated Language files and Translation tables for Template policies<br />
  Note: Conditional Access string has changed file in background. Please report if there is anything missing<br />

- **Custom ADMX Files**<br />
  - Fixed issues with migrating custom policies between environments<br />
  - Case reopened due to something broke the initial functionality<br />
  - Only custom ADMX policies with #Definition properties can be imported into a new environment<br />
  - Based on [Issue 124](https://github.com/Micke-K/IntuneManagement/issues/124)<br />

- **Scope Tags**<br />
  - Fixed issues with importing policies with Scope Tags but they were not set<br />
  - Based on [Issue 133](https://github.com/Micke-K/IntuneManagement/issues/133)<br />

**Generic**<br />
  - Remove invalid characters from path.<br />
  - Based on [Issue 150](https://github.com/Micke-K/IntuneManagement/issues/150)<br />
  <br />

## 3.9.0  - 2023-05-04

**New features**

- **Added support for Authentication Context objects**<br />
  - These are used by Conditional Access policies<br />
  Based on [Issue 109](https://github.com/Micke-K/IntuneManagement/issues/109)<br />
  
- **Added support for Windows 365 Cloud PC settings**<br />
  - Based on [Issue 125](https://github.com/Micke-K/IntuneManagement/issues/125)<br />

- **Added support for Export/Import Tennant Settings**<br />
  - This is added the Intune Info view for now (Views -> Intune Info)<br />
  This means that there is no support for Bulk Import/Export. It must be done manually<br />
  This is to minimize the risk of re-importing Tenant settings<br />
  Based on [Discussion 131](https://github.com/Micke-K/IntuneManagement/discussions/131)<br />

**Fixes**
- **Documentation**<br />
  - Added full documentation of Requirement and Detection rules for Win Apps<br />
  Based on [Issue 119](https://github.com/Micke-K/IntuneManagement/issues/119)<br />
  - Fixed issue were documentation could crash if Reusable Settings policies exists<br />
  Based on [Issue 123](https://github.com/Micke-K/IntuneManagement/issues/123)<br />
  - Regenerated Language files and Translation tables for Template policies<br />
- **Intunwin File Upload**<br />
  - Fixed issue when uploading very large files<br />
  Based on [Issue 112](https://github.com/Micke-K/IntuneManagement/issues/112)<br />
  - Fixed issue when IE not installed<br />
- **Compare**<br />
  - Fixed issue where Compare could generate an exception in the log<br />
  Based on [Issue 128](https://github.com/Micke-K/IntuneManagement/issues/128)<br />
  **Note:** Issue 128 is only partially fixed. Compare needs a major update to fix the rest<br />
- **Import**<br />
  - Fixed an issue when creating Cloud groups based on on-prem groups without MigTable<br />
  - Fixed an issue when importing groups with a space in the beginning<br />
  **Note:** Inital spaces will be removed when importing groups<br />
  - Fixed issue when importing Endpoint Status Page polices with applications defined<br />
  - Fixed issue when importing Proactive Remediations (Health Scripts) with assignments<br />
  - Fixed issue when importing a Conditional Policy with Session propery disableResilienceDefaults set to $false<br />
  - Fixed issue when importing WiFi profiles. Support for multiple references was added eg multiple server verification certificates<br />
  Based on [Issue 114](https://github.com/Micke-K/IntuneManagement/issues/114)<br />
  - Terms of Use was not visible in the menu<br />
  **Note:** This might generate a Consent prompt if Use Default Permissions is not enabled<br />
  Additional permission required on the Azure App: Agreement.ReadWrite.All<br />
<br />

## 3.8.1  - 2023-01-26

**New features**

- **Added support for Reusable Settings objects**<br />
  - These are used by some of the Endpoint Security polices like Firewall rules<br />
  Based on private request<br />
  Note: No documentation support yet<br />

- **Added support for custom Authentication Strengths objects**<br />
  - These can be used in Conditional Access policies<br />
  Based on [Issue 109](https://github.com/Micke-K/IntuneManagement/issues/109)<br />
  Note: Not all issues in 109 are fixed yet and no documentation support yet<br />

- **Export/Import**<br />
  - PowerShell files for Health Scripts exported to the Export folder<br />
  - PowerShell files for Application Detection scripts are exported to the Export folder<br />
  Both scripts exports are based on [Issue 103](https://github.com/Micke-K/IntuneManagement/issues/103)<br />  

- **Documentation**<br />
  - Documentation engine completely rewritten for Settings Catalog and had major updates for other object types<br />
  Please create an issue if there are any problems<br />
  - Added support for HTML output<br />
  - MD output is now official with included support for CSS and single file Output.<br />
  Based on [Issue 35](https://github.com/Micke-K/IntuneManagement/issues/35)<br />
  - Added support for indent on sub-properties so it will be visible that a property is set based on a parent<br />
  Based on [Discussion 90](https://github.com/Micke-K/IntuneManagement/discussions/90)<br />  
  - Added option to skip assignments in the documentation<br />
  Based on [Issue 102](https://github.com/Micke-K/IntuneManagement/issues/102)<br />
  - Moved some Output options to generic output settings; Document scripts and Remove script signature<br />

- **Generic**<br />
  - Added new property on applications, InstallerType. This can be added as a new column to the View for Applications.<br />
  It specifies the New Microsoft Store App type; UWP or Win32<br />
  Based on [Issue 101](https://github.com/Micke-K/IntuneManagement/issues/101)<br />
  - Added response information f an API call failed. The log should now have a better description on why an API failed.<br />

  
**Fixes**
- **Documentation**<br />
  - Lots of documentation issues fixed by the new Documentation engine<br />
  - Sections and policies should now be in correct alphabetic order<br />
  Based on [Discussion 90](https://github.com/Micke-K/IntuneManagement/discussions/90)<br />  
  - Fixed issues with assignments for Setting Catalog issues<br />
  Based on [Issue 102](https://github.com/Micke-K/IntuneManagement/issues/102)<br />
  - Translation files re-generated<br />
  - Fixed error message: "Invoke-WordTranslateColumnHeader is not recognized as the name of a cmdlet"
  Based on [Issue 99](https://github.com/Micke-K/IntuneManagement/issues/99)<br />


- **Authentication**<br />
  - Fixed an issue when authentication to China Cloud<br />
  Based on [Issue 106](https://github.com/Micke-K/IntuneManagement/issues/106)<br />  
<br />

## 3.7.4  - 2022-11-17

**Fixes**

Lots of these issues are based on [Issue 94](https://github.com/Micke-K/IntuneManagement/issues/94)<br />
Thank you **Dominique** for all the amazing help with testing!<br />


- **Import/Export**<br />
  - Added support for Export of TermsOfUse PDF files<br />
  Based on [Issue 27](https://github.com/Micke-K/IntuneManagement/issues/27)<br />
  - Fixed an issue where it failed to import .intunewin files during bulk import<br />
  - Fixed issue with importing Edge app assignments<br />
  - Changed the order for Bulk delete to make sure policies are deleted in the correct order<br />
  - Lots of logging fixes for Bulk Export - Logged error when exporting object types not used<br />
  - Business Store Apps will not be delete - Not supported<br />
  - No import of assignments for default policies (Enrollment Status Page and Enrolment Restrictions)<br />
  - Lots of logging fixes for Bulk Delete - Errors if deleting default policies, trying to delete object types that were not used etc.<br />

- **Documentation**<br />
  - Added intent for Win32 Assignments<br />
  Based on [Issue 98](https://github.com/Micke-K/IntuneManagement/issues/98)<br />
  - Fixed an error when documenting Assignments for Win32 apps - Invoke-WordTranslateColumnHeader was missing/removed<br />
  Based on [Issue 99](https://github.com/Micke-K/IntuneManagement/issues/99)<br />
<br />


- **Logging**<br />
  - Added additional response error information if it failed to call a Graph API<br />
  - Missing groups will now only generate a warning instead of Graph API error<br />
  - No error for users without a profile photo<br />



<br />
<br />

## 3.7.3  - 2022-10-24

**Fixes**

- **Import**<br />
  - Fixed a bug where it failed to import Endpoint Security policies<br />
  - Fixed an issue where it failed to import Assignment Filters. A new property was added that is not supported during the import<br />

<br />
<br />

## 3.7.2  - 2022-10-08

**New features**

- **Added support for ADMX Files (Preview)**<br />
  - First version of supporting the ADMX file import<br />
  - Support for export/import policies based on ADMX files<br />
  The import/export between environments is very tricky so please report any issues<br />
  **Note**: The ADMX/ADML files must be copied to the app package folder or the policy exported folder<br />
  The ADMX files imported is based on last modify date. This will make sure files are imported in the correct order eg Mozilla and Firefox ADMX files<br />
  Based on [Issue 84](https://github.com/Micke-K/IntuneManagement/issues/84)<br />
- **Added support for value output type when documenting Administrative Templates**<br />
  - Select Output value in the Documentation form. _Value with label_ will add the label when documenting sub-properties<br />
- **Translate TenantID when migrating policies between environment**<br />
  - Any policy with a Tenant ID value will be translated when importing to a new environment<br />
  Based on [Discussion 83](https://github.com/Micke-K/IntuneManagement/discussions/83)<br />

**Fixes**

- **Authentication**<br />
  - Fixed an issue when auhencating with certificates during batch jobs<br />
  Fixed by @cstaubli. Thank you!<br />
  Based on [Issue 85](https://github.com/Micke-K/IntuneManagement/issues/85)<br />

- **Export\Import Fixes**<br />
  - Fixed an issue when importing Microsoft Apps files and the default document format was not set<br />
  Based on [Issue 92](https://github.com/Micke-K/IntuneManagement/issues/92)<br />

- **Documentation**<br />
  - Fixed the order of sub-properties when documenting Administrative Templates<br />
  - Fixed an issue where some xml values were not documented eg taskbar xml
  - Translation files re-generated<br />

<br />
<br />

## 3.7.1  - 2022-08-08

**Fixes**

- **UI**<br />
  - Fixed a bug where the menu bar was empty if not logged in<br />

## 3.7.0  - 2022-08-02

**Breaking changes**
  - A third header level was added when documenting to word<br />
  This level is used during bulk documentation and a group has more than one object type<br />
  Eg. The Conditional Access group documents Conditional Access, Named Locations and Terms of Use<br />
  The document will now have one section for each object type as third header level<br /><br />
  This could break documentation if a custom word template is used, and it does not have a third level header named 'Heading 3'<br />
  Specify the name of the 'Header 3 style' value in the Word settings before documenting<br />


**New features**

- **Support for tenant menu colors**<br />
  - Set colors and add tenant name to the menu bar<br />
  - Configure this in Tenant Settings and use this to distinguish lab from production environments<br />
  Based on [Issue 63](https://github.com/Micke-K/IntuneManagement/issues/63)<br />

- **Support for Compliance Scripts**<br />
  - Added support to Export, Import and Document **Compliance Scripts** profiles<br />
  - Compliance Script will now be included when documenting Compliance policy objects<br />
  Based on [Issue 60](https://github.com/Micke-K/IntuneManagement/issues/60)<br />
  
**Setting changes**
  - 'Allow update on import (Preview)' is removed<br />
  The 'Import type' is now always available<br />
  Note that Replace/Update are not fully verified yet<br />
  Based on [Issue 68](https://github.com/Micke-K/IntuneManagement/issues/68)<br />

**Fixes**

- **Export\Import Fixes**<br />
  - Target app groups was not set properly for App Protection policies during import<br />
  Based on [Issue 67](https://github.com/Micke-K/IntuneManagement/issues/67)<br />
  - Scope Tags were not assigned to objects during import<br />
  This happened in environment where Scope Tags already existed before import<br />
  Labels renamed to clarify that Scope Tags are assigned and not imported during import<br />
  Based on [Issue 61](https://github.com/Micke-K/IntuneManagement/issues/61)<br />
  - Default branding file had double dots in the exported file name [Issue 64](https://github.com/Micke-K/IntuneManagement/issues/64)<br />
  - Added API throttling during batch mode<br />

- **Documentation**<br />
  - Some properties were not documented for Endpoint Security objects<br />
  - Authentication context name added to Conditional Access
  - Translation files re-generated. This might add support for updated settings eg DFCI objects now uses separate category files<br />

<br />
<br />

## 3.6.0  - 2022-06-29

**New features**

**Silent batch job**<br />
  - Added support for silent batch documentation<br />
  Based on [Issue 39](https://github.com/Micke-K/IntuneManagement/issues/39)<br />

**Support for Co-management Settings**<br />
  - Added support for Export,Import and Document **Co-management Settings** profiles<br />

**Documentation**<br />
  - Re-generated language files and translation files<br />
  Some changes in Android profiles, iOS VPN and Windows Wired Network
  - Add support for documenting the following profiles - [Issue 57](https://github.com/Micke-K/IntuneManagement/issues/57)<br />
    -  Intune Roles<br />
    -  Custom Device Type Restrictions<br />
<br /> 

**Fixes**

- **UI Fixes**<br />
  - View did not show properties below 10 levels<br />
- **Silent batch job - [Issue 39](https://github.com/Micke-K/IntuneManagement/issues/39)**<br />
  - Unchecking default values was not working<br />
  - Failed to start without configuration file<br />
  - Failed to authenticate with certificate<br />
- **Documentation Fixes**<br />
  - Autopilot Profiles ([Issue 50](https://github.com/Micke-K/IntuneManagement/issues/50))<br />
  - Kiosk Template Profiles ([Issue 49](https://github.com/Micke-K/IntuneManagement/issues/49))<br />
  - Endpoint Protection Template Profiles ([Issue 51](https://github.com/Micke-K/IntuneManagement/issues/51))<br />
  - All User/All Devices Assignments ([Issue 54](https://github.com/Micke-K/IntuneManagement/issues/54))<br /> 
  - Scope Tags for Filters and PolicySets ([Issue 52](https://github.com/Micke-K/IntuneManagement/issues/52))<br /> 
  - Local User Group Membership - members were not listed
- **Export\Import Fixes**<br />
  - AssignmentFilters were not assigned during import (Twitter reported issue)<br />
  - Failed to assigned dependencies during import when dependency objects existed in the environment (Twitter reported issue)<br />
  - Failed to import/export lots of policies (Twitter reported issue)<br />
    429 - Too many requests. Graph API throttling kicked in.<br />
  - PowerShell script exported with wrong encoding ([Issue 48](https://github.com/Micke-K/IntuneManagement/issues/48))<br />

<br />

## 3.5.0  - 2022-04-26

**New features**

- **Automatic update check**<br />
  The app will check GitHub at start-up if there is a new version available<br />
  This can be disabled in settings<br />

- **Use PowerShell 5**<br />
  Command files will now use PowerShell 5 (-version 5 in the command line)<br />
  This is based on [Issue 44](https://github.com/Micke-K/IntuneManagement/issues/44)<br />

- **Documentation**<br />
  - New Word settings: Table text style and table caption location<br />
  This is based on an additional request in [Issue 37](https://github.com/Micke-K/IntuneManagement/issues/37)<br />
  - Terms of Use info when documenting Conditional Access<br />
  - Added documentation support for Terms of Use<br />
  - Added additional support for offline documentation<br />
  **Note:** Offline is defined as documenting an exported folder while logged in to another tenant.<br /> 
  If logged in to the same tenant as the exported folder, "online" documentation will be used<br />
  - Changed the layout for the assignment table on Win32 Applications. There were too many columns so additional info is changed to a table in the value column<br />
  - Filter / Filter Mode column headers are now set from language files<br />

<br />

- **Export/Import**<br />
  - Users in Conditional Access are now added to the Migration Table<br />
  This is so the user IDs can be translated during Offline documentation<br />
  - Referenced settings are now included in the export<br />
  This is to support referenced settings during import, copy and offline documentation (Certs on VPN profiles etc.)<br />
  These properties are named #CustomRef_*PropertyName* in the json file<br />
  **Note:** This might cause export/copy to take longer once every second week since it requires the MetaData XML for Graph to be downloaded.<br />
  This feature can be turned off by unchecking 'Resolve reference info' in Settings<br />
<br />

- **Copy**<br />
  - New dialog when copying an object. Description can now be changed during the copy<br />
<br />
- **Authentication**<br />
  - Full authentication support for US Government and China clouds<br />
  This requires that 'Show Azure AD login menu' is enabled in Settings<br />
  - Consent can be requested for missing permissions. This can be triggered via the 'Request Consent' link in the user profile info<br />
  - New version of MSAL.DLL, version 4.42.1<br />
  - Object types with only Read permissions are now supported. These will be orange in the menu<br />
  Buttons like Import and Delete will still be available but they will not work<br />
<br />

- **List objects**<br />
  - IsAssigned column is added to objects that supports it (property on the Graph object)<br />
  - Enable 'Expand assignments' in Settings to include Assignments when getting a full list of objects from Graph<br />
  This can be used for adding Custom columns based on assignment info<br />
  It is also used for setting the IsAssigned column for objects that doesn't have the info in Graph<br />
  This is based on [Issue 30](https://github.com/Micke-K/IntuneManagement/issues/30)<br />
  - Apps can be filtered in the request<br />
  If there are more than 1000 applications in the environment, the filter box can be used to return only matched items<br />
  Enter the filter in the text box and press the Refresh button. Clear the filter box and click Refresh to reload other objects<br />
  This is based on [Issue 28](https://github.com/Micke-K/IntuneManagement/issues/28)<br />
<br />

**Fixes**

- **Documentation**<br />

  - Fixed bug in *Conditional Access* documentation that caused some Grant information to be excluded from documentation
  - Fixed missing properties when documenting *Device restrictions (Windows 10 Team)* profiles 
  - Fixed some Offline Documentation issues<br />
  Get dependency info from exported folders instead of Graph<br />
  Offline documentation is not 100% fully supported yet. Dependency applications for Win32 apps are not included in this version<br />
  and there might be more properties missing. Please report anything missing for offline documentation to [Issue 37](https://github.com/Micke-K/IntuneManagement/issues/37)<br />
  **Note** Offline documentation will always require online access. Some information like language text, Azure roles, Mobile apps etc. will use Graph API<br /><br />

- **Authentication**<br />
  - First login with last used account could fail if the user domain was changed after the initial token was cached<br />
<br />

## 3.4.0  - 2022-03-01

**New features**

- **Silent batch job**
  Export/Import can now be executed without UI<br />
  See documentation for full requirements
  
  **Note** Please report any issues to [Issue 39](https://github.com/Micke-K/IntuneManagement/issues/39)
  
  This is based on [Issue 39](https://github.com/Micke-K/IntuneManagement/issues/39)
  
- **Documentation**

  - Support for documenting an environment based on exported files<br />

    Select the **Source files** folder in the Documentation Types (Bulk menu) dialog.
    
    Note: Some values will NOT be included. These are referenced values and not a property on the object eg Certificate on a VPN profile, Root certificate on a SCEP profile etc. These values will be documented with ##TBD...<br />

    This is based on [Issue 37](https://github.com/Micke-K/IntuneManagement/issues/37)

  - Support for attaching the json file for the object in the word document

  - Support for documentation output level (Word)<br />
    Documenting the full environment can create a document with 1000+ pages depending on the amount of profiles and policies. The documentation output level can now be used to reduce the document size. The output level options are:

    - Full - Document every single value
    - Limited - Set max value and truncate size for documentation and as option, attach the original value as a text document to the value cell e.g. truncate all values over 500 characters to 10 characters and attach the full value as a text file in the document. This will reduce documentation size for profiles with large XML strings like ADMX ingestion
    - Basic - Only include the Basic and Assignments tables in the documentation 

  - Added support for documenting Filters

- Added UI for configuring custom columns<br />
  This can now be done in the Detail View<br />
  This is based on [Issue 30](https://github.com/Micke-K/IntuneManagement/issues/30)<br />

-  Added support for updating Name and Descriptions of the object in Detail View<br />
  This is As-Is functionality. Not all object types have been tested.<br />
  It is recommended to use the portal for this.<br />
  This is based on a private request<br />

- Added support for copying an app

  **Note** This requires that the **App packages folder** is specified in Settings and that the file for the app is available in that folder. If the app file is missing it can be uploaded manually in the Details view 

  This is based on [Issue 42](https://github.com/Micke-K/IntuneManagement/issues/42)

- Added support for manually upload an app file via the Details view 

**Fixes**

- **Documentation**<br />

  - Updated documentation files with support for new properties and removed unused values (Windows Updates, Windows Feature Updates etc.)
  - Fixed an issue where VPN profiles in some cases was missing the Base VPN settings
  - Fixed an issue when using a template<br />
  A table of content will no longer be created. That should be included in the template

- **Application import**

  - Minor change in the app Win32 upload functionality to align to portal APIs
  - The File Name is now updated to be based on the actual uploaded file 

    **Important** Please create an issue if there are any problems

- Fixed an issue where ESP and Enrolment Restriction objects were not listed
  The original filters stopped working 

  **Note** The Enrolment restrictions has changed in Graph. There is now one object for each OS type. So there will be multiple restriction objects exported. platformType column was added to identify each object   

  This is based on [Issue 41](https://github.com/Micke-K/IntuneManagement/issues/41) 

- Minor fixes in Import/Export extensions - Required for silent batch job support

- Fixed an issue where PostListCommand was not triggered

    - Additional Endpoint Security columns were not listed
    - Azure Branding objects was missing the language column <br />
<br />
- Fixed issue where the Document button was not enabled when **Select All** was clicked (without selecting an object first)

  This is based on [Issue 36](https://github.com/Micke-K/IntuneManagement/issues/36) 

- Other minor bug fixes to support the new features

## 3.3.3  - 2021-12-15

**Fixes**

- Fixed issue where displayName was missing in object list<br />
  Thank you Jason!

## 3.3.2  - 2021-12-14

**New features**

- Markdown support for documentation (Experimental)
  This will create a MD document in the Documents folder.
  **Note:** This is not working 100% at the moment. The script will create a MD document but it might be too large if all objects in the environment are documented. 

  Also note that HTML tables are used so that code can be documented as code blocks. This must be supported by the MD Viewer. The *Markdown Viewer* extension in Chrome was used during testing.

  Please report any suggestions to the issue.<br />
  This is based on [Issue 35](https://github.com/Micke-K/IntuneManagement/issues/35)

- Added support for batched export
  This will use batch API to request full info for up to 20 objects per batch to reduce export time
  This can be enabled in setting

- Added support for scrolling cached users and guest accounts in the profile info<br />
  This can be enabled in settings

- Added support for sorting cached users<br />
  This can be enabled in settings

**Fixes**

- Paged return of objects<br />
  Only first page of objects will be loaded by default. 
  Additional pages can be loaded with **Load More** or all available objects can be loaded with **Load All**.<br />
  This is based on [Issue 28](https://github.com/Micke-K/IntuneManagement/issues/28)

- Fixed an issue where a checkbox had to be clicked twice to be checked when the list was filtered <br />
  This is based on a known issue 

- Fixed an issue where buttons were not enabled when **Select All** was checked<br />
  This is based on [Issue 36](https://github.com/Micke-K/IntuneManagement/issues/36)

- Fixed an issue when adding object ID to the file name during export
  The separate settings file was not exported with the ID in the name which could cause issues during import

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

* **Use Default Permissions** is now set to Enabled by default. With the Tenant Specific Settings feature, this can now be enabled globally or per tenant. Consultants accessing multiple environments might not have permissions to grant consent requests so this could be enabled on a global level and then disabled for tenants where the permissions can be added.  

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
