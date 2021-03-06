# Release Notes

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
