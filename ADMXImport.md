# ADMX Ingestion

The script can create Custom Profiles based on ADMX ingestion. ADMX ingestion is a way to support existing ADMX files in an MDM environment. Windows uses the [Policy CSP](https://docs.microsoft.com/en-us/windows/client-management/mdm/policy-configuration-service-provider) to apply the configuration. 

Microsoft links:

* ADMX Ingestion documentation can be found [here](https://docs.microsoft.com/en-us/windows/client-management/mdm/understanding-admx-backed-policies) and [here](https://docs.microsoft.com/en-us/windows/client-management/mdm/enable-admx-backed-policies-in-mdm)
* Additional information including blocked registry keys can be found [here](https://docs.microsoft.com/en-us/windows/client-management/mdm/win32-and-centennial-app-policy-configuration)  
* Shema definition documentation can be found [here](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-gpreg/6e10478a-e9e6-4fdc-a1f6-bdd9bd7f2209)
* Old ADMX schema documentation (Vista) including attribute information can be found [here](http://download.microsoft.com/download/5/0/8/5081217f-4a2a-470e-a7fa-5976e40b0839/Group%20Policy%20ADMX%20Syntax%20Reference%20Guide.doc)
* Another schema documentation including attribute information can be found [here](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/policy/admx-schema)

## ADMX Import

The **ADMX Import** tool is used for configuring 3rd party applications e.g. Chrome, Google Update etc. These ADMX files are available from the software vendor. An ADMX can be loaded in the tool and all settings can be configured using a similar UI as GPMC. When the ADMX is loaded, the script will look for an ADML file that is either in the same directory or in the en-US subdirectory. An ADML file can also be loaded manually, if another language should be used in the UI.

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/ADMXExample.png" height="50%" width="50%"/>

The image above shows the the tool after the chrome.admx file was loaded. The tool supports delivering ADMX settings to HKLM (Computer Settings) and HKCU (User Settings). Categories will be added based on the Class attribute for each ADMX policy setting. The *All settings* category will display all the settings for the base category (Computer or User)

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/ADMXSettingExample.png" height="50%" width="50%"/>

A policy setting can either be edited via double-clicking an item or right-clicking and select Edit. 

The *Intune OMA-URI name* property specifies the name of the OMA-URI row in the Custom Profile. This is optional and if it is not specified, the script will use the name of the policy.

A policy must be set to Enabled before any changes can be made. The *Policy* tab will list all possible settings for the policy. This could be a dropdown box, text box, check box, numeric up-down box etc. The script creates the controls based on the presentation settings in the ADML file. An ADML is not mandatory but the controls and the UI could cause unpredictable results. Always use the associated ADML for correctly generated controls.

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/ADMXSettingExampleOMAURI.png" height="50%" width="50%"/>

The *OMA-URI Settings* tab contains the configured settings. This is the string that will be added to the enabled policy. This can be manually configured in case there is something that is not supported by the script. Do **NOT** add <enabled /> or <disabled /> to this text box. The script will add that automatically.  If *Manual configuration* is checked, the script will upload the text as it is specified, including additional manual changes. If it is not checked, the script will generate the text when importing the profile. If manual configuration is added and then checkbox is cleared, those changes will be lost during the upload.

The *XML Definition* tab contains the XML node for the ADMX policy. This is used for reference in case manual configuration is required.

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/ADMXImportSetting.png" height="50%" width="50%"/>

The *Import* tab  is used for creating the Custom Profile in Intune. The *Custom Profile Name* is mandatory and specifies the name of the profile.  

The *ADMX Policy File Name* is mandatory and this should be a globally unique name. This will generate an ADMX file on the client. See the Deep Dive section for more information.

The *ADMX App Id* is mandatory and this does not have to be unique but it is recommended in some circumstances e.g. if multiple versions of the Chrome ADMX file is uploaded, it is recommended to add the version to the *ADMX Policy File Name* and *ADMX App Id* e.g. Chrome91. See the Deep Dive section for more information.

The *Ingest ADMX file* is checked by default and this will included the ADMX file ingestion in the Custom Profile. If there will be multiple Custom Profiles based on the same ADMX file, it might be better to have one Custom Profile for the ADMX ingestion and one separate Custom Profile for each of the settings. This requires that the same *ADMX App Id* is used for each Custom Profile that is based in the ingested ADMX file. 

 The *OMA-URI Name  for the ADMX ingestion* specifies the name of the OMA-URI row inside the Custom Profile. This is optional. It will be set to a value based on the loaded file name by default e.g. chrome.admx Ingestion.

The *Import* button will create the Custom Profile in Intune. There is no visual information if the profile was created successfully but the log will display the name and id of the created profile. A message box will be displayed if it fails to create the Custom Profile. 

## Reg Values

The **Reg Values** tool can be used to create registry values in HKLM and HKCU. This uses the same functionality as the ADMX Import tool; ADMX ingestion. The difference is that the Reg Values tool builds the ADMX file in the background based on the added registry values. There are some benefits of using this over a PowerShell script e.g. Intune will state if the registry keys were applied successfully and if a conflict or an error occurred.  

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/ADMXRegValuePolicy.png" height="50%" width="50%"/>

The initial screen include the options of specifying the Custom Profile name and description. The *Policy type* can either be Policy or Preferences. See Known Issues and Limitations for more information about policy types.

The *Added reg values* list contains the ADMX policies. Each ADMX policy can contain one or more registry values.

The top part of the *Add new reg policy* form specifies the attributes on the policy node in the ADMX file. The *Policy name* identifies the policy. This cannot contain any spaces. The *Policy status* property specifies if the policy should be enabled or disabled. The hive and the key properties specifies where the registry values should be added. The *Reg key* property is a global value for all added registry values in the bottom section.

The *Policy value* is an optional value. This should only be used if the registry policy should add a value that specifies if it is enabled/disable e.g. 1 or 0. This will use the enabledValue and disabledValue nodes in the background. If this value is specified, a registry value (DWORD) will be set to 1 when the policy is enabled or 0 if the policy is disabled.

The lower part of the form specifies individual registry values. The tool support creating/setting the following type of registry values:

* String
* Expanded string (String with Expanded checked)
* Multi-string
* DWORD
* List - a key/value string list. Each key will be a string value.

The *Key* property is not required unless the value is located in a different location than specified in the *Reg key* property. This is used when specifying values for the List type. The List type values will then be added to a separate key.

**Note:** The List type does not support specifying the value name since all values are creating in a separate key.

*Value name* and *Value* properties specifies the values that should be created in the registry. 

Modify existing values by double-clicking on the value in the *Added reg values* list.

The *Additional value settings* section has additional settings for a policy type. This can be used to create a REG_EXPAND_SZ instead of a REG_SZ etc.

**Note:** The *Do not overwrite value* will set the soft attribute. This does **NOT** work, at least not outside the Software\Policies area on a cloud only joined device. This property is kept until more testing can confirm that it doesn't work at all.

Example of setting registry values for a device and a user: 

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegPolicyHKCU.png" height="50%" width="50%"/>

Example of adding HKCU settings

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegPolicyHKLM.png" height="50%" width="50%"/>

Example of adding HKLM settings

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegADMXFileContent.png" height="50%" width="50%"/>

Example of generated ADMX file

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegValueIntuneProfile.png" height="50%" width="50%"/>

Example of the created Custom Profile created. 

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegProfileADMXIngestion.png" height="50%" width="50%"/>

Example of the OMA-URI row for ADMX ingestion for a custom registry value.

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegProfileOMAURISetting.png" height="50%" width="50%"/>

Example of the OMA-URI row for specifying the registry values to set.

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegValuesHKCU.png" height="50%" width="50%"/>

Example of the implemented HKCU settings for a user

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegValuesHKLM.png" height="50%" width="50%"/>

Example of the implemented HKLM settings for a device

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegValuesHKLMList.png" height="50%" width="50%"/>

Example of the List type implementation on a device 

## Deep Dive

The [Policy CSP](https://docs.microsoft.com/en-us/windows/client-management/mdm/policy-configuration-service-provider) is used when ingesting ADMX files. When a Custom Profile with ADMX ingestion is assigned to a device, the Policy CSP will add information about the ADMX file in the registry and create an ADMX file in the file system.

**ADMX ingestion:**

An ADMX is ingested by using the following OMA-URI path:

./device/Vendor/MSFT/Policy/ConfigOperations/ADMXInstall/<AppID>/[Policy|Preference]/<ADMXFileName>

Example:

./device/Vendor/MSFT/Policy/ConfigOperations/ADMXInstall/Chrome/Policy/ChromeAdmx

This will generate an ADMX file, ChromeAdmx.admx, in the following folder on the device:

%ProgramData%\Microsoft\PolicyManager\ADMXIngestion\\<ProviderGuid>\\<AppID>\\[Policy|Preference]

The Reg Values tool will always use IntuneManagementReg as AppId. Each uploaded reg policy will have a unique named ADMX file, RegPolicy_<GUID>. 

The Policy CSP will create multiple registry values under HKLM\Software\Microsoft\PolicyManager

The ADMXDefault part will contain each Category, with full path, specified in the ADMX file. This is why the AppID must be unique when using multiple versions of the same admx file since each of these files will have the same category IDs e.g. each Chrome version should be named based on the version like Chromev91.

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegADMXDefault.png" height="50%" width="50%"/>

**Note:** If the same ADMX filename is used, the first file saved will win. It looks like Intune will not overwrite an existing ADMX file. That is why a unique name must be specified when different versions of the ADMX file is used.

There is an additional registry key added for the ADMX ingestion. Each ingested file will generate a key under AdmxInstalled.

AdmxInstalled\<ProviderGuid>\\<AppID>\\[Policy|Preference]\\<ADMXFileName>

Example:

\SOFTWARE\Microsoft\PolicyManager\AdmxInstalled\D12FCE57-F71E-4D0D-93EE-35C5E6F8C0D9\Chrome\Policy\ChromeAdmx

This key contains information when it was added, status and how many policies it has.

**Policy Settings**

Each setting is added based on the following OMA-URI path:

./[User|Device]/vendor/msft/policy/config/<AppID>~[Policy|Preference]~<CategoryPath>/<PolicyName>

Example:

./Device/Vendor/MSFT/Policy/Config/Chrome~Policy~googlechrome~Startup/ShowHomeButton

The CategoryPath is the full path to the category where the setting is defined. Each categoryId is separated with a ~. This should match the registry value specified in the image above in the ADMXDefault key. 

One registry key for each ADMX policy is created under the Provider path (PolicyManager\Provider\\<GUID>) . This includes the OME-URI settings configured in the Custom Profile. 

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegADMXProvider.png" height="50%" width="50%"/>

Applied settings are added to the Current registry key, Current\Device or Current\\<SID> depending if it is for the device or the user. There is one key for each policy category.

<img src="https://raw.githubusercontent.com/Micke-K/IntuneManagement/master/RegADMXCurrent.png" height="50%" width="50%"/>

**Troubleshooting**

The Event Viewer can be used for troubleshooting any issues with ADMX ingestion. The events are added to the following log:

Application and Services log\Microsoft\Windows\DeviceManagement-Enterprise-Diagnostics-Provider\Admin

The following events can be used for troubleshooting:

* 819 - Information: The policy settings was successfully deleted 
* 831 - Information: The policy settings was successfully added 
* 866 - Information: Update policy (This is followed after a 872 or 873 event)
* 872 - Information: Start updating existing ADMX ingestion
* 873 - Information: Starting new ADMX ingestion
* 865 - Error  Catastrophic Failure. This could be that the ADMX is invalid. 
* 404 - Error: This is generated for different reason
  * Generated after a 865 Catastrophic Failure error
  * The system cannot find the file. The CSP cannot file the specified ADMX file. This could happen when ADMX and policies are in separate Custom Profiles and the wrong AppID was specified in the OMA+URI paths or if the policy settings are applied before the ADMX ingestion.
* 454 - Error: This is listed when the Custom Profile is removed and the CSP cannot delete registry values outside Software\Policies.  

## Known Issues and Limitations

* The created ADMX ingestion profiles has only been tested on cloud only joined devices.  
* According to [this](https://docs.microsoft.com/en-us/windows/client-management/mdm/win32-and-centennial-app-policy-configuration) link, policies will NOT be enforced unless the device is domain joined. So only Hybrid devices would support enforced values. This means that all settings on cloud only joined devices will be set as Preference values e.g. set once and never updated.

**ADMX Import:**

* Only categories and policy names specified in the loaded ADMX/ADML file will be translated. If the ADMX uses strings outside the loaded ADML file, it might be blank or using the string id.  

**Reg Values**

* The Preferences type is supported by the tool but tests shows that there is no difference in functionality compared to the Policy type on cloud only joined devices.  
* The script will block some registry keys. These keys are blocked by Microsoft and a PowerShell script is required to write to these values. See [this](https://docs.microsoft.com/en-us/windows/client-management/mdm/win32-and-centennial-app-policy-configuration) link for more information.
* Values outside Software\Policies will **NOT** be deleted when the policy is removed.
* The tool supports all ADMX attributes specified in the schema but it looks like some functionalities are not supported by Windows or the Policy CSP e.g. the *soft* attribute should be set to true to avoid overwriting an existing value but all values were overwritten during the tests, even if the soft attribute was set. 
* QWORD is not supported. The ADMX schema definition includes longDecimal which would create QWORD values but this is not supported in the Policy CSP. It will generate a Catastrophic Failure event in the Event Log. 
* No support for enabledList/disabledList. This might be added in the future since this could make it very easy to create mapped drives via ADMX ingestion. 
