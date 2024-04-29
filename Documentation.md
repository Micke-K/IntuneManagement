# Documentation

The script can document most of the profiles and policies in Intune. The output can either be CSV or Word. Documenting to Word can either be to an existing Word template or empty document. There are many options for the documentation e.g. Language, Header styles, Table styles, managing not configured items etc.  

The idea behind the documentation method in the script is to output the information as close to the Endpoint Manager portal as possible. Some of the objects has a different property name or value in the Summary text vs Edit mode in the portal. The documentation will then use the Edit mode information when possible. Some policies and profiles might have a slightly different order than the portal.

The objects can be documented in three ways:

* One or more objects of a single object type

  Document one or more objects by selecting them and click the Documentation button 

* Selected objects of multiple object types

  Add objects via the Documentation button, Add to list and then document all selected objects via the Bulk menu

* All objects

  Document all supported objects. Initiated in the Bulk menu 

Documentation is a very complex process. There are multiple types of objects, different languages, different types of properties etc. The best output is often based on a personal opinion. Some parts can be configured but not everything can be personalized. The output of most properties is simple, there is a name and a value. This is not the case for some properties e.g. a firewall rule. The firewall rule is itself a table with lots of possible values. The documentation will add the multi-property values with a property separator, comma is the default setting. There are also properties that contains multiple values. These will be added with an object separator, new line as default. The separators can be changed e.g. it might be better to use | as a object separator when documenting to a CSV.  

**Note:** The word document might need some manual post updates. Tables are auto generated but they might have to be tweaked for personal preferences.      

Please read the [Deep Dive](#deep-dive) section below for a detailed description of the documentation process.

**Language Support**

The script can document the objects in any language supported by Intune. 

Note that some profiles and properties do not have language support. These will be documented in English.

**Known Issues**

This is the first version of the documentation support.  

* There are over 100 different object types. Not all are supported and not all supported object types are tested and verified. 

* Property based profiles might be missing some properties. Some properties in the translation files are generated at runtime. Support for these properties are added in the custom documentation provider.

* Some complex type properties are not translated in this version e.g. the screen layout of apps on the iPhone. 

* The script will generate warning that properties are missing. This could be caused by multiple reasons

  * The property must be added to the custom documentation provider

    Note that one missing property can cause other properties from not being documented.

  * The property does not exist until it is configured e.g. privacy settings in the Win 10 Device Restrictions profile.

* Property based profiles uses generated translation and static language files. These files needs to be generated and uploaded when new functionality is released in Intune

* Some Endpoint Security polices are NOT deviceManagementIntent objects. They are actually Settings Catalog objects and they will be listed with the Settings Catalog items e.g. *Antivirus - Windows 10 and Windows Server (ConfigMgr)* policies

* Some Endpoint Security/Settings Catalog items are not translated based on Graph API in the portal e.g. *Antivirus - Windows 10 and Windows Server (ConfigMgr)* policies.  These will be documented based on Graph API information which might be different compared to the portal

* Markdown is currently in experimental state. The script can document to an MD file created in the Documents folder but this can be to large in environments with many objects. The script will create HTML tables to support code blocks and column span. The MD View must support HTML tables to display the document. The *Markdown Viewer* extension in Chrome was used during testing.

Please create an [Issue](https://github.com/Micke-K/IntuneManagement/issues) if properties are documented incorrectly or missing. 

# Deep Dive

The documentation is based on a two step process

* Gather information about the object

  This will collection all the information about the object and add it to a PowerShell object in the code

* Send the information to the selected output provider

  This will send all the information gathered about the object to the selected output provider

These steps are then repeated for each object that is being documented.

An output provider then has an initial (PreProcess) and finish (PostProcess) step e.g. the Word provider will create a word document in the initial task, document all properties and then update content tables, word properties and save the document in the finish task.

**Object Types**

Intune has multiple object types, over 100 different profile/policy types. These could be objects based on a static set of properties e.g. Configuration Profiles, Compliance Policies etc. or it could be settings based objects like Settings Catalog, Endpoint Security policies and Administrative Templates. The settings based policies only store information about the configured settings and not all settings available for the object. 

**Settings Based Objects**

All the documentation for Settings objects is done via Microsoft Graph. The Graph APIs contains all the information about property types, enum and language strings etc.  

Language strings for column headers and basic information are documented from the static language files

**Property Based Objects**

Property based objects are very complex to document. These objects can be documented in different ways;

* Generated json files
* Manually created json files
* PowerShell function
* Or a combination of PowerShell and manually created json files 

Json files for translating property objects to documentation is located in the Documentation\ObjectInfo folder. All files that start with a # are manually created. These can either be based on the object type (@OData.Type) or the Object Type Id (specified in the EndpointManager.psm1 file). Files that does not start with a # are automatically generated and the `ObjectCategories.json` file contains the mapping between the Intune object and associated json files. One object type can be associated with multiple files. Each file represents one category of the object.

The generated files sometimes requires additional manually created properties. These could be properties in the UI that has a Yes/No, Enabled/Not Configured etc. trigger associated with a specific value. These properties must be manually added to the object before the documentation. The `DocumentationCustom.psm1` file takes care of this. This file is also used for overriding the documentation of specific values and other custom required processing.    

The json files contains a definition of each property to document. This includes information like type e.g. Boolean (Yes/No, Allow/Block, Enabled/Disabled etc.), Options, DataTable etc. The script will use these files to translate each property into a PSCustomObject that is then used by the output provider. The functionality of these files has been extended to enhance the documentation options for the manually created files. Data types of 100 or above is custom functionality. All data types below 100 is based on the same functionality as in the Intune portal. The data type engine in the script is created based on best effort of the generated json files. 

The `DocumentationCustom.psm1` file also takes care of custom documentation for some object types e.g. Conditional Access. App Configuration policies etc. These objects are documented via a PowerShell function in the script.

**Language Support**

The Settings based objects get their language strings from Graph APIs with a few exceptions.

The property based objects uses static language files. Each language file is generated based on multiple language files. Not all the language information is included in these files to reduce the output size. This could reduce the possibility to create custom documentation with multi language support.

These files will have to be re-generated when new functionality is released in Intune.

**Scripts for Generated Files**

The scripts that automatically generate language files, translation files, object info etc. are not included in the release. These scripts are currently not in a state that they can be released. The best would be if Microsoft released all the required information in Graph. A deep dive into graph suggests that it might be possible in the future since some information about the generated files are there but with some properties missing or language text missing. The information can't be accessed unless an API is called that gets the definition for all the profiles at the same time (the file is over 100MB). 

## Extending The Documentation

The documentation can be extended in multiple ways.

* Documentation Provider
* Documentation Output Provider

The two methods can be used to customize the documentation in every possible way e.g. an Excel provider can easily be created to support excel as an output type,  a PowerShell module can be created to translate any unsupported property.

The priority order for object documentation is:

* Script function in the documentation provider module
* Json file based on OData.Type
* Json file based on Object Type
* Generated json files 
* Settings Objects

**Documentation Provider**

The documentation provider takes care of collecting all the information about the object. The `DocumentationCustom.psm1` file is an example of this. This file has examples of custom translation of properties for json files and examples of custom translation of objects via PowerShell functions.

Documentation providers have a Priority property. This defines in what order the providers will be triggered. The provider with the lowest priority number will be executed first. The included custom documentation provider has a priority number of 1000. The information gathering of the provider can be overridden by creating a custom documentation provider with a lower priority number.

**Documentation Output Provider**

Once the script has finished gathering all the data of an object, it sends it to the Documentation Output provider.  This will then document it to the specific output type of the provider. Word and CSV are included. There is also a "None" provider included. This will only add the gathered information in the UI. This is used for quick information or when building the translation files. This is also used by the comparison functionality.

The `DocumentationWord.psm1` and `DocumentationWordOptions.xaml` files are example on how to create an output provider. The xaml file contains the configuration options in the UI. The psm1 file registers the provider and builds the Word document based on the gathered information.

**Translation Files**

Translation files can be created for objects that don't support documentation at the moment. Some of them would require a combination of a json file and adding additional properties via PowerShell. The files are read when the documentation is triggered so they can be updated and tested without restarting the script.

Do **NOT** update the generated files. These are automatically generated and will be reversed in every release.  

