@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'IntuneRemediation.psm1'
    
    # Version number of this module.
    ModuleVersion = '0.1.0'
    
    # ID used to uniquely identify this module
    GUID = '8e4f5a36-6ea0-4fa4-9f8c-5f65429d2419'
    
    # Author of this module
    Author = 'Intune Administrator'
    
    # Company or vendor of this module
    CompanyName = 'Your Company'
    
    # Copyright statement for this module
    Copyright = '(c) 2023 Your Company. All rights reserved.'
    
    # Description of the functionality provided by this module
    Description = 'PowerShell module for creating and managing Microsoft Intune remediation scripts'
    
    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '5.1'
    
    # Modules that must be imported into the global environment prior to importing this module
    # We're using dynamic module loading, so we don't require these upfront
    # RequiredModules = @(
    #     @{
    #         ModuleName = 'Microsoft.Graph.Authentication'
    #         ModuleVersion = '1.0.0'
    #         Guid = '883916f2-d041-46f9-b428-6a7916a6e26c'
    #     },
    #     @{
    #         ModuleName = 'Microsoft.Graph.DeviceManagement'
    #         ModuleVersion = '1.0.0'
    #         Guid = '60f7c2fb-646a-47bd-9cb3-90af0c32284f'
    #     }
    # )
    
    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport = @(
        'New-IntuneRemediationScript',
        'Connect-IntuneWithToken',
        'Test-IntuneRemediationScript',
        'Initialize-IntuneConnection'
    )
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData = @{
        PSData = @{
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags = @('Intune', 'Remediation', 'MDM')
            
            # A URL to the license for this module.
            LicenseUri = 'https://github.com/YourRepo/IntuneRemediation/blob/main/LICENSE'
            
            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/YourRepo/IntuneRemediation'
            
            # ReleaseNotes of this module
            ReleaseNotes = 'Initial release of IntuneRemediation module'
        }
    }
} 