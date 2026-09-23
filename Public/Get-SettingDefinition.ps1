<#
.SYNOPSIS
    List the settings the module knows about - keys, types, defaults and where they
    are stored.

.DESCRIPTION
    Exported as Get-IMSettingDefinition.

    Discovery for the settings API. Automation should not have to read the source or
    the settings dialog to find out that the export folder key is called
    "ExportFolder", that it is a Folder setting, or that it is stored under
    "IntuneManagerExportSettings".

    Only registered settings are listed - the ones the settings dialog shows. A few
    keys are deliberately unregistered (per-feature state, and hidden keys such as
    ExportReplaceTokens); those need an explicit -SubPath when read or written.

.PARAMETER Key
    Return one setting. Wildcards are supported.

.PARAMETER Section
    Return only the settings in one section (General, ImportExport, IntuneManager,
    ...). Wildcards are supported.

.EXAMPLE
    Get-IMSettingDefinition | Format-Table Key, Section, Type, DefaultValue

.EXAMPLE
    Get-IMSettingDefinition -Key *Export*

.EXAMPLE
    Get-IMSettingDefinition -Section General | Select-Object Key, Title, Description

.LINK
    Get-IMSetting
.LINK
    Set-IMSetting
#>
function Get-SettingDefinition
{
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]$Key,
        [string]$Section
    )

    foreach($settingSection in (Get-SettingsSections | Sort-Object Order, Title))
    {
        if($Section -and $settingSection.Id -notlike $Section -and $settingSection.Title -notlike $Section) { continue }

        foreach($definition in $settingSection.Values)
        {
            if($Key -and $definition.Key -notlike $Key) { continue }

            [PSCustomObject]@{
                Key          = $definition.Key
                Title        = $definition.Title
                Section      = $settingSection.Id
                SectionTitle = $settingSection.Title
                Type         = $definition.Type
                DefaultValue = $definition.DefaultValue
                SubPath      = $definition.SubPath
                Description  = $definition.Description
                # The allowed values for a list setting, so a caller can validate a
                # value before writing it.
                ItemsSource  = $definition.ItemsSource
            }
        }
    }
}
