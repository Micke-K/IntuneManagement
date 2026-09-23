# Shared UI settings registration - single source for settings that are consumed
# ONLY by UI code but apply to BOTH backends (WPF + Avalonia).
#
# Why here: these keys used to be registered twice (once per backend, kept in sync by
# hand) or in engine files despite having no engine consumer. UI/Classes/ loads for
# every backend (and headless - harmless: the Settings dialog never renders there,
# and registration keeps Get-SettingValue resolving the right SubPath everywhere).
#
# Why AppInitialized: UI/Classes/ dot-sources BEFORE Internal/, so the non-General
# sections (IntuneManager / GraphGeneral / IntuneTools / MSAL) do not exist yet at
# load time. All sections exist by AppInitialized, and both Settings dialogs read
# the registry at render time, long after that. (Same pattern as the IntuneManagerUI
# extensions' registrations.)
#
# Rule of thumb: a setting consumed by Internal/Classes/Public code must be
# registered in an engine file instead (see GraphPageSize in Internal/IntuneManager.ps1
# for the cautionary tale) - this file is only for UI-consumed settings.

function Add-CommonUISettings
{
    # ---- General section (UI appearance / behavior) ----

    # Default is not a palette of its own - it resolves to Light or Dark from the
    # Windows app theme at apply time (see UI/Classes/UIThemeCommon.ps1).
    $themeList = @(
        [PSCustomObject]@{ Name = "Default (follow Windows)"; Value = "Default" },
        [PSCustomObject]@{ Name = "Light";                    Value = "Light"   },
        [PSCustomObject]@{ Name = "Dark";                     Value = "Dark"    }
    )

    Add-SettingsObject -Title "Theme" -Key "AppTheme" -Type "List" `
        -ItemsSource $themeList `
        -Description "Application color theme. Default follows the Windows app theme." `
        -DefaultValue "Default" -Section "General"

    Add-SettingsObject -Title "Hide No-access items" -Key "HideNoAccess" -Type "Boolean" `
        -Description "Remove items from the menu if object permissions is missing. Default is to mark them with red" `
        -DefaultValue $false -Section "General"

    Add-SettingsObject -Title "Environment name" -Key "EnvironmentText" -Type "Text" `
        -Description "Label shown as a badge in the toolbar (e.g. Production, Lab). Leave empty to hide." `
        -Section "General"

    # Guarded: System.Drawing may be unavailable in a headless PS7 session; the badge
    # color picker then just offers the empty entry (the dialog never renders anyway).
    $colorsList = @([PSCustomObject]@{ Name = ""; Value = "" })
    if("System.Drawing.Color" -as [type]) {
        foreach($color in ([System.Drawing.Color].GetProperties() | Where-Object { $_.PropertyType -eq [System.Drawing.Color] } | Sort-Object -Property Name | Select-Object Name).Name) {
            $colorsList += [PSCustomObject]@{ Name = $color; Value = $color }
        }
    }

    Add-SettingsObject -Title "Environment color" -Key "EnvironmentColor" -Type "List" -ItemsSource $colorsList `
        -Description "Background color of the environment badge. Text color is auto-calculated for contrast." `
        -Section "General"

    Add-SettingsObject -Title "Show tenant name" -Key "MenuShowOrganizationName" -Type "Boolean" -DefaultValue $true `
        -Description "Adds the organization name next to the login info on the menu bar" `
        -Section "General"

    # ---- Intune section (list behavior) ----

    Add-SettingsObject -Title "Get all pages" -Key "GetAllPages" -Type "Boolean" `
        -Description "Get all pages when getting items in the UI. Note: This can take long time in environments with lots of policies and apps." `
        -DefaultValue $true -SubPath "IntuneManager" -Section "IntuneManager"

    Add-SettingsObject -Title "Single-line values in object list" -Key "ObjectListFirstLineOnly" -Type "Boolean" `
        -Description "Show only the first line of multi-line values (e.g. store app descriptions) in the object list, so every row is one line high. Hover a cell for the full text. Takes effect when the list is next loaded." `
        -DefaultValue $true -SubPath "IntuneManager" -Section "IntuneManager"

    Add-SettingsObject -Title "Max characters per cell" -Key "ObjectListMaxCellLength" -Type "Int" `
        -Description "With single-line values on, cut a cell's text at this many characters so one long description cannot push the other columns out of view. Hover the cell for the full text; sorting and the filter still use the whole value. 0 = no limit." `
        -DefaultValue 50 -SubPath "IntuneManager" -Section "IntuneManager"

    $viewMenuTypes = @(
        [PSCustomObject]@{ Name = "Group";      Value = "Group" },
        [PSCustomObject]@{ Name = "Type (API)"; Value = "Type"  }
    )

    Add-SettingsObject -Title "Menu Object Type" -Key "ObjectViewType" -Type "List" -ItemsSource $viewMenuTypes `
        -Description "Specify object type for the menu. Group: Groups items together like the portal. Type - Single item based on API. Some APIs are split into multiple menu items." `
        -DefaultValue "Group" -SubPath "IntuneManager" -Section "IntuneManager"

    # ---- MS Graph General section (UI action gating) ----

    Add-SettingsObject -Title "Refresh Objects after copy" -Key "RefreshObjectsAfterCopy" -Type "Boolean" `
        -Description "Reload the object list after copying an object" -DefaultValue $true `
        -SubPath "IntuneManager" -Section "GraphGeneral"

    Add-SettingsObject -Title "Show Delete button" -Key "AllowDelete" -Type "Boolean" `
        -Description "Allow deleting individual objectes" -DefaultValue $false `
        -SubPath "IntuneManager" -Section "GraphGeneral"

    Add-SettingsObject -Title "Show Bulk Delete" -Key "AllowBulkDelete" -Type "Boolean" `
        -Description "Allow using bulk delete to delete all objects of selected types" -DefaultValue $true `
        -SubPath "IntuneManager" -Section "GraphGeneral"

    # ---- Intune Tools section ----

    Add-SettingsObject -Title "Format OMA-URI Settings" -Key "FormatOMAURI" -Type "Boolean" `
        -Description "Automatically clean up XML formatting in OMA-URI and ADMX registry policies for consistent output" -DefaultValue $false `
        -SubPath "IntuneManager" -Section "IntuneTools"

    # ---- MSAL section (picker sort order) ----

    Add-SettingsObject -Title "Sort Account List" -Key "SortAccountList" -Type "Boolean" -DefaultValue $false `
        -Description "Sort the list of cached accounts based on user name. Updated at restart or account change" `
        -Section "MSAL"

    Add-SettingsObject -Title "Sort Tenant List" -Key "SortTenantList" -Type "Boolean" -DefaultValue $false `
        -Description "Sort the list of available tenants based on Tenant name. Updated at restart or account change" `
        -Section "MSAL"
}

Add-AppEventHandler "AppInitialized" "Add-CommonUISettings"
