#ImportOrder 6

# CLR-typed shape for items inside the Settings dialog's ComboBox controls
# (Type="List" settings). Existing setting registrations pass PSCustomObject
# arrays in -ItemsSource with .Name / .Value properties — those won't bind in
# Avalonia (see [[avalonia-binding-needs-clr-types]]). Add-SettingComboBox
# converts each PSCustomObject into one of these before assigning ItemsSource.

class SettingsListItem
{
    [string]$Name
    [object]$Value

    SettingsListItem() {}

    [string] ToString()
    {
        return [string]$this.Name
    }
}
