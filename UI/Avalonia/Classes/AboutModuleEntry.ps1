#ImportOrder 4

# CLR-typed shape for rows in the About dialog's loaded-modules ListBox.
# See [[avalonia-binding-needs-clr-types]] — Avalonia's binder does not
# special-case PSObject NoteProperties, so DataTemplate items must be real
# CLR types with real properties.

class AboutModuleEntry
{
    [string]$Name
    [string]$Version
    [string]$Type

    AboutModuleEntry() {}
}
