#ImportOrder 10

# CLR row shape for Avalonia ComboBoxes that need DisplayMemberBinding.
#
# Avalonia's data binder reflects on the runtime CLR type, not on PSObject
# ETS NoteProperties — so a ComboBox with ItemsSource = PSCustomObject[]
# and DisplayMemberBinding="{Binding Name}" silently renders blank cells
# because PSCustomObject has no real .Name property (see
# [[avalonia-binding-needs-clr-types]]).
#
# Wrap items in this class to get real CLR Name / Value / EnglishName /
# Source properties that the Avalonia binder + PowerShell readers (e.g.
# Get-AvaloniaDocumentationOptions reading SelectedItem.Value) both see.
# Source keeps the original PSCustomObject so callers that need extra
# fields (e.g. language metadata) can pull them back.

class NameValueComboItem
{
    [string] $Name
    [string] $EnglishName
    [object] $Value
    [object] $Source

    NameValueComboItem() {}

    # Backstop only. Each control picks its own field via DisplayMemberBinding
    # (13 combos bind Name, the language combo binds EnglishName), so this
    # exists purely so a control that was never given a binding degrades to
    # readable text instead of rendering "NameValueComboItem".
    [string] ToString() { return [string]$this.Name }
}
