#ImportOrder 9

# CLR-typed row shape for the Log DataGrid (LogInfo.axaml). The non-UI
# Internal/Core.ps1 keeps log entries as PSCustomObject in
# $script:LogItems; Avalonia's binder can't traverse PSCustomObject inside
# DataTemplates, so Get-LogViewPanel projects each entry into one of these
# at refresh time.
#
# RowForeground is computed at projection time so the DataGrid row style
# can bind directly to it (Type==2 -> orange, Type==3 -> red, default -> null).

class LogRowItem
{
    [int]      $ID
    [datetime] $DateTime
    [int]      $Type
    [string]   $TypeText
    [string]   $Text
    [object]   $RowForeground

    LogRowItem() {}
}
