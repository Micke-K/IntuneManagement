#ImportOrder 10

# CLR-typed Name/Value pair shape used by Show-MSALDecodedToken's DataGrid.
# The WPF version builds an array of PSCustomObjects with Name/Value; that
# would silently render blank in Avalonia (binder can't traverse NoteProperty),
# so the Avalonia tool converts to instances of this class.

class TokenInfoRow
{
    [string] $Name
    [object] $Value

    TokenInfoRow() {}
}
