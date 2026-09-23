#ImportOrder 9

# CLR-typed row shape for the Cached Objects DataGrid (CachedObjects.axaml).
# Update-CachedObjectsView builds these from $script:cacheObjects (a hashtable
# of PSCustomObject in non-UI code) — Avalonia's binder needs CLR types here
# rather than the raw cache entries.

class CacheRowItem
{
    [string]   $Name
    [object]   $Tags
    [string]   $TagsText
    [object]   $Value
    [string]   $ValueText
    [bool]     $Persistent
    [datetime] $TimeOut
    [long]     $Size
    [string]   $SizeText

    CacheRowItem() {}
}
