#ImportOrder 29

# Self-registration registry for the compare subsystem, mirroring
# [DocumentationRegistry]. Replaces the hardcoded $script:compare* arrays that
# Initialize-CompareModule used to build - which only the WPF AppInitialized
# handler ever ran, so in Avalonia mode the compare combos came up empty.
# Providers, output providers, and comparison types now register once at module
# load (Internal/Compare.ps1) so BOTH UI backends see the same catalog, and the
# documentation subsystem self-registers its own "doc" comparison type instead
# of Compare.ps1 reaching across with a Get-Command probe.
#
# Loaded at #ImportOrder 29 - before CompareClasses.ps1 (30) and
# CompareOutputClasses.ps1 (31) - so the provider/output classes it stores are
# already defined by the time Internal/Compare.ps1 registers instances.
#
# Dedupe is by .Value (a provider/output/type key), matching the
# DocumentationRegistry replace-by-identity semantics so Import-Module -Force
# re-registration never accumulates duplicates.
class CompareRegistry {
    static [System.Collections.Generic.List[object]]        $Providers       = [System.Collections.Generic.List[object]]::new()
    static [System.Collections.Generic.List[object]]        $OutputProviders = [System.Collections.Generic.List[object]]::new()
    static [System.Collections.Generic.List[PSCustomObject]] $ComparisonTypes = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ---- Compare providers (CompareProviderBase subclasses) ----
    static [void] RegisterProvider([object]$Provider) {
        if (-not $Provider.Value) { throw "Compare provider must have a Value" }
        $existing = [CompareRegistry]::Providers | Where-Object { $_.Value -eq $Provider.Value }
        if ($existing) { [CompareRegistry]::Providers.Remove($existing) | Out-Null }
        [CompareRegistry]::Providers.Add($Provider)
    }

    static [object] FindProvider([string]$Value) {
        return [CompareRegistry]::Providers | Where-Object { $_.Value -eq $Value } | Select-Object -First 1
    }

    # ---- Output providers (CompareOutputProviderBase subclasses) ----
    static [void] RegisterOutputProvider([object]$Provider) {
        if (-not $Provider.Value) { throw "Compare output provider must have a Value" }
        $existing = [CompareRegistry]::OutputProviders | Where-Object { $_.Value -eq $Provider.Value }
        if ($existing) { [CompareRegistry]::OutputProviders.Remove($existing) | Out-Null }
        [CompareRegistry]::OutputProviders.Add($Provider)
    }

    static [object] FindOutputProvider([string]$Value) {
        return [CompareRegistry]::OutputProviders | Where-Object { $_.Value -eq $Value } | Select-Object -First 1
    }

    static [object] FindOutputProviderByExtension([string]$Extension) {
        return [CompareRegistry]::OutputProviders | Where-Object { $_.Extension -eq $Extension } | Select-Object -First 1
    }

    # ---- Comparison types (PSCustomObject: Name, Value, [Compare], [RemoveProperties]) ----
    static [void] RegisterComparisonType([PSCustomObject]$Type) {
        if (-not $Type.Value) { throw "Comparison type must have a Value" }
        $existing = [CompareRegistry]::ComparisonTypes | Where-Object { $_.Value -eq $Type.Value }
        if ($existing) { [CompareRegistry]::ComparisonTypes.Remove($existing) | Out-Null }
        [CompareRegistry]::ComparisonTypes.Add($Type)
    }

    static [PSCustomObject] FindComparisonType([string]$Value) {
        return [CompareRegistry]::ComparisonTypes | Where-Object { $_.Value -eq $Value } | Select-Object -First 1
    }

    # Clear every collection. Called once at the top of the module-load
    # registration in Internal/Compare.ps1 so Import-Module -Force starts clean
    # (the static initializers above only run on first type load; the type is
    # cached across -Force reimports).
    static [void] Reset() {
        [CompareRegistry]::Providers.Clear()
        [CompareRegistry]::OutputProviders.Clear()
        [CompareRegistry]::ComparisonTypes.Clear()
    }
}
