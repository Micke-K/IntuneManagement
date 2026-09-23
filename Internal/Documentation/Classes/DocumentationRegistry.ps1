class DocumentationRegistry {
    # ---- Output providers (phase 1, populated) ----
    static [System.Collections.Generic.List[PSCustomObject]] $Outputs = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ---- Per-@odata.type custom handler classes (phase 4 populates) ----
    static [System.Collections.Generic.List[object]] $Handlers = [System.Collections.Generic.List[object]]::new()

    # ---- Schema-driven input providers (phase 3 populates) ----
    # Each provider is a PSCustomObject with: Name, Order (int, lower = higher
    # priority), Match (scriptblock taking $PolicyObject, $Context), Translate.
    # FindInputProvider evaluates providers in Order and returns the FIRST whose
    # Match claims the object; the rest are skipped. The generic fallback registers
    # with Order = [int]::MaxValue so it is always evaluated last.
    static [System.Collections.Generic.List[PSCustomObject]] $InputProviders = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Monotonic registration counter, used as a stable tiebreak when two providers
    # share the same Order (preserves registration order for equal-Order providers).
    static [int] $InputProviderSeq = 0

    # ---- ObjectInfo walker customizations ----
    # These augment schema-driven Manifest/Profile translation without claiming
    # the whole object like a DocumentationHandlerBase implementation does.
    static [System.Collections.Generic.List[PSCustomObject]] $ObjectInfoCustomizers = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ---- Outputs ----

    static [void] RegisterOutput([PSCustomObject]$Provider) {
        if (-not $Provider.Value) { throw "Output provider must have a Value" }
        $existing = [DocumentationRegistry]::Outputs | Where-Object { $_.Value -eq $Provider.Value }
        if ($existing) {
            [DocumentationRegistry]::Outputs.Remove($existing) | Out-Null
        }
        [DocumentationRegistry]::Outputs.Add($Provider)
    }

    static [PSCustomObject] FindOutput([string]$Value) {
        return [DocumentationRegistry]::Outputs | Where-Object { $_.Value -eq $Value } | Select-Object -First 1
    }

    # ---- Handlers ----

    static [void] RegisterHandler([object]$Handler) {
        if (-not $Handler.ODataTypes -or $Handler.ODataTypes.Count -eq 0) {
            throw "Documentation handler must declare ODataTypes"
        }
        # Replace any existing handler claiming the same @odata.type
        $newKeys = @($Handler.ODataTypes)
        $stale = [DocumentationRegistry]::Handlers | Where-Object {
            $existingKeys = @($_.ODataTypes)
            foreach ($k in $existingKeys) { if ($newKeys -contains $k) { return $true } }
            $false
        }
        foreach ($s in $stale) { [DocumentationRegistry]::Handlers.Remove($s) | Out-Null }
        [DocumentationRegistry]::Handlers.Add($Handler)
    }

    static [object] FindHandler([string]$ODataType) {
        if (-not $ODataType) { return $null }
        foreach ($h in [DocumentationRegistry]::Handlers) {
            if ($h.ODataTypes -contains $ODataType) { return $h }
        }
        return $null
    }

    # ---- ObjectInfo walker customizations ----

    static [void] RegisterObjectInfoCustomizer([PSCustomObject]$Customizer) {
        if (-not $Customizer.Name) { throw "ObjectInfo customizer must have a Name" }
        # A customizer either claims specific @odata.types (ODataTypes/ODataTypePatterns)
        # or opts into MatchAll to run for every object (its hooks branch internally,
        # matching the old DocumentationCustom.psm1 single-provider model).
        if (-not $Customizer.MatchAll -and
            (-not $Customizer.ODataTypes -or $Customizer.ODataTypes.Count -eq 0) -and
            (-not $Customizer.ODataTypePatterns -or $Customizer.ODataTypePatterns.Count -eq 0)) {
            throw "ObjectInfo customizer $($Customizer.Name) must declare MatchAll, ODataTypes or ODataTypePatterns"
        }
        $existing = [DocumentationRegistry]::ObjectInfoCustomizers | Where-Object Name -EQ $Customizer.Name
        if ($existing) {
            [DocumentationRegistry]::ObjectInfoCustomizers.Remove($existing) | Out-Null
        }
        [DocumentationRegistry]::ObjectInfoCustomizers.Add($Customizer)
    }

    static [object[]] FindObjectInfoCustomizers([string]$ODataType) {
        return @([DocumentationRegistry]::ObjectInfoCustomizers | Where-Object {
            if ($_.MatchAll) { return $true }
            if (-not $ODataType) { return $false }
            if ($_.ODataTypes -contains $ODataType) { return $true }
            foreach ($pattern in @($_.ODataTypePatterns)) {
                if ($ODataType -like $pattern) { return $true }
            }
            return $false
        })
    }

    # ---- Input providers ----

    static [void] RegisterInputProvider([PSCustomObject]$Provider) {
        if (-not $Provider.Name)     { throw "Input provider must have a Name" }
        if (-not $Provider.Match)    { throw "Input provider $($Provider.Name) must have a Match scriptblock" }
        if (-not $Provider.Translate){ throw "Input provider $($Provider.Name) must have a Translate scriptblock" }

        # Default Order for providers that don't declare one: after the specific
        # providers (which use < 100), before the fallback ([int]::MaxValue).
        if (-not $Provider.PSObject.Properties['Order'] -or $null -eq $Provider.Order) {
            $Provider | Add-Member -NotePropertyName Order -NotePropertyValue 100 -Force
        }
        # Stable tiebreak for equal Orders (registration order).
        $Provider | Add-Member -NotePropertyName _Seq -NotePropertyValue ([DocumentationRegistry]::InputProviderSeq++) -Force

        $existing = [DocumentationRegistry]::InputProviders | Where-Object { $_.Name -eq $Provider.Name }
        if ($existing) {
            [DocumentationRegistry]::InputProviders.Remove($existing) | Out-Null
        }
        [DocumentationRegistry]::InputProviders.Add($Provider)
    }

    static [PSCustomObject] FindInputProvider([object]$PolicyObject) {
        return [DocumentationRegistry]::FindInputProvider($PolicyObject, $null)
    }

    # Evaluate providers in ascending Order (ties broken by registration sequence)
    # and return the first whose Match claims the object. $Context is passed to
    # Match as a second argument; providers that declare only param($PolicyObject)
    # ignore it, while the fallback reads it to honor the opt-in option.
    static [PSCustomObject] FindInputProvider([object]$PolicyObject, [object]$Context) {
        $ordered = [DocumentationRegistry]::InputProviders | Sort-Object @{ Expression = { [int]$_.Order } }, @{ Expression = { [int]$_._Seq } }
        foreach ($p in $ordered) {
            try {
                if (& $p.Match $PolicyObject $Context) {
                    Write-Log "InputProvider found: $($p.Name) matched '$($PolicyObject.Name)' - Policy Type: $($PolicyObject.PolicyName)"
                    return $p
                }
            } catch {
                # A throwing Match must not silently skip the provider - that turns a
                # buggy Match into a mysterious "no provider matched / empty document".
                # Log it (warning) and keep evaluating the remaining providers.
                Write-Log "Documentation input provider '$($p.Name)' Match threw for '$($PolicyObject.Name)': $($_.Exception.Message)" 2
            }
        }
        return $null
    }

    # Clear every registry collection. The static List initializers above run
    # only on the FIRST type load; on Import-Module -Force the type is cached so
    # they do NOT re-run and the lists keep their previous-import entries. Each
    # provider/handler re-registers (replace-by-identity) on every import, but
    # one whose identity was EDITED between reloads (a handler's ODataTypes, a
    # provider's Name) would orphan its old entry because the replace can't
    # match the changed key. Resetting at load time guarantees a clean slate.
    static [void] Reset() {
        [DocumentationRegistry]::Outputs.Clear()
        [DocumentationRegistry]::Handlers.Clear()
        [DocumentationRegistry]::InputProviders.Clear()
        [DocumentationRegistry]::InputProviderSeq = 0
        [DocumentationRegistry]::ObjectInfoCustomizers.Clear()
    }
}

# Get-DocumentationOutputValues backs the -OutputFormat ArgumentCompleter on
# Start-GraphBulkDocumentation (invoked via & (Get-Module IntuneManagement) {
# Get-DocumentationOutputValues }). ArgumentCompleter is used instead of a
# [ValidateSet([IValidateSetValuesGenerator])] so the module still imports on
# Windows PowerShell 5.1 (that interface is PS7-only).

function Get-DocumentationOutputValues {
    # Runs in the main IntuneManagement module scope so [DocumentationRegistry] is
    # visible (invoked from the ArgumentCompleter via & (Get-Module ...)).
    try {
        return @([DocumentationRegistry]::Outputs | Sort-Object Value | Select-Object -ExpandProperty Value)
    }
    catch {
        return @()
    }
}

# Start each module import with empty collections (see Reset() above). Safe to
# run unconditionally: Documentation.ps1 dot-sources Classes/ before Core/,
# InputProviders/, OutputProviders/ and PolicyTypeHandlers/, so this executes
# before any Add-Documentation*/RegisterHandler call repopulates the lists.
[DocumentationRegistry]::Reset()
