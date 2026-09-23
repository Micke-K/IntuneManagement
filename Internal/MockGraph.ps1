# Offline Graph for the mock tenant.
#
# Serves Microsoft Graph requests from JSON files on disk so the whole
# application can run - sign in, browse, view, copy, import, delete, document -
# with no tenant and no network. Used for screenshots, demos and UI work on a
# train. Wired in through the AuthenticationMock provider
# (Classes/AuthenticationMock.ps1), which hands every request here via the
# RoutesAllRequests capability; nothing else in the product knows about it.
#
# Data folder layout (IM_MOCK_DATA):
#
#   settings.json                         settings store for the session (IM_SETTINGS_FILE)
#   tenant.json                           TenantId, TenantName, UPN, DisplayName, UserId
#   graph/<path>.json                     one file per Graph path, e.g.
#   graph/organization.json                 { "value": [ {...} ] }   a collection
#   graph/me.json                           { ... }                  a single object
#   graph/deviceManagement/deviceConfigurations.json
#
# A collection file's items are addressable as <path>/<id> and <path>/<id>/<prop>
# (assignments, settings, ...), $filter / $expand / $top are honoured on list
# calls, $batch is unpacked, and POST / PATCH / DELETE mutate the in-memory copy
# so Copy, Import and Delete work for the length of the session. The files on
# disk are never written to.
#
# Everything unknown answers 200 with an empty collection rather than 404:
# the product treats an empty list as "nothing here", which is the right
# rendering for a path the mock data simply does not cover.

$script:MockGraphStore = $null

# Load (or reload) the data folder. Keys are Graph paths without a leading
# slash: 'deviceManagement/deviceConfigurations'.
function Initialize-MockGraphStore {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)

    $graphRoot = Join-Path $Root "graph"
    $store = @{
        Root        = $Root
        Collections = @{}   # path -> List[object]
        Singles     = @{}   # path -> object
    }
    if(Test-Path -LiteralPath $graphRoot -PathType Container) {
        foreach($file in (Get-ChildItem -LiteralPath $graphRoot -Filter *.json -Recurse -File)) {
            $rel = $file.FullName.Substring($graphRoot.Length).TrimStart('\', '/')
            $key = ($rel -replace '\.json$', '') -replace '\\', '/'
            try {
                $data = [IO.File]::ReadAllText($file.FullName) | ConvertFrom-Json
            }
            catch {
                Write-Log "Mock Graph: cannot parse $($file.FullName): $($_.Exception.Message)" 3
                continue
            }
            if($data.PSObject.Properties['value'] -and $data.value -is [array]) {
                $list = [System.Collections.Generic.List[object]]::new()
                foreach($item in $data.value) { $list.Add($item) }
                $store.Collections[$key] = $list
            }
            else {
                $store.Singles[$key] = $data
            }
        }
    }
    $script:MockGraphStore = $store
    Write-Log "Mock Graph: loaded $($store.Collections.Count) collections and $($store.Singles.Count) objects from $graphRoot"
    return $store
}

function Get-MockGraphStore {
    return $script:MockGraphStore
}

# Split an absolute or relative Graph URL into its path (no host, no version,
# no leading slash) and a query hashtable with unescaped values.
function Split-MockGraphUrl {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Url)

    $path = $Url
    $query = ""
    $q = $path.IndexOf('?')
    if($q -ge 0) { $query = $path.Substring($q + 1); $path = $path.Substring(0, $q) }
    $path = $path -replace '^https?://[^/]+/', ''
    $path = $path -replace '^(beta|v1\.0)/', ''
    $path = $path.Trim('/')
    $path = [Uri]::UnescapeDataString($path)

    $params = @{}
    if($query) {
        foreach($pair in $query.Split('&')) {
            if(-not $pair) { continue }
            $eq = $pair.IndexOf('=')
            if($eq -lt 0) { $params[[Uri]::UnescapeDataString($pair)] = ""; continue }
            $name  = [Uri]::UnescapeDataString($pair.Substring(0, $eq))
            $value = [Uri]::UnescapeDataString($pair.Substring($eq + 1).Replace('+', ' '))
            $params[$name] = $value
        }
    }
    return [PSCustomObject]@{ Path = $path; Query = $params }
}

# ---------------------------------------------------------------------------
# OData $filter - the subset the policy types actually send:
#   a eq 'x'   a ne 'x'   a eq true    isof('microsoft.graph.t')
#   contains(tolower(a),'x')   startswith(a,'x')   endswith(a,'x')
#   and / or / not, parentheses, cast segments (microsoft.graph.t/prop)
# Anything the parser cannot read matches everything (logged once), which
# keeps a list visible rather than empty when a new clause shape appears.
# ---------------------------------------------------------------------------

function ConvertTo-MockFilterTokens {
    param([string]$Filter)
    $tokens = [System.Collections.Generic.List[object]]::new()
    $i = 0
    $n = $Filter.Length
    while($i -lt $n) {
        $c = $Filter[$i]
        if([char]::IsWhiteSpace($c)) { $i++; continue }
        if($c -eq '(' -or $c -eq ')' -or $c -eq ',') {
            $tokens.Add([PSCustomObject]@{ Type = 'punct'; Value = [string]$c }); $i++; continue
        }
        if($c -eq "'") {
            $sb = [System.Text.StringBuilder]::new()
            $i++
            while($i -lt $n) {
                if($Filter[$i] -eq "'") {
                    if($i + 1 -lt $n -and $Filter[$i + 1] -eq "'") { [void]$sb.Append("'"); $i += 2; continue }
                    break
                }
                [void]$sb.Append($Filter[$i]); $i++
            }
            $i++
            $tokens.Add([PSCustomObject]@{ Type = 'string'; Value = $sb.ToString() }); continue
        }
        $start = $i
        while($i -lt $n -and -not [char]::IsWhiteSpace($Filter[$i]) -and $Filter[$i] -notin @('(', ')', ',', "'")) { $i++ }
        $word = $Filter.Substring($start, $i - $start)
        $lower = $word.ToLowerInvariant()
        if($lower -in @('and', 'or', 'not', 'eq', 'ne', 'in')) { $tokens.Add([PSCustomObject]@{ Type = 'op'; Value = $lower }) }
        elseif($lower -eq 'true' -or $lower -eq 'false') { $tokens.Add([PSCustomObject]@{ Type = 'bool'; Value = ($lower -eq 'true') }) }
        elseif($lower -eq 'null') { $tokens.Add([PSCustomObject]@{ Type = 'null'; Value = $null }) }
        elseif($word -match '^-?\d+(\.\d+)?$') { $tokens.Add([PSCustomObject]@{ Type = 'number'; Value = [double]$word }) }
        else { $tokens.Add([PSCustomObject]@{ Type = 'ident'; Value = $word }) }
    }
    return ,$tokens
}

# Resolve 'a/b/c' against an item. A segment that names a type
# ('microsoft.graph.androidManagedStoreAppConfiguration') is a cast: it yields
# $null unless the item is of that type.
function Get-MockFilterPropertyValue {
    param($Item, [string]$Path)
    $current = $Item
    foreach($segment in $Path.Split('/')) {
        if($null -eq $current) { return $null }
        if($segment -like 'microsoft.graph.*') {
            $type = [string]$current.'@odata.type'
            if($type -ne "#$segment") { return $null }
            continue
        }
        $prop = $current.PSObject.Properties[$segment]
        if(-not $prop) { return $null }
        $current = $prop.Value
    }
    return $current
}

# Recursive-descent evaluator. $State is @{ Tokens; Pos; Item }.
function Invoke-MockFilterOr {
    param($State)
    $left = Invoke-MockFilterAnd $State
    while($State.Pos -lt $State.Tokens.Count -and $State.Tokens[$State.Pos].Type -eq 'op' -and $State.Tokens[$State.Pos].Value -eq 'or') {
        $State.Pos++
        $right = Invoke-MockFilterAnd $State
        $left = ($left -or $right)
    }
    return $left
}

function Invoke-MockFilterAnd {
    param($State)
    $left = Invoke-MockFilterNot $State
    while($State.Pos -lt $State.Tokens.Count -and $State.Tokens[$State.Pos].Type -eq 'op' -and $State.Tokens[$State.Pos].Value -eq 'and') {
        $State.Pos++
        $right = Invoke-MockFilterNot $State
        $left = ($left -and $right)
    }
    return $left
}

function Invoke-MockFilterNot {
    param($State)
    if($State.Pos -lt $State.Tokens.Count -and $State.Tokens[$State.Pos].Type -eq 'op' -and $State.Tokens[$State.Pos].Value -eq 'not') {
        $State.Pos++
        return (-not (Invoke-MockFilterNot $State))
    }
    return (Invoke-MockFilterComparison $State)
}

# A value expression: literal, property path, or a function call. Returns the
# value (not a boolean) - comparisons happen one level up.
function Invoke-MockFilterValue {
    param($State)
    $token = $State.Tokens[$State.Pos]
    $State.Pos++
    switch($token.Type) {
        'string' { return $token.Value }
        'bool'   { return $token.Value }
        'null'   { return $null }
        'number' { return $token.Value }
    }
    if($token.Type -ne 'ident') { throw "Unexpected token '$($token.Value)'" }

    $name = $token.Value
    $isCall = ($State.Pos -lt $State.Tokens.Count -and $State.Tokens[$State.Pos].Type -eq 'punct' -and $State.Tokens[$State.Pos].Value -eq '(')
    if(-not $isCall) { return (Get-MockFilterPropertyValue $State.Item $name) }

    $State.Pos++   # (
    $args = [System.Collections.Generic.List[object]]::new()
    while($true) {
        $args.Add((Invoke-MockFilterValue $State))
        $next = $State.Tokens[$State.Pos]
        $State.Pos++
        if($next.Type -eq 'punct' -and $next.Value -eq ')') { break }
        if(-not ($next.Type -eq 'punct' -and $next.Value -eq ',')) { throw "Expected ',' or ')' in $name(...)" }
    }
    switch($name.ToLowerInvariant()) {
        'tolower'    { return ([string]$args[0]).ToLowerInvariant() }
        'toupper'    { return ([string]$args[0]).ToUpperInvariant() }
        'contains'   { return (([string]$args[0]).IndexOf([string]$args[1], [StringComparison]::OrdinalIgnoreCase) -ge 0) }
        'startswith' { return (([string]$args[0]).StartsWith([string]$args[1], [StringComparison]::OrdinalIgnoreCase)) }
        'endswith'   { return (([string]$args[0]).EndsWith([string]$args[1], [StringComparison]::OrdinalIgnoreCase)) }
        'isof'       {
            $type = [string]$args[-1]
            return (([string]$State.Item.'@odata.type') -eq "#$type")
        }
    }
    throw "Unsupported filter function '$name'"
}

function Invoke-MockFilterComparison {
    param($State)
    $token = $State.Tokens[$State.Pos]
    if($token.Type -eq 'punct' -and $token.Value -eq '(') {
        $State.Pos++
        $inner = Invoke-MockFilterOr $State
        $State.Pos++   # )
        return $inner
    }
    $left = Invoke-MockFilterValue $State
    if($State.Pos -lt $State.Tokens.Count -and $State.Tokens[$State.Pos].Type -eq 'op' -and $State.Tokens[$State.Pos].Value -eq 'in') {
        # id in ('a','b') - the shape the group / filter name lookups send.
        $State.Pos += 2   # in (
        $found = $false
        while($State.Pos -lt $State.Tokens.Count) {
            $token = $State.Tokens[$State.Pos]
            $State.Pos++
            if($token.Type -eq 'punct' -and $token.Value -eq ')') { break }
            if($token.Type -eq 'punct' -and $token.Value -eq ',') { continue }
            if([string]$token.Value -ieq [string]$left) { $found = $true }
        }
        return $found
    }
    if($State.Pos -lt $State.Tokens.Count -and $State.Tokens[$State.Pos].Type -eq 'op' -and $State.Tokens[$State.Pos].Value -in @('eq', 'ne')) {
        $op = $State.Tokens[$State.Pos].Value
        $State.Pos++
        $right = Invoke-MockFilterValue $State
        $equal = if($null -eq $left -or $null -eq $right) { ($null -eq $left) -and ($null -eq $right) }
                 elseif($left -is [bool] -or $right -is [bool]) { [bool]$left -eq [bool]$right }
                 elseif($left -is [double] -or $right -is [double]) { [double]$left -eq [double]$right }
                 else { [string]$left -ieq [string]$right }
        if($op -eq 'eq') { return $equal } else { return (-not $equal) }
    }
    return [bool]$left
}

$script:MockGraphFilterWarned = @{}

function Test-MockGraphFilter {
    [CmdletBinding()]
    param($Item, [string]$Filter)

    if([string]::IsNullOrWhiteSpace($Filter)) { return $true }
    try {
        $state = @{ Tokens = (ConvertTo-MockFilterTokens $Filter); Pos = 0; Item = $Item }
        return [bool](Invoke-MockFilterOr $state)
    }
    catch {
        if(-not $script:MockGraphFilterWarned.ContainsKey($Filter)) {
            $script:MockGraphFilterWarned[$Filter] = $true
            Write-Log "Mock Graph: cannot evaluate `$filter '$Filter' ($($_.Exception.Message)) - returning every item" 2
        }
        return $true
    }
}

# ---------------------------------------------------------------------------
# Request handling
# ---------------------------------------------------------------------------

# Deep copy through JSON so a caller mutating the response never touches the store.
function Copy-MockGraphObject {
    param($Object)
    if($null -eq $Object) { return $null }
    return ($Object | ConvertTo-Json -Depth 50 -Compress | ConvertFrom-Json)
}

function New-MockGraphResult {
    param([int]$Status, $Body)
    return [PSCustomObject]@{ Status = $Status; Body = $Body }
}

function Get-MockGraphItem {
    param($Collection, [string]$Id)
    foreach($item in $Collection) {
        if([string]$item.id -eq $Id) { return $item }
    }
    return $null
}

# Longest known collection or single-object key that prefixes the path.
# Returns @{ Key; Kind ('collection'|'single'); Rest (remaining segments) }.
function Resolve-MockGraphPath {
    param([string]$Path)
    $store = $script:MockGraphStore
    $segments = @($Path.Split('/') | Where-Object { $_ -ne '' })
    for($take = $segments.Count; $take -ge 1; $take--) {
        $key = ($segments[0..($take - 1)] -join '/')
        $rest = if($take -lt $segments.Count) { @($segments[$take..($segments.Count - 1)]) } else { @() }
        if($store.Collections.ContainsKey($key)) { return @{ Key = $key; Kind = 'collection'; Rest = $rest } }
        if($store.Singles.ContainsKey($key))     { return @{ Key = $key; Kind = 'single';     Rest = $rest } }
    }
    return $null
}

function Add-MockGraphExpansion {
    param($Item, [string]$Expand)
    if(-not $Expand) { return $Item }
    foreach($clause in $Expand.Split(',')) {
        $name = ($clause.Trim() -split '\(')[0].Trim()
        if(-not $name) { continue }
        if(-not $Item.PSObject.Properties[$name]) {
            $Item | Add-Member -NotePropertyName $name -NotePropertyValue @() -Force
        }
    }
    return $Item
}

function Invoke-MockGraphList {
    param($Collection, [hashtable]$Query)
    $filter = $Query['$filter']
    $expand = $Query['$expand']
    $top    = 0
    if($Query['$top']) { try { $top = [int]$Query['$top'] } catch { } }

    $items = [System.Collections.Generic.List[object]]::new()
    foreach($item in $Collection) {
        if(-not (Test-MockGraphFilter -Item $item -Filter $filter)) { continue }
        $items.Add((Add-MockGraphExpansion (Copy-MockGraphObject $item) $expand))
        if($top -gt 0 -and $items.Count -ge $top) { break }
    }
    return (New-MockGraphResult 200 ([PSCustomObject]@{ value = $items.ToArray() }))
}

# Stamp the properties Graph fills in on create.
function Set-MockGraphCreatedProperties {
    param($Item)
    $now = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
    if(-not $Item.PSObject.Properties['id'] -or -not $Item.id) {
        $Item | Add-Member -NotePropertyName id -NotePropertyValue ([guid]::NewGuid().ToString()) -Force
    }
    foreach($name in 'createdDateTime', 'lastModifiedDateTime') {
        $Item | Add-Member -NotePropertyName $name -NotePropertyValue $now -Force
    }
    return $Item
}

function Invoke-MockGraphCollectionRequest {
    param([string]$Method, [string]$Key, [string[]]$Rest, [hashtable]$Query, $Body)
    $collection = $script:MockGraphStore.Collections[$Key]

    if($Rest.Count -eq 0) {
        switch($Method) {
            'GET'  { return (Invoke-MockGraphList -Collection $collection -Query $Query) }
            'POST' {
                $item = Set-MockGraphCreatedProperties (Copy-MockGraphObject $Body)
                $collection.Add($item)
                return (New-MockGraphResult 201 (Copy-MockGraphObject $item))
            }
        }
        return (New-MockGraphResult 200 ([PSCustomObject]@{}))
    }

    $id = $Rest[0]
    $item = Get-MockGraphItem -Collection $collection -Id $id
    if(-not $item) {
        Write-LogDebug "Mock Graph: $Method $Key/$id - no such item"
        if($Method -eq 'GET' -and $Rest.Count -gt 1) { return (New-MockGraphResult 200 ([PSCustomObject]@{ value = @() })) }
        return (New-MockGraphResult 404 ([PSCustomObject]@{ error = [PSCustomObject]@{ code = 'ResourceNotFound'; message = "Mock tenant has no $Key with id $id" } }))
    }

    if($Rest.Count -eq 1) {
        switch($Method) {
            'GET'    { return (New-MockGraphResult 200 (Add-MockGraphExpansion (Copy-MockGraphObject $item) $Query['$expand'])) }
            'PATCH'  {
                foreach($prop in $Body.PSObject.Properties) {
                    $item | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
                }
                $item | Add-Member -NotePropertyName lastModifiedDateTime -NotePropertyValue ([DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")) -Force
                return (New-MockGraphResult 200 (Copy-MockGraphObject $item))
            }
            'PUT'    {
                foreach($prop in $Body.PSObject.Properties) {
                    $item | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
                }
                return (New-MockGraphResult 200 (Copy-MockGraphObject $item))
            }
            'DELETE' { [void]$collection.Remove($item); return (New-MockGraphResult 204 $null) }
        }
        return (New-MockGraphResult 200 ([PSCustomObject]@{}))
    }

    # <collection>/<id>/<segment>: a navigation property or an action.
    $segment = $Rest[1]
    switch($segment) {
        'assign' {
            # POST .../assign { assignments: [...] } replaces the assignments.
            $assignments = @()
            if($Body -and $Body.PSObject.Properties['assignments']) { $assignments = @($Body.assignments) }
            foreach($a in $assignments) {
                if(-not $a.PSObject.Properties['id'] -or -not $a.id) { $a | Add-Member -NotePropertyName id -NotePropertyValue ([guid]::NewGuid().ToString()) -Force }
            }
            $item | Add-Member -NotePropertyName assignments -NotePropertyValue $assignments -Force
            return (New-MockGraphResult 200 ([PSCustomObject]@{ value = (Copy-MockGraphObject $assignments) }))
        }
        'assignments' {
            if($Method -eq 'POST') {
                $existing = @()
                if($item.PSObject.Properties['assignments']) { $existing = @($item.assignments) }
                $new = Copy-MockGraphObject $Body
                if(-not $new.PSObject.Properties['id'] -or -not $new.id) { $new | Add-Member -NotePropertyName id -NotePropertyValue ([guid]::NewGuid().ToString()) -Force }
                $item | Add-Member -NotePropertyName assignments -NotePropertyValue (@($existing) + @($new)) -Force
                return (New-MockGraphResult 201 (Copy-MockGraphObject $new))
            }
        }
    }

    $prop = $item.PSObject.Properties[$segment]
    if($Method -eq 'GET') {
        if(-not $prop) { return (New-MockGraphResult 200 ([PSCustomObject]@{ value = @() })) }
        $value = Copy-MockGraphObject $prop.Value
        if($null -eq $value) { return (New-MockGraphResult 200 ([PSCustomObject]@{ value = @() })) }
        if($prop.Value -is [array] -or $prop.Value -is [System.Collections.IList]) {
            $items = @($value)
            if($Query['$filter']) { $items = @($items | Where-Object { Test-MockGraphFilter -Item $_ -Filter $Query['$filter'] }) }
            return (New-MockGraphResult 200 ([PSCustomObject]@{ value = $items }))
        }
        return (New-MockGraphResult 200 $value)
    }
    if($Method -eq 'POST' -and $Rest.Count -eq 2) {
        # Action on the item (e.g. .../updateDefinitionValues, .../createCopy): accept.
        return (New-MockGraphResult 200 ([PSCustomObject]@{}))
    }
    return (New-MockGraphResult 200 ([PSCustomObject]@{}))
}

# One request in, one @{ Status; Body } out. The body is an object, not JSON.
function Resolve-MockGraphRequest {
    [CmdletBinding()]
    param([string]$Method, [string]$Url, $Body)

    if(-not $script:MockGraphStore) { throw "Mock Graph store is not initialized" }
    $Method = $Method.ToUpperInvariant()
    $parts = Split-MockGraphUrl $Url
    $path = $parts.Path

    if($path -eq '$batch') {
        $responses = [System.Collections.Generic.List[object]]::new()
        foreach($request in @($Body.requests)) {
            $sub = Resolve-MockGraphRequest -Method ([string]$request.method) -Url ([string]$request.url) -Body $request.body
            $responses.Add([PSCustomObject]@{
                id      = $request.id
                status  = $sub.Status
                headers = [PSCustomObject]@{ 'Content-Type' = 'application/json' }
                body    = $sub.Body
            })
        }
        return (New-MockGraphResult 200 ([PSCustomObject]@{ responses = $responses.ToArray() }))
    }

    $resolved = Resolve-MockGraphPath $path
    if(-not $resolved) {
        Write-LogDebug "Mock Graph: no data for $Method $path - answering with an empty collection"
        if($Method -eq 'GET') { return (New-MockGraphResult 200 ([PSCustomObject]@{ value = @() })) }
        return (New-MockGraphResult 200 ([PSCustomObject]@{}))
    }

    if($resolved.Kind -eq 'single') {
        $obj = $script:MockGraphStore.Singles[$resolved.Key]
        if($resolved.Rest.Count -eq 0) {
            if($Method -eq 'PATCH') {
                foreach($prop in $Body.PSObject.Properties) { $obj | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force }
            }
            return (New-MockGraphResult 200 (Copy-MockGraphObject $obj))
        }
        $prop = $obj.PSObject.Properties[$resolved.Rest[0]]
        if($prop) {
            $value = Copy-MockGraphObject $prop.Value
            if($prop.Value -is [array]) { return (New-MockGraphResult 200 ([PSCustomObject]@{ value = @($value) })) }
            return (New-MockGraphResult 200 $value)
        }
        return (New-MockGraphResult 200 ([PSCustomObject]@{ value = @() }))
    }

    return (Invoke-MockGraphCollectionRequest -Method $Method -Key $resolved.Key -Rest $resolved.Rest -Query $parts.Query -Body $Body)
}

# Shape a result like Invoke-WebRequest's response, which is what
# Invoke-MSGraphAPI reads (StatusCode / StatusDescription / Headers / Content /
# RawContentLength). A 4xx throws the way Invoke-WebRequest does, with the
# response on the exception, so the product's error handling runs unchanged
# and reports the status, error code and message of a mock failure.
function ConvertTo-MockWebResponse {
    param($Result)
    $json = if($null -eq $Result.Body) { "" } else { ($Result.Body | ConvertTo-Json -Depth 50 -Compress) }
    $description = switch($Result.Status) { 200 { 'OK' } 201 { 'Created' } 204 { 'No Content' } 400 { 'Bad Request' } 403 { 'Forbidden' } 404 { 'Not Found' } default { 'OK' } }
    $response = [PSCustomObject]@{
        StatusCode        = [int]$Result.Status
        StatusDescription = $description
        Headers           = @{ 'Content-Type' = 'application/json'; 'request-id' = [guid]::NewGuid().ToString() }
        Content           = $json
        RawContentLength  = [long][System.Text.Encoding]::UTF8.GetByteCount($json)
    }
    if($Result.Status -ge 400) {
        # Read-MSGraphErrorResponseContent takes the body from
        # Response.GetResponseStream(), as it does for an HttpWebResponse.
        $response | Add-Member -NotePropertyName ContentBytes -NotePropertyValue ([System.Text.Encoding]::UTF8.GetBytes($json))
        $response | Add-Member -MemberType ScriptMethod -Name GetResponseStream -Value { [System.IO.MemoryStream]::new([byte[]]$this.ContentBytes) }
        throw [MockGraphResponseException]::new("The remote server returned an error: ($($Result.Status)) $description.", $response)
    }
    return $response
}

# Entry point used by AuthenticationMock.InvokeWebRequest.
function Invoke-MockGraphRequest {
    [CmdletBinding()]
    param([string]$Url, [string]$Method, $Body)

    $bodyObject = $null
    if($null -ne $Body) {
        if($Body -is [string]) {
            if($Body.Trim()) { try { $bodyObject = $Body | ConvertFrom-Json } catch { $bodyObject = $null } }
        }
        elseif($Body -is [byte[]]) {
            try { $bodyObject = [System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json } catch { $bodyObject = $null }
        }
        else { $bodyObject = $Body }
    }
    $result = Resolve-MockGraphRequest -Method $Method -Url $Url -Body $bodyObject
    return (ConvertTo-MockWebResponse $result)
}
