# Update check - shared, UI-free.
#
# Both backends need the same two answers ("is a newer version published?" and
# "what do the published release notes say?"), and neither answer involves a UI
# toolkit. Keeping the network + version comparison here means the backends stay
# thin callers (R10) and no WPF type leaks into shared code (R12); the WPF and
# Avalonia sides only differ in how they show the result.
#
# Both versions live in ONE public repo: 3.x on the default branch, 4.x on the
# v4 branch until the two are renamed at release. So every lookup asks for the
# v4 ref first and falls back to no ref, which means the default branch. That
# ordering is self-healing: on the day v4 is renamed to the default branch the
# v4 ref stops existing, the first call 404s, and the fallback reads exactly the
# content that used to be on v4. Nothing here changes at the swap.

$script:AppUpdateRepo = 'Micke-K/IntuneManagement'

# Tried in order. $null means "no ref", i.e. whatever the default branch is.
$script:AppUpdateRefs = @('v4', $null)

# raw.githubusercontent.com fallback refs. Unlike the contents API there is no
# "omit the ref" form, so the default branch has to be named - hence master AND
# main, tried after v4. The list keeps the self-healing property described above:
# when v4 becomes the default branch its own ref stops resolving and the next
# entry serves exactly the content v4 used to.
$script:AppUpdateRawRefs = @('v4', 'master', 'main')

# Why a raw fallback exists at all: api.github.com allows 60 requests per hour per
# IP for unauthenticated callers, and that quota is counted per EGRESS ip - so every
# client behind one corporate proxy shares a single pool of 60 and a busy office
# never gets above it. raw.githubusercontent.com serves file content from a CDN and
# is not under that quota (verified live while the API was returning 403), so the
# two file lookups this module makes keep working when the API is exhausted.
# /releases has no raw equivalent and stays API-only.

# Human-readable reason the last GitHub lookup failed, or $null. Read by the UI so
# the dialog can say what actually happened instead of "not reachable" - the
# failures used to go only to Write-LogDebug, which is silent unless the Debug
# setting is on, so a rate-limited check looked exactly like a broken proxy.
$script:AppUpdateLastError = $null

function Get-AppUpdateLastError {
    return $script:AppUpdateLastError
}

function Clear-AppUpdateLastError {
    $script:AppUpdateLastError = $null
}

# Record (and log) why a lookup failed. First failure wins: it is the one that
# describes the primary cause, later refs usually fail the same way.
#
# A 403 from this API is almost always the rate limit rather than a permission
# problem, so it is named as such and the reset time is looked up from
# /rate_limit - that endpoint is explicitly exempt from the quota, so asking costs
# nothing even while everything else is being refused.
function Set-AppUpdateError {
    param([string]$Context, $ErrorRecord)

    $message = if ($ErrorRecord) { [string]$ErrorRecord.Exception.Message } else { 'unknown error' }
    $detail  = $message

    if ($message -match '(?i)rate limit') {
        $detail = 'GitHub API rate limit exceeded (60 requests per hour per IP for anonymous callers, shared by everyone behind the same proxy)'
        $reset = Get-AppUpdateRateLimitReset
        if ($reset) {
            $detail += ". Resets at $($reset.ToString('HH:mm'))"
        }
    }

    # Debug, not warning: a single lookup failing is not the outcome. The raw fallback
    # below often rescues it, and four warnings for a check that then succeeded is its
    # own kind of misleading. The entry points warn ONCE, with this reason, only when
    # the whole lookup ends up empty.
    Write-LogDebug "Update check: $Context failed. $detail"
    if (-not $script:AppUpdateLastError) { $script:AppUpdateLastError = $detail }
}

# Local time the anonymous quota resets, or $null. /rate_limit does not count
# against the quota, so this is safe to call from the failure path.
function Get-AppUpdateRateLimitReset {
    try {
        $params = Get-AppUpdateRestParams
        $limit = Invoke-RestMethod "https://api.github.com/rate_limit" @params
        $epoch = [long]$limit.resources.core.reset
        if ($epoch -gt 0) { return [DateTimeOffset]::FromUnixTimeSeconds($epoch).ToLocalTime() }
    }
    catch {
        Write-LogDebug "Update check: rate limit lookup failed: $($_.Exception.Message)"
    }
    return $null
}

# The git blob sha of some text, the same value the contents API reports in .sha.
#
# git hashes a blob as sha1("blob <bytecount>\0<bytes>") over LF-normalised content.
# Computing it here is what lets the raw fallback return a Sha at all - raw serves
# bytes with no metadata - so a caller comparing published against local keeps
# working whichever path the text arrived by.
#
# The length prefix counts BYTES, not characters, and the payload is hashed as UTF-8
# bytes. Those differ the moment the file contains one non-ASCII character ("cafe"
# with an e-acute is 5 characters but 6 bytes), which would silently yield a sha that
# matches nothing GitHub ever reports - so a published file identical to the local
# one would be reported as differing. Verified against git hash-object for both an
# ASCII and a non-ASCII payload (Tests/AppUpdateFallback.Tests.ps1).
function Get-AppGitBlobSha {
    param([string]$Text)

    if ($null -eq $Text) { return $null }

    $payload = [System.Text.Encoding]::UTF8.GetBytes($Text.Replace("`r`n", "`n"))
    $header  = [System.Text.Encoding]::ASCII.GetBytes("blob $($payload.Length)`0")

    $stream = $null
    try {
        $stream = [IO.MemoryStream]::new()
        $stream.Write($header, 0, $header.Length)
        $stream.Write($payload, 0, $payload.Length)
        [void]$stream.Seek(0, [IO.SeekOrigin]::Begin)
        return (Get-FileHash -InputStream $stream -Algorithm SHA1).Hash
    }
    finally {
        if ($stream) { $stream.Dispose() }
    }
}

# A file's text from raw.githubusercontent.com, or $null. Tried refs in order.
function Get-AppRawFileText {
    param([Parameter(Mandatory)][string]$Path)

    $params = Get-AppUpdateRestParams
    foreach ($ref in $script:AppUpdateRawRefs) {
        $url = "https://raw.githubusercontent.com/$($script:AppUpdateRepo)/$ref/$Path"
        try {
            $text = Invoke-RestMethod $url @params
            if ($text) {
                Write-LogDebug "Update check: '$Path' served by raw.githubusercontent.com at ref '$ref'"
                return [string]$text
            }
        }
        catch {
            Write-LogDebug "Update check: raw lookup of '$Path' at ref '$ref' failed: $($_.Exception.Message)"
        }
    }
    return $null
}

# Invoke-RestMethod splat honouring the configured proxy, if any.
function Get-AppUpdateRestParams {
    $params = @{}
    $proxyURI = Get-ProxyURI
    if ($proxyURI) {
        $params.Add('proxy', $proxyURI)
        $params.Add('UseBasicParsing', $true)
    }
    return $params
}

# A version that can carry a pre-release label, which [version] cannot:
# [version]'4.0.0-beta1' throws, and so does the 'v4.0.0-beta1' tag form.
#
# ModuleVersion alone cannot tell beta1 from beta2 - both are 4.0.0, because the
# label lives in PrivateData.PSData.Prerelease - so the label is carried here and
# compared, or a beta tester could never be told a newer beta exists.
function ConvertTo-AppVersion {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }

    $value = $Text.Trim()
    if ($value.StartsWith('v', [StringComparison]::OrdinalIgnoreCase)) { $value = $value.Substring(1) }

    $prerelease = $null
    $dash = $value.IndexOf('-')
    if ($dash -ge 0) {
        $prerelease = $value.Substring($dash + 1)
        $value = $value.Substring(0, $dash)
    }

    $release = $null
    try { $release = [version]$value } catch { return $null }

    $result = [PSCustomObject]@{
        Release    = $release
        Prerelease = $prerelease
    }
    # So the UI can keep calling .ToString() and get something a human reads.
    $result | Add-Member -MemberType ScriptMethod -Name ToString -Force -Value {
        if ($this.Prerelease) { "$($this.Release)-$($this.Prerelease)" } else { "$($this.Release)" }
    }
    return $result
}

# -1 / 0 / 1, SemVer-style: a pre-release is OLDER than the same release without
# one (4.0.0-beta1 is less than 4.0.0), and two labels compare numerically where
# they can, so beta10 never sorts below beta2.
function Compare-AppVersion {
    param($Left, $Right)

    if (-not $Left -and -not $Right) { return 0 }
    if (-not $Left)  { return -1 }
    if (-not $Right) { return 1 }

    $byRelease = $Left.Release.CompareTo($Right.Release)
    if ($byRelease -ne 0) { return [Math]::Sign($byRelease) }

    if (-not $Left.Prerelease -and -not $Right.Prerelease) { return 0 }
    if (-not $Left.Prerelease)  { return 1 }
    if (-not $Right.Prerelease) { return -1 }

    $leftParts  = $Left.Prerelease.Split('.')
    $rightParts = $Right.Prerelease.Split('.')
    $count = [Math]::Max($leftParts.Count, $rightParts.Count)
    for ($i = 0; $i -lt $count; $i++) {
        $a = if ($i -lt $leftParts.Count)  { $leftParts[$i] }  else { '' }
        $b = if ($i -lt $rightParts.Count) { $rightParts[$i] } else { '' }
        if ($a -eq $b) { continue }
        if (-not $a) { return -1 }
        if (-not $b) { return 1 }

        # beta2 vs beta10: split the trailing digits off and compare those as numbers.
        $matchA = [regex]::Match($a, '^(?<t>.*?)(?<n>[0-9]+)$')
        $matchB = [regex]::Match($b, '^(?<t>.*?)(?<n>[0-9]+)$')
        if ($matchA.Success -and $matchB.Success -and $matchA.Groups['t'].Value -eq $matchB.Groups['t'].Value) {
            return [Math]::Sign([int]$matchA.Groups['n'].Value - [int]$matchB.Groups['n'].Value)
        }
        return [Math]::Sign([string]::Compare($a, $b, [StringComparison]::OrdinalIgnoreCase))
    }
    return 0
}

# The module manifest published at one ref, as an AppVersion, or $null.
#
# Parsed by writing it to a temp file and reading it with Import-PowerShellDataFile
# - the manifest is PowerShell data, not JSON, and this avoids hand-rolling a
# parser or invoking the text as code.
function Get-AppRemoteManifestVersion {
    param([string]$Ref)

    $params = Get-AppUpdateRestParams
    $url = "https://api.github.com/repos/$($script:AppUpdateRepo)/contents/IntuneManagement.psd1"
    if ($Ref) { $url = $url + '?ref=' + $Ref }

    try {
        $content = Invoke-RestMethod $url @params
        if (-not $content.content) { return $null }

        $text = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($content.content))
        return (ConvertTo-AppVersionFromManifestText $text)
    }
    catch {
        Set-AppUpdateError "manifest lookup at ref '$Ref'" $_
        return $null
    }
}

# A manifest's ModuleVersion (plus pre-release label) from its TEXT.
#
# Written to a temp file and read with Import-PowerShellDataFile: the manifest is
# PowerShell data, not JSON, and this avoids hand-rolling a parser or invoking the
# text as code. Split out from the API path so the raw fallback parses identically.
function ConvertTo-AppVersionFromManifestText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }

    $tempManifest = $null
    try {
        $tempManifest = [IO.Path]::Combine([IO.Path]::GetTempPath(), "IM-update-$([Guid]::NewGuid().ToString('N')).psd1")
        [IO.File]::WriteAllText($tempManifest, $Text, (New-Object System.Text.UTF8Encoding($false)))

        $data = Import-PowerShellDataFile -Path $tempManifest -ErrorAction Stop
        if (-not $data.ModuleVersion) { return $null }

        $full = [string]$data.ModuleVersion
        $label = $data.PrivateData.PSData.Prerelease
        if ($label) { $full = "$full-$label" }
        return (ConvertTo-AppVersion $full)
    }
    catch {
        Write-LogDebug "Update check: manifest parse failed: $($_.Exception.Message)"
        return $null
    }
    finally {
        if ($tempManifest -and [IO.File]::Exists($tempManifest)) {
            try { [IO.File]::Delete($tempManifest) } catch { }
        }
    }
}

# Newest published release sharing $Major, as an AppVersion, or $null.
#
# The releases LIST, not releases/latest: that endpoint hides pre-releases by
# design, so a beta could never be told about a newer beta. Releases are not tied
# to a branch either, so this path is unaffected by the rename. tag_name, not
# name - one published release on this repo has an empty name field.
function Get-AppRemoteReleaseVersion {
    param([int]$Major = 0)

    try {
        $params = Get-AppUpdateRestParams
        $releases = Invoke-RestMethod "https://api.github.com/repos/$($script:AppUpdateRepo)/releases" @params
        $best = $null
        foreach ($release in @($releases)) {
            if ($release.draft) { continue }
            $candidate = ConvertTo-AppVersion $release.tag_name
            if (-not $candidate) { continue }
            if ($Major -gt 0 -and $candidate.Release.Major -ne $Major) { continue }
            if (-not $best -or (Compare-AppVersion $candidate $best) -gt 0) { $best = $candidate }
        }
        return $best
    }
    catch {
        Set-AppUpdateError 'releases lookup' $_
        return $null
    }
}

# Version published on GitHub, or $null when it cannot be determined.
function Get-AppRemoteVersion {
    param([int]$Major = 0)

    foreach ($ref in $script:AppUpdateRefs) {
        $fromManifest = Get-AppRemoteManifestVersion -Ref $ref
        if ($fromManifest -and ($Major -le 0 -or $fromManifest.Release.Major -eq $Major)) { return $fromManifest }
    }

    # Contents API gave nothing at any ref. Before falling back to /releases, try the
    # same manifest from raw: when the failure was the shared anonymous rate limit,
    # raw is not subject to it and answers normally.
    $rawText = Get-AppRawFileText -Path 'IntuneManagement.psd1'
    if ($rawText) {
        $fromRaw = ConvertTo-AppVersionFromManifestText $rawText
        if ($fromRaw -and ($Major -le 0 -or $fromRaw.Release.Major -eq $Major)) {
            $reason = Get-AppUpdateLastError
            if ($reason) { Write-Log "Version read from raw.githubusercontent.com instead of the API: $reason" }
            Clear-AppUpdateLastError
            return $fromRaw
        }
    }

    # No manifest at any ref matched. The 3.x line has no IntuneManagement.psd1
    # at all - its manifest is named differently - so ask the releases instead.
    return (Get-AppRemoteReleaseVersion -Major $Major)
}

# Version of the manifest sitting next to this module, or $null. Carries the
# pre-release label so one beta can be told from the next.
function Get-AppLocalVersion {
    try {
        $manifest = Join-Path $script:AppRootFolder 'IntuneManagement.psd1'
        if (-not [IO.File]::Exists($manifest)) { return $null }
        $data = Import-PowerShellDataFile -Path $manifest -ErrorAction Stop
        if (-not $data.ModuleVersion) { return $null }

        $full = [string]$data.ModuleVersion
        $label = $data.PrivateData.PSData.Prerelease
        if ($label) { $full = "$full-$label" }
        return (ConvertTo-AppVersion $full)
    }
    catch {
        Write-LogDebug "Update check: local manifest read failed: $($_.Exception.Message)"
    }
    return $null
}

# Both versions plus the verdict. IsOutdated is only ever $true when both
# versions were resolved, so a failed network call reads as "nothing to report"
# rather than "up to date" or a spurious upgrade prompt.
#
# Only a newer version of the SAME major is offered. Crossing a major is a
# migration with breaking changes, announced deliberately - not something to push
# at an installation through a startup dialog, and not something a branch rename
# should trigger for every 3.x user at once. Drop the -Major argument below to go
# back to "newest wins".
function Get-AppUpdateInfo {
    [CmdletBinding()]
    param()

    Clear-AppUpdateLastError

    $local  = Get-AppLocalVersion
    $major  = if ($local) { $local.Release.Major } else { 0 }
    $remote = Get-AppRemoteVersion -Major $major

    if (-not $local)  { Write-Log 'Failed to get version info from local file' 2 }
    if (-not $remote) {
        # Name the cause. 'Failed to get version info in GitHub' on its own reads as a
        # connectivity problem, which sent one report chasing a proxy setting for a
        # check that was simply rate limited.
        $reason = Get-AppUpdateLastError
        if ($reason) { Write-Log "Failed to get version info in GitHub: $reason" 2 }
        else         { Write-Log 'Failed to get version info in GitHub' 2 }
    }

    $outdated = ($local -and $remote -and (Compare-AppVersion $local $remote) -lt 0)

    if ($outdated) {
        Write-Log 'Local version and GitHub version does not match' 2
        Write-Log "Local version: $($local.ToString())"
        Write-Log "GitHub version: $($remote.ToString())"
    }
    elseif ($local -and $remote) {
        Write-Log "Running latest version: $($local.ToString())"
    }

    return [PSCustomObject]@{
        LocalVersion  = $local
        RemoteVersion = $remote
        IsOutdated    = [bool]$outdated
        Resolved      = [bool]($local -and $remote)
    }
}

# Published release notes, or $null when unreachable.
#
# Same ref order as the version lookup, and for the same reason: a 4.x
# installation must not be shown the 3.x notes, and after the rename the
# fallback lands on what used to be the v4 branch with no change here.
#
# Returns Text plus Sha: GitHub's blob sha lets a caller tell "the published notes
# differ from my local copy" without diffing the text, which is how the WPF
# Updates dialog decides whether to flag an update.
function Get-AppRemoteReleaseNotes {
    [CmdletBinding()]
    param()

    Clear-AppUpdateLastError

    $params = Get-AppUpdateRestParams
    foreach ($ref in $script:AppUpdateRefs) {
        $url = "https://api.github.com/repos/$($script:AppUpdateRepo)/contents/ReleaseNotes.md"
        if ($ref) { $url = $url + '?ref=' + $ref }
        try {
            $content = Invoke-RestMethod $url @params
            if ($content.content) {
                return [PSCustomObject]@{
                    Text = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($content.content))
                    Sha  = $content.sha
                }
            }
        }
        catch {
            Set-AppUpdateError "release notes lookup at ref '$ref'" $_
        }
    }

    # Raw fallback: same file, CDN instead of the rate-limited API. Sha is computed
    # locally (Get-AppGitBlobSha) so the caller's published-vs-local comparison
    # behaves exactly as it does on the API path.
    $rawText = Get-AppRawFileText -Path 'ReleaseNotes.md'
    if ($rawText) {
        $reason = Get-AppUpdateLastError
        if ($reason) { Write-Log "Release notes read from raw.githubusercontent.com instead of the API: $reason" }
        Clear-AppUpdateLastError
        return [PSCustomObject]@{
            Text = $rawText
            Sha  = Get-AppGitBlobSha $rawText
        }
    }

    $reason = Get-AppUpdateLastError
    if ($reason) { Write-Log "Failed to get release notes from GitHub: $reason" 2 }
    else         { Write-Log 'Failed to get release notes from GitHub' 2 }
    return $null
}
