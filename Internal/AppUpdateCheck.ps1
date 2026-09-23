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

    $tempManifest = $null
    try {
        $content = Invoke-RestMethod $url @params
        if (-not $content.content) { return $null }

        $text = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($content.content))
        $tempManifest = [IO.Path]::Combine([IO.Path]::GetTempPath(), "IM-update-$([Guid]::NewGuid().ToString('N')).psd1")
        [IO.File]::WriteAllText($tempManifest, $text, (New-Object System.Text.UTF8Encoding($false)))

        $data = Import-PowerShellDataFile -Path $tempManifest -ErrorAction Stop
        if (-not $data.ModuleVersion) { return $null }

        $full = [string]$data.ModuleVersion
        $label = $data.PrivateData.PSData.Prerelease
        if ($label) { $full = "$full-$label" }
        return (ConvertTo-AppVersion $full)
    }
    catch {
        Write-LogDebug "Update check: manifest lookup at ref '$Ref' failed: $($_.Exception.Message)"
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
        Write-LogDebug "Update check: releases lookup failed: $($_.Exception.Message)"
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

    $local  = Get-AppLocalVersion
    $major  = if ($local) { $local.Release.Major } else { 0 }
    $remote = Get-AppRemoteVersion -Major $major

    if (-not $local)  { Write-Log 'Failed to get version info from local file' 2 }
    if (-not $remote) { Write-Log 'Failed to get version info in GitHub' 2 }

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
            Write-LogDebug "Update check: release notes lookup at ref '$ref' failed: $($_.Exception.Message)"
        }
    }
    return $null
}
