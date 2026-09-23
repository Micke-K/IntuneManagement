[CmdletBinding()]
param(
    [string]$OutputDirectory,
    # Publish for a platform other than the one running this script. NuGet hands out
    # the RID-specific native packages regardless of host OS, so "osx-arm64" from a
    # Windows box produces the real macOS .dylib set - which is how the committed Mac
    # binaries were built without a Mac. Both osx RIDs yield the same universal
    # (x86_64 + arm64) dylibs, so either one covers every Mac.
    [ValidateSet('win-x64', 'win-arm64', 'osx-x64', 'osx-arm64', 'linux-x64', 'linux-arm64')]
    [string]$RuntimeIdentifier
)

$ErrorActionPreference = 'Stop'

$bootstrapDir = $PSScriptRoot
$avaloniaRoot = Split-Path -Parent $bootstrapDir
$uiRoot       = Split-Path -Parent $avaloniaRoot
$projectRoot  = Split-Path -Parent $uiRoot

if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $projectRoot 'Bin/Avalonia'
}

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "dotnet SDK is required. Install .NET 8 SDK from https://dot.net and retry."
}

# Publish for the current OS so the right native libs land in the output:
# Windows gets libSkiaSharp.dll/av_libglesv2.dll, Linux gets libSkiaSharp.so,
# macOS gets the .dylib. The managed Avalonia assemblies are identical across
# platforms. dotnet publish doesn't clean the output dir, so publishing on a
# second OS into the same Bin/Avalonia just adds that OS's natives alongside the
# existing ones — keeping the committed folder usable cross-platform.
$arch = [System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString().ToLowerInvariant()
$isWin = $IsWindows -or $PSVersionTable.PSEdition -eq 'Desktop'
$rid =
    if ($RuntimeIdentifier) { $RuntimeIdentifier }
    elseif ($isWin)   { "win-$arch" }
    elseif ($IsMacOS) { "osx-$arch" }
    else          { "linux-$arch" }

Write-Host "Publishing Avalonia binaries ($rid) to $OutputDirectory" -ForegroundColor Cyan

$publishArgs = @(
    'publish',
    (Join-Path $bootstrapDir 'AvaloniaPayload.csproj'),
    '-c', 'Release',
    '-r', $rid,
    '--self-contained', 'false',
    '-o', $OutputDirectory
)

& dotnet @publishArgs
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE"
}

Write-Host "Avalonia binaries published." -ForegroundColor Green
Write-Host "Set `$env:IM_UI_BACKEND = 'Avalonia' (or run UI/Avalonia/Start-Avalonia.ps1) to launch with Avalonia." -ForegroundColor Yellow
