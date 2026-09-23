[CmdletBinding()]
param()

# Rebuilds Bin/MainThreadHook/IntuneManagement.MainThreadHook.dll from
# MainThreadHook/StartupHook.cs. Needs the .NET 8+ SDK and NuGet access; the
# output is a single small DLL that IS committed (unlike the old SDK-bundling
# launcher), so run this only when the source changes and commit the result.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$project = Join-Path $PSScriptRoot 'MainThreadHook/MainThreadHook.csproj'
$build = Join-Path $PSScriptRoot 'MainThreadHook/bin/publish'
$output = Join-Path $root 'Bin/MainThreadHook'

& dotnet build $project -c Release -o $build --nologo
if ($LASTEXITCODE -ne 0) { throw "Main-thread hook build failed ($LASTEXITCODE)" }

New-Item -ItemType Directory -Path $output -Force | Out-Null
Copy-Item (Join-Path $build 'IntuneManagement.MainThreadHook.dll') $output -Force
Get-Item (Join-Path $output 'IntuneManagement.MainThreadHook.dll') | Select-Object FullName, Length
