# Where the MSAL.NET assemblies live for this PowerShell edition.
#
# Add-MSALPrereq loads from here, and Test-MSALResumeLikely checks the folder
# exists before any load is attempted: a deployment that only ever signs in with
# the OAuth provider is allowed to delete Bin/ entirely (README, Headless), and
# that must not produce an error per missing DLL at every import.
#
# Own file per R9 - AuthenticationMSALHelpers.ps1 is at its line budget.
function Get-MSALBinariesFolder
{
    [CmdletBinding()]
    param()

    $edition = if($PSVersionTable.PSVersion.Major -lt 7) { "MSAL_PS5" } else { "MSAL_PS7" }
    return (Join-Path (Join-Path $script:AppRootFolder "Bin") $edition)
}
