# Import all functions from Private and Public folders
$PublicFunctions = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue)
$PrivateFunctions = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)

# Dot source the functions
foreach ($Function in @($PrivateFunctions + $PublicFunctions)) {
    try {
        . $Function.FullName
    }
    catch {
        Write-Error -Message "Failed to import function $($Function.FullName): $_"
    }
}

# Export only the public functions
Export-ModuleMember -Function $PublicFunctions.BaseName
