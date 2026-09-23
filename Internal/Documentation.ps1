# Loader stub for the Documentation feature.
#
# The module loader at IntuneManagement.psm1:135 enumerates Internal/*.ps1 non-recursively,
# so files under Internal/Documentation/** would not auto-import. This stub recursively
# dot-sources them in lexical order. Folder names are chosen so the lexical sort loads
# dependencies first: Classes -> Core -> InputProviders -> OutputProviders -> PolicyTypeHandlers
# (Assets holds CSS only, no .ps1). "Classes" sorts before "Core" because 'l' < 'o'.

$documentationRoot = Join-Path $PSScriptRoot 'Documentation'

if ([IO.Directory]::Exists($documentationRoot)) {
    Get-ChildItem -Path $documentationRoot -Filter '*.ps1' -Recurse |
        Sort-Object FullName |
        ForEach-Object {
            Write-Verbose "Loading documentation file: $($_.FullName)"
            if ($script:IsWindowsOS) { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
            . $_.FullName
        }
}

# Contribute the "Documentation" comparison type to the compare subsystem when
# documentation is present. This inverts the old cross-subsystem coupling where
# Internal/Compare.ps1 probed Get-Command "Get-GraphDocumentation" to decide
# whether to add the type - documentation now owns and self-registers its own
# compare type (Internal/Compare.ps1 loads first, so the wrapper exists). No
# Compare scriptblock: doc compare needs the policy wrappers, and
# Compare-PolicyObjects routes Value "doc" straight to Compare-ObjectsBasedonDocumentation.
if (Get-Command Register-CompareComparisonType -ErrorAction SilentlyContinue) {
    Register-CompareComparisonType ([PSCustomObject]@{
        Name  = "Documentation"
        Value = "doc"
    })
}
