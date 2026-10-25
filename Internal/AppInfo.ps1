# Facts about this application, read from its own module manifest.
#
# The About dialog used to carry them as literal text in the XAML - in BOTH backends,
# separately - so the copyright line said "(c) 2022" in a 2026 build, and disagreed
# with the manifest's own Copyright field ("(c) 2019-2026 ..."), which had been kept
# current. Two hardcoded copies that nothing compares will always drift.
#
# IntuneManagement.psd1 is the single source of truth: it is the file the module is
# loaded from, PowerShell exposes its metadata on the module object, and it is already
# the thing maintained at release time.

# Copyright line for the About dialog, from the manifest.
#
# $ExecutionContext.SessionState.Module is this module, so no name lookup and nothing
# to keep in step if the module is ever renamed. The fallback is deliberately vague
# about the year rather than inventing one: a wrong year is worse than no year.
function Get-AppCopyright {
    [OutputType([string])]
    param()

    $text = $null
    try { $text = [string]$ExecutionContext.SessionState.Module.Copyright } catch { }
    if (-not [string]::IsNullOrWhiteSpace($text)) { return $text.Trim() }

    $author = $null
    try { $author = [string]$ExecutionContext.SessionState.Module.Author } catch { }
    if (-not [string]::IsNullOrWhiteSpace($author)) { return "(c) $author. Software released under MIT License." }

    return 'Software released under MIT License.'
}
