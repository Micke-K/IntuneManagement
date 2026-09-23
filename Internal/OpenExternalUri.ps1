# Opening a link in the default browser - shared by both UI backends.
#
# Start-Process <url> happens to work on Windows because the cmdlet falls
# back to ShellExecute for non-executables, but raw Process.Start(url) fails
# on .NET Core (UseShellExecute defaults to false there) and Start-Process
# itself fails off-Windows where a URL is not an executable file.
# ProcessStartInfo with UseShellExecute = true is the one form that opens the
# default browser on every supported host (PS 5.1, PS 7, and - once the
# Avalonia UI goes cross-platform - macOS/Linux via the OS launcher).
#
# Being a real module function also matters for the Avalonia side: event
# handlers there cannot rely on closures (ConvertTo-AvaloniaEventScriptBlock
# strips them), so link handlers must call a module-scope function by name.
function Open-ExternalUri
{
    param([Parameter(Mandatory)][string]$Uri)

    try {
        if($IsWindows -or $PSVersionTable.PSEdition -eq 'Desktop') {
            # Windows: ShellExecute opens the default handler for a URL or file.
            $startInfo = New-Object System.Diagnostics.ProcessStartInfo
            $startInfo.FileName = $Uri
            $startInfo.UseShellExecute = $true
            [void][System.Diagnostics.Process]::Start($startInfo)
            return
        }

        if($IsMacOS) {
            [void][System.Diagnostics.Process]::Start((New-OpenProcessInfo 'open' @($Uri)))
            return
        }

        # Linux. UseShellExecute is unreliable on .NET here (it can hang), so we
        # invoke the desktop's launcher directly. XFCE ignores xdg-settings and
        # routes links through exo-open -> ~/.config/xfce4/helpers.rc (WebBrowser),
        # so prefer `exo-open --launch WebBrowser` when present (it works for both
        # URLs and local file paths); otherwise fall back to xdg-open.
        #
        # Do NOT wrap the launch in `systemd-run --user --scope`: that puts a
        # snap-packaged browser (Firefox) in a generic run-<id>.scope which
        # snap-confine rejects ("... is not a snap cgroup"). What actually lets the
        # snap start is the correct DBUS_SESSION_BUS_ADDRESS + XDG_RUNTIME_DIR,
        # which the launchers (Start.ps1 / Start-Avalonia.ps1) export.
        if(Get-Command exo-open -ErrorAction SilentlyContinue) {
            [void][System.Diagnostics.Process]::Start((New-OpenProcessInfo 'exo-open' @('--launch','WebBrowser',$Uri)))
        }
        else {
            [void][System.Diagnostics.Process]::Start((New-OpenProcessInfo 'xdg-open' @($Uri)))
        }
    }
    catch {
        Write-LogError "Failed to open $Uri" $_.Exception
    }
}

# ArgumentList passes each argument verbatim (spaces in a file path are safe) and
# UseShellExecute=false keeps Process.Start non-blocking without the ShellExecute
# hang seen on Linux .NET.
function New-OpenProcessInfo
{
    param([string]$FileName, [string[]]$Arguments)

    $si = [System.Diagnostics.ProcessStartInfo]::new()
    $si.FileName = $FileName
    foreach($a in $Arguments) { $si.ArgumentList.Add($a) }
    $si.UseShellExecute = $false
    return $si
}
