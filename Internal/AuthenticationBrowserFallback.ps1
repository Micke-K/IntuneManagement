# Browserless-login support.
#
# On Linux/macOS (and Windows with "Use system browser") MSAL drives interactive
# sign-in through the system browser + a loopback redirect. On a machine with no
# usable default browser the redirect never arrives and the interactive poll loop
# waits out its whole timeout with the UI locked. This file provides the pieces
# that let the login flow (a) detect that situation up-front, (b) fail fast when a
# browser launch actually fails, and (c) fall back to device-code login instead.
#
# See docs/superpowers/specs/2026-09-05-browserless-login-device-code-fallback-design.md

# Marker embedded in every "could not open a browser" failure so the caller can
# tell a browser-launch failure apart from a normal auth error and fall back to
# device code.
$script:IMNoBrowserMarker = "IM_NO_BROWSER"

# Thin, mockable wrapper around Get-Command so the browser probe can be unit
# tested without shelling out to the real PATH.
function Test-CommandExists {
    param([string]$Name)
    if(-not $Name) { return $false }
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

# Wraps `xdg-settings get default-web-browser`. Returns the handler (e.g.
# 'firefox.desktop') or $null. Separated out so tests can mock it.
function Get-XdgDefaultWebBrowser {
    if(-not (Test-CommandExists "xdg-settings")) { return $null }
    try {
        $out = & xdg-settings get default-web-browser 2>$null
        if($out) {
            $line = ($out | Select-Object -First 1)
            if($line) { return ([string]$line).Trim() }
        }
    }
    catch { }
    return $null
}

# True when interactive browser sign-in has a realistic chance of working.
# Conservative by design: only returns $false when we are confident there is no
# browser, so a false negative can't send a browser-capable machine down the
# device-code path. The fast-fail launcher (IMMsalBrowserLauncher) covers the
# opposite case (a browser that is configured but broken).
function Test-InteractiveBrowserAvailable {
    # Windows: WAM / embedded WebView / Edge is always available.
    if($script:IsWindowsOS) { return $true }
    # macOS: LaunchServices always resolves a default handler.
    if($IsMacOS) { return $true }

    # Linux: probe. 1) $env:BROWSER pointing at a resolvable command.
    if($env:BROWSER) {
        $cmd = (($env:BROWSER -split ':')[0] -split '\s+')[0]
        if($cmd -and (Test-CommandExists $cmd)) { return $true }
    }

    # 2) A registered xdg default web browser.
    if(Get-XdgDefaultWebBrowser) { return $true }

    # 3) A known browser binary on PATH. NB: xdg-open is deliberately excluded -
    # it ships on headless boxes and is not itself a browser.
    foreach($b in @(
        "firefox","firefox-esr","google-chrome","google-chrome-stable","chrome",
        "chromium","chromium-browser","brave-browser","microsoft-edge",
        "microsoft-edge-stable","opera","vivaldi","vivaldi-stable","epiphany",
        "konqueror","falkon","midori")) {
        if(Test-CommandExists $b) { return $true }
    }

    return $false
}

# True when a failure (exception or string) carries the no-browser marker, i.e.
# the interactive browser launch could not open a browser at all.
function Test-IMNoBrowserFailure {
    param($Failure)
    if($null -eq $Failure) { return $false }

    $text = $null
    if($Failure -is [System.Exception]) {
        $text = $Failure.Message
        $inner = $Failure.InnerException
        $depth = 0
        while($inner -and $depth -lt 10) {
            $text = "$text $($inner.Message)"
            $inner = $inner.InnerException
            $depth++
        }
    }
    else {
        $text = [string]$Failure
    }

    return ($text -match $script:IMNoBrowserMarker)
}

# Compile the compiled-C# browser launcher used for MSAL's
# SystemWebViewOptions.OpenBrowserFunc. It MUST be compiled C# (not a ScriptBlock
# cast to a delegate): MSAL invokes OpenBrowserFunc on an internal thread with no
# PowerShell Runspace, where a ScriptBlock delegate throws "no Runspace
# available" - the same constraint documented for the device-code callback.
#
# The launcher owns the browser launch so it can DETECT failure: if the browser
# command can't start, or exits non-zero quickly (xdg-open returns 3 = "no
# handler", 4 = "action failed"), it throws with the IM_NO_BROWSER marker, which
# faults the MSAL task in ~1s instead of hanging the whole interactive timeout.
function Initialize-IMBrowserLauncher {
    if("IMMsalBrowserLauncher" -as [type]) { return }

    $refAssemblies = @(
        [System.Object].Assembly.Location,
        [System.Diagnostics.Process].Assembly.Location,
        [System.Threading.Tasks.Task].Assembly.Location,
        [System.Uri].Assembly.Location,
        # Process derives from System.ComponentModel.Component (System.ComponentModel.Primitives
        # on .NET Core); without this ref the compile fails with CS0012 on Process.Start.
        [System.ComponentModel.Component].Assembly.Location,
        [System.ComponentModel.Win32Exception].Assembly.Location
    ) | Where-Object { $_ } | Select-Object -Unique

    # -CompilerOptions is PS7-only: Windows PowerShell 5.1 compiles through CodeDom
    # and takes -CompilerParameters instead, so passing it there fails the whole
    # call ("a parameter cannot be found"). The type then never compiles, the
    # "already a type?" guard above never becomes true, and every call retried and
    # logged the same error - on 5.1 the fast-fail browser detection was simply
    # absent. The option only silences warnings 1701/1702, so 5.1 goes without it.
    $addTypeParams = @{
        ReferencedAssemblies = $refAssemblies
        ErrorAction          = 'Stop'
    }
    if($PSVersionTable.PSEdition -eq 'Core') {
        $addTypeParams['CompilerOptions'] = '/nowarn:1701;1702'
    }

    Add-Type -TypeDefinition @"
using System;
using System.Diagnostics;
using System.Threading.Tasks;

public static class IMMsalBrowserLauncher {
    // Default browser command when `$env:BROWSER` is unset. Set from PowerShell
    // (which knows the platform) in Initialize-IMBrowserLauncher.
    public static string DefaultCommand = "xdg-open";
    public const string Marker = "IM_NO_BROWSER";

    public static Task Open(Uri uri) {
        string exe = ResolveBrowserCommand();
        if (string.IsNullOrEmpty(exe)) {
            throw new Exception(Marker + ": no web browser command could be resolved");
        }

        Process p = null;
        try {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = exe;
            // AbsoluteUri, NOT Uri.ToString().
            //
            // Uri.ToString() unescapes the %20 separators MSAL puts in "scope".
            // ProcessStartInfo.Arguments then sees literal spaces and splits one URL
            // into several argv entries. AbsoluteUri keeps all whitespace escaped,
            // so the compatible Arguments property still passes one argument.
            //
            // Do not use ProcessStartInfo.ArgumentList here. It is preferable on
            // modern .NET, but is absent from .NET Framework and makes this Add-Type
            // fail during module import under Windows PowerShell 5.1.
            psi.Arguments = uri.AbsoluteUri;
            psi.UseShellExecute = false;
            p = Process.Start(psi);
        }
        catch (Exception ex) {
            throw new Exception(Marker + ": failed to launch '" + exe + "': " + ex.Message, ex);
        }

        if (p == null) {
            throw new Exception(Marker + ": failed to launch '" + exe + "'");
        }

        // A real browser forks and keeps running, so WaitForExit(3000) returns
        // false and we treat that as "launched OK". A quick non-zero exit means
        // the launch failed (no handler / action failed).
        if (p.WaitForExit(3000) && p.ExitCode != 0) {
            throw new Exception(Marker + ": '" + exe + "' exited with code " + p.ExitCode);
        }

        return Task.CompletedTask;
    }

    static string ResolveBrowserCommand() {
        string b = Environment.GetEnvironmentVariable("BROWSER");
        if (!string.IsNullOrEmpty(b)) {
            string first = b.Split(':')[0];
            string[] parts = first.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 0) { return parts[0]; }
        }
        return DefaultCommand;
    }
}
"@ @addTypeParams

    if($IsMacOS) {
        try { [IMMsalBrowserLauncher]::DefaultCommand = "open" } catch { }
    }
}

# Surface the pending device code (stashed by IMMsalDeviceCodeHelper) into the
# GUI: copy the user code to the clipboard and show a status message telling the
# user where to enter it. Called from the -DeviceCode poll loop in
# Get-MsalAuthenticationToken. No-ops cleanly on a console/headless host (no
# UIProvider), where MSAL's callback already printed the code to the console.
function Show-DeviceCodeInstruction {
    if(-not ("IMMsalDeviceCodeHelper" -as [type])) { return }

    $code = $null
    $url  = $null
    try { $code = [IMMsalDeviceCodeHelper]::UserCode } catch { }
    try { $url  = [IMMsalDeviceCodeHelper]::VerificationUrl } catch { }
    if(-not $url) { $url = "https://microsoft.com/devicelogin" }
    if(-not $code) { return }

    $copied = $false
    if($script:UIProvider) {
        try {
            $script:UIProvider.SetClipboardText($code)
            $copied = $true
        }
        catch {
            Write-LogDebug "Could not copy device code to clipboard: $($_.Exception.Message)"
        }
    }

    $prefix = if($copied) { "Sign-in code copied to clipboard. " } else { "" }
    # -CancelText surfaces the overlay's Cancel button. The ACTION was armed by the
    # wait loop that owns the cancellation token (Get-MsalAuthenticationToken), so
    # this only has to ask for the button - device-code waits run up to 15 minutes
    # and this is the only way out of an abandoned one.
    Write-Status "Waiting for sign-in" "${prefix}Open $url and enter code $code" -CancelText "Cancel"
}

# Compile at load so the type is available to the login flow (and tests) without
# a lazy first-call cost. Guarded, so re-import (PS7 -Force) is a no-op.
Initialize-IMBrowserLauncher
