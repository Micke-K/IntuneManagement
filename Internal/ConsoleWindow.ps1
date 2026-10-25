# The console window the application was started from - show, hide, toggle.
#
# Version 3 hid that console at startup unless -ShowConsoleWindow was passed
# to the launcher. Version 4 leaves it alone at startup and offers two things
# instead: File > Hide console / Show console for the session, and the
# HideConsoleWindow setting (UI/Classes/UICommonSettings.ps1) to hide it once
# the main window is up. The UI decides when; this file only knows how, and
# reads no setting itself.
#
# Where it works: Windows, and only when the process owns a real console
# window - the classic console host that Start.cmd and Start-PS7.cmd get.
# Under Windows Terminal the process owns a hidden pseudo-console and the
# visible window belongs to Terminal, so ShowWindow changes nothing the user
# can see; that case reports "not supported" and the UI hides the menu item.
# macOS and Linux have no console window of the process at all - the terminal
# emulator is another application and cannot be hidden from here - so the
# feature does not exist there and the menu item is never shown.

# Support is decided once, on first use, and BEFORE anything is hidden: a
# console that is not visible when we first look is either Windows Terminal's
# hidden pseudo-console or a -WindowStyle Hidden start, and neither is a window
# the user expects this menu to control.
$script:ConsoleWindowSupported = $null

function Initialize-ConsoleWindowApi
{
    if(-not $script:IsWindowsOS) { return $false }
    if("IMConsoleWindow" -as [type]) { return $true }
    try
    {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class IMConsoleWindow
{
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool   ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]   public static extern bool   IsWindowVisible(IntPtr hWnd);
}
'@ -ErrorAction Stop
        return $true
    }
    catch
    {
        Write-LogError "Console window API could not be loaded" $_.Exception
        return $false
    }
}

function Get-ConsoleWindowHandle
{
    if(-not (Initialize-ConsoleWindowApi)) { return [IntPtr]::Zero }
    return [IMConsoleWindow]::GetConsoleWindow()
}

# $true when Show/Hide-ConsoleWindow can do something the user will see.
function Test-ConsoleWindowSupported
{
    if($null -ne $script:ConsoleWindowSupported) { return $script:ConsoleWindowSupported }

    $supported = $false
    if($script:IsWindowsOS)
    {
        $handle = Get-ConsoleWindowHandle
        if($handle -ne [IntPtr]::Zero)
        {
            $supported = [bool][IMConsoleWindow]::IsWindowVisible($handle)
            if(-not $supported) { Write-LogDebug "Console window exists but is not visible (Windows Terminal or a hidden start) - hide/show console is off" }
        }
        else { Write-LogDebug "No console window - hide/show console is off" }
    }
    $script:ConsoleWindowSupported = $supported
    return $supported
}

function Get-ConsoleWindowVisible
{
    if(-not (Test-ConsoleWindowSupported)) { return $false }
    return [bool][IMConsoleWindow]::IsWindowVisible((Get-ConsoleWindowHandle))
}

# SW_SHOW (5): the user asked for the console, so it comes to the front.
function Show-ConsoleWindow
{
    if(-not (Test-ConsoleWindowSupported)) { return $false }
    [void][IMConsoleWindow]::ShowWindow((Get-ConsoleWindowHandle), 5)
    Write-Log "Console window shown"
    return $true
}

# SW_HIDE (0).
function Hide-ConsoleWindow
{
    if(-not (Test-ConsoleWindowSupported)) { return $false }
    [void][IMConsoleWindow]::ShowWindow((Get-ConsoleWindowHandle), 0)
    Write-Log "Console window hidden"
    return $true
}

# Toggles and returns the new state: $true visible, $false hidden.
function Switch-ConsoleWindow
{
    if(-not (Test-ConsoleWindowSupported)) { return $false }
    if(Get-ConsoleWindowVisible) { [void](Hide-ConsoleWindow); return $false }
    [void](Show-ConsoleWindow)
    return $true
}
