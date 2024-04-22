#region Console functions

# https://stackoverflow.com/questions/40617800/opening-powershell-script-and-hide-command-prompt-but-not-the-gui
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);

[DllImport("user32.dll")]
public static extern bool SetForegroundWindow(IntPtr hWnd);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleIcon(IntPtr hIcon);

[DllImport("user32.dll")] 
public static extern int SendMessage(int hWnd, uint wMsg, uint wParam, IntPtr lParam); 
'

function Show-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()

    # Hide = 0,
    # ShowNormal = 1,
    # ShowMinimized = 2,
    # ShowMaximized = 3,
    # Maximize = 3,
    # ShowNormalNoActivate = 4,
    # Show = 5,
    # Minimize = 6,
    # ShowMinNoActivate = 7,
    # ShowNoActivate = 8,
    # Restore = 9,
    # ShowDefault = 10,
    # ForceMinimized = 11

    [Console.Window]::ShowWindow($consolePtr, 4)
}

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
}

#endregion

# Unblock all files
# Not 100% OK but avoid issues with loading blocked files
function Unblock-AllFiles
{
    param($folder)

    (Get-ChildItem $folder -force | Where-Object {! $_.PSIsContainer}) | Unblock-File
    
    foreach($subFolder in (Get-ChildItem $folder -force | Where-Object {$_.PSIsContainer}))
    {
        Unblock-AllFiles $subFolder.FullName
    }
}

function Initialize-CloudAPIManagement
{
    [CmdletBinding(SupportsShouldProcess=$True)]
    param(
        [string]
        $View = "",
        [switch]
        $ShowConsoleWindow,
        [switch]
        $JSonSettings,
        [string]
        $JSonFile,
        [switch]
        $Silent,
        [string]
        $SilentBatchFile,
        [string]
        $tenantId,
        [string]
        $appId,
        [string]
        $secret,
        [string]
        $certificate
    )

    $PSModuleAutoloadingPreference = "none"

    $global:wpfNS = "xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'"

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    Add-Type -AssemblyName PresentationFramework

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $global:hideUI = ($Silent -eq $true)
    $global:SilentBatchFile = $SilentBatchFile

    if($tenantId)
    {
        $global:AzureAppId = $appId 
        $global:ClientSecret = $secret 
        $global:ClientCert = $certificate
    }

    if($global:hideUI -ne $true)
    {                
        # Run with UI
        try 
        {
            [xml]$xaml = Get-Content ([IO.Path]::GetDirectoryName($PSCommandPath) + "\Xaml\SplashScreen.xaml")
            $global:SplashScreen = ([Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml)))
            $global:txtSplashTitle = $global:SplashScreen.FindName("txtSplashTitle")
            $global:txtSplashText = $global:SplashScreen.FindName("txtSplashText")

            $global:txtSplashTitle.Text = ("Initializing Cloud API PowerShell Management")

            $global:SplashScreen.Show() | Out-Null
            [System.Windows.Forms.Application]::DoEvents()
        }
        catch 
        {
            
        }
    }
    else
    {
        # Run silent

        if(-not $tenantId)
        {
            # Core module not loaded yet so can't use log function
            Write-Error "Tenant Id is missing. Use -TenantId <Tenant-guid> on the command line to run silent batch jobs"
            return
        }
    }

    $global:TenantId = $tenantId


    if($ShowConsoleWindow -ne $true)
    {
        Hide-Console
    }

    if($JSonSettings -eq $true)
    {
        $global:UseJSonSettings = $true
        $global:JSonSettingFile = $JSonFile
    }
    else
    {
        $global:UseJSonSettings = $false
    }

    if($global:hideUI -ne $true)
    {
        $global:txtSplashText.Text = "Unblock files"
    } 
    [System.Windows.Forms.Application]::DoEvents()
    Unblock-AllFiles $PSScriptRoot

    if($global:hideUI -ne $true)
    {
        $global:txtSplashText.Text = "Load core module"
    }
    [System.Windows.Forms.Application]::DoEvents()
    Import-Module ($PSScriptRoot + "\Core.psm1") -Force -Global

    Start-CoreApp $View
}
