$script:MSALApps = @()
$script:MSALTokens = @{}
# Token ids are now allocated centrally by Get-NextAuthTokenId (AuthenticationCore.ps1)
# so they never collide with OAuth/MgGraph ids. The old per-provider counter is retired.

# Set once Initialize-MSALPrereq has successfully loaded the MSAL runtime. The DLLs
# are NOT loaded at module import - only on first actual MSAL use (New-MSALApp /
# New-MSALConfidentialApp / Get-TenantList fallback), so OAuth/MgGraph-only sessions
# and fresh no-cache startups never pay the load + Roslyn-compile cost.
$script:MSALPrereqLoaded = $false

function Invoke-MSALInitialize
{
    #Add Setting setion
    Add-SettingsSection -Title "MSAL" -Id "MSAL" -Order 8


    #Add MSAL Setting values

    Add-SettingsObject -Title "Remember Login" -Key "CacheMSALToken" -Type "Boolean" -DefaultValue $true `
        -Description "Store the MSAL token in an encrypted file and automatically log on when the script starts. The token is stored in the users profile and can only be decrypted by the user that created it. Note: Requires restart" `
        -Section "MSAL"

    Add-SettingsObject -Title "Get Tenant List" -Key "GetTenantList" -Type "Boolean" -DefaultValue $false `
        -Description "Get a list of all tenants the current user has access to. Only used when the user has access to multiple tenants. This may cause duplicate login/consent prompts first time" `
        -Section "MSAL"
    

    # SortAccountList / SortTenantList are UI-only settings (picker ordering) -
    # registered in UI/Classes/UICommonSettings.ps1, shared by both backends.

    # Off by default: WAM only keeps a session alive for the Windows account (README, Signing in).
    Add-SettingsObject -Title "Use Web Account Manager (WAM) for login" -Key "UseWAM" -Type "Boolean" -DefaultValue $false `
        -Description "Use the Windows Web Account Manager broker (Windows Hello, device compliance claims). Turn this on only when the account you sign in with is your Windows account or one added under Settings > Accounts - for any other account WAM cannot keep the session alive and you are prompted again about every hour. Requires PowerShell 7 and an app restart." `
        -Section "MSAL"

    Add-SettingsObject -Title "WAM: list OS accounts" -Key "WAMListOSAccounts" -Type "Boolean" -DefaultValue $true `
        -Description "When WAM is enabled, also surface machine-joined Entra accounts in the account picker (PS7+). Off-by-default on PS5." `
        -Section "MSAL"

    # On by default since 4.0.0-beta1. The embedded window cannot complete passkeys,
    # security keys or Windows Hello, which is where most sign-in reports on 3.x
    # ended up; the system browser can, and it reuses the browser session the admin
    # already has. The cost is one requirement on a CUSTOM app registration: the
    # http://localhost redirect URI. Microsoft's Graph PowerShell application, the
    # default, already allows it.
    Add-SettingsObject -Title "Use system browser for login" -Key "UseSystemBrowser" -Type "Boolean" -DefaultValue $true `
        -Description "Sign in using the default web browser instead of the embedded view or WAM. Enables passkey/FIDO2 login and browser extensions. Takes precedence over WAM. Custom app registrations need the http://localhost redirect URI (Mobile and desktop applications). Turn off to use the embedded window. Requires app restart." `
        -Section "MSAL"

    Add-SettingsObject -Title "Enable Continuous Access Evaluation (CAE)" -Key "EnableCAE" -Type "Boolean" -DefaultValue $true `
        -Description "Adds the 'cp1' client capability so Entra ID can revoke this session in near-real-time when admin policies change. Takes effect after an app restart / fresh login." `
        -Section "MSAL"

    # MSAL logging is implemented via a compiled C# bridge (see Enable-MSALLogging).
    # Earlier attempts used a PowerShell scriptblock as the LogCallback, but MSAL invokes
    # that delegate from background threads with no PS Runspace attached, which throws
    # PSInvalidOperationException and crashes the process. The C# bridge avoids that.
    Add-SettingsObject -Title "MSAL logging" -Key "MSALEnableLogging" -Type "Boolean" -DefaultValue $true `
        -Description "Write MSAL's internal log messages to msal.log under %LOCALAPPDATA%\IntuneManagement. Useful for diagnosing authentication failures." `
        -Section "MSAL"

    Add-SettingsObject -Title "MSAL log level" -Key "MSALLogLevel" -Type "List" -DefaultValue "Info" `
        -ItemsSource @(
            [PSCustomObject]@{ Name = "Error";   Value = "Error"   },
            [PSCustomObject]@{ Name = "Warning"; Value = "Warning" },
            [PSCustomObject]@{ Name = "Info";    Value = "Info"    },
            [PSCustomObject]@{ Name = "Verbose"; Value = "Verbose" }
        ) `
        -Description "Minimum severity of MSAL log messages to capture. Verbose is noisy." `
        -Section "MSAL"

    # Off by default and deliberately opt-in: MSAL's PII messages carry user names,
    # tenant/object ids and token details. The reason to turn it on is broker
    # diagnostics - MSALRuntime redacts its own error text as literal '(pii)' unless
    # PII logging is enabled, which is why a failed WAM silent refresh logs
    # "Error Message: (pii)" and tells you nothing. Enabling this unredacts both the
    # msal.log entries and the WAM exception message the app logs.
    Add-SettingsObject -Title "MSAL logging: include personal data (PII)" -Key "MSALEnablePiiLogging" -Type "Boolean" -DefaultValue $false `
        -Description "Include MSAL's PII log messages in msal.log and unredact broker/WAM error text (logged as '(pii)' otherwise). The log will then contain user names, tenant and object ids and token details - review it before sharing. Requires app restart." `
        -Section "MSAL"

    Add-SettingsObject -Title "Use MsalCacheHelper library" -Key "MSALUseCacheHelperLib" -Type "Boolean" -DefaultValue $true `
        -Description "Use Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper for the token cache. Provides cross-process file locking and is the supported library. Disable to fall back to the legacy TokenCacheHelperEx (DPAPI, process-local lock)." `
        -Section "MSAL"

    # MSAL 4.66+ supports forcing a regional ESTS endpoint. Reduces auth latency for
    # users in non-default Azure regions. Empty = let MSAL pick automatically.
    Add-SettingsObject -Title "MSAL region (optional)" -Key "MSALRegion" -Type "String" -DefaultValue "" `
        -Description "Force a regional ESTS endpoint (e.g. 'westeurope'). Leave empty for automatic. Applied via MSAL_FORCE_REGION; requires app restart." `
        -Section "MSAL"

    # When on, each AcquireToken call gets a fresh correlation Guid that's logged so
    # support tickets can be traced through MSAL logs and tenant audit logs.
    Add-SettingsObject -Title "Log MSAL correlation IDs" -Key "MSALCorrelationLogging" -Type "Boolean" -DefaultValue $false `
        -Description "Generate and log a correlation Guid per AcquireToken call. Useful when working with Microsoft support; off by default to keep logs quiet." `
        -Section "MSAL"

    # ToDo: when CS\HttpFactoryWithProxy.cs is wired into the PublicClientApplicationBuilder
    # (via .WithHttpClientFactory), add an "MSALDisableHttp2" boolean setting here and have
    # the factory return an HttpClient that pins HttpRequestMessage.Version to 1.1.
    # MSAL 4.73+ defaults to HTTP/2 for token requests; some inspection appliances break.
    # Until the factory is actually attached, exposing the setting would be misleading.

    # NOTE: the MSAL DLLs are deliberately NOT loaded here. Initialize-MSALPrereq
    # runs on first actual MSAL use (New-MSALApp and friends), so module import and
    # OAuth/MgGraph-only sessions skip the DLL load + Roslyn compile entirely.

    # Register this implementation in the multi-provider auth core. -SetActive makes
    # MSAL the default until the ActiveAuthProvider setting is applied at AppInitialized.
    # Invoke-MSGraphAPI routes through the registered provider (resolved per call by
    # TokenId), so this registration is on the live path, not just for discovery.
    # Registration is DLL-free: [AuthenticationMSAL]::new() and Initialize() touch no
    # MSAL types.
    if(Get-Command -Name Register-AuthProvider -ErrorAction SilentlyContinue) {
        try {
            Register-AuthProvider -Provider ([AuthenticationMSAL]::new()) -SetActive
        }
        catch {
            Write-LogError "Failed to register AuthenticationMSAL provider" $_.Exception
        }
    }
}

# Ensure the MSAL runtime is loaded, loading it on first call. Returns $true when the
# MSAL types are available. Safe to call from every MSAL entry point - the flag makes
# repeat calls O(1). The WAM / system-browser / region settings are resolved here (not
# at module init) so they reflect the settings at actual load time; all their consumers
# (Add-MSALPrereq's DLL list, New-MSALApp's builder options) run strictly downstream.
function Initialize-MSALPrereq
{
    if($script:MSALPrereqLoaded) { return $true }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $script:MSALUseWAM = Get-SettingValue "UseWAM"
    if($script:MSALUseWAM -and $PSVersionTable.PSVersion.Major -lt 7) {
        Write-Log "WAM is only supported in PowerShell 7 and later. Disabling WAM" 2
        $script:MSALUseWAM = $false
    }
    # WAM (the Web Account Manager broker) is a Windows-only component backed by
    # msalruntime.dll. There's no equivalent on Linux/macOS, so force it off and
    # fall back to the system-browser interactive flow.
    if($script:MSALUseWAM -and -not $script:IsWindowsOS) {
        Write-Log "WAM is only supported on Windows. Disabling WAM" 2
        $script:MSALUseWAM = $false
    }

    # System-browser interactive login (on by default). The default browser handles
    # the whole sign-in, so passkeys/FIDO2 and browser extensions work there - the
    # WAM pane and the embedded WebView don't reliably surface OS passkeys. Wins
    # over WAM when both are enabled. On Linux/macOS the system browser is already
    # the only interactive option, so the setting only changes behavior on Windows.
    $script:MSALUseSystemBrowser = (Get-SettingValue "UseSystemBrowser") -eq $true
    if($script:MSALUseSystemBrowser -and $script:MSALUseWAM) {
        Write-Log "Both UseWAM and UseSystemBrowser are enabled - the system browser takes precedence for interactive login" 2
        $script:MSALUseWAM = $false
    }

    # MSAL 4.66+ regional ESTS opt-in. MSAL reads the env var when the app is built,
    # so setting it here (before the first New-MSALApp) is early enough.
    $msalRegion = Get-SettingValue "MSALRegion" ""
    if($msalRegion) {
        $env:MSAL_FORCE_REGION = $msalRegion
        Write-LogDebug "MSAL region forced to '$msalRegion' via MSAL_FORCE_REGION"
    }

    Add-MSALPrereq

    # Only flag success when the core MSAL type actually resolved - a failed load
    # (missing Bin folder, blocked DLLs) leaves the flag unset so the next call can
    # retry after the user fixes the environment.
    if("Microsoft.Identity.Client.PublicClientApplicationBuilder" -as [type]) {
        $script:MSALPrereqLoaded = $true
        $sw.Stop()
        Write-Log "MSAL runtime loaded on first use ($([Math]::Round($sw.Elapsed.TotalMilliseconds))ms)"
    }
    else {
        Write-Log "MSAL runtime failed to load - MSAL authentication is unavailable (see log above)" 3
    }

    return $script:MSALPrereqLoaded
}

# Cheap pre-check for AuthenticationMSAL.TryResumeSession: is there plausibly a
# session to resume? Runs BEFORE the MSAL runtime is loaded (settings + file probes
# only), so fresh no-cache startups skip the DLL load entirely - the MgGraph
# provider's mg.authrecord.json probe is the precedent. Reads settings via
# Get-SettingValue rather than $script:MSALUseWAM, which is only resolved inside
# Initialize-MSALPrereq.
function Test-MSALResumeLikely
{
    # An OAuth-only deployment may delete Bin/ (README, Headless): nothing to load, nothing to resume.
    if(-not (Test-Path -LiteralPath (Get-MSALBinariesFolder) -PathType Container)) {
        Write-LogDebug "MSAL binaries not present under Bin/ - MSAL resume skipped (OAuth-only deployment?)"
        return $false
    }

    # The silent path matches cached accounts against the persisted last-user id;
    # without it there is nothing to resume.
    $lastUser = Get-SettingStoreValue "" "LastLoggedOnUserId"
    if(-not $lastUser) { return $false }

    # File-cache probe (current name + the legacy typo name; see Set-TokenCache).
    foreach($name in @("msalcache.bin3", "msalcahce.bin3")) {
        if(Test-Path (Join-Path $script:AppDataFolder $name)) { return $true }
    }

    # No file cache: the WAM broker can still silently reissue a token without one,
    # so allow the attempt when WAM would actually be active for this session.
    $wamActive = (Get-SettingValue "UseWAM") -and
                 ($PSVersionTable.PSVersion.Major -ge 7) -and
                 $script:IsWindowsOS -and
                 -not ((Get-SettingValue "UseSystemBrowser") -eq $true)
    return [bool]$wamActive
}

function Add-MSALPrereq
{
    $MSALBin  = Get-MSALBinariesFolder
    $DLLFiles = @()

    if($PSVersionTable.PSVersion.Major -lt 7) {

        # .NET Framework binds assembly references to the EXACT version, and the
        # bundled DLLs are not always the version MSAL.NET was compiled against
        # (Microsoft.Identity.Client 4.84 references IdentityModel.Abstractions
        # 8.14 while the bundle ships 8.18) - without a redirect every MSAL call
        # on PS5.1 dies with FileNotFoundException for the referenced version.
        # PS7/.NET rolls forward automatically. Register a C# AssemblyResolve
        # handler (NOT a scriptblock delegate - the event can fire on threads
        # with no runspace) that redirects Microsoft.Identity* requests to the
        # already-loaded assembly.
        if(-not ("IMAssemblyRedirect" -as [type])) {
            Add-Type -TypeDefinition @"
public static class IMAssemblyRedirect
{
    private static bool _registered;
    public static void Register()
    {
        if(_registered) { return; }
        _registered = true;
        System.AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
        {
            var requested = new System.Reflection.AssemblyName(args.Name).Name;
            if(!requested.StartsWith("Microsoft.Identity") &&
               !requested.StartsWith("Microsoft.Bcl.") &&
               !requested.StartsWith("System.Text.") &&
               !requested.StartsWith("System.Memory") &&
               !requested.StartsWith("System.Buffers") &&
               !requested.StartsWith("System.Numerics.Vectors") &&
               !requested.StartsWith("System.Runtime.CompilerServices.Unsafe") &&
               !requested.StartsWith("System.Threading.Tasks.Extensions")) { return null; }
            foreach(var asm in System.AppDomain.CurrentDomain.GetAssemblies())
            {
                if(asm.GetName().Name == requested) { return asm; }
            }
            return null;
        };
    }
}
"@
        }
        [IMAssemblyRedirect]::Register()
    }

    if($PSVersionTable.PSVersion.Major -lt 7) {
        # MSAL.NET 4.6x+ parses JSON with System.Text.Json on every target.
        # .NET Framework has none of this chain inbox, so the PS5 bundle ships
        # it (net461/netstandard2.0 builds); dependencies before consumers.
        foreach($netfxDep in @('System.Buffers.dll','System.Numerics.Vectors.dll',
                'System.Runtime.CompilerServices.Unsafe.dll','System.Memory.dll',
                'System.Threading.Tasks.Extensions.dll','Microsoft.Bcl.AsyncInterfaces.dll',
                'System.Text.Encodings.Web.dll','System.Text.Json.dll')) {
            $DLLFiles += [IO.FileInfo](Join-Path $MSALBin $netfxDep)
        }
    }
    $DLLFiles += [IO.FileInfo](Join-Path $MSALBin "Microsoft.IdentityModel.Abstractions.dll")
    $DLLFiles += [IO.FileInfo](Join-Path $MSALBin "Microsoft.Identity.Client.dll")
    # Desktop.dll provides the embedded WebView2 (and WebBrowser fallback) so non-loopback
    # redirect URIs like ".../oauth2/nativeclient" still work for AcquireTokenInteractive.
    # MSAL.NET ~4.79+ removed the built-in WinForms fallback; without this DLL plus
    # WithDesktopFeatures() in New-MSALApp the system browser is the only option and it
    # rejects anything that isn't http://localhost. Loaded unconditionally on Windows.
    $DLLFiles += [IO.FileInfo](Join-Path $MSALBin "Microsoft.Identity.Client.Desktop.dll")
    # MsalCacheHelper (the supported cross-platform token cache — DPAPI on Windows,
    # libsecret keyring on Linux, Keychain on macOS) lives in the Extensions.Msal
    # assembly. On Windows it's loaded with the WAM trio below; off Windows the
    # legacy DPAPI TokenCacheHelperEx can't run, so load Extensions.Msal there so
    # Set-TokenCache has a working helper.
    if($script:MSALUseWAM -or -not $script:IsWindowsOS) {
        $DLLFiles += [IO.FileInfo](Join-Path $MSALBin "Microsoft.Identity.Client.Extensions.Msal.dll")
    }
    if($script:MSALUseWAM) {
        $DLLFiles += [IO.FileInfo](Join-Path $MSALBin "Microsoft.Identity.Client.Broker.dll")
        $DLLFiles += [IO.FileInfo](Join-Path $MSALBin "Microsoft.Identity.Client.NativeInterop.dll")
    }

    $DLLFiles | ForEach-Object {
        $dllFile = $_
        # ToDo: Unblock files
        if($_.Exists) {
            try {
                [void][System.Reflection.Assembly]::LoadFrom($_.FullName)
                Write-Log "Loaded $($_.Name) version $($_.VersionInfo.FileVersion)"
            }
            catch [System.IO.FileLoadException] {
                # Common after an MSAL bundle upgrade: another PS host (VSCode PS extension,
                # other terminal) has the DLL mapped. The new file on disk can't be loaded
                # because a different assembly identity is already bound in this AppDomain.
                $loadedFile = [Appdomain]::CurrentDomain.GetAssemblies() | Where-Object Location -like "*\$($dllFile.Name)"
                if($loadedFile) {
                    $loadedFI = [IO.FileInfo]($loadedFile.Location)
                    Write-Log "Failed to load $($dllFile.Name) version $($dllFile.VersionInfo.FileVersion). A different version is already mapped: $($loadedFI.FullName) version $($loadedFI.VersionInfo.FileVersion). Restart your PowerShell session if you just upgraded the MSAL bundle." 2
                }
                else {
                    Write-LogError "Failed to load $($dllFile.Name) version $($dllFile.VersionInfo.FileVersion). The file may be locked by another process; restart your PowerShell session." $_.Exception
                }
            }
            catch {
                $loadedFile = [Appdomain]::CurrentDomain.GetAssemblies() | Where-Object Location -like "*\$($dllFile.Name)"
                if($loadedFile) {
                    $loadedFI = [IO.FileInfo]($loadedFile.Location)
                    Write-Log "Failed to load $($dllFile.Name) version $($dllFile.VersionInfo.FileVersion). File already loaded: $($loadedFI.FullName) version $($loadedFI.VersionInfo.FileVersion)" 2
                }
                else {
                    Write-LogError "Failed to load $($dllFile.Name) version $($dllFile.VersionInfo.FileVersion)" $_.Exception
                }
            }
        }
        else {
            Write-LogError "Microsoft.Identity file not found: $($_.FullName)"
        }
    }

    # TokenCacheHelperEx is the legacy DPAPI-backed cache helper. DPAPI
    # (System.Security.Cryptography.ProtectedData) is Windows-only and Set-TokenCache
    # never uses this helper off Windows (it uses MsalCacheHelper there), so skip the
    # compile entirely — it would only fail/log noise on Linux/macOS.
    if ($script:IsWindowsOS -and -not ("TokenCacheHelperEx" -as [type]))
    {
        [System.Collections.Generic.List[string]] $RequiredAssemblies = New-Object System.Collections.Generic.List[string]

        foreach($file in $DLLFiles) {
            $RequiredAssemblies.Add($file.FullName)
        }
        $RequiredAssemblies.Add('System.Security.dll')
        $RequiredAssemblies.Add('mscorlib.dll')
        if($PSVersionTable.PSVersion.Major -ge 7) {
            $RequiredAssemblies.Add('System.Security.Cryptography.ProtectedData.dll')
        }
        $RequiredAssemblies.Add('System.Threading.dll')

        try
        {
            # 3>$null suppresses csc CS1701 "Assuming assembly reference matches identity"
            # warnings that the upgraded MSAL.NET 4.8x DLLs trigger on .NET 9 hosts —
            # benign type-forwarding chatter (System.Runtime 8.0 -> 9.0). -IgnoreWarnings
            # only stops Add-Type throwing; it doesn't silence the warning text itself.
            Add-Type -Path (Join-Path $script:AppRootFolder "CS/TokenCacheHelperEx.cs") -ReferencedAssemblies $RequiredAssemblies -IgnoreWarnings 3>$null
        }
        catch
        {
            Write-LogError "Failed to compile TokenCacheHelperEx. The access token will not be cached. Check write access to the CS folder and ASR policies" $_.Exception
        }
    }

    # Guarded: Add-Type throws on a duplicate type name, so a re-entry after a partial
    # first attempt (e.g. a DLL failed to load) must not recompile MSALMethods.
    if($script:MSALUseWAM -and -not ("MSALMethods" -as [type])) {
        Write-Log "WAM is enabled. Adding WAM required native methods"
        [System.Collections.Generic.List[string]] $RequiredAssemblies = New-Object System.Collections.Generic.List[string]

        foreach($file in $DLLFiles) {
            $RequiredAssemblies.Add($file.FullName)    
        }
        $RequiredAssemblies.Add('mscorlib.dll')
        $RequiredAssemblies.Add('System.dll')
        # Import necessary methods from user32.dll and kernel32.dll
        Add-Type @"
            using System;
            using System.Runtime.InteropServices;
            using Microsoft.Identity.Client;
            //using Microsoft.Identity.Client.Desktop;
            using Microsoft.Identity.Client.Broker;

            public class MSALMethods
            {
                enum GetAncestorFlags {
                    GetParent = 1,
                    GetRoot = 2,
                    GetRootOwner = 3
                }

                [DllImport("user32.dll", ExactSpelling = true)]
                public static extern IntPtr GetAncestor(IntPtr hwnd, int flags);

                [DllImport("kernel32.dll")]
                public static extern IntPtr GetConsoleWindow();

                // This is your window handle!
                public static  IntPtr GetConsoleOrTerminalWindow()
                {
                    IntPtr consoleHandle = GetConsoleWindow();
                    IntPtr handle = GetAncestor(consoleHandle, (int)GetAncestorFlags.GetRootOwner );
                    
                    return handle;
                }

                public static void AddParentActivirtyOrWindow(PublicClientApplicationBuilder appBuilder)
                {
                    appBuilder.WithParentActivityOrWindow(GetConsoleOrTerminalWindow);
                }

                public static void AddWithBroker(PublicClientApplicationBuilder appBuilder, BrokerOptions options)
                {
                    appBuilder.WithBroker(options);
                }
            }

"@ -ReferencedAssemblies $RequiredAssemblies -IgnoreWarnings 3>$null
    }
}

function New-MSALApp {
    param(
        $EntraAppInfo = (Get-EntraApp),
        $Environment = (Get-EntraEnvironment)
    )

    # Lazy-load choke point: every MSAL flow (interactive connect, silent resume,
    # cached-account enumeration, UI click handlers) funnels through here. Throw
    # rather than return $null - callers would NRE confusingly at GetAccountsAsync.
    if(-not (Initialize-MSALPrereq)) {
        throw "MSAL libraries could not be loaded from Bin/MSAL_PS5|PS7 - see log for details."
    }

    $tenant = ?? $EntraAppInfo.TenantId "organizations"

    if($EntraAppInfo.Authority)
    {
        $authority = $EntraAppInfo.Authority
    }
    else
    {
        $authority = "https://$($Environment.URL)/$tenant/"
    }

    $msalApp = $script:MSALApps | Where-Object { $_.ClientId -eq  $entraAppInfo.ClientID  -and (-not $entraAppInfo.RedirectUri -or $_.AppConfig.RedirectUri -eq $entraAppInfo.RedirectUri)}
    if(-not $msalApp) {
        $appBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($entraAppInfo.ClientID)
        [void]$appBuilder.WithAuthority($authority)

        $redirectUri = $entraAppInfo.RedirectUri
        # Linux/macOS have no embedded WebView, so MSAL.NET drives interactive auth
        # through the system browser, which only accepts a loopback redirect. The
        # nativeclient / OOB redirects used on Windows can't work there, so unless the
        # caller already configured a loopback URI, substitute http://localhost (MSAL
        # picks a free port). First-party public clients such as Microsoft Graph
        # PowerShell permit loopback redirects. The same applies when the user opts
        # into the system browser on Windows (UseSystemBrowser) - passkey/FIDO2 login
        # works in a real browser but not in the WAM pane / embedded WebView.
        if((-not $script:IsWindowsOS -or $script:MSALUseSystemBrowser) -and ($redirectUri -notmatch '^http://localhost')) {
            $redirectUri = "http://localhost"
        }
        if ($redirectUri) { [void]$appBuilder.WithRedirectUri($redirectUri) }

        # Enable the embedded WebView (WebView2 with WebBrowser fallback) when not using
        # WAM. Required so AcquireTokenInteractive accepts the legacy nativeclient redirect
        # URI used by Microsoft Graph PowerShell and similar first-party apps; MSAL.NET
        # 4.79+ otherwise defaults to the system browser which only allows http://localhost.
        # WAM owns its own UI so we skip the call there; the extension is Windows-only, and
        # UseSystemBrowser wants that browser fallback anyway (loopback redirect set above).
        # It must be WithWindowsEmbeddedBrowserSupport and NOT WithDesktopFeatures: the
        # latter ALSO switches the WAM broker on (MSAL 4.88 leaves IsBrokerEnabled = True,
        # and is [Obsolete] pointing at WithBroker for that reason). As this branch runs
        # only when UseWAM is OFF, that silently re-enabled the broker the user had just
        # disabled, and startup silent-resume kept failing with an opaque "WAM Error (pii)".
        if(-not $script:MSALUseWAM -and -not $script:MSALUseSystemBrowser -and $script:IsWindowsOS) {
            try {
                [void][Microsoft.Identity.Client.Desktop.DesktopExtensions]::WithWindowsEmbeddedBrowserSupport($appBuilder)
            }
            catch {
                Write-LogDebug "WithWindowsEmbeddedBrowserSupport unavailable: $($_.Exception.Message)"
            }
        }

        [void] $appBuilder.WithClientName("IntuneManagement")
        # WithClientVersion expects a string; pass [Version].ToString() explicitly so we
        # don't rely on PowerShell's coercion (which silently broke once in 4.x history).
        [void] $appBuilder.WithClientVersion($PSVersionTable.PSVersion.ToString())

        # Continuous Access Evaluation: lets Entra ID invalidate the session near-real-time
        # when admin policies change. Capability "cp1" signals CAE-awareness to the IdP.
        if((Get-SettingValue "EnableCAE") -ne $false) {
            try { [void]$appBuilder.WithClientCapabilities([string[]]@("cp1")) }
            catch { Write-LogDebug "WithClientCapabilities not supported by this MSAL version" }
        }

        # MSAL internal logging via compiled C# bridge. Lazy: only enabled when the user
        # has the setting on. Failures are swallowed so a bad logger never breaks auth.
        if((Get-SettingValue "MSALEnableLogging") -ne $false) {
            try { Enable-MSALLogging $appBuilder }
            catch { Write-LogDebug "Failed to attach MSAL logging: $($_.Exception.Message)" }
        }

        if($script:MSALUseWAM) {
            [MSALMethods]::AddParentActivirtyOrWindow($appBuilder)

            $options = [Microsoft.Identity.Client.BrokerOptions]::new([Microsoft.Identity.Client.BrokerOptions+OperatingSystems]::Windows)
            $options.Title = "Intune Manager"
            if((Get-SettingValue "WAMListOSAccounts") -ne $false) {
                try { $options.ListOperatingSystemAccounts = $true }
                catch { Write-LogDebug "BrokerOptions.ListOperatingSystemAccounts not available in this MSAL version" }
            }

            [MSALMethods]::AddWithBroker($appBuilder, $options)
        }

        $msalApp = $appBuilder.Build()

        $script:MSALApps += $msalApp

        Set-TokenCache $msalApp
    }

    return $msalApp
}

# Attaches the MSAL LogCallback. Critical detail: the callback MUST be a pure managed
# method, NOT a PowerShell scriptblock-derived delegate. MSAL invokes the callback from
# background worker threads that have no PowerShell Runspace attached; a scriptblock
# delegate throws PSInvalidOperationException there and crashes the process.
#
# Strategy: compile a tiny C# class via Add-Type whose static method matches the
# LogCallback signature, then bind it as a delegate. The method writes directly to a
# file using BCL APIs only — no PowerShell runtime touched.
function Enable-MSALLogging
{
    param($AppBuilder)

    # Compile the bridge once per session. Re-Add-Type'ing the same name throws.
    if(-not ("MSALLogBridge" -as [type])) {
        $msalAsm = [AppDomain]::CurrentDomain.GetAssemblies() |
            Where-Object { $_.GetName().Name -eq "Microsoft.Identity.Client" } |
            Select-Object -First 1
        if(-not $msalAsm) {
            Write-LogDebug "MSAL assembly not loaded; cannot enable MSAL logging"
            return
        }

        # Use short-name refs the same way TokenCacheHelperEx.cs is compiled — PowerShell's
        # Add-Type resolves these to the correct reference-assembly facades on both
        # Desktop and Core. Tried fully-qualified paths to System.Private.CoreLib first;
        # Roslyn rejected them because LogLevel inherits from Enum, which is type-forwarded
        # from the legacy mscorlib identity. The short-name 'mscorlib.dll' satisfies that
        # lookup on .NET Core via netstandard.
        $refs = @('mscorlib.dll', 'System.dll', 'System.IO.dll', 'System.Threading.dll', $msalAsm.Location)

        # Bridge implementation:
        # - Configure() sets the destination file + min level (called from PS on each session)
        # - OnLog() is what MSAL invokes; runs entirely in managed code, lock-protected file write
        # - Size-based rotation at 5 MB to prevent unbounded growth
        # NOTE: comparison uses the enum directly (no int cast) to avoid pulling in the
        # full Enum surface in the C# compile.
        Add-Type -ReferencedAssemblies $refs -TypeDefinition @"
using System;
using System.IO;
using Microsoft.Identity.Client;

public static class MSALLogBridge
{
    private static readonly object _lock = new object();
    private static string _logFile;
    private static LogLevel _minLevel = LogLevel.Info;
    private const long MaxBytes = 5L * 1024 * 1024;

    // Set from PowerShell (MSALEnablePiiLogging). Kept as a settable field rather than a
    // Configure() parameter so a stale bridge type - the type is compiled once per
    // process and cannot be re-Add-Type'd - only costs PII detail, not all logging.
    public static volatile bool AllowPii = false;

    public static void Configure(string logFilePath, LogLevel minLevel)
    {
        lock (_lock)
        {
            _logFile = logFilePath;
            _minLevel = minLevel;
        }
    }

    public static void OnLog(LogLevel level, string message, bool containsPii)
    {
        if (containsPii && !AllowPii) return;
        if (level > _minLevel) return;
        if (string.IsNullOrEmpty(_logFile)) return;

        try
        {
            lock (_lock)
            {
                // Rotate when file grows past MaxBytes. Best-effort; failures swallowed.
                try
                {
                    var fi = new FileInfo(_logFile);
                    if (fi.Exists && fi.Length > MaxBytes)
                    {
                        var old = _logFile + ".old";
                        if (File.Exists(old)) File.Delete(old);
                        File.Move(_logFile, old);
                    }
                }
                catch { }

                File.AppendAllText(_logFile,
                    string.Format("[{0:yyyy-MM-dd HH:mm:ss.fff}] [{1,-7}] {2}{3}{4}",
                        DateTime.Now, level, containsPii ? "[PII] " : "", message, Environment.NewLine));
            }
        }
        catch { /* logger must never throw back into MSAL */ }
    }
}
"@ -IgnoreWarnings 3>$null
    }

    if(-not ("MSALLogBridge" -as [type])) {
        Write-LogDebug "MSALLogBridge type unavailable after Add-Type; skipping MSAL logging"
        return
    }

    # Resolve log file location and severity threshold from settings.
    $appData = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) "IntuneManagement"
    if(-not (Test-Path -LiteralPath $appData)) {
        [void](New-Item -ItemType Directory -Path $appData -Force -ErrorAction SilentlyContinue)
    }
    $logFile = Join-Path $appData "msal.log"

    $levelName = Get-SettingValue "MSALLogLevel" "Info"
    $msalLevel = switch ($levelName) {
        "Error"   { [Microsoft.Identity.Client.LogLevel]::Error   }
        "Warning" { [Microsoft.Identity.Client.LogLevel]::Warning }
        "Info"    { [Microsoft.Identity.Client.LogLevel]::Info    }
        "Verbose" { [Microsoft.Identity.Client.LogLevel]::Verbose }
        default   { [Microsoft.Identity.Client.LogLevel]::Info    }
    }

    [MSALLogBridge]::Configure($logFile, $msalLevel)

    # PII: two switches have to agree. The bridge drops containsPii messages unless
    # AllowPii is set, and MSAL itself only produces them when WithLogging is built with
    # enablePiiLogging - that same flag is what makes MSALRuntime hand over its real
    # broker error text instead of the literal '(pii)' placeholder.
    $enablePii = (Get-SettingValue "MSALEnablePiiLogging") -eq $true
    if($enablePii) {
        try { [MSALLogBridge]::AllowPii = $true }
        catch {
            # Stale bridge compiled before AllowPii existed (module re-imported in a
            # long-running session). Non-PII logging still works; restart to get PII.
            Write-Log "MSAL PII logging requested but the log bridge in this process predates it. Restart the app to capture PII messages." 2
            $enablePii = $false
        }
    }

    # Build a delegate that points at the static method. This is the key bit: the
    # delegate target is a CLR method, so MSAL can invoke it from any thread without
    # touching PowerShell.
    try {
        $method   = [MSALLogBridge].GetMethod("OnLog")
        $callback = [Delegate]::CreateDelegate([Microsoft.Identity.Client.LogCallback], $method)
        [void]$AppBuilder.WithLogging($callback, $msalLevel, $enablePii, $false)
        Write-LogDebug "MSAL logging enabled -> $logFile (level=$levelName, pii=$enablePii)"
        if($enablePii) {
            Write-Log "MSAL PII logging is ON - $logFile will contain user names, tenant/object ids and token details" 2
        }
    }
    catch {
        Write-LogError "Failed to bind MSAL log callback" $_.Exception
    }
}

function Set-TokenCache
{
    param($MsalApp)

    if(-not (Get-SettingValue "CacheMSALToken")) { return }

    $script:AppDataFolder = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) "IntuneManagement"

    # Cache filename. The legacy name "msalcahce.bin3" was a typo carried over from
    # very early versions; current installs have an existing file under that name.
    # We use the corrected "msalcache.bin3" going forward and do a one-shot rename
    # if the old file exists and the new one doesn't, so users don't have to re-auth
    # after this change.
    $cacheFileName = "msalcache.bin3"
    $oldFile = Join-Path $script:AppDataFolder "msalcahce.bin3"
    $newFile = Join-Path $script:AppDataFolder $cacheFileName
    if((Test-Path -LiteralPath $oldFile) -and -not (Test-Path -LiteralPath $newFile)) {
        try {
            Rename-Item -LiteralPath $oldFile -NewName $cacheFileName -ErrorAction Stop
            Write-Log "Migrated MSAL cache file: msalcahce.bin3 -> $cacheFileName"
        }
        catch {
            Write-LogDebug "MSAL cache rename failed ($($_.Exception.Message)) - falling back to legacy name"
            $cacheFileName = "msalcahce.bin3"
        }
    }

    # Roamed/UNC profiles are not safe — file-lock semantics over SMB are unreliable
    # (CrossPlatLock relies on FileShare.None / DeleteOnClose). %LOCALAPPDATA% is
    # non-roaming on a standard install; warn if someone has redirected it.
    if($script:AppDataFolder -like "\\*") {
        Write-Log "MSAL cache folder appears to be on a network share ($script:AppDataFolder). Cross-process locks may behave unreliably." 2
    }

    $useLib = (Get-SettingValue "MSALUseCacheHelperLib") -ne $false

    # Preferred path: Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper. Provides
    # cross-process file locking and is the supported library. Toggle the
    # "Use MsalCacheHelper library" setting off to fall back to the legacy helper.
    if($useLib -and ("Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper" -as [type])) {
        try {
            if(-not $script:cacheHelper) {
                $storageBuilder = [Microsoft.Identity.Client.Extensions.Msal.StorageCreationPropertiesBuilder]::new($cacheFileName, $script:AppDataFolder)
                # Windows encrypts the cache file with DPAPI automatically. Linux and
                # macOS need an explicit keyring/keychain configuration or
                # MsalCacheHelper throws "persistence is not configured". Linux uses
                # libsecret (gnome-keyring / KWallet); macOS uses the Keychain.
                if(-not $script:IsWindowsOS) {
                    $attr1 = [System.Collections.Generic.KeyValuePair[string,string]]::new("Version", "1")
                    $attr2 = [System.Collections.Generic.KeyValuePair[string,string]]::new("Product", "IntuneManagement")
                    if($IsMacOS) {
                        [void]$storageBuilder.WithMacKeyChain("com.intunemanagement.tokencache", "IntuneManagement")
                    }
                    else {
                        [void]$storageBuilder.WithLinuxKeyring(
                            "com.intunemanagement.tokencache",
                            "default",
                            "IntuneManagement MSAL token cache",
                            $attr1,
                            $attr2)
                    }
                }
                $storageProperties = $storageBuilder.Build()
                $script:cacheHelper = [Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper]::CreateAsync($storageProperties).GetAwaiter().GetResult()

                # Round-trip write/read/clear test — surfaces DPAPI breakage (corrupt
                # master key after machine join change, broken user profile, etc.)
                # immediately rather than on the first auth attempt. Per MSAL docs,
                # this never fails on Windows when the user profile is healthy.
                try {
                    $script:cacheHelper.VerifyPersistence()
                    Write-LogDebug "MSAL cache persistence verified"
                }
                catch {
                    Write-LogError "MSAL cache persistence verification failed - tokens may not persist across sessions" $_.Exception
                    # Keep the helper registered; even if persistence is degraded the
                    # in-memory cache still works for the current session.
                }
            }
            if($script:cacheHelper) {
                $script:cacheHelper.RegisterCache($MsalApp.UserTokenCache)
                Write-LogDebug "MSAL cache registered via MsalCacheHelper ($cacheFileName)"
                return
            }
        }
        catch {
            Write-LogError "MsalCacheHelper init failed; falling back to TokenCacheHelperEx" $_.Exception
            # Drop through to the legacy fallback below.
        }
    }

    # Legacy fallback: hand-rolled DPAPI helper. Process-local lock; Windows only —
    # DPAPI (ProtectedData) throws PlatformNotSupportedException off Windows, so don't
    # register it there (a broken helper would fail every cache read/write and could
    # break auth). On non-Windows we just run with the in-memory cache for the session.
    if($script:IsWindowsOS -and ("TokenCacheHelperEx" -as [type])) {
        [TokenCacheHelperEx]::EnableSerialization($MsalApp.UserTokenCache, (Join-Path $script:AppDataFolder $cacheFileName))
        Write-LogDebug "MSAL cache registered via legacy TokenCacheHelperEx ($cacheFileName)"
    }
    else {
        Write-Log "No MSAL cache helper available. Tokens will not persist across sessions." 2
    }
}

function Get-GraphDomain
{
    # Resolve the Graph hostname (graph.microsoft.com / graph.microsoft.us / etc.)
    # for the given MSAL TokenId. Callers use the result to build absolute Graph
    # URLs (@odata.bind / @odata.id). Returning $null silently here produces broken
    # "https:///<path>" URLs at the call site, so we fall back to a provider-aware
    # default instead of just warning and returning $null.
    param([int]$Id = (Get-DefaultTokenId))

    # Registry first: the central token registry records the Cloud for EVERY
    # provider's tokens (OAuth / MgGraph tokens are not in $script:MSALTokens, so
    # the MSAL lookup below misses them - that was why OAuth cloud resolution fell
    # through to defaults).
    if($Id -gt 0) {
        $rec = $script:AuthTokens[$Id]
        if($rec -and $rec.Cloud) {
            $cloud = Get-CloudByValue $rec.Cloud
            if($cloud -and $cloud.GraphHost) { return $cloud.GraphHost }
        }
    }

    $token = $script:MSALTokens.Values | Where-Object Id -eq $id
    if($token -and $token.CloudEntry) {
        return $token.CloudEntry.GraphHost
    }

    # No MSAL token is normal for MgGraph. Derive the cloud from the active SDK
    # context before falling back to persisted defaults.
    try {
        if(Get-Command Get-MgContext -ErrorAction Ignore) {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if($ctx -and $ctx.Environment) {
                $cloud = $script:Clouds | Where-Object MgEnvironment -eq $ctx.Environment | Select-Object -First 1
                if($cloud -and $cloud.GraphHost) { return $cloud.GraphHost }
            }
        }
    } catch { }

    try {
        $lastCloud = Get-SettingStoreValue "" "LastLoggedOnCloud" ""
        if($lastCloud) {
            $cloud = Get-CloudByValue $lastCloud
            if($cloud -and $cloud.GraphHost) { return $cloud.GraphHost }
        }
    } catch { }

    $defaultCloud = Get-CloudByValue (Get-DefaultCloud)
    if($defaultCloud -and $defaultCloud.GraphHost) { return $defaultCloud.GraphHost }

    Write-LogDebug "Get-GraphDomain: no cloud context for Id $Id; falling back to graph.microsoft.com"
    return "graph.microsoft.com"
}


# True when a failure is only our own cancellation echoing back (ending a wait cancels
# the token source; MSAL reports that as canceled). That artifact used to mask the real
# reason - a timeout - in both the log and the returned failure.
function Test-MsalCancelFailure
{
    param($Failure)

    if(-not $Failure) { return $false }
    return (($Failure -is [System.OperationCanceledException]) -or
            ([string]$Failure.ErrorCode -eq "authentication_canceled"))
}

function Get-MsalAuthenticationToken
{
    # -Interactive / -DeviceCode mark the flows a user is actually waiting in front
    # of, and only those get the status overlay's Cancel button armed. A silent
    # acquire has nothing worth cancelling (it fails in milliseconds) and arming it
    # would reset - and could steal the click from - a cancel already armed by
    # whatever operation the refresh happened inside.
    param($AquireTokenObj, [switch]$DeviceCode, [switch]$Interactive)

    $errorInfo = $null
    $authResult = $null
    $authenticationFailure = $null

    try 
    {        
        $tokenSource = New-Object System.Threading.CancellationTokenSource

        # Decouple MSAL's async continuations from this thread's SynchronizationContext
        # before kicking off the call. WPF and Avalonia both install a dispatcher-backed
        # SynchronizationContext on the UI thread; MSAL's internal 'await's would capture
        # it and post their completion continuation back to THIS thread, where it only
        # runs if this exact thread pumps that dispatcher queue. On the WAM-broker-failed
        # -> interactive fallback path that continuation was getting stranded: the token
        # was actually acquired and written to the cache, but the Task never transitioned
        # to Completed, so the poll loop below waited out the whole timeout and a
        # successful login was misreported as cancelled (the reported hang/regression).
        # Clearing Current so the first (synchronous) await captures a null context makes
        # every MSAL continuation resume on the thread pool, so the Task completes
        # independently of UI pumping. ExecuteAsync returns at the first await, so the
        # window where Current is null is just its synchronous prefix; restore right after.
        # SynchronizationContext is a System (not UI-framework) type, so this stays
        # R12-clean in Internal/. MSAL's embedded/system web dialog does not rely on the
        # caller's SynchronizationContext (it hosts its own STA loop / out-of-proc browser).
        $prevSyncContext = [System.Threading.SynchronizationContext]::Current
        [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
        try
        {
            $taskAuthenticationResult = $AquireTokenObj.ExecuteAsync($tokenSource.Token) #.GetAwaiter().GetResult()
        }
        finally
        {
            [System.Threading.SynchronizationContext]::SetSynchronizationContext($prevSyncContext)
        }

        # Wall-clock timeout so a killed/closed interactive (MFA / WAM broker) window
        # can't leave the task pending forever and hang the app. On timeout we cancel
        # the token source; the task then completes as canceled/faulted and the parsing
        # below returns a null $authResult - the caller treats that as a failed login.
        # With the SynchronizationContext decoupling above, a genuinely successful login
        # now flips IsCompleted promptly (on a thread-pool continuation), so this timeout
        # only fires on a truly abandoned window rather than on slow-but-successful ones.
        # Three waits, three budgets. One shared 180 s cap was both too short for a human
        # doing MFA (a WAM picker opening behind the app died at exactly 180 s and
        # surfaced as a user cancel) and far too long for a silent acquire, where a wedged
        # broker call held the calling Graph request for three minutes.
        if($DeviceCode) {
            # Device-code sign-in is completed on another device, so the interactive
            # cap does not apply. Start with a generous cap and, once MSAL reports the
            # code, switch to the code's own expiry (below).
            $timeoutSec = 900
        }
        elseif($Interactive) {
            # A human is in front of this one (picker, MFA push, number matching, passkey)
            # and the Cancel button is armed below, so an abandoned window has an exit that
            # is not the timeout - which is what lets the cap be generous. The setting's
            # registered default governs; this literal only covers an unreadable store.
            $timeoutSec = 600
            try { $t = [int](Get-SettingValue "MSGraphInteractiveTimeoutSec"); if($t -gt 0) { $timeoutSec = $t } } catch { }
        }
        else {
            # Silent acquire / client credentials: nobody waiting, normally ~1 s. The cap
            # only stops a wedged broker or hung socket holding the calling Graph request.
            # Not tighter than 60 s - MSAL's own per-request HTTP timeout is ~30 s and it
            # retries, so a shorter outer cap could kill a legitimate retry.
            $timeoutSec = 60
        }
        $deadline = [DateTime]::UtcNow.AddSeconds($timeoutSec)
        $timedOut = $false
        $cancelled = $false
        $deviceCodeSurfaced = $false
        # Arm the status overlay's Cancel button for the whole wait. Without it the
        # only way out of an abandoned sign-in is the timeout above - 3 minutes for
        # an interactive browser login, 15 for device code. Cancelling the token
        # source is what actually ends the wait; the loop below also polls the
        # request flag so it breaks immediately.
        $cancellable = ($Interactive -eq $true -or $DeviceCode -eq $true)
        if($cancellable)
        {
            # The click is dispatched later from an Avalonia/WPF event callback.
            # Capture the source explicitly instead of relying on PowerShell's
            # dynamic scope still containing this function's local variable.
            $cancelTokenSource = $tokenSource
            Set-StatusCancelAction ({ $cancelTokenSource.Cancel() }.GetNewClosure())
            # Arming only records the action - the button has to be asked for. Device
            # code asks later, from Show-DeviceCodeInstruction, because its message
            # cannot be written until MSAL hands over the code.
            if(-not $DeviceCode)
            {
                Write-Status "Waiting for sign-in" "Complete the sign-in in the browser window" -CancelText "Cancel"
            }
        }
        try
        {
            # Poll on a tight cadence (50 ms). Completion is observed via the thread-pool
            # continuation (see the context decoupling above); the UI pump here is only to
            # keep any caller-thread-hosted auth UI (embedded webview) responsive.
            while (!$taskAuthenticationResult.IsCompleted)
            {
                # Device code: once MSAL hands us the user_code + verification_uri,
                # surface it once (clipboard + status overlay) and adopt its expiry.
                if ($DeviceCode -and -not $deviceCodeSurfaced -and ('IMMsalDeviceCodeHelper' -as [type]) -and [IMMsalDeviceCodeHelper]::HasResult)
                {
                    $deviceCodeSurfaced = $true
                    Show-DeviceCodeInstruction
                    try {
                        $exp = [IMMsalDeviceCodeHelper]::ExpiresOn
                        if($exp -and $exp.UtcDateTime -gt [DateTime]::UtcNow) { $deadline = $exp.UtcDateTime }
                    } catch { }
                }
                if ($cancellable -and (Test-StatusCancelRequested))
                {
                    # Request-StatusCancel has already cancelled the token source; this
                    # only stops the poll so the caller does not wait for the task to
                    # observe it. Silent acquires ignore the flag - it belongs to
                    # whatever operation armed it, not to them.
                    $cancelled = $true
                    Write-Log "Login cancelled by the user" 2
                    break
                }
                if ([DateTime]::UtcNow -gt $deadline)
                {
                    $timedOut = $true
                    # Not "timed out or was cancelled": a real cancel sets $cancelled and
                    # is reported separately. Conflating them hid our own cap.
                    Write-Log "Login timed out after $timeoutSec s - giving up" 2
                    $tokenSource.Cancel()
                    break
                }
                Invoke-UIPump
                Start-Sleep -Milliseconds 50
            }
        }
        catch 
        {
            Write-LogError "Failed to generate token" $_.Exception
        }
        finally
        {
            # Cancellation can complete the task while the loop is sleeping, so the
            # loop may exit on IsCompleted without entering its request check. Capture
            # the sticky request before Clear-StatusCancelAction resets it.
            if($cancellable -and (Test-StatusCancelRequested))
            {
                $cancelled = $true
            }
            # Disarm before disposing: a click arriving after this point must not
            # reach a disposed token source. The button goes with it - -CancelText ""
            # hides it without touching the status text the caller had.
            if($cancellable)
            {
                Clear-StatusCancelAction
                Write-Status -CancelText ""
            }
            if (-not $taskAuthenticationResult.IsCompleted)
            {
                $tokenSource.Cancel()
            }
            $tokenSource.Dispose()
        }

        ## Parse task results
        if ($taskAuthenticationResult.IsFaulted) 
        {
            # ToDo Check if: $taskAuthenticationResult.Exception.InnerException -is [Microsoft.Identity.Client.MsalUiRequiredException]
            if($taskAuthenticationResult.Exception.InnerException.ResponseBody)
            {
                try 
                {
                    $errorInfo = $taskAuthenticationResult.Exception.InnerException.ResponseBody | ConvertFrom-Json
                }
                catch { }
            }
            $authenticationFailure = ?? $taskAuthenticationResult.Exception.InnerException $taskAuthenticationResult.Exception
            # A timeout or Cancel click also faults the task with MSAL's cancellation
            # error. Logging that as "Failed to login" buried the actual reason under what
            # looks like an auth problem; the branch below reports both properly.
            if(($timedOut -or $cancelled) -and (Test-MsalCancelFailure $authenticationFailure))
            {
                Write-LogDebug "Task faulted with a cancellation error after the wait ended - reported below"
            }
            elseif($errorInfo.error_description)
            {
                Write-LogError "Failed to login. Error: $($errorInfo.error). Description: $($errorInfo.error_description)" 3
            }
            else
            {
                Write-LogError "Failed to login" (?? $taskAuthenticationResult.Exception.InnerException $taskAuthenticationResult.Exception)
            }
        }

        if ($timedOut -or $cancelled -or $taskAuthenticationResult.IsCanceled)
        {
            if($timedOut) {
                # A timeout is a failure, not a decision: keep it loud and keep the reason.
                # The old '-not $authenticationFailure' guard let the faulted task's own
                # cancellation error win, so the seconds - the detail that identifies our
                # cap rather than a user's click - never reached the caller or the log.
                $hint = if($Interactive) { " Raise the 'Interactive login timeout' setting if the sign-in needs longer." } else { "" }
                Write-Log "The login timed out after $timeoutSec s.$hint" 2
                if(-not $authenticationFailure -or (Test-MsalCancelFailure $authenticationFailure)) {
                    $authenticationFailure = "Login timed out after $timeoutSec s"
                }
            }
            # A user cancel is not a failure to report as an error, but the caller
            # still needs a typed reason so it can suppress AuthenticationFailed.
            elseif($cancelled) {
                Write-Log "The login was canceled by the user" 2
                # Preserve cancellation as a typed outcome so the outer provider can
                # distinguish a user's decision from an authentication failure and
                # avoid raising AuthenticationFailed.
                $authenticationFailure = [System.OperationCanceledException]::new("Login was cancelled by the user")
            }
            else {
                # Cancelled without our timeout or Cancel button firing - MSAL or the
                # broker ended it (e.g. the sign-in window was closed).
                Write-Log "The login was canceled" 2
            }
        }
        elseif ($taskAuthenticationResult.IsCompleted -and -not $taskAuthenticationResult.IsFaulted)
        {
            # Only read .Result on a task that finished successfully. Reading it on a
            # still-running task (e.g. after a timeout break, where Cancel() is async and
            # may not have landed yet) would BLOCK and re-introduce the hang.
            $authResult = $taskAuthenticationResult.Result
        }
    }
    catch 
    {
        $authenticationFailure = ?? $_.Exception.InnerException $_.Exception
        Write-LogError "Failed to authenticate" (?? $_.Exception.InnerException $_.Exception)
    }
    
    $authResult, $authenticationFailure
}

function Get-TenantList
{
    # Login-time population of the tenant list the profile popups read from
    # $script:AccessibleTenants. Best-effort by design: it is gated by the
    # "Get Tenant List" setting and must never fail a sign-in.
    #
    # The Azure Resource Manager call itself lives in Internal/EntraTenantList.ps1,
    # shared with Get-AccessibleTenant, so both paths use the same per-cloud
    # endpoint and report the same reason when the app registration lacks the
    # delegated Azure Service Management permission. This function used to hold a
    # second copy that hardcoded the public-cloud host (so the list could never
    # work on a sovereign cloud) and ended in catch{} (so a missing permission
    # looked exactly like a tenant with no other tenants).
    param($TokenInfo)

    if($TokenInfo.Tenants) {
        $script:AccessibleTenants = $TokenInfo.Tenants
        return
    }

    try
    {
        Write-Log "Get tenant list"

        # Reuse the app the user already authenticated with. Previously this function
        # built a brand-new PublicClientApplication and re-attached the token cache,
        # which doubled cache helper registrations and forced the user through an
        # extra silent fetch. The existing app already has the right authority,
        # redirect URI, cache, broker, logging, and CAE capability wired up.
        $app = $TokenInfo.App
        if(-not $app) {
            Write-LogDebug "Get-TenantList: TokenInfo.App missing; falling back to a fresh app builder"
            # Lazy-load guard (cheap flag check; the primary path reuses TokenInfo.App,
            # which means the runtime is already loaded).
            if(-not (Initialize-MSALPrereq)) {
                throw "MSAL libraries could not be loaded from Bin/MSAL_PS5|PS7 - see log for details."
            }
            $redirectUri = (Get-EntraApp $TokenInfo.Token.ClientId).RedirectURI
            $appBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($TokenInfo.EntraApp.ClientId)
            [void]$appBuilder.WithTenantId($TokenInfo.token.TenantId)
            if($redirectUri) { [void]$appBuilder.WithRedirectUri($redirectUri) }
            $app = $appBuilder.Build()
            Set-TokenCache $app
        }

        $result = Get-EntraAccessibleTenant -App $app -Account $TokenInfo.Token.Account `
                    -TenantId $TokenInfo.token.TenantId -Cloud $TokenInfo.Cloud `
                    -AllowInteractive:($script:MainAppStarted -eq $true)

        if($result -and @($result.Tenants).Count -gt 0) {
            $script:AccessibleTenants = $result.Tenants
            $TokenInfo.Tenants = $result.Tenants
        }
    }
    catch {
        # Still best-effort, but no longer silent: the reason reaches the log.
        Write-LogError "Get-TenantList failed" $_.Exception
    }
}

function Connect-EntraEnvironment {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    param(
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        $AppId,

        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ClientSecret,

        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate', Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ClientCertificate,

        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent', Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret', Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate', Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]        
        [string]$Environment,

        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', Position = 4, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', Position = 4, ValueFromPipelineByPropertyName = $true)]
        [switch]$AuthenticationBroker,
        
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$User = $null,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$UserId = $null,

        [Parameter(Mandatory = $false, ParameterSetName = 'Token', Position = 5, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent', Position = 5, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [switch]$ForceRefresh,

        [Parameter(Mandatory = $true, ParameterSetName = 'Token', Position = 6, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Int]$TokenId,

        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', Position = 7, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', Position = 7, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent', Position = 7, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret', Position = 7, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate', Position = 7, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', Position = 8, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', Position = 8, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent', Position = 8, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret', Position = 8, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate', Position = 8, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$Authority,

        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent', Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret', Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate', Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$RedirectUri,

        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent', Position = 6, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'Token', ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [switch]$ForceSilent,

        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent', Position = 6, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'Token', ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [switch]$ForceInteractive,

        # Suppress the AuthenticationFailed app event when a silent refresh returns no
        # token. Used by the near-expiry background pre-flight refresh (GetAccessToken):
        # while the existing access token is still valid, a silent-refresh miss (e.g. the
        # WAM broker failing on the refresh round-trip) is a non-event - the caller keeps
        # using the still-valid token, so tearing the session down / flipping the UI to
        # Sign-in and logging a failure on every call is a false alarm. When the token has
        # actually expired the caller does NOT pass this, so the miss stays loud and the
        # UI reverts to Sign-in as designed (see Invoke-MSGraphAPI header comment).
        [switch]$SuppressFailedEvent,

        # Device code (RFC 8628) instead of a browser popup / WAM broker when
        # the interactive fallback fires. Silent-from-cache is still attempted
        # first — device code only runs if there's no valid cached token and
        # -ForceSilent isn't set. Same shape as -ForceInteractive; mutually
        # exclusive with it in practice (device code wins if both are passed).
        [switch]$DeviceCode,

        [switch]$DefaultToken,

        # CAE claims challenge from a Graph 401 (WWW-Authenticate header). When the caller
        # propagates this back into Connect-EntraEnvironment, MSAL re-fronts the token
        # request to satisfy the challenge in the silent path; without it, a CAE-revoked
        # session would loop forever between silent + interactive flows.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$ClaimsChallenge,

        # Phase 3 (2026-05-22) — flat Cloud taxonomy. When set, drives the MSAL
        # authority selection AND records the choice for later silent refresh +
        # next-launch resume. Takes precedence over -Environment (which is the legacy
        # shape; -Environment is still accepted for backward compat). When neither is
        # set, falls back to Get-StartupCloudHint (per-tenant memory → LastLoggedOnCloud
        # → DefaultCloud setting).
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [ValidateSet("Public", "USGov", "USGovDOD", "China")]
        [string]$Cloud
    )

    # Resolve Cloud → legacy Environment up-front so the rest of the function (which
    # still keys off $Environment / $msalEnvironment) keeps working unchanged. We
    # remember the resolved Cloud separately ($resolvedCloud) because USGov vs
    # USGovDOD collapse into the same legacy -Environment value; we need the original
    # to persist correctly after auth.
    $resolvedCloud = $null
    if($Cloud) {
        $entry = Get-CloudByValue $Cloud
        if($entry) {
            $Environment = $entry.LegacyEnv
            $resolvedCloud = $Cloud
        }
    }
    elseif(-not $Environment) {
        # Caller didn't pin a cloud — use the per-tenant / last-logged-on / default chain.
        $hintCloud = Get-StartupCloudHint -TenantId $TenantId
        if($hintCloud) {
            $entry = Get-CloudByValue $hintCloud
            if($entry) {
                $Environment = $entry.LegacyEnv
                $resolvedCloud = $hintCloud
            }
        }
    }

    $TokenInfo = $null
    if($TokenID) {
        $TokenInfo = $script:MSALTokens.Values | Where-Object Id -eq $TokenID
    }

    if($TokenInfo) {
        # Non-interactive auth types bypass the MSAL public-client flow entirely.
        if($TokenInfo.AuthType -eq "BYO") {
            if($TokenInfo.Token -and $TokenInfo.Token.ExpiresOn -lt [DateTimeOffset]::UtcNow) {
                # Evict the expired BYO token so consumers don't keep using it and getting 401s.
                # Caller must re-run Connect-IntuneManagement with a fresh token.
                Write-Log "BYO token for tenant '$($TokenInfo.Token.TenantId)' has expired. Removing it. Call Connect-IntuneManagement with a new token." 2
                # Unregister first (hydrates snapshot) then drop the MSAL entry.
                Unregister-AuthToken -TokenId $TokenInfo.Id
                Remove-MSALTokenEntry -TokenId $TokenInfo.Id
                Invoke-AuthTokenFailed -Provider "MSAL" -TenantId $TokenInfo.Token.TenantId -ErrorCode "TokenExpired" -Message "BYO token expired for tenant '$($TokenInfo.Token.TenantId)'"
                return $null
            }
            return Get-TokenInfo $TokenInfo.Id
        }

        if($TokenInfo.AuthType -eq "Confidential") {
            [string[]]$ccScopes = "https://$($TokenInfo.CloudEntry.GraphHost)/.default"
            $ccAuthResult, $ccAuthFailure = Get-MsalAuthenticationToken ($TokenInfo.App.AcquireTokenForClient($ccScopes))
            if($ccAuthResult) {
                $TokenInfo.Token = $ccAuthResult
                $TokenInfo.JWTAccessToken = Get-JWTtoken $ccAuthResult.AccessToken
            }
            elseif($ccAuthFailure) {
                Write-LogError "Failed to refresh confidential client token for app '$($TokenInfo.EntraApp.ClientId)'" $ccAuthFailure
            }
            return Get-TokenInfo $TokenInfo.Id
        }

        $msalApp = $TokenInfo.App
        $cloudEntry = $TokenInfo.CloudEntry
        $User = $TokenInfo.Token.Account.Username
        $account = $TokenInfo.Token.Account
        $TenantID = $TokenInfo.Token.TenantId
    }
    else {
        $msalEnvironment = Get-EntraEnvironment $Environment
        $EntraApp = Get-EntraApp $AppId $RedirectUri $Authority
        $msalApp = New-MSALApp $EntraApp $msalEnvironment
        # Phase 4: resolve the Cloud entry early so scope-building below + TokenInfo
        # persistence below both use a single source of truth.
        $cloudVal = $resolvedCloud
        if(-not $cloudVal) {
            $cloudVal = Convert-LegacyToCloud -GraphEnvironment $Environment -GCCType $null
        }
        $cloudEntry = Get-CloudByValue $cloudVal

        try {
            $script:MSALAccounts = $msalApp.GetAccountsAsync().GetAwaiter().GetResult()
        }
        catch {
            Write-LogError "Failed to get accounts" $_.Exception
        }
        $account = $null
        if($script:MSALAccounts.Count -gt 0 -and [String]::IsNullOrEmpty($User) -eq $false) {
            $account = $script:MSALAccounts | Where-Object userName -eq $User
        }
        
        if(-not $account -and ([String]::IsNullOrEmpty($User) -or $UserID)) {
            if($UserID) { $tmpId = $UserID }
            else { $tmpId = Get-SettingStoreValue "" "LastLoggedOnUserId" }
            if($tmpId) {
                # Try to get user based on Id - to allow alias login...
                $account = $script:MSALAccounts | Where-Object { $_.HomeAccountId.ObjectId -eq $tmpId }
            }
        }
        
        if(-not $account -and [String]::IsNullOrEmpty($User) -eq $false) {
            if($User) {$tmpUser = $User }
            else { $tmpUser = Get-SettingStoreValue "" "LastLoggedOnUserId" }
            if($tmpUser) {
                $account = $script:MSALAccounts | Where-Object { $_.HomeAccountId.Identifier -eq $tmpUser }
            }
        }        
    }

    [string[]] $scopes = "https://$($cloudEntry.GraphHost)/.default"

    # Generate a correlation id once per Connect-EntraEnvironment call so silent + any
    # interactive retry share the same id. Helps Microsoft support trace the flow end to
    # end; off by default to keep logs quiet.
    $logCorrelation = (Get-SettingValue "MSALCorrelationLogging") -eq $true
    $correlationId  = if($logCorrelation) { [Guid]::NewGuid() } else { $null }

    if($account -and $ForceInteractive -ne $true) {
        $AquireTokenObj = $msalApp.AcquireTokenSilent($scopes, $account)
        if($ForceRefresh) { [void]$AquireTokenObj.WithForceRefresh($ForceRefresh) }
        if($TenantID)     { [void]$AquireTokenObj.WithTenantId($TenantID) }
        # CAE: when the caller passes a Graph 401 claims challenge, propagate it into
        # the silent acquire so MSAL re-fronts the request and satisfies the challenge.
        if($ClaimsChallenge) {
            Write-LogDebug "Applying CAE claims challenge to silent acquire ($($ClaimsChallenge.Length) chars)"
            [void]$AquireTokenObj.WithClaims($ClaimsChallenge)
        }
        if($correlationId) {
            [void]$AquireTokenObj.WithCorrelationId($correlationId)
            Write-LogDebug "MSAL silent acquire correlation: $correlationId"
        }
        Write-LogDebug "MSAL silent acquire for '$($account.Username)' scopes=$($scopes -join ',') tenant=$TenantID"

        $authResult, $authenticationFailure = Get-MsalAuthenticationToken $AquireTokenObj
    }

    if (-not $authResult -and $ForceSilent -ne $true) {
        # Decide browser vs device code. On the system-browser path (Linux/macOS,
        # or Windows with UseSystemBrowser) a machine with no usable browser would
        # otherwise hang the interactive poll loop until timeout, so pre-check and
        # route straight to device code. This block is already gated on
        # -ForceSilent, so background/automatic refresh never lands here - only an
        # explicit interactive sign-in does.
        $onSystemBrowserPath = (-not $script:IsWindowsOS -or $script:MSALUseSystemBrowser)
        $useDeviceCode = [bool]$DeviceCode
        if(-not $useDeviceCode -and $onSystemBrowserPath -and -not (Test-InteractiveBrowserAvailable)) {
            Write-Log "No usable web browser detected - using device code sign-in instead" 2
            $useDeviceCode = $true
        }

        if(-not $useDeviceCode) {
            Write-Log "Initiate interactive login"
            $AquireTokenObj = $msalApp.AcquireTokenInteractive($scopes)

            if($User) {
                [void]$AquireTokenObj.WithLoginHint($User)
                [void]$AquireTokenObj.WithPrompt([Microsoft.Identity.Client.Prompt]::NoPrompt)
            }
            else
            {
                [void]$AquireTokenObj.WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount)
            }

            # Claims precedence: an explicit Graph-side challenge ($ClaimsChallenge) wins over
            # the MSAL-internal $authenticationFailure.Claims (which is set when the silent
            # path bounced with MsalUiRequiredException). Both ultimately resolve to the
            # same WithClaims call.
            $effectiveClaims = $ClaimsChallenge
            if(-not $effectiveClaims -and $authenticationFailure.Claims) {
                $effectiveClaims = $authenticationFailure.Claims
            }
            if($effectiveClaims)
            {
                Write-Log "Login claims: $effectiveClaims"
                [void]$AquireTokenObj.WithClaims($effectiveClaims)
            }

            if ($TenantID) { [void]$AquireTokenObj.WithTenantId($TenantID) }

            [IntPtr]$ParentWindow = Get-WindowHandle $PID
            if ($ParentWindow -ne [IntPtr]::Zero) { [void]$AquireTokenObj.WithParentActivityOrWindow($ParentWindow) }

            # Own the browser launch on the system-browser path so a failed launch
            # (no default browser / broken handler) faults the task in ~1s with the
            # IM_NO_BROWSER marker instead of hanging until the interactive timeout.
            # SystemWebViewOptions.OpenBrowserAsync is a Func<Uri,Task>; it must be a
            # compiled delegate because MSAL invokes it on a Runspace-less thread
            # where a ScriptBlock delegate would throw.
            if($onSystemBrowserPath) {
                try {
                    Initialize-IMBrowserLauncher
                    $swvOptions = [Microsoft.Identity.Client.SystemWebViewOptions]::new()
                    $openFuncType = [System.Func[System.Uri, System.Threading.Tasks.Task]]
                    $openMi = [IMMsalBrowserLauncher].GetMethod('Open')
                    $swvOptions.OpenBrowserAsync = [System.Delegate]::CreateDelegate($openFuncType, $openMi)
                    [void]$AquireTokenObj.WithSystemWebViewOptions($swvOptions)
                }
                catch {
                    Write-LogDebug "Could not attach custom browser launcher: $($_.Exception.Message)"
                }
            }

            if($correlationId) {
                [void]$AquireTokenObj.WithCorrelationId($correlationId)
                Write-LogDebug "MSAL interactive acquire correlation: $correlationId"
            }
            Write-LogDebug "MSAL interactive acquire scopes=$($scopes -join ',') tenant=$TenantID user='$User'"

            $authResult, $authenticationFailure = Get-MsalAuthenticationToken $AquireTokenObj -Interactive

            # Browser launch failed (no usable browser). Fall back to device code so
            # the user can still sign in from a phone / another device.
            if(-not $authResult -and $onSystemBrowserPath -and (Test-IMNoBrowserFailure $authenticationFailure)) {
                Write-Log "Browser could not be launched - falling back to device code sign-in" 2
                $useDeviceCode = $true
            }

            # A custom app registration without the loopback redirect. The browser
            # flow can only redirect to http://localhost, and Entra refuses with
            # AADSTS50011 - the one thing the system-browser default asks of a custom
            # registration. Say what to do instead of leaving the raw error.
            if(-not $authResult -and $onSystemBrowserPath) {
                $failureText = if($authenticationFailure -is [Exception]) { $authenticationFailure.Message } else { "$authenticationFailure" }
                if($failureText -match 'AADSTS50011') {
                    Write-Log "Sign-in failed because the app registration does not allow the redirect URI http://localhost. Add it under 'Mobile and desktop applications' on the app registration in Entra, or turn off 'Use system browser for login' in Settings to sign in through the embedded window." 3
                }
            }
        }

        if($useDeviceCode) {
            # Device code (RFC 8628): MSAL calls the callback with the user_code +
            # verification_uri + human-readable message. In the GUI we surface that
            # via the status overlay and copy the code to the clipboard (the
            # -DeviceCode path in Get-MsalAuthenticationToken below); a console host
            # also prints it. The callback returns a completed Task so MSAL polls.
            #
            # CRITICAL: MSAL invokes the callback on an internal HTTP polling
            # thread that has NO PowerShell Runspace. A ScriptBlock cast to
            # Func<DeviceCodeResult, Task> fails at invocation with "There is
            # no Runspace available to run scripts in this thread." Build the
            # callback as a compiled C# static method and bind it via
            # [Delegate]::CreateDelegate so no PowerShell scriptblock runs on
            # that thread - plain .NET calls only.
            Write-Log "Initiate device code login"

            if(-not ('IMMsalDeviceCodeHelper' -as [type])) {
                # Resolve reference assemblies from types already loaded so the
                # source compiles under both .NET Framework (PS 5.1, mscorlib
                # only) and .NET Core / .NET 8+ (PS 7+, split BCL: System.Runtime,
                # System.Console, System.Threading.Tasks). Add-Type in PS7
                # REPLACES the default reference set when -ReferencedAssemblies
                # is passed, so we must enumerate every ref needed by the source.
                $refAssemblies = @(
                    [System.Console].Assembly.Location,
                    [System.Threading.Tasks.Task].Assembly.Location,
                    [Microsoft.Identity.Client.DeviceCodeResult].Assembly.Location
                ) | Where-Object { $_ } | Select-Object -Unique
                Add-Type -TypeDefinition @"
using System;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
public static class IMMsalDeviceCodeHelper {
    // Latest device-code result, stashed for the PowerShell side to surface in
    // the GUI. The callback runs on a Runspace-less thread and can only touch
    // static fields (and the console), so it cannot show UI itself.
    public static string Message;
    public static string VerificationUrl;
    public static string UserCode;
    public static DateTimeOffset ExpiresOn;
    public static bool HasResult;
    public static Task Callback(DeviceCodeResult r) {
        try {
            Message = r.Message;
            VerificationUrl = r.VerificationUrl;
            UserCode = r.UserCode;
            ExpiresOn = r.ExpiresOn;
            HasResult = true;
        } catch { }
        try {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(r.Message);
            Console.ResetColor();
            Console.WriteLine();
        } catch { /* host may lack a console (WPF UI thread etc.) - swallow */ }
        return Task.CompletedTask;
    }
}
"@ -ReferencedAssemblies $refAssemblies -CompilerOptions '/nowarn:1701;1702' -ErrorAction Stop
            }

            # Backtick-arity syntax [System.Func`2[T,R]] fails to resolve as an
            # inline type literal in PS7's parser (see the FAIL when we tried).
            # Use the bracket-generic form [System.Func[T,R]] - PowerShell picks
            # the right arity from the type argument count.
            $callbackType = [System.Func[Microsoft.Identity.Client.DeviceCodeResult, System.Threading.Tasks.Task]]
            $callbackMi   = [IMMsalDeviceCodeHelper].GetMethod('Callback')
            $dcCallback   = [System.Delegate]::CreateDelegate($callbackType, $callbackMi)

            $AquireTokenObj = $msalApp.AcquireTokenWithDeviceCode([string[]]$scopes, $dcCallback)

            if($TenantID) { [void]$AquireTokenObj.WithTenantId($TenantID) }

            $effectiveClaims = $ClaimsChallenge
            if(-not $effectiveClaims -and $authenticationFailure.Claims) { $effectiveClaims = $authenticationFailure.Claims }
            if($effectiveClaims) {
                Write-Log "Login claims: $effectiveClaims"
                [void]$AquireTokenObj.WithClaims($effectiveClaims)
            }

            if($correlationId) {
                [void]$AquireTokenObj.WithCorrelationId($correlationId)
                Write-LogDebug "MSAL device-code acquire correlation: $correlationId"
            }
            Write-LogDebug "MSAL device-code acquire scopes=$($scopes -join ',') tenant=$TenantID"

            # Clear the previous attempt's stashed code so the poll loop only
            # surfaces the code for THIS acquire.
            try { [IMMsalDeviceCodeHelper]::HasResult = $false } catch { }

            $authResult, $authenticationFailure = Get-MsalAuthenticationToken $AquireTokenObj -DeviceCode
        }
    }

    if($authResult) {

        # Diagnostic: log the issued token lifetime (exp-iat) and the xms_cc claim so a
        # live-tenant run confirms whether CAE ('cp1') is still negotiated. If cp1 shows
        # here after EnableCAE was turned off, the MSAL app needs a rebuild (restart /
        # fresh login) - New-MSALApp bakes the capability in at build time.
        try {
            $jwtDbg = Get-JWTtoken $authResult.AccessToken
            if($jwtDbg -and $jwtDbg.Payload) {
                $lifeMin = try { [math]::Round((([int64]$jwtDbg.Payload.exp) - ([int64]$jwtDbg.Payload.iat)) / 60) } catch { '?' }
                $ccDbg   = @($jwtDbg.Payload.xms_cc) -join ','
                Write-LogDebug "MSAL token acquired: lifetime=$lifeMin min, xms_cc=[$ccDbg], tenant=$($authResult.TenantId)"
            }
        } catch { }

        $hashId = ($authResult.Account.HomeAccountId.Identifier + $authResult.TenantId + $authResult.ClientId)

        $tmpId = (Add-MSALTokenInfo $hashId $authResult -DefaultToken:($DefaultToken -eq $true) -Cloud $resolvedCloud)

        # Phase 3: persist per-tenant cloud memory. Prefer the Cloud the caller asked
        # for (most reliable — distinguishes USGov from USGovDOD which the JWT iss
        # can't). Fall back to detecting from the JWT iss claim for older call sites
        # that only know -Environment. Then save against the tenant we actually
        # authenticated against (authResult.TenantId — for guest scenarios this is
        # the resource tenant, which is the correct memory key for resume).
        try {
            $persistCloud = $resolvedCloud
            if(-not $persistCloud) {
                $iss = $null
                $jwt = Get-JWTtoken $authResult.AccessToken
                if($jwt -and $jwt.Payload) { $iss = [string]$jwt.Payload.iss }
                $persistCloud = Resolve-CloudFromIss -Iss $iss -FallbackCloud (Get-DefaultCloud)
            }
            if($persistCloud -and $authResult.TenantId) {
                Save-TenantCloud -TenantId $authResult.TenantId -Cloud $persistCloud
                if($DefaultToken -eq $true) {
                    Save-SettingStoreValue "" "LastLoggedOnCloud" $persistCloud
                }
            }
        }
        catch {
            Write-LogDebug "Phase 3 cloud memory write failed: $($_.Exception.Message)"
        }

        return (Get-TokenInfo $tmpId)
    }
    elseif ($authenticationFailure -is [System.OperationCanceledException]) {
        # User cancellation is an expected, non-failure outcome. In particular, do
        # not notify AuthenticationFailed subscribers: an existing session may still
        # be valid (for example when a consent prompt is cancelled).
        Write-Log "MSAL authentication cancelled by the user" 2
    }
    elseif ($SuppressFailedEvent) {
        # Background near-expiry refresh that missed while a valid token still exists.
        # Do not raise AuthenticationFailed - the caller keeps using the current token.
        Write-Log "Silent token refresh returned no token; keeping the existing valid token (background refresh, no failure raised)" 2
    }
    else {
        Invoke-AuthTokenFailed -Provider "MSAL" -Message "MSAL authentication did not return a token"
    }
}

function Add-MSALTokenInfo
{
    param(
        $TokenHashID,
        $AuthenticationResult,
        [switch]$DefaultToken,
        # Phase 4: caller passes the Cloud value resolved from -Cloud / -Environment /
        # Get-StartupCloudHint so the TokenInfo carries the right CloudEntry
        # (GraphHost / AADAuthority / MgEnvironment / legacy mapping) for resume.
        [string]$Cloud
    )

    if(-not $TokenInfo -and $script:MSALTokens.ContainsKey($TokenHashID)) {
        $TokenInfo = $script:MSALTokens[$TokenHashID]
    }

    $newToken = $false
    $newUser = $false

    if(-not $TokenInfo) {
        $cloudVal = if($Cloud) { $Cloud } else { Get-DefaultCloud }
        $cloudEntry = Get-CloudByValue $cloudVal
        $TokenInfo = [PSCustomObject]@{
            ID = (Get-NextAuthTokenId)  # central allocator -> globally unique across providers
            App = $msalApp
            Cloud = $cloudVal
            CloudEntry = $cloudEntry
            EntraApp = $EntraApp
            Token = $null
            JWTAccessToken = $null
            JWTIdToken = $null
            IsDefault = ($script:MSALTokens.Count -eq 0 -or $DefaultToken -eq $true) # First authentication is always default
            Tenants = $null
            Organization = $null
        }

        $script:MSALTokens.Add($TokenHashID, $TokenInfo) | Out-Null
        $newToken = $true
    }

    $TokenInfo.Token = $AuthenticationResult
    $TokenInfo.JWTAccessToken = Get-JWTtoken $AuthenticationResult.AccessToken
    $TokenInfo.JWTIdToken = Get-JWTtoken $AuthenticationResult.IdToken

    if($DefaultToken -eq $true) {
        $curDefaultToken = $script:MSALTokens.values | Where-Object { $_.IsDefault -eq $true -and $_.Id -ne $TokenInfo.Id }
        if($curDefaultToken.Id -ne $TokenInfo.Id)
        {
            if($curDefaultToken) { $curDefaultToken.IsDefault = $false }
            $TokenInfo.IsDefault = $true
            $newUser = $true
        }
    }

    if($newToken -eq $true -or $newUser -eq $true) {
        if(-not $TokenInfo.Organization) {
            $TokenInfo.Organization = (Get-MSALOrganizationInfo $TokenInfo)
        }

        # CRITICAL: $script:OrganizationId/Name are the *default* tenant snapshot for code
        # that doesn't carry a TokenId. Updating them on every new token (including
        # non-default secondary tenants) silently broke every cross-tenant flow that read
        # the globals. Only refresh when this token IS or BECOMES the default.
        if($TokenInfo.IsDefault) {
            # Organization can be $null if /Organization 401'd (revoked token, missing
            # Organization.Read.All scope, transient Graph error). Don't crash — fall
            # back to TenantId from the access token, and keep going. Auth completed
            # successfully; only the display-name lookup failed.
            if($TokenInfo.Organization -and $TokenInfo.Organization.Id) {
                $script:OrganizationId = $TokenInfo.Organization.Id
            }
            elseif($AuthenticationResult -and $AuthenticationResult.TenantId) {
                $script:OrganizationId = $AuthenticationResult.TenantId
                Write-Log "Organization info unavailable (Graph /Organization failed). Using TenantId '$($AuthenticationResult.TenantId)' as fallback." 2
            }

            if($TokenInfo.Organization -and $TokenInfo.Organization.displayName) {
                $script:OrganizationName = ([string]$TokenInfo.Organization.displayName).Trim()
            }
            else {
                $script:OrganizationName = $null
            }

            Save-SettingStoreValue "" "LastLoggedOnUserId" $AuthenticationResult.Account.HomeAccountId.ObjectId
        }

        if($newToken -eq $true -and (Get-SettingValue "GetTenantList" -TenantID $AuthenticationResult.Account.HomeAccountId.TenantId) -eq $true) {
            Get-TenantList $TokenInfo
        }

        Write-Log "Successfully authenticated user $($TokenInfo.Token.Account.Username)"

        # Diff the just-issued token's scopes against the union of every registered
        # policy type's declared _Permissions. We use ".default" as the login scope
        # so MSAL can't add scopes the Entra app isn't already consented to — the
        # check exists to surface what's missing in the Entra app config so the user
        # can fix it once instead of seeing scattered 403s on bulk operations.
        # Wrapped in try because a malformed JWT or a transient $script:IntuneTypes
        # state must not break the auth flow itself.
        try { Update-MissingPermissions $TokenInfo }
        catch { Write-LogDebug "Update-MissingPermissions threw: $($_.Exception.Message)" }

        # Realign the active auth provider with the auth that just succeeded. If the
        # user had MgGraph as active but did an MSAL flow (clicked a cached MSAL
        # account, used "Sign in with another account", etc.), subsequent
        # Invoke-MSGraphAPI calls would route to MgGraph and 401. Switch so the rest
        # of the session uses the provider that actually has a valid session.
        if(Get-Command Get-AuthProvider -ErrorAction SilentlyContinue) {
            $cur = (Get-AuthProvider)
            if($cur -and $cur.Id -ne "MSAL") {
                Write-Log "Active auth provider auto-switched from '$($cur.Id)' to 'MSAL' because MSAL authentication succeeded"
                Set-ActiveAuthProvider -Id "MSAL"
            }
        }

        # Register with the central token registry, which fires AuthenticatedNewToken
        # with the canonical [IMAuthToken] (sole event firer). Idempotent for an
        # already-registered token being promoted to default ($newUser) - it re-stamps
        # the record and sets the registry default. MSAL keeps its own per-entry
        # IsDefault flag for Update-MSALUserProfile / Get-FullToken; -Default keeps the
        # registry default in lockstep.
        Register-AuthToken -Provider (Get-AuthProvider "MSAL") -TokenId $TokenInfo.ID -Cloud $TokenInfo.Cloud -Default:([bool]$TokenInfo.IsDefault) | Out-Null
    }
    $refreshToken = $false
    if($authResult.AuthenticationResultMetadata.TokenSource -eq "Broker" -and $telemetry.broker_app_used -eq "true") {
        if($newToken -eq $false)
        {
            $refreshToken = $true
        }
        Write-Log "$($authResult.Account.UserName) authenticated with Broker. Time (ms): $($authResult.AuthenticationResultMetadata.DurationTotalInMs). CorrelationId: $($authResult.CorrelationId)"
    }
    elseif($authResult.AuthenticationResultMetadata.TokenSource -eq "IdentityProvider") {
        if($newToken -eq $false)
        {
            $refreshToken = $true
        }
        Write-Log "Token received from IdentityProvider. Time (ms) $($authResult.AuthenticationResultMetadata.DurationTotalInMs)"
    }
    else {
        # ToDo: Change to debug
        Write-Log "Token received from Cache. Time (ms) $($authResult.AuthenticationResultMetadata.DurationTotalInMs)"
    }
    if($refreshToken -eq $true) {
        Update-AuthToken -TokenId $TokenInfo.ID
    }

    $TokenInfo.ID
}

function Get-MSALOrganizationInfo
{
    param($TokenInfo)

    Write-Log "Get organization info"
    $organization = (Invoke-MSGraphAPI -Url "Organization" -SkipAuthentication -ODataMetadata "Skip" -TokenId $TokenInfo.Id).Value
    if($organization)
    {
        if($organization -is [array]) { $organization = $organization[0]}
        Save-SettingStoreValue $organization.Id "_Name" $organization.displayName
    }
    
    $organization
}

function Get-DefaultTokenId
{
    # Provider-agnostic: the default token now lives in the central registry (which
    # every provider feeds), not just the MSAL store. Returns 0 when nothing is
    # signed in - callers treat 0 as "no default" exactly as before.
    Get-DefaultAuthTokenId
}

function Get-FullToken
{
    # MSAL-only token-registry lookup. Callers commonly use this as a "probe" — if
    # the returned value is $null they fall back to another auth path (provider
    # facade, MgGraph SDK, etc.). Logging a WARNING for that normal case spammed
    # the log with "No token found with Id 0" every time the avatar refreshed,
    # the env badge ticked, or an assignment view validated. Demoted to debug so
    # real diagnostics aren't drowned out, and gated so callers explicitly asking
    # about a specific id still see something useful in debug.
    param([int]$Id = (Get-DefaultTokenId))

    # Token IDs are positive (allocated by Get-NextAuthTokenId). Provider-facade callers
    # use 0 as the "use default" sentinel ($provider.GetUserInfo(0), etc.) but a
    # PowerShell default param value only fires when the argument is OMITTED — an
    # explicit 0 leaves $Id at 0 and the lookup matches no token, so the Tenant
    # Settings menu was wrongly disabled, the avatar tooltip was empty, and so on.
    # Treat any non-positive id as "default".
    if($Id -le 0) { $Id = (Get-DefaultTokenId) }

    $token = $script:MSALTokens.Values | Where-Object Id -eq $id

    if(-not $token) {
        Write-LogDebug "Get-FullToken: no MSAL token registered for Id $Id (probe; caller handles null)"
    }

    $token
}

function Get-TokenInfo
{
    # Provider-agnostic shim over the central token registry. Projects each
    # [IMAuthToken] into the legacy Get-TokenInfo shape (Id / IsDefault / User /
    # TenantID / TenantName / ClientID / Expires) that existing callers - the
    # dependency cache, cross-tenant Copy dialogs, Get-GraphPolicies - already
    # bind to. Because it now sources the registry, those callers work for OAuth /
    # MgGraph tokens too, not just MSAL.
    param([Int]$TokenID = 0)

    $tokenObjects = Get-AuthTokenList | Select-Object `
        @{n="Id";e={$_.TokenId}},
        IsDefault,
        @{n="User";e={ if($_.UPN) { $_.UPN } else { $_.Account } }},
        @{n="TenantID";e={$_.TenantId}},
        TenantName,
        @{n="ClientID";e={$_.AppId}},
        @{n="Expires";e={$_.ExpiresOn}} |
        Sort-Object -Property Id

    if($null -ne $TokenID -and $TokenID -gt 0)
    {
        $tokenObjects = $tokenObjects | Where-Object Id -eq $TokenID
    }

    $tokenObjects
}

# The ONE token an operation runs against, as a single object.
#
# Get-TokenInfo above is a LIST accessor: 0 (its default) means "no filter, every
# registered token". Callers that hand it an operation's TokenId therefore get an
# array the moment a second tenant is signed in, because 0 is also the convention
# for "the default token" - and $arr.TenantId on that array is an ARRAY of tenant
# ids. Import fed exactly that into Convert-GraphOrganizationTokenToValue, so
# %OrganizationId% was restored as "id-a id-b" instead of the target tenant's id.
#
# 0 resolves to the default token here (IsDefault, then the lowest id as a
# fallback), which is what the callers meant by it all along. Anything else keeps
# Get-TokenInfo's behaviour, narrowed to one object.
function Get-OperationTokenInfo
{
    param([int]$TokenID = 0)

    if($TokenID -gt 0) { return (@(Get-TokenInfo $TokenID) | Select-Object -First 1) }

    $all = @(Get-TokenInfo)
    if($all.Count -eq 0) { return $null }

    $default = $all | Where-Object IsDefault | Select-Object -First 1
    if($default) { return $default }

    # No token is flagged default (Set-DefaultAuthToken never ran). Lowest id is
    # deterministic; Get-TokenInfo already sorts by Id.
    return $all[0]
}

function Get-TokenInfoForTenant
{
    param($TenantId)

    $tenantTokenObj = Get-AuthTokenList | Where-Object { $_.TenantId -eq $TenantId } | Select-Object -First 1

    if($tenantTokenObj) {
        Get-TokenInfo $tenantTokenObj.TokenId
    }
}

function Get-WindowHandle
{
    param($ProcessId)

    # A parent window handle is only meaningful for the Windows interactive flows
    # (WAM / embedded WebView). On Linux/macOS auth goes through the system browser,
    # there's no HWND, and the CIM cmdlets (Win32_Process) don't even exist — calling
    # Get-CimInstance there throws "not recognized". Return a null handle instead.
    # Callers must treat [IntPtr]::Zero as "no handle" (it's truthy in PowerShell, so
    # a bare `if ($handle)` check is not enough).
    if(-not $script:IsWindowsOS) { return [IntPtr]::Zero }

    if($null -eq $ProcessId) { return [IntPtr]::Zero }

    $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if(-not $process) { return [IntPtr]::Zero }

    if($process.MainWindowHandle -and $process.MainWindowHandle -ne 0) {
        return $process.MainWindowHandle
    }

    $cimProcess = Get-CimInstance -Class Win32_Process -Filter "ProcessID = '$ProcessId'" -ErrorAction SilentlyContinue
    if(-not $cimProcess) { return [IntPtr]::Zero }
    Get-WindowHandle $cimProcess.ParentProcessId
}

function Get-MSALAppAuthority
{
    # Returns the host portion of the authority URL for the current default environment.
    # Used by Start-MSALConsentPrompt when constructing a tenant-scoped authority.
    param($TokenId = (Get-DefaultTokenId))

    $token = $script:MSALTokens.Values | Where-Object Id -eq $TokenId
    if($token -and $token.CloudEntry -and $token.CloudEntry.AADAuthority) {
        return $token.CloudEntry.AADAuthority
    }

    $env = Get-EntraEnvironment "public"
    return $env.URL
}

function Update-MissingPermissions
{
    <#
    .SYNOPSIS
        Diff the API permissions declared by registered policy types against the
        scopes/roles actually present in the access token.

    .DESCRIPTION
        Login uses the MSAL ".default" scope, which means "give me a token containing
        every permission the Entra App registration is consented to." That covers the
        Entra-app side, but it tells us nothing about whether each individual policy
        type's required scope is actually present. This function does the post-auth
        check:
          * Collects the union of _Permissions across $script:IntuneTypes (each
            registered concrete policy type contributes whatever it declared in its
            class file — the dead Add-ViewItem code at CoreUI.ps1:102-129 used to do
            this and still doesn't).
          * Reads the granted permissions from the just-issued JWT:
              - delegated tokens (idtyp = 'user'): 'scp' claim, space-separated string
              - application tokens (idtyp = 'app'): 'roles' claim, array
            Both are checked so this works for either token type.
          * Stores the missing set in $script:missingPermissions so the existing
            Start-MSALConsentPrompt path can fire (or so the UI / a future feature
            can read it).
          * Logs a single warning enumerating the missing scopes — that is the cue to
            add them to the Entra App registration and grant admin consent. Without
            consent at the Entra-app level there is nothing the client can do; the
            ".default" silent acquire just returns the same scopes the cached token
            already has.

        Side note on permission strings: '_Permissions' uses Microsoft Graph short
        names ('DeviceManagementScripts.ReadWrite.All') and so does 'scp'/'roles', so a
        case-insensitive direct match is enough. We do NOT model implication chains
        (e.g. .Read.All being weaker than .ReadWrite.All) — Entra evaluates that on
        the server side and would reject the call if a weaker scope wasn't sufficient.
    #>
    param($TokenInfo)

    $script:missingPermissions = @()

    if(-not $TokenInfo -or -not $TokenInfo.JWTAccessToken -or -not $TokenInfo.JWTAccessToken.Payload) {
        Write-LogDebug "Update-MissingPermissions: no JWT payload available; skipping scope diff"
        return
    }

    # Collect distinct required permissions across every registered policy type.
    $required = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach($t in @($script:IntuneTypes)) {
        if(-not $t -or -not $t._Permissions) { continue }
        foreach($p in @($t._Permissions)) {
            if($p) { [void]$required.Add($p) }
        }
    }
    if($required.Count -eq 0) {
        Write-LogDebug "Update-MissingPermissions: no _Permissions declared on any registered type"
        return
    }

    # Read granted permissions from the token payload. Newer tokens carry an explicit
    # 'idtyp' discriminator (user/app) but we don't strictly need it — checking both
    # claims is harmless and covers older tokens that omit idtyp.
    $payload = $TokenInfo.JWTAccessToken.Payload
    $granted = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if($payload.scp) {
        foreach($s in ([string]$payload.scp).Split(' ')) {
            if($s) { [void]$granted.Add($s) }
        }
    }
    if($payload.roles) {
        foreach($r in @($payload.roles)) {
            if($r) { [void]$granted.Add($r) }
        }
    }

    if($granted.Count -eq 0) {
        Write-Log "Update-MissingPermissions: token has no scp/roles claims; cannot validate required permissions" 2
        return
    }

    $missing = @($required | Where-Object { -not $granted.Contains($_) }) | Sort-Object
    if($missing.Count -eq 0) {
        Write-LogDebug "All $($required.Count) declared permission(s) are present in the token"
        return
    }

    $script:missingPermissions = $missing
    Write-Log ("Token is missing $($missing.Count) of $($required.Count) declared API permission(s): $($missing -join ', '). " +
               "Add these to the Entra App registration (Microsoft Graph permissions) and grant admin consent, then sign out and back in. " +
               "Affected policy types will return 403 / be skipped on bulk operations until the scopes are present.") 2
}

function Start-MSALConsentPrompt
{
    <#
    .SYNOPSIS
        Trigger an interactive consent prompt for the scopes recorded in
        $script:missingPermissions, using the default-token's app + account.

    .DESCRIPTION
        Mirrors the OLD Extensions/MSALAuthentication.psm1 flow but reads from the
        new per-token registry ($script:MSALTokens via Get-FullToken) instead of the
        retired $global:MSALToken / $global:MSALApp globals. Called from the
        "Request Consent" link in ProfileInfo.Xaml when Update-MissingPermissions
        flagged scopes that the Entra app hasn't consented yet.

        Flow:
          1. Pull the active token info via Get-FullToken (default token by default;
             caller may pass an explicit -TokenId for cross-tenant scenarios).
          2. AcquireTokenInteractive against Graph .default + WithExtraScopesToConsent
             listing the missing scopes. WithExtraScopesToConsent surfaces the new
             scopes in the consent UI but doesn't put them on the returned token's
             scp claim — only the next .default acquire after consent will.
          3. If consent succeeds, replace TokenInfo.Token with the new auth result so
             subsequent Get-FullToken callers see the fresh token; re-run
             Update-MissingPermissions and refresh the consent-link visibility on
             the currently-open ProfileInfo popup.
    #>
    param([switch]$PassThru, [int]$TokenId = (Get-DefaultTokenId))
    Write-Log "Initiate consent prompt"

    if(($script:missingPermissions | Measure-Object).Count -eq 0) {
        Write-Log "No missing permissions; consent prompt has nothing to request" 2
        return
    }

    $tokenInfo = Get-FullToken $TokenId
    if(-not $tokenInfo -or -not $tokenInfo.Token -or -not $tokenInfo.App) {
        Write-Log "Cannot start consent prompt - no MSAL token registered for TokenId $TokenId" 3
        return
    }

    $loginHintName = $tokenInfo.Token.Account.UserName
    $tenantId      = $tokenInfo.Token.TenantId
    if(-not $loginHintName -or -not $tenantId) {
        Write-Log "Cannot start consent prompt - login hint or tenant id missing on token" 3
        return
    }

    # Pass the missing permissions as the PRIMARY scope to AcquireTokenInteractive.
    # When MSAL sees an unconsented scope as the primary, it forces Entra to surface
    # the actual consent screen (per-scope checkboxes / "Accept" button) — that's
    # what records the consent grant on the Entra app. Earlier we tried .default +
    # WithExtraScopesToConsent: .default returns whatever is already consented and
    # MSAL skipped the consent surface, so the prompt the user saw was just the
    # account picker — no consent grant was recorded.
    # Force Prompt.Consent so Entra always shows the consent UI (even for already-
    # consented scopes), which is the only reliable way to get the missing-scope
    # checkboxes to appear when the user is already signed in to that account.
    [string[]] $Scopes = [string[]]$script:missingPermissions
    $AquireTokenObj = $tokenInfo.App.AcquireTokenInteractive($Scopes)
    # WithAuthority on AcquireToken* is [Obsolete] in MSAL.NET (see
    # AbstractAcquireTokenParameterBuilder.cs L99-224). The authority host is
    # already on the app builder; per-request we only need to override the tenant.
    [void]$AquireTokenObj.WithTenantId($tenantId)
    [void]$AquireTokenObj.WithLoginHint($loginHintName)
    [void]$AquireTokenObj.WithPrompt([Microsoft.Identity.Client.Prompt]::Consent)

    [IntPtr]$ParentWindow = Get-WindowHandle $PID
    if($ParentWindow -ne [IntPtr]::Zero) {
        [void]$AquireTokenObj.WithParentActivityOrWindow($ParentWindow)
    }

    # Heads-up: if the user is signed in via a Microsoft first-party app (e.g.
    # 'Microsoft Graph PowerShell' = 14d82eec-…) the consent grant won't stick —
    # Microsoft owns those apps and you can't add scopes to their permission set.
    # The fix in that case is to switch to a custom Entra App registration in
    # Settings → Entra → Application Id, then re-login. We log a hint so the user
    # doesn't keep clicking Request Consent expecting it to take effect.
    $clientId = $tokenInfo.EntraApp.ClientId
    $msftFirstParty = @(
        '14d82eec-204b-4c2f-b7e8-296a70dab67e'  # Microsoft Graph PowerShell
        'd1ddf0e4-d672-4dae-b554-9d5bdfd93547'  # Microsoft Intune PowerShell (deprecated)
        '04b07795-8ddb-461a-bbee-02f9e1bf7b46'  # Azure CLI
    )
    if($clientId -and ($msftFirstParty -contains $clientId.ToLower())) {
        Write-Log ("Consent prompt is being run against a Microsoft first-party app ($clientId). New scopes typically cannot be added to those apps - " +
                   "if consent doesn't stick, register a custom Entra app, add the required Graph permissions to it, and switch to it via Settings -> Entra -> Application Id.") 2
    }

    Write-Log "Consent prompt for the following scopes: $($script:missingPermissions -join ', ')"

    $newAuthResult = $null
    $authFailure = $null
    try {
        $newAuthResult, $authFailure = Get-MsalAuthenticationToken $AquireTokenObj -Interactive
    }
    finally {
        # Get-MsalAuthenticationToken removes the Cancel button but intentionally
        # leaves its caller's status text intact. Consent has no outer sign-in wrapper,
        # so it owns clearing the full-window overlay on success, failure and cancel.
        Write-Status ""
    }
    if($newAuthResult) {
        Write-Log "Consent for additional scopes added successfully"
        # Refresh the per-token registry so subsequent Invoke-MSGraphAPI calls go
        # through the new token (with the freshly consented scopes).
        $tokenInfo.Token          = $newAuthResult
        $tokenInfo.JWTAccessToken = Get-JWTtoken $newAuthResult.AccessToken
        $tokenInfo.JWTIdToken     = Get-JWTtoken $newAuthResult.IdToken

        # Re-diff against the new token. If consent succeeded, $script:missingPermissions
        # should now be empty.
        try { Update-MissingPermissions $tokenInfo }
        catch { Write-LogDebug "Update-MissingPermissions threw after consent: $($_.Exception.Message)" }

        # Hide the consent link on any currently-open ProfileInfo popup. The grid
        # reference is stashed by MSGraphAuthenticationUI when the popup opens.
        if($script:userEllipsGrid -and $script:UIProvider) {
            try { $script:UIProvider.SetControlVisible($script:userEllipsGrid, "lnkRequestConsent", $false) }
            catch { Write-LogError "Failed to hide consent request link" $_.Exception }
        }

        # Re-evaluate which views the user can see — newly-consented scopes can
        # surface views that were previously greyed out.
        if(Get-Command Show-ViewMenu -ErrorAction SilentlyContinue) { Show-ViewMenu }

        if($PassThru -eq $true) {
            return $newAuthResult
        }
    }
    elseif($authFailure -is [System.OperationCanceledException]) {
        Write-Log "Consent prompt cancelled by the user" 2
    }
    elseif($authFailure) {
        Write-LogError "Consent prompt failed" $authFailure
    }
}

function Disconnect-EntraEnvironment
{
    param([Int]$TokenID = (Get-DefaultTokenId))

    if(-not $TokenID) { return }

    # Unregister FIRST, while the MSAL entry still exists, so the registry can
    # hydrate the disconnect snapshot (incl. TenantId) via MSAL.GetUserInfo and
    # promote the highest-id survivor as the new registry default. The registry
    # fires AuthenticationUserDisconnected (+ AuthenticatedNewToken for the
    # promoted survivor). Default authority is now the registry, so there is no
    # per-entry IsDefault flag to flip here.
    Unregister-AuthToken -TokenId $TokenID

    # Evict the MSAL account from the persistent cache (for interactive tokens)
    # and drop the in-memory entry so $script:MSALTokens reflects reality.
    Remove-MSALAccount -TokenId $TokenID
    Remove-MSALTokenEntry -TokenId $TokenID

    # If nothing is left, clear the default-tenant snapshot globals.
    if((Get-DefaultAuthTokenId) -le 0) {
        $script:OrganizationId   = $null
        $script:OrganizationName = $null
    }
}

function Remove-MSALTokenEntry
{
    # Removes a TokenInfo from $script:MSALTokens. Does NOT touch the on-disk MSAL
    # account cache — call Remove-MSALAccount for that.
    param([Int]$TokenId)

    if(-not $TokenId) { return }
    $entryKey = ($script:MSALTokens.GetEnumerator() | Where-Object { $_.Value.Id -eq $TokenId } | Select-Object -First 1).Key
    if($entryKey) {
        $script:MSALTokens.Remove($entryKey) | Out-Null
        Write-LogDebug "Removed MSAL token entry $entryKey from in-memory registry"
    }
}

function Remove-MSALAccount
{
    # Evicts the underlying MSAL IAccount from the persistent cache, so the user truly
    # disappears from the cached-users list. Only meaningful for interactive tokens —
    # confidential and BYO tokens have no MSAL account.
    param([Int]$TokenId, $Account)

    $msalApp = $null
    if($TokenId) {
        $tokenFull = $script:MSALTokens.Values | Where-Object Id -eq $TokenId
        if($tokenFull -and $tokenFull.AuthType -in @("Confidential","BYO")) { return }
        if($tokenFull) {
            $msalApp = $tokenFull.App
            if(-not $Account) { $Account = $tokenFull.Token.Account }
        }
    }
    if(-not $msalApp) {
        # Caller passed only an Account: pick any public-client app — they all share the cache.
        $msalApp = $script:MSALApps | Select-Object -First 1
    }
    if(-not $msalApp -or -not $Account) { return }

    try {
        $msalApp.RemoveAsync($Account).GetAwaiter().GetResult()
        Write-LogDebug "Removed MSAL account $($Account.Username) from persistent cache"
    }
    catch {
        Write-LogError "Failed to remove account $($Account.Username) from MSAL cache" $_.Exception
    }

    # Refresh the cached accounts list so the UI reflects the change.
    try {
        $script:MSALAccounts = $msalApp.GetAccountsAsync().GetAwaiter().GetResult()
    }
    catch { }
}

Invoke-MSALInitialize

#region Confidential Client / BYO Token

function New-MSALConfidentialApp
{
    param($TenantId, $AppId, $Secret, $Certificate, $Environment)

    # Lazy-load guard: client-credential connects reach MSAL through here without
    # ever touching New-MSALApp.
    if(-not (Initialize-MSALPrereq)) {
        throw "MSAL libraries could not be loaded from Bin/MSAL_PS5|PS7 - see log for details."
    }

    $authority = "https://$($Environment.URL)/$TenantId/"

    $appBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($AppId)
    [void]$appBuilder.WithAuthority($authority)
    [void]$appBuilder.WithClientName("IntuneManagement")
    # WithClientVersion expects a string; explicit ToString() to avoid coercion drift.
    [void]$appBuilder.WithClientVersion($PSVersionTable.PSVersion.ToString())

    if($Secret) {
        [void]$appBuilder.WithClientSecret($Secret)
    }
    elseif($Certificate) {
        [void]$appBuilder.WithCertificate($Certificate)
    }

    $appBuilder.Build()
}

function Connect-WithClientCredentials
{
    param($TenantId, $AppId, $Secret, $Certificate, $GraphEnvironment, $GCCType, [switch]$DefaultToken)

    # Phase 4: derive Cloud + CloudEntry once. CloudEntry.GraphHost gives the right
    # Graph hostname for ALL sovereign clouds — the legacy $msalEnvironment.GraphURL
    # was $null for "usGov" (the legacy taxonomy never carried GraphURL for that key),
    # which silently produced "https:///.default" scopes for GCC tenants.
    $cloudVal   = Convert-LegacyToCloud -GraphEnvironment $GraphEnvironment -GCCType $GCCType
    $cloudEntry = Get-CloudByValue $cloudVal
    $msalEnvironment = Get-EntraEnvironment $cloudEntry.LegacyEnv

    $app = New-MSALConfidentialApp -TenantId $TenantId -AppId $AppId `
                                   -Secret $Secret -Certificate $Certificate `
                                   -Environment $msalEnvironment

    if(-not $app) { return }

    [string[]]$scopes = "https://$($cloudEntry.GraphHost)/.default"
    $authResult, $authenticationFailure = Get-MsalAuthenticationToken ($app.AcquireTokenForClient($scopes))

    if(-not $authResult) {
        Write-Log "Failed to acquire token for app '$AppId' in tenant '$TenantId'" 3
        Invoke-AuthTokenFailed -Provider "MSAL" -TenantId $TenantId -Message "Failed to acquire client-credentials token for app '$AppId' in tenant '$TenantId'"
        return
    }

    $hashId = "CC|$TenantId|$AppId"

    if(-not $script:MSALTokens.ContainsKey($hashId)) {
        $TokenInfo = [PSCustomObject]@{
            ID             = (Get-NextAuthTokenId)
            App            = $app
            Cloud          = $cloudVal
            CloudEntry     = $cloudEntry
            EntraApp       = New-EntraApp -ClientId $AppId -TenantId $TenantId -RedirectUri $null -Authority $null
            Token          = $null
            JWTAccessToken = $null
            JWTIdToken     = $null
            IsDefault      = ($script:MSALTokens.Count -eq 0 -or $DefaultToken -eq $true)
            Tenants        = $null
            Organization   = $null
            AuthType       = "Confidential"
        }
        $script:MSALTokens.Add($hashId, $TokenInfo) | Out-Null
    }
    else {
        $TokenInfo = $script:MSALTokens[$hashId]
        if($DefaultToken) { $TokenInfo.IsDefault = $true }
    }

    $TokenInfo.Token          = $authResult
    $TokenInfo.JWTAccessToken = Get-JWTtoken $authResult.AccessToken

    if(-not $TokenInfo.Organization) {
        $TokenInfo.Organization = Get-MSALOrganizationInfo $TokenInfo
    }
    # Only refresh the global default-tenant snapshot when this token is the default.
    if($TokenInfo.IsDefault) {
        $script:OrganizationId   = $TokenInfo.Organization.Id
        $script:OrganizationName = $TokenInfo.Organization.displayName
    }

    Write-Log "Authenticated with client credentials: app '$AppId', tenant '$TenantId'"
    Register-AuthToken -Provider (Get-AuthProvider "MSAL") -TokenId $TokenInfo.ID -Cloud $TokenInfo.Cloud -Default:([bool]$TokenInfo.IsDefault) | Out-Null

    # Phase 3 + Phase 4: persist per-tenant cloud memory for the Confidential flow too.
    try {
        if($cloudVal -and $authResult.TenantId) {
            Save-TenantCloud -TenantId $authResult.TenantId -Cloud $cloudVal
            if($DefaultToken -eq $true) { Save-SettingStoreValue "" "LastLoggedOnCloud" $cloudVal }
        }
    }
    catch { Write-LogDebug "Confidential cloud memory write failed: $($_.Exception.Message)" }

    Get-TokenInfo $TokenInfo.Id
}

function Add-BYOTokenInfo
{
    param($Token, $TenantId, $GraphEnvironment, $GCCType, [switch]$DefaultToken)

    $jwt = Get-JWTtoken $Token
    if(-not $jwt) {
        Write-Log "Invalid token supplied to Connect-IntuneManagement" 3
        return
    }

    $payload    = $jwt.Payload
    $effectiveTenantId = ?? $TenantId $payload.tid
    $appId      = ?? $payload.appid $payload.azp
    $objectId   = $payload.oid
    $upn        = ?? $payload.upn (?? $payload.unique_name $objectId)
    $expiry     = if($payload.exp) { [DateTimeOffset]::FromUnixTimeSeconds([long]$payload.exp) } `
                  else             { [DateTimeOffset]::UtcNow.AddHours(1) }

    # Phase 4: derive Cloud + CloudEntry from the legacy pair. CloudEntry is the
    # single source of truth on TokenInfo for all sovereign-cloud routing.
    $cloudVal   = Convert-LegacyToCloud -GraphEnvironment $GraphEnvironment -GCCType $GCCType
    $cloudEntry = Get-CloudByValue $cloudVal

    # Synthetic token object shaped to match MSAL AuthenticationResult usage in Invoke-MSGraphAPI
    $syntheticToken = [PSCustomObject]@{
        AccessToken  = $Token
        ExpiresOn    = $expiry
        TenantId     = $effectiveTenantId
        ClientId     = $appId
        Account      = [PSCustomObject]@{
            Username      = $upn
            HomeAccountId = [PSCustomObject]@{ ObjectId = $objectId; Identifier = $objectId }
        }
        AuthenticationResultMetadata = [PSCustomObject]@{
            TokenSource       = "ExternalToken"
            DurationTotalInMs = 0
        }
    }

    $hashId = "BYO|$effectiveTenantId|$appId"

    if(-not $script:MSALTokens.ContainsKey($hashId)) {
        $TokenInfo = [PSCustomObject]@{
            ID             = (Get-NextAuthTokenId)
            App            = $null
            Cloud          = $cloudVal
            CloudEntry     = $cloudEntry
            EntraApp       = New-EntraApp -ClientId $appId -TenantId $effectiveTenantId -RedirectUri $null -Authority $null
            Token          = $null
            JWTAccessToken = $null
            JWTIdToken     = $null
            IsDefault      = ($script:MSALTokens.Count -eq 0 -or $DefaultToken -eq $true)
            Tenants        = $null
            Organization   = $null
            AuthType       = "BYO"
        }
        $script:MSALTokens.Add($hashId, $TokenInfo) | Out-Null
    }
    else {
        $TokenInfo = $script:MSALTokens[$hashId]
        if($DefaultToken) { $TokenInfo.IsDefault = $true }
    }

    $TokenInfo.Token          = $syntheticToken
    $TokenInfo.JWTAccessToken = $jwt

    if(-not $TokenInfo.Organization) {
        $TokenInfo.Organization = Get-MSALOrganizationInfo $TokenInfo
    }
    # Only refresh the global default-tenant snapshot when this token is the default.
    if($TokenInfo.IsDefault) {
        $script:OrganizationId   = $TokenInfo.Organization.Id
        $script:OrganizationName = $TokenInfo.Organization.displayName
    }

    Write-Log "BYO token registered for tenant '$effectiveTenantId' (expires: $($expiry.LocalDateTime))"
    Register-AuthToken -Provider (Get-AuthProvider "MSAL") -TokenId $TokenInfo.ID -Cloud $TokenInfo.Cloud -Default:([bool]$TokenInfo.IsDefault) | Out-Null

    # Phase 3 + Phase 4: persist per-tenant cloud memory for the BYO flow too.
    try {
        if($cloudVal -and $effectiveTenantId) {
            Save-TenantCloud -TenantId $effectiveTenantId -Cloud $cloudVal
            if($DefaultToken -eq $true) { Save-SettingStoreValue "" "LastLoggedOnCloud" $cloudVal }
        }
    }
    catch { Write-LogDebug "BYO cloud memory write failed: $($_.Exception.Message)" }

    Get-TokenInfo $TokenInfo.Id
}

function Resolve-MSALCertificate
{
    param($Certificate, $CertificatePath, [SecureString]$Password)

    # Already an X509Certificate2 — use directly
    if($Certificate -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) {
        return $Certificate
    }

    # String: try thumbprint lookup in both user and machine stores
    if($Certificate -is [string] -and -not [string]::IsNullOrEmpty($Certificate)) {
        foreach($store in @("Cert:\CurrentUser\My", "Cert:\LocalMachine\My")) {
            $cert = Get-Item "$store\$Certificate" -ErrorAction SilentlyContinue
            if($cert) { return $cert }
        }
        Write-Log "Certificate with thumbprint '$Certificate' not found in Cert:\CurrentUser\My or Cert:\LocalMachine\My" 2
    }

    # File path to .pfx
    if($CertificatePath -and [IO.File]::Exists($CertificatePath)) {
        try {
            if($Password) {
                return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                    $CertificatePath, $Password,
                    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet)
            }
            else {
                return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
            }
        }
        catch {
            Write-LogError "Failed to load certificate from '$CertificatePath'" $_.Exception
        }
    }
    elseif($CertificatePath) {
        Write-Log "Certificate file '$CertificatePath' not found" 3
    }

    return $null
}

#endregion Confidential Client / BYO Token

#region Auth-provider helpers shared by both UI trees
# Lifted out of UI/Extensions/MSGraphAuthenticationUI.ps1 + the Avalonia copy —
# neither function touches WPF/Avalonia controls, they just orchestrate
# providers and populate $script: state that the UI's Show-AuthenticationInfo /
# Set-EnvironmentInfo render. cleanup.md (R2) flagged the duplication.
# Show-AuthenticationInfo / Set-EnvironmentInfo are defined per-UI-tree in
# Extensions/CoreUI.ps1 and resolved at call time, so this Internal/ copy
# works regardless of which tree is loaded.

# Interactive sign-in routed through the active auth provider so the "Sign in with
# another account" buttons honour the user's provider selection (MSAL / MgGraph / ...).
#
# Cloud selection: defaults to the DefaultCloud setting (Phase 1, 2026-05-22). Pass
# -Cloud explicitly to target a different sovereign cloud one-off (e.g. from the
# "Sign in to a different cloud..." profile button). Removed the EntraLoginMenu
# per-login gate — the picker is now an opt-in action, not a per-sign-in modal.
function Invoke-AuthProviderInteractiveLogin
{
    param(
        # Optional one-off cloud override. Public / USGov / USGovDOD / China.
        # If omitted, providers pick up the DefaultCloud setting.
        [string]$Cloud
    )

    $authProvider = Get-AuthProvider
    if(-not $authProvider) {
        Write-Log "No active authentication provider - cannot sign in" 3
        return $null
    }

    # Providers wired into the built-in Connect-* entry points (UsesBuiltInConnectPath)
    # go through Connect-EntraEnvironment directly - it owns the rich sovereign-cloud
    # handling and keeps stack traces clean. Their Connect() forwards here anyway.
    if($authProvider.UsesBuiltInConnectPath) {
        $msalArgs = @{
            ForceInteractive = $true
            DefaultToken     = $true
        }
        # Phase 3: pass -Cloud directly so Connect-EntraEnvironment can distinguish
        # USGov from USGovDOD (legacy -Environment collapses both to "usGov") and
        # persist the right value into TenantCloud_<tid>.
        if($Cloud) { $msalArgs['Cloud'] = $Cloud }
        return (Connect-EntraEnvironment @msalArgs)
    }

    # Generic provider path. Hand an interactive-shaped hashtable to Connect(). Each
    # provider translates: MgGraph calls Connect-MgGraph with no auth args (interactive
    # device-code / browser). Cloud is read by the provider if it supports it;
    # ignored otherwise.
    $providerArgs = @{ DefaultToken = $true; Interactive = $true }
    if($Cloud) { $providerArgs['Cloud'] = $Cloud }
    return $authProvider.Connect($providerArgs)
}

function Update-MSALUserProfile
{
    <#
    .SYNOPSIS
        MSAL-specific UI enrichment: fetches the current MSAL user's /me profile,
        profile photo, and JWT-derived app display name; refreshes the UI's
        environment badge and sign-in bar.

    .DESCRIPTION
        Renamed from Get-MSALUserInfo. The provider-neutral responsibilities
        (setting $script:OrganizationId / $script:OrganizationName /
        $script:CurrentUser from the active provider's GetUserInfo) moved to
        Sync-AuthContextFromProvider in Internal/AuthenticationCore.ps1, which
        is wired to the "AuthenticatedNewToken" event and fires for every
        provider regardless of UI.

        Call this ONLY when the active provider is MSAL — it reads MSAL-only
        state ($script:MSALTokens, $script:MSALDefaultToken.JWTAccessToken
        payload). UI backends invoke it from Invoke-MSALUIEventNewAuthentication
        after guarding on provider Id.
    #>
    # The default token lives in the central registry now. Resolve it to the MSAL
    # entry via Get-FullToken; if the current default belongs to another provider
    # (OAuth/MgGraph) Get-FullToken returns $null and this MSAL-only enrichment
    # correctly no-ops.
    $script:MSALDefaultToken = Get-FullToken (Get-DefaultAuthTokenId)
    if(-not $script:MSALDefaultToken) {
        if($script:UIProvider) { $script:UIProvider.SetEnvironmentInfo() }
        $script:CurrentProfilePhoto = $null
        if($script:UIProvider) { $script:UIProvider.ShowAuthenticationInfo() }
        return
    }

    Write-Log "Get current user (MSAL profile enrichment)"

    if($script:MSALDefaultToken.JWTAccessToken.Payload.idtyp -ne "app")
    {
        $tmpMe = Invoke-MSGraphAPI -Url "ME" -SkipAuthentication -ODataMetadata "Skip"
        if($null -ne $tmpMe -and $tmpMe.creationType -ne "Invitation")
        {
            ### Only get user info from home tenant. Overwrites the minimal
            ### $script:CurrentUser set by Sync-AuthContextFromProvider with the
            ### richer Graph /me object (UI reads givenName / surname / mail / …).
            $script:CurrentUser = $tmpMe
            Write-Log "Get profile picture"
            $script:CurrentProfilePhoto = Join-Path (Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) "CloudAPIPowerShellManagement") "$($script:CurrentUser.Id).jpeg"
            Invoke-MSGraphAPI "me/photos/48x48/`$value" -OutFile $script:CurrentProfilePhoto -SkipAuthentication -NoError | Out-Null
        }
    }
    else
    {
        $script:CurrentProfilePhoto = $null
        $script:CurrentUser = $script:MSALDefaultToken.JWTAccessToken.Payload.app_displayname
    }

    if($script:UIProvider) { $script:UIProvider.SetEnvironmentInfo($script:MSALDefaultToken.Organization.displayName) }
    if($script:UIProvider) { $script:UIProvider.ShowAuthenticationInfo() }
}

# Backward-compat alias. Third-party scripts / older UI extensions may still
# call Get-MSALUserInfo; forward to the renamed function. Remove once nothing
# outside this module references the old name.
Set-Alias -Name Get-MSALUserInfo -Value Update-MSALUserProfile -Scope Script

#endregion Auth-provider helpers shared by both UI trees
