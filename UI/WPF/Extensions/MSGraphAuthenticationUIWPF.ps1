Add-AppEventHandler "AppInitialized" "Invoke-MSALUIAppInitialized"

# Invoke-AuthProviderInteractiveLogin moved to Internal/MSGraphAuthentication.ps1
# (shared by both UI trees — no WPF dependency).

function Show-CloudPickerMenu
{
    <#
    .SYNOPSIS
        Modal dialog to choose a target cloud before sign-in.
    .DESCRIPTION
        Single-combo picker over the flat $script:Clouds taxonomy (Public / USGov / USGovDOD / China).
        Returns the selected Cloud value (string) on OK, $null on Cancel.

        Defaults the selection to -Default if provided, otherwise to the current
        DefaultCloud setting. Pass -Persist to save the choice back to DefaultCloud
        (so subsequent zero-config sign-ins use it); without -Persist the choice is
        one-off and used only for the immediate sign-in.
    #>
    param(
        [string]$Default,
        [switch]$Persist
    )

    $script:loginMenuForm = Initialize-Window ($script:AppUIRootFolder + "\Xaml\CloudPickerMenu.xaml")
    if(-not $script:loginMenuForm) { return $null }

    $script:UIProvider.SetXamlProperty($script:loginMenuForm, "cbCloud", "ItemsSource", $script:Clouds)

    if(-not $Default) { $Default = Get-DefaultCloud }
    $script:UIProvider.SetXamlProperty($script:loginMenuForm, "cbCloud", "SelectedValue", $Default)

    $script:UIProvider.AddXamlEvent($script:loginMenuForm, "btnLogin", "Add_Click", {
        $script:loginMenuForm.DialogResult = $true
        $script:loginMenuForm.Close()
    })

    $script:UIProvider.AddXamlEvent($script:loginMenuForm, "btnCancel", "Add_Click", {
        $script:loginMenuForm.DialogResult = $false
        $script:loginMenuForm.Close()
    })

    $script:loginMenuForm.Owner = $script:window
    $script:loginMenuForm.Icon  = $script:Window.Icon
    $ok = $script:loginMenuForm.ShowDialog()
    if(-not $ok) { return $null }

    $cloudValue = $script:UIProvider.GetXamlProperty($script:loginMenuForm, "cbCloud", "SelectedValue")
    if(-not $cloudValue) { $cloudValue = "Public" }
    if($Persist) {
        Set-DefaultCloud $cloudValue
    }
    return [string]$cloudValue
}

function Get-MSALUserProfile
{
    param($Size = 32, $FontSize = 20, $Color = "Blue", [Switch]$Popup)

    Write-LogDebug "Create Profile Ellipse"

    # Refresh $script:MSALAccounts from the persistent MSAL cache whenever the popup is
    # built. Without this, BYO/Confidential sessions (which never go through the interactive
    # MSAL path) show an empty "cached users" list even when the on-disk cache has
    # interactive accounts from previous sessions. Cheap — Get-AccountsAsync reads the
    # already-deserialized cache, no network call.
    # Lazy-MSAL gate: this runs on every startup draw (window Loaded ->
    # Show-AuthenticationInfo), including signed-out sessions. When the MSAL runtime
    # was never loaded this session, skip the refresh instead of forcing the DLL
    # load - the login-click handler re-enumerates via GetCachedAccounts (which
    # ensure-loads), so the popup list is fully populated on first click.
    if($Popup -and ($script:MSALApps.Count -gt 0 -or $script:MSALPrereqLoaded)) {
        try {
            $publicClientApp = $script:MSALApps | Select-Object -First 1
            if(-not $publicClientApp) {
                # No public-client app exists yet (pure BYO/CC session). Create one so
                # the file cache deserializes; this does NOT trigger any auth.
                $publicClientApp = New-MSALApp
            }
            if($publicClientApp) {
                $script:MSALAccounts = $publicClientApp.GetAccountsAsync().GetAwaiter().GetResult()
            }
        }
        catch {
            Write-LogDebug "Failed to refresh cached account list: $($_.Exception.Message)"
        }
    }

    $grdUserProfile = $null

    # Treat an expired default token the same as "not signed in" so the title-bar
    # reverts to the Sign-in icon (which prompts interactive login on click) instead
    # of showing a stale logged-in avatar. Test-DefaultTokenExpired is provider-
    # agnostic and returns $false when expiry is unknown/SDK-managed.
    $tokenExpired = $false
    if(Get-Command Test-DefaultTokenExpired -ErrorAction SilentlyContinue) {
        try { $tokenExpired = Test-DefaultTokenExpired } catch { }
    }

    # Signed-in state is provider-agnostic: $script:CurrentUser is set by
    # Sync-AuthContextFromProvider for EVERY provider (MSAL/OAuth/MgGraph). Rich popup
    # fields below come from the provider's GetUserInfo() and the photo from
    # $script:CurrentProfilePhoto - this function reads no provider-specific globals.
    if(-not $script:CurrentUser -or $tokenExpired)
    {
        #########################################################################################################
        ### Build login button when no user is logged on (or the token has expired)
        #########################################################################################################

        Write-LogDebug "Add login button"
        $grdUserProfile = [System.Windows.Controls.Border]::new()
        $icon = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\Icons\Logon.xaml"))
        $icon.Width = $Size
        $icon.Height = $Size
        $grdUserProfile.Background = "#01000000"
        $grdUserProfile.Child = $icon

        $lnkButton = [System.Windows.Controls.Button]::new()
        $lnkButton.Content = $grdUserProfile
        $lnkButton.Cursor = "Hand"
        $lnkButton.Style = $script:window.TryFindResource("ContentButton") 
        $lnkButton.add_Click({
            Write-Log "Login icon clicked"
            try {
                # Cached-account list comes from the active provider so providers that don't
                # surface cached identities (MgGraph) just show an empty list and the user
                # is taken straight to the interactive flow.
                $authProviderClick = $null
                try { $authProviderClick = Get-AuthProvider } catch { Write-LogError "Get-AuthProvider failed in login-click handler" $_.Exception }

                Write-Log "Login-click: active provider = '$(if($authProviderClick){$authProviderClick.Id}else{'<none>'})'"

                $cachedRows = @()
                if($authProviderClick -and $authProviderClick.SupportsCachedUsers) {
                    try { $cachedRows = @($authProviderClick.GetCachedAccounts()) } catch { $cachedRows = @() }
                }

                if($cachedRows.Count -eq 0)
                {
                    # No cached users for the active provider — go straight to interactive
                    # auth via whichever provider is active.
                    Write-Status "Signing in..."
                    Write-Log "Login-click: calling Invoke-AuthProviderInteractiveLogin"
                    $tokenInfo = Invoke-AuthProviderInteractiveLogin
                    if($tokenInfo) {
                        Write-Log "Login-click: interactive login returned a token, refreshing UI"
                        try { Get-MSALUserInfo } catch { Write-LogError "Get-MSALUserInfo failed after interactive login" $_.Exception }
                        try { Show-AuthenticationInfo } catch { Write-LogError "Show-AuthenticationInfo failed after interactive login" $_.Exception }
                        try { Set-EnvironmentInfo } catch { Write-LogError "Set-EnvironmentInfo failed after interactive login" $_.Exception }
                    }
                    else {
                        Write-Log "Login-click: interactive login returned nothing (user cancelled or auth failed)" 2
                    }
                    Write-Status ""
                }
            else
            {
                # Add list of cached users + a 'Sign in with a different account' option

                $xaml = Get-Content ($script:AppUIRootFolder + "\Xaml\LoginPanel.Xaml") -Encoding UTF8
                $loginPanel = [Windows.Markup.XamlReader]::Parse($xaml)
                $otherLogins = $loginPanel.FindName("grdAccounts")
                foreach($row in $cachedRows)
                {
                    # Filter Microsoft personal accounts. .Native is the underlying
                    # provider-specific object (MSAL IAccount for the MSAL provider) —
                    # Add-CachedUser reads its UserName / HomeAccountId fields.
                    if($row.Native -and (Test-IsPersonalMSAAccount $row.Native)) { continue }
                    Add-CachedUser $row.Native $otherLogins
                }
                
                #########################################################################################################
                ### Add login button
                #########################################################################################################
                $grdAccount = [System.Windows.Controls.Grid]::new()
                $cd = [System.Windows.Controls.ColumnDefinition]::new()                
                $grdAccount.ColumnDefinitions.Add($cd)
                $cd = [System.Windows.Controls.ColumnDefinition]::new()
                $cd.Width = [double]::NaN   
                $grdAccount.ColumnDefinitions.Add($cd)

                $icon = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\Icons\Logon.xaml"))
                $icon.Width = 24
                $icon.Height = 24
                $icon.Margin = "0,0,5,0"
                $grdAccount.Children.Add($icon) | Out-Null

                $lbObj = [Windows.Markup.XamlReader]::Parse("<TextBlock $script:WPFNS>Sign in with a different account</TextBlock>")
                $lbObj.SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
                #$lbObj.Style = $script:window.TryFindResource("HoverUnderlineStyle")
                $grdAccount.Children.Add($lbObj) | Out-Null

                $lnkButton = [System.Windows.Controls.Button]::new()
                $lnkButton.Content = $grdAccount
                $lnkButton.Style = $script:window.TryFindResource("LinkButton")
                $lnkButton.Margin = "0,5,0,0"
                $lnkButton.Cursor = "Hand"
                $lnkButton.add_Click({
                    Write-Status "Logging in..."
                    $script:UIProvider.HidePopup()
                    if((Invoke-AuthProviderInteractiveLogin))
                    {
                        ## ToDo: Show graph objects
                    }
                    Write-Status ""
                })

                Add-GridObject $otherLogins $lnkButton

                $loginPanel.Tag = $this.Content

                $loginPanel.Add_Loaded({param($Obj, $E)
                    $point = $Obj.Tag.TransformToAncestor($script:window).Transform([System.Windows.Point]::new(0,0));
                    [System.Windows.Controls.Canvas]::SetLeft($Obj,($point.X - $Obj.ActualWidth + $Obj.Tag.ActualWidth))
                    [System.Windows.Controls.Canvas]::SetTop($Obj,($point.Y + $Obj.Tag.ActualHeight))
                })

                $script:UIProvider.ShowPopup($loginPanel)
            }
            }
            catch {
                Write-LogError "Login-click handler threw" $_.Exception
            }
        })

        return $lnkButton
    }

    #########################################################################################################
    ### Build the ellipse image for the Profile Info
    #########################################################################################################

    # Provider-agnostic user view. Drives org / app / auth / expiry / app-login below.
    # Identity name, initials and photo still come from $script:CurrentUser /
    # $script:CurrentProfilePhoto. No MSAL globals here, so OAuth/MgGraph render too.
    $provider = $null
    try { $provider = Get-AuthProvider } catch { }
    $userInfo = $null
    if($provider) { try { $userInfo = $provider.GetUserInfo((Get-DefaultAuthTokenId)) } catch { } }
    $isAppLogin = [bool]($userInfo -and $userInfo.AuthType -in @('Confidential','ClientCredential','ManagedIdentity','WorkloadFederation'))

    if($script:CurrentUser.givenName -and $script:CurrentUser.surname)
    {
        $initials = "$($script:CurrentUser.givenName[0])$($script:CurrentUser.surname[0])".ToUpper()
    }
    elseif($script:CurrentUser.userPrincipalName)
    {
        $initials = "$($script:CurrentUser.userPrincipalName[0])".ToUpper()
    }
    elseif($isAppLogin)
    {
        $initials = "APP"
    }

    $grdUserProfile = Get-MSALUserPhotoEllips -size $Size -fontSize $FontSize -Color $Color 
    
    if($Popup)
    {
        # Hide the popup when mouse button is clicked anywhere
        $grdUserProfile.add_MouseLeftButtonDown(({param($Obj, $E)
            if(-not $script:grdProfileInfo) { return }
            $script:UIProvider.ShowPopup($script:grdProfileInfo)
        }))

        try 
        {
            #########################################################################################################
            ### Build Profile Info for current user
            #########################################################################################################

            $script:grdProfileInfo = $null
            $xaml = Get-Content ($script:AppUIRootFolder + "\Xaml\ProfileInfo.Xaml") -Encoding UTF8
            $script:grdProfileInfo = [Windows.Markup.XamlReader]::Parse($xaml)
            $script:grdProfileInfo.Tag = $grdUserProfile
            $grdUserProfile.Tag = $script:grdProfileInfo
            # Org name from the provider's tenant name, falling back to the neutral
            # $script:OrganizationName (both set for every provider).
            $orgNameForPopup = if($userInfo -and $userInfo.TenantName) { $userInfo.TenantName } else { $script:OrganizationName }
            $script:UIProvider.SetXamlProperty($script:grdProfileInfo, "txtOrganization", "Text", ([string]$orgNameForPopup))
            if($isAppLogin)
            {
                $script:UIProvider.SetXamlProperty($script:grdProfileInfo, "txtUsername", "Text", "App Login")
            }
            else
            {
                $script:UIProvider.SetXamlProperty($script:grdProfileInfo, "txtUsername", "Text", $script:CurrentUser.displayName)
                $script:UIProvider.SetXamlProperty($script:grdProfileInfo, "txtLogonName", "Text", $script:CurrentUser.userPrincipalName)
            }

            # App info from the provider-agnostic user view (GetUserInfo resolves it
            # from the JWT app claims / EntraApp per provider).
            $appName = if($userInfo) { $userInfo.AppName } else { $null }
            $appId   = if($userInfo) { $userInfo.AppId }   else { $null }

            Write-LogDebug "Profile Info app=$appName id=$appId"
            $script:UIProvider.SetXamlProperty($script:grdProfileInfo, "txtAppName", "Text", $appName)
            $script:UIProvider.SetXamlProperty($script:grdProfileInfo, "txtAppId",   "Text", $appId)

            # Auth method + expiry from the provider-agnostic user view.
            $authLabel  = $null
            $expiryLabel = $null
            if($userInfo) {
                $authTypeRaw = if($userInfo.AuthType) { $userInfo.AuthType } else { "Interactive" }
                $modeName = switch ($authTypeRaw) {
                    "BYO"              { "Bring-your-own token" }
                    "Confidential"     { "Client credentials" }
                    "ClientCredential" { "Client credentials" }
                    "ManagedIdentity"  { "Managed identity" }
                    default            { "Interactive" }
                }
                $providerDisplay = if($provider) { $provider.DisplayName } else { "" }
                $authLabel = "$providerDisplay - $modeName"
                if($userInfo.ExpiresOn) { $expiryLabel = "Token expires $($userInfo.ExpiresOn.ToString('g'))" }
            }
            $script:UIProvider.SetXamlProperty($script:grdProfileInfo, "txtAuthMethod", "Text",       $authLabel)
            $script:UIProvider.SetXamlProperty($script:grdProfileInfo, "txtAuthMethod", "Visibility", "Visible")
            if($expiryLabel) {
                $script:UIProvider.SetXamlProperty($script:grdProfileInfo, "txtAuthExpiry", "Text",       $expiryLabel)
                $script:UIProvider.SetXamlProperty($script:grdProfileInfo, "txtAuthExpiry", "Visibility", "Visible")
            }

            $script:UIProvider.AddXamlEvent($script:grdProfileInfo, "btnProfileClose", "add_click", { $script:UIProvider.HidePopup() })

            # Get the elips with only the photo in a larger size for the popup info
            $tmpObj = Get-MSALUserPhotoEllips -size 64 -fontSize 32

            $profileGrid =  $script:grdProfileInfo.FindName("ProfileInfo")
            if($tmpObj -and $profileGrid)
            {
                $tmpObj.SetValue([System.Windows.Controls.Grid]::RowProperty,1)
                $tmpObj.SetValue([System.Windows.Controls.Grid]::RowSpanProperty,2)
            }

            if($tmpObj)
            {
                $profileGrid.Children.Add($tmpObj) | Out-Null
            }

            if($isAppLogin)
            {
                $tmpObj.Visibility = "Collapsed"
            }
        
            $script:grdProfileInfo.Add_Loaded({param($Obj, $E)
                $point = $Obj.Tag.TransformToAncestor($script:window).Transform([System.Windows.Point]::new(0,0));
                [System.Windows.Controls.Canvas]::SetLeft($Obj,($point.X - $Obj.ActualWidth + $Obj.Tag.ActualWidth)) 
                [System.Windows.Controls.Canvas]::SetTop($Obj,($point.Y + $Obj.Tag.ActualHeight))
            })

            # Consent prompt is delegated-only - for app-only tokens (client-credential
            # / certificate / managed identity) there is no interactive user to walk
            # through the consent UI, so the link is hidden. $isAppLogin derives from the
            # provider-agnostic GetUserInfo().AuthType.
            if(-not $isAppLogin)
            {
                #########################################################################################################
                ### Show / Hide consent button
                #########################################################################################################
                $script:userEllipsGrid = $tmpObj
                if(($script:missingPermissions | Measure-Object).Count -eq 0)
                {
                    $script:UIProvider.SetXamlProperty($script:userEllipsGrid, "lnkRequestConsent", "Visibility", "Collapsed")
                }
                else
                {
                    $script:UIProvider.SetXamlProperty($script:userEllipsGrid, "lnkRequestConsent", "Visibility", "Visible")
                }
                $script:UIProvider.AddXamlEvent($script:userEllipsGrid, "lnkRequestConsent", "add_Click", {
                    Start-MSALConsentPrompt
                })
                
                $otherLogins = $script:grdProfileInfo.FindName("grdCachedAccounts")

                #########################################################################################################
                ### Add cached users
                #########################################################################################################
                # Cached-account list comes from the active provider. Providers that don't
                # surface cached identities (MgGraph today, SupportsCachedUsers=$false)
                # return an empty array so the section stays empty — no leftover MSAL
                # users shown in MgGraph mode.
                $cachedRows = @()
                if($activeProvider -and $activeProvider.SupportsCachedUsers) {
                    try { $cachedRows = @($activeProvider.GetCachedAccounts()) } catch { $cachedRows = @() }
                }
                if((Get-SettingValue "SortAccountList") -eq $true) {
                    $cachedRows = @($cachedRows | Sort-Object -Property Username)
                }

                foreach($row in $cachedRows)
                {
                    if(-not $row.Native) { continue }
                    $Account = $row.Native

                    # Skip current logged on user (matched on UPN / user id from the
                    # provider-agnostic user view).
                    if($script:CurrentUser.userPrincipalName -eq $Account.Username -or
                    ($userInfo -and $userInfo.UserId -eq $Account.HomeAccountId.ObjectId)) { continue }

                    # Skip Microsoft personal (MSA) accounts — not usable for Intune/Graph
                    if(Test-IsPersonalMSAAccount $Account) { continue }

                    Add-CachedUser $Account $otherLogins
                }
                
                #########################################################################################################
                ### Add login with another user
                #########################################################################################################
                $grdAccount = [System.Windows.Controls.Grid]::new()
                $cd = [System.Windows.Controls.ColumnDefinition]::new()                
                $grdAccount.ColumnDefinitions.Add($cd)
                $cd = [System.Windows.Controls.ColumnDefinition]::new()
                $cd.Width = [double]::NaN   
                $grdAccount.ColumnDefinitions.Add($cd)

                $icon = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\Icons\Logon.xaml"))
                $icon.Width = 24
                $icon.Height = 24
                $icon.Margin = "0,0,5,0"
                $grdAccount.Children.Add($icon) | Out-Null

                $lbObj = [Windows.Markup.XamlReader]::Parse("<TextBlock $script:WPFNS>Sign in with a different account</TextBlock>")
                $lbObj.SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
                #$lbObj.Style = $script:window.TryFindResource("HoverUnderlineStyle")
                $grdAccount.Children.Add($lbObj) | Out-Null

                $lnkButton = [System.Windows.Controls.Button]::new()
                $lnkButton.Content = $grdAccount
                $lnkButton.Style = $script:window.TryFindResource("LinkButton")
                $lnkButton.Margin = "0,5,0,0"
                $lnkButton.Cursor = "Hand"
                $lnkButton.Tag = $Account
                $lnkButton.add_Click({
                    Write-Status "Logging in..."
                    $script:UIProvider.HidePopup()
                    if((Invoke-AuthProviderInteractiveLogin))
                    {
                        # ToDo Show-GraphObjects
                    }
                    Write-Status ""
                })                    

                $otherLogins = $script:grdProfileInfo.FindName("grdLoginAccount")

                Add-GridObject $otherLogins $lnkButton

                #########################################################################################################
                ### Add "Sign in to a different cloud..." (opt-in picker; replaces the old EntraLoginMenu gate)
                #########################################################################################################
                $grdCloud = [System.Windows.Controls.Grid]::new()
                $cd = [System.Windows.Controls.ColumnDefinition]::new()
                $grdCloud.ColumnDefinitions.Add($cd)
                $cd = [System.Windows.Controls.ColumnDefinition]::new()
                $cd.Width = [double]::NaN
                $grdCloud.ColumnDefinitions.Add($cd)

                # Reuse the Logon icon so the row visually pairs with "Sign in with a
                # different account" above it; the text on the right is what distinguishes them.
                $iconCloud = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\Icons\Logon.xaml"))
                $iconCloud.Width = 24
                $iconCloud.Height = 24
                $iconCloud.Margin = "0,0,5,0"
                $grdCloud.Children.Add($iconCloud) | Out-Null

                $lbCloud = [Windows.Markup.XamlReader]::Parse("<TextBlock $script:WPFNS>Sign in to a different cloud...</TextBlock>")
                $lbCloud.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)
                $grdCloud.Children.Add($lbCloud) | Out-Null

                $btnCloud = [System.Windows.Controls.Button]::new()
                $btnCloud.Content = $grdCloud
                $btnCloud.Style = $script:window.TryFindResource("LinkButton")
                $btnCloud.Margin = "0,5,0,0"
                $btnCloud.Cursor = "Hand"
                $btnCloud.ToolTip = "Sign in to US Government (GCC High / DoD) or China cloud. To make a choice permanent, change Default cloud in Settings."
                $btnCloud.add_Click({
                    $script:UIProvider.HidePopup()
                    $picked = Show-CloudPickerMenu
                    if($picked) {
                        Write-Status "Signing in to $picked cloud..."
                        if((Invoke-AuthProviderInteractiveLogin -Cloud $picked)) {
                            # ToDo Show-GraphObjects
                        }
                        Write-Status ""
                    }
                })
                Add-GridObject $otherLogins $btnCloud

                $otherLogins = $script:grdProfileInfo.FindName("grdTenantAccounts")
                
                if(($script:AccessibleTenants | Measure-Object).Count -gt 1)
                {
                    #########################################################################################################
                    ### Add switch to another tenant
                    #########################################################################################################
                    $lbObj = [Windows.Markup.XamlReader]::Parse("<TextBlock $script:WPFNS><Bold>Tenants:</Bold></TextBlock>")
                    $lbObj.Margin = "0,5,0,0"

                    if((Get-SettingValue "SortTenantList") -eq $true)
                    {
                        $tenants = $script:AccessibleTenants | Sort-Object -Property DisplayName
                    }
                    else
                    {
                        $tenants = $script:AccessibleTenants
                    }

                    Add-GridObject $otherLogins $lbObj
                    foreach($tenant in $tenants)
                    {
                        try
                        {
                            $lbObj = [Windows.Markup.XamlReader]::Parse("<TextBlock $script:WPFNS HorizontalAlignment=`"Stretch`"><Bold>$($tenant.DisplayName)</Bold><LineBreak/>$($tenant.defaultDomain)<LineBreak/>$($tenant.tenantId)</TextBlock>")

                            if($tenant.tenantId -ne $userInfo.TenantId)
                            {
                                $lbObj.Style = $script:window.TryFindResource("HoverUnderlineStyleWithBackground")
                                $lbObj.HorizontalAlignment = "Stretch"
                                $lnkButton = [System.Windows.Controls.Button]::new()
                                $lnkButton.Content = $lbObj
                                $lnkButton.HorizontalAlignment = "Stretch"
                                $lnkButton.Style = $script:window.TryFindResource("ContentButton") 
                                $lnkButton.Margin = "0,5,0,0"
                                $lnkButton.Cursor = "Hand"
                                $lnkButton.Tag = $tenant
                                $lnkButton.add_Click({
                                    Write-Status "Logging in to $($this.Tag.DisplayName)"
                                    # Set authority to selected tenant
                                    $script:UIProvider.HidePopup()
                                    if((Connect-EntraEnvironment -User ($script:MSALAccounts | Where-Object UserName -eq $script:CurrentUser.userPrincipalName).Username -TenantId $this.Tag.tenantId -DefaultToken))
                                    {
                                        # ToDo: Show-GraphObjects
                                    }
                                    Write-Status ""
                                })
                                Add-GridObject $otherLogins $lnkButton
                            }
                            else
                            {
                                $lbObj.Background = $script:window.TryFindResource("SelectedRowBackgroundColor")
                                $lbObj.Margin = "0,5,0,0"
                                Add-GridObject $otherLogins $lbObj
                            }                        
                        }
                        catch {}
                    }
                }
            }

            #########################################################################################################
            ### Add event handling
            #########################################################################################################
            # Token-info button. For MSAL, shows AuthenticationResult fields. For
            # other providers, shows whatever the provider's session exposes (e.g.
            # Get-MgContext for MgGraph).
            $script:UIProvider.AddXamlEvent($tmpObj, "lnkTokeninfo", "add_Click", {
                $tokenArr = @()
                $providerNow = $null
                try { $providerNow = Get-AuthProvider } catch { }

                # Provider-agnostic: each provider returns its own native session rows
                # (MSAL AuthenticationResult fields, MgGraph Get-MgContext, ...).
                if($providerNow) { $tokenArr = @($providerNow.GetSessionInfoRows()) }

                $dg = [System.Windows.Controls.DataGrid]::new()
                Set-GridPixelScrolling $dg
                $dg.ItemsSource = ($tokenArr | Select-Object Name, Value)
                $script:UIProvider.ShowModalForm("Session Info", $dg)
            })

            $script:UIProvider.AddXamlEvent($tmpObj, "lnkAccessTokenInfo", "add_Click", {
                # Decode the active provider's access token on demand (works for any
                # provider whose access token is a JWT - MSAL, OAuth).
                $providerNow = $null
                try { $providerNow = Get-AuthProvider } catch { }
                if($providerNow) {
                    $raw = $providerNow.GetAccessToken(0, "https://$(Get-GraphDomain)")
                    if($raw) {
                        $jwt = Get-JWTtoken $raw
                        Show-MSALDecodedToken $jwt "Access Token Info"
                    }
                    else {
                        $script:UIProvider.ShowMessageBox("No access token available from provider '$($providerNow.Id)'.", "Access Token Info", "OK", "Information") | Out-Null
                    }
                }
            })

            $script:UIProvider.AddXamlEvent($tmpObj, "lnkIdTokenInfo", "add_Click", {
                # ID token decode via the provider (only providers that surface an id
                # token return one; the button is hidden otherwise).
                $providerNow = $null
                try { $providerNow = Get-AuthProvider } catch { }
                $idJwt = if($providerNow) { $providerNow.GetIdTokenJwt((Get-DefaultAuthTokenId)) } else { $null }
                if($idJwt) {
                    Show-MSALDecodedToken $idJwt "Id Token Info"
                }
                else {
                    $script:UIProvider.ShowMessageBox("ID token is not available from the active provider.", "Id Token Info", "OK", "Information") | Out-Null
                }
            })

            # Permissions: token scopes x Intune role, per policy type (EffectivePermissionsUIWPF.ps1).
            $script:UIProvider.AddXamlEvent($tmpObj, "lnkEffectivePermissions", "add_Click", { Show-EffectivePermissionsDialog })

            # Phase 3: capability-flag-aware buttons. Disable buttons the active
            # provider doesn't support, so the user gets clear feedback instead of
            # silent no-ops.
            #
            # Refresh:
            #   * Required: provider.SupportsRefresh AND it's not a BYO token (BYO is
            #     externally-supplied; even on a provider that supports refresh, BYO
            #     can't be refreshed).
            if($provider) {
                $isByo = [bool]($userInfo -and $userInfo.AuthType -eq "BYO")
                if(-not $provider.SupportsRefresh) {
                    $script:UIProvider.SetXamlProperty($tmpObj, "lnkForceRefresh", "IsEnabled", $false)
                    $script:UIProvider.SetXamlProperty($tmpObj, "lnkForceRefresh", "ToolTip",   "$($provider.DisplayName) does not support manual refresh.")
                }
                elseif($isByo) {
                    $script:UIProvider.SetXamlProperty($tmpObj, "lnkForceRefresh", "IsEnabled", $false)
                    $script:UIProvider.SetXamlProperty($tmpObj, "lnkForceRefresh", "ToolTip",   "BYO tokens cannot be refreshed. Re-run Connect-IntuneManagement with a new token.")
                }

                # Inspector buttons are shown only when the provider actually has that
                # data (self-describing, provider-agnostic):
                #   * Session Info: only when the provider returns session rows (OAuth has none).
                #   * Access Token: always - decoded on demand from GetAccessToken.
                #   * Id Token: only when the provider surfaces an id token (MSAL).
                if(@($provider.GetSessionInfoRows()).Count -eq 0) {
                    $script:UIProvider.SetXamlProperty($tmpObj, "lnkTokeninfo", "Visibility", "Collapsed")
                }
                if($null -eq $provider.GetIdTokenJwt((Get-DefaultAuthTokenId))) {
                    $script:UIProvider.SetXamlProperty($tmpObj, "lnkIdTokenInfo", "Visibility", "Collapsed")
                }
            }

            $script:UIProvider.AddXamlEvent($tmpObj, "lnkForceRefresh", "add_Click", {
                Write-Status "Refreshing the token"

                $providerNow = $null
                try { $providerNow = Get-AuthProvider } catch { }

                $refreshed = $false
                if($providerNow -and $providerNow.UsesBuiltInConnectPath) {
                    # Providers on the built-in Connect path refresh a specific token id
                    # via Connect-EntraEnvironment -ForceRefresh.
                    $refreshed = [bool](Connect-EntraEnvironment -ForceRefresh -TokenId (Get-DefaultAuthTokenId))
                }
                elseif($providerNow) {
                    # Generic provider Refresh (e.g. MgGraph re-runs Connect-MgGraph).
                    # Pass the REAL default token id: ids are allocated from 1, so the
                    # literal 0 this used to send matched nothing in OAuth's token table
                    # and its Refresh returned $false before trying anything - the button
                    # just closed the popup. MgGraph ignores the id, so it is unaffected.
                    $refreshed = [bool]$providerNow.Refresh((Get-DefaultAuthTokenId))
                }

                if(-not $refreshed)
                {
                    $script:UIProvider.HidePopup()
                }
                else
                {
                    # Refresh title-bar profile and view menu after a successful refresh
                    # so any new role assignments / scope changes are reflected.
                    try { Get-MSALUserInfo } catch { Write-LogError "Get-MSALUserInfo failed after refresh" $_.Exception }
                    if(Get-Command Show-ViewMenu -ErrorAction SilentlyContinue) { Show-ViewMenu }
                }
                Write-Status ""
            })

            $script:UIProvider.AddXamlEvent($tmpObj, "lnkLogout", "add_Click", {
                $script:UIProvider.HidePopup()

                if(Disconnect-EntraEnvironment)
                {
                    # ToDo: Show-GraphObjects
                }
            })

            if($isAppLogin)
            {
                $script:UIProvider.SetXamlProperty($tmpObj, "lnkLogout", "Visibility", "Collapsed")
            }
        }
        catch {
            Write-LogError "Failed to create profile information object. Error: " $_.Exception   
        }
    }

    $grdUserProfile
}

# Filters out Microsoft personal / consumer (MSA) accounts. Those accounts live on the
# well-known tenant 9188040d-6c67-4c5b-b112-36a304b66dad and cannot be used to sign into
# Entra-tenant resources (Intune / Graph) — surfacing them in the picker only confuses
# the user. Keep work/school accounts plus accounts where tenant info isn't available.
function Test-IsPersonalMSAAccount {
    param($Account)
    if(-not $Account -or -not $Account.HomeAccountId) { return $false }
    $tid = $Account.HomeAccountId.TenantId
    return ($tid -eq "9188040d-6c67-4c5b-b112-36a304b66dad")
}

function Add-CachedUser 
{
    param($Account, $ParentObj)

    try
    {
        $grdAccount = [System.Windows.Controls.Grid]::new()

        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        $grdAccount.ColumnDefinitions.Add($cd) # Login

        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        $cd.Width = [double]::NaN   
        $grdAccount.ColumnDefinitions.Add($cd) # Forget

        $grdLogin = [System.Windows.Controls.Grid]::new()
        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        $grdLogin.ColumnDefinitions.Add($cd)
        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        $cd.Width = [double]::NaN   
        $grdLogin.ColumnDefinitions.Add($cd)

        $icon = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\Icons\LoggedOnUser.xaml"))
        $icon.Width = 24
        $icon.Height = 24
        $icon.Margin = "0,0,5,0"
        $grdLogin.Children.Add($icon) | Out-Null

        $tenantName = Get-SettingStoreValue $Account.HomeAccountId.TenantId "_Name" $Account.HomeAccountId.TenantId

        $lbObj = [Windows.Markup.XamlReader]::Parse("<TextBlock $script:WPFNS>$($Account.UserName)<LineBreak/>$($tenantName)</TextBlock>")
        $lbObj.SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
        $grdLogin.Children.Add($lbObj) | Out-Null

        $lnkButton = [System.Windows.Controls.Button]::new()
        $lnkButton.Content = $grdLogin
        $lnkButton.Style = $script:window.TryFindResource("LinkButton") 
        $lnkButton.Margin = "0,5,0,0"
        $lnkButton.Cursor = "Hand"
        $lnkButton.Tag = $Account
        $lnkButton.add_Click({
            Write-Status "Logging in with $($this.Tag.UserName)"
            $script:UIProvider.HidePopup()
            
            if((Connect-EntraEnvironment -User $this.Tag.UserName -DefaultToken))
            {
                ## ToDo: Show-GraphObjects
            }
            Write-Status ""
        })

        $grdAccount.Children.Add($lnkButton) | Out-Null

        # Add Forget user icon
        $icon = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\Icons\Bin.xaml"))
        $icon.Width = 16
        $icon.Height = 16
        $icon.Margin = "5,5,0,0"
        
        $lnkButton = [System.Windows.Controls.Button]::new()
        $lnkButton.ToolTip = "Forget"
        $lnkButton.Content = $icon
        $lnkButton.Style = $script:window.TryFindResource("LinkButton") 
        $lnkButton.Margin = "0,5,0,0"
        $lnkButton.Cursor = "Hand"
        $lnkButton.Tag = $Account
        $lnkButton.SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
        $lnkButton.add_Click({
            # Truly evict the MSAL account: removes from in-memory MSALTokens (if signed in
            # under this account) AND from the persistent on-disk cache so it stops showing
            # up in future sessions. The old code called Disconnect-EntraEnvironment with an
            # Account object, which coerced to int 0 (default token) — never the intended one.
            Write-Status "Removing $($this.Tag.UserName)"
            try {
                # If this account is currently logged in, disconnect properly first so we
                # promote a new default and fire the event.
                $matchingTokens = @($script:MSALTokens.Values | Where-Object {
                    $_.Token -and $_.Token.Account -and
                    $_.Token.Account.HomeAccountId.Identifier -eq $this.Tag.HomeAccountId.Identifier
                })
                foreach($tok in $matchingTokens) { Disconnect-EntraEnvironment -TokenID $tok.Id }

                # Always also evict from the persistent MSAL cache (covers cached-but-not-signed-in).
                Remove-MSALAccount -Account $this.Tag

                $this.Parent.Parent.Children.Remove($this.Parent)
            }
            catch {
                Write-LogError "Failed to forget account $($this.Tag.UserName)" $_.Exception
            }
            Write-Status ""
        })

        $grdAccount.Children.Add($lnkButton) | Out-Null
        
        Add-GridObject $ParentObj $grdAccount
    }
    catch {}     
}

function Get-MSALUserPhotoEllips
{
    param($Size = 32, $FontSize = 20, $Color = "Blue")

    $grdUserPhoto = [System.Windows.Controls.Grid]::new()

    $ellipse = [System.Windows.Shapes.Ellipse]::new()
    $ellipse.Width = $Size
    $ellipse.Height = $Size
    $ellipse.Fill = $Color
    $ellipse.Stroke = "#FFFF00FF"
    $ellipse.StrokeThickness = "0"

    $grdUserPhoto.Children.Add($ellipse) | Out-Null

    $tb = [System.Windows.Controls.TextBlock]::new()
    $tb.FontSize = $FontSize
    $tb.Foreground = "White"
    #$tb.FontFamily=""
    $tb.FontWeight = "Bold"
    #$tb.TextLineBounds="Tight"
    $tb.VerticalAlignment="Center"
    $tb.HorizontalAlignment="Center"
    #$tb.IsTextScaleFactorEnabled="False"
    $tb.Text = $initials

    $grdUserPhoto.Children.Add($tb) | Out-Null

    if($script:CurrentProfilePhoto -and [IO.File]::Exists($script:CurrentProfilePhoto))   
    {        
        Write-LogDebug "Create image"
        $img = [System.Windows.Media.Imaging.BitmapImage]::new()
        $img.BeginInit()
        $img.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $img.UriSource = [System.Uri]::new($script:CurrentProfilePhoto)
        $img.EndInit()        
        $ib = [System.Windows.Media.ImageBrush]::new()
        $ib.ImageSource = $img

        $ellipse = [System.Windows.Shapes.Ellipse]::new()
        $ellipse.Width = $Size
        $ellipse.Height = $Size
        $ellipse.FlowDirection="LeftToRight"
        $ellipse.Fill = $ib
        $grdUserPhoto.Children.Add($ellipse) | Out-Null
    }

    $grdUserPhoto
}

function Show-MSALDecodedToken {
    param (
        $TokenData,
        $Title
    )

    if(-not $TokenData.Header) { return }

    $tokenArr = @()
    foreach($prop in ($TokenData.Header | Get-Member | Where-Object MemberType -eq NoteProperty))
    {
        $tokenArr += New-Object PSObject -Property @{
            Name=$prop.Name
            Value=$TokenData.Header."$($prop.Name)"
        }
    }

    # Each claim yields one or more GRID ROWS rather than one row holding a
    # multi-line string. Newline-joining a long list makes a single row several
    # times taller than the rest (8 permissions measured 147px against 33px in
    # the Avalonia grid), and a virtualizing DataGrid estimates its scroll extent
    # from the rows it has realized - so that one outlier makes the extent swing
    # as it scrolls in and out of view, giving a jumpy scrollbar and a mouse
    # wheel that appears to run through the table several times. One value per
    # row keeps every row a uniform height and stays fully readable.
    foreach($prop in ($TokenData.Payload | Get-Member | Where-Object MemberType -eq NoteProperty))
    {
        $raw = $TokenData.Payload."$($prop.Name)"

        if($prop.Name -in @("exp","iat","nbf","xms_tcdt"))
        {
            $values = @([datetime]::new(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc).AddSeconds($raw).ToLocalTime())
        }
        elseif($prop.Name -in @("acrs","amr"))
        {
            # Short arrays; a single ";"-joined line stays well inside one row.
            $values = @($raw -join ";")
        }
        elseif($prop.Name -in @("wids"))
        {
            if(-not $script:AADRoles)
            {
                # This will fail if RoleManagement.Read.Directory permission is not granted. Use -NoError to hide any problems
                $script:AADRoles = (Invoke-MSGraphAPI -url "/directoryRoles?`$select=roleTemplateId,displayName" -ODataMetadata "minimal" -Noerror).value
            }
            $values = @()
            foreach($wid in $raw)
            {
                $text = $wid
                $role = ($script:AADRoles | Where-Object roleTemplateId -eq $wid)
                if($role)
                {
                    $text = ($text + " ($($role.displayName))")
                }
                $values += $text
            }
        }
        elseif($prop.Name -eq "scp")
        {
            # Delegated tokens pack the scopes into one space-separated string.
            $values = @(([string]$raw).Split(" ") | Where-Object { $_ })
        }
        else
        {
            # Covers "roles" (an array on app-only tokens) and any other array
            # claim: one row each, so nothing ever renders as "System.Object[]".
            $values = @($raw)
        }

        foreach($v in $values)
        {
            $tokenArr += New-Object PSObject -Property @{
                Name=$prop.Name
                Value=$v
            }
        }
    }
    $dg = [System.Windows.Controls.DataGrid]::new()
    Set-GridPixelScrolling $dg
    $dg.ItemsSource = ($tokenArr | Select-Object Name, Value)
    $script:UIProvider.ShowModalForm($Title, $dg)
}

# Get-MSALUserInfo moved to Internal/MSGraphAuthentication.ps1.

#region Event functions

function Invoke-MSALUIEventNewAuthentication
{
    [CmdLetbinding()]
    param($TokenInfo)

    # Provider-neutral tenant/user sync and the ScopeTags/AssignmentFilters
    # dependency-cache preload used to happen here. Both moved to
    # Invoke-AuthCoreOnNewToken in Internal/AuthenticationCore.ps1, which
    # subscribes to the same "AuthenticatedNewToken" event and runs regardless
    # of UI mode -- headless flows (OAuth AppId+Secret in Azure Automation)
    # get the same state updates that were previously UI-exclusive.
    #
    # This handler now does only UI-specific work: MSAL profile enrichment
    # (Graph /me, profile photo, JWT-derived app name) and the auth-info bar
    # refresh. Update-MSALUserProfile self-guards on $script:MSALDefaultToken
    # so it's a no-op for non-MSAL providers.
    if($TokenInfo.IsDefault) {
        Update-MSALUserProfile
        Show-AuthenticationInfo
    }
}

function Invoke-MSALUIEventTokenRefreshed
{
    [CmdLetbinding()]
    param($TokenInfo)

    # A silent renewal keeps the same user/tenant, so skip the heavy profile
    # enrichment (Graph /me + photo) and full object reload that a NEW token
    # triggers - just redraw the auth-info bar so the shown token expiry reflects
    # the renewed token. Only the default token drives the visible profile.
    if($TokenInfo.IsDefault) {
        try { Show-AuthenticationInfo }
        catch { Write-LogError "Show-AuthenticationInfo failed on token refresh" $_.Exception }
    }
}

function Invoke-MSALUIEventUserDisconnected
{
    [CmdLetbinding()]
    param($TokenInfo)

    # Wipe every cache entry tagged for the disconnected tenant so a future sign-in to
    # a different tenant doesn't see stale ScopeTags / Filters / AAD / baseline templates.
    if($TokenInfo -and $TokenInfo.TenantID) {
        try { Clear-TenantCache -TenantId $TokenInfo.TenantID }
        catch { Write-LogError "Failed to clear tenant cache on disconnect" $_.Exception }
    }

    Get-MSALUserInfo
    Show-AuthenticationInfo
}

function Invoke-MSALUIEventAuthenticationFailed
{
    # Profile enrichment is best-effort and must NEVER stop the auth-info bar from
    # redrawing - the redraw is the whole point of this handler: on a hard failure
    # (expired token / no session) Get-MSALUserProfile's expired/no-token branch
    # flips the avatar back to the Sign-in icon. Guard so a throw in Get-MSALUserInfo
    # can't skip Show-AuthenticationInfo.
    try { Get-MSALUserInfo } catch { Write-LogError "Get-MSALUserInfo failed on AuthenticationFailed" $_.Exception }
    Show-AuthenticationInfo
    # Also refresh the environment/tenant badge so it hides when the session is gone
    # (expired token still returns GetUserInfo, so Set-EnvironmentInfo's own expiry
    # guard does the hiding). Explicit here so it doesn't depend on the enrichment
    # above reaching its own Set-EnvironmentInfo call.
    try { Set-EnvironmentInfo } catch { Write-LogError "Set-EnvironmentInfo failed on AuthenticationFailed" $_.Exception }
}

function Invoke-MSALUIAppInitialized
{
    Add-AppEventHandler "AuthenticatedNewToken" "Invoke-MSALUIEventNewAuthentication"
    Add-AppEventHandler "AuthenticationTokenRefresh" "Invoke-MSALUIEventTokenRefreshed"
    Add-AppEventHandler "AuthenticationUserDisconnected" "Invoke-MSALUIEventUserDisconnected"
    Add-AppEventHandler "AuthenticationFailed" "Invoke-MSALUIEventAuthenticationFailed"
}

#Endregion
