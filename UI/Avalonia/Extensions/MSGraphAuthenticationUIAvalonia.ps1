# UI/Avalonia counterpart of UI/WPF/Extensions/MSGraphAuthenticationUI.ps1.
#
# Scope: event handlers + dialog ports that depend on Avalonia controls
# (Show-MSALDecodedToken, Show-CloudPickerMenu, the four AuthenticationXxx event
# handlers, Invoke-MSALUIAppInitialized). The non-UI auth helpers
# (Invoke-AuthProviderInteractiveLogin, Get-MSALUserInfo) live in
# Internal/MSGraphAuthentication.ps1 and are shared by both UI trees.
#
# Deliberately NOT ported here (yet):
#  - Get-MSALUserProfile — replaced by an inline Avalonia helper inside
#    Show-AuthenticationInfo for the simplified initials button. The full
#    cached-account picker / popup version lives in WPF only for now.
#  - Get-MSALUserPhotoEllips — deferred until the popup infrastructure
#    (Show-Popup) is ported.

Add-AppEventHandler "AppInitialized" "Invoke-MSALUIAppInitialized"

function Show-MSALDecodedToken
{
    <#
    .SYNOPSIS
        Render a decoded JWT (header + payload) in a modal DataGrid.
    .DESCRIPTION
        Port of UI/WPF/Extensions/MSGraphAuthenticationUI.ps1 Show-MSALDecodedToken.
        Same formatting rules: exp/iat/nbf/xms_tcdt -> local DateTime,
        acrs/amr/scp -> newline-joined, wids -> resolved Entra role displayName
        when available. Output is a list of [TokenInfoRow] CLR instances bound
        into a DataGrid with auto-generated columns.
    #>
    param($TokenData, $Title)

    $ui = $script:UIProvider
    if (-not $TokenData -or -not $TokenData.Header) { return }

    $tokenArr = @()
    foreach ($prop in ($TokenData.Header | Get-Member | Where-Object MemberType -eq NoteProperty)) {
        $tokenArr += [TokenInfoRow]@{
            Name  = $prop.Name
            Value = $TokenData.Header."$($prop.Name)"
        }
    }

    # Each claim yields one or more GRID ROWS rather than one row holding a
    # multi-line string. Newline-joining a long list put a single row several
    # times taller than the rest (8 permissions measured 147px against 33px),
    # and the DataGrid estimates its scroll extent from the row heights it has
    # realized - so that one outlier made the extent swing as it scrolled in and
    # out of view, giving a jumpy, flickering scrollbar and a mouse wheel that
    # appeared to run through the table several times. One value per row keeps
    # every row a uniform height and stays fully readable.
    foreach ($prop in ($TokenData.Payload | Get-Member | Where-Object MemberType -eq NoteProperty)) {
        $raw = $TokenData.Payload."$($prop.Name)"

        if ($prop.Name -in @('exp','iat','nbf','xms_tcdt')) {
            $values = @([datetime]::new(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc).AddSeconds($raw).ToLocalTime())
        }
        elseif ($prop.Name -in @('acrs','amr')) {
            # Short arrays; a single ';'-joined line stays well inside one row.
            $values = @($raw -join ';')
        }
        elseif ($prop.Name -in @('wids')) {
            if (-not $script:AADRoles) {
                # Will fail if RoleManagement.Read.Directory permission isn't
                # granted. -NoError swallows the failure; we just render the
                # raw GUIDs without a friendly name.
                $script:AADRoles = (Invoke-MSGraphAPI -url "/directoryRoles?`$select=roleTemplateId,displayName" -ODataMetadata 'minimal' -Noerror).value
            }
            $values = @()
            foreach ($wid in $raw) {
                $text = $wid
                $role = ($script:AADRoles | Where-Object roleTemplateId -eq $wid)
                if ($role) { $text = "$text ($($role.displayName))" }
                $values += $text
            }
        }
        elseif ($prop.Name -eq 'scp') {
            # Delegated tokens pack the scopes into one space-separated string.
            $values = @(([string]$raw).Split(' ') | Where-Object { $_ })
        }
        else {
            # Covers 'roles' (an array on app-only tokens) and any other array
            # claim: one row each, so nothing ever renders as "System.Object[]".
            $values = @($raw)
        }

        foreach ($v in $values) {
            $tokenArr += [TokenInfoRow]@{
                Name  = $prop.Name
                Value = $v
            }
        }
    }

    $dg = [Avalonia.Controls.DataGrid]::new()
    $dg.AutoGenerateColumns      = $true
    $dg.IsReadOnly               = $true
    $dg.CanUserSortColumns       = $true
    $dg.CanUserResizeColumns     = $true
    $dg.GridLinesVisibility      = [Avalonia.Controls.DataGridGridLinesVisibility]::Horizontal
    $dg.MinWidth                 = 600
    $dg.MinHeight                = 400
    $dg.ItemsSource              = $tokenArr

    $ui.ShowModalForm($Title, $dg)
}

function Show-CloudPickerMenu
{
    <#
    .SYNOPSIS
        Modal dialog to choose a target cloud before sign-in.
    .DESCRIPTION
        Avalonia port of UI/WPF/Extensions/MSGraphAuthenticationUI.ps1
        Show-CloudPickerMenu. Returns the selected Cloud value (string) on OK,
        $null on Cancel. -Persist saves the choice to DefaultCloud for future
        zero-config sign-ins.
    .NOTES
        $script:Clouds items are PSCustomObjects (Name / Value); converted to
        [SettingsListItem] CLR instances so Avalonia ComboBox bindings resolve.
    #>
    param(
        [string]$Default,
        [switch]$Persist
    )

    $ui = $script:UIProvider
    $dialog = $ui.GetXamlObject((Join-Path $script:AppUIRootFolder 'XAML/CloudPickerMenu.axaml'))
    if (-not $dialog) { return $null }

    $cbCloud = (Get-AvaloniaHost)::FindByName($dialog, 'cbCloud')
    if (-not $cbCloud) { return $null }

    $items = @()
    foreach ($cloud in @($script:Clouds)) {
        if (-not $cloud) { continue }
        $items += [SettingsListItem]@{
            Name  = [string]$cloud.Name
            Value = $cloud.Value
        }
    }
    $cbCloud.ItemsSource = $items


    if (-not $Default) { $Default = Get-DefaultCloud }
    $selected = $items | Where-Object { "$($_.Value)" -eq "$Default" } | Select-Object -First 1
    if (-not $selected -and $items.Count -gt 0) { $selected = $items[0] }
    if ($selected) { $cbCloud.SelectedItem = $selected }

    # Module scope, not captures: the handlers are re-bound and lose locals, so
    # Sign in / Cancel never closed the dialog and it always returned $null -
    # "Sign in to a different cloud..." could never actually sign in. Reset on
    # entry so a previous open cannot leak its answer into this one.
    $script:_cloudPickerDialog = $dialog
    $script:_cloudPickerCombo  = $cbCloud
    $script:_cloudPickerResult = $null

    $ui.AddXamlEvent($dialog, 'btnLogin', 'Add_Click', ({
        if ($script:_cloudPickerCombo -and $script:_cloudPickerCombo.SelectedItem) {
            $script:_cloudPickerResult = [string]$script:_cloudPickerCombo.SelectedItem.Value
        }
        if ($script:_cloudPickerDialog) { $script:_cloudPickerDialog.Close() }
    }))

    $ui.AddXamlEvent($dialog, 'btnCancel', 'Add_Click', ({
        $script:_cloudPickerResult = $null
        if ($script:_cloudPickerDialog) { $script:_cloudPickerDialog.Close() }
    }))

    (Get-AvaloniaHost)::ShowDialog($dialog, $script:Window)

    $picked = $script:_cloudPickerResult
    $script:_cloudPickerDialog = $null
    $script:_cloudPickerCombo  = $null
    $script:_cloudPickerResult = $null

    if (-not $picked) { return $null }
    if ($Persist) { Set-DefaultCloud $picked }
    return [string]$picked
}

# Filters out Microsoft personal / consumer (MSA) accounts. Those accounts live
# on the well-known tenant 9188040d-6c67-4c5b-b112-36a304b66dad and cannot be
# used to sign into Entra-tenant resources (Intune / Graph) — surfacing them in
# the picker only confuses the user. Keep work/school accounts plus accounts
# where tenant info isn't available. (Verbatim from WPF MSGraphAuthenticationUI.)
function Test-IsPersonalMSAAccount {
    param($Account)
    if(-not $Account -or -not $Account.HomeAccountId) { return $false }
    $tid = $Account.HomeAccountId.TenantId
    return ($tid -eq "9188040d-6c67-4c5b-b112-36a304b66dad")
}

function Invoke-AvaloniaCachedAccountSignIn {
    param($Account)
    if(-not $Account) { return }

    Write-Status "Logging in with $($Account.UserName)"
    try {
        if (Connect-EntraEnvironment -User $Account.UserName -DefaultToken) {
            # AuthenticatedNewToken event handler refreshes the UI.
        }
    }
    catch {
        Write-LogError "Cached-account sign-in failed" $_.Exception
    }
    finally {
        Write-Status ""
    }
}

# Build one row in the cached-account picker — Avalonia version of WPF
# Add-CachedUser. Each row is a 2-column Grid: a click-to-login button on the
# left that calls Connect-EntraEnvironment with the cached username, and a
# Forget button on the right that disconnects (if signed in under this account)
# and evicts the MSAL cache entry.
function Add-CachedUser {
    param($Account, $ParentObj)

    $ui = $script:UIProvider
    try {
        $row = [Avalonia.Controls.Grid]::new()
        $row.Margin = [Avalonia.Thickness]::new(0, 5, 0, 0)
        $row.ColumnDefinitions.Add([Avalonia.Controls.ColumnDefinition]::new([Avalonia.Controls.GridLength]::new(1, [Avalonia.Controls.GridUnitType]::Star)))
        $row.ColumnDefinitions.Add([Avalonia.Controls.ColumnDefinition]::new([Avalonia.Controls.GridLength]::Auto))

        $tenantName = Get-SettingStoreValue $Account.HomeAccountId.TenantId "_Name" $Account.HomeAccountId.TenantId

        $loginBtn = [Avalonia.Controls.Button]::new()
        $loginBtn.HorizontalAlignment        = [Avalonia.Layout.HorizontalAlignment]::Stretch
        $loginBtn.HorizontalContentAlignment = [Avalonia.Layout.HorizontalAlignment]::Left
        $loginBtn.Cursor = [Avalonia.Input.Cursor]::new([Avalonia.Input.StandardCursorType]::Hand)
        $loginBtn.Tag = $Account

        # TextBlock.Inlines + Run/LineBreak to render the two-line username/tenant
        # label without needing a separate StackPanel host.
        $loginText = [Avalonia.Controls.TextBlock]::new()
        $loginText.Inlines.Add([Avalonia.Controls.Documents.Run]::new($Account.UserName))
        $loginText.Inlines.Add([Avalonia.Controls.Documents.LineBreak]::new())
        $loginText.Inlines.Add([Avalonia.Controls.Documents.Run]::new([string]$tenantName))
        $loginBtn.Content = $loginText

        $loginBtn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            $acct = $S.Tag
            $ui.HidePopup()
            Invoke-AvaloniaDeferredAction -Argument $acct -Action {
                param($selectedAccount)
                Invoke-AvaloniaCachedAccountSignIn $selectedAccount
            }
        }))
        $row.Children.Add($loginBtn) | Out-Null

        $forgetBtn = [Avalonia.Controls.Button]::new()
        # U+1F5D1 (🗑) is outside the BMP and won't fit in a single 16-bit
        # [char]. ConvertFromUtf32 returns the proper surrogate-pair string.
        $forgetBtn.Content = [char]::ConvertFromUtf32(0x1F5D1)
        $forgetBtn.Margin  = [Avalonia.Thickness]::new(5, 0, 0, 0)
        [Avalonia.Controls.ToolTip]::SetTip($forgetBtn, 'Forget this account')
        $forgetBtn.Cursor = [Avalonia.Input.Cursor]::new([Avalonia.Input.StandardCursorType]::Hand)
        $forgetBtn.Tag = $Account
        [Avalonia.Controls.Grid]::SetColumn($forgetBtn, 1)
        $forgetBtn.add_Click((ConvertTo-AvaloniaEventScriptBlock {
            param($S, $E)
            $acct = $S.Tag
            Write-Status "Removing $($acct.UserName)"
            try {
                # If the account is currently signed in under any token, disconnect
                # properly first so the event fires and any default-token promotion
                # runs. Always also evict from the on-disk MSAL cache.
                $matchingTokens = @($script:MSALTokens.Values | Where-Object {
                    $_.Token -and $_.Token.Account -and
                    $_.Token.Account.HomeAccountId.Identifier -eq $acct.HomeAccountId.Identifier
                })
                foreach ($tok in $matchingTokens) { Disconnect-EntraEnvironment -TokenID $tok.Id }
                Remove-MSALAccount -Account $acct

                # Walk up to the StackPanel (row.Parent) and drop the row.
                $parent = $S.Parent  # the row Grid
                if ($parent -and $parent.Parent) {
                    [void]$parent.Parent.Children.Remove($parent)
                }
            } catch {
                Write-LogError "Failed to forget account $($acct.UserName)" $_.Exception
            }
            Write-Status ""
        }))
        $row.Children.Add($forgetBtn) | Out-Null

        $ParentObj.Children.Add($row) | Out-Null
    } catch {
        Write-LogError "Add-CachedUser failed" $_.Exception
    }
}

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
    $ui = $script:UIProvider
    if($TokenInfo.IsDefault) {
        Update-MSALUserProfile
        $ui.ShowAuthenticationInfo()
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
        $ui = $script:UIProvider
        if($ui) {
            try { $ui.ShowAuthenticationInfo() }
            catch { Write-LogError "ShowAuthenticationInfo failed on token refresh" $_.Exception }
        }
    }
}

function Invoke-MSALUIEventUserDisconnected
{
    [CmdLetbinding()]
    param($TokenInfo)

    $ui = $script:UIProvider
    if($TokenInfo -and $TokenInfo.TenantID) {
        try { Clear-TenantCache -TenantId $TokenInfo.TenantID }
        catch { Write-LogError "Failed to clear tenant cache on disconnect" $_.Exception }
    }

    Get-MSALUserInfo
    $ui.ShowAuthenticationInfo()
}

function Invoke-MSALUIEventAuthenticationFailed
{
    $ui = $script:UIProvider
    # Best-effort enrichment; must never stop the redraw. On a hard failure the
    # redraw reverts the avatar to the Sign-in button (Get-AvaloniaUserProfile's
    # expired/no-token branch). Guard so a throw can't skip ShowAuthenticationInfo.
    try { Get-MSALUserInfo } catch { Write-LogError "Get-MSALUserInfo failed on AuthenticationFailed" $_.Exception }
    $ui.ShowAuthenticationInfo()
    # Also refresh the environment/tenant badge so it hides when the session is gone
    # (Set-EnvironmentInfo's expiry guard does the hiding). Explicit here so it doesn't
    # depend on the enrichment above reaching its own SetEnvironmentInfo call.
    try { $ui.SetEnvironmentInfo() } catch { Write-LogError "SetEnvironmentInfo failed on AuthenticationFailed" $_.Exception }
}

function Invoke-MSALUIAppInitialized
{
    Add-AppEventHandler "AuthenticatedNewToken" "Invoke-MSALUIEventNewAuthentication"
    Add-AppEventHandler "AuthenticationTokenRefresh" "Invoke-MSALUIEventTokenRefreshed"
    Add-AppEventHandler "AuthenticationUserDisconnected" "Invoke-MSALUIEventUserDisconnected"
    Add-AppEventHandler "AuthenticationFailed" "Invoke-MSALUIEventAuthenticationFailed"
}

#endregion
