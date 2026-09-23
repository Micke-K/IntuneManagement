# "Permissions" button in the Profile popup: what the signed-in identity can
# actually do per policy type - the app's token scopes combined with the user's
# Intune role (Internal/EffectivePermissions.ps1). Thin caller of the public
# Get-GraphEffectivePermissions (R10); the popup wiring in
# MSGraphAuthenticationUIWPF.ps1 is a single AddXamlEvent line.

function Show-EffectivePermissionsDialog {
    param()

    $rows = @()
    $raw  = $null
    try {
        $rows = @(Get-GraphEffectivePermissions)
        $raw  = Get-GraphEffectivePermissions -Raw
    }
    catch { Write-LogError "Get-GraphEffectivePermissions failed" $_.Exception }

    if($rows.Count -eq 0) {
        $script:UIProvider.ShowMessageBox("No access token available. Sign in first.", "Permissions", "OK", "Information") | Out-Null
        return
    }

    $header = Get-EffectivePermissionsHeaderText $raw

    $grid = [System.Windows.Controls.Grid]::new()
    $grid.RowDefinitions.Add([System.Windows.Controls.RowDefinition]@{ Height = "Auto" })
    $grid.RowDefinitions.Add([System.Windows.Controls.RowDefinition]@{ Height = "*" })

    $txt = [System.Windows.Controls.TextBlock]::new()
    $txt.Text = $header
    $txt.TextWrapping = "Wrap"
    $txt.Margin = "5,5,5,10"
    [System.Windows.Controls.Grid]::SetRow($txt, 0)
    $grid.Children.Add($txt) | Out-Null

    $dg = [System.Windows.Controls.DataGrid]::new()
    Set-GridPixelScrolling $dg
    $dg.IsReadOnly = $true
    $dg.ItemsSource = @($rows | Sort-Object EffectiveLevel, Title |
        Select-Object @{ n = "Type"; e = { $_.Title } },
                      @{ n = "Category"; e = { $_.ResourceCategory } },
                      @{ n = "Required"; e = { $_.Required } },
                      @{ n = "Token"; e = { $_.TokenAccess } },
                      @{ n = "Intune role"; e = { $_.RoleAccess } },
                      @{ n = "Effective"; e = { $_.EffectiveAccess } },
                      @{ n = "Result"; e = { $_.Result } },
                      Reason)
    [System.Windows.Controls.Grid]::SetRow($dg, 1)
    $grid.Children.Add($dg) | Out-Null

    $script:UIProvider.ShowModalForm("Permissions", $grid)
}

# One-paragraph explanation of where the Intune-role half came from. Shared
# wording with the Avalonia dialog lives here in spirit only - each backend
# keeps its own copy under UI/<backend>/ (R12).
function Get-EffectivePermissionsHeaderText {
    param($Context)

    $refresh = "Changed roles (PIM, a new Intune role)? Click Refresh in the Profile popup - the app keeps using the token it signed in with until then."
    if(-not $Context) {
        return ("Token scopes only. The Intune role check does not apply to this token (app-only sign-in) or " +
                "could not be read; see the log. Scope tags are not evaluated. $refresh")
    }
    $asOf = if($Context.AsOf) { " Token issued $($Context.AsOf.ToString('yyyy-MM-dd HH:mm'))." } else { "" }
    if($Context.AllAllowed) {
        return ("Token scopes combined with the Intune Administrator / Global Administrator directory role " +
                "(full Intune access, no per-action lookup).$asOf $refresh")
    }
    # "N of M" only when the action catalogue was readable; M is how many actions
    # Intune defines, so N of N would mean full access, not an empty answer.
    $counts = "$($Context.Allowed.Count) resource actions allowed"
    if($Context.Catalog) { $counts = "$($Context.Allowed.Count) of $($Context.Catalog.Count) resource actions allowed" }
    return ("Token scopes combined with the signed-in user's Intune role permissions " +
            "($counts).$asOf " +
            "Scope tags are not evaluated: a user limited to some tags can still be refused on individual objects. $refresh")
}
