# Avalonia port of UI/WPF/Extensions/EffectivePermissionsUIWPF.ps1: the
# "Permissions" button in the Profile popup. Thin caller of the public
# Get-GraphEffectivePermissions (R10); rows go through the CLR class
# Classes/EffectivePermissionRowItem.ps1 because the Avalonia binder cannot
# read PSCustomObject note properties. Wired from Show-ProfilePopup
# (CoreUIAvalonia.ps1) with a single AddXamlEvent line.

function Show-EffectivePermissionsDialog {
    param()

    $ui = $script:UIProvider

    $rows = @()
    $raw  = $null
    try {
        $rows = @(Get-GraphEffectivePermissions)
        $raw  = Get-GraphEffectivePermissions -Raw
    }
    catch { Write-LogError "Get-GraphEffectivePermissions failed" $_.Exception }

    if ($rows.Count -eq 0) {
        $ui.ShowMessageBox('No access token available. Sign in first.', 'Permissions', 'OK', 'Information') | Out-Null
        return
    }

    $items = @()
    foreach ($r in ($rows | Sort-Object EffectiveLevel, Title)) {
        $items += [EffectivePermissionRowItem]@{
            Type       = [string]$r.Title
            Category   = [string]$r.ResourceCategory
            Required   = [string]$r.Required
            Token      = [string]$r.TokenAccess
            IntuneRole = [string]$r.RoleAccess
            Effective  = [string]$r.EffectiveAccess
            Result     = [string]$r.Result
            Reason     = [string]$r.Reason
        }
    }

    $grid = [Avalonia.Controls.Grid]::new()
    $rdHeader = [Avalonia.Controls.RowDefinition]::new(); $rdHeader.Height = [Avalonia.Controls.GridLength]::Auto
    $rdGrid   = [Avalonia.Controls.RowDefinition]::new(); $rdGrid.Height   = [Avalonia.Controls.GridLength]::new(1, [Avalonia.Controls.GridUnitType]::Star)
    $grid.RowDefinitions.Add($rdHeader)
    $grid.RowDefinitions.Add($rdGrid)

    $txt = [Avalonia.Controls.TextBlock]::new()
    $txt.Text         = Get-EffectivePermissionsHeaderText $raw
    $txt.TextWrapping = [Avalonia.Media.TextWrapping]::Wrap
    $txt.MaxWidth     = 900
    $txt.Margin       = [Avalonia.Thickness]::new(5, 5, 5, 10)
    [Avalonia.Controls.Grid]::SetRow($txt, 0)
    $grid.Children.Add($txt)

    $dg = [Avalonia.Controls.DataGrid]::new()
    $dg.AutoGenerateColumns  = $true
    $dg.IsReadOnly           = $true
    $dg.CanUserSortColumns   = $true
    $dg.CanUserResizeColumns = $true
    $dg.GridLinesVisibility  = [Avalonia.Controls.DataGridGridLinesVisibility]::Horizontal
    $dg.MinWidth             = 900
    $dg.MinHeight            = 450
    $dg.ItemsSource          = $items
    [Avalonia.Controls.Grid]::SetRow($dg, 1)
    $grid.Children.Add($dg)

    $ui.ShowModalForm('Permissions', $grid)
}

# Same wording as the WPF dialog; each backend keeps its own copy (R12).
function Get-EffectivePermissionsHeaderText {
    param($Context)

    $refresh = "Changed roles (PIM, a new Intune role)? Click Refresh in the Profile popup - the app keeps using the token it signed in with until then."
    if (-not $Context) {
        return ("Token scopes only. The Intune role check does not apply to this token (app-only sign-in) or " +
                "could not be read; see the log. Scope tags are not evaluated. $refresh")
    }
    $asOf = if ($Context.AsOf) { " Token issued $($Context.AsOf.ToString('yyyy-MM-dd HH:mm'))." } else { "" }
    if ($Context.AllAllowed) {
        return ("Token scopes combined with the Intune Administrator / Global Administrator directory role " +
                "(full Intune access, no per-action lookup).$asOf $refresh")
    }
    # "N of M" only when the action catalogue was readable (see the WPF copy).
    $counts = "$($Context.Allowed.Count) resource actions allowed"
    if ($Context.Catalog) { $counts = "$($Context.Allowed.Count) of $($Context.Catalog.Count) resource actions allowed" }
    return ("Token scopes combined with the signed-in user's Intune role permissions " +
            "($counts).$asOf " +
            "Scope tags are not evaluated: a user limited to some tags can still be refused on individual objects. $refresh")
}
