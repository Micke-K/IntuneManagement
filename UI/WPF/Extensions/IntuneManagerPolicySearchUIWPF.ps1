# Scoped policy picker (WPF).
#
# Thin wrapper over Show-ObjectPickerDialog: it supplies the tenant list, the
# scope list and a search handler, so the dialog queries Graph for the typed
# name inside the selected policy group / policy type instead of enumerating the
# whole tenant up front. The search itself lives in Internal/PolicySearch.ps1
# and is shared with the Avalonia backend (R10 - this file only collects input
# and renders).

function Show-PolicySearchDialog
{
    param(
        [string]$Title = "Select policy",
        # Preselects this policy type in the "Search in" dropdown.
        [string]$DefaultPolicyTypeId,
        # Ids to keep out of the results (e.g. the policy being compared).
        [string[]]$ExcludeIds,
        # Rows to show before the first search. The compare flows seed this
        # from the main view's already-loaded list so the dialog opens with
        # something useful in it.
        [object[]]$InitialItems
    )

    $scopes = Get-PolicySearchScopes -DefaultPolicyTypeId $DefaultPolicyTypeId
    if(@($scopes).Count -eq 0) {
        $script:UIProvider.ShowMessageBox("No policy types are available to search.", "Select policy", "OK", "Warning") | Out-Null
        return $null
    }

    $defaultKey = Get-PolicySearchDefaultScopeKey -Scopes $scopes -PolicyTypeId $DefaultPolicyTypeId

    # Empty unless more than one tenant is signed in; the dialog hides the combo
    # in that case and every search runs against the default token.
    $tenants          = Get-PolicySearchTenants
    $defaultTenantKey = Get-PolicySearchDefaultTenantKey -Tenants $tenants

    $cols = @(
        [PSCustomObject]@{ Header = "Name"; Binding = "Name" }
        [PSCustomObject]@{ Header = "Type"; Binding = "PolicyType.Title" }
    )

    # Both handlers read $script: state rather than capturing locals - see the
    # $script:_pickerState note in Show-ObjectPickerDialog.
    $script:_policySearchState = @{
        ExcludeIds = $ExcludeIds
        MinLength  = Get-PolicySearchMinLength
    }

    $searchHandler = {
        param($SearchText, $Scope)

        $st = $script:_policySearchState
        if(-not $Scope) { return @() }

        if([String]::IsNullOrWhiteSpace($SearchText) -or $SearchText.Trim().Length -lt $st.MinLength) {
            Set-PickerStatusText "Type at least $($st.MinLength) characters, or use 'Load all in scope'."
            return @()
        }

        # $null when the dialog is single-tenant, in which case the default
        # token is used.
        $tenant  = Get-PickerSelectedTenant
        $tokenId = if($tenant) { [int]$tenant.TokenId } else { Get-DefaultTokenId }
        $where   = if($tenant) { " in $($tenant.Title)" } else { "" }

        Write-Status "Searching $($Scope.Title.Trim())$where..."
        try {
            $found = @(Search-IntunePolicies -Scope $Scope -SearchText $SearchText.Trim() -ExcludeIds $st.ExcludeIds -TokenId $tokenId)
            if($found.Count -eq 0) {
                Set-PickerStatusText "No object in $($Scope.Title.Trim())$where matches '$($SearchText.Trim())'."
            }
            else {
                Set-PickerStatusText "$($found.Count) object(s) found$where."
            }
            return $found
        }
        finally {
            Write-Status ""
        }
    }

    $loadHandler = {
        $st = $script:_policySearchState
        $scope = Get-PickerSelectedScope
        if(-not $scope) { return @() }

        $tenant  = Get-PickerSelectedTenant
        $tokenId = if($tenant) { [int]$tenant.TokenId } else { Get-DefaultTokenId }
        $where   = if($tenant) { " in $($tenant.Title)" } else { "" }

        Write-Status "Loading all objects in $($scope.Title.Trim())$where..."
        try {
            $loaded = @(Search-IntunePolicies -Scope $scope -ExcludeIds $st.ExcludeIds -TokenId $tokenId)
            Set-PickerStatusText "$($loaded.Count) object(s) in $($scope.Title.Trim())$where."
            return $loaded
        }
        finally {
            Write-Status ""
        }
    }

    try {
        return Show-ObjectPickerDialog `
            -Title $Title `
            -Items $InitialItems `
            -DisplayColumns $cols `
            -Scopes $scopes `
            -SelectedScopeKey $defaultKey `
            -Tenants $tenants `
            -SelectedTenantKey $defaultTenantKey `
            -SearchHandler $searchHandler `
            -LoadHandler $loadHandler `
            -LoadLabel "Load all in scope"
    }
    finally {
        $script:_policySearchState = $null
    }
}
