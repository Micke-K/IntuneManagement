#ImportOrder 27

# Offline "mock tenant" provider.
#
# Signs in with no network and answers every Graph request from the JSON files
# under IM_MOCK_DATA (see Internal/MockGraph.ps1). Registered only when that
# variable points at a folder - a normal launch never sees it - and used for
# screenshots, demos and UI work without a tenant.
#
# The access token is a real-looking but unsigned JWT: the access-marking code
# reads scp / wids from it exactly as it would from Entra, so the menu renders
# with full access and no role lookup. Every request then goes through
# InvokeWebRequest (RoutesAllRequests) instead of HTTPS.

# What a mock 4xx throws. Invoke-MSGraphAPI reads the failure from
# $_.Exception.Response - StatusCode, Headers, and the body through
# GetResponseStream() - the same way it reads a WebException from
# Invoke-WebRequest, so a mock 404 reports its status, code and message like a
# real one. WebException.Response cannot be set from PowerShell, hence a class.
class MockGraphResponseException : System.Exception {
    [object]$Response

    MockGraphResponseException([string]$Message, [object]$Response) : base($Message) {
        $this.Response = $Response
    }
}

class AuthenticationMock : AuthenticationProvider {

    static [hashtable]$Tokens = @{}

    [string]$DataFolder

    AuthenticationMock() {
        $this.Id          = "Mock"
        $this.DisplayName = "Mock tenant (offline demo data)"

        $this.SupportsInteractive       = $true
        $this.SupportsClientSecret      = $false
        $this.SupportsCertificate       = $false
        $this.SupportsIdentityProvider  = $false
        $this.SupportsBYOToken          = $false
        $this.SupportsClaimsChallenge   = $false
        $this.SupportsMultiTenant       = $false
        $this.SupportsRefresh           = $true
        $this.SupportsForget            = $false
        $this.SupportsCachedUsers       = $false
        $this.RoutesAllRequests         = $true
    }

    [void] Initialize() {
        $this.DataFolder = [string]$env:IM_MOCK_DATA
    }

    # Sign in at startup - there is nothing to prompt for.
    [bool] TryResumeSession() {
        if(-not $this.DataFolder) { return $false }
        $token = $this.Connect(@{})
        return [bool]$token
    }

    [PSCustomObject] Connect([hashtable]$Arguments) {
        if(-not $this.DataFolder -or -not (Test-Path -LiteralPath $this.DataFolder -PathType Container)) {
            Write-Log "Mock provider: IM_MOCK_DATA does not point at a folder ('$($this.DataFolder)')" 3
            return $null
        }

        $tenant = Get-MockTenantProfile -Root $this.DataFolder
        Initialize-MockGraphStore -Root $this.DataFolder | Out-Null

        $tokenId = Get-NextAuthTokenId
        [AuthenticationMock]::Tokens[$tokenId] = @{
            Id          = $tokenId
            Tenant      = $tenant
            AccessToken = (New-MockAccessToken -Tenant $tenant)
            AcquiredAt  = [DateTime]::UtcNow
        }
        Write-Log "Mock provider: signed in to '$($tenant.TenantName)' as $($tenant.UPN) (TokenId=$tokenId)"

        try {
            $cur = Get-AuthProvider
            if($cur -and $cur.Id -ne $this.Id) { Set-ActiveAuthProvider -Id $this.Id }
        } catch { }

        return (Register-AuthToken -Provider $this -TokenId $tokenId -Cloud (Get-DefaultCloud))
    }

    [bool] Disconnect([int]$TokenId) {
        if(-not [AuthenticationMock]::Tokens.ContainsKey($TokenId)) { return $false }
        Unregister-AuthToken -TokenId $TokenId
        [AuthenticationMock]::Tokens.Remove($TokenId) | Out-Null
        return $true
    }

    [bool] Refresh([int]$TokenId) {
        return [AuthenticationMock]::Tokens.ContainsKey($TokenId)
    }

    [string] GetAccessToken([int]$TokenId, [string]$Resource) {
        $entry = $this.GetEntry($TokenId)
        if(-not $entry) { return $null }
        return [string]$entry.AccessToken
    }

    [datetime] GetAccessTokenExpiry([int]$TokenId, [string]$Resource) {
        return [datetime]::MaxValue
    }

    [object] InvokeWebRequest([string]$Url, [string]$Method, [object]$Body, [hashtable]$Headers) {
        return (Invoke-MockGraphRequest -Url $Url -Method $Method -Body $Body)
    }

    [PSCustomObject] GetUserInfo([int]$TokenId) {
        $entry = $this.GetEntry($TokenId)
        if(-not $entry) { return $null }
        $tenant = $entry.Tenant
        return [PSCustomObject]@{
            Provider    = $this.Id
            DisplayName = $tenant.DisplayName
            UPN         = $tenant.UPN
            UserId      = $tenant.UserId
            TenantId    = $tenant.TenantId
            TenantName  = $tenant.TenantName
            AppId       = $tenant.AppId
            AppName     = $tenant.AppName
            AuthType    = "Interactive"
            ExpiresOn   = [DateTime]::Now.AddYears(1)
        }
    }

    [PSCustomObject[]] GetSessionInfoRows() {
        $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
        $rows.Add([PSCustomObject]@{ Name = "Provider";    Value = "Mock (offline)" })
        $rows.Add([PSCustomObject]@{ Name = "Data folder"; Value = $this.DataFolder })
        return $rows.ToArray()
    }

    hidden [hashtable] GetEntry([int]$TokenId) {
        if($TokenId -le 0 -and [AuthenticationMock]::Tokens.Count -gt 0) {
            $TokenId = ([AuthenticationMock]::Tokens.Keys | Sort-Object -Descending | Select-Object -First 1)
        }
        if(-not [AuthenticationMock]::Tokens.ContainsKey($TokenId)) { return $null }
        return [AuthenticationMock]::Tokens[$TokenId]
    }
}
