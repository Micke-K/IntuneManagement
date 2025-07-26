function Test-AuthToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [SecureString]$SecureToken
    )

    try {
        # Convert SecureString to plain text for validation
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
        $TokenPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        
        # Basic validation - token should be a JWT
        if (-not $TokenPlainText.StartsWith('eyJ')) {
            Write-Error "The provided token does not appear to be a valid JWT token."
            return $false
        }

        # Split the token into its components
        $TokenParts = $TokenPlainText.Split('.')
        if ($TokenParts.Count -ne 3) {
            Write-Error "The provided token is not a valid JWT token (should have 3 parts)."
            return $false
        }

        # Decode the payload
        $Payload = $TokenParts[1].Replace('-', '+').Replace('_', '/')
        while ($Payload.Length % 4) { $Payload += "=" }
        $DecodedPayload = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($Payload))
        $PayloadJson = $DecodedPayload | ConvertFrom-Json

        # Check for token expiration
        $Now = [DateTime]::UtcNow
        $ExpiryTime = [DateTime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc).AddSeconds($PayloadJson.exp)
        
        if ($Now -ge $ExpiryTime) {
            Write-Error "The token has expired on $ExpiryTime UTC. Current time is $Now UTC."
            return $false
        }

        # Check for required scopes for Intune management
        $Scopes = $PayloadJson.scp -split " "
        $RequiredScopes = @(
            "DeviceManagementManagedDevices.Read.All",
            "DeviceManagementConfiguration.Read.All"
        )

        foreach ($Scope in $RequiredScopes) {
            if ($Scopes -notcontains $Scope) {
                Write-Warning "The token may not have all required permissions. Missing: $Scope"
            }
        }

        # Token appears valid
        Write-Verbose "Token validation successful. Token expires at $ExpiryTime UTC."
        return $true
    }
    catch {
        Write-Error "Error validating token: $_"
        return $false
    }
    finally {
        if ($BSTR) {
            # Clean up the unmanaged memory
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        }
    }
} 