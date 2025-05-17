function Initialize-IntuneConnection {
    <#
    .SYNOPSIS
        Sets up a direct connection to Intune using browser-acquired tokens.
        
    .DESCRIPTION
        This function implements the token-based authentication approach for connecting to Microsoft Intune.
        It handles installing required modules, validating tokens, and establishing a connection to Microsoft Graph.
        
    .PARAMETER Token
        Optional. The authentication token string obtained from a browser session. If not provided, the user will be prompted.
        
    .EXAMPLE
        Initialize-IntuneConnection
        
    .EXAMPLE
        Initialize-IntuneConnection -Token "eyJ0eXAiOiJKV..."
        
    .NOTES
        This function directly implements the token authentication approach without requiring external scripts.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$Token
    )
    
    try {
        # Check if Microsoft Graph modules are available
        $requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.DeviceManagement")
        $modulesAvailable = $true
        
        foreach ($module in $requiredModules) {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                $modulesAvailable = $false
                Write-Warning "Required module $module is not installed."
            }
        }
        
        if (-not $modulesAvailable) {
            Write-Host "Installing required Microsoft Graph modules..." -ForegroundColor Yellow
            
            try {
                foreach ($module in $requiredModules) {
                    Write-Host "Installing $module module..." -ForegroundColor Cyan
                    Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
                    Import-Module -Name $module -Force
                }
                Write-Host "All required modules installed and imported successfully." -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to install required modules: $_"
                Write-Host "Please install the modules manually using:" -ForegroundColor Yellow
                foreach ($module in $requiredModules) {
                    Write-Host "Install-Module -Name $module -Scope CurrentUser -Force" -ForegroundColor White
                }
                return $false
            }
        }
        
        # If token is not provided, prompt for it
        if ([string]::IsNullOrEmpty($Token)) {
            # Display instructions for obtaining a token
            Write-Host "`n=== ACQUIRING AUTHENTICATION TOKEN ===" -ForegroundColor Cyan
            Write-Host "To get your authentication token:" -ForegroundColor Yellow
            Write-Host "1. Open a browser and go to: https://endpoint.microsoft.com/" -ForegroundColor White
            Write-Host "2. Sign in with your admin account (if not already signed in)" -ForegroundColor White
            Write-Host "3. Press F12 to open developer tools" -ForegroundColor White
            Write-Host "4. Go to 'Network' tab" -ForegroundColor White
            Write-Host "5. Refresh the page (F5)" -ForegroundColor White
            Write-Host "6. Filter requests by typing 'graph.microsoft' in the filter box" -ForegroundColor White
            Write-Host "7. Click on any request to graph.microsoft.com" -ForegroundColor White
            Write-Host "8. In the Headers tab, scroll to find 'Authorization: Bearer eyJ...'" -ForegroundColor White
            Write-Host "9. Copy the entire token (starts with 'eyJ' and is very long)" -ForegroundColor White
            Write-Host ""
            
            # Prompt for token
            Write-Host "Paste your authentication token below and press Enter:" -ForegroundColor Green
            $Token = Read-Host
        }
        
        # Validate token format
        if (-not $Token.StartsWith("eyJ")) {
            Write-Error "The token doesn't appear to be valid. It should start with 'eyJ'."
            return $false
        }
        
        # Test the token with a basic API call
        $headers = @{
            'Authorization' = "Bearer $Token"
            'Content-Type' = 'application/json'
        }
        
        try {
            $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$top=1"
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
            Write-Host "✓ Token validated successfully with Intune API." -ForegroundColor Green
        }
        catch {
            Write-Error "Token validation failed: $($_.Exception.Message)"
            return $false
        }
        
        # Store token for later use
        $global:IntuneToken = $Token
        $global:IntuneHeaders = $headers
        
        # Connect to Microsoft Graph PowerShell
        try {
            # Convert the token to SecureString
            $secureToken = ConvertTo-SecureString -String $Token -AsPlainText -Force
            
            # Disconnect from any existing Graph connection
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            
            # Connect using the token
            Connect-MgGraph -AccessToken $secureToken -NoWelcome
            
            # Test the connection
            $context = Get-MgContext
            if ($context) {
                Write-Host "✓ Successfully connected to Microsoft Graph as $($context.Account)" -ForegroundColor Green
                
                # Test a device management command
                try {
                    $device = Get-MgDeviceManagementManagedDevice -Top 1 -ErrorAction Stop
                    Write-Host "✓ Successfully verified access to Intune managed devices!" -ForegroundColor Green
                    
                    # Store a script-scope copy of the token for our module functions
                    $script:IntuneToken = $Token
                    $script:IntuneHeaders = $headers
                    
                    return $true
                }
                catch {
                    Write-Host "× Could not access Intune devices: $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-Host "  This may indicate insufficient permissions, but token authentication was successful." -ForegroundColor Yellow
                    
                    # Still store the tokens since basic authentication worked
                    $script:IntuneToken = $Token
                    $script:IntuneHeaders = $headers
                    
                    return $true
                }
            }
            else {
                Write-Error "Could not establish Microsoft Graph PowerShell session"
                return $false
            }
        }
        catch {
            Write-Error "Error connecting to Microsoft Graph PowerShell: $_"
            
            # Provide guidance on module updates if needed
            Write-Host "If this error persists, try updating the Microsoft Graph modules:" -ForegroundColor Yellow
            Write-Host "Update-Module Microsoft.Graph -Force" -ForegroundColor White
            
            return $false
        }
    }
    catch {
        Write-Error "Error in Initialize-IntuneConnection: $_"
        return $false
    }
} 