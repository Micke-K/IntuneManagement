# Enhanced Intune Connection Script
# Uses your browser token with official Microsoft Graph PowerShell modules
# For maximum compatibility and scalability

# Function to test if the token is valid and can access Intune
function Test-IntuneToken {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Token
    )
    
    # Create headers with the token
    $headers = @{
        'Authorization' = "Bearer $Token"
        'Content-Type' = 'application/json'
    }
    
    # Test a basic Intune API call
    try {
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$top=1"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
        
        # Token works!
        return $headers
    }
    catch {
        Write-Host "Token is not valid for Intune access: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Function to setup Microsoft Graph PowerShell with the token
function Connect-GraphWithToken {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Token
    )
    
    # Check if required modules are installed
    $requiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.DeviceManagement",
        "Microsoft.Graph.Users"
    )
    
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing $module module..." -ForegroundColor Yellow
            Install-Module -Name $module -Scope CurrentUser -Force
        }
        
        # Import the module
        Write-Host "Importing $module module..." -ForegroundColor Cyan
        Import-Module $module -Force
    }
    
    try {
        # Convert the string token to a SecureString
        $secureToken = ConvertTo-SecureString -String $Token -AsPlainText -Force
        
        # Disconnect if already connected
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        
        # Connect using the secure token
        Write-Host "Connecting to Microsoft Graph PowerShell with token..." -ForegroundColor Cyan
        Connect-MgGraph -AccessToken $secureToken -NoWelcome
        
        # Verify the connection was successful
        $context = Get-MgContext
        if ($context) {
            Write-Host "✓ Successfully connected to Microsoft Graph PowerShell!" -ForegroundColor Green
            Write-Host "  Account: $($context.Account)" -ForegroundColor Green
            
            # Test connection by getting a device
            try {
                $device = Get-MgDeviceManagementManagedDevice -Top 1 -ErrorAction Stop
                if ($device) {
                    Write-Host "✓ Successfully verified access to Intune managed devices!" -ForegroundColor Green
                    return $true
                }
            }
            catch {
                Write-Host "× Could not access Intune devices with Microsoft Graph cmdlets: $($_.Exception.Message)" -ForegroundColor Yellow
                return $false
            }
        }
        else {
            Write-Host "× Could not establish Microsoft Graph PowerShell session" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "× Error connecting to Microsoft Graph PowerShell: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  If this error persists, try updating the Microsoft.Graph modules:" -ForegroundColor Yellow
        Write-Host "  Update-Module Microsoft.Graph.Authentication -Force" -ForegroundColor White
        return $false
    }
}

# Function to get available Microsoft Graph Intune cmdlets
function Get-IntuneCommands {
    $intuneCommands = Get-Command -Module Microsoft.Graph.DeviceManagement | 
                     Where-Object { $_.Name -like "*-MgDeviceManagement*" } |
                     Select-Object -First 10 |
                     ForEach-Object { $_.Name }
    return $intuneCommands
}

# Main script
try {
    Clear-Host
    Write-Host "=== ENHANCED INTUNE CONNECTION SCRIPT ===" -ForegroundColor Cyan
    Write-Host "This script provides access to Intune via Microsoft Graph PowerShell modules" -ForegroundColor Cyan
    Write-Host ""
    
    # Instructions for obtaining a token
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
    
    # Prompt for the token
    Write-Host "Paste your token below and press Enter:" -ForegroundColor Green
    $token = Read-Host
    
    # Save the token globally for reuse
    $global:IntuneToken = $token
    
    # Validate the token format
    if (-not $token.StartsWith("eyJ")) {
        Write-Host "The token doesn't appear to be valid. It should start with 'eyJ'." -ForegroundColor Red
        exit
    }
    
    # Test the token with REST API approach first
    Write-Host "`nValidating token with a basic API call..." -ForegroundColor Yellow
    $headers = Test-IntuneToken -Token $token
    
    if ($headers) {
        # Save the headers globally for REST API calls (as a fallback)
        $global:IntuneHeaders = $headers
        
        # Display success for validation
        Write-Host "✓ Success! The token is valid for Intune access." -ForegroundColor Green
        
        # Now connect to Microsoft Graph PowerShell with the token
        $graphSuccess = Connect-GraphWithToken -Token $token
        
        if ($graphSuccess) {
            # Get available commands
            $commands = Get-IntuneCommands
            
            Write-Host "`nYou now have access to Microsoft Graph PowerShell cmdlets!" -ForegroundColor Green
            Write-Host "Here are some examples of available Intune cmdlets:" -ForegroundColor Green
            
            foreach ($command in $commands) {
                Write-Host "  • $command" -ForegroundColor White
            }
            
            # Show usage examples
            Write-Host "`n====== USAGE EXAMPLES ======" -ForegroundColor Cyan
            
            # Microsoft Graph PowerShell examples
            Write-Host "METHOD 1: Microsoft Graph PowerShell Cmdlets (Recommended)" -ForegroundColor Green
            Write-Host "Use the official Microsoft Graph PowerShell cmdlets for maximum functionality:" -ForegroundColor White
            Write-Host ""
            Write-Host "# Get all devices" -ForegroundColor Yellow
            Write-Host 'Get-MgDeviceManagementManagedDevice | Select-Object DeviceName, OperatingSystem, OSVersion' -ForegroundColor White
            Write-Host ""
            Write-Host "# Get specific devices by filter" -ForegroundColor Yellow
            Write-Host '$filter = "startsWith(deviceName,''WIN'')"' -ForegroundColor White
            Write-Host 'Get-MgDeviceManagementManagedDevice -Filter $filter' -ForegroundColor White
            Write-Host ""
            Write-Host "# Get device configurations" -ForegroundColor Yellow
            Write-Host 'Get-MgDeviceManagementDeviceConfiguration | Select-Object Id, DisplayName' -ForegroundColor White
            Write-Host ""
            Write-Host "# Execute actions on devices" -ForegroundColor Yellow
            Write-Host '$deviceId = "DEVICE-ID-HERE"' -ForegroundColor White
            Write-Host 'Invoke-MgRebootManagedDevice -ManagedDeviceId $deviceId' -ForegroundColor White
            Write-Host ""
            
            # REST API examples (as fallback)
            Write-Host "METHOD 2: REST API (Fallback approach)" -ForegroundColor Green
            Write-Host "If needed, you can still use REST API calls directly:" -ForegroundColor White
            Write-Host ""
            Write-Host "# Get all devices" -ForegroundColor Yellow
            Write-Host '$uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"' -ForegroundColor White
            Write-Host '$response = Invoke-RestMethod -Uri $uri -Headers $global:IntuneHeaders -Method Get' -ForegroundColor White
            Write-Host '$devices = $response.value' -ForegroundColor White
            Write-Host ""
            
            # Function to reconnect
            Write-Host "# Reconnect function (use if token expires)" -ForegroundColor Yellow
            Write-Host 'function Reconnect-IntuneGraph {' -ForegroundColor White
            Write-Host '    # Update REST API headers with fresh token' -ForegroundColor White
            Write-Host '    $global:IntuneHeaders = @{' -ForegroundColor White
            Write-Host '        "Authorization" = "Bearer $global:IntuneToken"' -ForegroundColor White
            Write-Host '        "Content-Type" = "application/json"' -ForegroundColor White
            Write-Host '    }' -ForegroundColor White
            Write-Host '    ' -ForegroundColor White
            Write-Host '    # Update Graph PowerShell connection' -ForegroundColor White
            Write-Host '    $secureToken = ConvertTo-SecureString -String $global:IntuneToken -AsPlainText -Force' -ForegroundColor White
            Write-Host '    Disconnect-MgGraph -ErrorAction SilentlyContinue' -ForegroundColor White
            Write-Host '    Connect-MgGraph -AccessToken $secureToken -NoWelcome' -ForegroundColor White
            Write-Host '    Write-Host "Reconnected to Intune!" -ForegroundColor Green' -ForegroundColor White
            Write-Host '}' -ForegroundColor White
            Write-Host ""
            
            Write-Host "NOTE: The token will expire after about 1 hour." -ForegroundColor Yellow
            Write-Host "      Run this script again when needed to get a fresh token, or use the Reconnect-IntuneGraph function." -ForegroundColor Yellow
        }
        else {
            Write-Host "`nFalling back to REST API approach." -ForegroundColor Yellow
            Write-Host "Your token is valid, but there was an issue setting up Microsoft Graph PowerShell." -ForegroundColor Yellow
            Write-Host "You can still use the REST API approach with `$global:IntuneHeaders:" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "# Example: Get all devices" -ForegroundColor Cyan
            Write-Host '$uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"' -ForegroundColor White
            Write-Host '$response = Invoke-RestMethod -Uri $uri -Headers $global:IntuneHeaders -Method Get' -ForegroundColor White
            Write-Host '$devices = $response.value' -ForegroundColor White
        }
    }
    else {
        Write-Host "`nThe token doesn't have necessary permissions to access Intune." -ForegroundColor Red
    }
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
} 