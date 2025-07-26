function Connect-IntuneWithToken {
    <#
    .SYNOPSIS
        Connects to Microsoft Intune using a browser-acquired token.
        
    .DESCRIPTION
        This function establishes a connection to Microsoft Intune using a token acquired from a browser.
        It validates the token, ensures required modules are installed, and establishes both PowerShell cmdlet
        and REST API access.
        
    .PARAMETER Token
        The authentication token string obtained from a browser session.
        
    .EXAMPLE
        Connect-IntuneWithToken -Token "eyJ0eXAiOiJKV..."
        
    .NOTES
        Requires Microsoft.Graph.Authentication and Microsoft.Graph.DeviceManagement modules.
        The token expires after approximately one hour and will need to be refreshed.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Token
    )
    
    try {
        # Convert token to SecureString
        $SecureToken = ConvertTo-SecureString -String $Token -AsPlainText -Force
        
        # Validate the token
        if (-not (Test-AuthToken -SecureToken $SecureToken)) {
            Write-Error "Authentication token validation failed. Please acquire a new token."
            return $false
        }
        
        # Verify required modules are installed
        $RequiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.DeviceManagement")
        foreach ($Module in $RequiredModules) {
            if (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Warning "$Module module not found. Attempting to install..."
                try {
                    Install-Module -Name $Module -Scope CurrentUser -Force -AllowClobber
                }
                catch {
                    Write-Error "Failed to install $Module. Please install it manually using: Install-Module -Name $Module -Scope CurrentUser -Force"
                    return $false
                }
            }
        }
        
        # Disconnect from any existing Graph connection
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
        catch {
            # Ignore any errors from disconnection
        }
        
        # Connect to Microsoft Graph with the token
        try {
            Connect-MgGraph -AccessToken $SecureToken -NoWelcome
            Write-Host "Successfully connected to Microsoft Graph!" -ForegroundColor Green
            
            # Store REST API headers for direct API calls if needed
            $script:IntuneHeaders = @{
                "Authorization" = "Bearer $Token"
                "Content-Type" = "application/json"
            }
            
            # Store the token for potential reconnection
            $script:IntuneToken = $Token
            
            return $true
        }
        catch {
            Write-Error "Failed to connect to Microsoft Graph: $_"
            return $false
        }
    }
    catch {
        Write-Error "Error in Connect-IntuneWithToken: $_"
        return $false
    }
} 