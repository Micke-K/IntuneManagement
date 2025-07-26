function New-IntuneRemediationScript {
    <#
    .SYNOPSIS
        Creates a new remediation script package in Microsoft Intune.
        
    .DESCRIPTION
        This function creates a new remediation script package in Microsoft Intune with detection
        and remediation scripts. It requires a valid connection to Microsoft Graph.
        
    .PARAMETER DisplayName
        The display name for the remediation script package in Intune.
        
    .PARAMETER Description
        A description for the remediation script package.
        
    .PARAMETER Publisher
        The publisher of the remediation script package.
        
    .PARAMETER DetectionScriptContent
        The content of the detection script as a string.
        
    .PARAMETER RemediationScriptContent
        The content of the remediation script as a string.
        
    .PARAMETER RunAs32Bit
        If specified, runs the scripts in 32-bit PowerShell.
        
    .PARAMETER EnforceSignatureCheck
        If specified, enforces script signature check.
        
    .PARAMETER RunAsAccount
        The account context to run the scripts. Options are: System, User
        
    .EXAMPLE
        $detection = Get-Content -Path ".\DetectionScript.ps1" -Raw
        $remediation = Get-Content -Path ".\RemediationScript.ps1" -Raw
        New-IntuneRemediationScript -DisplayName "Fix Service State" -Description "Checks and fixes critical service state" -Publisher "IT Department" -DetectionScriptContent $detection -RemediationScriptContent $remediation -RunAsAccount "System"
        
    .NOTES
        Requires a connection to Microsoft Graph using Connect-IntuneWithToken.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $false)]
        [string]$Description = "",
        
        [Parameter(Mandatory = $false)]
        [string]$Publisher = "IT Department",
        
        [Parameter(Mandatory = $true)]
        [string]$DetectionScriptContent,
        
        [Parameter(Mandatory = $true)]
        [string]$RemediationScriptContent,
        
        [Parameter(Mandatory = $false)]
        [switch]$RunAs32Bit = $false,
        
        [Parameter(Mandatory = $false)]
        [switch]$EnforceSignatureCheck = $false,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("System", "User")]
        [string]$RunAsAccount = "System"
    )
    
    try {
        # Check if connected to Graph
        try {
            $GraphConnection = Get-MgContext
            if (-not $GraphConnection) {
                Write-Error "Not connected to Microsoft Graph. Please connect using Connect-IntuneWithToken first."
                return $false
            }
        }
        catch {
            Write-Error "Not connected to Microsoft Graph. Please connect using Connect-IntuneWithToken first."
            return $false
        }
        
        # Convert script contents to Base64
        $DetectionScriptB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($DetectionScriptContent))
        $RemediationScriptB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($RemediationScriptContent))
        
        # Determine run as account value for API
        $RunAsAccountValue = switch ($RunAsAccount) {
            "System" { "system" }
            "User" { "user" }
            default { "system" }
        }
        
        # Create the remediation script object
        $RemediationScriptBody = @{
            "@odata.type" = "#microsoft.graph.deviceHealthScript"
            "displayName" = $DisplayName
            "description" = $Description
            "publisher" = $Publisher
            "detectionScriptContent" = $DetectionScriptB64
            "remediationScriptContent" = $RemediationScriptB64
            "runAsAccount" = $RunAsAccountValue
            "enforceSignatureCheck" = $EnforceSignatureCheck.IsPresent
            "runAs32Bit" = $RunAs32Bit.IsPresent
            "roleScopeTagIds" = @("0")  # Default scope tag
        }
        
        Write-Verbose "Creating remediation script: $DisplayName"
        
        # POST to Microsoft Graph API
        $Url = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts"
        $RemediationScriptJson = $RemediationScriptBody | ConvertTo-Json -Depth 10
        
        try {
            $Response = Invoke-MgGraphRequest -Method POST -Uri $Url -Body $RemediationScriptJson -ContentType "application/json"
            
            Write-Host "Successfully created remediation script '$DisplayName' with ID: $($Response.id)" -ForegroundColor Green
            return $Response
        }
        catch {
            Write-Error "Failed to create remediation script: $_"
            
            # More detailed error information
            $ErrorDetails = $_.ErrorDetails | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($ErrorDetails) {
                Write-Error "Error details: $($ErrorDetails.error.message)"
            }
            return $false
        }
    }
    catch {
        Write-Error "Error in New-IntuneRemediationScript: $_"
        return $false
    }
} 