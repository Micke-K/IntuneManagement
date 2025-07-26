<#
.SYNOPSIS
    Detects if the Windows Time service is running.

.DESCRIPTION
    This script checks if the Windows Time (w32time) service is running.
    Returns compliance state to Intune based on the service status.

.NOTES
    File Name: Detect-CriticalService.ps1
    Author: Intune Administrator
    Created: 2023-11-09
    Version: 1.0
#>

# Define the service to check
$serviceName = "w32time"

try {
    # Get the service
    $service = Get-Service -Name $serviceName -ErrorAction Stop
    
    # Check if the service is running
    if ($service.Status -eq "Running") {
        Write-Host "The $serviceName service is running correctly."
        # Return compliant (exit code 0)
        exit 0
    }
    else {
        Write-Host "The $serviceName service is NOT running. Current status: $($service.Status)"
        # Return non-compliant (exit code 1)
        exit 1
    }
}
catch {
    Write-Error "Error checking service status: $_"
    # Return non-compliant on error
    exit 1
} 