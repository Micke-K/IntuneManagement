<#
.SYNOPSIS
    Remediates the Windows Time service by starting it if not running.

.DESCRIPTION
    This script starts the Windows Time (w32time) service if it's not running.
    Attempts to also set the service to automatic startup if it isn't already.

.NOTES
    File Name: Remediate-CriticalService.ps1
    Author: Intune Administrator
    Created: 2023-11-09
    Version: 1.0
#>

# Define the service to check
$serviceName = "w32time"

try {
    # Get the service
    $service = Get-Service -Name $serviceName -ErrorAction Stop
    
    # Log the current status
    Write-Host "Current status of $serviceName service: $($service.Status)"
    
    # Check if service is not running
    if ($service.Status -ne "Running") {
        Write-Host "Attempting to start the $serviceName service..."
        
        # Try to start the service
        Start-Service -Name $serviceName -ErrorAction Stop
        
        # Verify the service started successfully
        $service = Get-Service -Name $serviceName -ErrorAction Stop
        if ($service.Status -eq "Running") {
            Write-Host "The $serviceName service was started successfully."
            
            # Ensure the service is set to start automatically
            $startupType = (Get-WmiObject -Class Win32_Service -Filter "Name='$serviceName'").StartMode
            if ($startupType -ne "Auto") {
                Write-Host "Setting $serviceName service to start automatically..."
                Set-Service -Name $serviceName -StartupType Automatic
                Write-Host "Service startup type set to Automatic."
            }
            
            # Return success (exit code 0)
            exit 0
        }
        else {
            Write-Error "Failed to start the $serviceName service. Current status: $($service.Status)"
            # Return failure (exit code 1)
            exit 1
        }
    }
    else {
        Write-Host "The $serviceName service is already running."
        # Return success (exit code 0)
        exit 0
    }
}
catch {
    Write-Error "Error remediating service: $_"
    # Return failure (exit code 1)
    exit 1
} 