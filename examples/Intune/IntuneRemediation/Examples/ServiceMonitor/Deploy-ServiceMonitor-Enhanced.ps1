<#
.SYNOPSIS
    Tests and deploys the Windows Time service monitor remediation scripts to Intune.

.DESCRIPTION
    This script demonstrates how to use the IntuneRemediation module to:
    1. Test the detection and remediation scripts locally
    2. Connect to Intune using integrated browser-token authentication
    3. Upload the scripts to Intune as a remediation script package

.NOTES
    File Name: Deploy-ServiceMonitor-Enhanced.ps1
    Author: Intune Administrator
    Created: 2023-11-09
    Version: 1.0
#>

# Ensure we are in the script's directory
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDir = Split-Path $scriptPath -Parent
Set-Location -Path $scriptDir

# Import the IntuneRemediation module
# Assuming the module is already installed or available in a parent directory
try {
    $modulePath = (Get-Item -Path $scriptDir).Parent.Parent.FullName
    Import-Module -Name "$modulePath\IntuneRemediation.psd1" -Force -ErrorAction Stop
    Write-Host "IntuneRemediation module imported successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to import IntuneRemediation module: $_"
    Write-Host "Please ensure the module is installed or adjust the path accordingly." -ForegroundColor Yellow
    exit 1
}

# Step 1: Test the remediation scripts locally
Write-Host "`n=== TESTING REMEDIATION SCRIPTS LOCALLY ===" -ForegroundColor Cyan
$detectionScriptPath = "$scriptDir\Detect-CriticalService.ps1"
$remediationScriptPath = "$scriptDir\Remediate-CriticalService.ps1"

$testResults = Test-IntuneRemediationScript -DetectionScriptPath $detectionScriptPath -RemediationScriptPath $remediationScriptPath -Cycles 1 -ShowScriptOutput

if ($testResults.FinalStatus -eq "Compliant") {
    Write-Host "The remediation test completed successfully and the service is in a compliant state." -ForegroundColor Green
}
else {
    Write-Host "The remediation test completed, but the service is still in a non-compliant state." -ForegroundColor Red
    Write-Host "Review the test output above for more information." -ForegroundColor Yellow
    
    $proceed = Read-Host "Do you want to proceed with deploying to Intune anyway? (Y/N)"
    if ($proceed -ne "Y") {
        Write-Host "Deployment cancelled." -ForegroundColor Yellow
        exit 0
    }
}

# Step 2: Connect to Intune using the integrated authentication approach
Write-Host "`n=== CONNECTING TO INTUNE ===" -ForegroundColor Cyan
Write-Host "Using browser token authentication to connect to Intune..." -ForegroundColor Yellow

# You can optionally pre-supply a token if you have one
# $myToken = "eyJ0eXAiOi..."
# $connected = Initialize-IntuneConnection -Token $myToken

# Or just call without a token to be prompted
$connected = Initialize-IntuneConnection

if (-not $connected) {
    Write-Error "Failed to connect to Intune. Please check your token and try again."
    exit 1
}

# Step 3: Upload the scripts to Intune
Write-Host "`n=== UPLOADING REMEDIATION SCRIPTS TO INTUNE ===" -ForegroundColor Cyan

# Read the script contents
$detectionScriptContent = Get-Content -Path $detectionScriptPath -Raw
$remediationScriptContent = Get-Content -Path $remediationScriptPath -Raw

# Create the remediation script package in Intune
$remediationScriptParams = @{
    DisplayName = "Windows Time Service Monitor"
    Description = "Monitors the Windows Time service and starts it if it's not running."
    Publisher = "IT Department"
    DetectionScriptContent = $detectionScriptContent
    RemediationScriptContent = $remediationScriptContent
    RunAsAccount = "System"  # Run as system account
    RunAs32Bit = $false      # Run as 64-bit
    EnforceSignatureCheck = $false
}

$result = New-IntuneRemediationScript @remediationScriptParams

if ($result) {
    Write-Host "`nSuccessfully created remediation script package in Intune!" -ForegroundColor Green
    Write-Host "You can now assign this script to device groups through the Intune portal." -ForegroundColor Green
    Write-Host "Script ID: $($result.id)" -ForegroundColor Green
}
else {
    Write-Host "`nFailed to create remediation script package in Intune." -ForegroundColor Red
    Write-Host "Please check the error messages above for more information." -ForegroundColor Yellow
} 