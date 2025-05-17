# IntuneRemediation PowerShell Module

## Overview

The IntuneRemediation module provides an easy way to create, test, and deploy remediation scripts to Microsoft Intune using PowerShell. It uses browser-acquired tokens for authentication, eliminating the need for app registrations.

## Features

- **Token-based Authentication**: Connect to Intune using tokens from your browser session
- **Script Testing**: Test your remediation scripts locally before deploying to Intune
- **Script Deployment**: Easily upload detection and remediation scripts to Intune
- **Proper Structure**: Follows PowerShell best practices with Public/Private function separation
- **Integrated Authentication**: Self-contained implementation of browser token authentication

## Installation

### Manual Installation

1. Clone or download this repository
2. Copy the `IntuneRemediation` folder to one of your PowerShell module paths:
   - `$Home\Documents\WindowsPowerShell\Modules` (for current user)
   - `$env:ProgramFiles\WindowsPowerShell\Modules` (for all users)

### PowerShell Gallery (Recommended)

```powershell
Install-Module -Name IntuneRemediation -Scope CurrentUser
```

## Quick Start

### 1. Import the Module

```powershell
Import-Module IntuneRemediation
```

### 2. Connect to Intune

```powershell
# This will prompt you for a token with instructions
Initialize-IntuneConnection

# Or if you already have a token
Initialize-IntuneConnection -Token "eyJ0eXAiOi..."
```

### 3. Test a Remediation Script Locally

```powershell
Test-IntuneRemediationScript -DetectionScriptPath ".\Detect.ps1" -RemediationScriptPath ".\Remediate.ps1" -Cycles 1
```

### 4. Upload a Remediation Script to Intune

```powershell
$detection = Get-Content -Path ".\Detect.ps1" -Raw
$remediation = Get-Content -Path ".\Remediate.ps1" -Raw

New-IntuneRemediationScript -DisplayName "My Remediation Script" `
                           -Description "Fixes a common issue" `
                           -Publisher "IT Department" `
                           -DetectionScriptContent $detection `
                           -RemediationScriptContent $remediation `
                           -RunAsAccount "System"
```

## Sample Scripts

The module includes sample remediation scripts in the `Examples` folder:

- **ServiceMonitor**: Monitors and fixes the Windows Time service if it's not running
- **Deploy-ServiceMonitor.ps1**: Demonstrates how to test and deploy the scripts to Intune
- **Deploy-ServiceMonitor-Enhanced.ps1**: Demonstrates our simplified approach to token authentication

## How to Create a Remediation Script

### 1. Detection Script

The detection script should:
- Return exit code 0 (or $true) if the system is compliant
- Return exit code 1 (or $false) if the system is non-compliant and requires remediation

Example:
```powershell
# Check if a service is running
$service = Get-Service -Name "ServiceName"
if ($service.Status -eq "Running") {
    # Compliant
    exit 0
} else {
    # Non-compliant
    exit 1
}
```

### 2. Remediation Script

The remediation script should:
- Fix the issue identified by the detection script
- Return exit code 0 if remediation was successful
- Return exit code 1 if remediation failed

Example:
```powershell
# Start the service
try {
    Start-Service -Name "ServiceName" -ErrorAction Stop
    exit 0  # Success
} catch {
    exit 1  # Failed
}
```

## Getting Browser Tokens

The `Initialize-IntuneConnection` function will guide you through the token acquisition process, but here's how to do it manually:

1. Open a browser and go to https://endpoint.microsoft.com/
2. Sign in with your admin account
3. Press F12 to open developer tools
4. Go to the Network tab
5. Refresh the page (F5)
6. Filter requests by typing "graph.microsoft" in the filter box
7. Click on any request to graph.microsoft.com
8. In the Headers tab, find "Authorization: Bearer eyJ..."
9. Copy the entire token (starting with "eyJ")

## Functions

### Public Functions

- **Initialize-IntuneConnection**: Connect to Intune using browser token authentication
- **Connect-IntuneWithToken**: Legacy method to connect to Intune using a token
- **New-IntuneRemediationScript**: Create a new remediation script in Intune
- **Test-IntuneRemediationScript**: Test a remediation script locally before deployment

## Requirements

- PowerShell 5.1 or later
- Microsoft.Graph.Authentication module (installed automatically if needed)
- Microsoft.Graph.DeviceManagement module (installed automatically if needed)
- Administrative access to Microsoft Intune

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 