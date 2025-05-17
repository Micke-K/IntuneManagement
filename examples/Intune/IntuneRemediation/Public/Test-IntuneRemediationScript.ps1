function Test-IntuneRemediationScript {
    <#
    .SYNOPSIS
        Tests a remediation script pair locally before uploading to Intune.
        
    .DESCRIPTION
        This function tests a detection and remediation script pair locally to verify they work as expected.
        The detection script should return a boolean value (or exit code 0/1), and the remediation script
        should fix the issue if the detection script returns $false.
        
    .PARAMETER DetectionScriptPath
        The path to the detection script file.
        
    .PARAMETER RemediationScriptPath
        The path to the remediation script file.
        
    .PARAMETER Cycles
        The number of detection-remediation cycles to run. Default is 1.
        
    .PARAMETER NoRemediate
        If specified, only runs the detection script without remediation.
        
    .PARAMETER ShowScriptOutput
        If specified, displays the raw script output rather than just success/failure.
        
    .EXAMPLE
        Test-IntuneRemediationScript -DetectionScriptPath ".\Detect.ps1" -RemediationScriptPath ".\Remediate.ps1" -Cycles 2
        
    .NOTES
        This function does not upload anything to Intune, it only tests the scripts locally.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$DetectionScriptPath,
        
        [Parameter(Mandatory = $false)]
        [string]$RemediationScriptPath,
        
        [Parameter(Mandatory = $false)]
        [int]$Cycles = 1,
        
        [Parameter(Mandatory = $false)]
        [switch]$NoRemediate,
        
        [Parameter(Mandatory = $false)]
        [switch]$ShowScriptOutput
    )
    
    try {
        # Verify the detection script exists
        if (-not (Test-Path -Path $DetectionScriptPath)) {
            Write-Error "Detection script not found at path: $DetectionScriptPath"
            return $false
        }
        
        # Verify the remediation script exists (if not using NoRemediate)
        if (-not $NoRemediate -and -not [string]::IsNullOrEmpty($RemediationScriptPath)) {
            if (-not (Test-Path -Path $RemediationScriptPath)) {
                Write-Error "Remediation script not found at path: $RemediationScriptPath"
                return $false
            }
        }
        
        Write-Host "Starting Intune remediation script test" -ForegroundColor Cyan
        Write-Host "Detection script: $DetectionScriptPath" -ForegroundColor Cyan
        if (-not $NoRemediate -and -not [string]::IsNullOrEmpty($RemediationScriptPath)) {
            Write-Host "Remediation script: $RemediationScriptPath" -ForegroundColor Cyan
        }
        Write-Host "-------------------------------------------" -ForegroundColor Cyan
        
        $TestResults = @{
            Cycles = $Cycles
            DetectionResults = @()
            RemediationResults = @()
            FinalStatus = "Unknown"
        }
        
        for ($i = 1; $i -le $Cycles; $i++) {
            Write-Host "Cycle $i of $Cycles" -ForegroundColor Yellow
            
            # Run detection script
            Write-Host "  Running detection script..." -ForegroundColor Gray
            $DetectionOutput = $null
            $DetectionSuccess = $false
            
            try {
                # Execute the detection script and capture its output
                $DetectionOutput = & $DetectionScriptPath
                $DetectionExitCode = $LASTEXITCODE
                
                # Check if output is boolean or exit code indicates success/failure
                if ($DetectionOutput -is [bool]) {
                    $DetectionSuccess = $DetectionOutput
                }
                elseif ($DetectionExitCode -eq 0) {
                    $DetectionSuccess = $true
                }
                else {
                    $DetectionSuccess = $false
                }
                
                if ($ShowScriptOutput) {
                    Write-Host "    Output: $DetectionOutput" -ForegroundColor Gray
                }
                
                if ($DetectionSuccess) {
                    Write-Host "  Detection result: " -ForegroundColor Gray -NoNewline
                    Write-Host "COMPLIANT" -ForegroundColor Green
                }
                else {
                    Write-Host "  Detection result: " -ForegroundColor Gray -NoNewline
                    Write-Host "NON-COMPLIANT" -ForegroundColor Red
                }
                
                $TestResults.DetectionResults += @{
                    Cycle = $i
                    Success = $DetectionSuccess
                    Output = $DetectionOutput
                    ExitCode = $DetectionExitCode
                }
            }
            catch {
                Write-Host "  Detection script error: $_" -ForegroundColor Red
                $TestResults.DetectionResults += @{
                    Cycle = $i
                    Success = $false
                    Error = $_
                }
                continue
            }
            
            # Run remediation script if detection failed and not using NoRemediate
            if (-not $DetectionSuccess -and -not $NoRemediate -and -not [string]::IsNullOrEmpty($RemediationScriptPath)) {
                Write-Host "  Running remediation script..." -ForegroundColor Gray
                $RemediationOutput = $null
                
                try {
                    # Execute the remediation script and capture its output
                    $RemediationOutput = & $RemediationScriptPath
                    $RemediationExitCode = $LASTEXITCODE
                    
                    if ($ShowScriptOutput) {
                        Write-Host "    Output: $RemediationOutput" -ForegroundColor Gray
                    }
                    
                    Write-Host "  Remediation completed with exit code: $RemediationExitCode" -ForegroundColor Gray
                    
                    $TestResults.RemediationResults += @{
                        Cycle = $i
                        Output = $RemediationOutput
                        ExitCode = $RemediationExitCode
                    }
                    
                    # Run detection again to see if remediation fixed the issue
                    Write-Host "  Running detection script again to verify remediation..." -ForegroundColor Gray
                    $VerificationOutput = & $DetectionScriptPath
                    $VerificationExitCode = $LASTEXITCODE
                    
                    # Check if output is boolean or exit code indicates success/failure
                    if ($VerificationOutput -is [bool]) {
                        $VerificationSuccess = $VerificationOutput
                    }
                    elseif ($VerificationExitCode -eq 0) {
                        $VerificationSuccess = $true
                    }
                    else {
                        $VerificationSuccess = $false
                    }
                    
                    if ($ShowScriptOutput) {
                        Write-Host "    Output: $VerificationOutput" -ForegroundColor Gray
                    }
                    
                    if ($VerificationSuccess) {
                        Write-Host "  Verification result: " -ForegroundColor Gray -NoNewline
                        Write-Host "REMEDIATION SUCCESSFUL" -ForegroundColor Green
                    }
                    else {
                        Write-Host "  Verification result: " -ForegroundColor Gray -NoNewline
                        Write-Host "REMEDIATION FAILED" -ForegroundColor Red
                    }
                    
                    $TestResults.DetectionResults += @{
                        Cycle = "$i-verification"
                        Success = $VerificationSuccess
                        Output = $VerificationOutput
                        ExitCode = $VerificationExitCode
                    }
                }
                catch {
                    Write-Host "  Remediation script error: $_" -ForegroundColor Red
                    $TestResults.RemediationResults += @{
                        Cycle = $i
                        Error = $_
                    }
                }
            }
            
            Write-Host ""
        }
        
        # Determine final status
        $FinalDetection = $TestResults.DetectionResults | Select-Object -Last 1
        if ($FinalDetection.Success) {
            $TestResults.FinalStatus = "Compliant"
            Write-Host "Final status: " -ForegroundColor Cyan -NoNewline
            Write-Host "COMPLIANT" -ForegroundColor Green
        }
        else {
            $TestResults.FinalStatus = "Non-Compliant"
            Write-Host "Final status: " -ForegroundColor Cyan -NoNewline
            Write-Host "NON-COMPLIANT" -ForegroundColor Red
        }
        
        return $TestResults
    }
    catch {
        Write-Error "Error in Test-IntuneRemediationScript: $_"
        return $false
    }
} 