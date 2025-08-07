# RunTests.ps1
# Test runner script for the PowerShell automation project

param(
    [string]$TestPath = ".\Pester_tests.ps1",
    [switch]$Detailed,
    [switch]$InstallPester
)

# Check if Pester is installed
if (-not (Get-Module -ListAvailable -Name Pester)) {
    if ($InstallPester) {
        Write-Host "Installing Pester module..." -ForegroundColor Yellow
        try {
            Install-Module -Name Pester -Force -SkipPublisherCheck -Scope CurrentUser
            Write-Host "Pester installed successfully!" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install Pester: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Please run: Install-Module -Name Pester -Force -SkipPublisherCheck" -ForegroundColor Yellow
            exit 1
        }
        
    }
    else {
        Write-Host "Pester module not found. Please install it first:" -ForegroundColor Red
        Write-Host "Install-Module -Name Pester -Force -SkipPublisherCheck" -ForegroundColor Yellow
        Write-Host "Or run this script with -InstallPester switch" -ForegroundColor Yellow
        exit 1
    }
}

# Import Pester
Import-Module Pester -Force

Write-Host "=== PowerShell Automation Project - Test Suite ===" -ForegroundColor Cyan
Write-Host "Running tests from: $TestPath" -ForegroundColor White
Write-Host ""

# Configure Pester
$pesterConfig = New-PesterConfiguration
$pesterConfig.Run.Path = $TestPath
$pesterConfig.Run.PassThru = $true
$pesterConfig.Output.Verbosity = if ($Detailed) { 'Detailed' } else { 'Normal' }
$pesterConfig.CodeCoverage.Enabled = $false  # Disable code coverage for simplicity
$pesterConfig.TestResult.Enabled = $true
$pesterConfig.TestResult.OutputPath = ".\TestResults.xml"

# Run the tests
try {
    $result = Invoke-Pester -Configuration $pesterConfig
    
    # Display summary
    Write-Host ""
    Write-Host "=== TEST SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Total Tests: $($result.TotalCount)" -ForegroundColor White
    Write-Host "Passed: $($result.PassedCount)" -ForegroundColor Green
    Write-Host "Failed: $($result.FailedCount)" -ForegroundColor $(if ($result.FailedCount -gt 0) { 'Red' } else { 'Green' })
    Write-Host "Skipped: $($result.SkippedCount)" -ForegroundColor Yellow
    Write-Host "Duration: $($result.Duration)" -ForegroundColor White
    
    if ($result.FailedCount -gt 0) {
        Write-Host ""
        Write-Host "=== PASSED TESTS ===" -ForegroundColor Green
        $result.Passed | ForEach-Object {
            Write-Host " + $($_.Name)" -ForegroundColor Green
        }
        Write-Host ""
        Write-Host "=== FAILED TESTS ===" -ForegroundColor Red
        $result.Failed | ForEach-Object {
            Write-Host "  - $($_.Name)" -ForegroundColor Red
            Write-Host "    $($_.ErrorRecord.Exception.Message)" -ForegroundColor DarkRed
        }
        Write-Host ""
        Write-Host "=== SKIPPED TESTS ===" -ForegroundColor Yellow
        $result.Skipped | ForEach-Object {
            Write-Host " x  $($_.Name)" -ForegroundColor Yellow
        }
    }
    
    # Calculate success rate
    $successRate = if ($result.TotalCount -gt 0) {
        [Math]::Round(($result.PassedCount / $result.TotalCount) * 100, 1)
    }
    else { 0 }
    
    Write-Host ""
    Write-Host "Success Rate: $successRate%" -ForegroundColor $(if ($successRate -ge 80) { 'Green' } elseif ($successRate -ge 60) { 'Yellow' } else { 'Red' })
    
    # Provide feedback based on results
    if ($successRate -ge 90) {
        Write-Host "Excellent! Your code is well-tested and robust." -ForegroundColor Green
    }
    elseif ($successRate -ge 80) {
        Write-Host "Good test coverage. Consider addressing any failing tests." -ForegroundColor Yellow
    }
    elseif ($successRate -ge 60) {
        Write-Host "Moderate test coverage. Several tests need attention." -ForegroundColor Yellow
    }
    else {
        Write-Host "Low test coverage. Please review and fix failing tests." -ForegroundColor Red
    }
    
    # Exit with appropriate code
    exit $(if ($result.FailedCount -eq 0) { 0 } else { 1 })
}
catch {
    Write-Host "Error running tests: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}