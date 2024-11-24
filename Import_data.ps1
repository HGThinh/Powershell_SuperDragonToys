<#
.SYNOPSIS
    Chrome and WebDriver Path Discovery Script

.DESCRIPTION
    This script identifies Chrome and ChromeDriver installation paths, configures Chrome options, reads data from an Excel file, and automates web interactions using Selenium WebDriver.

.PARAMETER ChromePaths
    Array of potential paths for the Chrome executable.

.PARAMETER ChromeDriverPaths
    Array of potential paths for the ChromeDriver executable.

.EXAMPLE
    ./ChromeWebDriverSetup.ps1

.NOTES
    Author      : Hoàng Gia Thịnh
    Created     : 24/11/2024
    Modified    : 24/11/2024
    Version     : 1.0
    Dependencies: Selenium WebDriver, ImportExcel module

.LINK
    https://chromedriver.chromium.org/downloads
#>

# Enable strict mode and configure error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Define paths
$ChromePaths = @(
    "C:\Users\chrome.exe"
)

$ChromeDriverPaths = @(
    "C:\Users\chromedriver-win64"
)

# Function to find a file in multiple locations
function Find-File {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Paths
    )
    foreach ($path in $Paths) {
        $expandedPath = $ExecutionContext.InvokeCommand.ExpandString($path)
        if (Test-Path $expandedPath) {
            return $expandedPath
        }
    }
    return $null
}

# Find Chrome and ChromeDriver
$ChromePath = Find-File -Paths $ChromePaths
$ChromeDriverPath = Find-File -Paths $ChromeDriverPaths

# Validate findings
if (!$ChromePath) {
    Write-Error "Chrome executable not found. Please install Google Chrome."
    exit 1
}

if (!$ChromeDriverPath) {
    Write-Error "ChromeDriver not found. Download it from https://chromedriver.chromium.org/downloads"
    exit 1
}

# Import required modules
try {
    Import-Module -Name ImportExcel -ErrorAction Stop
    Add-Type -Path "C:\Users\WebDriver.dll"
    Add-Type -Path "C:\Users\WebDriver.Support.dll"
} catch {
    Write-Error "Required modules or assemblies could not be loaded: $($_.Exception.Message)"
    exit 1
}

# Main logic
try {
    # Configure Chrome options
    $ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $ChromeOptions.BinaryLocation = $ChromePath
    $ChromeOptions.AddArgument("--start-maximized")
    $ChromeOptions.AddArgument("--no-sandbox")
    $ChromeOptions.AddArgument("--disable-dev-shm-usage")

    # Initialize WebDriver
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($ChromeDriverPath)
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $ChromeOptions)

    # Define file paths
    $ExcelFilePath = "C:\Users\Data_powershell.xlsx"
    $WebsiteUrl = "http://127.0.0.1:5500/index.html"

    # Navigate to the website
    $driver.Navigate().GoToUrl($WebsiteUrl)

    # Read Excel data
    $ExcelData = Import-Excel -Path $ExcelFilePath

    # Process each row
    foreach ($row in $ExcelData) {
        try {
            # Fill form fields
            $driver.Manage().Timeouts().ImplicitWait = [TimeSpan]::FromSeconds(10)

            $driver.FindElementById("name").SendKeys($row.Name)
            $driver.FindElementById("overview").SendKeys($row.Overview)
            $driver.FindElementById("details").SendKeys($row.Details)
            $driver.FindElementById("tags").SendKeys($row.Tags)
            $driver.FindElementById("categories").SendKeys($row.Categories)
            $driver.FindElementById("cost").SendKeys($row.Cost)
            $driver.FindElementById("price").SendKeys($row.Price)

            # Click save button
            $driver.FindElementById("save").Click()

            # Pause for a moment
            Start-Sleep -Seconds 2
        } catch {
            Write-Warning "Error processing row: $($row.Name)"
            Write-Warning $_.Exception.Message
            continue
        }
    }
} catch {
    Write-Error "Automation Error: $($_.Exception.Message)"
} finally {
    # Clean up
    if ($driver) {
        $driver.Quit()
        $driver.Dispose()
    }
    Write-Host "Automation completed."
}
