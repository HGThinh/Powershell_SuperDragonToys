# Requires -Module System.Drawing, ImportExcel
<#
.SYNOPSIS
    Unified Script for:
        - Image Processing
        - Watermarking 
        - Data Import Automation

.DESCRIPTION
    This script combines two functionalities:
    1. Image processing with:
        - Resizing (default 600x600).
        - Centering
        - Watermarking.
    2. Web automation script identifies:
        - Chrome and ChromeDriver installation paths.
        - Configures Chrome options.
        - Reads data from an Excel file
        - Automates web interactions using Selenium WebDriver.

.NOTES
    Author      : Hoàng Gia Thịnh
    Created     : 24/11/2024
    Modified    : 05/01/2025
    Version     : 1.2
    Dependencies: System.Drawing, Selenium WebDriver, ImportExcel module
    
.PARAMETER ChromePaths
    Array of potential paths for the Chrome executable.

.PARAMETER ChromeDriverPaths
    Array of potential paths for the ChromeDriver executable.

.PARAMETER ExcelFilePath
    Path to the Excel file containing web form data.
    
.PARAMETER WebsiteUrl
    URL of the website to automate.
    
.PARAMETER InputPath
    Full path to the input image file that needs to be processed.
    
.PARAMETER OutputPath
    Full path where the processed image will be saved.
    
.PARAMETER WatermarkText
    Optional text to be added as a watermark on the image. Defaults to an empty string.
    
.PARAMETER TargetSize
    Optional target size for the square output image. Defaults to 600 pixels.
    
.EXAMPLE ProcessImage
    Process-Image -InputPath "C:\input\photo.jpg" -OutputPath "C:\output\processed_photo.jpg" -WatermarkText "© Your Name"

.EXAMPLE ImportData
  Import-Data -ExcelFilePath "C:\Users\Data_powershell.xlsx" -WebsiteUrl "http://127.0.0.1:5500/index.html"
  
.LINK Chromedriver
    https://chromedriver.chromium.org/downloads
    
.LINK SystemDrawing
    https://docs.microsoft.com/en-us/dotnet/api/system.drawing
#>

# Import necessary modules
Add-Type -AssemblyName System.Drawing
Import-Module -Name ImportExcel -ErrorAction Stop

# Section 1: Image Processing
function Process-Image {
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Path to the input Image file that needs to be processed")]
        [string]$InputPath,
        [Parameter(Mandatory = $true, HelpMessage = "Path for the output image file where the processed image will be saved")]
        [string]$OutputPath,
        [string]$WatermarkText = "",
        [int]$TargetSize = 600
    )
    try {
        $originalImage = [System.Drawing.Image]::FromFile($InputPath)
        $originalWidth = $originalImage.Width
        $originalHeight = $originalImage.Height

        $finalImage = New-Object System.Drawing.Bitmap($TargetSize, $TargetSize)
        $graphics = [System.Drawing.Graphics]::FromImage($finalImage)
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $graphics.Clear([System.Drawing.Color]::White)

        $ratio = [Math]::Min($TargetSize / $originalWidth, $TargetSize / $originalHeight)
        $newWidth = [int]($originalWidth * $ratio)
        $newHeight = [int]($originalHeight * $ratio)
        $x = ($TargetSize - $newWidth) / 2
        $y = ($TargetSize - $newHeight) / 2
        $graphics.DrawImage($originalImage, $x, $y, $newWidth, $newHeight)

        $font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold)
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(128, 0, 0, 0))
        $textSize = $graphics.MeasureString($WatermarkText, $font)
        $watermarkX = 20
        $watermarkY = $TargetSize - $textSize.Height - 20
        $graphics.DrawString($WatermarkText, $font, $brush, $watermarkX, $watermarkY)
        $finalImage.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    }
    finally {
        if ($graphics) { $graphics.Dispose() }
        if ($finalImage) { $finalImage.Dispose() }
        if ($originalImage) { $originalImage.Dispose() }
        if ($font) { $font.Dispose() }
        if ($brush) { $brush.Dispose() }
    }
}

# Section 2: Data Import Automation
function Import-Data {
    param (
        [Parameter(Mandatory = $false, HelpMessage = "Array of potential paths for the Chrome executable")]
        [string[]]$ChromePaths = @("C:\Users\chrome.exe"),
        [Parameter(Mandatory = $false, HelpMessage = "Array of potential paths for the ChromeDriver executable")]
        [string[]]$ChromeDriverPaths = @("C:\Users\chromedriver-win64"),
        [Parameter(Mandatory = $true, HelpMessage = "Path to the Excel file containing web form data")]
        [string]$ExcelFilePath = "C:\Users\Data_powershell.xlsx",
        [Parameter(Mandatory = $true, HelpMessage = "URL of the website to automate")]
        [string]$WebsiteUrl = "http://127.0.0.1:5500/index.html"
    )
    function Find-File {
        param ([string[]]$Paths)
        foreach ($path in $Paths) {
            $expandedPath = $ExecutionContext.InvokeCommand.ExpandString($path)
            if (Test-Path $expandedPath) {
                return $expandedPath
            }
        }
        return $null
    }

    $ChromePath = Find-File -Paths $ChromePaths
    $ChromeDriverPath = Find-File -Paths $ChromeDriverPaths

    if (!$ChromePath) {
        Write-Error "Chrome executable not found."
        exit 1
    }
    if (!$ChromeDriverPath) {
        Write-Error "ChromeDriver not found."
        exit 1
    }

    try {
        $ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
        $ChromeOptions.BinaryLocation = $ChromePath
        $ChromeOptions.AddArgument("--start-maximized")
        $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($ChromeDriverPath)
        $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $ChromeOptions)

        $driver.Navigate().GoToUrl($WebsiteUrl)
        $ExcelData = Import-Excel -Path $ExcelFilePath

        foreach ($row in $ExcelData) {
            try {
                $driver.FindElementById("name").SendKeys($row.Name)
                $driver.FindElementById("overview").SendKeys($row.Overview)
                $driver.FindElementById("details").SendKeys($row.Details)
                $driver.FindElementById("tags").SendKeys($row.Tags)
                $driver.FindElementById("categories").SendKeys($row.Categories)
                $driver.FindElementById("cost").SendKeys($row.Cost)
                $driver.FindElementById("price").SendKeys($row.Price)
                $driver.FindElementById("save").Click()
                Start-Sleep -Seconds 2
            } catch {
                Write-Warning "Error processing row: $($_.Exception.Message)"
            }
        }
    } finally {
        if ($driver) {
            $driver.Quit()
            $driver.Dispose()
        }
    }
}

# Example usage of each function:
# Process-Image -InputPath "C:\input\photo.jpg" -OutputPath "C:\output\processed_photo.jpg" -WatermarkText "Your Name"
# Import-Data -ExcelFilePath "C:\Users\Data_powershell.xlsx" -WebsiteUrl "http://127.0.0.1:5500/index.html"
