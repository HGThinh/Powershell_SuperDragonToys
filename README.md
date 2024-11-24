# PowerShell Automation Tools

A collection of PowerShell scripts for automation tasks including image processing and web automation using Selenium WebDriver.

## Scripts

### 1. Image Processor (Process-Image.ps1)

A PowerShell script for batch processing images with resizing and watermarking capabilities. This script provides a simple way to standardize image sizes while maintaining aspect ratios and optionally adding watermarks.

#### Features
- Resize images to a specified target size (default 600x600 pixels)
- Maintain aspect ratio during resizing
- Center images on a white background
- Add customizable watermark text
- High-quality image processing with bicubic interpolation
- Error handling and resource management

#### Prerequisites
- Windows PowerShell 5.1 or later
- System.Drawing assembly
- Sufficient permissions to read/write image files

#### Basic Usage
```powershell
./Process-Image.ps1 `
    -InputPath "C:\path\to\input\image.jpg" `
    -OutputPath "C:\path\to\output\image.jpg" `
    -WatermarkText "© Your Name" `
    -TargetSize 600
```

### 2. Chrome WebDriver Setup (ChromeWebDriverSetup.ps1)

A script for automating web interactions using Selenium WebDriver with Chrome, including path discovery and Excel data processing capabilities.

#### Features
- Automatic Chrome and ChromeDriver path detection
- Excel data import and processing
- Configurable Chrome options
- Robust error handling
- Form automation capabilities

#### Prerequisites
- Windows PowerShell 5.1 or later
- Selenium WebDriver
- ImportExcel PowerShell module
- Google Chrome browser
- ChromeDriver matching your Chrome version

#### Required Modules
```powershell
Import-Module -Name ImportExcel
Add-Type -AssemblyName System.Drawing
```

## Installation

1. Clone or download this repository
2. Install required dependencies:
   ```powershell
   Install-Module -Name ImportExcel
   Add-Type -AssemblyName System.Drawing
   ```
3. Download ChromeDriver matching your Chrome version from [ChromeDriver Downloads](https://chromedriver.chromium.org/downloads)

## Error Handling

Both scripts include comprehensive error handling:
- Input validation
- Resource management
- Detailed error messages
- Proper cleanup procedures

## Contributing

Feel free to submit issues and enhancement requests. Follow these steps to contribute:

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

## Author

Hoàng Gia Thịnh

## Acknowledgments

- Built using System.Drawing for .NET
- Selenium WebDriver for web automation
- ImportExcel module for Excel processing

## References

### Official Documentation
- [PowerShell Documentation](https://docs.microsoft.com/en-us/powershell/)
- [Selenium WebDriver Documentation](https://www.selenium.dev/documentation/webdriver/)
- [System.Drawing Namespace](https://docs.microsoft.com/en-us/dotnet/api/system.drawing)

### Tools and Dependencies
- [ChromeDriver Downloads](https://chromedriver.chromium.org/downloads)
- [ImportExcel Module](https://github.com/dfinke/ImportExcel)
- [Selenium WebDriver for PowerShell](https://www.powershellgallery.com/packages/Selenium)

### Useful Resources
- [PowerShell: Automatically Fill Online Form with data read from Spreadsheet] (https://www.youtube.com/watch?v=G6Ea3FCWLA4)
- [PowerShell: Resize-Image] (https://gist.github.com/someshinyobject/617bf00556bc43af87cd)
- [Powershell use .NET .DrawImage in System.Drawing] (https://stackoverflow.com/questions/55001057/powershell-use-net-drawimage-in-system-drawing)
- [How can I get PowerShell to read data in an Excel spreadsheet and apply to AD?] (https://community.spiceworks.com/t/how-can-i-get-powershell-to-read-data-in-an-excel-spreadsheet-and-apply-to-ad/818942/2)

