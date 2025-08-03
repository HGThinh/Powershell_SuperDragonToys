# Product Data Automation Script

## Overview

This PowerShell script automates the process of extracting product information from billing PDF files, processing product images, and maintaining a centralized product database in an Excel file. It continuously monitors designated folders for new PDF invoices and image files, updates product quantities and details, links images, and performs regular data analysis.

## Features

- **Automated PDF Parsing**: Extracts `ProductName`, `Quantity`, `Price`, and `Barcode` from billing PDFs using multiple robust extraction methods (iTextSharp, Microsoft Word COM, Adobe Reader COM, plain text fallback).
- **Dynamic Excel Database Management**:
  - Creates a new Excel database (`ProductDatabase.xlsx`) if one doesn't exist.
  - Updates existing product quantities based on new PDF data.
  - Adds new products to the database.
  - Detects and logs price changes for existing products.
- **Image Processing & Linking**:
  - Automatically resizes images to a specified `TargetSize` (default 600x600 pixels) and adds a configurable text watermark.
  - Links processed images to corresponding products in the Excel database using the image filename (assumed to be the product barcode).
- **Real-time Monitoring**: Continuously scans predefined input folders for new PDF and image files at a configurable interval.
- **Comprehensive Logging**: Maintains detailed logs of all operations, including file processing, database updates, and errors, in a `ProjectLog.txt` file.
- **Data Analysis & Reporting**: Periodically generates a detailed analytical report on product data (average price, top products, stock levels) and saves it to a separate timestamped log file.
- **Automated File Management**: Moves processed PDF and original image files to designated "Finished" folders.

## Prerequisites

- **PowerShell 5.1 or newer**: The script requires PowerShell version 5.1 or higher to run.
- **ImportExcel Module**: This PowerShell module is essential for reading from and writing to Excel files.
  - Install it by running PowerShell as Administrator and executing:
    ```powershell
    Install-Module -Name ImportExcel -Scope CurrentUser
    ```
- **iTextSharp (Optional but Recommended)**: The script attempts to automatically install or locate the `iTextSharp` library for PDF text extraction. If `iTextSharp` cannot be installed or found, the script will fall back to using Microsoft Word or Adobe Reader COM objects, which require the respective applications to be installed on the system.
  - For reliable PDF parsing, ensure your system has internet access for NuGet package installation or a compatible version of Microsoft Word or Adobe Acrobat/Reader.

## Configuration

All primary configuration settings are located at the beginning of the `Data_automation.ps1` script. You **must** adjust these paths and settings to match your environment.

```powershell
# **IMPORTANT: Make sure this path match**
$projectRoot = "C:\Project" # Base directory for all project files and folders.

$inputPdfFolder = $projectRoot # Directory where new PDF files are dropped for processing.
$finishedPdfFolder = Join-Path $projectRoot "FinishedPdfs" # Directory where processed PDF files are moved.
$finishedImageFolder = Join-Path $projectRoot "FinishedImages" # Directory where original image files are moved after processing.
$imagesWithWatermarkFolder = Join-Path $projectRoot "ImagesWithWatermark" # Directory where watermarked and resized images are saved.
$sourceImagesFolder = $projectRoot # Directory where new image files are dropped for processing.
$logFolder = Join-Path $projectRoot "Logs" # Directory for log files.
$dataFolder = $projectRoot # Directory for data files, including the Excel database.
$excelDatabasePath = Join-Path $dataFolder "ProductDatabase.xlsx" # Full path to the Excel database file.
$logFilePath = Join-Path $logFolder "ProjectLog.txt" # Full path to the main log file.

# Image Processing Configuration
$imageResizeWidth = 600 # pixels (Target width and height for resized images)
$defaultImageWatermarkText = "Name" # Text to be applied as a watermark on processed images.

# Scan interval for new PDFs and Images (in seconds)
$scanIntervalSeconds = 30 # Time in seconds the script waits before scanning for new files again.

## Author

Hoàng Gia Thịnh

## Acknowledgments

- Built using System.Drawing for .NET
- ImportExcel PowerShell Module: For providing robust and easy-to-use cmdlets for Excel file manipulation in PowerShell.
- iTextSharp: For its powerful PDF text extraction capabilities, which greatly enhance the functionality of script


## References

### Official Documentation
- [PowerShell Documentation](https://docs.microsoft.com/en-us/powershell/)
- [iTextSharp Documentation] (https://itextpdf.com/resources/api-documentation)
- [System.Drawing Namespace](https://docs.microsoft.com/en-us/dotnet/api/system.drawing)

### Tools and Dependencies
- [ImportExcel Module](https://github.com/dfinke/ImportExcel)
- [iTextSharp library, nugetURL]("https://www.nuget.org/api/v2/package/iTextSharp/5.5.13.3")

### Useful Resources
- [PowerShell: Resize-Image] (https://gist.github.com/someshinyobject/617bf00556bc43af87cd)
- [Powershell use .NET .DrawImage in System.Drawing] (https://stackoverflow.com/questions/55001057/powershell-use-net-drawimage-in-system-drawing)
- [How can I get PowerShell to read data in an Excel spreadsheet and apply to AD?] (https://community.spiceworks.com/t/how-can-i-get-powershell-to-read-data-in-an-excel-spreadsheet-and-apply-to-ad/818942/2)
-[Claude AI for PDF text extraction] (https://claude.ai/share/7fa07356-e21f-40a3-9f11-2b115bd8561d)
-[Gemini AI for comment base] (https://g.co/gemini/share/3848d71374af)
-[Gemini AI for loop to check new images] (https://g.co/gemini/share/317764724bc4)

```
