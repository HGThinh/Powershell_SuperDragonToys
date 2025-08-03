# Requires -Modules ImportExcel
# Requires -Version 5.1

#region --- Configuration ---

# **IMPORTANT: Make sure this path match**
$projectRoot = "C:\Users\hoang\Desktop\Test28_07_2025"

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
$imageResizeWidth = 600 # pixels
$defaultImageWatermarkText = "Thinh" 

# Scan interval for new PDFs and Images (in seconds)
$scanIntervalSeconds = 30 # Time in seconds the script waits before scanning for new files again.

#endregion
