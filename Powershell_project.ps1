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

#region --- PDF Text Extraction Functions ---

function Install-iTextSharp {
    <#
    .SYNOPSIS
    Attempts to install and load the iTextSharp assembly for PDF text extraction.
    .DESCRIPTION
    This function tries several methods to make iTextSharp available:
    1. Checks common NuGet package paths for an existing DLL.
    2. Attempts to install iTextSharp via NuGet (requires PackageManagement and PSGallery).
    3. If NuGet fails, it tries to directly download the iTextSharp NuGet package and extract the DLL.
    It sets a global variable `$global:iTextSharpAvailable` to indicate success or failure.
    .OUTPUTS
    [boolean] Returns $true if iTextSharp is successfully loaded, $false otherwise.
    #>
    Write-Host "Setting up PDF text extraction..." -ForegroundColor Yellow
    
    # First, try to find existing iTextSharp DLL
    $possiblePaths = @(
        "${env:USERPROFILE}\.nuget\packages\itextsharp\*\lib\*\itextsharp.dll", # Common NuGet package path for current user.
        "${env:ProgramFiles}\PackageManagement\NuGet\Packages\iTextSharp*\lib\*\itextsharp.dll", # Common NuGet package path for all users.
        ".\itextsharp.dll", # Current directory.
        ".\lib\itextsharp.dll", # 'lib' subdirectory.
        (Join-Path $projectRoot "itextsharp.dll"), # Project root directory.
        (Join-Path $projectRoot "lib\itextsharp.dll") # 'lib' subdirectory within project root.
    )
    
    foreach ($path in $possiblePaths) {
        $foundDll = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($foundDll) {
            try {
                Add-Type -Path $foundDll.FullName # Loads the found iTextSharp DLL.
                Write-Host "iTextSharp loaded from: $($foundDll.FullName)" -ForegroundColor Green
                return $true # Indicates successful loading.
            }
            catch {
                Write-Host "Failed to load from $($foundDll.FullName)" -ForegroundColor Yellow
            }
        }
    }
    
    # Try to install using NuGet
    try {
        Write-Host "Installing NuGet provider..." -ForegroundColor Green
        Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction Stop # Ensures NuGet package provider is installed.
        
        Write-Host "Registering PSGallery repository..." -ForegroundColor Green
        if (!(Get-PSRepository -Name "PSGallery" -ErrorAction SilentlyContinue)) {
            Register-PSRepository -Default -ErrorAction Stop # Registers PSGallery if not already present.
        }
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted # Sets PSGallery as a trusted repository.
        
        Write-Host "Installing iTextSharp..." -ForegroundColor Green
        Install-Package -Name iTextSharp -Force -Scope CurrentUser -ErrorAction Stop # Installs iTextSharp via NuGet.
        
        # Try to find the installed DLL
        $installedPaths = @(
            "${env:USERPROFILE}\.nuget\packages\itextsharp\*\lib\*\itextsharp.dll" # Expected path after NuGet installation.
        )
        
        foreach ($path in $installedPaths) {
            $foundDll = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($foundDll) {
                Add-Type -Path $foundDll.FullName # Loads the newly installed iTextSharp DLL.
                Write-Host "iTextSharp installed and loaded successfully!" -ForegroundColor Green
                return $true # Indicates successful installation and loading.
            }
        }
    }
    catch {
        Write-Host "Package installation failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # Alternative: Download directly from NuGet
    Write-Host "Trying direct download from NuGet..." -ForegroundColor Yellow
    try {
        $nugetUrl = "https://www.nuget.org/api/v2/package/iTextSharp/5.5.13.3" # Direct URL to a specific iTextSharp NuGet package version.
        $zipPath = Join-Path $projectRoot "iTextSharp.zip" # Temporary path for the downloaded NuGet package.
        $extractPath = Join-Path $projectRoot "iTextSharp" # Directory to extract the NuGet package content.
        
        # Download package
        Invoke-WebRequest -Uri $nugetUrl -OutFile $zipPath -ErrorAction Stop # Downloads the NuGet package.
        
        # Extract ZIP
        if (Test-Path $extractPath) {
            Remove-Item $extractPath -Recurse -Force # Cleans up any existing extraction directory.
        }
        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force # Extracts the downloaded ZIP file.
        
        # Find and load DLL
        $dllPath = Get-ChildItem -Path "$extractPath\lib\*\itextsharp.dll" -Recurse | Select-Object -First 1 # Finds the iTextSharp DLL within the extracted content.
        if ($dllPath) {
            Add-Type -Path $dllPath.FullName # Loads the directly downloaded iTextSharp DLL.
            Write-Host "iTextSharp downloaded and loaded successfully!" -ForegroundColor Green
            
            # Cleanup
            Remove-Item $zipPath -Force -ErrorAction SilentlyContinue # Removes the temporary ZIP file.
            
            return $true # Indicates successful direct download and loading.
        }
    }
    catch {
        Write-Host "Direct download failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    Write-Host "All installation methods failed. Trying COM object fallback..." -ForegroundColor Yellow
    return $false # All attempts to load iTextSharp failed.
}

function Extract-TextFromPdf-iTextSharp {
    <#
    .SYNOPSIS
    Extracts text from a PDF file using the iTextSharp library.
    .DESCRIPTION
    This function utilizes the iTextSharp .NET library to read a PDF file
    and extract all readable text content from its pages.
    .PARAMETER PdfFilePath
    The full path to the PDF file from which to extract text.
    .OUTPUTS
    [string] The concatenated text extracted from the PDF, or $null if extraction fails.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$PdfFilePath # Path to the PDF file.
    )
    
    try {
        # Create PDF reader
        $reader = New-Object iTextSharp.text.pdf.PdfReader($PdfFilePath) # Initializes a PdfReader object.
        $stringBuilder = New-Object System.Text.StringBuilder # Used for efficient string concatenation.
        
        # Extract text from each page
        for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
            $pageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page) # Extracts text from a specific page.
            $stringBuilder.AppendLine($pageText) | Out-Null # Appends the page text to the StringBuilder.
        }
        
        $extractedText = $stringBuilder.ToString() # Converts the StringBuilder content to a single string.
        $reader.Close() # Closes the PDF reader.
        
        return $extractedText # Returns the extracted text.
    }
    catch {
        Write-Host "iTextSharp extraction failed: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null # Returns null if an error occurs during extraction.
    }
}

function Extract-TextFromPdf-ComObject {
    <#
    .SYNOPSIS
    Extracts text from a PDF file using COM objects (Microsoft Word or Adobe Reader).
    .DESCRIPTION
    This function serves as a fallback for PDF text extraction if iTextSharp is not available.
    It first attempts to open the PDF with Microsoft Word and extract its content.
    If Word fails, it tries to use Adobe Reader's COM object model (if Adobe Reader is installed).
    .PARAMETER PdfFilePath
    The full path to the PDF file from which to extract text.
    .OUTPUTS
    [string] The extracted text content, or $null if both COM object methods fail.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$PdfFilePath # Path to the PDF file.
    )
    
    try {
        # Try with Microsoft Word
        $word = New-Object -ComObject Word.Application # Creates a COM object for Microsoft Word.
        $word.Visible = $false # Keeps Word application hidden.
        
        # Open PDF in Word (Word can open PDFs)
        $doc = $word.Documents.Open($PdfFilePath) # Opens the PDF file in Word.
        
        # Get text content
        $text = $doc.Content.Text # Extracts all text from the Word document.
        
        # Close document and quit Word
        $doc.Close() # Closes the document.
        $word.Quit() # Quits the Word application.
        
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null # Releases the document COM object.
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null # Releases the Word application COM object.
        
        return $text # Returns the extracted text.
    }
    catch {
        Write-Host "Word COM object extraction failed: $($_.Exception.Message)" -ForegroundColor Yellow
        
        # Try alternative method using Adobe Reader (if installed)
        try {
            $reader = New-Object -ComObject AcroExch.PDDoc # Creates a COM object for Adobe Acrobat/Reader document.
            if ($reader.Open($PdfFilePath)) {
                # Attempts to open the PDF.
                $jso = $reader.GetJSObject() # Gets the JavaScript Object (JSO) interface.
                $text = ""
                for ($i = 0; $i -lt $jso.numPages; $i++) {
                    $text += $jso.getPageNthWord($i, 0, $jso.getPageNumWords($i)) # Concatenates words from each page.
                }
                $reader.Close() # Closes the PDF document.
                return $text # Returns the extracted text.
            }
        }
        catch {
            Write-Host "Adobe Reader COM object also failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        return $null # Returns null if both COM object methods fail.
    }
}

#endregion

#region --- Main Functions (Updated) ---
function Write-Log {
    <#
    .SYNOPSIS
    Writes a timestamped log message to a file and the console.
    .DESCRIPTION
    This function creates a log entry with a timestamp and a specified level (INFO, WARN, ERROR, SUCCESS, DEBUG).
    It writes the entry to a central log file (`$logFilePath`) and also displays it on the console.
    It automatically creates the log folder if it doesn't exist.
    .PARAMETER Message
    The log message to write.
    .PARAMETER Level
    The severity level of the log message. Valid values are "INFO", "WARN", "ERROR", "SUCCESS", "DEBUG".
    Default is "INFO".
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$Message, # The message to be logged.
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO" # The severity level of the log message.
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" # Formats the current date and time.
    $logEntry = "[$timestamp] [$Level] $Message" # Constructs the log entry string.

    if (-not (Test-Path $logFolder)) {
        try {
            New-Item -ItemType Directory -Path $logFolder -Force | Out-Null # Creates the log folder if it doesn't exist.
        }
        catch {
            Write-Warning "Could not create log folder: $logFolder. Log messages will only go to console." # Warns if folder creation fails.
            Write-Host "$logEntry" # Outputs to console only if folder creation fails.
            return
        }
    }

    Add-Content -Path $logFilePath -Value $logEntry # Appends the log entry to the log file.
    Write-Host "$logEntry" # Outputs the log entry to the console.
}

function Create-ProjectFolders {
    <#
    .SYNOPSIS
    Ensures all necessary project folders exist.
    .DESCRIPTION
    This function iterates through a predefined list of project-specific directories
    and creates them if they do not already exist. It logs the creation of each folder
    or throws an error if folder creation fails.
    #>
    Write-Host "Ensuring project folders exist..."
    $folders = @(
        $inputPdfFolder, # Input directory for PDFs.
        $finishedPdfFolder, # Directory for finished PDFs.
        $finishedImageFolder, # Directory for finished images.
        $imagesWithWatermarkFolder, # Directory for watermarked images.
        $sourceImagesFolder, # Source directory for images.
        $logFolder, # Directory for logs.
        $dataFolder # Directory for data files.
    )
    foreach ($folder in $folders) {
        if (-not (Test-Path $folder)) {
            try {
                New-Item -ItemType Directory -Path $folder -Force | Out-Null # Creates the directory.
                Write-Host "Created folder: $folder" # Confirms folder creation.
            }
            catch {
                Write-Host "Failed to create folder $folder : $($_.Exception.Message)" -ForegroundColor Red # Reports creation failure.
                throw # Throws the error to stop script execution if a critical folder cannot be created.
            }
        }
    }
    Write-Log -Message "Project folders checked." # Logs that folder check is complete.
}
#endregion