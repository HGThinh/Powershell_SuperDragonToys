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

function Parse-BillingPdf {
    <#
    .SYNOPSIS
    Extracts product information from a billing PDF file.
    .DESCRIPTION
    This function attempts to extract text from a given PDF using three methods in order:
    1. iTextSharp (if available and loaded)
    2. COM objects (Microsoft Word or Adobe Reader)
    3. Direct plain text reading (as a last resort for simple text-based PDFs).
    Once text is extracted, it applies multiple regex patterns to find product details
    (ProductName, Quantity, Price, Barcode) and returns them as an array of custom objects.
    It logs each step and any failures.
    .PARAMETER PdfPath
    The full path to the billing PDF file to be parsed.
    .OUTPUTS
    [array] An array of PSCustomObject, each representing a product found in the PDF,
            or $null if no product data can be extracted.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$PdfPath # Path to the PDF file to parse.
    )
    Write-Log -Message "Attempting to parse PDF: $PdfPath" # Logs the start of PDF parsing.

    Try {
        # Extract text from PDF using multiple methods
        $pdfContent = $null # Initializes variable to store extracted PDF content.
        
        # Method 1: Try iTextSharp (if available)
        if ($global:iTextSharpAvailable) {
            Write-Log -Message "Extracting text using iTextSharp..." -Level "INFO"
            $pdfContent = Extract-TextFromPdf-iTextSharp -PdfFilePath $PdfPath # Tries iTextSharp extraction.
        }
        
        # Method 2: Fallback to COM objects if iTextSharp failed
        if (-not $pdfContent) {
            Write-Log -Message "Extracting text using COM objects..." -Level "INFO"
            $pdfContent = Extract-TextFromPdf-ComObject -PdfFilePath $PdfPath # Tries COM object extraction.
        }
        
        # Method 3: Last resort - try to read as plain text (if PDF is text-based)
        if (-not $pdfContent) {
            Write-Log -Message "Attempting to read PDF as plain text..." -Level "WARN"
            try {
                $pdfContent = Get-Content $PdfPath -Raw -Encoding UTF8 # Tries reading as plain text.
            }
            catch {
                Write-Log -Message "Plain text reading also failed." -Level "ERROR" # Logs plain text reading failure.
            }
        }
        
        if (-not $pdfContent) {
            Write-Log -Message "All PDF text extraction methods failed for: $PdfPath" -Level "ERROR" # Logs complete extraction failure.
            return $null # Returns null if no text could be extracted.
        }
        
        Write-Log -Message "Successfully extracted text from PDF. Content length: $($pdfContent.Length) characters" -Level "INFO" # Logs successful text extraction.
        
        # Parse the extracted content for product information
        $products = @() # Initializes an empty array to store product data.
        
        # Updated regex patterns to be more flexible
        $patterns = @(
            # Original pattern
            "(?s)Product:\s*(?<ProductName>[^,]+),\s*Quantity:\s*(?<Quantity>\d+),\s*Price:\s*(?<Price>[\d.]+),\s*Barcode:\s*(?<Barcode>\d+)", # Pattern 1
            
            # Alternative patterns for different PDF formats
            "(?i)(?<ProductName>[^\r\n]+)\s+Qty:\s*(?<Quantity>\d+)\s+Price:\s*\$?(?<Price>[\d.]+)\s+Barcode:\s*(?<Barcode>\d+)", # Pattern 2
            "(?i)Barcode:\s*(?<Barcode>\d+)\s+(?<ProductName>[^\r\n]+)\s+Quantity:\s*(?<Quantity>\d+)\s+Price:\s*\$?(?<Price>[\d.]+)", # Pattern 3
            
            # More flexible pattern
            "(?i)(?<ProductName>[A-Za-z0-9\s\-_]+)\s*.*?(?<Quantity>\d+)\s*.*?(?<Price>\d+\.?\d*)\s*.*?(?<Barcode>\d{8,})" # Pattern 4
        )
        
        $foundMatch = $false # Flag to track if any pattern found a match.
        foreach ($pattern in $patterns) {
            $matches = [regex]::Matches($pdfContent, $pattern) # Attempts to find matches using the current pattern.
            
            if ($matches.Count -gt 0) {
                Write-Log -Message "Found $($matches.Count) product matches using pattern: $($patterns.IndexOf($pattern) + 1)" -Level "INFO" # Logs successful pattern match.
                $foundMatch = $true # Sets flag to true.
                
                foreach ($match in $matches) {
                    $product = [PSCustomObject]@{
                        Barcode     = $match.Groups["Barcode"].Value.Trim() # Extracts and trims Barcode.
                        ProductName = $match.Groups["ProductName"].Value.Trim() # Extracts and trims Product Name.
                        Quantity    = [int]$match.Groups["Quantity"].Value # Extracts and converts Quantity to integer.
                        Price       = [double]$match.Groups["Price"].Value # Extracts and converts Price to double.
                    }
                    $products += $product # Adds the extracted product to the array.
                }
                break # Use the first pattern that works
            }
        }
        
        if (-not $foundMatch) {
            Write-Log -Message "No product data found in PDF using any pattern. PDF content preview (first 500 chars):" -Level "WARN" # Warns if no product data is found.
            Write-Log -Message $pdfContent.Substring(0, [Math]::Min(500, $pdfContent.Length)) -Level "DEBUG" # Logs a preview of the PDF content.
            
            # Save extracted content for debugging
            $debugPath = Join-Path $logFolder "extracted_content_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt" # Creates a debug file path.
            $pdfContent | Out-File -FilePath $debugPath -Encoding UTF8 # Saves the full extracted content to a debug file.
            Write-Log -Message "Full extracted content saved to: $debugPath" -Level "INFO" # Logs the debug file path.
            
            return $null # Returns null if no product data is found.
        }

        Write-Log -Message "Successfully extracted $($products.Count) products from $PdfPath" -Level "SUCCESS" # Logs successful product extraction.
        return $products # Returns the array of extracted products.
    }
    Catch {
        Write-Log -Message "Error parsing PDF '$PdfPath': $($_.Exception.Message)" -Level "ERROR" # Logs any error during PDF parsing.
        return $null # Returns null if an error occurs.
    }
}

#endregion