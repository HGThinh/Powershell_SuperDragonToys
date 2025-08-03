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

function Update-ProductDatabase {
    <#
    .SYNOPSIS
    Updates the product database (Excel file) with new product data.
    .DESCRIPTION
    This function takes an array of new product data, compares it with existing data
    in the Excel database, updates quantities for existing products, and adds new products.
    It also checks for price changes and logs them as warnings.
    If the Excel file doesn't exist, it creates a new one with an initial sample entry.
    Requires the `ImportExcel` module.
    .PARAMETER NewProductsData
    An array of PSCustomObject, each representing a new product with Barcode, ProductName, Quantity, and Price.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [array]$NewProductsData # Array of new product data to be added/updated.
    )
    Write-Log -Message "Updating product database..." # Logs the start of database update.

    if (-not (Test-Path $excelDatabasePath)) {
        Write-Log -Message "Excel database not found. Creating new file at $excelDatabasePath." # Logs that a new Excel file will be created.
        $initialData = @(
            [PSCustomObject]@{
                Barcode        = "00000"
                ProductName    = "Sample Product"
                Quantity       = 0
                Price          = 0.00
                LastUpdated    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                ImageLink      = ""
                WarehouseStock = 0
            }
        )
        $initialData | Export-Excel -Path $excelDatabasePath -AutoSize -NoHeader -ClearSheet # Creates a new Excel file with initial data.
        Write-Log -Message "Initial Excel database created." # Confirms initial creation.
    }

    $existingData = Import-Excel -Path $excelDatabasePath -ErrorAction SilentlyContinue # Imports existing data from Excel.
    if (-not $existingData -or ($existingData -isnot [array] -and $existingData.PSObject.TypeNames -notcontains "System.Collections.ArrayList")) {
        $existingData = @() # Initializes as empty array if Import-Excel returns nothing or a non-array.
        Write-Log -Message "Initialized existingData as empty array to prevent 'op_Addition' error." -Level "INFO" # Logs array initialization.
    }

    $updatedRows = 0 # Counter for updated rows.
    $addedRows = 0 # Counter for added rows.
    $priceChangeAlerts = @() # Array to store price change messages.

    foreach ($newProduct in $NewProductsData) {
        $found = $false # Flag to check if product exists.
        for ($i = 0; $i -lt $existingData.Count; $i++) {
            if ($existingData[$i].Barcode -eq $newProduct.Barcode) {
                $found = $true # Product found.
                if ($existingData[$i].Price -ne $newProduct.Price) {
                    $priceChangeAlerts += "Price change detected for $($newProduct.ProductName) (Barcode: $($newProduct.Barcode)): Old Price $($existingData[$i].Price) -> New Price $($newProduct.Price)" # Records price change.
                }
                $existingData[$i].Quantity = $existingData[$i].Quantity + $newProduct.Quantity # Updates quantity.
                $existingData[$i].Price = $newProduct.Price # Updates price.
                $existingData[$i].LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") # Updates last updated timestamp.
                $updatedRows++ # Increments updated count.
                Write-Log -Message "Updated quantity for '$($newProduct.ProductName)' (Barcode: $($newProduct.Barcode)). New Quantity: $($existingData[$i].Quantity)" -Level "INFO" # Logs quantity update.
                break # Exit loop once product is found and updated.
            }
        }
        if (-not $found) {
            $newProductRow = [PSCustomObject]@{
                Barcode        = $newProduct.Barcode
                ProductName    = $newProduct.ProductName
                Quantity       = $newProduct.Quantity
                Price          = $newProduct.Price
                LastUpdated    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                ImageLink      = ""
                WarehouseStock = 0
            }
            $existingData += $newProductRow # Adds new product to existing data.
            $addedRows++ # Increments added count.
            Write-Log -Message "Added new product: '$($newProduct.ProductName)' (Barcode: $($newProduct.Barcode)). Quantity: $($newProduct.Quantity)" -Level "INFO" # Logs new product addition.
        }
    }

    Try {
        $existingData | Export-Excel -Path $excelDatabasePath -AutoSize -ClearSheet # Exports all data back to Excel.
        Write-Log -Message "Excel database saved successfully. Updated $updatedRows items, Added $addedRows items." -Level "SUCCESS" # Confirms successful save.

        if ($priceChangeAlerts.Count -gt 0) {
            $alertBody = "The following price changes were detected during PDF processing:`n`n" + ($priceChangeAlerts -join "`n") # Formats price change alerts.
            Write-Log -Message "Price change alerts detected:`n$alertBody" -Level "WARN" # Logs price change alerts.
        }
    }
    Catch {
        Write-Log -Message "Error saving Excel database: $($_.Exception.Message)" -Level "ERROR" # Logs error during Excel save.
        throw # Throws the error.
    }
}

function Process-Image {
    <#
    .SYNOPSIS
    Resizes and optionally watermarks an image.
    .DESCRIPTION
    This function uses the System.Drawing namespace to load an image,
    resize it to a square target size while maintaining aspect ratio (padding with white),
    and optionally adds a text watermark to the bottom-left corner.
    The processed image is saved as a PNG file.
    .PARAMETER InputPath
    The full path to the input image file.
    .PARAMETER OutputPath
    The full path where the processed image will be saved.
    .PARAMETER WatermarkText
    Optional text to be added as a watermark. If not provided, no watermark is added.
    .PARAMETER TargetSize
    The desired width and height (in pixels) for the square output image. Default is 600.
    #>
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Path to the input Image file that needs to be processed")]
        [string]$InputPath, # Path to the original image file.
        [Parameter(Mandatory = $true, HelpMessage = "Path for the output image file where the processed image will be saved")]
        [string]$OutputPath, # Path where the processed image will be saved.
        [string]$WatermarkText = "", # Optional text watermark.
        [int]$TargetSize = 600 # Target size for the square output image.
    )
    $originalImage = $null # Variable to hold the original image object.
    $finalImage = $null # Variable to hold the processed image object.
    $graphics = $null # Variable to hold the graphics object.
    $font = $null # Variable to hold the font object for watermark.
    $brush = $null # Variable to hold the brush object for watermark.

    try {
        # Load the System.Drawing assembly (usually loaded by default in Windows PowerShell)
        Add-Type -AssemblyName System.Drawing # Ensures the System.Drawing assembly is loaded.

        $originalImage = [System.Drawing.Image]::FromFile($InputPath) # Loads the original image.
        $originalWidth = $originalImage.Width # Gets original image width.
        $originalHeight = $originalImage.Height # Gets original image height.

        $finalImage = New-Object System.Drawing.Bitmap($TargetSize, $TargetSize) # Creates a new blank bitmap for the final image.
        $graphics = [System.Drawing.Graphics]::FromImage($finalImage) # Creates a Graphics object to draw on the final image.
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic # Sets interpolation mode for quality.
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality # Sets smoothing mode for quality.
        $graphics.Clear([System.Drawing.Color]::White) # Fills the background with white.

        $ratio = [Math]::Min($TargetSize / $originalWidth, $TargetSize / $originalHeight) # Calculates scaling ratio to fit within target size.
        $newWidth = [int]($originalWidth * $ratio) # Calculates new width after scaling.
        $newHeight = [int]($originalHeight * $ratio) # Calculates new height after scaling.
        $x = ($TargetSize - $newWidth) / 2 # Calculates X offset for centering.
        $y = ($TargetSize - $newHeight) / 2 # Calculates Y offset for centering.
        $graphics.DrawImage($originalImage, $x, $y, $newWidth, $newHeight) # Draws the scaled original image onto the new bitmap.

        # Add watermark only if text is provided
        if (-not [string]::IsNullOrEmpty($WatermarkText)) {
            $font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold) # Creates a font for the watermark.
            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(128, 0, 0, 0)) # Creates a semi-transparent black brush.
            $textSize = $graphics.MeasureString($WatermarkText, $font) # Measures the size of the watermark text.
            $watermarkX = 20 # X position for the watermark.
            $watermarkY = $TargetSize - $textSize.Height - 20 # Y position for the watermark (bottom-left).
            $graphics.DrawString($WatermarkText, $font, $brush, $watermarkX, $watermarkY) # Draws the watermark text.
        }

        $finalImage.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png) # Saves the processed image as PNG.
        Write-Log -Message "Processed image using System.Drawing: '$InputPath' -> '$OutputPath'." -Level "INFO" # Logs successful image processing.

    }
    catch {
        Write-Log -Message "Error processing image '$InputPath' with System.Drawing: $($_.Exception.Message)" -Level "ERROR" # Logs error during image processing.
        throw # Throws the error.
    }
    finally {
        # Ensure all disposable objects are disposed
        if ($graphics) { $graphics.Dispose() } # Disposes the Graphics object.
        if ($finalImage) { $finalImage.Dispose() } # Disposes the final image bitmap.
        if ($originalImage) { $originalImage.Dispose() } # Disposes the original image.
        if ($font) { $font.Dispose() } # Disposes the font object.
        if ($brush) { $brush.Dispose() } # Disposes the brush object.
    }
}

function Process-Images {
    <#
    .SYNOPSIS
    Processes all PNG images in the source image folder and updates the product database.
    .DESCRIPTION
    This function scans the `$sourceImagesFolder` for PNG files.
    For each image, it calls `Process-Image` to resize and watermark it,
    then saves the processed image to `$imagesWithWatermarkFolder`.
    It attempts to link the processed image's path to a corresponding product
    in the Excel database based on the image's base name (assumed to be the barcode).
    Finally, it updates the Excel database with the image links.
    #>
    Write-Log -Message "Starting image processing..." # Logs the start of image processing.
    $imagesProcessedCount = 0 # Counter for processed images.
    $imageFiles = Get-ChildItem -Path $sourceImagesFolder -Filter "*.png" -File # Gets all PNG files in the source folder.

    if ($imageFiles.Count -eq 0) {
        Write-Log -Message "No image files found in '$sourceImagesFolder' to process." -Level "INFO" # Logs if no images are found.
        return # Exits the function.
    }

    $excelData = Import-Excel -Path $excelDatabasePath -ErrorAction SilentlyContinue # Imports existing Excel data.
    if (-not $excelData) {
        Write-Log -Message "Could not load Excel data for image linking. Skipping image linking." -Level "WARN" # Warns if Excel data cannot be loaded.
        $excelData = @() # Initializes as empty array to avoid errors.
    }

    foreach ($imageFile in $imageFiles) {
        $imageName = $imageFile.BaseName # Gets the file name without extension (assumed to be barcode).
        $outputImagePath = Join-Path $imagesWithWatermarkFolder "$($imageName)_processed.png" # Constructs output path for processed image.

        Try {
            Process-Image -InputPath $imageFile.FullName `
                -OutputPath $outputImagePath `
                -WatermarkText $defaultImageWatermarkText `
                -TargetSize $imageResizeWidth # Calls Process-Image to handle resizing and watermarking.
            
            Write-Log -Message "Image processed: '$($imageFile.Name)' -> '$($outputImagePath)'." -Level "INFO" # Logs successful image processing.

            # Link image to product in Excel (by Barcode matching BaseName)
            $barcodeToLink = $imageFile.BaseName # Barcode is assumed to be the image file's base name.

            for ($i = 0; $i -lt $excelData.Count; $i++) {
                if ($excelData[$i].Barcode -eq $barcodeToLink) {
                    $excelData[$i].ImageLink = $outputImagePath # Updates the ImageLink column in Excel data.
                    Write-Log -Message "Linked image '$($outputImagePath)' to barcode '$barcodeToLink' in Excel." -Level "INFO" # Logs image linking.
                    break # Exit loop once product is found.
                }
            }
            $imagesProcessedCount++ # Increments processed image count.
        }
        Catch {
            Write-Log -Message "Error processing image '$($imageFile.Name)': $($_.Exception.Message)" -Level "ERROR" # Logs error during image processing.
        }
    }

    if ($imagesProcessedCount -gt 0) {
        Try {
            $excelData | Export-Excel -Path $excelDatabasePath -AutoSize -ClearSheet # Exports updated Excel data.
            Write-Log -Message "Excel database updated with image links." -Level "SUCCESS" # Confirms Excel update.
        }
        Catch {
            Write-Log -Message "Error saving Excel database after image linking: $($_.Exception.Message)" -Level "ERROR" # Logs error saving Excel after image linking.
        }
    }
    Write-Log -Message "Image processing complete. Processed $imagesProcessedCount images." # Logs overall image processing completion.
}

function Analyze-Data {
    <#
    .SYNOPSIS
    Performs analytical reporting on product data from the Excel database.
    .DESCRIPTION
    This function reads the product data from the Excel database,
    calculates various statistics such as average price, highest/lowest priced products,
    total quantities, top products by quantity, and identifies products low on stock or out of stock.
    It captures all analysis output into a separate, timestamped log file within the Logs folder,
    while also sending critical messages to the main project log.
    .OUTPUTS
    A log file with detailed analysis (ProductAnalysis_YYYYMMDD_HHMMSS.txt) in the Logs folder.
    #>
    # Define a path for the specific analysis output file
    # This will place the analysis in the Logs folder with a timestamp
    $analysisOutputPath = Join-Path $logFolder "ProductAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt" # Path for the dedicated analysis log file.

    Write-Log -Message "Starting data analysis (Price, Products, Warehouse)..." -Level "INFO" # Logs the start of data analysis.

    # Use a StringBuilder to capture all analysis output
    $analysisContent = New-Object System.Text.StringBuilder # StringBuilder to accumulate analysis output.

    # Temporarily override Write-Log for analysis output capture
    # This block temporarily redefines Write-Log to also append to $analysisContent.
    # It saves the original Write-Log to restore it later.
    $script:OriginalWriteLog = Get-Item Function:\Write-Log # Saves the original Write-Log function definition.
    Function Write-Log {
        param(
            [Parameter(Mandatory = $true)][string]$Message,
            [string]$Level = "INFO"
        )
        # Append to our StringBuilder for the analysis file
        $analysisContent.AppendLine("[$Level] $Message") | Out-Null # Appends log message to StringBuilder.
        
        # Also call the original Write-Log function to keep logging to ProjectLog.txt
        # Ensure the original function can be called. This is a common pattern for "wrapping".
        # If your original Write-Log is defined globally, you can call it directly or via its fully qualified name.
        & $script:OriginalWriteLog -Message $Message -Level $Level # Calls the original Write-Log function.
    }

    Try {
        $data = Import-Excel -Path $excelDatabasePath -ErrorAction Stop # Imports all data from the Excel database.

        if (-not $data) {
            Write-Log -Message "No data found in Excel for analysis." -Level "INFO" # Logs if no data is found for analysis.
            return # Exits the function.
        }

        $averagePrice = ($data | Measure-Object -Property Price -Average).Average # Calculates average price.
        $highestPrice = $data | Sort-Object -Property Price -Descending | Select-Object -First 1 ProductName, Price # Finds product with highest price.
        $lowestPrice = $data | Sort-Object -Property Price | Select-Object -First 1 ProductName, Price # Finds product with lowest price.
        Write-Log -Message "Price Analysis:" -Level "INFO"
        Write-Log -Message "  Total Products: $($data.Count)" -Level "INFO"
        Write-Log -Message "  Average Price: $(Format-Number -Number $averagePrice -DecimalDigits 2)" -Level "INFO" # Logs average price (formatted).
        if ($highestPrice) { Write-Log -Message "  Highest Price Product: $($highestPrice.ProductName) ($($highestPrice.Price))" -Level "INFO" } # Logs highest price product.
        if ($lowestPrice) { Write-Log -Message "  Lowest Price Product: $($lowestPrice.ProductName) ($($lowestPrice.Price))" -Level "INFO" } # Logs lowest price product.

        $totalQuantity = ($data | Measure-Object -Property Quantity -Sum).Sum # Calculates total quantity across all products.
        $top5ProductsByQuantity = $data | Sort-Object -Property Quantity -Descending | Select-Object -First 5 ProductName, Quantity # Finds top 5 products by quantity.
        Write-Log -Message "Product Analysis:" -Level "INFO"
        Write-Log -Message "  Total Quantity Across All Products: $totalQuantity" -Level "INFO"
        $top5ProductsByQuantity | ForEach-Object { Write-Log -Message "    $($_.ProductName): $($_.Quantity)" -Level "INFO" } # Logs top 5 products.

        $totalWarehouseStock = ($data | Measure-Object -Property WarehouseStock -Sum).Sum # Calculates total warehouse stock.
        $productsLowOnStock = $data | Where-Object { $_.WarehouseStock -le 10 -and $_.WarehouseStock -gt 0 } # Finds products with low stock.
        $outOfStockProducts = $data | Where-Object { $_.WarehouseStock -eq 0 } # Finds out of stock products.

        Write-Log -Message "Warehouse Analysis:" -Level "INFO"
        if ($productsLowOnStock.Count -gt 0) {
            Write-Log -Message "  Products Low on Stock (<=10):" -Level "INFO"
            $productsLowOnStock | ForEach-Object { Write-Log -Message "    $($_.ProductName) (Barcode: $($_.Barcode)): $($_.WarehouseStock) in stock" -Level "INFO" } # Logs low stock products.
        }
        else {
            Write-Log -Message "  No products detected as 'low on stock'." -Level "INFO" # Logs if no low stock products.
        }
        if ($outOfStockProducts.Count -gt 0) {
            Write-Log -Message "  Products Out of Stock (0):" -Level "INFO"
            $outOfStockProducts | ForEach-Object { Write-Log -Message "    $($_.ProductName) (Barcode: $($_.Barcode))" -Level "INFO" } # Logs out of stock products.
        }
        else {
            Write-Log -Message "  No products detected as 'out of stock'." -Level "INFO" # Logs if no out of stock products.
        }

        Write-Log -Message "Data analysis complete." -Level "SUCCESS" # Logs completion of data analysis.
    }
    Catch {
        Write-Log -Message "Error during data analysis: $($_.Exception.Message)" -Level "ERROR" # Logs any error during analysis.
    }
    Finally {
        # Restore the original Write-Log function
        Remove-Item Function:\Write-Log -ErrorAction SilentlyContinue # Removes the temporary Write-Log function.
        Set-Item Function:\Write-Log $script:OriginalWriteLog -ErrorAction SilentlyContinue # Restores the original Write-Log function.
        Remove-Variable script:OriginalWriteLog -ErrorAction SilentlyContinue # Cleans up the temporary variable.

        # Save the captured analysis content to the dedicated file
        Try {
            $analysisContent.ToString() | Out-File -FilePath $analysisOutputPath -Encoding UTF8 # Saves the accumulated analysis content to its dedicated file.
            # Use original Write-Log here for consistency
            & $script:OriginalWriteLog -Message "Product analysis exported to: $analysisOutputPath" -Level "INFO" # Logs the path of the saved analysis file using the original Write-Log.
        }
        Catch {
            & $script:OriginalWriteLog -Message "Failed to export product analysis to $analysisOutputPath : $($_.Exception.Message)" -Level "ERROR" # Logs error if analysis file cannot be saved.
        }
        Remove-Variable analysisContent -ErrorAction SilentlyContinue # Cleans up the StringBuilder variable.
    }
}

#endregion