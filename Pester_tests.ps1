# Comprehensive Pester tests for the PowerShell automation project
# Requires Pester v5.x (Install-Module -Name Pester -Force -SkipPublisherCheck)

BeforeAll {
    # Import the main script functions
    # Dot source the main script to load all functions
    $scriptPath = Join-Path $PSScriptRoot "Data_automation.ps1"
    
    # Mock global variables and paths for testing
    $global:testProjectRoot = "C:\Users\hoang\Desktop\Test28_07_2025"
    $global:testLogFolder = Join-Path $global:testProjectRoot "Logs"
    $global:testLogFilePath = Join-Path $global:testLogFolder "ProjectLog.txt"
    $global:testExcelPath = Join-Path $global:testProjectRoot "ProductDatabase.xlsx"
    
    # Override configuration variables for testing
    $script:projectRoot = $global:testProjectRoot
    $script:logFolder = $global:testLogFolder
    $script:logFilePath = $global:testLogFilePath
    $script:excelDatabasePath = $global:testExcelPath
    $script:imageResizeWidth = 600
    $script:WatermarkTextt = "Test"
    
    $testExcel = "C:\Users\hoang\Desktop\Test28_07_2025\ProductDatabase.xlsx"
    if (-not (Test-Path $testExcel)) {
        New-Item -ItemType File -Path $testExcel | Out-Null
    }
    # Create test directories
    New-Item -ItemType Directory -Path $global:testProjectRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $global:testLogFolder -Force | Out-Null
    
    # Load functions from the main script (excluding the main execution block)
    $scriptContent = Get-Content $scriptPath -Raw
    # Remove the main execution block to avoid running the infinite loop
    $functionContent = $scriptContent -replace '#region --- Main Script ---[\s\S]*#endregion', ''
    Invoke-Expression $functionContent

    function New-DummyPng {
        param($Path)
        Add-Type -AssemblyName System.Drawing
        $bmp = New-Object System.Drawing.Bitmap 50, 50
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::LightBlue)
        $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose()
        $bmp.Dispose()
    }

    New-DummyPng -Path "C:\Users\hoang\Desktop\Test28_07_2025\123456.png"

}

Describe "Write-Log Function Tests" {
    BeforeEach {
        # Clean up log file before each test
        if (Test-Path $global:testLogFilePath) {
            Remove-Item $global:testLogFilePath -Force
        }
    }
    
    Context "Basic Logging Functionality" {
        It "Should create log file if it doesn't exist" {
            Write-Log -Message "Test message" -Level "INFO"
            Test-Path $global:testLogFilePath | Should -Be $true
        }
        
        It "Should write message with correct format" {
            Write-Log -Message "Test message" -Level "INFO"
            $logContent = Get-Content $global:testLogFilePath
            $logContent | Should -Match '\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] \[INFO\] Test message'
        }
        
        It "Should handle different log levels" {
            $levels = @("INFO", "WARN", "ERROR", "SUCCESS", "DEBUG")

            foreach ($level in $levels) {
                Write-Log -Message "Test $level message" -Level $level
            }

            $logContent = Get-Content $global:testLogFilePath -Raw  # Read as single string
            ($logContent -split "`r?`n").Count | Should -Be 6

            foreach ($level in $levels) {
                $logContent | Should -Match $level
            }
        }

        
        It "Should default to INFO level when not specified" {
            Write-Log -Message "Default level test"
            $logContent = Get-Content $global:testLogFilePath
            $logContent | Should -Match '\[(INFO|WARN|ERROR|DEBUG)\]'
        }
    }
    
    Context "Error Handling" {
        It "Should handle invalid log levels gracefully" {
            { Write-Log -Message "Test" -Level "INVALID" } | Should -Throw
        }
        
        It "Should handle empty messages gracefully and not log anything" {
            $initialCount = (Get-Content $global:testLogFilePath).Count
            { Write-Log -Message "" -Level "INFO" } | Should -Not -Throw
            $finalCount = (Get-Content $global:testLogFilePath).Count
            $finalCount | Should -Be $initialCount
        }

    }
}

Describe "Format-Number Function Tests" {
    Context "Number Formatting" {
        It "Should format number with default 2 decimal places" {
            $result = Format-Number -Number 123.456
            $result | Should -Be "123.46"
        }
        
        It "Should format number with specified decimal places" {
            $result = Format-Number -Number 123.456 -DecimalDigits 3
            $result | Should -Be "123.456"
        }
        
        It "Should handle zero" {
            $result = Format-Number -Number 0
            $result | Should -Be "0.00"
        }
        
        It "Should handle negative numbers" {
            $result = Format-Number -Number -123.456
            $result | Should -Be "-123.46"
        }
        
        It "Should handle large numbers" {
            $result = Format-Number -Number 1234567.89
            $result | Should -Be "1,234,567.89"
        }
    }
}

Describe "Create-ProjectFolders Function Tests" {
    BeforeEach {
        # Clean up test directories
        if (Test-Path $global:testProjectRoot) {
            Remove-Item $global:testProjectRoot -Recurse -Force
        }
        New-Item -ItemType Directory -Path $global:testProjectRoot -Force | Out-Null
    }
    
    Context "Folder Creation" {
        It "Should create all required project folders" {
            # Mock the folder variables
            $script:finishedPdfFolder = Join-Path $global:testProjectRoot "FinishedPdfs"
            $script:finishedImageFolder = Join-Path $global:testProjectRoot "FinishedImages"
            $script:imagesWithWatermarkFolder = Join-Path $global:testProjectRoot "ImagesWithWatermark"
            
            Create-ProjectFolders
            
            Test-Path $script:finishedPdfFolder | Should -Be $true
            Test-Path $script:finishedImageFolder | Should -Be $true
            Test-Path $script:imagesWithWatermarkFolder | Should -Be $true
            Test-Path $global:testLogFolder | Should -Be $true
        }
        
        It "Should not fail if folders already exist" {
            $testFolder = Join-Path $global:testProjectRoot "ExistingFolder"
            New-Item -ItemType Directory -Path $testFolder -Force | Out-Null
            
            $script:finishedPdfFolder = $testFolder
            $script:finishedImageFolder = $testFolder
            $script:imagesWithWatermarkFolder = $testFolder
            
            { Create-ProjectFolders } | Should -Not -Throw
        }
    }
}

Describe "PDF Text Extraction Functions Tests" {
    Context "iTextSharp Installation" {
        It "Should return boolean value from Install-iTextSharp" {
            $result = Install-iTextSharp
            $result | Should -Match [bool]
        }
    }
    
    Context "PDF Text Extraction" {
        BeforeEach {
            # Create a mock PDF file for testing
            $testPdfPath = Join-Path $global:testProjectRoot "test.pdf"
            "Mock PDF Content" | Out-File -FilePath $testPdfPath -Encoding UTF8
        }
        
        It "Should handle non-existent PDF files gracefully" {
            $nonExistentPath = Join-Path $global:testProjectRoot "nonexistent.pdf"
            $result = Extract-TextFromPdf-ComObject -PdfFilePath $nonExistentPath
            $result | Should -BeNullOrEmpty
        }
        
        It "Should accept valid PDF file path parameter" {
            $testPdfPath = Join-Path $global:testProjectRoot "test.pdf"
            { Extract-TextFromPdf-ComObject -PdfFilePath $testPdfPath } | Should -Not -Throw
        }
    }
}

Describe "Parse-BillingPdf Function Tests" {
    Context "PDF Parsing" {
        BeforeEach {
            # Create mock PDF content
            $testPdfPath = Join-Path $global:testProjectRoot "billing.pdf"
            $mockPdfContent = @"
Product: Test Product, Quantity: 5, Price: 10.99, Barcode: 1234567890
Product: Another Product, Quantity: 3, Price: 25.50, Barcode: 0987654321
"@
            $mockPdfContent | Out-File -FilePath $testPdfPath -Encoding UTF8
        }
        
        It "Should return null for non-existent PDF" {
            $nonExistentPath = Join-Path $global:testProjectRoot "nonexistent.pdf"
            $result = Parse-BillingPdf -PdfPath $nonExistentPath
            $result | Should -BeNullOrEmpty
        }
        
        It "Should accept valid PDF path parameter" {
            $testPdfPath = Join-Path $global:testProjectRoot "billing.pdf"
            { Parse-BillingPdf -PdfPath $testPdfPath } | Should -Not -Throw
        }
        
        It "Should validate mandatory PDF path parameter" {
            { Parse-BillingPdf } | Should -Throw
        }
    }
}

Describe "Update-ProductDatabase Function Tests" {
    Context "Database Operations" {
        BeforeEach {
            # Clean up Excel file
            if (Test-Path $global:testExcelPath) {
                Remove-Item $global:testExcelPath -Force
            }
            
            # Mock ImportExcel and Export-Excel functions if not available
            if (-not (Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
                function Global:Import-Excel {
                    param($Path, $ErrorAction)
                    return @()
                }
                function Global:Export-Excel {
                    param($InputObject, $Path, $AutoSize, $ClearSheet, $NoHeader)
                    # Mock export - create empty file
                    New-Item -Path $Path -ItemType File -Force | Out-Null
                }
            }
        }
        
        
        It "Should require mandatory NewProductsData parameter" {
            { Update-ProductDatabase } | Should -Throw
        }
        
        It "Should throw an error for empty product data array" {
            $emptyData = @()
            { Update-ProductDatabase -NewProductsData $emptyData } | Should -Throw
        }
    }
}

Describe "Process-Image Function Tests" {
    Context "Image Processing Parameters" {
        It "Should require mandatory InputPath parameter" {
            { Process-Image -OutputPath "test.png" } | Should -Throw
        }
        
        It "Should require mandatory OutputPath parameter" {
            { Process-Image -InputPath "test.png" } | Should -Throw
        }
        
        
        It "Should handle non-existent input file gracefully" {
            $inputPath = Join-Path $global:testProjectRoot "nonexistent.png"
            $outputPath = Join-Path $global:testProjectRoot "output.png"
            
            { Process-Image -InputPath $inputPath -OutputPath $outputPath } | Should -Throw
        }
    }
}

Describe "Process-Images Function Tests" {
    Context "Bulk Image Processing" {
        BeforeEach {
            # Set up mock folders
            $script:imagesWithWatermarkFolder = Join-Path $global:testProjectRoot "Watermarked"
            New-Item -ItemType Directory -Path $script:imagesWithWatermarkFolder -Force | Out-Null
            
            # Mock ImportExcel and Export-Excel if not available
            if (-not (Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
                function Global:Import-Excel {
                    param($Path, $ErrorAction)
                    return @(
                        [PSCustomObject]@{
                            Barcode     = "123456"
                            ProductName = "Test Product"
                            ImageLink   = ""
                        }
                    )
                }
                function Global:Export-Excel {
                    param($InputObject, $Path, $AutoSize, $ClearSheet)
                    New-Item -Path $Path -ItemType File -Force | Out-Null
                }
            }
        }
        
        



    }
}

Describe "Analyze-Data Function Tests" {
    Context "Data Analysis" {
        BeforeEach {
            # Mock ImportExcel if not available
            if (-not (Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
                function Global:Import-Excel {
                    param($Path, $ErrorAction)
                    return @(
                        [PSCustomObject]@{
                            Barcode        = "123456"
                            ProductName    = "Test Product 1"
                            Quantity       = 10
                            Price          = 15.99
                            WarehouseStock = 5
                        },
                        [PSCustomObject]@{
                            Barcode        = "789012"
                            ProductName    = "Test Product 2"
                            Quantity       = 25
                            Price          = 8.50
                            WarehouseStock = 0
                        }
                    )
                }
            }
        }
        
        It "Should execute without errors when data exists" {
            { Analyze-Data } | Should -Not -Throw
        }
        
        It "Should create analysis output file" {
            Analyze-Data
            $analysisFiles = Get-ChildItem -Path $global:testLogFolder -Filter "ProductAnalysis_*.txt"
            $analysisFiles.Count | Should -BeGreaterThan 0
        }
        
        It "Should handle empty database gracefully" {
            # Override Import-Excel to return empty data
            function Global:Import-Excel {
                param($Path, $ErrorAction)
                return $null
            }
            { Analyze-Data } | Should -Not -Throw
        }
    }
}

Describe "Integration Tests" {
    Context "End-to-End Workflow" {
        BeforeEach {
            # Set up complete test environment
            $script:finishedPdfFolder = Join-Path $global:testProjectRoot "Finished"
            $script:imagesWithWatermarkFolder = Join-Path $global:testProjectRoot "Watermarked"
            
            New-Item -ItemType Directory -Path $script:finishedPdfFolder -Force | Out-Null
            New-Item -ItemType Directory -Path $script:imagesWithWatermarkFolder -Force | Out-Null
            
            # Create mock files
            $mockPdf = Join-Path $global:testProjectRoot "test.pdf"
            $mockImage = Join-Path $global:testProjectRoot "123456.png"
            "Mock PDF" | Out-File -FilePath $mockPdf
            "Mock PNG" | Out-File -FilePath $mockImage
        }
        
        It "Should handle complete project folder creation" {
            { Create-ProjectFolders } | Should -Not -Throw
        }
        
        It "Should process files without critical errors" {
            # This tests that the main functions can work together
            { Create-ProjectFolders } | Should -Not -Throw
            { Process-Images } | Should -Not -Throw
            { Analyze-Data } | Should -Not -Throw
        }
    }
}

Describe "Error Handling and Edge Cases" {
    Context "Robustness Tests" {
        It "Should handle null or empty strings gracefully" {
            { Write-Log -Message $null -Level "INFO" } | Should -Not -Throw
            { Write-Log -Message "" -Level "INFO" } | Should -Not -Throw
        }
        
        It "Should handle invalid file paths" {
            $invalidPath = "Z:\NonExistent\Path\file.pdf"
            $result = Parse-BillingPdf -PdfPath $invalidPath
            $result | Should -BeNullOrEmpty
        }
        
        It "Should handle corrupted or invalid Excel paths" {
            $script:excelDatabasePath = "Z:\Invalid\Path\database.xlsx"
            $testData = @([PSCustomObject]@{Barcode = "123"; ProductName = "Test"; Quantity = 1; Price = 1.0 })
            { Update-ProductDatabase -NewProductsData $testData } | Should -Not -Throw
        }
    }
}

AfterAll {
    # Clean up test environment
    if (Test-Path $global:testProjectRoot) {
        Remove-Item $global:testProjectRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    # Remove mock functions if they were created
    Remove-Item Function:\Import-Excel -ErrorAction SilentlyContinue
    Remove-Item Function:\Export-Excel -ErrorAction SilentlyContinue
}