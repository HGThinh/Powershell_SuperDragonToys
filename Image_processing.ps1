# Requires -Module System.Drawing
.SYNOPSIS
    Image Processing and Watermarking PowerShell Script
.DESCRIPTION
    This script provides a function to process images by:
    - Resizing to a target square dimension (default 600x600)
    - Centering the original image 
    - Adding a watermark text
    - Preserving image quality with high-quality rendering
.PARAMETER InputPath
    Full path to the input image file that needs to be processed.
.PARAMETER OutputPath
    Full path where the processed image will be saved.
.PARAMETER WatermarkText
    Optional text to be added as a watermark on the image. Defaults to an empty string.
.PARAMETER TargetSize
    Optional target size for the square output image. Defaults to 600 pixels.
.EXAMPLE
    Process-Image -InputPath "C:\input\photo.jpg" -OutputPath "C:\output\processed_photo.jpg" -WatermarkText "© Your Name"
.NOTES
    Author      : Hoàng Gia Thịnh
    Created     : 24/11/2024
    Modified    : 16/12/2024
    Version     : 1.1
    Dependencies: System.Drawing assembly
.LINK
    https://docs.microsoft.com/en-us/dotnet/api/system.drawing
Add-Type -AssemblyName System.Drawing

function Process-Image {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InputPath,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        [string]$WatermarkText = "",
        [int]$TargetSize = 600
    )

    try {
        # Load the original image
        $originalImage = [System.Drawing.Image]::FromFile($InputPath)
        $originalWidth = $originalImage.Width
        $originalHeight = $originalImage.Height

        # Create new bitmap for the final image (600x600)
        $finalImage = New-Object System.Drawing.Bitmap($TargetSize, $TargetSize)
        $graphics = [System.Drawing.Graphics]::FromImage($finalImage)
        
        # Set high quality rendering
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        
        # Fill background with white
        $graphics.Clear([System.Drawing.Color]::White)

        # Calculate scaling and position for centered image
        $ratio = [Math]::Min($TargetSize / $originalWidth, $TargetSize / $originalHeight)
        $newWidth = [int]($originalWidth * $ratio)
        $newHeight = [int]($originalHeight * $ratio)
        
        # Calculate position to center the image
        $x = ($TargetSize - $newWidth) / 2
        $y = ($TargetSize - $newHeight) / 2

        # Draw the original image centered
        $graphics.DrawImage($originalImage, $x, $y, $newWidth, $newHeight)

        # Add watermark
        $font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold)
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(128, 0, 0, 0))
        
        # Measure watermark text to position it
        $textSize = $graphics.MeasureString($WatermarkText, $font)
        $watermarkX = 20
        $watermarkY = $TargetSize - $textSize.Height - 20
        
        # Draw watermark
        $graphics.DrawString($WatermarkText, $font, $brush, $watermarkX, $watermarkY)

        # Save the final image
        $finalImage.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    }
    finally {
        # Clean up resources
        if ($graphics) { $graphics.Dispose() }
        if ($finalImage) { $finalImage.Dispose() }
        if ($originalImage) { $originalImage.Dispose() }
        if ($font) { $font.Dispose() }
        if ($brush) { $brush.Dispose() }
    }
}

# Example usage:
# Process-Image -InputPath "C:\input\photo.jpg" -OutputPath "C:\output\processed_photo.jpg" -WatermarkText "Your Name"
