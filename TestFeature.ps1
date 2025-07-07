Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variable to hold imported file path
$global:importFilePath = ""

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "CSV Export/Import Tool"
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = "CenterScreen"

# Test button (was Export)
$testButton = New-Object System.Windows.Forms.Button
$testButton.Location = New-Object System.Drawing.Point(50, 50)
$testButton.Size = New-Object System.Drawing.Size(120, 40)
$testButton.Text = "Test"
$form.Controls.Add($testButton)

# Import button
$importButton = New-Object System.Windows.Forms.Button
$importButton.Location = New-Object System.Drawing.Point(200, 50)
$importButton.Size = New-Object System.Drawing.Size(120, 40)
$importButton.Text = "Import from CSV"
$form.Controls.Add($importButton)

# Test button event: export dummy data
$testButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Dummy Product Data"
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $products = @(
            [PSCustomObject]@{ Name = "Toy Car"; Barcode = "111"; Category = "Toys"; Quantity = 10 }
            [PSCustomObject]@{ Name = "Puzzle Box"; Barcode = "112"; Category = "Games"; Quantity = 5 }
            [PSCustomObject]@{ Name = "Lego Set"; Barcode = "113"; Category = "Building"; Quantity = 8 }
        )
        $products | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show("Dummy product data exported.", "Test", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Import button event
$importButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $openFileDialog.Title = "Open CSV File"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $global:importFilePath = $openFileDialog.FileName
        $importedData = Import-Csv -Path $global:importFilePath | ForEach-Object { $_ }

        if (-not $importedData) {
            [System.Windows.Forms.MessageBox]::Show("CSV file is empty or unreadable.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Create new form to show imported data
        $dataForm = New-Object System.Windows.Forms.Form
        $dataForm.Text = "Imported Products"
        $dataForm.Size = New-Object System.Drawing.Size(700, 400)

        $dataGrid = New-Object System.Windows.Forms.DataGridView
        $dataGrid.Dock = "Top"
        $dataGrid.Height = 300
        $dataGrid.ReadOnly = $false
        $dataGrid.AutoGenerateColumns = $true
        $dataGrid.DataSource = $importedData
        $dataGrid.Refresh()
        $dataForm.Controls.Add($dataGrid)

        # Create Export button inside imported view
        $innerExportButton = New-Object System.Windows.Forms.Button
        $innerExportButton.Text = "Export Updated"
        $innerExportButton.Width = 120
        $innerExportButton.Height = 40
        $innerExportButton.Top = 310
        $innerExportButton.Left = 270
        $dataForm.Controls.Add($innerExportButton)

        # Export event
        $innerExportButton.Add_Click({
            $newData = $dataGrid.DataSource

            # Path to result file
            $resultPath = [System.IO.Path]::Combine((Split-Path $global:importFilePath), "Result.csv")

            $existingData = @()
            if (Test-Path $resultPath) {
                $existingData = Import-Csv $resultPath | ForEach-Object { $_ }
            }

            $combined = @{}
            $deletedCount = 0
            $newProductCount = 0
            $updatedProductCount = 0

            # Store existing data by barcode
            foreach ($item in $existingData) {
                $combined[$item.Barcode] = $item
            }

            # Process new data
            foreach ($item in $newData) {
                $barcode = $item.Barcode
                $qty = [int]$item.Quantity

                if ($combined.ContainsKey($barcode)) {
                    $existingQty = [int]$combined[$barcode].Quantity
                    $newQty = $existingQty + $qty

                    if ($newQty -lt 0) {
                        $combined.Remove($barcode)
                        $deletedCount++
                    } else {
                        $combined[$barcode].Quantity = $newQty
                        $updatedProductCount++
                    }
                } else {
                    if ($qty -ge 0) {
                        $combined[$barcode] = $item
                        $newProductCount++
                    }
                }
            }

            $combined.Values | Export-Csv -Path $resultPath -NoTypeInformation

            $message = "Exported to Result.csv.`nNew products added: $newProductCount`nUpdated products: $updatedProductCount`nDeleted products: $deletedCount"
            [System.Windows.Forms.MessageBox]::Show($message, "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

            $dataForm.Close() # Close the import screen
        })

        $dataForm.ShowDialog()
    }
})

# Show the main form
[void]$form.ShowDialog()
