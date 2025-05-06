$systemName = $env:COMPUTERNAME

# Create Excel application object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.Name = "Installed Software"

# Set headers
$headers = @("Name", "Version", "Publisher", "Install Date")
for ($i = 0; $i -lt $headers.Count; $i++) {
    $worksheet.Cells.Item(1, $i + 1) = $headers[$i]
    $worksheet.Cells.Item(1, $i + 1).Font.Bold = $true
    $worksheet.Cells.Item(1, $i + 1).Interior.ColorIndex = 15  # Light gray background
}

# Function to get installed software from a registry path
function Get-InstalledSoftwareFromRegistry {
    param (
        [string]$registryPath
    )
    $items = @()
    $keys = Get-ChildItem -Path $registryPath -ErrorAction SilentlyContinue
    foreach ($key in $keys) {
        $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
        if ($props.DisplayName -and $props.DisplayName.Trim() -ne "") {
            $item = [PSCustomObject]@{
                Name        = $props.DisplayName
                Version     = $props.DisplayVersion
                Publisher   = $props.Publisher
                InstallDate = $props.InstallDate
            }
            $items += $item
        }
    }
    return $items
}

# Get installed software from both 64-bit and 32-bit registry locations
$softwareList = @()
$softwareList += Get-InstalledSoftwareFromRegistry -registryPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$softwareList += Get-InstalledSoftwareFromRegistry -registryPath "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
$softwareList += Get-InstalledSoftwareFromRegistry -registryPath "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

# Remove duplicates by Name and Version
$softwareList = $softwareList | Sort-Object Name, Version -Unique

# Output to Excel starting from row 2
$row = 2
foreach ($app in $softwareList) {
    $worksheet.Cells.Item($row, 1) = $app.Name
    $worksheet.Cells.Item($row, 2) = $app.Version
    $worksheet.Cells.Item($row, 3) = $app.Publisher

    # Convert InstallDate from yyyymmdd to DateTime and write to Excel
    if ($app.InstallDate -match '^\d{8}$') {
        $date = [datetime]::ParseExact($app.InstallDate, 'yyyyMMdd', $null)
        $worksheet.Cells.Item($row, 4).Value2 = $date
        $worksheet.Cells.Item($row, 4).NumberFormat = "yyyy-mm-dd"
    } elseif ($app.InstallDate) {
        # If InstallDate exists but is not in expected format, write as is
        $worksheet.Cells.Item($row, 4) = $app.InstallDate
    } else {
        # If InstallDate is missing, leave cell blank
        $worksheet.Cells.Item($row, 4) = ""
    }
    $row++
}

# Autofit columns for better appearance
$worksheet.Columns.AutoFit()

# Save the workbook to a file
$excelFilePath = "$env:USERPROFILE\OneDrive\Desktop\InstalledSoftware_$systemName.xlsx"
$workbook.SaveAs($excelFilePath)

# Cleanup
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Output "Installed software list exported to $excelFilePath"
