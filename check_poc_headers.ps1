$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
    
    $sheets = @("POC", "Sheet9")
    foreach ($sheetName in $sheets) {
        try {
            $sheet = $wb.Sheets.Item($sheetName)
            Write-Host "`n--- Sheet: $sheetName ---"
            for ($row = 1; $row -le 2; $row++) {
                $rowData = @()
                for ($col = 1; $col -le 35; $col++) {
                    $cell = $sheet.Cells.Item($row, $col)
                    $val = $cell.Value2
                    if ($val -eq $null) { $val = "" }
                    $rowData += $val
                }
                $line = "Row ${row}: " + ($rowData -join ' | ')
                Write-Host $line
            }
        } catch {
            Write-Host "Could not access sheet: $sheetName"
        }
    }
    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
