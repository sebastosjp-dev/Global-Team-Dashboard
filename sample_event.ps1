$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
    $sheet = $wb.Sheets.Item('EVENT')
    $range = $sheet.UsedRange
    $rows = $range.Rows.Count
    $cols = $range.Columns.Count

    Write-Host "Headers (Row 1):"
    for ($c = 1; $c -le $cols; $c++) {
        Write-Host "$($sheet.Cells.Item(1, $c).Value2)"
    }
    
    Write-Host "`nSample Data (Row 2):"
     for ($c = 1; $c -le $cols; $c++) {
        Write-Host "$($sheet.Cells.Item(2, $c).Value2)"
    }

    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
