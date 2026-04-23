$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
    $sheet = $wb.Sheets.Item("PIPELINE")
    Write-Host "--- Sheet: PIPELINE ---"
    $rowData = @()
    for ($col = 1; $col -le 50; $col++) {
        $cell = $sheet.Cells.Item(1, $col)
        $val = $cell.Value2
        if ($val -eq $null) { $val = "" }
        $rowData += $val
    }
    Write-Host ("Headers: " + ($rowData -join ' | '))
    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
