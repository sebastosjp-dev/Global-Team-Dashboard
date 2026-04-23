$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
$sheet = $wb.Sheets.Item('POC')

$headers = @{}
for ($c = 1; $c -le 20; $c++) {
    $h = $sheet.Cells.Item(1, $c).Value2
    if ($h) { $headers[$h.Trim()] = $c }
}

$statusCol = $headers['Current Status']
$notesCol = $headers['POC Notes']
$endCol = $headers['POC End']

Write-Host "Sample Rows (Current Status = Complete or Won):"
for ($r = 2; $r -le 50; $r++) {
    $status = [string]$sheet.Cells.Item($r, $statusCol).Value2
    $notes = [string]$sheet.Cells.Item($r, $notesCol).Value2
    $endVal = $sheet.Cells.Item($r, $endCol).Value2
    $endDateStr = "Empty"
    if ($endVal -is [double]) { $endDateStr = ([datetime]::FromOADate($endVal)).ToString("yyyy-MM-dd") }

    if ($status -like "*won*" -or $status -like "*complete*" -or $status -like "*success*") {
        Write-Host "Row $r | Status: $status | Notes: $notes | End Date: $endDateStr"
    }
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
