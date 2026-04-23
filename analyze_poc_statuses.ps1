$excel = New-Object -ComObject Excel.Application
$wb = $excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
$sheet = $wb.Sheets.Item('POC')
$usedRange = $sheet.UsedRange
$rows = $usedRange.Rows.Count
$cols = $usedRange.Columns.Count

$statusCol = -1
for ($c = 1; $c -le $cols; $c++) {
    if ($sheet.Cells.Item(1, $c).Text -like '*Current Status*') {
        $statusCol = $c
        break
    }
}

if ($statusCol -eq -1) {
    Write-Host "Current Status column not found"
} else {
    $counts = @{}
    for ($r = 2; $r -le $rows; $r++) {
        $status = $sheet.Cells.Item($r, $statusCol).Value
        if ($status -eq $null) { $status = "" }
        $status = $status.ToString().Trim()
        if ($status -ne "") {
            $counts[$status] = $counts[$status] + 1
        }
    }
    $counts | ConvertTo-Json
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
