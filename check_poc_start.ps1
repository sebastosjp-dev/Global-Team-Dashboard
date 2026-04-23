$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open("c:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx")

# List all sheet names
Write-Output "=== Sheet Names ==="
foreach ($ws in $wb.Sheets) {
    Write-Output ("  " + $ws.Name)
}

$ws = $wb.Sheets.Item("POC")
$lastRow = $ws.UsedRange.Rows.Count
$lastCol = $ws.UsedRange.Columns.Count

Write-Output ""
Write-Output "=== POC Sheet Headers ==="
$headers = @()
for ($col = 1; $col -le $lastCol; $col++) {
    $h = $ws.Cells.Item(1, $col).Text
    $headers += $h
    Write-Output ("  Col $col : $h")
}

# Find POC Start column
$pocStartCol = -1
for ($i = 0; $i -lt $headers.Count; $i++) {
    if ($headers[$i] -match "POC Start") {
        $pocStartCol = $i + 1
        break
    }
}

if ($pocStartCol -gt 0) {
    Write-Output ""
    Write-Output "=== POC Start Column Found: Col $pocStartCol ==="
    Write-Output "--- POC Start values (rows 2 to 60) ---"
    
    $maxRow = [Math]::Min($lastRow, 60)
    for ($row = 2; $row -le $maxRow; $row++) {
        $pocName = $ws.Cells.Item($row, 1).Text
        $pocStartVal = $ws.Cells.Item($row, $pocStartCol).Text
        Write-Output ("Row ${row}: [$pocName] -> POC Start=[$pocStartVal]")
    }
    
    # Count by month for 2026
    Write-Output ""
    Write-Output "=== Monthly POC Start Count (2026) ==="
    $monthlyCounts = @{}
    for ($m = 1; $m -le 12; $m++) { $monthlyCounts[$m] = 0 }
    
    for ($row = 2; $row -le $lastRow; $row++) {
        $val = $ws.Cells.Item($row, $pocStartCol).Value2
        if ($val -and $val -is [double] -and $val -gt 30000) {
            $dt = [DateTime]::FromOADate($val)
            if ($dt.Year -eq 2026) {
                $monthlyCounts[$dt.Month]++
            }
        }
    }
    
    $monthNames = @("", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    for ($m = 1; $m -le 12; $m++) {
        Write-Output ("  $($monthNames[$m]): $($monthlyCounts[$m])")
    }
} else {
    Write-Output "POC Start column NOT found in headers."
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
