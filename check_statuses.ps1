$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$filePath = 'C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx'

try {
    $wb = $excel.Workbooks.Open($filePath)
    $pocSheet = $wb.Sheets.Item('POC')
    $range = $pocSheet.UsedRange
    $rows = $range.Rows.Count
    $cols = $range.Columns.Count

    $headers = @{}
    for ($c = 1; $c -le $cols; $c++) {
        $val = $pocSheet.Cells.Item(1, $c).Value2
        if ($val) {
            $headers[$val.ToString().Trim()] = $c
        }
    }

    $statusCol = $headers['Current Status']
    if (-not $statusCol) {
        $statusKey = $headers.Keys | Where-Object { $_ -like '*Status*' } | Select-Object -First 1
        $statusCol = $headers[$statusKey]
    }

    if ($statusCol) {
        $counts = @{}
        for ($r = 2; $r -le $rows; $r++) {
            $val = $pocSheet.Cells.Item($r, $statusCol).Value2
            if ($val) {
                $status = $val.ToString().Trim()
                if ($counts.ContainsKey($status)) {
                    $counts[$status]++
                } else {
                    $counts[$status] = 1
                }
            }
        }
        $counts | ConvertTo-Json
    } else {
        Write-Host "Status column not found."
    }
    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
} finally {
    $excel.Quit()
    if ($wb) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
