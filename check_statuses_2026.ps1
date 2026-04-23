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
    $dateColName = $null
    if ($headers.ContainsKey('Date')) { $dateColName = 'Date' }
    elseif ($headers.ContainsKey('POC Date')) { $dateColName = 'POC Date' }
    else { $dateColName = $headers.Keys | Where-Object { $_ -like '*Date*' } | Select-Object -First 1 }
    $dateCol = $headers[$dateColName]

    if ($statusCol -and $dateCol) {
        $countsGlobal = @{}
        $counts2026 = @{}
        for ($r = 2; $r -le $rows; $r++) {
            $statusVal = $pocSheet.Cells.Item($r, $statusCol).Value2
            $dateVal = $pocSheet.Cells.Item($r, $dateCol).Value2
            
            if ($statusVal) {
                $status = $statusVal.ToString().Trim()
                $countsGlobal[$status] = ($countsGlobal[$status] || 0) + 1
                
                if ($dateVal -is [double]) {
                    $date = [datetime]::FromOADate($dateVal)
                    if ($date.Year -eq 2026) {
                        $counts2026[$status] = ($counts2026[$status] || 0) + 1
                    }
                }
            }
        }
        "Global Statuses:" | Out-Default
        $countsGlobal | ConvertTo-Json | Out-Default
        "2026 Statuses:" | Out-Default
        $counts2026 | ConvertTo-Json | Out-Default
    } else {
        Write-Host "Status or Date column not found."
    }
    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
} finally {
    $excel.Quit()
    [System.GC]::Collect()
}
