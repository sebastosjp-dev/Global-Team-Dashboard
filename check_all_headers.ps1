$Excel = New-Object -ComObject Excel.Application
try {
    $Workbook = $Excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
    foreach ($Sheet in $Workbook.Worksheets) {
        Write-Host "File1 Sheet: $($Sheet.Name)"
        $range = $Sheet.UsedRange
        $firstRow = $range.Rows.Item(1)
        $headers = @()
        if ($range.Rows.Count -ge 1) {
            for ($i = 1; $i -le $range.Columns.Count; $i++) {
                $val = $range.Cells.Item(1, $i).Value2
                if ($val -ne $null) { $headers += $val }
            }
        }
        Write-Host "Headers: $($headers -join ', ')"
    }
    $Workbook.Close($false)

    $Workbook2 = $Excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\Global MRR ARR.xlsx')
    foreach ($Sheet in $Workbook2.Worksheets) {
        Write-Host "File2 Sheet: $($Sheet.Name)"
        $range = $Sheet.UsedRange
        $headers = @()
        if ($range.Rows.Count -ge 1) {
            for ($i = 1; $i -le $range.Columns.Count; $i++) {
                $val = $range.Cells.Item(1, $i).Value2
                if ($val -ne $null) { $headers += $val }
            }
        }
        Write-Host "Headers: $($headers -join ', ')"
    }
    $Workbook2.Close($false)
} finally {
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}
