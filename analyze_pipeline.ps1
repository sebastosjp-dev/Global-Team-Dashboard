$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
    $sheet = $wb.Sheets.Item("PIPELINE")
    
    $stats = @{
        "Q1" = @{ "all" = 0; "y2026" = 0 }
        "Q2" = @{ "all" = 0; "y2026" = 0 }
        "Q3" = @{ "all" = 0; "y2026" = 0 }
        "Q4" = @{ "all" = 0; "y2026" = 0 }
    }

    $lastRow = $sheet.UsedRange.Rows.Count
    for ($row = 2; $row -le $lastRow; $row++) {
        $tcv = $sheet.Cells.Item($row, 5).Value2 # KOR TCV (USD) is 5th col
        $qValue = [string]$sheet.Cells.Item($row, 8).Value2 # Quarter is 8th col
        $closeDate = $sheet.Cells.Item($row, 7).Value2 # Close Date is 7th col
        
        if ($tcv -eq $null) { $tcv = 0 }
        
        $q = ""
        if ($qValue -like "*Q1*") { $q = "Q1" }
        elseif ($qValue -like "*Q2*") { $q = "Q2" }
        elseif ($qValue -like "*Q3*") { $q = "Q3" }
        elseif ($qValue -like "*Q4*") { $q = "Q4" }
        
        if ($q -ne "") {
            $stats[$q]["all"] += $tcv
            
            # Excel serial date for 2026 is between 46022 (2026-01-01) and 46386 (2026-12-31)
            if ($closeDate -ge 46022 -and $closeDate -le 46386) {
                $stats[$q]["y2026"] += $tcv
            }
        }
    }
    
    $stats | ConvertTo-Json
    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
