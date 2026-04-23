$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
    $sheet = $wb.Sheets.Item("PIPELINE")
    
    $stats = @{
        "Q1" = 0
        "Q2" = 0
        "Q3" = 0
        "Q4" = 0
    }

    $lastRow = $sheet.UsedRange.Rows.Count
    for ($row = 2; $row -le $lastRow; $row++) {
        $tcv = $sheet.Cells.Item($row, 5).Value2 # KOR TCV (USD)
        $qValue = [string]$sheet.Cells.Item($row, 8).Value2 # Quarter
        $closeDate = $sheet.Cells.Item($row, 7).Value2 # Close Date
        
        if ($tcv -eq $null) { $tcv = 0 }
        
        $q = ""
        if ($qValue -like "*Q1*") { $q = "Q1" }
        elseif ($qValue -like "*Q2*") { $q = "Q2" }
        elseif ($qValue -like "*Q3*") { $q = "Q3" }
        elseif ($qValue -like "*Q4*") { $q = "Q4" }
        
        if ($q -ne "" -and $closeDate -ge 46022 -and $closeDate -le 46386) { # 2026 range
            $stats[$q] += $tcv
        }
    }
    
    Write-Host "Expected TCV for 2026:"
    $stats | ConvertTo-Json
    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
