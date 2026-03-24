$excel = New-Object -ComObject Excel.Application
try {
    $wb = $excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
    foreach ($sheet in $wb.Sheets) {
        Write-Host "File1: $($sheet.Name)"
    }
    $wb.Close($false)
} catch {
    Write-Host "Error with File 1: $_"
}

try {
    $wb2 = $excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\Global MRR ARR.xlsx')
    foreach ($sheet in $wb2.Sheets) {
        Write-Host "File2: $($sheet.Name)"
    }
    $wb2.Close($false)
} catch {
    Write-Host "Error with File 2: $_"
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
