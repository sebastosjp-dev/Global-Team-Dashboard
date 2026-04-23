$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
    $sheet = $wb.Sheets.Item('POC')
    $rows = $sheet.UsedRange.Rows.Count
    $cols = $sheet.UsedRange.Columns.Count

    $headers = @{}
    for ($c = 1; $c -le $cols; $c++) {
        $h = $sheet.Cells.Item(1, $c).Value2
        if ($h) { $headers[$h.Trim()] = $c }
    }

    $statusCol = $headers['Current Status']
    $notesCol = $headers['POC Notes']
    $endCol = $headers['POC End']

    $oldWon = 0
    $newWon = 0

    $now = [datetime]"2026-03-24"

    for ($r = 2; $r -le $rows; $r++) {
        $status = [string]$sheet.Cells.Item($r, $statusCol).Value2
        $notes = [string]$sheet.Cells.Item($r, $notesCol).Value2
        $endVal = $sheet.Cells.Item($r, $endCol).Value2

        # Old Logic
        if ($status -like "*won*" -or $status -like "*complete*" -or $status -like "*success*") {
            $oldWon++
        }

        # New Logic
        $isNoteWon = $notes -like "*won*"
        $isEndDateOk = $false
        if ($endVal -is [double]) {
            $endDate = [datetime]::FromOADate($endVal)
            if ($endDate.Year -eq 2026 -and $endDate -le $now) {
                $isEndDateOk = $true
            }
        }

        if ($isNoteWon -and $isEndDateOk) {
            $newWon++
        }
    }

    Write-Host "Old Won Count: $oldWon"
    Write-Host "New Won Count: $newWon"
    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
