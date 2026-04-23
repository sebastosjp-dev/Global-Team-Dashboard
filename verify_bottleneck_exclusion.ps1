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

    $nameCol = $headers['CRM POC Name']
    $statusCol = $headers['Current Status']
    $notesCol = $headers['POC Notes']
    $wdCol = $headers['Working Days']
    $endCol = $headers['POC End']

    $now = [datetime]"2026-03-24"
    $analysisYear = 2026

    $runningListCount = 0
    $stalledListCount = 0

    for ($r = 2; $r -le $rows; $r++) {
        $name = [string]$sheet.Cells.Item($r, $nameCol).Value2
        $curStatus = ([string]$sheet.Cells.Item($r, $statusCol).Value2).ToLower()
        $notesStr = ([string]$sheet.Cells.Item($r, $notesCol).Value2).ToLower()
        $runningDays = [double]$sheet.Cells.Item($r, $wdCol).Value2
        $endVal = $sheet.Cells.Item($r, $endCol).Value2

        # isActuallyWon logic
        $isActuallyWon = $false
        if ($notesStr -like "*won*") {
            if ($endVal -is [double]) {
                $endDate = [datetime]::FromOADate($endVal)
                if ($endDate.Year -eq $analysisYear -and $endDate -le $now) {
                    $isActuallyWon = $true
                }
            }
        }

        # isWon logic from updated services.js
        $isWon = ($curStatus -like "*won*") -or ($curStatus -like "*complete*") -or $isActuallyWon

        if (-not $isWon -and (($curStatus -like "*running*") -or ($curStatus -like "*progress*") -or ($runningDays -ge 100))) {
            $runningListCount++
            if ($runningDays -ge 100) { $stalledListCount++ }
            # Write-Host "Included: $name (Days: $runningDays, Status: $curStatus)"
        } else {
            if ($name -like "*SimasJiwa*") {
                Write-Host "Excluded as expected: $name (IsWon: $isWon, Status: $curStatus, Days: $runningDays)"
            }
        }
    }

    Write-Host "Total Running List Count: $runningListCount"
    Write-Host "Total Stalled List Count: $stalledListCount"

    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
