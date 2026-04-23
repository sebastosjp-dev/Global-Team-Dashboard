$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false # 백그라운드에서 실행
$excel.DisplayAlerts = $false

$filePath = 'C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx'

try {
    # 워크북 열기
    $wb = $excel.Workbooks.Open($filePath)

    # --- 1단계: 'POC' 시트에서 데이터 읽기 ---
    $pocSheet = $wb.Sheets.Item('POC')
    if ($null -eq $pocSheet) {
        throw "시트 'POC'를 찾을 수 없습니다."
    }

    $pocRange = $pocSheet.UsedRange
    $pocRows = $pocRange.Rows.Count
    $pocCols = $pocRange.Columns.Count

    # 'POC' 시트에서 열 인덱스 찾기
    $pocHeaders = @{}
    for ($c = 1; $c -le $pocCols; $c++) {
        $headerName = $pocSheet.Cells.Item(1, $c).Value2
        if ($headerName) {
            $pocHeaders[$headerName.Trim()] = $c
        }
    }

    # 날짜 열 이름 추정 (Date가 포함된 첫 번째 열)
    # 우선순위: 정확히 'Date', 그 다음 'POC Date', 그 다음 '*Date*'
    $dateColName = $null
    if ($pocHeaders.ContainsKey('Date')) { $dateColName = 'Date' }
    elseif ($pocHeaders.ContainsKey('POC Date')) { $dateColName = 'POC Date' }
    else { $dateColName = $pocHeaders.Keys | Where-Object { $_ -like '*Date*' } | Select-Object -First 1 }
    
    if (-not $dateColName) { throw "'POC' 시트에서 날짜 열을 찾을 수 없습니다." }
    Write-Host "날짜 열로 '$dateColName'을(를) 사용합니다."

    $dateCol = $pocHeaders[$dateColName]
    $estValueCol = $pocHeaders['Estimated Value KOR (USD)']
    $weightedValueCol = $pocHeaders['Weighted Value KOR (USD)']
    $pocStartCol = $pocHeaders['POC Start']

    if (-not ($dateCol -and $estValueCol -and $weightedValueCol)) {
        throw "필수 열('*Date*', 'Estimated Value KOR (USD)', 'Weighted Value KOR (USD)') 중 일부를 'POC' 시트에서 찾을 수 없습니다."
    }
    if (-not $pocStartCol) { Write-Host "경고: 'POC Start' 열을 찾을 수 없습니다. New POC Starts 집계가 건너뛰어집니다." }

    # --- 2단계: 월별 합계 계산 ---
    $monthlyEstimated = @{}
    $monthlyWeighted = @{}
    $monthlyNewStarts = @{}

    # 1월부터 12월까지 합계 초기화
    1..12 | ForEach-Object {
        $monthlyEstimated[$_] = 0
        $monthlyWeighted[$_] = 0
        $monthlyNewStarts[$_] = 0
    }

    for ($r = 2; $r -le $pocRows; $r++) {
        $dateValue = $pocSheet.Cells.Item($r, $dateCol).Value2
        $estValue = $pocSheet.Cells.Item($r, $estValueCol).Value2
        $weightedValue = $pocSheet.Cells.Item($r, $weightedValueCol).Value2

        if ($dateValue -is [double]) {
            try {
                $date = [datetime]::FromOADate($dateValue)
                
                # 2026년 데이터만 필터링
                if ($date.Year -eq 2026) {
                    $month = $date.Month
                    if ($estValue -is [double]) {
                        $monthlyEstimated[$month] += $estValue
                    }
                    if ($weightedValue -is [double]) {
                        $monthlyWeighted[$month] += $weightedValue
                    }
                }
            }
            catch { }
        }
        
        # New POC Starts 집계 (POC Start 열 기준)
        if ($pocStartCol) {
            $startDateValue = $pocSheet.Cells.Item($r, $pocStartCol).Value2
            if ($startDateValue -is [double]) {
                try {
                    $startDate = [datetime]::FromOADate($startDateValue)
                    if ($startDate.Year -eq 2026) {
                        $monthlyNewStarts[$startDate.Month]++
                    }
                }
                catch {}
            }
        }
    }

    # --- 3단계: 'Monthly POC Analysis (2026)' 시트 업데이트 ---
    $analysisSheet = $wb.Sheets.Item('Monthly POC Analysis (2026)')
    if ($null -eq $analysisSheet) {
        throw "시트 'Monthly POC Analysis (2026)'를 찾을 수 없습니다."
    }

    $analysisRange = $analysisSheet.UsedRange
    $analysisRows = $analysisRange.Rows.Count
    $analysisCols = $analysisRange.Columns.Count

    # 분석 시트에서 헤더와 열 인덱스 찾기
    $headerRow, $monthCol, $totalEstCol, $totalWeightedCol, $newPocStartsCol = 0, 0, 0, 0, 0
    for ($r = 1; $r -le 10; $r++) {
        # 첫 10개 행에서 헤더 검색
        for ($c = 1; $c -le $analysisCols; $c++) {
            $cellValue = $analysisSheet.Cells.Item($r, $c).Value2
            if ($cellValue -is [string]) {
                if ($cellValue.Trim() -eq 'Month') { $monthCol = $c; $headerRow = $r }
                if ($cellValue.Trim() -eq 'Total Estimated (USD)') { $totalEstCol = $c }
                if ($cellValue.Trim() -eq 'Total Weighted (USD)') { $totalWeightedCol = $c }
                if ($cellValue.Trim() -eq 'New POC Starts') { $newPocStartsCol = $c }
            }
        }
        if ($monthCol -gt 0) { break }
    }
    
    if (-not ($headerRow -gt 0 -and $monthCol -gt 0 -and $totalEstCol -gt 0 -and $totalWeightedCol -gt 0)) {
        throw "필수 열('Month', 'Total Estimated (USD)', 'Total Weighted (USD)')을 'Monthly POC Analysis (2026)' 시트에서 찾을 수 없습니다."
    }

    # --- 4단계: 계산된 합계를 분석 시트에 쓰기 ---
    $monthMap = @{"January" = 1; "Jan" = 1; "February" = 2; "Feb" = 2; "March" = 3; "Mar" = 3; "April" = 4; "Apr" = 4; "May" = 5; "June" = 6; "Jun" = 6; "July" = 7; "Jul" = 7; "August" = 8; "Aug" = 8; "September" = 9; "Sep" = 9; "October" = 10; "Oct" = 10; "November" = 11; "Nov" = 11; "December" = 12; "Dec" = 12 }

    $lastDataRow = $headerRow
    for ($r = $headerRow + 1; $r -le $analysisRows; $r++) {
        $monthName = $analysisSheet.Cells.Item($r, $monthCol).Value2
        if ($monthName -and $monthMap.ContainsKey($monthName.Trim())) {
            $monthNum = $monthMap[$monthName.Trim()]
            $analysisSheet.Cells.Item($r, $totalEstCol).Value2 = $monthlyEstimated[$monthNum]
            $analysisSheet.Cells.Item($r, $totalWeightedCol).Value2 = $monthlyWeighted[$monthNum]
            if ($newPocStartsCol -gt 0) {
                $analysisSheet.Cells.Item($r, $newPocStartsCol).Value2 = $monthlyNewStarts[$monthNum]
            }
            $lastDataRow = $r
        }
    }

    # --- 5단계: 막대 그래프 차트 생성 ---
    Write-Host "차트를 업데이트합니다..."
    
    # 기존 차트 삭제 (중복 방지)
    $chartObjects = $analysisSheet.ChartObjects()
    $chartCount = $chartObjects.Count
    for ($i = $chartCount; $i -ge 1; $i--) {
        $chartObjects.Item($i).Delete()
    }

    # 데이터 범위 설정 (Month, Estimated, Weighted 열)
    # 헤더 행부터 데이터 마지막 행까지
    $rangeMonth = $analysisSheet.Range($analysisSheet.Cells.Item($headerRow, $monthCol), $analysisSheet.Cells.Item($lastDataRow, $monthCol))
    $rangeEst = $analysisSheet.Range($analysisSheet.Cells.Item($headerRow, $totalEstCol), $analysisSheet.Cells.Item($lastDataRow, $totalEstCol))
    $rangeWeighted = $analysisSheet.Range($analysisSheet.Cells.Item($headerRow, $totalWeightedCol), $analysisSheet.Cells.Item($lastDataRow, $totalWeightedCol))
    
    # 떨어진 범위를 하나로 결합
    $chartDataRange = $excel.Union($rangeMonth, $rangeEst, $rangeWeighted)

    # 차트 추가 (xlColumnClustered = 51)
    $shape = $analysisSheet.Shapes.AddChart2(201, 51)
    $chart = $shape.Chart
    $chart.SetSourceData($chartDataRange)
    $chart.ChartTitle.Text = "2026 Monthly POC Analysis (USD)"
    
    # 차트 위치 설정 (데이터 테이블 아래)
    $shape.Left = $analysisSheet.Cells.Item($headerRow, $monthCol).Left
    $shape.Top = $analysisSheet.Cells.Item($lastDataRow + 2, $monthCol).Top
    $shape.Width = 600
    $shape.Height = 350

    $wb.Close($true) # 변경사항 저장
    Write-Host "'Monthly POC Analysis (2026)' 시트가 성공적으로 업데이트되었습니다."
}
catch {
    Write-Host "오류가 발생했습니다: $_"
    if ($wb) { $wb.Close($false) } # 오류 발생 시 저장 안 함
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
}