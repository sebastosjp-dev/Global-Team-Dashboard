$Excel = New-Object -ComObject Excel.Application
try {
    $Workbook = $Excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\2026 Global Rev.01.xlsx')
    $Workbook.Worksheets | ForEach-Object { $O = $_.Name; $O } | Out-File -FilePath sheets_file1.txt -Encoding UTF8
    $Workbook.Close($false)

    $Workbook2 = $Excel.Workbooks.Open('C:\Users\LG\Downloads\global dashboard making\Global MRR ARR.xlsx')
    $Workbook2.Worksheets | ForEach-Object { $O = $_.Name; $O } | Out-File -FilePath sheets_file2.txt -Encoding UTF8
    $Workbook2.Close($false)
} finally {
    $Excel.Quit()
}
