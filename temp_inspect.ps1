$file = "C:\Users\acoscolin\Desktop\SEGUIMIENTO DESCUBIERTOS\SEGUIMIENTO DESCUBIERTOS TRAB-RUTA.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open($file)
$ws = $wb.Sheets.Item(1)

Write-Host "--- SHEET NAME: $($ws.Name) ---"

# Get Headers (First Row)
$lastCol = $ws.UsedRange.Columns.Count
$headers = @()
for ($i = 1; $i -le $lastCol; $i++) {
    $headers += $ws.Cells.Item(1, $i).Value2
}
Write-Host "--- HEADERS ---"
$headers -join ", "

# Get 5 rows of data
Write-Host "--- SAMPLE DATA (First 5 rows) ---"
for ($r = 2; $r -le 6; $r++) {
    $rowValues = @()
    for ($c = 1; $c -le $lastCol; $c++) {
        $val = $ws.Cells.Item($r, $c).Text
        $rowValues += "$val"
    }
    Write-Host ($rowValues -join " | ")
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
