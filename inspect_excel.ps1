param (
    [string]$FilePath = "C:\Users\acoscolin\Desktop\SEGUIMIENTO DESCUBIERTOS\SEGUIMIENTO DESCUBIERTOS TRAB-RUTA.xlsx"
)

Write-Host "--- INSPECTING EXCEL FILE ---" -ForegroundColor Cyan

if (-not (Test-Path $FilePath)) {
    Write-Error "File not found: $FilePath"
    exit 1
}

Write-Host "File exists: $FilePath" -ForegroundColor Green

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($FilePath)
    $ws = $wb.Sheets.Item(1)
    
    Write-Host "Worksheet Name: $($ws.Name)" -ForegroundColor Yellow
    
    # Check Header Row
    $expectedHeaders = @("CLIENTE", "FECHA", "ESTADO", "TRABAJADOR", "Total H")
    $headers = @()
    $lastCol = $ws.UsedRange.Columns.Count

    for ($i = 1; $i -le $lastCol; $i++) {
        $headers += $ws.Cells.Item(1, $i).Text
    }

    foreach ($exp in $expectedHeaders) {
        if ($headers -contains $exp) {
            Write-Host "[OK] Column found: $exp" -ForegroundColor Green
        }
        else {
            Write-Host "[WARN] Missing column: $exp" -ForegroundColor Red
        }
    }

    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "Inspection Complete." -ForegroundColor Cyan
}
catch {
    Write-Error "Error opening Excel: $_"
}
