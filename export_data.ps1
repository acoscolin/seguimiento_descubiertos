param (
    [string]$FilePath = "C:\Users\acoscolin\Desktop\SEGUIMIENTO DESCUBIERTOS\SEGUIMIENTO DESCUBIERTOS TRAB-RUTA.xlsx",
    [string]$OutputPath = "C:\Users\acoscolin\Desktop\SEGUIMIENTO DESCUBIERTOS\data.json"
)

Write-Host "--- STARTING DATA EXPORT ---" -ForegroundColor Cyan

if (-not (Test-Path $FilePath)) {
    Write-Error "File not found: $FilePath"
    exit 1
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($FilePath)
    $ws = $wb.Sheets.Item(1)

    $lastRow = $ws.UsedRange.Rows.Count
    $lastCol = $ws.UsedRange.Columns.Count
    
    # Map headers to indices
    $headerMap = @{}
    for ($c = 1; $c -le $lastCol; $c++) {
        $val = $ws.Cells.Item(1, $c).Text.Trim()
        $headerMap[$val] = $c
    }

    $data = @()
    $currentYear = (Get-Date).Year

    # Required columns
    if (-not ($headerMap.ContainsKey("CLIENTE") -and $headerMap.ContainsKey("FECHA"))) {
        throw "Missing critical columns (CLIENTE, FECHA). Headers found: $($headerMap.Keys)"
    }

    for ($r = 2; $r -le $lastRow; $r++) {
        $cliente = $ws.Cells.Item($r, $headerMap["CLIENTE"]).Text.Trim()
        
        # Skip empty rows
        if ([string]::IsNullOrWhiteSpace($cliente)) { continue }

        $fechaRaw = $ws.Cells.Item($r, $headerMap["FECHA"]).Text.Trim()
        $estado = $ws.Cells.Item($r, $headerMap["ESTADO"]).Text.Trim()
        $trabajador = $ws.Cells.Item($r, $headerMap["TRABAJADOR"]).Text.Trim()
        $tipo = $ws.Cells.Item($r, $headerMap["TIPO TRAB"]).Text.Trim()
        
        # Handle Total H (e.g. "3,0" -> 3.0)
        $hoursRaw = $ws.Cells.Item($r, $headerMap["Total H"]).Text.Trim()
        $hours = 0.0
        if ($hoursRaw) {
            $hoursRaw = $hoursRaw -replace ',', '.'
            try {
                $hours = [float]$hoursRaw
            }
            catch {
                Write-Warning "Row $($r): Could not parse hours '$hoursRaw'"
            }
        }

        # Handle Date (e.g. "01-dic" -> "2025-12-01")
        # Spanish mapping for months if needed, or rely on system locale if Excel parsed it.
        # If Excel shows "01-dic", .Text returns "01-dic". We need to parse.
        $dateIso = ""
        if ($fechaRaw) {
            try {
                # Simple parsing assuming format dd-MMM
                # Dictionary for Spanish months
                $months = @{
                    "ene" = "01"; "feb" = "02"; "mar" = "03"; "abr" = "04"; "may" = "05"; "jun" = "06";
                    "jul" = "07"; "ago" = "08"; "sep" = "09"; "oct" = "10"; "nov" = "11"; "dic" = "12"
                }
                
                $parts = $fechaRaw -split "-"
                if ($parts.Count -eq 2) {
                    $day = $parts[0].PadLeft(2, '0')
                    $monStr = $parts[1].ToLower().Substring(0, 3)
                    if ($months.ContainsKey($monStr)) {
                        $month = $months[$monStr]
                        $dateIso = "$currentYear-$month-$day"
                    }
                }
            }
            catch {
                Write-Warning "Row $($r): Could not parse date '$fechaRaw'"
            }
        }

        $record = @{
            info = @{
                cliente    = $cliente
                fecha      = $dateIso
                estado     = $estado
                trabajador = $trabajador
                tipo       = $tipo
                hours      = $hours
                horIn      = $ws.Cells.Item($r, $headerMap["Hor In"]).Text.Trim()
                horFin     = $ws.Cells.Item($r, $headerMap["Hor Fin"]).Text.Trim()
            }
        }
        $data += $record.info
    }

    $json = $data | ConvertTo-Json -Depth 3 -Compress
    # Force UTF8 NoBOM
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    
    # Write JSON
    [System.IO.File]::WriteAllText($OutputPath, $json, $utf8NoBom)
    
    # Export JS
    $jsPath = $OutputPath.Replace("json", "js")
    $jsContent = "window.dashboardData = $json;"
    [System.IO.File]::WriteAllText($jsPath, $jsContent, $utf8NoBom)

    Write-Host "Successfully exported $((Get-Date).ToString('HH:mm:ss')): $($data.Count) records to $OutputPath and $jsPath" -ForegroundColor Green

    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

}
catch {
    Write-Error "Fatal Error: $_"
    if ($excel) { $excel.Quit() }
}
