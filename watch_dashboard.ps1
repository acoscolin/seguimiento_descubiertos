$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = "C:\Users\acoscolin\Desktop\SEGUIMIENTO DESCUBIERTOS"
$watcher.Filter = "SEGUIMIENTO DESCUBIERTOS TRAB-RUTA.xlsx"
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

$action = {
    $path = $Event.SourceEventArgs.FullPath
    $changeType = $Event.SourceEventArgs.ChangeType
    Write-Host "File Changed: $path ($changeType)" -ForegroundColor Yellow
    
    # Debounce / Wait for lock release
    Start-Sleep -Seconds 2
    
    Write-Host "Triggering Export..." -ForegroundColor Cyan
    try {
        & "powershell" -ExecutionPolicy Bypass -File "C:\Users\acoscolin\Desktop\SEGUIMIENTO DESCUBIERTOS\export_data.ps1"
    }
    catch {
        Write-Error "Export Failed: $_"
    }
}

Register-ObjectEvent $watcher "Changed" -Action $action
Register-ObjectEvent $watcher "Created" -Action $action
Register-ObjectEvent $watcher "Renamed" -Action $action

Write-Host "--- DETECTOR DE CAMBIOS ACTIVO ---" -ForegroundColor Green
Write-Host "Monitorizando: $($watcher.Filter)"
Write-Host "Instrucciones: Guarda tu Excel y espera unos segundos."
Write-Host "Presiona Ctrl+C para detener."


# Keep alive
while ($true) { Start-Sleep -Seconds 5 }
