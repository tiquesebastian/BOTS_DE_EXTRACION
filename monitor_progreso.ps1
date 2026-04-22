$f = "c:/Users/ticdesarrollo09/Music/trabajo/bot/890701715_fechas_ingreso.csv"
$lastCount = 0
$start = Get-Date

while ($true) {
    if (Test-Path $f) {
        $count = (Get-Content $f | Measure-Object -Line).Lines - 1
        if ($count -ne $lastCount) {
            $con_fecha = (Get-Content $f | Where-Object { $_ -match '^\d+,\d{1,2}/\d{1,2}/\d{2}$' }).Count
            $sin_fecha = $count - $con_fecha
            $elapsed = [math]::Round(((Get-Date) - $start).TotalSeconds, 0)
            $velocidad = if ($elapsed -gt 0) { [math]::Round($count / $elapsed, 2) } else { 0 }
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Procesados: $count/811 | Con fecha: $con_fecha | Sin fecha: $sin_fecha | Vel: $velocidad rad/seg | Tiempo: ${elapsed}s" -ForegroundColor Cyan
            $lastCount = $count
        }
    }
    Start-Sleep -Seconds 3
}
