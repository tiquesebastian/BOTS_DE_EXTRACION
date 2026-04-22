$f = "c:/Users/ticdesarrollo09/Music/trabajo/bot/890701715_documentos.csv"
$totalObjetivo = 1309
$lastCount = -1
$start = Get-Date

while ($true) {
    if (Test-Path $f) {
        $rows = Import-Csv $f
        $count = $rows.Count

        if ($count -ne $lastCount) {
            $conDocumento = ($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.numero_documento) }).Count
            $sinDocumento = $count - $conDocumento
            $nitHospital = ($rows | Where-Object { $_.numero_documento -eq '890701715' }).Count

            $elapsed = [math]::Round(((Get-Date) - $start).TotalSeconds, 0)
            $velocidad = if ($elapsed -gt 0) { [math]::Round($count / $elapsed, 2) } else { 0 }
            $pct = if ($totalObjetivo -gt 0) { [math]::Round(($count * 100.0) / $totalObjetivo, 2) } else { 0 }

            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Procesados: $count/$totalObjetivo ($pct%) | Con doc: $conDocumento | Sin doc: $sinDocumento | NIT hospital: $nitHospital | Vel: $velocidad rad/seg | Tiempo: ${elapsed}s" -ForegroundColor Cyan
            $lastCount = $count
        }
    } else {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Esperando archivo de salida: $f" -ForegroundColor Yellow
    }

    Start-Sleep -Seconds 3
}
