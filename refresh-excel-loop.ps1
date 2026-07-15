# Actualizador automatico de consultas (Power Query) del Excel de CSAT
# No requiere abrir Excel manualmente: usa automatizacion COM en segundo plano.
# Requiere tener Microsoft Excel de escritorio instalado en este equipo.

$root = $PSScriptRoot
$intervalMinutes = 30
$maxRefreshSeconds = 180   # si Excel tarda mas de esto, se considera colgado y se mata el proceso

$bakeScript = Join-Path $root "bake-snapshot.ps1"
if (Test-Path $bakeScript) { . $bakeScript }  # carga Invoke-BakeSnapshot sin ejecutarla todavia

$pushScript = Join-Path $root "push-to-github.ps1"
if (Test-Path $pushScript) { . $pushScript }  # carga Invoke-PushToGitHub sin ejecutarla todavia

function Find-ExcelFile {
    param([string]$folder)
    # Busca por patron en vez de nombre exacto, para evitar problemas de
    # codificacion con tildes (o si el nombre cambia ligeramente).
    $match = Get-ChildItem -Path $folder -Filter "*recupero*CSAT*.xlsx" -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notlike '~$*' } |
             Select-Object -First 1
    if ($match) { return $match.FullName }
    return $null
}

function Refresh-ExcelQueries {
    param([string]$path, [int]$timeoutSeconds)

    if (-not $path -or -not (Test-Path $path)) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - No se encontro el archivo Excel en esta carpeta."
        return
    }

    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Actualizando consultas de Excel..."

    # Procesos EXCEL.EXE existentes antes de empezar, para no tocarlos si
    # hay que limpiar despues de un timeout (solo matamos el que abrimos aqui).
    $pidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

    $job = Start-Job -ScriptBlock {
        param($path)
        $excel = $null
        $wb = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel.AskToUpdateLinks = $false
            $excel.EnableEvents = $false

            $wb = $excel.Workbooks.Open($path, 0, $false)
            $wb.RefreshAll()

            foreach ($sheet in $wb.Worksheets) {
                foreach ($pt in $sheet.PivotTables()) {
                    try { $pt.RefreshTable() | Out-Null } catch {}
                }
            }

            try { $excel.CalculateUntilAsyncQueriesDone() } catch {}
            Start-Sleep -Seconds 8
            $wb.Save()
            "OK"
        } catch {
            "ERROR: $($_.Exception.Message)"
        } finally {
            if ($wb) {
                try { $wb.Close($true) } catch {}
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
            }
            if ($excel) {
                try { $excel.Quit() } catch {}
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
        }
    } -ArgumentList $path

    $finished = Wait-Job -Job $job -Timeout $timeoutSeconds

    if ($finished) {
        $result = Receive-Job -Job $job
        Remove-Job -Job $job -Force
        if ($result -like "OK") {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Consultas actualizadas y archivo guardado correctamente."
            if (Get-Command Invoke-BakeSnapshot -ErrorAction SilentlyContinue) {
                Invoke-BakeSnapshot | Out-Null
            }
            if (Get-Command Invoke-PushToGitHub -ErrorAction SilentlyContinue) {
                Invoke-PushToGitHub | Out-Null
            }
        } else {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Error al actualizar el Excel: $result"
            Write-Host "  (Si el archivo estaba abierto manualmente, cierralo y se reintentara en el siguiente ciclo)."
        }
    } else {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - El refresh de Excel tardo mas de $timeoutSeconds segundos (parece colgado)."
        Write-Host "  Cerrando el proceso de Excel automaticamente para liberar el archivo..."
        Remove-Job -Job $job -Force

        # Mata solo los EXCEL.EXE que aparecieron DESPUES de empezar este ciclo
        # (para no cerrar un Excel que el usuario tenia abierto de antes con otro archivo).
        $pidsAfter = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $newPids = $pidsAfter | Where-Object { $pidsBefore -notcontains $_ }
        foreach ($procId in $newPids) {
            try {
                Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
                Write-Host "  Proceso Excel (PID $procId) cerrado."
            } catch {}
        }
        Write-Host "  Se reintentara en el siguiente ciclo."
    }
}

Write-Host "=== Actualizador automatico de consultas Excel (CSAT) ==="
Write-Host "Carpeta: $root"
Write-Host "Frecuencia: cada $intervalMinutes minutos"
Write-Host "Limite por intento: $maxRefreshSeconds segundos"
Write-Host "Deja esta ventana abierta. Cierrala para detener las actualizaciones."
Write-Host ""

while ($true) {
    $excelFile = Find-ExcelFile -folder $root
    Refresh-ExcelQueries -path $excelFile -timeoutSeconds $maxRefreshSeconds
    Start-Sleep -Seconds ($intervalMinutes * 60)
}
