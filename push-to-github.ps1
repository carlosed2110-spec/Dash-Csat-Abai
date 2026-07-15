# Extrae los datos del Excel de CSAT y los sube a GitHub, para que el
# dashboard publicado en GitHub Pages se actualice solo (sin depender de
# OneDrive ni de un servidor local).
#
# Requiere:
#   - Microsoft Excel de escritorio instalado (para leer el archivo)
#   - Git para Windows instalado (https://git-scm.com/download/win)
#   - Un Personal Access Token de GitHub ya configurado en este equipo
#     (ver instrucciones que te compartieron junto con este script)

# ============================================================================
# CONFIGURACION - cambia esto por los datos de TU repositorio
# ============================================================================
$repoUrl = "https://github.com/TU-USUARIO/TU-REPO.git"
$repoFolder = Join-Path $PSScriptRoot "github-repo"
$dataJsonRelativePath = "docs\data.json"
# ============================================================================

$root = $PSScriptRoot

function Find-ExcelFile {
    param([string]$folder)
    $match = Get-ChildItem -Path $folder -Filter "*recupero*CSAT*.xlsx" -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notlike '~$*' } |
             Select-Object -First 1
    if ($match) { return $match.FullName }
    return $null
}

function ConvertExcelSerialToISO {
    param($value)
    if ($null -eq $value) { return $null }
    if ($value -is [double] -or $value -is [int]) {
        $ms = [math]::Round(($value - 25569.0) * 86400.0 * 1000.0)
        try {
            $dt = [DateTimeOffset]::FromUnixTimeMilliseconds($ms).UtcDateTime
            return $dt.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        } catch { return $null }
    }
    return $value
}

function Get-HeaderInfo {
    param($ws, [string[]]$mustContain)
    $used = $ws.UsedRange
    $rows = $used.Rows.Count
    $cols = $used.Columns.Count
    if ($rows -lt 1 -or $cols -lt 1) { return $null }
    $scanRows = [Math]::Min(3, $rows)
    $headerVals = $ws.Range($ws.Cells(1,1), $ws.Cells($scanRows, $cols)).Value2
    for ($r = 1; $r -le $scanRows; $r++) {
        $rowHeaders = @{}
        for ($c = 1; $c -le $cols; $c++) {
            $v = if ($scanRows -eq 1) { $headerVals[1,$c] } else { $headerVals[$r,$c] }
            if ($v) { $rowHeaders[[string]$v] = $c }
        }
        $hasAll = $true
        foreach ($m in $mustContain) { if (-not $rowHeaders.ContainsKey($m)) { $hasAll = $false; break } }
        if ($hasAll) {
            return @{ HeaderRow = $r; Cols = $rowHeaders; DataRows = ($rows - $r); TotalRows = $rows; TotalCols = $cols }
        }
    }
    return $null
}

$DETAIL_HEADERS = @{
    'Master.Nombre Completo'       = 'ag'
    'ID de empleado'               = 'id'
    'Master.Jefe Directo'          = 'jefe'
    'Master.Programa'              = 'prog'
    'Datos Para Matrix.Tipo'       = 'tipo'
    'Día'                          = 'dia'
    'Mes'                          = 'mes'
    'Hora'                         = 'hora'
    'Hora de inicio'               = 'fecha'
    'Duración de la llamada'       = 'dur'
    'Motivo de la llamada'         = 'motivo'
    'Comentarios de llamada'       = 'coment'
    'Cliente'                      = 'cliente'
    'Resultado comercial'          = 'resultado'
    'Cola de habilidades'          = 'cola'
    'Master.Estado Laboral'        = 'estado'
    'Número de llamada'            = 'num'
}

function Export-DataJson {
    param([string]$outputPath)

    $excelPath = Find-ExcelFile -folder $root
    if (-not $excelPath) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - No se encontro el archivo Excel."
        return $false
    }

    $excel = $null
    $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.EnableEvents = $false

        $wb = $excel.Workbooks.Open($excelPath, 0, $true)  # solo lectura

        $detailInfo = $null; $detailWs = $null
        $masterInfo = $null; $masterWs = $null
        foreach ($ws in $wb.Worksheets) {
            $info = Get-HeaderInfo -ws $ws -mustContain @('Nombre del agente (Empleado)')
            if ($info -and (-not $detailInfo -or $info.DataRows -gt $detailInfo.DataRows)) { $detailInfo = $info; $detailWs = $ws }
            $infoM = Get-HeaderInfo -ws $ws -mustContain @('Documento','Datos Matrix Empleados.ID de empleado')
            if ($infoM -and (-not $masterInfo -or $infoM.DataRows -gt $masterInfo.DataRows)) { $masterInfo = $infoM; $masterWs = $ws }
        }
        if (-not $detailInfo) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - No se encontro la hoja de detalle."
            return $false
        }

        $id2dni = @{}
        if ($masterInfo) {
            $mVals = $masterWs.Range($masterWs.Cells(1,1), $masterWs.Cells($masterInfo.TotalRows, $masterInfo.TotalCols)).Value2
            $colDoc = $masterInfo.Cols['Documento']
            $colId  = $masterInfo.Cols['Datos Matrix Empleados.ID de empleado']
            for ($r = $masterInfo.HeaderRow + 1; $r -le $masterInfo.TotalRows; $r++) {
                $idVal = $mVals[$r, $colId]
                if ($idVal) { $id2dni[[string]$idVal] = $mVals[$r, $colDoc] }
            }
        }

        $dVals = $detailWs.Range($detailWs.Cells(1,1), $detailWs.Cells($detailInfo.TotalRows, $detailInfo.TotalCols)).Value2
        $colMap = @{}
        foreach ($headerName in $DETAIL_HEADERS.Keys) {
            if ($detailInfo.Cols.ContainsKey($headerName)) { $colMap[$DETAIL_HEADERS[$headerName]] = $detailInfo.Cols[$headerName] }
        }

        $records = New-Object System.Collections.Generic.List[object]
        for ($r = $detailInfo.HeaderRow + 1; $r -le $detailInfo.TotalRows; $r++) {
            $ag = if ($colMap.ContainsKey('ag')) { $dVals[$r, $colMap['ag']] } else { $null }
            $tipo = if ($colMap.ContainsKey('tipo')) { $dVals[$r, $colMap['tipo']] } else { $null }
            if (-not $ag -and -not $tipo) { continue }
            $idemp = if ($colMap.ContainsKey('id')) { $dVals[$r, $colMap['id']] } else { $null }
            $dni = $null
            if ($idemp -and $id2dni.ContainsKey([string]$idemp)) { $dni = $id2dni[[string]$idemp] }

            $rec = [ordered]@{
                ag = $ag; id = $idemp; dni = $dni
                jefe = if ($colMap.ContainsKey('jefe')) { $dVals[$r, $colMap['jefe']] } else { $null }
                prog = if ($colMap.ContainsKey('prog')) { $dVals[$r, $colMap['prog']] } else { $null }
                tipo = $tipo
                dia = if ($colMap.ContainsKey('dia')) { $dVals[$r, $colMap['dia']] } else { $null }
                mes = if ($colMap.ContainsKey('mes')) { $dVals[$r, $colMap['mes']] } else { $null }
                hora = if ($colMap.ContainsKey('hora')) { $dVals[$r, $colMap['hora']] } else { $null }
                fecha = if ($colMap.ContainsKey('fecha')) { ConvertExcelSerialToISO $dVals[$r, $colMap['fecha']] } else { $null }
                dur = if ($colMap.ContainsKey('dur')) { $dVals[$r, $colMap['dur']] } else { $null }
                motivo = if ($colMap.ContainsKey('motivo')) { $dVals[$r, $colMap['motivo']] } else { $null }
                coment = if ($colMap.ContainsKey('coment')) { $dVals[$r, $colMap['coment']] } else { $null }
                cliente = if ($colMap.ContainsKey('cliente')) { $dVals[$r, $colMap['cliente']] } else { $null }
                resultado = if ($colMap.ContainsKey('resultado')) { $dVals[$r, $colMap['resultado']] } else { $null }
                cola = if ($colMap.ContainsKey('cola')) { $dVals[$r, $colMap['cola']] } else { $null }
                estado = if ($colMap.ContainsKey('estado')) { $dVals[$r, $colMap['estado']] } else { $null }
                num = if ($colMap.ContainsKey('num')) { $dVals[$r, $colMap['num']] } else { $null }
            }
            $records.Add($rec)
        }

        $seen = @{}
        $deduped = New-Object System.Collections.Generic.List[object]
        foreach ($rec in $records) {
            $key = "$($rec.fecha)|$($rec.num)|$($rec.dni)"
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            $deduped.Add($rec)
        }

        Write-Host "$(Get-Date -Format 'HH:mm:ss') - $($records.Count) registros leidos, $($deduped.Count) tras quitar duplicados."

        $json = $deduped | ConvertTo-Json -Depth 4 -Compress
        if ($deduped.Count -eq 1) { $json = "[$json]" }

        $outDir = Split-Path -Parent $outputPath
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
        [System.IO.File]::WriteAllText($outputPath, $json, [System.Text.Encoding]::UTF8)
        return $true
    } catch {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Error exportando datos: $($_.Exception.Message)"
        return $false
    } finally {
        if ($wb) {
            try { $wb.Close($false) } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        }
        if ($excel) {
            try { $excel.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
}

function Invoke-PushToGitHub {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Git no esta instalado. Instalalo desde https://git-scm.com/download/win"
        return $false
    }

    if (-not (Test-Path $repoFolder)) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Clonando el repositorio por primera vez..."
        git clone $repoUrl $repoFolder 2>&1 | ForEach-Object { Write-Host "  $_" }
        if (-not (Test-Path $repoFolder)) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - No se pudo clonar el repositorio. Revisa la URL y tus credenciales."
            return $false
        }
    }

    $dataJsonPath = Join-Path $repoFolder $dataJsonRelativePath
    $ok = Export-DataJson -outputPath $dataJsonPath
    if (-not $ok) { return $false }

    Push-Location $repoFolder
    try {
        git add $dataJsonRelativePath 2>&1 | Out-Null
        $status = git status --porcelain
        if (-not $status) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Sin cambios nuevos, no se sube nada."
            return $true
        }
        git commit -m "Actualizacion de datos $(Get-Date -Format 'yyyy-MM-dd HH:mm')" 2>&1 | ForEach-Object { Write-Host "  $_" }
        git push 2>&1 | ForEach-Object { Write-Host "  $_" }
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Datos subidos a GitHub. El sitio se actualizara en 1-2 minutos."
        return $true
    } catch {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Error subiendo a GitHub: $($_.Exception.Message)"
        return $false
    } finally {
        Pop-Location
    }
}

# Si se ejecuta este archivo directamente, sube los datos una vez.
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-PushToGitHub | Out-Null
}
