# Sube TODO el portal (CSAT + Gestion KPI + Evolutivo por Supervisor) a GitHub,
# para que el sitio publicado en GitHub Pages se actualice solo, sin depender
# de OneDrive ni de un servidor local corriendo en tu PC.
#
# Requiere:
#   - Microsoft Excel de escritorio instalado
#   - Git para Windows instalado (https://git-scm.com/download/win)
#   - Un Personal Access Token de GitHub ya configurado en este equipo
#
# Archivos que este script espera encontrar en la MISMA carpeta que el:
#   - cloud-portal.html
#   - cloud-dashboard-csat.html
#   - cloud-kpi-dashboard.html
#   - cloud-evolutivo-supervisor.html
#   - header-animation.mp4
#   - El Excel de CSAT (el que ya usabas)
#   - La carpeta "Improductivos" con los .xlsx/.xlsm mensuales

# ============================================================================
# CONFIGURACION - cambia esto por los datos de TU repositorio
# ============================================================================
$repoUrl = "https://github.com/TU-USUARIO/TU-REPO.git"
$repoFolder = Join-Path $PSScriptRoot "github-repo"
$excludedPlatforms = @("RETENCIONES_MOVIL")
# ============================================================================

$root = $PSScriptRoot
$monthAbbr = @('Ene.','Feb.','Mar.','Abr.','May.','Jun.','Jul.','Ago.','Sep.','Oct.','Nov.','Dic.')

# ---------------- Utilidades compartidas ----------------

function Find-CsatExcelFile {
    param([string]$folder)
    $match = Get-ChildItem -Path $folder -Filter "*recupero*CSAT*.xlsx" -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notlike '~$*' } |
             Select-Object -First 1
    if ($match) { return $match.FullName }
    return $null
}

function Find-ImproductivosFiles {
    param([string]$folder)
    $improFolder = Join-Path $folder "Improductivos"
    if (-not (Test-Path $improFolder)) { return @() }
    $files = Get-ChildItem -Path $improFolder -Include "*.xlsx","*.xlsm" -File -Recurse -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notlike '~$*' }
    return $files
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

function Escape-JsonScriptTags {
    param([string]$json)
    return ($json -replace '(?i)</script>', '<\/script>')
}

# ---------------- Exportar CSAT ----------------

$DETAIL_HEADERS_CSAT = @{
    'Master.Nombre Completo' = 'ag'
    'ID de empleado'         = 'id'
    'Master.Jefe Directo'    = 'jefe'
    'Master.Programa'        = 'prog'
    'Datos Para Matrix.Tipo' = 'tipo'
    'Día'                    = 'dia'
    'Mes'                    = 'mes'
    'Hora'                   = 'hora'
    'Hora de inicio'         = 'fecha'
    'Duración de la llamada' = 'dur'
    'Motivo de la llamada'   = 'motivo'
    'Comentarios de llamada' = 'coment'
    'Cliente'                = 'cliente'
    'Resultado comercial'    = 'resultado'
    'Cola de habilidades'    = 'cola'
    'Master.Estado Laboral'  = 'estado'
    'Número de llamada'      = 'num'
    'Colas.Skill'            = 'skill'
    'La respuesta_2'         = 'nota'
}

function Export-CsatDataJson {
    param([string]$outputPath)

    $excelPath = Find-CsatExcelFile -folder $root
    if (-not $excelPath) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - [CSAT] No se encontro el archivo Excel."
        return $false
    }

    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false; $excel.EnableEvents = $false
        $wb = $excel.Workbooks.Open($excelPath, 0, $true)

        $detailInfo = $null; $detailWs = $null
        $masterInfo = $null; $masterWs = $null
        foreach ($ws in $wb.Worksheets) {
            $info = Get-HeaderInfo -ws $ws -mustContain @('Nombre del agente (Empleado)')
            if ($info -and (-not $detailInfo -or $info.DataRows -gt $detailInfo.DataRows)) { $detailInfo = $info; $detailWs = $ws }
            $infoM = Get-HeaderInfo -ws $ws -mustContain @('Documento','Datos Matrix Empleados.ID de empleado')
            if ($infoM -and (-not $masterInfo -or $infoM.DataRows -gt $masterInfo.DataRows)) { $masterInfo = $infoM; $masterWs = $ws }
        }
        if (-not $detailInfo) { Write-Host "$(Get-Date -Format 'HH:mm:ss') - [CSAT] No se encontro la hoja de detalle."; return $false }

        $id2dni = @{}
        if ($masterInfo) {
            $mVals = $masterWs.Range($masterWs.Cells(1,1), $masterWs.Cells($masterInfo.TotalRows, $masterInfo.TotalCols)).Value2
            $colDoc = $masterInfo.Cols['Documento']; $colId = $masterInfo.Cols['Datos Matrix Empleados.ID de empleado']
            for ($r = $masterInfo.HeaderRow + 1; $r -le $masterInfo.TotalRows; $r++) {
                $idVal = $mVals[$r, $colId]
                if ($idVal) { $id2dni[[string]$idVal] = $mVals[$r, $colDoc] }
            }
        }

        $dVals = $detailWs.Range($detailWs.Cells(1,1), $detailWs.Cells($detailInfo.TotalRows, $detailInfo.TotalCols)).Value2
        $colMap = @{}
        foreach ($headerName in $DETAIL_HEADERS_CSAT.Keys) {
            if ($detailInfo.Cols.ContainsKey($headerName)) { $colMap[$DETAIL_HEADERS_CSAT[$headerName]] = $detailInfo.Cols[$headerName] }
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
                ag=$ag; id=$idemp; dni=$dni
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
                skill = if ($colMap.ContainsKey('skill')) { $dVals[$r, $colMap['skill']] } else { $null }
                nota = if ($colMap.ContainsKey('nota')) { $dVals[$r, $colMap['nota']] } else { $null }
            }
            $records.Add($rec)
        }

        $seen = @{}; $deduped = New-Object System.Collections.Generic.List[object]
        foreach ($rec in $records) {
            $key = "$($rec.fecha)|$($rec.num)|$($rec.dni)"
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            $deduped.Add($rec)
        }

        Write-Host "$(Get-Date -Format 'HH:mm:ss') - [CSAT] $($records.Count) registros leidos, $($deduped.Count) tras quitar duplicados."

        $json = $deduped | ConvertTo-Json -Depth 4 -Compress
        if ($deduped.Count -eq 1) { $json = "[$json]" }
        $json = Escape-JsonScriptTags $json

        $outDir = Split-Path -Parent $outputPath
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
        [System.IO.File]::WriteAllText($outputPath, $json, [System.Text.Encoding]::UTF8)
        return $true
    } catch {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - [CSAT] Error exportando datos: $($_.Exception.Message)"
        return $false
    } finally {
        if ($wb) { try { $wb.Close($false) } catch {}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
        if ($excel) { try { $excel.Quit() } catch {}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    }
}

# ---------------- Exportar Improductivos (KPI + Evolutivo Supervisor) ----------------

$DATA_COLS_IMPRO = @{
    'FECHA'='fecha'; 'PLATAFORMA'='plataforma'; 'Nombre Completo'='agente'; 'Jefe Directo'='jefe'
    'Programa'='programa'; 'Estado Laboral'='estado'; 'Antiguedad'='antiguedad'; 'Semana'='semana'
    'COPC_TOTAL_ATENDIDAS'='atendidas'; 'COPC_TOTAL_CULP_TRX'='culpTrx'; 'COPC_SHORT_CALL'='shortCall'
    'TOTAL_REITERO_VC_72HRS'='rell72'; 'TOTAL_ATENDIDAS_PARA_REITERO_VC_72HRS'='atendRell72'
    'TOTAL_REITERO_VC_30MIN'='rell30'; 'TOTAL_ATENDIDAS_PARA_REITERO_VC_30MIN'='atendRell30'
    'TOTAL_REITERO_VC_24HRS'='rell24'; 'TOTAL_ATENDIDAS_PARA_REITERO_VC_24HRS'='atendRell24'
    'TMO'='tmo'; 'Q'='q'; 'NPS_POND'='nps_pond'; 'POND'='pond'; 'RESOL_CALC'='resolCalc'
    'HOLD_TIME'='hold'; 'DURACION'='duracion'; 'TIPIFICACIONES'='tipificaciones'
    'CSAT_NUM'='csatNum'; 'CSAT_DEN'='csatDen'
}

function Get-DominantMonth {
    param($records)
    $counts = @{}
    foreach ($r in $records) { if ($r.mes) { $counts[$r.mes] = ($counts[$r.mes] + 1) } }
    $best = $null; $bestN = -1
    foreach ($m in $counts.Keys) { if ($counts[$m] -gt $bestN) { $best = $m; $bestN = $counts[$m] } }
    return $best
}

function Export-OneImproductivosFile {
    param([string]$excelPath)

    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false; $excel.EnableEvents = $false
        $wb = $excel.Workbooks.Open($excelPath, 0, $true)

        $detailInfo = $null; $detailWs = $null
        $objInfo = $null; $objWs = $null
        foreach ($ws in $wb.Worksheets) {
            $info = Get-HeaderInfo -ws $ws -mustContain @('Nombre Completo','PLATAFORMA')
            if ($info -and (-not $detailInfo -or $info.DataRows -gt $detailInfo.DataRows)) { $detailInfo = $info; $detailWs = $ws }
            if ($ws.Name.Trim() -eq 'IMPRODUCTIVAS FACTURABLES') { $objWs = $ws }
        }
        if (-not $detailInfo) { Write-Host "  No se encontro DATA_DIA_PG en $([System.IO.Path]::GetFileName($excelPath))"; return $null }

        $dVals = $detailWs.Range($detailWs.Cells(1,1), $detailWs.Cells($detailInfo.TotalRows, $detailInfo.TotalCols)).Value2
        $colMap = @{}
        foreach ($headerName in $DATA_COLS_IMPRO.Keys) {
            $trimmed = $headerName.Trim()
            if ($detailInfo.Cols.ContainsKey($trimmed)) { $colMap[$DATA_COLS_IMPRO[$headerName]] = $detailInfo.Cols[$trimmed] }
        }

        $records = New-Object System.Collections.Generic.List[object]
        for ($r = $detailInfo.HeaderRow + 1; $r -le $detailInfo.TotalRows; $r++) {
            $ag = if ($colMap.ContainsKey('agente')) { $dVals[$r, $colMap['agente']] } else { $null }
            if (-not $ag) { continue }
            $plat = if ($colMap.ContainsKey('plataforma')) { $dVals[$r, $colMap['plataforma']] } else { $null }
            if ($plat -and ($excludedPlatforms -contains $plat)) { continue }

            $fechaRaw = if ($colMap.ContainsKey('fecha')) { $dVals[$r, $colMap['fecha']] } else { $null }
            $fechaISO = ConvertExcelSerialToISO $fechaRaw
            $mes = $null
            if ($fechaISO) {
                $monthNum = [int]$fechaISO.Substring(5,2)
                if ($monthNum -ge 1 -and $monthNum -le 12) { $mes = $monthAbbr[$monthNum-1] }
            }

            $rec = [ordered]@{
                fecha = $fechaISO; mes = $mes; plataforma = $plat
                jefe = if ($colMap.ContainsKey('jefe')) { $dVals[$r, $colMap['jefe']] } else { $null }
                programa = if ($colMap.ContainsKey('programa')) { $dVals[$r, $colMap['programa']] } else { $null }
                estado = if ($colMap.ContainsKey('estado')) { $dVals[$r, $colMap['estado']] } else { $null }
                antiguedad = if ($colMap.ContainsKey('antiguedad')) { $dVals[$r, $colMap['antiguedad']] } else { $null }
                semana = if ($colMap.ContainsKey('semana')) { $dVals[$r, $colMap['semana']] } else { $null }
                atendidas = if ($colMap.ContainsKey('atendidas')) { $dVals[$r, $colMap['atendidas']] } else { 0 }
                culpTrx = if ($colMap.ContainsKey('culpTrx')) { $dVals[$r, $colMap['culpTrx']] } else { 0 }
                shortCall = if ($colMap.ContainsKey('shortCall')) { $dVals[$r, $colMap['shortCall']] } else { 0 }
                rell72 = if ($colMap.ContainsKey('rell72')) { $dVals[$r, $colMap['rell72']] } else { 0 }
                atendRell72 = if ($colMap.ContainsKey('atendRell72')) { $dVals[$r, $colMap['atendRell72']] } else { 0 }
                rell30 = if ($colMap.ContainsKey('rell30')) { $dVals[$r, $colMap['rell30']] } else { 0 }
                atendRell30 = if ($colMap.ContainsKey('atendRell30')) { $dVals[$r, $colMap['atendRell30']] } else { 0 }
                rell24 = if ($colMap.ContainsKey('rell24')) { $dVals[$r, $colMap['rell24']] } else { 0 }
                atendRell24 = if ($colMap.ContainsKey('atendRell24')) { $dVals[$r, $colMap['atendRell24']] } else { 0 }
                tmo = if ($colMap.ContainsKey('tmo')) { $dVals[$r, $colMap['tmo']] } else { 0 }
                q = if ($colMap.ContainsKey('q')) { $dVals[$r, $colMap['q']] } else { 0 }
                nps_pond = if ($colMap.ContainsKey('nps_pond')) { $dVals[$r, $colMap['nps_pond']] } else { 0 }
                pond = if ($colMap.ContainsKey('pond')) { $dVals[$r, $colMap['pond']] } else { 0 }
                resolCalc = if ($colMap.ContainsKey('resolCalc')) { $dVals[$r, $colMap['resolCalc']] } else { 0 }
                hold = if ($colMap.ContainsKey('hold')) { $dVals[$r, $colMap['hold']] } else { 0 }
                duracion = if ($colMap.ContainsKey('duracion')) { $dVals[$r, $colMap['duracion']] } else { 0 }
                tipificaciones = if ($colMap.ContainsKey('tipificaciones')) { $dVals[$r, $colMap['tipificaciones']] } else { 0 }
                csatNum = if ($colMap.ContainsKey('csatNum')) { $dVals[$r, $colMap['csatNum']] } else { 0 }
                csatDen = if ($colMap.ContainsKey('csatDen')) { $dVals[$r, $colMap['csatDen']] } else { 0 }
            }
            $records.Add($rec)
        }

        $objectives = @{}
        if ($objWs) {
            $usedObj = $objWs.UsedRange
            $objRows = $usedObj.Rows.Count; $objCols = $usedObj.Columns.Count
            $oVals = $objWs.Range($objWs.Cells(1,1), $objWs.Cells($objRows, $objCols)).Value2
            for ($r = 1; $r -le $objRows; $r++) {
                $cell0 = if ($objRows -eq 1) { $oVals } else { $oVals[$r,1] }
                if (-not $cell0) { continue }
                if (([string]$cell0).Trim().ToLower() -ne 'plataforma') { continue }
                $hdrRow = @{}
                for ($c = 2; $c -le $objCols; $c++) {
                    $hv = $oVals[$r,$c]
                    if ($hv) { $hdrRow[$c] = [string]$hv }
                }
                for ($r2 = $r+1; $r2 -le $objRows; $r2++) {
                    $platCell = $oVals[$r2,1]
                    if (-not $platCell) { break }
                    $platName = ([string]$platCell).Trim()
                    if (-not $objectives.ContainsKey($platName)) { $objectives[$platName] = @{} }
                    foreach ($c in $hdrRow.Keys) {
                        $key = $hdrRow[$c]
                        if (-not $objectives[$platName].ContainsKey($key)) {
                            $val = $oVals[$r2,$c]
                            $objectives[$platName][$key] = if ($val -eq '-' -or $val -eq '') { $null } else { $val }
                        }
                    }
                }
            }
        }

        return @{ records = $records; objectives = $objectives; month = (Get-DominantMonth $records) }
    } catch {
        Write-Host "  Error leyendo $([System.IO.Path]::GetFileName($excelPath)): $($_.Exception.Message)"
        return $null
    } finally {
        if ($wb) { try { $wb.Close($false) } catch {}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
        if ($excel) { try { $excel.Quit() } catch {}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    }
}

function Export-ImproductivosDataJson {
    param([string]$outputPath)

    $files = Find-ImproductivosFiles -folder $root
    if (-not $files -or $files.Count -eq 0) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - [KPI] No se encontraron archivos en la carpeta Improductivos."
        return $false
    }

    $allRecords = New-Object System.Collections.Generic.List[object]
    $objectivesByMonth = @{}

    foreach ($f in $files) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - [KPI] Leyendo $($f.Name)..."
        $result = Export-OneImproductivosFile -excelPath $f.FullName
        if (-not $result) { continue }
        foreach ($rec in $result.records) { $allRecords.Add($rec) }
        if ($result.month) {
            if (-not $objectivesByMonth.ContainsKey($result.month)) { $objectivesByMonth[$result.month] = @{} }
            foreach ($plat in $result.objectives.Keys) {
                $objectivesByMonth[$result.month][$plat] = $result.objectives[$plat]
            }
        }
    }

    if ($allRecords.Count -eq 0) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - [KPI] No se pudo leer ningun registro de Improductivos."
        return $false
    }

    Write-Host "$(Get-Date -Format 'HH:mm:ss') - [KPI] $($allRecords.Count) registros totales de $($files.Count) archivo(s)."

    $payload = [ordered]@{ records = $allRecords; objectivesByMonth = $objectivesByMonth }
    $json = $payload | ConvertTo-Json -Depth 6 -Compress
    $json = Escape-JsonScriptTags $json

    $outDir = Split-Path -Parent $outputPath
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    [System.IO.File]::WriteAllText($outputPath, $json, [System.Text.Encoding]::UTF8)
    return $true
}

# ---------------- Orquestacion / Git ----------------

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

    $docsFolder = Join-Path $repoFolder "docs"
    if (-not (Test-Path $docsFolder)) { New-Item -ItemType Directory -Path $docsFolder -Force | Out-Null }

    $okCsat = Export-CsatDataJson -outputPath (Join-Path $docsFolder "data.json")
    $okImpro = Export-ImproductivosDataJson -outputPath (Join-Path $docsFolder "improductivos-data.json")

    # Copiar los "cascarones" estaticos (portal + 3 paneles) tal cual, mas el video del header.
    $shellFiles = @{
        "cloud-portal.html"                = "index.html"
        "cloud-dashboard-csat.html"         = "dashboard_csat_recupero.html"
        "cloud-kpi-dashboard.html"          = "kpi-dashboard.html"
        "cloud-evolutivo-supervisor.html"   = "evolutivo-supervisor.html"
    }
    foreach ($src in $shellFiles.Keys) {
        $srcPath = Join-Path $root $src
        if (Test-Path $srcPath) {
            Copy-Item -Path $srcPath -Destination (Join-Path $docsFolder $shellFiles[$src]) -Force
        } else {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Aviso: no se encontro $src junto a este script."
        }
    }
    $videoPath = Join-Path $root "header-animation.mp4"
    if (Test-Path $videoPath) {
        Copy-Item -Path $videoPath -Destination (Join-Path $docsFolder "header-animation.mp4") -Force
    }

    if (-not $okCsat -and -not $okImpro) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - No se pudo exportar ningun dato, se cancela el push."
        return $false
    }

    Push-Location $repoFolder
    try {
        git add -A 2>&1 | Out-Null
        $status = git status --porcelain
        if (-not $status) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Sin cambios nuevos, no se sube nada."
            return $true
        }
        git commit -m "Actualizacion automatica $(Get-Date -Format 'yyyy-MM-dd HH:mm')" 2>&1 | ForEach-Object { Write-Host "  $_" }
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

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-PushToGitHub | Out-Null
}
