# --- Einstellungen -----------------------------------------------------------
$ExcelPath   = "C:\Pfad\zu\deiner\Liste.xlsx"   # Excel-Datei mit den Namen
$SheetName   = "Tabelle1"                       # Blattname
$ColumnName  = "Column1"                        # Spaltenname mit den Ordnernummern
$HasHeader   = $true                            # $false, wenn KEINE Kopfzeile vorhanden ist

$SourceDir   = "C:\Users\felix\Desktop\Quelle"  # Quellordner mit allen Unterordnern
$TargetDir   = "C:\Users\felix\Desktop\Ziel"    # Zielordner

$Prefix      = "ab"                             # Präfix, das vor die Excel-Werte gesetzt wird

$Overwrite   = $true    # vorhandene Ordner im Ziel überschreiben
$WhatIf      = $false   # $true = nur anzeigen, was passieren würde
# ---------------------------------------------------------------------------

# Zielordner anlegen, falls nötig
if (-not (Test-Path -LiteralPath $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir | Out-Null
}

# Excel öffnen (COM, kein Extra-Modul nötig)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open($ExcelPath)
    $ws = $wb.Worksheets.Item($SheetName)

    # benutzten Bereich holen
    $used = $ws.UsedRange
    $rowStart = if ($HasHeader) { 2 } else { 1 }
    $rowEnd   = $used.Rows.Count

    # Spaltenindex für $ColumnName ermitteln (oder 1 wenn keine Header)
    $colIdx = 1
    if ($HasHeader) {
        $header = @{}
        for ($c=1; $c -le $used.Columns.Count; $c++) {
            $name = ($ws.Cells.Item(1,$c).Text).Trim()
            if ($name) { $header[$name] = $c }
        }
        if (-not $header.ContainsKey($ColumnName)) {
            throw "Spalte '$ColumnName' wurde im Blatt '$SheetName' nicht gefunden."
        }
        $colIdx = $header[$ColumnName]
    }

    # Ordnernummern auslesen
    $names = @()
    for ($r=$rowStart; $r -le $rowEnd; $r++) {
        $val = ($ws.Cells.Item($r, $colIdx).Text).Trim()
        if ($val) { $names += $val }
    }

} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)   | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)   | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)| Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host "Gefundene Einträge in Excel: $($names.Count)"

$copyParams = @{Recurse = $true}
if ($Overwrite) { $copyParams['Force'] = $true }
if ($WhatIf)    { $copyParams['WhatIf'] = $true }

# Kopieren
foreach ($name in $names) {
    # Präfix voranstellen
    $folderName = "$Prefix$name"

    $sourcePath = Join-Path -Path $SourceDir -ChildPath $folderName
    $destPath   = Join-Path -Path $TargetDir -ChildPath $folderName

    if (Test-Path -LiteralPath $sourcePath -PathType Container) {
        Copy-Item -Path $sourcePath -Destination $destPath @copyParams
        if (-not $WhatIf) { Write-Host "Kopiert: $sourcePath -> $destPath" }
    } else {
        Write-Warning "Ordner nicht gefunden: '$folderName'"
    }
}
