Set-StrictMode -Version Latest

# --- Einstellungen -----------------------------------------------------------
$ExcelPath   = "C:\Pfad\zu\deiner\Liste.xlsx"   # Excel-Datei mit den Ordnernummern
$SheetName   = "Tabelle1"                       # Blattname
$ColumnName  = "Column1"                        # Spaltenname mit den Ordnernummern
$HasHeader   = $false                            # $false, wenn KEINE Kopfzeile vorhanden ist

$SourceRoot  = "C:\Users\felix\Desktop\Quelle"  # FIX: Quell-ROOT (unverändert lassen)
$TargetRoot  = "C:\Users\felix\Desktop\Ziel" # FIX: Ziel-ROOT (unverändert lassen)

$Prefix      = "ab"                               # z.B. "ab" oder "" für kein Präfix

# Dateiendungen: leer => alle Dateien kopieren
$AllowedExtensions = @(".txt", ".csv")          # Beispiel
# $AllowedExtensions = @()                      # alle Dateien

$Overwrite   = $true
$WhatIf      = $false
# ---------------------------------------------------------------------------

# Zielroot anlegen
if (-not (Test-Path -LiteralPath $TargetRoot)) {
    New-Item -ItemType Directory -Path $TargetRoot | Out-Null
}

# Excel öffnen (COM)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open($ExcelPath)
    $ws = $wb.Worksheets.Item($SheetName)

    $used = $ws.UsedRange
    $rowStart = if ($HasHeader) { 2 } else { 1 }
    $rowEnd   = $used.Rows.Count

    # Spaltenindex ermitteln
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

    # Werte lesen
    $names = @()
    for ($r=$rowStart; $r -le $rowEnd; $r++) {
        $val = ($ws.Cells.Item($r, $colIdx).Text).Trim()
        if ($val) { $names += $val }
    }
}
finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)   | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)   | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)| Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host "Gefundene Einträge in Excel: $($names.Count)"

# Hilfsfunktion: sichere Ordnererstellung
function Ensure-Dir($path) {
    if (-not (Test-Path -LiteralPath $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
}

foreach ($name in $names) {
    # WICHTIG: pro Iteration NEU berechnen, Root NIE überschreiben!
    $folderName = "$Prefix$name"
    $sourcePath = Join-Path -Path $SourceRoot -ChildPath $folderName
    $destPath   = Join-Path -Path $TargetRoot -ChildPath $folderName

    if (-not (Test-Path -LiteralPath $sourcePath -PathType Container)) {
        Write-Warning "Ordner nicht gefunden: $folderName"
        continue
    }

    # Ziel-Ordner für diesen Eintrag anlegen (flache Ebene unter TargetRoot)
    Ensure-Dir $destPath

    if ($AllowedExtensions.Count -eq 0) {
        # GANZEN ORDNER rekursiv kopieren (flach unter $destPath)
        Copy-Item -Path $sourcePath -Destination $destPath -Recurse -Force:($Overwrite) -WhatIf:$WhatIf
        if (-not $WhatIf) { Write-Host "Kopiert gesamten Ordner: $sourcePath -> $destPath" }
        continue
    }

    # Nur erlaubte Endungen kopieren (Ordnerstruktur UNTERHALB des jeweiligen Ordners beibehalten)
    $files = Get-ChildItem -Path $sourcePath -Recurse -File | Where-Object {
        $AllowedExtensions -contains $_.Extension.ToLower()
    }

    foreach ($file in $files) {
        # Relativ zum AKTUELLEN sourcePath berechnen (damit keine Fremdverschachtelung entsteht)
        $relativePath = $file.FullName.Substring($sourcePath.Length).TrimStart('\')
        $targetFile   = Join-Path -Path $destPath -ChildPath $relativePath

        $targetDir = Split-Path -Path $targetFile -Parent
        Ensure-Dir $targetDir

        Copy-Item -LiteralPath $file.FullName -Destination $targetFile -Force:($Overwrite) -WhatIf:$WhatIf
        if (-not $WhatIf) { Write-Host "Kopiert Datei: $($file.Name) -> $targetFile" }
    }
}
