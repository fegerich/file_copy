Set-StrictMode -Version Latest

# --- Einstellungen -----------------------------------------------------------
$ExcelPath   = "C:\Users\felix\VsCodeProjects\export_aramais-17-10-2022.xlsx"   # Excel-Datei mit den Ordnernummern
$SheetName   = "Tabelle1"                       # Blattname
$ColumnName  = "Auftragsnumer"                  # Spaltenname mit den Ordnernummern
$HasHeader   = $true                            # $false, wenn KEINE Kopfzeile vorhanden ist

$SourceRoot  = "C:\Users\root"    # Quell-ROOT
$TargetRoot  = "C:\Zieldaten" # Ziel-ROOT

$Prefix      = ""         # z. B. "ab" oder "" für kein Präfix
$SubfolderToCopy = "hallo"    # nur dieser Unterordner je Nummer kopieren

# Dateiendungen: @() = alle Dateien, sonst nur die angegebenen
$AllowedExtensions = @(".txt")                 # z. B. @(".pdf",".csv") oder @() für alle

$Overwrite   = $true
$WhatIf      = $false
# ---------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $TargetRoot)) {
    New-Item -ItemType Directory -Path $TargetRoot | Out-Null
}

# Excel lesen (COM)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open($ExcelPath)
    $ws = $wb.Worksheets.Item($SheetName)

    $used = $ws.UsedRange
    $rowStart = if ($HasHeader) { 2 } else { 1 }
    $rowEnd   = $used.Rows.Count

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

function Ensure-Dir([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
}

Write-Host "Gefundene Einträge in Excel: $($names.Count)"

foreach ($name in $names) {
    $folderName = "$Prefix$name"

    # Quelle: <SourceRoot>\<Nummer>\<SubfolderToCopy>
    $sourceBase = Join-Path $SourceRoot $folderName
    $sourcePath = Join-Path $sourceBase $SubfolderToCopy

    # Ziel:   <TargetRoot>\<Nummer>
    $destPath  = Join-Path $TargetRoot $folderName

    if (-not (Test-Path -LiteralPath $sourcePath -PathType Container)) {
        Write-Warning "Unterordner nicht gefunden: $sourcePath"
        continue
    }

    Ensure-Dir $destPath

    # Nur Dateien direkt im angegebenen Unterordner (ohne Unterordner)
    $files = Get-ChildItem -Path $sourcePath -File   # wichtig: kein -Recurse

    # ggf. nach Endungen filtern
    if ($AllowedExtensions.Count -gt 0) {
        $files = $files | Where-Object { $AllowedExtensions -contains $_.Extension.ToLower() }
    }

    foreach ($file in $files) {
        $targetFile = Join-Path $destPath $file.Name
        Copy-Item -LiteralPath $file.FullName -Destination $targetFile -Force:($Overwrite) -WhatIf:$WhatIf
        if (-not $WhatIf) { Write-Host "Kopiert Datei: $($file.Name) -> $targetFile" }
    }
}
