# --- START DES SKRIPTS ---

# Pruefe, ob das notwendige Excel-Modul installiert ist
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Das benoetigte PowerShell-Modul 'ImportExcel' wird nicht gefunden." -ForegroundColor Yellow
    $antwort = Read-Host "Soll versucht werden, es zu installieren? (j/n)"
    if ($antwort -eq 'j') {
        Write-Host "Installiere 'ImportExcel' von der PowerShell Gallery..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Repository PSGallery -Force
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Error "Installation fehlgeschlagen. Bitte PowerShell als Administrator neustarten und 'Install-Module ImportExcel' manuell ausfuehren."
            Start-Sleep -Seconds 10; exit
        }
    } else {
        Write-Error "Skript kann ohne das 'ImportExcel'-Modul nicht ausgefuehrt werden."
        Start-Sleep -Seconds 10; exit
    }
}

# Definiere Pfade
$Quellordner = $PSScriptRoot
$Zieldatei = Join-Path -Path $Quellordner -ChildPath "Gesamt-Report.xlsx"

Write-Host "Lese alle '*.csv'-Dateien im aktuellen Ordner..."

# --- HIER IST DIE FINALE KORREKTUR: Text-basiertes Zusammenfuegen ---
$csvDateien = Get-ChildItem -Path $Quellordner -Filter "*.csv"

if ($csvDateien) {
    # Nimm den Header von der ERSTEN Datei
    $header = Get-Content -Path $csvDateien[0].FullName -TotalCount 1

    # Sammle den Inhalt ALLER Dateien (ohne deren Header)
    $inhalt = $csvDateien | ForEach-Object {
        Get-Content -Path $_.FullName | Select-Object -Skip 1
    }

    # Fuege den einen Header und den gesamten Inhalt zusammen
    $alleDatenText = @($header) + @($inhalt)

    # Konvertiere den finalen Textblock in einem einzigen Schritt in Objekte
    $alleDaten = $alleDatenText | ConvertFrom-Csv -Delimiter ';'
    
    Write-Host "$($alleDaten.Count) Berichte gefunden. Erstelle Excel-Datei..."
    
    # --- STABILER EXCEL-EXPORT ---
    # Exportiere die Daten zuerst ohne die problematischen Features
    $ExcelPackage = $alleDaten | Export-Excel -Path $Zieldatei -WorksheetName "Win11_Inventur" -AutoSize -PassThru
    
    # Fuege die Tabelle und den AutoFilter nachtraeglich zum erstellten Paket hinzu
    $Worksheet = $ExcelPackage.Workbook.Worksheets["Win11_Inventur"]
    $TableName = "Inventurdaten"
    $Table = $Worksheet.Tables.Add($Worksheet.Dimension, $TableName)
    $Table.ShowFilter = $true

    # Speichere das finale Excel-Paket
    Close-ExcelPackage $ExcelPackage
    
    Write-Host -ForegroundColor Green "✅ Fertig! Die Datei '$Zieldatei' wurde erfolgreich erstellt."
    
    # Oeffne die fertige Excel-Datei direkt
    Invoke-Item $Zieldatei
} else {
    Write-Warning "Keine 'REPORT-*.csv'-Dateien im Ordner gefunden."
}

Start-Sleep -Seconds 5