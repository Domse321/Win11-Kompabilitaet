# --------------------------------------------------------------------------------------
#
#   BAT-Datei-Generator
#   Beschreibung: Dieses Skript liest die 'standorte.txt' aus und erstellt
#                 automatisch fuer jeden Eintrag eine passende .bat-Datei.
#
# --------------------------------------------------------------------------------------

try {
    # Lese alle Standorte aus der Textdatei
    $standorte = Get-Content -Path ".\standorte.txt" -ErrorAction Stop
} catch {
    Write-Error "Fehler: Die Datei 'standorte.txt' wurde im aktuellen Ordner nicht gefunden."
    Start-Sleep -Seconds 10
    exit
}

# Dies ist die Vorlage fuer jede .bat-Datei.
# Der Platzhalter {0} wird spaeter durch den Standortnamen ersetzt.
$batVorlage = @"
@echo off
rem ---------------------------------------------------
rem --- DIESE DATEI WURDE AUTOMATISCH ERSTELLT ---
rem ---------------------------------------------------

set "STANDORT={0}"

rem ---------------------------------------------------
rem --- AB HIER NICHTS MEHR AENDERN ---
rem ---------------------------------------------------

powershell.exe -ExecutionPolicy Bypass -File "%~dp0\#IGNORIEREN.ps1" -StandortName "%STANDORT%"

exit
"@

Write-Host "Erstelle .bat-Dateien..." -ForegroundColor Cyan

# Gehe jeden Standort in der Liste durch
foreach ($standort in $standorte) {
    # Ersetze den Platzhalter in der Vorlage durch den aktuellen Standort
    $batInhalt = $batVorlage -f $standort
    
    # Definiere den Dateinamen (z.B. "#Berlin-Mitte.bat")
    $batDateiname = "#$($standort).bat"
    
    # Speichere die neue .bat-Datei
    Set-Content -Path $batDateiname -Value $batInhalt
    
    Write-Host "✅ Datei '$batDateiname' wurde erstellt."
}

Write-Host "`nFertig! Alle .bat-Dateien wurden erfolgreich erstellt." -ForegroundColor Green
Start-Sleep -Seconds 5