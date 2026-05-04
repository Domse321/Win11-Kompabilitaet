# Win11-Kompatibilitäts-Check & Inventarisierung

![Windows 11](https://img.shields.io/badge/Windows-11-blue?style=for-the-badge&logo=windows)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?style=for-the-badge&logo=powershell)

Dieses Projekt bietet eine automatisierte Lösung, um die Hardware-Kompatibilität von Windows-Systemen mit Windows 11 zu prüfen und gleichzeitig wichtige Inventardaten zu erfassen. Es ist besonders für Umgebungen mit vielen Geräten an verschiedenen Standorten geeignet.

## 🚀 Funktionen

* **Detaillierter Hardware-Check:** Prüfung von CPU (Kerne/Takt), RAM, Speicherplatz, Partitionstil (GPT), Firmware (UEFI), Secure Boot und TPM 2.0.
* **Inventarisierung:** Erfassung von Seriennummer, MAC-Adresse, Hersteller, Modell und OS-Installationsdatum.
* **Automatisierung:** Generierung von standortspezifischen Starter-Dateien (.bat).
* **Zentrales Reporting:** Zusammenführung aller Einzelberichte in eine professionelle Excel-Datei mit Filterfunktionen.

## 📁 Dateistruktur

* `Win11-Check.ps1`: Das Kernskript zur Datensammlung und Kompatibilitätsprüfung.
* `Standort_BATs_erstellen.ps1`: Erzeugt automatisch Batch-Dateien für jeden Standort.
* `standorte.txt`: Steuerungsdatei mit einer Liste der Standorte.
* `Report.ps1`: Skript zur Erstellung des Gesamt-Excel-Reports aus den gesammelten CSV-Daten.

## 🛠 Voraussetzungen

* **Betriebssystem:** Windows 7 oder neuer.
* **Rechte:** Für vollständige Ergebnisse (TPM/Secure Boot) ist die Ausführung als **Administrator** erforderlich.
* **Module:** Für den Excel-Export wird das PowerShell-Modul `ImportExcel` benötigt (wird bei Bedarf vom Report-Skript installiert).

## 📖 Nutzung des Workflows

### 1. Vorbereitung
Trage deine Standorte zeilenweise in die `standorte.txt` ein. Führe anschließend `Standort_BATs_erstellen.ps1` aus, um für jeden Standort eine eigene `.bat`-Datei zu erhalten.

### 2. Durchführung am Client
Starte die entsprechende `.bat`-Datei (z.B. `#Berlin.bat`) auf dem Zielrechner, idealerweise als Administrator.
* Gib den Raumnamen/Nummer ein.
* Das Skript erstellt eine lokale CSV-Datei mit den Ergebnissen im selben Ordner.
* Eine visuelle Rückmeldung zeigt sofort an, ob das Gerät Windows 11 fähig ist.

### 3. Auswertung
Sammle alle erzeugten CSV-Dateien in einem Ordner und führe `Report.ps1` aus. 
* Das Skript führt alle Daten zusammen.
* Es wird eine Datei `Gesamt-Report.xlsx` erstellt, die automatisch geöffnet wird.

## ⚠️ Wichtige Hinweise zum TPM-Check
* **Admin-Modus:** Der TPM-Status wird zwingend für die Kompatibilitätswertung herangezogen.
* **Standard-Nutzer:** Falls keine Admin-Rechte vorhanden sind, wird der TPM-Check für die finale "Win11-Fähig" Wertung ignoriert, da der Zugriff auf die Hardware-Schnittstelle eingeschränkt ist.

---
**Entwickelt von:** Dominic Seiler
