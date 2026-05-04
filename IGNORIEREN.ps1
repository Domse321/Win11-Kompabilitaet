# --------------------------------------------------------------------------------------
#
#   Win11-Kompatibilitaets-Check & Inventarisierung
#   Beschreibung: Dieses Skript sammelt Hardware- und Systeminformationen,
#                 prueft die Kompatibilitaet mit Windows 11 und
#                 erstellt einen lokalen CSV-Report mit Inventardaten.
#
#   Autor: Dominic Seiler
#   Version: 2.4.0 (Strukturell optimiert)
#
#   Wichtige Hinweise:
#   - Das Skript sollte ueber die '*Standortname*.bat'-Datei ausgefuehrt werden.
#   - Fuer praeziseste Ergebnisse die '*Standortname*.bat' als Administrator starten.
#   - Mit dem Skript Report.ps1 kann man die CSV-Daten zu einer uebersichtlichen Excel umwandeln.
#
# --------------------------------------------------------------------------------------

# --- PARAMETER DEFINITION ---
param (
    [Parameter(Mandatory=$true)]
    [string]$StandortName
)

# --- HELFER-FUNKTIONEN ---
function Format-SensibleRam { 
    param($GB)
    if ($GB -gt 32) { return 64 } elseif ($GB -gt 24) { return 32 } elseif ($GB -gt 16) { return 24 }
    elseif ($GB -gt 12) { return 16 } elseif ($GB -gt 8) { return 12 } elseif ($GB -gt 6) { return 8 }
    elseif ($GB -gt 4) { return 6 } elseif ($GB -gt 2) { return 4 } else { return [math]::Round($GB) }
}

# --- START DER SCRIPT-LOGIK ---
$ErrorActionPreference = "SilentlyContinue"
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

# --- 1. RAUM ERFASSEN ---
Clear-Host 
Write-Host -NoNewline "Standort: "
Write-Host -NoNewline "'$StandortName'" -ForegroundColor Cyan
Write-Host " - richtig? Sonst bitte Fenster schliessen und richtige Datei ausfuehren!"
Write-Host "" # Leerzeile

if (-not $isAdmin) {
    Write-Host "Skript bitte als Admin ausfuehren, wenn moeglich (Aktuell = Standardnutzer)" -ForegroundColor Red
    Write-Host "" # Leerzeile
}

$Raum = Read-Host -Prompt "Bitte geben Sie den Raumnamen oder die Raumnummer ein (z.B. B-109 oder Turm 1B)"
while ([string]::IsNullOrWhiteSpace($Raum)) {
    $Raum = Read-Host -Prompt "Die Eingabe darf nicht leer sein. Bitte geben Sie den Raumnamen erneut ein"
}

# --- DATENSAMMLUNG ---
Write-Host "Sammle Systeminformationen, bitte warten..." -ForegroundColor Cyan
if ($isAdmin) { Write-Host "(Admin-Modus erkannt)" -ForegroundColor Gray }

# Basis-Informationen (einmalige Abfrage)
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$PCName = $env:COMPUTERNAME
$OSInfo = Get-WmiObject -Class Win32_OperatingSystem
$CSInfo = Get-WmiObject -Class Win32_ComputerSystem 
$BiosInfo = Get-WmiObject -Class Win32_BIOS       
$CPUInfo = Get-WmiObject -Class Win32_Processor
$GPUInfo = Get-WmiObject -Class Win32_VideoController | Select-Object -First 1

# Architektur
$Architektur = $OSInfo.OSArchitecture
$ArchitekturOK = $Architektur -eq '64-Bit'

# CPU
$CPUName = $CPUInfo.Name.Trim()
$CpuKerne = $CPUInfo.NumberOfCores
$CpuKerneOK = $CpuKerne -ge 2
$CpuTaktMHz = $CPUInfo.MaxClockSpeed
$CpuTaktOK = $CpuTaktMHz -ge 1000

# RAM
$RamModuleBytes = (Get-WmiObject -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum
$RamGB = Format-SensibleRam -GB ($RamModuleBytes / 1GB)
$RamOK = $RamModuleBytes -ge (4 * 1GB)

# GPU
$GpuName = $GPUInfo.Name
$GpuDX12WDDM2OK = $OSInfo.BuildNumber -ge 10240

# Disk, Firmware, Secure Boot & TPM (in Blöcken für Kompatibilität)
if (Get-Command 'Get-Disk' -ErrorAction SilentlyContinue) {
    $SystemDisk = Get-Disk | Where-Object IsSystem -eq $true | Select-Object -First 1
    $DiskGroesseGB = [math]::Round($SystemDisk.Size / 1GB); $DiskGroesseOK = $DiskGroesseGB -ge 64
    $PartitionStil = $SystemDisk.PartitionStyle; $PartitionStilOK = $PartitionStil -eq 'GPT'
} else {
    $DiskGroesseGB = "Fehlgeschlagen (altes OS)"; $DiskGroesseOK = $false
    $PartitionStil = "Fehlgeschlagen (altes OS)"; $PartitionStilOK = $false
}
if (Get-Command 'Confirm-SecureBootUEFI' -ErrorAction SilentlyContinue) {
    # Firmware Typ ermitteln
    try {
        $FirmwareTypValue = (Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control" -Name "PEFirmwareType" -ErrorAction Stop)
        if ($FirmwareTypValue -eq 2) { $FirmwareTyp = "UEFI" } elseif ($FirmwareTypValue -eq 1) { $FirmwareTyp = "Legacy" }
        else { $FirmwareTyp = "Unbekannt" } 
    } catch { 
        if ($env:firmware_type) { $FirmwareTyp = $env:firmware_type } 
        else { $FirmwareTyp = "Unbekannt (Fehler)" } 
    }
    # Secure Boot Status ermitteln
    $SecureBootStatus = "Deaktiviert"
    if ($isAdmin) { if (Confirm-SecureBootUEFI) { $SecureBootStatus = "Aktiviert" } } 
    else { if ((Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State" -Name "UEFISecureBootEnabled") -eq 1) { $SecureBootStatus = "Aktiviert" } else { $SecureBootStatus = "nicht abrufbar" } }
    
    # Firmware Typ korrigieren, falls Secure Boot aktiv aber Typ unbekannt
    if ($FirmwareTyp -eq "Unbekannt" -and $SecureBootStatus -eq "Aktiviert") { 
        $FirmwareTyp = "UEFI (abgeleitet)" 
    }
    $FirmwareTypOK = $FirmwareTyp -match "UEFI" 
    # Secure Boot gilt als OK, wenn die Firmware UEFI ist
    $SecureBootOK = $FirmwareTypOK
} else {
    $FirmwareTyp = "Legacy (vermutet)"; $FirmwareTypOK = $false
    $SecureBootStatus = "Fehlgeschlagen (altes OS)"; $SecureBootOK = $false
}
if (Get-Command 'Get-Tpm' -ErrorAction SilentlyContinue) {
    $TpmVorhanden = $false; $TpmVersion = "Nicht gefunden"; $TpmReady = $false
    if ($isAdmin) {
        $Tpm = Get-Tpm
        $TpmVorhanden = $Tpm.TpmPresent; $TpmReady = if ($TpmVorhanden) { $Tpm.TpmReady } else { $false }
        $TpmInfo = Get-WmiObject -Namespace "root\CIMV2\Security\MicrosoftTpm" -ClassName Win32_Tpm 
        if ($TpmInfo) { $TpmVersion = $TpmInfo.SpecVersion.Split(',')[0] }
        if ($TpmVorhanden -and $TpmReady -and [string]::IsNullOrWhiteSpace($TpmVersion)) { $TpmVersion = "2.x" }
    } else {
        $TpmVersion = "nicht abrufbar"; $TpmReady = "nicht abrufbar"
        $TpmInfoWmi = Get-WmiObject -Namespace "Root\CIMV2\Security\MicrosoftTpm" -ClassName Win32_Tpm
        if ($TpmInfoWmi) { 
            $TpmVorhanden = $true
            $TpmVersion = $TpmInfoWmi.SpecVersion.Split(',')[0]
            $TpmReady = ($TpmInfoWmi.IsEnabled() -and $TpmInfoWmi.IsActivated() -and $TpmInfoWmi.IsOwned()) 
        } else { 
            $TpmRegistryVersion = Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\TPM\WMI" -Name "SpecificationVersion"
            if ($TpmRegistryVersion) { 
                $TpmVorhanden = $true
                $TpmVersion = ($TpmRegistryVersion -split ',')[0..1] -join '.'
                $TpmReady = (Get-Service -Name "TPM").Status -eq 'Running' 
            }
        }
    }
} else {
    $TpmVersion = "Fehlgeschlagen (altes OS)"; $TpmReady = $false
}
# TPM gilt als OK, wenn ein 2.x Modul vorhanden ist (Version 2.x oder nicht abrufbar), unabhaengig von Ready
$TpmOK = $TpmVorhanden -and ($TpmVersion -match "^2\." -or $TpmVersion -eq "nicht abrufbar")

# --- Zusaetzliche Inventar-Daten ---
$Seriennummer = $BiosInfo.SerialNumber
$MacAdresse = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='TRUE'").MACAddress | Select-Object -First 1
$Hersteller = $CSInfo.Manufacturer
$Modell = $CSInfo.Model
$OSInstallationsDatum = ([WMI]'').ConvertToDateTime($OSInfo.InstallDate).ToString("yyyy-MM-dd") # Formatiertes Datum

Write-Host "Datensammlung abgeschlossen." -ForegroundColor Green

# --- Finale Kompatibilitaetspruefung (mit Admin-abhaengiger TPM-Logik) ---
$BasicsOK = ($ArchitekturOK -and $RamOK -and $CpuKerneOK -and $CpuTaktOK -and $DiskGroesseOK -and $PartitionStilOK -and $FirmwareTypOK -and $SecureBootOK -and $GpuDX12WDDM2OK)
if ($isAdmin) {
    # Im Admin-Modus muss AUCH der TPM-Check bestanden sein
    $Win11Faehig = $BasicsOK -and $TpmOK
} else {
    # Im Standard-Modus wird der TPM-Check ignoriert
    $Win11Faehig = $BasicsOK
}

# --- SPEICHERE DATEI LOKAL ---
try {
    $ErgebnisObjekt = [PSCustomObject]@{
        Standort           = $StandortName
        Raum               = $Raum
        Win11_Faehig       = $Win11Faehig
        Computername       = $PCName
        Architektur        = $Architektur
        CPU_Name           = $CPUName
        CPU_Kerne          = $CpuKerne
        CPU_Takt_MHz       = $CpuTaktMHz
        RAM_GB             = $RamGB
        Disk_Groesse_GB    = $DiskGroesseGB
        Firmware_Typ       = $FirmwareTyp
        Partition_Stil     = $PartitionStil
        SecureBoot_Status  = $SecureBootStatus
        TPM_Version        = $TpmVersion
        TPM_Ready          = $TpmReady
        GPU_Name           = $GpuName
        GPU_DX12_WDDM2_OK  = $GpuDX12WDDM2OK
        Seriennummer       = $Seriennummer
        MAC_Adresse        = $MacAdresse
        Hersteller         = $Hersteller
        Modell             = $Modell
        OS_Installiert_Am  = $OSInstallationsDatum
    }
    
    $SaubereSeriennummer = $Seriennummer -replace '[\\/:*?"<>|]', '-'
    $Dateiname = "$($StandortName)-$($PCName)-$($SaubereSeriennummer).csv"
    $Ausgabepfad = Join-Path -Path $PSScriptRoot -ChildPath $Dateiname
    
    $ErgebnisObjekt | Export-Csv -Path $Ausgabepfad -Delimiter ';' -Encoding utf8 -NoTypeInformation
    
    Write-Host -ForegroundColor Green "`n--------------------------------------------------"
    Write-Host -ForegroundColor Green "✅ Bericht erfolgreich erstellt/ueberschrieben:"
    Write-Host -ForegroundColor Cyan $Ausgabepfad
    Write-Host -ForegroundColor Green "--------------------------------------------------"
} catch {
    Write-Host -ForegroundColor Red "❌ FEHLER: Bericht konnte nicht erstellt werden."; Write-Host -ForegroundColor Yellow $_.Exception.Message
}

# --- FINALE ANZEIGE ---
Start-Sleep -Seconds 2
Clear-Host
if ($Win11Faehig) {
    Write-Host "`n*****************************************************" -ForegroundColor Green
    Write-Host; Write-Host "    Windows 11 ist kompatibel, Sie sind hier fertig." -ForegroundColor Green -BackgroundColor Black; Write-Host
    Write-Host "*****************************************************" -ForegroundColor Green
} else {
    Write-Host "`n********************************************************" -ForegroundColor Red
    Write-Host; Write-Host "    Computer bitte bekleben, nicht Windows 11 faehig!" -ForegroundColor Red -BackgroundColor Black; Write-Host
    Write-Host "********************************************************" -ForegroundColor Red
}

# Warte 30 Sekunden im sichtbaren Fenster.
Write-Host "`nDieses Fenster schliesst sich in Kuerze automatisch..." -ForegroundColor Gray
Start-Sleep -Seconds 15
