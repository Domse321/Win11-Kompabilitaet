# --------------------------------------------------------------------------------------
#
#   Win11-Kompatibilitaets-Check
#   Beschreibung: Dieses Skript sammelt Hardware- und Systeminformationen,
#                 um die Kompatibilitaet mit Windows 11 zu pruefen, und
#                 laedt einen CSV-Report in die Cloud hoch.
#
#   Autor: Dominic Seiler
#   Version: 1.5.2
#
#   Wichtige Hinweise:
#   - Das Skript sollte ueber die 'SKRIPT_START.bat'-Datei ausgefuehrt werden.
#   - Fuer praeziseste Ergebnisse die 'SKRIPT_START.bat' als Administrator starten.
#   - Mit dem Skript Report.ps1 kann man die CSV-Daten zu einer uebersichtlichen Excel umwandeln.
#
# --------------------------------------------------------------------------------------

# --- ANPASSEN ---
$StandortName = "ERS"
# Tragen Sie hier die Domain Ihrer Nextcloud ein.
$NextcloudDomain = "https://cloud.hapy-it.schule"
# Tragen Sie hier nur den Token des oeffentlichen Links ein.
$LinkToken = "nkSeLJqjdFYeeZj" 
# Tragen Sie hier das Passwort fuer den Link ein. Leer lassen, wenn es keins gibt.
$LinkPasswort = "" 

# --- 1. RAUM ERFASSEN ---
Clear-Host 
$Raum = Read-Host -Prompt "Bitte geben Sie den Raumnamen oder die Raumnummer ein (z.B. B-109 oder Turm 1B)"
while ([string]::IsNullOrWhiteSpace($Raum)) {
    $Raum = Read-Host -Prompt "Die Eingabe darf nicht leer sein. Bitte geben Sie den Raumnamen erneut ein"
}

# --- START DER SCRIPT-LOGIK ---
$ErrorActionPreference = "SilentlyContinue"
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# --- DATENSAMMLUNG ---
if ($isAdmin) { Write-Host "Sammle detaillierte Systeminformationen (Admin-Modus)..." }
else { Write-Host "Sammle Systeminformationen (Standard-Benutzer Modus)..." }
$PCName = $env:COMPUTERNAME; $OSInfo = Get-WmiObject -Class Win32_OperatingSystem; $Architektur = $OSInfo.OSArchitecture; $ArchitekturOK = $Architektur -eq '64-Bit'
function Format-SensibleRam { param($GB)
    if ($GB -gt 32) { return 64 } elseif ($GB -gt 24) { return 32 } elseif ($GB -gt 16) { return 24 }
    elseif ($GB -gt 12) { return 16 } elseif ($GB -gt 8) { return 12 } elseif ($GB -gt 6) { return 8 }
    elseif ($GB -gt 4) { return 6 } elseif ($GB -gt 2) { return 4 } else { return [math]::Round($GB) }
}
$RamModuleBytes = (Get-WmiObject -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum
$RamGB_Roh = $RamModuleBytes / 1GB; $RamGB = Format-SensibleRam -GB $RamGB_Roh; $RamOK = $RamModuleBytes -ge (4 * 1GB)
$CPU = Get-WmiObject -Class Win32_Processor; $CPUName = $CPU.Name.Trim(); $CpuKerne = $CPU.NumberOfCores; $CpuKerneOK = $CpuKerne -ge 2
$CpuTaktMHz = $CPU.MaxClockSpeed; $CpuTaktOK = $CpuTaktMHz -ge 1000
$GPU = Get-WmiObject -Class Win32_VideoController | Select-Object -First 1; $GpuName = $GPU.Name;
$GpuDX12WDDM2OK = $false; if ($OSInfo.BuildNumber -ge 10240) { $GpuDX12WDDM2OK = $true }
if (Get-Command 'Get-Disk' -ErrorAction SilentlyContinue) {
    $SystemDisk = Get-Disk | Where-Object IsSystem -eq $true | Select-Object -First 1
    $DiskGroesseGB = [math]::Round($SystemDisk.Size / 1GB); $DiskGroesseOK = $DiskGroesseGB -ge 64
    $PartitionStil = $SystemDisk.PartitionStyle; $PartitionStilOK = $PartitionStil -eq 'GPT'
} else {
    $DiskGroesseGB = "Fehlgeschlagen (altes OS)"; $DiskGroesseOK = $false
    $PartitionStil = "Fehlgeschlagen (altes OS)"; $PartitionStilOK = $false
}
if (Get-Command 'Confirm-SecureBootUEFI' -ErrorAction SilentlyContinue) {
    try {
        $FirmwareTypValue = (Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control" -Name "PEFirmwareType" -ErrorAction Stop)
        if ($FirmwareTypValue -eq 2) { $FirmwareTyp = "UEFI" } elseif ($FirmwareTypValue -eq 1) { $FirmwareTyp = "Legacy" } else { $FirmwareTyp = "Unbekannt" }
    } catch { if ($env:firmware_type) { $FirmwareTyp = $env:firmware_type } }
    $FirmwareTypOK = $FirmwareTyp -eq "UEFI"
    $SecureBootStatus = "Deaktiviert"
    if ($isAdmin) { if (Confirm-SecureBootUEFI) { $SecureBootStatus = "Aktiviert" } } 
    else { if ((Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State" -Name "UEFISecureBootEnabled" -ErrorAction SilentlyContinue) -eq 1) { $SecureBootStatus = "Aktiviert" } else { $SecureBootStatus = "nicht abrufbar" } }
    $SecureBootOK = $SecureBootStatus -eq "Aktiviert" -or $SecureBootStatus -eq "nicht abrufbar"
} else {
    $FirmwareTyp = "Legacy (vermutet)"; $FirmwareTypOK = $false
    $SecureBootStatus = "Fehlgeschlagen (altes OS)"; $SecureBootOK = $false
}
if (Get-Command 'Get-Tpm' -ErrorAction SilentlyContinue) {
    $TpmVorhanden = $false; $TpmVersion = "Nicht gefunden"; $TpmReady = $false
    if ($isAdmin) {
        $Tpm = Get-Tpm
        $TpmVorhanden = $Tpm.TpmPresent; $TpmReady = if ($TpmVorhanden) { $Tpm.TpmReady } else { $false }
        $TpmVersion = if ($TpmVorhanden) { $Tpm.SpecificationVersion.Split(',')[0] } else { "Nicht gefunden" }
        if ($TpmVorhanden -and $TpmReady -and [string]::IsNullOrWhiteSpace($TpmVersion)) { $TpmVersion = "2.x" }
    } else {
        $TpmVersion = "nicht abrufbar"; $TpmReady = "nicht abrufbar";
        $TpmInfoWmi = Get-CimInstance -Namespace "Root\CIMV2\Security\MicrosoftTpm" -ClassName Win32_Tpm -ErrorAction SilentlyContinue
        if ($TpmInfoWmi) { $TpmVorhanden = $true; $TpmVersion = $TpmInfoWmi.SpecVersion.Split(',')[0]; $TpmReady = $TpmInfoWmi.IsReady($true).IsReady }
        else { $TpmRegistryVersion = Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\TPM\WMI" -Name "SpecificationVersion" -ErrorAction SilentlyContinue
            if ($TpmRegistryVersion) { $TpmVorhanden = $true; $TpmVersion = ($TpmRegistryVersion -split ',')[0..1] -join '.'; $TpmReady = (Get-Service -Name "TPM" -ErrorAction SilentlyContinue).Status -eq 'Running' }
        }
    }
} else {
    $TpmVersion = "Fehlgeschlagen (altes OS)"; $TpmReady = $false
}
$Win11Faehig = ($ArchitekturOK -and $RamOK -and $CpuKerneOK -and $CpuTaktOK -and $DiskGroesseOK -and $PartitionStilOK -and $FirmwareTypOK -and $SecureBootOK -and $GpuDX12WDDM2OK)

# --- VERSUCHE UPLOAD ---
$UploadErfolgreich = $false
$ErgebnisObjekt = $null # Initialisiere das Objekt hier, damit es im Offline-Fall verfuegbar ist

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
    }

    $CsvInhalt = ($ErgebnisObjekt | ConvertTo-Csv -Delimiter ';' -NoTypeInformation) -join "`r`n"
    $Dateiname = "$($StandortName)-$($PCName)_report.csv" 
    $ZielUrl = "$NextcloudDomain/public.php/webdav/$Dateiname"
    $EncodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("${LinkToken}:${LinkPasswort}"))
    $Headers = @{ Authorization = "Basic $EncodedCredentials" }
    Invoke-RestMethod -Uri $ZielUrl -Method Put -Headers $Headers -Body $CsvInhalt -ContentType 'text/csv'
    $UploadErfolgreich = $true
}
catch {
    # Stelle sicher, dass das ErgebnisObjekt auch im Fehlerfall existiert,
    # falls die Datensammlung erfolgreich war, aber der Upload fehlschlug.
    if ($null -eq $ErgebnisObjekt) {
         $ErgebnisObjekt = [PSCustomObject]@{ Standort = $StandortName; Raum = $Raum; Computername = $PCName; Win11_Faehig = $false; # ... Fuelle ggf. Minimalwerte ein
         }
    }
    $UploadErfolgreich = $false
}

# --- FINALE ANZEIGE & OFFLINE-FALLBACK ---
Start-Sleep -Seconds 2
Clear-Host
if ($UploadErfolgreich) {
    if ($Win11Faehig) {
        Write-Host "`n*****************************************************" -ForegroundColor Green
        Write-Host; Write-Host "    Windows 11 ist kompatibel, Sie sind hier fertig." -ForegroundColor Green -BackgroundColor Black; Write-Host
        Write-Host "*****************************************************" -ForegroundColor Green
    } else {
        Write-Host "`n********************************************************" -ForegroundColor Red
        Write-Host; Write-Host "    Computer bitte bekleben, nicht Windows 11 faehig!" -ForegroundColor Red -BackgroundColor Black; Write-Host
        Write-Host "********************************************************" -ForegroundColor Red
    }
} else {
    # --- OFFLINE-FALLBACK ---
    $DesktopPfad = [System.Environment]::GetFolderPath('Desktop')
    
    # *** NEU: CSV-Datei auf Desktop speichern ***
    try {
        $CsvOfflineDateiname = "$($StandortName)-$($PCName)_report.csv"
        $CsvOfflinePfad = Join-Path -Path $DesktopPfad -ChildPath $CsvOfflineDateiname
        $ErgebnisObjekt | Export-Csv -Path $CsvOfflinePfad -Delimiter ';' -Encoding utf8 -NoTypeInformation
        Write-Host -ForegroundColor Cyan "`nCSV-Bericht wurde auf dem Desktop gespeichert: $CsvOfflinePfad"
    } catch {
        Write-Warning "Konnte CSV-Datei nicht auf dem Desktop speichern."
    }

    # Textdatei erstellen und oeffnen (wie bisher)
    $TextdateiInhalt = @"
**********************************
   WINDOWS 11 CHECK - OFFLINE
**********************************
Eine CSV-Datei mit den Details wurde auf dem Desktop gespeichert ($($CsvOfflineDateiname)).
Bitte kopieren Sie diese Datei (z.B. per USB-Stick) und senden Sie sie an seiler@it.hapy.schule.

Falls das nicht moeglich ist, machen Sie bitte ein Foto von diesem Text
und senden Sie es an seiler@it.hapy.schule.

STANDORT:       $($ErgebnisObjekt.Standort)
RAUM:           $($ErgebnisObjekt.Raum)
COMPUTERNAME:   $($ErgebnisObjekt.Computername)
----------------------------------
WIN11 FAEHIG?:  $($ErgebnisObjekt.Win11_Faehig)
----------------------------------
CPU Name:       $($ErgebnisObjekt.CPU_Name)
Kerne:          $($ErgebnisObjekt.CPU_Kerne)
Takt:           $($ErgebnisObjekt.CPU_Takt_MHz) MHz
RAM:            $($ErgebnisObjekt.RAM_GB) GB
Firmware:       $($ErgebnisObjekt.Firmware_Typ)
Secure Boot:    $($ErgebnisObjekt.SecureBoot_Status)
TPM Version:    $($ErgebnisObjekt.TPM_Version)
"@
    $TxtOfflineDateiname = "Win11-Report-OFFLINE_$($PCName).txt"
    $TxtOfflinePfad = Join-Path -Path $DesktopPfad -ChildPath $TxtOfflineDateiname
    $TextdateiInhalt | Out-File -FilePath $TxtOfflinePfad -Encoding utf8
    notepad.exe $TxtOfflinePfad
    
    Write-Host "`n****************************************************************" -ForegroundColor Yellow
    Write-Host; Write-Host "    KEINE INTERNETVERBINDUNG - Bitte CSV vom Desktop oder Foto senden!" -ForegroundColor Yellow -BackgroundColor Black; Write-Host
    Write-Host; Write-Host "****************************************************************" -ForegroundColor Yellow
}
# Warte 30 Sekunden im sichtbaren Fenster.
Write-Host "`nDieses Fenster schliesst sich in Kuerze automatisch..." -ForegroundColor Gray
Write-Host "`nDie Dateien loeschen sich danach selbst." -ForegroundColor Gray
Start-Sleep -Seconds 30
$ScriptPath = $MyInvocation.MyCommand.Path
$BatPath = Join-Path -Path (Split-Path -Parent $ScriptPath) -ChildPath "SKRIPT_START.bat"
$Command = "/c timeout /t 3 > NUL && del /f /q `"$BatPath`" && del /f /q `"$ScriptPath`""
Start-Process cmd.exe -ArgumentList $Command -WindowStyle Hidden