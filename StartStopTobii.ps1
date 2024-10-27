# Description : Ce script surveille la présence de l'executable de Star Citizen :
# Si l'exe est detecté les service Tobii sont lancés et les joystick Virpil USB sont activés.
# Lorsque Star Citizen est fermé, les joystick ainsi que les services tobii sont désactivés.
# Il est ensuite demandé à l'utilisateur si il veut fermer RSI Launcher.
# L'objectif est de permettre la bonne mise en veille de l'écran et de l'ordinateur qui sont bloqués par RSI Launcher, les joystick USB et les services Tobii. 

#bypass de la politique d'execution powershell
Set-ExecutionPolicy -Scope CurrentUser Bypass

# Nom du processus de Star Citizen (sans .exe)
$applicationName = "starcitizen"

# Nom du processus RSI Launcher à arrêter si l'utilisateur confirme
$rsiLauncherProcess = "RSI Launcher"

# Mot-clé des services à démarrer/arrêter
$serviceNamePattern = "tobii"

# Device IDs des périphériques USB à gérer
$usbDeviceIDs = @("USB\VID_3344&PID_03EB\FF", "USB\VID_3344&PID_43EB\FF")

# Définition du fichier de log
$logFile = Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "Log_EtatExecution.txt"

# Limite de lignes pour le fichier de log
$maxLogLines = 100

# Fonction pour écrire une entrée dans le fichier de log
function Write-Log {
    param ($message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -FilePath $logFile -Append

    # Limiter le fichier de log au nombre maximum de lignes ($maxLogLines)
    $logContent = Get-Content -Path $logFile
    if ($logContent.Count -gt $maxLogLines) {
        $logContent = $logContent[-$maxLogLines..-1]  # Garde seulement les $maxLogLines dernières lignes
        $logContent | Set-Content -Path $logFile      # Réécrit le fichier de log avec ces lignes
    }
}

# Fonction pour désactiver un périphérique USB via WMI si celui-ci est actif
function Disable-USBDevice {
    param ($deviceID)
    
    $device = Get-WmiObject Win32_PnPEntity | Where-Object { $_.DeviceID -like $deviceID}
    if ($device -and $device.Status -eq "OK") {
        $device.Disable() | Out-Null
        Write-Host "Périphérique USB $deviceID désactivé."
        Write-Log "Périphérique USB $deviceID désactivé."
    } else {
        Write-Host "Périphérique USB $deviceID est déjà désactivé ou introuvable."
        Write-Log "Périphérique USB $deviceID est déjà désactivé ou introuvable."
    }
}

# Fonction pour activer un périphérique USB via WMI
function Enable-USBDevice {
    param ($deviceID)
    
    $device = Get-WmiObject Win32_PnPEntity | Where-Object { $_.DeviceID -like $deviceID}
    if ($device) {
        $device.Disable() | Out-Null
        Start-Sleep -Milliseconds 500
        $device.Enable() | Out-Null
        Write-Host "Périphérique USB $deviceID activé."
        Write-Log "Périphérique USB $deviceID activé."
    } else {
        Write-Host "Périphérique USB $deviceID introuvable."
        Write-Log "Périphérique USB $deviceID introuvable."
    }
}

# État initial de l'application
$applicationRunning = $false

# Ajout d'une nouvelle occurence dans la log
"--- Nouvelle exécution du script : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ---" | Out-File -Append -FilePath $logFile

foreach ($usbDeviceID in $usbDeviceIDs) {
    Disable-USBDevice -deviceID $usbDeviceID
}

# Arrêter uniquement les services Tobii en cours d'exécution
$tobiiServicesRunning = Get-Service | Where-Object { $_.DisplayName -match $serviceNamePattern -and $_.Status -eq 'Running' }

if ($tobiiServicesRunning) {
    $tobiiServicesRunning | ForEach-Object {
        Stop-Service -Name $_.Name
        Write-Host "Arrêt du service $($_.DisplayName)"
        Write-Log "Arrêt du service $($_.DisplayName)"
    }
} else {
    Write-Host "Aucun service Tobii en cours d'exécution."
    Write-Log "Aucun service Tobii en cours d'exécution."
}

# Boucle principale de surveillance
while ($true) {
    $process = Get-Process -Name $applicationName -ErrorAction SilentlyContinue

    if ($process -and -not $applicationRunning) {
        Write-Host "Star Citizen détecté. Démarrage des services Tobii et activation des périphériques USB..."
        Write-Log "Star Citizen détecté. Démarrage des services Tobii et activation des périphériques USB..."
        
        Get-Service | Where-Object { $_.DisplayName -match $serviceNamePattern } | ForEach-Object {
            if ($_.Status -ne 'Running') {
                Start-Service -Name $_.Name
                Write-Host "Démarrage du service $($_.DisplayName)"
                Write-Log "Démarrage du service $($_.DisplayName)"
            }
        }

        foreach ($usbDeviceID in $usbDeviceIDs) {
            Enable-USBDevice -deviceID $usbDeviceID
        }

        $applicationRunning = $true
    }
    elseif (-not $process -and $applicationRunning) {
        Write-Host "Star Citizen fermé. Arrêt des services Tobii et désactivation des périphériques USB..."
        Write-Log "Star Citizen fermé. Arrêt des services Tobii et désactivation des périphériques USB..."
        
        Get-Service | Where-Object { $_.DisplayName -match $serviceNamePattern } | ForEach-Object {
            if ($_.Status -eq 'Running') {
                Stop-Service -Name $_.Name
                Write-Host "Arrêt du service $($_.DisplayName)"
                Write-Log "Arrêt du service $($_.DisplayName)"
            }
        }
        
        foreach ($usbDeviceID in $usbDeviceIDs) {
        Disable-USBDevice -deviceID $usbDeviceID
        }

        # Affiche une boîte de dialogue Oui/Non pour arrêter le RSI Launcher
        Add-Type -AssemblyName PresentationFramework
        $result = [System.Windows.MessageBox]::Show("Voulez-vous arrêter le processus RSI Launcher ?", "Confirmation", "YesNo", "Question")

        if ($result -eq 'Yes') {
            # Arrêter le processus RSI Launcher si l'utilisateur clique sur Oui
            $rsiProcess = Get-Process -Name $rsiLauncherProcess -ErrorAction SilentlyContinue
            if ($rsiProcess) {
                Stop-Process -Name $rsiLauncherProcess -Force
                Write-Host "Processus RSI Launcher arrêté."
                Write-log "Processus RSI Launcher arrêté."
            } else {
                Write-Host "Le processus RSI Launcher n'est pas en cours d'exécution."
                Write-log "Le processus RSI Launcher n'est pas en cours d'exécution."
            }
        } else {
            Write-Host "Le processus RSI Launcher continuera de fonctionner et empeche la mise en veille de l'ordinateur"
            Write-Host "Le processus RSI Launcher continuera de fonctionner et empeche la mise en veille de l'ordinateur"
        }

        $applicationRunning = $false
    }
    
    Start-Sleep -Seconds 5
}
