# Script d'installation de l'imprimante Toshiba Universal Printer 2
# Version adaptée pour exécution via irm | iex

# URL de base où sont hébergés les fichiers du pilote
$baseUrl = "https://asstec3.fr"  # À remplacer par votre URL réelle

# Créer un répertoire temporaire pour les fichiers
$tempDir = Join-Path $env:TEMP "ToshibaInstall_$(Get-Random)"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# Fonction pour télécharger un fichier
function Download-File {
    param(
        [string]$Url,
        [string]$Destination
    )
    try {
        Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing
        return $true
    } catch {
        Write-Host "Erreur lors du téléchargement de $Url : $_" -ForegroundColor Red
        return $false
    }
}

# Logo et informations
Clear-Host
$logo = @"
+------------------------------------------+
|     Installation Imprimante Toshiba      |
|         Universal Printer 2              |
+------------------------------------------+
"@

Write-Host $logo -ForegroundColor Cyan
Write-Host "Version: 2.0 (Web)`nDate: $(Get-Date -Format 'dd/MM/yyyy')`n" -ForegroundColor Gray

# Fonction pour afficher un message stylisé
function Show-StyledMessage {
    param (
        [string]$Message,
        [string]$Type = "Info"
    )
    
    $symbol = switch ($Type) {
        "Success" { "(+) " }
        "Warning" { "(!) " }
        "Error"   { "(x) " }
        default   { "(>) " }
    }
    
    $color = switch ($Type) {
        "Success" { "Green" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
        default   { "Cyan" }
    }
    
    Write-Host "$symbol$Message" -ForegroundColor $color
}

# Fonction pour afficher une barre de progression
function Show-Progress {
    param (
        [string]$Activity,
        [int]$PercentComplete
    )
    Write-Progress -Activity $Activity -PercentComplete $PercentComplete -Status "$PercentComplete% Complete"
}

# Vérifier si on est en mode administrateur
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Le script nécessite des droits administrateur." -ForegroundColor Yellow
    Write-Host "Relancez PowerShell en tant qu'administrateur et réexécutez la commande :" -ForegroundColor Yellow
    Write-Host "irm $baseUrl/toshiba.ps1 | iex" -ForegroundColor Cyan
    Write-Host "`nAppuyez sur une touche pour fermer..." -ForegroundColor White
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

# Télécharger le fichier ZIP contenant les pilotes
Write-Host "`nTéléchargement des pilotes Toshiba..." -ForegroundColor Yellow
$zipPath = Join-Path $tempDir "toshiba.zip"
if (-not (Download-File -Url "$baseUrl/toshiba.zip" -Destination $zipPath)) {
    Show-StyledMessage "Impossible de télécharger les pilotes. Vérifiez votre connexion internet." "Error"
    Remove-Item -Path $tempDir -Recurse -Force
    exit
}

# Extraire le fichier ZIP
Write-Host "Extraction des fichiers..." -ForegroundColor Yellow
try {
    Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force
    Show-StyledMessage "Fichiers extraits avec succès" "Success"
} catch {
    Show-StyledMessage "Erreur lors de l'extraction : $_" "Error"
    Remove-Item -Path $tempDir -Recurse -Force
    exit
}

# Définir les chemins des pilotes
$driverPath = Join-Path $tempDir "UNI\Driver\64bit"
$driverInfPath = Join-Path $driverPath "eSf6u.inf"

# Vérifier que les fichiers existent
if (-not (Test-Path $driverInfPath)) {
    Show-StyledMessage "Les fichiers du pilote sont introuvables après extraction" "Error"
    Remove-Item -Path $tempDir -Recurse -Force
    exit
}

# Fonction pour vérifier l'environnement Windows
function Test-WindowsEnvironment {
    Write-Host "`nVérification de l'environnement Windows..." -ForegroundColor Yellow
    
    $osInfo = Get-WmiObject -Class Win32_OperatingSystem
    Write-Host "Système d'exploitation : $($osInfo.Caption) $($osInfo.OSArchitecture)"
    
    if (-not [Environment]::Is64BitOperatingSystem) {
        Show-StyledMessage "ERREUR : Ce script nécessite un système Windows 64-bit." "Error"
        return $false
    }
    
    $systemDrive = Get-PSDrive -Name C
    $freeSpaceGB = [math]::Round($systemDrive.Free / 1GB, 2)
    Write-Host "Espace disque disponible : $freeSpaceGB GB"
    if ($freeSpaceGB -lt 1) {
        Show-StyledMessage "ATTENTION : Peu d'espace disque disponible ($freeSpaceGB GB). 1 GB minimum recommandé." "Warning"
    }
    
    return $true
}

# Fonction pour nettoyer les anciens pilotes Toshiba
function Remove-OldToshibaDrivers {
    Write-Host "`nRecherche d'anciens pilotes Toshiba..." -ForegroundColor Yellow
    $oldDrivers = Get-PrinterDriver | Where-Object { $_.Name -like "*TOSHIBA*" }
    
    if ($oldDrivers) {
        Write-Host "Pilotes Toshiba existants trouvés :" -ForegroundColor Yellow
        $oldDrivers | ForEach-Object { Write-Host "  - $($_.Name)" }
        Write-Host "Note: Les anciens pilotes ne seront pas supprimés pour éviter les conflits." -ForegroundColor Cyan
    } else {
        Write-Host "Aucun ancien pilote Toshiba trouvé." -ForegroundColor Green
    }
}

# Fonction principale d'installation
function Install-ToshibaPrinter {
    # Demander l'adresse IP
    Write-Host "`nÉtape 1: Configuration de l'adresse IP" -ForegroundColor Yellow
    do {
        $printerIP = Read-Host "Veuillez entrer l'adresse IP de l'imprimante"
        if ($printerIP -match "^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$") {
            Show-StyledMessage "Adresse IP valide" "Success"
            $validIP = $true
        } else {
            Show-StyledMessage "Adresse IP invalide. Format attendu: xxx.xxx.xxx.xxx" "Error"
            $validIP = $false
        }
    } while (-not $validIP)

    # Demander le nom
    Write-Host "`nÉtape 2: Nommage de l'imprimante" -ForegroundColor Yellow
    $printerName = Read-Host "Veuillez entrer le nom souhaité pour l'imprimante"

    # Vérifier si l'imprimante existe déjà
    if (Get-Printer -Name $printerName -ErrorAction SilentlyContinue) {
        Show-StyledMessage "ERREUR : Une imprimante avec le nom '$printerName' existe déjà." "Error"
        return $false
    }

    # Vérifications
    Show-Progress -Activity "Vérification de l'environnement" -PercentComplete 20
    if (-not (Test-WindowsEnvironment)) {
        return $false
    }

    Show-Progress -Activity "Recherche des pilotes existants" -PercentComplete 40
    Remove-OldToshibaDrivers

    # Test de connectivité
    Show-Progress -Activity "Test de la connectivité réseau" -PercentComplete 60
    Write-Host "`nVérification de l'accès réseau à l'imprimante..." -ForegroundColor Yellow
    $pingResult = Test-Connection -ComputerName $printerIP -Count 1 -Quiet
    if (-not $pingResult) {
        Show-StyledMessage "L'imprimante n'est pas accessible à l'adresse $printerIP" "Warning"
        Show-StyledMessage "Le pilote sera tout de même installé, mais vérifiez la connectivité." "Warning"
        
        $choice = Read-Host "Voulez-vous continuer malgré l'erreur de connectivité ? (O/N)"
        if ($choice -ne "O" -and $choice -ne "o") {
            Show-StyledMessage "Installation annulée par l'utilisateur" "Warning"
            return $false
        }
    }

    # Installation du pilote
    Write-Host "`nÉtape 3: Installation du pilote dans Windows" -ForegroundColor Yellow
    Write-Host "Installation du pilote Toshiba Universal Printer 2..."

    try {
        Write-Host "Installation du pilote avec printui.dll..." -ForegroundColor Yellow
        
        $printui = Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /ia /m `"TOSHIBA Universal Printer 2`" /f `"$driverInfPath`"" -NoNewWindow -Wait -PassThru
        
        if ($printui.ExitCode -ne 0) {
            Show-StyledMessage "ERREUR : L'installation du pilote a échoué avec printui.dll" "Error"
        } else {
            Write-Host "Le pilote a été installé avec succès via printui.dll" -ForegroundColor Green
        }
        
        # Vérifier l'installation
        $driver = Get-PrinterDriver -Name "TOSHIBA Universal Printer 2" -ErrorAction SilentlyContinue
        if (-not $driver) {
            Write-Host "Le pilote n'est pas détecté, tentative avec pnputil..." -ForegroundColor Yellow
            
            $pnputil = Start-Process -FilePath "pnputil" -ArgumentList "/add-driver `"$driverInfPath`" /install" -NoNewWindow -Wait -PassThru
            
            if ($pnputil.ExitCode -ne 0) {
                Show-StyledMessage "ERREUR : L'installation du pilote a échoué avec les deux méthodes." "Error"
            }
            
            # Réessayer avec printui
            $printui = Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /ia /m `"TOSHIBA Universal Printer 2`" /f `"$driverInfPath`"" -NoNewWindow -Wait -PassThru
            
            # Vérification finale
            $driver = Get-PrinterDriver -Name "TOSHIBA Universal Printer 2" -ErrorAction SilentlyContinue
            if (-not $driver) {
                Show-StyledMessage "ERREUR : Impossible d'installer le pilote." "Error"
                return $false
            }
        }
    } catch {
        Show-StyledMessage "ERREUR lors de l'installation du pilote: $($_.Exception.Message)" "Error"
        return $false
    }

    Show-Progress -Activity "Installation du pilote" -PercentComplete 80

    # Attendre l'installation
    Write-Host "`nAttente de la fin de l'installation du pilote..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5

    # Créer le port
    Write-Host "`nÉtape 4: Création du port réseau" -ForegroundColor Yellow
    $portName = "Port_$printerIP"
    $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue

    if ($existingPort) {
        Write-Host "Le port $portName existe déjà, il sera réutilisé." -ForegroundColor Yellow
    } else {
        Write-Host "Création du port TCP/IP pour l'adresse $printerIP..."
        try {
            Add-PrinterPort -Name $portName -PrinterHostAddress $printerIP -ErrorAction Stop
            Write-Host "Le port réseau a été créé avec succès." -ForegroundColor Green
        } catch {
            Show-StyledMessage "ERREUR : Impossible de créer le port réseau.`nDétails : $($_.Exception.Message)" "Error"
            return $false
        }
    }

    # Ajouter l'imprimante
    Write-Host "`nÉtape 5: Installation de l'imprimante" -ForegroundColor Yellow
    Write-Host "Configuration de l'imprimante $printerName..."
    try {
        Add-Printer -Name $printerName -DriverName "TOSHIBA Universal Printer 2" -PortName $portName -ErrorAction Stop
        Write-Host "L'imprimante a été configurée avec succès." -ForegroundColor Green
    } catch {
        Show-StyledMessage "ERREUR : Impossible d'installer l'imprimante.`nDétails : $($_.Exception.Message)" "Error"
        return $false
    }

    # Résumé
    $osInfo = Get-WmiObject -Class Win32_OperatingSystem
    $summary = @"
+------------- Résumé de l'Installation -------------+
| Système    : $($osInfo.Caption)
| Pilote     : TOSHIBA Universal Printer 2
| IP         : $printerIP
| Port       : $portName
| État       : $(if ($driver) { "(+) Installé" } else { "(x) Non installé" })
+----------------------------------------------------+
"@

    Write-Host "`n$summary" -ForegroundColor Cyan

    if ($driver) {
        Show-StyledMessage "Installation terminée avec succès !" "Success"
        Show-StyledMessage "L'imprimante est prête à être utilisée" "Success"
    } else {
        Show-StyledMessage "Des erreurs sont survenues pendant l'installation" "Error"
    }

    Show-Progress -Activity "Installation terminée" -PercentComplete 100
    return $true
}

# Boucle principale
do {
    Clear-Host
    Write-Host $logo -ForegroundColor Cyan
    Write-Host "Version: 2.0 (Web)`nDate: $(Get-Date -Format 'dd/MM/yyyy')`n" -ForegroundColor Gray

    # Installer l'imprimante
    $success = Install-ToshibaPrinter

    # Demander si l'utilisateur veut installer une autre imprimante
    Write-Host "`n"
    do {
        $response = Read-Host "Voulez-vous installer une autre imprimante ? (O/N)"
        switch ($response.ToUpper()) {
            "O" {
                Show-StyledMessage "Lancement d'une nouvelle installation..." "Info"
                Write-Host "`n=================================================`n"
                Start-Sleep -Seconds 2
                $continue = $true
                break
            }
            "N" {
                Show-StyledMessage "Merci d'avoir utilisé le script d'installation !" "Success"
                $continue = $false
                break
            }
            default {
                Show-StyledMessage "Réponse invalide. Veuillez répondre par O (Oui) ou N (Non)" "Warning"
            }
        }
    } while ($response -notmatch '^[OoNn]$')

} while ($continue)

# Nettoyage
Write-Host "`nNettoyage des fichiers temporaires..." -ForegroundColor Yellow
Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`nAppuyez sur une touche pour fermer la fenêtre..." -ForegroundColor White
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
