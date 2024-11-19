# Script d'installation de l'imprimante Toshiba Universal Printer 2

# Recuperer le chemin du script
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path -Parent $scriptPath

# Debug - Afficher les informations de chemin
Clear-Host
$logo = @"
+------------------------------------------+
|     Installation Imprimante Toshiba      |
|         Universal Printer 2              |
+------------------------------------------+
"@

Write-Host $logo -ForegroundColor Cyan
Write-Host "Version: 1.0`nDate: $(Get-Date -Format 'dd/MM/yyyy')`n" -ForegroundColor Gray

# Fonction pour afficher un message stylisé
function Show-StyledMessage {
    param (
        [string]$Message,
        [string]$Type = "Info" # Info, Success, Warning, Error
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

# Verifier si on est en mode administrateur
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Si on n'est pas admin, relancer le script en tant qu'admin avec le bon chemin
if (-not $isAdmin) {
    Write-Host "Le script necessite des droits administrateur. Demande d'elevation des privileges..." -ForegroundColor Yellow
    $arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"Set-Location '$scriptDirectory'; & '$scriptPath'`""
    Write-Host "Commande executee: $arguments" -ForegroundColor Magenta
    Start-Process powershell -Verb RunAs -ArgumentList $arguments -Wait
    exit
}

# S'assurer qu'on est dans le bon répertoire
Set-Location $scriptDirectory
Write-Host "Nouveau repertoire de travail: $(Get-Location)" -ForegroundColor Magenta

# Demander l'adresse IP de l'imprimante avec validation en temps réel
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

# Demander le nom de l'imprimante
Write-Host "`nÉtape 2: Nommage de l'imprimante" -ForegroundColor Yellow
$printerName = Read-Host "Veuillez entrer le nom souhaite pour l'imprimante"

# Verifier si une imprimante avec ce nom existe deja
if (Get-Printer -Name $printerName -ErrorAction SilentlyContinue) {
    Show-StyledMessage "ERREUR : Une imprimante avec le nom '$printerName' existe deja." "Error"
    exit
}

# Chemin absolu du fichier INF du pilote
$driverPath = Join-Path $scriptDirectory "UNI\Driver\64bit"
$driverInfPath = Join-Path $driverPath "eSf6u.inf"

Write-Host "`nVerification du chemin du pilote:" -ForegroundColor Yellow
Write-Host "Repertoire des pilotes: $driverPath"
Write-Host "Fichier INF: $driverInfPath"

# Verifier que le dossier des pilotes existe
if (-not (Test-Path $driverPath)) {
    Show-StyledMessage "ERREUR : Le dossier des pilotes n'existe pas: $driverPath" "Error"
    exit
}

# Verifier que le fichier INF existe
if (-not (Test-Path $driverInfPath)) {
    Show-StyledMessage "ERREUR : Le fichier pilote n'existe pas: $driverInfPath" "Error"
    exit
}

# Lister le contenu du dossier des pilotes
Write-Host "`nFichiers presents dans le dossier des pilotes:" -ForegroundColor Yellow
Get-ChildItem $driverPath | ForEach-Object {
    Write-Host "  $($_.Name)"
}

# Fonction pour vérifier l'environnement Windows
function Test-WindowsEnvironment {
    Write-Host "`nVérification de l'environnement Windows..." -ForegroundColor Yellow
    
    # Vérifier la version de Windows
    $osInfo = Get-WmiObject -Class Win32_OperatingSystem
    Write-Host "Système d'exploitation : $($osInfo.Caption) $($osInfo.OSArchitecture)"
    
    # Vérifier si nous sommes en 64-bit
    if (-not [Environment]::Is64BitOperatingSystem) {
        Show-StyledMessage "ERREUR : Ce script nécessite un système Windows 64-bit." "Error"
        return $false
    }
    
    # Vérifier l'espace disque disponible (minimum 1 GB recommandé)
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
        Write-Host "Note: Les anciens pilotes ne seront pas supprimés pour éviter les conflits avec d'autres imprimantes." -ForegroundColor Cyan
    } else {
        Write-Host "Aucun ancien pilote Toshiba trouvé." -ForegroundColor Green
    }
}

# Vérification de l'environnement avec barre de progression
Show-Progress -Activity "Vérification de l'environnement" -PercentComplete 20
if (-not (Test-WindowsEnvironment)) {
    exit 1
}

Show-Progress -Activity "Recherche des pilotes existants" -PercentComplete 40
Remove-OldToshibaDrivers

Show-Progress -Activity "Test de la connectivité réseau" -PercentComplete 60
# Vérifier l'accès réseau à l'imprimante
Write-Host "`nVérification de l'accès réseau à l'imprimante..." -ForegroundColor Yellow
$pingResult = Test-Connection -ComputerName $printerIP -Count 1 -Quiet
if (-not $pingResult) {
    Show-StyledMessage "L'imprimante n'est pas accessible à l'adresse $printerIP" "Warning"
    Show-StyledMessage "Le pilote sera tout de même installé, mais vérifiez la connectivité réseau." "Warning"
    
    $choice = Read-Host "Voulez-vous continuer malgré l'erreur de connectivité ? (O/N)"
    if ($choice -ne "O" -and $choice -ne "o") {
        Show-StyledMessage "Installation annulée par l'utilisateur" "Warning"
        exit
    }
}

# 1. Ajouter le pilote a la bibliotheque
Write-Host "`nÉtape 4: Installation du pilote dans Windows" -ForegroundColor Yellow
Write-Host "Installation du pilote Toshiba Universal Printer 2..."

try {
    # Installer le pilote d'imprimante avec rundll32
    Write-Host "Installation du pilote avec printui.dll..." -ForegroundColor Yellow
    
    $printui = Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /ia /m `"TOSHIBA Universal Printer 2`" /f `"$driverInfPath`"" -NoNewWindow -Wait -PassThru
    
    if ($printui.ExitCode -ne 0) {
        Show-StyledMessage "ERREUR : L'installation du pilote a echoue avec printui.dll" "Error"
    } else {
        Write-Host "Le pilote a ete installe avec succes via printui.dll" -ForegroundColor Green
    }
    
    # Vérifier que le pilote est bien installé
    $driver = Get-PrinterDriver -Name "TOSHIBA Universal Printer 2" -ErrorAction SilentlyContinue
    if (-not $driver) {
        # Si le pilote n'est pas trouvé, essayer avec pnputil comme solution de secours
        Write-Host "Le pilote n'est pas detecte, tentative avec pnputil..." -ForegroundColor Yellow
        
        $pnputil = Start-Process -FilePath "pnputil" -ArgumentList "/add-driver `"$driverInfPath`" /install" -NoNewWindow -Wait -PassThru
        
        if ($pnputil.ExitCode -ne 0) {
            Show-StyledMessage "ERREUR : L'installation du pilote a echoue avec les deux methodes." "Error"
        }
        
        # Réessayer d'installer le pilote avec printui après pnputil
        $printui = Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /ia /m `"TOSHIBA Universal Printer 2`" /f `"$driverInfPath`"" -NoNewWindow -Wait -PassThru
        
        # Vérification finale
        $driver = Get-PrinterDriver -Name "TOSHIBA Universal Printer 2" -ErrorAction SilentlyContinue
        if (-not $driver) {
            Show-StyledMessage "ERREUR : Impossible d'installer le pilote meme apres plusieurs tentatives." "Error"
        }
    }
} catch {
    Show-StyledMessage "ERREUR lors de l'installation du pilote: $($_.Exception.Message)" "Error"
}

Show-Progress -Activity "Installation du pilote" -PercentComplete 80

# Attendre quelques secondes pour que Windows termine l'installation
Write-Host "`nAttente de la fin de l'installation du pilote..." -ForegroundColor Yellow
Start-Sleep -Seconds 5

# Vérifier les pilotes disponibles
Write-Host "`nPilotes d'imprimante disponibles:"
$toshiba_drivers = Get-PrinterDriver | Where-Object { $_.Name -like "*TOSHIBA*" }
if ($toshiba_drivers) {
    $toshiba_drivers | Format-Table Name, Manufacturer, PrinterEnvironment
} else {
    Write-Host "Aucun pilote Toshiba trouve dans le systeme." -ForegroundColor Red
    Write-Host "Liste de tous les pilotes installes:" -ForegroundColor Yellow
    Get-PrinterDriver | Format-Table Name, Manufacturer, PrinterEnvironment
}

# 2. Verifier si le port existe deja
Write-Host "`nÉtape 5: Creation du port reseau" -ForegroundColor Yellow
$portName = "Port_$printerIP"
$existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue

if ($existingPort) {
    Write-Host "Le port $portName existe deja, il sera reutilise." -ForegroundColor Yellow
} else {
    Write-Host "Creation du port TCP/IP pour l'adresse $printerIP..."
    try {
        Add-PrinterPort -Name $portName -PrinterHostAddress $printerIP -ErrorAction Stop
        Write-Host "Le port reseau a ete cree avec succes." -ForegroundColor Green
    } catch {
        Show-StyledMessage "ERREUR : Impossible de creer le port reseau.`nDetails : $($_.Exception.Message)`nVerifiez que l'adresse IP est correcte et que le reseau est accessible." "Error"
    }
}

# 4. Ajouter l'imprimante
Write-Host "`nÉtape 6: Installation de l'imprimante" -ForegroundColor Yellow
Write-Host "Configuration de l'imprimante $printerName..."
try {
    Add-Printer -Name $printerName -DriverName "TOSHIBA Universal Printer 2" -PortName $portName -ErrorAction Stop
    Write-Host "L'imprimante a ete configuree avec succes." -ForegroundColor Green
} catch {
    Show-StyledMessage "ERREUR : Impossible d'installer l'imprimante.`nDetails : $($_.Exception.Message)" "Error"
}

# Résumé de l'installation avec style
$summary = @"
+------------- Resume de l'Installation -------------+
| Systeme    : $($osInfo.Caption)
| Pilote     : TOSHIBA Universal Printer 2
| IP         : $printerIP
| Port       : $portName
| Etat       : $(if ($driver) { "(+) Installe" } else { "(x) Non installe" })
+------------------------------------------------+
"@

Write-Host "`n$summary" -ForegroundColor Cyan

if ($driver) {
    Show-StyledMessage "Installation terminee avec succes !" "Success"
    Show-StyledMessage "L'imprimante est prete a etre utilisee" "Success"
} else {
    Show-StyledMessage "Des erreurs sont survenues pendant l'installation" "Error"
    Show-StyledMessage "Consultez le fichier README.md pour le depannage" "Info"
}

Show-Progress -Activity "Installation terminee" -PercentComplete 100

# Boucle principale d'installation
do {
    Clear-Host
    Write-Host $logo -ForegroundColor Cyan
    Write-Host "Version: 1.0`nDate: $(Get-Date -Format 'dd/MM/yyyy')`n" -ForegroundColor Gray

    # Demander l'adresse IP de l'imprimante avec validation en temps réel
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

    # Demander le nom de l'imprimante
    Write-Host "`nÉtape 2: Nommage de l'imprimante" -ForegroundColor Yellow
    $printerName = Read-Host "Veuillez entrer le nom souhaite pour l'imprimante"

    # Verifier si une imprimante avec ce nom existe deja
    if (Get-Printer -Name $printerName -ErrorAction SilentlyContinue) {
        Show-StyledMessage "ERREUR : Une imprimante avec le nom '$printerName' existe deja." "Error"
        exit
    }

    # Chemin absolu du fichier INF du pilote
    $driverPath = Join-Path $scriptDirectory "UNI\Driver\64bit"
    $driverInfPath = Join-Path $driverPath "eSf6u.inf"

    Write-Host "`nVerification du chemin du pilote:" -ForegroundColor Yellow
    Write-Host "Repertoire des pilotes: $driverPath"
    Write-Host "Fichier INF: $driverInfPath"

    # Verifier que le dossier des pilotes existe
    if (-not (Test-Path $driverPath)) {
        Show-StyledMessage "ERREUR : Le dossier des pilotes n'existe pas: $driverPath" "Error"
        exit
    }

    # Verifier que le fichier INF existe
    if (-not (Test-Path $driverInfPath)) {
        Show-StyledMessage "ERREUR : Le fichier pilote n'existe pas: $driverInfPath" "Error"
        exit
    }

    # Lister le contenu du dossier des pilotes
    Write-Host "`nFichiers presents dans le dossier des pilotes:" -ForegroundColor Yellow
    Get-ChildItem $driverPath | ForEach-Object {
        Write-Host "  $($_.Name)"
    }

    # Fonction pour vérifier l'environnement Windows
    function Test-WindowsEnvironment {
        Write-Host "`nVérification de l'environnement Windows..." -ForegroundColor Yellow
        
        # Vérifier la version de Windows
        $osInfo = Get-WmiObject -Class Win32_OperatingSystem
        Write-Host "Système d'exploitation : $($osInfo.Caption) $($osInfo.OSArchitecture)"
        
        # Vérifier si nous sommes en 64-bit
        if (-not [Environment]::Is64BitOperatingSystem) {
            Show-StyledMessage "ERREUR : Ce script nécessite un système Windows 64-bit." "Error"
            return $false
        }
        
        # Vérifier l'espace disque disponible (minimum 1 GB recommandé)
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
            Write-Host "Note: Les anciens pilotes ne seront pas supprimés pour éviter les conflits avec d'autres imprimantes." -ForegroundColor Cyan
        } else {
            Write-Host "Aucun ancien pilote Toshiba trouvé." -ForegroundColor Green
        }
    }

    # Vérification de l'environnement avec barre de progression
    Show-Progress -Activity "Vérification de l'environnement" -PercentComplete 20
    if (-not (Test-WindowsEnvironment)) {
        exit 1
    }

    Show-Progress -Activity "Recherche des pilotes existants" -PercentComplete 40
    Remove-OldToshibaDrivers

    Show-Progress -Activity "Test de la connectivité réseau" -PercentComplete 60
    # Vérifier l'accès réseau à l'imprimante
    Write-Host "`nVérification de l'accès réseau à l'imprimante..." -ForegroundColor Yellow
    $pingResult = Test-Connection -ComputerName $printerIP -Count 1 -Quiet
    if (-not $pingResult) {
        Show-StyledMessage "L'imprimante n'est pas accessible à l'adresse $printerIP" "Warning"
        Show-StyledMessage "Le pilote sera tout de même installé, mais vérifiez la connectivité réseau." "Warning"
        
        $choice = Read-Host "Voulez-vous continuer malgré l'erreur de connectivité ? (O/N)"
        if ($choice -ne "O" -and $choice -ne "o") {
            Show-StyledMessage "Installation annulée par l'utilisateur" "Warning"
            exit
        }
    }

    # 1. Ajouter le pilote a la bibliotheque
    Write-Host "`nÉtape 4: Installation du pilote dans Windows" -ForegroundColor Yellow
    Write-Host "Installation du pilote Toshiba Universal Printer 2..."

    try {
        # Installer le pilote d'imprimante avec rundll32
        Write-Host "Installation du pilote avec printui.dll..." -ForegroundColor Yellow
        
        $printui = Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /ia /m `"TOSHIBA Universal Printer 2`" /f `"$driverInfPath`"" -NoNewWindow -Wait -PassThru
        
        if ($printui.ExitCode -ne 0) {
            Show-StyledMessage "ERREUR : L'installation du pilote a echoue avec printui.dll" "Error"
        } else {
            Write-Host "Le pilote a ete installe avec succes via printui.dll" -ForegroundColor Green
        }
        
        # Vérifier que le pilote est bien installé
        $driver = Get-PrinterDriver -Name "TOSHIBA Universal Printer 2" -ErrorAction SilentlyContinue
        if (-not $driver) {
            # Si le pilote n'est pas trouvé, essayer avec pnputil comme solution de secours
            Write-Host "Le pilote n'est pas detecte, tentative avec pnputil..." -ForegroundColor Yellow
            
            $pnputil = Start-Process -FilePath "pnputil" -ArgumentList "/add-driver `"$driverInfPath`" /install" -NoNewWindow -Wait -PassThru
            
            if ($pnputil.ExitCode -ne 0) {
                Show-StyledMessage "ERREUR : L'installation du pilote a echoue avec les deux methodes." "Error"
            }
            
            # Réessayer d'installer le pilote avec printui après pnputil
            $printui = Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /ia /m `"TOSHIBA Universal Printer 2`" /f `"$driverInfPath`"" -NoNewWindow -Wait -PassThru
            
            # Vérification finale
            $driver = Get-PrinterDriver -Name "TOSHIBA Universal Printer 2" -ErrorAction SilentlyContinue
            if (-not $driver) {
                Show-StyledMessage "ERREUR : Impossible d'installer le pilote meme apres plusieurs tentatives." "Error"
            }
        }
    } catch {
        Show-StyledMessage "ERREUR lors de l'installation du pilote: $($_.Exception.Message)" "Error"
    }

    Show-Progress -Activity "Installation du pilote" -PercentComplete 80

    # Attendre quelques secondes pour que Windows termine l'installation
    Write-Host "`nAttente de la fin de l'installation du pilote..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5

    # Vérifier les pilotes disponibles
    Write-Host "`nPilotes d'imprimante disponibles:"
    $toshiba_drivers = Get-PrinterDriver | Where-Object { $_.Name -like "*TOSHIBA*" }
    if ($toshiba_drivers) {
        $toshiba_drivers | Format-Table Name, Manufacturer, PrinterEnvironment
    } else {
        Write-Host "Aucun pilote Toshiba trouve dans le systeme." -ForegroundColor Red
        Write-Host "Liste de tous les pilotes installes:" -ForegroundColor Yellow
        Get-PrinterDriver | Format-Table Name, Manufacturer, PrinterEnvironment
    }

    # 2. Verifier si le port existe deja
    Write-Host "`nÉtape 5: Creation du port reseau" -ForegroundColor Yellow
    $portName = "Port_$printerIP"
    $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue

    if ($existingPort) {
        Write-Host "Le port $portName existe deja, il sera reutilise." -ForegroundColor Yellow
    } else {
        Write-Host "Creation du port TCP/IP pour l'adresse $printerIP..."
        try {
            Add-PrinterPort -Name $portName -PrinterHostAddress $printerIP -ErrorAction Stop
            Write-Host "Le port reseau a ete cree avec succes." -ForegroundColor Green
        } catch {
            Show-StyledMessage "ERREUR : Impossible de creer le port reseau.`nDetails : $($_.Exception.Message)`nVerifiez que l'adresse IP est correcte et que le reseau est accessible." "Error"
        }
    }

    # 4. Ajouter l'imprimante
    Write-Host "`nÉtape 6: Installation de l'imprimante" -ForegroundColor Yellow
    Write-Host "Configuration de l'imprimante $printerName..."
    try {
        Add-Printer -Name $printerName -DriverName "TOSHIBA Universal Printer 2" -PortName $portName -ErrorAction Stop
        Write-Host "L'imprimante a ete configuree avec succes." -ForegroundColor Green
    } catch {
        Show-StyledMessage "ERREUR : Impossible d'installer l'imprimante.`nDetails : $($_.Exception.Message)" "Error"
    }

    # Résumé de l'installation avec style
    $summary = @"
+------------- Resume de l'Installation -------------+
| Systeme    : $($osInfo.Caption)
| Pilote     : TOSHIBA Universal Printer 2
| IP         : $printerIP
| Port       : $portName
| Etat       : $(if ($driver) { "(+) Installe" } else { "(x) Non installe" })
+------------------------------------------------+
"@

    Write-Host "`n$summary" -ForegroundColor Cyan

    if ($driver) {
        Show-StyledMessage "Installation terminee avec succes !" "Success"
        Show-StyledMessage "L'imprimante est prete a etre utilisee" "Success"
    } else {
        Show-StyledMessage "Des erreurs sont survenues pendant l'installation" "Error"
        Show-StyledMessage "Consultez le fichier README.md pour le depannage" "Info"
    }

    Show-Progress -Activity "Installation terminee" -PercentComplete 100

    # Demander si l'utilisateur veut installer une autre imprimante
    Write-Host "`n"
    do {
        $response = Read-Host "Voulez-vous installer une autre imprimante ? (O/N)"
        switch ($response.ToUpper()) {
            "O" {
                Show-StyledMessage "Lancement d'une nouvelle installation..." "Info"
                Write-Host "`n=================================================`n"
                Start-Sleep -Seconds 2  # Petite pause pour la lisibilité
                $continue = $true
                break
            }
            "N" {
                Show-StyledMessage "Merci d'avoir utilise le script d'installation !" "Success"
                $continue = $false
                break
            }
            default {
                Show-StyledMessage "Reponse invalide. Veuillez repondre par O (Oui) ou N (Non)" "Warning"
            }
        }
    } while ($response -notmatch '^[OoNn]$')

} while ($continue)

# Attendre avant de fermer
Write-Host "`nAppuyez sur une touche pour fermer la fenetre..." -ForegroundColor White
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')