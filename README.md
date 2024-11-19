# Script d'Installation Imprimante Toshiba Universal Printer 2

Ce script PowerShell automatise l'installation d'imprimantes réseau Toshiba utilisant le pilote Universal Printer 2. Il offre une interface interactive simple et sécurisée pour installer une ou plusieurs imprimantes en une seule session.

## Table des matières
- [Prérequis](#prérequis)
- [Installation rapide](#installation-rapide)
- [Guide détaillé](#guide-détaillé)
- [Fonctionnalités](#fonctionnalités)
- [Structure des fichiers](#structure-des-fichiers)
- [Guide de dépannage](#guide-de-dépannage)
- [FAQ](#faq)
- [Support technique](#support-technique)

## Prérequis

### Configuration système requise
- Système d'exploitation : Windows 64-bit
- Mémoire : 2 GB RAM minimum
- Espace disque : 1 GB minimum disponible
- Réseau : Connexion active avec accès à l'imprimante

### Droits et accès
- Droits administrateur Windows
- Accès réseau à l'imprimante (ports standards d'impression ouverts)
- Pare-feu Windows configuré pour autoriser l'impression

### Fichiers nécessaires
- Script `Installation.ps1`
- Dossier `UNI/Driver/64bit` contenant les pilotes Toshiba
- Fichier principal du pilote : `eSf6u.inf`

## Installation rapide

1. Clonez ou téléchargez le dépôt
2. Clic droit sur `Installation.ps1` > "Exécuter avec PowerShell"
3. Suivez les instructions à l'écran

## Guide détaillé

### Préparation
1. Vérifiez que vous avez tous les prérequis
2. Notez l'adresse IP de votre imprimante
3. Choisissez un nom unique pour l'imprimante
4. Fermez toutes les applications d'impression ouvertes

### Installation pas à pas
1. **Lancement du script**
   - Clic droit sur `Installation.ps1`
   - Sélectionnez "Exécuter avec PowerShell"
   - Confirmez l'élévation des privilèges

2. **Configuration de l'imprimante**
   - Entrez l'adresse IP de l'imprimante
   - Spécifiez le nom souhaité pour l'imprimante
   - Attendez la validation de la connectivité

3. **Installation du pilote**
   - Le script installe automatiquement le pilote
   - Vérifie la présence de pilotes existants
   - Configure le port réseau

4. **Vérification**
   - Consultez le résumé d'installation
   - Vérifiez que l'imprimante apparaît dans Windows
   - Optionnel : Installez une autre imprimante

## Fonctionnalités

### Fonctionnalités principales
- Installation automatisée du pilote
- Configuration du port réseau TCP/IP
- Support de plusieurs installations consécutives
- Interface utilisateur interactive

### Sécurité et validation
- Vérification des droits administrateur
- Validation de l'adresse IP
- Test de connectivité réseau
- Vérification de l'unicité du nom d'imprimante

### Interface utilisateur
- Messages colorés et informatifs
- Barre de progression
- Résumé détaillé de l'installation
- Option de réinstallation simplifiée

## Structure des fichiers

```
Toshiba/
├── Installation.ps1     # Script principal d'installation
├── README.md           # Documentation complète
└── UNI/
    └── Driver/
        └── 64bit/      # Pilotes Windows 64-bit
            ├── eSf6u.inf   # Fichier principal du pilote
            └── [autres fichiers du pilote]
```

## Guide de dépannage

### Problèmes courants et solutions

#### Le script ne démarre pas
- **Symptôme** : Message d'erreur PowerShell
- **Solution** : 
  1. Exécutez PowerShell en tant qu'administrateur
  2. Vérifiez la politique d'exécution : `Get-ExecutionPolicy`
  3. Si nécessaire : `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process`

#### Erreur de connexion réseau
- **Symptôme** : "L'imprimante n'est pas accessible"
- **Solution** :
  1. Vérifiez l'adresse IP (`ping [adresse_ip]`)
  2. Contrôlez la connexion réseau
  3. Vérifiez les pare-feu

#### Échec d'installation du pilote
- **Symptôme** : "Impossible d'installer le pilote"
- **Solution** :
  1. Vérifiez les fichiers du pilote
  2. Consultez les journaux Windows
  3. Désinstallez les anciens pilotes

### Codes d'erreur communs
- **Error 0x0000000A** : Pilote manquant
- **Error 0x00000057** : Paramètre invalide
- **Error 0x00000002** : Fichier non trouvé

## FAQ

**Q: Puis-je installer plusieurs imprimantes ?**
R: Oui, le script propose cette option à la fin de chaque installation.

**Q: Le script est-il compatible Windows 32-bit ?**
R: Non, uniquement Windows 64-bit est supporté.

**Q: Comment changer le nom après installation ?**
R: Utilisez le gestionnaire d'impression Windows.

## Support technique

### Ressources
- Documentation Toshiba : [lien]
- Support Windows : [lien]

### Journaux et diagnostics
1. Journaux d'événements Windows
   - `Event Viewer > Applications`
   - Filtrer par "Print Service"

2. Fichiers de diagnostic
   - Résumé d'installation
   - Logs PowerShell : `Get-EventLog -LogName Application`

### Contact
Pour le support technique :
1. Consultez d'abord ce guide
2. Vérifiez les journaux système
3. Contactez votre administrateur système
