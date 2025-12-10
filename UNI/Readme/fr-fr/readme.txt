TOSHIBA e-STUDIO Série

	Copyright(c) 2003-2022 Toshiba Tec Corporation, Tous droits réservés
			31 août, 2022

Ce produit présente les restrictions et problèmes suivants.  Nous vous recommandons de lire attentivement ce manuel avant l'utilisation du produit.


--------------------------------------------------------------------------------
    Pilote d'imprimante 	     		Version 7.222.5412.81
--------------------------------------------------------------------------------

Certains faits
==============

1. Windows

- Lorsqu'un pilote d'imprimante installé sur un serveur est installé sur un ordinateur client à l'aide de "Pointer et imprimer" et que la valeur par défaut est réglée sur [Paramètres] sur l'onglet [Général] dans le pilote d'imprimante de l'ordinateur client, il ne peut pas être le pilote par défaut dans l'ordinateur client même si l'option [Paramètres normaux] est modifié dans le pilote d'imprimante du serveur.

- Sous Windows 8 32 bits, les boîtes de dialogue apparaissaient à l'impression, par ex., "Mot de passe d'impression privée" et "Code départemental", s'affichent à l'état inactif. La raison : l'application est appelée via WOW64 sous Windows 8 32 bits. Veuillez vous assurer d'activer la boîte de dialogue avant le paramètre.

- Lorsque l'image Portrait et Paysage (Landscape) existent sur la même feuille "Plusieurs pages par feuille" , l'impression ne sera pas effectuée en tournant l'image de l'une ou de l'autre.  Elle sera exécutée par réduction.

- Lorsqu'une feuille de papier original contenant des images Portrait et Paysage est imprimée en mode Impression recto-verso, il se peut qu'une image s'imprime sur la page suivante lorsque l'orientation de l'image change.

- Le message d'aide n'apparaît pas même lorsque l'on fait clic sur le bouton "?" sur la partie droite supérieure de l'écran.

- Quand l'impression est exécutée dans un environnement NetWare, la communication SNMP ne peut pas être exécutée juste après que le pilote de l'imprimante est installé. Ouvrir Propriété du pilote de l'imprimante et faire clic sur l'onglet de Configuration, puis sur la touche Mettre à jour maintenant pour localiser e-STUDIO en usage courant.

- Lorsque le pilote d'imprimante PS3 est utilisé et Adobe TypeManager Font exécute l'impression, ceci présentera une erreur. Dans ce cas, ne pas utiliser Adobe TypeManager Font.  Utiliser le pilote d'imprimante PCL lorsque on doit utiliser Adobe TypeManager Font.

- Lorsque LPR est utilisé pour imprimer, invalider Byte Counting.

- Quand on imprime un Filigrane, le résultat pourrait ne pas être comme prévu selon les réglages. Dans ce cas, imprimer de nouveau le Filigrane après avoir changé la dimension et la position du Filigrane.  Il est efficace de positionner le Filigrane dans la zone imprimable.

- Incompatible avec TOSHIBA Universal Printer, TOSHIBA e-STUDIO Series PS3, et un pilote d'imprimante TOSHIBA e-STUDIO Series XPS

Les caracteristiques telles que "Profil", "Tatouage numerique", "Recouvrement", les fichiers de parametrage exportes crees avec les pilotes d'imprimante TOSHIBA Universal Printer, TOSHIBA e-STUDIO Series PS3, TOSHIBA e-STUDIO Series XPS ne fonctionnent pas avec ce pilote d'imprimante.
Recreez-les avec ce pilote d'imprimante.

- Limitation lors de l'utilisation de Profil
1. <Incompatibilité entre le modèle équipé de disque dur (HDD) et le modèle non équipé de disque dur>
1.1. Lorsque vous chargez un profil créé avec un modèle non équipé de disque dur sur un modèle équipé de disque dur, les travaux d'impression, excepté l'impression normale et l'impression par recouvrement, disparaissent.
Fermez le pilote d'imprimante GUI via Annuler et rouvrez-le pour continuer.

1.2 Lorsque vous chargez un profil créé par un modèle équipé de disque dur sur un modèle non équipé de disque dur, tous les travaux d'impression, y compris ceux requerrant un disque dur, peuvent être sélectionnés.
Ne sélectionnez pas les travaux d'impression requerrant un disque dur. Si vous envoyez un travail d'impression requerrant un disque dur, le travail sera supprimé dans le périphérique multifonction.
Pour plus d'informations sur les travaux d'impression requerrant un disque dur, reportez-vous au manuel d'impression.

2. <Incompatibilité entre le modèle équipé d'appareil de duplexage automatique (ADU) et le modèle non équipé d'appareil de duplexage automatique>
2.1. Lorsque vous chargez un profil créé par un modèle non équipé d'ADU sur un modèle équipé d'ADU, l'impression sur les deux faces est désactivée.
Fermez le pilote d'imprimante GUI via Annuler et le rouvrir pour continuer.

2.2. Lorsque vous chargez un profil créé par un modèle équipé d'ADU sur un modèle non équipé d'ADU, l'impression sur les deux faces est activée.
Ne sélectionnez pas l'impression sur les deux faces. Si vous envoyez un travail d'impression sur les deux faces, il sera imprimé sur une seule face.

- Impression à partir d'une application

(1) Vue d'ensemble
Bien que certaines options d'impression puissent être réglées à partir de votre application, ne pas utiliser votre application pour régler l'impression.  Utiliser votre pilote d'imprimante pour régler des options d'imprimante.
Lorsque l'impression est effectuée après réglage de votre application, on pourrait se retrouvez avec un résultat inattendu ou cela pourra prendre un bon moment pour le traitement.
Ne pas régler "Collate" (Recueillir) de votre application.
Lorsque l'on règle un magasin de papier de votre application, deux "Auto Select tray" peuvent être affichés.  Dans ce cas, sélectionner "AutoSelect tray" affiché dessous.

(2) Microsoft Excel
Excel traite chaque  feuille de travail comme une tâche d'impression. Les réglages sur l'écran du pilote d'imprimante s'appliquent seulement à la première feuille de travail.

Quand on entre plus d'une dans [Copies] sur le dialogue d'impression pour Excel, ne pas cocher [Assembler].

(3) Adobe Acrobat 
Lors de l'impression, si PS3 est utilisé, il peut se peut que l'image imprimée soit agrandie ou réduite. Dans ce cas, l'utilisation de PCL6 or bien encore de PS3 avec les configurations suivantes permet une impression de dimensions correctes.
Ouvrir la boîte de dialogue "impression" pour réaliser les réglages de PS3.
- Sélectionner [Adapter au papier] à partir du menu [Mise à l'échelle de la page].
- Cocher [Choisir source de papier par format de page PDF].
(Noter que la configuration du papier du pilote est invalide tant le format papier du fichier pdf est utilisé.)

(4) Adobe PageMaker
Pour l'impression privée et l'impression de code de département avec WIndowsPPD pour Adobe PageMaker, spécifiez toujours un code.  Autrement une feuille d'erreur sera imprimée. 


- Point and print à partir du Serveur Windows

L'installation d'impression qui a été réglé dans "Impression par défaut" du serveur ne peut pas réfléter au côté du client selon la combinaison du SE (Système d'Exploitation) du serveur et du SE du client. 


- Bluetooth
Certains adaptateurs Bluetooth USB disponibles dans le commerce risquent de ne pas pouvoir s'utiliser avec ce MFP.
Pour en savoir plus sur les adaptateurs Bluetooth USB recommandés, consultez le manuel d'utilisation de Bluetooth.


--------------------------------------------------------------------------------
Impression
--------------------------------------------------------------------------------

Certains faits
==============

1. Certains documents pourraient causer une insuffisance d'impression de la mémoire affectée à l'impression.

2. Si vous imprimez en mode couleur, l'encre noire est utilisée pour imprimer les zones grises puisque "Gris intégral" est réglé sur ON dans la boîte de dialogue "Paramétrage" sous "Qualité d'image" comme réglage par défaut du pilote d'imprimante. Par conséquent, il se peut que le résultat de l'impression ne soit pas celui escompté avec certaines données (par exemple du moiré au lieu d'une gradation douce). Dans ce cas, réglez "Gris intégral" sur OFF.

3. Lorsqu'un travail d'impression est envoyé depuis votre appareil (PC, tablette, etc.) à l'aide du pilote d'imprimante TOSHIBA Universal Printer 2 et désignant "Modèle générique" comme nom d'imprimante, les documents en sortie peuvent différer des paramètres / options spécifiés dans le pilote d'impression. (Par exemple, les paramètres / options d'impression tels que Couleur / Noir et Blanc, Source papier, Destination, Format papier / Types de papier, Agrafage et Perforation peuvent ne pas fonctionner comme spécifié, ce qui entraîne l'impression des travaux avec des paramètres par défaut).


---------------------------------------------------------
État de prise en charge d'appareil Citrix pour impression
---------------------------------------------------------

1. Noms des appareils pris en charge
 - Citrix XenApp 7.6 LTSR
 - Citrix XenApp 7.15 LTSR
 - Citrix Virtual Apps 1912 LTSR
 - Citrix Virtual Apps 2106
 - Citrix Virtual Apps 2109
 - Citrix Virtual Apps 2112
 - Citrix Virtual Apps 2206
 - Citrix Virtual Apps 2203 LTSR

2. Environnements d'exploitation confirmés
 - Server OS : Windows Server 2019 (Virtual Apps 2112)
               Windows Server 2012 R2 (Virtual Apps 1912 LTSR)
               Windows Server 2022 (Virtual Apps 2203 LTSR)
 - Client OS	: Windows 11, Windows 10, Windows 8.1


--------------------------------------------------------------------------------
Marques déposées
--------------------------------------------------------------------------------

- Le nom officiel de Windows Server 2012 est Système d'exploitation Microsoft Windows Server 2012.
- Le nom officiel de Windows Server 2019 est Système d'exploitation Microsoft Windows Server 2019.
- Le nom officiel de Windows Server 2022 est Système d'exploitation Microsoft Windows Server 2022.
- Le nom officiel de Windows 8 est Système d'exploitation Microsoft Windows 8.
- Le nom officiel de Windows 10 est Système d'exploitation Microsoft Windows 10.
- Le nom officiel de Windows 11 est Système d'exploitation Microsoft Windows 11.
- Microsoft, Windows et les marques de fabrique et d'autres produits Microsoft représentent les marques déposées de Microsoft Corporation aux États-Unis et autres pays.
- PostScript et les marques de fabrique et d'autres produits Adobe Systems représentent les marques déposées de Adobe Systems Incorporated aux États-Unis et autres pays.
- Novell, NetWare, et NDS sont les marques déposées de Novell, Inc.
- Citrix, Citrix XenApp, Citrix Virtual Apps sont des marques commerciales de Citrix Systems, Inc. et/ou d’une ou de plusieurs de ses filiales, et peuvent être déposées auprès du Patent and Trademark Office des Etats-Unis et dans d’autres pays.
- Bluetooth(R) est une marque déposée de la Bluetooth SIG, lnc.
- Les autres noms de compagnies et noms de produits mentionnés dans ce manuel sont les marques déposées de leurs compagnies respectives.
