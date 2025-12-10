TOSHIBA e-STUDIO-Serie

	Copyright(c) 2003-2022 Toshiba Tec Corporation, Alle Rechte vorbehalten.
			31. August 2022

Dieses Produkt unterliegt den folgenden Begrenzungen und Regeln.  Bitte lesen Sie diese Angaben vor der Verwendung durch.


--------------------------------------------------------------------------------
    Druckertreiber 		Version 7.222.5412.81
--------------------------------------------------------------------------------

Regeln
======

1. Windows

- Wenn mithilfe von "Point-and-Print" ein auf dem Server installierter Druckertreiber auf einem Client-PC installiert wird und auf dem Client-PC der Standardwert auf der Registerkarte "Allgemein" des Druckertreibers unter "Einstellungen" festgelegt wird, wird er nicht zum Standardwert auf dem Client-PC, auch wenn im Druckertreiber auf dem Server "Normale Einstellungen" geändert wird.

- In Windows 8 (32 Bit) werden die Dialogfelder, die beim Drucken eingeblendet werden (z. B. "Kennwort für vertraulichen Druck" und "Abteilungscode") in einem inaktiven Zustand angezeigt. Dies liegt daran, dass die Anwendung in Windows 8 (32 Bit) über WOW64 aufgerufen wird. Bitte machen Sie das angezeigte Dialogfeld aktiv, bevor Sie Einstellungen vornehmen.

- Wenn Bilder im Hochformat und Querformat auf den gleichen Blatt von "Mehrere Seiten pro Blatt" vorhanden sind, wird der Druck nicht durch Drehen eines der Bilder ausgeführt.  Er wird durch Verkleinerung ausgeführt.

- Wenn ein Originaldokument, das Grafiken im Hoch- und Querformat enthält, im Modus ‚2-seitiger Druck' gedruckt wird, kann es vorkommen, dass eine Grafik auf die nächste Seite gedruckt wird, sobald sich deren Ausrichtung ändert.

- Die Hilfe-Meldung erscheint nicht, auch wenn die Schaltfläche "?" oben rechts im Bildschirm geklickt wird.

- Wenn der Druck in einer NetWare-Umgebung ausgeführt wird, kann SNMP-Kommunikation nicht sofort nach der Installation des Druckers ausgeführt werden.  Öffnen Sie Eigenschaften am Drucker und klicken Sie auf das Register Konfiguration, und dann auf die Schaltfläche Jetzt Aktualisieren, um den aktuell verwendeten e-STUDIO zu finden.

- Wenn der PS3-Druckertreiber verwendet wird und der Adobe TypeManager Font den Druck ausführt, wird dies ein Fehler. Verwenden Sie deshalb nicht den Adobe TypeManager Font. Verwenden Sie den PCL-Druckertreiber, wenn Sie den Adobe TypeManager Font verwenden müssen.

- Wenn LPR zum Drucken verwendet wird, deaktivieren Sie das Byte Counting.

- Beim Drucken eines Wasserzeichens kann das Resultat sich je nach Einstellungen von der Erwartung unterscheiden.  In diesem Fall drucken Sie das Waserzeichen erneut nach dem Ändern von Größe und Position von Wasserzeichen.  Es ist wirksam, das Wasserzeichen im Druckbereich zu platzieren.

- Nicht kompatibel mit TOSHIBA Universal Printer, TOSHIBA e-STUDIO Series PS3 und dem Druckertreiber TOSHIBA e-STUDIO Series XPS

Funktionen wie „Profil“, „Wasserzeichen“, „Overlay“ und mit den Druckertreibern TOSHIBA Universal Printer, TOSHIBA e-STUDIO Series PS3, TOSHIBA e-STUDIO Series XPS erstellte exportierte Einstelldateien funktionieren nicht mit diesem Druckertreiber.
Erstellen Sie diese erneut mit diesem Druckertreiber.

- Beschränkungen bei der Verwendung von Profilen
1. <Inkompatibilität zwischen Modellen mit Festplatte (HDD)/Modellen ohne HDD>
1.1. Wenn Sie in einem Modell mit HDD ein Profil laden, welches in einem Modell ohne HDD erstellt wurde, gehen alle Druckaufträge mit Ausnahme des normalen und Overlay-Drucks verloren.
Beenden Sie das Druckertreiber-Programm über die Option Abbrechen und öffnen Sie es zum Fortfahren erneut.

1.2 Wenn Sie in einem Modell ohne HDD ein Profil laden, welches in einem Modell mit HDD erstellt wurde, stehen alle Druckaufträge einschließlich denen, die HDD benötigen, zur Auswahl.
Bitte wählen Sie nicht die Druckaufträge aus, die HDD benötigen. Wenn Sie Druckaufträge senden, die HDD benötigen, werden diese im MFP gelöscht.
Entnehmen Sie Details zu den Druckaufträgen, die HDD benötigen, der Druckanleitung.

2. <Inkompatibilität zwischen Modellen mit automatischer Duplexeinheit (ADU)/Modellen ohne ADU>
2.1. Wenn Sie in einem Modell mit ADU ein Profil laden, welches in einem Modell ohne ADU erstellt wurde, ist das 2-seitige Drucken deaktiviert.
Beenden Sie das Druckertreiber-Programm über die Option Abbrechen und öffnen Sie es zum Fortfahren erneut.

2.2. Wenn Sie in einem Modell ohne ADU ein Profil laden, welches in einem Modell mit ADU erstellt wurde, ist das 2-seitige Drucken aktiviert.
Wählen Sie das 2-seitige Drucken jedoch nicht. Wenn Sie einen 2-seitigen Druckauftrag senden, wird er nur 1-seitig gedruckt.

- Drucken aus einer Anwendung

(1) Übersicht
Obwohl einige Druckoptionen aus Ihrer Anwendung eingestellt werden können, verwenden Sie Ihre Anwendung nicht zur Druckeinstellung. Verwenden Sie Ihren Druckertreiber zum Einstellen von Druckoptionen.
Wenn Druck nach dem Einstellen Ihrer Anwendung ausgeführt wird, können unerwartete Ergebnisse auftreten oder die Verarbeitung kann lange dauern.
Stellen Sie nicht "Kollatieren" in Ihrer Anwendung ein.
Bei der Wahl eines Papierfachs für Ihre Anwendung können zwei "AutoSelect tray"-Anzeigen erscheinen. In diesem Fall wählen Sie das unten gezeigte "AutoSelect tray" (Autowahl-Fach).

(2) Microsoft Excel
Excel behandelt jedes Worksheet als einen Druckauftrag.  Die Einstellung im Druckertreiber-Bildschirm gelten nur für das erste Worksheet.

Wenn Sie mehr als eins in [Kopien] im Drucken-Dialog für Excel eingeben, markieren Sie nicht [Kollatieren].

(3) Adobe Acrobat 
Bei Verwendung von PS3 für den Druck können gedruckte Bilder vergrößert oder verkleinert werden. Durch Verwendung von PCL6 oder PS3 mit den folgenden Druckeinstellungen können auch in diesem Fall Bilder der richtigen Größe ausgedruckt werden.
Öffnen Sie das Druckdialogfeld, um die Einstellungen für PS3 vorzunehmen.
- Wählen Sie [An Papier anpassen] aus dem Menü [Seitenskalierung]. 
- Markieren Sie [Papierquelle nach PDF-Seitengröße wählen]. 
(Beachten Sie, dass die Papiereinstellungen des Treibers nicht gelten, da die Papiergröße der PDF Datei verwendet wird.)

(4) Adobe PageMaker
Bei Privatdruckvorgängen und Abteilungscode-Druckvorgängen unter Verwendung von WindowsPPD für Adobe PageMaker muss immer ein Code angegeben werden. Andernfalls wird eine Fehlernachricht gedruckt. 


- Point & Print vom Windows-Server

 Druck-Einrichtung, die "Standartwerte" des Servers eingestellt ist, kann möglicherweise nicht auf der Client-Seite reflektiert werden, je nach der Kombination von Server-Betriebssystem und Client-Betriebssystem.


- Bluetooth
Einige handelsübliche Bluetooth USB-Adapter können eventuell nicht mit diesem MFP zusammen benutzt werden.
Empfohlene Bluetooth USB-Adapter, entnehmen Sie der Bedienungsanleitung für Bluetooth.


--------------------------------------------------------------------------------
Drucken
--------------------------------------------------------------------------------

Regeln
======

1. Manche Dokumente können bewirken, das der für den Druck reservierte Speicher unzureichend wird.

2. Wenn das Drucken im Farbmodus erfolgt, wird der schwarze Toner zum Drucken der grauen Bereiche verwendet, da "Reines Grau" im Dialogfeld "Einstellung" unter "Bildqualität" als Standardeinstellung des Druckertreibers aktiviert ist. Infolgedessen kann es zu unerwarteter Ausgabe kommen, abhängig von den Druckdaten, (z. B. moiré statt glatter Farbverlauf). Deaktivieren Sie in diesem Fall "Reines Grau."

3. Wenn Sie über den Druckertreiber TOSHIBA Universal Printer 2 einen Druckjob senden (von PC, Tablet, etc.) und “Standardmodell” als Druckername definieren, kann das Druckergebnis von den im Druckertreiber festgelegten Einstellungen/Optionen abweichen. (Die Druckeinstellungen/Optionen wie beispielsweise Farbe/Schwarz & Weiß, Papierquelle, Ziel, Papiergröße/Papiertyp, Heften und Lochung funktionieren möglicherweise nicht wie vorgesehen, was dazu führt, dass der Druckjob in den Standardeinstellungen gedruckt wird)


---------------------------------------------------------
Support-Status der Citrix-Produkte für den Druckvorgang
---------------------------------------------------------

1. Namen der unterstützten Produkte
 - Citrix XenApp 7.6 LTSR
 - Citrix XenApp 7.15 LTSR
 - Citrix Virtual Apps 1912 LTSR
 - Citrix Virtual Apps 2106
 - Citrix Virtual Apps 2109
 - Citrix Virtual Apps 2112
 - Citrix Virtual Apps 2206
 - Citrix Virtual Apps 2203 LTSR

2. Getestete Betriebsumgebungen
 - Server OS : Windows Server 2019 (Virtual Apps 2112)
               Windows Server 2012 R2 (Virtual Apps 1912 LTSR)
               Windows Server 2022 (Virtual Apps 2203 LTSR)
 - Client OS	: Windows 11, Windows 10, Windows 8.1


--------------------------------------------------------------------------------
Warenzeichen
--------------------------------------------------------------------------------

- Der offizielle Name von Windows Server 2012 ist "Betriebssystem Microsoft Windows Server 2012".
- Der offizielle Name von Windows Server 2019 ist "Betriebssystem Microsoft Windows Server 2019".
- Der offizielle Name von Windows Server 2022 ist "Betriebssystem Microsoft Windows Server 2022".
- Der offizielle Name von Windows 8 ist "Betriebssystem Microsoft Windows 8".
- Der offizielle Name von Windows 10 ist "Betriebssystem Microsoft Windows 10".
- Der offizielle Name von Windows 11 ist "Betriebssystem Microsoft Windows 11".
- Microsoft, Windows und die Warenzeichen und Produktnamen anderer Microsoft-Produkte sind Warenzeichen der Microsoft Corporation in den USA und anderen Ländern.
- PostScript und die Warenzeichen und Produktnamen anderer Adobe Systems-Produkte sind Warenzeichen der Adobe Systems Incorporated in den USA und anderen Ländern.
- Novell, NetWare und NDS sind Warenzeichen von Novell, Inc.
- Citrix, Citrix XenApp, Citrix Virtual Apps sind Marken von Citrix Systems, Inc. und/oder einem oder mehreren Tochterunternehmen und sind möglicherweise beim Patent- und Markenamt der Vereinigten Staaten und in anderen Ländern eingetragen.
- Bluetooth(R) ist eine registrierte Handelsmarke und Eigentum von Bluetooth SIG, lnc.
- Etwaige bestehende Warenzeichen oder sonstige gewerbliche Schutzrechte an anderen hier zitierten Bezeichnungen erkennen wir an. 
