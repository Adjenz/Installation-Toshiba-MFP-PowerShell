TOSHIBA e-STUDIO Serie

	Copyright(c) 2003-2022 Toshiba Tec Corporation, Ogni diritto riservato
			31 agosto 2022

Questo prodotto presenta le restrizioni e problemi seguenti. Vi raccomandiamo di leggere attentamente questo manuale prima di utilizzare il prodotto.


--------------------------------------------------------------------------------
    Driver di stampa 		Versione 7.222.5412.81
--------------------------------------------------------------------------------

Problemi
========

1. Windows

- Quando il driver della stampante installato su un server viene installato su un client PC utilizzando "Seleziona e stampa" e il valore predefinito è impostato su [Impostazioni] nella scheda [Generale] nel driver della stampante nel client PC, questo non può essere applicato sul client PC anche se le [Impostazioni normali] vengono modificate nel driver della stampante del server.

- In Windows 8 a 32 bit, le finestre di dialogo sono visualizzate durante la stampa, ad esempio: "Password Stampa riservata" e "Codice reparto" sono visualizzate con stato inattivo. Questo si verifica perché l'applicazione viene invocata tramite WOW64 in Windows 8 a 32 bit. Rendere attiva la finestra di dialogo visualizzata prima dell'impostazione.

- Quando l'immagine Ritratto (Portrait) e Paesaggio (Landscape) esistono sullo stesso foglio "Pagine multiple per foglio", la stampa non sarà effettuata girando l'immagine dell'uno o dell'altro. Sarà effettuata con riduzione.

- eCopy sowie die Marken- und Produktnamen eCopy-zugehöriger Produkte sind in den Vereinigten Staaten und anderen Ländern Marken oder eingetragene Marken von eCopy.

- Il messaggio d'aiuto non appare anche quando si fa clic sul tasto "?"  sulla parte destra superiore dello schermo.

- In caso di stampa in ambiente NetWare non è possibile eseguire la comunicazione SNMP immediatamente dopo avere installato il driver di stampa.  Aprite la finestra Proprietà del driver di stampa, cliccate sul selettore Configurazione e quindi sul pulsante Aggiorna adesso in modo da individuare l'e-STUDIO attualmente in uso.

- Quando il driver della stampante PS3 è utilizzato ed Adobe Type Manager Font effettua la stampa, questo presenterà un errore. In questo caso, non utilizzare Adobe Type Manager Font. Utilizzare il driver di stampante PCL quando si deve utilizzare Adobe Type Manager Font.

- In caso di stampa con LPR, disabilitate il conteggio di byte.

- In caso di stampa di una Filigrana, in funzione delle impostazioni eseguite il risultato potrebbe differire da quello atteso. In tal caso stampate nuovamente la Filigrana dopo avere modificato in Filigrana il formato e la posizione.  Si rivela efficace posizionare la Filigrana nell'area stampabile.

- Incompatibilità con TOSHIBA Universal Printer, TOSHIBA e-STUDIO Series PS3 e un driver di stampa TOSHIBA e-STUDIO Series XPS

Le funzioni come "Profilo", "Filigrana", "Sovrapposizione" e i file di impostazione esportati creati con i driver di stampa di TOSHIBA Universal Printer, TOSHIBA e-STUDIO Series PS3, TOSHIBA e-STUDIO Series XPS non sono operativi con questo driver.
È necessario ricrearli mediante questo driver.

- Limitazione per l'uso della funzione Profilo
1. <Incompatibilità tra il modello dotato di disco rigido (HDD) e il modello privo di disco rigido>
1.1. Caricando in un modello dotato di disco rigido un Profilo creato mediante un modello privo di disco rigido, non saranno visibili i lavori di stampa, ad eccezione della stampa normale e della stampa in sovrapposizione.
Premere Annulla per chiudere l'interfaccia grafica del driver di stampa, quindi riaprire per continuare.

1.2 Caricando un Profilo creato mediante un modello dotato di disco rigido in un modello privo di disco rigido, saranno selezionabili tutti i lavori di stampa, compresi quelli che richiedono il disco rigido.
Non selezionare i lavori di stampa che richiedono il disco rigido. Se si invia un lavoro di stampa che richiede il disco rigido, tale lavoro verrà eliminato nell'MFP.
Per i dettagli sui lavori di stampa che richiedono il disco rigido, fare riferimento alla Guida alla stampa.

2. <Incompatibilità tra i modelli dotati di modulo fronte/retro automatico (ADU) e i modelli che ne sono privi>
2.1. Caricando in un modello dotato di modulo fronte/retro automatico un Profilo creato mediante un modello privo di tale modulo, verrà disabilitata la stampa fronte/retro.
Premere Annulla per chiudere l'interfaccia grafica del driver di stampa, quindi riaprire per continuare.

2.2. Caricando in un modello privo di modulo fronte/retro automatico un Profilo creato mediante un modello dotato di tale modulo, verrà abilitata la stampa fronte/retro.
Non selezionare Stampa fronte/retro. Se si invia un lavoro di stampa fronte/retro, quest'ultimo verrà stampato su lato singolo.

- Stampa a partire da un'applicazione

(1) Veduta d'insieme	
Benché alcune opzioni di stampa possano essere impostate a partire dalla vostra applicazione, non utilizzare la vostra applicazione per impostare la stampa. Utilizzare il vostro driver di stampante per impostare delle opzioni di stampante.
Quando la stampa viene effettuata dopo impostazione della vostra applicazione, si potrebbe ritrovarsi con un risultato inatteso o ciò potrà necessitare un pò di tempo per il trattamento.
Non impostare "Collate" (Raccogliere) della vostra applicazione.
Quando si imposta uno scomparto di carta della vostra applicazione, due "Auto Select tray" possono essere visualizzati. In questo caso, scegliere "Auto Select tray" visualizzato sotto.

(2) Microsoft Excel
Excel gestisce ciascun foglio di lavoro come un'unica operazione di stampa.  Le impostazioni eseguite nello schermo del driver di stampa sono pertanto valide solamente per il primo foglio di lavoro.

Quando nel campo [Copie] del dialogo di stampa di Excel si inserisce un valore superiore a 1, non si deve selezionare la casella [Fascicola].

(3) Adobe Acrobat 
Se si stampa con il PS3, l'immagine può apparire ingrandita oppure ridotta.  In tal caso, stampando con il PCL6 oppure con il PS3 usando le impostazioni qui di seguito indicate, l'immagine può essere stampata nel formato corretto.
Aprite il dialogo Stampa in modo da eseguire l'impostazione del PS3.
- Dal menu [Ridimensionamento pagina] selezionate [Adatta alla carta].
- Selezionate la casella [Scegli alimentazione da dimensioni pagina PDF].
(Vi preghiamo di notare che l'impostazione della carta nel driver non è valida poiché si utilizza il formato del file PDF.)

(4) Adobe PageMaker
Per la stampa personale e la stampa del codice reparto con WIndowsPPD per Adobe PageMaker, è necessario specificare sempre il codice.  Altrimenti, verrà stampato un foglio d'errore. 


- Point and print dal Server Windows

 La messa a punto della stampa che è stata regolata in "Predefinite)" del Server non può riflettere al lato del cliente secondo la combinazione del SO (Sistema Operativo) del Server e del SO del cliente.


- Bluetooth
Alcuni adattatori di Bluetooth USB disponibili in commercio non possono essere utilizzati con questo MFP.
Per gli adattatori di Bluetooth USB disponibili, fare riferimento al manuale operativo per il Bluetooth.


--------------------------------------------------------------------------------
Stampa
--------------------------------------------------------------------------------

Problemi
========

1. Alcuni documenti potrebbero rendere insufficiente la memoria destinata alle operazioni di stampa.

2. Quando si esegue la stampa in modalità a colori, il toner nero viene utilizzato per stampare le aree grigie, poiché la funzione "Grigio puro" è impostata su ON nella finestra di dialogo "Impostazione" sotto "Qualità immagine" come impostazione predefinita del driver di stampa. Di conseguenza, la stampa prodotta potrebbe presentare un risultato imprevisto a seconda dei dati da stampare (ad esempio, un effetto moiré anziché una gradazione omogenea). In questo caso, impostare "Grigio puro" su OFF.

3. Quando un processo di stampa viene inviato dal proprio dispositivo (PC, tablet, ecc.) utilizzando il driver di stampa TOSHIBA Universal Printer 2 della stampante universale TOSHIBA Universal 2 e denominato "Modello generico" come nome della stampante, l'output può differire dalle impostazioni/opzioni di stampa specificate nel driver di stampa. (Per esempio, impostazioni/opzioni di stampa come Colore/Bianco/Nero, Alimentazione, Destinazione, Dimensioni carta/Tipi di carta, Pinzatura e Foratura potrebbe non funzionare come specificato in modo che il lavoro venga stampato utilizzando le impostazioni predefinite.)



---------------------------------------------------------
Stato del supporto dei prodotti Citrix per la stampa
---------------------------------------------------------

1. Nomi dei prodotti supportati
 - Citrix XenApp 7.6 LTSR
 - Citrix XenApp 7.15 LTSR
 - Citrix Virtual Apps 1912 LTSR
 - Citrix Virtual Apps 2106
 - Citrix Virtual Apps 2109
 - Citrix Virtual Apps 2112
 - Citrix Virtual Apps 2206
 - Citrix Virtual Apps 2203 LTSR

2. Ambienti operativi verificati
 - Server OS : Windows Server 2019 (Virtual Apps 2112)
               Windows Server 2012 R2 (Virtual Apps 1912 LTSR)
               Windows Server 2022 (Virtual Apps 2203 LTSR)
 - Client OS	: Windows 11, Windows 10, Windows 8.1


--------------------------------------------------------------------------------
Marca depositata
--------------------------------------------------------------------------------

- Il nome ufficiale di Windows Server 2012 è Sistema operativo Microsoft Windows Server 2012.
- Il nome ufficiale di Windows Server 2019 è Sistema operativo Microsoft Windows Server 2019.
- Il nome ufficiale di Windows Server 2022 è Sistema operativo Microsoft Windows Server 2022.
- Il nome ufficiale di Windows 8 è Sistema operativo Microsoft Windows 8.
- Il nome ufficiale di Windows 10 è Sistema operativo Microsoft Windows 10.
- Il nome ufficiale di Windows 11 è Sistema operativo Microsoft Windows 11.
- Microsoft, Windows ed i marchi di fabbrica e di altri prodotti Microsoft rappresentano le marche depositate di Microsoft Corporation negli Stati Uniti ed altri paesi.	
- PostScript ed i marchi di fabbrica e di altri prodotti Adobe Systems rappresentano le marche depositate di Adobe Systems Incorporated negli Stati Uniti ed altri paesi.	
- Novell, NetWare, e NDS sono le marche depositate di Novell, Inc.
- Citrix, Citrix XenApp, Citrix Virtual Apps are trademarks of Citrix Systems, Inc. and/or one or more of its subsidiaries, and may be registered in the United States Patent and Trademark Office and in other countries.
- Bluetooth(R) est une marque déposée de la Bluetooth SIG, lnc.
- Gli altri nomi di società e nomi di prodotti citati in questo manuale sono le marche depositate delle loro società rispettive.
