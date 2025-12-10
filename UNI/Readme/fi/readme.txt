TOSHIBA e-STUDIO Series

	Copyright(c) 2003-2022 Toshiba Tec Corporation, All Rights Reserved.
			August 31 2022

This product has the following restrictions and issues.  Please be sure to read this before using.


--------------------------------------------------------------------------------
    Printer Driver		Version 7.222.5412.81
--------------------------------------------------------------------------------

Issues
======

1. Windows

- When a printer driver installed in a server is installed in a client PC using "Point and Print" and the default value is set in [Settings] on the [General] tab in the printer driver of the client PC, it cannot be reflected as a default in the client PC even if [Normal Settings] is changed in the printer driver of the server.

- In Windows 8 32bit, the dialogs popped up at printing, e.g. "Private Print Password" and "Department Code", are displayed in inactive state.  This is because the application is called via WOW64 in Windows 8 32 bit.  Please make the displayed dialog active before the setting.

- When Portrait and Landscape images exist on the same sheet in "Multiple Pages per Sheet" printing, printing will be performed by reducing, not by rotating the images.

- When an original paper, which contains Portrait and Landscape images, is printed in 2-Sided Printing mode, an image may be printed on the next page when an orientation of the image changes.

- Help message does not appear even when the "?" button on the upper right of the screen is clicked.

- When printing is performed in a NetWare environment, SNMP communication cannot be performed immediately after the printer driver is installed.  Open the Property of the printer driver and click on the Configuration tab, then Update Now button to find e-STUDIO currently being used.

- When the PS3 Printer Driver is used with Adobe TypeManager Font, an error will result.  So do not use Adobe TypeManager Font.  Use the PCL Printer Driver when you need to use Adobe TypeManager Font.

- When LPR is used for printing, disable Byte Counting.

- When printing a Watermark, the result may not be as you had expected depending on the settings.  In such a case, print the Watermark again after changing the size and position of the Watermark.  It is effective to position the Watermark in the printable area.

- Incompatibility with TOSHIBA Universal Printer, TOSHIBA e-STUDIO Series PS3, and a TOSHIBA e-STUDIO Series XPS printer driver
 Features like "Profile", "Watermark", "Overlay", exported setting files created with TOSHIBA Universal Printer, TOSHIBA e-STUDIO Series PS3, TOSHIBA e-STUDIO Series XPS printer drivers will not work with this version's driver. 
 Please recreate them with this version's driver.

- Limitation when using Profile
1. <Incompatibility between Harddisk (HDD) equipped model/HDD non-equipped model>
  1.1. When you will load a Profile created by HDD non-equipped model in HDD equipped model, Print Jobs except Normal Print and Print to Overlay will disappear.
  Please close printer driver GUI by Cancel and re-open to continue.

  1.2 When you will load a Profile created by HDD equipped model in HDD non-equipped model, all Print Jobs including that requires HDD will be selectable.
  Please do not select those Print Jobs that requires HDD. If you send Print Jobs that requires HDD, the job will be deleted in the MFP.
  For details about the Print Jobs that requires HDD, please refer to the Printing Guide.

2. <Incompatibility between Automatic Duplexing Unit (ADU) equipped model/ADU non-equipped model>
  2.1. When you will load a Profile created by ADU non-equipped model in ADU equipped model, 2-Sided Printing will be disabled.
  Please close printer driver GUI by Cancel and re-open to continue.

  2.2. When you will load a Profile created by ADU equipped model in ADU non-equipped model, 2-Sided Printing will be enabled.
  Please do not select 2-Sided Printing. If you send 2-Sided Print Job, it will be printed in 1-Side.


- Printing from an application

(1) Overview
Although some printing options can be set from your application, do not use your application to set printing. Use your printer driver to set printer options.
When printing is performed after setting such options in your application, you may end up with an unexpected result or it may take a long time to process.
Do not set "Collate" in your application.
When setting a paper tray in your application, two "AutoSelect tray"s may be displayed. In this case, select the lower "AutoSelect tray."

(2) Microsoft Excel
Excel handles each worksheet as one print job.  The settings on the printer driver screen apply to the first worksheet only.

When you enter more than one in [Copies] on the Print dialog for Excel, do not check [Collate].

(3) Adobe Acrobat 
When PS3 is used for printing, a printed image may be enlarged or reduced.  In such cases, using PCL6 or using PS3 with the following setting can print the image at a proper size.
Open the Print dialog box to perform the setting for PS3.
- Select [Fit to Paper] form from the [Page Scaling] menu.
- Check [Choose Paper Source by PDF page size].
(Note that the driver's paper setting is not valid as the PDF file's paper size is used.)

(4) Adobe PageMaker
For Private Print and Department Code Print using WIndowsPPD for Adobe PageMaker, always specify a code.  If not, an error sheet will be printed. 


- Point and print from the Windows Server

Printing setup that is set in "Printing Defaults" of the server cannot be reflected to the client side depending on the combination of the server OS and the client OS.


- Bluetooth
Some commercially available Bluetooth USB adapters may not be used with this MFP.
For recommended Bluetooth USB adapters, refer to the operation manual For Bluetooth.


--------------------------------------------------------------------------------
Print
--------------------------------------------------------------------------------

Issues
======

1. Some documents may cause print-assigned memory to become insufficient to print.

2. When printing is performed in color mode, the black toner is used to print gray areas, because "Pure Gray" is set to ON in the "Setting" dialog under "Image Quality" as the default setting of the printer driver. As a result, an unexpected output may be provided depending on print data (for example, moire instead of a smooth gradation). In this case, set "Pure Gray" to OFF.

3. When a print job is sent from your device (PC, tablet, etc.) using the TOSHIBA Universal Printer 2 printer driver and designating "Generic Model" as a printer name, the output may differ from the print settings/options specified in the printer driver. (For example, print settings/options such as Color/Black and White, Paper Source, Destination, Paper Size/Paper Types, Staple and Hole Punch may not work as specified resulting in the job being printed using the default settings.)


---------------------------------------------------------
Citrix Product Support Status for Printing
---------------------------------------------------------

1. Supported Product Names
 - Citrix XenApp 7.6 LTSR
 - Citrix XenApp 7.15 LTSR
 - Citrix Virtual Apps 1912 LTSR
 - Citrix Virtual Apps 2106
 - Citrix Virtual Apps 2109
 - Citrix Virtual Apps 2112
 - Citrix Virtual Apps 2206
 - Citrix Virtual Apps 2203 LTSR

2. Verified Operating Environments
 - Server OS : Windows Server 2019 (Virtual Apps 2112)
               Windows Server 2012 R2 (Virtual Apps 1912 LTSR)
               Windows Server 2022 (Virtual Apps 2203 LTSR)
 - Client OS : Windows 11, Windows 10, Windows 8.1


--------------------------------------------------------------------------------
Trademarks
--------------------------------------------------------------------------------

- The official name of Windows Server 2012 is Microsoft Windows Server 2012 Operating System.
- The official name of Windows Server 2019 is Microsoft Windows Server 2019 Operating System.
- The official name of Windows Server 2022 is Microsoft Windows Server 2022 Operating System.
- The official name of Windows 8 is Microsoft Windows 8 Operating System.
- The official name of Windows 10 is Microsoft Windows 10 Operating System.
- The official name of Windows 11 is Microsoft Windows 11 Operating System.
- Microsoft, Windows and the brand names and product names of other Microsoft products are trademarks of Microsoft Corporation in the US and other countries.
- Postscript, and the brand names and product names of other Adobe Systems products are trademarks of Adobe Systems Incorporated in the US and other countries.
- Novell, NetWare, and NDS are trademarks of Novell, Inc.
- Citrix, Citrix XenApp, Citrix Virtual Apps are trademarks of Citrix Systems, Inc. and/or one or more of its subsidiaries, and may be registered in the United States Patent and Trademark Office and in other countries.
- Bluetooth(R) is a registered trademark owned by Bluetooth SIG, lnc.
- Other company names and product names in this manual are the trademarks of their respective companies.
