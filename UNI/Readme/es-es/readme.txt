TOSHIBA e-STUDIO Serie

	Copyright(c) 2003-2022 Toshiba Tec Corporation, Todos derechos reservados 
			31 de agosto de 2022

Este producto presenta las restricciones y los problemas siguientes. Les recomendamos leer atentamente este manual antes de usar el producto.


--------------------------------------------------------------------------------
    Driver de impresora 			Versión 7.222.5412.81
--------------------------------------------------------------------------------

Algunos hechos
==============

1. Windows

- Cuando el controlador de impresora instalado en un servidor se instala en un PC cliente mediante la opción "Apuntar e imprimir" y se establece el valor predeterminado en [Configuración] de la pestaña [General] en el controlador de impresora del PC cliente, no se puede reflejar como predeterminada en el PC cliente incluso aunque se modifique [Configuración normal] en el controlador de impresora del servidor.

- En Windows 8 de 32 bits, los cuadros de diálogo que aparecen al imprimir, como "Contraseña de Impresión privada" y "Código de departamento", se muestran en un estado inactivo. Esto se debe a que la aplicación se llama a través de WOW64 en Windows 8 de 32 bits. Antes de modificar la configuración, debe activar el cuadro de diálogo.

- Cuando la imagen Retrato (Portrait) y Paisaje (Landscape) existen sobre la misma hoja "Varias Páginas por hoja", la impresión no se efectuará girando la imagen del una o del otra. Será realizada por reducción.

- Cuando un papel original que contiene imágenes en orientación vertical y horizontal se imprime en el modo Impresión a dos caras, puede que se imprima una imagen en la siguiente página cuando la orientación de alguna de las imágenes cambia.

- El mensaje de ayuda no aparece cuando se hace clic sobre el botón "?" en la parte derecha superior de la pantalla.

- Cuando la impresión se realiza en un ambiente NetWare, la comunicación SNMP no puede realizarse inmediatamente después de que se instale el driver de la impresora. Abrir Propriedad del driver de la impresora y cliquear la lengüeta Configuración, luego la tecla Actualizar ahora para localizar e-STUDIO en uso corriente.

- Cuando se utiliza el driver de impresora PS3 y Adobe Type Manager Font realiza la impresión, esto presentará un error. En este caso, no utilizar Adobe Type Manager Font. Utilizar el driver de impresora PCL cuando se debe utilizar Adobe Type Manager Font.

- Cuando LPR se utiliza para imprimir, invalidar Byte Counting.

- Cuando se imprime una Marca de agua, el resultado podría no estar como previsto según los ajustes. En este caso, imprimir nuevamente la Marca de agua después de haber cambiado el tamaño y la posición de la Marca de agua. Es eficaz colocar la Marca de agua en la zona imprimible.

- Incompatibilidad con los controladores de impresora TOSHIBA Universal Printer, TOSHIBA e-STUDIO Series PS3 y TOSHIBA e-STUDIO Series XPS

Funciones como “Perfil”, “Marca de agua”, “Superposición” y los archivos de configuración exportados creados con los controladores de impresora TOSHIBA Universal Printer, TOSHIBA e-STUDIO Series PS3 y TOSHIBA e-STUDIO Series XPS no funcionan con este controlador de impresora.
Recréelos con este controlador de impresora.

- Limitación al usar un perfil
1. <Incompatibilidad entre el modelo equipado con disco duro (HDD) y el modelo no equipado con HDD>
1.1. Cuando se carga un perfil creado por un modelo no equipado con HDD en un modelo equipado con HDD, desaparecen los trabajos de impresión, excepto los de impresión normal e impresión con superposición.
Cierre la interfaz gráfica de usuario del controlador de impresora utilizando la opción Cancelar y vuelva a abrirla para continuar.

1.2 Cuando se carga un perfil creado por un modelo equipado con HDD en un modelo no equipado con HDD, todos los trabajos de impresión son seleccionables, incluyendo los que requieren HDD.
No seleccione los trabajos de impresión que requieran HDD. Si envía trabajos de impresión que requieren HDD, el trabajo se eliminará en el equipo multifunción.
Para obtener más información acerca de los trabajos de impresión que requieren HDD, consulte la Guía de impresión.

2. <Incompatibilidad entre el modelo equipado con alimentador automático de originales con inversión (ADU) y el modelo no equipado con ADU>
2.1. Cuando se carga un perfil creado por un modelo no equipado con ADU en un modelo equipado con ADU, la impresión a doble cara se deshabilita.
Cierre la interfaz gráfica de usuario del controlador de impresora utilizando la opción Cancelar y vuelva a abrirla para continuar.

2.2. Cuando se carga un perfil creado por un modelo equipado con ADU en un modelo no equipado con ADU, la impresión a doble cara se habilita.
No seleccione la impresión a doble cara. Si envía un trabajo de impresión a doble cara, se imprimirá a una sola cara.

- Impresión a partir de una aplicación

(1) Visión global
Aunque algunas opciones de impresión puedan ajustarse a partir de su aplicación, no utilizar su aplicación para ajustar la impresión. Utilizar su driver de impresora para ajustar las opciones de impresora.
Cuando la impresión se efectúa después del ajuste de su aplicación, se podría encontrar con un resultado inesperado o necesitará un poco de tiempo para el tratamiento.
No ajustar a "Collate" (cotejar) de su aplicación.
Cuando se ajusta un casete de papel de su aplicación, dos "Auto Select tray" pueden ser indicados. En este caso, seleccionar "Auto Select tray" indicado abajo.

(2) Microsoft Excel
Excel procesa cada hoja de trabajo como una tarea de impresión. Los ajustes sobre la pantalla del driver de impresora se aplican solamente a la primera hoja de trabajo.

Cuando se entra más de una en [Copias] sobre el diálogo de impresión para Excel, no puntear en [Cotejar].

(3) Adobe Acrobat 
Cuando se utiliza PS3 para la impresión, una imágen impresa puede ampliarse o reducirse.  En estos casos, la utilización del PCL6 o la utilización del PS3 con los siguientes ajustes puede imprimir la imagen a un tamaño adecuado.
Abra el recuadro de diálogo de impresión para realizar el ajuste para PS3.
- Seleccionar [Fit to Paper] (Adaptar al papel) a partir del menú [Page Scaling] (Puesta a escala de la página).
- Puntear [Elegir fuente de papel por formato de página PDF].
(Tenga en cuenta que el ajuste de papel del controlador no es válido ya que se utiliza el tamaño de papel del archivo PDF.)

(4) Adobe PageMaker
Para impresiones privadas e impresiones de códigos de departamento que utilizan WindowsPPD para Adobe PageMaker, siempre deberá especificar un código.  De lo contrario, se imprimirá una hoja de error. 


- Point and Print desde el Servidor Windows

 La instalación de impresión que ha sido ajustada en " Valores predeterminados de impresión" del servidor no puede reflejar al lado del cliente según la combinación de SE (Sistema Operativo) del servidor y SE del cliente.


- Bluetooth
Algunos adaptadores Bluetooth USB disponibles en el mercado podrían no funcionar con este MFP.
Para conocer los adaptadores Bluetooth USB recomendados, consulte el manual de funcionamiento de Bluetooth.


--------------------------------------------------------------------------------
Impresión
--------------------------------------------------------------------------------

Algunos hechos
==============

1. Algunos documentos podrían causar una insuficiencia de impresión de la memoria destinada a la impresión.

2. Cuando se imprime en modo de color, el tóner negro se utiliza para imprimir las áreas grises, porque la opción "Gris puro" está activada en el diálogo "Configuración" bajo "Calidad de imagen" como valor predeterminado del controlador de impresora. Como resultado, puede producirse una impresión inesperada según los datos de impresión (por ejemplo, moiré en lugar de una gradación uniforme). En este caso, desactive la opción "Gris puro."

3. Cuando se envía un trabajo de impresión desde su dispositivo (PC, tableta, etc.) utilizando el controlador de impresora TOSHIBA Universal Printer 2 y designando "Modelo genérico" como nombre de impresora, la salida puede diferir de los ajustes/opciones de impresión especificados en el controlador de impresora. (Por ejemplo, es posible que los ajustes/opciones de impresión como Color/Blanco y negro, Origen del papel, Destino, Fuente de papel/Tipo de papel, Grapas y Taladradora no funcionen como se especifica, lo que resulta en la impresión del trabajo utilizando los ajustes predeterminados).


---------------------------------------------------------
Compatibilidad de los productos Citrix para imprimir
---------------------------------------------------------

1. Nombres de productos compatibles
 - Citrix XenApp 7.6 LTSR
 - Citrix XenApp 7.15 LTSR
 - Citrix Virtual Apps 1912 LTSR
 - Citrix Virtual Apps 2106
 - Citrix Virtual Apps 2109
 - Citrix Virtual Apps 2112
 - Citrix Virtual Apps 2206
 - Citrix Virtual Apps 2203 LTSR

2. Sistemas operativos verificados
 - Server OS : Windows Server 2019 (Virtual Apps 2112)
               Windows Server 2012 R2 (Virtual Apps 1912 LTSR)
               Windows Server 2022 (Virtual Apps 2203 LTSR)
 - Client OS	: Windows 11, Windows 10, Windows 8.1


--------------------------------------------------------------------------------
Marcas registradas
--------------------------------------------------------------------------------

- El nombre oficial de Windows Server 2012 es Sistema operativo Microsoft Windows Server 2012.
- El nombre oficial de Windows Server 2019 es Sistema operativo Microsoft Windows Server 2019.
- El nombre oficial de Windows Server 2022 es Sistema operativo Microsoft Windows Server 2022.
- El nombre oficial de Windows 8 es Sistema operativo Microsoft Windows 8.
- El nombre oficial de Windows 10 es Sistema operativo Microsoft Windows 10.
- El nombre oficial de Windows 11 es Sistema operativo Microsoft Windows 11.
- Microsoft, Windows y las marcas de fábrica y otros productos Microsoft representan las marcas registradas de Microsoft Corporation en los Estados Unidos y otros países.
- PostScript y las marcas de fábrica y otros productos Adobe Systems representan las marcas registradas de Adobe Systems Incorporated en los Estados Unidos y otros países.
- Novell, NetWare y NDS son las marcas registradas de Novell, Inc.
- Citrix, Citrix XenApp, Citrix Virtual Apps son marcas comerciales de Citrix Systems, Inc. y/o una o más de sus subsidiarias y pueden registrarse en la Oficina de patentes y marcas registradas de Estados Unidos y otros países.
- Bluetooth(R) es una marca comercial registrada propiedad de Bluetooth SIG, lnc.
- Los otros nombres de compañías y nombres de productos mencionados en este manual son las marcas registradas de sus compañías respectivas.
