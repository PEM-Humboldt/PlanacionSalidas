# PlanacionSalidas
código para la sistematización de solicitudes  salidas de campo anticipos y transporte.
La rutina aquí consignada genera un vinculo entre R y la tabla de planeción de salidas de campo depositada en google sheet, lo que permite la planeaión y actualización e todo el equipo Humboldt.

Una vez se tienen confirmadas y verificadas las condiciones y neesidades para la salida de campo, el código genera la informaión en el formato requerido por los apoyos logísticos para la solicitud en intranet. 

Esta rutina facilita y optimiza el seguimiento planeación y solicitud de las salidas de campo. 

####Auntenfificación de cuentas vinculo R - google

#######Ingresar a:  https://console.developers.google.com/project. desde la cuenta gmail a vincular.
 Buscar el proyecto "PalneacionSalidas" y con calidad de editor cree su propio ID de clientes OAuth 2.0

## Al crear la cuenta de OAuth 2.0 tenga en cuneta asignar la dirección URI de la aplicación.
descargue y asignele una carpeta al archivo "client secret" y copie la clave de la API, remplace y corra la siguiente linea (la dirección URI aparece en el final de la notificación del error de autentificación)


