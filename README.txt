*Navegando sobre una página*----

Descargue el archivo web_scraping1.xlsm, valla al VBE presionando la combinación de teclas ALt + F11,
una vez ahí asegurese de tener activada la referencia 'Microsoft Internet Controls',luego abra el Módulo1
dentro se encontrará con la rutina web_scraping(), asigne un valor válido a la variable dni, abra la venta
inmediato con la combinación Ctrl + g, luego ejecute la rutina.Puede ver los resultados en la ventana inmediato.

Este método utiliza el enlace anticipado haciendo referencia a la librería Microsoft Internet Controls, 
puede cambiar este enfoque utilizando la función 'CreateObject("")' para crear una instancia del objeto Internet Explorer.

Dentro de Módulo2 encontrará el mismo proceso convertido en función 'function web_scraping(strDNI)', y otras 2 rutinas
tanto para consulta individual 'Sub test()' y consulta masiva 'Sub test2()' para esta última rutina debe apoyarse en la hoja de cálculo Hoja1 donde los datos se tomarán desde la celda "A2".


Visite el siguiente enlace para mas detalles:
https://youtu.be/tjmKgmneeME




*Realizamos una petición*--

El segundo archivo,web_scraping2.xlsm,igual que la anterior en este caso extraeremos los datos de una consulta de dni pero se hará usando otro método.Inspeccionando la página de la cual estraeremos los datos me tope con que esta web
utiliza una api y realiza las peticiones mediante el método GET, a diferencia del archivo anterior que usa el método POST para realizar la consulta, nos valiamos de otras librerias para navegar y recuperar esos datos.Al asunto dentro del árbol de archivos en el apartado de módulos se topará con JsonConverter y Testing,este último es donde esta nuestro código para realizar la consulta de dni, el módulo de JsonConverter no lo valla a modificar a menos que sepa que es lo que esta haciendo, este módulo es de un tercero y el repositorio también se encuentra alojado en GitHub, este nos servirá para leer la respuesta del servidor que nos entregará en formato JSON, en resumen deberá tener lo siguiente:

Activar las siguientes referencias en la pestaña de herramientas:
-Microsoft XML v6.0
-Microsoft Scripting Runtime

Deberá descargar el JsonConverter de GitHub.
-https://github.com/VBA-tools/VBA-JSON

Eso es todo,aquí el enlace al video tutorial.
-https://youtu.be/TyvQJNE-61I



*Utilizamos una API, para consultar*--

Accedemos a la carpeta 'consulta dni' dentro tenemos dos archivos;consulta api dni.xlsm y tokenJson.json.

En este caso realizaremos una consulta de dni a través de una api. Para ello el proveedor que utilizaremos nos brindará un token para poder hacer las consultas, este token lo almacenaremos dentro del archivo 'tokenJson.json'

Proveedor para el API:
-https://apiperu.dev/

Activar las siguientes referencias en la pestaña de herramientas:
-Microsoft XML v6.0
-Microsoft Scripting Runtime

Deberá descargar el JsonConverter de GitHub.
-https://github.com/VBA-tools/VBA-JSON

Enlace al video tutorial.
https://youtu.be/VoxvbqspG9U


