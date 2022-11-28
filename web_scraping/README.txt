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

