# reportesInfrapy

Generar informes de capturas de chrome y ponerlos en un documento word y ppt

El flujo principal es:

* Cambiar a otra ventana mediante Alt + Tab, suponiendo que Chrome está abierto.
* Capturar una región fija de la pantalla.
* Guardar cada captura como tab_0.png, tab_1.png, etc.
* Cambiar a la siguiente pestaña con Ctrl + Tab.
* Comparar la captura actual con la primera captura para detectar si se volvió a la pestaña inicial.
* Detenerse cuando:
  Se regresa visualmente a la primera pestaña, o
  Se alcanzan 26 capturas.
* Insertar las imágenes en un documento Word.
* Insertar las mismas imágenes en una presentación PowerPoint.
* Guardar los archivos generados.
  Los archivos finales son:
      InformeMensualEstadoServidoresProduccion.docx
      InformeMensualEstadoServidoresProduccion.pptx

  Requisitos Python y las librerias para docx y pptx
