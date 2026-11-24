# Cambios

Formato: versión, qué cambia para quien lo usa y, aparte, lo técnico.

## 5.0 — 2026-11-24

**Para quien lo usa**

- Nuevo nombre: **PDF Master**. El archivo pasa a llamarse `pdf_master.pyw`; tus preferencias se conservan.
- Interfaz nueva, retro de terminal de fósforo: texto ámbar, lo que se puede tocar en cian, tipografías pixel incluidas en el mismo archivo.
- La pantalla se ordena en tres pasos numerados: elegir archivos, dónde se guardan, convertir.
- El menú «Más» pasa a ser el botón **MÁS…** arriba a la derecha.
- En pantallas bajas (notebooks) la lista se achica y el resto queda visible; al abrir el registro o las opciones de unir, la ventana crece sola si hay lugar.
- El registro marca en verde lo que salió bien y en rojo lo que falló.
- **Corregido:** el pie de página imprimía la raya larga (—) como un punto (·). Ahora sale bien, con acentos y comillas tipográficas.

**Técnico**

- 47 pruebas automáticas (pytest): motor, motor PDF, conversión real con LibreOffice e interfaz. Corren en Windows y Linux con Python 3.12 y 3.14.
- El pie usa una fuente incrustada con Unicode y recortada a los caracteres usados (unos 15 KB por PDF unido).
- El registro de errores al abrir pasa a `pdf_master_error.log`.

## 4 y anteriores — sin fecha registrada

Resumen reconstruido de la conversación en la que se hizo:

- **v4:** pantalla única minimalista; elegir dónde se guardan los PDF con vista previa de la ruta; herramientas y carpetas automáticas en un menú «Más».
- Herramientas PDF (dividir, comprimir, eliminar páginas con miniaturas) y carpetas automáticas.
- Unir PDF con página en blanco tras documentos impares, total par, pie con `[PAGINA]`, `[TOTAL]` y `[ARCHIVO]`; enviar a PDFsam; abrir el resultado al terminar.
- PowerPoint además de Word.
- LibreOffice como único motor, con instalación automática (MSI oficial verificado con SHA-256; si falla, winget; si falla, la página de descarga).
- **Corregidos en el camino:** la unión fallaba con solo PDF y subcarpeta («no se pudo crear `PDF\Fusionado.pdf.part`»); el porcentaje de compresión salía con el signo invertido; la raya larga del pie salía como «?».
- Primera versión: un script que convertía `.docx` a PDF con `docx2pdf`.
