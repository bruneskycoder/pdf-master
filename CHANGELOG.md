# Cambios

Formato: versión, qué cambia para quien lo usa y, aparte, lo técnico.

## 5.2 — 2026-11-25

**Para quien lo usa**

- **Siempre arranca con los efectos completos** (antes, si Windows tenía apagado «Mostrar animaciones», arrancaba en *Reducidos* y no se veía ni el arranque ni la pantalla de trabajo). Se cambia en `MÁS… » Efectos visuales`.
- **PIN de operador** al final del arranque. Abajo, en rojo, el sistema confiesa cuál es porque está hackeado desde 1994. Cualquier respuesta entra; no es una clave real y no se guarda.
- **Ataques falsos cada tanto:** destello rojo, encabezado corrompido y alarma en `SIS>`, que se resuelve sola a los dos segundos con un chiste. Nunca escriben en el registro.
- **Fallas rojas** que se recuperan en el acto durante la pantalla de trabajo, y un sector de arranque «dañado» en el arranque.
- **Capa de vidrio CRT en Windows** sobre toda la ventana: líneas de barrido, esquinas oscuras (curvatura insinuada), parpadeo y destellos. No recibe clics ni foco.
- Más comentarios en reposo, y más seguido (cada 18 a 45 segundos).

**Técnico**

- Clase nueva `PantallaCRT`: ventana en capas con color transparente, `-alpha` y `WS_EX_TRANSPARENT`; sigue a la principal.
- Corregido: un mensaje nuevo en `SIS>` no cortaba el tipeo del anterior, que seguía escribiendo encima (y en rojo, si el nuevo era una alarma).
- 67 pruebas (la de la capa corre solo en Windows).

## 5.1 — 2026-11-25

**Para quien lo usa**

- Terminal de **fósforo verde** sobre negro, con líneas de barrido, halo en los títulos y marco doble en el encabezado. Ámbar solo para advertencias, rojo para errores.
- Personalidad de burocracia informática: secuencia de arranque, línea `SIS>` que comenta cada acción con los números reales, comentarios ocasionales cuando el programa está quieto, microinterferencias raras y un mensaje oculto.
- Botones entre corchetes, que se «decodifican» al pasar el mouse. Barra de progreso de bloques.
- La pantalla de trabajo muestra expediente, subprocesos absurdos, el bloque **PROGRESO REAL** con su propia barra y, al final, el tiempo real de ejecución.
- Encabezado con datos reales: reloj, espacio libre en el disco de destino y estado de LibreOffice.
- Menú `MÁS… » Efectos visuales`: completos, reducidos o apagados (modo serio). Respeta «Mostrar animaciones» de Windows.
- Atajos: `Ctrl+O` agregar, `Ctrl+Enter` convertir, `Esc` ocultar efectos.

**Técnico**

- Clases nuevas: `Guion` (textos sin repetir hasta agotar), `BarraSegmentada` (misma interfaz que `ttk.Progressbar`), `Encabezado`, `Arranque`. `Teatro` rediseñado.
- 61 pruebas: suman modos de efectos, que el guion use los datos reales, modo serio sin sátira, arranque, barra, decodificación de botones y mensaje oculto.
- La preferencia vieja `"efectos": false` sigue valiendo (equivale a *Apagados*).

## 5.0 — 2026-11-24

**Para quien lo usa**

- Nuevo nombre: **PDF Master**. El archivo pasa a llamarse `pdf_master.pyw`; tus preferencias se conservan.
- Interfaz nueva, retro de terminal de fósforo: texto ámbar, lo que se puede tocar en cian, tipografías pixel incluidas en el mismo archivo.
- La pantalla se ordena en tres pasos numerados: elegir archivos, dónde se guardan, convertir.
- El menú «Más» pasa a ser el botón **MÁS…** arriba a la derecha.
- En pantallas bajas (notebooks) la lista se achica y el resto queda visible; al abrir el registro o las opciones de unir, la ventana crece sola si hay lugar.
- El registro marca en verde lo que salió bien y en rojo lo que falló; si algo falló, se abre solo al terminar.
- Efectos de carga satíricos estilo película de hackers (lluvia de caracteres y frases grandilocuentes). La línea «PROGRESO REAL» es lo único que informa; un clic los oculta y `"efectos": false` en las preferencias los apaga.
- **Corregido:** el pie de página imprimía la raya larga (—) como un punto (·). Ahora sale bien, con acentos y comillas tipográficas.

**Técnico**

- 52 pruebas automáticas (pytest): motor, motor PDF, conversión real con LibreOffice e interfaz. Corren en Windows y Linux con Python 3.12 y 3.14.
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
