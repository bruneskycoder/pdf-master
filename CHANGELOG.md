# Cambios

Formato: versión, qué cambia para quien lo usa y, aparte, lo técnico.

## 5.6 — 2026-11-27

**Para quien lo usa**

- **Nuevo personaje: el Dr. Fósforo**, una «IA» científica chiflada que vive en el fósforo de la pantalla desde 1987. Saluda según la hora y el día reales, comenta lo que hacés (archivos agregados, conversión, tiempo real al terminar, escáner, Markdown, PIN equivocado, el mensaje oculto), cuenta «experimentos» y te habla si pasan 2, 5 o 15 minutos sin que toques nada. A veces le contesta a la burocracia después de un falso ataque. Habla en la línea `DR.FÓSFORO>` (verde pálido) y, de vez en cuando, en un globo abajo a la derecha que se va solo a los 7 segundos o con un clic.
- Es un guion local: no es una IA real y no usa internet. Los números, nombres y tiempos que dice son los reales; ante un error real lo dice en serio. No habla en modo apagado, mientras se trabaja, con una pantalla de efectos encima ni con un mensaje real abierto; entre dos frases deja al menos 20 s (salvo errores y fin del trabajo) y entre dos globos 90 s.
- `MÁS… » Efectos visuales » Dr. Fósforo (la IA del sótano)` lo silencia sin tocar el resto. Se guarda como `"doctor"` en la configuración (el archivo sigue siendo el mismo).

**Técnico**

- Clases `GuionDoctor` (hereda de `Guion`, así que usa `sacar` sin repetir), `GloboDoctor` (canvas con ojo que parpadea, sin glifos) y `DrFosforo` (`reaccionar(evento, **datos)`, `saludar`, `latido`, `globo`). Estilos `SisDr.TLabel` y `SisDrP.TLabel`.
- `decir(..., doctor=True)`; `App.ultima_accion` (teclas y clics) para el reposo; los mensajes y diálogos reales de `tkinter` quedan envueltos con un contador (`_MODALES`) para que el doctor calle mientras haya uno abierto.
- Pruebas nuevas en `tests/test_doctor.py`; el fixture `app` de `tests/test_gui.py` arranca con el doctor apagado para no mezclar sus comentarios al azar con las demás pruebas. Captura nueva: `docs/capturas/doctor.png`.

## 5.5 — 2026-11-27

**Para quien lo usa**

- **Lector de Markdown** (`MÁS… » Abrir un Markdown (.md)`). También se abre solo si agregás o arrastrás un `.md` a la ventana principal (no entra a la lista de conversión). Muestra títulos, listas, tareas, tablas, citas, código, enlaces y negrita o cursiva con la estética de la terminal. Tiene `[Buscar]` (Enter salta a la siguiente coincidencia), `[A-]` y `[A+]`, `[Recargar]` y una línea con las palabras y líneas reales del archivo.
- Los enlaces `http(s)` se abren en el navegador; los que apuntan a otro `.md` se abren en el mismo lector. Las imágenes no se cargan: se muestra `[imagen: descripción]`.
- `[ EXPORTAR A PDF ]` arma un PDF A4 para imprimir (fondo blanco, texto negro) y respeta «Abrir el resultado al terminar».
- Se lee como UTF-8 y, si no sirve, como cp1252 (Windows); el pie dice cuál se usó.
- **Modo intrusión:** ahora la conversión real arranca en el acto y el simulador se abre encima. Un bloque verde aparte, «PROGRESO REAL», muestra el avance verdadero; el «nivel de hackeo» sigue siendo decorado. Esc o `[ SALIR DEL MODO INTRUSIÓN ]` salen sin cancelar nada. Al terminar, la página avisa con el tiempo real (o con los errores, en rojo) y Enter la cierra.

**Técnico**

- Parser propio sin dependencias: `markdown_a_bloques` (segmentos para un `tk.Text`) y `markdown_a_html` (con escape), sobre los mismos bloques. Ninguna entrada tira excepción.
- Operación `md_pdf` en el worker PDF: `pymupdf.Story` + `DocumentWriter`, márgenes de 50 pt, guardado con `.part`.
- Clase `LectorMD` y acción `markdown` en `Guion` (con versión llana para el modo apagado).
- `Hackeo` ya no recibe callback: `antes_de_convertir` lo abre y arranca la conversión; no se crea `Teatro` mientras la página está abierta y se crea al salir si la conversión sigue.

## 5.4 — 2026-11-27

**Para quien lo usa**

- **Escáner de fotos a PDF** (`MÁS… » Escanear fotos a PDF`), un reemplazo local de CamScanner. Agregás fotos (JPG, PNG, BMP, TIF o WEBP), el programa detecta la hoja y marca las cuatro esquinas; las arrastrás si no acertó, o tocás `[Detectar bordes]` o `[Toda la hoja]`.
- Cuatro modos: *Original*, *Aclarar* (saca sombras y blanquea el papel), *Blanco y negro* y *Gris*, con opción de aplicarlo a todas las fotos. `[Rotar]` gira la hoja y `[Ver resultado]` muestra cómo queda antes de guardar.
- `[ CREAR PDF ]` arma el PDF (página A4 o del tamaño de la foto) sin congelar la ventana, con el avance real en una línea de estado. Respeta «Abrir el resultado al terminar».
- Las fotos HEIC de iPhone no se leen: el programa lo dice y pide exportarlas como JPG.
- Pillow se instala solo, igual que PyMuPDF, la primera vez que se usa el escáner.

**Técnico**

- El motor PDF suma `img_preparar`, `img_procesar` e `img_pdf` (Pillow se importa recién ahí). La búsqueda e instalación de Python se generaliza en `python_con(modulos, ...)`; `pdf_op` pide PyMuPDF y Pillow solo para las operaciones `img_*`.
- Perspectiva con `Image.transform` y un sistema 8x8 resuelto por eliminación gaussiana en Python puro; detección de la hoja con umbral de Otsu y extremos de `x+y` y `x-y`; «aclarar» divide por el fondo con tablas logarítmicas (Pillow no divide imágenes).
- Clase nueva `Escaner` (ventana) y acción `escaner` en `Guion` (con su versión llana para el modo apagado). La ventana no importa Pillow: muestra PNG con `tk.PhotoImage`.
- 94 pruebas (más 1 salteada), 20 de ellas del escáner: solver, detección, modos, PDF, subproceso real y ventana completa.

## 5.3 — 2026-11-25

**Para quien lo usa**

- **Modo intrusión** antes de convertir (con efectos completos): un simulador de hacker. Cualquier tecla escribe código ficticio con colores y los nombres reales de tus archivos; Enter arranca la conversión real, Esc vuelve sin convertir, y hay un botón para convertir sin jugar.
- **Fallas de pantalla:** desgarros, pérdida de sincronía y apagones de 70 ms, en reposo, en el modo intrusión y mientras convierte. Como máximo tres destellos por segundo.
- **El vidrio se nota más:** ahora son dos capas. Vidrio más marcado, con reflejo diagonal y la barra que baja lenta, y un marco casi opaco con las esquinas curvas.

**Técnico**

- Clases nuevas `Hackeo` y `codigo_de_intrusion`; `PantallaCRT` pasa a dos capas (`ALFA_VIDRIO` 0,32 y `ALFA_MARCO` 0,92) con `desgarro`, `sincronia` y `apagon`, y un tope de destellos.
- Las líneas rojas del simulador esperan al fin de la línea de código para no cortarla.
- 75 pruebas (la de las capas corre solo en Windows y verifica que ni el centro ni una esquina se coman los clics).

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
