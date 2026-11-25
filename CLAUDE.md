# CLAUDE.md — cómo se trabaja en este repositorio

Instrucciones para Claude Code. Quien dirige no programa profesionalmente: tiene nociones básicas de desarrollo web.

## Registro

- Español rioplatense, directo, sin relleno. Toda respuesta empieza con la palabra «Bruno» (si falta, el contexto se degradó).
- Antes de tocar mucho código, explicar el plan en pocas líneas. Ante una duda razonable, declarar el supuesto y avanzar.
- Estados de lo que se afirma: `VERIFIED` · `PARTIALLY VERIFIED` · `ASSUMED` · `BLOCKED`. Nunca afirmar sin evidencia.

## Qué es

`pdf_master.pyw`: un solo archivo, Python + tkinter, para Windows. Convierte Word y PowerPoint a PDF con LibreOffice, une PDF y trae herramientas básicas. Detalle para usuarios en `README.md`; historia en `CHANGELOG.md`.

## Reglas

1. **Un solo archivo de doble clic.** No partirlo sin decisión explícita; no exigir consola al usuario final.
2. **Toda tarea declara antes cómo se demuestra que terminó**, y el reporte cierra con lo que no se pudo comprobar. Cambios chicos y verificables.
3. **Antes de commitear:** `python -m pytest -q` (en Linux, con `xvfb-run -a`). Si se toca la interfaz, regenerar y mirar las capturas: `python tests/captura.py docs/capturas`.
4. **No agregar dependencias sin avisar.** Hoy: LibreOffice, PyMuPDF (se instala solo) y tkinterdnd2 (opcional).
5. **Textos de la interfaz y errores en español rioplatense**, claros para alguien no técnico. Sin botones ni textos de más.
6. **Anticipar Windows:** rutas con espacios y «ñ», archivos descargados bloqueados, permisos, antivirus, WinError 10060, distintos intérpretes de Python.
7. **Estética de terminal de fósforo verde:** verde fuerte (`ACENTO`) y entre corchetes `[ ASÍ ]` = lo que se toca; verde suave (`FG`, `GRIS`) = información; verde pálido (`VERDE`) = lo importante o verificado; ámbar (`AMBAR`) = advertencia; rojo = error. Nada de azules ni violetas. Texto con contraste de 4,5:1 o más. Fuentes VT323 (solo tamaños grandes) y Share Tech Mono; ninguna tiene flechas, recuadros ni bloques: usar «»», palabras, o dibujarlos en un canvas.
8. **La sátira (clases `Guion`, `Arranque`, `Teatro`, línea `SIS>`) no puede mentir:** cantidades, rutas y tiempos que aparezcan son los reales; no se inventan indicadores de CPU, memoria ni seguridad; lo verdadero se distingue a la vista (verde pálido, «PROGRESO REAL»); los errores reales van en lenguaje llano y nunca se disfrazan. Inspiración en Fallout y Matrix sin copiar nombres, interfaces ni recursos (nada de RobCo, Vault-Tec, Pip-Boy, katakana). La institución es propia: «Dirección de Conversión Documental». Tres niveles (`MÁS… » Efectos visuales`): completos, reducidos (sin movimiento) y apagados (modo serio); se respeta «Mostrar animaciones» de Windows.
9. **Repositorio público.** Nada de datos personales (nombres de usuario de Windows, rutas reales, ubicación, correos) en código, capturas, commits ni issues.

## Bugs que no hay que reintroducir

- Unir con solo PDF y subcarpeta: el directorio de salida se crea siempre (test `test_fusion_crea_la_carpeta_de_salida`).
- Porcentaje de compresión con el signo invertido: `(d/a - 1)*100` (test `test_comprimir_informa_el_signo_correcto`).
- Raya larga del pie: fuente incrustada con Unicode, no Helvetica básica ni cp1252 (test `test_pie_con_raya_y_acentos`).
- La config sigue en `~/.convertidor_word_pdf.json` aunque el programa se llame PDF Master: cambiarla borra las preferencias de quien actualiza.
