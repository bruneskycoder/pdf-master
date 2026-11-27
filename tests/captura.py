"""Saca capturas de la interfaz con archivos de ejemplo (para el README y para revisar
cómo se ve en cada sistema). Uso: python tests/captura.py <carpeta-de-salida>
En Linux: xvfb-run -a python tests/captura.py docs/capturas"""
import importlib.machinery
import importlib.util
import os
import sys
import tempfile
from pathlib import Path

from PIL import ImageGrab

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = Path(sys.argv[1] if len(sys.argv) > 1 else "capturas")
SALIDA.mkdir(parents=True, exist_ok=True)

cargador = importlib.machinery.SourceFileLoader("pdf_master", str(RAIZ / "pdf_master.pyw"))
spec = importlib.util.spec_from_loader("pdf_master", cargador)
pm = importlib.util.module_from_spec(spec)
cargador.exec_module(pm)

tmp = Path(tempfile.mkdtemp())
pm.CONFIG = tmp / "config.json"
pm.App._preparar_dnd = lambda self: None
pm.App._estado_lo = lambda self: self.v_lo.set("LibreOffice: listo")
pm.modo_efectos = lambda cfg: "completos"
pm.buscar_soffice = lambda: "soffice"
demo = tmp / "Facultad" / "Física II"
demo.mkdir(parents=True)
for n in ("TP 1 - Introducción.docx", "TP 2 - Señales y sistemas.docx",
          "Clase 3 - Transformada.pptx", "TP 10 - Conclusión.docx", "Apunte de cátedra.pdf"):
    (demo / n).write_bytes(b"x" * 48000)


def foto(w, nombre):
    w.update()
    w.update_idletasks()
    x, y, a, h = w.winfo_rootx(), w.winfo_rooty(), w.winfo_width(), w.winfo_height()
    extra = {"xdisplay": os.environ["DISPLAY"]} if os.environ.get("DISPLAY") else {}
    ImageGrab.grab(bbox=(x, y, x + a, y + h), all_screens=sys.platform == "win32", **extra
                   ).save(SALIDA / f"{nombre}.png")
    print("captura:", SALIDA / f"{nombre}.png")


MD_EJEMPLO = """# Análisis Matemático I

Apuntes de la **clase 4**: límites y continuidad. *Borrador*, falta revisar los ejemplos.

## Qué vemos hoy

- Definición de límite
- Propiedades del límite
  - suma y producto
  - cociente (si el denominador no se anula)
- [x] Leer el capítulo 2
- [ ] Resolver la guía de ejercicios

> Un límite no dice qué vale la función en el punto, sino hacia dónde va.

### Valores de ejemplo

| Función | Punto | Límite |
|---|---|---:|
| f(x) = 2x + 1 | 3 | 7 |
| g(x) = x^2 | -2 | 4 |
| h(x) = 1 / x | 0 | no existe |

### En código

```python
def f(x):
    return 2 * x + 1

print(f(3))  # 7
```

---

Más material en la [página de la cátedra](https://ejemplo.org/catedra) y en `guia-4.md`.
"""


app = pm.App()
app.geometry(os.environ.get("GEO", "920x700+0+0"))


def arranque():
    foto(app, "arranque")
    app.after(900, pin)


def pin():
    foto(app, "pin")
    app.arranque.entrada.insert(0, "1111")
    app.arranque._validar()
    app.after(500, lambda: (foto(app, "pin_respuesta"), app.after(1700, vacia)))


def vacia():
    foto(app, "vacia")
    app.agregar_rutas([str(p) for p in demo.iterdir()])
    app.paginas[demo / "Apunte de cátedra.pdf"] = 12
    app.rangos[demo / "Apunte de cátedra.pdf"] = "1-4"
    app.v_salida_modo.set("subcarpeta")
    app.v_fusionar.set(True)
    app._alternar_fusion()
    app.v_nombre.set("Entrega final")
    app._llenar_lista()
    app.after(800, principal)


def principal():
    foto(app, "principal")
    app.guion.mazos["ATAQUES"] = [("SE DETECTÓ UN HACKER EN LA RED. PIDE QUE LE CONVIERTAN UN "
                                   "TRABAJO PRÁCTICO.", "SE LE ASIGNÓ TURNO PARA MARZO.")]
    app.ataque()
    app.after(420, lambda: (foto(app, "ataque"), app.after(2600, doctor)))


def doctor():
    """El Dr. Fósforo: su línea DR.FÓSFORO> y el globo (se fuerzan para la captura)."""
    d = app.doctor
    d.callar()
    d.saludado = True
    d.ultima_voz = d.ultimo_globo = -1e9
    app.decir("Eso no fue un virus: fui yo calibrando la antena con una cuchara.", doctor=True)
    d.globo(["La Dirección dice que soy", "«personal no autorizado».", "Yo digo que soy personal",
             "luminoso."], urgente=True)
    app.after(2300, lambda: (foto(app, "doctor"), d.callar(), app.after(300, hackear)))


def hackear():
    import types
    app.antes_de_convertir(sorted(demo.iterdir()), lambda: None)
    for i in range(330):
        app.hackeo._tecla(types.SimpleNamespace(keysym="a", char="a"))
    app.barra.config(maximum=5, value=2)
    app.v_status.set("OK  TP 2 - Señales y sistemas.docx")
    app.hackeo._tick_ahora()
    app.hackeo.pendiente = "KERNEL PANIC: el núcleo entró en pánico. Se lo calmó con un vaso de agua."
    for i in range(14):
        app.hackeo._tecla(types.SimpleNamespace(keysym="a", char="a"))
    app.hackeo._denegado()
    app.after(300, lambda: (foto(app, "hackeo"), app.hackeo.cerrar(), app.after(900, trabajar)))


def trabajar():
    app.trabajando = True
    app.btn_conv.config(state="disabled")
    app.v_fusionar.set(False)
    app._alternar_fusion()
    app.barra.config(maximum=5, value=3)
    app._alternar_log()
    for linea in ("--- Procesando ---", "OK  TP 1 - Introducción.docx", "OK  TP 2 - Señales y sistemas.docx",
                  "ERROR  TP 10 - Conclusión.docx: el archivo está abierto en otro programa"):
        app._log(linea)
    app.v_status.set("OK  TP 2 - Señales y sistemas.docx")
    app.teatro = pm.Teatro(app, app.medio)
    app.after(3200, trabajando)


def trabajando():
    foto(app, "trabajando")
    app.teatro.cerrar()
    app.trabajando = False
    app.abrir_herramientas()
    app._win_herr.geometry("+0+0")
    app.after(800, lambda: (foto(app._win_herr, "herramientas"), app.after(100, escanear)))


def escanear():
    """Ventana del escáner con dos fotos sintéticas difíciles (una con sombra fuerte arriba y
    otra girada), y después el diálogo de opciones del PDF."""
    sys.path.insert(0, str(RAIZ / "tests"))
    from sinteticas import foto_dificil
    fotos = []
    for i, caso in enumerate(("sombra", "rotada8")):
        f = tmp / f"Hoja {i + 1}.jpg"
        foto_dificil(caso)[0].save(f, quality=90)
        fotos.append(str(f))
    app.abrir_escaner()
    esc = app._win_esc
    esc.geometry("+0+0")
    esc.agregar_fotos(fotos)

    def listo():
        if len(esc.fotos) < 2:
            return app.after(200, listo)
        foto(esc, "escaner")
        esc.crear_pdf()
        esc.update()
        app.after(300, opciones)

    def opciones():
        foto(esc.dialogo, "escaner_opciones")
        esc.dialogo.destroy()
        esc.destroy()
        app.after(100, markdown)
    app.after(200, listo)


def markdown():
    """Lector de Markdown con apuntes de una materia ficticia."""
    f = tmp / "Apuntes de Análisis Matemático.md"
    f.write_text(MD_EJEMPLO, encoding="utf-8")
    app.abrir_markdown(str(f))
    w = app._win_md
    w.geometry("980x740+0+0")
    w.v_buscar.set("límite")
    w.buscar()
    w.caja.yview_moveto(0)
    app.after(600, lambda: (foto(w, "markdown"), app.destroy()))


app.after(700, arranque)
app.mainloop()
print("fuentes:", app.fam_titulo, "/", app.fam)
