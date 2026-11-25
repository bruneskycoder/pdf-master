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


app = pm.App()
app.geometry(os.environ.get("GEO", "920x700+0+0"))


def arranque():
    foto(app, "arranque")
    app.after(1700, vacia)


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
    app.after(800, lambda: (foto(app._win_herr, "herramientas"), app.destroy()))


app.after(950, arranque)
app.mainloop()
print("fuentes:", app.fam_titulo, "/", app.fam)
