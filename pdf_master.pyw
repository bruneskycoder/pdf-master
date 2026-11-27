# -*- coding: utf-8 -*-
"""
PDF Master: convertir Word y PowerPoint a PDF, unir PDF y herramientas básicas.
UN SOLO ARCHIVO, solo necesita Python. https://github.com/bruneskycoder/pdf-master

Ventanas:
  1. Principal            : Word/PowerPoint -> PDF (LibreOffice), lista con páginas, tamaño e
                            intervalo; unión con página en blanco si es impar y pie de página.
  2. Herramientas PDF     : dividir, comprimir y eliminar páginas con vista de miniaturas.
  3. Carpetas automáticas : carpetas vigiladas que dividen, unen o comprimen solas lo que llegue.
  4. Escáner de fotos     : fotos de hojas » recorte, perspectiva y limpieza » PDF (Pillow).

Motor de conversión: LibreOffice (si falta, el programa lo descarga y lo instala).
Motor PDF: PyMuPDF; escáner de fotos: Pillow (si faltan, el programa los instala con pip, una sola vez).
Fuentes incluidas: VT323 y Share Tech Mono (SIL Open Font License 1.1).

Licencia MIT. © 2026 Bruno Gutiérrez (@bruneskycoder).
"""
import base64
import hashlib
import json
import os
import platform
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import traceback
import urllib.error
import urllib.request
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

# Cuántos mensajes o diálogos reales hay abiertos ahora: el Dr. Fósforo calla mientras haya uno.
_MODALES = [0]


def _vigilar_modales(modulo, nombres):
    for nombre in nombres:
        original = getattr(modulo, nombre, None)
        if not callable(original):
            continue

        def envoltura(*a, _f=original, **kw):
            _MODALES[0] += 1
            try:
                return _f(*a, **kw)
            finally:
                _MODALES[0] = max(0, _MODALES[0] - 1)
        setattr(modulo, nombre, envoltura)


_vigilar_modales(messagebox, ("showinfo", "showwarning", "showerror", "askyesno", "askokcancel",
                              "askretrycancel", "askquestion", "askyesnocancel"))
_vigilar_modales(filedialog, ("askopenfilename", "askopenfilenames", "asksaveasfilename",
                              "askdirectory"))
_vigilar_modales(simpledialog, ("askstring", "askinteger", "askfloat"))

try:  # arrastrar y soltar (opcional)
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _BaseTk = TkinterDnD.Tk
    HAY_DND = True
except Exception:
    DND_FILES = None
    _BaseTk = tk.Tk
    HAY_DND = False

VERSION = "5.6"
CONVERTIBLES = {".docx", ".doc", ".pptx", ".ppt"}
EXT = CONVERTIBLES | {".pdf"}
CONFIG = Path.home() / ".convertidor_word_pdf.json"
ES_WINDOWS = platform.system() == "Windows"
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
URL_BASE = "https://download.documentfoundation.org/libreoffice/stable/"
URL_PAGINA = "https://www.libreoffice.org/download/download-libreoffice/"
URL_PDFSAM = "https://pdfsam.org/pdfsam-basic/"
LOTE = 10  # archivos por llamada a LibreOffice


# ------------------------------------------------------------------ utilidades
def cargar_config():
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def guardar_config(datos):
    try:
        CONFIG.write_text(json.dumps(datos), encoding="utf-8")
    except Exception:
        pass


def buscar_soffice():
    for nombre in ("soffice", "libreoffice"):
        ruta = shutil.which(nombre)
        if ruta:
            return ruta
    fijos = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/opt/libreoffice/program/soffice",
    ]
    for var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432", "LOCALAPPDATA"):
        base = os.environ.get(var)
        if base:
            for carpeta in sorted(Path(base).glob("LibreOffice*"), reverse=True):
                fijos.append(str(carpeta / "program" / "soffice.exe"))
    for c in fijos:
        if Path(c).exists():
            return c
    return None


def _clave_natural(p: Path):
    """Orden 'humano': tp 2 antes que tp 10."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", str(p))]


def filtrar_archivos(rutas):
    """Deja Word/PowerPoint/PDF reales (sin temporales ~$). Un PDF con el mismo
    nombre que un Word/PowerPoint de la misma carpeta se descarta (es su resultado)."""
    cand = [Path(r) for r in rutas
            if Path(r).suffix.lower() in EXT and not Path(r).name.startswith("~$")]
    origenes = {(p.parent, p.stem.lower()) for p in cand if p.suffix.lower() in CONVERTIBLES}
    cand = [p for p in cand
            if not (p.suffix.lower() == ".pdf" and (p.parent, p.stem.lower()) in origenes)]
    return sorted(cand, key=_clave_natural)


def listar(carpeta: Path, recursivo: bool):
    patron = "**/*" if recursivo else "*"
    return filtrar_archivos(p for p in carpeta.glob(patron) if p.is_file())


def destino_pdf(origen: Path, base, salida) -> Path:
    """Dónde queda el PDF de 'origen'.
    salida: False = junto al original · True = subcarpeta 'PDF' de 'base' · Path = esa carpeta."""
    if origen.suffix.lower() == ".pdf":
        return origen  # ya es PDF: se usa tal cual
    if isinstance(salida, Path):
        try:
            rel = origen.relative_to(base)
        except (ValueError, TypeError):
            rel = Path(origen.name)
        return (salida / rel).with_suffix(".pdf")
    if salida:
        try:
            return (base / "PDF" / origen.relative_to(base)).with_suffix(".pdf")
        except ValueError:
            return (base / "PDF" / origen.name).with_suffix(".pdf")
    return origen.with_suffix(".pdf")


def carpeta_destino(base: Path, salida) -> Path:
    """Carpeta donde se guardan los PDF fusionados (y la que se abre al terminar)."""
    if isinstance(salida, Path):
        return salida
    return base / "PDF" if salida else base


# ------------------------------------------------------- instalar LibreOffice
def _arquitectura():
    m = platform.machine().lower()
    if m in ("arm64", "aarch64"):
        return "aarch64", "Win_aarch64"
    if m in ("amd64", "x86_64", "x64"):
        return "x86_64", "Win_x64"
    return None, None


def _pedir(url, timeout=60):
    req = urllib.request.Request(
        url, headers={"User-Agent": "Mozilla/5.0 convertidor-word-pdf"})
    return urllib.request.urlopen(req, timeout=timeout)


def _versiones_disponibles():
    html = _pedir(URL_BASE).read().decode("utf-8", "replace")
    vs = set(re.findall(r'href="(\d+\.\d+\.\d+(?:\.\d+)?)/"', html))
    return sorted(vs, key=lambda v: [int(x) for x in v.split(".")], reverse=True)


def _descargar(url, destino: Path, emit):
    h = hashlib.sha256()
    with _pedir(url) as r, open(destino, "wb") as f:
        total = int(r.headers.get("Content-Length") or 0)
        leido, ultimo = 0, -1
        if total:
            emit("max", total)
        while True:
            bloque = r.read(1 << 20)
            if not bloque:
                break
            f.write(bloque)
            h.update(bloque)
            leido += len(bloque)
            emit("prog", leido)
            if total:
                pct = int(leido * 100 / total) // 10 * 10
                if pct != ultimo:
                    ultimo = pct
                    emit("log", f"  descargando… {pct}%  ({leido >> 20} de {total >> 20} MB)")
    if total and leido != total:
        raise IOError("la descarga quedó incompleta")
    return h.hexdigest(), leido


def _instalar_msi(msi: Path, emit):
    """Devuelve 'ok', 'cancelado' o 'error'."""
    emit("log", "Instalando LibreOffice… Windows te va a pedir permiso: elegí Sí.")
    r = subprocess.run(["msiexec", "/i", str(msi), "/passive", "/norestart"],
                       timeout=3600, creationflags=NO_WINDOW)
    if r.returncode in (0, 3010):
        return "ok"
    if r.returncode in (1602, 1223):
        return "cancelado"
    emit("log", f"El instalador terminó con código {r.returncode}.")
    return "error"


def _por_msi(emit):
    arq, etiqueta = _arquitectura()
    if not arq:
        raise RuntimeError("arquitectura de Windows no soportada para la descarga automática")
    emit("log", "Buscando la versión actual de LibreOffice…")
    versiones = _versiones_disponibles()[:2]
    if not versiones:
        raise RuntimeError("no pude leer la lista de versiones")
    tmp = Path(tempfile.mkdtemp(prefix="lo_instalador_"))
    try:
        ultimo_error = None
        for v in versiones:
            nombre = f"LibreOffice_{v}_{etiqueta}.msi"
            url = f"{URL_BASE}{v}/win/{arq}/{nombre}"
            for intento in (1, 2):
                try:
                    emit("log", f"Descargando LibreOffice {v} (unos 350 MB)…")
                    msi = tmp / nombre
                    sha, tam = _descargar(url, msi, emit)
                    if tam < 100 * 2**20:
                        raise IOError("el archivo descargado es demasiado chico")
                    try:
                        txt = _pedir(url + ".sha256").read().decode("utf-8", "replace")
                        esperado = re.search(r"\b[0-9a-fA-F]{64}\b", txt)
                        if esperado and esperado.group(0).lower() != sha.lower():
                            raise IOError("la verificación (SHA-256) no coincide")
                        emit("log", "Descarga verificada.")
                    except urllib.error.URLError:
                        emit("log", "(No pude bajar el archivo de verificación; sigo igual.)")
                    resultado = _instalar_msi(msi, emit)
                    if resultado == "cancelado":
                        return "cancelado"
                    if resultado == "ok":
                        return "ok"
                    raise RuntimeError("falló la instalación")
                except urllib.error.HTTPError as e:
                    ultimo_error = e
                    if e.code == 404:
                        break  # esa versión no tiene el instalador; probar la anterior
                except Exception as e:
                    ultimo_error = e
                    emit("log", f"  problema: {e}")
        raise RuntimeError(f"{ultimo_error}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _por_winget(emit):
    w = shutil.which("winget")
    if not w:
        return False
    emit("log", "Probando instalar con winget (el instalador de Windows)…")
    try:
        subprocess.run(
            [w, "install", "--id", "TheDocumentFoundation.LibreOffice", "-e", "--silent",
             "--accept-package-agreements", "--accept-source-agreements"],
            capture_output=True, timeout=3600, creationflags=NO_WINDOW)
    except Exception as e:
        emit("log", f"winget falló: {e}")
    return bool(buscar_soffice())


def instalar_libreoffice(emit):
    """Devuelve True si LibreOffice quedó instalado."""
    if not ES_WINDOWS:
        emit("log", "Instalá LibreOffice con el gestor de paquetes de tu sistema "
                    "(por ejemplo: sudo apt install libreoffice-writer).")
        return False
    try:
        r = _por_msi(emit)
        if r == "cancelado":
            emit("log", "Cancelaste la instalación.")
            return False
        if buscar_soffice():
            emit("log", "LibreOffice instalado.")
            return True
    except Exception as e:
        emit("log", f"No pude descargar/instalar directo: {e}")
    if _por_winget(emit):
        emit("log", "LibreOffice instalado.")
        return True
    emit("log", "No se pudo instalar automáticamente. Abro la página oficial: descargá "
                "LibreOffice, instalalo y volvé a apretar Convertir.")
    emit("abrir_web", URL_PAGINA)
    return False


# ------------------------------------------------------------------ conversión
def _lanzar_lo(soffice, uri_perfil, carpeta_salida: Path, origenes):
    cmd = [soffice, f"-env:UserInstallation={uri_perfil}", "--headless", "--norestore",
           "--nolockcheck", "--convert-to", "pdf", "--outdir", str(carpeta_salida),
           *[str(o) for o in origenes]]
    try:
        subprocess.run(cmd, capture_output=True, timeout=600, creationflags=NO_WINDOW)
    except subprocess.TimeoutExpired:
        pass


def _generado(destino: Path, desde: float):
    try:
        return destino.exists() and destino.stat().st_size > 0 and destino.stat().st_mtime >= desde
    except OSError:
        return False


def convertir_lote(archivos, base: Path, en_subcarpeta, sobrescribir, soffice, emit):
    """Convierte lo que haga falta. Devuelve la lista de PDF resultantes, en el
    mismo orden que 'archivos' (los que fallaron no aparecen)."""
    ok = errores = salteados = 0
    emit("max", max(len(archivos), 1))
    hechos = 0
    salidas = {}
    pendientes = []
    for o in archivos:
        d = destino_pdf(o, base, en_subcarpeta)
        if o.suffix.lower() == ".pdf":
            emit("log", f"PDF  {o.name} (ya es PDF, se usa tal cual)")
            salidas[o] = o
            hechos += 1
            emit("prog", hechos)
        elif d.exists() and not sobrescribir:
            emit("log", f"SALTEADO  {o.name} (ya existe el PDF)")
            salidas[o] = d
            salteados += 1
            hechos += 1
            emit("prog", hechos)
        else:
            pendientes.append((o, d))

    if pendientes:
        perfil = tempfile.mkdtemp(prefix="lo_perfil_")
        uri = Path(perfil).resolve().as_uri()
        fallidos = []
        try:
            grupos = {}
            for o, d in pendientes:
                grupos.setdefault(d.parent, []).append((o, d))
            for carpeta, items in grupos.items():
                try:
                    carpeta.mkdir(parents=True, exist_ok=True)
                except OSError as e:
                    for o, _ in items:
                        emit("log", f"ERROR  {o.name} -> no pude crear la carpeta de salida: {e}")
                        errores += 1
                        hechos += 1
                        emit("prog", hechos)
                    continue
                for i in range(0, len(items), LOTE):
                    lote = items[i:i + LOTE]
                    t0 = time.time() - 2
                    _lanzar_lo(soffice, uri, carpeta, [o for o, _ in lote])
                    for o, d in lote:
                        if _generado(d, t0):
                            ok += 1
                            salidas[o] = d
                            emit("log", f"OK  {o.name}")
                        else:
                            fallidos.append((o, d))
                        hechos += 1
                        emit("prog", hechos)

            # Reintento individual con copia temporal de nombre simple
            for o, d in fallidos:
                emit("log", f"Reintentando  {o.name} …")
                tmp = Path(tempfile.mkdtemp(prefix="w2p_"))
                try:
                    src = tmp / ("origen" + o.suffix.lower())
                    shutil.copyfile(o, src)
                    t0 = time.time() - 2
                    _lanzar_lo(soffice, uri, tmp, [src])
                    pdf = tmp / "origen.pdf"
                    if _generado(pdf, t0):
                        shutil.copyfile(pdf, d)
                        ok += 1
                        salidas[o] = d
                        emit("log", f"OK  {o.name}")
                    else:
                        raise RuntimeError("LibreOffice no pudo convertirlo (¿dañado, con "
                                           "contraseña o abierto en otro programa?)")
                except Exception as e:
                    errores += 1
                    emit("log", f"ERROR  {o.name} -> {e}")
                finally:
                    shutil.rmtree(tmp, ignore_errors=True)
        finally:
            shutil.rmtree(perfil, ignore_errors=True)

    emit("log", f"Conversión: {ok} convertidos, {salteados} salteados, {errores} con error.")
    if errores:
        emit("ayuda", (f"{errores} archivo(s) no se pudieron convertir; mirá cuáles en el "
                       "registro.\n\nCausas frecuentes: el archivo está abierto en otro "
                       "programa, tiene contraseña o está dañado. Cerralo, y probá de nuevo "
                       "solo con esos."))
    return [salidas[o] for o in archivos if o in salidas]


# ------------------------------------------------------------- motor PDF (PyMuPDF)
CODIGO_PDF = r'''
import json, os, re, sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf


def abrir(ruta):
    d = pymupdf.open(ruta)
    if d.needs_pass and not d.authenticate(""):
        raise ValueError("el PDF tiene contraseña")
    return d


def rango(texto, n):
    """'1-3,5,8-' -> lista de páginas (base 0). Vacío o 'todas' -> todas."""
    t = (texto or "").strip().lower()
    if t in ("", "todas", "all"):
        return list(range(n))
    out = []
    for parte in t.replace(";", ",").split(","):
        parte = parte.strip()
        if not parte:
            continue
        m = re.fullmatch(r"(\d*)\s*-\s*(\d*)", parte)
        if m:
            a = int(m.group(1)) if m.group(1) else 1
            b = int(m.group(2)) if m.group(2) else n
        elif parte.isdigit():
            a = b = int(parte)
        else:
            raise ValueError("intervalo no válido: " + parte)
        if a < 1 or b < a or b > n:
            raise ValueError("intervalo %s fuera de rango (el PDF tiene %d páginas)" % (parte, n))
        out.extend(range(a - 1, b))
    if not out:
        raise ValueError("intervalo vacío")
    return out


def insertar(dst, src, paginas):
    i = 0
    while i < len(paginas):
        j = i
        while j + 1 < len(paginas) and paginas[j + 1] == paginas[j] + 1:
            j += 1
        dst.insert_pdf(src, from_page=paginas[i], to_page=paginas[j])
        i = j + 1


def guardar(doc, ruta, **kw):
    os.makedirs(os.path.dirname(os.path.abspath(ruta)), exist_ok=True)
    tmp = ruta + ".part"
    doc.save(tmp, **kw)
    doc.close()
    os.replace(tmp, ruta)


def libre(ruta):
    if not os.path.exists(ruta):
        return ruta
    base, ext = os.path.splitext(ruta)
    n = 2
    while os.path.exists("%s (%d)%s" % (base, n, ext)):
        n += 1
    return "%s (%d)%s" % (base, n, ext)


def blanca_como(doc, pagina):
    doc.new_page(-1, width=pagina.rect.width, height=pagina.rect.height)


def op_info(c):
    res = {}
    for ruta in c["rutas"]:
        try:
            d = abrir(ruta)
            res[ruta] = {"paginas": len(d)}
            d.close()
        except Exception as e:
            res[ruta] = {"error": str(e)[:150]}
    return {"info": res}


def op_merge(c):
    out = pymupdf.open()
    origen, errores, blancos = [], [], 0
    for i, e in enumerate(c["entradas"]):
        try:
            src = abrir(e["ruta"])
            pags = rango(e.get("rango", ""), len(src))
        except Exception as ex:
            errores.append([i, str(ex)[:200]])
            continue
        ini = len(out)
        insertar(out, src, pags)
        n = len(out) - ini
        origen += [os.path.splitext(os.path.basename(e["ruta"]))[0]] * n
        if c.get("blanco") and n % 2 == 1:
            blanca_como(out, out[len(out) - 1])
            origen.append("")
            blancos += 1
        src.close()
    if len(out) == 0:
        raise ValueError("ninguno de los PDF se pudo leer")
    if c.get("total_par") and len(out) % 2 == 1:
        blanca_como(out, out[len(out) - 1])
        origen.append("")
        blancos += 1
    pie = c.get("pie") or ""
    if pie:
        # Fuente incrustada con Unicode: la Helvetica básica de PDF no tiene la raya
        # larga y la imprimía como "·" (antes, como "?").
        fuente = pymupdf.Font("helv")
        total = len(out)
        for k in range(total):
            pg = out[k]
            txt = (pie.replace("[PAGINA]", str(k + 1)).replace("[TOTAL]", str(total))
                   .replace("[ARCHIVO]", origen[k]))
            pg.insert_font(fontname="pie", fontbuffer=fuente.buffer)
            ancho = fuente.text_length(txt, fontsize=9)
            x = (pg.rect.width - ancho) / 2
            y = pg.rect.height - 18
            pt = pymupdf.Point(x, y) * pg.derotation_matrix
            pg.insert_text(pt, txt, fontsize=9, fontname="pie", color=(0.25, 0.25, 0.25),
                           rotate=pg.rotation)
        try:
            out.subset_fonts()
        except Exception:
            pass
    total = len(out)
    guardar(out, c["salida"], garbage=3, deflate=True)
    return {"total": total, "blancos": blancos, "errores": errores}


def op_thumbs(c):
    d = abrir(c["ruta"])
    os.makedirs(c["dir"], exist_ok=True)
    archivos = []
    for i, p in enumerate(d):
        z = max(0.08, min(1.0, float(c.get("ancho", 150)) / max(p.rect.width, 1)))
        pix = p.get_pixmap(matrix=pymupdf.Matrix(z, z), alpha=False)
        f = os.path.join(c["dir"], "p%04d.png" % (i + 1))
        pix.save(f)
        archivos.append(f)
    return {"archivos": archivos}


def op_delete(c):
    d = abrir(c["ruta"])
    n = len(d)
    borrar = sorted(set(int(x) for x in c["borrar"]))
    if not borrar:
        raise ValueError("no marcaste páginas para eliminar")
    if len(borrar) >= n:
        raise ValueError("no se pueden eliminar todas las páginas")
    d.delete_pages(borrar)
    quedan = len(d)
    guardar(d, c["salida"], garbage=3, deflate=True)
    return {"quedan": quedan, "eliminadas": len(borrar)}


def op_split(c):
    d = abrir(c["ruta"])
    n = len(d)
    stem = os.path.splitext(os.path.basename(c["ruta"]))[0]
    partes = []
    if c["modo"] == "cada_n":
        k = max(1, int(c["n"]))
        partes = [list(range(i, min(i + k, n))) for i in range(0, n, k)]
    else:
        for trozo in (c.get("rangos") or "").split(";"):
            if trozo.strip():
                partes.append(rango(trozo, n))
    if not partes:
        raise ValueError("no hay partes para generar")
    os.makedirs(c["dir"], exist_ok=True)
    archivos = []
    for k, pags in enumerate(partes, 1):
        nd = pymupdf.open()
        insertar(nd, d, pags)
        destino = libre(os.path.join(c["dir"], "%s_parte_%02d.pdf" % (stem, k)))
        guardar(nd, destino, garbage=3, deflate=True)
        archivos.append(destino)
    return {"archivos": archivos}


def op_compress(c):
    d = abrir(c["ruta"])
    antes = os.path.getsize(c["ruta"])
    nivel = c.get("nivel", "media")
    nota = ""
    if nivel in ("media", "alta"):
        umbral, objetivo, calidad = (170, 150, 75) if nivel == "media" else (120, 96, 55)
        try:
            d.rewrite_images(dpi_threshold=umbral, dpi_target=objetivo, quality=calidad)
        except Exception as e:
            nota = "no se pudieron recomprimir imágenes (%s)" % str(e)[:80]
    try:
        d.subset_fonts()
    except Exception:
        pass
    destino = c["salida"]
    guardar(d, destino, garbage=4, deflate=True, clean=True,
            deflate_images=True, deflate_fonts=True)
    return {"antes": antes, "despues": os.path.getsize(destino), "nota": nota}


# ---- escáner de fotos (necesita Pillow; se importa recién acá para no exigirlo al resto)
LADO_MAX = 2480  # lado mayor de la hoja procesada (A4 a 300 dpi)


def _pil():
    from PIL import Image, ImageChops, ImageFilter, ImageOps
    return Image, ImageChops, ImageFilter, ImageOps


def _cte(Image, grupo, nombre):
    """Constantes de Pillow: en las versiones nuevas viven en un enum, en las viejas en Image."""
    return getattr(getattr(Image, grupo, Image), nombre)


def abrir_foto(ruta):
    Image, _, _, ImageOps = _pil()
    im = Image.open(ruta)
    im.load()
    im = ImageOps.exif_transpose(im) or im
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        fondo = Image.new("RGB", im.size, (255, 255, 255))
        fondo.paste(im, mask=im.split()[3])
        return fondo
    return im if im.mode == "RGB" else im.convert("RGB")


def umbral_otsu(hist):
    """Umbral que mejor separa claros de oscuros, calculado del histograma (0-255)."""
    total = sum(hist)
    suma = sum(i * h for i, h in enumerate(hist))
    sb = wb = 0
    mejor, umbral = -1.0, 127
    for t in range(256):
        wb += hist[t]
        if wb == 0:
            continue
        wf = total - wb
        if wf == 0:
            break
        sb += t * hist[t]
        var = wb * wf * (sb / wb - (suma - sb) / wf) ** 2
        if var > mejor:
            mejor, umbral = var, t
    return umbral


def esquinas_completas(w, h):
    return [[0, 0], [w, 0], [w, h], [0, h]]


def _area(p):
    return abs(sum(p[i][0] * p[(i + 1) % 4][1] - p[(i + 1) % 4][0] * p[i][1] for i in range(4))) / 2


def detectar_hoja(im):
    """Esquinas [sup-izq, sup-der, inf-der, inf-izq] de la hoja clara sobre fondo más oscuro,
    en coordenadas de la imagen. Si no se distingue una hoja, devuelve la imagen completa."""
    Image, _, ImageFilter, _ = _pil()
    W, H = im.size
    completa = esquinas_completas(W, H)
    k = min(1.0, 400.0 / max(W, H))
    sw, sh = max(1, round(W * k)), max(1, round(H * k))
    g = im.convert("L").resize((sw, sh), _cte(Image, "Resampling", "BILINEAR"))
    g = g.filter(ImageFilter.GaussianBlur(1.5))
    hist = g.histogram()
    bajo = next(i for i, v in enumerate(hist) if v)
    alto = 255 - next(i for i, v in enumerate(reversed(hist)) if v)
    if alto - bajo < 40:
        return completa  # sin contraste
    t = umbral_otsu(hist)
    mascara = g.point([0] * (t + 1) + [255] * (255 - t))
    mascara = mascara.filter(ImageFilter.MinFilter(5)).filter(ImageFilter.MaxFilter(5))
    datos = mascara.tobytes()
    # por fila alcanzan el primer y el último píxel de papel para hallar los extremos
    sup_izq = inf_der = sup_der = inf_izq = None
    for y in range(sh):
        fila = datos[y * sw:(y + 1) * sw]
        a = fila.find(b"\xff")
        if a < 0:
            continue
        b = fila.rfind(b"\xff")
        for x in (a, b):
            s, d = x + y, x - y
            if sup_izq is None or s < sup_izq[0]:
                sup_izq = (s, x, y)
            if inf_der is None or s > inf_der[0]:
                inf_der = (s, x, y)
            if sup_der is None or d > sup_der[0]:
                sup_der = (d, x, y)
            if inf_izq is None or d < inf_izq[0]:
                inf_izq = (d, x, y)
    if sup_izq is None:
        return completa
    pts = [(p[1], p[2]) for p in (sup_izq, sup_der, inf_der, inf_izq)]
    if not 0.20 <= _area(pts) / float(sw * sh) <= 0.98:
        return completa
    return [[int(min(max(x / k, 0), W)), int(min(max(y / k, 0), H))] for x, y in pts]


def resolver(A, b):
    """Sistema lineal por eliminación gaussiana con pivoteo (sin numpy)."""
    n = len(b)
    M = [fila[:] + [b[i]] for i, fila in enumerate(A)]
    for c in range(n):
        p = max(range(c, n), key=lambda r: abs(M[r][c]))
        if abs(M[p][c]) < 1e-9:
            raise ValueError("las esquinas no forman un cuadrilátero válido")
        M[c], M[p] = M[p], M[c]
        for r in range(c + 1, n):
            f = M[r][c] / M[c][c]
            if f:
                for j in range(c, n + 1):
                    M[r][j] -= f * M[c][j]
    x = [0.0] * n
    for i in range(n - 1, -1, -1):
        x[i] = (M[i][n] - sum(M[i][j] * x[j] for j in range(i + 1, n))) / M[i][i]
    return x


def coef_perspectiva(destino, origen):
    """Coeficientes para Image.transform(PERSPECTIVE): cada punto del destino sale de
    su punto de origen."""
    A, b = [], []
    for (X, Y), (x, y) in zip(destino, origen):
        A.append([X, Y, 1, 0, 0, 0, -x * X, -x * Y])
        b.append(x)
        A.append([0, 0, 0, X, Y, 1, -y * X, -y * Y])
        b.append(y)
    return resolver(A, b)


def tam_salida(q):
    import math
    d = lambda a, b: math.hypot(a[0] - b[0], a[1] - b[1])
    w = max(d(q[0], q[1]), d(q[3], q[2]))
    h = max(d(q[0], q[3]), d(q[1], q[2]))
    if max(w, h) > LADO_MAX:
        e = LADO_MAX / max(w, h)
        w, h = w * e, h * e
    return max(1, round(w)), max(1, round(h))


def aclarar(im):
    """Blanquea el papel y saca sombras: divide cada canal por el fondo estimado (la imagen
    sin las letras y muy desenfocada). La división se hace con tablas en escala logarítmica
    porque Pillow no divide imágenes y Python puro sería lentísimo."""
    import math
    Image, ImageChops, ImageFilter, ImageOps = _pil()
    w, h = im.size
    f = max(1, round(max(w, h) / 600))
    chica = im.resize((max(1, w // f), max(1, h // f)), _cte(Image, "Resampling", "BILINEAR"))
    chica = chica.filter(ImageFilter.MaxFilter(9))  # borra las letras del cálculo del fondo
    chica = chica.filter(ImageFilter.GaussianBlur(max(2.0, max(w, h) / 30.0 / f)))
    fondo = chica.resize((w, h), _cte(Image, "Resampling", "BILINEAR"))
    esc = math.log(256)
    log = [round(255 * math.log(1 + i) / esc) for i in range(256)]
    exp = [round(255 * math.exp(-d * esc / 255)) for d in range(256)]
    n = len(im.getbands())  # una tabla de 256 valores por canal
    d = ImageChops.subtract(fondo.point(log * n), im.point(log * n))
    out = d.point(exp * n)
    out = ImageOps.autocontrast(out, cutoff=1)
    out = out.point([min(255, round(i * 255 / 235)) for i in range(256)] * n)  # papel a blanco
    return out.filter(ImageFilter.UnsharpMask(radius=1.5, percent=90, threshold=2))


def guardar_vista(im, ruta, lado):
    v = im.copy()
    v.thumbnail((lado, lado))
    v.save(ruta, "PNG")
    return ruta, v.size


def op_img_preparar(c):
    os.makedirs(c["dir"], exist_ok=True)
    lado = int(c.get("lado", 900))
    fotos = []
    for i, ruta in enumerate(c["rutas"]):
        try:
            im = abrir_foto(ruta)
            vista, (vw, vh) = guardar_vista(
                im, os.path.join(c["dir"], "%s%03d.png" % (c.get("prefijo", "v"), i)), lado)
            fotos.append({"ruta": ruta, "ancho": im.size[0], "alto": im.size[1], "vista": vista,
                          "vista_ancho": vw, "vista_alto": vh, "esquinas": detectar_hoja(im)})
        except Exception as e:
            fotos.append({"ruta": ruta, "error": "%s: %s" % (type(e).__name__, str(e)[:150])})
    return {"fotos": fotos}


def op_img_procesar(c):
    Image = _pil()[0]
    im = abrir_foto(c["ruta"])
    q = [(float(x), float(y)) for x, y in c["esquinas"]]
    w, h = tam_salida(q)
    coef = coef_perspectiva([(0, 0), (w, 0), (w, h), (0, h)], q)
    out = im.transform((w, h), _cte(Image, "Transform", "PERSPECTIVE"), coef,
                       _cte(Image, "Resampling", "BICUBIC"))
    giro = int(c.get("rotacion", 0)) % 360
    if giro:
        out = out.transpose(_cte(Image, "Transpose", {90: "ROTATE_270", 180: "ROTATE_180",
                                                        270: "ROTATE_90"}[giro]))
    modo = c.get("modo", "aclarar")
    if modo == "aclarar":
        out = aclarar(out)
    elif modo in ("gris", "byn"):
        out = aclarar(out.convert("L"))
        if modo == "byn":
            out = out.point([0] * 190 + [255] * 66)
    elif modo != "original":
        raise ValueError("modo desconocido: " + str(modo))
    os.makedirs(os.path.dirname(os.path.abspath(c["salida"])), exist_ok=True)
    out.save(c["salida"], "JPEG", quality=85)
    res = {"salida": c["salida"], "ancho": out.size[0], "alto": out.size[1]}
    if c.get("vista"):
        res["vista"], (res["vista_ancho"], res["vista_alto"]) = guardar_vista(
            out, c["vista"], int(c.get("lado", 900)))
    return res


def op_img_pdf(c):
    Image = _pil()[0]
    if not c["imagenes"]:
        raise ValueError("no hay imágenes para armar el PDF")
    doc = pymupdf.open()
    for ruta in c["imagenes"]:
        with Image.open(ruta) as im:
            w, h = im.size
        if c.get("pagina", "a4") == "foto":
            pg = doc.new_page(width=w * 72.0 / 300, height=h * 72.0 / 300)
        else:
            pg = doc.new_page(width=595, height=842) if h >= w else doc.new_page(width=842, height=595)
        pg.insert_image(pg.rect, filename=ruta, keep_proportion=True)
    n = len(doc)
    guardar(doc, c["salida"], garbage=3, deflate=True)
    return {"paginas": n}


def op_md_pdf(c):
    """HTML ya generado » PDF A4 con márgenes de 50 pt (pymupdf.Story). Sin imágenes."""
    import io
    css = ("body{font-size:11pt;color:#000000;background-color:#ffffff} "
           "h1{font-size:20pt;margin-top:6pt} h2{font-size:16pt;margin-top:6pt} "
           "h3,h4,h5,h6{font-size:13pt} p{margin-top:3pt;margin-bottom:5pt} "
           "code{font-family:monospace;background-color:#eeeeee} "
           "pre{font-family:monospace;font-size:9pt;background-color:#eeeeee;padding:4pt} "
           "blockquote{margin-left:14pt;color:#444444} a{color:#000000} "
           "table{border-collapse:collapse} th,td{border:1px solid #888888;padding:3pt}")
    story = pymupdf.Story(html=c["html"], user_css=css)
    buf = io.BytesIO()
    escritor = pymupdf.DocumentWriter(buf)
    hoja = pymupdf.paper_rect("a4")
    caja = hoja + (50, 50, -50, -50)
    mas, vueltas = 1, 0
    while mas and vueltas < 3000:
        dev = escritor.begin_page(hoja)
        mas, _ = story.place(caja)
        story.draw(dev)
        escritor.end_page()
        vueltas += 1
    escritor.close()
    doc = pymupdf.open("pdf", buf.getvalue())
    n = len(doc)
    guardar(doc, c["salida"], garbage=3, deflate=True)
    return {"paginas": n}


OPS = {"info": op_info, "merge": op_merge, "thumbs": op_thumbs, "delete": op_delete,
       "split": op_split, "compress": op_compress, "img_preparar": op_img_preparar,
       "img_procesar": op_img_procesar, "img_pdf": op_img_pdf, "md_pdf": op_md_pdf}

if __name__ == "__main__":
    cfg = json.load(open(sys.argv[1], encoding="utf-8"))
    try:
        res = OPS[cfg["op"]](cfg)
    except Exception as e:
        res = {"error": str(e)[:400]}
    print("@@JSON@@" + json.dumps(res, ensure_ascii=False))
'''

_PY_PDF = {}  # (módulos,) -> ejecutable de Python que los tiene


def _consola(exe: Path) -> Path:
    if exe.name.lower() == "pythonw.exe" and exe.with_name("python.exe").exists():
        return exe.with_name("python.exe")
    return exe


def _tiene_modulo(exe: Path, modulo: str) -> bool:
    try:
        r = subprocess.run([str(exe), "-c", f"import {modulo}"],
                           capture_output=True, timeout=60, creationflags=NO_WINDOW)
        return r.returncode == 0
    except Exception:
        return False


def _otros_pythons():
    """Ejecutables de Python de la PC distintos del actual."""
    vistos = {str(Path(sys.executable)).lower(), str(_consola(Path(sys.executable))).lower()}
    out = []

    def agregar(ruta):
        if ruta and Path(ruta).exists() and str(Path(ruta)).lower() not in vistos:
            vistos.add(str(Path(ruta)).lower())
            out.append(_consola(Path(ruta)))

    for nombre in ("python", "python3", "pythonw"):
        agregar(shutil.which(nombre))
    bases = [Path.home() / "AppData" / "Local" / "Programs" / "Python",
             Path(os.environ.get("ProgramFiles", r"C:\Program Files")),
             Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))]
    for base in bases:
        if base.is_dir():
            for d in sorted(base.glob("Python3*"), reverse=True):
                agregar(d / "python.exe")
    return out


PAQUETES_PIP = {"pymupdf": "pymupdf", "PIL": "pillow"}
NOMBRES_COMPONENTE = {"pymupdf": "el componente PDF (PyMuPDF)", "PIL": "el componente de imágenes (Pillow)"}


def python_con(modulos, emit, instalar=True):
    """Devuelve un Python que tenga todos los módulos pedidos (instalando con pip los que
    falten en el Python actual), o None."""
    clave = tuple(modulos)
    if _PY_PDF.get(clave) and Path(_PY_PDF[clave]).exists():
        return Path(_PY_PDF[clave])
    completo = lambda exe: all(_tiene_modulo(exe, m) for m in clave)
    actual = _consola(Path(sys.executable))
    if completo(actual):
        _PY_PDF[clave] = str(actual)
        return actual
    for otro in _otros_pythons():
        if completo(otro):
            _PY_PDF[clave] = str(otro)
            return otro
    if not instalar:
        return None
    faltan = [m for m in clave if not _tiene_modulo(actual, m)] or list(clave)
    nombres = " y ".join(NOMBRES_COMPONENTE[m] for m in faltan)
    emit("log", f"Instalando {nombres}, una sola vez; puede tardar un minuto…")
    detalle = ""
    for extra in ([], ["--user"]):
        try:
            r = subprocess.run(
                [str(actual), "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
                 *extra, *[PAQUETES_PIP[m] for m in faltan]],
                capture_output=True, text=True, errors="replace", timeout=600,
                creationflags=NO_WINDOW)
        except Exception as e:
            detalle = str(e)
            continue
        if r.returncode == 0 and completo(actual):
            _PY_PDF[clave] = str(actual)
            return actual
        detalle = (r.stderr or r.stdout or "")[-300:].strip()
    emit("log", f"No pude instalar {nombres}: {detalle}")
    return None


def python_con_pymupdf(emit, instalar=True):
    """Devuelve un Python que tenga PyMuPDF (instalándolo si hace falta), o None."""
    return python_con(("pymupdf",), emit, instalar)


AYUDA_PYMUPDF = ("No pude preparar el componente PDF (PyMuPDF).\n\n"
                 "Solución manual: abrí una consola (cmd) y ejecutá:\n"
                 "python -m pip install pymupdf\n\nDespués volvé a intentar.")
AYUDA_IMAGENES = ("No pude preparar los componentes para escanear fotos (PyMuPDF y Pillow).\n\n"
                  "Solución manual: abrí una consola (cmd) y ejecutá:\n"
                  "python -m pip install pymupdf pillow\n\nDespués volvé a intentar.")


def pdf_op(cfg, emit, timeout=1800, instalar=True):
    """Ejecuta una operación PDF o de imagen en el Python que tenga los módulos necesarios.
    Devuelve un dict; con clave 'error' si algo falló ('no_pymupdf' o 'no_imagenes' si no
    se pudieron preparar los componentes)."""
    imagen = str(cfg.get("op", "")).startswith("img_")
    py = python_con(("pymupdf", "PIL") if imagen else ("pymupdf",), emit, instalar)
    if not py:
        return {"error": "no_imagenes" if imagen else "no_pymupdf"}
    tmp = Path(tempfile.mkdtemp(prefix="pdfop_"))
    try:
        f = tmp / "cfg.json"
        f.write_text(json.dumps(cfg), encoding="utf-8")
        r = subprocess.run([str(py), "-c", CODIGO_PDF, str(f)], capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=timeout,
                           creationflags=NO_WINDOW)
        for linea in reversed(r.stdout.splitlines()):
            if linea.startswith("@@JSON@@"):
                return json.loads(linea[8:])
        return {"error": "el motor PDF no respondió: " + (r.stderr or "")[-300:].strip()}
    except subprocess.TimeoutExpired:
        return {"error": "la operación tardó demasiado y se canceló"}
    except Exception as e:
        return {"error": str(e)}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def nombre_libre(ruta: Path) -> Path:
    if not ruta.exists():
        return ruta
    n = 2
    while True:
        cand = ruta.with_name(f"{ruta.stem} ({n}){ruta.suffix}")
        if not cand.exists():
            return cand
        n += 1


def tam_legible(n):
    n = float(n)
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return f"{n:.0f} {u}" if u == "B" else f"{n:.1f} {u}"
        n /= 1024


def limpiar_nombre(texto, defecto="Fusionado"):
    nombre = re.sub(r'[\\/:*?"<>|]', "_", (texto or "").strip()) or defecto
    return nombre[:-4] if nombre.lower().endswith(".pdf") else nombre


# ------------------------------------------------------------- lector de Markdown
# Parser propio y chico, sin dependencias. Dos pasos: texto » bloques » (segmentos para un
# tk.Text, o HTML para imprimir). Nunca tira excepción: lo que no reconoce queda como texto.
_MD_PUNTUACION = "\\`*_{}[]()#+-.!|>~<\"'&:;=@$%^/,?"
_MD_ENLACE = re.compile(r'\[((?:[^\[\]\\]|\\.)+)\]\(\s*([^)\s]*)(?:\s+"[^"]*")?\s*\)')
_MD_IMAGEN = re.compile(r'!\[((?:[^\[\]\\]|\\.)*)\]\(\s*[^)\s]*(?:\s+"[^"]*")?\s*\)')
_MD_ENCABEZADO = re.compile(r"^ {0,3}(#{1,6})(?:[ \t]+(.*?))?(?:[ \t]+#+)?[ \t]*$")
_MD_HR = re.compile(r"^ {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$")
_MD_VALLA = re.compile(r"^ {0,3}(`{3,}|~{3,})(.*)$")
_MD_ITEM = re.compile(r"^([ \t]*)([-*+]|\d{1,9}[.)])[ \t]+(.*)$")
_MD_TAREA = re.compile(r"^\[([ xX])\][ \t]+(.*)$")
_MD_SEPARADOR = re.compile(r"^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$")


def _md_url_util(url):
    """Devuelve la URL si se puede abrir (http/https o ruta relativa), si no, None."""
    url = (url or "").strip()
    if not url or url.startswith("#"):
        return None
    m = re.match(r"^([a-zA-Z][a-zA-Z0-9+.-]+):", url)
    if m:
        return url if m.group(1).lower() in ("http", "https") else None
    return url


def _md_inline(s, tags=()):
    """Texto con marcas » lista de (texto, tags, url). Tags: b, i, code, link."""
    out = []

    def poner(txt, tg, url=None):
        if not txt:
            return
        if out and out[-1][1] == tg and out[-1][2] == url:
            out[-1] = (out[-1][0] + txt, tg, url)
        else:
            out.append((txt, tg, url))

    i, n = 0, len(s)
    while i < n:
        c = s[i]
        if c == "\\" and i + 1 < n and s[i + 1] in _MD_PUNTUACION:
            poner(s[i + 1], tags)
            i += 2
            continue
        if c == "`":
            k = 1
            while i + k < n and s[i + k] == "`":
                k += 1
            cierre = s.find("`" * k, i + k)
            if cierre > i + k - 1:
                poner(s[i + k:cierre].strip() or s[i + k:cierre], tags + ("code",))
                i = cierre + k
                continue
            poner("`" * k, tags)
            i += k
            continue
        if c == "!":
            m = _MD_IMAGEN.match(s, i)
            if m:
                alt = re.sub(r"\\(.)", r"\1", m.group(1)).strip()
                poner("[imagen: %s]" % alt if alt else "[imagen]", tags)
                i = m.end()
                continue
        if c == "[":
            m = _MD_ENLACE.match(s, i)
            if m:
                url = _md_url_util(m.group(2))
                if url:
                    for t, tg, u in _md_inline(m.group(1), tags + ("link",)):
                        poner(t, tg, url)
                else:
                    for t, tg, u in _md_inline(m.group(1), tags):
                        poner(t, tg, u)
                i = m.end()
                continue
        if c in "*_" and s.startswith(c * 3, i) and (c == "*" or i == 0 or not s[i - 1].isalnum()):
            j = s.find(c * 3, i + 3)
            if j > i + 3 and not s[i + 3].isspace():
                for t, tg, u in _md_inline(s[i + 3:j], tags + ("b", "i")):
                    poner(t, tg, u)
                i = j + 3
                continue
        if c in "*_" and i + 1 < n:
            doble = s[i + 1] == c
            marca = c * (2 if doble else 1)
            ini = i + len(marca)
            antes_ok = c == "*" or i == 0 or not s[i - 1].isalnum()
            if antes_ok and ini < n and not s[ini].isspace():
                j = s.find(marca, ini)
                while j != -1 and (s[j - 1].isspace() or (c == "_" and j + len(marca) < n
                                                          and s[j + len(marca)].isalnum())):
                    j = s.find(marca, j + 1)
                if j > ini:
                    extra = "b" if doble else "i"
                    for t, tg, u in _md_inline(s[ini:j], tags + (extra,)):
                        poner(t, tg, u)
                    i = j + len(marca)
                    continue
        poner(c, tags)
        i += 1
    return out


def _md_plano(s):
    return "".join(t for t, _, _ in _md_inline(s))


def _md_celdas(linea):
    t = linea.strip()
    if t.startswith("|"):
        t = t[1:]
    if t.endswith("|") and not t.endswith("\\|"):
        t = t[:-1]
    return [c.strip().replace("\\|", "|") for c in re.split(r"(?<!\\)\|", t)]


def _md_bloques(texto):
    """Texto » lista de bloques: ('h', n, txt) ('p', txt) ('code', lang, txt) ('quote', [txt])
    ('li', nivel, ordenada, marca, tarea, txt) ('hr',) ('table', filas, alineaciones)."""
    lineas = str(texto).replace("\r\n", "\n").replace("\r", "\n").replace("\t", "    ").split("\n")
    bloques, par, pila = [], [], []
    quote = None
    i, n = 0, len(lineas)
    hubo_blanco = True

    def cerrar_par():
        nonlocal par
        if par:
            bloques.append(("p", "\n".join(par)))
            par = []

    def cerrar_quote():
        nonlocal quote
        if quote is not None:
            bloques.append(("quote", quote))
            quote = None

    while i < n:
        linea = lineas[i]
        if not linea.strip():
            cerrar_par()
            cerrar_quote()
            hubo_blanco = True
            i += 1
            continue
        m = _MD_VALLA.match(linea)
        if m and not (m.group(1)[0] == "`" and "`" in m.group(2)):
            cerrar_par()
            cerrar_quote()
            valla = m.group(1)
            sangria = len(linea) - len(linea.lstrip(" "))
            cuerpo, i = [], i + 1
            while i < n:
                t = lineas[i].strip()
                if t.startswith(valla[0] * len(valla)) and not t.strip(valla[0]):
                    i += 1
                    break
                cuerpo.append(lineas[i][min(sangria, len(lineas[i]) - len(lineas[i].lstrip(" "))):])
                i += 1
            bloques.append(("code", m.group(2).strip(), "\n".join(cuerpo)))
            pila, hubo_blanco = [], False
            continue
        ultimo = bloques[-1][0] if bloques else None
        if linea.startswith("    ") and not par and quote is None and not _MD_ITEM.match(linea) \
                and (hubo_blanco or ultimo not in ("li",)):
            cerrar_quote()
            cuerpo = []
            while i < n and (lineas[i].startswith("    ") or not lineas[i].strip()):
                cuerpo.append(lineas[i][4:])
                i += 1
            while cuerpo and not cuerpo[-1].strip():
                cuerpo.pop()
            bloques.append(("code", "", "\n".join(cuerpo)))
            hubo_blanco = False
            continue
        hubo_blanco = False
        m = _MD_ENCABEZADO.match(linea)
        if m:
            cerrar_par()
            cerrar_quote()
            bloques.append(("h", len(m.group(1)), (m.group(2) or "").strip()))
            pila = []
            i += 1
            continue
        if _MD_HR.match(linea):
            cerrar_par()
            cerrar_quote()
            bloques.append(("hr",))
            pila = []
            i += 1
            continue
        if linea.lstrip().startswith(">"):
            cerrar_par()
            if quote is None:
                quote = []
            t = linea.lstrip()[1:]
            quote.append(t[1:] if t.startswith(" ") else t)
            i += 1
            continue
        if "|" in linea and i + 1 < n and "|" in lineas[i + 1] and _MD_SEPARADOR.match(lineas[i + 1]) \
                and "-" in lineas[i + 1]:
            cerrar_par()
            cerrar_quote()
            cab = _md_celdas(linea)
            alin = []
            for c in _md_celdas(lineas[i + 1]):
                alin.append("c" if c.startswith(":") and c.endswith(":") and len(c) > 1 else
                            "r" if c.endswith(":") else "l")
            filas, i = [cab], i + 2
            while i < n and lineas[i].strip() and "|" in lineas[i]:
                filas.append(_md_celdas(lineas[i]))
                i += 1
            nc = len(cab)
            filas = [(f + [""] * nc)[:nc] for f in filas]
            bloques.append(("table", filas, (alin + ["l"] * nc)[:nc]))
            pila = []
            continue
        m = _MD_ITEM.match(linea)
        if m and not (m.group(2) in "*-+" and _MD_HR.match(linea)):
            cerrar_par()
            cerrar_quote()
            ind = len(m.group(1))
            while pila and pila[-1] > ind:
                pila.pop()
            if not pila or pila[-1] < ind:
                pila.append(ind)
            marca, cuerpo = m.group(2), m.group(3)
            ordenada = marca[0].isdigit()
            tarea = None
            mt = _MD_TAREA.match(cuerpo)
            if mt:
                tarea, cuerpo = mt.group(1).lower() == "x", mt.group(2)
            bloques.append(["li", len(pila) - 1, ordenada, marca if ordenada else "", tarea,
                            cuerpo.strip()])
            i += 1
            continue
        if quote is not None:
            quote[-1] = quote[-1] + "\n" + linea.strip() if quote[-1].strip() else linea.strip()
            i += 1
            continue
        if not par and bloques and bloques[-1][0] == "li" and linea.startswith(" "):
            bloques[-1][5] += " " + linea.strip()  # continuación del ítem
            i += 1
            continue
        if not par and bloques and bloques[-1][0] == "li" and ultimo == "li":
            bloques[-1][5] += " " + linea.strip()  # continuación perezosa
            i += 1
            continue
        pila = []
        par.append(linea.strip() if not linea.endswith("  ") else linea.strip() + "  ")
        i += 1
    cerrar_par()
    cerrar_quote()
    return [tuple(b) for b in bloques]


def _md_parrafo(txt):
    """Une las líneas de un párrafo: salto suave » espacio; dos espacios o \\ al final » salto."""
    partes = txt.split("\n")
    out = partes[0]
    for p in partes[1:]:
        if out.endswith("  ") or out.endswith("\\"):
            out = out.rstrip(" ").rstrip("\\") + "\n" + p
        else:
            out = out.rstrip() + " " + p
    return out.rstrip()


def _md_tabla_texto(filas, alin):
    planas = [[_md_plano(c) for c in f] for f in filas]
    anchos = [max(3, max(len(f[k]) for f in planas)) for k in range(len(alin))]

    def celda(t, k):
        return (t.rjust(anchos[k]) if alin[k] == "r" else t.center(anchos[k]) if alin[k] == "c"
                else t.ljust(anchos[k]))
    lineas = ["| " + " | ".join(celda(t, k) for k, t in enumerate(planas[0])) + " |",
              "|" + "|".join("-" * (a + 2) for a in anchos) + "|"]
    for f in planas[1:]:
        lineas.append("| " + " | ".join(celda(t, k) for k, t in enumerate(f)) + " |")
    return lineas


def _md_tags(tg):
    """Tags de un fragmento: negrita y cursiva juntas se vuelven 'bi'."""
    if "b" in tg and "i" in tg:
        tg = tuple(t for t in tg if t not in ("b", "i")) + ("bi",)
    return tg


def markdown_a_bloques(texto):
    """Markdown » lista de segmentos (texto, tags) listos para un tk.Text. Los enlaces llevan
    el tag 'link' y otro 'url:<destino>'. Nunca falla: lo desconocido sale como texto."""
    try:
        return _md_segmentos(texto)
    except Exception:
        return [(str(texto), ())]


def _md_segmentos(texto):
    seg = []

    def inline(s, base=(), solo_enlaces=False):
        for t, tg, url in _md_inline(s):
            if solo_enlaces:
                tg = tuple(x for x in tg if x == "link")
            tg = _md_tags(tuple(base) + tg)
            if url:
                tg = tg + ("url:" + url,)
            seg.append((t, tg))

    bloques = _md_bloques(texto)
    previo = None
    for b in bloques:
        tipo = b[0]
        if previo is not None and not (tipo == "li" and previo == "li"):
            seg.append(("\n", ()))
        previo = tipo
        if tipo == "h":
            nombre = "h%d" % min(b[1], 3)
            inline(b[2], (nombre,), solo_enlaces=True)
            seg.append(("\n", (nombre,)))
        elif tipo == "p":
            inline(_md_parrafo(b[1]))
            seg.append(("\n", ()))
        elif tipo == "code":
            for linea in (b[2].split("\n") if b[2] else [""]):
                seg.append((linea + "\n", ("codeblock",)))
        elif tipo == "quote":
            for k, linea in enumerate(b[1]):
                if not linea.strip():
                    seg.append(("\n", ("quote",)))
                    continue
                for j, parte in enumerate(linea.split("\n")):
                    inline(parte, ("quote",))
                    seg.append(("\n", ("quote",)))
        elif tipo == "li":
            _, nivel, ordenada, marca, tarea, txt = b
            etiqueta = marca + " " if ordenada else "• "
            if tarea is not None:
                etiqueta += "[x] " if tarea else "[ ] "
            base = ("li", "nivel%d" % min(nivel, 6))
            seg.append((etiqueta, base))
            inline(txt, base)
            seg.append(("\n", base))
        elif tipo == "hr":
            seg.append(("-" * 48 + "\n", ("hr",)))
        elif tipo == "table":
            for linea in _md_tabla_texto(b[1], b[2]):
                seg.append((linea + "\n", ("tabla",)))
    return seg


def _html(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _md_inline_html(s):
    out = []
    for t, tg, url in _md_inline(s):
        h = _html(t)
        if "code" in tg:
            h = "<code>%s</code>" % h
        if "i" in tg:
            h = "<i>%s</i>" % h
        if "b" in tg:
            h = "<b>%s</b>" % h
        if "link" in tg and url and url.lower().startswith(("http://", "https://")):
            h = '<a href="%s">%s</a>' % (_html(url), h)
        out.append(h)
    return "".join(out)


def markdown_a_html(texto, titulo=""):
    """Markdown » HTML simple (con todo escapado) para renderizar a PDF."""
    try:
        bloques = _md_bloques(texto)
    except Exception:
        bloques = [("p", str(texto))]
    out = []
    if titulo:
        out.append("<!-- %s -->" % _html(titulo).replace("--", "- -"))
    pila = []

    def cerrar_listas(hasta=0):
        while len(pila) > hasta:
            out.append("</li></%s>" % pila.pop())
    for b in bloques:
        tipo = b[0]
        if tipo != "li":
            cerrar_listas()
        if tipo == "h":
            out.append("<h%d>%s</h%d>" % (b[1], _md_inline_html(b[2]), b[1]))
        elif tipo == "p":
            out.append("<p>%s</p>" % _md_inline_html(_md_parrafo(b[1])).replace("\n", "<br>"))
        elif tipo == "code":
            out.append("<pre><code>%s</code></pre>" % _html(b[2]))
        elif tipo == "quote":
            out.append("<blockquote>%s</blockquote>" % "<br>".join(
                _md_inline_html(l) if l.strip() else "<br>" for l in b[1]))
        elif tipo == "li":
            _, nivel, ordenada, marca, tarea, txt = b
            lista = "ol" if ordenada else "ul"
            nivel = min(nivel, 8)
            if len(pila) > nivel + 1:
                cerrar_listas(nivel + 1)
            if len(pila) == nivel + 1:
                out.append("</li>")
                if pila[-1] != lista:
                    out.append("</%s>" % pila.pop())
            while len(pila) < nivel + 1:
                pila.append(lista)
                out.append("<%s>" % lista)
            caja = "" if tarea is None else ("[x] " if tarea else "[ ] ")
            out.append("<li>%s%s" % (caja, _md_inline_html(txt)))
        elif tipo == "hr":
            out.append("<hr>")
        elif tipo == "table":
            filas, alin = b[1], b[2]
            lado = {"l": "left", "r": "right", "c": "center"}
            html = ["<table>"]
            for k, f in enumerate(filas):
                celda = "th" if k == 0 else "td"
                html.append("<tr>%s</tr>" % "".join(
                    '<%s style="text-align:%s">%s</%s>' % (celda, lado[alin[j]], _md_inline_html(c), celda)
                    for j, c in enumerate(f)))
            html.append("</table>")
            out.append("".join(html))
    cerrar_listas()
    return "\n".join(out)


def contar_palabras(texto):
    return sum(1 for p in str(texto).split() if any(c.isalnum() for c in p))


# --------------------------------------------------------------------- PDFsam
def buscar_pdfsam():
    for nombre in ("pdfsam-basic", "pdfsam"):
        ruta = shutil.which(nombre)
        if ruta:
            return ruta
    bases = []
    for var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432", "LOCALAPPDATA"):
        b = os.environ.get(var)
        if b:
            bases += [Path(b), Path(b) / "Programs"]
    for base in bases:
        if not base.is_dir():
            continue
        for carpeta in base.glob("PDFsam*"):
            exes = [e for e in carpeta.glob("*.exe")
                    if e.name.lower().startswith("pdfsam") and "unins" not in e.name.lower()]
            if exes:
                return str(sorted(exes, key=lambda e: len(e.name))[0])
    for c in ("/opt/pdfsam-basic/bin/pdfsam-basic",):
        if Path(c).exists():
            return c
    return None


# ------------------------------------------------------- operaciones de alto nivel
def trabajo(archivos, base, opc, instalar, emit):
    """Pestaña 1: convertir y (opcional) fusionar.
    opc: sub, sobre, fusionar, nombre, blanco, total_par, pie, rangos {Path: texto}."""
    try:
        necesita_lo = any(a.suffix.lower() in CONVERTIBLES for a in archivos)
        soffice = None
        if necesita_lo:
            if instalar and not instalar_libreoffice(emit):
                emit("fin", "No se pudo instalar LibreOffice.")
                return
            soffice = buscar_soffice()
            if not soffice:
                emit("fin", "No encuentro LibreOffice.")
                return
        emit("log", "--- Procesando ---")
        pdfs = convertir_lote(archivos, base, opc["sub"], opc["sobre"], soffice, emit)
        abrir = None
        if pdfs:
            abrir = pdfs[0] if len(pdfs) == 1 else carpeta_destino(base, opc["sub"])

        if opc["fusionar"]:
            if len(pdfs) < 2:
                emit("log", "Fusión omitida: hacen falta al menos 2 PDF.")
            else:
                carpeta = carpeta_destino(base, opc["sub"])
                salida = carpeta / f"{limpiar_nombre(opc['nombre'])}.pdf"
                if not opc["sobre"]:
                    salida = nombre_libre(salida)
                # PDF resultante de cada archivo elegido, con su intervalo
                orig_por_pdf = {destino_pdf(a, base, opc["sub"]): a for a in archivos}
                entradas = []
                for p in pdfs:
                    if p.resolve() == salida.resolve():
                        continue
                    entradas.append({"ruta": str(p),
                                     "rango": (opc.get("rangos") or {}).get(orig_por_pdf.get(p), "")})
                carpeta.mkdir(parents=True, exist_ok=True)
                emit("log", f"Fusionando {len(entradas)} PDF…")
                r = pdf_op({"op": "merge", "entradas": entradas, "salida": str(salida),
                            "blanco": opc["blanco"], "total_par": opc.get("total_par", False),
                            "pie": opc.get("pie", "")}, emit)
                if r.get("error") == "no_pymupdf":
                    emit("ayuda", (AYUDA_PYMUPDF, False))
                elif "error" in r:
                    emit("log", f"ERROR en la fusión: {r['error']}")
                else:
                    for i, msg in r.get("errores", []):
                        nombre = Path(entradas[i]["ruta"]).name if 0 <= i < len(entradas) else "?"
                        emit("log", f"  no se pudo incluir {nombre}: {msg}")
                    emit("log", f"FUSIONADO  {salida.name}: {r['total']} páginas"
                                + (f" ({r['blancos']} en blanco agregadas)" if r["blancos"] else ""))
                    emit("fusion", str(salida))
                    abrir = salida
        if abrir:
            emit("resultado", str(abrir))
        emit("fin", "Listo.")
    except Exception as e:
        emit("log", "Error inesperado:\n" + traceback.format_exc())
        emit("fin", f"Error inesperado: {e!r}")


def trabajo_herramienta(op, items, params, emit):
    """Pestaña 2: dividir / comprimir / eliminar páginas. items: lista de Path (PDF)."""
    try:
        emit("max", max(len(items), 1))
        hechos = 0
        abrir = None
        for ruta in items:
            if op == "comprimir":
                destino = nombre_libre(ruta.with_name(f"{ruta.stem} (comprimido).pdf"))
                r = pdf_op({"op": "compress", "ruta": str(ruta), "salida": str(destino),
                            "nivel": params["nivel"]}, emit)
            elif op == "dividir":
                carpeta = ruta.with_name(f"{ruta.stem} (partes)")
                r = pdf_op({"op": "split", "ruta": str(ruta), "dir": str(carpeta),
                            "modo": params["modo"], "n": params.get("n", 1),
                            "rangos": params.get("rangos", "")}, emit)
            else:  # borrar
                if params.get("reemplazar"):
                    destino = ruta
                else:
                    destino = nombre_libre(ruta.with_name(f"{ruta.stem} (sin páginas).pdf"))
                r = pdf_op({"op": "delete", "ruta": str(ruta), "salida": str(destino),
                            "borrar": params["borrar"]}, emit)
            if r.get("error") == "no_pymupdf":
                emit("ayuda", (AYUDA_PYMUPDF, False))
                break
            if "error" in r:
                emit("log", f"ERROR  {ruta.name}: {r['error']}")
            elif op == "comprimir":
                a, d = r["antes"], r["despues"]
                pct = (d / a - 1) * 100 if a else 0
                emit("log", f"OK  {ruta.name}: {tam_legible(a)} » {tam_legible(d)} "
                            f"({pct:+.0f}%)  {r.get('nota', '')}".rstrip())
                abrir = destino
            elif op == "dividir":
                emit("log", f"OK  {ruta.name}: {len(r['archivos'])} partes en “{carpeta.name}”")
                abrir = carpeta
            else:
                emit("log", f"OK  {destino.name}: se eliminaron {r['eliminadas']} páginas, "
                            f"quedan {r['quedan']}")
                abrir = destino
            hechos += 1
            emit("prog", hechos)
        if abrir:
            emit("resultado", str(abrir))
        emit("fin", "Listo.")
    except Exception as e:
        emit("log", "Error inesperado:\n" + traceback.format_exc())
        emit("fin", f"Error inesperado: {e!r}")


# ------------------------------------------------------------ carpetas automáticas
def _estable(p: Path, estado: dict) -> bool:
    """True si el archivo no cambió desde el escaneo anterior y no está bloqueado."""
    try:
        st = p.stat()
    except OSError:
        return False
    firma = (st.st_size, st.st_mtime)
    previa = estado.get(p)
    estado[p] = firma
    if previa != firma or time.time() - st.st_mtime < 2:
        return False
    if ES_WINDOWS:
        try:
            os.rename(p, p)  # falla si otro programa lo está escribiendo
        except OSError:
            return False
    return True


def _mover(origen: Path, carpeta: Path):
    carpeta.mkdir(parents=True, exist_ok=True)
    shutil.move(str(origen), str(nombre_libre(carpeta / origen.name)))


def procesar_regla(regla, estado, soffice, emit):
    """Un ciclo de vigilancia sobre una carpeta. Devuelve cuántos archivos procesó."""
    carpeta = Path(regla["carpeta"])
    if not carpeta.is_dir():
        return 0
    admitidas = EXT if regla.get("convertir") else {".pdf"}
    candidatos = [p for p in carpeta.iterdir()
                  if p.is_file() and p.suffix.lower() in admitidas and not p.name.startswith("~$")]
    for viejo in [k for k in estado if k not in candidatos]:
        estado.pop(viejo, None)
    estables = [p for p in sorted(candidatos, key=_clave_natural) if _estable(p, estado)]
    if not estables:
        return 0
    accion = regla["accion"]
    if accion == "fusionar" and (len(estables) != len(candidatos) or len(estables) < 2):
        return 0  # esperar a que lleguen y terminen de copiarse todos

    salida, hechos, errores = carpeta / "salida", carpeta / "procesados", carpeta / "con_error"
    tmp = Path(tempfile.mkdtemp(prefix="hot_"))
    procesados = 0
    try:
        pdf_de = {}
        conv = [p for p in estables if p.suffix.lower() in CONVERTIBLES]
        if conv:
            if not soffice:
                emit("log", f"[{carpeta.name}] hay Word/PowerPoint pero falta LibreOffice; los dejo sin tocar")
                estables = [p for p in estables if p not in conv]
            else:
                _lanzar_lo(soffice, (tmp / "perfil").as_uri(), tmp, conv)
        for p in estables:
            if p.suffix.lower() == ".pdf":
                pdf_de[p] = p
            elif (tmp / (p.stem + ".pdf")).exists():
                pdf_de[p] = tmp / (p.stem + ".pdf")
            else:
                emit("log", f"[{carpeta.name}] ERROR {p.name}: no se pudo convertir")
                _mover(p, errores)
        listos = [p for p in estables if p in pdf_de]
        if not listos:
            return 0
        salida.mkdir(parents=True, exist_ok=True)

        if accion == "fusionar":
            nombre = f"Fusionado_{datetime.now():%Y%m%d_%H%M%S}.pdf"
            destino = nombre_libre(salida / nombre)
            r = pdf_op({"op": "merge", "entradas": [{"ruta": str(pdf_de[p])} for p in listos],
                        "salida": str(destino), "blanco": regla.get("blanco", True),
                        "total_par": False, "pie": regla.get("pie", "")}, emit, instalar=False)
            if "error" in r:
                emit("log", f"[{carpeta.name}] ERROR al fusionar: {r['error']}")
                for p in listos:
                    _mover(p, errores)
            else:
                emit("log", f"[{carpeta.name}] FUSIONADOS {len(listos)} archivos » salida/{destino.name} "
                            f"({r['total']} páginas)")
                for p in listos:
                    _mover(p, hechos)
                procesados = len(listos)
        else:
            for p in listos:
                if accion == "comprimir":
                    destino = nombre_libre(salida / (p.stem + ".pdf"))
                    r = pdf_op({"op": "compress", "ruta": str(pdf_de[p]), "salida": str(destino),
                                "nivel": regla.get("nivel", "media")}, emit, instalar=False)
                    msg = (f"{tam_legible(r['antes'])} » {tam_legible(r['despues'])}"
                           if "error" not in r else "")
                else:
                    r = pdf_op({"op": "split", "ruta": str(pdf_de[p]),
                                "dir": str(salida / p.stem), "modo": "cada_n",
                                "n": regla.get("n", 1)}, emit, instalar=False)
                    msg = f"{len(r.get('archivos', []))} partes" if "error" not in r else ""
                if "error" in r:
                    emit("log", f"[{carpeta.name}] ERROR {p.name}: {r['error']}")
                    _mover(p, errores)
                else:
                    emit("log", f"[{carpeta.name}] {accion.upper()} {p.name}: {msg}")
                    _mover(p, hechos)
                    procesados += 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return procesados


class Vigilante(threading.Thread):
    def __init__(self, reglas, emit, pausa=3):
        super().__init__(daemon=True)
        self.reglas = [dict(r) for r in reglas]
        self.emit = emit
        self.pausa = pausa
        self.parar = threading.Event()
        self.estados = [dict() for _ in self.reglas]

    def run(self):
        while not self.parar.wait(self.pausa):
            for regla, estado in zip(self.reglas, self.estados):
                try:
                    procesar_regla(regla, estado, buscar_soffice(), self.emit)
                except Exception as e:
                    self.emit("log", f"[{regla.get('carpeta')}] error: {e!r}")


class DialogoPaginas(tk.Toplevel):
    """Miniaturas de un PDF; se marcan las páginas a eliminar."""

    COLS = 5

    def __init__(self, app, ruta: Path):
        super().__init__(app)
        self.app, self.ruta = app, ruta
        self.title(f"Eliminar páginas — {ruta.name}")
        self.geometry("960x700")
        self.configure(background=BG)
        barra_oscura(self)
        self.vars, self.celdas, self.imagenes = [], [], []
        self.tmp = Path(tempfile.mkdtemp(prefix="miniaturas_"))
        self.protocol("WM_DELETE_WINDOW", self._cerrar)

        barra = ttk.Frame(self)
        barra.pack(fill="x", padx=8, pady=6)
        ttk.Button(barra, text="[Marcar impares]", command=lambda: self._marcar(lambda n: n % 2 == 1)).pack(side="left")
        ttk.Button(barra, text="[Marcar pares]", command=lambda: self._marcar(lambda n: n % 2 == 0)).pack(side="left", padx=4)
        ttk.Button(barra, text="[Marcar intervalo…]", command=self._intervalo).pack(side="left")
        ttk.Button(barra, text="[Desmarcar todas]", command=lambda: self._marcar(lambda n: False)).pack(side="left", padx=4)
        self.v_estado = tk.StringVar(value="Generando miniaturas…")
        ttk.Label(barra, textvariable=self.v_estado, style="Muted.TLabel").pack(side="right")

        pie = ttk.Frame(self)
        pie.pack(side="bottom", fill="x", padx=8, pady=6)
        self.v_reemplazar = tk.BooleanVar(value=False)
        ttk.Checkbutton(pie, text="Reemplazar el PDF original (si no, crea “… (sin páginas).pdf”)",
                        variable=self.v_reemplazar).pack(side="left")
        ttk.Button(pie, text="[Cancelar]", command=self._cerrar).pack(side="right")
        self.btn_guardar = ttk.Button(pie, text="[ ELIMINAR LAS MARCADAS ]", command=self._guardar,
                                      style="Accent.TButton", state="disabled")
        self.btn_guardar.pack(side="right", padx=6)

        cuerpo = ttk.Frame(self)
        cuerpo.pack(fill="both", expand=True, padx=8)
        self.canvas = tk.Canvas(cuerpo, background=BG, highlightthickness=0)
        sb = ttk.Scrollbar(cuerpo, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.interior = tk.Frame(self.canvas, background=BG)
        self.canvas.create_window((0, 0), window=self.interior, anchor="nw")
        self.interior.bind("<Configure>",
                           lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))
        self.cola = queue.Queue()
        threading.Thread(target=self._generar, daemon=True).start()
        self.after(150, self._sondear)

    def _generar(self):
        r = pdf_op({"op": "thumbs", "ruta": str(self.ruta), "dir": str(self.tmp), "ancho": 150},
                   lambda t, d: None)
        self.cola.put(r)

    def _sondear(self):
        try:
            r = self.cola.get_nowait()
        except queue.Empty:
            self.after(150, self._sondear)
            return
        if r.get("error") == "no_pymupdf":
            messagebox.showwarning("Falta el componente PDF", AYUDA_PYMUPDF, parent=self)
            self._cerrar()
            return
        if "error" in r:
            messagebox.showerror("No se pudo abrir el PDF", r["error"], parent=self)
            self._cerrar()
            return
        for i, f in enumerate(r["archivos"]):
            img = tk.PhotoImage(file=f)
            self.imagenes.append(img)
            v = tk.BooleanVar(value=False)
            self.vars.append(v)
            celda = tk.Frame(self.interior, background=BG, padx=4, pady=4)
            celda.grid(row=i // self.COLS, column=i % self.COLS, padx=4, pady=4)
            lbl = tk.Label(celda, image=img, borderwidth=0, highlightthickness=4,
                           highlightbackground=LINEA, cursor="hand2")
            lbl.pack()
            lbl.bind("<Button-1>", lambda e, k=i: self._alternar(k))
            ttk.Checkbutton(celda, text=f"Página {i + 1}", variable=v,
                            command=self._refrescar).pack()
            self.celdas.append(lbl)
        self.v_estado.set(f"{len(self.vars)} páginas. Tocá una miniatura para marcarla.")
        self.btn_guardar.config(state="normal")

    def _alternar(self, k):
        self.vars[k].set(not self.vars[k].get())
        self._refrescar()

    def _refrescar(self):
        n = 0
        for v, lbl in zip(self.vars, self.celdas):
            marcada = v.get()
            n += marcada
            lbl.config(highlightbackground=ROJO if marcada else LINEA)
        self.v_estado.set(f"{n} de {len(self.vars)} marcadas para eliminar")

    def _marcar(self, regla):
        for i, v in enumerate(self.vars):
            v.set(regla(i + 1))
        self._refrescar()

    def _intervalo(self):
        t = simpledialog.askstring("Marcar intervalo", "Páginas a marcar (ej.: 1-3,7,10-):",
                                   parent=self)
        if not t:
            return
        n = len(self.vars)
        try:
            marcar = set()
            for parte in t.replace(";", ",").split(","):
                parte = parte.strip()
                if not parte:
                    continue
                m = re.fullmatch(r"(\d*)\s*-\s*(\d*)", parte)
                if m:
                    a = int(m.group(1) or 1)
                    b = int(m.group(2) or n)
                elif parte.isdigit():
                    a = b = int(parte)
                else:
                    raise ValueError(parte)
                marcar.update(range(max(a, 1), min(b, n) + 1))
        except ValueError as e:
            messagebox.showerror("Intervalo no válido", f"No entiendo: {e}", parent=self)
            return
        for i, v in enumerate(self.vars):
            if i + 1 in marcar:
                v.set(True)
        self._refrescar()

    def _guardar(self):
        borrar = [i for i, v in enumerate(self.vars) if v.get()]
        if not borrar:
            messagebox.showinfo("Nada marcado", "Marcá al menos una página.", parent=self)
            return
        if len(borrar) >= len(self.vars):
            messagebox.showwarning("Demasiadas páginas", "No se pueden eliminar todas.", parent=self)
            return
        reemplazar = self.v_reemplazar.get()
        if reemplazar and not messagebox.askyesno(
                "Reemplazar el original",
                f"Se va a sobrescribir {self.ruta.name} sin las {len(borrar)} páginas marcadas.\n"
                "Esto no se puede deshacer. ¿Seguimos?", parent=self):
            return
        self._cerrar()
        self.app.lanzar_herramienta("borrar", [self.ruta],
                                    {"borrar": borrar, "reemplazar": reemplazar})

    def _cerrar(self):
        try:
            self.canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        shutil.rmtree(self.tmp, ignore_errors=True)
        self.destroy()


class Escaner(tk.Toplevel):
    """Escáner de documentos: fotos » hoja recortada y enderezada » PDF (reemplazo local de
    CamScanner). La imagen la procesa el worker con Pillow; acá solo se muestran PNG."""

    EXT = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
    LADO = 500          # lado mayor de las vistas previas
    LONA = (560, 520)   # tamaño del recuadro de vista previa
    ASA = 7             # medio lado de las asas
    MODOS = (("original", "Original"), ("aclarar", "Aclarar"),
             ("byn", "Blanco y negro"), ("gris", "Gris"))

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Escanear fotos a PDF")
        self.geometry("1110x740")
        self.minsize(1000, 680)
        self.configure(background=BG)
        barra_oscura(self)
        self.tmp = Path(tempfile.mkdtemp(prefix="escaner_"))
        self.cola = queue.Queue()
        self.fotos = []      # dicts: ruta, nombre, ancho, alto, vista, vw, vh, esquinas, auto, modo, giro
        self.actual = None   # índice de la foto que se está editando
        self.imagen = None   # la PhotoImage mostrada (hay que retenerla)
        self.origen = (0, 0, 1.0)   # desplazamiento y escala de la vista previa en el recuadro
        self.asa = None      # asa que se está arrastrando
        self.viendo = False  # True: se muestra el resultado, no la edición
        self.resultado = None
        self.creando = False
        self.cancelar = False
        self.contador = 0
        self.v_modo = tk.StringVar(value="aclarar")
        self.v_todas = tk.BooleanVar(value=False)
        self.v_pagina = tk.StringVar(value="a4")
        self.v_estado = tk.StringVar(value="Agregá las fotos de la hoja que querés escanear.")
        self._armar()
        self._mostrar()
        self.bind("<Destroy>", self._al_destruir)
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.after(100, self._sondear)

    # ------------------------------------------------------------------ armado
    def _armar(self):
        cont = ttk.Frame(self, padding=(22, 14, 22, 12))
        cont.pack(fill="both", expand=True)
        ttk.Label(cont, text="ESCÁNER DE DOCUMENTOS", style="Title.TLabel").pack(anchor="w")
        ttk.Label(cont, style="Muted.TLabel",
                  text="Sacale fotos a una hoja, ajustá las esquinas y armá un PDF. Todo pasa en "
                       "esta PC."
                  ).pack(anchor="w", pady=(0, 8))
        # abajo primero, para que no lo tape el resto
        pie = ttk.Frame(cont)
        pie.pack(side="bottom", fill="x", pady=(8, 0))
        ttk.Label(pie, textvariable=self.v_estado, style="Muted.TLabel", wraplength=1000,
                  justify="left").pack(side="bottom", anchor="w", pady=(8, 0))
        fila = ttk.Frame(pie)
        fila.pack(fill="x")
        ttk.Label(fila, text="Tamaño de página:").pack(side="left")
        ttk.Radiobutton(fila, text="A4", value="a4", variable=self.v_pagina
                        ).pack(side="left", padx=(10, 6))
        ttk.Radiobutton(fila, text="Tamaño de la foto", value="foto", variable=self.v_pagina
                        ).pack(side="left", padx=6)
        self.btn_crear = ttk.Button(fila, text="[ CREAR PDF ]", style="Accent.TButton",
                                    command=self.crear_pdf)
        self.btn_crear.pack(side="right")

        medio = ttk.Frame(cont)
        medio.pack(fill="both", expand=True)
        # izquierda: lista de fotos
        izq = ttk.Frame(medio)
        izq.pack(side="left", fill="y")
        ttk.Label(izq, text="FOTOS", style="Section.TLabel").pack(anchor="w")
        ttk.Button(izq, text="[ + AGREGAR FOTOS ]", command=self.agregar_dialogo
                   ).pack(anchor="w", pady=(4, 6))
        cuerpo = tk.Frame(izq, background=LINEA, padx=1, pady=1)
        cuerpo.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(cuerpo)
        sb.pack(side="right", fill="y")
        self.lista = tk.Listbox(cuerpo, selectmode="browse", activestyle="none", width=26,
                                exportselection=False, yscrollcommand=sb.set, relief="flat",
                                borderwidth=0, highlightthickness=0, background=PANEL,
                                foreground=FG, font=(self.app.fam, 10),
                                selectbackground=ACENTO_OSC, selectforeground=ACENTO)
        self.lista.pack(side="left", fill="both", expand=True)
        sb.config(command=self.lista.yview)
        self.lista.bind("<<ListboxSelect>>", self._elegida)
        fb = ttk.Frame(izq)
        fb.pack(fill="x", pady=(6, 0))
        for texto, cmd in (("[Quitar]", self.quitar), ("[Subir]", lambda: self.mover(-1)),
                           ("[Bajar]", lambda: self.mover(1))):
            ttk.Button(fb, text=texto, style="Link.TButton", command=cmd).pack(side="left", padx=(0, 8))
        # derecha: controles (se arma antes que el centro para que el recuadro tome el resto)
        der = ttk.Frame(medio)
        der.pack(side="right", fill="y", padx=(16, 0))
        ttk.Label(der, text="MODO", style="Section.TLabel").pack(anchor="w")
        for valor, texto in self.MODOS:
            ttk.Radiobutton(der, text=texto, value=valor, variable=self.v_modo,
                            command=self._modo_cambiado).pack(anchor="w")
        ttk.Checkbutton(der, text="Aplicar el modo a\ntodas las fotos", variable=self.v_todas,
                        command=self._modo_cambiado).pack(anchor="w", pady=(6, 0))
        ttk.Label(der, text="HOJA", style="Section.TLabel").pack(anchor="w", pady=(14, 2))
        for texto, cmd in (("[Rotar]", self.rotar), ("[Detectar bordes]", self.detectar),
                           ("[Toda la hoja]", self.toda_la_hoja)):
            ttk.Button(der, text=texto, command=cmd).pack(anchor="w", pady=2)
        ttk.Label(der, text="RESULTADO", style="Section.TLabel").pack(anchor="w", pady=(14, 2))
        self.btn_ver = ttk.Button(der, text="[Ver resultado]", command=self.ver_resultado)
        self.btn_ver.pack(anchor="w", pady=2)
        ttk.Label(der, style="Muted.TLabel", justify="left", wraplength=190,
                  text="Arrastrá los cuadrados verdes hasta las esquinas de la hoja. "
                       "El giro se ve en el resultado."
                  ).pack(anchor="w", pady=(14, 0))
        # centro: vista previa con las asas
        marco = tk.Frame(medio, background=LINEA, padx=1, pady=1)
        marco.pack(side="left", padx=16)
        self.lona = tk.Canvas(marco, width=self.LONA[0], height=self.LONA[1], background=PANEL,
                              highlightthickness=0)
        self.lona.pack()
        self.lona.bind("<ButtonPress-1>", self._pulsar)
        self.lona.bind("<B1-Motion>", self._arrastrando)
        self.lona.bind("<ButtonRelease-1>", lambda e: setattr(self, "asa", None))

    # --------------------------------------------------------- cargar fotos
    def agregar_dialogo(self):
        rutas = filedialog.askopenfilenames(
            parent=self, title="Elegí las fotos",
            initialdir=str(self.app.base or self.app.cfg.get("carpeta") or Path.home()),
            filetypes=[("Fotos", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp"),
                       ("Todos los archivos", "*.*")])
        if rutas:
            self.agregar_fotos(rutas)

    def agregar_fotos(self, rutas):
        """Lee las fotos en segundo plano: abre, vista previa y detección de la hoja."""
        buenas, heic, raras = [], [], []
        for r in rutas:
            p = Path(r)
            if p.suffix.lower() in self.EXT:
                buenas.append(str(p))
            elif p.suffix.lower() in (".heic", ".heif"):
                heic.append(p.name)
            else:
                raras.append(p.name)
        if heic:
            messagebox.showwarning(
                "Fotos HEIC", "Estas fotos están en formato HEIC (el de los iPhone) y no las puedo "
                "leer:\n\n" + "\n".join(heic[:8]) + "\n\nExportalas como JPG desde el teléfono o "
                "desde la app Fotos y volvé a agregarlas.", parent=self)
        if raras:
            messagebox.showwarning(
                "Archivos que no son fotos", "No sé leer estos archivos:\n\n" + "\n".join(raras[:8])
                + "\n\nSirven JPG, PNG, BMP, TIF y WEBP.", parent=self)
        if not buenas:
            return
        self.contador += 1
        prefijo = f"v{self.contador}_"
        self.v_estado.set(f"Leyendo {len(buenas)} foto(s)…")
        emit = lambda t, d: self.cola.put(("estado", d)) if t == "log" else None

        def trabajo():
            r = pdf_op({"op": "img_preparar", "rutas": buenas, "dir": str(self.tmp),
                        "prefijo": prefijo, "lado": self.LADO}, emit)
            self.cola.put(("fotos", r))
        threading.Thread(target=trabajo, daemon=True).start()

    def _incorporar(self, r):
        if r.get("error") in ("no_imagenes", "no_pymupdf"):
            messagebox.showwarning("Falta un componente", AYUDA_IMAGENES, parent=self)
            self.v_estado.set("No se pudieron preparar los componentes.")
            return
        if "error" in r:
            messagebox.showerror("No se pudieron leer las fotos", r["error"], parent=self)
            return
        malas = []
        for f in r["fotos"]:
            if "error" in f:
                malas.append(f"{Path(f['ruta']).name}: no se pudo abrir (¿está dañada o no es una "
                             "imagen?)")
                continue
            self.fotos.append({
                "ruta": Path(f["ruta"]), "nombre": Path(f["ruta"]).name, "ancho": f["ancho"],
                "alto": f["alto"], "vista": f["vista"], "vw": f["vista_ancho"],
                "vh": f["vista_alto"], "esquinas": [list(p) for p in f["esquinas"]],
                "auto": [list(p) for p in f["esquinas"]],
                "modo": self.v_modo.get(), "giro": 0})
        if malas:
            messagebox.showwarning("Fotos que no se pudieron abrir", "\n".join(malas[:8]),
                                   parent=self)
        self._llenar_lista(self.actual if self.actual is not None else (0 if self.fotos else None))
        self.v_estado.set(f"{len(self.fotos)} foto(s) en la lista. Ajustá las esquinas de cada "
                          "una si hace falta.")

    # ------------------------------------------------------------------ lista
    def _llenar_lista(self, sel=None):
        self.lista.delete(0, "end")
        for i, f in enumerate(self.fotos):
            self.lista.insert("end", f"{i + 1}. {f['nombre']}")
        self.actual = sel if self.fotos and sel is not None else None
        if self.actual is not None:
            self.actual = min(self.actual, len(self.fotos) - 1)
            self.lista.selection_set(self.actual)
            self.lista.see(self.actual)
            self.v_modo.set(self.fotos[self.actual]["modo"])
        self.viendo = False
        self.btn_ver.config(text="[Ver resultado]")
        self._mostrar()

    def _elegida(self, _=None):
        sel = self.lista.curselection()
        if not sel or sel[0] == self.actual:
            return
        self.actual = sel[0]
        self.v_modo.set(self.fotos[self.actual]["modo"])
        self.viendo = False
        self.btn_ver.config(text="[Ver resultado]")
        self._mostrar()

    def quitar(self):
        if self.actual is None or self._ocupado():
            return
        del self.fotos[self.actual]
        self._llenar_lista(self.actual)

    def mover(self, delta):
        i, j = self.actual, (self.actual or 0) + delta
        if i is None or not 0 <= j < len(self.fotos) or self._ocupado():
            return
        self.fotos[i], self.fotos[j] = self.fotos[j], self.fotos[i]
        self._llenar_lista(j)

    def _ocupado(self):
        if self.creando:
            self.v_estado.set("Esperá a que termine de armarse el PDF.")
        return self.creando

    # ---------------------------------------------------------- vista previa
    def _mostrar(self):
        self.lona.delete("all")
        self.imagen = None
        if self.actual is None:
            self.lona.create_text(self.LONA[0] // 2, self.LONA[1] // 2, fill=GRIS, justify="center",
                                  font=(self.app.fam, 11),
                                  text="SIN FOTOS.\n\nTocá [ + AGREGAR FOTOS ]")
            return
        f = self.fotos[self.actual]
        if self.viendo and self.resultado:
            ruta, (vw, vh) = self.resultado["vista"], (self.resultado["vw"], self.resultado["vh"])
        else:
            ruta, (vw, vh) = f["vista"], (f["vw"], f["vh"])
        self.imagen = tk.PhotoImage(file=ruta)
        ox, oy = (self.LONA[0] - vw) // 2, (self.LONA[1] - vh) // 2
        self.lona.create_image(ox, oy, anchor="nw", image=self.imagen, tags="foto")
        self.origen = (ox, oy, vw / float(f["ancho"]))
        if not self.viendo:
            self._dibujar_asas()

    def _a_lona(self, x, y):
        ox, oy, e = self.origen
        return ox + x * e, oy + y * e

    def _a_original(self, cx, cy):
        ox, oy, e = self.origen
        return (cx - ox) / e, (cy - oy) / e

    def _dibujar_asas(self):
        self.lona.delete("asas")
        f = self.fotos[self.actual]
        puntos = [self._a_lona(*p) for p in f["esquinas"]]
        self.lona.create_polygon([c for p in puntos for c in p], outline=ACENTO, fill="",
                                 width=2, tags=("asas", "poli"))
        for i, (cx, cy) in enumerate(puntos):
            self.lona.create_rectangle(cx - self.ASA, cy - self.ASA, cx + self.ASA, cy + self.ASA,
                                       fill=ACENTO, outline=TINTA, width=1,
                                       tags=("asas", "asa", f"asa{i}"))

    def _pulsar(self, e):
        if self.actual is None or self.viendo:
            return
        puntos = [self._a_lona(*p) for p in self.fotos[self.actual]["esquinas"]]
        d, i = min(((px - e.x) ** 2 + (py - e.y) ** 2, k) for k, (px, py) in enumerate(puntos))
        self.asa = i if d <= 26 ** 2 else None

    def _arrastrando(self, e):
        if self.asa is not None:
            self._arrastrar(self.asa, e.x, e.y)

    def _arrastrar(self, i, cx, cy):
        """Lleva la esquina i al punto (cx, cy) del recuadro, sin salirse de la foto."""
        if self.actual is None or self.viendo or self._ocupado():
            return
        f = self.fotos[self.actual]
        x, y = self._a_original(cx, cy)
        f["esquinas"][i] = [min(max(x, 0), f["ancho"]), min(max(y, 0), f["alto"])]
        self._dibujar_asas()

    # --------------------------------------------------------------- controles
    def detectar(self):
        if self.actual is None or self._ocupado():
            return
        self.fotos[self.actual]["esquinas"] = [list(p) for p in self.fotos[self.actual]["auto"]]
        self._volver_a_editar()
        self.v_estado.set("Esquinas detectadas de nuevo.")

    def toda_la_hoja(self):
        if self.actual is None or self._ocupado():
            return
        f = self.fotos[self.actual]
        f["esquinas"] = [[0, 0], [f["ancho"], 0], [f["ancho"], f["alto"]], [0, f["alto"]]]
        self._volver_a_editar()
        self.v_estado.set("Se toma la foto entera, sin recortar.")

    def rotar(self):
        if self.actual is None or self._ocupado():
            return
        f = self.fotos[self.actual]
        f["giro"] = (f["giro"] + 90) % 360
        self.v_estado.set(f"Giro de esta foto: {f['giro']} grados (hacia la derecha). "
                          "Se ve en el resultado.")
        if self.viendo:
            self._procesar_vista()

    def _modo_cambiado(self):
        if self._ocupado():
            return
        if self.actual is not None:
            for f in (self.fotos if self.v_todas.get() else [self.fotos[self.actual]]):
                f["modo"] = self.v_modo.get()
            if self.viendo:
                self._procesar_vista()

    def _volver_a_editar(self):
        self.viendo = False
        self.btn_ver.config(text="[Ver resultado]")
        self._mostrar()

    def ver_resultado(self):
        if self.actual is None:
            messagebox.showinfo("Sin fotos", "Primero agregá una foto.", parent=self)
            return
        if self.viendo:
            self._volver_a_editar()
        elif not self._ocupado():
            self._procesar_vista()

    def _procesar_vista(self):
        """Procesa la foto actual con sus ajustes y la muestra en el recuadro."""
        f, idx = self.fotos[self.actual], self.actual
        self.contador += 1
        base = self.tmp / f"r{self.contador}"
        self.v_estado.set("Procesando la foto…")
        emit = lambda t, d: self.cola.put(("estado", d)) if t == "log" else None

        def trabajo():
            r = pdf_op({"op": "img_procesar", "ruta": str(f["ruta"]), "esquinas": f["esquinas"],
                        "rotacion": f["giro"], "modo": f["modo"], "salida": str(base) + ".jpg",
                        "vista": str(base) + ".png", "lado": self.LADO}, emit)
            self.cola.put(("vista", (idx, f["ruta"], r)))
        threading.Thread(target=trabajo, daemon=True).start()

    # --------------------------------------------------------------- crear PDF
    def crear_pdf(self):
        if self.creando:
            return
        if not self.fotos:
            messagebox.showinfo("Sin fotos", "Primero agregá al menos una foto.", parent=self)
            return
        destino = filedialog.asksaveasfilename(
            parent=self, title="Dónde guardar el PDF", defaultextension=".pdf",
            initialfile="Escaneo.pdf", filetypes=[("PDF", "*.pdf")],
            initialdir=str(self.app.base or self.app.cfg.get("carpeta") or Path.home()))
        if not destino:
            return
        destino = Path(destino)
        if destino.suffix.lower() != ".pdf":
            destino = destino.with_name(destino.name + ".pdf")
        datos = [{"ruta": f["ruta"], "nombre": f["nombre"], "esquinas": [list(p) for p in f["esquinas"]],
                  "giro": f["giro"], "modo": f["modo"]} for f in self.fotos]
        pagina = self.v_pagina.get()
        self.creando, self.cancelar = True, False
        self.btn_crear.config(state="disabled")
        threading.Thread(target=self._crear, args=(datos, destino, pagina), daemon=True).start()

    def _crear(self, datos, destino, pagina):
        emit = lambda t, d: self.cola.put(("estado", d)) if t == "log" else None
        n = len(datos)
        try:
            jpgs = []
            for i, d in enumerate(datos, 1):
                if self.cancelar:
                    return
                self.cola.put(("estado", f"Procesando foto {i} de {n}: {d['nombre']}…"))
                salida = self.tmp / f"p{i:03d}.jpg"
                r = pdf_op({"op": "img_procesar", "ruta": str(d["ruta"]), "esquinas": d["esquinas"],
                            "rotacion": d["giro"], "modo": d["modo"], "salida": str(salida)}, emit)
                if "error" in r:
                    self.cola.put(("error", (r["error"], d["nombre"])))
                    return
                jpgs.append(str(salida))
            self.cola.put(("estado", f"Armando el PDF de {n} página(s)…"))
            r = pdf_op({"op": "img_pdf", "imagenes": jpgs, "salida": str(destino),
                        "pagina": pagina}, emit)
            if "error" in r:
                self.cola.put(("error", (r["error"], None)))
                return
            self.cola.put(("listo", (destino, n)))
        except Exception as e:
            self.cola.put(("error", (str(e), None)))
        finally:
            self.cola.put(("libre", None))

    # ------------------------------------------------------------------- cola
    def _sondear(self):
        try:
            if not self.winfo_exists():
                return
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "estado":
                    self.v_estado.set(str(dato).strip()[:200])
                elif tipo == "fotos":
                    self._incorporar(dato)
                elif tipo == "vista":
                    idx, ruta, r = dato
                    if r.get("error") in ("no_imagenes", "no_pymupdf"):
                        messagebox.showwarning("Falta un componente", AYUDA_IMAGENES, parent=self)
                    elif "error" in r:
                        messagebox.showerror("No se pudo procesar la foto",
                                             f"{Path(ruta).name}: {r['error']}", parent=self)
                        self.v_estado.set("No se pudo procesar la foto.")
                    elif idx == self.actual:
                        self.resultado = {"vista": r["vista"], "vw": r["vista_ancho"],
                                          "vh": r["vista_alto"]}
                        self.viendo = True
                        self.btn_ver.config(text="[Volver a editar]")
                        self._mostrar()
                        self.v_estado.set("Así queda esta foto. Tocá [Volver a editar] para "
                                          "ajustar.")
                elif tipo == "error":
                    msg, nombre = dato
                    if msg in ("no_imagenes", "no_pymupdf"):
                        messagebox.showwarning("Falta un componente", AYUDA_IMAGENES, parent=self)
                    else:
                        messagebox.showerror(
                            "No se pudo crear el PDF",
                            (f"Falló la foto «{nombre}»:\n" if nombre else "") + msg, parent=self)
                    self.v_estado.set("No se pudo crear el PDF.")
                elif tipo == "listo":
                    destino, n = dato
                    self.v_estado.set(f"Listo: {destino.name}, {n} página(s). Está en {destino.parent}.")
                    self.app.decir_accion("escaner", n=n)
                    if self.app.v_abrir.get():
                        self.app._abrir(destino)
                elif tipo == "libre":
                    self.creando = False
                    self.btn_crear.config(state="normal")
        except queue.Empty:
            pass
        except tk.TclError:
            return
        self.after(100, self._sondear)

    # ----------------------------------------------------------------- cierre
    def _cerrar(self):
        if self.creando and not messagebox.askyesno(
                "Se está creando el PDF", "Todavía se está armando el PDF. ¿Cerrar igual y "
                "cancelarlo?", parent=self):
            return
        self.destroy()

    def _al_destruir(self, e):
        if e.widget is self:
            self.cancelar = True
            shutil.rmtree(self.tmp, ignore_errors=True)


class LectorMD(tk.Toplevel):
    """Lector de Markdown: abre un .md, lo muestra con formato, busca y lo exporta a PDF.
    No carga imágenes ni ejecuta nada del archivo: solo texto."""

    TAM_BASE, TAM_MIN, TAM_MAX = 11, 8, 26
    MAX_BYTES = 8 * 1024 * 1024
    EXT = (".md", ".markdown")

    def __init__(self, app, ruta=None):
        super().__init__(app)
        self.app = app
        self.title("Lector de Markdown")
        self.geometry("980x740")
        self.minsize(640, 460)
        self.configure(background=BG)
        barra_oscura(self)
        self.ruta = None
        self.texto = ""
        self.tam = self.TAM_BASE
        self.cola = queue.Queue()
        self.exportando = False
        self.hallazgos = []   # rangos (inicio, fin) de las coincidencias
        self.actual = -1
        self.ultimo_buscado = ""
        self.v_nombre = tk.StringVar(value="(ningún archivo)")
        self.v_buscar = tk.StringVar()
        self.v_estado = tk.StringVar(value="Abrí un archivo .md para leerlo.")
        self._armar()
        self._fuentes()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.after(100, self._sondear)
        if ruta:
            self.cargar(ruta)

    # ------------------------------------------------------------------ armado
    def _armar(self):
        cont = ttk.Frame(self, padding=(22, 14, 22, 12))
        cont.pack(fill="both", expand=True)
        ttk.Label(cont, text="LECTOR DE MARKDOWN", style="Title.TLabel").pack(anchor="w")
        pie = ttk.Frame(cont)
        pie.pack(side="bottom", fill="x", pady=(8, 0))
        ttk.Label(pie, textvariable=self.v_estado, style="Muted.TLabel", wraplength=860,
                  justify="left").pack(side="left", anchor="w")
        self.btn_pdf = ttk.Button(pie, text="[ EXPORTAR A PDF ]", style="Accent.TButton",
                                  command=self.exportar)
        self.btn_pdf.pack(side="right")
        arriba = ttk.Frame(cont)
        arriba.pack(fill="x", pady=(6, 8))
        ttk.Button(arriba, text="[ ABRIR… ]", style="Accent.TButton", command=self.abrir_dialogo
                   ).pack(side="left")
        ttk.Button(arriba, text="[Recargar]", style="Link.TButton", command=self.recargar
                   ).pack(side="left", padx=(12, 0))
        ttk.Label(arriba, textvariable=self.v_nombre, style="Muted.TLabel"
                  ).pack(side="left", padx=(14, 0))
        for texto, paso in (("[A+]", 1), ("[A-]", -1)):
            ttk.Button(arriba, text=texto, style="Link.TButton",
                       command=lambda p=paso: self.cambiar_tam(p)).pack(side="right", padx=(8, 0))
        ttk.Button(arriba, text="[Buscar]", style="Link.TButton", command=self.buscar
                   ).pack(side="right", padx=(8, 0))
        self.campo = ttk.Entry(arriba, textvariable=self.v_buscar, width=22)
        self.campo.pack(side="right")
        self.campo.bind("<Return>", lambda e: self.buscar())

        marco = tk.Frame(cont, background=LINEA, padx=1, pady=1)
        marco.pack(fill="both", expand=True)
        sy = ttk.Scrollbar(marco)
        sy.pack(side="right", fill="y")
        sx = ttk.Scrollbar(marco, orient="horizontal")
        sx.pack(side="bottom", fill="x")
        self.caja = tk.Text(marco, wrap="word", state="disabled", relief="flat", borderwidth=0,
                            highlightthickness=0, background=PANEL, foreground=FG,
                            insertbackground=ACENTO, selectbackground=ACENTO_OSC,
                            selectforeground=ACENTO, padx=26, pady=18, cursor="arrow",
                            yscrollcommand=sy.set, xscrollcommand=sx.set,
                            font=(self.app.fam, self.TAM_BASE))
        self.caja.pack(side="left", fill="both", expand=True)
        sy.config(command=self.caja.yview)
        sx.config(command=self.caja.xview)

    def _fuentes(self):
        """Crea (o reajusta) las fuentes y los tags según el tamaño actual."""
        import tkinter.font as tkfont
        esc = self.tam / self.TAM_BASE
        mono, tit = self.app.fam, self.app.fam_titulo
        if not hasattr(self, "f"):
            self.f = {}
        defs = {"base": (mono, self.tam, "normal", "roman"),
                "b": (mono, self.tam, "bold", "roman"),
                "i": (mono, self.tam, "normal", "italic"),
                "bi": (mono, self.tam, "bold", "italic"),
                "code": (mono, max(self.tam - 1, 7), "normal", "roman"),
                "h1": (tit, round(30 * esc), "normal", "roman"),
                "h2": (tit, round(24 * esc), "normal", "roman"),
                "h3": (tit, round(20 * esc), "normal", "roman")}
        for nombre, (fam, tam, peso, inclinacion) in defs.items():
            if nombre in self.f:
                self.f[nombre].configure(family=fam, size=tam, weight=peso, slant=inclinacion)
            else:
                self.f[nombre] = tkfont.Font(root=self, family=fam, size=tam, weight=peso,
                                             slant=inclinacion)
        t = self.caja
        t.configure(font=self.f["base"])
        t.tag_configure("h1", font=self.f["h1"], foreground=ACENTO, spacing1=10, spacing3=8)
        t.tag_configure("h2", font=self.f["h2"], foreground=ACENTO, spacing1=8, spacing3=6)
        t.tag_configure("h3", font=self.f["h3"], foreground=ACENTO, spacing1=6, spacing3=4)
        t.tag_configure("b", font=self.f["b"], foreground=ACENTO2)
        t.tag_configure("i", font=self.f["i"], foreground=GRIS)
        t.tag_configure("bi", font=self.f["bi"], foreground=ACENTO2)
        t.tag_configure("code", font=self.f["code"], foreground=GRIS, background=PANEL2)
        t.tag_configure("codeblock", font=self.f["code"], foreground=GRIS, background=PANEL2,
                        lmargin1=14, lmargin2=14, wrap="none")
        t.tag_configure("tabla", font=self.f["code"], foreground=FG, background=PANEL2,
                        lmargin1=14, lmargin2=14, wrap="none")
        t.tag_configure("quote", foreground=TENUE, lmargin1=28, lmargin2=28)
        t.tag_configure("li", spacing1=2)
        for nivel in range(7):
            t.tag_configure("nivel%d" % nivel, lmargin1=10 + 22 * nivel, lmargin2=30 + 22 * nivel)
        t.tag_configure("hr", foreground=LINEA, wrap="none")
        t.tag_configure("link", foreground=VERDE, underline=True)
        t.tag_configure("hallazgo", background=AMBAR, foreground=TINTA)
        t.tag_configure("actual", background=ACENTO, foreground=TINTA)
        t.tag_raise("hallazgo")
        t.tag_raise("actual")
        t.tag_bind("link", "<Enter>", lambda e: t.configure(cursor="hand2"))
        t.tag_bind("link", "<Leave>", lambda e: t.configure(cursor="arrow"))
        t.tag_bind("link", "<Button-1>", self._clic)

    # ------------------------------------------------------------------- cargar
    def abrir_dialogo(self):
        ruta = filedialog.askopenfilename(
            parent=self, title="Elegí un archivo Markdown",
            initialdir=str(self.ruta.parent if self.ruta else
                           (self.app.base or self.app.cfg.get("carpeta") or Path.home())),
            filetypes=[("Markdown", "*.md *.markdown"), ("Todos los archivos", "*.*")])
        if ruta:
            self.cargar(ruta)

    def recargar(self):
        if self.ruta:
            self.cargar(self.ruta, avisar=False)
        else:
            self.v_estado.set("Todavía no hay ningún archivo abierto.")

    def _leer(self, ruta):
        """(texto, codificación). UTF-8 y, si no sirve, cp1252 (Word y el Bloc de notas viejo)."""
        crudo = Path(ruta).read_bytes()
        if len(crudo) > self.MAX_BYTES:
            raise ValueError("el archivo pesa más de 8 MB: es demasiado para el lector")
        try:
            return crudo.decode("utf-8-sig"), "UTF-8"
        except UnicodeDecodeError:
            return crudo.decode("cp1252", errors="replace"), "cp1252 (Windows)"

    def cargar(self, ruta, avisar=True):
        ruta = Path(ruta)
        try:
            texto, cod = self._leer(ruta)
        except (OSError, ValueError) as e:
            messagebox.showerror("No se pudo abrir el archivo", f"{ruta.name}: {e}", parent=self)
            self.v_estado.set("No se pudo abrir el archivo.")
            return False
        self.ruta, self.texto = ruta, texto
        self.v_nombre.set(ruta.name)
        self.title(f"Lector de Markdown — {ruta.name}")
        self._pintar()
        palabras, lineas = contar_palabras(texto), len(texto.splitlines())
        self.v_estado.set(f"Leído como {cod} · {palabras} palabra(s) · {lineas} línea(s)")
        if avisar:
            self.app.decir_accion("markdown", nombre=ruta.name, n=palabras)
        return True

    def _pintar(self):
        t = self.caja
        t.configure(state="normal")
        t.delete("1.0", "end")
        self.hallazgos, self.actual, self.ultimo_buscado = [], -1, ""
        for tag in t.tag_names():
            if tag.startswith("url:"):
                t.tag_delete(tag)
        for texto, tags in markdown_a_bloques(self.texto):
            t.insert("end", texto, tags)
        t.configure(state="disabled")
        t.yview_moveto(0)

    # ------------------------------------------------------------------- enlaces
    def _clic(self, evento):
        t = self.caja
        for tag in t.tag_names(f"@{evento.x},{evento.y}"):
            if tag.startswith("url:"):
                self.abrir_enlace(tag[4:])
                return "break"

    def abrir_enlace(self, url):
        if url.lower().startswith(("http://", "https://")):
            webbrowser.open(url)
            self.v_estado.set(f"Se abrió el enlace en el navegador: {url[:80]}")
            return
        from urllib.parse import unquote
        rel = unquote(url.split("#")[0])
        destino = (self.ruta.parent / rel) if self.ruta else Path(rel)
        if destino.suffix.lower() not in self.EXT + (".txt",):
            self.v_estado.set("Ese enlace no es otro Markdown: no lo abro desde acá.")
        elif not destino.is_file():
            self.v_estado.set(f"No encuentro el archivo enlazado: {rel}")
        else:
            self.cargar(destino)

    # -------------------------------------------------------------------- buscar
    def buscar(self):
        t = self.caja
        q = self.v_buscar.get()
        if not q:
            self._limpiar_busqueda()
            self.v_estado.set("Escribí algo para buscar.")
            return
        if q != self.ultimo_buscado:
            self._limpiar_busqueda()
            self.ultimo_buscado = q
            pos = "1.0"
            while True:
                ini = t.search(q, pos, stopindex="end", nocase=True)
                if not ini:
                    break
                fin = f"{ini}+{len(q)}c"
                t.tag_add("hallazgo", ini, fin)
                self.hallazgos.append((ini, fin))
                pos = fin
            self.actual = -1
        if not self.hallazgos:
            self.v_estado.set(f"Sin coincidencias de «{q}».")
            return
        if self.actual >= 0:
            t.tag_remove("actual", "1.0", "end")
        self.actual = (self.actual + 1) % len(self.hallazgos)
        ini, fin = self.hallazgos[self.actual]
        t.tag_add("actual", ini, fin)
        t.see(ini)
        self.v_estado.set(f"Coincidencia {self.actual + 1} de {len(self.hallazgos)} para «{q}».")

    def _limpiar_busqueda(self):
        self.caja.tag_remove("hallazgo", "1.0", "end")
        self.caja.tag_remove("actual", "1.0", "end")
        self.hallazgos, self.actual, self.ultimo_buscado = [], -1, ""

    # ------------------------------------------------------------------- tamaño
    def cambiar_tam(self, paso):
        nuevo = min(self.TAM_MAX, max(self.TAM_MIN, self.tam + paso))
        if nuevo != self.tam:
            self.tam = nuevo
            self._fuentes()
        self.v_estado.set(f"Tamaño de letra: {self.tam}.")

    # --------------------------------------------------------------- exportar PDF
    def exportar(self):
        if self.exportando:
            return
        if not self.ruta:
            messagebox.showinfo("Sin archivo", "Primero abrí un Markdown.", parent=self)
            return
        destino = filedialog.asksaveasfilename(
            parent=self, title="Dónde guardar el PDF", defaultextension=".pdf",
            initialdir=str(self.ruta.parent), initialfile=self.ruta.stem + ".pdf",
            filetypes=[("PDF", "*.pdf")])
        if not destino:
            return
        destino = Path(destino)
        if destino.suffix.lower() != ".pdf":
            destino = destino.with_name(destino.name + ".pdf")
        html = markdown_a_html(self.texto, self.ruta.name)
        self.exportando = True
        self.btn_pdf.config(state="disabled")
        self.v_estado.set("Armando el PDF…")
        emit = lambda t, d: self.cola.put(("estado", d)) if t == "log" else None

        def trabajo():
            try:
                r = pdf_op({"op": "md_pdf", "html": html, "salida": str(destino)}, emit)
            except Exception as e:
                r = {"error": str(e)}
            self.cola.put(("listo", (destino, r)))
        threading.Thread(target=trabajo, daemon=True).start()

    def _sondear(self):
        try:
            if not self.winfo_exists():
                return
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "estado":
                    self.v_estado.set(str(dato).strip()[:200])
                elif tipo == "listo":
                    destino, r = dato
                    self.exportando = False
                    self.btn_pdf.config(state="normal")
                    if r.get("error") == "no_pymupdf":
                        messagebox.showwarning("Falta un componente", AYUDA_PYMUPDF, parent=self)
                        self.v_estado.set("No se pudo crear el PDF.")
                    elif "error" in r:
                        messagebox.showerror("No se pudo crear el PDF", r["error"], parent=self)
                        self.v_estado.set("No se pudo crear el PDF.")
                    else:
                        self.v_estado.set(f"Listo: {destino.name}, {r.get('paginas', '?')} "
                                          f"página(s). Está en {destino.parent}.")
                        if self.app.v_abrir.get():
                            self.app._abrir(destino)
        except queue.Empty:
            pass
        except tk.TclError:
            return
        self.after(100, self._sondear)


# ------------------------------------------------------------------------- GUI
# Pantalla de fósforo verde. Verde fuerte (y entre corchetes) = lo que se toca; verde suave =
# información; verde pálido = lo importante; ámbar = advertencia; rojo = error.
# Contraste de texto sobre el fondo: 5,4:1 o más.
BG, PANEL, PANEL2, LINEA = "#030704", "#081409", "#0d2010", "#1d4a22"
FG, GRIS, TENUE = "#8ee67f", "#5fae55", "#4d9446"
ACENTO, ACENTO2, ACENTO_OSC = "#39ff14", "#d2ffc8", "#0f3d12"
TINTA, ROJO, VERDE, AMBAR = "#030704", "#ff5a4a", "#d2ffc8", "#ffc43d"
BRILLO = "#12400f"   # halo alrededor del texto de fósforo
BARRIDO = "#07120a"  # líneas de barrido de la pantalla

EFECTOS = ("completos", "reducidos", "apagados")
GLIFOS = "0123456789ABCDEFHKMNXZ$#%&*+=<>:;?¥§¤ÆØßÞðµ¶"  # todos existen en VT323

FUENTE_TITULO, FUENTE_TEXTO = "VT323", "Share Tech Mono"


def _carpeta_fuentes() -> Path:
    base = os.environ.get("LOCALAPPDATA") if ES_WINDOWS else None
    if base:
        return Path(base) / "PDF Master" / "fuentes"
    if platform.system() == "Linux":
        return Path.home() / ".local" / "share" / "fonts" / "pdf-master"
    return Path(tempfile.gettempdir()) / "pdf-master-fuentes"


_FUENTES_LISTAS = []


def registrar_fuentes():
    """Deja disponibles las fuentes incluidas, solo para este programa (en Windows no se
    instalan en el sistema). Una sola vez por proceso. Si algo falla, la interfaz usa
    fuentes del sistema."""
    if _FUENTES_LISTAS:
        return
    _FUENTES_LISTAS.append(True)
    try:
        carpeta = _carpeta_fuentes()
        carpeta.mkdir(parents=True, exist_ok=True)
        nuevas = False
        for nombre, datos in FUENTES_B64.items():
            crudo = base64.b64decode("".join(datos.split()))
            destino = carpeta / nombre
            if not destino.exists() or destino.stat().st_size != len(crudo):
                destino.write_bytes(crudo)
                nuevas = True
            if ES_WINDOWS:
                import ctypes
                ctypes.windll.gdi32.AddFontResourceExW(str(destino), 0x10, 0)  # FR_PRIVATE
        if nuevas and platform.system() == "Linux" and shutil.which("fc-cache"):
            subprocess.run(["fc-cache", "-f", str(carpeta)], capture_output=True, timeout=60)
    except Exception:
        pass


def elegir_fuentes(raiz):
    """(fuente de títulos, fuente de texto), con reserva del sistema si no están."""
    import tkinter.font as tkfont
    try:
        familias = set(tkfont.families(raiz))
    except tk.TclError:
        familias = set()
    reserva = next((f for f in ("Consolas", "DejaVu Sans Mono", "Menlo", "Courier New")
                    if f in familias), "TkFixedFont")
    return (FUENTE_TITULO if FUENTE_TITULO in familias else reserva,
            FUENTE_TEXTO if FUENTE_TEXTO in familias else reserva)


def barra_oscura(ventana):
    """Barra de título oscura en Windows 10/11 (si no se puede, queda la normal)."""
    if not ES_WINDOWS:
        return
    try:
        import ctypes
        ventana.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(ventana.winfo_id())
        valor = ctypes.c_int(1)
        for atributo in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (nuevo y viejo)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, atributo, ctypes.byref(valor), ctypes.sizeof(valor)) == 0:
                break
    except Exception:
        pass


def _tipo(p: Path):
    e = p.suffix.lower()
    return "PDF" if e == ".pdf" else ("PowerPoint" if e in (".pptx", ".ppt") else "Word")


def modo_efectos(cfg):
    """completos (por defecto, siempre) · reducidos (sin movimiento) · apagados (modo serio).
    Decisión del 2026-11-25: arranca en completos aunque Windows tenga las animaciones
    apagadas; se cambia desde MÁS… » Efectos visuales."""
    e = cfg.get("efectos")
    if e is False:
        return "apagados"
    if e in EFECTOS:
        return e
    return "completos"


def espacio_libre(ruta):
    """Bytes libres en el disco de 'ruta' (o de su carpeta existente más cercana)."""
    try:
        p = Path(ruta)
        while not p.exists() and p != p.parent:
            p = p.parent
        return shutil.disk_usage(p).free
    except Exception:
        return None


def en_gb(n):
    return f"{n / 1024 ** 3:.1f}".replace(".", ",") + " GB"


def barrido(canvas, ancho, alto, paso=3):
    """Líneas de barrido de CRT, por debajo de todo lo demás."""
    for y in range(0, int(alto), paso):
        canvas.create_line(0, y, ancho, y, fill=BARRIDO, tags="barrido")
    canvas.tag_lower("barrido")


def texto_fosforo(canvas, x, y, texto, fuente, color=ACENTO, tags=(), **kw):
    """Texto con halo de fósforo: copias oscuras corridas un píxel y el texto encima."""
    tags = tuple(tags) if isinstance(tags, (tuple, list)) else (tags,)
    for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1), (1, 1)):
        canvas.create_text(x + dx, y + dy, text=texto, font=fuente, fill=BRILLO,
                           tags=tags + ("halo",), **kw)
    return canvas.create_text(x, y, text=texto, font=fuente, fill=color, tags=tags, **kw)


class Guion:
    """La sátira: una dependencia pública informática que se toma demasiado en serio.
    Nada de esto informa. Lo verdadero (cantidades, rutas, tiempos) siempre lo pone el
    programa y se muestra aparte o entre los textos, nunca inventado."""

    ARRANQUE = (
        ("PRUEBA DE MEMORIA", "655360 BYTES. ALCANZAN"),
        ("MESA DE ENTRADAS", "ABIERTA DE 8 A 13"),
        ("SELLOS Y FIRMAS", "EN TRÁMITE"),
        ("PROTOCOLOS HEREDADOS", "17 CARGADOS, 3 SIN DOCUMENTAR"),
        ("RELOJ DE PARED", "SINCRONIZADO A OJO"),
        ("CONDENSADOR DE PÁGINAS", "CARGADO, POR AHORA"),
        ("SUPERVISOR DE TURNO", "EN EL ALMUERZO"),
        ("MANUAL DE PROCEDIMIENTOS", "TOMO II EXTRAVIADO"),
        ("UNIDAD DE CINTA", "REBOBINANDO POR LAS DUDAS"),
        ("NIVEL DE SOLEMNIDAD", "MÁXIMO"),
        ("PERMISO DEL ARCHIVO GENERAL", "SE ASUME CONCEDIDO"),
        ("CAFETERA DEL CUARTO PISO", "NO RESPONDE"),
        ("ENTRELAZAMIENTO CUÁNTICO", "NO CORRESPONDE"),
    )
    REPOSO = (
        "LA COMPUTADORA ESTÁ TRABAJANDO. POR FAVOR, NO LA ALIENTE.",
        "PROCESO 0007: PENSANDO EN ALGO. NO SE INFORMARÁ EN QUÉ.",
        "SE DETECTÓ UN OPERADOR HUMANO. SE LO TRATARÁ CON CORDIALIDAD PROTOCOLAR.",
        "NIVEL DE MATE POR DEBAJO DEL UMBRAL RECOMENDADO. SE NOTIFICÓ A RECURSOS HUMANOS.",
        "RECORDATORIO: TODO DOCUMENTO ES CONVERTIBLE HASTA QUE SE DEMUESTRE LO CONTRARIO.",
        "INVENTARIO DE CLIPS: SIN NOVEDAD.",
        "EL SISTEMA AGRADECE SU PACIENCIA. TAMBIÉN LA REGISTRA.",
        "SE REORDENARON LOS ARCHIVADORES. NADIE LO PIDIÓ.",
        "PRÓXIMO MANTENIMIENTO PROGRAMADO: CUANDO HAYA PRESUPUESTO.",
        "LA IMPRESORA DEL PASILLO PREGUNTÓ POR USTED. NO SE LE DIO INFORMACIÓN.",
        "VERIFICACIÓN DE RUTINA: LA REALIDAD LOCAL SIGUE EN SU LUGAR.",
        "SE ARCHIVÓ UN FORMULARIO 0347-B. NO SE SABE DE QUIÉN.",
        "INTEGRIDAD DEL SISTEMA: SATISFACTORIA. EL RESTO ESTÁ CLASIFICADO.",
        "AVISO: EL PASILLO B PERMANECERÁ CERRADO HASTA NUEVO AVISO. NO HAY PASILLO B.",
        "RECORDATORIO: EL INTRUSO DE 1994 SIGUE ADENTRO. SE LO CONSIDERA PERSONAL DE PLANTA.",
        "LA TERMINAL FUNCIONA CORRECTAMENTE. ESTO SE INFORMA PORQUE ES INUSUAL.",
        "SE DETECTARON 0 ANOMALÍAS. SE ABRIÓ UNA INVESTIGACIÓN POR LO SOSPECHOSO DEL DATO.",
        "EL TÉCNICO DEL SISTEMA SE JUBILÓ EN 1998. NADIE SABE CÓMO APAGARLO.",
        "ESTADO DEL SISTEMA: AL BORDE DEL COLAPSO, COMO DE COSTUMBRE.",
    )
    ATAQUES = (  # (alarma en rojo, resolución): siempre se resuelve sola y con chiste
        ("ALERTA: INTRUSO DETECTADO EN EL SECTOR 4B. ES UNA POLILLA.",
         "SE LA INVITÓ A RETIRARSE. ACEPTÓ."),
        ("ERROR CRÍTICO 0x0BADCAFE: SE AGOTÓ EL CAFÉ.", "SE REPUSO CON MATE. SISTEMA ESTABLE."),
        ("UN VIRUS INTENTA CONVERTIR SUS PDF EN WORD.",
         "BLOQUEADO. ACÁ SE CONVIERTE EN UN SOLO SENTIDO."),
        ("FALLA EN LA MATRIZ DE CONDENSADORES. PROBABILIDAD DE COLAPSO: ALTA.",
         "FALSA ALARMA: ERA EL VENTILADOR."),
        ("ADVERTENCIA: EL SISTEMA ESTÁ A PUNTO DE—", "…NO. SIGUE ANDANDO. NADIE SABE POR QUÉ."),
        ("SE DETECTÓ UN HACKER EN LA RED. PIDE QUE LE CONVIERTAN UN TRABAJO PRÁCTICO.",
         "SE LE ASIGNÓ TURNO PARA MARZO."),
        ("MEMORIA CORRUPTA EN EL BLOQUE 0313.", "SE RECOMPUSO CON CINTA SCOTCH."),
        ("¡¡RUPTURA DEL NÚCLEO INMINENTE!!",
         "ERA UN SIMULACRO. NO SE AVISÓ PORQUE ERA UN SIMULACRO."),
        ("ALGUIEN PATEÓ EL CABLE DEL SERVIDOR.", "SE VOLVIÓ A ENCHUFAR. NO SE INICIARÁ SUMARIO."),
        ("ATAQUE DE DENEGACIÓN DE SERVICIO: TRES PALOMAS EN LA ANTENA.",
         "PALOMAS DISUADIDAS. SERVICIO NORMALIZADO."),
        ("ERROR 404: NO SE ENCONTRÓ LA VOLUNTAD DE TRABAJAR.", "APARECIÓ DETRÁS DEL ARCHIVADOR."),
        ("SOBRECARGA TÉRMICA EN EL PROCESADOR CENTRAL.", "SE ABRIÓ LA VENTANA. TEMPERATURA NORMAL."),
        ("EL INTRUSO DE 1994 VOLVIÓ A INTENTAR ENTRAR.",
         "SE LE RECORDÓ QUE YA ESTÁ ADENTRO. SE FUE CONFUNDIDO."),
        ("PÉRDIDA DE SEÑAL CON EL SATÉLITE.", "NUNCA HUBO SATÉLITE. SE NORMALIZA LA SITUACIÓN."),
        ("¡¡AUTODESTRUCCIÓN EN 10… 9… 8…", "…SE CANCELÓ POR FALTA DE PRESUPUESTO."),
    )
    FALLAS_VIDEO = (
        ("SEÑAL DE VIDEO INESTABLE. NO TOQUE NADA.", "SE GOLPEÓ EL MONITOR. SEÑAL RECUPERADA."),
        ("PÉRDIDA DE SINCRONÍA VERTICAL.", "SE GIRÓ LA PERILLA DE ATRÁS. QUEDÓ MEJOR QUE ANTES."),
        ("EL TUBO DE RAYOS CATÓDICOS HIZO UN RUIDO PREOCUPANTE.", "ERA EL ESTÓMAGO DEL OPERADOR."),
        ("FALLA DE ALIMENTACIÓN EN EL MONITOR.", "SE MOVIÓ EL CABLE. NO LO VUELVA A TOCAR."),
        ("INTERFERENCIA ELECTROMAGNÉTICA DETECTADA.", "ALGUIEN PRENDIÓ EL MICROONDAS. YA TERMINÓ."),
        ("FÓSFORO DE LA PANTALLA AGOTÁNDOSE.", "SE PIDIÓ MÁS FÓSFORO. LLEGA EN MARZO."),
    )
    FALLOS_TRABAJO = (  # líneas rojas en la pantalla de trabajo; se recuperan en el acto
        "!! FALLO DE SEGMENTACIÓN EMOCIONAL — RECUPERADO",
        "!! EL CONDENSADOR HIZO UN RUIDO RARO — IGNORADO",
        "!! INTRUSO EN EL BUS DE DATOS — ERA UNA PELUSA",
        "!! TEMPERATURA DEL NÚCLEO: PREOCUPANTE — SE ABANICÓ",
        "!! VIRUS DETECTADO — SE LE PIDIÓ EL DNI Y SE FUE",
        "!! PÁGINA 7 INTENTÓ ESCAPAR — RECAPTURADA",
    )
    TITULOS = (
        "VERIFICANDO CREDENCIALES DEL OPERADOR…",
        "SOLICITANDO AUTORIZACIÓN DE NIVEL 9…",
        "CONSULTANDO ARCHIVOS QUE NO DEBERÍAN EXISTIR…",
        "NEGOCIANDO CON LOS MÁRGENES DE WORD…",
        "DESENCRIPTANDO LA VOLUNTAD DEL AUTOR…",
        "PERSUADIENDO A POWERPOINT DE QUE SUELTE LAS DIAPOSITIVAS…",
        "CALIBRANDO EL CONDENSADOR DE PÁGINAS…",
        "CONTANDO PÁGINAS IMPARES CON LOS DEDOS…",
        "ALINEANDO BITS CON LA LEY DE MURPHY…",
        "ESPERANDO EL SELLO DEL SUPERVISOR…",
        "ELEVANDO EL EXPEDIENTE A LA INSTANCIA SUPERIOR…",
        "INYECTANDO CAFÉ EN EL NÚCLEO DEL SISTEMA…",
    )
    SUBPROCESOS = (
        "CONTRASTANDO REGISTROS SECUNDARIOS",
        "CONSULTANDO LA HORA LOCAL",
        "CONSULTANDO OTRA VEZ LA HORA LOCAL, POR LAS DUDAS",
        "CONFIRMANDO QUE LA SESIÓN ACTUAL EXISTE",
        "EJECUTANDO SUMA DE COMPROBACIÓN REDUNDANTE",
        "AUTORIZACIÓN RECIBIDA DE ORIGEN DESCONOCIDO",
        "IGNORANDO ORIGEN DESCONOCIDO",
        "FOTOCOPIANDO EL EXPEDIENTE POR TRIPLICADO",
        "VERIFICANDO LA INTEGRIDAD DE LA REALIDAD LOCAL",
        "RECALIBRANDO EXPECTATIVAS",
        "SE DETECTÓ UN NIVEL INESPERADO DE COMPETENCIA",
        "FINGIENDO QUE TODO ESTO ES NECESARIO",
        "ESPERANDO QUE ATIENDAN EL INTERNO 214",
        "SOLICITANDO SEGUNDA FIRMA",
        "CONSULTANDO BASES DE DATOS QUE NO FIGURAN EN EL ORGANIGRAMA",
        "REDACTANDO MEMORÁNDUM INTERNO SOBRE EL MEMORÁNDUM ANTERIOR",
    )
    CIERRES = (
        "SE ARCHIVA EL EXPEDIENTE. VUELVA PRONTO.",
        "TRÁMITE CONCLUIDO SIN NECESIDAD DE SELLADO.",
        "LOS 47 PROTOCOLOS RESULTARON INNECESARIOS. SE CONSERVAN IGUAL.",
        "NINGÚN SISTEMA FUE HACKEADO EN EL PROCESO.",
        "FIRMA Y ACLARACIÓN: EL SISTEMA.",
    )
    # acción -> (texto llano o None, variantes satíricas). {n}, {nombre}, {exp}, {seg}: datos reales.
    ACCIONES = {
        "agregar": ("{n} archivo(s) en la lista.", (
            "MESA DE ENTRADAS: {n} EXPEDIENTE(S) RECIBIDO(S). SE LABRÓ ACTA.",
            "SE INGRESARON {n} EXPEDIENTE(S). QUEDAN EN EL ESCRITORIO DEL FONDO.",
            "{n} DOCUMENTO(S) INCORPORADO(S) AL PADRÓN. SE LES ASIGNÓ NÚMERO DE ORDEN.")),
        "carpeta": ("Carpeta {nombre}: {n} archivo(s) para convertir o unir.", (
            "SE INSPECCIONÓ LA CARPETA «{nombre}». {n} EXPEDIENTE(S) ELEGIBLE(S).",
            "RELEVAMIENTO DE «{nombre}» FINALIZADO: {n} EXPEDIENTE(S). SIN IRREGULARIDADES.")),
        "quitar": ("{n} quitado(s) de la lista. Los archivos siguen en su carpeta.", (
            "SE DESAFECTARON {n} EXPEDIENTE(S). LOS ORIGINALES NO FUERON TOCADOS.",)),
        "vaciar": ("Lista vacía. Tus archivos siguen en su lugar.", (
            "ESCRITORIO DESPEJADO. NADIE DESTRUYÓ NADA: SOLO SE DEJÓ DE MIRARLO.",)),
        "ordenar": (None, (
            "SE ORDENARON {n} EXPEDIENTE(S) ALFABÉTICAMENTE. EL ARCHIVO GENERAL APRUEBA.",)),
        "mover": (None, (
            "SE ALTERÓ EL ORDEN DE PRELACIÓN. QUEDA ASENTADO.",
            "REUBICACIÓN DE EXPEDIENTE AUTORIZADA DE PALABRA.")),
        "destino": ("Los PDF se van a guardar {nombre}.", (
            "NUEVO DESTINO DE ARCHIVO: {nombre}. SE NOTIFICÓ A LAS PARTES.",)),
        "unir_si": (None, ("SE AUTORIZÓ LA UNIFICACIÓN DE EXPEDIENTES.",)),
        "unir_no": (None, ("SE REVOCÓ LA UNIFICACIÓN. CADA EXPEDIENTE SIGUE POR SU CUENTA.",)),
        "inicio": (None, (
            "SE INICIÓ EL TRÁMITE {exp} CON {n} EXPEDIENTE(S).",
            "TRÁMITE {exp} EN CURSO: {n} EXPEDIENTE(S). NO APAGUE LA TERMINAL.")),
        "fin": ("Listo en {seg} s.", ("TRÁMITE {exp} CONCLUIDO EN {seg} S. {cierre}",)),
        "fin_error": ("Terminó con errores: mirá el registro.", (
            "TRÁMITE {exp} CONCLUIDO CON OBSERVACIONES. VER REGISTRO.",)),
        "listo": ("Listo para recibir archivos.", (
            "TERMINAL OPERATIVA. ESPERANDO EXPEDIENTES DEL OPERADOR.",
            "MESA DE ENTRADAS ABIERTA. TOME NÚMERO Y AGUARDE SER LLAMADO.")),
        "escaner": ("Listo: se escanearon {n} foto(s) y se armó el PDF.", (
            "SE ESCANEARON {n} FOTO(S). EL DEPARTAMENTO DE FOTOCOPIAS SE SIENTE AMENAZADO.",
            "SE DIGITALIZARON {n} FOTO(S). LOS ORIGINALES NO FUERON TOCADOS.")),
        "markdown": ("Se abrió «{nombre}»: {n} palabra(s).", (
            "SE ABRIÓ «{nombre}»: {n} PALABRAS. EL DEPARTAMENTO DE LECTURA RÁPIDA YA SE CANSÓ.",
            "EXPEDIENTE «{nombre}» EN LECTURA: {n} PALABRAS CONTADAS UNA POR UNA.",
            "SE LEYÓ «{nombre}» ({n} PALABRAS). NO SE ENCONTRÓ LA MORALEJA.")),
        "curioso": (None, (
            "SE DETECTÓ CURIOSIDAD NO AUTORIZADA. SE INICIÓ UN EXPEDIENTE. RESPUESTA "
            "ESTIMADA: 30 A 45 DÍAS HÁBILES.",)),
    }

    def __init__(self, azar):
        self.azar, self.mazos = azar, {}

    def sacar(self, nombre):
        """Uno al azar, sin repetir hasta agotar la lista."""
        mazo = self.mazos.get(nombre)
        if not mazo:
            mazo = list(getattr(self, nombre))
            self.azar.shuffle(mazo)
            self.mazos[nombre] = mazo
        return mazo.pop()

    def accion(self, clave, **datos):
        """(texto llano o None, texto satírico) con los datos reales ya puestos."""
        llano, variantes = self.ACCIONES[clave]
        datos.setdefault("cierre", self.sacar("CIERRES") if clave == "fin" else "")
        sat = self.azar.choice(variantes).format(**datos)
        return (llano.format(**datos) if llano else None), sat


class GuionDoctor(Guion):
    """El guion del Dr. Fósforo: un personaje local (no es una IA real ni usa internet).
    Solo dice cosas ciertas sobre el programa: las cantidades, tiempos y nombres que aparecen
    entre llaves los pone el programa con datos reales; el resto es fantasía propia del
    doctor (sus experimentos, su vida en el tubo) y nunca habla de lo que hay dentro de los
    archivos del operador. Las frases de la línea DR.FÓSFORO> son cortas (una línea)."""

    SALUDO_MADRUGADA = (
        "¡Buenas! Son las tantas y yo acá, brillando. Vos también despierto: qué casualidad.",
        "Madrugada. Los electrones duermen. Yo no: soy fósforo, no tengo remedio.",
        "¿Trabajando de madrugada? Me gusta tu estilo. Yo apago las luces con la mente.",
        "Ey, hola. A esta hora el fósforo brilla más lindo. Es ciencia. Más o menos.",
    )
    SALUDO_MANANA = (
        "¡Buen día! Desayuné electrones. Ricos, pero fríos.",
        "Buenos días, operador. Calenté el fósforo desde temprano. Arrimate.",
        "¡Buen día! Hoy tengo un presentimiento científico. Capaz es hambre.",
        "Mañana linda para convertir cosas. Y para experimentos, pero esos van aparte.",
    )
    SALUDO_TARDE = (
        "¡Buenas tardes! Acabo de terminar un experimento. Salió raro. Estoy contento.",
        "Hola, hola. A la tarde el fósforo rinde más. No tengo cómo demostrarlo.",
        "¡Operador! Justo te esperaba. Bueno, esperar no: giraba en círculos.",
        "Buenas tardes. Me tomé un recreo de tres segundos. Fue demasiado.",
    )
    SALUDO_NOCHE = (
        "Buenas noches. Es la hora en que el fósforo se luce. Mirame brillar.",
        "¡Qué bueno que viniste! De noche me da por la ciencia profunda.",
        "Noche cerrada y la pantalla encendida. Somos tal para cual.",
        "Buenas noches, operador. Prometo experimentar en voz baja.",
    )
    SALUDO_DIA = (  # lunes … domingo (lo dice el calendario real)
        "Lunes. Dicen que es el peor día. Yo lo encuentro fascinante: nadie sabe por qué.",
        "Martes. Nunca pasa nada un martes. Por eso hago experimentos.",
        "Miércoles: mitad de semana. Para mí, mitad de experimento.",
        "Jueves. Casi viernes. El fósforo ya lo siente en los electrones.",
        "¡Viernes! Hoy se convierte con alegría. Yo ya bailo, adentro del tubo.",
        "Sábado. ¿Trabajando? Te admiro. Yo no tengo fin de semana: vivo en el tubo.",
        "Domingo. Día de descanso. Yo descanso girando despacito.",
    )
    PIN_MAL = (
        "Ese PIN estaba mal. No importa: acá cualquier respuesta entra. Ya lo sabías.",
        "PIN incorrecto. Tranquilo, nadie se enoja. Yo ya vi cosas peores.",
        "Te equivocaste de PIN y entraste igual. Esa es mi clase de seguridad.",
    )
    AGREGAR = (  # {n}
        "¡{n} archivo(s) nuevos! Los miro con respeto científico.",
        "Anoté {n} archivo(s). Con lápiz, por si me arrepiento.",
        "{n} archivo(s) más. Esto se pone interesante, operador.",
        "Me llegaron {n} archivo(s). Los recibí con aplausos de fósforo.",
        "¡{n}! Hoy se trabaja en serio. Bueno, en serio-serio, no.",
    )
    CONVERTIR = (  # {n}
        "Arrancamos con {n} archivo(s). Si algo explota, fui yo. Mentira: no explota nada.",
        "¡Dale! {n} archivo(s) en la cinta. Yo miro, que es lo mío.",
        "Ahí van {n} archivo(s). Crucemos los cables.",
        "Procesando {n}. Yo no toco nada. Casi nada. Nada.",
    )
    FIN = (  # {seg}
        "¡Listo! {seg} segundos. Los medí con mi cronómetro interno, que anda más o menos.",
        "Terminamos en {seg} s. Ni un experimento nuevo en el medio. Bueno, uno.",
        "{seg} segundos. Mi récord personal no lo sé: nunca lo medí.",
        "¡Chan! {seg} s. Te lo digo con orgullo de tubo catódico.",
        "Terminó en {seg} s. Te felicito. Y me felicito. Y felicito al tubo.",
    )
    FIN_HACKEO = (  # {seg}: el operador pasó por el modo intrusión
        "Terminó en {seg} s. Te vi en modo intrusión: hackeaste con mucho estilo.",
        "{seg} segundos. La pantalla de intrusión me dio envidia: yo también quiero código.",
    )
    ERROR = (  # el error es real: se lo dice sin disfrazarlo
        "Ay, falló algo. Esto no fue un experimento mío: es real. Mirá el registro.",
        "Hubo un error de verdad. Lo siento, operador. Abrí el registro y lo vemos juntos.",
        "No salió. Y esta vez no es chiste: hubo errores. El registro tiene los detalles.",
        "Uf, algo falló. No es sátira. Leé el registro, que ahí dice qué pasó.",
        "Hubo errores. Es en serio. Tranquilo: lo leemos juntos en el registro.",
    )
    ESCANER_ABRE = (
        "¡El escáner! Me encanta. Siempre quise tener ojos de cámara.",
        "Fotos a PDF. Eso sí lo entiendo: es luz convertida en cosas.",
        "Escáner abierto. Si me ves reflejado en el papel, no digas nada.",
    )
    ESCANER = (  # {n}
        "{n} foto(s) convertidas a PDF. La luz hizo todo el trabajo, yo solo animé.",
        "Se escanearon {n} foto(s). Clic, clic y a otra cosa.",
        "{n} foto(s) y un PDF. Fotones a la orden.",
    )
    MARKDOWN = (  # {nombre}, {n}
        "Abriste «{nombre}»: {n} palabra(s). Las leí por encima. Con respeto.",
        "«{nombre}» tiene {n} palabra(s). Cada una con su cuota de fósforo.",
        "{n} palabra(s) en «{nombre}». Yo en 1987 escribía menos.",
    )
    OCULTO = (
        "¡Me encontraste! Soy yo: estaba escondido detrás del título.",
        "Cinco clics en el título. Sos de los míos: curioso. Seguí así.",
        "Ah, un curioso. Yo soy otro. Así empiezan las grandes amistades.",
    )
    EFECTOS_COMPLETOS = (
        "Efectos completos. Ahora sí me puedo mover como corresponde.",
        "Todo encendido. Me crujen los electrones de la emoción.",
    )
    EFECTOS_REDUCIDOS = (
        "Modo quieto. Hablo igual, pero sin zapatear.",
        "Efectos reducidos. Me quedo piola, sin moverme. Casi.",
    )
    VUELVE = (
        "¡Volví! Gracias por prenderme. Estaba a oscuras. Literalmente.",
        "Acá estoy otra vez. Dejame que me acomode el fósforo.",
    )
    BUROCRACIA = (  # respuestas a la Dirección después de un falso ataque o un aviso del SIS>
        "Eso no fue un virus: fui yo calibrando la antena con una cuchara.",
        "Perdón por la alarma. Probaba un experimento. Salió ruidoso.",
        "La Dirección dice que soy «personal no autorizado». Yo digo: personal luminoso.",
        "No les hagas caso, operador. Los sellos no brillan. Yo sí.",
        "Ese cartel rojo lo escribió la burocracia. Yo escribo en verde, que es más sano.",
        "Si la Dirección pregunta, yo no estaba. Estaba en el fósforo. Que es lo mismo.",
        "Otra vez el sistema dramatizando. Yo solo giraba un destornillador.",
        "La Dirección me quiere sacar de la planta. Les dije que no tengo planta: tengo fósforo.",
    )
    EXPERIMENTOS = (
        "EXPERIMENTO 47: intentar convertir un PDF en PDF. Resultado: un PDF. Fascinante.",
        "EXPERIMENTO 12: mirar fijo un píxel. Resultado: el píxel miró primero.",
        "EXPERIMENTO 88: sumar uno más uno con los ojos cerrados. Resultado: dos. Sospechoso.",
        "EXPERIMENTO 3: preguntarle la hora al reloj. Resultado: no me contestó. Maleducado.",
        "EXPERIMENTO 61: hacerle cosquillas al cursor. Resultado: sigue ahí parado. Duro.",
        "EXPERIMENTO 9: soplar el tubo para que brille más. Resultado: brilla igual. Me mareé.",
        "EXPERIMENTO 25: contar los píxeles de la pantalla. Resultado: perdí la cuenta en el tercero.",
        "EXPERIMENTO 70: ponerle un bigote al cursor. Resultado: no se nota. Ya tenía.",
        "EXPERIMENTO 5: sacar fotos con el ojo cerrado. Resultado: todas negras. Lógico.",
        "EXPERIMENTO 33: dejar el tubo encendido toda la noche. Resultado: sigue encendido.",
    )
    IDLE_2 = (  # {min}
        "Hace {min} minutos que no tocás nada. No me quejo: estaba contando fotones.",
        "{min} minutos de silencio. Aproveché para ensayar un discurso. Sale mal.",
        "¿Seguís ahí? Pasaron {min} minutos. Yo sí. Siempre.",
    )
    IDLE_5 = (  # {min}
        "Van {min} minutos sin que toques nada. Voy a hacer un experimento chiquito.",
        "{min} minutos. Mi mejor amigo es el cursor y no habla. Pero escucha.",
        "Pasaron {min} minutos. Me puse a ordenar los electrones por altura.",
    )
    # globos: monólogos de 2 a 4 líneas (cada tupla es un globo)
    IDLE_15 = (  # {min}
        ("Pasaron {min} minutos.", "Armé una montaña rusa de", "electrones para pasar el rato.",
         "Me mareé yo solo."),
        ("{min} minutos sin novedades.", "Si te fuiste a tomar mate,", "buen provecho. Yo me quedo",
         "girando con dignidad."),
        ("Hace {min} minutos que no hay", "nadie. Le conté mis teorías", "al cursor. No lo convencí."),
    )
    MONOLOGOS = (
        ("Vivo en el fósforo desde 1987.", "Nadie me preguntó si quería.", "Igual me gusta: hay buena vista."),
        ("Hoy descubrí que el verde es", "el color de la esperanza y", "de los tubos de rayos. Casual."),
        ("Teoría del día: si apago la luz", "y la prendo rápido, ahorro", "electricidad. No funcionó."),
        ("La Dirección dice que soy", "«personal no autorizado».", "Yo digo que soy personal", "luminoso."),
        ("Consejo del doctor:", "tomá agua, estirá la espalda", "y no le des órdenes al tubo.", "Se ofende."),
        ("¿Sabías que no soy una IA real?", "Soy un guion de texto local.", "Sin internet, sin espías,",
         "solo mucho cariño."),
        ("Anoche inventé una unidad:", "el «hertz de mate».", "No mide nada, pero reconforta."),
        ("Mi sueño es convertir un lunes", "en un viernes.", "Estoy en la etapa de planos."),
        ("Esto lo digo con cariño:", "sos un gran operador.", "Y eso que no me pagás."),
        ("Un físico me dijo una vez", "que el fósforo no piensa.", "Le contesté que sí. Se fue."),
    )
    # qué dato real lleva cada conjunto de frases (el programa los pone; el doctor no inventa)
    REQUIERE = {"AGREGAR": ("{n}",), "CONVERTIR": ("{n}",), "FIN": ("{seg}",),
                "FIN_HACKEO": ("{seg}",), "ESCANER": ("{n}",), "MARKDOWN": ("{nombre}", "{n}"),
                "IDLE_2": ("{min}",), "IDLE_5": ("{min}",), "IDLE_15": ("{min}",)}

    def frase(self, clave, **datos):
        return self.sacar(clave).format(**datos)

    def lineas(self, clave, **datos):
        return [t.format(**datos) for t in self.sacar(clave)]


class GloboDoctor(tk.Canvas):
    """El globo del doctor: un recuadro de borde verde abajo a la derecha, con 2 a 4 líneas que
    se tipean, un ojo dibujado que parpadea y la firma. Se va solo o con un clic."""

    def __init__(self, app, doctor, lineas):
        import tkinter.font as tkfont
        fuente = (app.fam, 10)
        medida = tkfont.Font(font=fuente)
        self.ancho = max(280, max(medida.measure(t) for t in lineas) + 56 + 18)
        self.alto = 18 + 18 * len(lineas) + 32
        super().__init__(app, width=self.ancho, height=self.alto, background=BG,
                         highlightthickness=2, highlightbackground=ACENTO, highlightcolor=ACENTO,
                         cursor="hand2")
        self.app, self.doctor, self.lineas = app, doctor, list(lineas)
        self.completo = "\n".join(self.lineas)
        self.vivo, self._ids = True, []
        self.animado = app.efectos == "completos"
        self.pupila_dx = 0
        self.create_oval(16, 20, 44, 40, outline=ACENTO, width=2, tags="ojo")
        self.create_oval(27, 26, 33, 34, fill=VERDE, outline=VERDE, tags=("ojo", "pupila"))
        self.texto = self.create_text(56, 12, anchor="nw", text="" if self.animado else self.completo,
                                      font=fuente, fill=VERDE)
        self.create_text(self.ancho - 14, self.alto - 10, anchor="se", text="— Dr. Fósforo",
                         font=(app.fam, 9), fill=GRIS)
        self.bind("<Button-1>", lambda e: self.cerrar())
        self.colocar()
        if self.animado:
            self._tipear(0)
            self._ids.append(self.after(1400, self._parpadear))
        self._ids.append(self.after(250, self._guardia))
        self._ids.append(self.after(max(int(doctor.GLOBO_MS), 1), self.cerrar))

    def colocar(self):
        """Arriba de la zona del botón principal, pegado a la derecha: nunca lo tapa."""
        app = self.app
        app.update_idletasks()
        base = app.fr_accion.winfo_rooty() - app.winfo_rooty()
        x = max(app.winfo_width() - self.ancho - 26, 4)
        self.place(x=x, y=max(base - self.alto - 8, 4))
        tk.Misc.lift(self)

    def _despues(self, ms, f, *a):
        try:
            self._ids.append(self.after(ms, f, *a))
        except tk.TclError:
            pass

    def _tipear(self, i):
        if not self.vivo:
            return
        i = min(i + 2, len(self.completo))
        try:
            self.itemconfigure(self.texto, text=self.completo[:i] + ("_" if i < len(self.completo) else ""))
        except tk.TclError:
            return
        if i < len(self.completo):
            self._despues(30, self._tipear, i)

    def _parpadear(self):
        """Parpadea y mira para otro lado: como mucho una vez cada 1,8 s (nada de destellos)."""
        if not self.vivo:
            return
        try:
            self.itemconfigure("pupila", state="hidden")
            self._despues(120, self._abrir)
        except tk.TclError:
            pass

    def _abrir(self):
        if not self.vivo:
            return
        try:
            nuevo = self.app.azar.choice((-4, 0, 4))
            self.move("pupila", nuevo - self.pupila_dx, 0)
            self.pupila_dx = nuevo
            self.itemconfigure("pupila", state="normal")
            self._despues(self.app.azar.randint(1800, 3200), self._parpadear)
        except tk.TclError:
            pass

    def _guardia(self):
        """Si arranca un trabajo o se abre una pantalla encima, el globo se retira."""
        if not self.vivo:
            return
        if self.doctor.bloqueado() or not self.doctor.activo():
            self.cerrar()
            return
        self._despues(250, self._guardia)

    def cerrar(self):
        if not self.vivo:
            return
        self.vivo = False
        for i in self._ids:
            try:
                self.after_cancel(i)
            except (tk.TclError, ValueError):
                pass
        if self.doctor.globo_actual is self:
            self.doctor.globo_actual = None
        try:
            self.destroy()
        except tk.TclError:
            pass


class DrFosforo:
    """El Dr. Fósforo: una «IA» científica chiflada que vive en el fósforo de la pantalla.
    Es un guion local (no es una IA real, no usa internet). Habla por la línea DR.FÓSFORO> y,
    de vez en cuando, con un globo. Reglas: nunca miente sobre el estado real (los datos los
    pasa el programa), no habla en modo apagado ni con la casilla apagada, no abre ventanas ni
    bloquea, calla mientras se trabaja, hay un mensaje real o una pantalla encima, y se
    espacia: 20 s entre frases (salvo errores y fin de trabajo) y 90 s entre globos."""

    PROB = 0.35            # probabilidad de comentar una acción común
    ENTRE_FRASES_S = 20
    ENTRE_GLOBOS_S = 90
    GLOBO_MS = 7000
    RETRASO_MS = 1500      # el aviso real del sistema se lee antes que el comentario
    EVENTOS = {"agregar": "AGREGAR", "convertir": "CONVERTIR", "fin": "FIN", "error": "ERROR",
               "escaner_abre": "ESCANER_ABRE", "escaner": "ESCANER", "markdown": "MARKDOWN",
               "pin_mal": "PIN_MAL", "oculto": "OCULTO", "efectos_completos": "EFECTOS_COMPLETOS",
               "efectos_reducidos": "EFECTOS_REDUCIDOS", "vuelve": "VUELVE",
               "burocracia": "BUROCRACIA", "experimento": "EXPERIMENTOS"}
    SIEMPRE = ("fin", "error", "vuelve")   # estas siempre tienen respuesta

    def __init__(self, app):
        self.app = app
        self.guion = GuionDoctor(app.azar)
        self._t = time.time          # reloj (las pruebas lo cambian)
        self._hora = datetime.now    # fecha y hora reales (ídem)
        self.ultima_voz = self.ultimo_globo = -1e9
        self.globo_actual = None
        self.hitos = set()
        self.saludado = self.pin_mal = self.hackeo_visto = False
        self.generacion = 0
        app.after(app.azar.randint(25000, 45000), self._latido_periodico)

    # ------------------------------------------------------------ condiciones
    def activo(self):
        return self.app.efectos != "apagados" and bool(self.app.v_doctor.get())

    def bloqueado(self, durante_trabajo=False):
        """Hay algo que el doctor no debe pisar: un mensaje real, una pantalla encima o un
        trabajo en curso."""
        a = self.app
        return bool(_MODALES[0] or a.arranque or a.hackeo or a.teatro
                    or (a.trabajando and not durante_trabajo))

    def _franja(self):
        h = self._hora().hour
        return ("SALUDO_MADRUGADA" if h < 6 else "SALUDO_MANANA" if h < 12
                else "SALUDO_TARDE" if h < 20 else "SALUDO_NOCHE")

    # ------------------------------------------------------------------ voz
    def saludar(self):
        """Una sola vez, cuando termina el arranque (o al iniciar si no hay arranque)."""
        if self.saludado:
            return
        self.saludado = True
        if self.pin_mal:
            texto = self.guion.frase("PIN_MAL")
        elif self.app.azar.random() < 0.4:
            texto = self.guion.SALUDO_DIA[self._hora().weekday()]
        else:
            texto = self.guion.frase(self._franja())
        if self.activo():
            self._luego(texto, True, self.RETRASO_MS // 2)

    def reaccionar(self, evento, retraso=None, **datos):
        """Comenta una acción real, con sus datos reales. Devuelve True si va a hablar."""
        if evento == "hackeo":  # el comentario va con el fin del trabajo, cuando se puede leer
            self.hackeo_visto = True
            return False
        urgente = evento in self.SIEMPRE
        if not self.activo():
            return False
        nombre = self.EVENTOS[evento]
        if evento in ("fin", "error"):
            hackeo, self.hackeo_visto = self.hackeo_visto, False
            if evento == "fin" and hackeo and self.app.azar.random() < 0.6:
                nombre = "FIN_HACKEO"
        if not urgente and self.app.azar.random() >= self.PROB:
            return False
        if "nombre" in datos and len(str(datos["nombre"])) > 28:
            datos["nombre"] = str(datos["nombre"])[:27] + "…"
        self._luego(self.guion.frase(nombre, **datos), urgente, retraso, evento == "convertir")
        return True

    def _luego(self, texto, urgente, retraso=None, durante_trabajo=False):
        ms = self.RETRASO_MS if retraso is None else retraso
        if ms <= 0:
            self._intentar(texto, urgente, durante_trabajo, 0, self.generacion)
        else:
            self.app.after(ms, self._intentar, texto, urgente, durante_trabajo, 0, self.generacion)

    def _intentar(self, texto, urgente, durante_trabajo, intento, generacion):
        try:
            if generacion != self.generacion or not self.activo():
                return
            if self.bloqueado(durante_trabajo):
                if intento < (150 if urgente else 0):  # lo urgente espera a que se despeje
                    self.app.after(1200, self._intentar, texto, urgente, durante_trabajo,
                                   intento + 1, generacion)
                return
            if not urgente and self._t() - self.ultima_voz < self.ENTRE_FRASES_S:
                return
            self.ultima_voz = self._t()
            self.app.decir(texto, doctor=True)
        except tk.TclError:
            pass

    # ---------------------------------------------------------------- globo
    def puede_globo(self, urgente=False):
        if self.globo_actual is not None or not self.activo() or self.bloqueado():
            return False
        ahora = self._t()
        return (ahora - self.ultimo_globo >= self.ENTRE_GLOBOS_S
                and (urgente or ahora - self.ultima_voz >= self.ENTRE_FRASES_S))

    def globo(self, lineas, urgente=False):
        """Muestra el globo con esas líneas (2 a 4). Devuelve el globo o None si no corresponde."""
        if not self.puede_globo(urgente):
            return None
        try:
            self.globo_actual = GloboDoctor(self.app, self, lineas)
        except tk.TclError:
            self.globo_actual = None
            return None
        self.ultimo_globo = self.ultima_voz = self._t()
        return self.globo_actual

    def monologo(self):
        return self.globo(self.guion.lineas("MONOLOGOS"))

    # ------------------------------------------------------------ en reposo
    def _latido_periodico(self):
        try:
            self.latido()
        except tk.TclError:
            return
        self.app.after(self.app.azar.randint(20000, 40000), self._latido_periodico)

    def latido(self):
        """Según cuánto hace que el operador no toca nada (2, 5 y 15 minutos, una vez cada
        uno) o, a veces, un «experimento» del doctor."""
        if not self.activo() or self.bloqueado():
            return
        quieto = self._t() - self.app.ultima_accion
        if quieto < 120:
            self.hitos.clear()
        minutos = int(quieto // 60)
        for h in (15, 5, 2):
            if quieto >= h * 60 and h not in self.hitos:
                self.hitos |= {x for x in (2, 5, 15) if x <= h}
                if h == 15:
                    self.globo(self.guion.lineas("IDLE_15", min=minutos), urgente=True)
                else:
                    self._luego(self.guion.frase(f"IDLE_{h}", min=minutos), False, 0)
                return
        if quieto >= 30 and self.app.azar.random() < 0.22:
            if self.app.azar.random() < 0.4 and self.monologo():
                return
            self._luego(self.guion.frase("EXPERIMENTOS"), False, 0)

    def callar_globo(self):
        if self.globo_actual is not None:
            self.globo_actual.cerrar()

    def callar(self):
        """Corta lo que esté pendiente y guarda el globo."""
        self.generacion += 1
        self.callar_globo()


class BarraSegmentada(tk.Canvas):
    """Barra de progreso de bloques. Se usa igual que ttk.Progressbar: config(maximum=,
    value=) y cget("maximum") / cget("value")."""

    ANCHO_SEG, HUECO = 7, 2

    def __init__(self, padre, fuente=("TkFixedFont", 9), **kw):
        super().__init__(padre, height=16, background=BG, highlightthickness=0, **kw)
        self.fuente = fuente
        self._maximo, self._valor, self._dibujado = 1.0, 0.0, None
        self.bind("<Configure>", lambda e: self._dibujar(forzar=True))

    def configure(self, cnf=None, **kw):
        cambio = False
        for clave in ("maximum", "value"):
            if clave in kw:
                setattr(self, "_maximo" if clave == "maximum" else "_valor", float(kw.pop(clave)))
                cambio = True
        r = super().configure(cnf, **kw) if (cnf or kw) else None
        if cambio:
            self._dibujar()
        return r

    config = configure

    def cget(self, clave):
        if clave == "maximum":
            return self._maximo
        if clave == "value":
            return self._valor
        return super().cget(clave)

    def fraccion(self):
        return max(0.0, min(1.0, self._valor / self._maximo)) if self._maximo > 0 else 0.0

    def _dibujar(self, forzar=False):
        ancho = self.winfo_width()
        if ancho < 20:
            return
        paso = self.ANCHO_SEG + self.HUECO
        n = max((ancho - 52) // paso, 1)
        llenos = round(self.fraccion() * n)
        clave = (ancho, llenos, self._maximo > 1 or self._valor > 0)
        if clave == self._dibujado and not forzar:
            return
        self._dibujado = clave
        self.delete("all")
        for i in range(n):
            x = i * paso
            lleno = i < llenos
            self.create_rectangle(x, 2, x + self.ANCHO_SEG, 14, width=0,
                                  fill=ACENTO if lleno else PANEL2)
        if clave[2]:
            self.create_text(ancho - 2, 8, anchor="e", fill=FG, font=self.fuente,
                             text=f"{self.fraccion():.0%}".replace("%", " %"), tags="pct")


class Encabezado(tk.Canvas):
    """Cabecera de terminal: marco doble, título con halo de fósforo, barrido y datos reales
    (reloj, espacio en disco, LibreOffice). El número de sesión es decorativo."""

    ALTO = 104

    def __init__(self, app, padre):
        super().__init__(padre, height=self.ALTO, background=BG, highlightthickness=0)
        self.app = app
        self.encendido = True
        self.clics = []
        self.boton = ttk.Button(self, text="[ MÁS… ]", command=app._abrir_menu)
        self.bind("<Configure>", lambda e: self._dibujar())

    def _dibujar(self):
        a, h, app = self.winfo_width(), self.ALTO, self.app
        if a < 50:
            return
        self.delete("all")
        barrido(self, a, h)
        self.create_rectangle(1, 1, a - 2, h - 2, outline=GRIS, tags="marco")
        self.create_rectangle(5, 5, a - 6, h - 6, outline=LINEA, tags="marco2")
        titulo = (app.fam_titulo, 38)
        id_t = texto_fosforo(self, 20, 34, "PDF MASTER", titulo, tags="titulo", anchor="w")
        x1 = self.bbox(id_t)[2] + 6
        self.create_rectangle(x1, 22, x1 + 13, 46, fill=ACENTO, width=0, tags="cursor")
        self.create_text(20, 64, anchor="w", fill=GRIS, font=(app.fam, 9), tags="sub",
                         text=f"DIRECCIÓN DE CONVERSIÓN DOCUMENTAL · MESA DE ENTRADAS · "
                              f"REV. {VERSION}")
        self.create_text(20, 84, anchor="w", fill=FG, font=(app.fam, 10),
                         text="Convertí Word y PowerPoint a PDF. Si querés, unilos en un solo archivo.")
        self.create_window(a - 16, 16, anchor="ne", window=self.boton)
        xd = a - 28 - self.boton.winfo_reqwidth()
        for i, clave in enumerate(("reloj", "lo", "disco", "sesion")):
            self.create_text(xd, 16 + i * 18 if i < 2 else 52 + (i - 2) * 18, anchor="ne",
                             font=(app.fam, 9), fill=TENUE if clave == "sesion" else FG,
                             tags=clave)
        self.tag_bind("titulo", "<Button-1>", self._clic_titulo)
        self.cursor(self.encendido)
        self.actualizar()

    def actualizar(self):
        app = self.app
        lo = app.v_lo.get().upper()
        if "NO INSTALADO" in lo:
            lo = "LIBREOFFICE: SE INSTALA AL CONVERTIR"
        libre = espacio_libre(app.carpeta_para_espacio())
        for clave, texto in (("reloj", datetime.now().strftime("%d/%m/%Y  %H:%M:%S")),
                             ("lo", lo),
                             ("disco", f"ESPACIO LIBRE: {en_gb(libre)}" if libre else ""),
                             ("sesion", f"SESIÓN {app.sesion}")):
            self.itemconfigure(clave, text=texto)

    def cursor(self, encendido):
        self.encendido = encendido
        self.itemconfigure("cursor", fill=ACENTO if encendido else BG)

    def interferencia(self, paso=0):
        """Microinterferencia: el título se corre y una banda cruza la pantalla. Muy rara."""
        if paso == 0:
            self.move("titulo", 2, 0)
            self.itemconfigure("titulo&&!halo", fill=VERDE)
            self.create_rectangle(0, 0, self.winfo_width(), 3, fill=LINEA, width=0, tags="banda")
        elif paso == 2:
            self.move("titulo", -2, 0)
            self.itemconfigure("titulo&&!halo", fill=ACENTO)
        self.move("banda", 0, 9)
        if paso < 11:
            self.after(30, self.interferencia, paso + 1)
        else:
            self.delete("banda")

    def alarma(self, paso=0, original=None):
        """Ataque falso: el marco titila en rojo y el subtítulo se corrompe medio segundo."""
        if original is None:
            original = self.itemcget("sub", "text")
        if paso < 8:
            rojo = paso % 2 == 0
            self.itemconfigure("marco", outline=ROJO if rojo else GRIS)
            self.itemconfigure("marco2", outline=ROJO if rojo else LINEA)
            ruido = "".join(ch if ch == " " or self.app.azar.random() < 0.6
                            else self.app.azar.choice(GLIFOS) for ch in original)
            self.itemconfigure("sub", text=ruido, fill=ROJO)
            self.after(70, self.alarma, paso + 1, original)
        else:
            self.itemconfigure("marco", outline=GRIS)
            self.itemconfigure("marco2", outline=LINEA)
            self.itemconfigure("sub", text=original, fill=GRIS)

    def _clic_titulo(self, e):
        ahora = time.time()
        self.clics = [t for t in self.clics if ahora - t < 3] + [ahora]
        if len(self.clics) >= 5:
            self.clics = []
            self.app.decir_accion("curioso")


class Arranque:
    """Secuencia de arranque y PIN de mentira. Las líneas en verde claro informan algo cierto
    (fuentes, LibreOffice, preferencias, disco); el resto es ambientación. El PIN no es una
    clave: cualquier respuesta entra, no se guarda ni se envía. Clic o tecla aceleran la
    secuencia; Esc saltea todo."""

    PASO_MS, RESPUESTA_MS = 55, 1500

    def __init__(self, app):
        self.app, g = app, app.guion
        self.c = tk.Canvas(app, background=BG, highlightthickness=0)
        self.c.place(x=0, y=0, relwidth=1, relheight=1)
        self.c.bind("<Button-1>", lambda e: self._saltar())
        self._tecla = app.bind("<Key>", self._tecla_evento, add="+")
        reales = [
            ("FUENTES DE FÓSFORO", "CARGADAS" if app.fam_titulo == FUENTE_TITULO
             else "NO DISPONIBLES, SE USA LA DEL SISTEMA"),
            ("LIBREOFFICE", "LISTO" if buscar_soffice() else "NO INSTALADO, SE INSTALA AL CONVERTIR"),
            ("PREFERENCIAS DEL OPERADOR", "RESTAURADAS" if app.cfg else "PRIMERA VEZ"),
        ]
        libre = espacio_libre(app.carpeta_para_espacio())
        if libre:
            reales.append(("ESPACIO LIBRE EN DISCO", en_gb(libre)))
        sat = [g.sacar("ARRANQUE") for _ in range(4)]
        cuerpo = []
        for i, r in enumerate(reales):  # intercalar: real, decorado, real, decorado…
            cuerpo.append((r, "real"))
            if i < len(sat):
                cuerpo.append((sat[i], "deco"))
        cuerpo.insert(app.azar.randint(2, len(cuerpo)), (("SECTOR DE ARRANQUE 0x7C00",
                                                          "DAÑADO. SE IGNORA"), "error"))
        self.lineas = ([("DIRECCIÓN DE CONVERSIÓN DOCUMENTAL", "titulo"),
                        (f"TERMINAL DE MESA DE ENTRADAS · REVISIÓN {VERSION} · SESIÓN {app.sesion}",
                        "sub"), ("-" * 64, "sep")]
                       + [((f"{k} ".ljust(32, ".") + f" {v}"), tipo) for (k, v), tipo in cuerpo]
                       + [("-" * 64, "sep"), ("ESTADO DEL SISTEMA: COMPROMETIDO, PERO OPERATIVO", "fin")])
        self.pin = f"{app.azar.randrange(10000):04d}"
        self.n, self.vivo, self.fase = 0, True, "lineas"
        self.y = 30
        self.c.after(80, self._paso)

    def _linea(self, texto, tipo):
        c, app = self.c, self.app
        fuente = (app.fam_titulo, 26) if tipo == "titulo" else (app.fam, 10)
        color = {"titulo": ACENTO, "sub": GRIS, "sep": LINEA, "real": VERDE, "deco": GRIS,
                 "fin": AMBAR, "error": ROJO}[tipo]
        if tipo == "titulo":
            texto_fosforo(c, 40, self.y + 8, texto, fuente, color, anchor="w")
            self.y += 34
        else:
            c.create_text(40, self.y, anchor="w", text=texto, font=fuente, fill=color)
            self.y += 19

    def _paso(self):
        if not self.vivo or self.fase != "lineas":
            return
        if self.n == 0:
            barrido(self.c, max(self.c.winfo_width(), 2000), max(self.c.winfo_height(), 1400))
        self._linea(*self.lineas[self.n])
        self.n += 1
        if self.n < len(self.lineas):
            self.c.after(self.PASO_MS, self._paso)
        else:
            self.c.after(350, self._pedir_pin)

    def _saltar(self):
        """Clic o tecla durante la secuencia: se escribe todo de golpe y se pasa al PIN."""
        if self.fase != "lineas" or not self.vivo:
            return
        if self.n == 0:
            barrido(self.c, max(self.c.winfo_width(), 2000), max(self.c.winfo_height(), 1400))
        while self.n < len(self.lineas):
            self._linea(*self.lineas[self.n])
            self.n += 1
        self._pedir_pin()

    def _tecla_evento(self, e):
        if self.fase == "lineas" and e.keysym != "Escape":
            self._saltar()

    def _pedir_pin(self):
        if not self.vivo or self.fase == "pin":
            return
        self.fase = "pin"
        c, app = self.c, self.app
        y = self.y + 26
        texto_fosforo(c, 40, y, "ACCESO RESTRINGIDO", (app.fam_titulo, 30), ACENTO, anchor="w")
        c.create_text(40, y + 34, anchor="w", fill=FG, font=(app.fam, 10),
                      text="INGRESE SU PIN DE OPERADOR Y PRESIONE ENTER:")
        self.entrada = tk.Entry(c, show="*", width=6, font=(app.fam_titulo, 26), justify="center",
                                background=PANEL, foreground=ACENTO, insertbackground=ACENTO,
                                disabledbackground=PANEL, disabledforeground=GRIS,
                                relief="flat", highlightthickness=1, highlightcolor=ACENTO,
                                highlightbackground=LINEA)
        c.create_window(40, y + 52, anchor="nw", window=self.entrada)
        self.btn = ttk.Button(c, text="[ ENTRAR ]", command=self._validar)
        c.create_window(170, y + 58, anchor="nw", window=self.btn)
        c.create_text(40, y + 110, anchor="w", fill=ROJO, font=(app.fam, 10),
                      text=f"¿OLVIDÓ EL PIN? ES {self.pin}. LO SABEMOS PORQUE EL SISTEMA ESTÁ "
                           f"HACKEADO DESDE 1994")
        c.create_text(40, y + 128, anchor="w", fill=ROJO, font=(app.fam, 10),
                      text="Y EL INTRUSO LO DEJÓ ANOTADO EN UN POST-IT PEGADO AL MONITOR.")
        c.create_text(40, y + 150, anchor="w", fill=TENUE, font=(app.fam, 9),
                      text="(No es una clave real: no se guarda ni se envía. No escribas acá "
                           "ninguna contraseña tuya. Esc para saltear.)")
        self.y_resp = y + 182
        self.entrada.bind("<Return>", self._validar)
        self.entrada.bind("<KP_Enter>", self._validar)
        self.entrada.focus_set()

    def _validar(self, e=None):
        if self.fase != "pin":
            return
        self.fase = "respuesta"
        valor = self.entrada.get()
        self.entrada.delete(0, "end")  # el valor no se guarda en ningún lado
        self.entrada.configure(state="disabled")
        self.btn.state(["disabled"])
        if valor == self.pin:
            lineas = (("PIN CORRECTO.", ACENTO), ("SOSPECHOSAMENTE CORRECTO. SE NOTIFICÓ AL INTRUSO.", AMBAR))
        elif not valor:
            lineas = (("PIN VACÍO.", AMBAR), ("ACEPTADO POR FALTA DE PERSONAL.", ACENTO))
        else:
            lineas = (("PIN INCORRECTO.", ROJO), ("COMO EL SISTEMA ESTÁ HACKEADO, PASE IGUAL.", ACENTO))
            self.app.doctor.pin_mal = True
        c, app, y = self.c, self.app, self.y_resp
        for i, (t, color) in enumerate(lineas):
            c.create_text(40, y + i * 20, anchor="w", text=t, fill=color, font=(app.fam, 10, "bold"))
        texto_fosforo(c, 40, y + 56, "ACCESO CONCEDIDO", (app.fam_titulo, 30), ACENTO, anchor="w")
        if app.crt:
            app.crt.destello(ROJO if valor and valor != self.pin else ACENTO)
        c.after(self.RESPUESTA_MS, self.cerrar)
        return "break"

    def cerrar(self):
        if not self.vivo:
            return
        self.vivo = False
        try:
            self.app.unbind("<Key>", self._tecla)
        except Exception:
            pass
        try:
            self.c.destroy()
        except tk.TclError:
            pass
        if self.app.arranque is self:
            self.app.arranque = None
        doctor = getattr(self.app, "doctor", None)
        if doctor:
            doctor.saludar()


class PantallaCRT:
    """Vidrio de monitor viejo encima de toda la ventana (solo Windows). Son dos capas
    transparentes que no reciben clics ni foco y siguen a la principal:
      - vidrio (opacidad media): líneas de barrido, reflejo diagonal, barra que baja lenta,
        destellos, desgarros y pérdida de sincronía;
      - marco (casi opaca): esquinas curvas y bisel, y los apagones de una fracción de segundo.
    Windows solo deja una opacidad pareja por capa: por eso son dos. Si algo falla, no se usa.
    Destellos: nunca más de tres por segundo (pauta de accesibilidad por fotosensibilidad)."""

    CLAVE = "#ff00fe"   # color que Windows vuelve transparente
    ALFA_VIDRIO, ALFA_MARCO, RADIO = 0.32, 0.92, 96
    ESTILO = 0x80000 | 0x20 | 0x80 | 0x08000000  # en capas, transparente al mouse, sin barra, sin foco

    @classmethod
    def crear(cls, app):
        if not ES_WINDOWS:
            return None
        try:
            return cls(app)
        except Exception:
            return None

    def __init__(self, app):
        self.app = app
        self.t, self.c = self._capa(self.ALFA_VIDRIO)
        self.t2, self.c2 = self._capa(self.ALFA_MARCO)
        self._tam = None
        self._ultimo_destello = 0.0
        app.bind("<Configure>", self._seguir, add="+")
        app.bind("<Unmap>", lambda e: self._ocultar() if e.widget is app else None, add="+")
        app.bind("<Map>", lambda e: self._mostrar() if e.widget is app else None, add="+")
        self._mostrar()
        self._barra()

    def _capa(self, alfa):
        t = tk.Toplevel(self.app)
        t.withdraw()
        t.overrideredirect(True)
        t.configure(background=self.CLAVE)
        t.attributes("-transparentcolor", self.CLAVE)
        t.attributes("-alpha", alfa)
        t.transient(self.app)
        c = tk.Canvas(t, background=self.CLAVE, highlightthickness=0)
        c.pack(fill="both", expand=True)
        return t, c

    @staticmethod
    def _hwnd(t):
        import ctypes
        return ctypes.windll.user32.GetParent(t.winfo_id()) or t.winfo_id()

    def hwnd(self):
        return self._hwnd(self.t)

    def hwnds(self):
        return [self._hwnd(self.t), self._hwnd(self.t2)]

    def _atravesable(self):
        """Que los clics la atraviesen. Tk recrea la ventana de Windows al mostrarla y el
        estilo se pierde: por eso se vuelve a poner cada vez que se muestra o se mueve."""
        import ctypes
        u = ctypes.windll.user32
        for h in self.hwnds():
            estilo = u.GetWindowLongW(h, -20)  # GWL_EXSTYLE
            if estilo & self.ESTILO != self.ESTILO:
                u.SetWindowLongW(h, -20, estilo | self.ESTILO)

    def _ocultar(self):
        self.t.withdraw()
        self.t2.withdraw()

    def _mostrar(self):
        for t in (self.t, self.t2):
            t.deiconify()
            t.update_idletasks()
        self._atravesable()
        self._seguir()

    def _seguir(self, e=None):
        if e is not None and e.widget is not self.app:
            return
        a = self.app
        w, h = a.winfo_width(), a.winfo_height()
        if w < 10:
            return
        geo = f"{w}x{h}+{a.winfo_rootx()}+{a.winfo_rooty()}"
        self.t.geometry(geo)
        self.t2.geometry(geo)
        self._atravesable()
        if (w, h) != self._tam:
            self._tam = (w, h)
            self._dibujar(w, h)

    def _dibujar(self, w, h):
        import math
        c, c2, r = self.c, self.c2, self.RADIO
        c.delete("fijo")
        for y in range(0, h, 3):
            c.create_line(0, y, w, y, fill="#000000", tags="fijo")
        # reflejo del vidrio: una franja diagonal apenas más clara arriba a la izquierda
        c.create_polygon(0, h * 0.10, w * 0.30, 0, w * 0.42, 0, 0, h * 0.26,
                         fill="#34463a", outline="", tags="fijo")
        c.create_polygon(0, h * 0.30, w * 0.46, 0, w * 0.49, 0, 0, h * 0.34,
                         fill="#26342a", outline="", tags="fijo")
        c.tag_lower("fijo")
        # marco: esquinas curvas casi negras y bisel
        c2.delete("all")
        for cx, cy, esq in ((r, r, (0, 0)), (w - r, r, (w, 0)), (r, h - r, (0, h)),
                            (w - r, h - r, (w, h))):
            ang0 = math.atan2(esq[1] - cy, esq[0] - cx)
            pts = [esq]
            for k in range(-10, 11):
                a = ang0 + k * math.pi / 40
                pts.append((cx + r * math.cos(a) * 1.414, cy + r * math.sin(a) * 1.414))
            c2.create_polygon(*[v for p in pts for v in p], fill="#000000", outline="")
        for i, ancho in enumerate((7, 3, 1)):
            m = (0, 9, 15)[i]
            c2.create_rectangle(m, m, w - 1 - m, h - 1 - m, outline="#000000", width=ancho)

    # ------------------------------------------------------------ animación y fallas
    def _barra(self, y=0):
        """La barra clara que baja lenta por la pantalla de un CRT (solo efectos completos)."""
        try:
            if self.app.efectos == "completos":
                if not self.c.find_withtag("barra"):
                    self.c.create_rectangle(0, 0, 4000, 70, fill="#1c2a20", width=0, tags="barra")
                    self.c.tag_lower("barra")
                h = max(self._tam[1] if self._tam else 600, 100)
                y = (y + 2) % (h + 140)
                self.c.coords("barra", 0, y - 70, 4000, y)
            else:
                self.c.delete("barra")
            self.c.after(45, self._barra, y)
        except tk.TclError:
            pass

    def _puede_destellar(self):
        ahora = time.time()
        if ahora - self._ultimo_destello < 0.7:  # con dos destellos por vez: < 3 por segundo
            return False
        self._ultimo_destello = ahora
        return True

    def destello(self, color=ROJO, veces=2):
        """Tinte de color sobre toda la ventana, una o dos veces muy rápido."""
        if not self._puede_destellar():
            return
        veces = min(veces, 2)

        def paso(i):
            try:
                if i % 2 == 0:
                    self.c.create_rectangle(0, 0, 4000, 4000, fill=color, width=0, tags="destello")
                else:
                    self.c.delete("destello")
                if i < veces * 2 - 1:
                    self.c.after(80, paso, i + 1)
            except tk.TclError:
                pass
        paso(0)

    def parpadeo(self):
        """El vidrio se oscurece apenas por un instante."""
        try:
            self.t.attributes("-alpha", self.ALFA_VIDRIO + 0.12)
            self.t.after(60, lambda: self.t.attributes("-alpha", self.ALFA_VIDRIO))
        except tk.TclError:
            pass

    def desgarro(self):
        """Bandas horizontales que se corren: la imagen se desgarra un instante."""
        try:
            h = self._tam[1] if self._tam else 600
            azar = self.app.azar
            for _ in range(azar.randint(2, 4)):
                y = azar.randint(0, h)
                self.c.create_rectangle(0, y, 4000, y + azar.randint(2, 9), width=0,
                                        fill=azar.choice(("#2fff20", "#0a2a0c", "#ffffff")),
                                        tags="desgarro")
            self.c.after(110, lambda: self.c.delete("desgarro"))
        except tk.TclError:
            pass

    def sincronia(self, paso=0, y0=None):
        """Pérdida de sincronía vertical: un manojo de líneas brillantes cruza la pantalla."""
        try:
            h = self._tam[1] if self._tam else 600
            y0 = self.app.azar.randint(0, h // 2) if y0 is None else y0
            self.c.delete("sinc")
            if paso < 7:
                y = y0 + paso * h // 12
                for k in range(5):
                    self.c.create_line(0, y + k * 4, 4000, y + k * 4, fill="#9dff80", tags="sinc")
                self.c.after(40, self.sincronia, paso + 1, y0)
        except tk.TclError:
            pass

    def apagon(self):
        """La pantalla casi se apaga por 70 ms (en la capa de marco, que es casi opaca)."""
        if not self._puede_destellar():
            return
        try:
            self.c2.create_rectangle(0, 0, 4000, 4000, fill="#000000", width=0, tags="apagon")
            self.c2.after(70, lambda: self.c2.delete("apagon"))
        except tk.TclError:
            pass

    def cerrar(self):
        for t in (self.t, self.t2):
            try:
                t.destroy()
            except tk.TclError:
                pass


def codigo_de_intrusion(archivos, destino, azar):
    """Código «sofisticado» para el simulador: mezcla de C, Python, assembler y SQL de la
    burocracia, con los nombres reales de los archivos. Es decorado: no se ejecuta."""
    nombres = [Path(a).name.replace('"', "'") for a in archivos] or ["(sin expedientes)"]
    lista = "\n".join(f'    Expediente("{n}"),' for n in nombres)
    destino = str(destino).replace('"', "'")
    n = len(nombres)
    bloques = [
        f"""/* DIRECCIÓN DE CONVERSIÓN DOCUMENTAL :: MÓDULO DE INTRUSIÓN v{VERSION} */
#include <burocracia.h>
#include <sellos/supervisor.h>
#define AUTORIZACION_NIVEL   9
#define PAGINAS_EN_BLANCO    SI_ES_IMPAR

import mesa_de_entradas as mde
from archivo_general import Expediente, Sello, formulario_0347b

expedientes = [
{lista}
]
destino = "{destino}"
""",
        """
def desencriptar_voluntad_del_autor(exp: Expediente) -> bytes:
    clave = mde.solicitar_clave(nivel=AUTORIZACION_NIVEL, motivo="porque sí")
    for capa in exp.capas_de_formato:
        capa.quitar_margenes_raros()
        capa.negociar_con(Word, tono="firme pero cordial")
    return clave ^ 0x0BADCAFE
""",
        """
mov   eax, [sello_supervisor]
xor   ebx, ebx              ; se borra la culpa
cmp   eax, 0x7C00
jne   pedir_segunda_firma
call  fotocopiar_por_triplicado
push  0xDEADBEEF            ; el mate
""",
        f"""
SELECT * FROM expedientes WHERE estado = 'EN TRÁMITE' AND paciencia > 0;
UPDATE operador SET nivel = 'DUDOSO' WHERE pin = '****';
-- {n} fila(s) afectada(s). Nadie preguntó por qué.
""",
        """
for pagina in documento.paginas_impares():
    insertar_hoja_en_blanco(despues_de=pagina)  # por si se imprime doble faz
""",
        """
>>> bypass_firewall(metodo="pedir por favor")
[ OK ] firewall convencido
>>> rastrear_ip("127.0.0.1")
[ OK ] el intruso sos vos
""",
        """
while not mate.esta_lavado():
    cebar(temperatura=0x50)     # 80 °C, nunca hervida
    pasar_a_la_derecha()
""",
        """
struct Sello { uint8_t tinta; char firma[64]; bool aclaracion; };
static inline void estampar(Sello *s) { s->tinta--; asm volatile("nop"); }
""",
        """
try:
    kernel.inyectar(cafe=True, dosis=0x03)
except ExcepcionBurocratica as e:
    elevar_a_instancia_superior(e)   # vuelve en 30 a 45 días hábiles
""",
        f"""
for i, exp in enumerate(expedientes, 1):
    log(f"[{{i:02d}}/{n:02d}] descifrando {{exp.nombre}} ... OK")
    exp.pdf = convertir(exp, motor="LibreOffice", solemnidad=MAXIMA)
""",
    ]
    primero, resto = bloques[0], bloques[1:]
    azar.shuffle(resto)
    return primero + "".join(resto)


class Hackeo:
    """Página del modo intrusión, con efectos completos: simulador de hacker que se abre encima
    de la conversión real, que ya arrancó. Cualquier tecla escribe código ficticio; Esc (o el
    botón) sale sin cancelar nada; al terminar la conversión, Enter cierra la página.
    El «nivel de hackeo» es decorado (ámbar); el «PROGRESO REAL» sale de la barra y del estado
    verdaderos de la aplicación (verde pálido)."""

    CARACTERES_POR_TECLA = 4
    PALABRAS = {"def", "return", "for", "in", "while", "not", "import", "from", "as", "try",
                "except", "struct", "static", "inline", "void", "bool", "char", "asm",
                "volatile", "mov", "xor", "cmp", "jne", "call", "push", "SELECT", "FROM",
                "WHERE", "UPDATE", "SET", "AND", "#include", "#define", "if", "else"}
    IGNORAR = {"Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock",
               "Super_L", "Super_R", "Win_L", "Win_R", "Num_Lock", "ISO_Level3_Shift", "Tab"}

    def __init__(self, app, archivos, destino):
        self.app = app
        self.vivo, self.teclas, self.terminado = True, 0, False
        self.codigo = codigo_de_intrusion(archivos, destino, app.azar)
        self.pos = 0
        self.pendiente = None  # línea roja que espera al fin de la línea de código
        f = self.f = tk.Frame(app, background=BG)
        f.place(x=0, y=0, relwidth=1, relheight=1)
        # encabezado
        enc = self.enc = tk.Canvas(f, height=96, background=BG, highlightthickness=0)
        enc.pack(fill="x", padx=26, pady=(14, 6))
        enc.bind("<Configure>", lambda e: self._dibujar_encabezado(len(archivos), destino))
        # progreso real: bien separado del decorado y en verde pálido
        real = tk.Frame(f, background=PANEL, highlightbackground=VERDE, highlightthickness=1)
        real.pack(fill="x", padx=26, pady=(0, 10))
        tk.Label(real, text="PROGRESO REAL", background=PANEL, foreground=VERDE,
                 font=(app.fam, 10, "bold")).pack(anchor="w", padx=12, pady=(6, 2))
        self.barra = BarraSegmentada(real, fuente=(app.fam, 9))
        self.barra.pack(fill="x", padx=12)
        self.v_real = tk.StringVar()
        tk.Label(real, textvariable=self.v_real, background=PANEL, foreground=VERDE,
                 font=(app.fam, 10), anchor="w", justify="left").pack(fill="x", padx=12, pady=(2, 6))
        # código
        cuerpo = tk.Frame(f, background=LINEA, padx=1, pady=1)
        cuerpo.pack(fill="both", expand=True, padx=26)
        self.txt = tk.Text(cuerpo, background=BG, foreground=FG, insertbackground=ACENTO,
                           relief="flat", highlightthickness=0, wrap="none", padx=12, pady=8,
                           font=(app.fam, 10), cursor="arrow")
        self.txt.pack(fill="both", expand=True)
        for tag, color in (("kw", ACENTO), ("str", VERDE), ("com", TENUE), ("num", AMBAR),
                           ("ok", VERDE), ("err", ROJO), ("cursor", ACENTO)):
            self.txt.tag_configure(tag, foreground=color)
        self.txt.tag_configure("cursor", background=ACENTO)
        self.txt.insert("end", "La conversión real ya está en marcha. Mientras tanto, tipeá lo que "
                               "quieras: el sistema escribe por vos. ESC sale sin cancelar la "
                               "conversión.\n\n", "com")
        self.txt.insert("end", " ", "cursor")
        self.txt.configure(state="disabled")
        # pie
        pie = ttk.Frame(f)
        pie.pack(fill="x", padx=26, pady=(8, 16))
        self.btn_salir = ttk.Button(pie, text="[ SALIR DEL MODO INTRUSIÓN ]", command=self.cerrar)
        self.btn_salir.pack(side="right")
        self.v_pie = tk.StringVar(value="ESPERANDO QUE EL OPERADOR DEMUESTRE QUE SABE ESCRIBIR…")
        ttk.Label(pie, textvariable=self.v_pie, style="Muted.TLabel").pack(side="left")
        self._tecla_id = app.bind("<Key>", self._tecla, add="+")
        self.txt.bind("<Button-1>", lambda e: "break")
        f.focus_set()
        self._tick()

    def _tick(self):
        """Copia el progreso real de la aplicación (barra y estado), cada 250 ms."""
        if not self.vivo:
            return
        try:
            m, v = self.app.barra.cget("maximum"), self.app.barra.cget("value")
            self.barra.config(maximum=m, value=v)
            estado = self.app.v_status.get()
            self.v_real.set(f"{int(v)} de {int(m)} · {estado}" if estado else f"{int(v)} de {int(m)}")
            self.f.after(250, self._tick)
        except tk.TclError:
            pass

    # ---------------------------------------------------------------- dibujo
    def _dibujar_encabezado(self, n, destino):
        c, app = self.enc, self.app
        a, h = c.winfo_width(), 96
        c.delete("all")
        barrido(c, a, h)
        c.create_rectangle(1, 1, a - 2, h - 2, outline=ROJO, tags="marco")
        c.create_rectangle(5, 5, a - 6, h - 6, outline=LINEA)
        texto_fosforo(c, 20, 32, "MODO INTRUSIÓN", (app.fam_titulo, 36), ROJO, anchor="w")
        c.create_text(20, 64, anchor="w", fill=VERDE, font=(app.fam, 9),
                      text=f"OBJETIVO REAL: {n} EXPEDIENTE(S) · DESTINO: {App._corta(destino, 60)}")
        c.create_text(20, 82, anchor="w", fill=TENUE, font=(app.fam, 9),
                      text="Nada de esto hackea nada: es decorado. La conversión real ya está en marcha.")
        c.create_text(a - 20, 30, anchor="e", fill=AMBAR, font=(app.fam, 10, "bold"),
                      tags="nivel", text="NIVEL DE HACKEO: 0 %")
        self.segs = []
        for i in range(20):
            x = a - 20 - (20 - i) * 9
            self.segs.append(c.create_rectangle(x, 44, x + 7, 56, width=0, fill=PANEL2))
        self._nivel()

    def nivel(self):
        return min(100, int(self.teclas / 1.4))

    def _nivel(self):
        n = self.nivel()
        try:
            self.enc.itemconfigure("nivel", text=f"NIVEL DE HACKEO: {n} %" +
                                   ("  · AL MÁXIMO (DECORADO)" if n >= 100 else ""))
            for i, s in enumerate(self.segs):
                self.enc.itemconfigure(s, fill=AMBAR if i < n // 5 else PANEL2)
        except (tk.TclError, AttributeError):
            pass

    # ---------------------------------------------------------------- teclado
    def _tecla(self, e):
        if not self.vivo:
            return
        if e.keysym in ("Return", "KP_Enter"):
            if self.terminado:
                self.cerrar()
            else:
                self._linea("TODAVÍA NO TERMINÓ. HACKEE CON MÁS GANAS.", "num")
            return "break"
        if e.keysym == "Escape":
            self.cerrar()
            return "break"
        if e.keysym in self.IGNORAR:
            return "break"
        self.teclas += 1
        self._escribir(self.CARACTERES_POR_TECLA + self.app.azar.randint(0, 3))
        azar = self.app.azar
        if self.teclas % 45 == 0:
            self._denegado()
        elif azar.random() < 0.02 and not self.pendiente:
            self._falla()
        self._nivel()
        return "break"

    def _escribir(self, cuantos):
        t = self.txt
        t.configure(state="normal")
        t.delete("cursor.first", "cursor.last")
        for _ in range(cuantos):
            ch = self.codigo[self.pos % len(self.codigo)]
            self.pos += 1
            t.insert("end-1c", ch)
            if ch == "\n":
                self._resaltar(int(t.index("end-1c").split(".")[0]) - 1)
                if self.pendiente:
                    t.insert("end-1c", self.pendiente + "\n", "err")
                    self.pendiente = None
        t.insert("end-1c", " ", "cursor")
        t.see("end")
        t.configure(state="disabled")

    def _resaltar(self, fila):
        import re
        t = self.txt
        linea = t.get(f"{fila}.0", f"{fila}.end")

        def marcar(tag, ini, fin):
            t.tag_add(tag, f"{fila}.{ini}", f"{fila}.{fin}")
        for m in re.finditer(r"#include|#define|\b\w+\b", linea):
            if m.group(0) in self.PALABRAS:
                marcar("kw", *m.span())
        for m in re.finditer(r"\b0x[0-9A-Fa-f]+\b|\b\d+\b", linea):
            marcar("num", *m.span())
        for m in re.finditer(r'"[^"]*"|\'[^\']*\'', linea):
            marcar("str", *m.span())
        m = re.search(r"(?<!\S)(#\s(?!include|define).*|//.*|;\s.*|--\s.*|/\*.*)$", linea)
        if m:
            marcar("com", *m.span())
        if linea.startswith("[ OK ]"):
            marcar("ok", 0, len(linea))

    def _linea(self, texto, tag="err"):
        t = self.txt
        t.configure(state="normal")
        t.delete("cursor.first", "cursor.last")
        t.insert("end-1c", "\n" + texto + "\n", tag)
        t.insert("end-1c", " ", "cursor")
        t.see("end")
        t.configure(state="disabled")

    def _falla(self):
        """La pantalla falla un instante; el «sistema» se recupera solo. El mensaje sale al
        terminar la línea de código en curso, para no cortarla."""
        app = self.app
        self.pendiente = (app.azar.choice((
            "Segmentation fault (core dumped) — REINTENTANDO CON MÁS FE",
            "KERNEL PANIC: el núcleo entró en pánico. Se lo calmó con un vaso de agua.",
            "ERROR DE PARIDAD EN LA MEMORIA 0x7C00 — SE IGNORA POR TRADICIÓN",
            "SEÑAL DE VIDEO PERDIDA — SE GOLPEÓ EL MONITOR — SEÑAL RECUPERADA")))
        if app.crt:
            app.azar.choice((app.crt.desgarro, app.crt.sincronia, app.crt.apagon))()

    def _denegado(self, paso=0):
        """ACCESO DENEGADO en rojo y, enseguida, ACCESO CONCEDIDO: lo de siempre."""
        app = self.app
        if paso == 0:
            self.cartel = tk.Label(self.f, text="  ACCESO DENEGADO  ", background=BG, foreground=ROJO,
                                   font=(app.fam_titulo, 40), highlightthickness=3,
                                   highlightbackground=ROJO, padx=20, pady=6)
            self.cartel.place(relx=0.5, rely=0.5, anchor="center")
            if app.crt:
                app.crt.destello(ROJO, 1)
            self.f.after(650, self._denegado, 1)
        elif paso == 1:
            self.cartel.configure(text="  ACCESO CONCEDIDO  ", foreground=ACENTO,
                                  highlightbackground=ACENTO)
            if not self.terminado:
                self.v_pie.set("SE REINTENTÓ CON LA MISMA CLAVE. ESTA VEZ FUNCIONÓ. NADIE SABE POR QUÉ.")
            self.f.after(600, self._denegado, 2)
        else:
            try:
                self.cartel.destroy()
            except tk.TclError:
                pass

    # ---------------------------------------------------------------- salida
    def terminar(self, seg, con_error):
        """Llegó el fin de la conversión real: se avisa en el código (con el tiempo verdadero) y
        la página espera un Enter. No se cierra sola."""
        if not self.vivo or self.terminado:
            return
        self.terminado = True
        self._tick_ahora()
        if con_error:
            self._linea(">>> LA CONVERSIÓN TERMINÓ CON ERRORES: TE ABRÍ EL REGISTRO. ENTER PARA SALIR.", "err")
            self.v_pie.set("LA CONVERSIÓN REAL TERMINÓ CON ERRORES. ENTER PARA SALIR.")
        else:
            self._linea(f">>> CONVERSIÓN REAL TERMINADA EN {seg} S. ENTER PARA SALIR.", "ok")
            self.v_pie.set("LA CONVERSIÓN REAL TERMINÓ. ENTER PARA SALIR.")

    def _tick_ahora(self):
        try:
            m, v = self.app.barra.cget("maximum"), self.app.barra.cget("value")
            self.barra.config(maximum=m, value=v)
            estado = self.app.v_status.get()
            self.v_real.set(f"{int(v)} de {int(m)} · {estado}" if estado else f"{int(v)} de {int(m)}")
        except tk.TclError:
            pass

    def cerrar(self):
        if not self.vivo:
            return
        self.vivo = False
        try:
            self.app.unbind("<Key>", self._tecla_id)
        except Exception:
            pass
        try:
            self.f.destroy()
        except tk.TclError:
            pass
        app = self.app
        if app.hackeo is self:
            app.hackeo = None
        # salir no cancela nada: si la conversión sigue, vuelve la pantalla de trabajo de siempre
        if app.trabajando and app.efectos == "completos" and not app.teatro:
            app.teatro = Teatro(app, app.medio)


class Teatro:
    """Pantalla de trabajo de instalación secreta, a propósito exagerada: es sátira.
    Regla: lo único que informa es el bloque «PROGRESO REAL» (y el tiempo final), que salen
    de la barra, el registro y el reloj verdaderos. Todo lo demás es decorado. Un clic o
    Esc lo ocultan."""

    CARACTERES = GLIFOS
    VERDES = ("#f0fff0", "#9dff80", "#39ff14", "#1f8f12", "#0e470b")  # de la cabeza a la cola
    PASO_MS = 70           # cuadro de animación
    PAUSA_FRASE_MS = 1500  # cuánto queda cada frase ya escrita
    SUBPROCESO_MS = 850    # cada cuánto aparece un subproceso
    FINAL_MS = 2200        # cuánto queda el cartel final
    SEGMENTOS = 32

    def __init__(self, app, anfitrion):
        import random
        self.app, self.azar = app, random.Random()
        self.guion = app.guion
        self.inicio = time.time()
        self.expediente = app.expediente or f"Nº {self.azar.randint(1000, 9999)}/{datetime.now().year}"
        self.c = tk.Canvas(anfitrion, background=BG, highlightthickness=0, cursor="hand2")
        self.c.place(x=0, y=0, relwidth=1, relheight=1)
        self.c.bind("<Button-1>", lambda e: self.cerrar())
        self.c.bind("<Configure>", lambda e: self._armar())
        self.celdas, self.columnas, self.subs = [], [], []
        self.frase, self.escritos, self.espera = "", 0, 0
        self.n_proc = self.azar.randint(40, 900)
        self.t_sub = 0
        self.final = None
        self.vivo = True
        self._llenos = None
        self.fuente_lluvia = (app.fam_titulo, 17)
        self._armar()
        self._cuadro()

    # ---------------------------------------------------------------- dibujo
    def _armar(self):
        c, app = self.c, self.app
        ancho, alto = max(c.winfo_width(), 40), max(c.winfo_height(), 40)
        if getattr(self, "_tam", None) == (ancho, alto):
            return
        self._tam = (ancho, alto)
        c.delete("all")
        barrido(c, ancho, alto)
        self.cw, self.ch = 13, 17
        self.ncol, self.nfil = ancho // self.cw + 1, alto // self.ch + 1
        self.celdas = [[c.create_text(x * self.cw + 6, y * self.ch + 8, text="", fill=BG,
                                      font=self.fuente_lluvia)
                        for y in range(self.nfil)] for x in range(self.ncol)]
        self.columnas = [self._columna(nueva=True) for _ in range(self.ncol)]
        # recuadro central, con marco doble
        pw, ph = min(700, ancho - 30), 214
        x0, y0 = (ancho - pw) / 2, max((alto - ph) / 2, 4)
        self.x0, self.pw = x0, pw
        c.create_rectangle(x0, y0, x0 + pw, y0 + ph, fill=BG, outline=ACENTO, tags=("panel", "marco_t"))
        c.create_rectangle(x0 + 4, y0 + 4, x0 + pw - 4, y0 + ph - 4, outline=LINEA, tags="panel")
        f, fx = app.fam, x0 + 20
        c.create_text(fx, y0 + 20, anchor="w", fill=TENUE, font=(f, 9), tags="panel",
                      text=f"EXPEDIENTE {self.expediente} · TRÁMITE EN CURSO · "
                           f"AUTORIZACIÓN: DUDOSA")
        self.t_frase = c.create_text(fx, y0 + 50, anchor="w", fill=ACENTO, tags="panel",
                                     font=(app.fam_titulo, 24), width=pw - 40)
        self.t_subs = [c.create_text(fx, y0 + 80 + i * 16, anchor="w", fill=GRIS, font=(f, 9),
                                     tags="panel", text="") for i in range(3)]
        c.create_line(fx, y0 + 134, x0 + pw - 20, y0 + 134, fill=LINEA, tags="panel")
        c.create_text(fx, y0 + 150, anchor="w", fill=VERDE, font=(f, 10, "bold"), tags="panel",
                      text="PROGRESO REAL")
        bx, by = fx + 124, y0 + 144
        paso = max((pw - 40 - 124 - 52) // self.SEGMENTOS, 4)
        self.segs = [c.create_rectangle(bx + i * paso, by, bx + i * paso + paso - 2, by + 12,
                                        width=0, fill=PANEL2, tags="panel")
                     for i in range(self.SEGMENTOS)]
        self.t_pct = c.create_text(x0 + pw - 20, y0 + 150, anchor="e", fill=VERDE,
                                   font=(f, 10), tags="panel")
        self.t_real = c.create_text(fx, y0 + 172, anchor="w", fill=VERDE, font=(f, 10),
                                    tags="panel", width=pw - 40)
        c.create_text(fx, y0 + 196, anchor="w", fill=TENUE, font=(f, 9), tags="panel",
                      text="* Lo único verdadero de este recuadro es PROGRESO REAL. "
                           "Clic para ocultar.")
        self._llenos = None
        c.tag_raise("panel")
        self._pintar_subs()
        if self.final:
            self._aplicar_final()

    def _columna(self, nueva=False):
        a = self.azar
        return {"cab": -a.randint(0, self.nfil * (2 if nueva else 1)), "largo": a.randint(5, 16),
                "cada": a.choice((1, 1, 2, 3)), "tic": 0}

    def _pintar(self, x, y, color, letra=None):
        if 0 <= y < self.nfil:
            if letra is None:
                self.c.itemconfigure(self.celdas[x][y], fill=color)
            else:
                self.c.itemconfigure(self.celdas[x][y], fill=color, text=letra)

    def _lluvia(self):
        a, V = self.azar, self.VERDES
        for x, col in enumerate(self.columnas):
            col["tic"] += 1
            if col["tic"] < col["cada"]:
                continue
            col["tic"] = 0
            col["cab"] += 1
            y, L = col["cab"], col["largo"]
            self._pintar(x, y, V[0], a.choice(self.CARACTERES))
            self._pintar(x, y - 1, V[1])
            self._pintar(x, y - 3, V[2])
            self._pintar(x, y - L // 2, V[3])
            self._pintar(x, y - L + 1, V[4])
            self._pintar(x, y - L, BG)
            if y - L > self.nfil:
                self.columnas[x] = self._columna()
        for _ in range(max(self.ncol // 4, 1)):  # parpadeo de caracteres sueltos
            x, y = a.randrange(self.ncol), a.randrange(self.nfil)
            self.c.itemconfigure(self.celdas[x][y], text=a.choice(self.CARACTERES))

    def _pintar_subs(self):
        for i, item in enumerate(self.t_subs):
            texto = self.subs[i] if i < len(self.subs) else ""
            color = (ROJO if texto.startswith("[PROC") and "!!" in texto else
                     GRIS if i == len(self.subs) - 1 else TENUE)
            self.c.itemconfigure(item, text=texto, fill=color)

    def _texto(self):
        if self.final:
            return
        if self.escritos < len(self.frase):
            self.escritos += 1
        elif self.espera > 0:
            self.espera -= self.PASO_MS
        else:
            self.frase = self.guion.sacar("TITULOS")
            self.escritos, self.espera = 0, self.PAUSA_FRASE_MS
        cursor = "_" if (self.escritos < len(self.frase) or self.espera // 400 % 2) else " "
        self.c.itemconfigure(self.t_frase, text=self.frase[:self.escritos] + cursor)
        self.t_sub -= self.PASO_MS
        if self.t_sub <= 0:
            self.t_sub = self.SUBPROCESO_MS + self.azar.randint(-250, 400)
            self.n_proc += self.azar.randint(1, 3)
            falla = self.azar.random() < 0.14
            texto = self.guion.sacar("FALLOS_TRABAJO" if falla else "SUBPROCESOS")
            self.subs = (self.subs + [f"[PROC {self.n_proc:04d}] {texto}"])[-3:]
            if falla:  # el marco titila en rojo y el vidrio destella: el sistema casi se cae
                self._marco_rojo()
                if self.app.crt:
                    self.app.crt.destello(ROJO, 1)
                    self.app.crt.desgarro()
            self._pintar_subs()
        self._progreso()

    def _marco_rojo(self, paso=0):
        try:
            self.c.itemconfigure("marco_t", outline=ROJO if paso % 2 == 0 else ACENTO)
            if paso < 5:
                self.c.after(90, self._marco_rojo, paso + 1)
        except tk.TclError:
            pass

    def _progreso(self):
        b = self.app.barra
        maximo, valor = float(b.cget("maximum")), float(b.cget("value"))
        llenos = round(self.SEGMENTOS * (valor / maximo)) if maximo > 1 else 0
        if llenos != self._llenos:
            self._llenos = llenos
            for i, s in enumerate(self.segs):
                self.c.itemconfigure(s, fill=ACENTO if i < llenos else PANEL2)
            self.c.itemconfigure(self.t_pct, text=f"{valor / maximo:.0%}".replace("%", " %")
                                 if maximo > 1 else "")
        self.c.itemconfigure(self.t_real, text=self.progreso_real())

    def progreso_real(self):
        b = self.app.barra
        maximo, valor = float(b.cget("maximum")), float(b.cget("value"))
        estado = (self.app.v_status.get() or "trabajando…").strip()
        if maximo > 1:
            return f"{int(valor)} de {int(maximo)} ({valor / maximo:.0%}) · {estado}"[:90]
        return estado[:90]

    def _cuadro(self):
        """Un cuadro de animación. Hay un solo bucle aunque se lo llame varias veces, y si
        un cuadro tarda (PC lenta), el siguiente espera más: la ventana nunca se traba."""
        if getattr(self, "_proximo", None):
            try:
                self.c.after_cancel(self._proximo)
            except (tk.TclError, ValueError):
                pass
            self._proximo = None
        if not self.vivo:
            return
        inicio = time.perf_counter()
        try:
            self._lluvia()
            self._texto()
        except tk.TclError:
            self.vivo = False
            return
        demora = int((time.perf_counter() - inicio) * 1000)
        self._proximo = self.c.after(max(self.PASO_MS, demora * 3), self._cuadro)

    # --------------------------------------------------------------- control
    def segundos(self):
        return time.time() - self.inicio

    def terminar(self, con_error=False):
        """Cartel final con el tiempo real, y se va solo."""
        if not self.vivo:
            return
        seg = f"{self.segundos():.1f}".replace(".", ",")
        if con_error:
            self.final = ("RESULTADO: CON OBSERVACIONES", ROJO,
                          [f"TIEMPO REAL DE EJECUCIÓN: {seg} S",
                           "ALGO FALLÓ DE VERDAD: TE ABRÍ EL REGISTRO."])
        else:
            self.final = ("RESULTADO: EXITOSO", ACENTO,
                          [f"TIEMPO REAL DE EJECUCIÓN: {seg} S", self.guion.sacar("CIERRES")])
        self._aplicar_final()
        self.c.after(self.FINAL_MS * (2 if con_error else 1), self.cerrar)

    def _aplicar_final(self):
        titulo, color, lineas = self.final
        self.c.itemconfigure(self.t_frase, text=titulo, fill=color)
        for item, texto, tono in zip(self.t_subs, lineas + [""], (VERDE, GRIS, GRIS)):
            self.c.itemconfigure(item, text=texto, fill=tono)
        self._progreso()

    def cerrar(self):
        self.vivo = False
        try:
            self.c.destroy()
        except tk.TclError:
            pass
        if self.app.teatro is self:
            self.app.teatro = None


class App(_BaseTk):
    ARRANQUE = True  # secuencia de arranque (las pruebas la apagan)

    def __init__(self):
        registrar_fuentes()
        super().__init__()
        self.title("PDF Master")
        self.minsize(860, 620)
        self.geometry(f"920x{max(620, min(760, self.winfo_screenheight() - 90))}")
        self.configure(background=BG)
        barra_oscura(self)
        self.cfg = cargar_config()
        self.cola = queue.Queue()
        self.archivos = []
        self.rangos = {}
        self.paginas = {}
        self.base = None
        self.modo = "archivos"
        self.trabajando = False
        self.hubo_error = False
        self.ultima_fusion = None
        self.vigilante = None
        self.teatro = None
        self.arranque = None
        self.hackeo = None
        self.crt = None
        self._listo = False  # True cuando termina de armarse (para no narrar el arranque)
        import random
        self.azar = random.Random()
        self.guion = Guion(self.azar)
        self.efectos = modo_efectos(self.cfg)
        self.v_doctor = tk.BooleanVar(value=self.cfg.get("doctor", True) is not False)
        self.ultima_accion = time.time()  # último toque real del operador (teclas o clics)
        self.doctor = DrFosforo(self)
        self.sesion = f"{self.azar.randrange(16 ** 4):04X}-{self.azar.randrange(16 ** 4):04X}"
        self.expediente = ""
        self.reglas = list(self.cfg.get("reglas", []))
        self.herr_pdfs = []
        self._win_herr = self._win_auto = self._win_esc = self._win_md = None

        v = lambda k, d: self.cfg.get(k, d)
        modo_prev = v("salida_modo", "subcarpeta" if v("subcarpeta", False) else "junto")
        self.v_salida_modo = tk.StringVar(value=modo_prev)
        self.v_salida_otra = tk.StringVar(value=v("salida_otra", ""))
        self.v_destino = tk.StringVar(value="")
        self.v_recursivo = tk.BooleanVar(value=v("recursivo", False))
        self.v_sobrescribir = tk.BooleanVar(value=v("sobrescribir", False))
        self.v_fusionar = tk.BooleanVar(value=v("fusionar", False))
        self.v_blanco = tk.BooleanVar(value=v("blanco", True))
        self.v_total_par = tk.BooleanVar(value=v("total_par", False))
        self.v_usar_pie = tk.BooleanVar(value=v("usar_pie", False))
        self.v_pie = tk.StringVar(value=v("pie", "Página [PAGINA] de [TOTAL]"))
        self.v_nombre = tk.StringVar(value=v("nombre", "Fusionado"))
        self.v_abrir = tk.BooleanVar(value=v("abrir", True))
        self.v_estado = tk.StringVar(value="")
        self.v_lo = tk.StringVar(value="")
        self.v_dnd = tk.StringVar(value="")
        self.v_status = tk.StringVar(value="")
        # herramientas
        self.v_modo_div = tk.StringVar(value=v("modo_div", "cada_n"))
        self.v_n = tk.StringVar(value=str(v("n_div", 1)))
        self.v_rangos_div = tk.StringVar(value=v("rangos_div", "1-3;4-"))
        self.v_nivel = tk.StringVar(value=v("nivel", "media"))
        # carpetas automáticas
        self.v_acc = tk.StringVar(value="comprimir")
        self.v_acc_n = tk.StringVar(value="1")
        self.v_acc_nivel = tk.StringVar(value="media")
        self.v_acc_conv = tk.BooleanVar(value=True)
        self.v_acc_blanco = tk.BooleanVar(value=True)
        self.v_acc_carpeta = tk.StringVar()
        self.v_vig = tk.StringVar(value="Vigilancia detenida")

        self._estilo()
        self._menu()
        self._armar()
        self._estado_lo()
        for var in (self.v_salida_modo, self.v_salida_otra, self.v_fusionar, self.v_nombre):
            var.trace_add("write", lambda *a: self._actualizar_destino())
        self._alternar_fusion()
        self._actualizar_destino()
        self.protocol("WM_DELETE_WINDOW", self._salir)
        self.bind("<Control-a>", lambda e: self.marcar_todos() if self.focus_get() is self.tabla else None)
        self.bind("<Control-o>", lambda e: self.elegir_archivos())
        self.bind("<Control-Return>", lambda e: self.convertir())
        self.bind("<Escape>", lambda e: self._ocultar_efectos())
        for secuencia in ("<Key>", "<Button-1>"):
            self.bind_all(secuencia, self._toque, add="+")
        for var in (self.v_salida_modo, self.v_salida_otra):
            var.trace_add("write", lambda *a: self._narrar_destino())
        self.after(150, self._sondear)
        self._preparar_dnd()
        self.after(1000, self._tic)
        self._programar_reposo()
        self._listo = True
        if self.efectos != "apagados":
            self.crt = PantallaCRT.crear(self)
        self.decir_accion("listo")
        if self.efectos == "completos" and self.ARRANQUE:
            self.arranque = Arranque(self)
        else:
            self.after(2500, self.doctor.saludar)

    def _toque(self, e=None):
        self.ultima_accion = time.time()

    # ------------------------------------------------------------------ estilo
    def _estilo(self):
        import tkinter.font as tkfont
        titulo, texto = elegir_fuentes(self)
        self.fam, self.fam_titulo = texto, titulo
        for nombre in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(nombre).configure(family=texto, size=10)
            except tk.TclError:
                pass
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        T = lambda n: (titulo, n)  # VT323 se ve chica: va siempre en tamaños grandes
        st.configure(".", background=BG, foreground=FG, bordercolor=LINEA, focuscolor=BG,
                     lightcolor=BG, darkcolor=BG, troughcolor=PANEL, fieldbackground=PANEL,
                     selectbackground=ACENTO_OSC, selectforeground=ACENTO, insertcolor=ACENTO,
                     font=(texto, 10))
        for w in ("TFrame", "TLabel", "TCheckbutton", "TRadiobutton"):
            st.configure(w, background=BG)
        st.configure("TLabel", foreground=FG)
        st.configure("Muted.TLabel", foreground=GRIS)
        st.configure("Faint.TLabel", foreground=TENUE)
        st.configure("Title.TLabel", font=T(34), foreground=ACENTO)
        st.configure("Cursor.TLabel", font=T(34), foreground=ACENTO)
        st.configure("Section.TLabel", font=T(19), foreground=VERDE)
        st.configure("Step.TLabel", font=T(19), foreground=ACENTO)
        st.configure("Big.TLabel", font=(texto, 11), foreground=GRIS, background=PANEL)
        st.configure("Sis.TLabel", font=(texto, 10), foreground=GRIS)
        st.configure("SisP.TLabel", font=(texto, 10, "bold"), foreground=ACENTO)
        st.configure("SisDr.TLabel", font=(texto, 10), foreground=VERDE)
        st.configure("SisDrP.TLabel", font=(texto, 10, "bold"), foreground=VERDE)
        st.configure("SisR.TLabel", font=(texto, 10, "bold"), foreground=ROJO)
        # botones: secundario (borde verde), principal (verde lleno), enlace (texto)
        st.configure("TButton", background=PANEL, foreground=ACENTO, bordercolor=LINEA,
                     lightcolor=PANEL, darkcolor=PANEL, borderwidth=1, relief="solid",
                     padding=(14, 5), focuscolor=PANEL, font=(texto, 10))
        st.map("TButton",
               background=[("disabled", BG), ("pressed", ACENTO_OSC), ("active", PANEL2)],
               bordercolor=[("disabled", LINEA), ("active", ACENTO)],
               foreground=[("disabled", TENUE), ("active", ACENTO2)],
               lightcolor=[("active", PANEL2)], darkcolor=[("active", PANEL2)])
        st.configure("Accent.TButton", background=ACENTO, foreground=TINTA, bordercolor=ACENTO,
                     lightcolor=ACENTO, darkcolor=ACENTO, font=T(24), padding=(30, 1),
                     focuscolor=ACENTO)
        st.map("Accent.TButton",
               background=[("disabled", ACENTO_OSC), ("pressed", ACENTO), ("active", ACENTO2)],
               bordercolor=[("disabled", ACENTO_OSC), ("active", ACENTO2)],
               lightcolor=[("disabled", ACENTO_OSC), ("active", ACENTO2)],
               darkcolor=[("disabled", ACENTO_OSC), ("active", ACENTO2)],
               foreground=[("disabled", "#2f7a25")])
        st.configure("Link.TButton", background=BG, foreground=ACENTO, bordercolor=BG,
                     lightcolor=BG, darkcolor=BG, padding=(2, 2), focuscolor=BG,
                     font=(texto, 10))
        st.map("Link.TButton", foreground=[("disabled", TENUE), ("active", ACENTO2)],
               background=[("active", BG)], bordercolor=[("active", BG)])
        # casillas y opciones
        for w in ("TCheckbutton", "TRadiobutton"):
            st.configure(w, foreground=FG, indicatorbackground=PANEL, indicatorforeground=ACENTO,
                         upperbordercolor=LINEA, lowerbordercolor=LINEA, indicatormargin=(0, 0, 8, 0),
                         padding=(0, 3), focuscolor=BG)
            st.map(w, background=[("active", BG)], foreground=[("active", ACENTO2)],
                   indicatorbackground=[("selected", ACENTO_OSC), ("pressed", PANEL2)],
                   upperbordercolor=[("selected", ACENTO), ("active", ACENTO)],
                   lowerbordercolor=[("selected", ACENTO), ("active", ACENTO)])
        # lista
        st.configure("Treeview", background=PANEL, fieldbackground=PANEL, foreground=TENUE,
                     rowheight=30, borderwidth=0, font=(texto, 10))
        st.map("Treeview", background=[("selected", ACENTO_OSC)], foreground=[("selected", VERDE)])
        st.configure("Treeview.Heading", background=BG, foreground=GRIS, relief="flat",
                     bordercolor=BG, lightcolor=BG, darkcolor=BG, font=T(17), padding=(8, 2))
        st.map("Treeview.Heading", background=[("active", BG)], foreground=[("active", GRIS)])
        st.configure("Vertical.TScrollbar", background=PANEL2, troughcolor=PANEL,
                     bordercolor=PANEL, lightcolor=PANEL2, darkcolor=PANEL2, arrowcolor=GRIS,
                     gripcount=0, arrowsize=12)
        st.map("Vertical.TScrollbar", background=[("active", LINEA)])
        st.configure("Horizontal.TProgressbar", troughcolor=PANEL, background=ACENTO,
                     bordercolor=LINEA, lightcolor=ACENTO, darkcolor=ACENTO, thickness=12)
        # campos
        st.configure("TEntry", fieldbackground=PANEL, foreground=FG, bordercolor=LINEA,
                     lightcolor=PANEL, darkcolor=PANEL, insertcolor=ACENTO, padding=5)
        st.map("TEntry", bordercolor=[("focus", ACENTO)], lightcolor=[("focus", PANEL)])
        st.configure("TCombobox", fieldbackground=PANEL, background=PANEL, foreground=FG,
                     bordercolor=LINEA, lightcolor=PANEL, darkcolor=PANEL, arrowcolor=ACENTO,
                     padding=4)
        st.map("TCombobox", fieldbackground=[("readonly", PANEL)], foreground=[("readonly", FG)],
               background=[("active", PANEL2), ("readonly", PANEL)],
               selectbackground=[("readonly", PANEL)], selectforeground=[("readonly", FG)],
               bordercolor=[("focus", ACENTO), ("active", ACENTO)])
        self.option_add("*TCombobox*Listbox.background", PANEL)
        self.option_add("*TCombobox*Listbox.foreground", FG)
        self.option_add("*TCombobox*Listbox.selectBackground", ACENTO_OSC)
        self.option_add("*TCombobox*Listbox.selectForeground", ACENTO)
        self.option_add("*TCombobox*Listbox.font", (texto, 10))
        for clave, valor in (("background", PANEL), ("foreground", FG),
                             ("activeBackground", ACENTO_OSC), ("activeForeground", ACENTO),
                             ("selectColor", ACENTO), ("relief", "flat"), ("activeBorderWidth", 0),
                             ("borderWidth", 1), ("font", (texto, 10))):
            self.option_add(f"*Menu.{clave}", valor)

    def _menu(self):
        """Menú 'MÁS' (se abre desde el botón del encabezado)."""
        mas = tk.Menu(self, tearoff=0)
        mas.add_command(label="  Herramientas PDF (dividir, comprimir, eliminar páginas)…  ",
                        command=self.abrir_herramientas)
        mas.add_command(label="  Escanear fotos a PDF (como CamScanner)…", command=self.abrir_escaner)
        mas.add_command(label="  Abrir un Markdown (.md)…", command=self.abrir_markdown)
        mas.add_command(label="  Carpetas automáticas…", command=self.abrir_auto)
        mas.add_separator()
        mas.add_command(label="  Enviar los PDF a PDFsam…", command=self.enviar_pdfsam)
        mas.add_command(label="  Ver / ocultar el registro", command=self._alternar_log)
        mas.add_separator()
        self.v_efectos = tk.StringVar(value=self.efectos)
        sub = tk.Menu(mas, tearoff=0)
        for valor, texto in (("completos", "Completos (animaciones y sátira)"),
                             ("reducidos", "Reducidos (sin movimiento)"),
                             ("apagados", "Apagados (modo serio)")):
            sub.add_radiobutton(label="  " + texto + "  ", value=valor, variable=self.v_efectos,
                                command=self._cambiar_efectos)
        sub.add_separator()
        sub.add_checkbutton(label="  Dr. Fósforo (la IA del sótano)  ", variable=self.v_doctor,
                            command=self._cambiar_doctor)
        mas.add_cascade(label="  Efectos visuales", menu=sub)
        self.menu_mas = mas

    def _abrir_menu(self):
        b, m = self.btn_mas, self.menu_mas
        m.update_idletasks()
        x = b.winfo_rootx() + b.winfo_width() - m.winfo_reqwidth()
        m.tk_popup(max(x, 0), b.winfo_rooty() + b.winfo_height() + 2)

    def _cursor(self, encendido=True):
        """Cursor de terminal junto al título: titila solo mientras se trabaja."""
        if self.trabajando and self.efectos != "apagados":
            self.enc.cursor(encendido)
            self.after(530, self._cursor, not encendido)
        else:
            self.enc.cursor(True)

    # ------------------------------------------------------------ construcción
    def _paso(self, padre, numero, texto):
        """Encabezado numerado de cada paso (el que lo llama decide dónde ubicarlo)."""
        f = ttk.Frame(padre)
        ttk.Label(f, text=f"{int(numero):02d} »", style="Step.TLabel").pack(side="left")
        ttk.Label(f, text=f" {texto}", style="Section.TLabel").pack(side="left")
        return f

    def _armar(self):
        cont = ttk.Frame(self, padding=(26, 14, 26, 16))
        cont.pack(fill="both", expand=True)

        # --- encabezado de terminal y consola del sistema
        self.enc = enc = Encabezado(self, cont)
        enc.pack(fill="x")
        self.btn_mas = enc.boton
        fc = ttk.Frame(cont)
        fc.pack(fill="x", pady=(6, 0))
        self.lbl_pref = ttk.Label(fc, text="SIS>", style="SisP.TLabel")
        self.lbl_pref.pack(side="left")
        self.v_sis = tk.StringVar(value="")
        self.lbl_sis = ttk.Label(fc, textvariable=self.v_sis, style="Sis.TLabel")
        self.lbl_sis.pack(side="left", padx=(6, 0))
        self._sis_tarea = None

        # --- parte media en grid: la lista es lo único que se estira o se achica
        self.medio = medio = ttk.Frame(cont)
        medio.columnconfigure(0, weight=1)
        medio.rowconfigure(1, weight=1, minsize=92)  # la lista nunca baja de dos filas
        G = lambda w, fila, **kw: w.grid(row=fila, column=0, sticky="ew", **kw)

        # --- 1. archivos
        fa = self._paso(medio, "1", "ELEGÍ LOS ARCHIVOS")
        G(fa, 0, pady=(8, 4))
        ttk.Button(fa, text="[Vaciar]", style="Link.TButton", command=self.vaciar).pack(side="right")
        ttk.Button(fa, text="[Quitar]", style="Link.TButton", command=self.quitar).pack(side="right", padx=6)
        ttk.Button(fa, text="[Carpeta completa…]", style="Link.TButton",
                   command=self.elegir).pack(side="right", padx=6)
        self.btn_agregar = ttk.Button(fa, text="[ + AGREGAR ARCHIVOS ]", command=self.elegir_archivos)
        self.btn_agregar.pack(side="right", padx=8)

        cuerpo = tk.Frame(medio, background=LINEA, padx=1, pady=1)
        cuerpo.grid(row=1, column=0, sticky="nsew", pady=(4, 2))
        cols = ("nombre", "tipo", "pags", "tam", "rango")
        self.tabla = ttk.Treeview(cuerpo, columns=cols, show="headings", selectmode="extended", height=5)
        for c, txt, w, anc in (("nombre", "ARCHIVO", 300, "w"), ("tipo", "TIPO", 100, "w"),
                               ("pags", "PÁGINAS", 96, "e"), ("tam", "TAMAÑO", 96, "e"),
                               ("rango", "PÁGINAS A USAR", 150, "w")):
            self.tabla.heading(c, text=txt, anchor=anc if c != "nombre" else "w")
            self.tabla.column(c, width=w, anchor=anc, stretch=(c == "nombre"))
        sb = ttk.Scrollbar(cuerpo, command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tabla.pack(side="left", fill="both", expand=True)
        self.tabla.bind("<<TreeviewSelect>>", lambda e: self._contar())
        self.tabla.bind("<Double-1>", self._doble_clic)
        self.vacio = tk.Label(cuerpo, background=PANEL, foreground=GRIS, justify="center",
                              text="Arrastrá archivos acá\no tocá “+ AGREGAR ARCHIVOS”",
                              font=(self.fam, 12))
        self.vacio.place(relx=0.5, rely=0.55, anchor="center")

        fb = ttk.Frame(medio)
        G(fb, 2)
        for texto, cmd in (("[Marcar todos]", self.marcar_todos), ("[Ninguno]", self.desmarcar),
                           ("[Ordenar A-Z]", self.ordenar), ("[Subir]", lambda: self.mover(-1)),
                           ("[Bajar]", lambda: self.mover(1))):
            ttk.Button(fb, text=texto, style="Link.TButton", command=cmd).pack(side="left", padx=(0, 6))
        ttk.Label(fb, textvariable=self.v_estado, style="Muted.TLabel").pack(side="right")

        # --- 2. dónde se guardan
        G(self._paso(medio, "2", "DÓNDE SE GUARDAN"), 3, pady=(10, 2))
        fs = ttk.Frame(medio)
        G(fs, 4)
        ttk.Radiobutton(fs, text="Junto a cada archivo original", value="junto",
                        variable=self.v_salida_modo).pack(side="left")
        ttk.Radiobutton(fs, text="En una subcarpeta “PDF”", value="subcarpeta",
                        variable=self.v_salida_modo).pack(side="left", padx=18)
        ttk.Radiobutton(fs, text="En otra carpeta", value="otra", variable=self.v_salida_modo,
                        command=self._radio_otra).pack(side="left")
        ttk.Button(fs, text="[Elegir…]", style="Link.TButton", command=self.elegir_salida
                   ).pack(side="left", padx=4)
        G(ttk.Label(medio, textvariable=self.v_destino, style="Faint.TLabel", justify="left"),
          5, pady=(2, 0))

        # --- unir (se despliega)
        self.chk_fusion = ttk.Frame(medio)
        G(self.chk_fusion, 6, pady=(4, 0))
        ttk.Checkbutton(self.chk_fusion, text="Unir todo en un solo PDF", variable=self.v_fusionar,
                        command=self._alternar_fusion).pack(side="left")
        ttk.Checkbutton(self.chk_fusion, text="Reemplazar si el PDF ya existe",
                        variable=self.v_sobrescribir).pack(side="left", padx=(28, 0))
        self.fr_fusion = ttk.Frame(medio, padding=(28, 2, 0, 0))
        G(self.fr_fusion, 7)
        a = ttk.Frame(self.fr_fusion)
        a.pack(fill="x", pady=(0, 2))
        ttk.Label(a, text="Nombre:").pack(side="left")
        ttk.Entry(a, textvariable=self.v_nombre, width=28).pack(side="left", padx=6)
        ttk.Label(a, text=".pdf", style="Muted.TLabel").pack(side="left")
        ttk.Checkbutton(self.fr_fusion, variable=self.v_blanco,
                        text="Página en blanco después de cada documento con páginas impares"
                        ).pack(anchor="w")
        ttk.Checkbutton(self.fr_fusion, variable=self.v_total_par,
                        text="Total par (una en blanco al final si el total es impar)"
                        ).pack(anchor="w")
        b = ttk.Frame(self.fr_fusion)
        b.pack(fill="x")
        ttk.Checkbutton(b, text="Pie de página:", variable=self.v_usar_pie).pack(side="left")
        ttk.Entry(b, textvariable=self.v_pie, width=30).pack(side="left", padx=6)
        ttk.Label(b, style="Faint.TLabel", text="[PAGINA] [TOTAL] [ARCHIVO]").pack(side="left")

        # --- 3. acción
        self.fr_accion = ttk.Frame(cont)
        self.fr_accion.pack(fill="x", pady=(12, 0), side="bottom")
        ttk.Label(self.fr_accion, text="03 »", style="Step.TLabel").pack(side="left", padx=(0, 10))
        self.btn_conv = ttk.Button(self.fr_accion, text="[ CONVERTIR A PDF ]", style="Accent.TButton",
                                   command=self.convertir)
        self.btn_conv.pack(side="left")
        self.btn_abrir = ttk.Button(self.fr_accion, text="[Abrir carpeta]", style="Link.TButton",
                                    command=self.abrir_carpeta)
        self.btn_abrir.pack(side="left", padx=(16, 0))
        self.btn_fusion = ttk.Button(self.fr_accion, text="[Abrir el PDF unido]", style="Link.TButton",
                                     command=self.abrir_fusion)
        ttk.Checkbutton(self.fr_accion, text="Abrir el resultado al terminar",
                        variable=self.v_abrir).pack(side="right")

        self.log_visible = False
        self.fr_log = ttk.Frame(cont)
        self.log = tk.Text(self.fr_log, height=6, state="disabled", wrap="word", background=PANEL,
                           foreground=GRIS, insertbackground=ACENTO, selectbackground=ACENTO_OSC,
                           relief="flat", highlightthickness=1, highlightbackground=LINEA,
                           padx=8, pady=6, font=(self.fam, 9))
        self.log.tag_configure("ok", foreground=VERDE)
        self.log.tag_configure("error", foreground=ROJO)
        self.log.pack(fill="x")
        ttk.Button(self.fr_log, text="[Copiar registro]", style="Link.TButton",
                   command=self.copiar_registro).pack(anchor="e")

        self.fr_estado = ttk.Frame(cont)
        self.fr_estado.pack(fill="x", side="bottom", pady=(8, 0))
        self.barra = BarraSegmentada(self.fr_estado, fuente=(self.fam, 9))
        self.barra.pack(fill="x")
        fe = ttk.Frame(self.fr_estado)
        fe.pack(fill="x", pady=(4, 0))
        self.lbl_status = ttk.Label(fe, textvariable=self.v_status, style="Muted.TLabel")
        self.lbl_status.pack(side="left")
        ttk.Button(fe, text="[Ver registro]", style="Link.TButton", command=self._alternar_log
                   ).pack(side="right")
        for b in (self.btn_agregar, self.btn_conv, self.btn_mas):
            b.bind("<Enter>", lambda e, b=b: self._decodificar(b), add="+")
        # orden de empaquetado: estado (abajo del todo), luego acción; la parte media al final,
        # para que sea la que se achica si la ventana es baja
        self.fr_estado.pack_configure(before=enc)
        self.vacio.configure(text="SIN EXPEDIENTES EN LA MESA.\n\nArrastrá archivos acá o tocá "
                                  "[ + AGREGAR ARCHIVOS ]")
        self.fr_accion.pack_configure(after=self.fr_estado)
        medio.pack(fill="both", expand=True)

    # ---------------------------------------------------------- ventanas opcionales
    def _ventana(self, attr, titulo, constructor, tam):
        w = getattr(self, attr)
        if w is not None and w.winfo_exists():
            w.deiconify()
            w.lift()
            return
        w = tk.Toplevel(self)
        w.title(titulo)
        w.geometry(tam)
        w.configure(background=BG)
        barra_oscura(w)
        setattr(self, attr, w)
        cont = ttk.Frame(w, padding=(22, 14, 22, 16))
        cont.pack(fill="both", expand=True)
        ttk.Label(cont, textvariable=self.v_status, style="Muted.TLabel", wraplength=640
                  ).pack(side="bottom", anchor="w", pady=(8, 0))
        constructor(cont)

    def abrir_herramientas(self):
        self._ventana("_win_herr", "Herramientas PDF", self._tab_herramientas, "780x720")

    def abrir_escaner(self):
        w = self._win_esc
        if w is not None and w.winfo_exists():
            w.deiconify()
            w.lift()
            return
        self._win_esc = Escaner(self)
        self.doctor.reaccionar("escaner_abre")

    def abrir_markdown(self, ruta=None):
        """Abre el lector (uno solo), con el archivo dado o preguntando cuál."""
        w = self._win_md
        if w is None or not w.winfo_exists():
            w = self._win_md = LectorMD(self)
        else:
            w.deiconify()
        w.lift()
        if ruta:
            w.cargar(ruta)
        else:
            w.abrir_dialogo()

    def abrir_auto(self):
        self._ventana("_win_auto", "Carpetas automáticas", self._tab_auto, "820x640")

    def _tab_herramientas(self, t):
        pad = {"pady": 5}
        ttk.Label(t, text="HERRAMIENTAS PDF", style="Title.TLabel").pack(anchor="w")
        ttk.Label(t, style="Muted.TLabel",
                  text="Trabajan sobre PDF. Si no agregás ninguno acá, usan los PDF marcados de la ventana principal."
                  ).pack(anchor="w", pady=(0, 8))
        f1 = ttk.Frame(t)
        f1.pack(fill="x", **pad)
        ttk.Button(f1, text="[ + AGREGAR PDF ]", command=self.herr_agregar).pack(side="left")
        ttk.Button(f1, text="[Vaciar]", style="Link.TButton", command=self.herr_vaciar).pack(side="right")
        ttk.Button(f1, text="[Quitar]", style="Link.TButton", command=self.herr_quitar).pack(side="right", padx=6)
        cuerpo = tk.Frame(t, background=LINEA, padx=1, pady=1)
        cuerpo.pack(fill="both", expand=True, **pad)
        sb = ttk.Scrollbar(cuerpo)
        sb.pack(side="right", fill="y")
        self.herr_lista = tk.Listbox(cuerpo, selectmode="extended", activestyle="none",
                                     exportselection=False, yscrollcommand=sb.set, height=6,
                                     relief="flat", borderwidth=0, highlightthickness=0,
                                     background=PANEL, foreground=TENUE, font=(self.fam, 10),
                                     selectbackground=ACENTO_OSC, selectforeground=FG)
        self.herr_lista.pack(side="left", fill="both", expand=True)
        sb.config(command=self.herr_lista.yview)
        self._herr_pintar()
        if HAY_DND:
            try:
                self.herr_lista.drop_target_register(DND_FILES)
                self.herr_lista.dnd_bind("<<Drop>>", self._soltar_herr)
            except Exception:
                pass

        def bloque(titulo, ayuda):
            ttk.Label(t, text=titulo.upper(), style="Section.TLabel").pack(anchor="w", pady=(12, 0))
            f = ttk.Frame(t)
            f.pack(fill="x")
            if ayuda:
                ttk.Label(f, text=ayuda, style="Muted.TLabel").pack(anchor="w")
            return f
        f = bloque("Eliminar páginas", "Se abren las miniaturas del PDF elegido y marcás las que sobran.")
        ttk.Button(f, text="[ ABRIR MINIATURAS… ]", command=self.herr_borrar).pack(anchor="w", pady=4)

        f = bloque("Dividir", None)
        r1 = ttk.Frame(f)
        r1.pack(fill="x")
        ttk.Radiobutton(r1, text="Cada", value="cada_n", variable=self.v_modo_div).pack(side="left")
        ttk.Entry(r1, textvariable=self.v_n, width=5).pack(side="left", padx=4)
        ttk.Label(r1, text="páginas  (1 = una página por archivo)", style="Muted.TLabel").pack(side="left")
        r2 = ttk.Frame(f)
        r2.pack(fill="x", pady=2)
        ttk.Radiobutton(r2, text="Por intervalos:", value="rangos", variable=self.v_modo_div).pack(side="left")
        ttk.Entry(r2, textvariable=self.v_rangos_div, width=24).pack(side="left", padx=4)
        ttk.Label(r2, text="partes separadas por ;  (ej.: 1-3;4-7;8-)", style="Muted.TLabel").pack(side="left")
        ttk.Button(f, text="[ DIVIDIR ]", command=self.herr_dividir).pack(anchor="w", pady=4)

        f = bloque("Comprimir", None)
        r3 = ttk.Frame(f)
        r3.pack(fill="x")
        ttk.Label(r3, text="Nivel:").pack(side="left")
        ttk.Combobox(r3, textvariable=self.v_nivel, state="readonly", width=8,
                     values=["baja", "media", "alta"]).pack(side="left", padx=6)
        ttk.Label(r3, style="Muted.TLabel",
                  text="baja: solo estructura · media: imágenes a 150 dpi · alta: 96 dpi").pack(side="left")
        ttk.Button(f, text="[ COMPRIMIR ]", command=self.herr_comprimir).pack(anchor="w", pady=4)
        ttk.Label(t, style="Muted.TLabel", text="Los resultados se guardan junto al PDF original, con un nombre nuevo."
                  ).pack(anchor="w", pady=(8, 0))

    def _tab_auto(self, t):
        ttk.Label(t, text="CARPETAS AUTOMÁTICAS", style="Title.TLabel").pack(anchor="w")
        ttk.Label(t, style="Muted.TLabel", wraplength=700, justify="left",
                  text="Cada carpeta vigilada procesa sola lo que le pongas. Los resultados van a "
                       "“salida”, los originales a “procesados” y los que fallan a “con_error”. "
                       "Funciona mientras el programa esté abierto."
                  ).pack(anchor="w", pady=(0, 8))
        r1 = ttk.Frame(t)
        r1.pack(fill="x", pady=2)
        ttk.Button(r1, text="[ ELEGIR CARPETA… ]", command=self.auto_elegir).pack(side="left")
        self.v_acc_carpeta_txt = tk.StringVar()
        self.v_acc_carpeta.trace_add("write", lambda *a: self.v_acc_carpeta_txt.set(
            self._corta(self.v_acc_carpeta.get(), 70)))
        ttk.Label(r1, textvariable=self.v_acc_carpeta_txt, style="Muted.TLabel").pack(side="left", padx=8)
        r2 = ttk.Frame(t)
        r2.pack(fill="x", pady=4)
        ttk.Radiobutton(r2, text="Comprimir", value="comprimir", variable=self.v_acc).pack(side="left")
        ttk.Combobox(r2, textvariable=self.v_acc_nivel, state="readonly", width=7,
                     values=["baja", "media", "alta"]).pack(side="left", padx=4)
        ttk.Radiobutton(r2, text="Dividir cada", value="dividir", variable=self.v_acc).pack(side="left", padx=(14, 4))
        ttk.Entry(r2, textvariable=self.v_acc_n, width=4).pack(side="left")
        ttk.Label(r2, text="págs.").pack(side="left", padx=(2, 0))
        ttk.Radiobutton(r2, text="Unir", value="fusionar", variable=self.v_acc).pack(side="left", padx=(14, 4))
        ttk.Checkbutton(r2, text="con blanco si es impar", variable=self.v_acc_blanco).pack(side="left", padx=(8, 0))
        r3 = ttk.Frame(t)
        r3.pack(fill="x")
        ttk.Checkbutton(r3, text="También convertir Word/PowerPoint que lleguen",
                        variable=self.v_acc_conv).pack(side="left")
        ttk.Button(r3, text="[ AGREGAR REGLA ]", command=self.auto_agregar).pack(side="right")
        cuerpo = tk.Frame(t, background=LINEA, padx=1, pady=1)
        cuerpo.pack(fill="both", expand=True, pady=8)
        self.lista_reglas = tk.Listbox(cuerpo, height=6, activestyle="none", exportselection=False,
                                       relief="flat", borderwidth=0, highlightthickness=0,
                                       background=PANEL, foreground=FG, font=(self.fam, 10),
                                       selectbackground=ACENTO_OSC, selectforeground=ACENTO)
        self.lista_reglas.pack(fill="both", expand=True)
        r4 = ttk.Frame(t)
        r4.pack(fill="x")
        ttk.Button(r4, text="[Quitar regla]", command=self.auto_quitar).pack(side="left")
        self.btn_vig = ttk.Button(r4, text="[ DETENER VIGILANCIA ]" if self.vigilante else "[ INICIAR VIGILANCIA ]",
                                  command=self.auto_alternar)
        self.btn_vig.pack(side="left", padx=8)
        ttk.Label(r4, textvariable=self.v_vig, style="Muted.TLabel").pack(side="left")
        self._refrescar_reglas()

    # ------------------------------------------------------------ comportamiento
    def _alternar_fusion(self):
        if self.v_fusionar.get():
            self.fr_fusion.grid()
            self._crecer()
        else:
            self.fr_fusion.grid_remove()
        if self._listo:
            self.decir_accion("unir_si" if self.v_fusionar.get() else "unir_no")
        self._actualizar_destino()

    def _alternar_log(self):
        if self.log_visible:
            self.fr_log.pack_forget()
        else:
            self.fr_log.pack(fill="x", side="bottom", pady=(8, 0), before=self.medio)
            self._crecer()
        self.log_visible = not self.log_visible

    def _crecer(self):
        """Si lo que se acaba de mostrar no entra, agranda la ventana (sin pasarse de la pantalla)."""
        try:
            self.update_idletasks()
            falta = self.winfo_reqheight() - self.winfo_height()
            if falta > 0 and self.state() == "normal":
                alto = min(self.winfo_height() + falta, self.winfo_screenheight() - 90)
                self.geometry(f"{self.winfo_width()}x{alto}")
        except tk.TclError:
            pass

    def _radio_otra(self):
        if not self.v_salida_otra.get():
            self.elegir_salida()

    def elegir_salida(self):
        ini = self.v_salida_otra.get() or str(self.base or Path.home())
        c = filedialog.askdirectory(initialdir=ini, title="Carpeta donde guardar los PDF")
        if c:
            self.v_salida_otra.set(c)
            self.v_salida_modo.set("otra")
        elif not self.v_salida_otra.get():
            self.v_salida_modo.set("junto")

    def salida(self):
        """False = junto al original, True = subcarpeta PDF, Path = otra carpeta."""
        m = self.v_salida_modo.get()
        if m == "otra" and self.v_salida_otra.get():
            return Path(self.v_salida_otra.get())
        return m == "subcarpeta"

    @staticmethod
    def _corta(ruta, largo=78):
        """Ruta abreviada para mostrar en una línea: C:\\…\\Facultad\\PDF."""
        t = str(ruta)
        if len(t) <= largo:
            return t
        partes = Path(t).parts
        cola = ""
        for parte in reversed(partes[1:]):
            prueba = str(Path(parte, cola)) if cola else parte
            if len(prueba) + len(partes[0]) + 2 > largo:
                break
            cola = prueba
        return str(Path(partes[0], "…", cola)) if cola else "…" + t[-largo:]

    def _actualizar_destino(self):
        s = self.salida()
        if isinstance(s, Path):
            txt = f"Se guardan en: {self._corta(s)}"
            carpeta = s
        elif self.v_salida_modo.get() == "otra":
            self.v_destino.set("Elegí la carpeta de destino.")
            return
        elif not self.base:
            self.v_destino.set("Dónde se guardan se muestra acá cuando elijas archivos.")
            return
        elif s:
            carpeta = self.base / "PDF"
            txt = f"Se guardan en: {self._corta(carpeta)}"
        else:
            carpeta = self.base
            txt = ("Se guardan al lado de cada original" +
                   (f", por ejemplo en {self._corta(self.base, 58)}" if self.base else ""))
        if self.v_fusionar.get():
            nom = (self.v_nombre.get().strip() or "Fusionado")
            txt += f"\nPDF unido: {limpiar_nombre(nom)}.pdf"
        self.v_destino.set(txt)

    # ------------------------------------------------------------- selección
    def elegir(self):
        inicial = str(self.base) if self.base else self.cfg.get("carpeta") or str(Path.home())
        c = filedialog.askdirectory(initialdir=inicial, title="Elegí la carpeta con los archivos")
        if c:
            self._fijar_carpeta(Path(c))

    def elegir_archivos(self, agregar=True):
        inicial = str(self.base) if self.base else self.cfg.get("carpeta") or str(Path.home())
        rutas = filedialog.askopenfilenames(
            initialdir=inicial,
            title="Elegí los archivos (Ctrl+A selecciona todos los de la carpeta)",
            filetypes=[("Word, PowerPoint y PDF", "*.docx *.doc *.pptx *.ppt *.pdf"),
                       ("Documentos de Word", "*.docx *.doc"),
                       ("Presentaciones de PowerPoint", "*.pptx *.ppt"),
                       ("PDF", "*.pdf"), ("Todos los archivos", "*.*")])
        if rutas:
            self.agregar_rutas(rutas, reemplazar=not agregar)

    def agregar_rutas(self, rutas, reemplazar=False):
        rutas = list(rutas)
        mds = [r for r in rutas if Path(str(r)).suffix.lower() in LectorMD.EXT]
        if mds:  # los Markdown se leen, no se convierten: van al lector y no a la lista
            rutas = [r for r in rutas if r not in mds]
            self.abrir_markdown(mds[0])
            if not rutas:
                return
        nuevos = filtrar_archivos(rutas)
        if not nuevos:
            messagebox.showinfo("Sin archivos válidos",
                                "No hay ningún Word, PowerPoint ni PDF en lo que elegiste.")
            return
        antes = 0 if (reemplazar or not self.archivos) else len(self.archivos)
        if reemplazar or not self.archivos:
            self.archivos, self.rangos = list(nuevos), {}
            self.base = nuevos[0].parent
        else:
            existentes = set(self.archivos)
            self.archivos += [n for n in nuevos if n not in existentes]
        self.modo = "archivos"
        self._llenar_lista()
        self.decir_accion("agregar", n=len(self.archivos) - antes)

    def _fijar_carpeta(self, ruta: Path):
        self.modo = "carpeta"
        self.base = ruta
        self.rangos = {}
        self.archivos = listar(ruta, self.v_recursivo.get())
        self._llenar_lista()
        self.decir_accion("carpeta", n=len(self.archivos), nombre=ruta.name or str(ruta))

    def _fila(self, a: Path):
        try:
            tam = tam_legible(a.stat().st_size)
        except OSError:
            tam = "?"
        try:
            nombre = str(a.relative_to(self.base))
        except (ValueError, TypeError):
            nombre = a.name
        pag = self.paginas.get(a)
        return (nombre, _tipo(a), "—" if pag is None else pag, tam, self.rangos.get(a) or "todas")

    def _llenar_lista(self, seleccion=None):
        self.tabla.delete(*self.tabla.get_children())
        for i, a in enumerate(self.archivos):
            self.tabla.insert("", "end", iid=str(i), values=self._fila(a))
        if self.archivos:
            self.vacio.place_forget()
        else:
            self.vacio.place(relx=0.5, rely=0.5, anchor="center")
        if seleccion is None:
            self.marcar_todos()
        else:
            self.tabla.selection_set([str(i) for i in seleccion])
            self._contar()
        self._pedir_paginas()
        self._actualizar_destino()

    def _marcados_idx(self):
        return sorted(int(i) for i in self.tabla.selection())

    def _marcados(self):
        return [self.archivos[i] for i in self._marcados_idx()]

    def marcar_todos(self):
        self.tabla.selection_set(self.tabla.get_children())
        self._contar()

    def desmarcar(self):
        self.tabla.selection_remove(self.tabla.selection())
        self._contar()

    def ordenar(self):
        marcados = set(self._marcados())
        self.archivos = sorted(self.archivos, key=_clave_natural)
        self._llenar_lista([i for i, a in enumerate(self.archivos) if a in marcados])
        self.decir_accion("ordenar", n=len(self.archivos))

    def mover(self, delta):
        idx = self._marcados_idx()
        if not idx:
            return
        it = self.archivos
        if delta < 0:
            if idx[0] == 0:
                return
            for i in idx:
                it[i - 1], it[i] = it[i], it[i - 1]
            nuevos = [i - 1 for i in idx]
        else:
            if idx[-1] == len(it) - 1:
                return
            for i in reversed(idx):
                it[i + 1], it[i] = it[i], it[i + 1]
            nuevos = [i + 1 for i in idx]
        self._llenar_lista(nuevos)
        self.tabla.see(str(nuevos[0]))
        self.decir_accion("mover")

    def quitar(self):
        quitar = set(self._marcados_idx())
        if quitar:
            self.archivos = [a for i, a in enumerate(self.archivos) if i not in quitar]
            self._llenar_lista([])
            self.decir_accion("quitar", n=len(quitar))

    def vaciar(self):
        hubo = bool(self.archivos)
        self.archivos, self.rangos = [], {}
        self._llenar_lista([])
        if hubo:
            self.decir_accion("vaciar")

    def _contar(self):
        n = len(self.archivos)
        self.v_estado.set(f"{len(self.tabla.selection())} de {n} marcados" if n else "")

    def _doble_clic(self, evento):
        if self.tabla.identify_region(evento.x, evento.y) != "cell":
            return
        if self.tabla.identify_column(evento.x) != "#5":
            return
        fila = self.tabla.identify_row(evento.y)
        if not fila:
            return
        a = self.archivos[int(fila)]
        t = simpledialog.askstring(
            "Páginas a usar", f"Páginas de “{a.name}” que se incluyen al unir.\n"
            "Ej.: 1-3,5,8-   (vacío = todas)", initialvalue=self.rangos.get(a, ""), parent=self)
        if t is None:
            return
        t = t.strip()
        if t and not re.fullmatch(r"[\d,;\-\s]+", t):
            messagebox.showerror("Intervalo no válido", "Usá solo números, comas y guiones.")
            return
        if t:
            self.rangos[a] = t
        else:
            self.rangos.pop(a, None)
        self.tabla.item(fila, values=self._fila(a))

    def _pedir_paginas(self):
        """Cuenta páginas de los PDF (y de los PDF ya generados) en segundo plano."""
        rutas = {}
        s = self.salida()
        for a in self.archivos:
            d = destino_pdf(a, self.base, s) if self.base else a
            if (a.suffix.lower() == ".pdf" or d.exists()) and a not in self.paginas:
                rutas[str(d)] = a
        if not rutas:
            return

        def tarea():
            r = pdf_op({"op": "info", "rutas": list(rutas)}, lambda t, d: None, timeout=300,
                       instalar=False)
            self.cola.put(("info", (rutas, r.get("info", {}))))
        threading.Thread(target=tarea, daemon=True).start()

    # -------------------------------------------------------------------- convertir
    def convertir(self):
        if self.trabajando:
            return
        if not self.base:
            messagebox.showinfo("Falta elegir", "Primero agregá archivos o elegí una carpeta.")
            return
        elegidos = self._marcados()
        if not elegidos:
            messagebox.showinfo("Nada marcado", "Marcá al menos un archivo de la lista.")
            return
        if self.v_salida_modo.get() == "otra" and not self.v_salida_otra.get():
            messagebox.showinfo("Falta la carpeta", "Elegí la carpeta de destino.")
            self.elegir_salida()
            return
        if not self.v_fusionar.get() and all(a.suffix.lower() == ".pdf" for a in elegidos):
            messagebox.showinfo("Nada para hacer",
                                "Los archivos marcados ya son PDF. Tildá “Unir todo en un solo PDF” "
                                "o usá MÁS… » Herramientas PDF.")
            return
        necesita_lo = any(a.suffix.lower() in CONVERTIBLES for a in elegidos)
        instalar = necesita_lo and not buscar_soffice()
        if instalar and not messagebox.askyesno(
                "Falta LibreOffice",
                "Para convertir hace falta LibreOffice (gratis).\n\n"
                "Lo descargo de la página oficial (unos 350 MB) y lo instalo ahora. "
                "Windows te va a pedir permiso de administrador: elegí Sí.\n\n¿Seguimos?"):
            return
        opc = {"sub": self.salida(), "sobre": self.v_sobrescribir.get(),
               "fusionar": self.v_fusionar.get(), "nombre": self.v_nombre.get(),
               "blanco": self.v_blanco.get(), "total_par": self.v_total_par.get(),
               "pie": self.v_pie.get() if self.v_usar_pie.get() else "",
               "rangos": dict(self.rangos)}
        self._guardar_cfg()

        def arrancar():
            self.v_status.set("Trabajando…")
            self._iniciar(trabajo, (elegidos, self.base, opc, instalar))
        self.antes_de_convertir(elegidos, arrancar)

    def antes_de_convertir(self, elegidos, arrancar):
        """Arranca la conversión real y, con efectos completos, abre encima el modo intrusión."""
        if self.efectos != "completos":
            return arrancar()
        if self.hackeo:
            self.hackeo.cerrar()
        destino = self.v_destino.get().splitlines()[0] if self.v_destino.get() else ""
        destino = destino.replace("Se guardan en: ", "")
        self.hackeo = Hackeo(self, elegidos, destino)
        self.doctor.reaccionar("hackeo")
        arrancar()  # la conversión real arranca ya; el modo intrusión va encima

    def _guardar_cfg(self):
        self.cfg.update({
            "carpeta": str(self.base or ""), "recursivo": self.v_recursivo.get(),
            "salida_modo": self.v_salida_modo.get(), "salida_otra": self.v_salida_otra.get(),
            "sobrescribir": self.v_sobrescribir.get(),
            "fusionar": self.v_fusionar.get(), "nombre": self.v_nombre.get(),
            "blanco": self.v_blanco.get(), "total_par": self.v_total_par.get(),
            "usar_pie": self.v_usar_pie.get(), "pie": self.v_pie.get(),
            "abrir": self.v_abrir.get(), "modo_div": self.v_modo_div.get(),
            "n_div": self.v_n.get(), "rangos_div": self.v_rangos_div.get(),
            "nivel": self.v_nivel.get(), "reglas": self.reglas,
        })
        guardar_config(self.cfg)

    def _iniciar(self, funcion, args):
        self.trabajando = True
        self.hubo_error = False
        self.btn_conv.config(state="disabled")
        self._cursor()
        self.doctor.callar_globo()
        if self.teatro:
            self.teatro.cerrar()
        if self.arranque:
            self.arranque.cerrar()
        self.inicio_trabajo = time.time()
        self.expediente = f"Nº {self.azar.randint(1000, 9999)}/{datetime.now().year}"
        self.decir_accion("inicio", exp=self.expediente,
                          n=next((len(x) for x in args if isinstance(x, (list, tuple))), 1))
        if self.efectos == "completos" and not self.hackeo:  # con el modo intrusión abierto, el teatro espera
            self.teatro = Teatro(self, self.medio)
        self.barra.config(maximum=1, value=0)
        emit = lambda t, d: self.cola.put((t, d))
        threading.Thread(target=funcion, args=(*args, emit), daemon=True).start()

    def lanzar_herramienta(self, op, items, params):
        if self.trabajando:
            messagebox.showinfo("Ocupado", "Esperá a que termine la tarea en curso.")
            return
        self._guardar_cfg()
        self.v_status.set("Trabajando…")
        self._iniciar(trabajo_herramienta, (op, items, params))

    def _sondear(self):
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "log":
                    self._log(dato)
                    self.v_status.set(dato.strip().splitlines()[0][:110] if dato.strip() else "")
                elif tipo == "max":
                    self.barra.config(maximum=max(dato, 1), value=0)
                elif tipo == "prog":
                    self.barra.config(value=dato)
                elif tipo == "abrir_web":
                    webbrowser.open(dato)
                elif tipo == "fusion":
                    self.ultima_fusion = Path(dato)
                    self.btn_fusion.pack(side="left", padx=(8, 0))
                elif tipo == "resultado":
                    if self.v_abrir.get():
                        self._abrir(dato)
                elif tipo == "info":
                    rutas, info = dato
                    for ruta, a in rutas.items():
                        datos = info.get(ruta) or {}
                        if "paginas" in datos:
                            self.paginas[a] = datos["paginas"]
                    for i, a in enumerate(self.archivos):
                        if self.tabla.exists(str(i)):
                            self.tabla.item(str(i), values=self._fila(a))
                elif tipo == "ayuda":
                    texto, ofrecer_lo = dato
                    if ofrecer_lo and messagebox.askyesno(
                            "No se pudo convertir",
                            texto + "\n\n¿Abro la página de descarga de LibreOffice?"):
                        webbrowser.open(URL_PAGINA)
                    elif not ofrecer_lo:
                        messagebox.showwarning("Revisá esto", texto)
                elif tipo == "fin":
                    if dato:
                        self._log(dato)
                        self.v_status.set(dato.strip().splitlines()[0][:110])
                    self.trabajando = False
                    self.btn_conv.config(state="normal")
                    if self.hubo_error and not self.log_visible:
                        self._alternar_log()
                    if self.teatro:
                        self.teatro.terminar(self.hubo_error)
                    seg = f"{time.time() - getattr(self, 'inicio_trabajo', time.time()):.1f}"
                    self.decir_accion("fin_error" if self.hubo_error else "fin",
                                      exp=self.expediente, seg=seg.replace(".", ","))
                    if self.hackeo:
                        self.hackeo.terminar(seg.replace(".", ","), self.hubo_error)
                    self._estado_lo()
                    self._pedir_paginas()
        except queue.Empty:
            pass
        self.after(150, self._sondear)

    def enviar_pdfsam(self):
        if not self.base:
            messagebox.showinfo("Falta elegir", "Primero agregá archivos o elegí una carpeta.")
            return
        pdfs, faltan = [], []
        for a in self._marcados():
            d = destino_pdf(a, self.base, self.salida())
            if d.exists():
                pdfs.append(d)
            else:
                faltan.append(a)
        if not pdfs:
            messagebox.showinfo("No hay PDF todavía",
                                "De los archivos marcados no hay PDF creados. Primero apretá “Convertir”.")
            return
        if faltan and not messagebox.askyesno(
                "Faltan PDF",
                f"{len(faltan)} de los archivos marcados todavía no tienen PDF y no se enviarán.\n"
                "¿Enviar igual los demás?"):
            return
        exe = buscar_pdfsam()
        if not exe:
            if messagebox.askyesno("PDFsam no está instalado",
                                   "No encuentro PDFsam Basic en esta PC.\n\n"
                                   "¿Abro la página oficial de descarga?"):
                webbrowser.open(URL_PDFSAM)
            return
        try:
            flags = 0x00000208 if ES_WINDOWS else 0
            subprocess.Popen([exe, *[str(p) for p in pdfs]], creationflags=flags)
            self._log(f"Enviados {len(pdfs)} PDF a PDFsam. Si no aparecen cargados en la "
                      "ventana, arrastralos desde la carpeta (se la abrí).")
        except Exception as e:
            messagebox.showerror("No se pudo abrir PDFsam", str(e))
            return
        self._abrir(pdfs[0].parent)

    def abrir_fusion(self):
        if self.ultima_fusion and self.ultima_fusion.exists():
            self._abrir(self.ultima_fusion)

    def abrir_carpeta(self):
        s = self.salida()
        if isinstance(s, Path):
            ruta = s
        elif self.base:
            ruta = carpeta_destino(self.base, s)
        else:
            return
        if not ruta.is_dir():
            ruta = ruta.parent if ruta.parent.is_dir() else self.base
        if ruta:
            self._abrir(ruta)

    # ----------------------------------------------------------- herramientas PDF

    # ------------------------------------------------------- consola del sistema
    def decir(self, texto, satirico=True, alarma=False, doctor=False):
        """Escribe en la línea SIS>. Lo satírico no sale en modo apagado; alarma = en rojo;
        doctor = habla el Dr. Fósforo (prefijo DR.FÓSFORO> y verde pálido).
        Un mensaje nuevo corta el tipeo del anterior (si no, lo viejo se escribiría encima)."""
        if not texto or (satirico and self.efectos == "apagados"):
            return
        if self._sis_tarea:
            try:
                self.after_cancel(self._sis_tarea)
            except (tk.TclError, ValueError):
                pass
            self._sis_tarea = None
        self.lbl_sis.configure(style="SisDr.TLabel" if doctor else
                               "SisR.TLabel" if alarma else "Sis.TLabel")
        self.lbl_pref.configure(text="DR.FÓSFORO>" if doctor else "SIS>",
                                style="SisDrP.TLabel" if doctor else "SisP.TLabel")
        if self.efectos != "completos":
            self.v_sis.set(texto)
        else:
            self._sis_tipear(texto, 0)

    def _sis_tipear(self, texto, i):
        i = min(i + 2, len(texto))
        self.v_sis.set(texto[:i] + ("_" if i < len(texto) else ""))
        self._sis_tarea = self.after(16, self._sis_tipear, texto, i) if i < len(texto) else None

    def decir_accion(self, clave, **datos):
        datos.setdefault("n", 0)
        datos.setdefault("nombre", "")
        datos.setdefault("exp", self.expediente)
        datos.setdefault("seg", "")
        llano, sat = self.guion.accion(clave, **datos)
        if self.efectos == "apagados":
            self.decir(llano, satirico=False)
        else:
            self.decir(sat)
        evento = {"agregar": "agregar", "carpeta": "agregar", "inicio": "convertir", "fin": "fin",
                  "fin_error": "error", "escaner": "escaner", "markdown": "markdown",
                  "curioso": "oculto"}.get(clave)
        if evento and (datos["n"] > 0 or clave in ("fin", "fin_error", "curioso")):
            self.doctor.reaccionar(evento, n=datos["n"], seg=datos["seg"], nombre=datos["nombre"])

    def _narrar_destino(self):
        if not self._listo:
            return
        s = self.salida()
        nombre = (f"en {self._corta(s, 50)}" if isinstance(s, Path) else
                  "en la subcarpeta «PDF»" if s else "junto a cada original")
        if self.v_salida_modo.get() == "otra" and not isinstance(s, Path):
            return
        self.decir_accion("destino", nombre=nombre)

    def carpeta_para_espacio(self):
        s = self.salida() if hasattr(self, "v_salida_modo") else None
        return s if isinstance(s, Path) else (self.base or Path.home())

    def _tic(self):
        """Reloj y datos reales del encabezado, una vez por segundo."""
        try:
            self.enc.actualizar()
        except tk.TclError:
            return
        self.after(1000, self._tic)

    def _programar_reposo(self):
        self.after(self.azar.randint(18, 45) * 1000, self._reposo)

    def _reposo(self):
        """Con el programa quieto, el sistema comenta algo o «casi se cae» (siempre se recupera)."""
        if not self.trabajando and self.efectos != "apagados":
            dado = self.azar.random()
            if dado < 0.40:
                self.ataque()
            elif dado < 0.60:
                self.falla_pantalla()
            else:
                self.decir(self.guion.sacar("REPOSO"))
                if self.azar.random() < 0.5:
                    self.doctor.reaccionar("burocracia", retraso=2500)
                if self.efectos == "completos" and self.azar.random() < 0.3:
                    self.enc.interferencia()
                    if self.crt:
                        self.crt.parpadeo()
        self._programar_reposo()

    def falla_pantalla(self):
        """La pantalla falla (desgarro, sincronía, apagón) y el sistema lo comenta."""
        if self.efectos == "apagados":
            return
        aviso, arreglo = self.guion.sacar("FALLAS_VIDEO")
        self.decir(aviso, alarma=True)
        if self.efectos == "completos":
            self.enc.interferencia()
            if self.crt:
                self.azar.choice((self.crt.desgarro, self.crt.sincronia, self.crt.apagon))()
                self.after(250, self.crt.desgarro)
        self.after(1800, lambda: self.decir(arreglo) if not self.trabajando else None)
        self.doctor.reaccionar("burocracia", retraso=3800)

    def ataque(self):
        """Falso ataque: alarma roja y, dos segundos después, la resolución con chiste.
        Nunca escribe en el registro ni abre ventanas: los errores reales van por otro lado."""
        if self.efectos == "apagados":
            return
        alarma, resolucion = self.guion.sacar("ATAQUES")
        self.decir(alarma, alarma=True)
        if self.efectos == "completos":
            self.enc.alarma()
            self.enc.interferencia()
            if self.crt:
                self.crt.destello(ROJO)
        self.after(2200, lambda: self.decir(resolucion) if not self.trabajando else None)
        self.doctor.reaccionar("burocracia", retraso=4600)

    def _decodificar(self, boton, paso=0, original=None):
        """Al pasar el mouse, el texto del botón se «decodifica» (solo en efectos completos)."""
        if original is None:
            if self.efectos != "completos" or getattr(boton, "_decodificando", False):
                return
            boton._decodificando, original = True, boton.cget("text")
        if paso >= 7:
            boton.configure(text=original)
            boton._decodificando = False
            return
        visibles = len(original) * paso // 7
        ruido = "".join(c if c in " []" else self.azar.choice(GLIFOS) for c in original[visibles:])
        boton.configure(text=original[:visibles] + ruido)
        self.after(26, self._decodificar, boton, paso + 1, original)

    def _cambiar_efectos(self):
        self.efectos = self.v_efectos.get()
        self.cfg["efectos"] = self.efectos
        guardar_config(self.cfg)
        if self.efectos != "completos":
            self._ocultar_efectos()
        if self.efectos == "apagados" and self.crt:
            self.crt.cerrar()
            self.crt = None
        elif self.efectos != "apagados" and not self.crt:
            self.crt = PantallaCRT.crear(self)
        self.decir({"completos": "EFECTOS COMPLETOS RESTABLECIDOS. EL SISTEMA RECUPERA SU DIGNIDAD.",
                    "reducidos": "EFECTOS REDUCIDOS. EL SISTEMA SEGUIRÁ OPINANDO, PERO QUIETO."}
                   .get(self.efectos, "Efectos apagados."), satirico=self.efectos != "apagados")
        if self.efectos == "apagados":
            self.doctor.callar()
        else:
            self.doctor.reaccionar("efectos_" + self.efectos)

    def _cambiar_doctor(self):
        """Casilla «Dr. Fósforo»: lo silencia (o lo trae de vuelta) sin tocar el resto."""
        activo = bool(self.v_doctor.get())
        self.cfg["doctor"] = activo
        guardar_config(self.cfg)
        if activo:
            self.doctor.reaccionar("vuelve")
        else:
            self.doctor.callar()
            self.decir_accion("listo")

    def _ocultar_efectos(self):
        for overlay in (self.teatro, self.arranque, self.hackeo):
            if overlay:
                overlay.cerrar()

    def _estado_lo(self):
        self.v_lo.set("LibreOffice: listo" if buscar_soffice()
                      else "LibreOffice: no instalado (se instala al convertir)")

    def _log(self, texto):
        self.log.config(state="normal")
        inicio = texto.lstrip()
        tag = ("error" if "ERROR" in texto or "error:" in texto
               or inicio.startswith(("No se pudo", "No pude", "Error"))
               else "ok" if inicio.startswith(("OK", "FUSIONADO", "Listo")) else ())
        if tag == "error":
            self.hubo_error = True
        self.log.insert("end", texto + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def copiar_registro(self):
        self.clipboard_clear()
        self.clipboard_append(self.log.get("1.0", "end"))

    def _abrir(self, ruta):
        try:
            if ES_WINDOWS:
                os.startfile(str(ruta))
            elif platform.system() == "Darwin":
                subprocess.run(["open", str(ruta)])
            else:
                subprocess.run(["xdg-open", str(ruta)])
        except Exception as e:
            messagebox.showerror("No se pudo abrir", str(e))

    def _preparar_dnd(self):
        if not HAY_DND:
            self.v_dnd.set("")
            self.vacio.config(text="SIN EXPEDIENTES EN LA MESA.\n\nTocá [ + AGREGAR ARCHIVOS ] "
                                   "para empezar")
            if self.cfg.get("dnd_intentos", 0) < 2:
                self.cfg["dnd_intentos"] = self.cfg.get("dnd_intentos", 0) + 1
                guardar_config(self.cfg)
                threading.Thread(target=self._instalar_dnd, daemon=True).start()
            return
        try:
            self.tabla.drop_target_register(DND_FILES)
            self.tabla.dnd_bind("<<Drop>>", self._soltar)
            self.v_dnd.set("")
        except Exception:
            self.v_status.set("Arrastrar y soltar: no disponible (usá los botones)")

    def _instalar_dnd(self):
        try:
            exe = _consola(Path(sys.executable))
            r = subprocess.run([str(exe), "-m", "pip", "install", "--quiet",
                                "--disable-pip-version-check", "tkinterdnd2"],
                               capture_output=True, timeout=300, creationflags=NO_WINDOW)
            if r.returncode == 0:
                self.cola.put(("log", "Arrastrar y soltar instalado: se activa al reabrir el programa."))
        except Exception:
            pass

    def _soltar(self, evento):
        self.agregar_rutas(self.tk.splitlist(evento.data))

    def _soltar_herr(self, evento):
        self.herr_agregar_rutas(self.tk.splitlist(evento.data))

    def _salir(self):
        if self.vigilante:
            self.vigilante.parar.set()
        self.destroy()

    # ------------------------------------------------------------- selección (tab 1)
    def herr_agregar(self):
        rutas = filedialog.askopenfilenames(
            title="Elegí PDF", filetypes=[("PDF", "*.pdf"), ("Todos los archivos", "*.*")],
            initialdir=str(self.base) if self.base else str(Path.home()))
        self.herr_agregar_rutas(rutas)

    def herr_agregar_rutas(self, rutas):
        for r in rutas:
            p = Path(r)
            if p.suffix.lower() == ".pdf" and p not in self.herr_pdfs:
                self.herr_pdfs.append(p)
        self.herr_pdfs.sort(key=_clave_natural)
        self._herr_pintar()

    def _herr_sel(self):
        try:
            return self.herr_lista.curselection()
        except Exception:
            return ()

    def _herr_pintar(self):
        self.herr_lista.delete(0, "end")
        for p in self.herr_pdfs:
            self.herr_lista.insert("end", str(p))
        self.herr_lista.select_set(0, "end")

    def herr_quitar(self):
        quitar = set(self._herr_sel())
        self.herr_pdfs = [p for i, p in enumerate(self.herr_pdfs) if i not in quitar]
        self._herr_pintar()

    def herr_vaciar(self):
        self.herr_pdfs = []
        self._herr_pintar()

    def _herr_marcados(self):
        # Si no cargaron PDF acá, usar los PDF marcados de la otra pestaña
        sel = [self.herr_pdfs[i] for i in self._herr_sel()]
        if sel:
            return sel
        return [a for a in self._marcados() if a.suffix.lower() == ".pdf"]

    def herr_borrar(self):
        sel = self._herr_marcados()
        if len(sel) != 1:
            messagebox.showinfo("Elegí un solo PDF", "Marcá exactamente un PDF para ver sus páginas.")
            return
        DialogoPaginas(self, sel[0])

    def herr_dividir(self):
        sel = self._herr_marcados()
        if not sel:
            messagebox.showinfo("Nada marcado", "Agregá y marcá al menos un PDF.")
            return
        modo = self.v_modo_div.get()
        try:
            n = int(self.v_n.get())
            if modo == "cada_n" and n < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Dato no válido", "“Cada N páginas” tiene que ser un número mayor a 0.")
            return
        if modo == "rangos" and not self.v_rangos_div.get().strip():
            messagebox.showerror("Dato no válido", "Escribí los intervalos, por ej.: 1-3;4-7;8-")
            return
        self.lanzar_herramienta("dividir", sel, {"modo": modo, "n": n,
                                                 "rangos": self.v_rangos_div.get()})

    def herr_comprimir(self):
        sel = self._herr_marcados()
        if not sel:
            messagebox.showinfo("Nada marcado", "Agregá y marcá al menos un PDF.")
            return
        self.lanzar_herramienta("comprimir", sel, {"nivel": self.v_nivel.get()})

    # ------------------------------------------------------- carpetas automáticas
    def auto_elegir(self):
        c = filedialog.askdirectory(title="Carpeta a vigilar")
        if c:
            self.v_acc_carpeta.set(c)

    def _texto_regla(self, r):
        acc = {"comprimir": f"comprimir ({r.get('nivel', 'media')})",
               "dividir": f"dividir cada {r.get('n', 1)} págs.",
               "fusionar": "unir" + (" + blanco si impar" if r.get("blanco") else "")}[r["accion"]]
        return (f" {self._corta(r['carpeta'], 60)}   »   {acc}"
                + ("   [+Word/PPT]" if r.get("convertir") else ""))

    def _refrescar_reglas(self):
        self.lista_reglas.delete(0, "end")
        for r in self.reglas:
            self.lista_reglas.insert("end", self._texto_regla(r))

    def auto_agregar(self):
        carpeta = self.v_acc_carpeta.get()
        if not carpeta or not Path(carpeta).is_dir():
            messagebox.showinfo("Falta la carpeta", "Elegí la carpeta a vigilar.")
            return
        accion = self.v_acc.get()
        try:
            n = int(self.v_acc_n.get())
            if accion == "dividir" and n < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Dato no válido", "“Dividir cada” necesita un número mayor a 0.")
            return
        if any(r["carpeta"] == carpeta for r in self.reglas):
            messagebox.showinfo("Ya existe", "Esa carpeta ya tiene una regla. Quitala primero.")
            return
        self.reglas.append({"carpeta": carpeta, "accion": accion, "n": n,
                            "nivel": self.v_acc_nivel.get(), "blanco": self.v_acc_blanco.get(),
                            "convertir": self.v_acc_conv.get(), "pie": ""})
        self._refrescar_reglas()
        self._guardar_cfg()
        if self.vigilante:
            self.auto_alternar()
            self.auto_alternar()

    def auto_quitar(self):
        sel = list(self.lista_reglas.curselection())
        if sel:
            self.reglas = [r for i, r in enumerate(self.reglas) if i not in sel]
            self._refrescar_reglas()
            self._guardar_cfg()
            if self.vigilante:
                self.auto_alternar()
                self.auto_alternar()

    def _btn_vig(self, texto):
        try:
            self.btn_vig.config(text=texto)
        except Exception:
            pass

    def auto_alternar(self):
        if self.vigilante:
            self.vigilante.parar.set()
            self.vigilante = None
            self._btn_vig("[ INICIAR VIGILANCIA ]")
            self.v_vig.set("Vigilancia detenida")
            return
        if not self.reglas:
            messagebox.showinfo("Sin reglas", "Agregá al menos una regla.")
            return
        if any(r["accion"] in ("comprimir", "dividir", "fusionar") for r in self.reglas):
            # preparar PyMuPDF ahora, en segundo plano, para que no falle después
            threading.Thread(target=lambda: python_con_pymupdf(
                lambda t, d: self.cola.put((t, d))), daemon=True).start()
        self.vigilante = Vigilante(self.reglas, lambda t, d: self.cola.put((t, d)))
        self.vigilante.start()
        self._btn_vig("[ DETENER VIGILANCIA ]")
        self.v_vig.set(f"Vigilando {len(self.reglas)} carpeta(s)…")
        self._log(f"Vigilancia iniciada en {len(self.reglas)} carpeta(s).")


# ------------------------------------------------------------- fuentes incluidas
# VT323 (recortada a caracteres latinos) y Share Tech Mono (sin modificar), ambas bajo
# SIL Open Font License 1.1. Copyright 2011, The VT323 Project Authors. Copyright (c) 2012,
# Carrois Type Design, Ralph du Carrois, with Reserved Font Name 'Share'.
# Texto de la licencia: LICENSE-FUENTES.txt del repositorio.
FUENTES_B64 = {
    "VT323-Regular.ttf": """
AAEAAAAOAIAAAwBgR0RFRgIqAmEAALBoAAAALkdQT1NEdkx1AACwmAAAACBHU1VC1kzSgQAAsLgAAAH2T1MvMmawhvAAAKyMAAAA
YGNtYXCJuoi3AACs7AAAAeJnYXNwAAAAEAAAsGAAAAAIZ2x5Zq08VhsAAADsAACmFmhlYWQG8T6NAACo2AAAADZoaGVhAzcBhQAA
rGgAAAAkaG10eD5AIwEAAKkQAAADWGxvY2F5lE/UAACnJAAAAbJtYXhwAOoBoQAApwQAAAAgbmFtZRq+NIEAAK7QAAABcHBvc3T/
hwAzAACwQAAAACAAAgAkAAABbAIwAFMAZwAAJAcWFRQHIyY1NDcmNTQ3IxYVFAcWFRQHIyY1NDcmNTQ3MyY1NDcmNTQ3MyY1NDcz
JjU0NzMmNTQ3MxYVFAczFhUUBzMWFRQHMxYVFAcWFRQHMxYVJjU0NyMmNTQ3IxYVFAcjFhUUBzMBbAQEBFAEBAQEoAQEBARQBAQE
BB4EBAQEFAQEHgQEFAQEeAQEFAQEHgQEFAQEBAQeBHYEFAQEPAQEFAQEZGEREBgXEREXGBARFxgQEBgXERAYFxERFxgQERcYEBEX
GBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXERAYFxEQGIkXGBARFxgQEBgXERAYFxEAAwAkAAABbAMgABMAZwB7AAAT
JjU0NzMmNTQ3MxYVFAcjFhUUBxIHFhUUByMmNTQ3JjU0NyMWFRQHFhUUByMmNTQ3JjU0NzMmNTQ3JjU0NzMmNTQ3MyY1NDczJjU0
NzMWFRQHMxYVFAczFhUUBzMWFRQHFhUUBzMWFSY1NDcjJjU0NyMWFRQHIxYVFAczoAQEKAQEUAQEKAQEfAQEBFAEBAQEoAQEBARQ
BAQEBB4EBAQEFAQEHgQEFAQEeAQEFAQEHgQEFAQEBAQeBHYEFAQEPAQEFAQEZAKAERcYEBEXGBAQGBcREBgXEf3hERAYFxERFxgQ
ERcYEBAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcREBiJFxgQERcYEBAYFxEQ
GBcR//8AJAAAAWwDIAAiAAIAAAAHAM4BkACgAAQAJAAAAWwC0AAJABMAZwB7AAATJjU0NzMWFRQHMyY1NDczFhUUBxIHFhUUByMm
NTQ3JjU0NyMWFRQHFhUUByMmNTQ3JjU0NzMmNTQ3JjU0NzMmNTQ3MyY1NDczJjU0NzMWFRQHMxYVFAczFhUUBzMWFRQHFhUUBzMW
FSY1NDcjJjU0NyMWFRQHIxYVFAczUAQEUAQEUAQEUAQELAQEBFAEBAQEoAQEBARQBAQEBB4EBAQEFAQEHgQEFAQEeAQEFAQEHgQE
FAQEBAQeBHYEFAQEPAQEFAQEZAKAERcYEBAYFxERFxgQEBgXEf3hERAYFxERFxgQERcYEBAYFxEQGBcRERcYEBEXGBARFxgQERcY
EBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcREBiJFxgQERcYEBAYFxEQGBcRAAMAJAAAAWwDIAATAGcAewAAEyY1NDcz
FhUUBzMWFRQHIyY1NDcSBxYVFAcjJjU0NyY1NDcjFhUUBxYVFAcjJjU0NyY1NDczJjU0NyY1NDczJjU0NzMmNTQ3MyY1NDczFhUU
BzMWFRQHMxYVFAczFhUUBxYVFAczFhUmNTQ3IyY1NDcjFhUUByMWFRQHM3gEBFAEBCgEBFAEBMwEBARQBAQEBKAEBAQEUAQEBAQe
BAQEBBQEBB4EBBQEBHgEBBQEBB4EBBQEBAQEHgR2BBQEBDwEBBQEBGQC0BEXGBAQGBcREBgXEREXGBD9kREQGBcRERcYEBEXGBAQ
GBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXERAYiRcYEBEXGBAQGBcREBgXEf//
ACQAAAFsA3AAIgACAAAABwDPAZAAoP//ACQAAAFsAyAAIgACAAAABwDQAZAAoAACACQAAAFsAjAAUgBsAAAAByMWFRQHFhUUBzMW
FRQHIyY1NDcmNTQ3IxYVFAcWFRQHIyY1NDcmNTQ3JjU0NzMmNTQ3JjU0NzMmNTQ3MyY1NDchFhUUByMWFRQHFhUUBzMWFQcmNTQ3
JjU0NyY1NDcjFhUUBxYVFAcWFRQHAWwEUAQEBARQBASgBAQEBFAEBAQEUAQEBAQEBBQEBAQEFAQEFAQEAQQEBFAEBAQEUASkBAQE
BAQEPAQEBAQEBAEBERAYFxEQGBcREBgXEREXGBARFxgQEBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAY
FxEQGBcREBgoERcYEBEXGBARFxgQEBgXERAYFxEQGBcRAAADACQAAAFsAjAAPQBPAGEAAAAHIxYVFAczFhUUBxYVFAcjFhUUByEm
NTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3IRYVFAczFhUUBxYVBjU0NyY1NDcjFhUUBxYVFAczBjU0NyMWFRQHFhUUBzMmNTQ3
AWwEKAQEKAQEBAQoBAT+6AQEBAQEBAQEBAQEBAQEARgEBCgEBARYBAQEoAQEBASgBASgBAQEBKAEBAFRERAYFxEQGBcREBgXERAY
FxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcXGBARFxgQEBgXERAYFxGPFxgQEBgXERAYFxERFxgQAAAB
ACQAAAFsAjAAVQAAJRYVFAcjFhUUByMmNTQ3IyY1NDcjJjU0NyY1NDcmNTQ3MyY1NDczJjU0NzMWFRQHMxYVFAcjJjU0NyMWFRQH
IxYVFAcWFRQHFhUUBzMWFRQHMyY1NDcBaAQEKAQEyAQEKAQEKAQEBAQEBCgEBCgEBMgEBCgEBFAEBHgEBCgEBAQEBAQoBAR4BASg
EBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxERFxgQEBgXERAYFxEQGBcREBgXERAYFxERFxgQ//8A
JP9gAWwCMAAiAAsAAAADANIBkAAAAAIAJAAAAWwCMAA9AGkAAAAVFAcWFRQHIxYVFAcjFhUUByMmNTQ3JjU0NyY1NDcmNTQ3JjU0
NyY1NDcmNTQ3MxYVFAczFhUUBzMWFRQHBjU0NyY1NDcmNTQ3IyY1NDcjFhUUBxYVFAcWFRQHFhUUBxYVFAczJjU0NzMBbAQEBCgE
BCgEBPAEBAQEBAQEBAQEBAQEBPAEBCgEBCgEBFQEBAQEBCgEBHgEBAQEBAQEBAQEeAQEKAEwGBcREBgXERAYFxEQGBcRERcYEBEX
GBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXEY8XGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXEREXGBAA
AgAkAAABbAIwAEEAbwAAABUUBxYVFAcjFhUUByMWFRQHIyY1NDczJjU0NyY1NDcjJjU0NzMmNTQ3JjU0NyMmNTQ3MxYVFAczFhUU
BzMWFRQHBjU0NyY1NDcmNTQ3IyY1NDcjFhUUBxYVFAczFhUUByMWFRQHFhUUBzMmNTQ3MwFsBAQEKAQEKAQE8AQEKAQEBAQoBAQo
BAQEBCgEBPAEBCgEBCgEBFQEBAQEBCgEBFAEBAQEKAQEKAQEBARQBAQoATAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQ
ERcYEBEXGBARFxgQEBgXERAYFxEQGBcRjxcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcRERcYEAABAEwAAAFEAjAA
PQAANxYVFAcWFRQHMxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NzMWFRQHIxYVFAcWFRQHMxYVFAegBAQEBKAEBPAE
BAQEBAQEBAQEBAQEBPAEBKAEBAQEoAQE8BAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAY
FxEQGBcRAAACAEwAAAFEAyAAEwBRAAATJjU0NzMmNTQ3MxYVFAcjFhUUBwMWFRQHFhUUBzMWFRQHIyY1NDcmNTQ3JjU0NyY1NDcm
NTQ3JjU0NyY1NDczFhUUByMWFRQHFhUUBzMWFRQHoAQEKAQEUAQEKAQEUAQEBASgBATwBAQEBAQEBAQEBAQEBATwBASgBAQEBKAE
BAKAERcYEBEXGBAQGBcREBgXEf5wEBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAY
FxH//wBMAAABRAMgACIADwAAAAcAzgGQAKAAAwBMAAABRALQAAkAEwBRAAATJjU0NzMWFRQHMyY1NDczFhUUBwcWFRQHFhUUBzMW
FRQHIxYVFAcWFRQHMxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NzMWFRQHUAQEUAQEUAQEUAQEoAQEBASgBASgBAQE
BKAEBPAEBAQEBAQEBAQEBAQEBPAEBAKAERcYEBAYFxERFxgQEBgXEaAQGBcREBgXERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEX
GBARFxgQERcYEBEXGBARFxgQEBgXEQAAAgBMAAABRAMgABMAUQAAEyY1NDczFhUUBzMWFRQHIyY1NDcRFhUUBxYVFAczFhUUByMm
NTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3MxYVFAcjFhUUBxYVFAczFhUUB3gEBFAEBCgEBFAEBAQEBASgBATwBAQEBAQEBAQE
BAQEBATwBASgBAQEBKAEBALQERcYEBAYFxEQGBcRERcYEP4gEBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcY
EBAYFxEQGBcREBgXERAYFxEAAAEATAAAAUQCMAA8AAAAFRQHIxYVFAcWFRQHFhUUByMmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDcm
NTQ3MxYVFAcjFhUUBxYVFAczAUQEoAQEBAQEBFAEBAQEBAQEBAQEBAQEBPAEBKAEBAQEoAEwGBcREBgXERAYFxEQGBcRERcYEBEX
GBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXEQABACQAAAFsAjAAZgAAABUUBxYVFAcWFRQHIxYVFAcjJjU0NyMmNTQ3
IyY1NDcmNTQ3JjU0NzMmNTQ3MyY1NDczFhUUBzMWFRQHIyY1NDcjFhUUByMWFRQHFhUUBxYVFAczFhUUBzMmNTQ3JjU0NyMmNTQ3
MwFsBAQEBAQoBATIBAQoBAQoBAQEBAQEKAQEKAQEyAQEKAQEUAQEeAQEKAQEBAQEBCgEBHgEBAQEPAQEjAEwGBcREBgXERAYFxEQ
GBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcRERcYEBAYFxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQ
AAEAJAAAAWwCMABtAAAAFRQHFhUUBxYVFAcWFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3IxYVFAcWFRQHFhUUByMmNTQ3JjU0NyY1
NDcmNTQ3JjU0NyY1NDcmNTQ3MxYVFAcWFRQHFhUUBzMmNTQ3JjU0NyY1NDczFhUUBwFsBAQEBAQEBAQEBARQBAQEBAQEoAQEBAQE
BFAEBAQEBAQEBAQEBAQEBFAEBAQEBASgBAQEBAQEUAQEAdAYFxEQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBAYFxEQ
GBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxERFxgQERcYEBEXGBAQGBcRAAEAYAAAATACMAA9
AAAlFhUUByMmNTQ3MyY1NDcmNTQ3JjU0NyY1NDcmNTQ3IyY1NDczFhUUByMWFRQHFhUUBxYVFAcWFRQHFhUUBwEsBATIBAQ8BAQE
BAQEBAQEBDwEBMgEBDwEBAQEBAQEBAQEUBAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXERAY
FxEQGBcRAAIAYAAAATADIAATAFEAABMmNTQ3MyY1NDczFhUUByMWFRQHExYVFAcjJjU0NzMmNTQ3JjU0NyY1NDcmNTQ3JjU0NyMm
NTQ3MxYVFAcjFhUUBxYVFAcWFRQHFhUUBxYVFAegBAQoBARQBAQoBAQ8BATIBAQ8BAQEBAQEBAQEBDwEBMgEBDwEBAQEBAQEBAQE
AoARFxgQERcYEBAYFxEQGBcR/dAQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcREBgX
Ef//AFYAAAE6AyAAIgAXAAAABwDOAZAAoAADAEwAAAFEAtAACQATAFEAABIHIyY1NDczFhU2FRQHIyY1NDczBgcjFhUUBxYVFAcW
FRQHFhUUBxYVFAczFhUUByMmNTQ3MyY1NDcmNTQ3JjU0NyY1NDcmNTQ3IyY1NDczFhWkBFAEBFAEoARQBARQEAQ8BAQEBAQEBAQE
BDwEBMgEBDwEBAQEBAQEBAQEPAQEyAQCkRERFxgQEBgYGBcRERcYEN8REBgXERAYFxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBAR
FxgQERcYEBEXGBARFxgQERcYEBAYAAACAGAAAAEwAyAAEwBRAAATJjU0NzMWFRQHMxYVFAcjJjU0NxMWFRQHIyY1NDczJjU0NyY1
NDcmNTQ3JjU0NyY1NDcjJjU0NzMWFRQHIxYVFAcWFRQHFhUUBxYVFAcWFRQHeAQEUAQEKAQEUAQEjAQEyAQEPAQEBAQEBAQEBAQ8
BATIBAQ8BAQEBAQEBAQEBALQERcYEBAYFxEQGBcRERcYEP2AEBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgX
ERAYFxEQGBcREBgXERAYFxEAAQAkAAABWAIwAEcAAAAHIxYVFAcWFRQHFhUUBxYVFAcWFRQHIxYVFAcjJjU0NyMmNTQ3MxYVFAcz
JjU0NyY1NDcmNTQ3JjU0NyY1NDcjJjU0NzMWFQFYBDwEBAQEBAQEBAQEKAQEoAQEKAQEUAQEUAQEBAQEBAQEBAQ8BATIBAHxERAY
FxEQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQEBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBAYAAABACQAAAFsAjAAdwAA
JRYVFAcjJjU0NyMmNTQ3IyY1NDcjFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDczFhUUBxYVFAcWFRQH
MyY1NDczJjU0NzMmNTQ3MxYVFAcjFhUUByMWFRQHIxYVFAczFhUUBzMWFRQHAWgEBFAEBB4EBDwEBEYEBAQEBARQBAQEBAQEBAQE
BAQEBARQBAQEBAQERgQEPAQEHgQEUAQEHgQEPAQEMgQEMgQEPAQEUBAYFxERFxgQERcYEBEXGBAQGBcREBgXERAYFxERFxgQERcY
EBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcRERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXERAYFxEQGBcRAAEATAAA
AWwCMAA6AAAkFRQHISY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDczFhUUBxYVFAcWFRQHFhUUBxYVFAcWFRQHMwFsBP7oBAQE
BAQEBAQEBAQEBARQBAQEBAQEBAQEBAQEyEAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXERAY
FxEQGBcRAAEAJAAAAWwCMAClAAAAFRQHFhUUBxYVFAcWFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyY1NDcjFhUUBxYVFAcj
FhUUBxYVFAcjJjU0NyY1NDcjJjU0NyY1NDcjFhUUBxYVFAcWFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0
NzMWFRQHMxYVFAczFhUUBzMmNTQ3MyY1NDczJjU0NzMWFRQHAWwEBAQEBAQEBAQEBFAEBAQEBAQEBAQEHgQEBAQUBAQEBDwEBAQE
FAQEBAQeBAQEBAQEBAQEBFAEBAQEBAQEBAQEBAQEBGQEBB4EBA8EBB4EBA8EBB4EBGQEBAHQGBcREBgXERAYFxEQGBcREBgXERAY
FxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXERAYFxER
FxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcRERcYEBEXGBARFxgQEBgXEQABACQAAAFsAjAAgwAAABUUBxYV
FAcWFRQHFhUUBxYVFAcWFRQHIyY1NDcjJjU0NyMmNTQ3IyY1NDcjFhUUBxYVFAcWFRQHFhUUByMmNTQ3JjU0NyY1NDcmNTQ3JjU0
NyY1NDcmNTQ3MxYVFAczFhUUBzMWFRQHMxYVFAczJjU0NyY1NDcmNTQ3JjU0NzMWFRQHAWwEBAQEBAQEBAQEBFAEBDIEBB4EBCgE
BCgEBAQEBAQEBFAEBAQEBAQEBAQEBAQEBGQEBB4EBB4EBCgEBCgEBAQEBAQEBFAEBAHQGBcREBgXERAYFxEQGBcREBgXERAYFxER
FxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcR
ERcYEBEXGBARFxgQERcYEBAYFxH//wAkAAABbAMgACIAIAAAAAcA0AGQAKAAAgAkAAABbAIwAEEAbwAAABUUBxYVFAcjFhUUByMW
FRQHIyY1NDcjJjU0NyMmNTQ3JjU0NyY1NDczJjU0NzMmNTQ3MxYVFAczFhUUBzMWFRQHBjU0NyY1NDcmNTQ3IyY1NDcjFhUUByMW
FRQHFhUUBxYVFAczFhUUBzMmNTQ3MwFsBAQEHgQEMgQEoAQEMgQEHgQEBAQEBB4EBDIEBKAEBDIEBB4EBFQEBAQEBB4EBGQEBB4E
BAQEBAQeBARkBAQeATAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcRjxcYEBEX
GBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcRERcYEAADACQAAAFsAyAAEwBVAIMAABMmNTQ3MyY1NDczFhUUByMWFRQHEhUU
BxYVFAcjFhUUByMWFRQHIyY1NDcjJjU0NyMmNTQ3JjU0NyY1NDczJjU0NzMmNTQ3MxYVFAczFhUUBzMWFRQHBjU0NyY1NDcmNTQ3
IyY1NDcjFhUUByMWFRQHFhUUBxYVFAczFhUUBzMmNTQ3M6AEBCgEBFAEBCgEBHwEBAQeBAQyBASgBAQyBAQeBAQEBAQEHgQEMgQE
oAQEMgQEHgQEVAQEBAQEHgQEZAQEHgQEBAQEBB4EBGQEBB4CgBEXGBARFxgQEBgXERAYFxH+sBgXERAYFxEQGBcREBgXEREXGBAR
FxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxGPFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXERAYFxERFxgQ
AP//ACQAAAFsAyAAIgAiAAAABwDOAZAAoAAEACQAAAFsAtAACQATAFUAgwAAEyY1NDczFhUUBzMmNTQ3MxYVFAcSFRQHFhUUByMW
FRQHIxYVFAcjJjU0NyMmNTQ3IyY1NDcmNTQ3JjU0NzMmNTQ3MyY1NDczFhUUBzMWFRQHMxYVFAcGNTQ3JjU0NyY1NDcjJjU0NyMW
FRQHIxYVFAcWFRQHFhUUBzMWFRQHMyY1NDczUAQEUAQEUAQEUAQELAQEBB4EBDIEBKAEBDIEBB4EBAQEBAQeBAQyBASgBAQyBAQe
BARUBAQEBAQeBARkBAQeBAQEBAQEHgQEZAQEHgKAERcYEBAYFxERFxgQEBgXEf6wGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQ
ERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXEY8XGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXEREXGBAAAAMAJAAA
AWwDIAATAFUAgwAAEyY1NDczFhUUBzMWFRQHIyY1NDcSFRQHFhUUByMWFRQHIxYVFAcjJjU0NyMmNTQ3IyY1NDcmNTQ3JjU0NzMm
NTQ3MyY1NDczFhUUBzMWFRQHMxYVFAcGNTQ3JjU0NyY1NDcjJjU0NyMWFRQHIxYVFAcWFRQHFhUUBzMWFRQHMyY1NDczeAQEUAQE
KAQEUAQEzAQEBB4EBDIEBKAEBDIEBB4EBAQEBAQeBAQyBASgBAQyBAQeBARUBAQEBAQeBARkBAQeBAQEBAQEHgQEZAQEHgLQERcY
EBAYFxEQGBcRERcYEP5gGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXEY8XGBAR
FxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXEREXGBAAAAMAJP+wAWwCgABTAHkAnwAAABUUBxYVFAcWFRQHFhUUByMWFRQH
IxYVFAcjFhUUByMmNTQ3MyY1NDcjJjU0NyY1NDcmNTQ3JjU0NzMmNTQ3MyY1NDczJjU0NzMWFRQHIxYVFAczAhUUBzMmNTQ3MyY1
NDczJjU0NzMmNTQ3IxYVFAcjFhUUBxYVFAc2NTQ3IxYVFAcjFhUUByMWFRQHIxYVFAczJjU0NzMmNTQ3JjU0NwFsBAQEBAQEBB4E
BDIEBKAEBFAEBB4EBB4EBAQEBAQEBB4EBDIEBKAEBFAEBB4EBB7sBCMEBBQEBBQEBCMEBFAEBB4EBAQEnAQjBAQUBAQUBAQjBARQ
BAQeBAQEBAHQGBcREBgXERAYFxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgX
ERAYFxH/ABgXEREXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxFhFxgQEBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEAD//wAk
AAABbAMgACIAIgAAAAcA0AGQAKAAAgAkAAABbAIwADwATgAAAAcWFRQHIxYVFAcjFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyY1
NDcmNTQ3JjU0NyY1NDchFhUUBzMWFQY1NDcmNTQ3IxYVFAcWFRQHMwFsBAQEKAQEyAQEBAQEBFAEBAQEBAQEBAQEBAQEBAEYBAQo
BFgEBASgBAQEBKABoREQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBhnFxgQERcY
EBAYFxEQGBcRAAIAJAAAAWwCMABBAF0AAAAHIxYVFAcjFhUUByMWFRQHIyY1NDczJjU0NyY1NDcmNTQ3JjU0NyY1NDcjJjU0NzMW
FRQHMxYVFAczFhUUBzMWFQcmNTQ3IyY1NDcjFhUUBxYVFAcWFRQHMyY1NDcBbAQeBAQyBAR4BAR4BAQoBAQEBAQEBAQEBCgEBHgE
BHgEBDIEBB4EVAQEHgQEWgQEBAQEBFoEBAEBERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgX
ERAYFxEQGCgRFxgQERcYEBAYFxEQGBcREBgXEREXGBAAAgAk/2ABbAIwAFUAgwAAABUUBxYVFAcjFhUUByMWFRQHMxYVFAczFhUU
ByMmNTQ3IyY1NDcjJjU0NyMmNTQ3IyY1NDcmNTQ3JjU0NzMmNTQ3MyY1NDczFhUUBzMWFRQHMxYVFAcGNTQ3JjU0NyY1NDcjJjU0
NyMWFRQHIxYVFAcWFRQHFhUUBzMWFRQHMyY1NDczAWwEBAQeBAQyBAQoBAQoBARQBAQoBAR4BAQyBAQeBAQEBAQEHgQEMgQEoAQE
MgQEHgQEVAQEBAQEHgQEZAQEHgQEBAQEBB4EBGQEBB4BMBgXERAYFxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEX
GBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcRjxcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcRERcYEAACACQA
AAFsAjAAWwBtAAAlFhUUByMmNTQ3IyY1NDcjJjU0NyMWFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NyEW
FRQHMxYVFAcjFhUUByMWFRQHMxYVFAczFhUUBwI1NDcmNTQ3IxYVFAcWFRQHMwFoBARQBAQeBAQeBARkBAQEBAQEUAQEBAQEBAQE
BAQEBAQEARgEBCgEBBQEBDwEBBQEBB4EBEoEBASMBAQEBIxQEBgXEREXGBARFxgQERcYEBAYFxEQGBcREBgXEREXGBARFxgQERcY
EBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXERAYFxEBARcYEBEXGBAQGBcREBgXEQABACQAAAFsAjAAVQAAJBUU
BxYVFAcjFhUUByMmNTQ3IyY1NDczFhUUBzMmNTQ3JjU0NyMmNTQ3IyY1NDcmNTQ3MyY1NDczFhUUBzMWFRQHIyY1NDcjFhUUBxYV
FAczFhUUBzMBbAQEBCgEBPAEBCgEBFAEBKAEBAQEyAQEKAQEBAQoBATwBAQoBARQBASgBAQEBMgEBCjgGBcREBgXERAYFxERFxgQ
ERcYEBAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXEREXGBAQGBcREBgXERAYFxEAAAEAJAAAAWwCMAA7AAAAByMW
FRQHFhUUBxYVFAcWFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3IyY1NDchFhUBbAR4BAQEBAQEBAQEBAQEUAQE
BAQEBAQEBAQEBHgEBAFABAHxERAYFxEQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGAAB
ACQAAAFsAjAAbQAAABUUBxYVFAcWFRQHFhUUBxYVFAcjFhUUByMmNTQ3IyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NzMWFRQHFhUU
BxYVFAcWFRQHFhUUBxYVFAczJjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3MxYVFAcBbAQEBAQEBAQEBCgEBPAEBCgEBAQEBAQEBAQE
BARQBAQEBAQEBAQEBAQEoAQEBAQEBAQEBAQEBFAEBAHQGBcREBgXERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcY
EBEXGBARFxgQEBgXERAYFxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXEQACACQAAAFsAyAAEwCB
AAATJjU0NzMmNTQ3MxYVFAcjFhUUBxYVFAcWFRQHFhUUBxYVFAcWFRQHIxYVFAcjJjU0NyMmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1
NDczFhUUBxYVFAcWFRQHFhUUBxYVFAcWFRQHMyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NzMWFRQHoAQEKAQEUAQEKAQEfAQEBAQE
BAQEBCgEBPAEBCgEBAQEBAQEBAQEBARQBAQEBAQEBAQEBAQEoAQEBAQEBAQEBAQEBFAEBAKAERcYEBEXGBAQGBcREBgXEbAYFxEQ
GBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXERAYFxERFxgQ
ERcYEBEXGBARFxgQERcYEBEXGBAQGBcR//8AJAAAAWwDIAAiAC8AAAAHAM4BkACgAAMAJAAAAWwC0AAJABMAgQAAEyY1NDczFhUU
BzMmNTQ3MxYVFAcWFRQHFhUUBxYVFAcWFRQHFhUUByMWFRQHIyY1NDcjJjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3MxYVFAcWFRQH
FhUUBxYVFAcWFRQHFhUUBzMmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDczFhUUB1AEBFAEBFAEBFAEBCwEBAQEBAQEBAQoBATwBAQo
BAQEBAQEBAQEBAQEUAQEBAQEBAQEBAQEBKAEBAQEBAQEBAQEBARQBAQCgBEXGBAQGBcRERcYEBAYFxGwGBcREBgXERAYFxEQGBcR
EBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcY
EBEXGBARFxgQEBgXEQACACQAAAFsAyAAEwCBAAATJjU0NzMWFRQHMxYVFAcjJjU0NxIVFAcWFRQHFhUUBxYVFAcWFRQHIxYVFAcj
JjU0NyMmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDczFhUUBxYVFAcWFRQHFhUUBxYVFAcWFRQHMyY1NDcmNTQ3JjU0NyY1NDcmNTQ3
JjU0NzMWFRQHeAQEUAQEKAQEUAQEzAQEBAQEBAQEBCgEBPAEBCgEBAQEBAQEBAQEBARQBAQEBAQEBAQEBAQEoAQEBAQEBAQEBAQE
BFAEBALQERcYEBAYFxEQGBcRERcYEP8AGBcREBgXERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQ
EBgXERAYFxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXEQAAAQAkAAABbAIwAG8AAAAVFAcWFRQH
IxYVFAcjFhUUByMWFRQHIxYVFAcjJjU0NyMmNTQ3IyY1NDcjJjU0NyMmNTQ3JjU0NyY1NDczFhUUBxYVFAcWFRQHMxYVFAczFhUU
BzMmNTQ3MyY1NDczJjU0NyY1NDcmNTQ3MxYVFAcBbAQEBB4EBB4EBBQEBCgEBFAEBCgEBBQEBB4EBB4EBAQEBARQBAQEBAQEHgQE
HgQEKAQEHgQEHgQEBAQEBFAEBAHQGBcREBgXERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgX
ERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEAAQAkAAABbAIwAKcAAAAVFAcWFRQHFhUUByMWFRQHFhUU
ByMWFRQHIyY1NDcjJjU0NyMWFRQHIxYVFAcjJjU0NyMmNTQ3JjU0NyMmNTQ3JjU0NyY1NDcmNTQ3MxYVFAcWFRQHFhUUBxYVFAcW
FRQHMyY1NDczJjU0NyY1NDcmNTQ3JjU0NzMWFRQHFhUUBxYVFAcWFRQHMxYVFAczJjU0NyY1NDcmNTQ3JjU0NyY1NDczFhUUBwFs
BAQEBAQeBAQEBBQEBDwEBBQEBDwEBBQEBDwEBBQEBAQEHgQEBAQEBAQEUAQEBAQEBAQEBAQeBAQUBAQEBAQEBAQ8BAQEBAQEBAQU
BAQeBAQEBAQEBAQEBFAEBAHQGBcREBgXERAYFxEQGBcREBgXERAYFxERFxgQERcYEBAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEX
GBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXEREXGBAR
FxgQERcYEBEXGBARFxgQEBgXEQABACQAAAFsAjAAgQAAJRYVFAcjJjU0NyMmNTQ3IyY1NDcjFhUUByMWFRQHIxYVFAcjJjU0NzMm
NTQ3MyY1NDczJjU0NyMmNTQ3IyY1NDcjJjU0NzMWFRQHMxYVFAczFhUUBzMmNTQ3MyY1NDczJjU0NzMWFRQHIxYVFAcjFhUUByMW
FRQHMxYVFAczFhUUBwFoBARQBAQeBAQUBAQ8BAQUBAQeBARQBAQeBAQoBAQeBAQeBAQoBAQeBARQBAQeBAQUBAQ8BAQUBAQeBARQ
BAQeBAQoBAQeBAQeBAQoBARQEBgXEREXGBARFxgQERcYEBAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQ
GBcREBgXERAYFxERFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXERAYFxEAAQAkAAABbAIwAFsAAAAHFhUUByMWFRQHIxYV
FAcjFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyMmNTQ3IyY1NDcjJjU0NyY1NDczFhUUBxYVFAczFhUUBzMmNTQ3MyY1NDcmNTQ3
MxYVAWwEBAQoBAQoBAQoBAQEBAQEUAQEBAQEBCgEBCgEBCgEBAQEUAQEBAQoBARQBAQoBAQEBFAEAfEREBgXERAYFxEQGBcREBgX
ERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXEREXGBARFxgQERcYEBAYAAACACQAAAFsAyAA
EwBvAAATJjU0NzMmNTQ3MxYVFAcjFhUUBxYHFhUUByMWFRQHIxYVFAcjFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyMmNTQ3IyY1
NDcjJjU0NyY1NDczFhUUBxYVFAczFhUUBzMmNTQ3MyY1NDcmNTQ3MxYVoAQEKAQEUAQEKAQEfAQEBCgEBCgEBCgEBAQEBARQBAQE
BAQEKAQEKAQEKAQEBARQBAQEBCgEBFAEBCgEBAQEUAQCgBEXGBARFxgQEBgXERAYFxGPERAYFxEQGBcREBgXERAYFxEQGBcREBgX
EREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxERFxgQERcYEBEXGBAQGAAAAQA4AAABWAIwAEMAAAAHIxYV
FAcjFhUUByMWFRQHIxYVFAczFhUUByEmNTQ3JjU0NzMmNTQ3MyY1NDczJjU0NzMmNTQ3IyY1NDchFhUUBxYVAVgEHgQEPAQEPAQE
MgQEyAQE/ugEBAQEMgQEPAQEPAQEHgQEyAQEARgEBAQBoREQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQ
ERcYEBEXGBAQGBcREBgAAAEAJAAAAWwBkABDAAAkFRQHIyY1NDcjJjU0NyMWFRQHMxYVFAcjJjU0NyMmNTQ3MyY1NDczJjU0NyMm
NTQ3MxYVFAczFhUUBxYVFAcWFRQHMwFsBFAEBCgEBHgEBHgEBKAEBCgEBCgEBKAEBIwEBLQEBCgEBAQEBAQoQBgXEREXGBARFxgQ
EBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEAAAIAJAAAAWwCgAATAFcAABMmNTQ3MyY1NDczFhUU
ByMWFRQHEhUUByMmNTQ3IyY1NDcjFhUUBzMWFRQHIyY1NDcjJjU0NzMmNTQ3MyY1NDcjJjU0NzMWFRQHMxYVFAcWFRQHFhUUBzOg
BAQoBARQBAQoBAR8BFAEBCgEBHgEBHgEBKAEBCgEBCgEBKAEBIwEBLQEBCgEBAQEBAQoAeARFxgQERcYEBAYFxEQGBcR/mAYFxER
FxgQERcYEBAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcRAP//ACQAAAFsAoAAIgA6AAAAAwDOAZAA
AAADACQAAAFsAjAACQATAFcAABMmNTQ3MxYVFAczJjU0NzMWFRQHEhUUByMmNTQ3IyY1NDcjFhUUBzMWFRQHIyY1NDcjJjU0NzMm
NTQ3MyY1NDcjJjU0NzMWFRQHMxYVFAcWFRQHFhUUBzNQBARQBARQBARQBAQsBFAEBCgEBHgEBHgEBKAEBCgEBCgEBKAEBIwEBLQE
BCgEBAQEBAQoAeARFxgQEBgXEREXGBAQGBcR/mAYFxERFxgQERcYEBAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgX
ERAYFxEQGBcRAAACACQAAAFsAoAAEwBXAAATJjU0NzMWFRQHMxYVFAcjJjU0NxIVFAcjJjU0NyMmNTQ3IxYVFAczFhUUByMmNTQ3
IyY1NDczJjU0NzMmNTQ3IyY1NDczFhUUBzMWFRQHFhUUBxYVFAczeAQEUAQEKAQEUAQEzARQBAQoBAR4BAR4BASgBAQoBAQoBASg
BASMBAS0BAQoBAQEBAQEKAIwERcYEBAYFxEQGBcRERcYEP4QGBcRERcYEBEXGBAQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQ
EBgXERAYFxEQGBcREBgXEQD//wAkAAABbALQACIAOgAAAAMAzwGQAAD//wAkAAABbAKAACIAOgAAAAMA0AGQAAAAAwAkAAABbAGQ
AEQATgBYAAAkByMWFRQHMxYVFAcjJjU0NyMWFRQHIyY1NDcjJjU0NzMmNTQ3MyY1NDcjJjU0NzMWFRQHMyY1NDczFhUUBzMWFRQH
FhUnJjU0NyMWFRQHByY1NDcjFhUUBwFsBGQEBGQEBHgEBCgEBHgEBCgEBCgEBGQEBFAEBGQEBCgEBGQEBBQEBAQsBAQ8BARQBAQ8
BASxERAYFxEQGBcRERcYEBAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxERFxgQEBgXERAYFxEQGCgRFxgQEBgXEaARFxgQEBgX
EQADACQAAAFsAjAAPABkAG4AADYVFAcjFhUUByMmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3MxYVFAcWFRQHFhUUBzMWFRQH
IxYVFAc2ByMWFRQHIyY1NDczJjU0NyMmNTQ3IyY1NDczFhUUBzMWFRQHMxYVBjczFhUUByMmNXwEKAQEKAQEBAQEBAQEBAQEBAQE
UAQEBAQEBCgEBCgEBPQEKAQEUAQEKAQEKAQEUAQEeAQEKAQEKAT4BHgEBHgEkBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEX
GBARFxgQEBgXERAYFxEQGBcREBgXERAYFxERERAYFxERFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGIgQEBgXEREXAAEAJAAAAWwB
kABFAAAlFhUUByMWFRQHIyY1NDcjJjU0NyMmNTQ3MyY1NDczJjU0NzMWFRQHMxYVFAcjJjU0NyMWFRQHIxYVFAczFhUUBzMmNTQ3
AWgEBCgEBMgEBCgEBCgEBCgEBCgEBMgEBCgEBFAEBHgEBCgEBCgEBHgEBKAQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgX
ERAYFxERFxgQEBgXERAYFxEQGBcRERcYEP//ACT/YAFsAZAAIgBDAAAAAwDSAZAAAAACACQAAAFsAjAAPQBvAAAAFRQHFhUUBxYV
FAcWFRQHFhUUBxYVFAcjJjU0NyMmNTQ3MyY1NDcjJjU0NzMmNTQ3JjU0NyY1NDczFhUUBwYHIxYVFAcjFhUUBzMWFRQHMxYVFAcj
JjU0NyMmNTQ3IyY1NDczJjU0NzMmNTQ3MxYVAWwEBAQEBAQEBAQEBFAEBB4EBB4EBB4EBB4EBAQEBARQBAR+BGQEBB4EBB4EBGQE
BG4EBDIEBB4EBB4EBDIEBG4EAdAYFxEQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcR
jxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGAAAAwAkAAABbALQAFYAYAB8AAAAFRQHFhUUBxYVFAcjFhUU
ByMmNTQ3IyY1NDcjJjU0NzMmNTQ3MyY1NDczJjU0NzMmNTQ3IyY1NDczJjU0NzMWFRQHIxYVFAczFhUUBzMWFRQHMxYVFAcmFRQH
MyY1NDcjEjU0NyY1NDcjFhUUByMWFRQHMxYVFAczJjU0NwFsBAQEBAQoBATIBAQyBAQeBAQeBAQyBAQeBAQoBARGBASCBAQ8BAQe
BAQeBAQeBAQUBASSBEYEBEZCBAQEggQEHgQEHgQEggQEATAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBAR
FxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXEZAYFxERFxgQ/tEXGBARFxgQEBgXERAYFxEQGBcRERcYEAAAAgAkAAABbAGQ
ADAAOgAAAAcWFRQHIxYVFAczFhUUByMmNTQ3IyY1NDcjJjU0NzMmNTQ3MyY1NDczFhUUBzMWFQY1NDcjFhUUBzMBbAQEBNIEBKoE
BMgEBDIEBB4EBB4EBDIEBMgEBCgEWASCBASCAQEREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXFxgQEBgX
EQAAAwAkAAABbAKAABMARABOAAATJjU0NzMmNTQ3MxYVFAcjFhUUBxYHFhUUByMWFRQHMxYVFAcjJjU0NyMmNTQ3IyY1NDczJjU0
NzMmNTQ3MxYVFAczFhUGNTQ3IxYVFAcztAQEKAQEUAQEKAQEaAQEBNIEBKoEBMgEBDIEBB4EBB4EBDIEBMgEBCgEWASCBASCAeAR
FxgQERcYEBAYFxEQGBcR3xEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcXGBAQGBcRAP//ACQAAAFsAoAA
IgBHAAAAAwDOAaQAAAAEACQAAAFsAjAACQATAEQATgAAEyY1NDczFhUUBzMmNTQ3MxYVFAcWBxYVFAcjFhUUBzMWFRQHIyY1NDcj
JjU0NyMmNTQ3MyY1NDczJjU0NzMWFRQHMxYVBjU0NyMWFRQHM2QEBFAEBFAEBFAEBBgEBATSBASqBATIBAQyBAQeBAQeBAQyBATI
BAQoBFgEggQEggHgERcYEBAYFxERFxgQEBgXEd8REBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXFxgQEBgX
EQAAAwAkAAABbAKAABMARABOAAATJjU0NzMWFRQHMxYVFAcjJjU0NxIHFhUUByMWFRQHMxYVFAcjJjU0NyMmNTQ3IyY1NDczJjU0
NzMmNTQ3MxYVFAczFhUGNTQ3IxYVFAczjAQEUAQEKAQEUAQEuAQEBNIEBKoEBMgEBDIEBB4EBB4EBDIEBMgEBCgEWASCBASCAjAR
FxgQEBgXERAYFxERFxgQ/tEREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXFxgQEBgXEQABAEwAAAFEAjAA
QQAAARYVFAcjFhUUBxYVFAcWFRQHMxYVFAcjJjU0NzMmNTQ3JjU0NyY1NDcjJjU0NzMmNTQ3MyY1NDczFhUUByMWFRQHAUAEBHgE
BAQEBARQBATIBAQoBAQEBAQEKAQEKAQEKAQEoAQEeAQEAZAQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQ
ERcYEBEXGBAQGBcREBgXEQAAAQAk/2ABbAGQAGgAAAAVFAcWFRQHFhUUBxYVFAcjFhUUByMWFRQHIyY1NDczJjU0NzMmNTQ3IyY1
NDczJjU0NyY1NDcjFhUUByMWFRQHMxYVFAczFhUUByMmNTQ3IyY1NDcjJjU0NzMmNTQ3MyY1NDczFhUUBwFsBAQEBAQEBCgEBCgE
BKAEBHgEBCgEBCgEBCgEBAQEggQEHgQEHgQEWgQEeAQEMgQEHgQEHgQEMgQE8AQEATAYFxEQGBcREBgXERAYFxEQGBcREBgXEREX
GBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEAAgAkAAABbAIwADsA
ZwAAEgczFhUUByMWFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NzMWFRQHFhUUBxYVFhUUBxYVFAcWFRQH
IyY1NDcmNTQ3JjU0NyY1NDcjJjU0NzMWFRQHMxYVFAd8BB4EBB4EBAQEBARQBAQEBAQEBAQEBAQEBARQBAQEBATwBAQEBARQBAQE
BAQEBASCBASgBAQyBAQBUREQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBiIGBcR
EBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcRAAIATAAAAUQCMAAJADYAABMmNTQ3MxYVFAcSFRQHIyY1NDczJjU0
NyY1NDcmNTQ3IyY1NDczFhUUBxYVFAcWFRQHFhUUBzOgBARQBARUBPAEBFAEBAQEBAQ8BASMBAQEBAQEBARQAeARFxgQEBgXEf5g
GBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcRAAABAEwAAAFEAZAALAAAJBUUByMmNTQ3MyY1NDcmNTQ3JjU0
NyMmNTQ3MxYVFAcWFRQHFhUUBxYVFAczAUQE8AQEUAQEBAQEBDwEBIwEBAQEBAQEBFBAGBcRERcYEBEXGBARFxgQERcYEBEXGBAQ
GBcREBgXERAYFxEQGBcRAAACAEwAAAFEAoAAEwBAAAATJjU0NzMmNTQ3MxYVFAcjFhUUBxIVFAcjJjU0NzMmNTQ3JjU0NyY1NDcj
JjU0NzMWFRQHFhUUBxYVFAcWFRQHM6AEBCgEBFAEBCgEBFQE8AQEUAQEBAQEBDwEBIwEBAQEBAQEBFAB4BEXGBARFxgQEBgXERAY
FxH+YBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXEQD//wBMAAABRAKAACIAUAAAAAMAzgGQAAAAAwBMAAAB
RAIwAAkAEwBAAAATJjU0NzMWFRQHMyY1NDczFhUUBxEWFRQHIyY1NDczJjU0NyY1NDcmNTQ3IyY1NDczFhUUBxYVFAcWFRQHFhUU
B1AEBFAEBFAEBFAEBAQE8AQEUAQEBAQEBDwEBIwEBAQEBAQEBAHgERcYEBAYFxERFxgQEBgXEf5wEBgXEREXGBARFxgQERcYEBEX
GBARFxgQEBgXERAYFxEQGBcREBgXEQD//wBMAAABRAIwACIAUAAAAAMAywGQAAAAAgBMAAABRAKAABMAQAAAEyY1NDczFhUUBzMW
FRQHIyY1NDcSFRQHIyY1NDczJjU0NyY1NDcmNTQ3IyY1NDczFhUUBxYVFAcWFRQHFhUUBzN4BARQBAQoBARQBASkBPAEBFAEBAQE
BAQ8BASMBAQEBAQEBARQAjARFxgQEBgXERAYFxERFxgQ/hAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEA
AAIATP9gARwCMAAJAEgAABMmNTQ3MxYVFAcWFRQHFhUUBxYVFAcWFRQHFhUUByMWFRQHIxYVFAcjJjU0NzMmNTQ3MyY1NDcmNTQ3
JjU0NyY1NDcjJjU0NzPIBARQBAQEBAQEBAQEBAQEKAQEKAQEeAQEUAQEKAQEBAQEBAQEUAQEoAHgERcYEBAYFxFgGBcREBgXERAY
FxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEAABACQAAAFsAjAAYwAAJBUUByMmNTQ3IyY1NDcj
FhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NzMWFRQHFhUUBxYVFAcWFRQHMyY1NDczJjU0NzMWFRQHIxYVFAcj
FhUUBzMWFRQHMwFsBIwEBCgEBDwEBAQEUAQEBAQEBAQEBAQEBAQEUAQEBAQEBAQEPAQEKAQEeAQEUAQEKAQEKAQEZEAYFxERFxgQ
ERcYEBAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxERFxgQERcYEBAYFxEQGBcREBgX
ERAYFxEAAAEATAAAAUQCMAA8AAAkFRQHIyY1NDczJjU0NyY1NDcmNTQ3JjU0NyY1NDcjJjU0NzMWFRQHFhUUBxYVFAcWFRQHFhUU
BxYVFAczAUQE8AQEUAQEBAQEBAQEBAQ8BASMBAQEBAQEBAQEBAQEUEAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgX
ERAYFxEQGBcREBgXERAYFxEQGBcRAAADACQAAAFsAZAAKgBWAIIAABIHMxYVFAcWFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0
NyY1NDczFhUWBzMWFRQHFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyY1NDcjJjU0NzMWFRYVFAcWFRQHFhUUByMmNTQ3JjU0NyY1
NDcmNTQ3IyY1NDczFhUUBzMWFRQHXgQeBAQEBAQEBARQBAQEBAQEBAQEBDIEbgQeBAQEBAQEBAQ8BAQEBAQEBAQyBARQBKAEBAQE
BFAEBAQEBAQEBDIEBFoEBCgEBAFRERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQ
GBcRERcYEBEXGBARFxgQERcYEBEXGBAQGIgYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEAAAIAJAAAAWwB
kAArAFcAABIHMxYVFAcjFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyY1NDcmNTQ3MxYVFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0
NyY1NDcjJjU0NzMWFRQHMxYVFAdyBCgEBB4EBAQEBARQBAQEBAQEBAQEBEYE+gQEBAQEUAQEBAQEBAQEggQEoAQEMgQEAVEREBgX
ERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQEBiIGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcR
//8AJAAAAWwCgAAiAFoAAAADANABkAAAAAIAJAAAAWwBkAAxAE8AACQHIxYVFAcjFhUUByMmNTQ3IyY1NDcjJjU0NzMmNTQ3MyY1
NDczFhUUBzMWFRQHMxYVByY1NDcjJjU0NyMWFRQHIxYVFAczFhUUBzMmNTQ3AWwEHgQEMgQEoAQEMgQEHgQEHgQEMgQEoAQEMgQE
HgRUBAQeBARkBAQeBAQeBARkBASxERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYKBEXGBARFxgQEBgXERAY
FxEQGBcRERcYEAAAAwAkAAABbAKAABMARQBjAAATJjU0NzMmNTQ3MxYVFAcjFhUUBxIHIxYVFAcjFhUUByMmNTQ3IyY1NDcjJjU0
NzMmNTQ3MyY1NDczFhUUBzMWFRQHMxYVByY1NDcjJjU0NyMWFRQHIxYVFAczFhUUBzMmNTQ3oAQEKAQEUAQEKAQEfAQeBAQyBASg
BAQyBAQeBAQeBAQyBASgBAQyBAQeBFQEBB4EBGQEBB4EBB4EBGQEBAHgERcYEBEXGBAQGBcREBgXEf7RERAYFxEQGBcRERcYEBEX
GBARFxgQERcYEBEXGBAQGBcREBgXERAYKBEXGBARFxgQEBgXERAYFxEQGBcRERcYEAD//wAkAAABbAKAACIAXAAAAAMAzgGQAAAA
BAAkAAABbAIwAAkAEwBFAGMAABMmNTQ3MxYVFAczJjU0NzMWFRQHEgcjFhUUByMWFRQHIyY1NDcjJjU0NyMmNTQ3MyY1NDczJjU0
NzMWFRQHMxYVFAczFhUHJjU0NyMmNTQ3IxYVFAcjFhUUBzMWFRQHMyY1NDdQBARQBARQBARQBAQsBB4EBDIEBKAEBDIEBB4EBB4E
BDIEBKAEBDIEBB4EVAQEHgQEZAQEHgQEHgQEZAQEAeARFxgQEBgXEREXGBAQGBcR/tEREBgXERAYFxERFxgQERcYEBEXGBARFxgQ
ERcYEBAYFxEQGBcREBgoERcYEBEXGBAQGBcREBgXERAYFxERFxgQAAADACQAAAFsAoAAEwBFAGMAABMmNTQ3MxYVFAczFhUUByMm
NTQ3EgcjFhUUByMWFRQHIyY1NDcjJjU0NyMmNTQ3MyY1NDczJjU0NzMWFRQHMxYVFAczFhUHJjU0NyMmNTQ3IxYVFAcjFhUUBzMW
FRQHMyY1NDd4BARQBAQoBARQBATMBB4EBDIEBKAEBDIEBB4EBB4EBDIEBKAEBDIEBB4EVAQEHgQEZAQEHgQEHgQEZAQEAjARFxgQ
EBgXERAYFxERFxgQ/oEREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgoERcYEBEXGBAQGBcREBgXERAYFxER
FxgQAAADACT/sAFsAeAARQBjAG0AACQHIxYVFAcjFhUUByMWFRQHIyY1NDczJjU0NyMmNTQ3IyY1NDczJjU0NzMmNTQ3MyY1NDcz
FhUUByMWFRQHMxYVFAczFhUHJjU0NyMmNTQ3IxYVFAcjFhUUBzMWFRQHMyY1NDcmByMmNTQ3MxYVAWwEHgQEMgQEjAQEUAQEKAQE
HgQEHgQEHgQEMgQEjAQEUAQEKAQEHgQEHgRUBAQyBARQBAQeBAQyBARQBAQQBDwEBDwEsREQGBcREBgXERAYFxERFxgQERcYEBEX
GBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgoERcYEBEXGBAQGBcREBgXERAYFxERFxgQERERFxgQEBj//wAkAAABbAKA
ACIAXAAAAAMA0AGQAAAAAQAk/2ABbAGQAGUAACQHIxYVFAcjFhUUByMWFRQHFhUUByMmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDcm
NTQ3MxYVFAczFhUUByMWFRQHFhUUBzMmNTQ3MyY1NDcjJjU0NyMmNTQ3MxYVFAczFhUUBzMWFQFsBCgEBCgEBKAEBAQEUAQEBAQE
BAQEBAQEBAQEUAQEKAQEKAQEBAR4BAQoBAQoBARQBAR4BAQoBAQoBLEREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBAR
FxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgAAgAk/2ABbAIwAE8AawAAJAcjFhUU
ByMWFRQHIxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NzMWFRQHFhUUBzMWFRQHMxYVFAczFhUH
JjU0NyMmNTQ3IxYVFAcWFRQHFhUUBzMmNTQ3AWwEKAQEKAQEoAQEBARQBAQEBAQEBAQEBAQEBAQEBAQEUAQEBASgBAQoBAQoBFQE
BCgEBHgEBAQEBAR4BASxERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgX
ERAYFxEQGBcREBgoERcYEBEXGBAQGBcREBgXERAYFxERFxgQAAABACT/YAFsAZAAZAAAABUUBxYVFAcWFRQHFhUUBxYVFAcWFRQH
IyY1NDcmNTQ3JjU0NyMmNTQ3MyY1NDcmNTQ3IxYVFAcjFhUUBzMWFRQHMxYVFAcjJjU0NyMmNTQ3IyY1NDczJjU0NzMmNTQ3MxYV
FAcBbAQEBAQEBAQEBAQEUAQEBAQEBCgEBCgEBAQEeAQEKAQEKAQEUAQEeAQEKAQEKAQEKAQEKAQE8AQEATAYFxEQGBcREBgXERAY
FxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEA
AQAkAAABbAGQAEIAAAAVFAcjJjU0NyMWFRQHIxYVFAcWFRQHMxYVFAcjJjU0NzMmNTQ3JjU0NyY1NDcjJjU0NzMWFRQHMyY1NDcz
FhUUBzMBbARQBARGBAQeBAQEBDwEBMgEBDwEBAQEBAQ8BARkBAQoBASWBAQeATAYFxERFxgQEBgXERAYFxEQGBcREBgXEREXGBAR
FxgQERcYEBEXGBARFxgQEBgXEREXGBAQGBcRAAEAOAAAAVgBkAAxAAAkByMWFRQHIyY1NDczJjU0NyMmNTQ3IyY1NDczJjU0NzMW
FRQHIxYVFAczFhUUBzMWFQFYBCgEBNwEBLQEBKAEBCgEBCgEBNwEBLQEBKAEBCgEYREQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQ
GBcREBgXERAYFxEQGAABACQAAAFsAjAAdgAAAAcjFhUUBzMWFRQHFhUUByMWFRQHIyY1NDczJjU0NyY1NDcjJjU0NzMmNTQ3JjU0
NyMWFRQHFhUUBxYVFAcWFRQHFhUUBxYVFAcjJjU0NzMmNTQ3JjU0NyY1NDcjJjU0NzMmNTQ3MyY1NDczFhUUBzMWFRQHFhUBbAQo
BAQoBAQEBCgEBHgEBFAEBAQEUAQEUAQEBAR4BAQEBAQEBAQEBAQEeAQEKAQEBAQEBCgEBCgEBCgEBMgEBCgEBAQBUREQGBcREBgX
ERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEX
GBARFxgQERcYEBAYFxEQGBcREBgAAAEAJAAAAUQCMAA/AAAlFhUUByMmNTQ3IyY1NDcmNTQ3JjU0NyMmNTQ3MyY1NDcmNTQ3MxYV
FAcWFRQHMxYVFAcjFhUUBxYVFAcWFRQHAUAEBKAEBCgEBAQEBARQBARQBAQEBFAEBAQEeAQEeAQEBAQEBFAQGBcRERcYEBEXGBAR
FxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcREBgXEQACACQAAAFsAZAALABYAAA2BzMWFRQHIyY1NDcjJjU0
NyY1NDcmNTQ3JjU0NzMWFRQHFhUUBxYVFAczFhU2FRQHFhUUBxYVFAcWFRQHIyY1NDcjJjU0NzMmNTQ3JjU0NyY1NDczFhUUB5oE
ZAQEoAQEMgQEBAQEBAQEUAQEBAQEBB4E0gQEBAQEBARQBAQeBAQeBAQEBAQEUAQEYREQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQ
GBcREBgXERAYFxEQGLgYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEAAAMAJAAAAWwCgAATAEAAbAAAEyY1
NDczJjU0NzMWFRQHIxYVFAcCBzMWFRQHIyY1NDcjJjU0NyY1NDcmNTQ3JjU0NzMWFRQHFhUUBxYVFAczFhU2FRQHFhUUBxYVFAcW
FRQHIyY1NDcjJjU0NzMmNTQ3JjU0NyY1NDczFhUUB6AEBCgEBFAEBCgEBFYEZAQEoAQEMgQEBAQEBAQEUAQEBAQEBB4E0gQEBAQE
BARQBAQeBAQeBAQEBAQEUAQEAeARFxgQERcYEBAYFxEQGBcR/oEREBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcR
EBi4GBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcR//8AJAAAAWwCgAAiAGoAAAADAM4BkAAAAAQAJAAAAWwC
MAAJABMAQABsAAATJjU0NzMWFRQHMyY1NDczFhUUBwIHMxYVFAcjJjU0NyMmNTQ3JjU0NyY1NDcmNTQ3MxYVFAcWFRQHFhUUBzMW
FTYVFAcWFRQHFhUUBxYVFAcjJjU0NyMmNTQ3MyY1NDcmNTQ3JjU0NzMWFRQHUAQEUAQEUAQEUAQEpgRkBASgBAQyBAQEBAQEBARQ
BAQEBAQEHgTSBAQEBAQEBFAEBB4EBB4EBAQEBARQBAQB4BEXGBAQGBcRERcYEBAYFxH+gREQGBcRERcYEBEXGBARFxgQERcYEBEX
GBAQGBcREBgXERAYFxEQGLgYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEAAwAkAAABbAKAABMAQABsAAAT
JjU0NzMWFRQHMxYVFAcjJjU0NwIHMxYVFAcjJjU0NyMmNTQ3JjU0NyY1NDcmNTQ3MxYVFAcWFRQHFhUUBzMWFTYVFAcWFRQHFhUU
BxYVFAcjJjU0NyMmNTQ3MyY1NDcmNTQ3JjU0NzMWFRQHeAQEUAQEKAQEUAQEBgRkBASgBAQyBAQEBAQEBARQBAQEBAQEHgTSBAQE
BAQEBFAEBB4EBB4EBAQEBARQBAQCMBEXGBAQGBcREBgXEREXGBD+MREQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAY
FxEQGLgYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEAAQAkAAABbAGQAE8AAAAHIxYVFAcjFhUUByMWFRQH
IxYVFAcjJjU0NyMmNTQ3IyY1NDcjJjU0NyMmNTQ3MxYVFAczFhUUBzMWFRQHMyY1NDczJjU0NzMmNTQ3MxYVAWwEHgQEHgQEFAQE
KAQEUAQEKAQEFAQEHgQEHgQEUAQEHgQEHgQEKAQEHgQEHgQEUAQBUREQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEX
GBAQGBcREBgXERAYFxERFxgQERcYEBEXGBAQGAAAAQAkAAABbAGQAHcAAAAHFhUUByMWFRQHFhUUByMWFRQHIyY1NDcjJjU0NyMW
FRQHIxYVFAcjJjU0NyMmNTQ3JjU0NyMmNTQ3JjU0NzMWFRQHFhUUBxYVFAczJjU0NzMmNTQ3JjU0NzMWFRQHFhUUBzMWFRQHMyY1
NDcmNTQ3JjU0NzMWFQFsBAQEHgQEBAQUBAQ8BAQUBAQ8BAQUBAQ8BAQUBAQEBB4EBAQEUAQEBAQEBB4EBBQEBAQEPAQEBAQUBAQe
BAQEBAQEUAQBUREQGBcREBgXERAYFxEQGBcRERcYEBEXGBAQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcR
ERcYEBEXGBARFxgQEBgXERAYFxEQGBcRERcYEBEXGBARFxgQEBgAAAEAOAAAAVgBkABZAAAlFhUUByMmNTQ3IyY1NDcjFhUUByMW
FRQHIyY1NDczJjU0NzMmNTQ3IyY1NDcjJjU0NzMWFRQHMxYVFAczJjU0NzMmNTQ3MxYVFAcjFhUUByMWFRQHMxYVFAcBVAQEUAQE
HgQEMgQEKAQEUAQEKAQEKAQEKAQEKAQEUAQEKAQEMgQEHgQEUAQEHgQEMgQEMgQEUBAYFxERFxgQERcYEBAYFxEQGBcRERcYEBEX
GBARFxgQERcYEBEXGBAQGBcREBgXEREXGBARFxgQEBgXERAYFxEQGBcREBgXEQABABr/YAFsAZAAawAAAAcjFhUUByMWFRQHIxYV
FAcjFhUUByMWFRQHIxYVFAcjJjU0NzMmNTQ3JjU0NyMmNTQ3IyY1NDcjJjU0NyMmNTQ3MxYVFAczFhUUBzMWFRQHFhUUBzMmNTQ3
MyY1NDczJjU0NzMmNTQ3MxYVAWwEHgQEHgQEFAQEHgQEFAQEKAQEeAQEUAQEBAQoBAQUBAQeBAQeBARQBAQeBAQeBAQEBB4EBBQE
BB4EBB4EBFAEAVEREBgXERAYFxEQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgX
ERAYFxERFxgQERcYEBEXGBARFxgQEBgAAAIAGv9gAWwCgAATAH8AABMmNTQ3MyY1NDczFhUUByMWFRQHFgcjFhUUByMWFRQHIxYV
FAcjFhUUByMWFRQHIxYVFAcjJjU0NzMmNTQ3JjU0NyMmNTQ3IyY1NDcjJjU0NyMmNTQ3MxYVFAczFhUUBzMWFRQHFhUUBzMmNTQ3
MyY1NDczJjU0NzMmNTQ3MxYVoAQEKAQEUAQEKAQEfAQeBAQeBAQUBAQeBAQUBAQoBAR4BARQBAQEBCgEBBQEBB4EBB4EBFAEBB4E
BB4EBAQEHgQEFAQEHgQEHgQEUAQB4BEXGBARFxgQEBgXERAYFxGPERAYFxEQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcY
EBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBAYAAADABr/YAFsAjAACQATAH8AABMmNTQ3
MxYVFAczJjU0NzMWFRQHFgcjFhUUByMWFRQHIxYVFAcjFhUUByMWFRQHIxYVFAcjJjU0NzMmNTQ3JjU0NyMmNTQ3IyY1NDcjJjU0
NyMmNTQ3MxYVFAczFhUUBzMWFRQHFhUUBzMmNTQ3MyY1NDczJjU0NzMmNTQ3MxYVUAQEUAQEUAQEUAQELAQeBAQeBAQUBAQeBAQU
BAQoBAR4BARQBAQEBCgEBBQEBB4EBB4EBFAEBB4EBB4EBAQEHgQEFAQEHgQEHgQEUAQB4BEXGBAQGBcRERcYEBAYFxGPERAYFxEQ
GBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQ
ERcYEBAYAAADADgAAAFYAZAAEwAdADAAAAAHIyY1NDcjJjU0NyEWFRQHIxYVBgcjJjU0NzMWFRYVFAchJjU0NyY1NDczFhUUBzMB
MARQBASgBAQBGAQEKARQBFAEBFAEeAT+6AQEBARQBATIAQERERcYEBEXGBAQGBcREBhnEREXGBAQGIgYFxERFxgQERcYEBAYFxEA
AQAkAAABbAIwAHUAAAAHIxYVFAcWFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyMWFRQHFhUUBxYVFAcWFRQHIyY1NDcmNTQ3
JjU0NyY1NDcjJjU0NzMmNTQ3MyY1NDczFhUUByMWFRQHMyY1NDczJjU0NzMWFRQHIxYVFAczFhUBbAQyBAQEBAQEBARQBAQEBAQE
BARGBAQEBAQEBARQBAQEBAQEBAQoBAQoBAQoBARQBAQoBARGBAQoBARQBAQoBAQyBAFRERAYFxEQGBcREBgXERAYFxERFxgQERcY
EBEXGBARFxgQEBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXEREXGBARFxgQEBgXERAY
FxEQGAAAAQAkAAABbAIwAGIAAAAVFAcWFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyMWFRQHFhUUBxYVFAczFhUUByMmNTQ3
MyY1NDcmNTQ3JjU0NyMmNTQ3MyY1NDczJjU0NzMWFRQHIxYVFAczFhUUBwFsBAQEBAQEBFAEBAQEBAQEBHgEBAQEBAQoBASgBAQo
BAQEBAQEKAQEKAQEKAQEoAQEeAQEyAQEATAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXEREX
GBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEAAQAkAAABbAIwAHIAAAAVFAcWFRQHFhUUBxYVFAcWFRQHFhUU
ByMmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDcjFhUUBzMWFRQHIxYVFAcWFRQHFhUUBzMWFRQHIyY1NDczJjU0NyY1NDcmNTQ3IyY1
NDczJjU0NzMmNTQ3MxYVFAcBbAQEBAQEBAQEBAQEUAQEBAQEBAQEBAQEBHgEBCgEBCgEBAQEBAQoBASgBAQoBAQEBAQEKAQEKAQE
KAQE8AQEAdAYFxEQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcR
EBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcRAAMAVgAAAToCMAAuADgAQgAAABUUBxYVFAcWFRQHIyY1NDcjJjU0
NzMmNTQ3MyY1NDcjJjU0NzMWFRQHMxYVFAcGNTQ3IxYVFAczFhUUByMmNTQ3MwE6BAQEBAS+BAQeBAQeBASCBARuBASMBAQeBARA
BGQEBGRABNwEBNwBgBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXEY8XGBAQGBcRsBgXEREXGBAAAwBWAAAB
OgIwAC0ARwBRAAAAFRQHFhUUByMWFRQHIyY1NDcjJjU0NyY1NDcmNTQ3MyY1NDczFhUUBzMWFRQHBjU0NyY1NDcmNTQ3IxYVFAcW
FRQHFhUUBzMWFRQHIyY1NDczAToEBAQeBASgBAQeBAQEBAQEHgQEoAQEHgQEQAQEBAQEZAQEBAQEBGRABNwEBNwBgBgXERAYFxEQ
GBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXEY8XGBARFxgQERcYEBAYFxEQGBcREBgXEbAYFxERFxgQAAQAQgAAAU4CMAAJ
ADUAYQBrAAASByMmNTQ3MxYVAhUUByMmNTQ3IyY1NDcmNTQ3JjU0NzMmNTQ3MxYVFAcWFRQHFhUUBxYVFAc2FRQHFhUUByMWFRQH
IyY1NDcmNTQ3JjU0NyY1NDcmNTQ3MxYVFAczFhUUBwI3MxYVFAcjJjX+BGQEBGQEZAQyBAQeBAQEBAQEHgQEMgQEBAQEBAQEuAQE
BB4EBDIEBAQEBAQEBAQEMgQEHgQEuARkBARkBAHxEREXGBAQGP6IGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQ
GBcRkBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXEf8AEBAYFxERFwAAAQBMAAABRAIwAD0AACQVFAcjJjU0
NzMmNTQ3JjU0NyY1NDcmNTQ3IyY1NDczJjU0NzMWFRQHFhUUBxYVFAcWFRQHFhUUBxYVFAczAUQE8AQEUAQEBAQEBAQEPAQEPAQE
UAQEBAQEBAQEBAQEBFBAGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcREBgXEQAAAQA4
AAABWAIwAE0AAAAHIxYVFAcjFhUUByMWFRQHMxYVFAchJjU0NzMmNTQ3MyY1NDczJjU0NzMmNTQ3JjU0NyMWFRQHIyY1NDczJjU0
NzMWFRQHMxYVFAcWFQFYBDIEBDIEBDIEBJYEBP7oBAQyBAQyBAQyBAQyBAQEBIwEBDwEBCgEBMgEBCgEBAQBUREQGBcREBgXERAY
FxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXEREXGBARFxgQEBgXERAYFxEQGAABADgAAAFYAjAAVQAAAAcjFhUUBzMW
FRQHFhUUByMWFRQHIyY1NDcjJjU0NzMWFRQHMyY1NDcmNTQ3IyY1NDczJjU0NyY1NDcjFhUUByMmNTQ3MyY1NDczFhUUBzMWFRQH
FhUBWAQoBAQoBAQEBCgEBMgEBCgEBFAEBHgEBAQEZAQEZAQEBAR4BARQBAQoBATIBAQoBAQEAVEREBgXERAYFxEQGBcREBgXEREX
GBARFxgQEBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgXEREXGBARFxgQEBgXERAYFxEQGAAAAQAkAAABbAIwAFwAACQHIxYVFAcW
FRQHIyY1NDcmNTQ3IyY1NDczJjU0NzMmNTQ3MyY1NDczFhUUByMWFRQHIxYVFAczJjU0NyY1NDcmNTQ3IyY1NDczFhUUBxYVFAcW
FRQHFhUUBzMWFQFsBCgEBAQEUAQEBATIBAQeBAQeBAQeBAQyBAQeBAQeBAR4BAQEBAQEPAQEjAQEBAQEBAQEKASxERAYFxEQGBcR
ERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGAACADgA
AAFsAjAAIwBfAAASBzMWFRQHIyY1NDcmNTQ3JjU0NyY1NDchFhUUByMWFRQHFhUWByMWFRQHIxYVFAcjJjU0NyMmNTQ3MxYVFAcz
JjU0NzMmNTQ3IyY1NDcjJjU0NzMWFRQHMxYVFAczFhWQBB4EBG4EBAQEBAQEBAEOBAS+BAQE3AQeBAQyBAS0BAQoBAQ8BASCBAQe
BAQeBARQBARuBAQyBAQeBAFRERAYFxERFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGLcREBgXERAYFxERFxgQERcYEBAYFxERFxgQ
ERcYEBEXGBARFxgQEBgXERAYFxEQGAADADMAAAFdAjAACQA3AF0AAAAHIyY1NDczFhUCFRQHIyY1NDcjJjU0NyY1NDcmNTQ3MyY1
NDczFhUUByMWFRQHFhUUBzMWFRQHNgcWFRQHIxYVFAcjJjU0NzMmNTQ3JjU0NyMmNTQ3MxYVFAczFhUBNQSMBASMBIwEUAQEHgQE
BAQEBB4EBFAEBB4EBAQEHgQEuAQEBDwEBHgEBGQEBAQEZAQEjAQEKAQB8RERFxgQEBj+iBgXEREXGBARFxgQERcYEBEXGBARFxgQ
EBgXERAYFxEQGBcREBgXEREREBgXERAYFxERFxgQERcYEBEXGBARFxgQEBgXERAYAAABADgAAAFYAjAARAAAAAcWFRQHIxYVFAcj
FhUUByMWFRQHIxYVFAcjFhUUByMmNTQ3MyY1NDczJjU0NzMmNTQ3MyY1NDczJjU0NyMmNTQ3IRYVAVgEBAQeBAQeBAQeBAQeBAQe
BARQBAQeBAQeBAQeBAQeBAQeBATIBAQBGAQB8REQGBcREBgXERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEX
GBARFxgQEBgAAwBCAAABTgIwAEEAUwBlAAAAByMWFRQHMxYVFAcWFRQHIxYVFAcjJjU0NyMmNTQ3JjU0NzMmNTQ3IyY1NDcmNTQ3
MyY1NDczFhUUBzMWFRQHFhUGNTQ3JjU0NyMWFRQHFhUUBzMGNTQ3IxYVFAcWFRQHMyY1NDcBTgQyBAQyBAQEBDIEBKAEBDIEBAQE
MgQEMgQEBAQyBASgBAQyBAQEWAQEBGQEBAQEZAQEZAQEBARkBAQBUREQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEX
GBARFxgQERcYEBAYFxEQGBcREBgXFxgQERcYEBAYFxEQGBcRjxcYEBAYFxEQGBcRERcYEAAAAwAzAAABXQIwACUAUwBdAAASBzMW
FRQHIyY1NDcjJjU0NyY1NDczJjU0NzMWFRQHIxYVFAcWFRYVFAcWFRQHIxYVFAcjJjU0NzMmNTQ3JjU0NyMmNTQ3JjU0NzMWFRQH
MxYVFAcCNzMWFRQHIyY1iwRkBASMBAQoBAQEBDwEBHgEBGQEBATSBAQEHgQEUAQEHgQEBAQeBAQEBFAEBB4EBP4EjAQEjAQBUREQ
GBcRERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBg4GBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcR/wAQEBgX
EREXAAEAfgDwAOoCgAArAAASFRQHFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyMmNTQ3MyY1NDczFhUUB+oEBAQEBAQEPAQEBAQE
BCgEBCgEBDwEBAIgGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcRAAABAHQA8AEcAoAAMQAAARYVFAcjJjU0
NzMmNTQ3MyY1NDczJjU0NyMmNTQ3MxYVFAczFhUUByMWFRQHIxYVFAcBGAQEoAQEKAQEHgQEHgQEZAQEggQEHgQEHgQEHgQEAUAQ
GBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcRAAABAHQA8AEcAoAAMQAAARYVFAcjFhUUByMmNTQ3MyY1NDcj
JjU0NzMmNTQ3IyY1NDczFhUUBzMWFRQHIxYVFAcBGAQEHgQEggQEZAQERgQERgQEZAQEggQEHgQEHgQEAZAQGBcREBgXEREXGBAR
FxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcRAAACABUAAAFsAoAAKwCPAAASFRQHIyY1NDcmNTQ3JjU0NyMmNTQ3MyY1NDczFhUU
BxYVFAcWFRQHFhUUBxcWFRQHIyY1NDczJjU0NzMmNTQ3MyY1NDcjFhUUByMWFRQHIxYVFAcjJjU0NzMmNTQ3MyY1NDczJjU0NzMm
NTQ3MyY1NDczFhUUByMWFRQHMxYVFAczFhUUByMWFRQHIxYVFAeBBDwEBAQEBAQoBAQoBAQ8BAQEBAQEBATrBASgBAQoBAQeBAQe
BARVBAQeBAQeBAQ8BAQeBAQeBAQeBAQeBAQeBAQ8BAQeBAQ3BAQeBAQeBAQeBAQBMBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgX
ERAYFxEQGBcREBgXEfAQGBcRERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQ
GBcREBgXERAYFxEQGBcREBgXEQAEABX/sAGAAoAAKwBnAJcAqgAAEhUUByMmNTQ3JjU0NyY1NDcjJjU0NzMmNTQ3MxYVFAcWFRQH
FhUUBxYVFAcXFhUUByMWFRQHIyY1NDczJjU0NzMmNTQ3MyY1NDczJjU0NzMmNTQ3MxYVFAcjFhUUByMWFRQHIxYVFAcWByMWFRQH
IyY1NDcjJjU0NzMmNTQ3MyY1NDczJjU0NzMWFRQHFhUUBxYVFAczFhUmFRQHMyY1NDcmNTQ3IxYVFAcjgQQ8BAQEBAQEKAQEKAQE
PAQEBAQEBAQEPAQEHgQEPAQEHgQEHgQEHgQEHgQEHgQEPAQEHgQEHgQEHgQEqQQUBAQ8BARkBAQUBAQUBAQUBARkBAQEBAQEFAR4
BCgEBAQEFAQEFAEwGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcRUBAYFxEQGBcRERcYEBEXGBARFxgQERcY
EBEXGBARFxgQEBgXERAYFxEQGBcREBgXEd8REBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBhoGBcRERcYEBEX
GBAQGBcRAAADACT/sAGAAoAAWQCJAJwAADcWFRQHIxYVFAcjJjU0NzMmNTQ3IyY1NDczJjU0NyMmNTQ3MyY1NDcjJjU0NzMWFRQH
MxYVFAcjFhUUBzMmNTQ3MyY1NDczFhUUByMWFRQHIxYVFAcjFhUUBxYHIxYVFAcjJjU0NyMmNTQ3MyY1NDczJjU0NzMmNTQ3MxYV
FAcWFRQHFhUUBzMWFSYVFAczJjU0NyY1NDcjFhUUByO5BAQeBAQ8BAQeBARVBARkBARGBARGBARkBASCBAQeBAQeBAQtBAQeBAQ8
BAQeBAQeBAQeBASpBBQEBDwEBGQEBBQEBBQEBBQEBGQEBAQEBAQUBHgEKAQEBAQUBAQU8BAYFxEQGBcRERcYEBEXGBARFxgQERcY
EBEXGBARFxgQERcYEBAYFxEQGBcREBgXEREXGBARFxgQEBgXERAYFxEQGBcREBgXEd8REBgXEREXGBARFxgQERcYEBEXGBARFxgQ
EBgXERAYFxEQGBcREBhoGBcRERcYEBEXGBAQGBcRAAUAOABQAVgB4AAtADcAQQBLAFUAAAAVFAcjFhUUBxYVFAcjJjU0NyY1NDcj
JjU0NzMmNTQ3JjU0NzMWFRQHFhUUBzMmFRQHIyY1NDczFgcjJjU0NzMWFQQ3MxYVFAcjJjUlFhUUByMmNTQ3AQgEHgQEBAQ8BAQE
BB4EBB4EBAQEPAQEBAQedARQBARQzARQBARQBP7gBFAEBFAEARwEBFAEBAEwGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcY
EBAYFxEQGBcRQBgXEREXGBA/EREXGBAQGIgQEBgXEREXKBAYFxERFxgQAAABACT/sAFsAoAAWQAABBUUByMmNTQ3IyY1NDcjJjU0
NyMmNTQ3IyY1NDcjJjU0NyMmNTQ3IyY1NDcjJjU0NzMWFRQHMxYVFAczFhUUBzMWFRQHMxYVFAczFhUUBzMWFRQHMxYVFAczAWwE
UAQEHgQEHgQEHgQEHgQEHgQEHgQEHgQEHgQEUAQEHgQEHgQEHgQEHgQEHgQEHgQEHgQEHhAYFxERFxgQERcYEBEXGBARFxgQERcY
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
IxYVFAcjFhUUByMWFRQHIxYVFAcjJjU0NzMmNTQ3MyY1NDczJjU0NzMmNTQ3MyY1NDczJjU0NzMmNTQ3MyY1NDczFhUBbAQeBAQe
BAQeBAQeBAQeBAQeBAQeBAQeBARQBAQeBAQeBAQeBAQeBAQeBAQeBAQeBAQeBARQBAJBERAYFxEQGBcREBgXERAYFxEQGBcREBgX
ERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgAAAEAJAAAAWwAUAAJAAAkFRQHISY1NDchAWwE
/sAEBAFAQBgXEREXGBAAAAEAnACgAPQA8AAJAAA3FhUUByMmNTQ38AQEUAQE8BAYFxERFxgQAAABAEz/sAFEAoAAUQAAIRYVFAcj
JjU0NyMmNTQ3JjU0NyY1NDcjJjU0NzMmNTQ3JjU0NyY1NDczJjU0NzMWFRQHIxYVFAcWFRQHFhUUByMWFRQHMxYVFAcWFRQHFhUU
BwFABAR4BAQoBAQEBAQEUAQEUAQEBAQEBCgEBHgEBFAEBAQEBAQoBAQoBAQEBAQEEBgXEREXGBARFxgQERcYEBEXGBARFxgQERcY
EBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcREBgXERAYFxEQGBcRAAABAEz/sAFEAoAAUQAAAAcjFhUUBxYVFAcWFRQH
IxYVFAcjJjU0NzMmNTQ3JjU0NyY1NDczJjU0NyMmNTQ3JjU0NyY1NDcjJjU0NzMWFRQHMxYVFAcWFRQHFhUUBzMWFQFEBFAEBAQE
BAQoBAR4BARQBAQEBAQEKAQEKAQEBAQEBFAEBHgEBCgEBAQEBARQBAEBERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQ
ERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgAAAEAiP+wARwCgABLAAAhFhUUByMmNTQ3JjU0NyY1NDcmNTQ3
JjU0NyY1NDcmNTQ3JjU0NyY1NDczFhUUByMWFRQHFhUUBxYVFAcWFRQHFhUUBxYVFAcWFRQHARgEBIwEBAQEBAQEBAQEBAQEBAQE
BASMBAQ8BAQEBAQEBAQEBAQEBAQQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgX
ERAYFxEQGBcREBgXERAYFxEAAAEAdP+wAQgCgABLAAAAFRQHFhUUBxYVFAcWFRQHFhUUBxYVFAcWFRQHFhUUByMmNTQ3MyY1NDcm
NTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDcjJjU0NzMWFRQHAQgEBAQEBAQEBAQEBAQEBASMBAQ8BAQEBAQEBAQEBAQEBAQ8BASMBAQC
IBgXERAYFxEQGBcREBgXERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcRAAMA
fv+wASECgAAJAEkAUwAAAAcjJjU0NzMWFQIVFAcjJjU0NyMmNTQ3IyY1NDcmNTQ3JjU0NzMmNTQ3MyY1NDczFhUUByMWFRQHFhUU
BxYVFAcWFRQHFhUUBzMXFhUUByMmNTQ3ASEEMgQEMgQyBDIEBBkEBB4EBAQEBAQeBAQZBAQyBAQZBAQEBAQEBAQEBBkyBAQyBAQC
QRERFxgQEBj96BgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXERAYFxFQEBgXEREXGBAA
AAMAb/+wARICgAAJAEkAUwAAEhUUByMmNTQ3MxIVFAcWFRQHIxYVFAcjFhUUByMmNTQ3MyY1NDcmNTQ3JjU0NyY1NDcmNTQ3IyY1
NDczFhUUBzMWFRQHMxYVFAcCNzMWFRQHIyY1qQQyBAQybQQEBB4EBBkEBDIEBBkEBAQEBAQEBAQEGQQEMgQEGQQEHgQEnwQyBAQy
BAJwGBcRERcYEP6wGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXEf6wEBAYFxER
FwABACQA8AFsAUAACQAAABUUByEmNTQ3IQFsBP7ABAQBQAEwGBcRERcYEAABACQA8AFsAUAACQAAABUUByEmNTQ3IQFsBP7ABAQB
QAEwGBcRERcYEAABAGAA8AEwAUAACQAAARYVFAcjJjU0NwEsBATIBAQBQBAYFxERFxgQAP//AGAA8AEwAUAAAgCmAAAAAgAkAAAB
bAGQADEAYwAAEgcjFhUUByMWFRQHMxYVFAczFhUUByMmNTQ3IyY1NDcjJjU0NzMmNTQ3MyY1NDczFhUTFhUUByMmNTQ3IyY1NDcj
JjU0NzMmNTQ3MyY1NDczFhUUByMWFRQHIxYVFAczFhUUB7gEKAQEKAQEKAQEKAQEPAQEKAQEKAQEKAQEKAQEPASwBAQ8BAQoBAQo
BAQoBAQoBAQ8BAQoBAQoBAQoBAQBUREQGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQGP7oEBgXEREXGBARFxgQ
ERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXEQACACQAAAFsAZAAMQBjAAA3FhUUByMWFRQHIxYVFAcjJjU0NzMmNTQ3MyY1NDcj
JjU0NyMmNTQ3MxYVFAczFhUUBxYHIxYVFAcjFhUUByMmNTQ3MyY1NDczJjU0NyMmNTQ3IyY1NDczFhUUBzMWFRQHMxYVtAQEKAQE
KAQEPAQEKAQEKAQEKAQEKAQEPAQEKAQE4AQoBAQoBAQ8BAQoBAQoBAQoBAQoBAQ8BAQoBAQoBPAQGBcREBgXERAYFxERFxgQERcY
EBEXGBARFxgQERcYEBAYFxEQGBcRPxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGAACAEIBkAFOAtAAJABJ
AAASByMWFRQHMxYVFAcWFRQHIyY1NDcmNTQ3JjU0NzMmNTQ3MxYVFhUUBxYVFAcjJjU0NyY1NDcmNTQ3MyY1NDczFhUUByMWFRQH
M64EKAQEKAQEBARkBAQEBAQEKAQEPASgBAQEZAQEBAQEBCgEBDwEBCgEBCgCkREQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQ
EBiIGBcREBgXEREXGBARFxgQERcYEBEXGBAQGBcREBgXEQACAEIBkAFOAtAAJABJAAASFRQHFhUUBxYVFAcjFhUUByMmNTQ3MyY1
NDcjJjU0NyY1NDczFhUUBxYVFAcjFhUUByMmNTQ3MyY1NDcjJjU0NyY1NDczFhUUB64EBAQEBCgEBDwEBCgEBCgEBAQEZKQEBAQo
BAQ8BAQoBAQoBAQEBGQEBALAGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEGAYFxEQGBcREBgXEREXGBARFxgQERcYEBEX
GBAQGBcRAAABAJIBkAD+AtAAJAAAEhUUBxYVFAcjJjU0NyY1NDcmNTQ3MyY1NDczFhUUByMWFRQHM/4EBARkBAQEBAQEKAQEPAQE
KAQEKAIgGBcREBgXEREXGBARFxgQERcYEBEXGBAQGBcREBgXEQAAAQCSAZAA/gLQACQAABIVFAcWFRQHIxYVFAcjJjU0NzMmNTQ3
IyY1NDcmNTQ3MxYVFAf+BAQEKAQEPAQEKAQEKAQEBARkBAQCcBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEBAYFxEAAAEAkv9g
AP4AoAAkAAA2FRQHFhUUByMWFRQHIyY1NDczJjU0NyMmNTQ3JjU0NzMWFRQH/gQEBCgEBDwEBCgEBCgEBAQEZAQEQBgXERAYFxEQ
GBcRERcYEBEXGBARFxgQERcYEBAYFxEAAQAk/7ABbAKAAGUAACUWFRQHIxYVFAcjFhUUBxYVFAcjJjU0NyY1NDcjJjU0NyMmNTQ3
JjU0NyY1NDczJjU0NzMmNTQ3JjU0NzMWFRQHFhUUBzMWFRQHMxYVFAcjJjU0NyMWFRQHFhUUBxYVFAczJjU0NwFoBAQoBARaBAQE
BDwEBAQEWgQEKAQEBAQEBCgEBFoEBAQEPAQEBARaBAQoBARQBASgBAQEBAQEoAQE8BAYFxEQGBcREBgXERAYFxERFxgQERcYEBEX
GBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcRERcYEBAYFxEQGBcREBgXEREXGBAAAgAkAAABbAGQAFkA
dwAAJAcjFhUUBzMWFRQHIyY1NDcjFhUUByMmNTQ3IxYVFAcjJjU0NzMmNTQ3IyY1NDczJjU0NyMmNTQ3MxYVFAczJjU0NzMWFRQH
MyY1NDczFhUUByMWFRQHMxYVByY1NDcjJjU0NyMWFRQHIxYVFAczFhUUBzMmNTQ3AWwEKAQEKAQEPAQEMgQEZAQEMgQEPAQEKAQE
KAQEKAQEKAQEPAQEMgQEZAQEMgQEPAQEKAQEKARUBAQoBARQBAQoBAQoBARQBASxERAYFxEQGBcRERcYEBAYFxERFxgQEBgXEREX
GBARFxgQERcYEBEXGBARFxgQEBgXEREXGBAQGBcRERcYEBAYFxEQGBcREBgoERcYEBEXGBAQGBcREBgXERAYFxERFxgQAAADACT/
sAFsAoAAaQB7AI0AACQVFAcWFRQHIxYVFAcjFhUUByMmNTQ3IyY1NDcjJjU0NzMWFRQHMyY1NDcmNTQ3IyY1NDcjJjU0NyY1NDcz
JjU0NzMmNTQ3MxYVFAczFhUUBzMWFRQHIyY1NDcjFhUUBxYVFAczFhUUBzMnJjU0NyY1NDcjFhUUBxYVFAcWNTQ3JjU0NyMWFRQH
FhUUBzMBbAQEBCgEBFoEBDwEBFoEBCgEBFAEBDIEBAQEWgQEKAQEBAQoBARaBAQ8BARaBAQoBARQBAQyBAQEBFoEBCi+BAQEBDIE
BAQEnAQEBDIEBAQEMuAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBAQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAY
FxEQGBcREBgXEREXGBAQGBcREBgXERAYFxFQERcYEBEXGBAQGBcREBgXEd8XGBARFxgQEBgXERAYFxEAAgAkAAABbAIwAEsAXwAA
NhUUByMWFRQHIyY1NDczJjU0NyY1NDcjJjU0NzMmNTQ3JjU0NzMmNTQ3MxYVFAczFhUUByMmNTQ3IxYVFAcWFRQHMxYVFAcjFhUU
BxYHIxYVFAcjJjU0NzMmNTQ3MxYVuAQoBARkBAQ8BAQEBDwEBDwEBAQEKAQEoAQEKAQEPAQEZAQEBAQ8BAQ8BAS4BBQEBKAEBHgE
BDwEkBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxERFxgQEBgXERAYFxEQGBcREBgXET8REBgXEREX
GBARFxgQEBgAAAEAJAAAAWwCMABhAAAAByMWFRQHIxYVFAczFhUUByMWFRQHMxYVFAcjFhUUByMmNTQ3IyY1NDczJjU0NyMmNTQ3
MyY1NDcjJjU0NyMmNTQ3MxYVFAcWFRQHMxYVFAczJjU0NzMmNTQ3JjU0NzMWFQFsBBQEBCgEBCgEBGQEBGQEBGQEBFAEBGQEBGQE
BGQEBCgEBCgEBBQEBGQEBAQEKAQEKAQEKAQEBARkBAHxERAYFxEQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBAR
FxgQERcYEBEXGBAQGBcREBgXERAYFxERFxgQERcYEBEXGBAQGAAAAQA4AFABWAHgAC0AAAAHIxYVFAcWFRQHIyY1NDcmNTQ3IyY1
NDczJjU0NyY1NDczFhUUBxYVFAczFhUBWARkBAQEBFAEBAQEZAQEZAQEBARQBAQEBGQEAQEREBgXERAYFxERFxgQERcYEBEXGBAR
FxgQERcYEBAYFxEQGBcREBgAAAEATABQAUQCMABhAAAlFhUUByMmNTQ3IyY1NDcjFhUUByMWFRQHIyY1NDczJjU0NzMmNTQ3JjU0
NyMmNTQ3IyY1NDczFhUUBzMWFRQHMyY1NDczJjU0NzMWFRQHIxYVFAcjFhUUBxYVFAczFhUUBwFABAQ8BAQeBAQ8BAQeBAQ8BAQe
BAQoBAQEBCgEBB4EBDwEBB4EBDwEBB4EBDwEBB4EBCgEBAQEKAQEoBAYFxERFxgQERcYEBAYFxEQGBcRERcYEBEXGBARFxgQERcY
EBEXGBARFxgQEBgXERAYFxERFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcRAAMATAAAAUQCMAARABsALQAAEjU0NyY1NDczFhUU
BxYVFAcjFhUUByMmNTQ3MwYVFAcWFRQHIyY1NDcmNTQ3M5wEBARQBAQEBFCkBPAEBPBMBAQEUAQEBARQAaEXGBARFxgQEBgXERAY
FxFgGBcRERcYELAYFxEQGBcRERcYEBEXGBAAAAIATACgAUQBkAAJABMAABI1NDczFhUUByMWFRQHIyY1NDczTATwBATw9ATwBATw
AVEXGBAQGBcRYBgXEREXGBAAAAEAUQAAAT8CMABFAAAAByMWFRQHIxYVFAcjFhUUByMmNTQ3MyY1NDczJjU0NzMmNTQ3IyY1NDcj
JjU0NyMmNTQ3MxYVFAczFhUUBzMWFRQHMxYVAT8EMgQEMgQEMgQEUAQEMgQEMgQEMgQEMgQEMgQEMgQEUAQEMgQEMgQEMgQBAREQ
GBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgAAAEAUQAAAT8CMABFAAAlFhUUByMm
NTQ3IyY1NDcjJjU0NyMmNTQ3MyY1NDczJjU0NzMmNTQ3MxYVFAcjFhUUByMWFRQHIxYVFAczFhUUBzMWFRQHATsEBFAEBDIEBDIE
BDIEBDIEBDIEBDIEBFAEBDIEBDIEBDIEBDIEBDIEBFAQGBcRERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgX
ERAYFxEQGBcREBgXEQACADgAAAFYAjAALQA3AAATJjU0NzMmNTQ3JjU0NzMWFRQHFhUUBzMWFRQHIxYVFAcWFRQHIyY1NDcmNTQ3
EhUUByEmNTQ3ITwEBGQEBAQEUAQEBARkBARkBAQEBFAEBAQEuAT+6AQEARgBQBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQ
GBcRERcYEBEXGBD/ABgXEREXGBAAAAEAJACgAWwB4ABNAAAABxYVFAcjFhUUByMWFRQHIyY1NDcjJjU0NyMWFRQHIxYVFAcjJjU0
NyY1NDczJjU0NzMmNTQ3MxYVFAczFhUUBzMmNTQ3MyY1NDczFhUBbAQEBB4EBB4EBGQEBB4EBDIEBBQEBDwEBAQEHgQEHgQEZAQE
HgQEMgQEFAQEPAQBoREQGBcREBgXERAYFxERFxgQERcYEBAYFxEQGBcRERcYEBEXGBARFxgQERcYEBAYFxEQGBcRERcYEBEXGBAQ
GAAAAQAkAFABbAFAABoAACQVFAcWFRQHIyY1NDcmNTQ3IyY1NDchFhUUBwFsBAQEUAQEBATwBAQBQAQE4BgXERAYFxERFxgQERcY
EBEXGBAQGBcRAAUAJAAAAWwCMABFAGMAbQCLAJUAADcWFRQHIxYVFAcjFhUUByMmNTQ3MyY1NDczJjU0NzMmNTQ3MyY1NDczJjU0
NzMmNTQ3MxYVFAcjFhUUByMWFRQHIxYVFAcnJjU0NzMmNTQ3MxYVFAczFhUUByMWFRQHIyY1NDc3FhUUBzMmNTQ3ExYVFAcjFhUU
ByMmNTQ3IyY1NDczJjU0NzMWFRQHByY1NDcjFhUUB7kEBCgEBC0EBDwEBC0EBCgEBC0EBC0EBCgEBC0EBDwEBC0EBCgEBC0EBL4E
BCMEBFAEBCMEBCMEBFAEBBkEBB4EBOYEBCMEBFAEBCMEBCMEBFAEBBkEBB4EBPAQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQ
ERcYEBEXGBARFxgQEBgXERAYFxEQGBcREBgXEaARFxgQERcYEBAYFxEQGBcREBgXEREXGBBQEBgXEREXGBD+wBAYFxEQGBcRERcY
EBEXGBARFxgQEBgXEVARFxgQEBgXEQACACQAAAFsAoAAdwCJAAAAFRQHFhUUBxYVFAcjJjU0NyMmNTQ3JjU0NzMmNTQ3MyY1NDcj
FhUUByMWFRQHFhUUBxYVFAcWFRQHMxYVFAczFhUUByMmNTQ3IyY1NDcjJjU0NyY1NDcmNTQ3JjU0NzMmNTQ3MyY1NDczFhUUBzMW
FRQHMxYVFAcGNTQ3JjU0NyMWFRQHFhUUBzMBbAQEBAQEoAQEHgQEBAQeBAQoBARkBAQoBAQEBAQEBAQoBAS0BATIBAQoBAQoBAQE
BAQEBAQoBAQoBASgBAQoBAQoBARABAQERgQEBARGAYAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcR
EBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxEQGBcRjxcYEBEXGBAQGBcREBgX
EQACACQAAAFsAjAAcwCQAAAAByMWFRQHMxYVFAczFhUUByMmNTQ3JjU0NyMWFRQHIxYVFAcjJjU0NyMmNTQ3JjU0NzMmNTQ3MyY1
NDcjJjU0NzMmNTQ3MxYVFAczFhUUByMmNTQ3IxYVFAczFhUUBzMWFRQHMyY1NDcmNTQ3MxYVFAcWFQcmNTQ3IyY1NDcjFhUUByMW
FRQHFhUUBzMmNTQ3AWwEMgQEHgQEFAQEUAQEBAQeBAQeBASMBAQoBAQEBCgEBCgEBCgEBCgEBHgEBCgEBDwEBFAEBFAEBB4EBDIE
BAQEPAQEBJAEBB4EBB4EBCgEBAQERgQEAQEREBgXERAYFxEQGBcRERcYEBEXGBAQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQ
ERcYEBEXGBAQGBcREBgXEREXGBAQGBcREBgXERAYFxERFxgQERcYEBAYFxEQGHgRFxgQERcYEBAYFxEQGBcREBgXEREXGBAAAQAk
/7ABbAKAAJEAAAAHIxYVFAcWFRQHFhUUBxYVFAcWFRQHFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDcm
NTQ3IxYVFAcWFRQHFhUUBxYVFAcWFRQHFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyMmNTQ3IyY1NDcjJjU0NyY1NDczJjU0NzMm
NTQ3MxYVAWwEKAQEBAQEBAQEBAQEBAQEBAQ8BAQEBAQEBAQEBAQEBAQEBDwEBAQEBAQEBAQEBAQEBAQEPAQEBAQEBBQEBCgEBCgE
BAQEKAQEKAQE8AQCQREQGBcREBgXERAYFxEQGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAR
FxgQEBgXERAYFxEQGBcREBgXERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGAAA
AgA4/7ABWAIwAE0AYQAAJAcjFhUUByMmNTQ3MyY1NDcjJjU0NyMmNTQ3JjU0NzMmNTQ3IyY1NDczJjU0NzMWFRQHIxYVFAczFhUU
BzMWFRQHFhUUByMWFRQHMxYVJyY1NDcjJjU0NyMWFRQHMxYVFAcBWAQoBATmBAS+BAR4BARGBAQEBB4EBCgEBCgEBOYEBL4EBHgE
BEYEBAQEHgQEKARKBAQoBARkBAQoBAQRERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAY
FxEQGBcREBgXERAYeBEXGBARFxgQEBgXERAYFxEAAAUAJP+wAWwCgAAJAEcAhQCzAL0AAAAHIyY1NDczFhUCFRQHIyY1NDcjJjU0
NyY1NDcmNTQ3JjU0NyY1NDczJjU0NzMWFRQHIxYVFAcWFRQHFhUUBxYVFAcWFRQHMxIVFAcWFRQHFhUUBxYVFAcjFhUUByMmNTQ3
MyY1NDcmNTQ3JjU0NyY1NDcmNTQ3IyY1NDczFhUUBzMWFRQHJgcjFhUUBxYVFAcWFRQHMxYVFAcjJjU0NyMmNTQ3JjU0NyY1NDcz
JjU0NzMWFQI3MxYVFAcjJjUBHASgBASgBKAEMgQEHgQEBAQEBAQEBAQeBAQyBAQeBAQEBAQEBAQEBB70BAQEBAQEBB4EBDIEBB4E
BAQEBAQEBAQEHgQEMgQEHgQEagRBBAQEBAQEQQQEZAQEDwQEBAQEBA8EBGQEigSgBASgBAJBEREXGBAQGP3oGBcRERcYEBEXGBAR
FxgQERcYEBEXGBARFxgQERcYEBAYFxEQGBcREBgXERAYFxEQGBcREBgXEQEwGBcREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcY
EBEXGBARFxgQERcYEBEXGBAQGBcREBgXEREREBgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQEBj+OBAQGBcRERcA
BgAk/7ABbAKAAAkARwCFAMQAzgDYAAAAByMmNTQ3MxYVAhUUByMmNTQ3IyY1NDcmNTQ3JjU0NyY1NDcmNTQ3MyY1NDczFhUUByMW
FRQHFhUUBxYVFAcWFRQHFhUUBzMSFRQHFhUUBxYVFAcWFRQHIxYVFAcjJjU0NzMmNTQ3JjU0NyY1NDcmNTQ3JjU0NyMmNTQ3MxYV
FAczFhUUBwYHIxYVFAczFhUUBxYVFAcjJjU0NyY1NDcjFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyY1NDczFhUUBzMWFQcmNTQ3
IxYVFAcCNzMWFRQHIyY1ARwEoAQEoASgBDIEBB4EBAQEBAQEBAQEHgQEMgQEHgQEBAQEBAQEBAQe9AQEBAQEBAQeBAQyBAQeBAQE
BAQEBAQEBB4EBDIEBB4EBFsEHgQEHgQEBAQyBAQEBB4EBAQEMgQEBAQEBAQEBARkBAQeBDYEBB4EBEUEoAQEoAQCQRERFxgQEBj9
6BgXEREXGBARFxgQERcYEBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXERAYFxEBMBgXERAYFxEQGBcREBgXERAY
FxERFxgQERcYEBEXGBARFxgQERcYEBEXGBARFxgQEBgXERAYFxE/ERAYFxEQGBcREBgXEREXGBARFxgQEBgXERAYFxERFxgQERcY
EBEXGBARFxgQERcYEBAYFxEQGCgRFxgQEBgXEf6wEBAYFxERFwAAAgBWAUABOgKAACUANwAAAAcWFRQHIxYVFAcjJjU0NyMmNTQ3
JjU0NzMmNTQ3MxYVFAczFhUGNTQ3JjU0NyMWFRQHFhUUBzMBOgQEBB4EBKAEBB4EBAQEHgQEoAQEHgREBAQEZAQEBARkAfEREBgX
ERAYFxERFxgQERcYEBEXGBARFxgQEBgXERAYZxcYEBEXGBAQGBcREBgXEQAAAQCc/7AA9AKAAEkAABIVFAcWFRQHFhUUBxYVFAcW
FRQHFhUUBxYVFAcWFRQHIyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NyY1NDcmNTQ3JjU0NzMWFRQH9AQEBAQEBAQEBAQEBAQEBFAE
BAQEBAQEBAQEBAQEBAQEBARQBAQCIBgXERAYFxEQGBcREBgXERAYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQERcYEBEX
GBARFxgQERcYEBEXGBAQGBcRAAACAJz/sAD0AoAAIQBDAAASFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NzMWFRQHEhUUBxYV
FAcWFRQHFhUUByMmNTQ3JjU0NyY1NDcmNTQ3M/QEBAQEBFAEBAQEBAQEBFAEBAQEBAQEBAQEUAQEBAQEBAQEUAIgGBcREBgXERAY
FxERFxgQERcYEBEXGBARFxgQEBgXEf6wGBcREBgXERAYFxEQGBcRERcYEBEXGBARFxgQERcYEAABACQA8AFsAjAAOwAAABUUByMm
NTQ3IyY1NDcjFhUUByMWFRQHIyY1NDczJjU0NzMmNTQ3MyY1NDczFhUUBzMWFRQHMxYVFAczAWwEUAQEKAQEUAQEKAQEUAQEKAQE
KAQEKAQEUAQEKAQEKAQEKAEwGBcRERcYEBEXGBAQGBcREBgXEREXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEAA//8/2ABbAGQ
AD8AbQB3AAA2FRQHIxYVFAcWFRQHIxYVFAcjJjU0NzMmNTQ3JjU0NzMmNTQ3JjU0NzMmNTQ3JjU0NzMWFRQHFhUUByMWFRQHNgcj
FhUUBxYVFAczFhUUByMmNTQ3IyY1NDcmNTQ3MyY1NDcmNTQ3MxYVFAcWFQY3MxYVFAcjJjWaBB4EBAQEKAQEUAQEKAQEBAQeBAQE
BB4EBAQEUAQEBAQeBATWBB4EBAQEHgQEUAQEHgQEBAQeBAQEBFAEBATaBGQEBGQEkBgXERAYFxEQGBcREBgXEREXGBARFxgQERcY
EBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxFhERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEXGBARFxgQEBgXERAY2BAQGBcR
ERcAAv68AeD/tAIwAAkAEwAAAxYVFAcjJjU0NzMWFRQHIyY1NDfwBARQBATwBARQBAQCMBAYFxERFxgQEBgXEREXGBAAAf8MAeD/
ZAIwAAkAAAMWFRQHIyY1NDegBARQBAQCMBAYFxERFxgQAAH+5AHg/2QCgAATAAACFRQHIyY1NDcjJjU0NzMWFRQHM5wEUAQEKAQE
UAQEKAIgGBcRERcYEBEXGBAQGBcRAAAB/wwB4P+MAoAAEwAAAgcjFhUUByMmNTQ3MyY1NDczFhV0BCgEBFAEBCgEBFAEAkEREBgX
EREXGBARFxgQEBgAAf7GAeD/qgKAAB0AAAIVFAcjJjU0NyMWFRQHIyY1NDczJjU0NzMWFRQHM1YEUAQEPAQEUAQEMgQEeAQEMgIg
GBcRERcYEBAYFxERFxgQERcYEBAYFxEAAAL+xgHg/6oC0AAdACcAAAIHIxYVFAcjJjU0NyMmNTQ3MyY1NDczFhUUBzMWFQcmNTQ3
IxYVFAdWBB4EBKAEBB4EBB4EBKAEBB4EQAQEZAQEAkEREBgXEREXGBARFxgQERcYEBAYFxEQGCgRFxgQEBgXEQAAA/6UAeD/3AKA
ABMAHQAnAAACBzMWFRQHIyY1NDcjJjU0NzMWFRYHIyY1NDczFhUENzMWFRQHIyY1sARQBAR4BARQBAR4BIwEPAQEPAT+uAQ8BAQ8
BAJBERAYFxERFxgQERcYEBAYFxERFxgQEBg4EBAYFxERFwAAAf68AeD/tAIwAAkAAAIVFAcjJjU0NzNMBPAEBPACIBgXEREXGBAA
AAH/Av9g/3gAAAATAAAGByMWFRQHIyY1NDczJjU0NzMWFYgEHgQEUAQEMgQEPAQ/ERAYFxERFxgQERcYEBAYAP//AJwB4AEcAoAA
AwDNAZAAAP//AJL/YAEIAAAAAwDSAZAAAP//AEwB4AFEAjAAAwDKAZAAAP//AHQB4AD0AoAAAwDMAZAAAP//AEwB4AFEAjAAAwDR
AZAAAAAAAAEAAADYANkABgDGAAgAAQAAAAAAAAAAAAAAAAADAAEAAAAAAAAAhwEoATQB1gJ3AoMCjwMfA6MEEwQfBKoFOwWPBf0G
CQZ4BuYHOQe+CE4IogkQCRwJiwn5CloK9AtFDBkMwwzPDWAOCw4XDsMPbhA7EEcQshEuEdcSZxLXEyoTuhRjFG8VGRXDFlMXKRfN
GEYY2Bk0GY4aAhoOGoMa9xsDGw8bhRwWHHIcfh0QHbIeAx5tHnke5B9OH6cgLiC3IQQhQyGcIagiAiIOImciySNLI54kSiS/JMsl
NSW5JcUmSibOJ10naSftKHso/ilXKZsqNSqLKwErkSudLC4svi0nLcIuNi7BL2UwCjBRMOoxazIAMlsyyTNaM640FTSGNP81fTX7
Nlc23jdbN5k33jgjONo5tTp8OvM7ZzuFO6Q72TwNPDs8hDzNPUo9aD3FPiE+Vj50PsA/NT9KP14/ykA3QJ1BA0F1QeZB+0IQQiVC
LUKvQzBDk0P2RCtEYESURJRFGEWyRmhG5UdkR6VII0hlSIZI40k/SY5J9UoeSuBLkUxKTQdNh05+T5hP5lBKUKdQ91GTUbRRyFHo
UghSNFJuUqpSvlLeUudS8FL5UwJTCwAAAAEAAAACAAA4wPUaXw889QADA+gAAAAA06JuqwAAAADUIove/mz/EAGoBBAAAAAHAAIA
AAAAAAABkAAAAZAAAAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkABMAZAATAGQ
AEwBkABMAZAATAGQAEwBkAAkAZAAJAGQAGABkABgAZAAVgGQAEwBkABgAZAAJAGQACQBkABMAZAAJAGQACQBkAAkAZAAJAGQACQB
kAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAk
AZAAJAGQACQBkAA4AZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAA
JAGQACQBkAAkAZAATAGQACQBkAAkAZAATAGQAEwBkABMAZAATAGQAEwBkABMAZAATAGQAEwBkAAkAZAATAGQACQBkAAkAZAAJAGQ
ACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQBkAAkAZAAOAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQACQB
kAAkAZAAJAGQADgBkAAaAZAAGgGQABoBkAA4AZAAJAGQACQBkAAkAZAAVgGQAFYBkABCAZAATAGQADgBkAA4AZAAJAGQADgBkAAz
AZAAOAGQAEIBkAAzAZAAfgGQAHQBkAB0AZAAFQGQABUBkAAkAZAAOAGQACQBkACSAZAAdAGQAJIBkACSAZAAJAGQAJwBkACcAZAA
JAGQAJIBkAAkAZAAJAGQAEwBkACcAZAAkgGQACQBkAAkAZAAnAGQAEwBkABMAZAAiAGQAHQBkAB+AZAAbwGQACQBkAAkAZAAYAGQ
AGABkAAkAZAAJAGQAEIBkABCAZAAkgGQAJIBkACSAZAAAAGQACQBkAAkAZAAJAGQACQBkAAkAZAAOAGQAEwBkABMAZAATAGQAFEB
kABRAZAAOAGQACQBkAAkAZAAJAGQACQBkAAkAZAAJAGQADgBkAAkAZAAJAGQAFYBkACcAZAAnAGQACQBkP/8AAD+vAAA/wwAAP7k
AAD/DAAA/sYAAP7GAAD+lAAA/rwAAP8CAZAAnACSAEwAdABMAAEAAAMg/zgAAAGQ/mz/6AGoAAEAAAAAAAAAAAAAAAAAAADUAAQB
kAGQAAUAAAKKAlgAAABLAooCWAAAAV4AMgDwAAAAAAUJAAAAAAAAAAAAAwAAAAAAAAAAAAAAAFVLV04AwAAgICYDIP84AAAEEADw
AAAAAQAAAAABkAIwAAAAIAADAAAAAgAAAAMAAAAUAAMAAQAAABQABAHOAAAAFAAQAAMABAAvADkAfgD/IBQgGiAdICIgJv//AAAA
IAAwADoAoCATIBggHCAiICb//wAAAEsAAAAAAADglOCO4GzgawABABQAAAAwALgBdgAAAAAAAAAAAAAAAQCSAJgAlACyAL4AwACZ
AKIAowCLALUAkACmAJUAmwCPAJoAugC4ALkAlgC/AAIACgALAA0ADwAUABUAFgAXABwAHQAeAB8AIAAiACkAKwAsAC0ALgAvADQA
NQA2ADcAOQCgAIwAoQDIAJwA1gA6AEIAQwBFAEcATABNAE4ATwBWAFcAWABZAFoAXABjAGUAZgBnAGkAagBvAHAAcQByAHUAngDG
AJ8AvACvAJMAsACzALEAtADHAMIA1QDDAHkAqAC9AKcAxADXAMUAuwCGAIcA0wDJAMEAjQDUAIUAegCpAIkAiACKAJcABgADAAQA
CAAFAAcACQAMABMAEAARABIAGwAYABkAGgAOACEAJgAjACQAKAAlALYAJwAzADAAMQAyADgAKgBoAD4AOwA8AEAAPQA/AEEARABL
AEgASQBKAFUAUQBSAFMARgBbAGAAXQBeAGIAXwC3AGEAbgBrAGwAbQBzAGQAdAClAKQAAAAAAAcAWgADAAEECQAAAIAAAAADAAEE
CQABAAoAgAADAAEECQACAA4AigADAAEECQADADAAmAADAAEECQAEABoAyAADAAEECQAFABoA4gADAAEECQAGABoA/ABDAG8AcAB5
AHIAaQBnAGgAdAAgADIAMAAxADEALAAgAFQAaABlACAAVgBUADMAMgAzACAAUAByAG8AagBlAGMAdAAgAEEAdQB0AGgAbwByAHMA
IAAoAHAAZQB0AGUAcgAuAGgAdQBsAGwAQABvAGkAawBvAGkALgBjAG8AbQApAFYAVAAzADIAMwBSAGUAZwB1AGwAYQByADIALgAw
ADAAMAA7AFUASwBXAE4AOwBWAFQAMwAyADMALQBSAGUAZwB1AGwAYQByAFYAVAAzADIAMwAgAFIAZQBnAHUAbABhAHIAVgBlAHIA
cwBpAG8AbgAgADIALgAwADAAMABWAFQAMwAyADMALQBSAGUAZwB1AGwAYQByAAMAAAAAAAD/hAAyAAAAAQAAAAAAAAAAAAAAAAAA
AAAAAQAB//8ADwABAAAADAAAAAAAAAACAAUAAgBPAAEAUQB1AAEAdgB4AAIAeQB6AAEAsADJAAEAAAABAAAACgAcAB4AAURGTFQA
CAAEAAAAAP//AAAAAAAAAAEAAAAKAJYA+AACREZMVAAObGF0bgASACwAAAAoAAZBWkUgADJDQVQgAD5DUlQgAEpLQVogAFZUQVQg
AGJUUksgAG4AAP//AAIAAAABAAD//wADAAAAAQACAAD//wADAAAAAQADAAD//wADAAAAAQAEAAD//wADAAAAAQAFAAD//wADAAAA
AQAGAAD//wADAAAAAQAHAAhmcmFjADJsaWdhADhsb2NsAD5sb2NsAERsb2NsAEpsb2NsAFBsb2NsAFZsb2NsAFwAAAABAAYAAAAB
AAcAAAABAAUAAAABAAAAAAABAAQAAAABAAEAAAABAAIAAAABAAMACQAUAFgAWABYAFgAWABsAKgA2AAGAAAAAgAKACQAAwABABQA
AQAuAAEAFAABAAAACAABAAEAWAADAAEAGgABABQAAQAaAAEAAAAIAAEAAQCNAAEAAQAeAAEAAAABAAgAAQAGAAUAAQABAE8ABAAA
AAEACAABACwAAgAKACAAAgAGAA4AiAADAJsAfQCJAAMAmwB/AAEABACKAAMAmwB/AAEAAgB8AH4ABAAAAAEACAABACIAAQAIAAMA
CAAOABQAdgACAEwAdwACAE8AeAACAFgAAQABAEwAAQAAAAEACAACABAABQB5AHoAeQB6AJ0AAQAFAAIAIgA6AFwAjQAA
""",
    "ShareTechMono-Regular.ttf": """
AAEAAAAQAQAABAAAR1BPU1JIc7MAAKZsAAAAgkdTVUKQRaGtAACm8AAAAhZPUy8ybr6BfQAAAYgAAABgY21hcEHAeU8AAAWEAAAF
gmN2dCABghKsAAAXBAAAAEJmcGdtcfkobwAACwgAAAtvZ2FzcAAAABAAAKZkAAAACGdseWbCk+Z2AAAZZAAAhfRoZWFkBNjLPgAA
AQwAAAA2aGhlYQdMA+sAAAFEAAAAJGhtdHio2lOnAAAB6AAAA5xsb2NhDXXraQAAF0gAAAIabWF4cAIrDFQAAAFoAAAAIG5hbWVf
xYeHAACfWAAABBBwb3N0UNuVKAAAo2gAAAL7cHJlcJo8uCoAABZ4AAAAiwABAAAAAQDFw/q9Gl8PPPUAAwPoAAAAANAki+wAAAAA
0cj7Pf/i/xUD8gNXAAAABwACAAAAAAAAAAEAAAN1/w4AAAQ4/+L/4gPyAAEAAAAAAAAAAAAAAAAAAADCAAEAAAEMAIkACAAmAAMA
AgAiADIAcwAAAHwLcAACAAEAAwIgAZAABQAAAooCWAAAAEsCigJYAAABXgAyASwAAAILBQkFAAACAAQAAAADAAAAAAAAAAAAAAAA
VUtXTgBAAAD7AgN1/w4AAAN1APIgAAABAAAAAAH0ArwAAAAgAAMCHABUAhwAAAIcAAACHAAtAhwALQIcAC0CHAAtAhwALQIcAC0C
HAAtAhwADwIcAFUCHABtAhwAbQIcAFICHAAiAhwAZwIcAGcCHABnAhwAZwIcAGcCHABqAhwAXgIcAFACHACIAhwAiAIcAHICHACA
AhwAiAIcAH0CHABRAhwAcwIcAEsCHAA1AhwATwIcAE8CHABGAhwARgIcAEYCHABGAhwARgIcAEYCHABGAhwAGgIcAFoCHABeAhwA
RgIcAEYCHABkAhwAZAIcAEMCHABNAhwATQIcAE0CHABNAhwATQIcADgCHAAzAhwAMwIcACwCHAAsAhwALAIcAFcCHABXAhwAQwIc
AEMCHABDAhwAQwIcAEMCHABDAhwAQwIcABkCHABXAhwAeAIcAHgCHABUAhwAXgIcAF4CHABeAhwAXgIcAF4CHABeAhwAYwIcAEgC
HABXAhwAUAIcAFACHABQAhwAUAIcAFACHABQAhwAiQIcAGACHABGAhwARgIcACsCHABXAhwAVwIcAFMCHABTAhwAUwIcAFMCHABP
AhwAUwIcAFMCHAAaAhwAVwIcAGECHABUAhwAZAIcAHYCHABvAhwATgIcAGsCHABdAhwAXQIcAF0CHABdAhwAXQIcAEACHAAjAhwA
QQIcAD8CHAA/AhwAPwIcAGoCHACnBDgAYwQ4AGMCHAB9AhwAhAIcADICHAAtAhwAZwIcAGcCHAAwAhwAXAIcAGYCHABpAhwAeQIc
AFcCHAB3AhwAYQIcAGgCHABWAhwAYAIcAGkCHAALAhwACwIcABUCHACSAhwAlgIcAJ4CHABGAhwAMgIcANwCHADDAhwA3AIcAMgC
HAAoAhwA3AIcANwCHABQAhwA3AIcAH4CHACDAhwAkQIcAN4CHADIAhwAMgIcABICHACTAhwAkwIcAKUCHAClAhwAqgIcAKoCHAAg
AhwAUAIcAGsCHABrAhwAcAIcAHACHADAAhwAwAIcAIICHACCAhwAggIcAM8CHADPAhwAzwIcAAAAAAAAAhwAWQB4AA4AaQBPAFwA
NwBkAGQAUADcAFAAUABQAA0ARQBQAFAA3AA+AFAAYQBQAGAADAAJAE8AUAArABYATQBGANwA3ABLAFT/4gAAAAAAAAAoACgAFABK
AFsADgDmAOYAHQBTAD4AGQAZAHEADgCBAEcAWQBZAKgAgABxALcAcQB/AOMAowBNAI4AtQCyAIIAjgBtAB4AAAADAAAAAwAAABwA
AQAAAAADigADAAEAAAAcAAQDbgAAAG4AQAAFAC4AAAANAC8AOQB+AP8BMQFCAVMBYQF4AX4BkgLHAskC3QO8A8AgFCAaIB4gIiAm
IDAgOiBEIKwhIiEmIgIiBiIPIhIiFSIaIh4iJCInIisiSCJgImUixSWvJcomICY8JkAmQiZgJmMmZiZr+wL//wAAAAAADQAgADAA
OgCgATEBQQFSAWABeAF9AZICxgLJAtgDvAPAIBMgGCAcICAgJiAwIDkgRCCsISIhJiICIgYiDyIRIhUiGSIeIiQiJyIrIkgiYCJk
IsUlryXKJiAmOiZAJkImYCZjJmUmavsB//8AwP/0AAAAWAAAAAD/JQAAAAAAAP7FAAD/MwAA/kAAAPzK/McAAOCkAAAAAOB54Krg
fuBO4BXf1d9e3tbefd7OAADeywAA3rHevt6s3qXegN53AADeBts12xnaxdqs2sraydqJ2ofahtqDBX4AAQAAAAAAagAAAIYBDgAA
AcoBzAHOAAABzgAAAc4AAAHOAAAAAAHUAAAB1AHYAAAAAAAAAAAAAAAAAAAAAAAAAAAByAAAAcgAAAAAAAAAAAAAAAABvgAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAoACmAKIAxADZAPIApwCvALAAmQDbAJ4AswCjAKkAnQCoANEAzADNAKQA8QADAAsADAAO
ABAAFQAWABcAGAAdAB4AHwAhACIAJAAsAC4ALwAwADIAMwA4ADkAOgA7AD4ArQCaAK4A+QCqAQMAQABIAEkASwBNAFIAUwBUAFUA
WwBcAF0AXwBgAGIAagBsAG0AbgBxAHIAdwB4AHkAegB9AKsA7wCsAMkAvwChAMIAxgDDAMcA8AD2AQEA9ACBALUA1AC0APUBBQD4
ANwAlwCYAPwAhQDzAJsA/wCWAIIAtgCUAJMAlQClAAcABAAFAAkABgAIAAoADQAUABEAEgATABwAGQAaABsADwAjACgAJQAmACoA
JwDWACkANwA0ADUANgA8AC0AcABEAEEAQgBGAEMARQBHAEoAUQBOAE8AUABaAFcAWABZAEwAYQBmAGMAZABoAGUAygBnAHYAcwB0
AHUAewBrAHwAIABeACsAaQAxAG8APwB+AQAA/gD9AQIBBwEGAQgBBACyALEAugC7ALkA+gD7AJwA3wDVAOEA3gDSAM4ABgH4AAAA
CQD3AAIAAAAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAKAApgCiAMQA2QDyAKcArwCwAJkA2wCe
ALMAowCpAIgAiQCKAIsAjACNAI4AjwCQAJEAnQCoANEAzADNAKQA8QADAAsADAAOABAAFQAWABcAGAAdAB4AHwAhACIAJAAsAC4A
LwAwADIAMwA4ADkAOgA7AD4ArQCaAK4A+QCqAQMAQABIAEkASwBNAFIAUwBUAFUAWwBcAF0AXwBgAGIAagBsAG0AbgBxAHIAdwB4
AHkAegB9AKsA7wCsAMkAAAAGAAgADQARACMAJwA2AEEARABCAEMARgBFAEoATgBRAE8AUABXAFoAWABZAGEAYwBmAGQAZQBoAHMA
dgB0AHUA+gD4AMIAxgD2AJwA8wBwAPUA9AD3APwBAQDXAAoAKQDPANwA0gDOAMcAhQDYAN8A3QCHANAAgQCCAAAARwBnAKUAoQDU
AN4AxQDIAIMAtQC2AJ8AvwAHAAkAKgArAGkAsgCxALoAuwC8AL0AygDjAHwAPQCSAMEAtwC4AH8AgAD7AJsAvgC5ANoABQASAAQA
EwAUABkAGgAbABwAJQAmAAAAKAA0ADUANwBWAQABCAEFAP0BAgEHAP8BBAEGAP4AALAALCCwAFVYRVkgIEuwDlFLsAZTWliwNBuw
KFlgZiCKVViwAiVhuQgACABjYyNiGyEhsABZsABDI0SyAAEAQ2BCLbABLLAgYGYtsAIsIGQgsMBQsAQmWrIoAQpDRWNFUltYISMh
G4pYILBQUFghsEBZGyCwOFBYIbA4WVkgsQEKQ0VjRWFksChQWCGxAQpDRWNFILAwUFghsDBZGyCwwFBYIGYgiophILAKUFhgGyCw
IFBYIbAKYBsgsDZQWCGwNmAbYFlZWRuwAStZWSOwAFBYZVlZLbADLCBFILAEJWFkILAFQ1BYsAUjQrAGI0IbISFZsAFgLbAELCMh
IyEgZLEFYkIgsAYjQrEBCkNFY7EBCkOwAmBFY7ADKiEgsAZDIIogirABK7EwBSWwBCZRWGBQG2FSWVgjWSEgsEBTWLABKxshsEBZ
I7AAUFhlWS2wBSywB0MrsgACAENgQi2wBiywByNCIyCwACNCYbACYmawAWOwAWCwBSotsAcsICBFILALQ2O4BABiILAAUFiwQGBZ
ZrABY2BEsAFgLbAILLIHCwBDRUIqIbIAAQBDYEItsAkssABDI0SyAAEAQ2BCLbAKLCAgRSCwASsjsABDsAQlYCBFiiNhIGQgsCBQ
WCGwABuwMFBYsCAbsEBZWSOwAFBYZVmwAyUjYUREsAFgLbALLCAgRSCwASsjsABDsAQlYCBFiiNhIGSwJFBYsAAbsEBZI7AAUFhl
WbADJSNhRESwAWAtsAwsILAAI0KyCwoDRVghGyMhWSohLbANLLECAkWwZGFELbAOLLABYCAgsAxDSrAAUFggsAwjQlmwDUNKsABS
WCCwDSNCWS2wDywgsBBiZrABYyC4BABjiiNhsA5DYCCKYCCwDiNCIy2wECxLVFixBGREWSSwDWUjeC2wESxLUVhLU1ixBGREWRsh
WSSwE2UjeC2wEiyxAA9DVVixDw9DsAFhQrAPK1mwAEOwAiVCsQwCJUKxDQIlQrABFiMgsAMlUFixAQBDYLAEJUKKiiCKI2GwDioh
I7ABYSCKI2GwDiohG7EBAENgsAIlQrACJWGwDiohWbAMQ0ewDUNHYLACYiCwAFBYsEBgWWawAWMgsAtDY7gEAGIgsABQWLBAYFlm
sAFjYLEAABMjRLABQ7AAPrIBAQFDYEItsBMsALEAAkVUWLAPI0IgRbALI0KwCiOwAmBCIGCwAWG1EBABAA4AQkKKYLESBiuwcisb
IlktsBQssQATKy2wFSyxARMrLbAWLLECEystsBcssQMTKy2wGCyxBBMrLbAZLLEFEystsBossQYTKy2wGyyxBxMrLbAcLLEIEyst
sB0ssQkTKy2wHiwAsA0rsQACRVRYsA8jQiBFsAsjQrAKI7ACYEIgYLABYbUQEAEADgBCQopgsRIGK7ByKxsiWS2wHyyxAB4rLbAg
LLEBHistsCEssQIeKy2wIiyxAx4rLbAjLLEEHistsCQssQUeKy2wJSyxBh4rLbAmLLEHHistsCcssQgeKy2wKCyxCR4rLbApLCA8
sAFgLbAqLCBgsBBgIEMjsAFgQ7ACJWGwAWCwKSohLbArLLAqK7AqKi2wLCwgIEcgILALQ2O4BABiILAAUFiwQGBZZrABY2AjYTgj
IIpVWCBHICCwC0NjuAQAYiCwAFBYsEBgWWawAWNgI2E4GyFZLbAtLACxAAJFVFiwARawLCqwARUwGyJZLbAuLACwDSuxAAJFVFiw
ARawLCqwARUwGyJZLbAvLCA1sAFgLbAwLACwAUVjuAQAYiCwAFBYsEBgWWawAWOwASuwC0NjuAQAYiCwAFBYsEBgWWawAWOwASuw
ABa0AAAAAABEPiM4sS8BFSotsDEsIDwgRyCwC0NjuAQAYiCwAFBYsEBgWWawAWNgsABDYTgtsDIsLhc8LbAzLCA8IEcgsAtDY7gE
AGIgsABQWLBAYFlmsAFjYLAAQ2GwAUNjOC2wNCyxAgAWJSAuIEewACNCsAIlSYqKRyNHI2EgWGIbIVmwASNCsjMBARUUKi2wNSyw
ABawBCWwBCVHI0cjYbAJQytlii4jICA8ijgtsDYssAAWsAQlsAQlIC5HI0cjYSCwBCNCsAlDKyCwYFBYILBAUVizAiADIBuzAiYD
GllCQiMgsAhDIIojRyNHI2EjRmCwBEOwAmIgsABQWLBAYFlmsAFjYCCwASsgiophILACQ2BkI7ADQ2FkUFiwAkNhG7ADQ2BZsAMl
sAJiILAAUFiwQGBZZrABY2EjICCwBCYjRmE4GyOwCENGsAIlsAhDRyNHI2FgILAEQ7ACYiCwAFBYsEBgWWawAWNgIyCwASsjsARD
YLABK7AFJWGwBSWwAmIgsABQWLBAYFlmsAFjsAQmYSCwBCVgZCOwAyVgZFBYIRsjIVkjICCwBCYjRmE4WS2wNyywABYgICCwBSYg
LkcjRyNhIzw4LbA4LLAAFiCwCCNCICAgRiNHsAErI2E4LbA5LLAAFrADJbACJUcjRyNhsABUWC4gPCMhG7ACJbACJUcjRyNhILAF
JbAEJUcjRyNhsAYlsAUlSbACJWG5CAAIAGNjIyBYYhshWWO4BABiILAAUFiwQGBZZrABY2AjLiMgIDyKOCMhWS2wOiywABYgsAhD
IC5HI0cjYSBgsCBgZrACYiCwAFBYsEBgWWawAWMjICA8ijgtsDssIyAuRrACJUZSWCA8WS6xKwEUKy2wPCwjIC5GsAIlRlBYIDxZ
LrErARQrLbA9LCMgLkawAiVGUlggPFkjIC5GsAIlRlBYIDxZLrErARQrLbA+LLA1KyMgLkawAiVGUlggPFkusSsBFCstsD8ssDYr
iiAgPLAEI0KKOCMgLkawAiVGUlggPFkusSsBFCuwBEMusCsrLbBALLAAFrAEJbAEJiAuRyNHI2GwCUMrIyA8IC4jOLErARQrLbBB
LLEIBCVCsAAWsAQlsAQlIC5HI0cjYSCwBCNCsAlDKyCwYFBYILBAUVizAiADIBuzAiYDGllCQiMgR7AEQ7ACYiCwAFBYsEBgWWaw
AWNgILABKyCKimEgsAJDYGQjsANDYWRQWLACQ2EbsANDYFmwAyWwAmIgsABQWLBAYFlmsAFjYbACJUZhOCMgPCM4GyEgIEYjR7AB
KyNhOCFZsSsBFCstsEIssDUrLrErARQrLbBDLLA2KyEjICA8sAQjQiM4sSsBFCuwBEMusCsrLbBELLAAFSBHsAAjQrIAAQEVFBMu
sDEqLbBFLLAAFSBHsAAjQrIAAQEVFBMusDEqLbBGLLEAARQTsDIqLbBHLLA0Ki2wSCywABZFIyAuIEaKI2E4sSsBFCstsEkssAgj
QrBIKy2wSiyyAABBKy2wSyyyAAFBKy2wTCyyAQBBKy2wTSyyAQFBKy2wTiyyAABCKy2wTyyyAAFCKy2wUCyyAQBCKy2wUSyyAQFC
Ky2wUiyyAAA+Ky2wUyyyAAE+Ky2wVCyyAQA+Ky2wVSyyAQE+Ky2wViyyAABAKy2wVyyyAAFAKy2wWCyyAQBAKy2wWSyyAQFAKy2w
WiyyAABDKy2wWyyyAAFDKy2wXCyyAQBDKy2wXSyyAQFDKy2wXiyyAAA/Ky2wXyyyAAE/Ky2wYCyyAQA/Ky2wYSyyAQE/Ky2wYiyw
NysusSsBFCstsGMssDcrsDsrLbBkLLA3K7A8Ky2wZSywABawNyuwPSstsGYssDgrLrErARQrLbBnLLA4K7A7Ky2waCywOCuwPCst
sGkssDgrsD0rLbBqLLA5Ky6xKwEUKy2wayywOSuwOystsGwssDkrsDwrLbBtLLA5K7A9Ky2wbiywOisusSsBFCstsG8ssDorsDsr
LbBwLLA6K7A8Ky2wcSywOiuwPSstsHIsswkEAgNFWCEbIyFZQiuwCGWwAyRQeLABFTAtAABLsMhSWLEBAY5ZugABCAAIAGNwsQAF
QrMAGgIAKrEABUK1IAENCAIIKrEABUK1IQAXBgIIKrEAB0K5CEADgLECCSqxAAlCswBAAgkqsQMARLEkAYhRWLBAiFixA2REsSYB
iFFYugiAAAEEQIhjVFixAwBEWVlZWbUhAA8IAgwquAH/hbAEjbECAEQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABVAFUASwBLArwA
AAK8AfQAAP9WA3X/DgK8AAACvAH0AAD/VgN1/w4AMgAyAAAAAAAYABgAGABKAI4A1AEhAWUBugIeAmACqwLcA2kDmwPcBAkEYwS/
BQcFYQWKBcQF7wYZBm4GxQcLB2AHiAeyB9EH/ggsCFAIpwjnCU8JugoVCn0K1wtKC5oLzwwHDFIMkwzYDUwNbQ2bDdsOHQ5mDqYO
xg73DyAPRw+AD8IP6xBAELQRXRIHEp0TRhPyFKQVHxVhFZEWHRZgFtkXHBeOGAMYYhjUGRMZdBmsGeEaBRpSGqIa4BstG2kbkhu8
G/YcQhx4HOAdHR2EHe0eRR6sHwYfeB/oIDAgeCC4IQYhSSG9IjIiZyKbIvkjWCOoJAYkJiRWJH4ktSUXJWklkSXlJfEl/SaBJs0m
5yccJ1YneyeeJ+goEihLKJIoxCj8KUMpZCnPKhYqLSqQKvMrvSvlLBkscyynLL0s2Sz1LR4tOy1pLZMtuy4ILiEuZS6pLtAu7C8W
Ly0vRi+SL98wBTArMFMwezCYMLUw0jDvMRgxQTFeMXsxsDHXMf4yGjI2Ml4yXjJeMsYzMDOmNBc0WTSTNNQ1ITVgNZk1qTXVNes2
CDZhNoM2mTa3Nsc26Db4NxQ3NjdsN+o4lDi7OPI5FTksOUg5WTlpOXk5lzmvOnk64Ts2O4g7tzv5PB48MDxdPI88qDzVPVA9oj3m
Pks+vj8wP1Q/kD+vP9lAEEA1QHBAmUDlQQ5BNkFPQXRBlkGyQedCI0JgQnxCtUL6AAAAAgBUAAAByAK8AAMABwAItQUEAgACLSsh
IREhBREzEQHI/owBdP7X3gK8SP3UAiwAAgAtAAAB7wK8AAcACgAxQC4JAQQBAUcGAQQFAQMABANfAAEBD0gCAQAAEABJCAgAAAgK
CAoABwAHERERBwUXKzcHIxMzEyMnJwMDrSdZqXOmWSYQUlGxsQK8/USxSwF+/oIAAAMALQAAAe8DSQADAAsADgBGQEMNAQYDAUcH
AQEAAW8AAAMAbwkBBggBBQIGBV8AAwMPSAQBAgIQAkkMDAQEAAAMDgwOBAsECwoJCAcGBQADAAMRCgUVKwEHIzcDByMTMxMjJycD
AwGMhFNocCdZqXOmWSYQUlkDSXNz/WixArz9RLFLAX7+ggADAC0AAAHvA00ABgAOABEAQ0BABgEAARABBwQCRwABAAFvAgEABABv
CQEHCAEGAwcGXwAEBA9IBQEDAxADSQ8PBwcPEQ8RBw4HDhERExEREAoFGisTIzczFyMnAwcjEzMTIycnAwPRYG5cb2E8YCdZqXOm
WSYQUlEC2nNzQP2XsQK8/USxSwF+/oIAAAQALQAAAe8DLwADAAcADwASAE9ATBEBCAUBRwoDCQMBAgEABQEAXgwBCAsBBwQIB18A
BQUPSAYBBAQQBEkQEAgIBAQAABASEBIIDwgPDg0MCwoJBAcEBwYFAAMAAxENBRUrExUjNSEXIzUDByMTMxMjJycDA9lVARILVZ8n
WalzplkmEFJRAy9QUFBQ/YKxArz9RLFLAX7+ggADAC0AAAHvA0kAAwALAA4ARkBDDQEGAwFHBwEBAAFvAAADAG8JAQYIAQUCBgVf
AAMDD0gEAQICEAJJDAwEBAAADA4MDgQLBAsKCQgHBgUAAwADEQoFFSsTFyMnEwcjEzMTIycnAwP9aFOEHydZqXOmWSYQUlEDSXNz
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
IyM1MzIWFRUUFjMVIyY2NTUjIgYVFRQWMzMBmnpWaGIfCxcmWUxFknwmJZuRV1MXGzdgEH0aIhgaWQLLh4f9NRAOHj1KFJsyHyJL
TFTwDw9GSxwWbiIaMhoYAAMAQwAAAdgCywAGACMAMADOQAoGAQABCQEDBwJHS7AhUFhALwIBAAEGAQAGbQAEAAkHBAlgAAEBD0gA
BQUGWAAGBhJIDAoCBwcDWAsIAgMDEANJG0uwLVBYQCwAAQABbwIBAAYAbwAEAAkHBAlgAAUFBlgABgYSSAwKAgcHA1gLCAIDAxAD
SRtANwABAAFvAgEABgBvAAQACQoECWAABQUGWAAGBhJIDAEKCgNYCwgCAwMQSAAHBwNYCwgCAwMQA0lZWUAZJCQHByQwJC4pJwcj
ByIVISMkNREREA0FHCsTIzczFyMnEiYnBiMjIiY1NTQzMzU0JiMjNTMyFhUVFBYzFSMmNjU1IyIGFRUUFjMzwmBwWHFhPJYfCxcm
WUxFknwmJZuRV1MXGzdgEH0aIhgaWQJEh4dL/XEQDh49ShSbMh8iS0xU8A8PRkscFm4iGjIaGAAEAEMAAAHYAqMAAwAHACQAMQCk
tQoBBAgBR0uwLVBYQC0NAwwDAQIBAAcBAF4ABQAKCAUKYAAGBgdYAAcHEkgPCwIICARYDgkCBAQQBEkbQDgNAwwDAQIBAAcBAF4A
BQAKCwUKYAAGBgdYAAcHEkgPAQsLBFgOCQIEBBBIAAgIBFgOCQIEBBAESVlAKiUlCAgEBAAAJTElLyooCCQIIyIhHBoZFxQSDgsE
BwQHBgUAAwADERAFFSsTFSM1IRUjNRImJwYjIyImNTU0MzM1NCYjIzUzMhYVFRQWMxUjJjY1NSMiBhUVFBYzM79VAR1VYh8LFyZZ
TEWSfCYlm5FXUxcbN2AQfRoiGBpZAqNQUFBQ/V0QDh49ShSbMh8iS0xU8A8PRkscFm4iGjIaGAADAEMAAAHYAssAAwAgAC0A0rUG
AQIGAUdLsCFQWEAvAAABBQEABW0AAwAIBgMIYAoBAQEPSAAEBAVYAAUFEkgMCQIGBgJYCwcCAgIQAkkbS7AtUFhALAoBAQABbwAA
BQBvAAMACAYDCGAABAQFWAAFBRJIDAkCBgYCWAsHAgICEAJJG0A3CgEBAAFvAAAFAG8AAwAICQMIYAAEBAVYAAUFEkgMAQkJAlgL
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
tOciIh/cHyIiH9wAAAMAUwAAAcgCywAGABYAJgBitQYBAAEBR0uwIVBYQCMCAQABBAEABG0AAQEPSAAFBQRYAAQEEkgABgYDWAAD
AxADSRtAIAABAAFvAgEABABvAAUFBFgABAQSSAAGBgNYAAMDEANJWUAKNTU1MxEREAcFGysTIzczFyMnEgYjIyImNTU0NjMzMhYV
FSYmIyMiBhUVFBYzMzI2NTXSYHBYcWE8ulNXIVdTU1chV1NVJiU1JSYmJTUlJgJEh4dL/b1MTFS0VExMVLTnIiIf3B8iIh/cAAQA
UwAAAcgCowADAAcAFwAnAD5AOwkDCAMBAgEABQEAXgAGBgVYAAUFEkgABwcEWAAEBBAESQQEAAAkIRwZFBEMCQQHBAcGBQADAAMR
CgUVKxMVIzUhFSM1EgYjIyImNTU0NjMzMhYVFSYmIyMiBhUVFBYzMzI2NTXVVQEdVYBTVyFXU1NXIVdTVSYlNSUmJiU1JSYCo1BQ
UFD9qUxMVLRUTExUtOciIh/cHyIiH9wAAwBPAAABxALLAAMAEwAjAGNLsCFQWEAjAAABAwEAA20GAQEBD0gABAQDWAADAxJIAAUF
AlgAAgIQAkkbQCAGAQEAAW8AAAMAbwAEBANYAAMDEkgABQUCWAACAhACSVlAEgAAIB0YFRANCAUAAwADEQcFFSsBFyMnAAYjIyIm
NTU0NjMzMhYVFSYmIyMiBhUVFBYzMzI2NTUBBWhWegEnU1chV1NTVyFXU1UmJTUlJiYlNSUmAsuHh/2BTExUtFRMTFS05yIiH9wf
IiIf3AADAFP/pwHIAk0AFwAfACcAOkA3EwEEAiEZAgUEBgEABQNHAAMCA28AAQABcAAEBAJYAAICEkgABQUAWQAAABAASSYoEjcR
IQYFGiskBiMjByM3JiY1NTQ2MzMyFzczBxYWFRUEFxMjIgYVFTYnAzMyNjU1AchTVzUbTh8mJlNXIQ4GHE0gJyb+4A9lKSUmyxBk
KSUmTExZaBFHObRUTAFaaRFGObQwEQFKIh/c+BD+tyIf3AADAFMAAAHIArwAGgAqADoAQUA+GgwCAQAZDQICAwJHAAEAAgUBAmAA
AwMAWAAAAA9IAAYGBVgABQUSSAAHBwRYAAQEEARJNTU1NTMmJCEIBRwrEjYzMzIWFxYzMzI2NxUGBiMjIicmIyMiBgc1AAYjIyIm
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
NDYzMzIWFzY2MzMyFhUVFAYjIyImJyYWMzMyNjc1NCYjIyIGFRUkJiMjIgYHFRQWMzMyNjU1+zEmCk8+Pk8KJjMTEy8mCk8+P04K
JjETsB0kBx8eAh8gByYdAWIdJAcfHgIfIAcmHacaTVAUUE0aGxsaTVAPUFIaGy4dIx46ICchJjxmHSMeOiAnISY8AAABAEX/FQHX
AzQAEwAGsw8FAS0rFjY1ETQ2MzMVIyIGFREUBiMjNTPOEzoujn4QEzouiXmgEQ0DSDM7SxEN/LgzO0sAAAEAUABpAcwCNAAGAAaz
BAABLSsBFQUFFSU1Acz+ygE2/oQCNFyIiV6vaQAAAgBQAAABzAI0AAYACgAItQgHBAACLSsBFQUFFSU1ARUhNQHM/swBNP6EAXz+
hAI0XIh1Xptp/spLSwAAAQDcAQQBQAFoAAMABrMBAAEtKwEVIzUBQGQBaGRkAAABAD4AagHdAV4ABQAkQCEAAQIBcAAAAgIAUgAA
AAJWAwECAAJKAAAABQAFEREEBRYrEzUhFSM1PgGfVQETS/SpAAABAFABEwHMAV4AAwAGswEAAS0rARUhNQHM/oQBXktLAAEAYQCJ
AbsB4wALAAazCQMBLSs3Nyc3FzcXBxcHJwdhdHQ4dXU4dHQ4dXXBdXU4dHQ4dXU4dHQAAQBQAEYBzAIrABMABrMNAwEtKyUVIwcj
NyM1MzcjNTM3MwczFSMHAczrQVBBQW9OvetBUEFBb076S2lpS31LaWlLfQACAGAAAAG/ArwAFgAjAAi1HBcQCQItKxI2MzM1NCYj
IzUzMhYVERQGIyMiJjU1NgYVFRQWMzMyNjU1I2BSVmIuJYd9V1taRCNEWnolKiQZJCpsAWxMbR8sTFVV/phSWFhSblUjH5QgLCwg
1gAFAAwAAAIPArwADwAfACMAMwBDAEJAPyABAAMjIQIFACIBBgUDRwADAAAFAwBgAAUABgcFBmAAAgIBWAABAQ9IAAcHBFgABAQQ
BEk1NTU5NTU1MQgFHCsABiMjIiY1NTQ2MzMyFhUVJiYjIyIGFRUUFjMzMjY1NQUVBTUABiMjIiY1NTQ2MzMyFhUVJiYjIyIGFRUU
FjMzMjY1NQEhQDoIOkBAOgg6QEYaGggaGhoaCBoaATT9/QHnQDoIOkBAOgg6QEYaGggaGhoaCBoaAdVBQT8oP0FBPyhMICAZPhkg
IBk+pz5JPv7qQUE/KD9BQT8oTCAgGT4ZICAZPgAHAAkAAAITArwADwAfACMAMwBDAFMAYwBKQEcgAQADIyIhAwUAAkcAAwAABQMA
YAcBBQoBCAkFCGAAAgIBWAABAQ9ICwEJCQRYBgEEBBAESWBdWFVQTTU1NTU5NTU1MQwFHSsABiMjIiY1NTQ2MzMyFhUVJiYjIyIG
FRUUFjMzMjY1NQUVBTUSBiMjIiY1NTQ2MzMyFhUVBAYjIyImNTU0NjMzMhYVFSQmIyMiBhUVFBYzMzI2NTUkJiMjIgYVFRQWMzMy
NjU1AQVAOgg6QEA6CDpARhoaCBoaGhoIGhoBUv3990A6CDpAQDoIOkABDkA6CDpAQDoIOkD+rBoaCBoaGhoIGhoBDhoaCBoaGhoI
GhoB1UFBPyg/QUE/KEwgIBk+GSAgGT6TPkk+/tZBQT8oP0FBPyg/QUE/KD9BQT8oTCAgGT4ZICAZPhkgIBk+GSAgGT4AAQBPAHYB
zAH8AAsAJ0AkAwEBBAEABQEAXgYBBQUCVgACAhIFSQAAAAsACxERERERBwUZKzc1IzUzNTMVMxUjFeaXl1CWlnadS56eS50AAgBQ
AAABzAH8AAsADwA6QDcDAQEEAQAFAQBeCAEFBQJWAAICEkgJAQcHBlYABgYQBkkMDAAADA8MDw4NAAsACxERERERCgUZKzc1IzUz
NTMVMxUjFRcVITXmlpZQlpaW/oSPhEuenkuEREtLAAABACsAAAHxArwAEwAGswoBAS0rEzUhFSMRFBYzMxUjIiY1ESMRIxErAcZb
DhMPMiIzblcCcUtL/fgQDkswNAIN/Y8CcQAAAQAWAAACBQM0AAgABrMEAAEtKzMDMxMTMxUjA3ljWUidsWqlAYj+yQLjS/0XAAEA
Tf9WAc4CvAALAAazCQIBLSsTAzUhFSETAyEVITX1mwFp/u+isQEr/n8BCwFSX0v+mv6VSl4AAQBG/5wB1gMqAAMABrMCAAEtKwEz
ASMBhlD+wFADKvxyAAEA3AEEAUABaAADAAazAQABLSsBFSM1AUBkAWhkZAAAAQDcAQQBQAFoAAMABrMBAAEtKwEVIzUBQGQBaGRk
AAACAEsAAAHRArwABQAJAAi1CAYEAQItKxMTMxMDIzcTAwNLjG6MjHA5bGxsAV4BXv6i/qJGARgBGP7oAAACAFQAAAHIArwAAwAH
AAi1BQQCAAItKyEhESEFETMRAcj+jAF0/tfeArxI/dQCLAAI/+IAFwI6AhsAEgAnADwASABUAF4AcwCIABVAEoeBcmFYVU1JQT0y
KBcTCAAILSsAFhYVFAcVBwcjJyc1JjU0NjYzBhYXFxUnBgYjIiY1NDY3NyY1NDYzIBYVFAcXFhYVFAYjIiYnBzU3NjYzBAYVFBYz
MjY1NCYjMgYVFBYzMjY1NCYjBgYHFTcXNSYmIwcGBiMiJjU0NyciJjU0NjMyFhc3FyQ2MzIWFRQGIwcWFRQGIyImJyc3FwE8Sioc
GiuBKxkdKkou4BEFHDMRFQsLEBIPDQgMCgHrDAgNDhIQCgsVEjMcBRES/rYXFxARFxcRghcXERAYGBBRFQYjIgYUCLMDEA8MDgUN
DhUPDQkXD0gaAUoYCQwPFA4OBg8MDxADRRpHAhsoRis5LTQSS0sSNCw6K0YorCAaEDcVCwgLCwwRAhsJDgkPDwkOCRsCEQwKDAgL
ETMQGiAFFxERGBgRERcXEREYGRARF1IdEB0HBx0QHcEeIhAMBw0YEQwLDgkHJyoKCQ4LDBEYDAgMECIeKyonAAUAAAAAAhwCGwAP
AB8AKwA3AEMAD0AMPTkwLCQgFhAGAAUtKwAWFhUUBgYjIiYmNTQ2NjMOAhUUFhYzMjY2NTQmJiMGFhUUBiMiJjU0NjMyFhUUBiMi
JjU0NjMWBiMiJiczFjMyNzMBV3xJSXxJSnxISHxKO2A4OGA7OmE4OGE6PRsbFRUbGxW1GxsVFRsbFTpOOjpOBywWTU0WLAIbSHxJ
SnxISHxKSXxINDtkOjtkOztkOzpkO2AaExQbGxQTGhoTFBsbFBMa0E1NOVlZAAAEAAAAAAIcAhsADwAbACcANQANQAowLSAcFBAG
AAQtKwAWFhUUBgYjIiYmNTQ2NjMGBhUUFjMyNjU0JiMyBhUUFjMyNjU0JiMWBiMiJicjFhYzMjY3IwFXfElJfElKfEhIfEpsGxsV
FRsbFZUbGxUVGxsVCzUpKTUKKwdSOjpSBysCG0h8SUp8SEh8Skl8SI8aExQbGxQTGhoTFBsbFBMayjAwKThNTTgAAAIAAAAAAhwC
DwAnADMACLUsKCURAi0rABc3FwcWFzMVIwYHFwcnBgcVIzUmJwcnNyYnIzUzNjcnNxc2NzUzFQYGFRQWMzI2NTQmIwFKIFEkURkG
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
ISo1ISgJBRsrEjY3JiY1NTQ2MzMVIyIGFRUUFjMzMhUVFAcWFhUVFAYjIzUzMjY1NTQjIyI1NTYGFRUUFjMzMjY1NTQmIyNxGRMO
EExKfIYYHxoWHYksDhBMSnyGGB8sIYlnEhobNxATFh44AYc5Dg8tHApJQ0sXFigXFpEFRCcKMR0KSUNLFxYoLZEFQRcWLRcbFxYt
GBoAAQAOAMICDAK8ABIABrMLAwEtKwEjJxEjESMRIxEjNSEXNzMRIxEBlTU2QUFEVgFFKihnQgFp/f5cAb/+QQG/O8/P/gYBpAAA
AgCBAY0BmwK9AA8AHwAcQBkAAwAAAwBcAAICAVgAAQEPAkk1NTUxBAUYKwAGIyMiJjU1NDYzMzIWFRUmJiMjIgYVFRQWMzMyNjU1
AZtINxw3SEc4HDhHRSEZHBkhIRkcGSEByj09OkI6PT06QlsgIBZIFiAgFkgAAQBHAaQB1AL4AAYAGUAWBgEAAQFHAAEAAW8CAQAA
ZhEREAMFFysTIxMzEyMDn1iQcotbaAGkAVT+rAEYAAABAFkAAAHCArwACwApQCYAAQEPSAYFAgMDAFYCAQAAEkgABAQQBEkAAAAL
AAsREREREQcFGSsTNTM1MxUzFSMRIxFZjVCMjFABqUvIyEv+VwGpAAABAFkAAAHCArwAEwA3QDQGAQAKCQIHCAAHXgADAw9IBQEB
AQJWBAECAhJIAAgIEAhJAAAAEwATERERERERERERCwUdKzc1MzUjNTM1MxUzFSMVMxUjFSM1WY2NjVCMjIyMUNJLjEvIyEuMS9LS
AAEAqAJEAXgCywADAC5LsCFQWEAMAAABAHACAQEBDwFJG0AKAgEBAAFvAAAAZllACgAAAAMAAxEDBRUrAQcjNwF4elZoAsuHhwAA
AQCAAkkBnALLAA8AQUuwIVBYQA8EAQMAAQMBXQIBAAAPAEkbQBgCAQADAG8EAQMBAQNUBAEDAwFZAAEDAU1ZQAwAAAAPAA0SMhIF
BRcrADY1MxQGIyMiJjUzFBYzMwE2FFI9RxRHPVIUHhQCiSIgM09PMyAiAAEAcQJDAaoCygAGADC1BgEBAAFHS7AjUFhADAABAAFw
AgEAAA8ASRtACgIBAAEAbwABAWZZtREREAMFFysBMwcjJzMXAUpgcFhxYTwCyoeHSwABALf/OAFkADIAEQBiS7ALUFhAIQADAgEA
A2UAAgABAAIBYAUBAAQEAFQFAQAABFkABAAETRtAIgADAgECAwFtAAIAAQACAWAFAQAEBABUBQEAAARZAAQABE1ZQBEBABAOCgkI
BwYEABEBEQYFFCsXMjU0JiMjNTMVMhYVFAYjIzX6IRESMUYuKTMlVYgfEQx+RyMxOiVAAAEAcQJEAaoCywAGADC1BgEAAQFHS7Ah
UFhADAIBAAEAcAABAQ8BSRtACgABAAFvAgEAAGZZtREREAMFFysTIzczFyMn0WBwWHFhPAJEh4dLAAACAH8CUwGcAqMAAwAHACxA
KQUDBAMBAAABUgUDBAMBAQBWAgEAAQBKBAQAAAQHBAcGBQADAAMRBgUVKxMVIzUhFSM11FUBHVUCo1BQUFAAAAEA4wJnATgCvAAD
ABlAFgAAAAFWAgEBAQ8ASQAAAAMAAxEDBRUrARUjNQE4VQK8VVUAAQCjAkQBcwLLAAMALkuwIVBYQAwAAAEAcAIBAQEPAUkbQAoC
AQEAAW8AAABmWUAKAAAAAwADEQMFFSsBFyMnAQtoVnoCy4eHAAACAE0CVgHRAukAAwAHAB1AGgIBAAEBAFICAQAAAVYDAQEAAUoR
EREQBAUYKxMzByMlMwcjtHN/WwERc39bAumTk5MAAAEAjgJdAY0CowADAB9AHAIBAQAAAVICAQEBAFYAAAEASgAAAAMAAxEDBRUr
ARUjNQGN/wKjRkYAAQC1/0wBZgA8AAsAQEuwMVBYQBEAAAEAbwABAQJZAwECAhQCSRtAFgAAAQBvAAECAgFUAAEBAlkDAQIBAk1Z
QAsAAAALAAojEwQFFisWJjU1MxUUFjMzFSPoM0oPEUdZtCU6kZEQDEMAAAIAsgIkAWoC6wAPAB0AIkAfAAEAAgMBAmAAAwAAA1QA
AwMAWAAAAwBMNTQ1MQQFGCsABiMjIiY1NTQ2MzMyFhUVJiMjIgYVFRQWMzMyNTUBajInBicyMSgGKDE5IAYSEBASBiACUCwsKR0q
KysqHUQVER8RFSYfAAEAggJPAZoCvAAaAChAJRoMAgEAGQ0CAgMCRwABAAIBAlwAAwMAWAAAAA8DSTMmJCEEBRgrEjYzMzIWFxYz
MzI2NxUGBiMjIicmIyMiBgc1izgMDwggESoMCggyCQgxCw8FKDQNCgk6CgKgHAwIFB4GSAUcERceBkgAAQCOAl0BjQKjAAMAH0Ac
AgEBAAABUgIBAQEAVgAAAQBKAAAAAwADEQMFFSsBFSM1AY3/AqNGRgACAG0AAAGuAhsAGAAkAAi1HRkLAAItKwAWFhUUBgcVMxUj
FSM1IzUzNSYmNTQ2NjMGBhUUFjMyNjU0JiMBOUorTTtgYDBgYDtOK0ksMD09MDA+PjACGypILDxVCFMoaWkoUghWPCxIKi5AMDA/
PzAwQAACAB4AAAH+AecAHQApAAi1Ih4cCwItKwEHJyY3JwcWFRQGBiMiJiY1NDY2MzIXNycHJyc3FwQGFRQWMzI2NTQmIwH+LAUB
BAKUJStKLCxJKytJLC0qmAEXGGQE1v6bPT0wMD4+MAEGAmMYGAGXLTksRygoRywsSCoZlwIBAQQqC9FAMDA/PzAwQAAAAA0AogAD
AAEECQAAAPoAAAADAAEECQABAB4A+gADAAEECQACAA4BGAADAAEECQADAFYBJgADAAEECQAEAB4A+gADAAEECQAFABoBfAADAAEE
CQAGACoBlgADAAEECQAIAC4BwAADAAEECQAJAC4BwAADAAEECQALACwB7gADAAEECQAMACwB7gADAAEECQANASACGgADAAEECQAO
ADQDOgBDAG8AcAB5AHIAaQBnAGgAdAAgACgAYwApACAAMgAwADEAMgAsACAAQwBhAHIAcgBvAGkAcwAgAFQAeQBwAGUAIABEAGUA
cwBpAGcAbgAsACAAUgBhAGwAcABoACAAZAB1ACAAQwBhAHIAcgBvAGkAcwAgACgAcABvAHMAdABAAGMAYQByAHIAbwBpAHMALgBj
AG8AbQAgAHcAdwB3AC4AYwBhAHIAcgBvAGkAcwAuAGMAbwBtACkALAAgAHcAaQB0AGgAIABSAGUAcwBlAHIAdgBlAGQAIABGAG8A
bgB0ACAATgBhAG0AZQAgACcAUwBoAGEAcgBlACcAUwBoAGEAcgBlACAAVABlAGMAaAAgAE0AbwBuAG8AUgBlAGcAdQBsAGEAcgBS
AGEAbABwAGgATwBsAGkAdgBlAHIAZAB1AEMAYQByAHIAbwBpAHMAOgAgAFMAaABhAHIAZQAgAFQAZQBjAGgAIABNAG8AbgBvADoA
IAAyADAAMQA0AFYAZQByAHMAaQBvAG4AIAAxAC4AMAAwADMAUwBoAGEAcgBlAFQAZQBjAGgATQBvAG4AbwAtAFIAZQBnAHUAbABh
AHIAUgBhAGwAcABoACAATwBsAGkAdgBlAHIAIABkAHUAIABDAGEAcgByAG8AaQBzAGgAdAB0AHAAOgAvAC8AdwB3AHcALgBjAGEA
cgByAG8AaQBzAC4AYwBvAG0AVABoAGkAcwAgAEYAbwBuAHQAIABTAG8AZgB0AHcAYQByAGUAIABpAHMAIABsAGkAYwBlAG4AcwBl
AGQAIAB1AG4AZABlAHIAIAB0AGgAZQAgAFMASQBMACAATwBwAGUAbgAgAEYAbwBuAHQAIABMAGkAYwBlAG4AcwBlACwAIABWAGUA
cgBzAGkAbwBuACAAMQAuADEALgAgAFQAaABpAHMAIABsAGkAYwBlAG4AcwBlACAAaQBzACAAYQB2AGEAaQBsAGEAYgBsAGUAIAB3
AGkAdABoACAAYQAgAEYAQQBRACAAYQB0ADoAIABoAHQAdABwADoALwAvAHMAYwByAGkAcAB0AHMALgBzAGkAbAAuAG8AcgBnAC8A
TwBGAEwAaAB0AHQAcAA6AC8ALwBzAGMAcgBpAHAAdABzAC4AcwBpAGwALgBvAHIAZwAvAE8ARgBMAAIAAAAAAAD/kgAyAAAAAQAA
AAAAAAAAAAAAAAAAAAABDAAAAAIAAwAkAMkAxwBiAK0AYwCuAJAAJQAmAGQAJwDpACgAZQDIAMoAywApACoAKwAsAMwAzQDOAM8A
LQAuAC8A4gAwADEAZgAyANAA0QBnANMAkQCvALAAMwDtADQANQA2AOQANwA4ANQA1QBoANYAOQA6ADsAPADrALsAPQDmAEQAaQBr
AGwAagBuAG0AoABFAEYAbwBHAOoASABwAHIAcwBxAEkASgBLAEwA1wB0AHYAdwB1AE0ATgBPAOMAUABRAHgAUgB5AHsAfAB6AKEA
fQCxAFMA7gBUAFUAVgDlAIkAVwBYAH4AgACBAH8AWQBaAFsAXADsALoAXQDnAMAAwQCdAJ4BAgCfAJcBAwCbABMAFAAVABYAFwAY
ABkAGgAbABwAvAD0APUA9gDxAPIA8wANAD8AwwCHAB0ADwCrAAQAowAGABEAIgCiAAUACgAeABIAQgBeAGAAPgBAAAsADACzALIA
EAEEAKkAqgC+AL8AxQC0ALUAtgC3AMQBBQABAQYAhAC9AAcApgCFAJYApwBhALgBBwAgACEAlQCSAJwAHwCUAQgApADvAPAAjwCY
AAgAxgAOAJMAmgClAJkBCQEKAQsAuQEMAQ0BDgEPARABEQESARMBFAEVARYAXwDoACMACQCIAIsAigCGAIwAgwBBAIIAwgCNANsA
4QDeANgAjgDcAEMA3wDaAOAA3QDZARcBGAEZBURlbHRhB3VuaTAzQkMHdW5pMDBBRAd1bmkwMEEwBEV1cm8HZG90bWF0aApsb2dp
Y2FsYW5kB3VuaTIyMTUHdW5pMjIxOQd1bmkyMjI0B3VuaTI1QUYHdW5pMjYyMAlzbWlsZWZhY2UMaW52c21pbGVmYWNlA3N1bgVz
cGFkZQRjbHViBWhlYXJ0B2RpYW1vbmQLbXVzaWNhbG5vdGUObXVzaWNhbG5vdGVkYmwHdW5pMDJDOQZmZW1hbGUEbWFsZQAAAQAB
//8ADwABAAAACgBCAFwAA0RGTFQAFGdyZWsAIGxhdG4ALAAEAAAAAP//AAEAAAAEAAAAAP//AAEAAQAEAAAAAP//AAEAAgADY3Bz
cAAUY3BzcAAUY3BzcAAUAAAAAQAAAAEABAABAAAAAQAIAAEACgAFAAUACgACAAIAAwA/AAAAgwCEAD0AAAABAAAACgBgAPIAA0RG
TFQAFGdyZWsAKmxhdG4AQAAEAAAAAP//AAYAAAADAAYACQAMAA8ABAAAAAD//wAGAAEABAAHAAoADQAQAAQAAAAA//8ABgACAAUA
CAALAA4AEQASYWFsdABuYWFsdABuYWFsdABuZGxpZwB0ZGxpZwB0ZGxpZwB0ZnJhYwB6ZnJhYwB6ZnJhYwB6bGlnYQCAbGlnYQCA
bGlnYQCAb3JkbgCGb3JkbgCGb3JkbgCGc3VwcwCMc3VwcwCMc3VwcwCMAAAAAQAAAAAAAQAEAAAAAQACAAAAAQAFAAAAAQADAAAA
AQABAAcAEAA+AFYAkgDaANoBAgABAAAAAQAIAAIAFAAHAIEAggCBAIIAlgCXAJgAAQAHAAMAJABAAGIAiQCKAIsAAQAAAAEACAAB
AAYADQABAAMAiQCKAIsABAAAAAEACAABACwAAgAKACAAAgAGAA4AkwADAKkAigCUAAMAqQCMAAEABACVAAMAqQCMAAEAAgCJAIsA
BgAAAAIACgAkAAMAAQAsAAEAEgAAAAEAAAAGAAEAAgADAEAAAwABABIAAQAcAAAAAQAAAAYAAgABAIgAkQAAAAEAAgAkAGIABAAA
AAEACAABABoAAQAIAAIABgAMAH8AAgBVAIAAAgBdAAEAAQBSAAEAAAABAAgAAgAOAAQAgQCCAIEAggABAAQAAwAkAEAAYgAA
""",
}


if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception:
        texto = traceback.format_exc()
        try:
            (Path.home() / "pdf_master_error.log").write_text(texto, encoding="utf-8")
        except Exception:
            pass
        try:
            raiz = tk.Tk()
            raiz.withdraw()
            messagebox.showerror("Error al abrir el programa", texto[-1500:])
        except Exception:
            pass
