# -*- coding: utf-8 -*-
"""
PDF Master: convertir Word y PowerPoint a PDF, unir PDF y herramientas básicas.
UN SOLO ARCHIVO, solo necesita Python. https://github.com/bruneskycoder/pdf-master

Ventanas:
  1. Principal            : Word/PowerPoint -> PDF (LibreOffice), lista con páginas, tamaño e
                            intervalo; unión con página en blanco si es impar y pie de página.
  2. Herramientas PDF     : dividir, comprimir y eliminar páginas con vista de miniaturas.
  3. Carpetas automáticas : carpetas vigiladas que dividen, unen o comprimen solas lo que llegue.

Motor de conversión: LibreOffice (si falta, el programa lo descarga y lo instala).
Motor PDF: PyMuPDF (si falta, el programa lo instala con pip, una sola vez).
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

try:  # arrastrar y soltar (opcional)
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _BaseTk = TkinterDnD.Tk
    HAY_DND = True
except Exception:
    DND_FILES = None
    _BaseTk = tk.Tk
    HAY_DND = False

VERSION = "5.0"
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


OPS = {"info": op_info, "merge": op_merge, "thumbs": op_thumbs, "delete": op_delete,
       "split": op_split, "compress": op_compress}

if __name__ == "__main__":
    cfg = json.load(open(sys.argv[1], encoding="utf-8"))
    try:
        res = OPS[cfg["op"]](cfg)
    except Exception as e:
        res = {"error": str(e)[:400]}
    print("@@JSON@@" + json.dumps(res, ensure_ascii=False))
'''

_PY_PDF = {"exe": None}


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


def python_con_pymupdf(emit, instalar=True):
    """Devuelve un Python que tenga PyMuPDF (instalándolo si hace falta), o None."""
    if _PY_PDF["exe"] and Path(_PY_PDF["exe"]).exists():
        return Path(_PY_PDF["exe"])
    actual = _consola(Path(sys.executable))
    if _tiene_modulo(actual, "pymupdf"):
        _PY_PDF["exe"] = str(actual)
        return actual
    for otro in _otros_pythons():
        if _tiene_modulo(otro, "pymupdf"):
            _PY_PDF["exe"] = str(otro)
            return otro
    if not instalar:
        return None
    emit("log", "Instalando el componente PDF (PyMuPDF, una sola vez; puede tardar un minuto)…")
    detalle = ""
    for extra in ([], ["--user"]):
        try:
            r = subprocess.run(
                [str(actual), "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
                 *extra, "pymupdf"],
                capture_output=True, text=True, errors="replace", timeout=600,
                creationflags=NO_WINDOW)
        except Exception as e:
            detalle = str(e)
            continue
        if r.returncode == 0 and _tiene_modulo(actual, "pymupdf"):
            _PY_PDF["exe"] = str(actual)
            return actual
        detalle = (r.stderr or r.stdout or "")[-300:].strip()
    emit("log", f"No pude instalar PyMuPDF: {detalle}")
    return None


AYUDA_PYMUPDF = ("No pude preparar el componente PDF (PyMuPDF).\n\n"
                 "Solución manual: abrí una consola (cmd) y ejecutá:\n"
                 "python -m pip install pymupdf\n\nDespués volvé a intentar.")


def pdf_op(cfg, emit, timeout=1800, instalar=True):
    """Ejecuta una operación PDF en el Python que tenga PyMuPDF. Devuelve un dict;
    con clave 'error' si algo falló."""
    py = python_con_pymupdf(emit, instalar)
    if not py:
        return {"error": "no_pymupdf"}
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
        ttk.Button(barra, text="Marcar impares", command=lambda: self._marcar(lambda n: n % 2 == 1)).pack(side="left")
        ttk.Button(barra, text="Marcar pares", command=lambda: self._marcar(lambda n: n % 2 == 0)).pack(side="left", padx=4)
        ttk.Button(barra, text="Marcar intervalo…", command=self._intervalo).pack(side="left")
        ttk.Button(barra, text="Desmarcar todas", command=lambda: self._marcar(lambda n: False)).pack(side="left", padx=4)
        self.v_estado = tk.StringVar(value="Generando miniaturas…")
        ttk.Label(barra, textvariable=self.v_estado, style="Muted.TLabel").pack(side="right")

        pie = ttk.Frame(self)
        pie.pack(side="bottom", fill="x", padx=8, pady=6)
        self.v_reemplazar = tk.BooleanVar(value=False)
        ttk.Checkbutton(pie, text="Reemplazar el PDF original (si no, crea “… (sin páginas).pdf”)",
                        variable=self.v_reemplazar).pack(side="left")
        ttk.Button(pie, text="Cancelar", command=self._cerrar).pack(side="right")
        self.btn_guardar = ttk.Button(pie, text="ELIMINAR LAS MARCADAS", command=self._guardar,
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


# ------------------------------------------------------------------------- GUI
# Paleta de fósforo: ámbar = información, cian = lo que se puede tocar.
BG, PANEL, PANEL2, LINEA = "#0d1117", "#141b24", "#1b2531", "#2a3a4a"
FG, GRIS, TENUE = "#ffb347", "#c09254", "#9a7b4c"  # contraste 10,6 · 6,7 · 4,8
ACENTO, ACENTO2, ACENTO_OSC = "#4dd9e8", "#9beef6", "#163b42"
TINTA, ROJO, VERDE = "#071014", "#ff6b6b", "#7ee787"

FUENTE_TITULO, FUENTE_TEXTO = "VT323", "Share Tech Mono"


def _carpeta_fuentes() -> Path:
    base = os.environ.get("LOCALAPPDATA") if ES_WINDOWS else None
    if base:
        return Path(base) / "PDF Master" / "fuentes"
    if platform.system() == "Linux":
        return Path.home() / ".local" / "share" / "fonts" / "pdf-master"
    return Path(tempfile.gettempdir()) / "pdf-master-fuentes"


def registrar_fuentes():
    """Deja disponibles las fuentes incluidas, solo para este programa (en Windows no se
    instalan en el sistema). Si algo falla, la interfaz usa fuentes del sistema."""
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


class App(_BaseTk):
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
        self.reglas = list(self.cfg.get("reglas", []))
        self.herr_pdfs = []
        self._win_herr = self._win_auto = None

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
        self.after(150, self._sondear)
        self._preparar_dnd()

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
        st.configure("Title.TLabel", font=T(34), foreground=FG)
        st.configure("Cursor.TLabel", font=T(34), foreground=ACENTO)
        st.configure("Section.TLabel", font=T(19), foreground=FG)
        st.configure("Step.TLabel", font=T(19), foreground=ACENTO)
        st.configure("Big.TLabel", font=(texto, 11), foreground=GRIS, background=PANEL)
        # botones: secundario (borde cian), principal (cian lleno), enlace (texto)
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
               foreground=[("disabled", "#3f7b84")])
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
        st.map("Treeview", background=[("selected", ACENTO_OSC)], foreground=[("selected", FG)])
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
        mas.add_command(label="  Carpetas automáticas…", command=self.abrir_auto)
        mas.add_separator()
        mas.add_command(label="  Enviar los PDF a PDFsam…", command=self.enviar_pdfsam)
        mas.add_command(label="  Ver / ocultar el registro", command=self._alternar_log)
        self.menu_mas = mas

    def _abrir_menu(self):
        b, m = self.btn_mas, self.menu_mas
        m.update_idletasks()
        x = b.winfo_rootx() + b.winfo_width() - m.winfo_reqwidth()
        m.tk_popup(max(x, 0), b.winfo_rooty() + b.winfo_height() + 2)

    def _cursor(self, encendido=True):
        """Cursor de terminal junto al título: titila solo mientras se trabaja."""
        if self.trabajando:
            self.cursor.configure(background=ACENTO if encendido else BG)
            self.after(530, self._cursor, not encendido)
        else:
            self.cursor.configure(background=ACENTO)

    # ------------------------------------------------------------ construcción
    def _paso(self, padre, numero, texto):
        """Encabezado numerado de cada paso (el que lo llama decide dónde ubicarlo)."""
        f = ttk.Frame(padre)
        ttk.Label(f, text=f"{numero}", style="Step.TLabel").pack(side="left")
        ttk.Label(f, text=f" {texto}", style="Section.TLabel").pack(side="left")
        return f

    def _armar(self):
        cont = ttk.Frame(self, padding=(26, 14, 26, 16))
        cont.pack(fill="both", expand=True)

        # --- encabezado
        enc = ttk.Frame(cont)
        enc.pack(fill="x")
        ttk.Label(enc, text="PDF MASTER", style="Title.TLabel").pack(side="left")
        self.cursor = tk.Frame(enc, background=ACENTO, width=13, height=24)
        self.cursor.pack(side="left", padx=(6, 0), pady=(4, 0))
        self.btn_mas = ttk.Button(enc, text="MÁS…", command=self._abrir_menu)
        self.btn_mas.pack(side="right")
        ttk.Label(enc, textvariable=self.v_lo, style="Faint.TLabel").pack(side="right", padx=14)
        regla = tk.Canvas(cont, height=6, background=BG, highlightthickness=0)
        regla.pack(fill="x", pady=(0, 2))

        def dibujar(e=None):
            regla.delete("all")
            ancho = regla.winfo_width()
            regla.create_line(0, 2, ancho, 2, fill=LINEA)
            regla.create_line(0, 2, 120, 2, fill=ACENTO, width=3)
            for x in range(ancho - 4, ancho - 60, -9):
                regla.create_line(x, 0, x, 5, fill=LINEA)
        regla.bind("<Configure>", dibujar)
        ttk.Label(cont, text="Convertí Word y PowerPoint a PDF. Si querés, unilos en un solo archivo.",
                  style="Muted.TLabel").pack(anchor="w", pady=(2, 4))

        # --- parte media en grid: la lista es lo único que se estira o se achica
        self.medio = medio = ttk.Frame(cont)
        medio.columnconfigure(0, weight=1)
        medio.rowconfigure(1, weight=1, minsize=92)  # la lista nunca baja de dos filas
        G = lambda w, fila, **kw: w.grid(row=fila, column=0, sticky="ew", **kw)

        # --- 1. archivos
        fa = self._paso(medio, "1", "ELEGÍ LOS ARCHIVOS")
        G(fa, 0, pady=(8, 4))
        ttk.Button(fa, text="Vaciar", style="Link.TButton", command=self.vaciar).pack(side="right")
        ttk.Button(fa, text="Quitar", style="Link.TButton", command=self.quitar).pack(side="right", padx=8)
        ttk.Button(fa, text="Carpeta completa…", style="Link.TButton",
                   command=self.elegir).pack(side="right", padx=8)
        ttk.Button(fa, text="+ AGREGAR ARCHIVOS", command=self.elegir_archivos).pack(side="right", padx=8)

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
        for texto, cmd in (("Marcar todos", self.marcar_todos), ("Ninguno", self.desmarcar),
                           ("Ordenar A-Z", self.ordenar), ("Subir", lambda: self.mover(-1)),
                           ("Bajar", lambda: self.mover(1))):
            ttk.Button(fb, text=texto, style="Link.TButton", command=cmd).pack(side="left", padx=(0, 10))
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
        ttk.Button(fs, text="Elegir…", style="Link.TButton", command=self.elegir_salida
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
        ttk.Label(self.fr_accion, text="3", style="Step.TLabel").pack(side="left", padx=(0, 10))
        self.btn_conv = ttk.Button(self.fr_accion, text="CONVERTIR A PDF", style="Accent.TButton",
                                   command=self.convertir)
        self.btn_conv.pack(side="left")
        self.btn_abrir = ttk.Button(self.fr_accion, text="Abrir carpeta", style="Link.TButton",
                                    command=self.abrir_carpeta)
        self.btn_abrir.pack(side="left", padx=(16, 0))
        self.btn_fusion = ttk.Button(self.fr_accion, text="Abrir el PDF unido", style="Link.TButton",
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
        ttk.Button(self.fr_log, text="Copiar registro", style="Link.TButton",
                   command=self.copiar_registro).pack(anchor="e")

        self.fr_estado = ttk.Frame(cont)
        self.fr_estado.pack(fill="x", side="bottom", pady=(8, 0))
        self.barra = ttk.Progressbar(self.fr_estado, mode="determinate")
        self.barra.pack(fill="x")
        fe = ttk.Frame(self.fr_estado)
        fe.pack(fill="x", pady=(4, 0))
        self.lbl_status = ttk.Label(fe, textvariable=self.v_status, style="Muted.TLabel")
        self.lbl_status.pack(side="left")
        ttk.Button(fe, text="Ver registro", style="Link.TButton", command=self._alternar_log
                   ).pack(side="right")
        # orden de empaquetado: estado (abajo del todo), luego acción; la parte media al final,
        # para que sea la que se achica si la ventana es baja
        self.fr_estado.pack_configure(before=enc)
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
        ttk.Button(f1, text="+ AGREGAR PDF", command=self.herr_agregar).pack(side="left")
        ttk.Button(f1, text="Vaciar", style="Link.TButton", command=self.herr_vaciar).pack(side="right")
        ttk.Button(f1, text="Quitar", style="Link.TButton", command=self.herr_quitar).pack(side="right", padx=6)
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
        ttk.Button(f, text="Abrir miniaturas…", command=self.herr_borrar).pack(anchor="w", pady=4)

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
        ttk.Button(f, text="Dividir", command=self.herr_dividir).pack(anchor="w", pady=4)

        f = bloque("Comprimir", None)
        r3 = ttk.Frame(f)
        r3.pack(fill="x")
        ttk.Label(r3, text="Nivel:").pack(side="left")
        ttk.Combobox(r3, textvariable=self.v_nivel, state="readonly", width=8,
                     values=["baja", "media", "alta"]).pack(side="left", padx=6)
        ttk.Label(r3, style="Muted.TLabel",
                  text="baja: solo estructura · media: imágenes a 150 dpi · alta: 96 dpi").pack(side="left")
        ttk.Button(f, text="Comprimir", command=self.herr_comprimir).pack(anchor="w", pady=4)
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
        ttk.Button(r1, text="Elegir carpeta…", command=self.auto_elegir).pack(side="left")
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
        ttk.Button(r3, text="Agregar regla", command=self.auto_agregar).pack(side="right")
        cuerpo = tk.Frame(t, background=LINEA, padx=1, pady=1)
        cuerpo.pack(fill="both", expand=True, pady=8)
        self.lista_reglas = tk.Listbox(cuerpo, height=6, activestyle="none", exportselection=False,
                                       relief="flat", borderwidth=0, highlightthickness=0,
                                       background=PANEL, foreground=FG, font=(self.fam, 10),
                                       selectbackground=ACENTO_OSC, selectforeground=ACENTO)
        self.lista_reglas.pack(fill="both", expand=True)
        r4 = ttk.Frame(t)
        r4.pack(fill="x")
        ttk.Button(r4, text="Quitar regla", command=self.auto_quitar).pack(side="left")
        self.btn_vig = ttk.Button(r4, text="Detener vigilancia" if self.vigilante else "Iniciar vigilancia",
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
        nuevos = filtrar_archivos(rutas)
        if not nuevos:
            messagebox.showinfo("Sin archivos válidos",
                                "No hay ningún Word, PowerPoint ni PDF en lo que elegiste.")
            return
        if reemplazar or not self.archivos:
            self.archivos, self.rangos = list(nuevos), {}
            self.base = nuevos[0].parent
        else:
            existentes = set(self.archivos)
            self.archivos += [n for n in nuevos if n not in existentes]
        self.modo = "archivos"
        self._llenar_lista()

    def _fijar_carpeta(self, ruta: Path):
        self.modo = "carpeta"
        self.base = ruta
        self.rangos = {}
        self.archivos = listar(ruta, self.v_recursivo.get())
        self._llenar_lista()

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

    def quitar(self):
        quitar = set(self._marcados_idx())
        if quitar:
            self.archivos = [a for i, a in enumerate(self.archivos) if i not in quitar]
            self._llenar_lista([])

    def vaciar(self):
        self.archivos, self.rangos = [], {}
        self._llenar_lista([])

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
        self.v_status.set("Trabajando…")
        self._iniciar(trabajo, (elegidos, self.base, opc, instalar))

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
        self.btn_conv.config(state="disabled")
        self._cursor()
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

    def _estado_lo(self):
        self.v_lo.set("LibreOffice: listo" if buscar_soffice()
                      else "LibreOffice: no instalado (se instala al convertir)")

    def _log(self, texto):
        self.log.config(state="normal")
        inicio = texto.lstrip()
        tag = ("error" if "ERROR" in texto or "error:" in texto
               or inicio.startswith(("No se pudo", "No pude", "Error"))
               else "ok" if inicio.startswith(("OK", "FUSIONADO", "Listo")) else ())
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
            self.vacio.config(text="Tocá “Agregar archivos” para empezar")
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
            self._btn_vig("Iniciar vigilancia")
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
        self._btn_vig("Detener vigilancia")
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
EBEXGBARFxgQERcYEBEXGBAQGBcREBgXERAYFxEQGBcREBgXERAYFxEQGBcREBgXEQAAAQCSAKAA/gFAABEAABIHFhUUByMmNTQ3
JjU0NzMWFf4EBARkBAQEBGQEAQEREBgXEREXGBARFxgQEBgAAQB0AKABHAFAABEAAAAHFhUUByMmNTQ3JjU0NzMWFQEcBAQEoAQE
BASgBAEBERAYFxERFxgQERcYEBAYAAACAJIAAAD+AeAAEQAjAAASBxYVFAcjJjU0NyY1NDczFhUQFRQHFhUUByMmNTQ3JjU0NzP+
BAQEZAQEBARkBAQEBGQEBAQEZAGhERAYFxERFxgQERcYEBAY/tgYFxEQGBcRERcYEBEXGBAAAQCS/2AA/gCgACQAADYVFAcWFRQH
IxYVFAcjJjU0NzMmNTQ3IyY1NDcmNTQ3MxYVFAf+BAQEKAQEPAQEKAQEKAQEBARkBARAGBcREBgXERAYFxERFxgQERcYEBEXGBAR
FxgQEBgXEQADACQAAAFsAFAACQATAB0AADcWFRQHIyY1NDczFhUUByMmNTQ3MxYVFAcjJjU0N3gEBFAEBMgEBFAEBMgEBFAEBFAQ
GBcRERcYEBAYFxERFxgQEBgXEREXGBAAAAIAnAAAAPQCMAApADMAABIVFAcWFRQHFhUUBxYVFAcjJjU0NyY1NDcmNTQ3JjU0NyY1
NDczFhUUBxEWFRQHIyY1NDf0BAQEBAQEBFAEBAQEBAQEBAQEUAQEBARQBAQB0BgXERAYFxEQGBcREBgXEREXGBARFxgQERcYEBEX
GBARFxgQEBgXEf5wEBgXEREXGBAAAgCc/7AA9AHgAAkAMwAAEyY1NDczFhUUBxYVFAcWFRQHFhUUBxYVFAcWFRQHIyY1NDcmNTQ3
JjU0NyY1NDcmNTQ3M6AEBFAEBAQEBAQEBAQEBARQBAQEBAQEBAQEBFABkBEXGBAQGBcRYBgXERAYFxEQGBcREBgXERAYFxERFxgQ
ERcYEBEXGBARFxgQERcYEAACACQAUAFsAjAATQBfAAAlFhUUByMWFRQHIyY1NDcjFhUUByMmNTQ3IyY1NDczJjU0NyY1NDcjJjU0
NzMmNTQ3MxYVFAczJjU0NzMWFRQHMxYVFAcjFhUUBxYVFAcmNTQ3JjU0NyMWFRQHFhUUBzMBaAQEMgQEUAQEPAQEUAQEMgQEMgQE
BAQyBAQyBARQBAQ8BARQBAQyBAQyBAQEBFQEBAQ8BAQEBDzwEBgXERAYFxERFxgQEBgXEREXGBARFxgQERcYEBEXGBARFxgQERcY
EBAYFxERFxgQEBgXERAYFxEQGBcREBgXEREXGBARFxgQEBgXERAYFxEAAQCSAAAA/gCgABEAADYHFhUUByMmNTQ3JjU0NzMWFf4E
BARkBAQEBGQEYREQGBcRERcYEBEXGBAQGAAAAgAkAAABWAIwADkAQwAAAAcWFRQHIxYVFAcjFhUUByMmNTQ3MyY1NDczJjU0NyY1
NDcjFhUUByMmNTQ3MyY1NDczFhUUBzMWFQMWFRQHIyY1NDcBWAQEBDIEBDIEBFAEBDIEBDIEBAQEjAQEUAQEPAQEyAQEKARoBARQ
BAQBoREQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQEBgXEREXGBARFxgQEBgXERAY/pgQGBcRERcYEAAAAgAk/7ABWAHgAAkA
QwAAEyY1NDczFhUUBxIHIxYVFAcjJjU0NyMmNTQ3JjU0NzMmNTQ3MyY1NDczFhUUByMWFRQHIxYVFAcWFRQHMyY1NDczFhWMBARQ
BAR8BDwEBMgEBCgEBAQEMgQEMgQEUAQEMgQEMgQEBASMBARQBAGQERcYEBAYFxH+gREQGBcRERcYEBEXGBARFxgQERcYEBEXGBAQ
GBcREBgXERAYFxEQGBcRERcYEBAYAAIATAHgAUQCgAARACMAABIVFAcWFRQHIyY1NDcmNTQ3MxYHFhUUByMmNTQ3JjU0NzMWFaQE
BARQBAQEBFCkBAQEUAQEBARQBAJwGBcREBgXEREXGBARFxgQPxEQGBcRERcYEBEXGBAQGAABAJwB4AD0AoAAEQAAEgcWFRQHIyY1
NDcmNTQ3MxYV9AQEBFAEBAQEUAQCQREQGBcRERcYEBEXGBAQGAACAJL/YAD+AeAAEQA2AAASBxYVFAcjJjU0NyY1NDczFhUQFRQH
FhUUBxYVFAcjFhUUByMmNTQ3MyY1NDcjJjU0NyY1NDcz/gQEBGQEBAQEZAQEBAQEBCgEBDwEBCgEBCgEBAQEZAGhERAYFxERFxgQ
ERcYEBAY/tgYFxEQGBcREBgXERAYFxERFxgQERcYEBEXGBARFxgQAAEAJP+wAWwCgABZAAAAByMWFRQHIxYVFAcjFhUUByMWFRQH
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
