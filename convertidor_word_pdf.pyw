# -*- coding: utf-8 -*-
"""
Convertidor a PDF + herramientas PDF (estilo PDFsam).  UN SOLO ARCHIVO, solo necesita Python.

Pestañas:
  1. Convertir y fusionar : Word/PowerPoint -> PDF (LibreOffice), tabla con detalles
                            (tamaño, páginas, intervalo, fecha), fusión con página en blanco
                            si es impar y pie de página, enviar a PDFsam.
  2. Herramientas PDF     : dividir, comprimir y eliminar páginas con vista de miniaturas.
  3. Carpetas automáticas : carpetas vigiladas ("hot folders") que dividen, fusionan o
                            comprimen solas todo lo que les pongas.

Motor de conversión: LibreOffice (si falta, el programa lo descarga y lo instala).
Motor PDF: PyMuPDF (si falta, el programa lo instala con pip, una sola vez).
"""
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
        total = len(out)
        for k in range(total):
            pg = out[k]
            txt = (pie.replace("[PAGINA]", str(k + 1)).replace("[TOTAL]", str(total))
                   .replace("[ARCHIVO]", origen[k]))
            txt = txt.encode("cp1252", "replace").decode("cp1252")
            ancho = pymupdf.get_text_length(txt, fontname="helv", fontsize=9)
            x = (pg.rect.width - ancho) / 2
            y = pg.rect.height - 18
            pt = pymupdf.Point(x, y) * pg.derotation_matrix
            pg.insert_text(pt, txt, fontsize=9, fontname="helv", color=(0.25, 0.25, 0.25),
                           rotate=pg.rotation)
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
                emit("log", f"OK  {ruta.name}: {tam_legible(a)} → {tam_legible(d)} "
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
                emit("log", f"[{carpeta.name}] FUSIONADOS {len(listos)} archivos → salida/{destino.name} "
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
                    msg = (f"{tam_legible(r['antes'])} → {tam_legible(r['despues'])}"
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
        self.geometry("920x680")
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
        ttk.Label(barra, textvariable=self.v_estado).pack(side="right")

        pie = ttk.Frame(self)
        pie.pack(side="bottom", fill="x", padx=8, pady=6)
        self.v_reemplazar = tk.BooleanVar(value=False)
        ttk.Checkbutton(pie, text="Reemplazar el PDF original (si no, crea “… (sin páginas).pdf”)",
                        variable=self.v_reemplazar).pack(side="left")
        ttk.Button(pie, text="Cancelar", command=self._cerrar).pack(side="right")
        self.btn_guardar = ttk.Button(pie, text="Eliminar las marcadas", command=self._guardar,
                                      state="disabled")
        self.btn_guardar.pack(side="right", padx=6)

        cuerpo = ttk.Frame(self)
        cuerpo.pack(fill="both", expand=True, padx=8)
        self.canvas = tk.Canvas(cuerpo, background="#e8e8e8", highlightthickness=0)
        sb = ttk.Scrollbar(cuerpo, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.interior = tk.Frame(self.canvas, background="#e8e8e8")
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
            celda = tk.Frame(self.interior, background="#e8e8e8", padx=4, pady=4)
            celda.grid(row=i // self.COLS, column=i % self.COLS, padx=4, pady=4)
            lbl = tk.Label(celda, image=img, borderwidth=0, highlightthickness=4,
                           highlightbackground="#e8e8e8", cursor="hand2")
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
            lbl.config(highlightbackground="#d32f2f" if marcada else "#e8e8e8")
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
BG, FG, GRIS, LINEA = "#f6f7f9", "#1f2328", "#6b7280", "#e3e5ea"
ACENTO, ACENTO2, ROJO, VERDE = "#2563eb", "#1d4ed8", "#b42318", "#15803d"


def _tipo(p: Path):
    e = p.suffix.lower()
    return "PDF" if e == ".pdf" else ("PowerPoint" if e in (".pptx", ".ppt") else "Word")


class App(_BaseTk):
    def __init__(self):
        super().__init__()
        self.title("Convertir a PDF")
        self.minsize(800, 560)
        self.configure(background=BG)
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
        familia = "Segoe UI" if ES_WINDOWS else None
        for nombre in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                f = tkfont.nametofont(nombre)
                if familia:
                    f.configure(family=familia)
                f.configure(size=10)
            except tk.TclError:
                pass
        fam = familia or tkfont.nametofont("TkDefaultFont").actual("family")
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure(".", background=BG, foreground=FG, bordercolor=LINEA, focuscolor=BG)
        for w in ("TFrame", "TLabel", "TCheckbutton", "TRadiobutton"):
            st.configure(w, background=BG)
        st.configure("TLabel", foreground=FG)
        st.configure("Muted.TLabel", foreground=GRIS)
        st.configure("Title.TLabel", font=(fam, 17, "bold"))
        st.configure("Section.TLabel", font=(fam, 10, "bold"))
        st.configure("TButton", background="#e8eaee", foreground=FG, borderwidth=0,
                     padding=(12, 6), focuscolor="#e8eaee")
        st.map("TButton", background=[("active", "#dcdfe5"), ("disabled", "#eef0f3")],
               foreground=[("disabled", "#9ca3af")])
        st.configure("Accent.TButton", background=ACENTO, foreground="white",
                     font=(fam, 11, "bold"), padding=(26, 10), focuscolor=ACENTO)
        st.map("Accent.TButton", background=[("active", ACENTO2), ("disabled", "#a9c0f3")],
               foreground=[("disabled", "white")])
        st.configure("Link.TButton", background=BG, foreground=ACENTO, padding=(2, 2),
                     focuscolor=BG)
        st.map("Link.TButton", foreground=[("active", ACENTO2), ("disabled", "#9ca3af")],
               background=[("active", BG)])
        st.configure("Treeview", background="white", fieldbackground="white", rowheight=30,
                     borderwidth=0, foreground=FG)
        st.map("Treeview", background=[("selected", "#dbe7ff")], foreground=[("selected", FG)])
        st.configure("Treeview.Heading", background="#eef0f3", foreground=GRIS, relief="flat",
                     font=(fam, 9, "bold"), padding=(8, 6))
        st.map("Treeview.Heading", background=[("active", "#eef0f3")])
        st.configure("Horizontal.TProgressbar", troughcolor="#e3e5ea", background=ACENTO,
                     borderwidth=0, thickness=6)
        st.configure("TEntry", fieldbackground="white", padding=4)
        st.configure("TCombobox", fieldbackground="white", padding=4)
        st.configure("TNotebook", background=BG)
        self.fam = fam

    def _menu(self):
        m = tk.Menu(self)
        mas = tk.Menu(m, tearoff=0)
        mas.add_command(label="Herramientas PDF (dividir, comprimir, eliminar páginas)…",
                        command=self.abrir_herramientas)
        mas.add_command(label="Carpetas automáticas…", command=self.abrir_auto)
        mas.add_separator()
        mas.add_command(label="Enviar los PDF a PDFsam…", command=self.enviar_pdfsam)
        mas.add_command(label="Ver / ocultar el registro", command=self._alternar_log)
        m.add_cascade(label="Más", menu=mas)
        self.config(menu=m)

    # ------------------------------------------------------------ construcción
    def _armar(self):
        cont = ttk.Frame(self, padding=(22, 16, 22, 14))
        cont.pack(fill="both", expand=True)

        enc = ttk.Frame(cont)
        enc.pack(fill="x")
        ttk.Label(enc, text="Convertir a PDF", style="Title.TLabel").pack(side="left")
        ttk.Label(enc, textvariable=self.v_lo, style="Muted.TLabel").pack(side="right")
        ttk.Label(cont, text="Word, PowerPoint y PDF. Podés convertirlos y unirlos en uno solo.",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 10))

        # --- archivos
        fa = ttk.Frame(cont)
        fa.pack(fill="x")
        ttk.Button(fa, text="+  Agregar archivos", command=self.elegir_archivos).pack(side="left")
        ttk.Button(fa, text="Carpeta completa…", style="Link.TButton",
                   command=self.elegir).pack(side="left", padx=10)
        ttk.Button(fa, text="Vaciar", style="Link.TButton", command=self.vaciar).pack(side="right")
        ttk.Button(fa, text="Quitar", style="Link.TButton", command=self.quitar).pack(side="right", padx=6)

        cuerpo = tk.Frame(cont, background=LINEA, padx=1, pady=1)
        cuerpo.pack(fill="both", expand=True, pady=(8, 2))
        cols = ("nombre", "tipo", "pags", "tam", "rango")
        self.tabla = ttk.Treeview(cuerpo, columns=cols, show="headings", selectmode="extended", height=6)
        for c, txt, w, anc in (("nombre", "Archivo", 340, "w"), ("tipo", "Tipo", 90, "w"),
                               ("pags", "Páginas", 80, "e"), ("tam", "Tamaño", 80, "e"),
                               ("rango", "Páginas a usar", 130, "w")):
            self.tabla.heading(c, text=txt, anchor=anc if c != "nombre" else "w")
            self.tabla.column(c, width=w, anchor=anc, stretch=(c == "nombre"))
        sb = ttk.Scrollbar(cuerpo, command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tabla.pack(side="left", fill="both", expand=True)
        self.tabla.bind("<<TreeviewSelect>>", lambda e: self._contar())
        self.tabla.bind("<Double-1>", self._doble_clic)
        self.vacio = tk.Label(cuerpo, background="white", foreground=GRIS, justify="center",
                              text="Arrastrá archivos acá\no tocá “Agregar archivos”",
                              font=(self.fam, 11))
        self.vacio.place(relx=0.5, rely=0.5, anchor="center")

        fb = ttk.Frame(cont)
        fb.pack(fill="x")
        for texto, cmd in (("Marcar todos", self.marcar_todos), ("Ninguno", self.desmarcar),
                           ("Ordenar A-Z", self.ordenar), ("▲", lambda: self.mover(-1)),
                           ("▼", lambda: self.mover(1))):
            ttk.Button(fb, text=texto, style="Link.TButton", command=cmd).pack(side="left", padx=(0, 8))
        ttk.Label(fb, textvariable=self.v_estado, style="Muted.TLabel").pack(side="right")
        ttk.Label(cont, style="Muted.TLabel", textvariable=self.v_dnd).pack(anchor="w")

        # --- dónde se guardan
        ttk.Label(cont, text="Guardar los PDF", style="Section.TLabel").pack(anchor="w", pady=(12, 2))
        fs = ttk.Frame(cont)
        fs.pack(fill="x")
        ttk.Radiobutton(fs, text="Junto a cada archivo original", value="junto",
                        variable=self.v_salida_modo).pack(anchor="w")
        ttk.Radiobutton(fs, text="En una subcarpeta “PDF”", value="subcarpeta",
                        variable=self.v_salida_modo).pack(anchor="w")
        fo = ttk.Frame(fs)
        fo.pack(fill="x", anchor="w")
        ttk.Radiobutton(fo, text="En otra carpeta:", value="otra", variable=self.v_salida_modo,
                        command=self._radio_otra).pack(side="left")
        ttk.Button(fo, text="Elegir…", style="Link.TButton", command=self.elegir_salida
                   ).pack(side="left", padx=6)
        ttk.Label(cont, textvariable=self.v_destino, style="Muted.TLabel", wraplength=820,
                  justify="left").pack(anchor="w", pady=(4, 0))
        ttk.Checkbutton(cont, text="Reemplazar si el PDF ya existe",
                        variable=self.v_sobrescribir).pack(anchor="w", pady=(4, 0))

        # --- fusión (se despliega)
        self.chk_fusion = ttk.Checkbutton(cont, text="Unir todo en un solo PDF", variable=self.v_fusionar,
                                          command=self._alternar_fusion)
        self.chk_fusion.pack(anchor="w", pady=(12, 0))
        self.fr_fusion = ttk.Frame(cont, padding=(22, 4, 0, 0))
        a = ttk.Frame(self.fr_fusion)
        a.pack(fill="x")
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
        ttk.Label(b, style="Muted.TLabel", text="[PAGINA] [TOTAL] [ARCHIVO]").pack(side="left")

        # --- acción
        self.fr_accion = ttk.Frame(cont)
        self.fr_accion.pack(fill="x", pady=(16, 0), side="bottom")
        ttk.Checkbutton(self.fr_accion, text="Abrir el resultado al terminar",
                        variable=self.v_abrir).pack(side="right")
        self.btn_conv = ttk.Button(self.fr_accion, text="Convertir", style="Accent.TButton",
                                   command=self.convertir)
        self.btn_conv.pack(side="left")
        self.btn_abrir = ttk.Button(self.fr_accion, text="Abrir carpeta", style="Link.TButton",
                                    command=self.abrir_carpeta)
        self.btn_abrir.pack(side="left", padx=(14, 0))
        self.btn_fusion = ttk.Button(self.fr_accion, text="Abrir fusionado", style="Link.TButton",
                                     command=self.abrir_fusion)

        self.log_visible = False
        self.fr_log = ttk.Frame(cont)
        self.log = tk.Text(self.fr_log, height=7, state="disabled", wrap="word", background="white",
                           relief="flat", highlightthickness=1, highlightbackground=LINEA,
                           font=(self.fam, 9))
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
        # orden de empaquetado: estado (abajo del todo), luego acción
        self.fr_estado.pack_configure(before=enc)
        self.fr_accion.pack_configure(after=self.fr_estado)

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
        setattr(self, attr, w)
        cont = ttk.Frame(w, padding=18)
        cont.pack(fill="both", expand=True)
        ttk.Label(cont, textvariable=self.v_status, style="Muted.TLabel", wraplength=640
                  ).pack(side="bottom", anchor="w", pady=(8, 0))
        constructor(cont)

    def abrir_herramientas(self):
        self._ventana("_win_herr", "Herramientas PDF", self._tab_herramientas, "720x640")

    def abrir_auto(self):
        self._ventana("_win_auto", "Carpetas automáticas", self._tab_auto, "760x620")

    def _tab_herramientas(self, t):
        pad = {"pady": 5}
        ttk.Label(t, text="Herramientas PDF", style="Title.TLabel").pack(anchor="w")
        ttk.Label(t, style="Muted.TLabel",
                  text="Trabajan sobre PDF. Si no agregás ninguno acá, usan los PDF marcados de la ventana principal."
                  ).pack(anchor="w", pady=(0, 8))
        f1 = ttk.Frame(t)
        f1.pack(fill="x", **pad)
        ttk.Button(f1, text="+  Agregar PDF", command=self.herr_agregar).pack(side="left")
        ttk.Button(f1, text="Vaciar", style="Link.TButton", command=self.herr_vaciar).pack(side="right")
        ttk.Button(f1, text="Quitar", style="Link.TButton", command=self.herr_quitar).pack(side="right", padx=6)
        cuerpo = tk.Frame(t, background=LINEA, padx=1, pady=1)
        cuerpo.pack(fill="both", expand=True, **pad)
        sb = ttk.Scrollbar(cuerpo)
        sb.pack(side="right", fill="y")
        self.herr_lista = tk.Listbox(cuerpo, selectmode="extended", activestyle="none",
                                     exportselection=False, yscrollcommand=sb.set, height=6,
                                     relief="flat", borderwidth=0, highlightthickness=0,
                                     selectbackground="#dbe7ff", selectforeground=FG)
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
            ttk.Label(t, text=titulo, style="Section.TLabel").pack(anchor="w", pady=(12, 0))
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
        ttk.Label(t, text="Carpetas automáticas", style="Title.TLabel").pack(anchor="w")
        ttk.Label(t, style="Muted.TLabel", wraplength=700, justify="left",
                  text="Cada carpeta vigilada procesa sola lo que le pongas. Los resultados van a "
                       "“salida”, los originales a “procesados” y los que fallan a “con_error”. "
                       "Funciona mientras el programa esté abierto."
                  ).pack(anchor="w", pady=(0, 8))
        r1 = ttk.Frame(t)
        r1.pack(fill="x", pady=2)
        ttk.Button(r1, text="Elegir carpeta…", command=self.auto_elegir).pack(side="left")
        ttk.Label(r1, textvariable=self.v_acc_carpeta, style="Muted.TLabel").pack(side="left", padx=8)
        r2 = ttk.Frame(t)
        r2.pack(fill="x", pady=4)
        ttk.Radiobutton(r2, text="Comprimir", value="comprimir", variable=self.v_acc).pack(side="left")
        ttk.Combobox(r2, textvariable=self.v_acc_nivel, state="readonly", width=7,
                     values=["baja", "media", "alta"]).pack(side="left", padx=4)
        ttk.Radiobutton(r2, text="Dividir cada", value="dividir", variable=self.v_acc).pack(side="left", padx=(14, 4))
        ttk.Entry(r2, textvariable=self.v_acc_n, width=4).pack(side="left")
        ttk.Label(r2, text="págs.").pack(side="left", padx=(2, 0))
        ttk.Radiobutton(r2, text="Fusionar", value="fusionar", variable=self.v_acc).pack(side="left", padx=(14, 4))
        ttk.Checkbutton(r2, text="con blanco si es impar", variable=self.v_acc_blanco).pack(side="left")
        r3 = ttk.Frame(t)
        r3.pack(fill="x")
        ttk.Checkbutton(r3, text="También convertir Word/PowerPoint que lleguen",
                        variable=self.v_acc_conv).pack(side="left")
        ttk.Button(r3, text="Agregar regla", command=self.auto_agregar).pack(side="right")
        cuerpo = tk.Frame(t, background=LINEA, padx=1, pady=1)
        cuerpo.pack(fill="both", expand=True, pady=8)
        self.lista_reglas = tk.Listbox(cuerpo, height=6, activestyle="none", exportselection=False,
                                       relief="flat", borderwidth=0, highlightthickness=0,
                                       selectbackground="#dbe7ff", selectforeground=FG)
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
            self.fr_fusion.pack(anchor="w", fill="x", after=self.chk_fusion)
        else:
            self.fr_fusion.pack_forget()
        self._actualizar_destino()

    def _alternar_log(self):
        if self.log_visible:
            self.fr_log.pack_forget()
        else:
            self.fr_log.pack(fill="x", side="bottom", pady=(8, 0), after=self.fr_accion)
        self.log_visible = not self.log_visible

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

    def _actualizar_destino(self):
        s = self.salida()
        if isinstance(s, Path):
            txt = f"Se guardan en: {s}"
            carpeta = s
        elif self.v_salida_modo.get() == "otra":
            self.v_destino.set("Elegí la carpeta de destino.")
            return
        elif not self.base:
            self.v_destino.set("Dónde se guardan se muestra acá cuando elijas archivos.")
            return
        elif s:
            carpeta = self.base / "PDF"
            txt = f"Se guardan en: {carpeta}"
        else:
            carpeta = self.base
            txt = ("Se guardan al lado de cada original" +
                   (f" (por ejemplo en {self.base})" if self.base else ""))
        if self.v_fusionar.get():
            nom = (self.v_nombre.get().strip() or "Fusionado")
            txt += f"\nEl PDF unido: {carpeta / (nom + '.pdf')}"
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
                                "o usá Más ▸ Herramientas PDF.")
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
        self.log.insert("end", texto + "\n")
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
            self.v_dnd.set("Arrastrar y soltar: no disponible (usá los botones)")

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
               "fusionar": "fusionar" + (" + blanco si impar" if r.get("blanco") else "")}[r["accion"]]
        return f"{r['carpeta']}   →   {acc}" + ("   [+Word/PPT]" if r.get("convertir") else "")

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


if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception:
        texto = traceback.format_exc()
        try:
            (Path.home() / "convertidor_error.log").write_text(texto, encoding="utf-8")
        except Exception:
            pass
        try:
            raiz = tk.Tk()
            raiz.withdraw()
            messagebox.showerror("Error al abrir el programa", texto[-1500:])
        except Exception:
            pass
