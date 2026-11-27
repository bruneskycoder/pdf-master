"""Carga pdf_master.pyw como módulo y expone el motor PDF (CODIGO_PDF) para los tests."""
import importlib.machinery
import importlib.util
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent


def _cargar():
    ruta = RAIZ / "pdf_master.pyw"
    cargador = importlib.machinery.SourceFileLoader("pdf_master", str(ruta))
    spec = importlib.util.spec_from_loader("pdf_master", cargador)
    modulo = importlib.util.module_from_spec(spec)
    cargador.exec_module(modulo)
    return modulo


PM = _cargar()


@pytest.fixture(scope="session")
def pm():
    return PM


CARTELES = []  # carteles que el programa intentó mostrar durante la prueba actual


@pytest.fixture(autouse=True)
def carteles_sin_bloquear(monkeypatch):
    """En las pruebas nadie cierra un cartel: uno inesperado colgaría la suite (pasó en la
    CI de Windows). Se registra acá y se devuelve la respuesta neutra. Una prueba que espera
    un cartel puede reemplazarlo con su propio monkeypatch."""
    import types
    from tkinter import messagebox
    CARTELES.clear()
    falsos = types.SimpleNamespace()
    nombres = []
    for nombre, respuesta in (("showinfo", "ok"), ("showwarning", "ok"), ("showerror", "ok"),
                              ("askyesno", False), ("askokcancel", False),
                              ("askretrycancel", False), ("askquestion", "no"),
                              ("askyesnocancel", None)):
        def cartel(*a, _n=nombre, _r=respuesta, **kw):
            CARTELES.append((_n, a, kw))
            return _r
        setattr(falsos, nombre, cartel)
        nombres.append(nombre)
    PM._vigilar_modales(falsos, nombres)  # misma envoltura que en el programa: cuenta abiertos
    for nombre in nombres:
        monkeypatch.setattr(messagebox, nombre, getattr(falsos, nombre))
    yield CARTELES


@pytest.fixture(scope="session")
def motor():
    """Las funciones del worker PDF, ejecutadas en este mismo proceso."""
    pytest.importorskip("pymupdf")
    espacio = {"__name__": "motor_pdf"}
    exec(PM.CODIGO_PDF, espacio)
    return espacio


@pytest.fixture
def hacer_pdf(tmp_path):
    """Crea un PDF de n páginas con su número escrito en cada una."""
    pymupdf = pytest.importorskip("pymupdf")

    def _hacer(nombre, n, ancho=595, alto=842):
        doc = pymupdf.open()
        for i in range(n):
            pg = doc.new_page(width=ancho, height=alto)
            pg.insert_text((72, 72), f"{nombre} p{i + 1}")
        ruta = tmp_path / nombre
        ruta.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(ruta))
        doc.close()
        return ruta

    return _hacer
