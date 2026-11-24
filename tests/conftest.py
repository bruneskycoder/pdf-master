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
