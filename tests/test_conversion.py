"""De punta a punta con LibreOffice real: Word -> PDF -> unido. Se saltea si no está instalado."""
import shutil
import subprocess

import pytest

pymupdf = pytest.importorskip("pymupdf")


@pytest.fixture
def docx(pm, tmp_path):
    soffice = pm.buscar_soffice()
    if not soffice:
        pytest.skip("LibreOffice no está instalado")
    carpeta = tmp_path / "Trabajos prácticos"
    carpeta.mkdir()
    for nombre, texto in (("TP 1 - Introducción", "uno"), ("TP 2 - Señales", "dos")):
        (carpeta / f"{nombre}.txt").write_text(texto, encoding="utf-8")
    subprocess.run([soffice, "--headless", f"-env:UserInstallation={(tmp_path / 'perfil').as_uri()}",
                    "--convert-to", "docx:MS Word 2007 XML", "--outdir", str(carpeta),
                    *[str(p) for p in carpeta.glob("*.txt")]], capture_output=True, timeout=180)
    for p in carpeta.glob("*.txt"):
        p.unlink()
    archivos = sorted(carpeta.glob("*.docx"))
    if len(archivos) != 2:
        pytest.skip("no se pudieron preparar los .docx de prueba")
    return carpeta, archivos


def test_convertir_y_unir_en_subcarpeta(pm, docx):
    carpeta, archivos = docx
    eventos = []
    opc = {"sub": True, "sobre": False, "fusionar": True, "nombre": "Entrega final",
           "blanco": True, "total_par": False, "pie": "[ARCHIVO] — [PAGINA]/[TOTAL]", "rangos": {}}
    pm.trabajo(archivos, carpeta, opc, False, lambda t, d: eventos.append((t, d)))
    assert eventos[-1] == ("fin", "Listo."), eventos
    for a in archivos:
        assert (carpeta / "PDF" / f"{a.stem}.pdf").exists()
    unido = carpeta / "PDF" / "Entrega final.pdf"
    with pymupdf.open(str(unido)) as d:
        textos = [p.get_text() for p in d]
    assert len(textos) == 4  # 1 + blanca + 1 + blanca
    assert "TP 1 - Introducción — 1/4" in textos[0]
    assert ("fusion", str(unido)) in eventos
