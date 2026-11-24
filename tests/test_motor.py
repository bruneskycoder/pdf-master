"""Funciones del motor sin GUI: nombres de destino, orden natural, filtrado."""
from pathlib import Path


# ------------------------------------------------------------- orden natural
def test_orden_natural_numeros(pm):
    nombres = ["tp 10.docx", "tp 2.docx", "TP 1.docx", "tp 9.pptx"]
    orden = sorted((Path(n) for n in nombres), key=pm._clave_natural)
    assert [p.name for p in orden] == ["TP 1.docx", "tp 2.docx", "tp 9.pptx", "tp 10.docx"]


def test_orden_natural_ñ_y_espacios(pm):
    nombres = ["Año 2/clase 10.pdf", "Año 2/clase 2.pdf", "Año 1/clase 3.pdf"]
    orden = sorted((Path(n) for n in nombres), key=pm._clave_natural)
    assert [str(p) for p in orden] == [str(Path("Año 1/clase 3.pdf")),
                                      str(Path("Año 2/clase 2.pdf")),
                                      str(Path("Año 2/clase 10.pdf"))]


# ------------------------------------------------------------- filtrado
def test_filtrar_descarta_temporales_y_extensiones_ajenas(pm, tmp_path):
    rutas = [tmp_path / n for n in ("a.docx", "~$a.docx", "b.txt", "c.PPTX", "d.pdf", "e.xlsx")]
    assert [p.name for p in pm.filtrar_archivos(rutas)] == ["a.docx", "c.PPTX", "d.pdf"]


def test_filtrar_descarta_pdf_que_es_resultado_de_un_word(pm, tmp_path):
    rutas = [tmp_path / "Informe.docx", tmp_path / "informe.pdf", tmp_path / "Otro.pdf"]
    assert [p.name for p in pm.filtrar_archivos(rutas)] == ["Informe.docx", "Otro.pdf"]


def test_filtrar_conserva_pdf_homónimo_de_otra_carpeta(pm, tmp_path):
    rutas = [tmp_path / "a" / "x.docx", tmp_path / "b" / "x.pdf"]
    assert len(pm.filtrar_archivos(rutas)) == 2


def test_listar_recursivo(pm, tmp_path):
    (tmp_path / "sub").mkdir()
    for n in ("uno.docx", "sub/dos.pptx", "sub/tres.pdf", "nota.txt"):
        (tmp_path / n).write_bytes(b"x")
    assert [p.name for p in pm.listar(tmp_path, False)] == ["uno.docx"]
    assert sorted(p.name for p in pm.listar(tmp_path, True)) == ["dos.pptx", "tres.pdf", "uno.docx"]


# ------------------------------------------------------------- destino
def test_destino_junto_al_original(pm, tmp_path):
    origen = tmp_path / "Mis cosas" / "Trabajo práctico Nº 1.docx"
    assert pm.destino_pdf(origen, tmp_path, False) == origen.with_suffix(".pdf")


def test_destino_subcarpeta_pdf_respeta_estructura(pm, tmp_path):
    origen = tmp_path / "unidad 2" / "clase.pptx"
    assert pm.destino_pdf(origen, tmp_path, True) == tmp_path / "PDF" / "unidad 2" / "clase.pdf"


def test_destino_subcarpeta_fuera_de_base(pm, tmp_path):
    origen = Path("/otro/lugar/clase.docx")
    assert pm.destino_pdf(origen, tmp_path, True) == tmp_path / "PDF" / "clase.pdf"


def test_destino_otra_carpeta(pm, tmp_path):
    salida = tmp_path / "Salida PDF"
    origen = tmp_path / "base" / "sub" / "x.doc"
    assert pm.destino_pdf(origen, tmp_path / "base", salida) == salida / "sub" / "x.pdf"
    assert pm.destino_pdf(Path("/lejos/y.doc"), tmp_path / "base", salida) == salida / "y.pdf"


def test_destino_de_un_pdf_es_el_mismo(pm, tmp_path):
    origen = tmp_path / "ya.pdf"
    for salida in (False, True, tmp_path / "otra"):
        assert pm.destino_pdf(origen, tmp_path, salida) == origen


def test_carpeta_destino(pm, tmp_path):
    assert pm.carpeta_destino(tmp_path, False) == tmp_path
    assert pm.carpeta_destino(tmp_path, True) == tmp_path / "PDF"
    assert pm.carpeta_destino(tmp_path, tmp_path / "z") == tmp_path / "z"


# ------------------------------------------------------------- nombres
def test_limpiar_nombre(pm):
    assert pm.limpiar_nombre('TP: final/v2?') == "TP_ final_v2_"
    assert pm.limpiar_nombre("Fusión.PDF") == "Fusión"
    assert pm.limpiar_nombre("   ") == "Fusionado"


def test_nombre_libre(pm, tmp_path):
    ruta = tmp_path / "Fusionado.pdf"
    assert pm.nombre_libre(ruta) == ruta
    ruta.write_bytes(b"x")
    (tmp_path / "Fusionado (2).pdf").write_bytes(b"x")
    assert pm.nombre_libre(ruta) == tmp_path / "Fusionado (3).pdf"


def test_tam_legible(pm):
    assert pm.tam_legible(512) == "512 B"
    assert pm.tam_legible(1536) == "1.5 KB"
    assert pm.tam_legible(5 * 1024 ** 3) == "5.0 GB"
