"""Worker PDF (PyMuPDF): rangos, fusión con blancas, pie, dividir, eliminar, comprimir."""
import pytest

pymupdf = pytest.importorskip("pymupdf")


def paginas(ruta):
    with pymupdf.open(str(ruta)) as d:
        return [p.get_text().strip() for p in d]


# ------------------------------------------------------------- rangos
@pytest.mark.parametrize("texto, esperado", [
    ("", [0, 1, 2, 3, 4]),
    ("todas", [0, 1, 2, 3, 4]),
    ("1-3", [0, 1, 2]),
    ("2,4", [1, 3]),
    ("4-", [3, 4]),
    ("-2", [0, 1]),
    ("1-2; 5", [0, 1, 4]),
    (" 3 - 3 ", [2]),
])
def test_rango_valido(motor, texto, esperado):
    assert motor["rango"](texto, 5) == esperado


@pytest.mark.parametrize("texto", ["0", "6", "3-2", "a", "1-9", ","])
def test_rango_invalido(motor, texto):
    with pytest.raises(ValueError):
        motor["rango"](texto, 5)


# ------------------------------------------------------------- fusión
def test_fusion_agrega_blanca_tras_impares(motor, hacer_pdf, tmp_path):
    a, b, c = hacer_pdf("a.pdf", 3), hacer_pdf("b.pdf", 2), hacer_pdf("c.pdf", 1)
    salida = tmp_path / "out.pdf"
    r = motor["op_merge"]({"entradas": [{"ruta": str(x)} for x in (a, b, c)],
                           "salida": str(salida), "blanco": True})
    assert r == {"total": 8, "blancos": 2, "errores": []}
    assert paginas(salida) == ["a.pdf p1", "a.pdf p2", "a.pdf p3", "",
                               "b.pdf p1", "b.pdf p2", "c.pdf p1", ""]


def test_fusion_sin_blancas(motor, hacer_pdf, tmp_path):
    a, b = hacer_pdf("a.pdf", 3), hacer_pdf("b.pdf", 1)
    r = motor["op_merge"]({"entradas": [{"ruta": str(a)}, {"ruta": str(b)}],
                           "salida": str(tmp_path / "o.pdf")})
    assert (r["total"], r["blancos"]) == (4, 0)


def test_fusion_total_par(motor, hacer_pdf, tmp_path):
    a, b = hacer_pdf("a.pdf", 2), hacer_pdf("b.pdf", 1)
    r = motor["op_merge"]({"entradas": [{"ruta": str(a)}, {"ruta": str(b)}],
                           "salida": str(tmp_path / "o.pdf"), "total_par": True})
    assert (r["total"], r["blancos"]) == (4, 1)


def test_fusion_respeta_rango_y_la_blanca_cuenta_lo_usado(motor, hacer_pdf, tmp_path):
    a, b = hacer_pdf("a.pdf", 4), hacer_pdf("b.pdf", 2)
    salida = tmp_path / "o.pdf"
    r = motor["op_merge"]({"entradas": [{"ruta": str(a), "rango": "2-4"}, {"ruta": str(b)}],
                           "salida": str(salida), "blanco": True})
    assert paginas(salida) == ["a.pdf p2", "a.pdf p3", "a.pdf p4", "", "b.pdf p1", "b.pdf p2"]
    assert r["blancos"] == 1


def test_fusion_blanca_del_tamaño_de_la_última_página(motor, hacer_pdf, tmp_path):
    a = hacer_pdf("apaisado.pdf", 1, ancho=842, alto=595)
    salida = tmp_path / "o.pdf"
    motor["op_merge"]({"entradas": [{"ruta": str(a)}], "salida": str(salida), "blanco": True})
    with pymupdf.open(str(salida)) as d:
        assert (round(d[1].rect.width), round(d[1].rect.height)) == (842, 595)


def test_fusion_informa_errores_sin_cortar(motor, hacer_pdf, tmp_path):
    a = hacer_pdf("a.pdf", 2)
    r = motor["op_merge"]({"entradas": [{"ruta": str(tmp_path / "no-existe.pdf")},
                                        {"ruta": str(a), "rango": "1"}, {"ruta": str(a)}],
                           "salida": str(tmp_path / "o.pdf")})
    assert r["total"] == 3
    assert [i for i, _ in r["errores"]] == [0]


def test_fusion_crea_la_carpeta_de_salida(motor, hacer_pdf, tmp_path):
    """Bug histórico: 'no se pudo crear PDF\\Fusionado.pdf.part'."""
    a = hacer_pdf("a.pdf", 1)
    salida = tmp_path / "PDF" / "sub carpeta" / "Fusionado.pdf"
    motor["op_merge"]({"entradas": [{"ruta": str(a)}], "salida": str(salida)})
    assert salida.exists()
    assert not salida.with_name("Fusionado.pdf.part").exists()


def test_pie_con_raya_y_acentos(motor, hacer_pdf, tmp_path):
    """Bug histórico: la raya larga salía como '?' (se codifica con cp1252)."""
    a, b = hacer_pdf("Introducción.pdf", 1), hacer_pdf("b.pdf", 1)
    salida = tmp_path / "o.pdf"
    motor["op_merge"]({"entradas": [{"ruta": str(a)}, {"ruta": str(b)}], "salida": str(salida),
                       "pie": "[ARCHIVO] — página [PAGINA] de [TOTAL]"})
    textos = paginas(salida)
    assert "Introducción — página 1 de 2" in textos[0]
    assert "b — página 2 de 2" in textos[1]


# ------------------------------------------------------------- dividir y eliminar
def test_dividir_cada_n(motor, hacer_pdf, tmp_path):
    a = hacer_pdf("doc.pdf", 5)
    r = motor["op_split"]({"ruta": str(a), "dir": str(tmp_path / "partes"),
                           "modo": "cada_n", "n": 2})
    assert [len(paginas(f)) for f in r["archivos"]] == [2, 2, 1]
    assert r["archivos"][0].endswith("doc_parte_01.pdf")


def test_dividir_por_intervalos(motor, hacer_pdf, tmp_path):
    a = hacer_pdf("doc.pdf", 6)
    r = motor["op_split"]({"ruta": str(a), "dir": str(tmp_path / "p"),
                           "modo": "rangos", "rangos": "1-2;3-;"})
    assert [paginas(f) for f in r["archivos"]] == [
        ["doc.pdf p1", "doc.pdf p2"], ["doc.pdf p3", "doc.pdf p4", "doc.pdf p5", "doc.pdf p6"]]


def test_eliminar_paginas(motor, hacer_pdf, tmp_path):
    a = hacer_pdf("doc.pdf", 4)
    salida = tmp_path / "sin.pdf"
    r = motor["op_delete"]({"ruta": str(a), "salida": str(salida), "borrar": [0, 2]})
    assert r == {"quedan": 2, "eliminadas": 2}
    assert paginas(salida) == ["doc.pdf p2", "doc.pdf p4"]


def test_no_deja_eliminar_todas(motor, hacer_pdf, tmp_path):
    a = hacer_pdf("doc.pdf", 2)
    with pytest.raises(ValueError):
        motor["op_delete"]({"ruta": str(a), "salida": str(tmp_path / "x.pdf"), "borrar": [0, 1]})


# ------------------------------------------------------------- de punta a punta
def test_comprimir_informa_el_signo_correcto(pm, hacer_pdf, tmp_path):
    """Bug histórico: el porcentaje salía con el signo invertido. Pasa por el subproceso real."""
    a = hacer_pdf("grande.pdf", 3)
    eventos = []
    pm.trabajo_herramienta("comprimir", [a], {"nivel": "baja"}, lambda t, d: eventos.append((t, d)))
    logs = [d for t, d in eventos if t == "log"]
    assert any(l.startswith("OK  grande.pdf") for l in logs), logs
    linea = next(l for l in logs if l.startswith("OK"))
    antes = a.stat().st_size
    despues = (tmp_path / "grande (comprimido).pdf").stat().st_size
    signo = "-" if despues < antes else "+"
    assert f"({signo}" in linea
    assert eventos[-1] == ("fin", "Listo.")
