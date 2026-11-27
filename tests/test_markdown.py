"""Lector de Markdown: parser, HTML, exportación a PDF con el worker real y ventana."""
import random
import time

import pytest

from test_gui import app  # noqa: F401  (la ventana principal de las pruebas de interfaz)


def segs(pm, texto):
    return pm.markdown_a_bloques(texto)


def con_tag(seg, tag):
    return "".join(t for t, tg in seg if tag in tg)


# ------------------------------------------------------------------ parser
def test_encabezados(pm):
    s = segs(pm, "# Uno\n## Dos\n### Tres\n#### Cuatro\n###### Seis\n#sin espacio")
    assert con_tag(s, "h1") == "Uno\n"
    assert con_tag(s, "h2") == "Dos\n"
    assert con_tag(s, "h3") == "Tres\nCuatro\nSeis\n"
    assert "#sin espacio" in "".join(t for t, _ in s)


def test_parrafo_negrita_cursiva_codigo(pm):
    s = segs(pm, "Hola **fuerte** y *suave* y _otra_ y `x = 1`.\nSigue la línea.")
    assert con_tag(s, "b") == "fuerte"
    assert con_tag(s, "i") == "suave" + "otra"
    assert con_tag(s, "code") == "x = 1"
    plano = "".join(t for t, _ in s)
    assert "**" not in plano and "`" not in plano
    assert "Hola fuerte y suave y otra y x = 1. Sigue la línea." in plano


def test_negrita_y_cursiva_juntas_y_snake_case(pm):
    s = segs(pm, "***ambas*** y una_variable_larga")
    assert con_tag(s, "bi") == "ambas"
    assert "una_variable_larga" in "".join(t for t, _ in s)


def test_bloque_de_codigo_con_y_sin_lenguaje_y_sangria(pm):
    s = segs(pm, "```python\nx = 1\n# no es titulo\n```\n\n```\nsin lenguaje\n```\n\n    sangrado\n    **no**\n")
    assert con_tag(s, "codeblock") == "x = 1\n# no es titulo\nsin lenguaje\nsangrado\n**no**\n"
    assert con_tag(s, "h1") == ""


def test_bloque_de_codigo_sin_cerrar(pm):
    s = segs(pm, "```\nabierto\nsigue")
    assert con_tag(s, "codeblock") == "abierto\nsigue\n"


def test_citas(pm):
    s = segs(pm, "> primera línea\n> con **negrita**\n\nTexto aparte")
    assert con_tag(s, "quote") == "primera línea\ncon negrita\n"
    assert con_tag(s, "b") == "negrita"


def test_listas_con_niveles_y_numeradas(pm):
    s = segs(pm, "- uno\n  - dos\n    - tres\n* cuatro\n+ cinco\n\n1. a\n2. b\n    1. c\n")
    items = [(t, tg) for t, tg in s if "li" in tg and t.strip() and t != "\n"]
    textos = [t for t, _ in items]
    assert textos[:2] == ["• ", "uno"]
    assert "• " in textos and "1. " in textos and "2. " in textos
    niveles = {t: [x for x in tg if x.startswith("nivel")][0] for t, tg in items}
    assert niveles["uno"] == "nivel0" and niveles["dos"] == "nivel1" and niveles["tres"] == "nivel2"
    assert niveles["c"] == "nivel1"
    assert niveles["cuatro"] == "nivel0"


def test_listas_con_sangria_de_cuatro(pm):
    s = segs(pm, "- a\n    - b\n        - c\n")
    n = {t: [x for x in tg if x.startswith("nivel")][0] for t, tg in s if t in ("a", "b", "c")}
    assert n == {"a": "nivel0", "b": "nivel1", "c": "nivel2"}


def test_tareas(pm):
    plano = "".join(t for t, _ in segs(pm, "- [ ] pendiente\n- [x] hecha\n- [X] también"))
    assert "• [ ] pendiente" in plano and "• [x] hecha" in plano and "• [x] también" in plano


def test_lineas_horizontales(pm):
    s = segs(pm, "antes\n\n---\n\n***\n\n- - -\n\ndespués")
    assert sum(1 for t, tg in s if "hr" in tg) == 3
    assert not any(tg == () and t.startswith("---") for t, tg in s)


def test_enlaces(pm):
    s = segs(pm, "[web](https://ejemplo.com/a) y [otro](notas.md) y [mal](javascript:alert(1)) y [ancla](#x)")
    web = [(t, tg) for t, tg in s if "link" in tg]
    assert [t for t, _ in web] == ["web", "otro"]
    assert "url:https://ejemplo.com/a" in web[0][1]
    assert "url:notas.md" in web[1][1]
    plano = "".join(t for t, _ in s)
    assert "mal" in plano and "ancla" in plano and "javascript" not in plano
    assert "](" not in plano


def test_imagenes_no_se_cargan(pm):
    s = segs(pm, "![gráfico de barras](img/g.png) y ![](x.png)")
    plano = "".join(t for t, _ in s)
    assert "[imagen: gráfico de barras]" in plano and "[imagen]" in plano
    assert "img/g.png" not in plano


def test_tabla_alineada(pm):
    s = segs(pm, "| Nombre | Nota |\n|---|---:|\n| Ana | 9 |\n| Bernardita | 10 |\n")
    filas = con_tag(s, "tabla").splitlines()
    assert len(filas) == 4
    assert len({len(f) for f in filas}) == 1  # todas del mismo ancho
    assert filas[0].startswith("| Nombre ") and set(filas[1]) <= set("|-")
    assert filas[2].endswith("|  9 |") or filas[2].rstrip().endswith("9 |")


def test_escapes(pm):
    plano = "".join(t for t, _ in segs(pm, r"\*no es cursiva\* y \# tampoco titulo y \`nada\`"))
    assert plano.startswith("*no es cursiva* y # tampoco titulo y `nada`")


def test_entradas_raras_no_tiran_excepcion(pm):
    raras = ["", "\n\n\n", "#", "# ", "```", "|", "|||\n|-|\n", "[", "[]()", "![", "**", "*", "_", "``",
             "> ", "- ", "1.", "- [", "\\", "[a](", "[a](b", "***\n***", "| a |\n|---|\n", "\x00\x01\ufffd",
             "\t\t- x\n\t\t\t- y", "    ", "~~~\n~~~", "![a](b)[c](d)![e", "`" * 50, "*" * 33]
    azar = random.Random(1234)
    alfabeto = list("#*_`>-+[]()!|\\\n \t:~.1a\u00f1<&\"'")
    raras += ["".join(azar.choice(alfabeto) for _ in range(azar.randint(0, 200))) for _ in range(400)]
    for r in raras:
        assert isinstance(pm.markdown_a_bloques(r), list)
        assert isinstance(pm.markdown_a_html(r), str)
    for r in (None, 12, b"xx"):
        pm.markdown_a_bloques(r)


# -------------------------------------------------------------------- HTML
def test_html_escapa_y_genera_estructura(pm):
    md = ("# Título <b> & más\n\nTexto con <script>alert(1)</script> y **fuerte** y [s](https://a.b/?x=1&y=2)\n\n"
          "- uno\n  - anidado\n- dos\n\n1. a\n2. b\n\n```\nif a < b && c > d:\n```\n\n"
          "| A | B |\n|---|---|\n| 1 | 2 |\n\n> cita\n\n---\n")
    h = pm.markdown_a_html(md)
    assert "<h1>Título &lt;b&gt; &amp; más</h1>" in h
    assert "<script>" not in h and "&lt;script&gt;" in h
    assert "<b>fuerte</b>" in h
    assert '<a href="https://a.b/?x=1&amp;y=2">s</a>' in h
    assert h.count("<ul>") == 2 and h.count("</ul>") == 2 and "<ol>" in h and "</ol>" in h
    assert "<pre><code>if a &lt; b &amp;&amp; c &gt; d:</code></pre>" in h
    assert "<table>" in h and "<th" in h and "<td" in h
    assert "<blockquote>cita</blockquote>" in h and "<hr>" in h


def test_html_no_enlaza_javascript(pm):
    h = pm.markdown_a_html("[x](javascript:alert(1)) y [y](notas.md)")
    assert "href" not in h and "javascript" not in h


# --------------------------------------------------------- PDF con el worker
MD_PRUEBA = ("# Apuntes de Física\n\nUn párrafo con ñandú, acentos y **negrita**.\n\n"
             "- uno\n- dos\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n```\ncodigo = 1\n```\n")


def test_md_pdf_de_punta_a_punta(pm, tmp_path):
    pymupdf = pytest.importorskip("pymupdf")
    salida = tmp_path / "Año nuevo" / "apuntes.pdf"
    r = pm.pdf_op({"op": "md_pdf", "html": pm.markdown_a_html(MD_PRUEBA), "salida": str(salida)},
                  lambda t, d: None, timeout=120, instalar=False)
    assert "error" not in r, r
    assert r["paginas"] >= 1 and salida.exists()
    assert not salida.with_name(salida.name + ".part").exists()
    d = pymupdf.open(str(salida))
    assert len(d) == r["paginas"]
    texto = d[0].get_text()
    assert "Apuntes de Física" in texto and "ñandú" in texto and "codigo = 1" in texto


def test_md_pdf_documento_largo_tiene_varias_paginas(pm, tmp_path):
    pymupdf = pytest.importorskip("pymupdf")
    md = "# Largo\n\n" + "\n\n".join(f"Párrafo número {i} " + "palabra " * 60 for i in range(60))
    salida = tmp_path / "largo.pdf"
    r = pm.pdf_op({"op": "md_pdf", "html": pm.markdown_a_html(md), "salida": str(salida)},
                  lambda t, d: None, timeout=120, instalar=False)
    assert "error" not in r and r["paginas"] >= 2
    assert len(pymupdf.open(str(salida))) == r["paginas"]


# ------------------------------------------------------------------ ventana
def esperar(app, cond, tope=20):
    fin = time.time() + tope
    while time.time() < fin:
        app.update()
        if cond():
            return True
        time.sleep(0.05)
    return False


def md_con_ñ(tmp_path, contenido=None):
    carpeta = tmp_path / "Año 2026 apuntes"
    carpeta.mkdir()
    f = carpeta / "Física ñandú.md"
    f.write_text(contenido or ("# Título grande\n\nHola **mundo** con `código` y un [enlace](https://ejemplo.com).\n\n"
                               "- uno\n- dos\n\nmundo otra vez, mundo.\n"), encoding="utf-8")
    return f


def test_ventana_del_lector(app, tmp_path):
    f = md_con_ñ(tmp_path)
    app.abrir_markdown(str(f))
    app.update()
    w = app._win_md
    plano = w.caja.get("1.0", "end")
    assert "Título grande" in plano and "Hola mundo con código" in plano
    assert "**" not in plano and "`" not in plano and "# " not in plano
    assert "h1" in w.caja.tag_names()
    assert w.caja.tag_ranges("h1")
    assert w.caja.tag_cget("h1", "foreground") == "#39ff14"
    assert str(w.caja.cget("state")) == "disabled"
    assert "UTF-8" in w.v_estado.get() and "palabra" in w.v_estado.get() and "línea" in w.v_estado.get()
    assert w.v_nombre.get() == f.name
    # buscar
    w.v_buscar.set("mundo")
    w.buscar()
    assert len(w.caja.tag_ranges("hallazgo")) == 6  # 3 coincidencias
    assert w.caja.tag_ranges("actual")
    primera = w.caja.tag_ranges("actual")[0]
    w.buscar()
    assert w.caja.tag_ranges("actual")[0] != primera
    # tamaño
    antes = w.f["base"].cget("size")
    w.cambiar_tam(1)
    assert w.f["base"].cget("size") == antes + 1
    w.cambiar_tam(-1)
    w.cambiar_tam(-1)
    assert w.f["base"].cget("size") == antes - 1
    # recargar toma el cambio
    f.write_text("# Otro título\n", encoding="utf-8")
    w.recargar()
    assert "Otro título" in w.caja.get("1.0", "end")


def test_lector_lee_cp1252_e_informa(app, tmp_path):
    f = tmp_path / "viejo.md"
    f.write_bytes("# Canción\n\nseñor ñandú\n".encode("cp1252"))
    app.abrir_markdown(str(f))
    w = app._win_md
    assert "señor ñandú" in w.caja.get("1.0", "end")
    assert "cp1252" in w.v_estado.get()


def test_lector_exporta_a_pdf(app, tmp_path, monkeypatch, pm):
    pymupdf = pytest.importorskip("pymupdf")
    f = md_con_ñ(tmp_path)
    app.v_abrir.set(False)
    app.abrir_markdown(str(f))
    w = app._win_md
    destino = tmp_path / "salida con ñ" / "Física ñandú.pdf"
    pedido = {}

    def pedir(**kw):
        pedido.update(kw)
        return str(destino)
    monkeypatch.setattr(pm.filedialog, "asksaveasfilename", pedir)
    w.exportar()
    assert pedido["initialfile"] == "Física ñandú.pdf"
    assert pedido["initialdir"] == str(f.parent)
    assert esperar(app, lambda: destino.exists() and not w.exportando, 20), w.v_estado.get()
    assert "Listo" in w.v_estado.get()
    assert "Título grande" in pymupdf.open(str(destino))[0].get_text()


def test_enlace_local_abre_en_el_mismo_lector(app, tmp_path):
    f = md_con_ñ(tmp_path, "# Uno\n\n[ir](otro doc.md)\n")
    (f.parent / "otro doc.md").write_text("# Segundo\n", encoding="utf-8")
    app.abrir_markdown(str(f))
    w = app._win_md
    w.abrir_enlace("otro%20doc.md")
    assert "Segundo" in w.caja.get("1.0", "end") and w.v_nombre.get() == "otro doc.md"
    w.abrir_enlace("no_existe.md")
    assert "No encuentro" in w.v_estado.get()
    w.abrir_enlace("programa.exe")
    assert "no lo abro" in w.v_estado.get()


def test_enlace_web_usa_el_navegador(app, tmp_path, monkeypatch, pm):
    abiertos = []
    monkeypatch.setattr(pm.webbrowser, "open", lambda u: abiertos.append(u))
    app.abrir_markdown(str(md_con_ñ(tmp_path)))
    app._win_md.abrir_enlace("https://ejemplo.com")
    assert abiertos == ["https://ejemplo.com"]


def test_agregar_un_md_abre_el_lector_y_no_entra_a_la_lista(app, tmp_path):
    f = md_con_ñ(tmp_path)
    app.agregar_rutas([str(f)])
    app.update()
    assert app._win_md is not None and app._win_md.winfo_exists()
    assert app.archivos == []
    assert "Título grande" in app._win_md.caja.get("1.0", "end")
    # mezclado con un Word: el Word va a la lista, el .md al lector
    docx = tmp_path / "TP 1.docx"
    docx.write_bytes(b"x")
    app.agregar_rutas([str(f), str(docx)])
    assert [a.name for a in app.archivos] == ["TP 1.docx"]


def test_accion_markdown_llana_y_satirica(pm):
    import random
    g = pm.Guion(random.Random(1))
    llano, sat = g.accion("markdown", nombre="a.md", n=12, exp="X", seg="")
    assert llano == "Se abrió «a.md»: 12 palabra(s)."
    assert "a.md" in sat and "12" in sat
