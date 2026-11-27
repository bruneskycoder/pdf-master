"""Escáner de fotos a PDF: perspectiva, detección de la hoja, modos, PDF y ventana."""
import time

import pytest

Image = pytest.importorskip("PIL.Image")
pytest.importorskip("pymupdf")
from PIL import ImageDraw, ImageFilter  # noqa: E402

from sinteticas import (CASOS_DIFICILES, documento_renglones, foto_dificil,  # noqa: E402
                        pagina_con_sombra_y_sellos)
from test_gui import app  # noqa: E402,F401  (la ventana principal de las pruebas de interfaz)

CUADRILATERO = [(120, 90), (690, 130), (660, 520), (90, 480)]


def hoja_sintetica(ruta, tam=(800, 600), cuadrilatero=CUADRILATERO):
    im = Image.new("RGB", tam, (40, 40, 40))
    d = ImageDraw.Draw(im)
    d.polygon(cuadrilatero, fill=(238, 235, 225))
    for i in range(10):
        d.line([(160, 150 + i * 30), (600, 158 + i * 30)], fill=(25, 25, 25), width=3)
    im.save(ruta)
    return ruta


def aplicar(coef, x, y):
    a, b, c, d, e, f, g, h = coef
    den = g * x + h * y + 1
    return (a * x + b * y + c) / den, (d * x + e * y + f) / den


def test_perspectiva_identidad(motor):
    cuadrado = [(0, 0), (100, 0), (100, 100), (0, 100)]
    coef = motor["coef_perspectiva"](cuadrado, cuadrado)
    for x, y in ((0, 0), (100, 0), (50, 50), (30, 80)):
        sx, sy = aplicar(coef, x, y)
        assert abs(sx - x) < 0.01 and abs(sy - y) < 0.01


def test_perspectiva_lleva_las_esquinas_al_rectangulo(motor):
    destino = [(0, 0), (500, 0), (500, 700), (0, 700)]
    coef = motor["coef_perspectiva"](destino, CUADRILATERO)
    for d, o in zip(destino, CUADRILATERO):
        sx, sy = aplicar(coef, *d)
        assert abs(sx - o[0]) < 1 and abs(sy - o[1]) < 1


def test_perspectiva_degenerada_falla_con_mensaje(motor):
    with pytest.raises(ValueError):
        motor["coef_perspectiva"]([(0, 0), (1, 0), (2, 0), (3, 0)], [(5, 5)] * 4)


def test_detecta_la_hoja(motor, tmp_path):
    im = Image.open(hoja_sintetica(tmp_path / "h.png"))
    esquinas = motor["detectar_hoja"](im)
    for e, r in zip(esquinas, CUADRILATERO):
        assert abs(e[0] - r[0]) < 15 and abs(e[1] - r[1]) < 15, (esquinas, CUADRILATERO)


@pytest.mark.parametrize("color", [(255, 255, 255), (120, 120, 120), (0, 0, 0)])
def test_sin_contraste_devuelve_la_imagen_completa(motor, color):
    im = Image.new("RGB", (800, 600), color)
    assert motor["detectar_hoja"](im) == [[0, 0], [800, 0], [800, 600], [0, 600]]


def test_preparar_devuelve_vista_y_esquinas(motor, tmp_path):
    ruta = hoja_sintetica(tmp_path / "h.png")
    r = motor["op_img_preparar"]({"rutas": [str(ruta), str(tmp_path / "no_existe.jpg")],
                                  "dir": str(tmp_path / "v")})
    ok, mala = r["fotos"]
    assert (ok["ancho"], ok["alto"]) == (800, 600)
    assert max(ok["vista_ancho"], ok["vista_alto"]) == 800  # la foto es más chica que 900: no se agranda
    assert Image.open(ok["vista"]).size == (ok["vista_ancho"], ok["vista_alto"])
    assert len(ok["esquinas"]) == 4
    assert "error" in mala


@pytest.mark.parametrize("modo", ["original", "aclarar", "byn", "gris"])
def test_procesar_en_cada_modo(motor, tmp_path, modo):
    ruta = hoja_sintetica(tmp_path / "h.png")
    esq = [[0, 0], [800, 0], [800, 600], [0, 600]]
    r = motor["op_img_procesar"]({"ruta": str(ruta), "esquinas": esq, "modo": modo,
                                  "salida": str(tmp_path / "sal" / "o.jpg"),
                                  "vista": str(tmp_path / "sal" / "o.png"), "lado": 400,
                                  "enderezar": False})  # acá se prueba la geometría, no el giro fino
    im = Image.open(r["salida"])
    im.load()
    assert im.format == "JPEG" and im.size == (800, 600)
    assert Image.open(r["vista"]).size == (400, 300)


def test_procesar_recorta_y_rota(motor, tmp_path):
    ruta = hoja_sintetica(tmp_path / "h.png")
    base = {"ruta": str(ruta), "esquinas": [list(p) for p in CUADRILATERO], "modo": "original",
            "enderezar": False}
    a = motor["op_img_procesar"](dict(base, salida=str(tmp_path / "a.jpg")))
    b = motor["op_img_procesar"](dict(base, salida=str(tmp_path / "b.jpg"), rotacion=90))
    assert (b["ancho"], b["alto"]) == (a["alto"], a["ancho"])


def test_procesar_limita_el_lado_mayor(motor, tmp_path):
    ruta = tmp_path / "grande.png"
    Image.new("RGB", (4000, 3000), (200, 200, 200)).save(ruta)
    r = motor["op_img_procesar"]({"ruta": str(ruta), "modo": "original",
                                  "esquinas": [[0, 0], [4000, 0], [4000, 3000], [0, 3000]],
                                  "salida": str(tmp_path / "g.jpg")})
    assert (r["ancho"], r["alto"]) == (2480, 1860)


def test_byn_solo_blanco_y_negro(motor, tmp_path):
    ruta = hoja_sintetica(tmp_path / "h.png")
    r = motor["op_img_procesar"]({"ruta": str(ruta), "esquinas": [list(p) for p in CUADRILATERO],
                                  "modo": "byn", "salida": str(tmp_path / "bn.jpg")})
    im = Image.open(r["salida"]).convert("L")
    h = im.histogram()
    extremos = sum(h[:20]) + sum(h[236:])
    assert extremos / float(im.size[0] * im.size[1]) > 0.98


def test_aclarar_blanquea_el_papel(motor, tmp_path):
    ruta = hoja_sintetica(tmp_path / "h.png")
    r = motor["op_img_procesar"]({"ruta": str(ruta), "esquinas": [list(p) for p in CUADRILATERO],
                                  "modo": "aclarar", "salida": str(tmp_path / "ac.jpg")})
    im = Image.open(r["salida"]).convert("L")
    assert im.getpixel((10, 10)) > 235  # el papel quedó blanco


def test_pdf_de_tres_imagenes(motor, tmp_path):
    pymupdf = pytest.importorskip("pymupdf")
    rutas = []
    for i, tam in enumerate([(800, 600), (600, 800), (500, 500)]):
        f = tmp_path / f"i{i}.jpg"
        Image.new("RGB", tam, (200, 220, 200)).save(f, "JPEG")
        rutas.append(str(f))
    salida = tmp_path / "nueva" / "carpeta" / "doc.pdf"
    r = motor["op_img_pdf"]({"imagenes": rutas, "salida": str(salida), "pagina": "a4"})
    assert r["paginas"] == 3 and salida.exists() and not salida.with_name("doc.pdf.part").exists()
    d = pymupdf.open(str(salida))
    assert [(round(p.rect.width), round(p.rect.height)) for p in d] == [(842, 595), (595, 842), (595, 842)]
    d.close()
    foto = tmp_path / "foto.pdf"
    motor["op_img_pdf"]({"imagenes": rutas[:1], "salida": str(foto), "pagina": "foto"})
    d = pymupdf.open(str(foto))
    assert (round(d[0].rect.width), round(d[0].rect.height)) == (192, 144)  # 800x600 px a 300 dpi
    d.close()


def test_pdf_op_pide_imagenes_en_un_subproceso(pm, tmp_path):
    """El camino real: pdf_op arranca el worker en otro proceso y vuelve el JSON."""
    ruta = hoja_sintetica(tmp_path / "h.png")
    r = pm.pdf_op({"op": "img_preparar", "rutas": [str(ruta)], "dir": str(tmp_path / "v")},
                  lambda t, d: None, instalar=False)
    assert "error" not in r, r
    assert len(r["fotos"][0]["esquinas"]) == 4


def _esperar(app, condicion, tope=20):
    fin = time.time() + tope
    while time.time() < fin:
        app.update()
        if condicion():
            return True
        time.sleep(0.05)
    return False


def test_ventana_del_escaner(app, pm, tmp_path, monkeypatch):
    app.abrir_escaner()
    esc = app._win_esc
    app.update()
    fotos = [hoja_sintetica(tmp_path / f"f{i}.png") for i in range(2)]
    esc.agregar_fotos([str(f) for f in fotos])
    assert _esperar(app, lambda: len(esc.fotos) == 2)
    assert esc.lista.size() == 2 and esc.actual == 0
    assert len(esc.lona.find_withtag("asa")) == 4
    # arrastre de un asa: va al punto pedido (en coordenadas de la foto original)
    ox, oy, e = esc.origen
    esc._arrastrar(2, ox + 300 * e, oy + 250 * e)
    assert [round(v) for v in esc.fotos[0]["esquinas"][2]] == [300, 250]
    esc.detectar()
    assert esc.fotos[0]["esquinas"] == esc.fotos[0]["auto"]
    esc.toda_la_hoja()
    assert esc.fotos[0]["esquinas"][2] == [800, 600]
    # el modo se aplica a todas si se pide
    esc.v_modo.set("byn")
    esc.v_todas.set(True)
    esc._modo_cambiado()
    assert [f["modo"] for f in esc.fotos] == ["byn", "byn"]
    # ver el resultado y volver
    esc.ver_resultado()
    assert _esperar(app, lambda: esc.viendo)
    assert not esc.lona.find_withtag("asa")
    esc.ver_resultado()
    assert not esc.viendo and len(esc.lona.find_withtag("asa")) == 4
    # crear el PDF: se abre el diálogo de opciones y [ GUARDAR PDF ] pide dónde guardar
    destino = tmp_path / "salida" / "Escaneo.pdf"
    monkeypatch.setattr(pm.filedialog, "asksaveasfilename", lambda **kw: str(destino))
    app.v_abrir.set(False)
    esc.crear_pdf()
    app.update()
    assert esc.dialogo is not None and esc.dialogo.winfo_exists()
    esc.dialogo.guardar()
    assert _esperar(app, lambda: destino.exists())
    assert _esperar(app, lambda: not esc.creando)
    pymupdf = pytest.importorskip("pymupdf")
    d = pymupdf.open(str(destino))
    assert len(d) == 2
    d.close()
    assert "Listo" in esc.v_estado.get()
    tmp = esc.tmp
    esc.destroy()
    app.update()
    assert not tmp.exists()  # los temporales se borran al cerrar


def test_heic_se_explica(app, pm, tmp_path, monkeypatch):
    avisos = []
    monkeypatch.setattr(pm.messagebox, "showwarning", lambda t, m, **kw: avisos.append((t, m)))
    app.abrir_escaner()
    app._win_esc.agregar_fotos([str(tmp_path / "IMG_0001.HEIC")])
    assert avisos and "JPG" in avisos[0][1]
    assert app._win_esc.fotos == []
    app._win_esc.destroy()


# ---------------------------------------------------------------- detección difícil
def _error_esquinas(esquinas, esperado, tam):
    diag = (tam[0] ** 2 + tam[1] ** 2) ** 0.5
    return max(((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2) ** 0.5 / diag * 100
               for a, b in zip(esquinas, esperado))


@pytest.mark.parametrize("caso", CASOS_DIFICILES)
def test_detecta_hojas_dificiles(motor, caso):
    im, esperado = foto_dificil(caso)
    esquinas, metodo, puntaje = motor["detectar_hoja_info"](im)
    err = _error_esquinas(esquinas, esperado, im.size)
    assert metodo in ("bordes", "mascara"), (caso, metodo, puntaje)
    assert err < 2.0, (caso, metodo, esquinas, esperado, err)


def test_detecta_con_sombra_que_la_mascara_sola_no_resuelve(motor):
    """El caso real: la parte alta de la hoja en sombra, casi tan oscura como la mesa."""
    from PIL import ImageChops
    from sinteticas import MESA, sombra_vertical
    im, esperado = foto_dificil("sombra")
    # se repinta la hoja con 66 % de sombra arriba (gris ~ 80 contra mesa ~ 35)
    im2 = Image.new("RGB", im.size, MESA)
    capa = Image.new("RGB", im.size, (240, 237, 228))
    capa = ImageChops.multiply(capa, sombra_vertical(im.size, 0.66, hasta=200.0))
    mascara = Image.new("L", im.size, 0)
    ImageDraw.Draw(mascara).polygon(esperado, fill=255)
    im2.paste(capa, (0, 0), mascara)
    esquinas, metodo, _ = motor["detectar_hoja_info"](im2)
    assert metodo == "bordes"  # la máscara sola toma la zona sombreada por mesa
    assert _error_esquinas(esquinas, esperado, im.size) < 2.0, (metodo, esquinas, esperado)


def test_detector_devuelve_metodo_completa_sin_hoja(motor):
    esquinas, metodo, _ = motor["detectar_hoja_info"](Image.new("RGB", (800, 600), (120, 120, 120)))
    assert metodo == "completa" and esquinas == [[0, 0], [800, 0], [800, 600], [0, 600]]


def test_detector_es_rapido_con_una_foto_de_4000_por_3000(motor):
    im, esperado = foto_dificil("sombra", tam=(4000, 3000))
    t = time.time()
    esquinas, metodo, _ = motor["detectar_hoja_info"](im)
    seg = time.time() - t
    print(f"detección 4000x3000: {seg:.2f} s ({metodo})")
    assert seg < 3.0  # en la máquina de desarrollo ronda 0,3 s; el tope es holgado para CI lentos
    assert _error_esquinas(esquinas, esperado, im.size) < 2.0


def test_preparar_informa_el_metodo(motor, tmp_path):
    im, _ = foto_dificil("rotada8")
    im.save(tmp_path / "r.jpg", quality=92)
    f = motor["op_img_preparar"]({"rutas": [str(tmp_path / "r.jpg")], "dir": str(tmp_path / "v")})["fotos"][0]
    assert f["metodo"] in ("bordes", "mascara") and len(f["esquinas"]) == 4


# ---------------------------------------------------------------------- enderezar
@pytest.mark.parametrize("grados", [4.0, -6.0])
def test_estima_y_corrige_la_inclinacion(motor, grados):
    im = documento_renglones(grados)
    ang = motor["estimar_inclinacion"](im)
    assert abs(ang + grados) <= 0.5, (grados, ang)  # el giro que lo endereza es el opuesto
    recto = motor["enderezar_imagen"](im, ang)
    assert abs(motor["estimar_inclinacion"](recto)) < 0.3
    assert recto.size[0] <= im.size[0] + 1 and recto.size[1] <= im.size[1] + 1


def test_documento_sin_texto_no_rota(motor, tmp_path):
    blanco = Image.new("RGB", (900, 1200), (250, 250, 250))
    assert motor["estimar_inclinacion"](blanco) == 0.0
    blanco.save(tmp_path / "b.png")
    r = motor["op_img_procesar"]({"ruta": str(tmp_path / "b.png"), "modo": "aclarar",
                                  "esquinas": [[0, 0], [900, 0], [900, 1200], [0, 1200]],
                                  "salida": str(tmp_path / "s.jpg")})
    assert r["angulo"] == 0.0 and (r["ancho"], r["alto"]) == (900, 1200)


def test_procesar_endereza_por_defecto_y_se_puede_apagar(motor, tmp_path):
    documento_renglones(4.0).save(tmp_path / "d.png")
    base = {"ruta": str(tmp_path / "d.png"), "modo": "aclarar",
            "esquinas": [[0, 0], [900, 0], [900, 1200], [0, 1200]]}
    r = motor["op_img_procesar"](dict(base, salida=str(tmp_path / "a.jpg")))
    assert abs(r["angulo"] + 4.0) <= 0.5 and r["ancho"] < 900
    r2 = motor["op_img_procesar"](dict(base, salida=str(tmp_path / "b.jpg"), enderezar=False))
    assert r2["angulo"] == 0.0 and (r2["ancho"], r2["alto"]) == (900, 1200)
    # y el resultado quedó derecho
    assert abs(motor["estimar_inclinacion"](Image.open(tmp_path / "a.jpg"))) < 0.3


# ----------------------------------------------------------------------- aclarar
def test_aclarar_con_sombra_fuerte_y_sellos_de_color(motor):
    im = pagina_con_sombra_y_sellos()
    out = motor["aclarar"](im)

    def luminancia(caja):
        c = out.convert("L").crop(caja)
        h = c.histogram()
        return sum(i * v for i, v in enumerate(h)) / float(c.size[0] * c.size[1])
    zonas = {"arriba (sombra)": (20, 20, 300, 120), "medio": (20, 600, 100, 900),
             "abajo (clara)": (20, 1330, 900, 1390), "derecha": (930, 300, 990, 1000)}
    lum = {k: luminancia(v) for k, v in zonas.items()}
    assert all(v > 235 for v in lum.values()), lum
    assert max(lum.values()) - min(lum.values()) < 10, lum
    # la tinta queda negra
    assert max(out.getpixel((300, 200))) < 40
    # sellos: azul y rojo siguen siendo de color
    r, g, b = out.getpixel((780, 140))
    assert b > r + 100 and b > g + 100, (r, g, b)
    r, g, b = out.getpixel((690, 1190))
    assert r > g + 100 and r > b + 100, (r, g, b)
    r, g, b = out.getpixel((104, 1230))  # aro violeta
    assert r > g + 60 and b > g + 60, (r, g, b)


# ------------------------------------------------------------------ exportación
def _jpgs(tmp_path, tams=((800, 600), (600, 800), (500, 500))):
    rutas = []
    for i, tam in enumerate(tams):
        f = tmp_path / f"e{i}.jpg"
        Image.new("RGB", tam, (200, 220, 200)).save(f, "JPEG")
        rutas.append(str(f))
    return rutas


def _paginas(ruta):
    pymupdf = pytest.importorskip("pymupdf")
    d = pymupdf.open(str(ruta))
    try:
        return [(p.rect.width, p.rect.height) for p in d]
    finally:
        d.close()


@pytest.mark.parametrize("tamano,esperado", [("a4", (595, 842)), ("carta", (612, 792)),
                                             ("oficio", (612, 1008))])
def test_pdf_cada_tamano(motor, tmp_path, tamano, esperado):
    rutas = _jpgs(tmp_path, ((600, 800), (600, 800)))
    sal = tmp_path / "t.pdf"
    motor["op_img_pdf"]({"imagenes": rutas, "salida": str(sal), "tamano": tamano})
    for w, h in _paginas(sal):
        assert abs(w - esperado[0]) <= 1 and abs(h - esperado[1]) <= 1, (w, h)


def test_pdf_orientacion(motor, tmp_path):
    rutas = _jpgs(tmp_path)  # horizontal, vertical, cuadrada
    sal = tmp_path / "o.pdf"
    motor["op_img_pdf"]({"imagenes": rutas, "salida": str(sal), "tamano": "carta"})
    assert [w > h for w, h in _paginas(sal)] == [True, False, False]  # automática
    motor["op_img_pdf"]({"imagenes": rutas, "salida": str(sal), "tamano": "carta",
                         "orientacion": "horizontal"})
    assert all(abs(w - 792) <= 1 and abs(h - 612) <= 1 for w, h in _paginas(sal))
    motor["op_img_pdf"]({"imagenes": rutas, "salida": str(sal), "tamano": "carta",
                         "orientacion": "vertical"})
    assert all(abs(w - 612) <= 1 and abs(h - 792) <= 1 for w, h in _paginas(sal))


def test_pdf_margen_de_un_centimetro(motor, tmp_path):
    pymupdf = pytest.importorskip("pymupdf")
    rutas = _jpgs(tmp_path, ((600, 800),))
    m = 72 / 2.54
    for margen, minimo in ((0, 0.0), (1, m)):
        sal = tmp_path / f"m{margen}.pdf"
        motor["op_img_pdf"]({"imagenes": rutas, "salida": str(sal), "tamano": "a4",
                             "margen_cm": margen})
        d = pymupdf.open(str(sal))
        x0, y0, x1, y1 = d[0].get_image_info()[0]["bbox"]
        w, h = d[0].rect.width, d[0].rect.height
        d.close()
        assert x0 >= minimo - 1 and y0 >= minimo - 1, (margen, x0, y0)
        assert x1 <= w - minimo + 1 and y1 <= h - minimo + 1, (margen, x1, y1)
        if margen:
            # la imagen llena el hueco reducido en alguno de los dos sentidos
            assert abs(x0 - m) <= 1.5 or abs(y0 - m) <= 1.5
        else:
            assert x1 - x0 > w - 3 or y1 - y0 > h - 3


def test_pdf_marca_de_agua_con_acentos_y_en_diagonal(motor, tmp_path):
    pymupdf = pytest.importorskip("pymupdf")
    rutas = _jpgs(tmp_path, ((600, 800),))
    sal = tmp_path / "w.pdf"
    motor["op_img_pdf"]({"imagenes": rutas, "salida": str(sal), "tamano": "a4",
                         "marca": "COPIA ñandú"})
    d = pymupdf.open(str(sal))
    assert "COPIA ñandú" in d[0].get_text()
    lineas = [l for b in d[0].get_text("dict")["blocks"] for l in b.get("lines", [])]
    dx, dy = lineas[0]["dir"]
    assert abs(abs(dx) - 0.7071) < 0.01 and abs(abs(dy) - 0.7071) < 0.01  # 45 grados
    # centrada
    x0, y0, x1, y1 = lineas[0]["bbox"]
    assert abs((x0 + x1) / 2 - 297.5) < 10 and abs((y0 + y1) / 2 - 421) < 10
    d.close()
    sin = tmp_path / "sin.pdf"
    motor["op_img_pdf"]({"imagenes": rutas, "salida": str(sin), "tamano": "a4"})
    d = pymupdf.open(str(sin))
    assert d[0].get_text().strip() == ""
    d.close()


def test_pdf_calidad_liviana_pesa_menos_que_alta(motor, tmp_path):
    import random
    azar = random.Random(3)
    grande = Image.frombytes("RGB", (3500, 2600), azar.randbytes(3500 * 2600 * 3)).filter(
        ImageFilter.GaussianBlur(2))
    grande.save(tmp_path / "g.png")
    pesos = {}
    for nombre, lado, cal in (("alta", 3300, 90), ("liviana", 1600, 70)):
        r = motor["op_img_procesar"]({"ruta": str(tmp_path / "g.png"), "modo": "original",
                                      "esquinas": [[0, 0], [3500, 0], [3500, 2600], [0, 2600]],
                                      "salida": str(tmp_path / f"{nombre}.jpg"), "lado_max": lado,
                                      "calidad": cal, "enderezar": False})
        assert max(r["ancho"], r["alto"]) == lado
        sal = tmp_path / f"{nombre}.pdf"
        motor["op_img_pdf"]({"imagenes": [r["salida"]], "salida": str(sal), "tamano": "a4"})
        pesos[nombre] = sal.stat().st_size
    assert pesos["liviana"] < pesos["alta"], pesos


def test_pdf_tamano_desconocido_falla_con_mensaje(motor, tmp_path):
    with pytest.raises(ValueError):
        motor["op_img_pdf"]({"imagenes": _jpgs(tmp_path, ((10, 10),)), "salida": str(tmp_path / "x.pdf"),
                             "tamano": "a0"})


def test_rango_paginas(pm):
    assert pm.rango_paginas("1,3", 3) == [0, 2]
    assert pm.rango_paginas("2-3, 1", 3) == [1, 2, 0]
    assert pm.rango_paginas("2-", 4) == [1, 2, 3]
    for malo in ("9", "0", "3-1", "a", "", "1-9"):
        with pytest.raises(ValueError):
            pm.rango_paginas(malo, 2)


# ------------------------------------------------------------------------- diálogo
def _abrir_con_fotos(app, tmp_path, n):
    app.abrir_escaner()
    esc = app._win_esc
    app.update()
    esc.agregar_fotos([str(hoja_sintetica(tmp_path / f"f{i}.png")) for i in range(n)])
    assert _esperar(app, lambda: len(esc.fotos) == n)
    return esc


def test_dialogo_valida_el_rango_y_no_se_cierra(app, pm, tmp_path, monkeypatch):
    esc = _abrir_con_fotos(app, tmp_path, 2)
    pedidos = []
    monkeypatch.setattr(pm.filedialog, "asksaveasfilename", lambda **kw: pedidos.append(kw) or "")
    esc.crear_pdf()
    app.update()
    dlg = esc.dialogo
    assert dlg.winfo_exists() and dlg.grab_current() in (dlg, None)
    assert dlg.v_calidad.get() == "media" and dlg.v_marca_txt.get() == "COPIA"
    assert dlg.v_endereza.get() is True
    dlg.v_rango.set("personalizado")
    dlg.v_rango_txt.set("9")
    dlg.guardar()
    app.update()
    assert dlg.winfo_exists() and "2 foto" in dlg.v_error.get()
    assert pedidos == []  # no se llegó a pedir el archivo
    dlg.v_rango_txt.set("1-2")
    dlg.v_marca.set("texto")
    dlg.v_marca_txt.set("   ")
    dlg.guardar()
    assert dlg.winfo_exists() and "marca" in dlg.v_error.get().lower()
    dlg.destroy()
    esc.destroy()


def test_dialogo_recuerda_las_opciones_y_genera_el_pdf(app, pm, tmp_path, monkeypatch):
    import json
    pymupdf = pytest.importorskip("pymupdf")
    esc = _abrir_con_fotos(app, tmp_path, 3)
    destino = tmp_path / "sal" / "Doc.pdf"
    monkeypatch.setattr(pm.filedialog, "asksaveasfilename", lambda **kw: str(destino))
    app.v_abrir.set(False)
    esc.crear_pdf()
    app.update()
    dlg = esc.dialogo
    dlg.v_tamano.set("carta")
    dlg.v_orient.set("vertical")
    dlg.v_margen.set(1)
    dlg.v_rango.set("personalizado")
    dlg.v_rango_txt.set("1,3")
    dlg.v_marca.set("texto")
    dlg.v_marca_txt.set("BORRADOR")
    dlg.v_calidad.set("liviana")
    dlg.v_endereza.set(False)
    dlg.guardar()
    assert not dlg.winfo_exists()
    assert _esperar(app, lambda: destino.exists(), tope=30)
    assert _esperar(app, lambda: not esc.creando, tope=30)
    d = pymupdf.open(str(destino))
    assert len(d) == 2  # el rango 1,3 de 3 fotos
    assert abs(d[0].rect.width - 612) <= 1 and abs(d[0].rect.height - 792) <= 1
    assert "BORRADOR" in d[0].get_text()
    d.close()
    # quedó en la config, con prefijo esc_
    guardado = json.loads(pm.CONFIG.read_text(encoding="utf-8"))
    assert guardado["esc_tamano"] == "carta" and guardado["esc_orient"] == "vertical"
    assert guardado["esc_margen"] == 1 and guardado["esc_marca"] is True
    assert guardado["esc_marca_texto"] == "BORRADOR" and guardado["esc_calidad"] == "liviana"
    assert guardado["esc_enderezar"] is False
    # y el diálogo siguiente arranca con esas opciones
    esc.crear_pdf()
    app.update()
    dlg2 = esc.dialogo
    assert (dlg2.v_tamano.get(), dlg2.v_orient.get(), dlg2.v_margen.get()) == ("carta", "vertical", 1)
    assert dlg2.v_marca.get() == "texto" and dlg2.v_marca_txt.get() == "BORRADOR"
    assert dlg2.v_calidad.get() == "liviana" and dlg2.v_endereza.get() is False
    dlg2.destroy()
    esc.destroy()


def test_la_ventana_del_escaner_ya_no_pregunta_el_tamano(app, tmp_path):
    app.abrir_escaner()
    esc = app._win_esc
    app.update()
    assert not hasattr(esc, "v_pagina")
    esc.destroy()


@pytest.mark.parametrize("modo", ["aclarar", "original"])
def test_el_resultado_no_tiene_bordes_oscuros(motor, tmp_path, modo):
    """Sin restos de mesa de 1 o 2 px en los bordes tras la perspectiva (y el giro)."""
    im, _ = foto_dificil("sombra")
    im.save(tmp_path / "s.png")
    esq = motor["detectar_hoja"](im)
    r = motor["op_img_procesar"]({"ruta": str(tmp_path / "s.png"), "esquinas": esq, "modo": modo,
                                  "salida": str(tmp_path / "r.png.jpg")})
    L = Image.open(r["salida"]).convert("L")
    w, h = L.size
    franjas = {"arriba": (0, 0, w, 4), "abajo": (0, h - 4, w, h), "izq": (0, 0, 4, h),
               "der": (w - 4, 0, w, h)}
    px = L.load()
    for nombre, caja in franjas.items():
        if modo == "aclarar":
            minimo = min(L.crop(caja).getdata())
            assert minimo >= 200, (nombre, minimo)
            continue
        # el papel sintético tiene sombra: cada píxel de la franja se compara con el papel
        # 20 px más adentro, en la misma fila o columna
        peor = 9.9
        for x in range(caja[0], caja[2]):
            for y in range(caja[1], caja[3]):
                ix = x + {"izq": 20, "der": -20}.get(nombre, 0)
                iy = y + {"arriba": 20, "abajo": -20}.get(nombre, 0)
                peor = min(peor, px[x, y] / float(max(1, px[ix, iy])))
        assert peor > 0.8, (nombre, peor)
