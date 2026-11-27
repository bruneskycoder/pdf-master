"""Escáner de fotos a PDF: perspectiva, detección de la hoja, modos, PDF y ventana."""
import time

import pytest

Image = pytest.importorskip("PIL.Image")
pytest.importorskip("pymupdf")
from PIL import ImageDraw  # noqa: E402

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
                                  "vista": str(tmp_path / "sal" / "o.png"), "lado": 400})
    im = Image.open(r["salida"])
    im.load()
    assert im.format == "JPEG" and im.size == (800, 600)
    assert Image.open(r["vista"]).size == (400, 300)


def test_procesar_recorta_y_rota(motor, tmp_path):
    ruta = hoja_sintetica(tmp_path / "h.png")
    base = {"ruta": str(ruta), "esquinas": [list(p) for p in CUADRILATERO], "modo": "original"}
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
    # crear el PDF
    destino = tmp_path / "salida" / "Escaneo.pdf"
    monkeypatch.setattr(pm.filedialog, "asksaveasfilename", lambda **kw: str(destino))
    app.v_abrir.set(False)
    esc.crear_pdf()
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
