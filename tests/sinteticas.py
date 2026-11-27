"""Imágenes sintéticas para las pruebas y las capturas del escáner. Todas se dibujan con Pillow:
no hay fotos reales ni datos de nadie (el texto son renglones, sin letras)."""
import math
import random

from PIL import Image, ImageChops, ImageDraw, ImageFilter

HOJA = (240, 237, 228)
MESA = (38, 32, 28)


def girar(puntos, grados, centro):
    a = math.radians(grados)
    c, s = math.cos(a), math.sin(a)
    return [(centro[0] + (x - centro[0]) * c - (y - centro[1]) * s,
             centro[1] + (x - centro[0]) * s + (y - centro[1]) * c) for x, y in puntos]


def renglones(d, cuad, n=14, color=(30, 30, 30)):
    """Renglones paralelos al borde superior de la hoja (interpolación bilineal)."""
    tl, tr, br, bl = cuad

    def pt(u, v):
        arriba = (tl[0] + (tr[0] - tl[0]) * u, tl[1] + (tr[1] - tl[1]) * u)
        abajo = (bl[0] + (br[0] - bl[0]) * u, bl[1] + (br[1] - bl[1]) * u)
        return (arriba[0] + (abajo[0] - arriba[0]) * v, arriba[1] + (abajo[1] - arriba[1]) * v)
    for i in range(n):
        v = 0.12 + 0.76 * i / (n - 1)
        d.line([pt(0.12, v), pt(0.88 if i % 3 else 0.6, v)], fill=color, width=3)


def sombra_vertical(tam, oscuridad, hasta=140.0):
    """Máscara multiplicadora: (1 - oscuridad) arriba, 1 desde `hasta` filas (de 256) hacia abajo."""
    g = Image.linear_gradient("L").resize(tam)
    k = g.point([round(255 * (1 - oscuridad + oscuridad * min(1.0, i / hasta))) for i in range(256)])
    return Image.merge("RGB", (k, k, k))


CASOS_DIFICILES = ["sombra", "toca_borde", "rotada3", "rotada8", "rotada15", "trapecio",
                   "ruido", "sellos"]


def foto_dificil(nombre, tam=(1200, 900)):
    """(imagen, esquinas esperadas) de una hoja difícil sobre una mesa oscura. Las esquinas
    siguen el orden sup-izq, sup-der, inf-der, inf-izq; si la hoja toca el borde del encuadre,
    la esquina esperada es la intersección con ese borde."""
    W, H = tam
    centro = (W / 2, H / 2)
    q = [(0.18 * W, 0.10 * H), (0.82 * W, 0.12 * H), (0.80 * W, 0.90 * H), (0.20 * W, 0.88 * H)]
    if nombre.startswith("rotada"):
        grados = int(nombre[6:])
        m = {3: 0.12, 8: 0.15, 15: 0.18}[grados]
        q = girar([(0.2 * W, m * H), (0.8 * W, m * H), (0.8 * W, (1 - m) * H),
                   (0.2 * W, (1 - m) * H)], grados, centro)
    elif nombre == "trapecio":
        q = [(0.28 * W, 0.10 * H), (0.72 * W, 0.10 * H), (0.90 * W, 0.90 * H), (0.10 * W, 0.90 * H)]
    elif nombre == "toca_borde":
        q = [(0.10 * W, 0.08 * H), (0.90 * W, 0.11 * H), (0.93 * W, 1.10 * H), (0.07 * W, 1.07 * H)]
    elif nombre in ("sombra", "sellos"):
        q = [(0.09 * W, 0.04 * H), (0.93 * W, 0.07 * H), (0.95 * W, 0.97 * H), (0.06 * W, 0.94 * H)]
    im = Image.new("RGB", tam, MESA)
    if nombre == "ruido":
        azar = random.Random(7)
        ruido = Image.frombytes("L", tam, azar.randbytes(W * H)).point(
            [30 + i // 10 for i in range(256)])  # 30..55 al azar, píxel por píxel
        im = Image.merge("RGB", (ruido, ruido.point(lambda v: v - 4), ruido.point(lambda v: v - 8)))
        im = im.filter(ImageFilter.GaussianBlur(0.6))
    capa = Image.new("RGB", tam, HOJA)
    if nombre in ("sombra", "sellos"):
        capa = ImageChops.multiply(capa, sombra_vertical(tam, 0.40))  # 40 % más oscura arriba
    d = ImageDraw.Draw(capa)
    renglones(d, q)
    if nombre == "sellos":  # sellos grandes y oscuros cerca del borde
        for cx, cy, r, col in ((0.14 * W, 0.20 * H, 0.09 * W, (60, 40, 150)),
                               (0.86 * W, 0.82 * H, 0.10 * W, (170, 40, 50)),
                               (0.13 * W, 0.86 * H, 0.08 * W, (30, 30, 30)),
                               (0.87 * W, 0.14 * H, 0.07 * W, (20, 20, 20))):
            d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=col, width=7)
            d.ellipse([cx - r * 0.6, cy - r * 0.6, cx + r * 0.6, cy + r * 0.6], fill=col)
    mascara = Image.new("L", tam, 0)
    ImageDraw.Draw(mascara).polygon(q, fill=255)
    im.paste(capa, (0, 0), mascara)
    if nombre == "ruido":  # un objeto claro chico al costado
        ImageDraw.Draw(im).rectangle([0.02 * W, 0.40 * H, 0.07 * W, 0.55 * H], fill=(200, 200, 190))
    esperado = q
    if nombre == "toca_borde":
        def corta(a, b, y):
            t = (y - a[1]) / (b[1] - a[1])
            return (a[0] + (b[0] - a[0]) * t, y)
        esperado = [q[0], q[1], corta(q[1], q[2], H), corta(q[0], q[3], H)]
    return im, [(float(x), float(y)) for x, y in esperado]


def documento_renglones(grados=0.0, tam=(900, 1200)):
    """Página blanca con renglones negros finos (rectángulos), girada `grados` (antihorario)."""
    im = Image.new("RGB", tam, (255, 255, 255))
    d = ImageDraw.Draw(im)
    W, H = tam
    for i in range(36):
        y = int(0.10 * H + i * 0.022 * H)
        d.rectangle([int(0.12 * W), y, int(0.88 * W if i % 5 else 0.55 * W), y + 2], fill=(10, 10, 10))
    if grados:
        im = im.rotate(grados, resample=Image.BICUBIC, fillcolor=(255, 255, 255))
    return im


def pagina_con_sombra_y_sellos(tam=(1000, 1400)):
    """Hoja cremosa, 40 % más oscura arriba, con renglones y sellos azul, violeta y rojo."""
    W, H = tam
    capa = Image.new("RGB", tam, (238, 235, 225))
    d = ImageDraw.Draw(capa)
    for i in range(30):
        d.line([(120, 200 + i * 30), (880 if i % 4 else 500, 200 + i * 30)], fill=(25, 25, 25),
               width=3)
    azul, violeta, rojo = (40, 60, 170), (150, 40, 170), (190, 40, 50)
    d.ellipse([700, 60, 860, 220], outline=azul, width=8)
    d.ellipse([740, 100, 820, 180], fill=azul)
    d.ellipse([100, 1150, 260, 1310], outline=violeta, width=8)
    d.ellipse([600, 1100, 780, 1280], outline=rojo, width=8)
    d.ellipse([650, 1150, 730, 1230], fill=rojo)
    return ImageChops.multiply(capa, sombra_vertical(tam, 0.40))
