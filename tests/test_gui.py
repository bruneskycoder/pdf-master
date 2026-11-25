"""La interfaz arranca, usa las fuentes incluidas y responde a lo básico.
Se saltea si no hay pantalla (en Linux, correr con xvfb-run)."""
import sys

import pytest

tk = pytest.importorskip("tkinter")
ES_WINDOWS = sys.platform.startswith("win")


@pytest.fixture
def app(pm, tmp_path, monkeypatch):
    monkeypatch.setattr(pm, "CONFIG", tmp_path / "config.json")
    monkeypatch.setattr(pm, "HAY_DND", False)
    monkeypatch.setattr(pm.App, "_preparar_dnd", lambda self: None)
    monkeypatch.setattr(pm.App, "ARRANQUE", False)
    monkeypatch.setattr(pm, "modo_efectos", lambda cfg: cfg.get("efectos", "completos")
                        if cfg.get("efectos") is not False else "apagados")
    try:
        a = pm.App()
    except tk.TclError as e:
        pytest.skip(f"sin pantalla: {e}")
    a.update()
    yield a
    a.destroy()


def test_usa_las_fuentes_incluidas(app):
    if not ES_WINDOWS and app.fam_titulo != "VT323":
        pytest.skip("este Tk de Linux no usa fontconfig (sin Xft)")
    assert (app.fam_titulo, app.fam) == ("VT323", "Share Tech Mono")


def test_agregar_unir_y_destino(app, tmp_path):
    carpeta = tmp_path / "Año 2026" / "Física"
    carpeta.mkdir(parents=True)
    for n in ("TP 10.docx", "TP 2.pptx", "~$TP 2.pptx", "notas.txt"):
        (carpeta / n).write_bytes(b"x")
    app.agregar_rutas([str(p) for p in carpeta.iterdir()])
    assert [a.name for a in app.archivos] == ["TP 2.pptx", "TP 10.docx"]
    assert app.v_estado.get() == "2 de 2 marcados"
    app.v_salida_modo.set("subcarpeta")
    app.v_fusionar.set(True)
    app._alternar_fusion()
    app.update()
    assert app.fr_fusion.winfo_ismapped()
    assert "PDF unido: Fusionado.pdf" in app.v_destino.get()
    app._alternar_log()
    app.update()
    assert app.tabla.winfo_height() >= 60  # la lista no desaparece


def test_ventanas_secundarias_abren(app, pm, tmp_path):
    app.abrir_herramientas()
    app.abrir_auto()
    app.update()
    assert app._win_herr.winfo_exists() and app._win_auto.winfo_exists()


def test_ruta_corta(pm):
    largo = "C:/Users/Alguien/Documents/Facultad/2026/Física II/Trabajos prácticos/PDF"
    corta = pm.App._corta(largo, 40)
    assert len(corta) <= 42 and corta.endswith("PDF") and "…" in corta
    assert pm.App._corta("C:/corta", 40) == "C:/corta"


# ------------------------------------------------------------- efectos (sátira)
def _esperar(app, condicion, segundos=5):
    import time
    fin = time.time() + segundos
    while time.time() < fin and not condicion():
        app.update()
        time.sleep(0.02)
    return condicion()


def test_efectos_aparecen_al_trabajar_y_se_van_solos(app, pm, monkeypatch):
    monkeypatch.setattr(pm.Teatro, "FINAL_MS", 200)

    def falso(emit):
        emit("max", 2)
        emit("prog", 1)
        emit("log", "OK  uno.docx")

    app._iniciar(falso, ())
    assert app.teatro is not None
    assert _esperar(app, lambda: "1 de 2 (50%)" in app.teatro.progreso_real())
    app.cola.put(("fin", "Listo."))
    assert _esperar(app, lambda: app.teatro is None)
    assert str(app.btn_conv.cget("state")) == "normal"


def test_con_error_abre_el_registro(app, pm, monkeypatch):
    monkeypatch.setattr(pm.Teatro, "FINAL_MS", 200)
    app._iniciar(lambda emit: (emit("log", "ERROR  x.docx: dañado"), emit("fin", "Listo.")), ())
    assert _esperar(app, lambda: app.log_visible)
    assert app.teatro is None or app.teatro.final[0] == "RESULTADO: CON OBSERVACIONES"


def test_un_clic_oculta_los_efectos(app):
    app._iniciar(lambda emit: None, ())
    app.update()
    app.teatro.c.event_generate("<Button-1>", x=5, y=5, when="now")
    app.update()
    assert app.teatro is None


def test_efectos_se_pueden_apagar(app):
    app.v_efectos.set("apagados")
    app._cambiar_efectos()
    app._iniciar(lambda emit: emit("fin", "Listo."), ())
    assert app.teatro is None


def test_lluvia_aguanta_ventanas_chicas(app, pm):
    app.geometry("860x620")
    app._alternar_log()
    app._iniciar(lambda emit: None, ())
    for _ in range(60):
        app.teatro._lluvia()
        app.teatro._texto()
    app.update()
    assert app.teatro.vivo
    assert set(pm.Teatro.CARACTERES) <= set(map(chr, range(32, 256)))  # todos están en VT323


def test_la_animacion_es_un_solo_bucle(app):
    """Llamarla varias veces no apila bucles: antes eso colgó la CI de Windows."""
    app._iniciar(lambda emit: None, ())
    pendientes_antes = len(app.tk.call("after", "info"))
    for _ in range(20):
        app.teatro._cuadro()
    assert len(app.tk.call("after", "info")) == pendientes_antes
    app.update()
    assert app.teatro.vivo


# ------------------------------------------------------------- estética 5.1
def test_modo_efectos(pm, monkeypatch):
    assert pm.modo_efectos({"efectos": False}) == "apagados"  # configuración de la 5.0
    assert pm.modo_efectos({"efectos": True}) == "completos"
    assert pm.modo_efectos({"efectos": "reducidos"}) == "reducidos"
    monkeypatch.setattr(pm, "animaciones_del_sistema", lambda: False)
    assert pm.modo_efectos({}) == "reducidos"  # respeta «Mostrar animaciones» de Windows


def test_guion_usa_los_datos_reales(pm):
    import random
    g = pm.Guion(random.Random(1))
    for clave, (llano, variantes) in pm.Guion.ACCIONES.items():
        datos = {"n": 37, "nombre": "Física II", "exp": "Nº 1234/2026", "seg": "4,2"}
        for v in variantes:
            texto = v.format(cierre="", **datos)
            for campo, valor in datos.items():
                if "{" + campo + "}" in v:
                    assert str(valor) in texto, (clave, v)
        assert g.accion(clave, **datos)[1]
    assert set("".join(pm.GLIFOS)) <= set(map(chr, range(32, 256)))


def test_modo_serio_no_muestra_satira(app, tmp_path):
    app.v_efectos.set("apagados")
    app._cambiar_efectos()
    for n in ("a.docx", "b.pdf"):
        (tmp_path / n).write_bytes(b"x")
    app.agregar_rutas([str(tmp_path / "a.docx"), str(tmp_path / "b.pdf")])
    assert app.v_sis.get() == "2 archivo(s) en la lista."
    app._reposo()
    assert app.v_sis.get() == "2 archivo(s) en la lista."


def test_cada_accion_deja_constancia(app, tmp_path):
    app.v_efectos.set("reducidos")  # sin animación de tipeo, para leerlo enseguida
    app._cambiar_efectos()
    (tmp_path / "x.docx").write_bytes(b"x")
    import re
    for _ in range(12):  # todas las variantes, que salen al azar, llevan el número real
        app.agregar_rutas([str(tmp_path / "x.docx")], reemplazar=True)
        assert re.search(r"\b1 (EXPEDIENTE|DOCUMENTO)", app.v_sis.get()), app.v_sis.get()
    app.quitar()
    assert "SE DESAFECTARON 1" in app.v_sis.get()


def test_arranque_se_saltea_con_un_clic(app, pm):
    app.arranque = pm.Arranque(app)
    app.update()
    textos = [l for l, _ in app.arranque.lineas]
    assert any(t.startswith("FUENTES DE FÓSFORO") for t in textos)
    assert any(t.startswith("LIBREOFFICE") for t in textos)
    app.arranque.c.event_generate("<Button-1>", x=5, y=5, when="now")
    app.update()
    assert app.arranque is None


def test_barra_segmentada_se_usa_como_progressbar(app):
    app.barra.config(maximum=4, value=1)
    assert (app.barra.cget("maximum"), app.barra.cget("value")) == (4, 1)
    assert app.barra.fraccion() == 0.25
    app.update()


def test_botones_se_decodifican_y_vuelven(app):
    original = app.btn_conv.cget("text")
    app._decodificar(app.btn_conv)
    assert _esperar(app, lambda: not getattr(app.btn_conv, "_decodificando", False), 2)
    assert app.btn_conv.cget("text") == original


def test_mensaje_oculto_y_reposo(app):
    for _ in range(5):
        app.enc._clic_titulo(None)
    assert _esperar(app, lambda: "CURIOSIDAD NO AUTORIZADA" in app.v_sis.get() and
                    not app.v_sis.get().endswith("_"), 3)
    app._reposo()
    app.enc.interferencia()
    _esperar(app, lambda: False, 0.5)
    assert app.enc.winfo_exists()
