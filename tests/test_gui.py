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
    assert app.teatro is None or app.teatro.final[0] == "COMPLETADO CON OBSERVACIONES"


def test_un_clic_oculta_los_efectos(app):
    app._iniciar(lambda emit: None, ())
    app.update()
    app.teatro.c.event_generate("<Button-1>", x=5, y=5, when="now")
    app.update()
    assert app.teatro is None


def test_efectos_se_pueden_apagar(app):
    app.cfg["efectos"] = False
    app._iniciar(lambda emit: emit("fin", "Listo."), ())
    assert app.teatro is None


def test_lluvia_aguanta_ventanas_chicas(app, pm):
    app.geometry("860x620")
    app._alternar_log()
    app._iniciar(lambda emit: None, ())
    for _ in range(60):
        app.teatro._cuadro()
    app.update()
    assert app.teatro.vivo
    assert set(pm.Teatro.CARACTERES) <= set(map(chr, range(32, 256)))  # todos están en VT323
