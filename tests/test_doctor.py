"""El Dr. Fósforo: personaje con guion local. No miente, no molesta y se puede silenciar."""
import re
from datetime import datetime

import pytest

from test_gui import app, _esperar  # noqa: F401  (el fixture `app` se reutiliza)

PERMITIDOS = set(map(chr, range(32, 256))) | set("«»—…“”‘’•·")


class Reloj:
    def __init__(self, t=1_000_000.0):
        self.t = t

    def __call__(self):
        return self.t


@pytest.fixture
def dr(app, pm, monkeypatch):
    """El doctor encendido, con reloj propio y sin demoras."""
    monkeypatch.setattr(pm.DrFosforo, "RETRASO_MS", 0)
    app.v_doctor.set(True)
    app.efectos = "reducidos"  # sin tipeo: el texto se lee enseguida
    if app._sis_tarea:  # el tipeo del «listo» inicial no debe pisar lo que se mide
        app.after_cancel(app._sis_tarea)
        app._sis_tarea = None
    app.v_sis.set("")
    d = app.doctor
    d.reloj = Reloj()
    d._t = d.reloj
    app.ultima_accion = d.reloj()
    d.saludado = True
    return d


def _frases(pm):
    g = pm.GuionDoctor
    for nombre in vars(g):  # solo lo del doctor, no lo heredado de Guion
        if nombre.isupper() and nombre != "REQUIERE":
            for item in getattr(g, nombre):
                yield nombre, item


def test_saludo_segun_la_franja_horaria(dr, app, pm, monkeypatch):
    for hora, franja in ((0, "MADRUGADA"), (5, "MADRUGADA"), (6, "MANANA"), (11, "MANANA"),
                         (12, "TARDE"), (19, "TARDE"), (20, "NOCHE"), (23, "NOCHE")):
        dr._hora = lambda h=hora: datetime(2026, 11, 25, h, 30)  # un miércoles
        dr.saludado = False
        dr.ultima_voz = -1e9
        monkeypatch.setattr(app.azar, "random", lambda: 0.9)  # no el saludo del día
        dr.saludar()
        assert app.v_sis.get() in getattr(pm.GuionDoctor, "SALUDO_" + franja), (hora, app.v_sis.get())
        assert app.lbl_pref.cget("text") == "DR.FÓSFORO>"


def test_saludo_del_dia_de_la_semana(dr, app, pm, monkeypatch):
    dr._hora = lambda: datetime(2026, 11, 27, 15, 0)  # viernes
    dr.saludado = False
    monkeypatch.setattr(app.azar, "random", lambda: 0.1)
    dr.saludar()
    assert app.v_sis.get() == pm.GuionDoctor.SALUDO_DIA[4]
    assert app.v_sis.get().startswith("¡Viernes!")


def test_saludo_una_sola_vez_y_pin_incorrecto(dr, app, pm):
    dr.saludado = False
    dr.pin_mal = True
    dr.saludar()
    assert app.v_sis.get() in pm.GuionDoctor.PIN_MAL
    app.v_sis.set("otra cosa")
    dr.saludar()
    assert app.v_sis.get() == "otra cosa"


def test_el_doctor_habla_con_su_estilo(dr, app):
    dr.reaccionar("fin", seg="3,2")
    assert str(app.lbl_sis.cget("style")) == "SisDr.TLabel"
    assert app.lbl_pref.cget("text") == "DR.FÓSFORO>"
    app.decir("algo burocrático")
    assert app.lbl_pref.cget("text") == "SIS>"
    assert str(app.lbl_sis.cget("style")) == "Sis.TLabel"


def test_reaccionar_incluye_los_datos_reales(dr, app):
    dr.reloj.t += 100
    assert dr.reaccionar("fin", seg="12,5")
    assert "12,5" in app.v_sis.get()
    dr.reloj.t += 100
    for _ in range(8):  # las acciones comunes hablan a veces: se fuerza la suerte
        app.azar.random = lambda: 0.0
        dr.reloj.t += 100
        assert dr.reaccionar("agregar", n=7)
        assert "7" in app.v_sis.get()
    dr.reloj.t += 100
    assert dr.reaccionar("markdown", n=321, nombre="notas.md")
    assert "321" in app.v_sis.get() and "notas.md" in app.v_sis.get()
    dr.reloj.t += 100
    assert dr.reaccionar("escaner", n=4)
    assert "4" in app.v_sis.get()


def test_el_error_se_dice_sin_disfrazarlo(dr, app, pm):
    assert dr.reaccionar("error")
    assert app.v_sis.get() in pm.GuionDoctor.ERROR
    assert re.search(r"error|falló|errores|No salió", app.v_sis.get(), re.I)


def test_fin_y_error_siempre_responden(dr, app, monkeypatch):
    monkeypatch.setattr(app.azar, "random", lambda: 0.99)  # la suerte más mala
    assert dr.reaccionar("fin", seg="1,0")
    assert dr.reaccionar("error")
    assert not dr.reaccionar("agregar", n=2)  # una acción común, con mala suerte, calla


def test_con_hackeo_visto_el_fin_lo_comenta(dr, app, monkeypatch):
    dr.reaccionar("hackeo")
    monkeypatch.setattr(app.azar, "random", lambda: 0.0)
    dr.reaccionar("fin", seg="2,0")
    assert "2,0" in app.v_sis.get() and ("intrusión" in app.v_sis.get())


def test_no_habla_en_apagados_ni_con_la_casilla_apagada(dr, app):
    app.efectos = "apagados"
    assert not dr.reaccionar("fin", seg="1,0")
    assert not dr.reaccionar("error")
    dr.saludado = False
    dr.saludar()
    assert dr.globo(["a", "b"]) is None
    app.efectos = "reducidos"
    app.v_doctor.set(False)
    assert not dr.reaccionar("fin", seg="1,0")
    assert dr.globo(["a", "b"]) is None
    assert dr.monologo() is None


def test_la_casilla_del_menu_guarda_en_la_config(dr, app, pm):
    app.v_doctor.set(False)
    app._cambiar_doctor()
    assert app.cfg["doctor"] is False
    assert pm.cargar_config()["doctor"] is False
    app.v_doctor.set(True)
    app._cambiar_doctor()
    assert app.cfg["doctor"] is True
    assert pm.cargar_config()["doctor"] is True


def test_modo_apagado_desde_el_menu_lo_calla(app, pm):
    app.v_doctor.set(True)
    app.v_efectos.set("apagados")
    app._cambiar_efectos()
    assert app.doctor.activo() is False


def test_globo_aparece_no_tapa_el_boton_y_se_va_con_clic(dr, app):
    app.efectos = "completos"
    app.update()
    g = dr.globo(["Una línea de prueba.", "Otra línea."])
    assert g is not None and dr.globo_actual is g
    app.update()
    assert g.winfo_ismapped()
    gx, gy, gw, gh = g.winfo_rootx(), g.winfo_rooty(), g.winfo_width(), g.winfo_height()
    b = app.btn_conv
    bx, by, bw, bh = b.winfo_rootx(), b.winfo_rooty(), b.winfo_width(), b.winfo_height()
    assert not (gx < bx + bw and bx < gx + gw and gy < by + bh and by < gy + gh)
    for w in (app.btn_agregar, app.btn_abrir):
        if w.winfo_ismapped():
            x, y, a, h = w.winfo_rootx(), w.winfo_rooty(), w.winfo_width(), w.winfo_height()
            assert not (gx < x + a and x < gx + gw and gy < y + h and y < gy + gh)
    g.event_generate("<Button-1>", x=5, y=5, when="now")
    app.update()
    assert dr.globo_actual is None


def test_globo_se_va_solo(dr, app, pm, monkeypatch):
    monkeypatch.setattr(pm.DrFosforo, "GLOBO_MS", 150)
    g = dr.globo(["Hola.", "Chau."])
    assert g is not None
    assert _esperar(app, lambda: dr.globo_actual is None, 3)


def test_globo_en_reducidos_aparece_completo_y_sin_animacion(dr, app):
    g = dr.globo(["Una línea.", "Otra línea."])
    assert g.animado is False
    assert "Una línea.\nOtra línea." == g.itemcget(g.texto, "text")


def test_globo_respeta_el_intervalo_de_90_s(dr, app):
    assert dr.globo(["a.", "b."]) is not None
    dr.globo_actual.cerrar()
    dr.reloj.t += 89
    assert dr.globo(["a.", "b."]) is None
    dr.reloj.t += 2
    assert dr.globo(["a.", "b."]) is not None


def test_respeta_el_intervalo_entre_frases_salvo_urgentes(dr, app, monkeypatch):
    monkeypatch.setattr(app.azar, "random", lambda: 0.0)
    assert dr.reaccionar("agregar", n=1)
    primera = app.v_sis.get()
    dr.reloj.t += 5
    dr.reaccionar("agregar", n=2)
    assert app.v_sis.get() == primera  # 5 s después: callado
    dr.reaccionar("error")  # lo urgente pasa igual
    assert app.v_sis.get() != primera
    dr.reloj.t += 21
    dr.reaccionar("agregar", n=3)
    assert "3" in app.v_sis.get()


def test_no_habla_mientras_se_trabaja(dr, app):
    app.trabajando = True
    app.v_sis.set("trabajando")
    dr.reaccionar("fin", seg="1,0")
    assert app.v_sis.get() == "trabajando"
    assert dr.globo(["a.", "b."]) is None
    dr.latido()
    assert app.v_sis.get() == "trabajando"
    app.trabajando = False


@pytest.mark.parametrize("overlay", ["arranque", "hackeo", "teatro"])
def test_no_habla_con_overlays_abiertos(dr, app, overlay):
    setattr(app, overlay, object())
    app.v_sis.set("overlay")
    dr.reaccionar("fin", seg="1,0")
    assert app.v_sis.get() == "overlay"
    assert dr.globo(["a.", "b."]) is None
    setattr(app, overlay, None)


def test_no_habla_con_un_mensaje_real_abierto(dr, app, pm):
    pm._MODALES[0] = 1
    try:
        app.v_sis.set("mensaje")
        dr.reaccionar("error")
        assert app.v_sis.get() == "mensaje"
        assert dr.globo(["a.", "b."]) is None
    finally:
        pm._MODALES[0] = 0


def test_los_mensajes_reales_cuentan_como_abiertos(pm):
    visto = []
    # la envoltura cuenta al entrar y descuenta al salir: se prueba con una función propia
    falso = type("M", (), {"f": lambda *a, **k: visto.append(pm._MODALES[0])})
    pm._vigilar_modales(falso, ("f",))
    falso.f()
    assert visto == [1] and pm._MODALES[0] == 0
    for f in (pm.messagebox.showinfo, pm.messagebox.askyesno, pm.filedialog.askopenfilenames):
        assert f.__name__ == "envoltura"  # los mensajes y diálogos reales están vigilados


def test_el_globo_se_retira_si_arranca_un_trabajo(dr, app):
    g = dr.globo(["a.", "b."])
    assert g is not None
    app._iniciar(lambda emit: None, ())
    app.update()
    assert dr.globo_actual is None


def test_reposo_2_5_y_15_minutos(dr, app, pm, monkeypatch):
    monkeypatch.setattr(app.azar, "random", lambda: 0.99)
    t0 = dr.reloj.t
    app.ultima_accion = t0
    dr.reloj.t = t0 + 60
    dr.latido()
    assert not app.v_sis.get()
    dr.reloj.t = t0 + 130
    dr.latido()
    assert "2" in app.v_sis.get() and app.v_sis.get() in [f.format(min=2) for f in pm.GuionDoctor.IDLE_2]
    dr.reloj.t = t0 + 400
    dr.latido()
    assert app.v_sis.get() in [f.format(min=6) for f in pm.GuionDoctor.IDLE_5]
    dr.reloj.t = t0 + 420
    antes = app.v_sis.get()
    dr.latido()
    assert app.v_sis.get() == antes  # cada hito, una sola vez
    dr.reloj.t = t0 + 15 * 60 + 5
    dr.latido()
    g = dr.globo_actual
    assert g is not None and "15" in g.itemcget(g.texto, "text")
    g.cerrar()
    app.ultima_accion = dr.reloj.t  # el operador vuelve: se reinicia la cuenta
    dr.latido()  # el latido siguiente ve que hubo actividad y reinicia los hitos
    app.v_sis.set("")
    dr.reloj.t += 130
    dr.latido()
    assert app.v_sis.get() in [f.format(min=2) for f in pm.GuionDoctor.IDLE_2]


def test_el_toque_real_actualiza_la_ultima_interaccion(app):
    app.ultima_accion = 0
    app.event_generate("<Key>", keysym="a", when="now")
    app.update()
    assert app.ultima_accion > 0


def test_decir_accion_dispara_las_reacciones_con_datos_reales(dr, app, monkeypatch):
    monkeypatch.setattr(app.azar, "random", lambda: 0.0)
    app.decir_accion("fin", exp="Nº 1", seg="8,1")
    assert "8,1" in app.v_sis.get()


def test_el_guion_solo_usa_glifos_de_las_fuentes(pm):
    for nombre, item in _frases(pm):
        for texto in (item if isinstance(item, tuple) else (item,)):
            malos = set(texto) - PERMITIDOS
            assert not malos, (nombre, texto, malos)


def test_las_frases_con_marcadores_los_usan(pm):
    g = pm.GuionDoctor
    marcadores = re.compile(r"\{\w+\}")
    for nombre, item in _frases(pm):
        texto = " ".join(item) if isinstance(item, tuple) else item
        pedidos = g.REQUIERE.get(nombre, ())
        for m in pedidos:
            assert m in texto, (nombre, texto, m)
        if not pedidos:
            assert not marcadores.findall(texto), (nombre, texto)


def test_formato_de_las_frases_y_los_globos(pm):
    g = pm.GuionDoctor
    assert len(g.EXPERIMENTOS) >= 10 and len(g.BUROCRACIA) >= 8
    total = sum(len(getattr(g, n)) for n in vars(g) if n.isupper() and n != "REQUIERE")
    assert total >= 80
    for nombre, item in _frases(pm):
        if isinstance(item, tuple):  # globo: 2 a 4 líneas cortas
            assert 2 <= len(item) <= 4, (nombre, item)
            assert all(len(t) <= 36 for t in item), (nombre, item)
        else:  # línea SIS>: cabe en una línea
            assert len(item) <= 96, (nombre, len(item), item)
    assert "EXPERIMENTO 47: intentar convertir un PDF en PDF. Resultado: un PDF. Fascinante." \
        in g.EXPERIMENTOS


def test_el_guion_no_repite_hasta_agotar(pm):
    import random
    g = pm.GuionDoctor(random.Random(1))
    vistas = [g.sacar("EXPERIMENTOS") for _ in range(len(g.EXPERIMENTOS))]
    assert len(set(vistas)) == len(g.EXPERIMENTOS)
