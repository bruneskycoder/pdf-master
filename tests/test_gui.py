"""La interfaz arranca, usa las fuentes incluidas y responde a lo básico.
Se saltea si no hay pantalla (en Linux, correr con xvfb-run)."""
import sys

import pytest

tk = pytest.importorskip("tkinter")
ES_WINDOWS = sys.platform.startswith("win")


@pytest.fixture
def app(pm, tmp_path, monkeypatch):
    monkeypatch.setattr(pm, "CONFIG", tmp_path / "config.json")
    (tmp_path / "config.json").write_text('{"doctor": false}', encoding="utf-8")  # el Dr. Fósforo tiene sus pruebas aparte
    monkeypatch.setattr(pm, "HAY_DND", False)
    monkeypatch.setattr(pm.App, "_preparar_dnd", lambda self: None)
    monkeypatch.setattr(pm.App, "ARRANQUE", False)
    monkeypatch.setattr(pm, "modo_efectos", lambda cfg: cfg.get("efectos", "completos")
                        if cfg.get("efectos") is not False else "apagados")
    try:
        a = pm.App()
    except tk.TclError as e:
        if not ES_WINDOWS:
            pytest.skip(f"sin pantalla: {e}")
        # En Windows siempre hay pantalla: un fallo acá no se saltea en silencio. Los runners
        # de GitHub a veces no pueden leer los archivos de arranque de Tcl; se reintenta una
        # sola vez y, si vuelve a fallar, la prueba falla.
        import time
        time.sleep(1)
        try:
            a = pm.App()
        except tk.TclError as e2:
            pytest.fail(f"Tk no arrancó en Windows dos veces seguidas: {e} / {e2}")
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

    def bucles():  # tareas pendientes que son un cuadro de la animación
        return sum("_cuadro" in str(app.tk.call("after", "info", t)[0])
                   for t in app.tk.call("after", "info"))
    assert bucles() == 1
    for _ in range(20):
        app.teatro._cuadro()
    assert bucles() == 1
    app.update()
    assert app.teatro.vivo


# ------------------------------------------------------------- estética 5.1
def test_modo_efectos(pm, monkeypatch):
    assert pm.modo_efectos({"efectos": False}) == "apagados"  # configuración de la 5.0
    assert pm.modo_efectos({"efectos": True}) == "completos"
    assert pm.modo_efectos({"efectos": "reducidos"}) == "reducidos"
    assert pm.modo_efectos({}) == "completos"  # siempre completos, salvo que se elija otro


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


def test_arranque_clic_adelanta_hasta_el_pin(app, pm):
    app.arranque = pm.Arranque(app)
    app.update()
    textos = [l for l, _ in app.arranque.lineas]
    assert any(t.startswith("FUENTES DE FÓSFORO") for t in textos)
    assert any(t.startswith("LIBREOFFICE") for t in textos)
    app.arranque.c.event_generate("<Button-1>", x=5, y=5, when="now")
    app.update()
    assert app.arranque.fase == "pin"


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


# ------------------------------------------------------------- 5.2: PIN, ataques, CRT
def test_pin_deja_pasar_cualquier_respuesta_y_no_guarda_nada(app, pm):
    for valor, esperado in (("1", "PIN INCORRECTO."), ("", "PIN VACÍO."), (None, "PIN CORRECTO.")):
        app.arranque = a = pm.Arranque(app)
        a._saltar()
        assert a.fase == "pin"
        a.entrada.insert(0, a.pin if valor is None else valor)
        a._validar()
        app.update()
        textos = [a.c.itemcget(i, "text") for i in a.c.find_all() if a.c.type(i) == "text"]
        assert esperado in textos and "ACCESO CONCEDIDO" in textos
        assert a.entrada.get() == ""  # el valor tipeado no queda en ningún lado
        assert any("HACKEADO DESDE 1994" in t and a.pin in t for t in textos)
        a.cerrar()
        assert app.arranque is None


def test_escape_saltea_el_arranque(app, pm):
    app.arranque = pm.Arranque(app)
    app._ocultar_efectos()
    assert app.arranque is None


def test_mensaje_nuevo_corta_el_tipeo_anterior(app):
    app.decir("MENSAJE LARGO " * 8)
    app.decir("CORTO", alarma=True)
    _esperar(app, lambda: False, 0.4)
    assert app.v_sis.get() == "CORTO"
    assert str(app.lbl_sis.cget("style")) == "SisR.TLabel"


def test_ataque_falso_se_resuelve_solo_y_no_ensucia_el_registro(app, pm, monkeypatch):
    app.v_efectos.set("reducidos")  # sin tipeo, para leer enseguida
    app._cambiar_efectos()
    registro = app.log.get("1.0", "end")
    app.ataque()
    alarma = app.v_sis.get()
    assert str(app.lbl_sis.cget("style")) == "SisR.TLabel" and alarma
    assert _esperar(app, lambda: app.v_sis.get() != alarma, 4)
    assert str(app.lbl_sis.cget("style")) == "Sis.TLabel"
    assert app.log.get("1.0", "end") == registro
    pares = dict(pm.Guion.ATAQUES)
    assert pares[alarma] == app.v_sis.get()


def test_falla_roja_en_la_pantalla_de_trabajo(app, pm):
    app._iniciar(lambda emit: None, ())
    t = app.teatro
    t.azar.random = lambda: 0.0
    t.t_sub = 0
    t._texto()
    assert "!!" in t.subs[-1]
    item = t.t_subs[len(t.subs) - 1]
    assert t.c.itemcget(item, "fill") == pm.ROJO


def test_capa_crt_en_windows(app, pm):
    if not pm.ES_WINDOWS:
        assert app.crt is None
        pytest.skip("la capa CRT es solo para Windows")
    import ctypes
    assert app.crt is not None
    app.update()
    app.crt._seguir()
    app.update()
    t = app.crt.t
    assert (t.winfo_width(), t.winfo_height()) == (app.winfo_width(), app.winfo_height())
    hwnd = app.crt.hwnd()
    estilo = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
    assert estilo & 0x20 and estilo & 0x80000, hex(estilo)  # no recibe clics y es una capa
    # comportamiento real: la ventana que Windows ve bajo el mouse no puede ser la capa
    from ctypes import wintypes
    centro = wintypes.POINT(app.winfo_rootx() + app.winfo_width() // 2,
                            app.winfo_rooty() + app.winfo_height() // 2)
    ctypes.windll.user32.WindowFromPoint.argtypes = [wintypes.POINT]
    ctypes.windll.user32.WindowFromPoint.restype = wintypes.HWND
    esquina = wintypes.POINT(app.winfo_rootx() + 12, app.winfo_rooty() + 12)
    for punto in (centro, esquina):  # en la esquina está la capa casi opaca del marco
        debajo = ctypes.windll.user32.WindowFromPoint(punto)
        raiz = ctypes.windll.user32.GetAncestor(debajo, 2)  # GA_ROOT
        assert raiz not in app.crt.hwnds(), "la capa CRT se está comiendo los clics"
    for h in app.crt.hwnds():
        assert ctypes.windll.user32.GetWindowLongW(h, -20) & 0x20
    # minimizar y volver a mostrar no le saca la transparencia a los clics
    app.withdraw(); app.update(); app.deiconify(); app.update()
    estilo = ctypes.windll.user32.GetWindowLongW(app.crt.hwnd(), -20)
    assert estilo & 0x20, hex(estilo)
    app.crt.destello()
    app.crt.parpadeo()
    app.crt.desgarro()
    app.crt.sincronia()
    app.crt.apagon()
    _esperar(app, lambda: False, 0.5)
    app.v_efectos.set("apagados")
    app._cambiar_efectos()
    assert app.crt is None


# ------------------------------------------------------------- 5.3: simulador de hacker
def _evento(tecla, caracter=""):
    import types
    return types.SimpleNamespace(keysym=tecla, char=caracter)


def _hackeo(app, pm, tmp_path, llamadas):
    archivos = []
    for n in ("TP 1 - Introducción.docx", "Clase 'Final' ñ.pptx"):  # sin comillas dobles: Windows no las permite
        (tmp_path / n).write_bytes(b"x")
        archivos.append(tmp_path / n)
    app.antes_de_convertir(archivos, lambda: llamadas.append(1))
    return app.hackeo


def test_hackeo_cualquier_tecla_escribe_codigo(app, pm, tmp_path):
    llamadas = []
    h = _hackeo(app, pm, tmp_path, llamadas)
    assert h is not None
    antes = h.txt.get("1.0", "end")
    for _ in range(400):
        h._tecla(_evento("a", "a"))
    texto = h.txt.get("1.0", "end")
    assert len(texto) > len(antes) + 1500
    assert "TP 1 - Introducción.docx" in texto  # los nombres reales aparecen en el código
    assert h.nivel() == 100 and llamadas == [1]
    assert h.txt.tag_ranges("kw") and h.txt.tag_ranges("str")  # hay colores


def test_hackeo_la_conversion_real_arranca_al_abrir(app, pm, tmp_path):
    llamadas = []
    h = _hackeo(app, pm, tmp_path, llamadas)
    assert llamadas == [1] and app.hackeo is h  # arrancó una sola vez, con la página abierta
    h._tecla(_evento("x", "x"))
    _esperar(app, lambda: False, 0.5)
    assert llamadas == [1]


def test_hackeo_enter_antes_de_terminar_no_cierra_ni_reconvierte(app, pm, tmp_path):
    llamadas = []
    h = _hackeo(app, pm, tmp_path, llamadas)
    h._tecla(_evento("Return"))
    h._tecla(_evento("Return"))
    _esperar(app, lambda: False, 0.6)
    assert app.hackeo is h and h.vivo and llamadas == [1]
    assert "TODAVÍA NO TERMINÓ" in h.txt.get("1.0", "end")


def test_hackeo_fin_exitoso_espera_enter(app, pm, tmp_path):
    h = _hackeo(app, pm, tmp_path, [])
    h.terminar("3,4", False)
    texto = h.txt.get("1.0", "end")
    assert ">>> CONVERSIÓN REAL TERMINADA EN 3,4 S. ENTER PARA SALIR." in texto
    assert "TERMINÓ. ENTER PARA SALIR" in h.v_pie.get()
    _esperar(app, lambda: False, 0.5)
    assert app.hackeo is h  # no se cierra sola
    h._tecla(_evento("Return"))
    assert app.hackeo is None


def test_hackeo_fin_con_errores_en_rojo(app, pm, tmp_path):
    h = _hackeo(app, pm, tmp_path, [])
    h.terminar("1,0", True)
    texto = h.txt.get("1.0", "end")
    assert ">>> LA CONVERSIÓN TERMINÓ CON ERRORES: TE ABRÍ EL REGISTRO. ENTER PARA SALIR." in texto
    rojas = "".join(h.txt.get(a, b) for a, b in zip(*[iter(h.txt.tag_ranges("err"))] * 2))
    assert "CON ERRORES" in rojas
    h._tecla(_evento("KP_Enter"))
    assert app.hackeo is None


def test_hackeo_esc_sale_sin_cancelar_y_crea_el_teatro(app, pm, tmp_path):
    llamadas = []
    h = _hackeo(app, pm, tmp_path, llamadas)
    assert app.teatro is None
    app.trabajando = True  # la conversión sigue en marcha
    h._tecla(_evento("Escape"))
    assert app.hackeo is None and llamadas == [1]
    assert app.trabajando and app.teatro is not None
    app.teatro.cerrar()
    app.trabajando = False


def test_hackeo_salir_sin_conversion_en_marcha_no_crea_teatro(app, pm, tmp_path):
    h = _hackeo(app, pm, tmp_path, [])
    h.btn_salir.invoke()
    assert app.hackeo is None and app.teatro is None


def test_hackeo_progreso_real_refleja_la_barra(app, pm, tmp_path):
    h = _hackeo(app, pm, tmp_path, [])
    app.barra.config(maximum=5, value=2)
    app.v_status.set("OK  TP 2 - Señales.docx")
    assert _esperar(app, lambda: "2 de 5" in h.v_real.get(), 3)
    assert "TP 2 - Señales.docx" in h.v_real.get()
    assert (h.barra.cget("maximum"), h.barra.cget("value")) == (5, 2)
    app.barra.config(value=4)
    assert _esperar(app, lambda: "4 de 5" in h.v_real.get(), 3)


def test_hackeo_el_teatro_no_se_crea_con_la_pagina_abierta(app, pm, tmp_path):
    h = _hackeo(app, pm, tmp_path, [])
    app._iniciar(lambda emit: emit("fin", "Listo."), ())
    assert app.teatro is None and app.hackeo is h
    assert _esperar(app, lambda: not app.trabajando, 3)
    assert _esperar(app, lambda: h.terminado, 3)
    assert app.hackeo is h  # el fin no cierra la página


def test_hackeo_denegado_y_fallas_no_rompen(app, pm, tmp_path):
    h = _hackeo(app, pm, tmp_path, [])
    h._denegado()
    h._falla()
    for _ in range(120):  # la línea roja sale al terminar la línea de código en curso
        h._tecla(_evento("a", "a"))
    assert _esperar(app, lambda: not h.cartel.winfo_exists(), 3)
    rojas = [h.txt.get(a, b) for a, b in zip(*[iter(h.txt.tag_ranges("err"))] * 2)]
    assert any(r.strip() for r in rojas)
    assert all(h.txt.index(a).endswith(".0") for a in h.txt.tag_ranges("err")[::2])
    assert "Segmentation" in h.txt.get("1.0", "end") or "KERNEL" in h.txt.get("1.0", "end") \
        or "PARIDAD" in h.txt.get("1.0", "end") or "SEÑAL" in h.txt.get("1.0", "end")


def test_sin_efectos_completos_no_hay_simulador(app, pm, tmp_path):
    for modo in ("reducidos", "apagados"):
        app.v_efectos.set(modo)
        app._cambiar_efectos()
        llamadas = []
        assert _hackeo(app, pm, tmp_path, llamadas) is None
        assert llamadas == [1]


def test_convertir_pasa_por_el_simulador(app, pm, tmp_path, monkeypatch):
    llamadas = []
    monkeypatch.setattr(app, "_iniciar", lambda funcion, args: llamadas.append(args))
    for n in ("a.pdf", "b.pdf"):
        (tmp_path / n).write_bytes(b"%PDF-1.4")
    app.agregar_rutas([str(tmp_path / "a.pdf"), str(tmp_path / "b.pdf")])
    app.v_fusionar.set(True)
    app.convertir()
    assert app.hackeo is not None and len(llamadas) == 1  # la conversión real ya arrancó
    app.hackeo._tecla(_evento("Return"))  # Enter antes de terminar no reconvierte
    assert len(llamadas) == 1


def test_falla_de_pantalla_se_resuelve(app):
    app.v_efectos.set("reducidos")
    app._cambiar_efectos()
    app.falla_pantalla()
    alarma = app.v_sis.get()
    assert alarma and str(app.lbl_sis.cget("style")) == "SisR.TLabel"
    assert _esperar(app, lambda: app.v_sis.get() != alarma, 4)
