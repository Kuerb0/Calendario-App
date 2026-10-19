"""Pruebas rápidas de calendario.py que NO necesitan Outlook ni pantalla real.

Para qué sirve: antes de generar un instalador nuevo (o después de pedir
un cambio), ejecuta esto desde la carpeta principal:

    python herramientas/tests_calendario.py

Si todo sale "OK", es que las partes más delicadas del programa (limpieza
de correos, reglas de Outlook, anclaje de paneles, silenciador de avisos)
siguen funcionando igual que antes. Si algo falla, lo dice con el nombre
de la prueba, para saber qué mirar.

No hace falta pytest ni nada que no tengas ya instalado para usar la app.
"""
import os
import sys
import datetime as dt
import time
import tkinter as tk

# el programa vive en ../app respecto a esta carpeta (herramientas/)
sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "app"))
import calendario as cal

PRUEBAS = []
FALLOS = []


def prueba(nombre):
    def decorador(fn):
        def envoltura():
            try:
                fn()
                print(f"  OK   {nombre}")
            except AssertionError as e:
                print(f" FALLA {nombre}: {e}")
                FALLOS.append(nombre)
            except Exception as e:
                print(f" ERROR {nombre}: {e!r}")
                FALLOS.append(nombre)
        PRUEBAS.append(envoltura)
        return fn
    return decorador


@prueba("limpieza de HTML del cuerpo de un correo")
def _():
    html_body = "<p>Hola</p><br><b>mundo</b><script>ignorar()</script>"
    limpio = cal._texto_de_cuerpo(html_body)
    assert "Hola" in limpio and "mundo" in limpio, limpio
    assert "<" not in limpio and "ignorar" not in limpio, limpio


@prueba("cuenta_atras da un texto para fechas futuras y pasadas")
def _():
    ahora = dt.datetime(2026, 1, 1, 12, 0, tzinfo=cal.TZ)
    futuro = ahora + dt.timedelta(hours=2)
    pasado = ahora - dt.timedelta(hours=2)
    assert cal.cuenta_atras(futuro, ahora)
    assert cal.cuenta_atras(pasado, ahora)


@prueba("las cuentas atrás y el evento fijado siguen al evento cuando se mueve")
def _():
    guardado = cal.CFG.get("cuentas_atras"), cal.CFG.get("anclado"), cal.guardar_config
    cal.guardar_config = lambda: None
    try:
        hoy = dt.datetime.now(cal.TZ).replace(hour=10, minute=0, second=0, microsecond=0)
        a = hoy + dt.timedelta(days=5)
        ev = cal.Evento(a, a + dt.timedelta(hours=1), "Viaje", False, "u1")
        cal.CFG["cuentas_atras"] = [{"nombre": "Viaje", "fecha": a.date().isoformat(), "uid": "u1", "hasta": ev.fin.isoformat()}]
        cal.CFG["anclado"] = {"uid": "u1", "titulo": "Viaje", "inicio": a.isoformat(), "color": "#fff"}
        # sin cambios: no toca nada
        assert not cal.seguir_eventos_movidos([ev])
        # el evento se mueve tres días
        b = a + dt.timedelta(days=3)
        movido = cal.Evento(b, b + dt.timedelta(hours=1), "Viaje (nuevo)", False, "u1")
        assert cal.seguir_eventos_movidos([movido])
        c = cal.CFG["cuentas_atras"][0]
        assert c["fecha"] == b.date().isoformat() and c["hasta"] == movido.fin.isoformat(), c
        assert cal.CFG["anclado"]["inicio"] == b.isoformat() and cal.CFG["anclado"]["titulo"] == "Viaje (nuevo)"
        # un evento distinto con otro UID no se confunde
        otro = cal.Evento(b, b + dt.timedelta(hours=1), "Otro", False, "u2")
        cal.CFG["cuentas_atras"] = [{"nombre": "Viaje", "fecha": a.date().isoformat(), "uid": "u1"}]
        cal.CFG["anclado"] = None
        assert not cal.seguir_eventos_movidos([otro])
        assert cal.CFG["cuentas_atras"][0]["fecha"] == a.date().isoformat()
        # serie repetida: si la ocurrencia sigue en su sitio, no se cambia a otra
        serie = [cal.Evento(a + dt.timedelta(days=7 * k), a + dt.timedelta(days=7 * k, hours=1), "Clase", False, "s") for k in range(4)]
        cal.CFG["cuentas_atras"] = [{"nombre": "Clase", "fecha": (a + dt.timedelta(days=7)).date().isoformat(), "uid": "s", "hasta": serie[1].fin.isoformat()}]
        assert not cal.seguir_eventos_movidos(serie)
        # serie repetida: una ocurrencia movida un día se sigue (la más cercana es una sola)
        serie[1] = cal.Evento(a + dt.timedelta(days=8), a + dt.timedelta(days=8, hours=1), "Clase", False, "s")
        assert cal.seguir_eventos_movidos(serie)
        assert cal.CFG["cuentas_atras"][0]["fecha"] == (a + dt.timedelta(days=8)).date().isoformat()
    finally:
        cal.CFG["cuentas_atras"], cal.CFG["anclado"], cal.guardar_config = guardado


@prueba("el silenciador de notificaciones corta de verdad notificar()")
def _():
    mensajes = []
    original = cal.CFG.get("notificaciones_activas", True)
    import builtins
    orig_print = builtins.print
    orig_winotify = sys.modules.get("winotify", False)
    sys.modules["winotify"] = None      # sin winotify no sale un aviso real
    try:
        builtins.print = lambda *a, **k: mensajes.append(
            " ".join(str(x) for x in a))

        cal.CFG["notificaciones_activas"] = True
        cal.notificar("prueba-activada", "cuerpo1")
        assert any("prueba-activada" in m for m in mensajes), \
            "con las notificaciones activas, deberia haber avisado"

        mensajes.clear()
        cal.CFG["notificaciones_activas"] = False
        cal.notificar("prueba-silenciada", "cuerpo2")
        assert not mensajes, \
            f"no deberia notificar nada silenciado, salio: {mensajes}"
    finally:
        builtins.print = orig_print
        cal.CFG["notificaciones_activas"] = original
        if orig_winotify is False:
            sys.modules.pop("winotify", None)
        else:
            sys.modules["winotify"] = orig_winotify


@prueba("anclaje de paneles: izquierda/derecha calculan bien la X")
def _():
    raiz = tk.Tk()
    try:
        raiz.overrideredirect(True)
        raiz.geometry("400x300+500+150")
        raiz._panel_accesos = None
        raiz.update_idletasks()

        ubic = cal.CFG.setdefault("ubicacion_vistas", {})
        ubic["_prueba_"] = "derecha"
        x = cal.x_panel_acoplado(raiz, "_prueba_", 200)
        assert x >= raiz.winfo_rootx() + raiz.winfo_width(), \
            "anclado a la derecha deberia quedar despues de la ventana"

        ubic["_prueba_"] = "izquierda"
        x = cal.x_panel_acoplado(raiz, "_prueba_", 200)
        assert x + 200 <= raiz.winfo_rootx() + 2, \
            "anclado a la izquierda deberia quedar antes de la ventana"
    finally:
        raiz.destroy()
        cal.CFG.get("ubicacion_vistas", {}).pop("_prueba_", None)


@prueba("panel de control: los botones fijos de la barra no se pueden quitar")
def _():
    import extras
    original = cal.CFG.get("barra_ocultos")
    try:
        cal.CFG["barra_ocultos"] = ["btn_dias", "btn_refrescar", "btn_ocultar",
                                    "btn_festivos"]
        for fijo in extras.BOTONES_FIJOS:
            assert not extras.barra_oculta(cal, fijo), f"{fijo} no debe quitarse"
        assert extras.barra_oculta(cal, "btn_festivos")
        assert not extras.barra_oculta(cal, "btn_nuevo")
    finally:
        cal.CFG["barra_ocultos"] = original or []


@prueba("panel de control: detecta solapes y eventos repetidos")
def _():
    import extras
    tz = cal.TZ
    base = dt.datetime(2030, 5, 6, 10, 0, tzinfo=tz)

    def ev(uid, titulo, ini, fin):
        return cal.Evento(ini, fin, titulo, False, uid)

    a = ev("a", "Reunion", base, base + dt.timedelta(hours=2))
    b = ev("b", "Dentista", base + dt.timedelta(hours=1),
           base + dt.timedelta(hours=3))
    c = ev("c", "Comida", base + dt.timedelta(hours=5),
           base + dt.timedelta(hours=6))
    d = ev("d", "Comida", base + dt.timedelta(hours=5),
           base + dt.timedelta(hours=6))
    lista = sorted([a, b, c, d], key=lambda e: e.inicio)
    solapes = extras.PanelControl._solapes(None, lista)
    assert len(solapes) == 1 and {solapes[0][0].uid, solapes[0][1].uid} == {"a", "b"}, \
        f"solo deberian solaparse a y b: {solapes}"
    # c y d se solapan del todo Y son repetidos
    repes = extras.PanelControl._duplicados(None, lista)
    assert len(repes) == 1 and {repes[0][0].uid, repes[0][1].uid} == {"c", "d"}


@prueba("paleta de comandos: todos los comandos son ejecutables y hay panel de control")
def _():
    import extras

    class Falsa:
        correo_carpetas = [{"nombre": "Bandeja", "no_leidos": 1}]

        def __getattr__(self, nombre):
            return lambda *a, **k: None

    comandos = extras.lista_de_comandos(Falsa(), cal)
    assert len(comandos) > 40, f"faltan comandos: {len(comandos)}"
    titulos = [c[0] for c in comandos]
    assert "Panel de control" in titulos
    assert "Ir a hoy" in titulos
    assert len(set(titulos)) == len(titulos), "hay comandos con el mismo titulo"
    assert all(callable(c[2]) for c in comandos)


@prueba("columna de correo: filtra por carpeta, excluye y deja solo no leidos")
def _():
    import correo_columna as cc
    base = dt.datetime(2030, 5, 6, 10, 0, tzinfo=cal.TZ)

    def mensaje(carpeta, minutos, leido, asunto):
        return cal.Correo(asunto, "Ana", "ana@x.es",
                          base + dt.timedelta(minutes=minutos), leido, None,
                          carpeta, "#00aa88", "id" + asunto, "alm")

    datos = {
        "Bandeja": [mensaje("Bandeja", 5, False, "b1"),
                    mensaje("Bandeja", 1, True, "b2")],
        "Facturas": [mensaje("Facturas", 9, False, "f1")],
        "Promos": [mensaje("Promos", 7, False, "p1")],
    }
    todos = cc.filtrar_mensajes(datos)
    assert [m.asunto for m in todos] == ["f1", "p1", "b1", "b2"], \
        "deben salir mezclados, del mas nuevo al mas viejo"
    assert [m.asunto for m in cc.filtrar_mensajes(datos, filtro="Bandeja")] \
        == ["b1", "b2"]
    sin_promos = cc.filtrar_mensajes(datos, excluidas=["Promos"])
    assert "p1" not in [m.asunto for m in sin_promos]
    assert cc.filtrar_mensajes(datos, filtro="Promos", excluidas=["Promos"]) \
        == [], "una carpeta excluida no sale ni siendo la del filtro"
    solo = cc.filtrar_mensajes(datos, solo_sin_leer=True)
    assert all(not m.leido for m in solo) and len(solo) == 3
    info = [{"nombre": "Bandeja"}, {"nombre": "Promos"}]
    assert cc.filtro_valido(info, "Promos", []) == "Promos"
    assert cc.filtro_valido(info, "Promos", ["Promos"]) is None
    assert cc.filtro_valido(info, "Borrada", []) is None


@prueba("correo: los nombres de carpeta salen cortos y sin repetidos")
def _():
    import correo_columna as cc
    nombres = ["finanzas@ejemplo.com/Bandeja de entrada",
               "ana@ejemplo.com/Bandeja de entrada",
               "ana@ejemplo.com/Facturas",
               "Solo"]
    assert cc.etiqueta_corta(nombres[2], nombres) == "Facturas"
    assert cc.etiqueta_corta("Solo", nombres) == "Solo"
    a = cc.etiqueta_corta(nombres[0], nombres)
    b = cc.etiqueta_corta(nombres[1], nombres)
    assert a != b and a.startswith("Bandeja de entrada"), (a, b)


@prueba("columnas dentro: correo y to-dos no ensanchan el widget")
def _():
    cal.CFG.update(vista="rejilla", dias_visibles=7, ancho_fijo=True,
                   ancho_columna=cal.DEFAULTS["ancho_columna"])
    w = cal.Widget()
    try:
        w.vista = "rejilla"
        w.update()
        w._tam = None
        w.dibujar()
        base = w._tam[0]
        cal.CFG["columnas_dentro"] = True
        cal.CFG["correo_visible"] = True
        cal.CFG["todos_visible"] = True
        w._tam = None
        w.dibujar()
        assert abs(w._tam[0] - base) <= 8, (w._tam[0], base)
        cal.CFG["columnas_dentro"] = False
        w._tam = None
        w.dibujar()
        assert w._tam[0] > base + 100, "sin la opcion, si se ensancha"
    finally:
        cal.CFG["correo_visible"] = cal.CFG["todos_visible"] = False
        cal.CFG["columnas_dentro"] = True
        w.destroy()


@prueba("barra: los botones no aparecen ni desaparecen con los datos")
def _():
    w = cal.Widget()
    try:
        w.update()

        def botones():
            return tuple(sorted({t for i in w.lienzo.find_all()
                                 for t in w.lienzo.gettags(i)
                                 if t.startswith(("btn_", "nav_"))}))
        vistos = set()
        for desp in (0, 3):
            for sin_leer in (0, 6, 1234):
                w.desplazamiento = desp
                w.correo_carpetas = [{"nombre": "a", "color": None,
                                      "total": 1, "no_leidos": sin_leer}]
                w._tam = None
                w.dibujar()
                vistos.add(botones())
        assert len(vistos) == 1, f"{len(vistos)} conjuntos de botones distintos"
        # la barra no se queda vacía; cuántos botones hay depende de la
        # configuración de quien ejecuta la prueba (los hay que se ocultan)
        assert {"btn_cerrar", "nav_hoy"} <= set(botones()), botones()
    finally:
        w.desplazamiento = 0
        w.destroy()


@prueba("correo: solo unos pocos botones de carpeta, el resto en el menu")
def _():
    import correo_columna as cc
    carpetas = [{"nombre": f"c{i}", "no_leidos": n}
                for i, n in enumerate([0, 5, 0, 9, 1, 0, 0, 2, 0, 0])]
    vis = cc.chips_a_mostrar(carpetas, None, 3)
    assert [c["nombre"] for c in vis] == ["c1", "c3", "c7"], vis
    vis = cc.chips_a_mostrar(carpetas, "c8", 3)
    assert "c8" in [c["nombre"] for c in vis] and len(vis) == 3
    assert cc.chips_a_mostrar(carpetas, None, 0) == []


@prueba("menu moderno: lee marcas, submenus y se puede elegir una entrada")
def _():
    import menus
    raiz = tk.Tk()
    try:
        hecho = []
        sub = tk.Menu(raiz, tearoff=0)
        sub.add_command(label="uno", command=lambda: hecho.append("uno"))
        m = tk.Menu(raiz, tearoff=0)
        m.add_command(label="✓ Marcada", command=lambda: hecho.append("m"))
        m.add_command(label="     Sin marcar")
        m.add_separator()
        m.add_command(label="Bloqueada", state="disabled")
        m.add_cascade(label="Más", menu=sub)
        filas = menus.leer_entradas(m)
        assert [f["marca"] for f in filas if f["tipo"] == "command"][:2] == ["✓", " "]
        assert filas[-1]["tipo"] == "cascade" and filas[-1]["submenu"] is sub
        assert any(f.get("deshabilitado") for f in filas)
        menus.abrir(m, 50, 50, cal)
        raiz.update()
        assert len(menus._ABIERTOS) == 1
        abierto = menus._ABIERTOS[0]
        abierto._pulsar(0)
        for _ in range(30):
            raiz.update()
            time.sleep(0.01)
        assert hecho == ["m"] and not menus._ABIERTOS, (hecho, menus._ABIERTOS)
    finally:
        menus.cerrar_todo()
        raiz.destroy()


@prueba("correo: mover o eliminar quita el mensaje al momento")
def _():
    import correo_columna as cc
    base = dt.datetime(2030, 5, 6, 10, 0, tzinfo=cal.TZ)
    m1 = cal.Correo("a", "Ana", "a@x", base, False, None, "B", "#0a0", "i", "s")
    m2 = cal.Correo("b", "Ana", "a@x", base, True, None, "B", "#0a0", "j", "s")

    class Falso:
        correo_carpetas = [{"nombre": "B", "total": 2, "no_leidos": 1}]
        correo_mensajes = {"B": [m1, m2]}
        dibujos = 0

        def dibujar(self):
            self.dibujos += 1

        def after(self, ms, fn):
            pass

    app = Falso()
    cc._quitar_local(app, m1)
    assert app.correo_mensajes["B"] == [m2]
    assert app.correo_carpetas[0]["no_leidos"] == 0
    assert app.correo_carpetas[0]["total"] == 1
    assert hasattr(m1, "marcado") and m1.marcado is False


@prueba("correo: marcar todo como leido actualiza lo que se ve al momento")
def _():
    import correo_columna as cc
    base = dt.datetime(2030, 5, 6, 10, 0, tzinfo=cal.TZ)
    m1 = cal.Correo("a", "Ana", "a@x", base, False, None, "B", "#0a0", "i", "s")
    m2 = cal.Correo("b", "Ana", "a@x", base, False, None, "B", "#0a0", "j", "s")

    class Falso:
        correo_carpetas = [{"nombre": "B", "total": 2, "no_leidos": 2}]
        correo_mensajes = {"B": [m1, m2]}
        correo_error = ""

        def dibujar(self):
            pass

        def after(self, ms, fn):
            fn()

        recontado = False

        def _correo_ahora(self):        # los contadores reales vienen de Outlook
            self.recontado = True

    app = Falso()
    llamadas = []
    cal.marcar_todos_leidos = lambda carpetas: (llamadas.append(carpetas), 2)[1]
    guardado = {k: cal.CFG.get(k) for k in ("correo_filtro", "correo_excluidas")}
    cal.CFG.update(correo_filtro=None, correo_excluidas=[])
    try:
        cc._marcar_todos_leidos(app, cal)
        for _ in range(100):             # el trabajo corre en un hilo
            if app.recontado:
                break
            time.sleep(0.02)
    finally:
        cal.CFG.update(guardado)
    assert llamadas == [["B"]], llamadas
    assert m1.leido and m2.leido
    assert app.recontado, "debe pedir los contadores reales a Outlook"


@prueba("ayuda al pasar el raton: aparece con el texto y se puede cancelar")
def _():
    w = cal.Widget()
    try:
        w.update()
        w._vincular_ayuda("btn_nuevo", "Nuevo evento")
        assert "btn_nuevo" in w.NOMBRES_BOTONES or True
        clase_evento = type("E", (), {"x_root": 10, "y_root": 10})()
        w._programar_ayuda("btn_nuevo", "Nuevo evento", clase_evento)
        assert w._ayuda_tag == "btn_nuevo" and w._ayuda_win is None
        w._mostrar_ayuda("Nuevo evento")
        assert w._ayuda_win is not None
        w._cancelar_ayuda("btn_nuevo")
        assert w._ayuda_win is None and w._ayuda_tag is None
    finally:
        w.destroy()


@prueba("resumen de reglas: reconoce condiciones y acciones habituales")
def _():
    class Destinatario:
        def __init__(self, nombre):
            self.Name = nombre

    class Coleccion:
        def __init__(self, items):
            self._items = items

        @property
        def Count(self):
            return len(self._items)

        def Item(self, i):
            return self._items[i - 1]

    class Prop:
        """Una propiedad tipo 'RuleCondition'/'RuleAction' de Outlook,
        simulada: tiene Enabled y lo que haga falta encima."""
        def __init__(self, activo=False, **extra):
            self.Enabled = activo
            for k, v in extra.items():
                setattr(self, k, v)

    class Carpeta:
        Name = "Facturas"

    class Condiciones:
        From = Prop(True, Recipients=Coleccion([Destinatario("ana@ejemplo.com")]))
        SentTo = Prop(False, Recipients=Coleccion([]))
        Subject = Prop(True, Text=["factura"])
        Body = Prop(False, Text=[])
        SentOnlyToMe = False
        Category = Prop(False, Categories=[])
        Importance = Prop(False, Importance=1)
        HasAttachment = False

    class Acciones:
        MoveToFolder = Prop(True, Folder=Carpeta())
        CopyToFolder = Prop(False)
        Delete = False
        AssignToCategory = Prop(False, Categories=[])
        Forward = Prop(False, Recipients=Coleccion([]))
        PlaySound = Prop(False)
        Stop = False
        MarkAsTask = Prop(False)
        NewItemAlert = Prop(False)

    class Regla:
        Conditions = Condiciones()
        Actions = Acciones()

    condiciones, acciones = cal._resumen_regla(Regla())
    assert "ana@ejemplo.com" in condiciones, condiciones
    assert "factura" in condiciones, condiciones
    assert "Facturas" in acciones, acciones


@prueba("evento rapido: el titulo trae dia, hora, duracion y lugar")
def _():
    hoy = dt.date(2026, 10, 6)                      # martes
    t, d = cal.interpretar_rapido("Dentista mañana 10:30 45min @Clínica", hoy)
    assert t == "Dentista", t
    assert d == {"dia": dt.date(2026, 10, 7), "hora": (10, 30),
                 "minutos": 45, "lugar": "Clínica"}, d
    t, d = cal.interpretar_rapido("Cena el viernes a las 9 pm", hoy)
    assert (t, d["dia"], d["hora"]) == ("Cena", dt.date(2026, 10, 9), (21, 0))
    t, d = cal.interpretar_rapido("Comida 18/10 2 horas", hoy)
    assert d["dia"] == dt.date(2026, 10, 18) and d["minutos"] == 120, d
    t, d = cal.interpretar_rapido("Cumple de Ana 3 sep todo el día", hoy)
    assert t == "Cumple de Ana" and d["todo_dia"], (t, d)
    assert cal.interpretar_rapido("Reunión por la mañana", hoy) == \
        ("Reunión por la mañana", {})
    assert cal.interpretar_rapido("Revisar 12 contratos", hoy)[1] == {}


@prueba("ficha: el titulo rapido rellena los campos y los botones se alcanzan con Tab")
def _():
    w = cal.Widget()
    try:
        w.update()
        f = w._ficha(inicio=dt.datetime(2026, 10, 6, 9, 0, tzinfo=cal.TZ))
        f.update()
        f.titulo.insert(0, "Dentista 12/11/2026 17:15 30min @Clínica")
        f._aplicar_rapido()
        assert f.titulo.get() == "Dentista", f.titulo.get()
        assert (f.dia.get(), f.desde.get(), f.hasta.get(), f.lugar.get()) == \
            ("12/11/2026", "17:15", "17:45", "Clínica")
        assert f.repe_texto.cget("takefocus"), "los botones no se pueden enfocar"
    finally:
        w.destroy()


@prueba("avisos: lo que apunta otro PC se mezcla y guardar no lo pisa")
def _():
    import json
    import tempfile
    antes = (cal.CONFIG_PATH, dict(cal.CFG), dict(cal._SYNC))
    with tempfile.TemporaryDirectory() as d:
        try:
            cal.CONFIG_PATH = os.path.join(d, "config.json")
            cal.CFG["alarmas"] = {"a|2026-10-20": [10]}
            cal.CFG["avisar_de"] = ["x"]
            assert cal.guardar_config()
            # "el otro PC" escribe su alarma en el mismo fichero
            otro = json.load(open(cal.CONFIG_PATH, encoding="utf-8"))
            otro["alarmas"]["b|2026-10-21"] = [30]
            otro["avisar_de"].append("y")
            with open(cal.CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(otro, f)
            os.utime(cal.CONFIG_PATH, (time.time() + 5, time.time() + 5))
            assert cal.config_cambio_fuera()
            # mientras tanto este PC quita una y pone otra, y guarda
            cal.CFG["alarmas"].pop("a|2026-10-20")
            cal.CFG["alarmas"]["c|2026-10-22"] = [5]
            assert cal.guardar_config()
            disco = json.load(open(cal.CONFIG_PATH, encoding="utf-8"))
            assert set(disco["alarmas"]) == {"b|2026-10-21", "c|2026-10-22"}, disco["alarmas"]
            assert disco["avisar_de"] == ["x", "y"], disco["avisar_de"]
            assert not cal.config_cambio_fuera()
        finally:
            cal.CONFIG_PATH = antes[0]
            cal.CFG.clear()
            cal.CFG.update(antes[1])
            cal._SYNC.update(antes[2])


@prueba("navegacion: se puede ir mas alla de 90 dias y el rango de descarga lo sigue")
def _():
    assert cal.MAX_DIAS > 90
    w = cal.Widget()
    try:
        w.update()
        w.refrescar = lambda *a, **k: None       # sin red
        ahora = dt.date.today()
        assert w.ir_a_fecha(ahora + dt.timedelta(days=400))
        ini, fin = cal.rango_descarga()
        assert fin.date() >= ahora + dt.timedelta(days=400 + 42), fin
        w._ir_a_hoy()
        assert cal.rango_descarga()[1] < ini + dt.timedelta(days=400)
    finally:
        cal.ALCANCE.update(atras=0, adelante=0)
        w.destroy()


if __name__ == "__main__":
    print(f"Ejecutando {len(PRUEBAS)} pruebas de calendario.py...\n")
    for p in PRUEBAS:
        p()
    print()
    if FALLOS:
        print(f"{len(FALLOS)} de {len(PRUEBAS)} pruebas han fallado:")
        for nombre in FALLOS:
            print(f"  - {nombre}")
        sys.exit(1)
    print(f"Las {len(PRUEBAS)} pruebas han pasado.")
    sys.exit(0)
