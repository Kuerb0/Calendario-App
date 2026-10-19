"""Extras del widget: Panel de control y Paleta de comandos.

Los dos se pueden abrir como ventana aparte o anclados a un lado del widget
(la misma opción que ya tienen Correo, Hábitos, Cumpleaños...). Esa elección
se hace con el botón de la cabecera de cada panel o en Panel de control >
Sistema, y se guarda en CFG["ubicacion_vistas"].

Este módulo no importa calendario.py (así no se ejecuta dos veces): recibe el
módulo ya cargado como 'M' y la ventana principal como 'app'.
"""
import datetime as dt
import os
import threading
import time
import tkinter as tk
from tkinter import messagebox

# --------------------------------------------------------------------------
# Botones de la barra superior que el usuario puede quitar o poner
# (etiqueta, texto que se lee en el panel). La clave es la etiqueta de canvas
# del botón; "btn_orient" agrupa los botones H y V.
# --------------------------------------------------------------------------

BOTONES_BARRA = (
    ("btn_orient", "H / V · orientación"),
    ("btn_nivel", "Fondo / Normal / Encima"),
    ("btn_avisos", "🔔 · avisos"),
    ("btn_cuenta", "⏳ · cuenta atrás"),
)

# los que nunca se pueden quitar: sin ellos no habría cómo volver atrás
BOTONES_FIJOS = ("btn_dias", "btn_refrescar", "btn_ocultar")

PESTANAS = (("vista", "Vista"), ("barra", "Barra"), ("avisos", "Avisos"),
            ("agenda", "Agenda"), ("correo", "Correo"),
            ("eventos", "Eventos"), ("cuenta", "Cuenta atrás"),
            ("sistema", "Sistema"))

# paneles cuya ubicación se puede elegir desde Sistema: (clave, nombre, atributo)
PANELES_UBICABLES = (
    ("control", "Panel de control", "_panel_control"),
    ("paleta", "Paleta de comandos", "_paleta"),
    ("correo", "Correo", "_panel_correo"),
    ("habitos", "Hábitos", "_panel_habitos"),
    ("cumpleanos", "Cumpleaños", "_panel_cumpleanos"),
    ("plantillas", "Plantillas", "_panel_plantillas"),
    ("tareas", "Tareas permanentes", "_panel_tareas"),
)


def abrir_ruta(ruta):
    """Abre un fichero o carpeta con el programa de Windows que toque."""
    try:
        os.startfile(ruta)
    except AttributeError:               # por si se prueba fuera de Windows
        import subprocess
        subprocess.Popen(["xdg-open", ruta])


def dias_hasta(fecha, hoy=None):
    """Días que faltan hasta 'fecha' (dt.date); negativo si ya pasó."""
    return (fecha - (hoy or dt.date.today())).days


def color_cuenta(dias):
    """Rojo 0-5 días, ámbar 6-10, verde más de 10; gris si ya pasó."""
    if dias < 0:
        return "#7a7a7e"
    return "#ff453a" if dias <= 5 else "#ff9f0a" if dias <= 10 else "#30d158"


def leer_fecha(texto):
    """'25/12/2026', '25-12-2026' o '2026-12-25' -> dt.date (ValueError si no)."""
    texto = texto.strip().replace("/", "-")
    for formato in ("%d-%m-%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(texto, formato).date()
        except ValueError:
            pass
    raise ValueError(texto)


def barra_oculta(M, etiqueta):
    """¿El usuario ha quitado este botón de la barra?"""
    if etiqueta in BOTONES_FIJOS:
        return False
    return etiqueta in (M.CFG.get("barra_ocultos") or [])


# --------------------------------------------------------------------------
# Base común: ventana sin bordes, cabecera con título, ✕ y botón de ubicación
# --------------------------------------------------------------------------

class PanelBase(tk.Toplevel):
    CLAVE = ""            # clave en CFG["ubicacion_vistas"]
    ATRIBUTO = ""         # dónde la ventana principal guarda este panel
    TITULO = ""
    ANCHO = 340

    def __init__(self, app, M):
        super().__init__(app)
        self.app = app
        self.M = M
        self._clave_ubicacion = self.CLAVE
        self._colocado_ya = False
        self._abierto_en = time.monotonic()
        self.overrideredirect(True)
        self.configure(bg=M.C["rejilla"])
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass
        self.marco = tk.Frame(self, bg=M.C["fondo"], padx=8, pady=6)
        self.marco.pack(fill="both", expand=True, padx=1, pady=1)
        self.bind("<Escape>", lambda e: self.destroy())

    # -- piezas de dibujo ----------------------------------------------------

    def cabecera(self, padre):
        M = self.M
        cab = tk.Frame(padre, bg=M.C["fondo"], cursor="fleur")
        cab.pack(fill="x")
        tk.Label(cab, text=self.TITULO, bg=M.C["fondo"], fg=M.C["texto"],
                 font=M.letra(10, "bold"), cursor="fleur").pack(side="left")
        cerrar = tk.Label(cab, text="✕", bg=M.C["fondo"],
                          fg=M.C["texto_tenue"], font=M.letra(9),
                          cursor="hand2")
        cerrar.pack(side="right")
        cerrar.bind("<Button-1>", lambda e: self.destroy())
        M.boton_ubicacion_panel(cab, self.CLAVE, self)
        cab.bind("<Button-1>", self._agarrar)
        cab.bind("<B1-Motion>", self._mover_panel)
        cab.bind("<Button-3>",
                 lambda e: M.menu_ubicacion_panel(self.CLAVE, self, e))
        return cab

    def etiqueta(self, padre, texto, tenue=False, negrita=False, ancho=None):
        M = self.M
        return tk.Label(
            padre, text=texto, bg=M.C["fondo"],
            fg=M.C["texto_tenue"] if tenue else M.C["texto"],
            font=M.letra(7 if tenue else 8, "bold" if negrita else None),
            anchor="w", justify="left",
            wraplength=ancho or self.ANCHO - 24)

    def seccion(self, padre, texto):
        self.etiqueta(padre, texto, negrita=True).pack(fill="x", pady=(8, 2))

    def boton(self, padre, texto, accion, activo=False, lado=None,
              expandir=False, ancho=None):
        M = self.M
        b = tk.Label(padre, text=texto,
                     bg=M.C["evento"] if activo else M.C.get("boton", "#2f2f33"),
                     fg=M.C["evento_texto"] if activo else M.C["texto"],
                     font=M.letra(8, "bold" if activo else None),
                     padx=7, pady=3, cursor="hand2", width=ancho or 0)
        if lado:
            b.pack(side=lado, padx=(0, 3), pady=1, fill="x" if expandir else None,
                   expand=expandir)
        else:
            b.pack(fill="x", pady=1)
        b.bind("<Button-1>", lambda e: self._pulsar(accion))
        return b

    def casilla(self, padre, texto, valor, accion):
        """Fila con ☑/☐ que cambia al pulsarla. 'valor' es una función."""
        M = self.M
        fila = tk.Label(padre, anchor="w", bg=M.C["fondo"], fg=M.C["texto"],
                        font=M.letra(8), cursor="hand2", padx=2, pady=2)

        def poner():
            fila.config(text=("☑  " if valor() else "☐  ") + texto)

        def alternar(_e=None):
            self._pulsar(accion)
            poner()

        poner()
        fila.pack(fill="x")
        fila.bind("<Button-1>", alternar)
        return fila

    def fila_botones(self, padre):
        f = tk.Frame(padre, bg=self.M.C["fondo"])
        f.pack(fill="x", pady=1)
        return f

    def _pulsar(self, accion):
        # ignora clics fantasma que llegan justo al abrirse el panel
        if time.monotonic() - self._abierto_en < 0.25:
            return
        try:
            accion()
        except Exception as exc:
            self.M.log(f"{self.TITULO}: {exc!r}")

    # -- colocación ----------------------------------------------------------

    def actualizar_posicion(self):
        M = self.M
        if M.ubicacion_panel(self.CLAVE) == "flotante" and self._colocado_ya:
            return
        try:
            self.update_idletasks()
            app = self.app
            x = M.x_panel_acoplado(app, self.CLAVE, self.ANCHO)
            y = app.winfo_rooty()
            alto = self.winfo_reqheight()
            self.geometry(f"{self.ANCHO}x{alto}"
                          + M.encajar_en_pantalla(app, x, y, self.ANCHO, alto))
            self._colocado_ya = True
        except Exception:
            pass

    def _agarrar(self, e):
        self._agarre = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    def _mover_panel(self, e):
        if getattr(self, "_agarre", None):
            self.geometry(f"+{e.x_root - self._agarre[0]}"
                          f"+{e.y_root - self._agarre[1]}")

    def destroy(self):
        try:
            if getattr(self.app, self.ATRIBUTO, None) is self:
                setattr(self.app, self.ATRIBUTO, None)
            self.app.dibujar()          # el botón de la barra deja de estar marcado
        except Exception:
            pass
        super().destroy()


# --------------------------------------------------------------------------
# Panel de control
# --------------------------------------------------------------------------

class PanelControl(PanelBase):
    CLAVE = "control"
    ATRIBUTO = "_panel_control"
    TITULO = "⚙ Panel de control"
    ANCHO = 340

    def __init__(self, app, M, pestana=None):
        super().__init__(app, M)
        validas = [c for c, _ in PESTANAS]
        pestana = pestana or M.CFG.get("control_pestana") or "vista"
        self.pestana = pestana if pestana in validas else "vista"
        self._modo_eventos = "proximos"
        self._filtro_cal = None
        self._reglas = None            # None = sin cargar; lista = cargadas
        self._reglas_estado = ""
        self._reglas_hilo = None
        self._lista_carpetas = None    # todas las carpetas de Outlook
        self._elegir_carpetas = False  # ¿está abierto el selector?
        self._pendiente = None         # resultado del hilo, lo recoge after()
        self._texto_busqueda = ""
        self.pintar()
        self.actualizar_posicion()

    # -- estructura ----------------------------------------------------------

    def cambiar_pestana(self, clave):
        self.pestana = clave
        self.M.CFG["control_pestana"] = clave
        self.pintar()
        self.actualizar_posicion()

    def pintar(self):
        M = self.M
        for hijo in self.marco.winfo_children():
            hijo.destroy()
        self.cabecera(self.marco)
        # pestañas en dos filas para que quepan en un panel estrecho
        for grupo in (PESTANAS[:4], PESTANAS[4:]):
            fila = self.fila_botones(self.marco)
            for clave, texto in grupo:
                self.boton(fila, texto, lambda c=clave: self.cambiar_pestana(c),
                           activo=clave == self.pestana, lado="left",
                           expandir=True)
        tk.Frame(self.marco, bg=M.C["rejilla"], height=1).pack(fill="x",
                                                               pady=(6, 2))
        cuerpo = tk.Frame(self.marco, bg=M.C["fondo"])
        cuerpo.pack(fill="both")
        getattr(self, "_pestana_" + self.pestana)(cuerpo)

    def _repintar(self):
        self.pintar()
        self.actualizar_posicion()

    def _y_dibujar(self, fn):
        """Ejecuta un cambio del widget, redibuja el panel y la ventana."""
        def hacer():
            fn()
            try:
                self.app.dibujar()
            except Exception:
                pass
            self._repintar()
        return hacer

    # -- pestaña Vista ---------------------------------------------------------

    def _pestana_vista(self, p):
        M, app = self.M, self.app
        self.seccion(p, "Vista principal")
        f = self.fila_botones(p)
        for clave, texto in (("rejilla", "Calendario"), ("ambas", "Cal + lista"),
                             ("lista", "Lista"), ("mes", "Mes")):
            self.boton(f, texto,
                       self._y_dibujar(lambda c=clave: app._cambiar_vista(c)),
                       activo=app.vista == clave, lado="left", expandir=True)
        self.casilla(p, "Lista, correo y to-dos dentro del tamaño actual "
                        "(sin ensanchar el widget)",
                     lambda: M.CFG.get("columnas_dentro", True),
                     lambda: app._alternar_ajuste("columnas_dentro"))
        self.etiqueta(p, "Abrir en ventana aparte", tenue=True).pack(
            fill="x", pady=(4, 0))
        f = self.fila_botones(p)
        for clave, texto in (("rejilla", "Calendario"), ("ambas", "Cal + lista"),
                             ("lista", "Lista"), ("mes", "Mes")):
            self.boton(f, texto, lambda c=clave: app.abrir_vista_aparte(c),
                       lado="left", expandir=True)

        self.seccion(p, "Días a la vista")
        f = self.fila_botones(p)
        for n in (1, 3, 5, 7, 14, 21):
            self.boton(f, str(n),
                       self._y_dibujar(lambda n=n: app._cambiar_dias(n)),
                       activo=M.CFG.get("dias_visibles") == n, lado="left",
                       expandir=True)

        self.seccion(p, "Orientación")
        f = self.fila_botones(p)
        for clave, texto in (("horizontal", "Horizontal"),
                             ("vertical", "Vertical")):
            self.boton(f, texto,
                       self._y_dibujar(lambda c=clave: app._cambiar_orientacion(c)),
                       activo=app.orientacion == clave, lado="left",
                       expandir=True)

        self.seccion(p, "Posición respecto a las demás ventanas")
        f = self.fila_botones(p)
        for clave, texto in M.NIVELES.items():
            self.boton(f, texto,
                       self._y_dibujar(lambda c=clave: app._cambiar_nivel(c)),
                       activo=app.nivel == clave, lado="left", expandir=True)

        self.seccion(p, "Tema")
        f = self.fila_botones(p)
        for clave, texto in (("auto", "Como Windows"), ("claro", "Claro"),
                             ("oscuro", "Oscuro")):
            self.boton(f, texto,
                       self._y_dibujar(lambda c=clave: app._cambiar_tema(c)),
                       activo=M.CFG.get("tema", "oscuro") == clave,
                       lado="left", expandir=True)

        self.seccion(p, "Tamaño y transparencia")
        f = self.fila_botones(p)
        self.boton(f, "−", self._y_dibujar(lambda: app._cambiar_escala(-0.1)),
                   lado="left", ancho=3)
        self.etiqueta(f, f"Tamaño {int(round(M.ESCALA * 100))} %").pack(
            side="left", padx=6)
        self.boton(f, "+", self._y_dibujar(lambda: app._cambiar_escala(0.1)),
                   lado="left", ancho=3)
        self.etiqueta(p, "Transparencia", tenue=True).pack(fill="x",
                                                           pady=(4, 0))
        escala = tk.Scale(p, from_=30, to=100, orient="horizontal",
                          showvalue=True, length=self.ANCHO - 30,
                          bg=M.C["fondo"], fg=M.C["texto"], troughcolor=M.C["rejilla"],
                          highlightthickness=0, bd=0, font=M.letra(7))
        escala.set(int(round(M.CFG.get("opacidad", 1.0) * 100)))
        escala.pack(fill="x")
        escala.config(command=lambda v: self._vista_previa_opacidad(v))
        escala.bind("<ButtonRelease-1>",
                    lambda e: app._cambiar_opacidad(escala.get() / 100))

    def _vista_previa_opacidad(self, valor):
        try:
            self.app.attributes("-alpha", max(0.3, min(1.0, float(valor) / 100)))
        except Exception:
            pass

    # -- pestaña Barra -----------------------------------------------------------

    def _pestana_barra(self, p):
        M, app = self.M, self.app
        self.etiqueta(p, "Elige qué botones salen en la barra de arriba. "
                         "Si la ventana es estrecha, los menos importantes se "
                         "esconden solos.", tenue=True).pack(fill="x",
                                                             pady=(0, 4))

        def alternar(etiqueta):
            ocultos = list(M.CFG.get("barra_ocultos") or [])
            if etiqueta in ocultos:
                ocultos.remove(etiqueta)
            else:
                ocultos.append(etiqueta)
            M.CFG["barra_ocultos"] = ocultos
            M.guardar_config()
            app.dibujar()

        for etiqueta, texto in BOTONES_BARRA:
            self.casilla(p, texto, lambda e=etiqueta: not barra_oculta(M, e),
                         lambda e=etiqueta: alternar(e))
        self.seccion(p, "Tamaño de la barra")
        fz = self.fila_botones(p)

        def poner_zoom(v):
            M.CFG["barra_zoom"] = v
            M.guardar_config()
            M.aplicar_escala()
            app.dibujar()
        for nombre, v in (("Pequeña", 1.0), ("Media", 1.5), ("Grande", 2.0)):
            self.boton(fz, nombre, lambda v=v: poner_zoom(v), lado="left")
        self.etiqueta(p, "Los botones de días, actualizar y ocultar no se "
                         "pueden quitar. Todo sigue disponible con el clic "
                         "derecho y con Ctrl+K.", tenue=True).pack(fill="x",
                                                                  pady=(6, 2))
        f = self.fila_botones(p)

        def mostrar_todos():
            M.CFG["barra_ocultos"] = []
            M.guardar_config()
            app.dibujar()
        self.boton(f, "Mostrar todos", self._y_dibujar(mostrar_todos),
                   lado="left")

    # -- pestaña Avisos ------------------------------------------------------------

    def _pestana_avisos(self, p):
        M, app = self.M, self.app
        self.casilla(p, "Avisos activados (la campana)",
                     lambda: M.CFG.get("notificaciones_activas", True),
                     lambda: app._alternar_notificaciones())
        self.casilla(p, "«Ahora» y «luego» en la franja de abajo",
                     lambda: M.CFG.get("barra_proximos", True),
                     lambda: app._alternar_ajuste("barra_proximos"))

        self.seccion(p, "Avisar antes del evento (minutos)")
        self.etiqueta(p, "Separa con comas, por ejemplo 60, 10", tenue=True
                      ).pack(fill="x")
        f = self.fila_botones(p)
        self._var_avisos = tk.StringVar(
            value=", ".join(str(n) for n in M.CFG.get("avisos_minutos", [])))
        e = tk.Entry(f, textvariable=self._var_avisos,
                     bg=M.C.get("boton", "#2f2f33"), fg=M.C["texto"],
                     insertbackground=M.C["texto"], relief="flat",
                     font=M.letra(8), width=16)
        e.pack(side="left", padx=(0, 4), ipady=3)
        self.mensaje_avisos = self.etiqueta(p, "", tenue=True)
        self.boton(f, "Guardar", self._guardar_avisos, lado="left")
        e.bind("<Return>", lambda ev: self._guardar_avisos())
        self.mensaje_avisos.pack(fill="x")

        self.seccion(p, "Avisos de todo el día")
        f = self.fila_botones(p)
        self._var_todo_dia = tk.StringVar(
            value=M.CFG.get("aviso_todo_dia") or "")
        e2 = tk.Entry(f, textvariable=self._var_todo_dia,
                      bg=M.C.get("boton", "#2f2f33"), fg=M.C["texto"],
                      insertbackground=M.C["texto"], relief="flat",
                      font=M.letra(8), width=8)
        e2.pack(side="left", padx=(0, 4), ipady=3)
        self.boton(f, "Guardar", self._guardar_todo_dia, lado="left")
        self.etiqueta(p, "Hora en formato HH:MM. Vacío = sin aviso.",
                      tenue=True).pack(fill="x")

        self.seccion(p, "Probar")
        f = self.fila_botones(p)
        self.boton(f, "Ver un aviso de ejemplo",
                   lambda: app._probar_aviso(), lado="left")

    def _guardar_avisos(self):
        M = self.M
        try:
            valores = [int(t) for t in self._var_avisos.get().replace(";", ",")
                       .split(",") if t.strip()]
            if any(v < 0 or v > 10080 for v in valores):
                raise ValueError
        except ValueError:
            self.mensaje_avisos.config(text="Escribe solo números, "
                                            "por ejemplo 60, 10")
            return
        M.CFG["avisos_minutos"] = sorted(set(valores), reverse=True)
        M.guardar_config()
        self.mensaje_avisos.config(text="Guardado")

    def _guardar_todo_dia(self):
        M = self.M
        texto = self._var_todo_dia.get().strip()
        if texto:
            try:
                dt.datetime.strptime(texto, "%H:%M")
            except ValueError:
                return
        M.CFG["aviso_todo_dia"] = texto or None
        M.guardar_config()

    # -- pestaña Agenda (calendarios, origen y festivos) ----------------------------

    def _pestana_agenda(self, p):
        M, app = self.M, self.app
        self.seccion(p, "De dónde salen los eventos")
        f = self.fila_botones(p)
        for clave, texto in (("auto", "Automático"), ("outlook", "Outlook"),
                             ("ics", "Enlace ICS")):
            self.boton(f, texto,
                       self._y_dibujar(lambda c=clave: app._cambiar_origen(c)),
                       activo=M.CFG.get("origen", "auto") == clave,
                       lado="left", expandir=True)

        self.seccion(p, "Calendarios (enlaces ICS)")
        calendarios = M.CFG.get("calendarios") or []
        if not any(c.get("url") for c in calendarios):
            self.etiqueta(p, "Ninguno configurado", tenue=True).pack(fill="x")
        for c in calendarios:
            if not c.get("url"):
                continue
            fila = tk.Frame(p, bg=M.C["fondo"])
            fila.pack(fill="x")
            tk.Frame(fila, bg=c.get("color", M.C["evento"]), width=10,
                     height=10).pack(side="left", padx=(2, 4))
            self.casilla(fila, c.get("nombre", "Calendario"),
                         lambda c=c: c.get("activo", True),
                         lambda c=c: self._alternar_calendario(c))

        self.seccion(p, "Festivos")
        hay = False
        for clave, grupo in (M.CFG.get("festivos") or {}).items():
            if not isinstance(grupo, dict):
                continue
            hay = True
            nombre = grupo.get("nombre") or clave.capitalize()
            self.casilla(p, nombre, lambda g=grupo: g.get("activo", True),
                         lambda k=clave, g=grupo:
                         app._cambiar_festivo(k, not g.get("activo", True)))
        if not hay:
            self.etiqueta(p, "Sin grupos de festivos", tenue=True).pack(fill="x")

        self.seccion(p, "Acciones")
        f = self.fila_botones(p)
        self.boton(f, "Actualizar ahora", lambda: app.refrescar(), lado="left")
        self.boton(f, "Abrir el asistente…", lambda: app._abrir_asistente(),
                   lado="left")

    def _alternar_calendario(self, c):
        c["activo"] = not c.get("activo", True)
        self.M.guardar_config()
        self.app.refrescar()

    # -- pestaña Correo (bandeja y reglas de Outlook) -------------------------------

    def _pestana_correo(self, p):
        M, app = self.M, self.app
        import correo_columna as cc
        self.seccion(p, "Correo dentro del widget")
        self.casilla(p, "Mostrar la columna de correo",
                     lambda: M.CFG.get("correo_visible", False),
                     lambda: cc.alternar_columna(app, M))
        self.casilla(p, "Solo los no leídos",
                     lambda: M.CFG.get("correo_solo_sin_leer", False),
                     lambda: cc.alternar_sin_leer(app, M))
        self.etiqueta(p, "Botones de carpeta a la vista en la columna (el "
                         "resto va al menú ▾)", tenue=True).pack(
            fill="x", pady=(4, 0))
        fz = self.fila_botones(p)
        for n in (0, 2, 3, 5, 8):
            self.boton(fz, str(n),
                       lambda n=n: (cc._poner_chips(app, M, n), self._repintar()),
                       activo=M.CFG.get("correo_chips_max", 3) == n,
                       lado="left", expandir=True)
        self._selector_carpetas(p)
        filtro, fuera, _ = cc.estado(app, M)
        self.etiqueta(p, "Carpetas: marca las que quieres ver. «Solo» deja "
                         "únicamente esa.", tenue=True).pack(fill="x",
                                                            pady=(4, 0))
        if not app.correo_carpetas:
            self.etiqueta(p, "Aún no se han cargado las carpetas.",
                          tenue=True).pack(fill="x")
        for f in app.correo_carpetas:
            nombre = f["nombre"]
            corto = cc.etiqueta_corta(
                nombre, [g["nombre"] for g in app.correo_carpetas])
            fila = tk.Frame(p, bg=M.C["fondo"])
            fila.pack(fill="x")
            solo = tk.Label(fila, text="Solo", padx=6, pady=1, cursor="hand2",
                            bg=M.C["evento"] if filtro == nombre
                            else M.C.get("boton", "#2f2f33"),
                            fg=M.C["evento_texto"] if filtro == nombre
                            else M.C["texto"], font=M.letra(7))
            solo.pack(side="right")
            solo.bind("<Button-1>", lambda e, n=nombre: self._pulsar(
                lambda: (cc.poner_filtro(app, M, n), self._repintar())))
            cuenta = f" ({f['no_leidos']})" if f.get("no_leidos") else ""
            self.casilla(fila, corto + cuenta,
                         lambda n=nombre: n not in
                         (M.CFG.get("correo_excluidas") or []),
                         lambda n=nombre: cc.alternar_exclusion(app, M, n))
        f = self.fila_botones(p)
        self.boton(f, "Quitar filtros", lambda: (cc.limpiar_filtros(app, M),
                                                 self._repintar()),
                   lado="left")

        self.seccion(p, "Bandeja y reglas")
        f = self.fila_botones(p)
        self.boton(f, "Abrir bandeja", lambda: app._abrir_correo(modo="bandeja"),
                   lado="left", expandir=True)
        self.boton(f, "Panel de reglas", lambda: app._abrir_correo(modo="reglas"),
                   lado="left", expandir=True)
        f = self.fila_botones(p)
        self.boton(f, "Actualizar correo ahora", lambda: app.refrescar_correo(),
                   lado="left")

        self.seccion(p, "Reglas de Outlook")
        if self._reglas is None:
            if self._reglas_hilo is None:
                self._cargar_reglas()
            self.etiqueta(p, "Leyendo las reglas de Outlook…",
                          tenue=True).pack(fill="x")
            return
        f = self.fila_botones(p)
        self.boton(f, "Recargar", self._cargar_reglas, lado="left")
        if self._reglas_estado:
            self.etiqueta(p, self._reglas_estado, tenue=True).pack(fill="x")
        if not self._reglas:
            self.etiqueta(p, "No hay reglas, o no se puede leer Outlook "
                             "(hace falta el Outlook clásico).",
                          tenue=True).pack(fill="x", pady=(4, 0))
            return
        for r in self._reglas:
            self._fila_regla(p, r)

    # -- qué carpetas de Outlook se leen ----------------------------------------

    def _selector_carpetas(self, p):
        """Elegir a mano las carpetas que se leen (en vez de todas las que
        haya creado la persona)."""
        M, app = self.M, self.app
        conf = [c for c in (M.CFG.get("carpetas_correo") or [])
                if c.get("activo", True)]
        self.seccion(p, "Carpetas que se leen de Outlook")
        self.etiqueta(p, (f"Elegidas por ti: {len(conf)}" if conf else
                          "Automático: bandeja de entrada y las carpetas que "
                          "has creado"), tenue=True).pack(fill="x")
        f = self.fila_botones(p)
        import correo_columna as cc
        self.boton(f, "Elegir carpetas…",
                   lambda: cc.abrir_selector_carpetas(app, M), lado="left")
        if conf:
            self.boton(f, "Automático", self._carpetas_automatico, lado="left")
        if not self._elegir_carpetas:
            return
        if self._lista_carpetas is None:
            self.etiqueta(p, "Leyendo las carpetas de Outlook…",
                          tenue=True).pack(fill="x")
            return
        if not self._lista_carpetas:
            self.etiqueta(p, "No se han podido leer las carpetas (hace falta "
                             "el Outlook clásico abierto).",
                          tenue=True).pack(fill="x")
            return
        self.etiqueta(p, "Marca las carpetas que quieres (Ctrl o Mayús para "
                         "varias) y pulsa Guardar.", tenue=True
                      ).pack(fill="x", pady=(2, 2))
        marco = tk.Frame(p, bg=M.C["fondo"])
        marco.pack(fill="x")
        barra = tk.Scrollbar(marco)
        barra.pack(side="right", fill="y")
        lista = tk.Listbox(
            marco, selectmode="multiple", exportselection=False,
            height=max(4, min(14, len(self._lista_carpetas))), activestyle="none",
            bg=M.C.get("boton", "#2f2f33"), fg=M.C["texto"],
            selectbackground=M.C["evento"], selectforeground=M.C["evento_texto"],
            relief="flat", highlightthickness=0, font=M.letra(8),
            yscrollcommand=barra.set)
        barra.config(command=lista.yview)
        lista.pack(side="left", fill="x", expand=True)
        actuales = ({c.get("buscar") for c in conf} if conf
                    else {c["nombre"] for c in app.correo_carpetas})
        for i, c in enumerate(self._lista_carpetas):
            n = c["nombre"]
            extra = f"  ({c['no_leidos']})" if c["no_leidos"] else ""
            lista.insert("end", n.replace("/", " › ") + extra)
            if n in actuales:
                lista.selection_set(i)
        self._lista_widget = lista
        f = self.fila_botones(p)
        self.boton(f, "Guardar selección", self._guardar_carpetas, lado="left")

    def _abrir_selector(self):
        M = self.M
        self._elegir_carpetas = not self._elegir_carpetas
        if self._elegir_carpetas and self._lista_carpetas is None:
            self._pendiente = None

            def trabajo():
                try:
                    self._pendiente = ("carpetas",
                                       M.listar_carpetas_correo_outlook())
                except Exception as exc:
                    M.log(f"Carpetas de correo: {exc!r}")
                    self._pendiente = ("carpetas", [])

            threading.Thread(target=trabajo, daemon=True).start()
            self.after(250, self._recoger)
        self._repintar()

    def _guardar_carpetas(self):
        M, app = self.M, self.app
        lista = getattr(self, "_lista_widget", None)
        if lista is None or not self._lista_carpetas:
            return
        elegidas = [self._lista_carpetas[i]["nombre"]
                    for i in lista.curselection()]
        M.CFG["carpetas_correo"] = [
            {"nombre": n, "buscar": n, "activo": True} for n in elegidas]
        self._tras_cambiar_carpetas()

    def _carpetas_automatico(self):
        self.M.CFG["carpetas_correo"] = []
        self._tras_cambiar_carpetas()

    def _tras_cambiar_carpetas(self):
        M, app = self.M, self.app
        M.CFG["correo_filtro"] = None
        M.CFG["correo_excluidas"] = []
        M.guardar_config()
        self._elegir_carpetas = False
        app._correo_ahora()
        self._repintar()
        # cuando llegue el correo nuevo se repinta solo (ver _descargar_correo)
        self.after(4000, lambda: self.winfo_exists() and self.pestana == "correo"
                   and self._repintar())

    def _fila_regla(self, p, r):
        M = self.M
        caja = tk.Frame(p, bg=M.C.get("boton", "#2f2f33"), padx=6, pady=4)
        caja.pack(fill="x", pady=2)
        marca = "●" if r["activa"] else "○"
        titulo = self.etiqueta(caja, f"{marca} {r['nombre']}", negrita=True)
        titulo.config(bg=M.C.get("boton", "#2f2f33"))
        titulo.pack(fill="x")
        detalle = " · ".join(x for x in (r.get("cuenta"), r.get("condiciones"),
                                         r.get("acciones")) if x)
        if detalle:
            t = self.etiqueta(caja, detalle[:140], tenue=True)
            t.config(bg=M.C.get("boton", "#2f2f33"))
            t.pack(fill="x")
        f = tk.Frame(caja, bg=M.C.get("boton", "#2f2f33"))
        f.pack(fill="x", pady=(3, 0))
        for texto, fn in (
                ("Desactivar" if r["activa"] else "Activar",
                 lambda r=r: self._accion_regla(
                     "alternar", r, activar=not r["activa"])),
                ("Ejecutar ahora", lambda r=r: self._accion_regla("ejecutar", r)),
                ("Eliminar", lambda r=r: self._borrar_regla(r))):
            b = tk.Label(f, text=texto, bg=M.C["fondo"], fg=M.C["texto"],
                         font=M.letra(7), padx=6, pady=2, cursor="hand2")
            b.pack(side="left", padx=(0, 3))
            b.bind("<Button-1>", lambda e, fn=fn: self._pulsar(fn))

    def _cargar_reglas(self):
        """Lee las reglas en un hilo (Outlook puede tardar) y avisa al acabar."""
        M = self.M
        self._reglas = None
        self._pendiente = None

        def trabajo():
            try:
                self._pendiente = ("reglas", M.listar_reglas_outlook())
            except Exception as exc:
                M.log(f"Reglas de Outlook: {exc!r}")
                self._pendiente = ("reglas", [])

        self._reglas_hilo = threading.Thread(target=trabajo, daemon=True)
        self._reglas_hilo.start()
        self.after(250, self._recoger)

    def _accion_regla(self, tipo, r, activar=None):
        M = self.M
        self._pendiente = None
        self._reglas_estado = "Trabajando…"

        def trabajo():
            try:
                if tipo == "alternar":
                    ok = M.alternar_regla(r["cuenta"], r["nombre"], activar)
                elif tipo == "ejecutar":
                    ok = M.ejecutar_regla_ahora(r["cuenta"], r["nombre"])
                else:
                    ok = M.borrar_regla(r["cuenta"], r["nombre"])
                self._pendiente = ("hecho", tipo, bool(ok), r["nombre"])
            except Exception as exc:
                M.log(f"Regla '{r['nombre']}': {exc!r}")
                self._pendiente = ("hecho", tipo, False, r["nombre"])

        threading.Thread(target=trabajo, daemon=True).start()
        self.after(250, self._recoger)

    def _borrar_regla(self, r):
        if messagebox.askyesno(
                "Eliminar regla",
                f"¿Eliminar la regla «{r['nombre']}» de Outlook?\n"
                "No se puede deshacer.", parent=self):
            self._accion_regla("borrar", r)

    def _recoger(self):
        """Recoge lo que haya dejado el hilo; si aún no, vuelve a mirar."""
        try:
            if not self.winfo_exists():
                return
        except Exception:
            return
        res = self._pendiente
        if res is None:
            self.after(250, self._recoger)
            return
        self._pendiente = None
        if res[0] == "carpetas":
            self._lista_carpetas = res[1]
        elif res[0] == "reglas":
            self._reglas = res[1]
            self._reglas_hilo = None
            self._reglas_estado = ""
        else:
            _, tipo, ok, nombre = res
            self._reglas_estado = ("Hecho: " if ok else "No se pudo: ") + nombre
            self._reglas = None
            self._reglas_hilo = None
        if self.pestana == "correo":
            self._repintar()

    # -- pestaña Eventos ---------------------------------------------------------------

    def _eventos_proximos(self, dias=30):
        M = self.M
        ahora = dt.datetime.now(M.TZ)
        limite = ahora + dt.timedelta(days=dias)
        return sorted(
            (e for e in self.app.eventos
             if not e.parte and not M.es_festivo(e.categoria)
             and e.fin >= ahora and e.inicio <= limite),
            key=lambda e: e.inicio)

    def _solapes(self, eventos):
        pares = []
        con_hora = [e for e in eventos if not e.todo_dia
                    and e.estado not in ("libre",)]
        for i, a in enumerate(con_hora):
            for b in con_hora[i + 1:]:
                if b.inicio >= a.fin:
                    break
                if a.uid == b.uid or not (a.inicio < b.fin and b.inicio < a.fin):
                    continue
                if (a.titulo.strip().lower() == b.titulo.strip().lower()
                        and a.inicio == b.inicio and a.fin == b.fin):
                    continue          # idénticos: salen en «Repetidos»
                pares.append((a, b))
        return pares

    def _duplicados(self, eventos):
        vistos, pares = {}, []
        for e in eventos:
            clave = (e.titulo.strip().lower(), e.inicio, e.fin)
            if clave in vistos and vistos[clave].uid != e.uid:
                pares.append((vistos[clave], e))
            else:
                vistos.setdefault(clave, e)
        return pares

    def _pestana_eventos(self, p):
        M, app = self.M, self.app
        proximos = self._eventos_proximos()
        solapes = self._solapes(proximos)
        duplicados = self._duplicados(proximos)

        f = self.fila_botones(p)
        for clave, texto in (("proximos", f"Próximos ({len(proximos)})"),
                             ("solapes", f"Solapes ({len(solapes)})"),
                             ("duplicados", f"Repetidos ({len(duplicados)})")):
            self.boton(f, texto,
                       lambda c=clave: (setattr(self, "_modo_eventos", c),
                                        self._repintar()),
                       activo=self._modo_eventos == clave, lado="left",
                       expandir=True)
        self.etiqueta(p, "Próximos 30 días. Pulsa un evento para ir a su día; "
                         "⋯ abre las acciones.", tenue=True).pack(
            fill="x", pady=(4, 2))

        if self._modo_eventos == "proximos":
            self._buscador_eventos(p, proximos)
        else:
            pares = solapes if self._modo_eventos == "solapes" else duplicados
            if not pares:
                self.etiqueta(p, "Nada que revisar", tenue=True).pack(fill="x")
            for a, b in pares[:20]:
                self.etiqueta(p, f"{M.fecha_corta(a.inicio)}", negrita=True
                              ).pack(fill="x", pady=(4, 0))
                self._fila_evento(p, a)
                self._fila_evento(p, b)

    def _buscador_eventos(self, p, proximos):
        M = self.M
        calendarios = sorted({e.calendario for e in proximos if e.calendario})
        if len(calendarios) > 1:
            f = self.fila_botones(p)
            self.boton(f, "Todos", lambda: self._poner_filtro(None),
                       activo=self._filtro_cal is None, lado="left")
            for nombre in calendarios[:4]:
                self.boton(f, nombre[:14],
                           lambda n=nombre: self._poner_filtro(n),
                           activo=self._filtro_cal == nombre, lado="left")
        self._var_busca = tk.StringVar(value=self._texto_busqueda)
        e = tk.Entry(p, textvariable=self._var_busca,
                     bg=M.C.get("boton", "#2f2f33"), fg=M.C["texto"],
                     insertbackground=M.C["texto"], relief="flat",
                     font=M.letra(8))
        e.pack(fill="x", ipady=3, pady=(2, 4))
        self._lista_eventos = tk.Frame(p, bg=M.C["fondo"])
        self._lista_eventos.pack(fill="x")
        e.bind("<KeyRelease>", lambda ev: self._filtrar_eventos(proximos))
        self._filtrar_eventos(proximos)

    def _poner_filtro(self, nombre):
        self._filtro_cal = nombre
        self._repintar()

    def _filtrar_eventos(self, proximos):
        M = self.M
        self._texto_busqueda = self._var_busca.get()
        for hijo in self._lista_eventos.winfo_children():
            hijo.destroy()
        lista = proximos
        if self._filtro_cal:
            lista = [e for e in lista if e.calendario == self._filtro_cal]
        if self._texto_busqueda.strip():
            aguja = M.sin_acentos(self._texto_busqueda).strip()
            lista = [e for e in lista if aguja in M.sin_acentos(
                " ".join((e.titulo, e.lugar or "", e.calendario or "")))]
        if not lista:
            self.etiqueta(self._lista_eventos, "Sin resultados",
                          tenue=True).pack(fill="x")
        for ev in lista[:25]:
            self._fila_evento(self._lista_eventos, ev)
        if len(lista) > 25:
            self.etiqueta(self._lista_eventos,
                          f"… y {len(lista) - 25} más (usa el buscador)",
                          tenue=True).pack(fill="x")

    def _fila_evento(self, padre, ev):
        M = self.M
        fila = tk.Frame(padre, bg=M.C["fondo"])
        fila.pack(fill="x", pady=1)
        hora = "todo el día" if ev.todo_dia else f"{ev.inicio:%H:%M}"
        texto = f"{M.fecha_corta(ev.inicio)}  {hora}  {ev.titulo}"
        t = tk.Label(fila, text=texto[:44], bg=M.C["fondo"], fg=M.C["texto"],
                     font=M.letra(8), anchor="w", cursor="hand2")
        t.pack(side="left", fill="x", expand=True)
        t.bind("<Button-1>", lambda e, ev=ev: self._pulsar(
            lambda: self.app.ir_a_fecha(ev.inicio.date())))
        mas = tk.Label(fila, text="⋯", bg=M.C["fondo"], fg=M.C["texto_tenue"],
                       font=M.letra(10, "bold"), cursor="hand2", padx=4)
        mas.pack(side="right")
        mas.bind("<Button-1>", lambda e, ev=ev: self._menu_evento(e, ev))

    def _menu_evento(self, e, ev):
        M, app = self.M, self.app
        m = tk.Menu(self, tearoff=0)
        m.add_command(label="Ir a su día",
                      command=lambda: app.ir_a_fecha(ev.inicio.date()))
        m.add_command(label="Editar…", command=lambda: app._editar_evento(ev))
        m.add_command(label="Abrir en Outlook",
                      command=lambda: M.abrir_cita_en_outlook(ev))
        m.add_separator()
        m.add_command(label="Posponer 1 hora",
                      command=lambda: self._tras(app._posponer, ev, minutos=60))
        m.add_command(label="Pasar a mañana",
                      command=lambda: self._tras(app._posponer, ev, dias=1))
        m.add_separator()
        m.add_command(label="Eliminar (a Elementos eliminados)…",
                      command=lambda: self._eliminar_evento(ev))
        m.tk_popup(e.x_root, e.y_root)

    def _tras(self, fn, ev, **kw):
        fn(ev, **kw)
        self.after(600, self._repintar)

    def _eliminar_evento(self, ev):
        M = self.M
        if not messagebox.askyesno(
                "Eliminar evento",
                f"¿Mandar «{ev.titulo}» a Elementos eliminados de Outlook?",
                parent=self):
            return
        ok, motivo = M.borrar_en_outlook(ev)
        if ok:
            self.app.refrescar()
        else:
            messagebox.showinfo("No se ha podido", motivo or "Error",
                                parent=self)

    # -- pestaña Sistema ---------------------------------------------------------------

    # -- pestaña Cuenta atrás ------------------------------------------------------

    def _pestana_cuenta(self, p):
        M = self.M
        cuentas = M.CFG.setdefault("cuentas_atras", [])
        self.etiqueta(p, "Pon un evento en un día y verás cuántos faltan. "
                         "Rojo 0-5 · ámbar 6-10 · verde más de 10.",
                      tenue=True).pack(fill="x", pady=(0, 4))
        self._var_cu_nombre = tk.StringVar()
        self._var_cu_fecha = tk.StringVar()
        for var, ancho, aviso in ((self._var_cu_nombre, 24, "Nombre del evento"),
                                  (self._var_cu_fecha, 12, "Día (DD/MM/AAAA)")):
            self.etiqueta(p, aviso, tenue=True).pack(fill="x")
            e = tk.Entry(p, textvariable=var, bg=M.C.get("boton", "#2f2f33"),
                         fg=M.C["texto"], insertbackground=M.C["texto"],
                         relief="flat", font=M.letra(8), width=ancho)
            e.pack(fill="x", pady=(0, 3), ipady=3)
            e.bind("<Return>", lambda ev: self._anadir_cuenta())
        f = self.fila_botones(p)
        self.boton(f, "Añadir", self._anadir_cuenta, lado="left")
        self.mensaje_cuenta = self.etiqueta(p, "", tenue=True)
        self.mensaje_cuenta.pack(fill="x")

        self.seccion(p, "Cuenta atrás")
        if not cuentas:
            self.etiqueta(p, "Aún no hay eventos", tenue=True).pack(fill="x")
        hoy = dt.date.today()
        for c in sorted(cuentas, key=lambda c: c["fecha"]):
            n = dias_hasta(dt.date.fromisoformat(c["fecha"]), hoy)
            texto = ("hoy" if n == 0 else f"{n} d" if n > 0
                     else f"hace {-n} d")
            fila = tk.Frame(p, bg=M.C["fondo"])
            fila.pack(fill="x", pady=1)
            tk.Label(fila, text=texto, bg=color_cuenta(n), fg="#101010",
                     font=M.letra(8, "bold"), width=8).pack(side="left")
            tk.Label(fila, text=f"{c['nombre']} · {c['fecha']}",
                     bg=M.C["fondo"], fg=M.C["texto"], font=M.letra(8),
                     anchor="w").pack(side="left", padx=6, fill="x",
                                      expand=True)
            borrar = tk.Label(fila, text="✕", bg=M.C["fondo"],
                              fg=M.C["texto_tenue"], font=M.letra(9),
                              cursor="hand2")
            borrar.pack(side="right")
            borrar.bind("<Button-1>", lambda e, c=c: self._borrar_cuenta(c))

    def _anadir_cuenta(self):
        nombre = self._var_cu_nombre.get().strip()
        try:
            fecha = leer_fecha(self._var_cu_fecha.get())
        except ValueError:
            self.mensaje_cuenta.config(text="Día no válido (DD/MM/AAAA)")
            return
        if not nombre:
            self.mensaje_cuenta.config(text="Falta el nombre")
            return
        self.M.CFG.setdefault("cuentas_atras", []).append(
            {"nombre": nombre, "fecha": fecha.isoformat()})
        self.M.guardar_config()
        self.app._rehacer_barra()
        self._repintar()

    def _borrar_cuenta(self, c):
        self.M.CFG["cuentas_atras"].remove(c)
        self.M.guardar_config()
        self.app._rehacer_barra()
        self._repintar()

    def _pestana_sistema(self, p):
        M, app = self.M, self.app
        self.casilla(p, "Arrancar con Windows", lambda: M.arranque_activo(),
                     lambda: app._alternar_arranque())

        self.seccion(p, "Perfiles")
        perfiles = M.listar_perfiles()
        if not perfiles:
            self.etiqueta(p, "Sin perfiles guardados (clic derecho > Perfil)",
                          tenue=True).pack(fill="x")
        f = self.fila_botones(p)
        for nombre in perfiles[:6]:
            self.boton(f, nombre[:14],
                       self._y_dibujar(lambda n=nombre: app._cambiar_perfil(n)),
                       activo=(M.CFG.get("perfil") or "") == nombre,
                       lado="left")

        self.seccion(p, "Dónde se abre cada panel")
        self.etiqueta(p, "Ventana aparte, o anclado a un lado del widget.",
                      tenue=True).pack(fill="x")
        for clave, nombre, atributo in PANELES_UBICABLES:
            fila = tk.Frame(p, bg=M.C["fondo"])
            fila.pack(fill="x", pady=1)
            tk.Label(fila, text=nombre, bg=M.C["fondo"], fg=M.C["texto"],
                     font=M.letra(8), anchor="w", width=17).pack(side="left")
            actual = M.ubicacion_panel(clave)
            for valor, texto in (("flotante", "Aparte"), ("izquierda", "◀"),
                                 ("derecha", "▶")):
                self.boton(fila, texto,
                           lambda c=clave, v=valor, a=atributo:
                           self._poner_ubicacion(c, v, a),
                           activo=actual == valor, lado="left")

        self.seccion(p, "Post-It")
        self.casilla(p, "Unir las notas de Post-It con el Calendario",
                     lambda: M.CFG.get("sincronizar_postit", True),
                     lambda: app._alternar_sync_postit())
        self.etiqueta(p, "Carpeta: " + (app._carpeta_compartida() or "ninguna"),
                      tenue=True).pack(fill="x")
        f = self.fila_botones(p)
        self.boton(f, "Cambiar carpeta compartida",
                   self._y_dibujar(app._elegir_carpeta_sync),
                   lado="left")

        self.seccion(p, "Copias y carpetas")
        f = self.fila_botones(p)
        self.boton(f, "Copia de la configuración", lambda: app._copia_config(),
                   lado="left")
        f = self.fila_botones(p)
        self.boton(f, "Carpeta de datos", lambda: abrir_ruta(M.DATOS),
                   lado="left")
        self.boton(f, "Ver el registro", lambda: abrir_ruta(M.LOG_PATH),
                   lado="left")
        self.etiqueta(p, f"Calendario {M.VERSION}", tenue=True).pack(
            fill="x", pady=(8, 0))

    def _poner_ubicacion(self, clave, valor, atributo):
        M = self.M
        panel = getattr(self.app, atributo, None)
        try:
            abierto = panel is not None and panel.winfo_exists()
        except Exception:
            abierto = False
        if abierto:
            M.fijar_ubicacion_panel(clave, valor, panel)
        else:
            M.CFG.setdefault("ubicacion_vistas", {})[clave] = valor
            M.guardar_config()
        self._repintar()



# --------------------------------------------------------------------------
# Paleta de comandos: Ctrl+K, escribe y pulsa Intro
# --------------------------------------------------------------------------

def _cc():
    import correo_columna
    return correo_columna


def lista_de_comandos(app, M):
    """[(título, palabras clave, función)]. Todo lo que se puede hacer."""
    c = []

    def add(titulo, palabras, fn):
        c.append((titulo, palabras, fn))

    add("Ir a hoy", "hoy actual volver", app._ir_a_hoy)
    add("Actualizar ahora", "refrescar recargar descargar",
        lambda: app.refrescar())
    add("Nuevo evento", "crear añadir cita", lambda: app._nuevo_evento())
    add("Buscar evento o ir a una fecha", "buscar fecha saltar",
        lambda: app._abrir_panel_ir())

    for clave, nombre in (("rejilla", "Calendario"), ("ambas", "Calendario + lista"),
                          ("lista", "Lista"), ("mes", "Mes")):
        add(f"Vista: {nombre}", "cambiar ver vista",
            lambda k=clave: app._cambiar_vista(k))
        add(f"Abrir {nombre} en ventana aparte", "vista ventana flotante",
            lambda k=clave: app.abrir_vista_aparte(k))
    for n in (1, 3, 5, 7, 14, 21):
        add(f"Ver {n} {'día' if n == 1 else 'días'}", "dias semana rango",
            lambda n=n: app._cambiar_dias(n))
    add("Orientación horizontal", "horizontal vertical girar",
        lambda: app._cambiar_orientacion("horizontal"))
    add("Orientación vertical", "horizontal vertical girar",
        lambda: app._cambiar_orientacion("vertical"))
    for clave, nombre in M.NIVELES.items():
        add(f"Ventana: {nombre}", "nivel encima fondo normal siempre visible",
            lambda k=clave: app._cambiar_nivel(k))
    for clave, nombre in (("auto", "como Windows"), ("claro", "claro"),
                          ("oscuro", "oscuro")):
        add(f"Tema {nombre}", "tema color aspecto",
            lambda k=clave: app._cambiar_tema(k))

    add("Abrir hábitos", "habitos rachas", lambda: app._abrir_habitos())
    add("Abrir correo (bandeja)", "correo mail email bandeja",
        lambda: app._abrir_correo(modo="bandeja"))
    add("Abrir reglas de correo", "correo reglas outlook filtros",
        lambda: app._abrir_correo(modo="reglas"))
    add("Correo en el widget: mostrar u ocultar", "correo columna widget mail",
        lambda: app._abrir_correo_columna())
    add("Correo en el widget: solo no leídos", "correo columna sin leer filtro",
        lambda: _cc().alternar_sin_leer(app, M))
    add("Correo en el widget: todas las carpetas", "correo columna filtros quitar",
        lambda: _cc().limpiar_filtros(app, M))
    for f in app.correo_carpetas:
        nombre = f["nombre"]
        add(f"Correo: ver solo {nombre}", "correo carpeta filtrar",
            lambda n=nombre: _cc().poner_filtro(app, M, n))
        add(f"Correo: excluir o incluir {nombre}", "correo carpeta excluir ocultar",
            lambda n=nombre: _cc().alternar_exclusion(app, M, n))
    add("Abrir cumpleaños", "cumpleaños fiestas", lambda: app._abrir_cumpleanos())
    add("Abrir plantillas", "plantillas modelos",
        lambda: app._abrir_plantillas("plantillas"))
    add("Abrir tareas permanentes", "tareas pendientes",
        lambda: app._abrir_plantillas("tareas"))
    add("Mostrar u ocultar to-dos", "todos tareas", lambda: app._abrir_todos())
    add("Abrir la barra de vistas", "vistas accesos", lambda: app._abrir_accesos())

    add("Panel de control", "ajustes configuracion opciones preferencias",
        lambda: app._abrir_panel_control())
    for clave, nombre in PESTANAS:
        add(f"Panel de control: {nombre}", "ajustes configuracion opciones",
            lambda k=clave: app._abrir_panel_control(pestana=k, forzar=True))

    add("Activar o desactivar los avisos", "avisos campana notificaciones silenciar",
        lambda: app._alternar_notificaciones())
    add("Ver un aviso de ejemplo", "probar aviso", lambda: app._probar_aviso())
    add("Arrancar con Windows: activar o desactivar", "inicio arranque autoarranque",
        lambda: app._alternar_arranque())
    add("Copia de seguridad de la configuración", "copia backup guardar",
        lambda: app._copia_config())
    add("Abrir el asistente de configuración", "asistente configurar outlook ics",
        lambda: app._abrir_asistente())
    for nombre in M.listar_perfiles():
        add(f"Cambiar al perfil {nombre}", "perfil cuenta usuario",
            lambda n=nombre: app._cambiar_perfil(n))
    add("Abrir la carpeta de datos", "carpeta datos configuracion",
        lambda: abrir_ruta(M.DATOS))
    add("Ver el registro", "registro log errores",
        lambda: abrir_ruta(M.LOG_PATH))
    add("Ocultar el calendario", "ocultar esconder minimizar",
        lambda: app._alternar_visible(False))
    add("Salir del calendario", "cerrar salir terminar", lambda: app._salir())
    return c


class PaletaComandos(PanelBase):
    CLAVE = "paleta"
    ATRIBUTO = "_paleta"
    TITULO = "⌘ Comandos"
    ANCHO = 420
    FILAS = 9

    def __init__(self, app, M):
        super().__init__(app, M)
        self.comandos = lista_de_comandos(app, M)
        self.resultados = []          # [(texto a mostrar, función)]
        self.seleccion = 0
        cab = self.cabecera(self.marco)
        M_ = M
        self.entrada = tk.Entry(self.marco, bg=M_.C.get("boton", "#2f2f33"),
                                fg=M_.C["texto"], insertbackground=M_.C["texto"],
                                relief="flat", font=M_.letra(10))
        self.entrada.pack(fill="x", ipady=5, pady=(6, 2))
        self.pista = self.etiqueta(
            self.marco, "Escribe una acción, un evento o una fecha "
                        "(18/9, mañana)", tenue=True, ancho=self.ANCHO - 24)
        self.pista.pack(fill="x", pady=(0, 4))
        self.lista = tk.Frame(self.marco, bg=M_.C["fondo"])
        self.lista.pack(fill="x")
        self.entrada.bind("<KeyRelease>", self._al_escribir)
        self.entrada.bind("<Down>", lambda e: self._mover(1))
        self.entrada.bind("<Up>", lambda e: self._mover(-1))
        self.entrada.bind("<Return>", lambda e: self._ejecutar())
        self.entrada.bind("<Escape>", lambda e: self.destroy())
        self._buscar("")
        self.actualizar_posicion()
        self.entrada.focus_force()

    # -- búsqueda --------------------------------------------------------------

    def _buscar(self, texto):
        M = self.M
        aguja = M.sin_acentos(texto).strip()
        res = []
        if aguja:
            fecha = M.interpretar_fecha(texto, dt.datetime.now(M.TZ).date())
            if fecha:
                res.append((f"📅  Ir al {fecha:%d/%m/%Y}",
                            lambda f=fecha: self.app.ir_a_fecha(f)))
        palabras = aguja.split()
        puntuados = []
        for i, (titulo, claves, fn) in enumerate(self.comandos):
            pajar = M.sin_acentos(titulo + " " + claves)
            if all(p in pajar for p in palabras):
                # los que empiezan por lo escrito, primero; sin escribir nada,
                # se respeta el orden de la lista
                orden = 0 if (not aguja or
                              M.sin_acentos(titulo).startswith(aguja)) else 1
                puntuados.append((orden, len(titulo) if aguja else i,
                                  titulo, fn))
        puntuados.sort(key=lambda x: x[:2])
        limite = self.FILAS if aguja else 7      # sin escribir, solo los primeros
        for _, _, titulo, fn in puntuados[:limite]:
            res.append((f"›  {titulo}", fn))
        if aguja:
            for ev in M.buscar_eventos(self.app.eventos, texto, tope=5):
                hora = "todo el día" if ev.todo_dia else f"{ev.inicio:%H:%M}"
                res.append((f"🗓  {M.fecha_corta(ev.inicio)} {hora}  {ev.titulo[:40]}",
                            lambda ev=ev: self.app.ir_a_fecha(ev.inicio.date())))
        self.resultados = res[:self.FILAS + 5]
        self.seleccion = 0
        self._pintar_resultados()
        if aguja and not self.resultados:
            self.pista.config(text="Sin coincidencias")
        else:
            self.pista.config(text="Intro para ejecutar · ↑ ↓ para moverte · "
                                   "Esc para cerrar")

    def _al_escribir(self, e):
        if e.keysym in ("Up", "Down", "Return", "Escape"):
            return
        self._buscar(self.entrada.get())

    def _pintar_resultados(self):
        M = self.M
        for hijo in self.lista.winfo_children():
            hijo.destroy()
        for i, (texto, fn) in enumerate(self.resultados):
            marcado = i == self.seleccion
            fila = tk.Label(
                self.lista, text=texto, anchor="w", padx=6, pady=3,
                bg=M.C["evento"] if marcado else M.C["fondo"],
                fg=M.C["evento_texto"] if marcado else M.C["texto"],
                font=M.letra(9), cursor="hand2")
            fila.pack(fill="x")
            fila.bind("<Button-1>", lambda e, i=i: self._ejecutar(i))
        self.update_idletasks()
        self.actualizar_posicion_alto()

    def actualizar_posicion_alto(self):
        try:
            self.geometry(f"{self.ANCHO}x{self.winfo_reqheight()}")
        except Exception:
            pass

    def _mover(self, paso):
        if self.resultados:
            self.seleccion = (self.seleccion + paso) % len(self.resultados)
            self._pintar_resultados()
        return "break"

    def _ejecutar(self, indice=None):
        i = self.seleccion if indice is None else indice
        if not self.resultados:
            return "break"
        _, fn = self.resultados[i]
        app, M = self.app, self.M
        self.destroy()
        try:
            app.after(60, lambda: self._lanzar(fn))
        except Exception:
            self._lanzar(fn)
        return "break"

    def _lanzar(self, fn):
        try:
            fn()
        except Exception as exc:
            self.M.log(f"Paleta de comandos: {exc!r}")
            try:
                self.M.AvisoFlotante(self.app, "No se ha podido",
                                     str(exc)[:80], self.M.C["ahora"], 6)
            except Exception:
                pass

    # -- colocación: centrada arriba si es flotante, o anclada -------------------------

    def actualizar_posicion(self):
        M = self.M
        try:
            self.update_idletasks()
            app = self.app
            alto = self.winfo_reqheight()
            if M.ubicacion_panel(self.CLAVE) == "flotante":
                if self._colocado_ya:
                    return
                x = app.winfo_rootx() + max(0, (app.winfo_width() - self.ANCHO) // 2)
                y = app.winfo_rooty() + M.BARRA + 6
            else:
                x = M.x_panel_acoplado(app, self.CLAVE, self.ANCHO)
                y = app.winfo_rooty()
            self.geometry(f"{self.ANCHO}x{alto}"
                          + M.encajar_en_pantalla(app, x, y, self.ANCHO, alto))
            self._colocado_ya = True
        except Exception:
            pass
