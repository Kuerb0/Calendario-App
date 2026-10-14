"""Columna de correo dentro del widget, junto a la lista y los to-dos.

Enseña los correos más recientes de todas las carpetas mezclados, con chips
para ver solo una carpeta y una lista de carpetas excluidas. Los datos son los
mismos que usa el panel de correo (app.correo_carpetas / app.correo_mensajes),
así que no descarga nada por su cuenta.

Ajustes en CFG:
    correo_visible       la columna está a la vista
    ancho_correo         ancho en píxeles a escala 100 %
    correo_filtro        nombre de la carpeta que se ve sola (None = todas)
    correo_excluidas     carpetas que no salen en la columna
    correo_solo_sin_leer solo los no leídos
    correo_filas_max     filas que hacen crecer la ventana (el resto, si hay sitio)

Este módulo no importa calendario.py: recibe el módulo ya cargado como 'M' y la
ventana principal como 'app'.
"""
import datetime as dt
import threading
import tkinter as tk

ETIQUETA = "correo_col"          # todo lo dibujado lleva esta etiqueta


# --------------------------------------------------------------------------
# Lógica pura (se puede probar sin ventana)
# --------------------------------------------------------------------------

def filtrar_mensajes(por_carpeta, filtro=None, excluidas=(),
                     solo_sin_leer=False):
    """Los correos de todas las carpetas mezclados, del más nuevo al más viejo.

    'filtro' deja solo esa carpeta; 'excluidas' quita carpetas; una carpeta
    excluida no sale aunque sea la del filtro.
    """
    excluidas = set(excluidas or ())
    salida = []
    for nombre, mensajes in por_carpeta.items():
        if nombre in excluidas:
            continue
        if filtro and nombre != filtro:
            continue
        for m in mensajes:
            if solo_sin_leer and m.leido:
                continue
            salida.append(m)
    salida.sort(key=lambda m: m.fecha, reverse=True)
    return salida


def etiqueta_corta(nombre, todos=()):
    """Nombre corto para chips y menús: solo la carpeta, sin la cuenta.

    «finanzas@x.com/Bandeja de entrada» -> «Bandeja de entrada». Si dos
    carpetas acabarían con el mismo nombre, se les añade la cuenta.
    """
    nombre = str(nombre)
    if "/" not in nombre:
        return nombre
    cuenta, _, resto = nombre.partition("/")
    ultimo = nombre.rsplit("/", 1)[1]
    repetidas = sum(1 for n in todos
                    if str(n).rsplit("/", 1)[-1] == ultimo)
    if repetidas > 1:
        return f"{ultimo} ({cuenta.split('@')[0][:10]})"
    return ultimo


def chips_a_mostrar(visibles, filtro=None, maximo=3):
    """Las carpetas que llevan botón propio: las que tienen correo sin leer
    (las que más tienen, primero), la del filtro actual y, si sobra sitio,
    las primeras. El resto va al menú «▾ más» para que los filtros no se
    coman la columna."""
    maximo = max(0, int(maximo))
    orden = sorted(range(len(visibles)),
                   key=lambda i: (-visibles[i].get("no_leidos", 0), i))
    elegidos = [i for i in orden[:maximo]]
    for i, f in enumerate(visibles):
        if filtro and f["nombre"] == filtro and i not in elegidos:
            if elegidos and len(elegidos) >= maximo:
                elegidos.pop()          # el último de los elegidos cede sitio
            elegidos.append(i)
    return [visibles[i] for i in sorted(elegidos)]


def carpetas_visibles(info, excluidas=()):
    """Las carpetas (dicts de app.correo_carpetas) que no están excluidas."""
    fuera = set(excluidas or ())
    return [c for c in info if c["nombre"] not in fuera]


def filtro_valido(info, filtro, excluidas=()):
    """El filtro tal cual, o None si esa carpeta ya no existe o está excluida."""
    nombres = {c["nombre"] for c in carpetas_visibles(info, excluidas)}
    return filtro if filtro in nombres else None


def estado(app, M):
    """(filtro, excluidas, solo_sin_leer) de ahora mismo."""
    cfg = M.CFG
    fuera = [n for n in (cfg.get("correo_excluidas") or [])
             if isinstance(n, str)]
    filtro = filtro_valido(app.correo_carpetas, cfg.get("correo_filtro"), fuera)
    return filtro, fuera, bool(cfg.get("correo_solo_sin_leer", False))


def mensajes_de_la_columna(app, M):
    filtro, fuera, solo = estado(app, M)
    return filtrar_mensajes(app.correo_mensajes, filtro, fuera, solo)


# --------------------------------------------------------------------------
# Cambios de estado (guardan y repintan)
# --------------------------------------------------------------------------

def _guardar_y_pintar(app, M):
    M.guardar_config()
    app.dibujar()
    for atributo in ("_panel_control",):     # el panel de control, si está abierto
        panel = getattr(app, atributo, None)
        try:
            if panel is not None and panel.winfo_exists() and \
                    panel.pestana == "correo":
                panel._repintar()
        except Exception:
            pass


def poner_filtro(app, M, nombre):
    """Ver solo una carpeta (o todas con None). Pulsar la activa la suelta."""
    if nombre and M.CFG.get("correo_filtro") == nombre:
        nombre = None
    M.CFG["correo_filtro"] = nombre
    if nombre:
        fuera = [n for n in (M.CFG.get("correo_excluidas") or []) if n != nombre]
        M.CFG["correo_excluidas"] = fuera
    _guardar_y_pintar(app, M)


def alternar_exclusion(app, M, nombre):
    fuera = list(M.CFG.get("correo_excluidas") or [])
    if nombre in fuera:
        fuera.remove(nombre)
    else:
        fuera.append(nombre)
        if M.CFG.get("correo_filtro") == nombre:
            M.CFG["correo_filtro"] = None
    M.CFG["correo_excluidas"] = fuera
    _guardar_y_pintar(app, M)


def alternar_sin_leer(app, M):
    M.CFG["correo_solo_sin_leer"] = not M.CFG.get("correo_solo_sin_leer", False)
    _guardar_y_pintar(app, M)


def limpiar_filtros(app, M):
    M.CFG["correo_filtro"] = None
    M.CFG["correo_excluidas"] = []
    M.CFG["correo_solo_sin_leer"] = False
    _guardar_y_pintar(app, M)


def alternar_columna(app, M):
    """Enseña u oculta la columna. Al enseñarla pide el correo si no hay."""
    M.CFG["correo_visible"] = not M.CFG.get("correo_visible", False)
    M.guardar_config()
    if M.CFG["correo_visible"] and not app.correo_carpetas \
            and not app.correo_error:
        app._correo_ahora()
    app._tam = None
    app.dibujar()
    return "break"


def _poner_chips(app, M, n):
    M.CFG["correo_chips_max"] = int(n)
    _guardar_y_pintar(app, M)


def cambiar_ancho(app, M, px):
    M.CFG["ancho_correo"] = max(160, min(700, int(px)))
    M.guardar_config()
    app._tam = None
    app.dibujar()


# --------------------------------------------------------------------------
# Acciones sobre un mensaje
# --------------------------------------------------------------------------

def alternar_leido(app, M, m):
    nuevo = not m.leido
    if not M.marcar_correo(m.id_outlook, m.almacen_outlook, leido=nuevo):
        return
    m.leido = nuevo
    for c in app.correo_carpetas:
        if c["nombre"] == m.carpeta:
            c["no_leidos"] = max(0, c["no_leidos"] + (-1 if nuevo else 1))
    app.dibujar()


def abrir_mensaje(app, M, m):
    M.VentanaCorreo(app, m, None)


# -- control del correo: acciones sobre un mensaje ----------------------------

def _repintar_todo(app):
    """Redibuja la columna y, si está abierto, el panel de correo aparte."""
    app.dibujar()
    panel = getattr(app, "_panel_correo", None)
    try:
        if panel is not None and panel.winfo_exists():
            panel.pintar()
            panel.actualizar_posicion()
    except Exception:
        pass


def _aviso(app, M, titulo, texto, ok=True):
    try:
        M.AvisoFlotante(app, titulo, texto, "#2ecc71" if ok else "#ff453a", 4)
    except Exception:
        pass


def _en_hilo(app, trabajo, despues):
    """Outlook puede tardar: se le pide en un hilo y se avisa al acabar."""
    def correr():
        try:
            res = trabajo()
        except Exception as exc:                       # pragma: no cover
            res = (False, str(exc)[:60])
        try:
            app.after(0, lambda: despues(res))
        except Exception:
            pass
    threading.Thread(target=correr, daemon=True).start()


def _quitar_local(app, m):
    """Quita el mensaje de lo que se ve, sin esperar a Outlook."""
    lista = app.correo_mensajes.get(m.carpeta, [])
    if m in lista:
        lista.remove(m)
    for c in app.correo_carpetas:
        if c["nombre"] == m.carpeta:
            c["total"] = max(0, c.get("total", 0) - 1)
            if not m.leido:
                c["no_leidos"] = max(0, c.get("no_leidos", 0) - 1)


def mover_mensaje(app, M, m, destino=None, borrar=False):
    """Archiva (destino None), mueve a 'destino' o elimina el mensaje."""
    _quitar_local(app, m)
    _repintar_todo(app)

    def despues(res):
        ok, texto = res
        _aviso(app, M, "✉ Correo" if ok else "✉ No se pudo", texto, ok)
        if not ok:
            app._correo_ahora()          # vuelve a como estaba en Outlook
    _en_hilo(app, lambda: M.mover_correo(
        m.id_outlook, m.almacen_outlook, destino, borrar), despues)


def alternar_bandera(app, M, m):
    nuevo = not getattr(m, "marcado", False)
    m.marcado = nuevo
    _repintar_todo(app)
    _en_hilo(app, lambda: (M.marcar_bandera_correo(
        m.id_outlook, m.almacen_outlook, nuevo), ""),
        lambda res: None if res[0] else (
            setattr(m, "marcado", not nuevo), _repintar_todo(app)))


def crear_evento_desde(app, M, m):
    inicio = M.redondear_media_hora(dt.datetime.now(M.TZ))
    app._ficha(plantilla={
        "titulo": m.asunto, "inicio": inicio,
        "fin": inicio + dt.timedelta(hours=1),
        "notas": f"Correo de {m.remitente} ({m.fecha:%d/%m/%Y %H:%M})"})


def crear_todo_desde(app, M, m):
    app._todo_añadir(f"{m.asunto} ({m.remitente})")
    _aviso(app, M, "📋 To-do creado", m.asunto[:60])


def menu_mensaje(app, M, e, m):
    menu = tk.Menu(app, tearoff=0)
    menu.add_command(label="Leer", command=lambda: abrir_mensaje(app, M, m))
    menu.add_command(
        label="Marcar como no leído" if m.leido else "Marcar como leído",
        command=lambda: alternar_leido(app, M, m))
    menu.add_command(
        label=("Quitar la bandera" if getattr(m, "marcado", False)
               else "Marcar con bandera"),
        command=lambda: alternar_bandera(app, M, m))
    menu.add_separator()
    responder = tk.Menu(menu, tearoff=0)
    for etiqueta, modo in (("Responder", "responder"),
                           ("Responder a todos", "responder_todos"),
                           ("Reenviar", "reenviar")):
        responder.add_command(
            label=etiqueta,
            command=lambda modo=modo: M.responder_o_reenviar_correo(
                m.id_outlook, m.almacen_outlook, modo))
    menu.add_cascade(label="Responder o reenviar", menu=responder)
    menu.add_separator()
    menu.add_command(label="Archivar",
                     command=lambda: mover_mensaje(app, M, m))
    mover = tk.Menu(menu, tearoff=0)
    todos = [c["nombre"] for c in app.correo_carpetas]
    for c in app.correo_carpetas:
        if c["nombre"] != m.carpeta:
            mover.add_command(
                label=etiqueta_corta(c["nombre"], todos),
                command=lambda d=c["nombre"]: mover_mensaje(app, M, m, d))
    if len(todos) > 1:
        mover.add_separator()
    mover.add_command(label="Otra carpeta…",
                      command=lambda: SelectorCarpetas(
                          app, M, "Mover a…",
                          lambda d: mover_mensaje(app, M, m, d)))
    menu.add_cascade(label="Mover a", menu=mover)
    menu.add_command(label="Eliminar",
                     command=lambda: mover_mensaje(app, M, m, borrar=True))
    menu.add_separator()
    menu.add_command(label="Crear un evento con este correo…",
                     command=lambda: crear_evento_desde(app, M, m))
    menu.add_command(label="Añadir a mis to-dos",
                     command=lambda: crear_todo_desde(app, M, m))
    menu.add_separator()
    menu.add_command(label="Abrir en Outlook",
                     command=lambda: M.abrir_correo_en_outlook(
                         m.id_outlook, m.almacen_outlook))
    corto = etiqueta_corta(m.carpeta, [c["nombre"] for c in app.correo_carpetas])
    carpeta = tk.Menu(menu, tearoff=0)
    carpeta.add_command(label="Ver solo esta carpeta",
                        command=lambda: poner_filtro(app, M, m.carpeta))
    carpeta.add_command(label="Ocultarla de la columna",
                        command=lambda: alternar_exclusion(app, M, m.carpeta))
    carpeta.add_command(label="Dejar de leerla",
                        command=lambda: quitar_carpeta_leida(app, M, m.carpeta))
    carpeta.add_separator()
    carpeta.add_command(label="Elegir carpetas…",
                        command=lambda: abrir_selector_carpetas(app, M))
    menu.add_cascade(label=f"Carpeta «{corto}»", menu=carpeta)
    menu.tk_popup(e.x_root, e.y_root)


# -- qué carpetas de Outlook se leen ------------------------------------------

def carpetas_leidas(app, M):
    """Los nombres de las carpetas que se están leyendo ahora."""
    conf = [c.get("buscar") for c in (M.CFG.get("carpetas_correo") or [])
            if c.get("activo", True)]
    return set(conf) if conf else {c["nombre"] for c in app.correo_carpetas}


def fijar_carpetas_leidas(app, M, nombres):
    nombres = list(dict.fromkeys(nombres))
    M.CFG["carpetas_correo"] = [
        {"nombre": n, "buscar": n, "activo": True} for n in nombres]
    M.CFG["correo_filtro"] = None
    M.CFG["correo_excluidas"] = [
        n for n in (M.CFG.get("correo_excluidas") or []) if n in nombres]
    M.guardar_config()
    app._correo_ahora()


def quitar_carpeta_leida(app, M, nombre):
    fijar_carpetas_leidas(app, M, [n for n in carpetas_leidas(app, M)
                                   if n != nombre])


def abrir_selector_carpetas(app, M):
    SelectorCarpetas(app, M, "Carpetas de correo que se leen", None)


class SelectorCarpetas(tk.Toplevel):
    """Lista de todas las carpetas de Outlook con buscador.

    Sin 'alelegir', cada clic pone o quita una carpeta de las que se leen
    (se aplica al momento). Con 'alelegir', un doble clic elige una carpeta
    y se la pasa a esa función (por ejemplo, para mover un correo)."""

    def __init__(self, app, M, titulo, alelegir):
        super().__init__(app)
        self.app, self.M, self.alelegir = app, M, alelegir
        self.title(titulo)
        self.configure(bg=M.C["fondo"])
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass
        self.transient(app)
        self.geometry(f"320x420+{app.winfo_rootx() + 30}"
                      f"+{app.winfo_rooty() + 40}")
        tk.Label(self, text=("Doble clic en la carpeta de destino" if alelegir
                             else "Clic en una carpeta para leerla o dejar de "
                                  "leerla"),
                 bg=M.C["fondo"], fg=M.C["texto_tenue"], font=M.letra(8),
                 anchor="w").pack(fill="x", padx=8, pady=(8, 2))
        self.filtro = tk.StringVar()
        entrada = tk.Entry(self, textvariable=self.filtro,
                           bg=M.C.get("boton", "#2f2f33"), fg=M.C["texto"],
                           insertbackground=M.C["texto"], relief="flat",
                           font=M.letra(9))
        entrada.pack(fill="x", padx=8, pady=2, ipady=4)
        entrada.focus_set()
        self.filtro.trace_add("write", lambda *_: self._pintar())
        marco = tk.Frame(self, bg=M.C["fondo"])
        marco.pack(fill="both", expand=True, padx=8, pady=4)
        barra = tk.Scrollbar(marco)
        barra.pack(side="right", fill="y")
        self.lista = tk.Listbox(
            marco, activestyle="none", exportselection=False,
            bg=M.C.get("boton", "#2f2f33"), fg=M.C["texto"],
            selectbackground=M.C["evento"], selectforeground=M.C["evento_texto"],
            relief="flat", highlightthickness=0, font=M.letra(9),
            yscrollcommand=barra.set)
        barra.config(command=self.lista.yview)
        self.lista.pack(side="left", fill="both", expand=True)
        self.lista.bind("<ButtonRelease-1>", self._clic)
        self.lista.bind("<Double-Button-1>", self._doble)
        pie = tk.Frame(self, bg=M.C["fondo"])
        pie.pack(fill="x", padx=8, pady=(2, 8))
        if not alelegir:
            tk.Button(pie, text="Automático", relief="flat", command=self._auto,
                      bg=M.C.get("boton", "#2f2f33"), fg=M.C["texto"],
                      font=M.letra(8)).pack(side="left")
        tk.Button(pie, text="Cerrar", relief="flat", command=self.destroy,
                  bg=M.C.get("boton", "#2f2f33"), fg=M.C["texto"],
                  font=M.letra(8)).pack(side="right")
        self.bind("<Escape>", lambda e: self.destroy())
        self.carpetas = None
        self.marcadas = carpetas_leidas(app, M)
        self._visibles = []
        self._pendiente = None
        self._aplicar_id = None
        self.lista.insert("end", "Leyendo las carpetas de Outlook…")

        def leer():
            try:
                res = M.listar_carpetas_correo_outlook()
            except Exception:
                res = []
            self._pendiente = res
        threading.Thread(target=leer, daemon=True).start()
        self.after(200, self._recoger)

    def _recoger(self):
        try:
            if not self.winfo_exists():
                return
        except Exception:
            return
        if self._pendiente is None:
            self.after(200, self._recoger)
            return
        self.carpetas = self._pendiente
        self._pintar()

    def _pintar(self):
        if self.carpetas is None:
            return
        texto = self.filtro.get().strip().lower()
        self._visibles = [c for c in self.carpetas
                          if not texto or texto in c["nombre"].lower()]
        self.lista.delete(0, "end")
        if not self._visibles:
            self.lista.insert("end", "No hay carpetas que coincidan")
        for c in self._visibles:
            n = c["nombre"]
            marca = "" if self.alelegir else (
                "☑  " if n in self.marcadas else "☐  ")
            extra = f"  ({c['no_leidos']})" if c["no_leidos"] else ""
            self.lista.insert("end", marca + n.replace("/", " › ") + extra)

    def _indice(self, e):
        i = self.lista.nearest(e.y)
        return i if 0 <= i < len(self._visibles) else None

    def _clic(self, e):
        if self.alelegir or self.carpetas is None:
            return
        i = self._indice(e)
        if i is None:
            return
        n = self._visibles[i]["nombre"]
        if n in self.marcadas:
            self.marcadas.discard(n)
        else:
            self.marcadas.add(n)
        arriba = self.lista.yview()[0]
        self._pintar()
        self.lista.yview_moveto(arriba)
        # se aplica cuando dejas de pulsar (varios clics seguidos = una recarga)
        if self._aplicar_id:
            self.after_cancel(self._aplicar_id)
        self._aplicar_id = self.after(800, self._aplicar)

    def _aplicar(self):
        self._aplicar_id = None
        orden = [c["nombre"] for c in self.carpetas if c["nombre"] in self.marcadas]
        fijar_carpetas_leidas(self.app, self.M, orden)

    def _doble(self, e):
        if not self.alelegir:
            return
        i = self._indice(e)
        if i is not None:
            n = self._visibles[i]["nombre"]
            self.destroy()
            self.alelegir(n)

    def _auto(self):
        self.M.CFG["carpetas_correo"] = []
        self.M.guardar_config()
        self.app._correo_ahora()
        self.destroy()


# --------------------------------------------------------------------------
# Menús de la columna
# --------------------------------------------------------------------------

def construir_menu(app, M, padre):
    """El menú completo de la columna (cabecera ⋯, clic derecho y comandos)."""
    m = tk.Menu(padre, tearoff=0)
    filtro, fuera, solo = estado(app, M)
    visible = bool(M.CFG.get("correo_visible"))
    m.add_command(label=("Ocultar" if visible else "Mostrar")
                  + " el correo en el widget",
                  command=lambda: alternar_columna(app, M))
    m.add_separator()
    m.add_command(label=("✓ " if solo else "     ") + "Solo los no leídos",
                  command=lambda: alternar_sin_leer(app, M))
    m.add_command(label=("● " if not filtro else "     ")
                  + "Todas las carpetas",
                  command=lambda: (poner_filtro(app, M, None)
                                   if filtro else None))
    ver = tk.Menu(m, tearoff=0)
    quitar = tk.Menu(m, tearoff=0)
    if not app.correo_carpetas:
        ver.add_command(label="(aún no hay carpetas cargadas)", state="disabled")
        quitar.add_command(label="(aún no hay carpetas cargadas)",
                           state="disabled")
    todos = [c["nombre"] for c in app.correo_carpetas]
    for c in app.correo_carpetas:
        nombre = c["nombre"]
        corto = etiqueta_corta(nombre, todos)
        cuenta = f"  ({c['no_leidos']})" if c.get("no_leidos") else ""
        ver.add_command(label=("● " if filtro == nombre else "     ")
                        + corto + cuenta,
                        command=lambda n=nombre: poner_filtro(app, M, n))
        quitar.add_command(label=("✓ " if nombre in fuera else "     ")
                           + corto,
                           command=lambda n=nombre: alternar_exclusion(app, M, n))
    m.add_command(label="Elegir qué carpetas se leen…",
                  command=lambda: abrir_selector_carpetas(app, M))
    m.add_cascade(label="Ver solo una carpeta", menu=ver)
    m.add_cascade(label="Excluir carpetas", menu=quitar)
    m.add_command(label="Quitar todos los filtros",
                  command=lambda: limpiar_filtros(app, M))
    m.add_separator()
    m.add_command(label="Actualizar el correo", command=app._correo_ahora)
    m.add_command(label="Abrir la bandeja completa",
                  command=lambda: app._abrir_correo(modo="bandeja"))
    ancho = tk.Menu(m, tearoff=0)
    actual = M.CFG.get("ancho_correo", 260)
    for px in (200, 240, 260, 320, 400, 500):
        ancho.add_command(label=("● " if px == actual else "     ") + f"{px} px",
                          command=lambda v=px: cambiar_ancho(app, M, v))
    m.add_cascade(label="Ancho de la columna", menu=ancho)
    chips = tk.Menu(m, tearoff=0)
    actual_c = M.CFG.get("correo_chips_max", 3)
    for n in (0, 2, 3, 5, 8):
        chips.add_command(
            label=("● " if n == actual_c else "     ")
            + (f"{n} carpetas" if n else "Ninguno (todo en el menú ▾)"),
            command=lambda v=n: _poner_chips(app, M, v))
    m.add_cascade(label="Botones de carpeta a la vista", menu=chips)
    return m


# --------------------------------------------------------------------------
# Dibujo
# --------------------------------------------------------------------------

def ancho_columna(M):
    return round(M.CFG.get("ancho_correo", 260) * M.ESCALA)


def _alto_fila(M):
    return round(34 * M.ESCALA)


def alto_deseado(app, M):
    """Alto que le gustaría a la columna, para que la ventana crezca si hace
    falta (con un tope de filas: el resto se ve si la ventana es más alta)."""
    n = min(len(mensajes_de_la_columna(app, M)),
            max(1, int(M.CFG.get("correo_filas_max", 8))))
    fila = _alto_fila(M)
    chips = 2 * round(22 * M.ESCALA)
    return (M.Y_LISTA + 36 + round(6 * M.ESCALA) + chips + max(n, 1) * fila
            + round(10 * M.ESCALA) + M.ALTO_ESTADO + 4)


def _recortar(app, texto, fuente, ancho):
    if ancho <= 10:
        return ""
    texto = " ".join(str(texto).split())[:120]
    if app._ancho_texto(texto, fuente) <= ancho:
        return texto
    while len(texto) > 1 and app._ancho_texto(texto + "…", fuente) > ancho:
        texto = texto[:-1]
    return texto.rstrip() + "…"


def _divisor(app, M, x, alto):
    """Borde izquierdo de la columna: se arrastra para cambiar su ancho."""
    c = app.lienzo
    y0, y1 = M.BARRA, alto - M.ALTO_ESTADO
    media = app._grosor_divisor() / 2
    c.create_rectangle(x - media, y0, x + media, y1, fill=M.C["fondo"],
                       outline="", tags=("divisor_correo",))
    c.create_line(x, y0, x, y1, fill=M.C["texto_tenue"],
                  tags=("divisor_correo", "divisor_correo_tinta"))
    app._tirador(x, (y0 + y1) / 2, ("divisor_correo", "divisor_correo_tinta"))


def _es_oscuro(M):
    try:
        h = M.C["fondo"].lstrip("#")
        r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
        return (0.299 * r + 0.587 * g + 0.114 * b) < 128
    except Exception:
        return True


# color de cada icono: vivo sobre fondo oscuro, más profundo sobre fondo claro
_COLOR_ICONO = {
    "escribir": ("#5aa9ff", "#0b5fd6"),
    "leidos": ("#3ddc84", "#12813f"),
    "carpeta": ("#ffc857", "#a86a00"),
}


def _dibujar_icono_accion(c, nombre, cx, cy, tam, tinta, tag):
    """Iconos vectoriales gruesos y con color, nítidos a cualquier tamaño."""
    g = max(2, round(tam * 0.11))
    kw = dict(fill=tinta, width=g, capstyle="round", joinstyle="round",
              tags=(tag,))
    k = tam * 0.40
    if nombre == "escribir":                    # sobre con un «+»
        ax, ay = k * 1.05, k * 0.72
        c.create_rectangle(cx - ax, cy - ay, cx + ax, cy + ay, outline=tinta,
                           width=g, tags=(tag,))
        c.create_line(cx - ax, cy - ay, cx, cy + ay * 0.15, cx + ax, cy - ay,
                      **kw)
        px, py, r = cx + ax, cy + ay, k * 0.5   # «+» en la esquina
        c.create_oval(px - r, py - r, px + r, py + r, fill=tinta, outline="",
                      tags=(tag,))
        c.create_line(px - r * 0.55, py, px + r * 0.55, py, fill="#ffffff",
                      width=max(2, g - 1), tags=(tag,))
        c.create_line(px, py - r * 0.55, px, py + r * 0.55, fill="#ffffff",
                      width=max(2, g - 1), tags=(tag,))
    elif nombre == "leidos":                     # doble check grueso
        g2 = g + 1
        kw2 = dict(kw, width=g2)
        for dx in (-k * 0.42, k * 0.42):
            c.create_line(cx + dx - k * 0.62, cy + k * 0.02,
                          cx + dx - k * 0.16, cy + k * 0.52,
                          cx + dx + k * 0.66, cy - k * 0.52, **kw2)
    elif nombre == "carpeta":                    # carpeta rellena
        ancho, alto = k * 1.1, k * 0.85
        c.create_polygon(cx - ancho, cy - alto, cx - ancho * 0.2, cy - alto,
                         cx, cy - alto * 0.55, cx + ancho, cy - alto * 0.55,
                         cx + ancho, cy + alto, cx - ancho, cy + alto,
                         fill=tinta, outline=tinta, width=1, tags=(tag,))


def _boton_icono(app, M, x, y, ancho, alto, icono, tag, accion, ayuda=None):
    """Botón con un icono propio y de color (para la fila de acciones)."""
    c = app.lienzo
    oscuro = _es_oscuro(M)
    claro_c, oscuro_c = _COLOR_ICONO.get(icono, ("#ffffff", "#000000"))
    tinta = claro_c if oscuro else oscuro_c
    c.create_rectangle(x, y, x + ancho, y + alto,
                       fill=M.C.get("boton", "#2f2f33"), outline=tinta,
                       width=1, tags=(ETIQUETA, tag))
    _dibujar_icono_accion(c, icono, x + ancho / 2, y + alto / 2,
                          alto * 0.95, tinta, tag)
    c.tag_bind(tag, "<Button-1>", lambda e: accion(e))
    c.tag_bind(tag, "<Enter>",
              lambda e: (c.config(cursor="hand2"),
                        [c.itemconfig(i, fill=M.C.get("boton_raton", "#40404a"))
                         for i in c.find_withtag(tag)
                         if c.type(i) == "rectangle"]))
    c.tag_bind(tag, "<Leave>",
              lambda e: (c.config(cursor=""),
                        [c.itemconfig(i, fill=M.C.get("boton", "#2f2f33"))
                         for i in c.find_withtag(tag)
                         if c.type(i) == "rectangle"]))
    if ayuda:
        app._vincular_ayuda(tag, ayuda)
    return x + ancho


def _marcar_todos_leidos(app, M):
    filtro, fuera, _ = estado(app, M)
    carpetas = [f["nombre"] for f in carpetas_visibles(app.correo_carpetas, fuera)
               if not filtro or f["nombre"] == filtro]
    if not carpetas:
        return

    esperados = sum(c.get("no_leidos", 0) for c in app.correo_carpetas
                    if c["nombre"] in carpetas)

    def trabajo():
        return M.marcar_todos_leidos(carpetas), ""

    def despues(res):
        n, _ = res
        # Solo se enseña como leído lo que Outlook confirma. Antes se ponía todo
        # a 0 a ciegas y, al refrescar, lo que no se había marcado reaparecía.
        if n:
            for lista in app.correo_mensajes.values():
                for m in lista:
                    if m.carpeta in carpetas:
                        m.leido = True
        _repintar_todo(app)
        if n and n >= esperados:
            _aviso(app, M, "✉ Correo", f"{n} correo(s) marcados como leídos",
                   True)
        elif n:
            _aviso(app, M, "✉ Correo",
                   f"{n} marcados de {esperados}. Compruebo el resto…", False)
        elif esperados:
            _aviso(app, M, "✉ No se pudo",
                   "Outlook no ha dejado marcarlos (mira el registro)", False)
        else:
            _aviso(app, M, "✉ Correo", "No había nada sin leer", True)
        app._correo_ahora()              # los contadores reales, desde Outlook
    _en_hilo(app, trabajo, despues)


def _nuevo_correo(app, M):
    _en_hilo(app, lambda: (M.nuevo_correo(), ""),
             lambda res: None if res[0] else
             _aviso(app, M, "✉ No se pudo", "Outlook no ha respondido", False))


def _boton_texto(app, M, x, y, texto, tag, accion, derecha=False, **kw):
    c = app.lienzo
    item = c.create_text(x, y, text=texto, anchor="ne" if derecha else "nw",
                         tags=(ETIQUETA, tag), **kw)
    c.tag_bind(tag, "<Button-1>", lambda e: accion(e))
    c.tag_bind(tag, "<Enter>", lambda e: c.config(cursor="hand2"))
    c.tag_bind(tag, "<Leave>", lambda e: c.config(cursor=""))
    return item


def pintar(app, M, x0, ancho, alto_total):
    """Dibuja la columna de correo entre x0 y x0+ancho."""
    c, C, letra, esc = app.lienzo, M.C, M.letra, M.ESCALA
    if x0:
        _divisor(app, M, x0, alto_total)
    filtro, fuera, solo = estado(app, M)
    info = app.correo_carpetas
    sin_leer = sum(f.get("no_leidos", 0)
                   for f in carpetas_visibles(info, fuera))

    # cabecera
    y = M.Y_LISTA + 2
    c.create_text(x0 + 10, y, text="✉ Correo", anchor="nw",
                  font=letra(9, "bold"), fill=C["texto"], tags=(ETIQUETA,))
    _boton_texto(app, M, x0 + ancho - 10, y, "⋯", "correo_menu",
                 lambda e: (construir_menu(app, M, app.lienzo)
                            .tk_popup(e.x_root, e.y_root), "break")[1],
                 derecha=True, font=letra(11, "bold"), fill=C["texto_tenue"])
    c.create_line(x0 + 8, M.Y_LISTA + 20, x0 + ancho - 8, M.Y_LISTA + 20,
                  fill=C["rejilla"], tags=(ETIQUETA,))
    y = M.Y_LISTA + 26

    # fila de acciones (no son filtros): escribir, marcar todo leído, carpetas
    if M.CFG.get("correo_activo", True) and not app.correo_error:
        alto_ac = round(28 * esc)
        acciones = (
            ("escribir", "correo_ac_nuevo", lambda e: _nuevo_correo(app, M),
             "Escribir un correo nuevo"),
            ("leidos", "correo_ac_leidos", lambda e: _marcar_todos_leidos(app, M),
             "Marcar todo como leído (según el filtro de ahora)"),
            ("carpeta", "correo_ac_carpetas", lambda e: abrir_selector_carpetas(app, M),
             "Elegir qué carpetas de Outlook se leen"),
        )
        xa = x0 + 10
        for icono, tag, accion, ayuda in acciones:
            w = round(44 * esc)
            _boton_icono(app, M, xa, y, w, alto_ac, icono, tag, accion, ayuda)
            xa += w + round(4 * esc)
        y += alto_ac + 6

    tope = alto_total - M.ALTO_ESTADO - 4
    if not M.CFG.get("correo_activo", True):
        c.create_text(x0 + ancho / 2, y + 10, text="El correo está desactivado",
                      fill=C["texto_tenue"], font=letra(8), tags=(ETIQUETA,))
        return
    if app.correo_error and not info:
        c.create_text(x0 + 10, y + 4, anchor="nw", width=ancho - 20,
                      text=f"⚠ {app.correo_error}", fill=C["ahora"],
                      font=letra(8), tags=(ETIQUETA,))
        return
    if not info:
        c.create_text(x0 + ancho / 2, y + 10, text="Cargando el correo…",
                      fill=C["texto_tenue"], font=letra(8), tags=(ETIQUETA,))
        return

    # chips de carpetas (con salto de línea si no caben)
    fuente_chip = letra(7, "bold")
    alto_chip = round(16 * esc)
    x = x0 + 10
    limite_x = x0 + ancho - 8

    def chip(texto, activo, tag, accion_izq, accion_der=None,
             con_evento=False):
        nonlocal x, y
        w = app._ancho_texto(texto, fuente_chip) + round(14 * esc)
        if x + w > limite_x and x > x0 + 10:
            x = x0 + 10
            y += alto_chip + 3
        fondo = C["evento"] if activo else C.get("boton", "#2f2f33")
        tinta = C["evento_texto"] if activo else C["texto"]
        c.create_rectangle(x, y, x + w, y + alto_chip, fill=fondo, outline="",
                           tags=(ETIQUETA, tag))
        c.create_text(x + w / 2, y + alto_chip / 2, text=texto, fill=tinta,
                      font=fuente_chip, tags=(ETIQUETA, tag))
        c.tag_bind(tag, "<Button-1>",
                   (lambda e: accion_izq(e)) if con_evento
                   else (lambda e: accion_izq()))
        if accion_der:
            c.tag_bind(tag, "<Button-3>", lambda e: accion_der(e))
        c.tag_bind(tag, "<Enter>", lambda e: c.config(cursor="hand2"))
        c.tag_bind(tag, "<Leave>", lambda e: c.config(cursor=""))
        x += w + 3

    chip("Todas", filtro is None, "correo_chip_todas",
         lambda: poner_filtro(app, M, None))
    nombres = [f["nombre"] for f in info]
    visibles = carpetas_visibles(info, fuera)
    maximo = M.CFG.get("correo_chips_max", 3)
    con_boton = chips_a_mostrar(visibles, filtro, maximo)
    for i, f in enumerate(con_boton):
        nombre = f["nombre"]
        texto = (etiqueta_corta(nombre, nombres)[:18]
                 + (f" {f['no_leidos']}" if f.get("no_leidos") else ""))
        chip(texto, filtro == nombre, f"correo_chip_{i}",
             lambda n=nombre: poner_filtro(app, M, n),
             lambda e, n=nombre: _menu_chip(app, M, e, n))
    resto = len(visibles) - len(con_boton)
    if resto > 0:
        chip(f"▾ {resto} más", False, "correo_chip_mas",
             lambda e: _menu_carpetas(app, M, e), con_evento=True)
    chip("🔴 Sin leer" if solo else "○ Sin leer", solo, "correo_chip_sin_leer",
         lambda: alternar_sin_leer(app, M))
    y += alto_chip + 4

    if fuera:
        _boton_texto(
            app, M, x0 + 10, y,
            f"🚫 {len(fuera)} excluida{'s' if len(fuera) != 1 else ''}"
            " · gestionar", "correo_excluidas",
            lambda e: (construir_menu(app, M, app.lienzo)
                       .tk_popup(e.x_root, e.y_root), "break")[1],
            font=letra(7), fill=C["texto_tenue"])
        y += round(14 * esc)
    c.create_line(x0 + 8, y, x0 + ancho - 8, y, fill=C["rejilla"],
                  tags=(ETIQUETA,))
    y += 4

    # mensajes
    mensajes = mensajes_de_la_columna(app, M)
    if not mensajes:
        c.create_text(x0 + ancho / 2, y + 14, text="Nada con este filtro",
                      fill=C["texto_tenue"], font=letra(8), tags=(ETIQUETA,))
        return
    fila = _alto_fila(M)
    hoy = dt.datetime.now(M.TZ).date()
    f_remite = letra(8, "bold")
    f_asunto = letra(8)
    pintados = 0
    for n, m in enumerate(mensajes):
        if y + fila > tope - round(14 * esc):
            break
        tag = f"correo_msg_{n}"
        etiquetas = (ETIQUETA, tag)
        color = m.color_carpeta or C["evento"]
        # barra de color de la carpeta: sólida si está sin leer
        c.create_rectangle(x0 + 8, y + 3, x0 + 11, y + fila - 5,
                           fill=color if not m.leido
                           else M.mezclar(color, C["fondo"], 0.3),
                           outline="", tags=etiquetas)
        hora = (f"{m.fecha:%H:%M}" if m.fecha.date() == hoy
                else f"{m.fecha:%d/%m}")
        ancho_hora = app._ancho_texto(hora, letra(7)) + 6
        c.create_text(x0 + ancho - 10, y + 3, text=hora, anchor="ne",
                      font=letra(7), fill=C["texto_tenue"], tags=etiquetas)
        remite = _recortar(app, m.remitente, f_remite if not m.leido
                           else f_asunto, ancho - 26 - ancho_hora)
        c.create_text(x0 + 16, y + 2, text=remite, anchor="nw",
                      font=f_remite if not m.leido else f_asunto,
                      fill=C["texto"], tags=etiquetas)
        asunto = (("🚩 " if getattr(m, "marcado", False) else "") + m.asunto
                  + (" 📎" if m.adjuntos else ""))
        c.create_text(x0 + 16, y + 3 + round(14 * esc),
                      text=_recortar(app, asunto, f_asunto, ancho - 26),
                      anchor="nw", font=f_asunto,
                      fill=C["texto"] if not m.leido else C["texto_tenue"],
                      tags=etiquetas)
        # zona sensible a toda la fila
        c.create_rectangle(x0 + 6, y, x0 + ancho - 6, y + fila - 2, fill="",
                           outline="", tags=etiquetas)
        c.tag_bind(tag, "<Button-1>", lambda e, m=m: abrir_mensaje(app, M, m))
        c.tag_bind(tag, "<Button-3>", lambda e, m=m: menu_mensaje(app, M, e, m))
        c.tag_bind(tag, "<Enter>", lambda e: c.config(cursor="hand2"))
        c.tag_bind(tag, "<Leave>", lambda e: c.config(cursor=""))
        y += fila
        pintados += 1
    resto = len(mensajes) - pintados
    if resto > 0:
        _boton_texto(app, M, x0 + 10, y + 2,
                     f"+{resto} más · abrir la bandeja", "correo_mas",
                     lambda e: app._abrir_correo(modo="bandeja"),
                     font=letra(7), fill=C["texto_tenue"])


def _menu_carpetas(app, M, e):
    """Todas las carpetas en un menú: ver solo una, o excluirla."""
    filtro, fuera, _ = estado(app, M)
    todos = [c["nombre"] for c in app.correo_carpetas]
    menu = tk.Menu(app, tearoff=0)
    for c in carpetas_visibles(app.correo_carpetas, fuera):
        nombre = c["nombre"]
        cuenta = f"  ({c['no_leidos']})" if c.get("no_leidos") else ""
        menu.add_command(
            label=("● " if filtro == nombre else "     ")
            + etiqueta_corta(nombre, todos) + cuenta,
            command=lambda n=nombre: poner_filtro(app, M, n))
    menu.add_separator()
    menu.add_command(label="Elegir qué carpetas se leen…",
                     command=lambda: abrir_selector_carpetas(app, M))
    menu.tk_popup(e.x_root, e.y_root)
    return "break"


def _menu_chip(app, M, e, nombre):
    menu = tk.Menu(app, tearoff=0)
    corto = etiqueta_corta(nombre, [c["nombre"] for c in app.correo_carpetas])
    menu.add_command(label=f"Ver solo «{corto}»",
                     command=lambda: poner_filtro(app, M, nombre))
    menu.add_command(label=f"Excluir «{corto}»",
                     command=lambda: alternar_exclusion(app, M, nombre))
    menu.tk_popup(e.x_root, e.y_root)
    return "break"
