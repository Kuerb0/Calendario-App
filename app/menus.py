"""Menú emergente propio (clic derecho) para toda la aplicación.

El menú de Windows que trae Tk es lento de abrir, no sigue el tema del
widget y se ve antiguo. Este lo sustituye sin tocar el resto del programa:
`instalar(M)` cambia `tk.Menu.tk_popup`, así que TODOS los menús que ya
existen (con sus submenús, casillas ✓ y puntos ●) se dibujan con el estilo
del widget, aparecen al instante y se pueden manejar con el teclado.

    ↑ ↓ moverse · Enter elegir · → abre submenú · ← / Esc cierra

Se apaga con CFG["menu_moderno"] = False (vuelve el menú de Windows).

Este módulo no importa calendario.py: recibe el módulo ya cargado como 'M'.
"""
import tkinter as tk

_ORIGINAL = tk.Menu.tk_popup
_ABIERTOS = []           # cadena de menús abiertos (raíz, submenú, ...)


# --------------------------------------------------------------------------
# Lectura de un tk.Menu
# --------------------------------------------------------------------------

def leer_entradas(menu):
    """Lista de dicts con lo que hay que dibujar de cada entrada del menú."""
    try:
        ultimo = menu.index("end")
    except tk.TclError:
        ultimo = None
    if ultimo is None:
        return []
    filas = []
    for i in range(ultimo + 1):
        try:
            tipo = menu.type(i)
        except tk.TclError:
            continue
        if tipo == "tearoff":
            continue
        if tipo == "separator":
            filas.append({"tipo": "sep"})
            continue
        try:
            etiqueta = str(menu.entrycget(i, "label"))
            estado = str(menu.entrycget(i, "state"))
        except tk.TclError:
            continue
        marca = ""
        if etiqueta.startswith(("✓ ", "● ")):        # convenciones del programa
            marca, etiqueta = etiqueta[0], etiqueta[2:]
        elif etiqueta.startswith("     "):            # «sin marcar» alineado
            etiqueta = etiqueta.lstrip()
            marca = " "
        if tipo in ("checkbutton", "radiobutton"):
            try:
                var = menu.entrycget(i, "variable")
                valor = menu.getvar(var)
                if tipo == "checkbutton":
                    activo = str(valor) == str(menu.entrycget(i, "onvalue"))
                else:
                    activo = str(valor) == str(menu.entrycget(i, "value"))
                marca = ("✓" if tipo == "checkbutton" else "●") if activo else " "
            except tk.TclError:
                pass
        try:
            color = str(menu.entrycget(i, "foreground")) or None
        except tk.TclError:
            color = None
        try:
            atajo = str(menu.entrycget(i, "accelerator"))
        except tk.TclError:
            atajo = ""
        sub = None
        if tipo == "cascade":
            try:
                sub = menu.nametowidget(str(menu.entrycget(i, "menu")))
            except (tk.TclError, KeyError):
                sub = None
        filas.append({"tipo": tipo, "indice": i, "etiqueta": etiqueta,
                      "marca": marca, "deshabilitado": estado == "disabled",
                      "color": color, "atajo": atajo, "submenu": sub})
    # un separador al principio, al final o repetido no aporta nada
    limpias = []
    for f in filas:
        if f["tipo"] == "sep" and (not limpias or limpias[-1]["tipo"] == "sep"):
            continue
        limpias.append(f)
    while limpias and limpias[-1]["tipo"] == "sep":
        limpias.pop()
    return limpias


def _area_pantalla(root, x, y):
    """(izq, arriba, der, abajo) del área útil del monitor donde está (x, y)."""
    try:
        import ctypes
        from ctypes import wintypes

        class MONITORINFO(ctypes.Structure):
            _fields_ = [("cb", wintypes.DWORD), ("mon", wintypes.RECT),
                        ("work", wintypes.RECT), ("flags", wintypes.DWORD)]
        user32 = ctypes.windll.user32
        h = user32.MonitorFromPoint(wintypes.POINT(int(x), int(y)), 2)
        mi = MONITORINFO()
        mi.cb = ctypes.sizeof(MONITORINFO)
        if user32.GetMonitorInfoW(h, ctypes.byref(mi)):
            w = mi.work
            return w.left, w.top, w.right, w.bottom
    except Exception:
        pass
    return 0, 0, root.winfo_screenwidth(), root.winfo_screenheight()


# --------------------------------------------------------------------------
# La ventanita del menú
# --------------------------------------------------------------------------

class MenuModerno(tk.Toplevel):
    FILAS_MAX = 22               # más filas que estas: el menú se desplaza

    def __init__(self, maestro, menu, M, padre=None):
        super().__init__(maestro)
        self.M, self.menu, self.padre = M, menu, padre
        self.hijo = None
        self.filas = leer_entradas(menu)
        self.marcos = {}           # posición -> Frame de la fila
        self.sel = None
        self._cierre_id = None
        self.withdraw()            # no se ve hasta estar colocado
        self.overrideredirect(True)
        C = M.C
        self.configure(bg=C["rejilla"])
        try:
            self.attributes("-topmost", True)
        except tk.TclError:
            pass
        self.fuente = M.letra(9)
        self.alto_fila = max(24, round(26 * M.ESCALA * max(1.0, M.LETRA)))
        self.hay_marca = any(f.get("marca") for f in self.filas)
        self.hay_flecha = any(f["tipo"] == "cascade" for f in self.filas)
        self.ancho_fila = self._medir()
        self._construir()

    def _medir(self):
        """Ancho que necesita la fila más ancha (texto, marca, atajo, flecha)."""
        import tkinter.font as tkfont
        fuente = tkfont.Font(root=self.master, font=self.fuente)
        pequena = tkfont.Font(root=self.master, font=self.M.letra(8))
        mayor = 0
        for f in self.filas:
            if f["tipo"] == "sep":
                continue
            w = fuente.measure(f["etiqueta"]) + 26
            if f["atajo"]:
                w += pequena.measure(f["atajo"]) + 24
            mayor = max(mayor, w)
        extra = (30 if self.hay_marca else 0) + (26 if self.hay_flecha else 0)
        return max(round(180 * self.M.ESCALA), mayor + extra)

    # -- construcción ------------------------------------------------------

    def _construir(self):
        C = self.M.C
        n_filas = sum(1 for f in self.filas if f["tipo"] != "sep")
        exterior = tk.Frame(self, bg=C["rejilla"])
        exterior.pack(fill="both", expand=True, padx=1, pady=1)
        self.lienzo = tk.Canvas(exterior, bg=C["fondo"], highlightthickness=0,
                                bd=0)
        self.lienzo.pack(fill="both", expand=True)
        self.cuerpo = tk.Frame(self.lienzo, bg=C["fondo"])
        self.ventana_cuerpo = self.lienzo.create_window(
            0, 0, window=self.cuerpo, anchor="nw")
        self.cuerpo.pack_propagate(True)
        tk.Frame(self.cuerpo, bg=C["fondo"], height=4).pack(fill="x")
        for pos, f in enumerate(self.filas):
            if f["tipo"] == "sep":
                tk.Frame(self.cuerpo, bg=C["rejilla"], height=1).pack(
                    fill="x", padx=10, pady=4)
                continue
            self.marcos[pos] = self._fila(pos, f)
        tk.Frame(self.cuerpo, bg=C["fondo"], height=4).pack(fill="x")
        self.update_idletasks()
        ancho = self.ancho_fila + 8
        alto = self.cuerpo.winfo_reqheight()
        tope = self.alto_fila * self.FILAS_MAX + 8
        self.desplazable = alto > tope
        if self.desplazable:
            alto = tope
        self.lienzo.configure(width=ancho, height=alto,
                              scrollregion=(0, 0, ancho,
                                            self.cuerpo.winfo_reqheight()))
        self.lienzo.itemconfigure(self.ventana_cuerpo, width=ancho)
        self.tam = (ancho + 2, alto + 2)
        self.geometry(f"{self.tam[0]}x{self.tam[1]}")
        for w in (self, self.lienzo, self.cuerpo):
            w.bind("<MouseWheel>", self._rueda)
        self.bind("<Key>", self._tecla)

    def _fila(self, pos, f):
        C = self.M.C
        marco = tk.Frame(self.cuerpo, bg=C["fondo"], height=self.alto_fila,
                         width=self.ancho_fila,
                         cursor="hand2" if not f["deshabilitado"] else "")
        marco.pack(fill="x", padx=4)
        marco.pack_propagate(False)
        tinta = (C["texto_tenue"] if f["deshabilitado"]
                 else (f["color"] or C["texto"]))
        partes = []
        if self.hay_marca:
            partes.append(tk.Label(
                marco, text=f["marca"], bg=C["fondo"], fg=C["evento"],
                font=self.fuente, width=2, anchor="center"))
            partes[-1].pack(side="left", padx=(4, 0))
        etiqueta = tk.Label(marco, text=f["etiqueta"], bg=C["fondo"], fg=tinta,
                            font=self.fuente, anchor="w")
        etiqueta.pack(side="left", padx=(8 if not self.hay_marca else 2, 8),
                      fill="y")
        partes.append(etiqueta)
        if f["atajo"]:
            t = tk.Label(marco, text=f["atajo"], bg=C["fondo"],
                         fg=C["texto_tenue"], font=self.M.letra(8))
            t.pack(side="right", padx=(0, 10))
            partes.append(t)
        if f["tipo"] == "cascade":
            t = tk.Label(marco, text="›", bg=C["fondo"], fg=C["texto_tenue"],
                         font=self.M.letra(11, "bold"))
            t.pack(side="right", padx=(0, 8))
            partes.append(t)
        marco._partes = partes
        for w in [marco] + partes:
            w.bind("<Enter>", lambda e, p=pos: self._entrar(p))
            w.bind("<Button-1>", lambda e, p=pos: self._pulsar(p))
            w.bind("<MouseWheel>", self._rueda)
        return marco

    # -- resaltado y submenús ---------------------------------------------

    def _resaltar(self, pos):
        C = self.M.C
        if self.sel is not None and self.sel in self.marcos:
            self._pintar_fila(self.marcos[self.sel], C["fondo"], None)
        self.sel = pos
        if pos is not None and pos in self.marcos:
            f = self.filas[pos]
            self._pintar_fila(
                self.marcos[pos],
                C.get("boton_raton", "#40404a") if not f["deshabilitado"]
                else C["fondo"], f)

    def _pintar_fila(self, marco, fondo, f):
        marco.configure(bg=fondo)
        for w in marco._partes:
            w.configure(bg=fondo)

    def _entrar(self, pos):
        f = self.filas[pos]
        if f["deshabilitado"]:
            self._cerrar_hijo()
            self._resaltar(None)
            return
        self._resaltar(pos)
        if f["tipo"] == "cascade" and f["submenu"] is not None:
            if self._cierre_id:
                self.after_cancel(self._cierre_id)
            self._cierre_id = self.after(140, lambda: self._abrir_sub(pos))
        else:
            if self._cierre_id:
                self.after_cancel(self._cierre_id)
                self._cierre_id = None
            self._cerrar_hijo()

    def _cerrar_hijo(self):
        if self.hijo is not None:
            hijo, self.hijo = self.hijo, None
            hijo._cerrar_cadena()

    def _abrir_sub(self, pos):
        self._cierre_id = None
        try:
            if not self.winfo_exists() or self.sel != pos:
                return
        except tk.TclError:
            return
        f = self.filas[pos]
        if self.hijo is not None and self.hijo.menu is f["submenu"]:
            return
        self._cerrar_hijo()
        marco = self.marcos[pos]
        x = self.winfo_rootx() + self.winfo_width() - 4
        y = marco.winfo_rooty() - 4
        sub = MenuModerno(self.master, f["submenu"], self.M, padre=self)
        if not sub.filas:
            sub.destroy()
            return
        self.hijo = sub
        sub._colocar(x, y, x_izq=self.winfo_rootx() + 4)
        _ABIERTOS.append(sub)

    # -- acciones ----------------------------------------------------------

    def _pulsar(self, pos):
        f = self.filas[pos]
        if f["deshabilitado"]:
            return "break"
        if f["tipo"] == "cascade":
            self._abrir_sub(pos)
            return "break"
        menu, indice = self.menu, f["indice"]
        raiz = self.master
        cerrar_todo()
        # se ejecuta ya con el menú cerrado (así los cuadros que abra la
        # acción no quedan por detrás)
        raiz.after(10, lambda: _invocar(menu, indice))
        return "break"

    def _rueda(self, e):
        if self.desplazable:
            self.lienzo.yview_scroll(-1 if e.delta > 0 else 1, "units")
            self.lienzo.yview_scroll(-2 if e.delta > 0 else 2, "units")
        return "break"

    def _tecla(self, e):
        tecla = e.keysym
        if tecla == "Escape":
            cerrar_todo()
            return "break"
        activos = [p for p, f in enumerate(self.filas)
                   if f["tipo"] != "sep" and not f["deshabilitado"]]
        if not activos:
            return "break"
        if tecla in ("Down", "Up"):
            if self.sel in activos:
                i = activos.index(self.sel) + (1 if tecla == "Down" else -1)
            else:
                i = 0 if tecla == "Down" else len(activos) - 1
            self._cerrar_hijo()
            self._resaltar(activos[i % len(activos)])
            self._asegurar_visible(self.sel)
        elif tecla in ("Return", "KP_Enter", "space") and self.sel is not None:
            self._pulsar(self.sel)
        elif tecla == "Right" and self.sel is not None:
            f = self.filas[self.sel]
            if f["tipo"] == "cascade":
                self._abrir_sub(self.sel)
                if self.hijo is not None:
                    self.hijo.focus_force()
                    self.hijo._resaltar(next(
                        (p for p, g in enumerate(self.hijo.filas)
                         if g["tipo"] != "sep" and not g["deshabilitado"]),
                        None))
        elif tecla == "Left" and self.padre is not None:
            padre = self.padre
            self._cerrar_cadena()
            padre.hijo = None
            padre.focus_force()
        return "break"

    def _asegurar_visible(self, pos):
        if not self.desplazable or pos not in self.marcos:
            return
        marco = self.marcos[pos]
        total = max(self.cuerpo.winfo_reqheight(), 1)
        alto = self.lienzo.winfo_height()
        arriba = self.lienzo.canvasy(0)
        y0, y1 = marco.winfo_y(), marco.winfo_y() + marco.winfo_height()
        if y0 < arriba:
            self.lienzo.yview_moveto(max(0, (y0 - 4) / total))
        elif y1 > arriba + alto:
            self.lienzo.yview_moveto(min(1, (y1 - alto + 4) / total))

    # -- colocación y cierre -----------------------------------------------

    def _colocar(self, x, y, x_izq=None):
        izq, arriba, der, abajo = _area_pantalla(self.master, x, y)
        ancho, alto = self.tam
        if x + ancho > der:
            x = (x_izq - ancho) if x_izq is not None else der - ancho
            if x_izq is None:
                x = max(izq, x)
        if y + alto > abajo:
            y = max(arriba, abajo - alto)
        x, y = max(izq, int(x)), max(arriba, int(y))
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        try:
            self.attributes("-alpha", 0.0)
        except tk.TclError:
            pass
        self.deiconify()
        self.lift()
        self._aparecer(0)

    def _aparecer(self, paso):
        """Un fundido cortísimo (unos 60 ms): se nota fluido, no brusco."""
        try:
            self.attributes("-alpha", min(1.0, (paso + 1) / 3))
        except tk.TclError:
            return
        if paso < 2:
            self.after(20, lambda: self._aparecer(paso + 1))

    def _cerrar_cadena(self):
        if self._cierre_id:
            try:
                self.after_cancel(self._cierre_id)
            except tk.TclError:
                pass
        if self.hijo is not None:
            hijo, self.hijo = self.hijo, None
            hijo._cerrar_cadena()
        if self in _ABIERTOS:
            _ABIERTOS.remove(self)
        try:
            self.destroy()
        except tk.TclError:
            pass


def _invocar(menu, indice):
    try:
        menu.invoke(indice)
    except tk.TclError:
        pass


_CIERRE = {"raiz": None, "clic": None, "vigilante": None}


def cerrar_todo():
    """Cierra el menú y todos sus submenús."""
    raiz = _CIERRE["raiz"]
    for m in list(_ABIERTOS):
        try:
            m._cerrar_cadena()
        except Exception:
            pass
    _ABIERTOS.clear()
    if raiz is not None:
        try:
            if _CIERRE["clic"]:
                raiz.unbind_all("<ButtonPress>")
                _CIERRE["clic"] = None
            if _CIERRE["vigilante"]:
                raiz.after_cancel(_CIERRE["vigilante"])
        except tk.TclError:
            pass
    _CIERRE["vigilante"] = None
    _CIERRE["raiz"] = None


def _clic_fuera(e):
    """Un clic en cualquier sitio que no sea el menú lo cierra."""
    try:
        top = str(e.widget.winfo_toplevel())
    except Exception:
        top = ""
    if top not in {str(m) for m in _ABIERTOS}:
        cerrar_todo()


def _vigilar(raiz):
    """Si la aplicación pierde el foco (clic en otro programa), se cierra."""
    if not _ABIERTOS:
        return
    try:
        if raiz.focus_displayof() is None:
            cerrar_todo()
            return
    except (tk.TclError, KeyError):
        pass
    _CIERRE["vigilante"] = raiz.after(250, lambda: _vigilar(raiz))


def abrir(menu, x, y, M):
    """Muestra 'menu' (un tk.Menu ya montado) con el estilo moderno."""
    cerrar_todo()
    raiz = menu.winfo_toplevel()
    m = MenuModerno(raiz, menu, M)
    if not m.filas:
        m.destroy()
        return
    _CIERRE["raiz"] = raiz
    _ABIERTOS.append(m)
    m._colocar(x, y)
    try:
        m.focus_force()
    except tk.TclError:
        pass
    # el clic que abre el menú sigue repartiéndose entre los bindtags: si el
    # vigilante ya existiera, ese mismo clic lo cerraría al instante
    def armar():
        if m in _ABIERTOS and not _CIERRE["clic"]:
            _CIERRE["clic"] = raiz.bind_all("<ButtonPress>", _clic_fuera,
                                            add="+")
    raiz.after_idle(armar)
    _CIERRE["vigilante"] = raiz.after(400, lambda: _vigilar(raiz))


def instalar(M):
    """Cambia el menú emergente de toda la aplicación por el moderno."""
    def tk_popup(self, x, y, entry=""):
        if M.CFG.get("menu_moderno", True):
            try:
                abrir(self, x, y, M)
                return
            except Exception as exc:               # nunca dejar sin menú
                try:
                    M.log(f"Menú moderno: {exc!r}")
                except Exception:
                    pass
                cerrar_todo()
        _ORIGINAL(self, x, y, entry)
    tk.Menu.tk_popup = tk_popup
