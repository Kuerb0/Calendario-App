"""Asistente de instalacion con ventana: rellena config.json paso a paso.

Doble clic (o "python configurar.py"). Cada paso tiene un enlace "¿cómo se
hace?" que abre un .txt de la carpeta ayuda con las instrucciones.
"""
import json
import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk

BASE = os.path.dirname(os.path.abspath(__file__))
RECURSOS = getattr(sys, "_MEIPASS", BASE)      # en el .exe la ayuda va dentro; la config y el programa, junto al .exe
if getattr(sys, "frozen", False):
    BASE = os.path.dirname(os.path.abspath(sys.executable))


def _carpeta_datos():
    """Igual que en calendario.py: con app/ + datos/ la config va en ../datos."""
    if os.path.basename(BASE).lower() == "app":
        ruta = os.path.join(os.path.dirname(BASE), "datos")
        try:
            os.makedirs(ruta, exist_ok=True)
            return ruta
        except Exception:
            pass
    return BASE


CONFIG = os.path.join(_carpeta_datos(), "config.json")
CONFIG_ANTIGUA = os.path.join(os.path.dirname(BASE), "config.json")
AYUDA = os.path.join(RECURSOS, "ayuda")

COMUNIDADES = [
    ("AN", "Andalucía"), ("AR", "Aragón"), ("AS", "Asturias"),
    ("CB", "Cantabria"), ("CE", "Ceuta"), ("CL", "Castilla y León"),
    ("CM", "Castilla-La Mancha"), ("CN", "Canarias"), ("CT", "Cataluña"),
    ("EX", "Extremadura"), ("GA", "Galicia"), ("IB", "Islas Baleares"),
    ("MC", "Murcia"), ("MD", "Comunidad de Madrid"), ("ML", "Melilla"),
    ("NC", "Navarra"), ("PV", "País Vasco"), ("RI", "La Rioja"),
    ("VC", "Comunidad Valenciana"),
]

# Colores, a juego con el widget
FONDO = "#1c1c1e"
CAJA = "#26262b"
TEXTO = "#d0d0d2"
TENUE = "#8a8a90"
AZUL = "#3a86ff"


def abrir_ayuda(fichero):
    ruta = os.path.join(AYUDA, fichero)
    if not os.path.exists(ruta):
        messagebox.showinfo("Ayuda", f"No encuentro {ruta}")
        return
    try:
        os.startfile(ruta)                       # Windows
    except AttributeError:
        subprocess.Popen(["xdg-open", ruta])     # por si se prueba fuera


def carpetas_de_outlook():
    """[(nombre, resumen)] de los calendarios del Outlook clasico."""
    try:
        import pythoncom
        import win32com.client
        sys.path.insert(0, BASE)
        import calendario as cal
        pythoncom.CoInitialize()
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        salida = []
        for nombre, carpeta in cal.carpetas_de_calendario(ns):
            cuenta, total, proximas = cal.resumen_de_carpeta(carpeta)
            partes = []
            if cuenta and cuenta.lower() not in nombre.lower():
                partes.append(cuenta)
            partes.append("? elementos" if total < 0 else f"{total} elementos")
            if proximas >= 0:
                partes.append(f"{proximas} en 30 días")
            salida.append((nombre, " · ".join(partes)))
        return salida
    except BaseException:
        # BaseException: 'calendario' hace sys.exit() si aun no hay config
        return []


class Asistente(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Configurar el widget de calendario")
        self.configure(bg=FONDO)
        self.resizable(False, False)
        self.cfg = self._config_previa()
        self.carpetas = carpetas_de_outlook()

        self._cabecera()
        cuerpo = tk.Frame(self, bg=FONDO, padx=18, pady=4)
        cuerpo.pack(fill="both", expand=True)
        self._paso_origen(cuerpo)
        self._paso_festivos(cuerpo)
        self._paso_gustos(cuerpo)
        self._pie()
        self._al_cambiar_origen()

    # -- utilidades de dibujo ---------------------------------------------

    def _config_previa(self):
        for ruta in (CONFIG, CONFIG_ANTIGUA):
            try:
                with open(ruta, encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                continue
        return {}

    def _cabecera(self):
        marco = tk.Frame(self, bg=FONDO, padx=18, pady=14)
        marco.pack(fill="x")
        tk.Label(marco, text="Widget de calendario", bg=FONDO, fg=TEXTO,
                 font=("Segoe UI", 15, "bold")).pack(anchor="w")
        tk.Label(marco, bg=FONDO, fg=TENUE, font=("Segoe UI", 9),
                 justify="left",
                 text="Rellena lo de abajo y listo. Cada paso tiene un enlace "
                      "azul con las instrucciones detalladas.").pack(anchor="w")

    def _titulo(self, padre, numero, texto, ayuda=None, enlace=""):
        fila = tk.Frame(padre, bg=FONDO)
        fila.pack(fill="x", pady=(14, 4))
        tk.Label(fila, text=f"{numero}. {texto}", bg=FONDO, fg=TEXTO,
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        if ayuda:
            enlace_widget = tk.Label(fila, text=f"  {enlace}", bg=FONDO,
                                     fg=AZUL, cursor="hand2",
                                     font=("Segoe UI", 9, "underline"))
            enlace_widget.pack(side="left")
            enlace_widget.bind("<Button-1>", lambda e, f=ayuda: abrir_ayuda(f))

    def _entrada(self, padre, ancho=58):
        e = tk.Entry(padre, bg=CAJA, fg=TEXTO, insertbackground=TEXTO,
                     relief="flat", width=ancho, font=("Segoe UI", 10))
        e.pack(anchor="w", ipady=5, pady=2)
        return e

    def _rebuscar_carpetas(self):
        """Vuelve a preguntarle a Outlook, por si estaba sincronizando."""
        self.carpetas = carpetas_de_outlook()
        if getattr(self, "combo", None) is not None:
            self.combo["values"] = [f"{n}   ({c})" for n, c in self.carpetas]
            if self.carpetas:
                self.combo.current(0)

    def _pista(self, padre, texto):
        tk.Label(padre, text=texto, bg=FONDO, fg=TENUE, justify="left",
                 font=("Segoe UI", 8)).pack(anchor="w")

    # -- pasos --------------------------------------------------------------

    def _paso_origen(self, padre):
        self._titulo(padre, 1, "¿De dónde saco tus eventos?")
        self.origen = tk.StringVar(
            value="outlook" if self.carpetas else "ics")

        marco_out = tk.Frame(padre, bg=FONDO)
        marco_out.pack(fill="x", anchor="w")
        tk.Radiobutton(
            marco_out, text="Del Outlook instalado en este ordenador",
            variable=self.origen, value="outlook", bg=FONDO, fg=TEXTO,
            selectcolor=CAJA, activebackground=FONDO, activeforeground=TEXTO,
            state="normal" if self.carpetas else "disabled",
            font=("Segoe UI", 10), command=self._al_cambiar_origen
        ).pack(side="left")
        self._enlace_en(marco_out, "2-outlook-instalado.txt",
                        "¿qué es esto?")

        if self.carpetas:
            self.combo = ttk.Combobox(
                padre, width=56, state="readonly",
                values=[f"{n}   ({c})" for n, c in self.carpetas])
            self.combo.current(0)
            self.combo.pack(anchor="w", pady=(2, 0), padx=(22, 0))
            tk.Button(padre, text="Volver a buscar calendarios",
                      command=self._rebuscar_carpetas, bg=CAJA, fg=TEXTO,
                      relief="flat", font=("Segoe UI", 8),
                      activebackground=CAJA).pack(anchor="w", padx=(22, 0),
                                                  pady=(3, 0))
            self._pista(padre, "    Elige el calendario que es el tuyo. "
                               "Ventaja: se ven las etiquetas y va al instante.")
            self._pista(padre, "    Si pone 0 elementos, mira la cifra de «en "
                               "30 días»: las series cuentan como un elemento.")
        else:
            self.combo = None
            self._pista(padre, "    (No se ha encontrado el Outlook clásico "
                               "en este ordenador)")

        marco_ics = tk.Frame(padre, bg=FONDO)
        marco_ics.pack(fill="x", anchor="w", pady=(8, 0))
        tk.Radiobutton(
            marco_ics, text="De un enlace de calendario publicado (ICS)",
            variable=self.origen, value="ics", bg=FONDO, fg=TEXTO,
            selectcolor=CAJA, activebackground=FONDO, activeforeground=TEXTO,
            font=("Segoe UI", 10), command=self._al_cambiar_origen
        ).pack(side="left")
        self._enlace_en(marco_ics, "1-enlace-ics.txt",
                        "cómo consigo el enlace")

        self.ics = self._entrada(padre)
        self.ics.insert(0, self.cfg.get("ics_url", ""))
        self._pista(padre, "    Pega aquí el enlace que acaba en .ics")

    def _enlace_en(self, padre, fichero, texto):
        e = tk.Label(padre, text=f"  {texto}", bg=FONDO, fg=AZUL,
                     cursor="hand2", font=("Segoe UI", 9, "underline"))
        e.pack(side="left")
        e.bind("<Button-1>", lambda ev, f=fichero: abrir_ayuda(f))

    def _paso_festivos(self, padre):
        self._titulo(padre, 2, "Festivos", "3-festivos.txt",
                     "de dónde salen y cómo cambiarlos")
        fest = self.cfg.get("festivos", {})
        fila = tk.Frame(padre, bg=FONDO)
        fila.pack(fill="x", anchor="w")
        tk.Label(fila, text="Comunidad autónoma:", bg=FONDO, fg=TEXTO,
                 font=("Segoe UI", 10)).pack(side="left")
        self.comunidad = ttk.Combobox(
            fila, width=28, state="readonly",
            values=[nombre for _, nombre in COMUNIDADES])
        actual = fest.get("regional", {}).get("comunidad", "MD")
        self.comunidad.current(
            next((i for i, (c, _) in enumerate(COMUNIDADES) if c == actual), 13))
        self.comunidad.pack(side="left", padx=8)

        fila2 = tk.Frame(padre, bg=FONDO)
        fila2.pack(fill="x", anchor="w", pady=(6, 0))
        tk.Label(fila2, text="Municipio:", bg=FONDO, fg=TEXTO,
                 font=("Segoe UI", 10)).pack(side="left")
        self.municipio = tk.Entry(fila2, bg=CAJA, fg=TEXTO, relief="flat",
                                  insertbackground=TEXTO, width=22,
                                  font=("Segoe UI", 10))
        self.municipio.pack(side="left", padx=8, ipady=4)
        local = fest.get("local", {})
        self.municipio.insert(0, local.get("nombre", ""))

        tk.Label(fila2, text="sus 2 festivos (dd/mm):", bg=FONDO, fg=TEXTO,
                 font=("Segoe UI", 10)).pack(side="left", padx=(10, 4))
        fijos = list(local.get("fijos", {}).items())
        self.fiestas = []
        for i in range(2):
            e = tk.Entry(fila2, bg=CAJA, fg=TEXTO, relief="flat", width=7,
                         insertbackground=TEXTO, font=("Segoe UI", 10))
            e.pack(side="left", padx=3, ipady=4)
            if i < len(fijos):
                mes, dia = fijos[i][0].split("-")
                e.insert(0, f"{dia}/{mes}")
            self.fiestas.append(e)
        self._pista(padre, "    Déjalo en blanco si no quieres los locales.")

    def _paso_gustos(self, padre):
        self._titulo(padre, 3, "Cómo quieres verlo")
        fila = tk.Frame(padre, bg=FONDO)
        fila.pack(fill="x", anchor="w")

        tk.Label(fila, text="Días a la vista:", bg=FONDO, fg=TEXTO,
                 font=("Segoe UI", 10)).pack(side="left")
        self.dias = ttk.Combobox(fila, width=4, state="readonly",
                                 values=["1", "3", "5", "7", "14", "21"])
        self.dias.set(str(self.cfg.get("dias_visibles", 7)))
        self.dias.pack(side="left", padx=(6, 16))

        tk.Label(fila, text="Tema:", bg=FONDO, fg=TEXTO,
                 font=("Segoe UI", 10)).pack(side="left")
        self.tema = ttk.Combobox(fila, width=14, state="readonly",
                                 values=["Oscuro", "Claro", "Según Windows"])
        self.tema.set({"claro": "Claro", "auto": "Según Windows"}.get(
            self.cfg.get("tema", "oscuro"), "Oscuro"))
        self.tema.pack(side="left", padx=6)

        fila2 = tk.Frame(padre, bg=FONDO)
        fila2.pack(fill="x", anchor="w", pady=(8, 0))
        tk.Label(fila2, text="Avisarme con esta antelación (minutos):",
                 bg=FONDO, fg=TEXTO, font=("Segoe UI", 10)).pack(side="left")
        self.avisos = tk.Entry(fila2, bg=CAJA, fg=TEXTO, relief="flat", width=12,
                               insertbackground=TEXTO, font=("Segoe UI", 10))
        self.avisos.pack(side="left", padx=8, ipady=4)
        self.avisos.insert(0, ", ".join(
            str(x) for x in self.cfg.get("avisos_minutos", [60, 10])))
        self._pista(padre, "    Separados por comas. Déjalo vacío para no "
                           "recibir avisos.")

    def _pie(self):
        marco = tk.Frame(self, bg=FONDO, padx=18, pady=16)
        marco.pack(fill="x")
        self.aviso = tk.Label(marco, text="", bg=FONDO, fg="#ff6b5e",
                              font=("Segoe UI", 9))
        self.aviso.pack(anchor="w", pady=(0, 6))
        botones = tk.Frame(marco, bg=FONDO)
        botones.pack(fill="x")
        tk.Button(botones, text="Guardar y arrancar el widget",
                  command=self.guardar, bg=AZUL, fg="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=16, pady=8,
                  cursor="hand2").pack(side="left")
        tk.Button(botones, text="Cancelar", command=self.destroy, bg=CAJA,
                  fg=TEXTO, relief="flat", font=("Segoe UI", 10), padx=14,
                  pady=8, cursor="hand2").pack(side="left", padx=8)
        e = tk.Label(botones, text="qué puedo hacer después", bg=FONDO, fg=AZUL,
                     cursor="hand2", font=("Segoe UI", 9, "underline"))
        e.pack(side="right")
        e.bind("<Button-1>", lambda ev: abrir_ayuda("4-despues.txt"))

    # -- logica -------------------------------------------------------------

    def _al_cambiar_origen(self):
        usa_ics = self.origen.get() == "ics"
        self.ics.configure(state="normal" if usa_ics else "disabled")
        if self.combo is not None:
            self.combo.configure(state="disabled" if usa_ics else "readonly")

    def guardar(self):
        cfg = dict(self.cfg)
        if self.origen.get() == "outlook" and self.carpetas:
            nombre = self.carpetas[self.combo.current()][0]
            cfg["origen"] = "outlook"
            cfg["carpetas_outlook"] = [{"nombre": "Mi calendario",
                                        "buscar": nombre, "color": "#3a86ff",
                                        "activo": True}]
        else:
            url = self.ics.get().strip()
            if not url.startswith("http"):
                self.aviso.config(
                    text="Falta el enlace ICS (o elige leer del Outlook "
                         "instalado). Pulsa el enlace azul si no sabes de "
                         "dónde sacarlo.")
                return
            cfg["origen"] = "ics"
            cfg["ics_url"] = url
            cfg["calendarios"] = [{"nombre": "Mi calendario", "url": url,
                                   "color": "#3a86ff", "activo": True}]

        codigo = COMUNIDADES[self.comunidad.current()][0]
        festivos = cfg.setdefault("festivos", {})
        festivos.setdefault("nacional", {}).update(
            {"activo": True, "color": "#e5533d", "nombre": "Festivos nacionales"})
        festivos.setdefault("regional", {}).update(
            {"activo": True, "comunidad": codigo,
             "color": festivos.get("regional", {}).get("color", "#a06bff"),
             "nombre": COMUNIDADES[self.comunidad.current()][1]})

        municipio = self.municipio.get().strip()
        fijos = {}
        for e in self.fiestas:
            texto = e.get().strip().replace("-", "/").replace(".", "/")
            if not texto:
                continue
            try:
                dia, mes = (int(x) for x in texto.split("/")[:2])
                fijos[f"{mes:02d}-{dia:02d}"] = f"Fiesta local ({municipio})"
            except ValueError:
                self.aviso.config(text=f"La fecha '{texto}' no vale. "
                                       "Escríbela como dd/mm, por ejemplo 6/8.")
                return
        festivos.setdefault("local", {}).update(
            {"activo": bool(municipio and fijos), "nombre": municipio or "Locales",
             "color": festivos.get("local", {}).get("color", "#2ec4a6"),
             "fijos": fijos, "extra": festivos.get("local", {}).get("extra", {})})

        cfg["dias_visibles"] = int(self.dias.get())
        cfg["tema"] = {"Claro": "claro", "Según Windows": "auto"}.get(
            self.tema.get(), "oscuro")
        avisos = []
        for trozo in self.avisos.get().replace(";", ",").split(","):
            trozo = trozo.strip()
            if trozo.isdigit():
                avisos.append(int(trozo))
        cfg["avisos_minutos"] = sorted(avisos, reverse=True)
        cfg.setdefault("zona_horaria", "Europe/Madrid")

        try:
            with open(CONFIG, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
        except Exception as exc:
            self.aviso.config(text=f"No se ha podido guardar: {exc}")
            return

        self.destroy()
        try:
            if getattr(sys, "frozen", False):
                subprocess.Popen([sys.executable], cwd=BASE)       # el .exe vuelve a abrirse, ya con config.json
            else:
                subprocess.Popen([sys.executable.replace("python.exe", "pythonw.exe"),
                                  os.path.join(BASE, "calendario.py")], cwd=BASE)
        except Exception:
            pass


if __name__ == "__main__":
    Asistente().mainloop()
