"""
Widget de calendario de Outlook para el escritorio de Windows.

- Lee el calendario desde un enlace ICS publicado en Outlook web.
- Dibuja una rejilla de dias con horas, tema oscuro.
- Ventana sin bordes, siempre visible pero POR DEBAJO de cualquier app
  (y por ENCIMA del fondo de escritorio, para que no se pierda).
- Notificaciones nativas de Windows con antelacion configurable.
- Cachea los eventos en disco: al arrancar muestra algo aunque no haya red.

Configuracion: config.json (se crea solo en el primer arranque).

Barra superior:  ‹ ›  navegar     hoy  volver al dia actual
                 7 / 14  dias que se ven
                 Fest  festivos: nacionales, Madrid y Alcala (se apagan uno a uno)
                 Cal / Lista / Ambas  vista
                 (flecha circular) actualizar    (raya) ocultar
Mover: arrastrar con el raton.  Clic derecho: mismas opciones + salir.
Icono junto al reloj: clic para mostrar u ocultar.
"""

import copy
import datetime as dt
from functools import lru_cache
import html
import time
import json
import os
import pickle
import re
import shutil
import sys
import threading
import math
import tkinter as tk
import unicodedata
import webbrowser
from zoneinfo import ZoneInfo

import requests
import icalendar
import recurring_ical_events

import sincronizar

VERSION = "2.4.0"

# --------------------------------------------------------------------------
# Configuracion
# --------------------------------------------------------------------------

if getattr(sys, "frozen", False):
    # compilado con PyInstaller: los datos van junto al .exe, no en el temporal
    BASE = os.path.dirname(os.path.abspath(sys.executable))
    RECURSOS = getattr(sys, "_MEIPASS", BASE)
else:
    BASE = RECURSOS = os.path.dirname(os.path.abspath(__file__))

def carpeta_local():
    """Sitio propio de este ordenador, fuera de OneDrive o del repositorio.

    La caché, el registro y la señal cambian a cada momento y son distintos en
    cada equipo: si viven en una carpeta sincronizada, OneDrive se pasa el día
    subiéndolos y acaba creando copias en conflicto.
    """
    try:
        raiz = (os.environ.get("LOCALAPPDATA")
                or os.path.join(os.path.expanduser("~"), ".cache"))
        ruta = os.path.join(raiz, "WidgetCalendario")
        os.makedirs(ruta, exist_ok=True)
        return ruta
    except Exception:
        return BASE


LOCAL = carpeta_local()
def carpeta_datos():
    """Donde viven config.json y perfiles/: tus datos, aparte del programa.

    Con la estructura ordenada (app/ + datos/) van en ../datos. Compilado en un
    .exe (que vive en local) van en OneDrive\\Calendario\\datos si ya existe, para
    seguir compartiéndose entre equipos; si no, junto al .exe. Suelto: junto a él.

    Con una carpeta de sincronización elegida (la del instalador, la del Panel de control o la de Post-It) mandan
    «<carpeta>\\Calendario»: ahí viven config.json y perfiles/ para todos los equipos. La primera vez se llevan los de antes.
    """
    sync = sincronizar.carpeta_elegida(LOCAL)
    if sync:
        antes = []
        if os.path.basename(BASE).lower() == "app" and not getattr(sys, "frozen", False):
            antes.append(os.path.join(os.path.dirname(BASE), "datos"))
        antes.append(sincronizar.datos_compartidos())
        ruta = sincronizar.datos_en_sync(sync, antes)
        if ruta:
            return ruta
    if getattr(sys, "frozen", False):
        return sincronizar.datos_compartidos(espera=40) or BASE
    if (os.path.basename(BASE).lower() == "app"
            and not getattr(sys, "frozen", False)):
        ruta = os.path.join(os.path.dirname(BASE), "datos")
        try:
            os.makedirs(ruta, exist_ok=True)
            return ruta
        except Exception:
            return BASE
    return BASE


DATOS = carpeta_datos()
CONFIG_PATH = os.path.join(DATOS, "config.json")     # este sí se comparte
EQUIPO_PATH = os.path.join(LOCAL, "equipo.json")     # lo de ESTE ordenador: no se comparte
CACHE_PATH = os.path.join(LOCAL, "cache.pkl")
LOG_PATH = os.path.join(LOCAL, "registro.txt")
ICONO_PATH = (os.path.join(BASE, "icono.ico") if os.path.exists(
    os.path.join(BASE, "icono.ico")) else os.path.join(RECURSOS, "icono.ico"))
SENAL_PATH = os.path.join(LOCAL, "alternar.senal")
CACHE_VERSION = 7


def pedir_alternar():
    """Deja una senal para que la instancia que ya corre se muestre u oculte."""
    try:
        with open(SENAL_PATH, "w") as f:
            f.write("1")
        return True
    except Exception:
        return False


def consumir_alternar():
    """True si habia una senal pendiente (y la borra)."""
    try:
        if os.path.exists(SENAL_PATH):
            os.remove(SENAL_PATH)
            return True
    except OSError:
        pass
    return False


def log(mensaje):
    """Deja constancia en registro.txt (con pythonw no hay consola que mirar)."""
    try:
        if os.path.exists(LOG_PATH) and os.path.getsize(LOG_PATH) > 200_000:
            os.remove(LOG_PATH)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}  {mensaje}\n")
    except Exception:
        pass

DEFAULTS = {
    # se rellena solo a partir de "ics_url" la primera vez
    "calendarios": [],
    "ics_url": "",
    "zona_horaria": "Europe/Madrid",
    "refresco_minutos": 30,
    # Post-It <-> Calendario por una carpeta compartida (sincronizar.py)
    "sincronizar_postit": True,
    "carpeta_compartida": "",
    "color_postit": "#e0b400",
    "avisos_minutos": [60, 10],
    "hora_inicio": 8,
    "hora_fin": 22,
    "dias_visibles": 7,
    # "hoy" -> ventana deslizante que empieza en el dia actual
    # "lunes" -> semana natural, empieza siempre en lunes
    "inicio_ventana": "hoy",
    # cuantos dias se descargan por delante y por detras (para poder navegar)
    "dias_cache_atras": 7,
    "dias_cache_adelante": 90,
    # "rejilla" (calendario por horas), "lista" (agenda) o "ambas"
    "vista": "rejilla",
    # con 14 dias las columnas se estrechan para que el widget no crezca
    "ancho_fijo": True,
    # lineas finas de la rejilla (las de semana y mes se mantienen)
    "lineas_rejilla": True,
    "division_minutos": 60,
    "etiquetas_recientes": [],
    "hechos": [],
    # habitos: casillas por dia que se marcan a mano (ver PanelHabitos)
    # cada uno: {"id": ..., "nombre": ..., "icono": ..., "color": ...}
    "habitos": [],
    # marcas: {"id del habito": ["2026-08-29", ...]}
    "habitos_hechos": {},
    # la rejilla de habitos: cuantos dias se ven antes y despues de hoy
    "habitos_dias_atras": 10,
    "habitos_dias_delante": 4,
    "habitos_activo": True,
    "habitos_en_mes": True,      # puntitos de habitos en la vista de mes
    "tachar_pasados": False,
    "plantillas": [],
    "tareas": [],
    # cumpleaños guardados a mano: [{"nombre":..,"dia":..,"mes":..}, ...]
    "cumpleanos": [],
    "color_cumpleanos": "#a855f7",
    "dibujar_huecos": True,
    "lugares_recientes": [],
    "duracion_minutos": 60,
    "lineas_lista": 2,
    "lado_lista": "derecha",
    # anchura de la franja agarrable del divisor entre calendario y lista;
    # la raya que se ve es fina, esto es solo la zona sensible al raton
    "grosor_divisor": 11,
    "guardar_en_segundo_plano": True,
    "leer_notas": False,
    "barra_proximos": True,
    # botones de la barra de arriba que el usuario ha quitado (extras.py)
    "barra_ocultos": [],
    "arrastrar_eventos": True,
    "linea_ahora_lista": True,
    "avisos_flotantes": True,
    "avisos_solo_etiqueta": "",
    "aviso_flotante_segundos": 12,
    # apaga/enciende TODOS los avisos (toast, flotante, sonido) de un golpe
    "notificaciones_activas": True,
    # avisos antes de cada evento: apagados. Los eventos que de verdad
    # importan se marcan uno a uno (clic derecho -> «Avisarme de este»),
    # y esos avisan aunque esto siga apagado.
    "avisos_eventos": False,
    "avisar_de": [],
    # alarmas con sonido por evento: {"uid|fecha": [minutos antes, ...]}
    "alarmas": {},
    # sonido de avisos y alarmas: "" (el de siempre), "ninguno" o ruta a un .wav
    "sonido_aviso": "",
    # recordatorio de marcar los habitos
    "recordar_habitos": True,
    "hora_recordatorio": 21,
    # atenuar los eventos que ya han terminado
    "atenuar_pasado": True,
    # cada cuanto se reintenta cuando no hay red (minutos)
    "reintento_minutos": 1,
    # tamano general del widget (0.7 a 2.0); Ctrl + rueda lo cambia en caliente
    "escala": 1.0,
    "escala_letra": 1.0,
    "perfil": "",
    "ancho_lista": 320,
    "alto_celda_mes": 88,
    "alto_max_lista": 620,
    # festivos: se pueden apagar por separado desde el boton "Fest"
    "festivos": {
        "nacional": {"activo": True, "color": "#e5533d",
                     "nombre": "Festivos nacionales"},
        # "comunidad" es el codigo ISO de tu autonomia: MD Madrid, CT Cataluna,
        # AN Andalucia, VC Valencia, PV Pais Vasco, GA Galicia, CL Castilla y
        # Leon, CM Castilla-La Mancha, AR Aragon, AS Asturias, CN Canarias,
        # CB Cantabria, EX Extremadura, IB Baleares, MC Murcia, NC Navarra,
        # RI La Rioja, CE Ceuta, ML Melilla
        "regional": {"activo": True, "color": "#a06bff",
                     "comunidad": "MD", "nombre": "Comunidad de Madrid"},
        # los dos festivos de tu municipio: fechas fijas cada anio (mes-dia)
        # y excepciones puntuales en "extra" (anio-mes-dia)
        "local": {
            "activo": False, "color": "#2ec4a6",
            "nombre": "",          # lo rellena cada uno en el asistente
            "fijos": {},
            "extra": {},
        },
        # se pueden seguir varios municipios a la vez: cada uno se enciende
        # y se apaga por separado desde el boton «Fest».
        # Ojo: si la fecha fija cae en domingo el ayuntamiento la traslada;
        # esa excepcion se apunta en "extra" con el año completo.
        "municipios": [
            {"clave": "alcala", "nombre": "Alcalá de Henares",
             "activo": True, "color": "#2ec4a6",
             "fijos": {"08-06": "Santos Niños",
                       "10-09": "Bautismo de Cervantes"},
             "extra": {}},
            {"clave": "madrid", "nombre": "Madrid capital",
             "activo": False, "color": "#f08c2e",
             "fijos": {"05-15": "San Isidro",
                       "11-09": "Nuestra Señora de la Almudena"},
             "extra": {}},
        ],
    },
    # color por etiqueta (categoria) de Outlook: "Trabajo": "#f08c2e"
    "colores_categoria": {},
    # lista de to-dos independiente del calendario (a salvo de los refrescos)
    # cada uno: {"texto": str, "hecho": bool, "etiqueta": str}
    "todos": [],
    # columna de to-dos acoplada a la derecha (como la lista)
    "todos_visible": False,
    # donde vive cada panel secundario: "flotante" (ventana aparte, por
    # defecto), "izquierda" o "derecha" (anclado, sigue a la ventana
    # principal). Claves: correo, habitos, cumpleanos, plantillas, tareas
    "ubicacion_vistas": {},
    "ancho_todos": 210,
    # hora a la que avisar de los eventos de todo el dia (null = no avisar)
    "aviso_todo_dia": "09:00",
    # resumen del dia en cuanto arranca el widget
    "resumen_al_arrancar": True,
    # carpetas de Outlook a leer, cada una con su nombre y color:
    # [{"nombre": "Personal", "buscar": "usuario@outlook.com/Calendario",
    #   "color": "#3a86ff", "activo": true}]
    "carpetas_outlook": [],
    "carpeta_outlook": "",      # formato antiguo, una sola carpeta
    # carpetas de CORREO a leer (no de calendario): misma idea que las de
    # arriba, [{"nombre":.., "buscar":.., "color":.., "activo":true}].
    # vacio = solo la bandeja de entrada de la cuenta por defecto.
    "carpetas_correo": [],
    "correo_activo": True,
    "correo_limite_lista": 10,      # cuantos correos se piden por carpeta
    "correo_refresco_minutos": 3,
    "correo_marcar_leido_al_abrir": True,
    # columna de correo dentro del widget (correo_columna.py)
    "correo_visible": False,
    "ancho_correo": 260,
    "correo_filtro": None,          # una sola carpeta a la vista (None = todas)
    "correo_excluidas": [],         # carpetas que no salen en la columna
    "correo_solo_sin_leer": False,
    "correo_filas_max": 8,
    "correo_chips_max": 3,
    "barra_zoom": 1.5,
    "columnas_dentro": True,
    "menu_moderno": True,
    "correo_avisos": True,      # avisar (Windows + flotante) si sube el no leídos
    # de donde salen los eventos:
    #   "ics"     -> los enlaces publicados (por defecto)
    #   "outlook" -> el Outlook clasico instalado (trae las categorias y va al día)
    #   "auto"    -> Outlook si responde, si no los ICS
    "origen": "ics",
    # que el doble clic en un hueco cree un evento (desactivado por defecto)
    "doble_clic_crea": False,
    # como se crean los eventos:
    #   "auto"/"web" -> Outlook web con el evento ya rellenado (recomendado)
    #   "nuevo"      -> abre la app del Outlook nuevo (sin rellenar: no lo admite)
    #   "outlook"    -> abre el Outlook clasico de escritorio, ya rellenado
    "crear_con": "ficha",
    # atajo global para mostrar u ocultar (null para desactivarlo)
    "atajo_global": "ctrl+alt+c",
    # ver solo una etiqueta o un calendario (null = ver todo)
    "filtro": None,
    # evento fijado abajo con su cuenta atras
    "anclado": None,
    # donde vive la ventana: "fondo" (debajo de las apps, sobre el escritorio),
    # "normal" (una ventana mas) o "encima" (siempre visible sobre todo)
    "nivel_ventana": "fondo",
    # que el widget solo se vea en el escritorio virtual donde se activó
    "solo_un_escritorio": True,
    "escritorio_id": None,
    # raya blanca a la izquierda de cada lunes (si no, la tenue de siempre)
    "semana_blanca": True,
    "posicion_x": 40,
    "posicion_y": 40,
    # posicion recordada para cada combinacion de pantallas
    "posiciones": {},
    "ancho_columna": 118,
    "alto_hora": 30,
    # "horizontal" (dias en columnas) o "vertical" (dias en filas)
    "orientacion": "horizontal",
    # medidas propias del modo vertical, independientes de las de arriba
    "ancho_hora": 58,
    "alto_fila_v": 62,
    "opacidad": 0.92,
    "colores": {
        "fondo": "#1c1c1e",
        "rejilla": "#3a3a3c",
        "texto": "#d0d0d2",
        "texto_tenue": "#7a7a7e",
        "evento": "#3a86ff",
        "evento_texto": "#ffffff",
        "ahora": "#ff453a",
        "hoy": "#2c2c2e",
        "finde": "#242426",
        "boton": "#2f2f33",
        "boton_raton": "#40404a",
        "semana": "#54545c",
    },
    # "oscuro", "claro" o "auto" (sigue al tema de Windows)
    "tema": "oscuro",
    # solo hace falta cambiar lo que difiere del tema oscuro
    "colores_claro": {
        "fondo": "#f4f4f6",
        "rejilla": "#d2d2d8",
        "texto": "#26262a",
        "texto_tenue": "#76767e",
        "evento": "#2f6fe0",
        "evento_texto": "#ffffff",
        "ahora": "#d92b1c",
        "hoy": "#e6edfb",
        "finde": "#eaeaee",
        "boton": "#e2e2e8",
        "boton_raton": "#cfcfd8",
        "semana": "#b8b8c2",
    },
}


# --------------------------------------------------------------------------
# Perfiles: los datos de la cuenta de cada persona, aparte del resto
# --------------------------------------------------------------------------

DIR_PERFILES = os.path.join(DATOS, "perfiles")


def migrar_datos_antiguos():
    """Trae config.json y perfiles/ de la estructura antigua (todo junto).

    Antes el programa y tus datos vivían en la misma carpeta; ahora el programa
    está en app/ y los datos en datos/. Se copian una sola vez y los originales
    se renombran a «.migrado» para que no se editen por error.
    """
    if DATOS == BASE:
        return
    raiz = os.path.dirname(BASE)
    try:
        viejo = os.path.join(raiz, "config.json")
        if os.path.exists(viejo) and not os.path.exists(CONFIG_PATH):
            shutil.copy2(viejo, CONFIG_PATH)
            os.replace(viejo, viejo + ".migrado")
            log("Configuración traída de la carpeta antigua a datos/")
        viejos = os.path.join(raiz, "perfiles")
        if os.path.isdir(viejos) and not os.path.exists(DIR_PERFILES):
            shutil.copytree(viejos, DIR_PERFILES)
            os.replace(viejos, viejos + ".migrado")
            log("Perfiles traídos de la carpeta antigua a datos/")
    except Exception as exc:
        log(f"No se pudieron migrar los datos antiguos: {exc}")


migrar_datos_antiguos()
# lo unico que cambia de una persona a otra; el resto de ajustes son comunes
CLAVES_PERFIL = ("origen", "ics_url", "calendarios", "carpetas_outlook",
                 "carpeta_outlook", "carpetas_correo", "festivos",
                 "zona_horaria")
# lo que depende de las pantallas de ESTE ordenador: vive en equipo.json (local). Mover la ventana ya no reescribe el
# config.json que usan los dos PCs, que es lo que hacia que OneDrive creara copias en conflicto
CLAVES_EQUIPO = ("posicion_x", "posicion_y", "posiciones", "anclado", "escala", "escala_letra", "nivel_ventana",
                 "escritorio_id", "solo_un_escritorio", "atajo_global", "ancho_columna", "alto_hora", "ancho_hora",
                 "alto_fila_v", "alto_celda_mes", "ancho_lista", "alto_max_lista", "ancho_todos", "ancho_correo",
                 "barra_zoom")


def leer_equipo():
    """Los ajustes de este ordenador. Vacio si aun no hay equipo.json: valen los que traiga config.json."""
    try:
        with open(EQUIPO_PATH, encoding="utf-8") as f:
            return {k: v for k, v in json.load(f).items() if k in CLAVES_EQUIPO}
    except Exception:
        return {}


def nombre_de_fichero(nombre):
    """'Pepe Pérez' -> 'pepe-perez.json', sin sorpresas en el disco."""
    plano = "".join(c for c in unicodedata.normalize("NFD", str(nombre).lower())
                    if unicodedata.category(c) != "Mn")
    limpio = "".join(c if c.isalnum() else "-" for c in plano.strip())
    limpio = "-".join(t for t in limpio.split("-") if t)
    return (limpio or "perfil") + ".json"


def ruta_perfil(nombre):
    return os.path.join(DIR_PERFILES, nombre_de_fichero(nombre))


def listar_perfiles():
    """Nombres de los perfiles guardados, por orden alfabético."""
    try:
        nombres = []
        for f in sorted(os.listdir(DIR_PERFILES)):
            if not f.endswith(".json"):
                continue
            try:
                with open(os.path.join(DIR_PERFILES, f), encoding="utf-8") as h:
                    nombres.append(json.load(h).get("nombre") or f[:-5])
            except Exception:
                nombres.append(f[:-5])
        return nombres
    except FileNotFoundError:
        return []


def guardar_perfil(nombre, cfg=None):
    """Copia los datos de la cuenta que hay ahora al perfil indicado."""
    cfg = CFG if cfg is None else cfg
    if not nombre:
        return False
    datos = {"nombre": nombre}
    for clave in CLAVES_PERFIL:
        if clave in cfg:
            datos[clave] = cfg[clave]
    try:
        os.makedirs(DIR_PERFILES, exist_ok=True)
        with open(ruta_perfil(nombre), "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)
        log(f"Perfil guardado: {nombre}")
        return True
    except Exception as exc:
        log(f"No se pudo guardar el perfil '{nombre}': {exc}")
        return False


def leer_perfil(nombre):
    try:
        with open(ruta_perfil(nombre), encoding="utf-8") as f:
            return json.load(f)
    except Exception as exc:
        log(f"No se pudo leer el perfil '{nombre}': {exc}")
        return None


def aplicar_perfil_activo(cfg):
    """Si hay un perfil elegido, sus datos de cuenta mandan sobre config.json."""
    nombre = cfg.get("perfil")
    if not nombre:
        return False
    datos = leer_perfil(nombre)
    if not datos:
        return False
    for clave in CLAVES_PERFIL:
        if clave in datos:
            cfg[clave] = datos[clave]
    return True


def asistente_si_es_exe():
    """Compilado en .exe no hay consola donde pegar el enlace: se abre el asistente de configuracion y el .exe se reabre al terminar."""
    if getattr(sys, "frozen", False):
        import configurar
        configurar.Asistente().mainloop()
        sys.exit(0)


def cargar_config():
    if not os.path.exists(CONFIG_PATH):
        asistente_si_es_exe()
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(DEFAULTS, f, indent=2, ensure_ascii=False)
        print("Se ha creado config.json. Pega ahi tu enlace ICS y vuelve a ejecutar.")
        sys.exit(0)
    with open(CONFIG_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg.update(leer_equipo())       # lo de este ordenador manda sobre lo que dejo el otro en config.json
    # completa claves que falten si actualizas el script
    for k, v in DEFAULTS.items():
        cfg.setdefault(k, v)
    for k, v in DEFAULTS["colores"].items():
        cfg["colores"].setdefault(k, v)
    # nombres antiguos de los grupos de festivos
    festivos = cfg.setdefault("festivos", {})
    for viejo, nuevo in (("madrid", "regional"), ("alcala", "local")):
        if viejo in festivos and nuevo not in festivos:
            festivos[nuevo] = festivos.pop(viejo)
    for clave, valores in DEFAULTS["festivos"].items():
        if not isinstance(valores, dict):
            # "municipios" es una lista, no un grupo: se copia entera solo
            # si el usuario todavia no la tiene
            festivos.setdefault(clave, valores)
            continue
        for k, v in valores.items():
            festivos.setdefault(clave, {}).setdefault(k, v)
    # un solo enlace antiguo -> lista de calendarios
    if not cfg.get("calendarios"):
        cfg["calendarios"] = [{"nombre": "Mi calendario",
                               "url": cfg.get("ics_url", ""),
                               "color": cfg["colores"]["evento"],
                               "activo": True}]
    for c in cfg["calendarios"]:
        c.setdefault("activo", True)
        c.setdefault("color", cfg["colores"]["evento"])
        c.setdefault("nombre", "Calendario")
    aplicar_perfil_activo(cfg)
    if cfg.get("origen") != "outlook" and not any(
            c.get("url") for c in cfg["calendarios"]):
        asistente_si_es_exe()
        print("Falta la URL del calendario en config.json.")
        sys.exit(1)
    return cfg


CFG = cargar_config()
TZ = ZoneInfo(CFG["zona_horaria"])
C = {}                      # paleta en uso, la rellena aplicar_tema()


def windows_en_claro():
    """True si Windows esta en tema claro (si no se sabe, False)."""
    try:
        import winreg
        clave = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        valor, _ = winreg.QueryValueEx(clave, "AppsUseLightTheme")
        return bool(valor)
    except Exception:
        return False


def tema_claro_activo():
    tema = CFG.get("tema", "oscuro")
    return tema == "claro" or (tema == "auto" and windows_en_claro())


def color_semana():
    """Raya que marca el inicio de semana: blanca (o la tinta, en tema claro)."""
    if not CFG.get("semana_blanca", True):
        return C.get("semana", "#54545c")
    return C["texto"] if tema_claro_activo() else "#ffffff"


def escritorios_virtuales():
    """(id del escritorio actual, [ids de todos]) en hexadecimal; (None, []) si no se sabe."""
    try:
        import winreg
        with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops") as k:
            try:
                todos = winreg.QueryValueEx(k, "VirtualDesktopIDs")[0]
            except OSError:
                todos = b""
            try:
                actual = winreg.QueryValueEx(k, "CurrentVirtualDesktop")[0]
            except OSError:     # Windows no lo escribe hasta el primer cambio
                actual = todos[:16]
        return actual.hex(), [todos[i:i + 16].hex() for i in range(0, len(todos), 16)]
    except Exception:
        return None, []


def aplicar_tema():
    """Rellena C con la paleta que toque, sin perder los colores del config."""
    C.clear()
    C.update(DEFAULTS["colores"])
    C.update(CFG.get("colores", {}))
    if tema_claro_activo():
        C.update(DEFAULTS["colores_claro"])
        C.update(CFG.get("colores_claro", {}))


aplicar_tema()
DIAS = ["Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"]
# La inicial de miercoles es X, que la M ya se la lleva el martes. No se puede
# sacar de DIAS con [0] porque saldrian dos emes seguidas.
INICIALES_DIA = ["L", "M", "X", "J", "V", "S", "D"]
MESES = ["ene", "feb", "mar", "abr", "may", "jun",
         "jul", "ago", "sep", "oct", "nov", "dic"]
MESES_LARGOS = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                "julio", "agosto", "septiembre", "octubre", "noviembre",
                "diciembre"]


def columnas_de_semana(ini, dias):
    """Columnas (menos la primera) en las que empieza una semana nueva."""
    return [i for i in range(1, dias)
            if (ini + dt.timedelta(days=i)).weekday() == 0]


def grupos_de_mes(ini, dias):
    """Tramos de columnas consecutivas del mismo mes: [((anio, mes), c1, c2)]."""
    grupos = []
    for i in range(dias):
        fecha = (ini + dt.timedelta(days=i)).date()
        clave = (fecha.year, fecha.month)
        if grupos and grupos[-1][0] == clave:
            grupos[-1][2] = i
        else:
            grupos.append([clave, i, i])
    return [tuple(g) for g in grupos]

# Medidas base (a escala 1.0). aplicar_escala() las recalcula.
ESCALA = 1.0
LETRA = 1.0             # tamano del texto, aparte del tamano general
MARGEN_IZQ = 46         # ancho de la columna de las horas
BARRA = 24              # franja superior con los botones
BZ = 1.5                # zoom de la barra superior (botones e iconos)
CABECERA = 52           # barra + fila con los nombres de los dias
Y_LISTA = 30            # donde empieza el contenido en la vista de lista
ALTO_CHIP = 18          # alto de cada fila de la banda de "todo el dia"
ALTO_ESPECIAL = 16     # alto de cada recuadro de festivo/cumpleaños
BANDA_MINIMA = 22       # alto de la banda cuando no hay eventos de todo el dia
ALTO_FILA_DIA = 22      # vista de lista: alto de la cabecera de cada dia
ALTO_FILA_EV = 19       # vista de lista: alto de cada evento
ALTO_ESTADO = 16        # franja inferior con la hora de actualizacion
ALTO_MESES = 15         # franja con el nombre del mes bajo las columnas


def aplicar_escala():
    """Recalcula todas las medidas segun 'escala' del config."""
    global ESCALA, LETRA, MARGEN_IZQ, BARRA, CABECERA, Y_LISTA, ALTO_CHIP, BZ
    global BANDA_MINIMA, ALTO_FILA_DIA, ALTO_FILA_EV, ALTO_ESTADO, ALTO_ESPECIAL
    try:
        ESCALA = max(0.7, min(2.0, float(CFG.get("escala", 1.0))))
    except (TypeError, ValueError):
        ESCALA = 1.0
    try:
        LETRA = max(0.7, min(1.8, float(CFG.get("escala_letra", 1.0))))
    except (TypeError, ValueError):
        LETRA = 1.0
    try:
        BZ = max(1.0, min(2.5, float(CFG.get("barra_zoom", 1.5))))
    except (TypeError, ValueError):
        BZ = 1.5
    # las medidas que llevan texto dentro crecen tambien con la letra
    e = ESCALA
    t = e * max(1.0, LETRA)
    MARGEN_IZQ = round(46 * t)
    BARRA = round((36 if CFG.get("cuentas_atras") else 24) * t * BZ)
    CABECERA = BARRA + round(28 * t)
    Y_LISTA = BARRA + round(6 * t)
    ALTO_CHIP = round(18 * t)
    ALTO_ESPECIAL = round(16 * t)
    BANDA_MINIMA = round(22 * t)
    ALTO_FILA_DIA = round(22 * t)
    ALTO_FILA_EV = round(19 * t)
    ALTO_ESTADO = round(16 * t)
    global ALTO_MESES
    ALTO_MESES = round(15 * e)


def mezclar(color, fondo, peso=0.45):
    """Mezcla dos colores hex: peso 0 = fondo, 1 = color."""
    try:
        a = [int(color[i:i + 2], 16) for i in (1, 3, 5)]
        b = [int(fondo[i:i + 2], 16) for i in (1, 3, 5)]
    except (ValueError, IndexError, TypeError):
        return color
    return "#" + "".join(f"{int(x * peso + y * (1 - peso)):02x}"
                         for x, y in zip(a, b))


def letra(tam, estilo=None):
    """Fuente ya escalada (tamano general x tamano de letra)."""
    tam = max(6, round(tam * ESCALA * LETRA))
    return ("Segoe UI", tam) if estilo is None else ("Segoe UI", tam, estilo)


aplicar_escala()
NIVELES = {"fondo": "Fondo", "normal": "Normal", "encima": "Encima"}
MAX_DIAS = 365          # tope de dias que se pueden mostrar a la vez
TOPE_NAVEGACION = 3650  # hasta donde se puede viajar (dias) desde hoy
BORDE = 8               # margen sensible para redimensionar con el raton


def fecha_corta(momento):
    """'Mar 18/08' sin depender del idioma del sistema."""
    return f"{DIAS[momento.weekday()]} {momento.day:02d}/{momento.month:02d}"


def escribir_atomico(ruta, datos):
    """Escribe entero o no escribe: nunca deja el fichero a medias.

    Se vuelca a un temporal al lado, se fuerza a disco y se renombra encima.
    os.replace es atomico en Windows y en POSIX, asi que quien lea el fichero
    ve la version vieja o la nueva, jamas media. Importa porque config.json
    vive en carpeta sincronizada y un fichero truncado deja la app sin abrir.
    """
    tmp = f"{ruta}.tmp"
    modo, kw = (("wb", {}) if isinstance(datos, bytes)
                else ("w", {"encoding": "utf-8"}))
    try:
        with open(tmp, modo, **kw) as f:
            f.write(datos)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, ruta)
    except Exception:
        try:
            os.remove(tmp)            # que no queden .tmp sueltos
        except OSError:
            pass
        raise


# Avisos y alarmas: viajan por OneDrive entre ordenadores. Cada copia del
# programa tiene CFG en memoria, asi que sin esto guardar desde un PC pisaba lo
# que habia apuntado el otro, y nadie se enteraba de los cambios ajenos.
CLAVES_AVISOS = ("alarmas", "avisar_de", "hechos", "avisos_eventos",
                 "avisos_solo_etiqueta", "avisos_minutos", "aviso_todo_dia",
                 "notificaciones_activas", "correo_avisos", "recordar_habitos",
                 "hora_recordatorio")
_SYNC = {"base": {}, "mtime": None}    # lo ultimo que vimos en disco


def _fusionar(base, local, disco):
    """Tres vias: aplica sobre `disco` lo que `local` cambio respecto a `base`."""
    if local == base:
        return disco
    if disco == base:
        return local
    if isinstance(local, dict) and isinstance(disco, dict):
        base = base if isinstance(base, dict) else {}
        res = dict(disco)
        res.update({k: v for k, v in local.items() if base.get(k) != v})
        for k in base:
            if k not in local:
                res.pop(k, None)
        return res
    if isinstance(local, list) and isinstance(disco, list):
        base = base if isinstance(base, list) else []
        quitados = [x for x in base if x not in local]
        nuevos = [x for x in local if x not in base and x not in disco]
        return [x for x in disco if x not in quitados] + nuevos
    return local        # valor suelto cambiado en los dos sitios: gana este PC


def _mtime_config():
    try:
        return os.path.getmtime(CONFIG_PATH)
    except OSError:
        return None


def recordar_disco(cfg=None):
    """Anota el estado de los avisos que acaba de quedar en disco."""
    cfg = CFG if cfg is None else cfg
    _SYNC["base"] = {k: copy.deepcopy(cfg[k]) for k in CLAVES_AVISOS if k in cfg}
    _SYNC["mtime"] = _mtime_config()


def sincronizar_avisos():
    """Mezcla en CFG los avisos que otro ordenador dejo en config.json."""
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            disco = json.load(f)
    except Exception:
        return False                  # a medias o sin leer: se reintenta luego
    _SYNC["mtime"] = _mtime_config()
    cambio = False
    for k in CLAVES_AVISOS:
        if k not in disco:
            continue
        nuevo = _fusionar(_SYNC["base"].get(k), CFG.get(k), disco[k])
        cambio = cambio or nuevo != CFG.get(k)
        CFG[k] = nuevo
        _SYNC["base"][k] = copy.deepcopy(disco[k])
    return cambio


def config_cambio_fuera():
    """True si config.json ya no es el que escribimos o leimos nosotros."""
    return _mtime_config() != _SYNC["mtime"]


def guardar_config():
    """Guarda config.json sin riesgo de dejarlo a medias.

    Se serializa ANTES de tocar el disco: si CFG tuviera algo que json no
    sabe escribir, falla aqui y el fichero bueno sigue intacto.
    """
    sincronizar_avisos()        # que no pise lo que apuntaron en otro PC
    datos = b""
    try:
        with open(CONFIG_PATH, "rb") as viejo:
            datos = viejo.read()
    except OSError:
        pass
    try:                        # lo de este ordenador va a equipo.json; en config.json se deja lo que ya hubiera (un PC con la
        en_disco = json.loads(datos) if datos.strip() else {}     # version anterior lo sigue leyendo) para no reescribirlo al mover la ventana
        compartido = {k: (en_disco[k] if k in CLAVES_EQUIPO and k in en_disco else v) for k, v in CFG.items()}
        texto = json.dumps(compartido, indent=2, ensure_ascii=False)
        equipo = json.dumps({k: CFG[k] for k in CLAVES_EQUIPO if k in CFG}, indent=2, ensure_ascii=False)
    except Exception as exc:
        log(f"NO SE PUDO SERIALIZAR la configuración ({exc}); "
            "no se toca config.json")
        return False
    try:
        escribir_atomico(EQUIPO_PATH, equipo)
        if datos.replace(b"\r\n", b"\n") != texto.encode("utf-8"):   # si no cambia nada compartido no se toca (OneDrive no sincroniza en balde)
            if datos.strip():                 # copia del anterior, por si acaso
                try:
                    escribir_atomico(CONFIG_PATH + ".bak", datos)
                except Exception as exc:
                    log(f"No se pudo copiar config.json.bak ({exc})")
            escribir_atomico(CONFIG_PATH, texto)
        recordar_disco()
        return True
    except Exception as exc:
        log(f"NO SE PUDO GUARDAR config.json ({exc}); "
            "¿está la carpeta en un sitio sin permisos de escritura?")
        return False


recordar_disco()


# Guardado agrupado: marcar seis casillas de habitos seguidas no debe ser seis
# ciclos completos de leer + .bak + escribir, con OneDrive sincronizando cada
# vez. Se apunta que hay cambios y se vuelca una sola vez poco despues.
_PENDIENTE = {"hay": False, "raiz": None, "id": None}


def guardar_config_pronto(raiz=None, espera=1500):
    """Pide guardar dentro de un momento, agrupando los cambios seguidos."""
    raiz = raiz or _PENDIENTE["raiz"]
    if raiz is None:
        return guardar_config()       # sin ventana todavia: al disco y ya
    _PENDIENTE["raiz"] = raiz
    _PENDIENTE["hay"] = True
    if _PENDIENTE["id"] is not None:
        try:
            raiz.after_cancel(_PENDIENTE["id"])
        except Exception:
            pass
    _PENDIENTE["id"] = raiz.after(espera, volcar_config)
    return True


def volcar_config():
    """Escribe ya lo que estuviera pendiente (al cerrar, o al cumplirse el plazo)."""
    _PENDIENTE["id"] = None
    if not _PENDIENTE["hay"]:
        return False
    _PENDIENTE["hay"] = False
    return guardar_config()


def copia_fechada():
    """Guarda una copia con la fecha en el nombre. Devuelve la ruta."""
    destino = os.path.join(
        DATOS, f"config-{dt.datetime.now():%Y%m%d-%H%M}.json")
    with open(destino, "w", encoding="utf-8") as f:
        json.dump(CFG, f, indent=2, ensure_ascii=False)
    log(f"Copia de la configuración en {os.path.basename(destino)}")
    return destino


# --------------------------------------------------------------------------
# Datos
# --------------------------------------------------------------------------

class Evento:
    __slots__ = ("inicio", "fin", "titulo", "todo_dia", "uid",
                 "col", "ncols", "parte", "viene", "sigue",
                 "categoria", "color", "lugar", "notas", "calendario",
                 "estado", "id_outlook", "almacen_outlook")

    def __init__(self, inicio, fin, titulo, todo_dia, uid,
                 parte=0, viene=False, sigue=False,
                 categoria=None, color=None, lugar="", notas="",
                 calendario="", estado="ocupado", id_outlook="",
                 almacen_outlook=""):
        self.inicio = inicio
        self.fin = fin
        self.titulo = titulo
        self.todo_dia = todo_dia
        self.uid = uid
        self.col = 0
        self.ncols = 1
        self.parte = parte      # 0 = trozo original; >0 = continuacion en otro dia
        self.viene = viene      # el evento venia del dia anterior
        self.sigue = sigue      # el evento continua al dia siguiente
        self.categoria = categoria   # etiqueta de Outlook o tipo de festivo
        self.color = color           # None = color normal de evento
        self.lugar = lugar
        self.notas = notas
        self.calendario = calendario
        self.estado = estado         # ocupado / provisional / libre / fuera
        self.id_outlook = id_outlook          # EntryID, para poder abrirla
        self.almacen_outlook = almacen_outlook


def limites_ventana(ref=None, desplazamiento=0):
    """Rango de fechas que se muestra.

    'desplazamiento' son dias sumados al inicio: permite moverse a semanas
    anteriores o posteriores sin volver a descargar nada.
    """
    ahora = ref or dt.datetime.now(TZ)
    ini = ahora.replace(hour=0, minute=0, second=0, microsecond=0)
    if CFG.get("inicio_ventana", "hoy") == "lunes":
        ini -= dt.timedelta(days=ahora.weekday())
    ini += dt.timedelta(days=desplazamiento)
    return ini, ini + dt.timedelta(days=CFG["dias_visibles"])


# Hasta donde ha viajado el usuario mas alla del cache normal (dias desde hoy).
ALCANCE = {"atras": 0, "adelante": 0}


def rango_descarga(ref=None):
    """Ventana amplia que se pide al servidor, para poder navegar sin red."""
    ini, _ = limites_ventana(ref)
    adelante = max(CFG["dias_cache_adelante"], CFG["dias_visibles"] + 14,
                   ALCANCE["adelante"])
    atras = max(CFG["dias_cache_atras"], ALCANCE["atras"])
    return (ini - dt.timedelta(days=atras),
            ini + dt.timedelta(days=adelante))


def _medianoche_siguiente(momento):
    """Primer instante del dia siguiente a 'momento' (a prueba de cambios de hora)."""
    dia_sig = momento.date() + dt.timedelta(days=1)
    return dt.datetime.combine(dia_sig, dt.time(0), tzinfo=TZ)


def partir_por_dias(ev, maximo=32):
    """Divide un evento que cruza la medianoche en un trozo por dia."""
    if ev.todo_dia or ev.fin <= _medianoche_siguiente(ev.inicio):
        return [ev]

    piezas = []
    ini = ev.inicio
    n = 0
    while n < maximo:
        corte = _medianoche_siguiente(ini)
        ultimo = ev.fin <= corte
        piezas.append(Evento(
            ini, ev.fin if ultimo else corte, ev.titulo, False, ev.uid,
            parte=n, viene=n > 0, sigue=not ultimo,
            categoria=ev.categoria, color=ev.color, lugar=ev.lugar,
            notas=ev.notas, calendario=ev.calendario, estado=ev.estado,
            id_outlook=ev.id_outlook, almacen_outlook=ev.almacen_outlook,
        ))
        if ultimo:
            break
        ini = corte
        n += 1
    return piezas


PROP_CUERPO = "http://schemas.microsoft.com/mapi/proptag/0x1000001F"  # PR_BODY_W


PROP_EMAIL_REMITENTE = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"


def _email_remitente(correo):
    """Dirección del remitente por la vía que NO despierta al guardián.

    Leer 'SenderEmailAddress' provoca el aviso «un programa intenta obtener
    acceso a direcciones de correo» cada vez que se descarga el correo (al
    iniciar sesión, por ejemplo). El PropertyAccessor da lo mismo sin aviso;
    si falla, se deja vacío antes que arriesgarse a que vuelva a saltar.
    """
    try:
        return str(correo.PropertyAccessor.GetProperty(PROP_EMAIL_REMITENTE) or "")
    except Exception:
        return ""


def notas_de_la_cita(cita):
    """Notas de la cita sin despertar al guardián de Outlook.

    Leer 'cita.Body' hace que Outlook saque el aviso de 'un programa intenta
    obtener acceso a direcciones de correo'. Por eso, de serie no se leen las
    notas; y si se piden, primero se intenta por la vía que no avisa.
    """
    if not CFG.get("leer_notas", False):
        return ""
    try:
        return str(cita.PropertyAccessor.GetProperty(PROP_CUERPO) or "")
    except Exception:
        try:
            return str(cita.Body or "")
        except Exception as exc:
            log(f"No se pudieron leer las notas: {exc}")
            return ""


def _limpiar(texto, tope=280):
    """Quita etiquetas HTML y espacios sobrantes de las notas del evento."""
    texto = re.sub(r"<[^>]+>", " ", texto or "")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto[:tope] + ("…" if len(texto) > tope else "")


def _descargar_uno(conf, ini, fin):
    """Eventos de un calendario concreto, ya con su color y su nombre."""
    resp = requests.get(conf["url"], timeout=25)
    resp.raise_for_status()
    cal = icalendar.Calendar.from_ical(resp.content)
    nombre = conf.get("nombre", "Calendario")
    log(f"'{nombre}': {len(resp.content)} bytes descargados")

    eventos, etiquetas = [], set()
    # recurring_ical_events expande RRULE y aplica EXDATE y las excepciones
    # (los VEVENT con RECURRENCE-ID que sustituyen a una ocurrencia concreta)
    for e in recurring_ical_events.of(cal).between(ini, fin):
        inicio = e["DTSTART"].dt
        fin_ev = e["DTEND"].dt if "DTEND" in e else None
        titulo = str(e.get("SUMMARY", "(sin titulo)"))
        uid = str(e.get("UID", titulo))
        cats = categorias_de(e)
        etiquetas.update(cats)
        categoria = cats[0] if cats else None
        # la etiqueta manda sobre el color del calendario
        color = color_de_categoria(categoria) or conf.get("color")
        estado = ESTADOS_ICS.get(
            str(e.get("X-MICROSOFT-CDO-BUSYSTATUS", "")).upper(),
            "libre" if str(e.get("TRANSP", "")).upper() == "TRANSPARENT"
            else "ocupado")
        extra = dict(categoria=categoria, color=color, calendario=nombre,
                     estado=estado,
                     lugar=_limpiar(str(e.get("LOCATION", "")), 90),
                     notas=_limpiar(str(e.get("DESCRIPTION", ""))))

        if not isinstance(inicio, dt.datetime):  # evento de todo el dia
            fecha_ini = inicio
            # DTEND en eventos de dia completo es exclusivo
            fecha_fin = fin_ev if isinstance(fin_ev, dt.date) and not isinstance(
                fin_ev, dt.datetime) else None
            if fecha_fin is None:
                fecha_fin = fecha_ini + dt.timedelta(days=1)
            eventos.append(Evento(
                dt.datetime.combine(fecha_ini, dt.time(0), tzinfo=TZ),
                dt.datetime.combine(fecha_fin, dt.time(0), tzinfo=TZ),
                titulo, True, uid, **extra))
            continue

        inicio = inicio.astimezone(TZ)
        fin_ev = fin_ev.astimezone(TZ) if isinstance(fin_ev, dt.datetime) \
            else inicio + dt.timedelta(hours=1)
        if fin_ev <= inicio:
            fin_ev = inicio + dt.timedelta(minutes=30)
        eventos.extend(partir_por_dias(
            Evento(inicio, fin_ev, titulo, False, uid, **extra)))

    log(f"'{nombre}': {len(eventos)} eventos. Etiquetas: "
        + (", ".join(sorted(etiquetas)) if etiquetas else "ninguna"))
    return eventos


# Colores de las categorias de Outlook (enum OlCategoryColor, 1..25).
COLORES_CAT_OUTLOOK = {
    1: "#e5533d", 2: "#f08c2e", 3: "#f4b183", 4: "#e0c341", 5: "#3fa34d",
    6: "#2ec4a6", 7: "#8a9a3b", 8: "#3a86ff", 9: "#a06bff", 10: "#a04a5e",
    11: "#8a8a90", 12: "#5a5a62", 13: "#9a9aa0", 14: "#6a6a70", 15: "#4a4a4e",
    16: "#a83228", 17: "#c06a1e", 18: "#d09a6a", 19: "#b09a20", 20: "#2a7a34",
    21: "#1e8a76", 22: "#6a7a2b", 23: "#2a5ecc", 24: "#7a4ac0", 25: "#7a3444",
}


def _colores_de_categorias(fuente, aviso=""):
    """Categorias -> color. 'fuente' es el Namespace o el Store de la cuenta.

    Ojo: en Outlook las categorias son POR CUENTA, asi que hay que pedirselas
    al almacen de la carpeta que se esta leyendo, no al de por defecto.
    """
    colores = {}
    try:
        for cat in fuente.Categories:
            color = COLORES_CAT_OUTLOOK.get(int(cat.Color))
            if color:
                colores[str(cat.Name)] = color
    except Exception as exc:
        log(f"Sin colores de categorías{aviso}: {exc}")
    return colores


def carpetas_de_calendario(ns, carpeta=None, ruta="", salida=None, prof=0):
    """Todas las carpetas de calendario de todas las cuentas de Outlook."""
    salida = [] if salida is None else salida
    if prof > 5:            # algunos calendarios cuelgan bastante hondo
        return salida
    try:
        coleccion = ns.Folders if carpeta is None else carpeta.Folders
        total = coleccion.Count
    except Exception:
        return salida
    for i in range(1, total + 1):
        try:
            f = coleccion.Item(i)
            nombre = f"{ruta}/{f.Name}" if ruta else str(f.Name)
            try:
                if int(f.DefaultItemType) == 1:      # olAppointmentItem
                    salida.append((nombre, f))
            except Exception:
                pass
            carpetas_de_calendario(ns, f, nombre, salida, prof + 1)
        except Exception:
            continue
    return salida


def _seleccion_por_fechas(citas, ini, fin):
    """Filtra por rango probando los formatos de fecha que acepta Outlook."""
    for formato in ("%d/%m/%Y %H:%M", "%m/%d/%Y %H:%M", "%m/%d/%Y %I:%M %p"):
        try:
            seleccion = citas.Restrict(
                f"[Start] < '{fin.strftime(formato)}' AND "
                f"[End] > '{ini.strftime(formato)}'")
            _ = seleccion.Count
            return seleccion
        except Exception:
            continue
    return None


def resumen_de_carpeta(carpeta, dias=30):
    """(cuenta, elementos guardados, citas en los próximos días).

    'Elementos' puede ser 0 aunque el calendario tenga cosas: las series se
    guardan como un solo elemento y las carpetas recién sincronizadas tardan.
    Por eso se mira también cuántas citas caen en los próximos días, que es
    justo lo que luego enseña el widget.
    """
    try:
        cuenta = str(getattr(getattr(carpeta, "Store", None), "DisplayName", "")
                     or "")
    except Exception:
        cuenta = ""
    try:
        total = int(carpeta.Items.Count)
    except Exception:
        total = -1
    proximas = -1
    try:
        ahora = dt.datetime.now()
        citas = carpeta.Items
        citas.IncludeRecurrences = True
        citas.Sort("[Start]")
        seleccion = _seleccion_por_fechas(citas, ahora - dt.timedelta(days=1),
                                          ahora + dt.timedelta(days=dias))
        if seleccion is not None:
            proximas = 0
            for _ in seleccion:            # contar a mano: Count miente con
                proximas += 1              # recurrencias incluidas
                if proximas > 500:
                    break
    except Exception as exc:
        log(f"No se pudo resumir la carpeta: {exc}")
    return cuenta, total, proximas


def _localizar(todas, patron):
    """Carpeta cuyo nombre coincide (exacto primero, si no por trozo)."""
    patron = str(patron or "").strip().lower()
    if not patron:
        return None
    for nombre, carpeta in todas:
        if nombre.lower() == patron:
            return nombre, carpeta
    for nombre, carpeta in todas:
        if patron in nombre.lower():
            return nombre, carpeta
    return None


def _carpetas_elegidas(ns):
    """[(nombre a mostrar, color, carpeta)] segun el config."""
    conf = list(CFG.get("carpetas_outlook") or [])
    suelta = str(CFG.get("carpeta_outlook") or "").strip()
    if not conf and suelta:                      # formato antiguo, una sola
        conf = [{"nombre": "Outlook", "buscar": suelta}]
    if not conf:
        return [("Outlook", None, ns.GetDefaultFolder(9))]   # 9 = olFolderCalendar

    todas = carpetas_de_calendario(ns)
    salida = []
    for c in conf:
        if not c.get("activo", True):
            continue
        hallada = _localizar(todas, c.get("buscar"))
        if not hallada:
            log(f"Carpeta de Outlook no encontrada: '{c.get('buscar')}'")
            continue
        nombre_real, carpeta = hallada
        salida.append((c.get("nombre") or nombre_real, c.get("color"), carpeta))
        log(f"Carpeta de Outlook: {nombre_real}")
    return salida or [("Outlook", None, ns.GetDefaultFolder(9))]


class OutlookCerrado(Exception):
    """Outlook no esta abierto: no se lanza solo, se espera a que lo abras y se reintenta."""


def outlook_abierto():
    """El Outlook que ya esta en marcha. Dispatch lo lanzaria el solo, antes de que OneDrive y tu perfil esten listos."""
    import pythoncom
    import win32com.client
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except pythoncom.com_error:
        raise OutlookCerrado("Outlook no está abierto: se espera a que lo abras") from None


def leer_de_outlook(ini, fin):
    """Lee el calendario del Outlook clasico instalado (con sus categorias)."""
    import pythoncom
    pythoncom.CoInitialize()
    app = outlook_abierto()
    ns = app.GetNamespace("MAPI")
    colores_generales = _colores_de_categorias(ns, " de la cuenta por defecto")

    eventos, etiquetas = [], set()
    for nombre_cal, color_cal, carpeta in _carpetas_elegidas(ns):
        colores_cat = dict(colores_generales)
        try:                       # las de la cuenta de esa carpeta mandan
            colores_cat.update(_colores_de_categorias(
                carpeta.Store, f" de '{nombre_cal}'"))
        except Exception as exc:
            log(f"No se pudo acceder a las categorías de '{nombre_cal}': {exc}")
        if colores_cat:
            log(f"Categorías con color en '{nombre_cal}': "
                + ", ".join(f"{k} {v}" for k, v in colores_cat.items()))
        eventos.extend(_leer_carpeta(carpeta, nombre_cal, color_cal, ini, fin,
                                     colores_cat, etiquetas))
    log("Etiquetas en Outlook: "
        + (", ".join(sorted(etiquetas)) if etiquetas else "ninguna"))
    return eventos


def reloj_de(valor):
    """Los numeros de una fecha de Outlook, sin zona horaria."""
    return dt.datetime(int(valor.year), int(valor.month), int(valor.day),
                       int(valor.hour), int(valor.minute))


def hora_de_cita(cita, campo="Start"):
    """Hora de una cita de Outlook, en nuestra zona.

    Start y End ya vienen en hora local (las 18:00 de Outlook son las 18:00),
    aunque pywin32 les cuelgue una zona que no es la suya. Por eso se coge el
    reloj tal cual: convertir por el instante las adelantaba una o dos horas.
    """
    return reloj_de(getattr(cita, campo)).replace(tzinfo=TZ)


def _leer_carpeta(carpeta, nombre_cal, color_cal, ini, fin, colores_cat, etiquetas):
    citas = carpeta.Items
    citas.IncludeRecurrences = True
    citas.Sort("[Start]")

    # el formato de fecha del filtro depende del idioma de Outlook
    seleccion = None
    for formato in ("%d/%m/%Y %H:%M", "%m/%d/%Y %H:%M", "%m/%d/%Y %I:%M %p"):
        try:
            seleccion = citas.Restrict(
                f"[Start] < '{fin.strftime(formato)}' AND "
                f"[End] > '{ini.strftime(formato)}'")
            _ = seleccion.Count
            break
        except Exception:
            seleccion = None
    if seleccion is None:
        raise RuntimeError("Outlook no aceptó ningún formato de fecha")

    eventos = []
    for cita in seleccion:
        try:
            inicio = hora_de_cita(cita, "Start")
            final = hora_de_cita(cita, "End")
            titulo = str(cita.Subject or "(sin titulo)")
            cats = [c.strip() for c in str(cita.Categories or "").split(";")
                    if c.strip()]
            etiquetas.update(cats)
            categoria = cats[0] if cats else None
            color = (CFG.get("colores_categoria", {}).get(categoria)
                     or colores_cat.get(categoria)
                     or color_de_categoria(categoria)
                     or color_cal)
            extra = dict(categoria=categoria, color=color,
                         calendario=nombre_cal,
                         estado=ESTADOS_COM.get(int(getattr(cita, "BusyStatus", 2)),
                                                "ocupado"),
                         id_outlook=str(getattr(cita, "EntryID", "")),
                         almacen_outlook=str(getattr(carpeta, "StoreID", "")),
                         lugar=_limpiar(str(cita.Location or ""), 90),
                         notas=_limpiar(notas_de_la_cita(cita)))
            uid = str(getattr(cita, "GlobalAppointmentID", titulo))
            if cita.AllDayEvent:
                eventos.append(Evento(inicio, final, titulo, True, uid, **extra))
            else:
                eventos.extend(partir_por_dias(
                    Evento(inicio, final, titulo, False, uid, **extra)))
        except Exception as exc:
            log(f"Cita ilegible en Outlook: {exc}")
    log(f"'{nombre_cal}': {len(eventos)} eventos leídos de Outlook")
    return eventos


ESTADO = {"origen": None}      # de dónde salieron los últimos eventos: "outlook", "ics" o "cerrado" (Outlook sin abrir)


def descargar_eventos():
    """Junta los eventos de todos los calendarios activos."""
    ini, fin = rango_descarga()
    origen = CFG.get("origen", "ics")
    if origen in ("auto", "outlook"):
        try:
            eventos = leer_de_outlook(ini, fin)
            if eventos or origen == "outlook":
                eventos.sort(key=lambda x: x.inicio)
                log(f"Origen: Outlook instalado, {len(eventos)} eventos")
                ESTADO["origen"] = "outlook"
                return eventos
            log("Outlook no devolvió nada; se usan los enlaces ICS")
        except OutlookCerrado as exc:
            ESTADO["origen"] = "cerrado"
            if origen == "outlook":
                raise                      # solo Outlook: sin el no hay datos; se reintenta en reintento_minutos
            log(f"{exc}; se usan los ICS")
        except Exception as exc:
            log(f"No se pudo leer del Outlook instalado ({exc}); se usan los ICS")
    eventos = []
    for conf in CFG["calendarios"]:
        if not conf.get("activo", True) or not conf.get("url"):
            continue
        try:
            eventos.extend(_descargar_uno(conf, ini, fin))
        except Exception as exc:
            # que falle uno no debe dejar sin datos a los demas
            log(f"ERROR en '{conf.get('nombre', '?')}': {exc}")
    eventos.sort(key=lambda x: x.inicio)
    if ESTADO["origen"] != "cerrado":
        ESTADO["origen"] = "ics"
    log(f"Total entre {ini:%d/%m} y {fin:%d/%m}: {len(eventos)} eventos")
    return eventos


def repartir_columnas(eventos):
    """Agrupa eventos que se solapan y les asigna columna para que no se pisen."""
    grupo, fin_grupo = [], None
    for ev in eventos:
        if fin_grupo is not None and ev.inicio >= fin_grupo:
            _asignar(grupo)
            grupo, fin_grupo = [], None
        grupo.append(ev)
        fin_grupo = ev.fin if fin_grupo is None else max(fin_grupo, ev.fin)
    _asignar(grupo)


def _asignar(grupo):
    if not grupo:
        return
    columnas = []  # hora de fin ocupada por cada columna
    for ev in grupo:
        colocado = False
        for i, libre_desde in enumerate(columnas):
            if ev.inicio >= libre_desde:
                columnas[i] = ev.fin
                ev.col = i
                colocado = True
                break
        if not colocado:
            columnas.append(ev.fin)
            ev.col = len(columnas) - 1
    for ev in grupo:
        ev.ncols = len(columnas)


def guardar_cache(eventos):
    eventos = [e for e in eventos if e.categoria != "postit"]
    try:
        # atomico: el hilo de descarga y el principal pueden coincidir, y una
        # cache pickle a medias hace que al arrancar no se vea nada
        escribir_atomico(CACHE_PATH,
                         pickle.dumps({"v": CACHE_VERSION, "eventos": eventos}))
    except Exception as exc:
        print("No se pudo guardar la cache:", exc)
        log(f"No se pudo guardar la caché: {exc}")


def leer_cache():
    try:
        with open(CACHE_PATH, "rb") as f:
            datos = pickle.load(f)
        if isinstance(datos, dict) and datos.get("v") == CACHE_VERSION:
            return datos["eventos"]
        return []  # cache de una version anterior: se ignora
    except Exception:
        return []


# --------------------------------------------------------------------------
# Festivos (nacionales, Comunidad de Madrid y Alcala de Henares)
# --------------------------------------------------------------------------

# Tabla de respaldo por si no esta instalada la libreria "holidays".
RESPALDO_NACIONAL = {
    "2026-01-01": "Año Nuevo", "2026-01-06": "Epifanía del Señor",
    "2026-04-03": "Viernes Santo", "2026-05-01": "Fiesta del Trabajo",
    "2026-08-15": "Asunción de la Virgen",
    "2026-10-12": "Fiesta Nacional de España",
    "2026-12-08": "Inmaculada Concepción", "2026-12-25": "Navidad",
    "2027-01-01": "Año Nuevo", "2027-01-06": "Epifanía del Señor",
    "2027-03-26": "Viernes Santo", "2027-05-01": "Fiesta del Trabajo",
    "2027-10-12": "Fiesta Nacional de España",
    "2027-11-01": "Todos los Santos",
    "2027-12-06": "Día de la Constitución",
    "2027-12-08": "Inmaculada Concepción", "2027-12-25": "Navidad",
}
RESPALDO_MADRID = {
    "2026-04-02": "Jueves Santo",
    "2026-05-02": "Fiesta de la Comunidad de Madrid",
    "2026-11-02": "Lunes siguiente a Todos los Santos",
    "2026-12-07": "Lunes siguiente al Día de la Constitución",
    "2027-03-19": "San José", "2027-03-25": "Jueves Santo",
    "2027-05-03": "Lunes siguiente a la Fiesta de la Comunidad de Madrid",
}


def _de_tabla(tabla, anios):
    salida = {}
    for iso, nombre in tabla.items():
        fecha = dt.date.fromisoformat(iso)
        if fecha.year in anios:
            salida[fecha] = nombre
    return salida


@lru_cache(maxsize=16)
def _oficiales_cacheados(comunidad, anios):
    """Lo caro de verdad: construir las tablas de la libreria holidays.

    Los diccionarios que devuelve son compartidos entre llamadas: quien los
    reciba puede recorrerlos, pero no debe modificarlos.
    """
    try:
        import holidays
        nacional, regional = {}, {}
        for anio in anios:
            nac = holidays.Spain(years=anio)
            reg = holidays.Spain(subdiv=comunidad, years=anio)
            nacional.update(nac)
            regional.update({d: n for d, n in reg.items() if d not in nac})
        return nacional, regional
    except Exception as exc:
        log(f"Sin librería holidays, se usa la tabla interna: {exc}")
        if comunidad != "MD":
            # la tabla de respaldo solo tiene los de Madrid: mejor ninguno
            # que enseñar los de otra comunidad como si fueran los tuyos
            log(f"Sin festivos de '{comunidad}': instala la librería holidays")
            return _de_tabla(RESPALDO_NACIONAL, anios), {}
        return (_de_tabla(RESPALDO_NACIONAL, anios),
                _de_tabla(RESPALDO_MADRID, anios))


def _festivos_oficiales(anios):
    """Nacionales y de tu comunidad autonoma, por separado."""
    comunidad = CFG["festivos"].get("regional", {}).get("comunidad", "MD")
    return _oficiales_cacheados(comunidad, tuple(sorted(anios)))


def _festivos_de(conf, anios):
    """Fechas fijas cada anio (mes-dia) mas las excepciones (anio-mes-dia)."""
    salida = {}
    for anio in anios:
        for mes_dia, nombre in (conf.get("fijos") or {}).items():
            try:
                mes, dia = (int(x) for x in str(mes_dia).split("-"))
                salida[dt.date(anio, mes, dia)] = nombre
            except (ValueError, TypeError):
                continue
    for iso, nombre in (conf.get("extra") or {}).items():
        try:
            salida[dt.date.fromisoformat(iso)] = nombre
        except (ValueError, TypeError):
            continue
    return salida


def _festivos_locales(anios):
    """Los del municipio 'local' de siempre (se mantiene por compatibilidad)."""
    return _festivos_de(CFG["festivos"].get("local", {}), anios)


def grupos_de_festivos():
    """Todos los grupos que hay que pintar: los tres de siempre y, detrás,
    un grupo por cada municipio de la lista 'municipios'.

    Devuelve pares (clave, conf) donde la clave identifica al grupo en el
    menú y en la categoría del evento.
    """
    fest = CFG.get("festivos") or {}
    grupos = [(c, fest.get(c, {})) for c in ("nacional", "regional", "local")]
    for m in (fest.get("municipios") or []):
        if not isinstance(m, dict):
            continue
        clave = str(m.get("clave") or m.get("nombre") or "").strip()
        if clave:
            grupos.append((f"municipio:{clave}", m))
    return grupos


CATEGORIAS_FESTIVO = ("nacional", "regional", "local")


def es_festivo(categoria):
    """True si esa categoria es de un festivo (de cualquier ambito).

    Los municipios usan claves del tipo "municipio:alcala", asi que no vale
    con comparar contra la lista de siempre.
    """
    cat = str(categoria or "")
    return cat in CATEGORIAS_FESTIVO or cat.startswith("municipio:")


def generar_festivos(ini, fin):
    """Festivos activos del rango, como eventos de todo el dia."""
    anios = set(range(ini.year, fin.year + 1))
    nacional, regional = _festivos_oficiales(anios)

    eventos = []
    for clave, conf in grupos_de_festivos():
        if not conf.get("activo", True):
            continue
        if clave == "nacional":
            dias = nacional
        elif clave == "regional":
            dias = regional
        else:
            dias = _festivos_de(conf, anios)
        color = conf.get("color", C["evento"])
        for fecha, nombre in dias.items():
            if not ini.date() <= fecha < fin.date():
                continue
            eventos.append(Evento(
                dt.datetime.combine(fecha, dt.time(0), tzinfo=TZ),
                dt.datetime.combine(fecha + dt.timedelta(days=1), dt.time(0),
                                    tzinfo=TZ),
                str(nombre), True, f"festivo-{clave}-{fecha}",
                categoria=clave, color=color))
    return eventos


def generar_cumpleanos(ini, fin):
    """Cumpleaños guardados, como eventos de todo el dia, cada año."""
    if not CFG.get("cumpleanos"):
        return []
    color = CFG.get("color_cumpleanos", "#a855f7")
    eventos = []
    for anio in range(ini.year - 1, fin.year + 2):
        for c in CFG["cumpleanos"]:
            try:
                fecha = dt.date(anio, int(c["mes"]), int(c["dia"]))
            except (ValueError, KeyError):
                continue  # ej. 29 de febrero en año no bisiesto
            if not ini.date() <= fecha < fin.date():
                continue
            nombre = c.get("nombre", "").strip() or "Cumpleaños"
            eventos.append(Evento(
                dt.datetime.combine(fecha, dt.time(0), tzinfo=TZ),
                dt.datetime.combine(fecha + dt.timedelta(days=1), dt.time(0),
                                    tzinfo=TZ),
                f"🎂 {nombre}", True, f"cumple-{nombre}-{anio}",
                categoria="cumpleanos", color=color))
    return eventos


# --------------------------------------------------------------------------
# Etiquetas (categorias) de Outlook
# --------------------------------------------------------------------------

# Nombres estandar de las categorias de color de Outlook, en es y en en.
COLORES_OUTLOOK = {
    "rojo": "#e5533d", "roja": "#e5533d", "red": "#e5533d",
    "naranja": "#f08c2e", "orange": "#f08c2e",
    "amarillo": "#e0c341", "amarilla": "#e0c341", "yellow": "#e0c341",
    "verde": "#3fa34d", "green": "#3fa34d",
    "azul": "#3a86ff", "blue": "#3a86ff",
    "morado": "#a06bff", "morada": "#a06bff",
    "purpura": "#a06bff", "púrpura": "#a06bff", "purple": "#a06bff",
    "gris": "#8a8a90", "grey": "#8a8a90", "gray": "#8a8a90",
}


ESTADOS_ICS = {"FREE": "libre", "TENTATIVE": "provisional", "BUSY": "ocupado",
               "OOF": "fuera", "WORKINGELSEWHERE": "fuera"}
ESTADOS_COM = {0: "libre", 1: "provisional", 2: "ocupado", 3: "fuera",
               4: "fuera"}


def categorias_de(evento_ical):
    """Lista de etiquetas del VEVENT (vacia si el ICS no las trae)."""
    valor = evento_ical.get("CATEGORIES")
    if valor is None:
        return []
    try:
        return [str(v).strip() for v in valor.cats if str(v).strip()]
    except AttributeError:
        return [t.strip() for t in str(valor).split(",") if t.strip()]


def color_de_categoria(nombre):
    """Color configurado para esa etiqueta, o el de la categoria de Outlook."""
    if not nombre:
        return None
    conf = CFG.get("colores_categoria", {})
    if nombre in conf:
        return conf[nombre]
    minus = nombre.lower()
    for palabra, color in COLORES_OUTLOOK.items():
        if palabra in minus:
            return color
    return None


# --------------------------------------------------------------------------
# Notificaciones
# --------------------------------------------------------------------------

def notificar(titulo, cuerpo, mudo=False):
    """Aviso de Windows. Respeta la campana: si está silenciada, no sale
    ninguno (antes solo lo comprobaba el aviso de eventos en tic(), y se
    colaban el resumen del día, "Configuración guardada", "Copiado"...)."""
    if not CFG.get("notificaciones_activas", True):
        return
    try:
        from winotify import Notification, audio
        aviso = Notification(app_id="Calendario", title=titulo, msg=cuerpo)
        if mudo:        # suena el sonido elegido, no el de Windows
            aviso.set_audio(audio.Silent, loop=False)
        aviso.show()
    except Exception:
        print(f"[AVISO] {titulo} - {cuerpo}")


CARPETA_SONIDOS = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "Media")


def reproducir_sonido(ruta, veces=1):
    """Suena un .wav en segundo plano."""
    def sonar():
        try:
            import winsound
            for _ in range(veces):
                winsound.PlaySound(ruta, winsound.SND_FILENAME)
        except Exception:
            pass
    threading.Thread(target=sonar, daemon=True).start()


def sonido_de_aviso(veces=1):
    """Suena el sonido elegido en el menú de avisos.

    True si ya se ha ocupado (sonido propio o «ninguno»); False si toca el
    sonido de siempre."""
    ruta = CFG.get("sonido_aviso", "")
    if ruta == "ninguno":
        return True
    if ruta and os.path.isfile(ruta):
        reproducir_sonido(ruta, veces)
        return True
    return False


def sonido_alarma():
    """Pitidos de alarma en segundo plano (Beep bloquea, por eso el hilo)."""
    if sonido_de_aviso(3):
        return
    def sonar():
        try:
            import winsound
            for _ in range(4):
                for frecuencia in (880, 1175):
                    winsound.Beep(frecuencia, 220)
                time.sleep(0.4)
        except Exception:
            pass
    threading.Thread(target=sonar, daemon=True).start()


# --------------------------------------------------------------------------
# Habitos: casillas por dia
# --------------------------------------------------------------------------
#
# Las casillas se marcan a mano y se guardan en config.json. Se puede marcar
# y desmarcar cualquier dia pasado, que es justo lo que hace falta para
# llevar la cuenta de si hiciste esto o aquello.

COLORES_HABITO = ["#2ec4a6", "#3a86ff", "#f08c2e", "#a855f7", "#e5533d",
                  "#e0c341", "#3fa34d", "#ff2d55"]

# Los tres estados de una casilla. El ciclo al hacer clic es este mismo orden,
# porque lo normal es marcar «cumplido»: un solo clic para el caso frecuente.
ESTADOS_HABITO = ("", "ok", "no")
MARCA_HABITO = {"": "", "ok": "\u2713", "no": "\u2715"}
# «pendiente» no se guarda: se deduce de que ese dia haya un evento que hable
# del habito y todavia no lo hayas marcado. Se pinta como un circulo hueco.
MARCA_PENDIENTE = "\u25cb"


def habito_palabras(h):
    """Palabras que delatan al habito en el titulo de un evento.

    Por defecto, el propio nombre. Con el campo "buscar" se pueden poner
    varias separadas por comas: «gym, gimnasio, pesas».
    """
    texto = str(h.get("buscar") or h.get("nombre") or "")
    return [sin_acentos(p.strip().lower()) for p in texto.split(",")
            if p.strip()]


def eventos_de_habito(eventos, h, desde, hasta):
    """Eventos del rango cuyo titulo habla de ese habito."""
    palabras = habito_palabras(h)
    if not palabras or not eventos:
        return []
    encontrados = []
    for ev in eventos:
        try:
            fecha = ev.inicio.date()
        except Exception:
            continue
        if not desde <= fecha <= hasta:
            continue
        titulo = sin_acentos(str(ev.titulo or "").lower())
        if any(p in titulo for p in palabras):
            encontrados.append(ev)
    return encontrados


def dias_planificados(eventos, h, desde, hasta):
    """Dias del rango en los que el calendario habla de ese habito."""
    return {ev.inicio.date() for ev in eventos_de_habito(eventos, h, desde, hasta)}


def habito_vencido(eventos, h, ahora):
    """True si hoy lo tenias planificado a una hora que ya paso.

    Es lo que distingue «aun no lo he hecho» de «se me ha pasado». Un habito
    sin marcar no cuenta como incompleto hasta que llega su hora; los de todo
    el dia no vencen hasta que acaba el dia.
    """
    hoy = ahora.date()
    for ev in eventos_de_habito(eventos, h, hoy, hoy):
        try:
            if ev.todo_dia:
                continue          # no vence hasta la noche: lo ve el recordatorio
            if ev.fin <= ahora:
                return True
        except Exception:
            continue
    return False


def habitos_lista():
    """Los habitos guardados, ya saneados: siempre con id, nombre y color."""
    salida = []
    for i, h in enumerate(CFG.get("habitos") or []):
        if not isinstance(h, dict):
            continue
        nombre = str(h.get("nombre", "")).strip()
        if not nombre:
            continue
        salida.append({
            "id": str(h.get("id") or "").strip() or f"h{i}",
            "nombre": nombre,
            "icono": str(h.get("icono", "") or "").strip(),
            "buscar": str(h.get("buscar", "") or "").strip(),
            "dias": _dias_habito(h.get("dias")),
            "color": h.get("color") or COLORES_HABITO[i % len(COLORES_HABITO)],
        })
    return salida


def _dias_habito(valor):
    """Dias de la semana en que toca (0=lunes .. 6=domingo).

    Sin nada guardado se entiende que toca todos los dias, que es como se
    comportaba antes de existir este campo.
    """
    if not isinstance(valor, (list, tuple)) or not valor:
        return [0, 1, 2, 3, 4, 5, 6]
    dias = sorted({int(d) for d in valor
                   if isinstance(d, (int, float)) and 0 <= int(d) <= 6})
    return dias or [0, 1, 2, 3, 4, 5, 6]


def toca_habito(h, fecha):
    """True si ese dia de la semana le toca a este habito."""
    return fecha.weekday() in h.get("dias", [0, 1, 2, 3, 4, 5, 6])


def resumen_dias(dias):
    """'todos los días', 'de lunes a viernes' o 'L·X·V'."""
    dias = sorted(dias or [])
    if len(dias) == 7:
        return "todos los días"
    if dias == [0, 1, 2, 3, 4]:
        return "de lunes a viernes"
    if dias == [5, 6]:
        return "fines de semana"
    return "·".join(INICIALES_DIA[d] for d in dias) or "nunca"


def nuevo_id_habito():
    """Identificador estable y unico; el nombre puede cambiar, el id no."""
    usados = {h["id"] for h in habitos_lista()}
    base = f"h{int(time.time())}"
    if base not in usados:
        return base
    for n in range(1, 1000):
        if f"{base}-{n}" not in usados:
            return f"{base}-{n}"
    return f"{base}-{os.getpid()}"


def marcas_habito(hid):
    """Marcas de ese habito: {"2026-08-29": "ok" | "no"}.

    Los dias vacios sencillamente no estan. Se admite tambien el formato
    antiguo (una lista de fechas cumplidas) para no perder lo ya apuntado.
    """
    dato = (CFG.get("habitos_hechos") or {}).get(hid)
    if isinstance(dato, dict):
        return {k: v for k, v in dato.items()
                if isinstance(k, str) and v in ("ok", "no")}
    if isinstance(dato, list):          # formato antiguo: todo eran cumplidos
        return {d: "ok" for d in dato if isinstance(d, str)}
    return {}


def estado_habito(hid, fecha):
    """"ok", "no" o "" (sin marcar) para ese dia."""
    return marcas_habito(hid).get(fecha.isoformat(), "")


def poner_estado_habito(hid, fecha, estado):
    """Fija el estado de un dia. Con "" se borra la marca."""
    marcas = marcas_habito(hid)
    iso = fecha.isoformat()
    if estado in ("ok", "no"):
        marcas[iso] = estado
    else:
        marcas.pop(iso, None)
    # ordenadas y con tope: unos 11 años por habito, y config.json no crece
    # sin fin aunque el habito dure media vida
    recortadas = dict(sorted(marcas.items())[-4000:])
    CFG.setdefault("habitos_hechos", {})[hid] = recortadas
    guardar_config_pronto()      # se marcan varias casillas seguidas
    return estado


def girar_habito(hid, fecha):
    """Clic en la casilla: vacio -> cumplido -> fallido -> vacio."""
    actual = estado_habito(hid, fecha)
    siguiente = ESTADOS_HABITO[
        (ESTADOS_HABITO.index(actual) + 1) % len(ESTADOS_HABITO)]
    return poner_estado_habito(hid, fecha, siguiente)


def habito_racha(hid, hoy=None, h=None):
    """Dias cumplidos seguidos hasta hoy, contando solo los que tocan.

    Si el habito es de lunes, miercoles y viernes, el domingo ni suma ni
    rompe: se salta. Un fallo si rompe, y un dia que tocaba y quedo sin
    marcar tambien, porque no consta que se hiciera.

    Si hoy aun no esta marcado se empieza por ayer, para que la racha no se
    caiga sola a las 00:01.
    """
    hoy = hoy or dt.date.today()
    marcas = marcas_habito(hid)
    if h is None:
        h = next((x for x in habitos_lista() if x["id"] == hid), None)
    dias_semana = (h or {}).get("dias", [0, 1, 2, 3, 4, 5, 6])

    dia = hoy
    if marcas.get(hoy.isoformat()) != "ok":
        dia = hoy - dt.timedelta(days=1)
    racha = 0
    # tope de seguridad: 10 años hacia atras, por si el habito no toca nunca
    for _ in range(3700):
        if dia.weekday() not in dias_semana:
            dia -= dt.timedelta(days=1)      # no tocaba: ni suma ni rompe
            continue
        if marcas.get(dia.isoformat()) != "ok":
            break
        racha += 1
        dia -= dt.timedelta(days=1)
    return racha


def habito_cuenta(hid, desde, hasta):
    """(cumplidos, fallidos) entre esas dos fechas, ambas incluidas."""
    a, b = desde.isoformat(), hasta.isoformat()
    marcas = marcas_habito(hid)
    ok = sum(1 for d, v in marcas.items() if a <= d <= b and v == "ok")
    no = sum(1 for d, v in marcas.items() if a <= d <= b and v == "no")
    return ok, no


def fondo_dia(i):
    """Gris claro y gris oscuro alternos: cada día se distingue del vecino."""
    if i % 2 == 0:
        return mezclar("#ffffff", C["fondo"], 0.16)
    return mezclar("#000000", C["fondo"], 0.30)


def seguir_eventos_movidos(eventos, rango=None):
    """Las cuentas atrás y el evento fijado guardan una copia de la fecha y el nombre del evento. Aquí se vuelve a mirar el evento
    (por su UID) y se ponen su fecha y su nombre de ahora. En una serie repetida solo se sigue al evento si el más cercano es uno y
    se ve claro (a menos de una semana).
    rango: (ini, fin) que cubre 'eventos' entero. Si se da, lo que estaba dentro y ya no aparece se ha borrado: se quita la cuenta
    atrás (o el fijado). Sin rango no se quita nada, porque el evento podría estar fuera de lo descargado. Devuelve True si cambió algo."""
    originales = [e for e in eventos if not e.parte]
    uids = {e.uid for e in originales}
    cambio = False

    def borrado(uid, momento):
        # «ficha-…» es el UID provisional de un evento recién creado aquí: tras descargar, Outlook le da otro
        return (rango is not None and uid not in uids and not str(uid).startswith("ficha-")
                and rango[0] <= momento < rango[1])

    def adoptar(uid, nombre, fecha):
        """UID real del evento que tenía el provisional «ficha-…»: el del mismo nombre ese día o, si no, el único con ese nombre."""
        if not str(uid).startswith("ficha-") or uid in uids:
            return uid
        iguales = [e for e in originales if sin_acentos(e.titulo).strip() == sin_acentos(nombre).strip()]
        ese_dia = [e for e in iguales if e.inicio.date().isoformat() == fecha]
        candidatos = ese_dia or iguales
        return candidatos[0].uid if len({e.uid for e in candidatos}) == 1 else uid

    def buscar(uid, momento):
        propios = [e for e in originales if e.uid == uid]
        if not propios:
            return None                          # borrado, o fuera de lo que se descarga: se queda como estaba
        propios.sort(key=lambda e: abs((e.inicio - momento).total_seconds()))
        if len(propios) == 1:
            return propios[0]
        if abs((propios[0].inicio - momento).total_seconds()) == 0:
            return propios[0]
        cerca = abs((propios[0].inicio - momento).total_seconds())
        if cerca <= 7 * 86400 and abs((propios[1].inicio - momento).total_seconds()) > cerca:
            return propios[0]
        return None

    for c in list(CFG.get("cuentas_atras") or []):
        try:
            if not c.get("uid"):
                continue
            viejo = dt.datetime.combine(dt.date.fromisoformat(c["fecha"]), dt.time(), tzinfo=TZ)
            real = adoptar(c["uid"], c.get("nombre", ""), c["fecha"])
            if real != c["uid"]:
                c["uid"] = real
                cambio = True
            if borrado(c["uid"], viejo + dt.timedelta(hours=12)):
                CFG["cuentas_atras"].remove(c)
                cambio = True
                continue
            exacto = [e for e in originales if e.uid == c["uid"] and e.inicio.date().isoformat() == c["fecha"]]
            ev = exacto[0] if exacto else buscar(c["uid"], viejo + dt.timedelta(hours=12))
            if ev is None:
                continue
            nuevo = {"fecha": ev.inicio.date().isoformat(), "hasta": ev.fin.isoformat(), "nombre": ev.titulo}
            if any(c.get(k) != v for k, v in nuevo.items()):
                c.update(nuevo)
                cambio = True
        except (ValueError, KeyError):
            continue
    a = CFG.get("anclado")
    if a and a.get("uid"):
        try:
            momento = dt.datetime.fromisoformat(a["inicio"])
            real = adoptar(a["uid"], a.get("titulo", ""), momento.date().isoformat())
            if real != a["uid"]:
                a["uid"] = real
                cambio = True
            ev = buscar(a["uid"], momento)
            if borrado(a["uid"], momento):
                CFG["anclado"] = None
                cambio = True
            elif ev is not None and (a.get("inicio") != ev.inicio.isoformat() or a.get("titulo") != ev.titulo):
                a["inicio"], a["titulo"] = ev.inicio.isoformat(), ev.titulo
                cambio = True
        except (ValueError, KeyError):
            pass
    if cambio:
        guardar_config()
    return cambio


def cuentas_vigentes():
    """Cuentas atrás que aún no han pasado (las pasadas se borran solas)."""
    ahora = dt.datetime.now(TZ)
    vivas = []
    for c in CFG.get("cuentas_atras") or []:
        try:
            if c.get("hasta"):
                ok = dt.datetime.fromisoformat(c["hasta"]) > ahora
            else:
                ok = dt.date.fromisoformat(c["fecha"]) >= ahora.date()
        except (ValueError, KeyError):
            ok = False
        if ok:
            vivas.append(c)
    if len(vivas) != len(CFG.get("cuentas_atras") or []):
        CFG["cuentas_atras"] = vivas
        guardar_config()
    return sorted(vivas, key=lambda c: c["fecha"])


def habitos_estado(eventos=None, ahora=None):
    """Resumen para el chip de la barra de arriba.

    Un hábito sin marcar NO cuenta como incompleto: puede que todavia no le
    toque. Solo cuentan en contra los marcados con cruz y los que tenias
    planificados a una hora que ya ha pasado. Asi el chip esta en verde
    mientras el dia va bien, y solo se enciende cuando de verdad falta algo.
    """
    if not CFG.get("habitos_activo", True):
        return {"activo": False, "hechos": 0, "fallados": 0, "vencidos": 0,
                "pendientes": 0, "total": 0, "racha": 0, "al_dia": True}
    ahora = ahora or dt.datetime.now(TZ)
    hoy = ahora.date()
    lista = [h for h in habitos_lista() if toca_habito(h, hoy)]
    estados = [estado_habito(h["id"], hoy) for h in lista]

    vencidos = pendientes = 0
    for h, estado in zip(lista, estados):
        if estado:
            continue
        if habito_vencido(eventos, h, ahora):
            vencidos += 1             # tenia hora y ya paso: cuenta en contra
        elif dias_planificados(eventos, h, hoy, hoy):
            pendientes += 1           # planificado, pero aun no le toca
    fallados = estados.count("no")
    return {"activo": True,
            "hechos": estados.count("ok"),
            "fallados": fallados,
            "vencidos": vencidos,
            "pendientes": pendientes,
            "total": len(lista),
            "racha": max((habito_racha(h["id"], hoy, h) for h in lista),
                         default=0),
            # «al dia» = nada fallado y nada que se haya pasado de hora
            "al_dia": not fallados and not vencidos}


# --------------------------------------------------------------------------
# Interfaz
# --------------------------------------------------------------------------

def sin_acentos(texto):
    return "".join(c for c in unicodedata.normalize("NFD", str(texto).lower())
                   if unicodedata.category(c) != "Mn")


def interpretar_fecha(texto, hoy=None):
    """'18/9', '18 sep', '18/09/2026', 'hoy', 'mañana' -> fecha (o None)."""
    hoy = hoy or dt.date.today()
    t = sin_acentos(texto).strip()
    if not t:
        return None
    if t in ("hoy", "ahora"):
        return hoy
    if t in ("manana", "mnn"):
        return hoy + dt.timedelta(days=1)
    if t in ("ayer",):
        return hoy - dt.timedelta(days=1)

    # 2026-09-01: el formato ISO. Va antes que el resto porque ahi el
    # primer numero es el año, no el dia.
    m = re.fullmatch(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", t)
    if m:
        try:
            return dt.date(int(m[1]), int(m[2]), int(m[3]))
        except ValueError:
            return None

    # 18/9, 18-9-2026, 18.9.26
    m = re.fullmatch(r"(\d{1,2})[/\-. ](\d{1,2})(?:[/\-. ](\d{2,4}))?", t)
    if m:
        dia, mes, anio = int(m[1]), int(m[2]), m[3]
    else:
        # 18 sep / 18 septiembre / sep 18
        m = re.fullmatch(r"(\d{1,2})\s+([a-z]+)(?:\s+(\d{2,4}))?", t) or \
            re.fullmatch(r"([a-z]+)\s+(\d{1,2})(?:\s+(\d{2,4}))?", t)
        if not m:
            return None
        if m[1].isdigit():
            dia, nombre, anio = int(m[1]), m[2], m[3]
        else:
            dia, nombre, anio = int(m[2]), m[1], m[3]
        mes = next((i + 1 for i, x in enumerate(MESES)
                    if nombre.startswith(x)), None)
        if mes is None:
            return None

    if anio:
        anio = int(anio)
        anio += 2000 if anio < 100 else 0
    else:
        anio = hoy.year
    try:
        fecha = dt.date(anio, mes, dia)
    except ValueError:
        return None
    # sin año, si ya pasó hace mucho se entiende el año que viene
    if not m[3] and (hoy - fecha).days > 180:
        try:
            fecha = fecha.replace(year=anio + 1)
        except ValueError:
            pass
    return fecha


def buscar_eventos(eventos, texto, tope=40):
    """Eventos cuyo titulo, lugar, notas o calendario contienen el texto."""
    aguja = sin_acentos(texto).strip()
    if not aguja:
        return []
    encontrados = []
    vistos = set()
    for ev in sorted(eventos, key=lambda e: e.inicio):
        if ev.parte:                      # trozos de continuacion, una sola vez
            continue
        pajar = sin_acentos(" ".join(
            (ev.titulo, ev.lugar or "", ev.notas or "", ev.calendario or "")))
        if aguja in pajar:
            clave = (ev.uid, ev.inicio.date())
            if clave in vistos:
                continue
            vistos.add(clave)
            encontrados.append(ev)
            if len(encontrados) >= tope:
                break
    return encontrados


def enlace_nuevo_evento(inicio, fin=None, titulo=""):
    """Enlace de Outlook web con el evento ya prerrellenado."""
    from urllib.parse import urlencode
    fin = fin or inicio + dt.timedelta(hours=1)
    parametros = {
        "path": "/calendar/action/compose",
        "rru": "addevent",
        "startdt": inicio.strftime("%Y-%m-%dT%H:%M:%S"),
        "enddt": fin.strftime("%Y-%m-%dT%H:%M:%S"),
    }
    if titulo:
        parametros["subject"] = titulo
    return ("https://outlook.live.com/calendar/0/deeplink/compose?"
            + urlencode(parametros))


def crear_en_outlook_escritorio(inicio, fin=None, titulo=""):
    """Abre la ventana de cita nueva del Outlook instalado. True si lo logra."""
    fin = fin or inicio + dt.timedelta(hours=1)
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Outlook.Application")
        cita = app.CreateItem(1)                      # olAppointmentItem
        cita.Start = inicio.strftime("%Y-%m-%d %H:%M")
        cita.Duration = max(15, int((fin - inicio).total_seconds() // 60))
        if titulo:
            cita.Subject = titulo
        cita.Display(False)                           # no bloquea el widget
        return True
    except Exception as exc:
        log(f"Outlook de escritorio no disponible ({exc}); se usa la web")
        return False


# Identificador de la app "nuevo Outlook" (la de la Store) para lanzarla.
AUMID_NUEVO_OUTLOOK = ("Microsoft.OutlookForWindows_8wekyb3d8bbwe"
                       "!Microsoft.OutlookforWindows")


def abrir_nuevo_outlook():
    """Arranca el Outlook nuevo (olk.exe). No admite prerrellenar la cita."""
    import subprocess
    for orden in (["explorer.exe", "shell:AppsFolder\\" + AUMID_NUEVO_OUTLOOK],
                  ["cmd", "/c", "start", "", "olk.exe"]):
        try:
            subprocess.Popen(orden, shell=False)
            return True
        except Exception as exc:
            log(f"No se pudo abrir el Outlook nuevo con {orden[0]}: {exc}")
    return False


def calendarios_para_escribir():
    """Nombres de los calendarios de Outlook donde se puede crear."""
    nombres = [c.get("nombre") or c.get("buscar")
               for c in (CFG.get("carpetas_outlook") or [])
               if c.get("activo", True)]
    return nombres or ["Outlook"]


# olRecursDaily=0, olRecursWeekly=1, olRecursMonthly=2, olRecursYearly=5
PATRONES = {"diaria": (0, 1, None), "laborables": (1, 1, 62),
            "semanal": (1, 1, "dia"), "quincenal": (1, 2, "dia"),
            "mensual": (2, 1, None), "anual": (5, 1, None)}
# mascara de dias de Outlook: domingo=1, lunes=2, martes=4... (62 = L a V)
DIAS_OUTLOOK = (2, 4, 8, 16, 32, 64, 1)


def aplicar_repeticion(cita, repeticion, inicio, fin):
    """Convierte la repetición elegida en el patrón de Outlook."""
    if not repeticion or repeticion.get("tipo", "no") == "no":
        return False
    tipo, cada, dias = PATRONES[repeticion["tipo"]]
    patron = cita.GetRecurrencePattern()
    patron.RecurrenceType = tipo
    patron.Interval = cada
    if dias == "dia":
        patron.DayOfWeekMask = DIAS_OUTLOOK[inicio.weekday()]
    elif dias:
        patron.DayOfWeekMask = dias
    patron.PatternStartDate = fecha_para_outlook(inicio)
    patron.StartTime = fecha_para_outlook(inicio)
    patron.EndTime = fecha_para_outlook(fin)
    if repeticion.get("veces"):
        patron.Occurrences = int(repeticion["veces"])
    elif repeticion.get("hasta"):
        patron.PatternEndDate = fecha_para_outlook(
            dt.datetime.combine(repeticion["hasta"], dt.time(0)))
    else:
        patron.NoEndDate = True
    return True


def fecha_para_outlook(momento):
    """La hora que hay que asignarle a Outlook para que guarde 'momento'.

    Al asignar una fecha, pywin32 la da por local y la pasa a UTC; Outlook
    luego se queda con ese numero como hora local. Resultado: si le damos las
    17:00 a secas, guarda las 15:00. Etiquetandola como UTC no la toca nadie
    y Outlook guarda exactamente el reloj que queriamos. Se usan fechas y no
    cadenas 'aaaa-mm-dd' porque en algunos Windows el idioma las rompe.
    """
    return momento.replace(tzinfo=dt.timezone.utc, second=0, microsecond=0)


def reloj_pedido(momento):
    """El reloj que queremos ver en Outlook, sin zona."""
    return momento.replace(tzinfo=None, second=0, microsecond=0)


def escribir_inicio(cita, inicio):
    """Pone la hora de comienzo, con la cadena como plan B."""
    try:
        cita.Start = fecha_para_outlook(inicio)
    except Exception as exc:
        log(f"Start como fecha falló ({exc}); se prueba con texto")
        cita.Start = inicio.strftime("%Y-%m-%d %H:%M")


def asegurar_hora(cita, inicio):
    """Relee la cita y, si Outlook no se quedo con la hora, la corrige.

    Asignar fechas a Outlook depende de la version de pywin32 y de la zona
    del equipo, asi que en vez de fiarnos se comprueba: si el reloj guardado
    no es el que pedimos, se vuelve a escribir desplazado justo esa
    diferencia. Devuelve True si al final quedo bien.
    """
    quiero = reloj_pedido(inicio)
    for intento in range(2):
        try:
            puesto = reloj_de(cita.Start)
        except Exception:
            return True                # si no se puede releer, no molestamos
        desfase = (quiero - puesto).total_seconds()
        if abs(desfase) <= 60:
            return True
        if intento:
            break
        log(f"Outlook guardo {puesto:%Y-%m-%d %H:%M} en vez de "
            f"{quiero:%Y-%m-%d %H:%M}; se corrige {desfase / 3600:+.0f} h")
        try:
            escribir_inicio(cita, inicio + dt.timedelta(seconds=desfase))
            cita.Save()
        except Exception as exc:
            log(f"No se pudo corregir la hora: {exc}")
            return False
    log(f"Outlook sigue sin quedarse con {quiero:%Y-%m-%d %H:%M} "
        f"(tal cual venia: {cita.Start!r})")
    return False


def guardar_cita_outlook(datos, ev=None, calendario=None):
    """Crea o cambia una cita en el Outlook instalado.

    'datos' lleva titulo, inicio, fin, todo_dia, lugar, notas y categoria.
    Devuelve (hecho, motivo si falla / EntryID si va bien).
    """
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
        if ev is not None and ev.id_outlook:
            cita = (ns.GetItemFromID(ev.id_outlook, ev.almacen_outlook)
                    if ev.almacen_outlook else ns.GetItemFromID(ev.id_outlook))
            if getattr(cita, "IsRecurring", False):
                return False, ("Se repite: cámbialo desde Outlook para no "
                               "liar la serie")
            es_nueva = False
        else:
            carpetas = _carpetas_elegidas(ns)
            elegida = next((c for nombre, _c, c in carpetas
                            if calendario in (None, nombre)), None)
            cita = (elegida or carpetas[0][2]).Items.Add(1)
            es_nueva = True

        inicio, fin = datos["inicio"], datos["fin"]
        minutos = max(int((fin - inicio).total_seconds() // 60), 15)
        cita.Subject = datos.get("titulo") or "(sin título)"
        cita.AllDayEvent = bool(datos.get("todo_dia"))
        escribir_inicio(cita, inicio)
        cita.Duration = (max(1440, (minutos // 1440) * 1440)
                         if datos.get("todo_dia") else minutos)
        cita.Location = datos.get("lugar") or ""
        cita.Body = datos.get("notas") or ""
        cita.Categories = datos.get("categoria") or ""
        aplicar_repeticion(cita, datos.get("repeticion"), inicio, fin)
        cita.Save()
        if not asegurar_hora(cita, inicio):
            if es_nueva:
                # si no, queda una cita fantasma en Outlook que el widget
                # cree que no existe
                try:
                    cita.Delete()
                except Exception as exc:
                    log(f"No se pudo deshacer la cita a medio crear: {exc}")
            return False, ("Outlook no se ha quedado con la hora nueva; "
                           "mira registro.txt")
        log(f"Guardado en Outlook: {cita.Subject} {inicio:%d/%m %H:%M}")
        return True, str(cita.EntryID)
    except Exception as exc:
        log(f"No se pudo guardar la cita: {exc}")
        texto = str(exc)
        if "win32com" in texto or "pythoncom" in texto or "Outlook.Application" in texto:
            return False, ("No encuentro el Outlook clásico instalado: sin él "
                           "no se puede escribir en el calendario")
        return False, texto[:80]


# olFree=0, olTentative=1, olBusy=2, olOutOfOffice=3
ESTADOS_OUTLOOK = {"libre": 0, "provisional": 1, "ocupado": 2, "fuera": 3}


def cambiar_en_outlook(ev, **campos):
    """Toca campos sueltos de una cita ya existente (etiqueta, estado…)."""
    if not ev.id_outlook:
        return False, ("Solo se pueden cambiar los eventos leídos del Outlook "
                       "instalado")
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        cita = (ns.GetItemFromID(ev.id_outlook, ev.almacen_outlook)
                if ev.almacen_outlook else ns.GetItemFromID(ev.id_outlook))
        if "categoria" in campos:
            cita.Categories = campos["categoria"] or ""
        if "estado" in campos:
            cita.BusyStatus = ESTADOS_OUTLOOK.get(campos["estado"], 2)
        cita.Save()
        log(f"Cambiado en Outlook: {ev.titulo} {campos}")
        return True, ""
    except Exception as exc:
        log(f"No se pudo cambiar la cita: {exc}")
        return False, str(exc)[:70]


def borrar_en_outlook(ev):
    """Manda la cita a 'Elementos eliminados' de Outlook (se puede recuperar)."""
    if not ev.id_outlook:
        return False, ("Solo se pueden borrar los eventos leídos del Outlook "
                       "instalado")
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        cita = (ns.GetItemFromID(ev.id_outlook, ev.almacen_outlook)
                if ev.almacen_outlook else ns.GetItemFromID(ev.id_outlook))
        if getattr(cita, "IsRecurring", False):
            return False, ("Se repite: bórralo desde Outlook para elegir si es "
                           "solo este día o toda la serie")
        cita.Delete()
        log(f"Enviado a Elementos eliminados: {ev.titulo} {ev.inicio:%d/%m %H:%M}")
        return True, ""
    except Exception as exc:
        log(f"No se pudo borrar la cita: {exc}")
        return False, str(exc)[:70]


def mover_en_outlook(ev, nuevo_inicio):
    """Cambia la hora de la cita en Outlook. Devuelve (hecho, motivo)."""
    if not ev.id_outlook:
        return False, ("Solo se pueden mover los eventos leídos del Outlook "
                       "instalado, no los del enlace ICS")
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        cita = (ns.GetItemFromID(ev.id_outlook, ev.almacen_outlook)
                if ev.almacen_outlook else ns.GetItemFromID(ev.id_outlook))
        if getattr(cita, "IsRecurring", False):
            return False, "Se repite: cámbialo desde Outlook para no liar la serie"
        escribir_inicio(cita, nuevo_inicio)
        cita.Save()
        if not asegurar_hora(cita, nuevo_inicio):
            return False, ("Outlook no se ha quedado con la hora nueva; "
                           "mira registro.txt")
        log(f"Movido en Outlook: {ev.titulo} -> {nuevo_inicio:%d/%m %H:%M}")
        return True, ""
    except Exception as exc:
        log(f"No se pudo mover la cita: {exc}")
        return False, str(exc)[:70]


def abrir_cita_en_outlook(ev):
    """Abre en Outlook la cita concreta. True si lo consigue."""
    if not ev.id_outlook:
        return False
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        cita = (ns.GetItemFromID(ev.id_outlook, ev.almacen_outlook)
                if ev.almacen_outlook else ns.GetItemFromID(ev.id_outlook))
        cita.Display(False)
        return True
    except Exception as exc:
        log(f"No se pudo abrir la cita en Outlook: {exc}")
        return False


# --------------------------------------------------------------------------
# Correo (Outlook clasico instalado): bandejas, mensajes y reglas
# --------------------------------------------------------------------------

class Correo:
    """Un mensaje de correo, ya recortado a lo que enseña el panel."""
    __slots__ = ("asunto", "remitente", "remitente_email", "fecha", "leido",
                 "categoria", "carpeta", "color_carpeta", "id_outlook",
                 "almacen_outlook", "adjuntos", "importancia", "marcado")

    def __init__(self, asunto, remitente, remitente_email, fecha, leido,
                 categoria, carpeta, color_carpeta, id_outlook,
                 almacen_outlook, adjuntos=0, importancia=1):
        self.asunto = asunto
        self.remitente = remitente
        self.remitente_email = remitente_email
        self.fecha = fecha
        self.leido = leido
        self.categoria = categoria
        self.carpeta = carpeta
        self.color_carpeta = color_carpeta
        self.id_outlook = id_outlook
        self.almacen_outlook = almacen_outlook
        self.adjuntos = adjuntos
        self.importancia = importancia
        self.marcado = False          # con bandera de seguimiento


def carpetas_de_correo(ns, carpeta=None, ruta="", salida=None, prof=0,
                      excluir=None):
    """Todas las carpetas de correo (no de calendario, contactos...) de
    todas las cuentas de Outlook. Mismo patrón que carpetas_de_calendario.

    Si 'excluir' trae EntryID de carpetas, esas ni se enseñan ni se mira
    dentro (para no bajar a Elementos eliminados, Spam...).
    """
    salida = [] if salida is None else salida
    if prof > 5:
        return salida
    try:
        coleccion = ns.Folders if carpeta is None else carpeta.Folders
        total = coleccion.Count
    except Exception:
        return salida
    for i in range(1, total + 1):
        try:
            f = coleccion.Item(i)
            if excluir and str(getattr(f, "EntryID", "")) in excluir:
                continue
            nombre = f"{ruta}/{f.Name}" if ruta else str(f.Name)
            try:
                if int(f.DefaultItemType) == 0:      # olMailItem
                    salida.append((nombre, f))
            except Exception:
                pass
            carpetas_de_correo(ns, f, nombre, salida, prof + 1, excluir)
        except Exception:
            continue
    return salida


def resumen_de_carpeta_correo(carpeta):
    """(cuenta, total de elementos, no leidos) de una carpeta de correo."""
    try:
        cuenta = str(getattr(getattr(carpeta, "Store", None), "DisplayName", "")
                     or "")
    except Exception:
        cuenta = ""
    try:
        total = int(carpeta.Items.Count)
    except Exception:
        total = -1
    try:
        no_leidos = int(carpeta.UnReadItemCount)
    except Exception:
        no_leidos = -1
    return cuenta, total, no_leidos


# carpetas «de serie» que no cuentan como creadas por la persona: salida,
# enviados, borradores, eliminados, spam, RSS (el 30 no existe en todas
# las versiones, por eso va aparte y con su propio try/except)
_CONST_CARPETAS_ESTANDAR = (4, 5, 16, 3, 23, 30)


def _ids_carpetas_estandar(ns):
    """EntryID de las carpetas de serie de cada cuenta (Store), para no
    confundirlas con las que la persona se ha creado a mano."""
    ids = set()
    try:
        total = ns.Stores.Count
    except Exception:
        return ids
    for i in range(1, total + 1):
        try:
            store = ns.Stores.Item(i)
        except Exception:
            continue
        for constante in _CONST_CARPETAS_ESTANDAR:
            try:
                ids.add(str(store.GetDefaultFolder(constante).EntryID))
            except Exception:
                pass
    return ids


def _carpetas_correo_automaticas(ns):
    """La bandeja de entrada de cada cuenta y las carpetas que cuelgan de
    ella (las que la persona se ha creado). Se dejan fuera las carpetas
    internas de Outlook (Conflictos, Errores del servidor, RSS, Yammer...)."""
    try:
        almacenes = [ns.Stores.Item(i) for i in range(1, ns.Stores.Count + 1)]
    except Exception:
        almacenes = []
    salida = []
    varias = len(almacenes) > 1
    for almacen in almacenes:
        try:
            bandeja = almacen.GetDefaultFolder(6)      # olFolderInbox
        except Exception:
            continue                                   # sin bandeja de entrada
        try:
            cuenta = str(almacen.DisplayName)
        except Exception:
            cuenta = "Cuenta"
        nombre = (f"{cuenta}/Bandeja de entrada" if varias
                  else "Bandeja de entrada")
        salida.append((nombre, None, bandeja))
        for n, carpeta in carpetas_de_correo(ns, bandeja, nombre):
            salida.append((n, None, carpeta))
    if not salida:
        try:
            salida = [("Bandeja de entrada", None, ns.GetDefaultFolder(6))]
        except Exception:
            pass
    return salida


def _carpetas_correo_elegidas(ns):
    """[(nombre a mostrar, color, carpeta)] segun el config.

    Sin nada configurado se enseñan la bandeja de entrada y las carpetas
    que la propia persona se haya creado (no las de serie de Outlook).
    """
    conf = list(CFG.get("carpetas_correo") or [])
    if not conf:
        return _carpetas_correo_automaticas(ns)

    todas = carpetas_de_correo(ns)
    salida = []
    for c in conf:
        if not c.get("activo", True):
            continue
        hallada = _localizar(todas, c.get("buscar"))
        if not hallada:
            log(f"Carpeta de correo no encontrada: '{c.get('buscar')}'")
            continue
        nombre_real, carpeta = hallada
        salida.append((c.get("nombre") or nombre_real, c.get("color"), carpeta))
    return salida or _carpetas_correo_automaticas(ns)


def _fecha_de_correo(it):
    """La fecha del mensaje, con la misma cautela que hora_de_cita."""
    for campo in ("ReceivedTime", "SentOn", "CreationTime"):
        try:
            return reloj_de(getattr(it, campo)).replace(tzinfo=TZ)
        except Exception:
            continue
    return dt.datetime.now(TZ)


def leer_correos_de_carpeta(carpeta, nombre_cf, color_cf, limite=15):
    """Los correos mas recientes de una carpeta, del mas nuevo al mas viejo."""
    correos = []
    try:
        items = carpeta.Items
        items.Sort("[ReceivedTime]", True)
    except Exception as exc:
        log(f"No se pudo listar '{nombre_cf}': {exc}")
        return correos
    n = 0
    for it in items:
        if n >= max(1, int(limite)):
            break
        try:
            cats = [c.strip() for c in str(getattr(it, "Categories", "") or "")
                    .split(";") if c.strip()]
            categoria = cats[0] if cats else None
            try:
                adjuntos = int(it.Attachments.Count)
            except Exception:
                adjuntos = 0
            try:
                marcado = int(getattr(it, "FlagStatus", 0)) == 2
            except Exception:
                marcado = False
            correos.append(Correo(
                asunto=str(getattr(it, "Subject", "") or "(sin asunto)"),
                remitente=str(getattr(it, "SenderName", "") or "?"),
                remitente_email=_email_remitente(it),
                fecha=_fecha_de_correo(it),
                leido=not bool(getattr(it, "UnRead", False)),
                categoria=categoria, carpeta=nombre_cf, color_carpeta=color_cf,
                id_outlook=str(getattr(it, "EntryID", "")),
                almacen_outlook=str(getattr(carpeta, "StoreID", "")),
                adjuntos=adjuntos,
                importancia=int(getattr(it, "Importance", 1))))
            correos[-1].marcado = marcado
        except Exception as exc:
            log(f"Correo ilegible en '{nombre_cf}': {exc}")
        n += 1
    return correos


def listar_carpetas_correo_outlook():
    """Todas las carpetas de correo de Outlook, para que la persona elija
    cuáles se leen: [{"nombre": "cuenta/Carpeta", "total", "no_leidos"}]."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    salida = []
    for nombre, carpeta in carpetas_de_correo(ns):
        _, total, no_leidos = resumen_de_carpeta_correo(carpeta)
        salida.append(dict(nombre=nombre, total=max(0, total),
                           no_leidos=max(0, no_leidos)))
    return salida


def descargar_correo():
    """(info de carpetas, {carpeta: [Correo,...]}) de todas las cuentas."""
    import pythoncom
    pythoncom.CoInitialize()
    ns = outlook_abierto().GetNamespace("MAPI")
    limite = CFG.get("correo_limite_lista", 15)
    info, mensajes = [], {}
    for nombre_cf, color_cf, carpeta in _carpetas_correo_elegidas(ns):
        _, total, no_leidos = resumen_de_carpeta_correo(carpeta)
        info.append(dict(nombre=nombre_cf, color=color_cf,
                          total=max(0, total), no_leidos=max(0, no_leidos)))
        mensajes[nombre_cf] = leer_correos_de_carpeta(
            carpeta, nombre_cf, color_cf, limite)
    log(f"Correo: {len(info)} carpeta(s), "
        f"{sum(c['no_leidos'] for c in info)} sin leer")
    return info, mensajes


def _item_de_correo(id_outlook, almacen_outlook):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    return (ns.GetItemFromID(id_outlook, almacen_outlook) if almacen_outlook
            else ns.GetItemFromID(id_outlook))


def _texto_de_cuerpo(texto):
    """El cuerpo del correo, legible, sin las etiquetas HTML si las trae."""
    if not texto:
        return ""
    if "<" in texto and ">" in texto:
        texto = re.sub(r"(?is)<(script|style).*?</\1>", " ", texto)
        texto = re.sub(r"(?i)<br\s*/?>", "\n", texto)
        texto = re.sub(r"(?i)</(p|div|tr|li)>", "\n", texto)
        texto = re.sub(r"<[^>]+>", "", texto)
        texto = html.unescape(texto)
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n[ \t]+", "\n", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


def cuerpo_y_marcar_correo(id_outlook, almacen_outlook):
    """Cuerpo legible del correo; lo marca como leído si así está configurado.

    Se lee con el PropertyAccessor (igual que notas_de_la_cita) para no
    despertar el aviso de seguridad de Outlook al leer '.Body'.
    """
    try:
        it = _item_de_correo(id_outlook, almacen_outlook)
    except Exception as exc:
        log(f"No se pudo abrir el correo: {exc}")
        return ""
    try:
        cuerpo = str(it.PropertyAccessor.GetProperty(PROP_CUERPO) or "")
    except Exception:
        try:
            cuerpo = str(it.Body or "")
        except Exception as exc:
            log(f"No se pudo leer el cuerpo del correo: {exc}")
            cuerpo = ""
    if CFG.get("correo_marcar_leido_al_abrir", True):
        try:
            if it.UnRead:
                it.UnRead = False
                it.Save()
        except Exception as exc:
            log(f"No se pudo marcar el correo como leído: {exc}")
    return _texto_de_cuerpo(cuerpo)


def nuevo_correo():
    """Abre una ventana de correo nuevo en Outlook, en blanco."""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Outlook.Application")
        correo = app.CreateItem(0)                    # olMailItem
        correo.Display(False)
        return True
    except Exception as exc:
        log(f"No se pudo abrir un correo nuevo: {exc}")
        return False


def marcar_todos_leidos(carpetas):
    """Marca como leídos todos los correos de esas carpetas (rutas
    «cuenta/Carpeta»). Devuelve cuántos ha podido marcar."""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as exc:
        log(f"No se pudo marcar el correo como leído: {exc}")
        return 0
    # Mismos nombres y mismas carpetas que enseña la columna (cada cuenta tiene
    # su propia «Bandeja de entrada»: antes todas se resolvían a la de la cuenta
    # por defecto y las demás no se marcaban nunca).
    por_nombre = {}
    try:
        for nombre_cf, _color, carpeta in _carpetas_correo_elegidas(ns):
            por_nombre[nombre_cf] = carpeta
    except Exception as exc:
        log(f"Marcar leídos: no se pudieron listar las carpetas: {exc}")
    todas = None
    hechos = 0
    for nombre in carpetas:
        carpeta = por_nombre.get(nombre)
        if carpeta is None:
            if todas is None:
                todas = carpetas_de_correo(ns)
            hallada = _localizar(todas, nombre)
            carpeta = hallada[1] if hallada else None
        if carpeta is None:
            log(f"Marcar leídos: carpeta no encontrada: '{nombre}'")
            continue
        try:
            items = carpeta.Items
        except Exception as exc:
            log(f"Marcar leídos: no se abre '{nombre}': {exc}")
            continue
        try:                                # solo los no leídos, no todo el buzón
            pendientes = items.Restrict("[UnRead] = True")
            n = pendientes.Count
        except Exception:
            pendientes, n = items, items.Count
        marcados_aqui = 0
        for i in range(n, 0, -1):          # de atrás hacia delante: más fiable
            try:
                it = pendientes.Item(i)
                if getattr(it, "UnRead", False):
                    it.UnRead = False
                    it.Save()
                    hechos += 1
                    marcados_aqui += 1
            except Exception:
                continue
        log(f"Marcar leídos: '{nombre}' -> {marcados_aqui} de {n}")
    return hechos


def marcar_correo(id_outlook, almacen_outlook, leido=True):
    """Marca a mano un correo como leído o no leído."""
    try:
        it = _item_de_correo(id_outlook, almacen_outlook)
        it.UnRead = not leido
        it.Save()
        return True
    except Exception as exc:
        log(f"No se pudo cambiar el estado de leído: {exc}")
        return False


def abrir_correo_en_outlook(id_outlook, almacen_outlook):
    """Abre el correo en su propia ventana de Outlook."""
    try:
        _item_de_correo(id_outlook, almacen_outlook).Display(False)
        return True
    except Exception as exc:
        log(f"No se pudo abrir el correo en Outlook: {exc}")
        return False


def responder_o_reenviar_correo(id_outlook, almacen_outlook, modo="responder"):
    """Abre en Outlook una respuesta o un reenvío, ya redactado y listo
    para que la persona lo revise y lo mande ella misma: esto NUNCA envía
    nada solo, solo abre la ventana de Outlook con el borrador."""
    try:
        it = _item_de_correo(id_outlook, almacen_outlook)
        if modo == "reenviar":
            nuevo = it.Forward()
        elif modo == "responder_todos":
            nuevo = it.ReplyAll()
        else:
            nuevo = it.Reply()
        nuevo.Display(False)
        return True
    except Exception as exc:
        log(f"No se pudo {modo.replace('_', ' ')} el correo: {exc}")
        return False


def marcar_bandera_correo(id_outlook, almacen_outlook, marcar=True):
    """Pone o quita la bandera de seguimiento de un correo."""
    try:
        it = _item_de_correo(id_outlook, almacen_outlook)
        if marcar:
            it.FlagStatus = 2                  # olFlagMarked
        else:
            it.ClearTaskFlag()
        it.Save()
        return True
    except Exception as exc:
        log(f"No se pudo cambiar la bandera del correo: {exc}")
        return False


_NOMBRES_ARCHIVO = ("archivo", "archivado", "archivados", "archive",
                    "archived", "archivo de correo")


def _carpeta_de_archivo(ns, item):
    """La carpeta «Archivo» de la misma cuenta que el correo, si existe."""
    try:
        raiz = item.Parent.Store.GetRootFolder()
        for i in range(1, raiz.Folders.Count + 1):
            f = raiz.Folders.Item(i)
            if str(f.Name).strip().lower() in _NOMBRES_ARCHIVO:
                return f
    except Exception:
        pass
    return None


def mover_correo(id_outlook, almacen_outlook, destino=None, borrar=False):
    """Archiva (destino=None), mueve a la carpeta 'destino' (ruta completa
    «cuenta/Carpeta») o elimina el correo (va a Elementos eliminados: se
    puede recuperar desde Outlook). Devuelve (ok, texto para la persona)."""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        it = (ns.GetItemFromID(id_outlook, almacen_outlook) if almacen_outlook
              else ns.GetItemFromID(id_outlook))
        if borrar:
            it.Delete()
            return True, "Correo eliminado (está en Elementos eliminados)"
        if destino is None:
            carpeta = _carpeta_de_archivo(ns, it)
            if carpeta is None:
                return False, ("No hay carpeta «Archivo» en esa cuenta: "
                               "usa «Mover a…»")
        else:
            hallada = _localizar(carpetas_de_correo(ns), destino)
            if not hallada:
                return False, f"No encuentro la carpeta «{destino}»"
            carpeta = hallada[1]
        it.Move(carpeta)
        return True, f"Movido a «{str(carpeta.Name)}»"
    except Exception as exc:
        log(f"No se pudo mover/eliminar el correo: {exc}")
        return False, "Outlook no ha podido hacerlo"


# -- reglas de correo ---------------------------------------------------

def _nombres_destinatarios(recipients):
    try:
        return [str(recipients.Item(i).Name)
                for i in range(1, recipients.Count + 1)]
    except Exception:
        return []


def _resumen_regla(regla):
    """(condiciones, acciones) de una regla, en una frase legible.

    El modelo de reglas de Outlook tiene muchos tipos de condicion y accion;
    aqui se reconocen los mas habituales y, si la regla usa algo mas raro,
    se avisa en vez de fallar.
    """
    si = []
    try:
        cond = regla.Conditions
        if cond.From.Enabled:
            si.append("de " + ", ".join(
                _nombres_destinatarios(cond.From.Recipients) or ["alguien"]))
    except Exception:
        pass
    try:
        if cond.SentTo.Enabled:
            si.append("para " + ", ".join(
                _nombres_destinatarios(cond.SentTo.Recipients) or ["alguien"]))
    except Exception:
        pass
    try:
        if cond.Subject.Enabled:
            si.append("asunto contiene " + ", ".join(
                f"«{t}»" for t in cond.Subject.Text))
    except Exception:
        pass
    try:
        if cond.Body.Enabled:
            si.append("cuerpo contiene " + ", ".join(
                f"«{t}»" for t in cond.Body.Text))
    except Exception:
        pass
    try:
        if cond.SentOnlyToMe:
            si.append("enviado solo a mí")
    except Exception:
        pass
    try:
        if cond.Category.Enabled:
            si.append("categoría " + ", ".join(
                str(c) for c in cond.Category.Categories))
    except Exception:
        pass
    try:
        if cond.Importance.Enabled:
            nombre_imp = {0: "baja", 1: "normal", 2: "alta"}.get(
                int(cond.Importance.Importance), "distinta de la normal")
            si.append(f"importancia {nombre_imp}")
    except Exception:
        pass
    try:
        if cond.HasAttachment:
            si.append("tiene adjuntos")
    except Exception:
        pass
    condiciones = ("; ".join(si) if si
                   else "condición no reconocida (revísala en Outlook)")

    no = []
    try:
        acc = regla.Actions
        if acc.MoveToFolder.Enabled:
            try:
                no.append(f"mover a «{acc.MoveToFolder.Folder.Name}»")
            except Exception:
                no.append("mover de carpeta")
    except Exception:
        pass
    try:
        if acc.CopyToFolder.Enabled:
            try:
                no.append(f"copiar a «{acc.CopyToFolder.Folder.Name}»")
            except Exception:
                no.append("copiar a otra carpeta")
    except Exception:
        pass
    try:
        if acc.Delete:
            no.append("eliminar")
    except Exception:
        pass
    try:
        if acc.AssignToCategory.Enabled:
            no.append("etiquetar como " + ", ".join(
                str(c) for c in acc.AssignToCategory.Categories))
    except Exception:
        pass
    try:
        if acc.Forward.Enabled:
            no.append("reenviar a " + ", ".join(
                _nombres_destinatarios(acc.Forward.Recipients) or ["alguien"]))
    except Exception:
        pass
    try:
        if acc.PlaySound.Enabled:
            no.append("reproducir un sonido")
    except Exception:
        pass
    try:
        if acc.Stop:
            no.append("dejar de aplicar más reglas")
    except Exception:
        pass
    try:
        if acc.MarkAsTask.Enabled:
            no.append("marcar como tarea")
    except Exception:
        pass
    try:
        if acc.NewItemAlert.Enabled:
            no.append("mostrar un aviso en pantalla")
    except Exception:
        pass
    acciones = ("; ".join(no) if no
                else "acción no reconocida (revísala en Outlook)")
    return condiciones, acciones


def listar_reglas_outlook():
    """[{cuenta, orden, nombre, activa, condiciones, acciones}] de todas
    las cuentas de Outlook."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        cuentas = ns.Stores
        total = cuentas.Count
    except Exception as exc:
        log(f"No se pudieron leer las cuentas de Outlook: {exc}")
        return []
    salida = []
    for i in range(1, total + 1):
        try:
            store = cuentas.Item(i)
            reglas = store.GetRules()
        except Exception:
            continue
        for j in range(1, reglas.Count + 1):
            try:
                r = reglas.Item(j)
                condiciones, acciones = _resumen_regla(r)
                salida.append(dict(cuenta=str(store.DisplayName),
                                   orden=int(r.ExecutionOrder),
                                   nombre=str(r.Name), activa=bool(r.Enabled),
                                   condiciones=condiciones, acciones=acciones))
            except Exception as exc:
                log(f"Regla ilegible: {exc}")
    return salida


def _con_regla(cuenta, nombre_regla, funcion):
    """Localiza una regla por cuenta+nombre y le aplica 'funcion(reglas, indice)'."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        for i in range(1, ns.Stores.Count + 1):
            store = ns.Stores.Item(i)
            if str(store.DisplayName) != cuenta:
                continue
            reglas = store.GetRules()
            for j in range(1, reglas.Count + 1):
                if str(reglas.Item(j).Name) == nombre_regla:
                    return funcion(reglas, j)
    except Exception as exc:
        log(f"No se pudo acceder a la regla '{nombre_regla}': {exc}")
    return False


def alternar_regla(cuenta, nombre_regla, activar):
    def hacer(reglas, j):
        reglas.Item(j).Enabled = bool(activar)
        reglas.Save()
        return True
    return _con_regla(cuenta, nombre_regla, hacer)


def borrar_regla(cuenta, nombre_regla):
    def hacer(reglas, j):
        reglas.Remove(j)
        reglas.Save()
        return True
    return _con_regla(cuenta, nombre_regla, hacer)


def ejecutar_regla_ahora(cuenta, nombre_regla):
    def hacer(reglas, j):
        reglas.Item(j).Execute(ShowProgress=False)
        return True
    return _con_regla(cuenta, nombre_regla, hacer)


def cuenta_atras(momento, ahora=None):
    """'en 12 dias', 'en 2 h 15 min', 'en 8 min' o 'ya ha pasado'."""
    ahora = ahora or dt.datetime.now(TZ)
    segundos = (momento - ahora).total_seconds()
    if segundos <= 0:
        return "ya ha pasado"
    dias = int(segundos // 86400)
    if dias >= 1:
        return f"en {dias} día" + ("s" if dias > 1 else "")
    horas = int(segundos // 3600)
    minutos = int((segundos % 3600) // 60)
    if horas:
        return f"en {horas} h {minutos:02d} min"
    return f"en {max(1, minutos)} min"


def crear_evento(inicio, fin=None, titulo=""):
    """Segun 'crear_con': web prerrellenada (por defecto), Outlook nuevo o clasico."""
    modo = CFG.get("crear_con", "ficha")
    if modo == "outlook" and crear_en_outlook_escritorio(inicio, fin, titulo):
        return "clasico"
    if modo == "nuevo" and abrir_nuevo_outlook():
        return "nuevo"
    # "auto" y "web": el navegador es la unica via que llega ya rellenada
    webbrowser.open(enlace_nuevo_evento(inicio, fin, titulo))
    return "web"


def redondear_media_hora(momento):
    """Sube al siguiente :00 o :30."""
    minuto = 0 if momento.minute < 30 else 30
    base = momento.replace(minute=minuto, second=0, microsecond=0)
    return base + dt.timedelta(minutes=30)


def carpeta_de_inicio():
    """Carpeta 'Inicio' de Windows, donde van los programas que arrancan solos."""
    base = os.environ.get("APPDATA", "")
    return os.path.join(base, "Microsoft", "Windows", "Start Menu",
                        "Programs", "Startup")


def accesos_de_arranque():
    """Los ficheros que harían arrancar el widget con Windows."""
    carpeta = carpeta_de_inicio()
    return (os.path.join(carpeta, "Calendario.lnk"),
            os.path.join(carpeta, "Calendario.vbs"))


def arranque_activo():
    return any(os.path.exists(r) for r in accesos_de_arranque())


def _lanzador():
    """(ejecutable sin consola, guion) con los que arranca el widget."""
    guion = os.path.join(BASE, "calendario.py")
    ejecutable = sys.executable or "pythonw.exe"
    if ejecutable.lower().endswith("python.exe"):
        alterno = ejecutable[:-10] + "pythonw.exe"
        if os.path.exists(alterno):
            ejecutable = alterno
    return ejecutable, guion


def poner_arranque(activar=True):
    """Crea o quita el arranque con Windows. Devuelve (hecho, explicación)."""
    lnk, vbs = accesos_de_arranque()
    if not activar:
        borrados = 0
        for ruta in (lnk, vbs):
            try:
                if os.path.exists(ruta):
                    os.remove(ruta)
                    borrados += 1
            except Exception as exc:
                log(f"No se pudo quitar el arranque: {exc}")
                return False, str(exc)[:70]
        log(f"Arranque con Windows quitado ({borrados} accesos)")
        return True, "Ya no arranca solo"

    carpeta = carpeta_de_inicio()
    if not os.path.isdir(carpeta):
        return False, "No encuentro la carpeta de Inicio de Windows"
    ejecutable, guion = _lanzador()
    if not os.path.exists(guion):
        return False, "No encuentro calendario.py en esta carpeta"

    try:                                   # lo normal: un acceso directo
        import win32com.client
        atajo = win32com.client.Dispatch("WScript.Shell").CreateShortCut(lnk)
        atajo.TargetPath = ejecutable
        atajo.Arguments = f'"{guion}"'
        atajo.WorkingDirectory = BASE
        if os.path.exists(ICONO_PATH):
            atajo.IconLocation = ICONO_PATH
        atajo.WindowStyle = 7              # minimizado, sin robar el foco
        atajo.Description = "Widget de calendario"
        atajo.save()
        log(f"Arranque con Windows creado en {lnk}")
        return True, "Arrancará solo al encender el ordenador"
    except Exception as exc:
        log(f"Sin acceso directo ({exc}); se prueba con un .vbs")

    try:                                   # plan B, sin pywin32 y sin consola
        with open(vbs, "w", encoding="utf-8") as f:
            f.write('Set s = CreateObject("WScript.Shell")\n')
            f.write(f's.CurrentDirectory = "{BASE}"\n')
            f.write(f's.Run """{ejecutable}"" ""{guion}""", 0, False\n')
        log(f"Arranque con Windows creado en {vbs}")
        return True, "Arrancará solo al encender el ordenador"
    except Exception as exc:
        log(f"No se pudo crear el arranque: {exc}")
        return False, str(exc)[:70]


def encajar_en_pantalla(app, x, y, ancho, alto):
    """Devuelve la geometría metida dentro del escritorio visible."""
    try:
        x0, y0, an, al = app._area_visible()
    except Exception:
        x0, y0, an, al = 0, 0, app.winfo_screenwidth(), app.winfo_screenheight()
    x = max(x0 + 4, min(int(x), x0 + an - int(ancho) - 4))
    y = max(y0 + 4, min(int(y), y0 + al - int(alto) - 4))
    return f"+{x}+{y}"


# --------------------------------------------------------------------------
# Anclaje de paneles: cada panel «secundario» (correo, hábitos, cumpleaños,
# plantillas, tareas) puede ser una ventana aparte de toda la vida, o quedar
# anclado a la izquierda o a la derecha de la ventana principal, siguiéndola
# cuando esta se mueve. Se elige con el menú del propio panel (clic derecho
# en su cabecera) y se guarda por panel en CFG["ubicacion_vistas"].
# --------------------------------------------------------------------------

UBICACIONES_PANEL = (("flotante", "Ventana aparte"),
                     ("izquierda", "Anclar a la izquierda"),
                     ("derecha", "Anclar a la derecha"))


def ubicacion_panel(clave):
    valor = CFG.get("ubicacion_vistas", {}).get(clave, "flotante")
    return valor if valor in ("flotante", "izquierda", "derecha") else "flotante"


def fijar_ubicacion_panel(clave, ubicacion, panel):
    CFG.setdefault("ubicacion_vistas", {})[clave] = ubicacion
    guardar_config()
    try:
        panel._colocado_ya = False   # que en el nuevo sitio si se recoloque
        panel.actualizar_posicion()
    except Exception as exc:
        log(f"No se pudo mover el panel '{clave}': {exc}")


def menu_ubicacion_panel(clave, panel, e):
    """Clic derecho en la cabecera de un panel: elegir dónde vive."""
    m = tk.Menu(panel, tearoff=0)
    actual = ubicacion_panel(clave)
    for valor, etiqueta in UBICACIONES_PANEL:
        m.add_command(
            label=("● " if actual == valor else "     ") + etiqueta,
            command=lambda v=valor: fijar_ubicacion_panel(clave, v, panel))
    try:
        m.tk_popup(e.x_root, e.y_root)
    finally:
        m.grab_release()
    return "break"


SIMBOLO_UBICACION = {"flotante": "❐", "izquierda": "⇤", "derecha": "⇥"}
ORDEN_UBICACION = ("flotante", "izquierda", "derecha")


def boton_ubicacion_panel(cabecera, clave, panel):
    """Botón visible en la cabecera: un clic rota flotante→izquierda→derecha,
    el clic derecho abre el menú para elegir directamente. Se coloca a la
    derecha de lo que ya haya en 'cabecera' (normalmente, antes de "✕")."""
    actual = ubicacion_panel(clave)
    b = tk.Label(cabecera, text=SIMBOLO_UBICACION.get(actual, "❐"),
                bg=C["fondo"], fg=C["texto_tenue"], font=letra(9, "bold"),
                cursor="hand2", padx=4)
    b.pack(side="right")

    def _rotar(e=None):
        siguiente = ORDEN_UBICACION[
            (ORDEN_UBICACION.index(ubicacion_panel(clave)) + 1)
            % len(ORDEN_UBICACION)]
        fijar_ubicacion_panel(clave, siguiente, panel)
        try:
            b.config(text=SIMBOLO_UBICACION.get(siguiente, "❐"))
        except Exception:
            pass
        return "break"

    b.bind("<Button-1>", _rotar)
    b.bind("<Button-3>", lambda e: menu_ubicacion_panel(clave, panel, e))
    return b


def x_panel_acoplado(app, clave, ancho):
    """La X para un panel anclado (o su primera posición si es flotante)."""
    barra = getattr(app, "_panel_accesos", None)
    if ubicacion_panel(clave) == "izquierda":
        x = app.winfo_rootx() - ancho - 6
        try:
            if barra is not None and barra.winfo_exists():
                x = min(x, barra.winfo_x() - ancho - 6)
        except Exception:
            pass
        return x
    x = app.winfo_rootx() + app.winfo_width() + 6
    try:
        if barra is not None and barra.winfo_exists():
            x = max(x, barra.winfo_x() + barra.winfo_width() + 8)
    except Exception:
        pass
    return x


class AvisoFlotante(tk.Toplevel):
    """Recuadro pequeno en una esquina que aparece antes del evento y se va solo."""

    abiertos = []

    def __init__(self, app, titulo, detalle, color=None, segundos=12, al_pulsar=None):
        super().__init__(app)
        self.app = app
        self.al_pulsar = al_pulsar
        self.overrideredirect(True)
        self.configure(bg=C["rejilla"])
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass

        ancho, alto = round(300 * ESCALA), round(78 * ESCALA)
        c = tk.Canvas(self, width=ancho - 2, height=alto - 2, bg=C["fondo"],
                      highlightthickness=0)
        c.pack(padx=1, pady=1)
        c.create_rectangle(0, 0, 4, alto, fill=color or C["evento"], outline="")
        c.create_text(14, 10, text=titulo[:70], anchor="nw", fill=C["texto"],
                      font=letra(10, "bold"), width=ancho - 44)
        c.create_text(14, alto - round(26 * ESCALA), text=detalle, anchor="nw",
                      fill=C["texto_tenue"], font=letra(8), width=ancho - 24)
        c.create_text(ancho - 14, 8, text="✕", anchor="ne",
                      fill=C["texto_tenue"], font=letra(9), tags="cerrar")
        c.tag_bind("cerrar", "<Button-1>", lambda e: self.cerrar())
        c.bind("<Button-1>", self._pulsado)

        AvisoFlotante.abiertos.append(self)
        x, y = self._sitio(ancho, alto)
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.after(int(max(2, segundos) * 1000), self.cerrar)

    def _sitio(self, ancho, alto):
        """Abajo a la derecha del escritorio, apilandose hacia arriba."""
        try:
            x0, y0, an, al = self.app._area_visible()
        except Exception:
            x0, y0, an, al = 0, 0, self.winfo_screenwidth(), self.winfo_screenheight()
        hueco = len(AvisoFlotante.abiertos) - 1
        return (x0 + an - ancho - 24,
                y0 + al - alto - 60 - hueco * (alto + 8))

    def _pulsado(self, _=None):
        if self.al_pulsar:
            try:
                self.al_pulsar()
            except Exception as exc:
                log(f"Aviso flotante: {exc}")
        self.cerrar()

    def cerrar(self):
        if self in AvisoFlotante.abiertos:
            AvisoFlotante.abiertos.remove(self)
        try:
            self.destroy()
        except Exception:
            pass


def interpretar_hora(texto, por_defecto=None):
    """'9', '9:30', '0930', '9h' -> (hora, minuto), o None si no se entiende."""
    t = str(texto or "").strip().lower().replace("h", ":").replace(".", ":")
    if not t:
        return por_defecto
    m = re.fullmatch(r"(\d{1,2})(?::(\d{1,2}))?:?", t)
    if not m:
        m = re.fullmatch(r"(\d{2})(\d{2})", t)      # 0930
        if not m:
            return por_defecto
    hora, minuto = int(m[1]), int(m[2] or 0)
    if not (0 <= hora <= 23 and 0 <= minuto <= 59):
        return por_defecto
    return hora, minuto


DIAS_SEMANA = ("lunes", "martes", "mi[eé]rcoles", "jueves", "viernes",
               "s[aá]bado", "domingo")


def interpretar_rapido(texto, hoy=None):
    """'Dentista mañana 10:30 45min @Clínica' -> (título, datos).

    datos solo trae lo que reconoce: dia, hora (h, m), minutos, todo_dia y
    lugar. Reconoce hoy, mañana, pasado mañana, un día de la semana, 18/9,
    18 sep, 10:30, 'a las 5 pm', '45min', '2 horas', 'todo el día' y '@sitio'.
    Lo que sobra queda como título. Una hora suelta como '10h' no se toca.
    """
    hoy = hoy or dt.date.today()
    datos = {}

    def quitar(patron):
        nonlocal texto
        m = re.search(patron, texto, re.I)
        if m:
            texto = texto[:m.start()] + " " + texto[m.end():]
        return m

    if quitar(r"(?<!\w)todo el d[ií]a(?!\w)"):
        datos["todo_dia"] = True
    if quitar(r"(?<!\w)pasado ma[ñn]ana(?!\w)"):
        datos["dia"] = hoy + dt.timedelta(days=2)
    elif m := quitar(r"(?<!\w)(?<!la )(?<!de )(?<!esta )(hoy|ma[ñn]ana)(?!\w)"):
        datos["dia"] = hoy + dt.timedelta(days=int(m[1].lower() != "hoy"))
    elif m := quitar(r"(?<!\w)(?:el )?(" + "|".join(DIAS_SEMANA) + r")(?!\w)"):
        i = next(i for i, d in enumerate(DIAS_SEMANA)
                 if re.fullmatch(d, m[1], re.I))
        datos["dia"] = hoy + dt.timedelta(days=(i - hoy.weekday()) % 7 or 7)
    else:
        m = re.search(r"(?<![\d:])(\d{1,2}/\d{1,2}(?:/\d{2,4})?)(?![\d:])"
                      r"|(?<!\d)(\d{1,2}\s+(?:de\s+)?[a-záéíóú]{3,})(?!\w)",
                      texto, re.I)
        fecha = m and interpretar_fecha(
            re.sub(r"\s+de\s+", " ", m[1] or m[2], flags=re.I), hoy)
        if fecha:
            datos["dia"] = fecha
            texto = texto[:m.start()] + " " + texto[m.end():]

    if m := quitar(r"(?<!\d)(\d{1,2}):(\d{2})(?!\d)\s*(am|pm)?(?!\w)"
                   r"|(?<!\w)a las? (\d{1,2})(?::(\d{2}))?\s*(am|pm)?(?!\w)"):
        hora = int(m[1] or m[4])
        minuto = int(m[2] or m[5] or 0)
        if (m[3] or m[6] or "").lower() == "pm" and hora < 12:
            hora += 12
        if hora <= 23 and minuto <= 59:
            datos["hora"] = (hora, minuto)
    if m := quitar(r"(?<!\w)(\d+)\s*min(?:utos?)?(?!\w)"):
        datos["minutos"] = int(m[1])
    elif m := quitar(r"(?<!\w)(\d+(?:[.,]\d+)?)\s*horas?(?!\w)"):
        datos["minutos"] = round(float(m[1].replace(",", ".")) * 60)

    titulo, _, lugar = re.sub(r"\s+", " ", texto).strip().partition("@")
    if lugar.strip():
        datos["lugar"] = lugar.strip()
    return titulo.strip(" ,-·"), datos


class Ficha(tk.Toplevel):
    """Formulario propio para crear o cambiar un evento sin salir del widget."""

    ANCHO = 300

    def __init__(self, app, ev=None, inicio=None, plantilla=None):
        super().__init__(app)
        self.app = app
        self.ev = ev
        base = dict(plantilla or {})
        if base and inicio:            # la plantilla se coloca en ese hueco
            duracion = base["fin"] - base["inicio"] if base.get("fin") else \
                dt.timedelta(minutes=int(CFG.get("duracion_minutos", 60) or 60))
            base["inicio"], base["fin"] = inicio, inicio + duracion
        self.overrideredirect(True)
        self.configure(bg=C["rejilla"])
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass

        arranque = base.get("inicio") or inicio or (
            ev.inicio if ev else redondear_media_hora(dt.datetime.now(TZ)))
        final = base.get("fin") or (ev.fin if ev else arranque + dt.timedelta(
            minutes=int(CFG.get("duracion_minutos", 60) or 60)))

        marco = tk.Frame(self, bg=C["fondo"], padx=10, pady=8)
        marco.pack(fill="both", expand=True, padx=1, pady=1)
        cabecera = tk.Frame(marco, bg=C["fondo"], cursor="fleur")
        cabecera.pack(fill="x", pady=(0, 6))
        titulo_panel = tk.Label(
            cabecera, text=("Cambiar evento" if ev else
                            "Copia del evento" if base else "Evento nuevo"),
            bg=C["fondo"], fg=C["texto"], font=letra(10, "bold"), anchor="w",
            cursor="fleur")
        titulo_panel.pack(side="left")
        cerrar = tk.Label(cabecera, text="✕", bg=C["fondo"], fg=C["texto_tenue"],
                          font=letra(9), cursor="hand2")
        cerrar.pack(side="right")
        cerrar.bind("<Button-1>", lambda e: self.destroy())
        for zona in (cabecera, titulo_panel):
            zona.bind("<Button-1>", self._agarrar)
            zona.bind("<B1-Motion>", self._mover_panel)

        self.titulo = self._campo(marco, "Título",
                                  ev.titulo if ev else base.get("titulo", ""))
        fila = tk.Frame(marco, bg=C["fondo"])
        fila.pack(fill="x")
        self.dia = self._campo(fila, "Día", f"{arranque:%d/%m/%Y}", 12, "left")
        self.desde = self._campo(fila, "De", f"{arranque:%H:%M}", 6, "left")
        self.hasta = self._campo(fila, "a", f"{final:%H:%M}", 6, "left")

        atajos = tk.Frame(marco, bg=C["fondo"])
        atajos.pack(fill="x", pady=(6, 0))
        tk.Label(atajos, text="Dura", bg=C["fondo"], fg=C["texto_tenue"],
                 font=letra(7)).pack(side="left", padx=(0, 4))
        for minutos, texto in ((15, "15m"), (30, "30m"), (60, "1 h"),
                               (90, "1½ h"), (120, "2 h")):
            self._chip(atajos, texto, lambda m=minutos: self._duracion(m))
        self._chip(atajos, "＋día", lambda: self._correr_dia(1))
        self._chip(atajos, "−día", lambda: self._correr_dia(-1))

        self.todo_dia = tk.BooleanVar(
            value=bool(ev.todo_dia) if ev else bool(base.get("todo_dia")))
        tk.Checkbutton(marco, text="Todo el día", variable=self.todo_dia,
                       bg=C["fondo"], fg=C["texto_tenue"], selectcolor=C["fondo"],
                       activebackground=C["fondo"], activeforeground=C["texto"],
                       font=letra(8), anchor="w", bd=0,
                       highlightthickness=0).pack(fill="x", pady=(4, 0))

        self.lugar = self._campo(
            marco, "Dónde", (ev.lugar if ev else base.get("lugar", "")) or "",
            sugerencias=app.lugares_usados())
        self.etiqueta = self._campo(
            marco, "Etiqueta",
            (ev.categoria if ev else base.get("categoria", "")) or "",
            sugerencias=app.etiquetas_usadas())
        tk.Label(marco, text="Notas", bg=C["fondo"], fg=C["texto_tenue"],
                 font=letra(7), anchor="w").pack(fill="x", pady=(6, 1))
        self.notas = tk.Text(marco, height=3, bg=C.get("boton", "#2f2f33"),
                             fg=C["texto"], insertbackground=C["texto"],
                             relief="flat", font=letra(8), wrap="word")
        self.notas.pack(fill="x")
        if (ev.notas if ev else base.get("notas")):
            self.notas.insert("1.0", ev.notas if ev else base["notas"])

        self.repeticion = None
        if ev is None:
            self.repeticion = {"tipo": "no"}
            self.repe_texto = tk.Label(
                marco, text="No se repite  ▾", bg=C.get("boton", "#2f2f33"),
                fg=C["texto"], font=letra(8), anchor="w", padx=8, pady=4)
            self.repe_texto.pack(fill="x", pady=(8, 0))
            self._activable(self.repe_texto, self._menu_repeticion)

        self.calendario = None
        opciones = calendarios_para_escribir()
        if ev is None and len(opciones) > 1:
            self.calendario = tk.StringVar(value=opciones[0])
            tk.OptionMenu(marco, self.calendario, *opciones).pack(
                fill="x", pady=(6, 0))

        self.previa = tk.Label(marco, text="", bg=C["fondo"],
                               fg=C["texto_tenue"], font=letra(7), anchor="w",
                               wraplength=self.ANCHO)
        self.previa.pack(fill="x", pady=(4, 0))
        self.aviso = tk.Label(marco, text="", bg=C["fondo"], fg=C["ahora"],
                              font=letra(7), anchor="w", wraplength=self.ANCHO)
        self.aviso.pack(fill="x")
        if ev is None:           # «Cena mañana 21:00 @Casa» rellena el resto
            self.titulo.bind("<KeyRelease>", self._vista_previa, add="+")
            self.previa.config(
                text="Tip: escribe «Dentista mañana 10:30 @Clínica» y Enter")

        botones = tk.Frame(marco, bg=C["fondo"])
        botones.pack(fill="x", pady=(8, 0))
        self._boton(botones, "Guardar", self.guardar, principal=True)
        self._boton(botones, "Cancelar", self.destroy)
        if ev is not None and ev.id_outlook:
            self._boton(botones, "Abrir en Outlook",
                        lambda: (abrir_cita_en_outlook(ev), self.destroy()))
            self._boton(botones, "Eliminar", self._eliminar)

        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Control-Return>", lambda e: self.guardar())
        self.update_idletasks()
        self.geometry(f"+{app.winfo_rootx() + 60}+{app.winfo_rooty() + BARRA + 8}")
        self.titulo.focus_force()

    # -- mover la ventanita --------------------------------------------------

    def _agarrar(self, e):
        self._agarre = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    def _mover_panel(self, e):
        if not getattr(self, "_agarre", None):
            return
        self.geometry(f"+{e.x_root - self._agarre[0]}+{e.y_root - self._agarre[1]}")

    # -- trocitos de interfaz ------------------------------------------------

    def _campo(self, padre, etiqueta, valor="", ancho=None, lado=None,
               sugerencias=()):
        caja = tk.Frame(padre, bg=C["fondo"])
        caja.pack(side=lado, fill="x", expand=not lado, pady=(6, 0),
                  padx=(0, 6) if lado else 0)
        tk.Label(caja, text=etiqueta, bg=C["fondo"], fg=C["texto_tenue"],
                 font=letra(7), anchor="w").pack(fill="x")
        linea = tk.Frame(caja, bg=C["fondo"])
        linea.pack(fill="x")
        entrada = tk.Entry(linea, bg=C.get("boton", "#2f2f33"), fg=C["texto"],
                           insertbackground=C["texto"], relief="flat",
                           font=letra(9), width=ancho or 1)
        entrada.pack(side="left", fill="x", expand=True, ipady=3)
        entrada.insert(0, valor)
        entrada.bind("<Return>", lambda e: self.guardar())
        if sugerencias:
            flecha = tk.Label(linea, text="▾", bg=C.get("boton", "#2f2f33"),
                              fg=C["texto"], font=letra(8), padx=6)
            flecha.pack(side="left", fill="y")
            flecha.bind("<Button-1>",
                        lambda e, en=entrada, s=tuple(sugerencias):
                        self._desplegar(e, en, s))
            entrada.bind("<Down>",
                         lambda e, en=entrada, s=tuple(sugerencias):
                         self._desplegar(e, en, s))
        return entrada

    def _desplegar(self, e, entrada, opciones):
        """Lista con lo que ya se ha usado antes; al elegir, rellena el campo."""
        m = tk.Menu(self, tearoff=0)
        for op in opciones:
            m.add_command(label=op, command=lambda v=op: (
                entrada.delete(0, "end"), entrada.insert(0, v)))
        m.add_separator()
        m.add_command(label="(sin etiqueta)",
                      command=lambda: entrada.delete(0, "end"))
        # bajo el campo, venga de un clic o de la tecla ↓
        m.tk_popup(entrada.winfo_rootx(),
                   entrada.winfo_rooty() + entrada.winfo_height())
        return "break"

    def _activable(self, etiqueta, accion):
        """Un Label que se usa como botón: con Tab se llega a él, se ve dónde
        está el foco, y Enter o Espacio lo pulsan igual que el ratón."""
        etiqueta.config(takefocus=True, cursor="hand2", highlightthickness=1,
                        highlightbackground=etiqueta.cget("bg"),
                        highlightcolor=C["texto"])
        pulsar = lambda e: (accion(), "break")[1]
        for tecla in ("<Button-1>", "<Return>", "<space>"):
            etiqueta.bind(tecla, pulsar)
        return etiqueta

    def _chip(self, padre, texto, accion):
        """Botoncito de los atajos (duración, correr un día…)."""
        b = tk.Label(padre, text=texto, bg=C.get("boton", "#2f2f33"),
                     fg=C["texto"], font=letra(7), padx=6, pady=2)
        b.pack(side="left", padx=(0, 3))
        return self._activable(b, accion)

    def _vista_previa(self, e=None):
        """Debajo del título: cómo se ha entendido lo que se va escribiendo."""
        _, d = interpretar_rapido(self.titulo.get())
        partes = []
        if "dia" in d:
            partes.append(f"{d['dia']:%d/%m/%Y}")
        if d.get("todo_dia"):
            partes.append("todo el día")
        if "hora" in d:
            partes.append("%02d:%02d" % d["hora"])
        if "minutos" in d:
            partes.append(f"{d['minutos']} min")
        if "lugar" in d:
            partes.append("en " + d["lugar"])
        self.previa.config(text="→ " + " · ".join(partes) if partes else "")

    def _aplicar_rapido(self):
        """Pasa a los campos lo que el título traía dentro (solo en eventos
        nuevos; al cambiar uno ya hecho el título se respeta tal cual)."""
        if self.ev is not None:
            return
        titulo, d = interpretar_rapido(self.titulo.get())
        if not d:
            return

        def poner(entrada, valor):
            entrada.delete(0, "end")
            entrada.insert(0, valor)

        antes = (interpretar_hora(self.desde.get()),
                 interpretar_hora(self.hasta.get()))
        poner(self.titulo, titulo)
        if "dia" in d:
            poner(self.dia, f"{d['dia']:%d/%m/%Y}")
        if d.get("todo_dia"):
            self.todo_dia.set(True)
        if "lugar" in d:
            poner(self.lugar, d["lugar"])
        if "hora" in d:
            poner(self.desde, "%02d:%02d" % d["hora"])
        if "hora" in d or "minutos" in d:
            h0 = interpretar_hora(self.desde.get()) or (9, 0)
            minutos = d.get("minutos")
            if not minutos:           # conserva la duración que ya había
                a, b = antes
                minutos = ((b[0] * 60 + b[1]) - (a[0] * 60 + a[1])
                           if a and b else 0)
                if minutos <= 0:
                    minutos = int(CFG.get("duracion_minutos", 60) or 60)
            fin = dt.datetime(2000, 1, 1, *h0) + dt.timedelta(minutes=minutos)
            poner(self.hasta, f"{fin:%H:%M}")

    def _duracion(self, minutos):
        """Deja la hora de fin a X minutos de la de inicio."""
        h0 = interpretar_hora(self.desde.get())
        if not h0:
            return
        fin = (dt.datetime(2000, 1, 1, *h0) + dt.timedelta(minutes=minutos))
        self.hasta.delete(0, "end")
        self.hasta.insert(0, f"{fin:%H:%M}")
        CFG["duracion_minutos"] = minutos

    def _correr_dia(self, dias):
        """Mueve el día del evento hacia delante o hacia atrás."""
        fecha = interpretar_fecha(self.dia.get())
        if not fecha:
            return
        self.dia.delete(0, "end")
        self.dia.insert(0, f"{fecha + dt.timedelta(days=dias):%d/%m/%Y}")

    # -- repeticion ----------------------------------------------------------

    def _menu_repeticion(self):
        m = tk.Menu(self, tearoff=0)
        for tipo, etiqueta in (("no", "No se repite"),
                               ("diaria", "Todos los días"),
                               ("laborables", "De lunes a viernes"),
                               ("semanal", "Cada semana"),
                               ("quincenal", "Cada dos semanas"),
                               ("mensual", "Cada mes, el mismo día"),
                               ("anual", "Cada año")):
            m.add_command(label=etiqueta,
                          command=lambda t=tipo: self._elegir_repeticion(t))
        m.tk_popup(self.repe_texto.winfo_rootx(),
                   self.repe_texto.winfo_rooty() + self.repe_texto.winfo_height())

    def _elegir_repeticion(self, tipo):
        """Tras elegir cada cuánto, pregunta hasta cuándo."""
        if tipo == "no":
            self.repeticion = {"tipo": "no"}
            self.repe_texto.config(text="No se repite  ▾")
            return
        hasta = veces = None
        try:
            from tkinter import simpledialog
            texto = simpledialog.askstring(
                "¿Hasta cuándo?",
                "Escribe una fecha final (31/12/2026) o un número de veces (10).\n"
                "Déjalo vacío para que no acabe.", parent=self)
        except Exception as exc:
            log(f"No se pudo preguntar la repetición: {exc}")
            texto = ""
        texto = (texto or "").strip()
        if texto.isdigit() and len(texto) <= 3:
            veces = int(texto)
        elif texto:
            hasta = interpretar_fecha(texto)
            if not hasta:
                self.aviso.config(text="No entiendo esa fecha final")
                return
        self.repeticion = {"tipo": tipo, "hasta": hasta, "veces": veces}
        self.repe_texto.config(text=self._texto_repeticion() + "  ▾")

    def _texto_repeticion(self):
        r = self.repeticion or {"tipo": "no"}
        nombres = {"no": "No se repite", "diaria": "Todos los días",
                   "laborables": "De lunes a viernes", "semanal": "Cada semana",
                   "quincenal": "Cada dos semanas", "mensual": "Cada mes",
                   "anual": "Cada año"}
        texto = nombres.get(r["tipo"], "No se repite")
        if r.get("veces"):
            texto += f", {r['veces']} veces"
        elif r.get("hasta"):
            texto += f", hasta el {r['hasta']:%d/%m/%Y}"
        elif r["tipo"] != "no":
            texto += ", sin fin"
        return texto

    def _boton(self, padre, texto, accion, principal=False):
        b = tk.Label(padre, text=texto, bg=C["evento"] if principal
                     else C.get("boton", "#2f2f33"),
                     fg=C["evento_texto"] if principal else C["texto"],
                     font=letra(8, "bold" if principal else "normal"),
                     padx=10, pady=4)
        b.pack(side="left", padx=(0, 6))
        return self._activable(b, accion)

    # -- guardar -------------------------------------------------------------

    def _leer(self):
        """Convierte lo escrito en fechas; devuelve el dict o None si algo falla."""
        fecha = interpretar_fecha(self.dia.get())
        if not fecha:
            self.aviso.config(text="No entiendo el día (prueba 18/09/2026)")
            return None
        if self.todo_dia.get():
            inicio = dt.datetime.combine(fecha, dt.time(0), tzinfo=TZ)
            return {"titulo": self.titulo.get().strip(), "inicio": inicio,
                    "fin": inicio + dt.timedelta(days=1), "todo_dia": True,
                    "lugar": self.lugar.get().strip(),
                    "categoria": self.etiqueta.get().strip(),
                    "notas": self.notas.get("1.0", "end").strip()}
        h0 = interpretar_hora(self.desde.get())
        h1 = interpretar_hora(self.hasta.get())
        if not h0 or not h1:
            self.aviso.config(text="No entiendo las horas (prueba 9:30)")
            return None
        inicio = dt.datetime.combine(fecha, dt.time(*h0), tzinfo=TZ)
        fin = dt.datetime.combine(fecha, dt.time(*h1), tzinfo=TZ)
        if fin <= inicio:                      # cruza la medianoche
            fin += dt.timedelta(days=1)
        return {"titulo": self.titulo.get().strip(), "inicio": inicio,
                "fin": fin, "todo_dia": False,
                "lugar": self.lugar.get().strip(),
                "categoria": self.etiqueta.get().strip(),
                "notas": self.notas.get("1.0", "end").strip()}

    def _eliminar(self):
        """Borra el evento que se está editando, preguntando antes."""
        if self.ev is None:
            return None
        if self.app._borrar_evento(self.ev):
            self.destroy()
        return None

    def guardar(self):
        self._aplicar_rapido()
        datos = self._leer()
        if not datos:
            return None
        if not datos["titulo"]:
            self.aviso.config(text="Ponle un título")
            return None
        datos["repeticion"] = self.repeticion
        calendario = self.calendario.get() if self.calendario else None
        self.app.aplicar_ficha(datos, self.ev, calendario=calendario)
        self.destroy()
        return True


class VistaAparte(tk.Toplevel):
    """Una vista del calendario en su propia ventana.

    Reutiliza el dibujo del widget: se le presta el lienzo y la vista, se
    pinta, y se devuelve todo como estaba, para no tener dos motores.
    """

    def __init__(self, app, vista, dias=None):
        super().__init__(app)
        self.app = app
        self.vista = vista
        self.dias = dias
        self.desplazamiento = app.desplazamiento
        self._tam = None
        self.nombre = app.NOMBRES_VISTA.get(vista, "Calendario")
        self.overrideredirect(True)
        self.configure(bg=C["rejilla"])
        try:
            self.attributes("-topmost", app.nivel == "encima")
        except Exception:
            pass
        self.lienzo = tk.Canvas(self, bg=C["fondo"], highlightthickness=0)
        self.lienzo.pack(fill="both", expand=True, padx=1, pady=1)
        self.lienzo.bind("<Button-1>", self._agarrar)
        self.lienzo.bind("<B1-Motion>", self._mover_ventana)
        self.lienzo.bind("<Button-3>", self._menu)
        self.lienzo.bind("<Double-Button-1>", self._doble)
        self.lienzo.bind("<MouseWheel>",
                         lambda e: self._mover(-1 if e.delta > 0 else 1))
        self.bind("<Escape>", lambda e: self.destroy())
        app._vistas_aparte.append(self)
        self.pintar()
        self.geometry(f"+{app.winfo_rootx() + 40}+{app.winfo_rooty() + 40}")

    def pintar(self):
        """Pide al widget que dibuje su vista en este lienzo."""
        app = self.app
        if getattr(app, "_pintando_aparte", False):
            return
        guardado = (app.lienzo, app.vista, app._tam, app.desplazamiento,
                    CFG.get("dias_visibles"))
        app._pintando_aparte = True
        app.lienzo, app.vista = self.lienzo, self.vista
        app._tam, app.desplazamiento = self._tam, self.desplazamiento
        app._destino = self
        if self.dias:
            CFG["dias_visibles"] = self.dias
        try:
            app.dibujar()
            self._tam = app._tam
        except Exception as exc:
            log(f"No se pudo pintar la vista aparte: {exc}")
        finally:
            app._destino = None
            app._pintando_aparte = False
            app.lienzo, app.vista, app._tam, app.desplazamiento = guardado[:4]
            if guardado[4] is not None:
                CFG["dias_visibles"] = guardado[4]

    def _mover(self, pasos):
        self.desplazamiento += pasos * max(1, self.dias or CFG["dias_visibles"])
        self.pintar()
        return "break"

    def _agarrar(self, e):
        self._agarre = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    def _mover_ventana(self, e):
        if getattr(self, "_agarre", None):
            self.geometry(f"+{e.x_root - self._agarre[0]}"
                          f"+{e.y_root - self._agarre[1]}")

    def _con_lienzo(self, fn, e):
        """Ejecuta una accion del widget mirando ESTE lienzo."""
        app = self.app
        app._lienzo_activo = self.lienzo
        try:
            return fn(e)
        finally:
            app._lienzo_activo = None

    def _menu(self, e):
        ev = self._con_lienzo(lambda _e: self.app._evento_bajo_raton(), e)
        if ev is not None:
            return self._con_lienzo(self.app._menu_evento, e)
        m = tk.Menu(self, tearoff=0)
        m.add_command(label=self.nombre, state="disabled")
        m.add_separator()
        m.add_command(label="Ir a hoy",
                      command=lambda: (setattr(self, "desplazamiento", 0),
                                       self.pintar()))
        m.add_command(label="Actualizar", command=self.pintar)
        m.add_command(label="Cerrar esta ventana", command=self.destroy)
        m.tk_popup(e.x_root, e.y_root)
        return "break"

    def _doble(self, e):
        ev = self._con_lienzo(lambda _e: self.app._evento_bajo_raton(), e)
        if ev is not None:
            return self.app._editar_evento(ev)
        return "break"

    def destroy(self):
        try:
            self.app._vistas_aparte.remove(self)
        except (ValueError, AttributeError):
            pass
        super().destroy()


class PanelAccesos(tk.Toplevel):
    """Barra estrecha pegada al widget con un acceso para cada vista.

    Cada botón abre lo suyo: las vistas cambian la ventana principal y las
    demás cosas (plantillas, buscar) se abren en su propia ventana.
    """

    ANCHO = 252

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.overrideredirect(True)
        self.configure(bg=C["rejilla"])
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass
        self.marco = tk.Frame(self, bg=C["fondo"], padx=6, pady=6)
        self.marco.pack(fill="both", expand=True, padx=1, pady=1)
        self.botones = {}
        self.ctrl = False
        self.pintar()
        self._abierto_en = time.monotonic()
        self.actualizar_posicion()

    def pintar(self):
        for hijo in self.marco.winfo_children():
            hijo.destroy()
        self.botones = {}
        cab = tk.Label(self.marco, text="Vistas", bg=C["fondo"], fg=C["texto"],
                       font=letra(9, "bold"), anchor="w", cursor="fleur")
        cab.pack(fill="x", pady=(0, 4))
        cab.bind("<Button-1>", self._agarrar)
        cab.bind("<B1-Motion>", self._mover_panel)

        abiertas = {v.vista for v in getattr(self.app, "_vistas_aparte", [])}
        for clave, texto in (("rejilla", "Calendario"), ("ambas", "Cal + lista"),
                             ("lista", "Lista"), ("mes", "Mes")):
            self.botones[clave] = self._boton(
                texto, lambda c=clave: self._vista(c),
                activo=clave in abiertas or self.app.vista == clave)
        for modo, texto in (("plantillas", "Plantillas"),
                            ("tareas", "Tareas permanentes")):
            atributo = ("_panel_plantillas" if modo == "plantillas"
                        else "_panel_tareas")
            abierto = bool(getattr(self.app, atributo, None))
            self.botones[modo] = self._boton(
                texto, lambda m=modo: self._abrir_tablero_aparte(m),
                activo=abierto)
        abierto_cumple = bool(getattr(self.app, "_panel_cumpleanos", None))
        self.botones["cumpleanos"] = self._boton(
            "🎂 Cumpleaños", lambda: self._abrir_cumpleanos_aparte(),
            activo=abierto_cumple)
        abierto_habitos = bool(getattr(self.app, "_panel_habitos", None))
        self.botones["habitos"] = self._boton(
            "✅ Hábitos", lambda: self._abrir_habitos_aparte(),
            activo=abierto_habitos)
        abierto_correo = bool(getattr(self.app, "_panel_correo", None))
        self.botones["correo"] = self._boton(
            "✉️ Correo", lambda: self._abrir_correo_aparte(),
            activo=abierto_correo)
        abierto_todos = bool(CFG.get("todos_visible"))
        self.botones["todos"] = self._boton(
            "📋 To-dos", lambda: self._abrir_todos_aparte(),
            activo=abierto_todos)
        tk.Frame(self.marco, bg=C["rejilla"], height=1).pack(fill="x", pady=5)
        self._boton("Buscar", lambda: self.app._abrir_panel_ir())
        self._boton("✉ Correo en el widget",
                    lambda: self.app._abrir_correo_columna(),
                    activo=bool(CFG.get("correo_visible")))
        self._boton("⌘ Comandos (Ctrl+K)", lambda: self.app._abrir_paleta())
        self._boton("⚙ Panel de control",
                    lambda: self.app._abrir_panel_control(),
                    activo=bool(getattr(self.app, "_panel_control", None)))
        tk.Frame(self.marco, bg=C["rejilla"], height=1).pack(fill="x", pady=5)
        tk.Label(self.marco, text="Cada vista se abre en su ventana.\n"
                                  "Con Ctrl cambia la principal.",
                 bg=C["fondo"], fg=C["texto_tenue"], font=letra(6),
                 justify="left", anchor="w").pack(fill="x", pady=(0, 4))
        self._boton("Cerrar esta barra", self.destroy)

    def _boton(self, texto, accion, activo=False):
        b = tk.Label(self.marco, text=texto,
                     bg=C["evento"] if activo else C.get("boton", "#2f2f33"),
                     fg=C["evento_texto"] if activo else C["texto"],
                     font=letra(8, "bold" if activo else None),
                     anchor="w", padx=8, pady=4)
        b.pack(fill="x", pady=1)
        b.bind("<Button-1>", lambda e: self._pulsar(e, accion))
        b.bind("<Control-Button-1>", lambda e: self._pulsar(e, accion))
        return b

    def _pulsar(self, e, accion):
        # ignora clics que lleguen justo al abrirse el panel (clics fantasma)
        if time.monotonic() - getattr(self, "_abierto_en", 0) < 0.35:
            return "break"
        self.ctrl = bool(e.state & 0x0004)
        try:
            accion()
        finally:
            self.ctrl = False
        return "break"

    def _abrir_tablero_aparte(self, modo):
        """Abre (o cierra) Plantillas/Tareas permanentes en su propia
        ventana, igual que el resto de vistas."""
        self.app._abrir_plantillas(modo)
        self.pintar()
        self.actualizar_posicion()

    def _abrir_cumpleanos_aparte(self):
        """Abre (o cierra) el panel de cumpleaños, igual que el resto."""
        self.app._abrir_cumpleanos()
        self.pintar()
        self.actualizar_posicion()

    def _abrir_todos_aparte(self):
        """Enseña/esconde la columna de to-dos, igual que el resto."""
        self.app._abrir_todos()
        self.pintar()
        self.actualizar_posicion()

    def _abrir_habitos_aparte(self):
        """Abre (o cierra) el panel de hábitos, igual que el resto."""
        self.app._abrir_habitos()
        self.pintar()
        self.actualizar_posicion()

    def _abrir_correo_aparte(self):
        """Abre (o cierra) el panel de correo, igual que el resto."""
        self.app._abrir_correo()
        self.pintar()
        self.actualizar_posicion()

    def refrescar_tablero(self):
        pass

    def _vista(self, clave):
        """Cada vista sale en su propia ventana; la principal no se toca.

        Si esa vista ya está abierta aparte, un segundo clic la cierra.
        Con Ctrl pulsado, cambia la ventana principal como antes.
        """
        if self.ctrl:
            self.app._cambiar_vista(clave)
        else:
            for v in list(getattr(self.app, "_vistas_aparte", [])):
                try:
                    if v.vista == clave and v.winfo_exists():
                        v.destroy()
                        self.pintar()
                        return
                except Exception:
                    pass
            self.app.abrir_vista_aparte(clave)
        self.pintar()

    def actualizar_posicion(self):
        """Se queda pegada al lado derecho del widget.

        Solo se mueve si de verdad cambia de sitio: cada llamada a geometry()
        la pone por delante y taparía las ventanas que haya abiertas encima.
        """
        try:
            self.update_idletasks()
            app = self.app
            x = app.winfo_rootx() + app.winfo_width() + 6
            y = app.winfo_rooty()
            alto = self.winfo_reqheight()
            if getattr(self, "_sitio", None) == (x, y, alto):
                return
            self._sitio = (x, y, alto)
            sitio = encajar_en_pantalla(app, x, y, self.ANCHO, alto)
            self.geometry(f"{self.ANCHO}x{alto}{sitio}")
        except Exception:
            pass

    def _agarrar(self, e):
        self._agarre = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    def _mover_panel(self, e):
        if getattr(self, "_agarre", None):
            self.geometry(f"+{e.x_root - self._agarre[0]}"
                          f"+{e.y_root - self._agarre[1]}")


class TableroPlantillas(tk.Canvas):
    """Las plantillas, dibujadas donde haga falta: en la barra o en su ventana.

    Arriba, las que has guardado a mano; abajo, lo que has hecho hace poco.
    Se arrastran al calendario para crear ahi una copia.
    """

    ANCHO = 236

    def __init__(self, padre, app, alto=None, modo="plantillas"):
        super().__init__(padre, bg=C["fondo"], highlightthickness=0,
                         width=self.ANCHO,
                         height=alto or round(300 * ESCALA))
        self.app = app
        self.modo = modo
        self.arrastrando = None
        self.zonas = []
        self.bind("<Button-1>", self._empezar)
        self.bind("<B1-Motion>", self._arrastrar)
        self.bind("<ButtonRelease-1>", self._soltar)
        self.bind("<Button-3>", self._menu_fila)
        self.pintar()

    # -- contenido -----------------------------------------------------------

    def _filas(self):
        """[(tipo, dato)] de lo que toca pintar según el modo."""
        if self.modo == "tareas":
            tareas = CFG.get("tareas") or []
            pendientes = [t for t in tareas if not t.get("hecha")]
            hechas = [t for t in tareas if t.get("hecha")]
            filas = [("titulo", "Tareas permanentes")]
            filas += [("tarea", t) for t in pendientes] or [
                ("vacio", "Ninguna pendiente")]
            if hechas:
                filas.append(("titulo", "Hechas"))
                filas += [("tarea", t) for t in hechas]
            return filas
        filas = [("titulo", "Mis plantillas")]
        propias = CFG.get("plantillas") or []
        filas += [("mia", p) for p in propias] or [("vacio", "Ninguna todavía")]
        filas.append(("titulo", "Recientes"))
        recientes = self.app.plantillas_recientes()
        filas += [("reciente", p) for p in recientes] or [("vacio", "Nada aún")]
        return filas

    def pintar(self):
        self.delete("all")
        self.zonas = []
        y = 4
        for tipo, dato in self._filas():
            if tipo == "titulo":
                self.create_text(4, y + 6, text=dato, anchor="w",
                                 fill=C["texto"], font=letra(8, "bold"))
                self.create_line(4, y + 17, self.ANCHO - 4, y + 17,
                                 fill=C["rejilla"])
                y += 22
                continue
            if tipo == "vacio":
                self.create_text(12, y + 7, text=dato, anchor="w",
                                 fill=C["texto_tenue"], font=letra(8))
                y += 20
                continue
            alto = round(23 * ESCALA)
            color = (self.app._color_de_categoria(dato.get("categoria"))
                     or C["evento"])
            self.create_rectangle(4, y, self.ANCHO - 4, y + alto - 3,
                                  fill=mezclar(color, C["fondo"], 0.25),
                                  outline="")
            self.create_rectangle(4, y, 8, y + alto - 3, fill=color, outline="")
            hecha = bool(dato.get("hecha"))
            self.create_text(14, y + (alto - 3) / 2, anchor="w",
                             text=("✓ " if hecha else "")
                                  + dato.get("titulo", "")[:23],
                             fill=C["texto_tenue"] if hecha else C["texto"],
                             font=letra(8, "overstrike" if hecha else None))
            self.create_text(self.ANCHO - 8, y + (alto - 3) / 2, anchor="e",
                             text=f"{int(dato.get('minutos', 60))}m",
                             fill=C["texto_tenue"], font=letra(7))
            self.zonas.append((y, y + alto - 3, tipo, dato))
            y += alto
        self.config(height=max(y + 6, round(80 * ESCALA)))

    def _fila_en(self, y):
        for y1, y2, tipo, dato in self.zonas:
            if y1 <= y <= y2:
                return tipo, dato
        return None, None

    # -- arrastrar al calendario --------------------------------------------

    def _empezar(self, e):
        _tipo, dato = self._fila_en(e.y)
        self.arrastrando = dato or None

    def _arrastrar(self, e):
        if self.arrastrando:
            self.config(cursor="hand2")

    def _soltar(self, e):
        plantilla, self.arrastrando = self.arrastrando, None
        self.config(cursor="")
        if not plantilla:
            return
        app = self.app
        x = e.x_root - app.winfo_rootx()
        y = e.y_root - app.winfo_rooty()
        if not (0 <= x <= app.winfo_width() and 0 <= y <= app.winfo_height()):
            return                      # se ha soltado fuera del calendario
        hueco = app._hueco_en(x, y)
        if hueco is None:
            return
        app.crear_desde_plantilla(plantilla, hueco)

    # -- crear, fijar y borrar ----------------------------------------------

    def nueva(self):
        if self.modo == "tareas":
            return self.nueva_tarea()
        try:
            from tkinter import simpledialog
            titulo = simpledialog.askstring("Plantilla nueva",
                                            "¿Cómo se llama la tarea?",
                                            parent=self)
            if not titulo:
                return
            minutos = simpledialog.askinteger(
                "Plantilla nueva", "¿Cuántos minutos dura?", parent=self,
                initialvalue=60, minvalue=5, maxvalue=1440) or 60
            etiqueta = simpledialog.askstring(
                "Plantilla nueva", "Etiqueta (puedes dejarlo vacío):",
                parent=self) or ""
        except Exception as exc:
            log(f"No se pudo crear la plantilla: {exc}")
            return
        otras = [p for p in (CFG.get("plantillas") or [])
                 if p.get("titulo") != titulo]
        CFG["plantillas"] = ([{"titulo": titulo.strip(), "minutos": int(minutos),
                               "lugar": "", "categoria": etiqueta.strip(),
                               "notas": "", "todo_dia": False}] + otras)[:20]
        guardar_config()
        self.pintar()

    def nueva_tarea(self):
        """Una tarea que se queda ahí hasta que la borres."""
        try:
            from tkinter import simpledialog
            titulo = simpledialog.askstring("Tarea permanente",
                                            "¿Qué hay que hacer?", parent=self)
            if not titulo:
                return
            minutos = simpledialog.askinteger(
                "Tarea permanente", "¿Cuánto suele llevarte? (minutos)",
                parent=self, initialvalue=30, minvalue=5, maxvalue=1440) or 30
        except Exception as exc:
            log(f"No se pudo crear la tarea: {exc}")
            return
        tareas = list(CFG.get("tareas") or [])
        tareas.append({"titulo": titulo.strip(), "minutos": int(minutos),
                       "lugar": "", "categoria": "", "notas": "",
                       "todo_dia": False, "hecha": False})
        CFG["tareas"] = tareas[:60]
        guardar_config()
        self.pintar()

    def _alternar_hecha(self, tarea):
        tarea["hecha"] = not tarea.get("hecha")
        guardar_config()
        self.pintar()

    def _borrar_tarea(self, tarea):
        CFG["tareas"] = [t for t in (CFG.get("tareas") or [])
                         if t is not tarea]
        guardar_config()
        self.pintar()

    def _menu_fila(self, e):
        tipo, dato = self._fila_en(e.y)
        if not dato:
            return
        if tipo == "tarea":
            m = tk.Menu(self, tearoff=0)
            m.add_command(label=dato.get("titulo", "")[:34], state="disabled")
            m.add_separator()
            m.add_command(label="Ponerla en el próximo hueco",
                          command=lambda: self.app._desde_plantilla(dato))
            m.add_command(label=("Marcarla como pendiente"
                                 if dato.get("hecha") else "Marcarla como hecha"),
                          command=lambda: self._alternar_hecha(dato))
            m.add_command(label="Guardarla como plantilla",
                          command=lambda: self._fijar(dato))
            m.add_separator()
            m.add_command(label="Borrar la tarea",
                          command=lambda: self._borrar_tarea(dato))
            m.tk_popup(e.x_root, e.y_root)
            return
        m = tk.Menu(self, tearoff=0)
        m.add_command(label=dato.get("titulo", "")[:34], state="disabled")
        m.add_separator()
        m.add_command(label="Crear ahora (próximo hueco)",
                      command=lambda: self.app._desde_plantilla(dato))
        if tipo == "reciente":
            m.add_command(label="Guardar en mis plantillas",
                          command=lambda: self._fijar(dato))
        else:
            m.add_command(label="Borrar la plantilla",
                          command=lambda: (self.app._borrar_plantilla(
                              dato.get("titulo")), self.pintar()))
        m.tk_popup(e.x_root, e.y_root)

    def _fijar(self, plantilla):
        otras = [p for p in (CFG.get("plantillas") or [])
                 if p.get("titulo") != plantilla.get("titulo")]
        CFG["plantillas"] = ([dict(plantilla)] + otras)[:20]
        guardar_config()
        self.pintar()


class PanelPlantillas(tk.Toplevel):
    """El tablero de plantillas (o de tareas permanentes) en su propia ventana."""

    TITULOS = {"plantillas": "Plantillas", "tareas": "Tareas permanentes"}

    def __init__(self, app, modo="plantillas"):
        super().__init__(app)
        self.app = app
        self.modo = modo
        self._clave_ubicacion = modo
        self._colocado_ya = False
        self.overrideredirect(True)
        self.configure(bg=C["rejilla"])
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass
        marco = tk.Frame(self, bg=C["fondo"], padx=8, pady=6)
        marco.pack(fill="both", expand=True, padx=1, pady=1)

        cabecera = tk.Frame(marco, bg=C["fondo"], cursor="fleur")
        cabecera.pack(fill="x")
        titulo = tk.Label(cabecera, text=self.TITULOS.get(modo, "Plantillas"),
                          bg=C["fondo"], fg=C["texto"], font=letra(10, "bold"),
                          cursor="fleur")
        titulo.pack(side="left")
        cerrar = tk.Label(cabecera, text="✕", bg=C["fondo"], fg=C["texto_tenue"],
                          font=letra(9), cursor="hand2")
        cerrar.pack(side="right")
        cerrar.bind("<Button-1>", lambda e: self.destroy())
        boton_ubicacion_panel(cabecera, modo, self)
        for zona in (cabecera, titulo):
            zona.bind("<Button-1>", self._agarrar)
            zona.bind("<B1-Motion>", self._mover_panel)
        cabecera.bind("<Button-3>",
                      lambda e: menu_ubicacion_panel(modo, self, e))

        if modo == "plantillas":
            ayuda = "Arrástralas al calendario para crear el evento"
        else:
            ayuda = "Tus tareas de siempre, para marcarlas como hechas"
        tk.Label(marco, text=ayuda,
                 bg=C["fondo"], fg=C["texto_tenue"], font=letra(7),
                 wraplength=TableroPlantillas.ANCHO,
                 justify="left").pack(fill="x", pady=(2, 6))

        self.tablero = TableroPlantillas(marco, app, modo=modo)
        self.tablero.pack()

        pie = tk.Frame(marco, bg=C["fondo"])
        pie.pack(fill="x", pady=(6, 0))
        for texto, accion in (("+ Nueva", self.tablero.nueva),
                              ("Actualizar", self.tablero.pintar)):
            b = tk.Label(pie, text=texto, bg=C.get("boton", "#2f2f33"),
                         fg=C["texto"], font=letra(8), padx=8, pady=3)
            b.pack(side="left", padx=(0, 6))
            b.bind("<Button-1>", lambda e, f=accion: f())

        self.bind("<Escape>", lambda e: self.destroy())
        self.actualizar_posicion()
        self.lift()
        self.after(60, self.lift)

    def pintar(self):
        self.tablero.pintar()

    def actualizar_posicion(self):
        if (ubicacion_panel(self._clave_ubicacion) == "flotante"
                and self._colocado_ya):
            return
        try:
            self.update_idletasks()
            app = self.app
            ancho = self.winfo_reqwidth()
            x = x_panel_acoplado(app, self._clave_ubicacion, ancho)
            self.geometry(encajar_en_pantalla(app, x, app.winfo_rooty(),
                                              ancho, self.winfo_reqheight()))
            self._colocado_ya = True
        except Exception as exc:
            log(f"Plantillas/tareas: no se pudo colocar el panel ({exc})")

    def _agarrar(self, e):
        self._agarre = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    def _mover_panel(self, e):
        if getattr(self, "_agarre", None):
            self.geometry(f"+{e.x_root - self._agarre[0]}"
                          f"+{e.y_root - self._agarre[1]}")


class PanelCumpleanos(tk.Toplevel):
    """Lista de cumpleaños guardados, con alta y baja rápida.

    Cada uno se dibuja también en el calendario como un evento de todo el
    día, coloreado, cada año (ver 'generar_cumpleanos').
    """

    ANCHO = 252

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self._clave_ubicacion = "cumpleanos"
        self._colocado_ya = False
        self.overrideredirect(True)
        self.configure(bg=C["rejilla"])
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass
        self.marco = tk.Frame(self, bg=C["fondo"], padx=8, pady=6)
        self.marco.pack(fill="both", expand=True, padx=1, pady=1)
        self.pintar()
        self.bind("<Escape>", lambda e: self.destroy())
        self._abierto_en = time.monotonic()
        self.actualizar_posicion()

    def pintar(self):
        for hijo in self.marco.winfo_children():
            hijo.destroy()
        cab = tk.Frame(self.marco, bg=C["fondo"], cursor="fleur")
        cab.pack(fill="x")
        tk.Label(cab, text="🎂 Cumpleaños", bg=C["fondo"], fg=C["texto"],
                 font=letra(10, "bold"), cursor="fleur").pack(side="left")
        cerrar = tk.Label(cab, text="✕", bg=C["fondo"], fg=C["texto_tenue"],
                          font=letra(9), cursor="hand2")
        cerrar.pack(side="right")
        cerrar.bind("<Button-1>", lambda e: self.destroy())
        boton_ubicacion_panel(cab, "cumpleanos", self)
        cab.bind("<Button-1>", self._agarrar)
        cab.bind("<B1-Motion>", self._mover_panel)
        cab.bind("<Button-3>",
                 lambda e: menu_ubicacion_panel("cumpleanos", self, e))

        tk.Label(self.marco, text="Salen coloreados cada año en el calendario",
                 bg=C["fondo"], fg=C["texto_tenue"], font=letra(7),
                 wraplength=self.ANCHO, justify="left").pack(fill="x",
                                                             pady=(2, 6))

        lista = sorted(CFG.get("cumpleanos", []),
                      key=lambda c: (c.get("mes", 1), c.get("dia", 1)))
        cuerpo = tk.Frame(self.marco, bg=C["fondo"])
        cuerpo.pack(fill="both")
        if not lista:
            tk.Label(cuerpo, text="Sin cumpleaños guardados todavía",
                     bg=C["fondo"], fg=C["texto_tenue"],
                     font=letra(8)).pack(anchor="w", pady=(0, 4))
        for c in lista:
            fila = tk.Frame(cuerpo, bg=C["fondo"])
            fila.pack(fill="x", pady=1)
            texto = f"{c.get('dia', 0):02d}/{c.get('mes', 0):02d}  {c.get('nombre', '')}"
            tk.Label(fila, text=texto, bg=C["fondo"], fg=C["texto"],
                     font=letra(8), anchor="w").pack(side="left", fill="x",
                                                     expand=True)
            quitar = tk.Label(fila, text="✕", bg=C["fondo"],
                              fg=C["texto_tenue"], font=letra(8),
                              cursor="hand2")
            quitar.pack(side="right")
            quitar.bind("<Button-1>", lambda e, c=c: self._quitar(c))

        tk.Frame(self.marco, bg=C["rejilla"], height=1).pack(fill="x",
                                                             pady=(6, 6))
        formulario = tk.Frame(self.marco, bg=C["fondo"])
        formulario.pack(fill="x")
        self.nombre_var = tk.StringVar()
        self.fecha_var = tk.StringVar()
        e_nombre = tk.Entry(formulario, textvariable=self.nombre_var,
                            bg=C.get("boton", "#2f2f33"), fg=C["texto"],
                            insertbackground=C["texto"], relief="flat",
                            font=letra(8), width=13)
        e_nombre.pack(side="left", padx=(0, 4))
        e_fecha = tk.Entry(formulario, textvariable=self.fecha_var,
                           bg=C.get("boton", "#2f2f33"), fg=C["texto"],
                           insertbackground=C["texto"], relief="flat",
                           font=letra(8), width=6)
        e_fecha.pack(side="left", padx=(0, 4))
        anadir = tk.Label(formulario, text="+ Añadir",
                          bg=C.get("boton", "#2f2f33"), fg=C["texto"],
                          font=letra(8), padx=6, pady=2, cursor="hand2")
        anadir.pack(side="left")
        anadir.bind("<Button-1>", lambda e: self._anadir())
        e_fecha.bind("<Return>", lambda e: self._anadir())
        e_nombre.bind("<Return>", lambda e: self._anadir())
        tk.Label(self.marco, text="Nombre y fecha en dd/mm (el año no hace falta)",
                 bg=C["fondo"], fg=C["texto_tenue"], font=letra(6),
                 wraplength=self.ANCHO, justify="left").pack(fill="x",
                                                             pady=(3, 0))

    def _anadir(self):
        nombre = self.nombre_var.get().strip()
        fecha = self.fecha_var.get().strip()
        m = re.fullmatch(r"(\d{1,2})[/\-.](\d{1,2})", fecha)
        if not nombre:
            self.nombre_var.set("")
            return
        if not m:
            self.fecha_var.set("")
            return
        dia, mes = int(m[1]), int(m[2])
        try:
            dt.date(2024, mes, dia)   # 2024: valida tambien el 29 de febrero
        except ValueError:
            self.fecha_var.set("")
            return
        CFG.setdefault("cumpleanos", []).append(
            {"nombre": nombre, "dia": dia, "mes": mes})
        guardar_config()
        self.pintar()
        self.actualizar_posicion()
        self.app.dibujar()

    def _quitar(self, c):
        try:
            CFG.get("cumpleanos", []).remove(c)
        except ValueError:
            pass
        guardar_config()
        self.pintar()
        self.actualizar_posicion()
        self.app.dibujar()

    def actualizar_posicion(self):
        if ubicacion_panel("cumpleanos") == "flotante" and self._colocado_ya:
            return
        try:
            self.update_idletasks()
            app = self.app
            x = x_panel_acoplado(app, "cumpleanos", self.ANCHO)
            y = app.winfo_rooty()
            alto = self.winfo_reqheight()
            self.geometry(encajar_en_pantalla(app, x, y, self.ANCHO, alto))
            self._colocado_ya = True
        except Exception:
            pass

    def _agarrar(self, e):
        self._agarre = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    def _mover_panel(self, e):
        if getattr(self, "_agarre", None):
            self.geometry(f"+{e.x_root - self._agarre[0]}"
                          f"+{e.y_root - self._agarre[1]}")


class PanelHabitos(tk.Toplevel):
    """Casillas de habitos: una fila por habito, una columna por dia.

    Cada casilla tiene tres estados y se gira con el raton:
        vacia  ->  cumplido (verde)  ->  fallido (rojo)  ->  vacia
    Se puede marcar cualquier dia pasado, no solo el de hoy, y todo queda
    guardado en config.json.
    """

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self._clave_ubicacion = "habitos"
        self._colocado_ya = False
        self.overrideredirect(True)
        self.configure(bg=C["rejilla"])
        self.desplazamiento = 0       # en dias, 0 = la rejilla acaba hoy
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass
        self.marco = tk.Frame(self, bg=C["fondo"], padx=8, pady=6)
        self.marco.pack(fill="both", expand=True, padx=1, pady=1)
        self.pintar()
        self.bind("<Escape>", lambda e: self.destroy())
        self._abierto_en = time.monotonic()
        self.actualizar_posicion()

    # -- ajustes de la rejilla ----------------------------------------------

    @staticmethod
    def _entero(clave, por_defecto, minimo, maximo):
        try:
            return max(minimo, min(maximo, int(CFG.get(clave, por_defecto))))
        except (TypeError, ValueError):
            return por_defecto

    def _atras(self):
        # respeta el ajuste antiguo de quien ya tuviera config guardada
        viejo = CFG.get("habitos_dias_rejilla")
        return self._entero("habitos_dias_atras", viejo or 10, 1, 45)

    def _delante(self):
        return self._entero("habitos_dias_delante", 4, 0, 21)

    def _dias(self):
        return self._atras() + 1 + self._delante()

    def _rango(self):
        """Primera y ultima fecha de la rejilla tal y como esta ahora."""
        centro = dt.date.today() - dt.timedelta(days=self.desplazamiento)
        return (centro - dt.timedelta(days=self._atras()),
                centro + dt.timedelta(days=self._delante()))

    def _mover(self, pasos):
        self.desplazamiento = max(0, self.desplazamiento + pasos * self._atras())
        self.pintar()
        self.actualizar_posicion()

    # -- dibujo --------------------------------------------------------------

    def pintar(self):
        for hijo in self.marco.winfo_children():
            hijo.destroy()
        hoy = dt.date.today()
        desde, hasta = self._rango()
        n = self._dias()
        lista = habitos_lista()

        cab = tk.Frame(self.marco, bg=C["fondo"], cursor="fleur")
        cab.pack(fill="x")
        tk.Label(cab, text="\u2705 H\u00e1bitos", bg=C["fondo"], fg=C["texto"],
                 font=letra(10, "bold"), cursor="fleur").pack(side="left")
        cerrar = tk.Label(cab, text="\u2715", bg=C["fondo"], fg=C["texto_tenue"],
                          font=letra(9), cursor="hand2")
        cerrar.pack(side="right")
        cerrar.bind("<Button-1>", lambda e: self.destroy())
        boton_ubicacion_panel(cab, "habitos", self)
        for texto, pasos in (("\u203a", -1), ("\u2039", 1)):
            if pasos < 0 and not self.desplazamiento:
                continue          # ya estamos en hoy: no hay futuro que ver
            b = tk.Label(cab, text=texto, bg=C["fondo"], fg=C["texto_tenue"],
                         font=letra(9, "bold"), cursor="hand2", padx=4)
            b.pack(side="right")
            b.bind("<Button-1>", lambda e, p=pasos: self._mover(p))
        cab.bind("<Button-1>", self._agarrar)
        cab.bind("<B1-Motion>", self._mover_panel)
        cab.bind("<Button-3>", lambda e: menu_ubicacion_panel("habitos", self, e))

        if self.desplazamiento:
            pista = f"{desde:%d/%m} \u2013 {hasta:%d/%m}"
        else:
            pista = f"{self._atras()} d\u00edas atr\u00e1s y {self._delante()} adelante"
        tk.Label(self.marco,
                 text=f"{pista}  \u00b7  clic: vac\u00edo \u2192 \u2713 \u2192 \u2715"
                      f"  \u00b7  {MARCA_PENDIENTE} planificado",
                 bg=C["fondo"], fg=C["texto_tenue"], font=letra(7), anchor="w"
                 ).pack(fill="x", pady=(2, 6))

        rejilla = tk.Frame(self.marco, bg=C["fondo"])
        rejilla.pack(fill="x")
        # cabecera: inicial del dia de la semana y numero
        for i in range(n):
            fecha = desde + dt.timedelta(days=i)
            tinta = C["texto"] if fecha == hoy else C["texto_tenue"]
            tk.Label(rejilla, text=INICIALES_DIA[fecha.weekday()], bg=C["fondo"],
                     fg=C["evento"] if fecha.weekday() >= 5 else tinta,
                     font=letra(6)).grid(row=0, column=i + 1)
            tk.Label(rejilla, text=f"{fecha.day:02d}", bg=C["fondo"], fg=tinta,
                     font=letra(6, "bold" if fecha == hoy else None)
                     ).grid(row=1, column=i + 1)
        tk.Label(rejilla, text="racha", bg=C["fondo"], fg=C["texto_tenue"],
                 font=letra(6)).grid(row=1, column=n + 1, padx=(6, 0))

        if not lista:
            tk.Label(self.marco, text="A\u00fan no hay h\u00e1bitos. A\u00f1ade uno abajo.",
                     bg=C["fondo"], fg=C["texto_tenue"], font=letra(8),
                     anchor="w").pack(fill="x", pady=(8, 2))

        for f, h in enumerate(lista):
            fila = f + 2
            marcas = marcas_habito(h["id"])
            planificados = dias_planificados(
                getattr(self.app, "eventos", []), h, desde, hasta)
            etiqueta = f"{h['icono']} {h['nombre']}".strip()
            nombre = tk.Label(rejilla, text=etiqueta[:18], bg=C["fondo"],
                              fg=C["texto"], font=letra(8), anchor="w",
                              cursor="hand2")
            nombre.grid(row=fila, column=0, sticky="w", padx=(0, 6))
            for boton in ("<Button-1>", "<Button-3>"):
                nombre.bind(boton, lambda e, x=h: self._menu_habito(e, x))

            for i in range(n):
                fecha = desde + dt.timedelta(days=i)
                estado = marcas.get(fecha.isoformat(), "")
                futuro = fecha > hoy
                toca = toca_habito(h, fecha)
                # sin marcar pero con un evento que habla del habito: pendiente
                pendiente = not estado and fecha in planificados
                marca = MARCA_HABITO.get(estado, "")
                if estado == "ok":
                    fondo, tinta = h["color"], C["evento_texto"]
                elif estado == "no":
                    fondo, tinta = C["ahora"], "#ffffff"
                elif pendiente:
                    fondo = mezclar(h["color"], C["fondo"], 0.22)
                    tinta, marca = h["color"], MARCA_PENDIENTE
                elif not toca:
                    # ese dia de la semana no toca: casi invisible, para que
                    # se lea de un vistazo el ritmo del habito
                    fondo, tinta = C["fondo"], C["texto_tenue"]
                else:
                    fondo = C["fondo"] if futuro else C.get("boton", "#2f2f33")
                    tinta = C["texto_tenue"]
                celda = tk.Label(rejilla, text=marca,
                                 width=2, bg=fondo, fg=tinta,
                                 font=letra(8, "bold"),
                                 cursor="" if futuro else "hand2",
                                 highlightthickness=1 if fecha == hoy else 0,
                                 highlightbackground=C["texto"],
                                 highlightcolor=C["texto"])
                celda.grid(row=fila, column=i + 1, padx=1, pady=1)
                if futuro:
                    continue
                # izquierdo gira; derecho vacia de golpe, sin dar dos vueltas
                celda.bind("<Button-1>", lambda e, x=h["id"], d=fecha:
                           self._girar(x, d))
                celda.bind("<Button-3>", lambda e, x=h["id"], d=fecha:
                           self._poner(x, d, ""))

            ok, fallos = habito_cuenta(h["id"], desde, hasta)
            racha = habito_racha(h["id"], hoy, h)
            tk.Label(rejilla, text=str(racha) if racha else "\u00b7",
                     bg=C["fondo"],
                     fg=h["color"] if racha else C["texto_tenue"],
                     font=letra(8, "bold" if racha else None)
                     ).grid(row=fila, column=n + 1, padx=(6, 0))
            resumen = f"{ok}\u2713 {fallos}\u2715"
            sueltos = sum(1 for d in planificados
                          if not marcas.get(d.isoformat()))
            if sueltos:
                resumen += f" {sueltos}{MARCA_PENDIENTE}"
            tk.Label(rejilla, text=resumen, bg=C["fondo"],
                     fg=C["texto_tenue"], font=letra(6)
                     ).grid(row=fila, column=n + 2, padx=(6, 0))

        tk.Frame(self.marco, bg=C["rejilla"], height=1).pack(fill="x",
                                                             pady=(8, 6))
        formulario = tk.Frame(self.marco, bg=C["fondo"])
        formulario.pack(fill="x")
        self.nombre_var = tk.StringVar()
        self.icono_var = tk.StringVar()
        e_nombre = tk.Entry(formulario, textvariable=self.nombre_var,
                            bg=C.get("boton", "#2f2f33"), fg=C["texto"],
                            insertbackground=C["texto"], relief="flat",
                            font=letra(8), width=16)
        e_nombre.pack(side="left", padx=(0, 4))
        e_icono = tk.Entry(formulario, textvariable=self.icono_var,
                           bg=C.get("boton", "#2f2f33"), fg=C["texto"],
                           insertbackground=C["texto"], relief="flat",
                           font=letra(8), width=3)
        e_icono.pack(side="left", padx=(0, 4))
        anadir = tk.Label(formulario, text="+ A\u00f1adir",
                          bg=C.get("boton", "#2f2f33"), fg=C["texto"],
                          font=letra(8), padx=6, pady=2, cursor="hand2")
        anadir.pack(side="left")
        anadir.bind("<Button-1>", lambda e: self._anadir())
        e_nombre.bind("<Return>", lambda e: self._anadir())
        e_icono.bind("<Return>", lambda e: self._anadir())
        tk.Label(self.marco,
                 text="Nombre y, si quieres, un emoji  \u00b7  clic en el nombre "
                      "del h\u00e1bito para quitarlo o cambiar los d\u00edas que se ven",
                 bg=C["fondo"], fg=C["texto_tenue"], font=letra(6),
                 wraplength=max(self.winfo_reqwidth(), 260), justify="left"
                 ).pack(fill="x", pady=(3, 0))

    # -- acciones ------------------------------------------------------------

    def _refrescar(self):
        self.pintar()
        self.actualizar_posicion()
        self.app.dibujar()

    def _girar(self, hid, fecha):
        girar_habito(hid, fecha)
        self._refrescar()

    def _poner(self, hid, fecha, estado):
        poner_estado_habito(hid, fecha, estado)
        self._refrescar()

    def _menu_habito(self, e, h):
        m = tk.Menu(self, tearoff=0)
        hoy = dt.date.today()
        m.add_command(
            label=f"{h['nombre']}  \u00b7  racha {habito_racha(h['id'], hoy, h)}"
                  f"  \u00b7  {resumen_dias(h['dias'])}",
            state="disabled")
        m.add_separator()
        for estado, texto in (("ok", "Marcar hoy como cumplido \u2713"),
                              ("no", "Marcar hoy como fallido \u2715"),
                              ("", "Dejar hoy sin marcar")):
            m.add_command(label=texto,
                          command=lambda s=estado: self._poner(h["id"], hoy, s))
        m.add_separator()
        colores = tk.Menu(m, tearoff=0)
        for color in COLORES_HABITO:
            colores.add_command(
                label=("\u25cf " if h["color"] == color else "     ") + color,
                command=lambda c=color: self._recolorear(h["id"], c))
        m.add_cascade(label="Color", menu=colores)
        cuando = tk.Menu(m, tearoff=0)
        for etiqueta, dias in (("Todos los días", [0, 1, 2, 3, 4, 5, 6]),
                               ("De lunes a viernes", [0, 1, 2, 3, 4]),
                               ("Fines de semana", [5, 6]),
                               ("Lunes, miércoles y viernes", [0, 2, 4]),
                               ("Martes y jueves", [1, 3])):
            cuando.add_command(
                label=("\u25cf " if h["dias"] == dias else "     ") + etiqueta,
                command=lambda d=dias: self._cambiar_dias_habito(h["id"], d))
        cuando.add_separator()
        for i, nombre in enumerate(DIAS):
            marcado = i in h["dias"]
            cuando.add_command(
                label=("\u2713 " if marcado else "     ") + nombre,
                command=lambda d=i: self._alternar_dia_habito(h["id"], d))
        m.add_cascade(label=f"Qué días ({resumen_dias(h['dias'])})", menu=cuando)
        m.add_command(
            label=f"Palabras que lo detectan ({h['buscar'] or h['nombre']})\u2026",
            command=lambda: self._pedir_palabras(h))
        atras = tk.Menu(m, tearoff=0)
        for valor in (7, 10, 14, 21, 30, 45):
            atras.add_command(
                label=("\u25cf " if self._atras() == valor else "     ")
                      + f"{valor} d\u00edas",
                command=lambda v=valor: self._cambiar_dias("atras", v))
        m.add_cascade(label=f"D\u00edas hacia atr\u00e1s ({self._atras()})", menu=atras)
        delante = tk.Menu(m, tearoff=0)
        for valor in (0, 2, 4, 7, 14):
            delante.add_command(
                label=("\u25cf " if self._delante() == valor else "     ")
                      + (f"{valor} d\u00edas" if valor else "ninguno"),
                command=lambda v=valor: self._cambiar_dias("delante", v))
        m.add_cascade(label=f"D\u00edas hacia delante ({self._delante()})",
                      menu=delante)
        m.add_separator()
        m.add_command(label=f"Quitar \u00ab{h['nombre']}\u00bb y su historial",
                      command=lambda: self._quitar(h))
        m.tk_popup(e.x_root, e.y_root)

    def _cambiar_dias_habito(self, hid, dias):
        for h in (CFG.get("habitos") or []):
            if isinstance(h, dict) and h.get("id") == hid:
                h["dias"] = sorted(set(dias)) or [0, 1, 2, 3, 4, 5, 6]
        guardar_config()
        self._refrescar()

    def _alternar_dia_habito(self, hid, dia):
        actual = next((h["dias"] for h in habitos_lista() if h["id"] == hid),
                      [0, 1, 2, 3, 4, 5, 6])
        nuevos = [d for d in actual if d != dia]
        if len(nuevos) == len(actual):
            nuevos.append(dia)
        if not nuevos:
            return              # un habito sin ningun dia no tendria sentido
        self._cambiar_dias_habito(hid, nuevos)

    def _recolorear(self, hid, color):
        for h in (CFG.get("habitos") or []):
            if isinstance(h, dict) and h.get("id") == hid:
                h["color"] = color
        guardar_config()
        self.pintar()
        self.actualizar_posicion()

    def _cambiar_dias(self, lado, valor):
        CFG[f"habitos_dias_{lado}"] = valor
        CFG.pop("habitos_dias_rejilla", None)   # ya no se usa el ajuste viejo
        guardar_config()
        self.pintar()
        self.actualizar_posicion()

    def _pedir_palabras(self, h):
        """Que palabras del titulo de un evento delatan a este habito."""
        try:
            from tkinter import simpledialog
            texto = simpledialog.askstring(
                f"Detectar «{h['nombre']}»",
                "Palabras que buscar en el título de los eventos,\n"
                "separadas por comas (por ejemplo: gym, gimnasio, pesas).\n"
                "Déjalo vacío para usar solo el nombre del hábito:",
                parent=self, initialvalue=h["buscar"])
        except Exception as exc:
            log(f"Hábitos: no se pudieron pedir las palabras ({exc})")
            return
        if texto is None:
            return
        for guardado in (CFG.get("habitos") or []):
            if isinstance(guardado, dict) and guardado.get("id") == h["id"]:
                guardado["buscar"] = texto.strip()
        guardar_config()
        self._refrescar()

    def _anadir(self):
        nombre = self.nombre_var.get().strip()
        if not nombre:
            return
        actuales = CFG.setdefault("habitos", [])
        actuales.append({"id": nuevo_id_habito(), "nombre": nombre[:30],
                         "icono": self.icono_var.get().strip()[:2],
                         "color": COLORES_HABITO[len(actuales)
                                                 % len(COLORES_HABITO)]})
        guardar_config()
        log(f"H\u00e1bito nuevo: {nombre}")
        self._refrescar()

    def _quitar(self, h):
        CFG["habitos"] = [x for x in (CFG.get("habitos") or [])
                          if not (isinstance(x, dict)
                                  and x.get("id") == h["id"])]
        marcas = len(marcas_habito(h["id"]))
        (CFG.get("habitos_hechos") or {}).pop(h["id"], None)
        guardar_config()
        log(f"H\u00e1bito \u00ab{h['nombre']}\u00bb borrado junto con {marcas} marcas")
        self._refrescar()

    # -- colocacion ----------------------------------------------------------

    def actualizar_posicion(self):
        if ubicacion_panel("habitos") == "flotante" and self._colocado_ya:
            return
        try:
            self.update_idletasks()
            app = self.app
            ancho = max(self.winfo_reqwidth(), 260)
            x = x_panel_acoplado(app, "habitos", ancho)
            y = app.winfo_rooty()
            self.geometry(encajar_en_pantalla(
                app, x, y, ancho, self.winfo_reqheight()))
            self._colocado_ya = True
        except Exception as exc:
            log(f"H\u00e1bitos: no se pudo colocar el panel ({exc})")

    def _agarrar(self, e):
        self._agarre = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    def _mover_panel(self, e):
        if getattr(self, "_agarre", None):
            self.geometry(f"+{e.x_root - self._agarre[0]}"
                          f"+{e.y_root - self._agarre[1]}")


class PanelCorreo(tk.Toplevel):
    """Bandeja de correo y reglas de Outlook, en su propia ventana.

    Dos modos, con pestañas arriba:
      - «bandeja»: las carpetas configuradas y sus correos mas recientes.
      - «reglas»: las reglas de Outlook, para activarlas/desactivarlas,
        ejecutarlas a mano o borrarlas (se cargan al entrar en este modo).
    """

    ANCHO = 340

    def __init__(self, app, modo="bandeja"):
        super().__init__(app)
        self.app = app
        self.modo = modo
        self.carpeta_activa = None
        self.reglas = []
        self.busqueda_var = tk.StringVar()
        self.busqueda_var.trace_add("write", lambda *a: self._al_escribir_busqueda())
        self._clave_ubicacion = "correo"
        self._colocado_ya = False
        self.overrideredirect(True)
        self.configure(bg=C["rejilla"])
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass
        self.marco = tk.Frame(self, bg=C["fondo"], padx=8, pady=6)
        self.marco.pack(fill="both", expand=True, padx=1, pady=1)
        self.pintar()
        self.bind("<Escape>", lambda e: self.destroy())
        self._abierto_en = time.monotonic()
        self.actualizar_posicion()
        if modo == "reglas":
            self._cargar_reglas()

    def cambiar_modo(self, modo):
        self.modo = modo
        if modo == "reglas" and not self.reglas:
            self._cargar_reglas()
        else:
            self.pintar()
        self.actualizar_posicion()

    # -- cabecera comun a los dos modos --------------------------------------

    def _cabecera(self):
        cab = tk.Frame(self.marco, bg=C["fondo"], cursor="fleur")
        cab.pack(fill="x")
        tk.Label(cab, text="✉ Correo", bg=C["fondo"], fg=C["texto"],
                 font=letra(10, "bold"), cursor="fleur").pack(side="left")
        cerrar = tk.Label(cab, text="✕", bg=C["fondo"], fg=C["texto_tenue"],
                          font=letra(9), cursor="hand2")
        cerrar.pack(side="right")
        cerrar.bind("<Button-1>", lambda e: self.destroy())
        refrescar = tk.Label(cab, text="⟳", bg=C["fondo"], fg=C["texto_tenue"],
                             font=letra(9, "bold"), cursor="hand2", padx=4)
        refrescar.pack(side="right")
        refrescar.bind("<Button-1>", lambda e: self._refrescar())
        boton_ubicacion_panel(cab, "correo", self)
        cab.bind("<Button-1>", self._agarrar)
        cab.bind("<B1-Motion>", self._mover_panel)
        cab.bind("<Button-3>", lambda e: menu_ubicacion_panel("correo", self, e))

        pestañas = tk.Frame(self.marco, bg=C["fondo"])
        pestañas.pack(fill="x", pady=(4, 6))
        for clave, texto in (("bandeja", "Bandeja"), ("reglas", "Reglas")):
            activa = self.modo == clave
            b = tk.Label(pestañas, text=texto,
                         bg=C["evento"] if activa else C.get("boton", "#2f2f33"),
                         fg=C["evento_texto"] if activa else C["texto"],
                         font=letra(8, "bold" if activa else None),
                         padx=10, pady=3, cursor="hand2")
            b.pack(side="left", padx=(0, 4))
            b.bind("<Button-1>", lambda e, c=clave: self.cambiar_modo(c))

    def _refrescar(self):
        if self.modo == "reglas":
            self._cargar_reglas()
        else:
            self.app._correo_ahora()   # el propio refresco repinta al terminar

    def pintar(self):
        for hijo in self.marco.winfo_children():
            hijo.destroy()
        self._cabecera()

        zona = tk.Frame(self.marco, bg=C["fondo"])
        zona.pack(fill="both", expand=True)
        self._lienzo_scroll = tk.Canvas(zona, bg=C["fondo"],
                                        highlightthickness=0)
        barra = tk.Scrollbar(zona, orient="vertical",
                             command=self._lienzo_scroll.yview)
        self._lienzo_scroll.configure(yscrollcommand=barra.set)
        self._lienzo_scroll.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.contenido = tk.Frame(self._lienzo_scroll, bg=C["fondo"])
        id_ventana = self._lienzo_scroll.create_window(
            (0, 0), window=self.contenido, anchor="nw")

        if self.modo == "reglas":
            self._pintar_reglas()
        else:
            self._pintar_bandeja()

        def _al_cambiar_contenido(e=None):
            try:
                self._lienzo_scroll.configure(
                    scrollregion=self._lienzo_scroll.bbox("all"))
            except Exception:
                pass
        self.contenido.bind("<Configure>", _al_cambiar_contenido)

        def _al_cambiar_ancho(e):
            self._lienzo_scroll.itemconfig(id_ventana, width=e.width)
        self._lienzo_scroll.bind("<Configure>", _al_cambiar_ancho)

        def _rueda(e):
            self._lienzo_scroll.yview_scroll(-1 * int(e.delta / 120), "units")

        def _entrar(e):
            self._lienzo_scroll.bind_all("<MouseWheel>", _rueda)

        def _salir(e):
            self._lienzo_scroll.unbind_all("<MouseWheel>")
        self._lienzo_scroll.bind("<Enter>", _entrar)
        self._lienzo_scroll.bind("<Leave>", _salir)

        self.update_idletasks()
        alto_contenido = self.contenido.winfo_reqheight()
        self._lienzo_scroll.configure(height=min(420, max(alto_contenido, 40)))

    # -- modo bandeja ---------------------------------------------------------

    def _pintar_bandeja(self):
        carpetas = self.app.correo_carpetas
        if self.app.correo_error:
            tk.Label(self.contenido, text=f"⚠ {self.app.correo_error}",
                     bg=C["fondo"], fg=C["ahora"], font=letra(8),
                     wraplength=self.ANCHO - 20, justify="left"
                     ).pack(fill="x", pady=(0, 6))
        if not carpetas:
            tk.Label(self.contenido,
                     text=("Cargando…" if not self.app.correo_error
                          else "Sin datos todavía."),
                     bg=C["fondo"], fg=C["texto_tenue"], font=letra(8)
                     ).pack(fill="x", pady=8)
            return

        nombres = {c["nombre"] for c in carpetas}
        if self.carpeta_activa not in nombres:
            self.carpeta_activa = carpetas[0]["nombre"]

        if len(carpetas) > 1:
            tk.Label(self.contenido, text=f"Carpetas ({len(carpetas)})",
                     bg=C["fondo"], fg=C["texto_tenue"], font=letra(7, "bold"),
                     anchor="w").pack(fill="x", pady=(0, 2))
            for c in carpetas:
                activa = c["nombre"] == self.carpeta_activa
                fondo_fila = C["evento"] if activa else C["fondo"]
                fila_c = tk.Frame(self.contenido, bg=fondo_fila, cursor="hand2")
                fila_c.pack(fill="x", pady=1)
                tk.Label(fila_c, text=c["nombre"], bg=fondo_fila,
                         fg=C["evento_texto"] if activa else C["texto"],
                         font=letra(8, "bold" if activa else None),
                         anchor="w", padx=6, pady=3
                         ).pack(side="left", fill="x", expand=True)
                if c["no_leidos"]:
                    tk.Label(fila_c, text=str(c["no_leidos"]),
                             bg=fondo_fila if activa else "#ff453a",
                             fg="#ffffff", font=letra(7, "bold"), padx=5
                             ).pack(side="right", padx=(0, 6))
                for w in (fila_c, *fila_c.winfo_children()):
                    w.bind("<Button-1>",
                           lambda e, n=c["nombre"]: self._cambiar_carpeta(n))
            tk.Frame(self.contenido, bg=C["rejilla"], height=1).pack(
                fill="x", pady=(4, 6))

        fila_busq = tk.Frame(self.contenido, bg=C["fondo"])
        fila_busq.pack(fill="x", pady=(0, 4))
        tk.Label(fila_busq, text="🔎", bg=C["fondo"], fg=C["texto_tenue"],
                 font=letra(8)).pack(side="left")
        entrada = tk.Entry(fila_busq, textvariable=self.busqueda_var,
                           bg=C.get("boton", "#2f2f33"), fg=C["texto"],
                           insertbackground=C["texto"], relief="flat",
                           font=letra(8))
        entrada.pack(side="left", fill="x", expand=True, padx=(4, 4))
        quitar = tk.Label(fila_busq, text="✕", bg=C["fondo"],
                          fg=C["texto_tenue"], font=letra(8), cursor="hand2")
        quitar.pack(side="right")
        quitar.bind("<Button-1>", lambda e: self.busqueda_var.set(""))

        self._marco_lista = tk.Frame(self.contenido, bg=C["fondo"])
        self._marco_lista.pack(fill="both", expand=True)
        self._pintar_lista_mensajes()

    def _mensajes_filtrados(self):
        todos = self.app.correo_mensajes.get(self.carpeta_activa, [])
        termino = self.busqueda_var.get().strip().lower()
        if not termino:
            return todos, termino
        filtrados = [m for m in todos if termino in m.remitente.lower()
                    or termino in m.asunto.lower()
                    or termino in (m.remitente_email or "").lower()]
        return filtrados, termino

    def _pintar_lista_mensajes(self):
        """Solo la lista de correos (no la cabecera ni el buscador), para
        poder refiltrar sin que el cuadro de búsqueda pierda el foco."""
        marco = getattr(self, "_marco_lista", None)
        if marco is None or not marco.winfo_exists():
            return
        for hijo in marco.winfo_children():
            hijo.destroy()

        mensajes, termino = self._mensajes_filtrados()
        if not mensajes:
            tk.Label(marco, text=("Sin resultados." if termino
                                  else "No hay correos aquí."),
                     bg=C["fondo"], fg=C["texto_tenue"], font=letra(8)
                     ).pack(fill="x", pady=8)
            return

        no_leidos_aqui = [m for m in mensajes if not m.leido]
        if no_leidos_aqui:
            marcar = tk.Label(
                marco, text=f"Marcar como leídos los {len(no_leidos_aqui)} sin leer",
                bg=C["fondo"], fg=C["texto_tenue"], font=letra(7, "underline"),
                cursor="hand2")
            marcar.pack(fill="x", pady=(0, 4), anchor="w")
            marcar.bind("<Button-1>", lambda e: self._marcar_todos_leidos())

        lista = tk.Frame(marco, bg=C["fondo"])
        lista.pack(fill="x")
        hoy = dt.date.today()
        for m in mensajes:
            fila = tk.Frame(lista, bg=C["fondo"], cursor="hand2")
            fila.pack(fill="x", pady=1)
            punto = "●" if not m.leido else " "
            tk.Label(fila, text=punto, bg=C["fondo"],
                     fg=m.color_carpeta or C["evento"], font=letra(8), width=2
                     ).pack(side="left")
            cabecera = tk.Frame(fila, bg=C["fondo"])
            cabecera.pack(side="left", fill="x", expand=True)
            estilo = "bold" if not m.leido else None
            tk.Label(cabecera, text=m.remitente[:30], bg=C["fondo"],
                     fg=C["texto"], font=letra(8, estilo), anchor="w"
                     ).pack(fill="x", anchor="w")
            asunto = m.asunto[:40] + (" 📎" if m.adjuntos else "")
            tk.Label(cabecera, text=asunto, bg=C["fondo"],
                     fg=C["texto"] if not m.leido else C["texto_tenue"],
                     font=letra(8, estilo), anchor="w").pack(fill="x", anchor="w")
            fecha_txt = (f"{m.fecha:%H:%M}" if m.fecha.date() == hoy
                        else f"{m.fecha:%d/%m}")
            tk.Label(fila, text=fecha_txt, bg=C["fondo"], fg=C["texto_tenue"],
                     font=letra(7)).pack(side="right", padx=(4, 2))
            for w in (fila, cabecera, *cabecera.winfo_children()):
                w.bind("<Button-1>", lambda e, msg=m: self._abrir_mensaje(msg))
                w.bind("<Button-3>", lambda e, msg=m: self._menu_mensaje(e, msg))
            tk.Frame(lista, bg=C["rejilla"], height=1).pack(fill="x", pady=1)

    def _al_escribir_busqueda(self):
        if self.modo != "bandeja":
            return
        try:
            self._pintar_lista_mensajes()
            self.actualizar_posicion()
        except Exception:
            pass

    def _marcar_todos_leidos(self):
        mensajes = self.app.correo_mensajes.get(self.carpeta_activa, [])
        cuantos = 0
        for m in mensajes:
            if not m.leido and marcar_correo(m.id_outlook, m.almacen_outlook,
                                             leido=True):
                m.leido = True
                cuantos += 1
        if cuantos:
            for c in self.app.correo_carpetas:
                if c["nombre"] == self.carpeta_activa:
                    c["no_leidos"] = max(0, c["no_leidos"] - cuantos)
        self._pintar_lista_mensajes()
        self.actualizar_posicion()
        self.app.dibujar()

    def _cambiar_carpeta(self, nombre):
        self.carpeta_activa = nombre
        self.busqueda_var.set("")
        self.pintar()
        self.actualizar_posicion()

    def _abrir_mensaje(self, m):
        VentanaCorreo(self.app, m, self)

    def _menu_mensaje(self, e, m):
        import correo_columna
        correo_columna.menu_mensaje(self.app, sys.modules[__name__], e, m)

    def _responder(self, m, modo):
        responder_o_reenviar_correo(m.id_outlook, m.almacen_outlook, modo)

    def _alternar_leido(self, m):
        if not marcar_correo(m.id_outlook, m.almacen_outlook, leido=not m.leido):
            return
        m.leido = not m.leido
        for c in self.app.correo_carpetas:
            if c["nombre"] == m.carpeta:
                c["no_leidos"] = max(0, c["no_leidos"] + (1 if not m.leido else -1))
        self.pintar()
        self.actualizar_posicion()
        self.app.dibujar()

    # -- modo reglas ------------------------------------------------------

    def _cargar_reglas(self):
        def hacer():
            reglas = listar_reglas_outlook()
            self.after(0, lambda: self._reglas_cargadas(reglas))
        self.reglas = None    # None = cargando
        if self.modo == "reglas":
            self.pintar()
        threading.Thread(target=hacer, daemon=True).start()

    def _reglas_cargadas(self, reglas):
        self.reglas = reglas
        try:
            if self.winfo_exists() and self.modo == "reglas":
                self.pintar()
                self.actualizar_posicion()
        except Exception:
            pass

    def _pintar_reglas(self):
        if self.reglas is None:
            tk.Label(self.contenido, text="Cargando reglas…", bg=C["fondo"],
                     fg=C["texto_tenue"], font=letra(8)).pack(pady=8)
            return
        if not self.reglas:
            tk.Label(self.contenido, text="No hay reglas en Outlook.",
                     bg=C["fondo"], fg=C["texto_tenue"], font=letra(8)
                     ).pack(fill="x", pady=8)
            return
        cuenta_actual = None
        for r in sorted(self.reglas, key=lambda r: (r["cuenta"], r["orden"])):
            if r["cuenta"] != cuenta_actual:
                cuenta_actual = r["cuenta"]
                tk.Label(self.contenido, text=cuenta_actual, bg=C["fondo"],
                         fg=C["texto_tenue"], font=letra(7, "bold"), anchor="w"
                         ).pack(fill="x", pady=(6, 2))
            fila = tk.Frame(self.contenido, bg=C["fondo"])
            fila.pack(fill="x", pady=1)
            var = tk.BooleanVar(value=r["activa"])
            chk = tk.Checkbutton(
                fila, variable=var, bg=C["fondo"], activebackground=C["fondo"],
                selectcolor=C.get("boton", "#2f2f33"), bd=0,
                command=lambda r=r, v=var: self._alternar_regla(r, v.get()))
            chk.pack(side="left")
            info = tk.Frame(fila, bg=C["fondo"])
            info.pack(side="left", fill="x", expand=True)
            tk.Label(info, text=r["nombre"], bg=C["fondo"],
                     fg=C["texto"] if r["activa"] else C["texto_tenue"],
                     font=letra(8, "bold"), anchor="w").pack(fill="x", anchor="w")
            tk.Label(info, text=f"Si {r['condiciones']}", bg=C["fondo"],
                     fg=C["texto_tenue"], font=letra(7), anchor="w",
                     wraplength=self.ANCHO - 60, justify="left"
                     ).pack(fill="x", anchor="w")
            tk.Label(info, text=f"→ {r['acciones']}", bg=C["fondo"],
                     fg=C["texto_tenue"], font=letra(7), anchor="w",
                     wraplength=self.ANCHO - 60, justify="left"
                     ).pack(fill="x", anchor="w")
            menu_btn = tk.Label(fila, text="⋮", bg=C["fondo"],
                                fg=C["texto_tenue"], font=letra(9, "bold"),
                                cursor="hand2", padx=4)
            menu_btn.pack(side="right")
            menu_btn.bind("<Button-1>", lambda e, r=r: self._menu_regla(e, r))
            tk.Frame(self.contenido, bg=C["rejilla"], height=1).pack(fill="x", pady=2)

    def _alternar_regla(self, r, activa):
        if alternar_regla(r["cuenta"], r["nombre"], activa):
            r["activa"] = activa
            self.pintar()
            self.actualizar_posicion()
        else:
            self._cargar_reglas()

    def _menu_regla(self, e, r):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Ejecutar ahora",
                          command=lambda: self._ejecutar_regla(r))
        menu.add_separator()
        menu.add_command(label=f"Eliminar «{r['nombre']}»",
                          command=lambda: self._borrar_regla(r))
        menu.tk_popup(e.x_root, e.y_root)

    def _ejecutar_regla(self, r):
        if ejecutar_regla_ahora(r["cuenta"], r["nombre"]):
            self.app._correo_ahora()

    def _borrar_regla(self, r):
        from tkinter import messagebox
        if not messagebox.askyesno(
                "Eliminar regla",
                f"¿Eliminar la regla «{r['nombre']}» de Outlook? "
                "No se puede deshacer.", parent=self):
            return
        if borrar_regla(r["cuenta"], r["nombre"]):
            self._cargar_reglas()

    # -- colocacion -------------------------------------------------------

    def actualizar_posicion(self):
        if ubicacion_panel("correo") == "flotante" and self._colocado_ya:
            return
        try:
            self.update_idletasks()
            app = self.app
            x = x_panel_acoplado(app, "correo", self.ANCHO)
            y = app.winfo_rooty()
            alto = min(self.winfo_reqheight(),
                      round(app.winfo_screenheight() * 0.85))
            sitio = encajar_en_pantalla(app, x, y, self.ANCHO, alto)
            self.geometry(f"{self.ANCHO}x{alto}{sitio}")
            self._colocado_ya = True
        except Exception as exc:
            log(f"Correo: no se pudo colocar el panel ({exc})")

    def _agarrar(self, e):
        self._agarre = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    def _mover_panel(self, e):
        if getattr(self, "_agarre", None):
            self.geometry(f"+{e.x_root - self._agarre[0]}"
                          f"+{e.y_root - self._agarre[1]}")


class VentanaCorreo(tk.Toplevel):
    """El cuerpo de un correo, para leerlo sin salir del widget.

    Al terminar de cargar el cuerpo se marca como leído (si así está
    configurado) y se avisa al panel de la lista para que se actualice.
    """

    ANCHO, ALTO = 480, 460

    def __init__(self, app, mensaje, panel_lista=None):
        super().__init__(app)
        self.app = app
        self.mensaje = mensaje
        self.panel_lista = panel_lista
        self.overrideredirect(True)
        self.configure(bg=C["rejilla"])
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass
        marco = tk.Frame(self, bg=C["fondo"], padx=10, pady=8)
        marco.pack(fill="both", expand=True, padx=1, pady=1)

        cab = tk.Frame(marco, bg=C["fondo"], cursor="fleur")
        cab.pack(fill="x", pady=(0, 6))
        tk.Label(cab, text=mensaje.asunto or "(sin asunto)", bg=C["fondo"],
                 fg=C["texto"], font=letra(10, "bold"), anchor="w",
                 wraplength=self.ANCHO - 40, justify="left", cursor="fleur"
                 ).pack(fill="x", anchor="w")
        cerrar = tk.Label(cab, text="✕", bg=C["fondo"], fg=C["texto_tenue"],
                          font=letra(9), cursor="hand2")
        cerrar.place(relx=1.0, x=0, y=0, anchor="ne")
        cerrar.bind("<Button-1>", lambda e: self.destroy())
        cab.bind("<Button-1>", self._agarrar)
        cab.bind("<B1-Motion>", self._mover_panel)

        remite = mensaje.remitente
        if mensaje.remitente_email and mensaje.remitente_email != remite:
            remite += f"  <{mensaje.remitente_email}>"
        tk.Label(marco, text=remite, bg=C["fondo"], fg=C["texto_tenue"],
                 font=letra(8), anchor="w", wraplength=self.ANCHO - 40,
                 justify="left").pack(fill="x", anchor="w")
        tk.Label(marco, text=f"{mensaje.fecha:%A %d de %B, %H:%M}",
                 bg=C["fondo"], fg=C["texto_tenue"], font=letra(8), anchor="w"
                 ).pack(fill="x", anchor="w", pady=(0, 6))

        cuerpo_marco = tk.Frame(marco, bg=C["fondo"])
        cuerpo_marco.pack(fill="both", expand=True)
        barra = tk.Scrollbar(cuerpo_marco)
        barra.pack(side="right", fill="y")
        self.texto = tk.Text(cuerpo_marco, bg=C["fondo"], fg=C["texto"],
                             font=letra(9), wrap="word", relief="flat",
                             insertbackground=C["texto"],
                             yscrollcommand=barra.set)
        self.texto.pack(side="left", fill="both", expand=True)
        barra.config(command=self.texto.yview)
        self.texto.insert("1.0", "Cargando…")
        self.texto.config(state="disabled")

        pie = tk.Frame(marco, bg=C["fondo"], pady=(8, 0))
        pie.pack(fill="x")
        self._boton(pie, "Responder",
                   lambda: self._responder("responder"), principal=True)
        self._boton(pie, "Reenviar", lambda: self._responder("reenviar"))
        self._boton(pie, "No leído", self._marcar_no_leido)
        self._boton(pie, "Abrir en Outlook", self._abrir_en_outlook)

        self.bind("<Escape>", lambda e: self.destroy())
        self.update_idletasks()
        self.geometry(
            f"{self.ANCHO}x{self.ALTO}"
            + encajar_en_pantalla(app, app.winfo_rootx() + 60,
                                 app.winfo_rooty() + BARRA + 8,
                                 self.ANCHO, self.ALTO))
        threading.Thread(target=self._cargar_cuerpo, daemon=True).start()

    def _boton(self, padre, texto, accion, principal=False):
        b = tk.Label(padre, text=texto,
                    bg=C["evento"] if principal else C.get("boton", "#2f2f33"),
                    fg=C["evento_texto"] if principal else C["texto"],
                    font=letra(8, "bold" if principal else None),
                    padx=10, pady=4, cursor="hand2")
        b.pack(side="left", padx=(0, 6))
        b.bind("<Button-1>", lambda e: accion())
        return b

    def _responder(self, modo):
        responder_o_reenviar_correo(self.mensaje.id_outlook,
                                    self.mensaje.almacen_outlook, modo)

    def _cargar_cuerpo(self):
        cuerpo = cuerpo_y_marcar_correo(self.mensaje.id_outlook,
                                        self.mensaje.almacen_outlook)
        self.after(0, lambda: self._mostrar_cuerpo(cuerpo))

    def _mostrar_cuerpo(self, cuerpo):
        try:
            self.texto.config(state="normal")
            self.texto.delete("1.0", "end")
            self.texto.insert("1.0", cuerpo or "(sin contenido)")
            self.texto.config(state="disabled")
        except Exception:
            return
        if self.mensaje.leido:
            return
        self.mensaje.leido = True
        for c in self.app.correo_carpetas:
            if c["nombre"] == self.mensaje.carpeta:
                c["no_leidos"] = max(0, c["no_leidos"] - 1)
        self.app.dibujar()
        self._avisar_panel()

    def _abrir_en_outlook(self):
        abrir_correo_en_outlook(self.mensaje.id_outlook,
                                self.mensaje.almacen_outlook)

    def _marcar_no_leido(self):
        if not marcar_correo(self.mensaje.id_outlook, self.mensaje.almacen_outlook,
                             leido=False):
            return
        self.mensaje.leido = False
        for c in self.app.correo_carpetas:
            if c["nombre"] == self.mensaje.carpeta:
                c["no_leidos"] += 1
        self.app.dibujar()
        self._avisar_panel()
        self.destroy()

    def _avisar_panel(self):
        panel = self.panel_lista
        try:
            if panel is not None and panel.winfo_exists():
                panel.pintar()
                panel.actualizar_posicion()
        except Exception:
            pass

    def _agarrar(self, e):
        self._agarre = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    def _mover_panel(self, e):
        if getattr(self, "_agarre", None):
            self.geometry(f"+{e.x_root - self._agarre[0]}"
                          f"+{e.y_root - self._agarre[1]}")


class PanelIr(tk.Toplevel):
    """Buscador de eventos y salto a una fecha, con mini calendario del mes."""

    ANCHO = 268
    CELDA = 34
    ALTO_MES = 6 * 26 + 44

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.hoy = dt.datetime.now(TZ).date()
        self.mes = self.hoy.replace(day=1)
        self.resultados = []

        self.overrideredirect(True)
        self.configure(bg=C["rejilla"])
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass

        marco = tk.Frame(self, bg=C["fondo"], padx=8, pady=8)
        marco.pack(fill="both", expand=True, padx=1, pady=1)

        self.entrada = tk.Entry(marco, bg=C.get("boton", "#2f2f33"),
                                fg=C["texto"], insertbackground=C["texto"],
                                relief="flat", font=letra(9))
        self.entrada.pack(fill="x", ipady=4)
        self.entrada.bind("<KeyRelease>", self._al_escribir)
        self.entrada.bind("<Return>", self._aceptar)
        self.entrada.bind("<Escape>", lambda e: self.destroy())

        self.pista = tk.Label(marco, bg=C["fondo"], fg=C["texto_tenue"],
                              font=letra(7), anchor="w",
                              text="Busca un evento, o escribe una fecha "
                                   "(18/9, 18 sep, hoy)")
        self.pista.pack(fill="x", pady=(4, 6))

        self.lienzo = tk.Canvas(marco, bg=C["fondo"], highlightthickness=0,
                                width=self.ANCHO, height=self.ALTO_MES)
        self.lienzo.pack()
        self.lienzo.bind("<Button-1>", self._clic)

        self.lista = tk.Listbox(marco, bg=C["fondo"], fg=C["texto"],
                                selectbackground=C["evento"], relief="flat",
                                highlightthickness=0, font=letra(8), height=8,
                                activestyle="none")
        self.lista.bind("<Double-Button-1>", self._aceptar)
        self.lista.bind("<Return>", self._aceptar)

        self._pintar_mes()
        self.update_idletasks()
        x = app.winfo_rootx() + 40
        y = app.winfo_rooty() + BARRA + 6
        self.geometry(f"+{x}+{y}")
        self.entrada.focus_force()
        self.bind("<FocusOut>", self._quizas_cerrar)

    def _quizas_cerrar(self, _=None):
        # se cierra al perder el foco, pero no al pulsar dentro del panel
        self.after(150, lambda: (self.focus_displayof() or self.destroy()))

    # -- mini calendario ----------------------------------------------------

    def _pintar_mes(self):
        c = self.lienzo
        c.delete("all")
        prim = self.mes
        c.create_text(20, 12, text="‹", fill=C["texto"], font=letra(11),
                      tags="mes_ant")
        c.create_text(self.ANCHO - 20, 12, text="›", fill=C["texto"],
                      font=letra(11), tags="mes_sig")
        c.create_text(self.ANCHO / 2, 12,
                      text=f"{MESES_LARGOS[prim.month - 1].capitalize()} "
                           f"{prim.year}",
                      fill=C["texto"], font=letra(9, "bold"))

        ancho_col = self.ANCHO / 7
        for i, ini in enumerate("LMXJVSD"):
            c.create_text(ancho_col * (i + 0.5), 32, text=ini,
                          fill=C["texto_tenue"], font=letra(7))

        con_eventos = {ev.inicio.date() for ev in self.app._mostrables()}
        con_eventos |= {ev.inicio.date() for ev in self.app.eventos}
        dia = prim - dt.timedelta(days=prim.weekday())
        self._celdas = {}
        for fila in range(6):
            for col in range(7):
                x = ancho_col * (col + 0.5)
                y = 48 + fila * 26
                de_otro_mes = dia.month != prim.month
                if dia == self.hoy:
                    c.create_oval(x - 11, y - 10, x + 11, y + 10,
                                  fill=C["evento"], outline="")
                color = (C["evento_texto"] if dia == self.hoy
                         else C["texto_tenue"] if de_otro_mes else C["texto"])
                c.create_text(x, y, text=str(dia.day), fill=color, font=letra(8))
                if dia in con_eventos and dia != self.hoy:
                    c.create_oval(x - 2, y + 9, x + 2, y + 13,
                                  fill=C["evento"], outline="")
                self._celdas[(fila, col)] = dia
                dia += dt.timedelta(days=1)

    def _clic(self, e):
        item = self.lienzo.find_withtag("current")
        etiquetas = self.lienzo.gettags(item[0]) if item else ()
        if "mes_ant" in etiquetas or "mes_sig" in etiquetas:
            paso = -1 if "mes_ant" in etiquetas else 1
            mes = self.mes.month + paso
            anio = self.mes.year + (1 if mes > 12 else -1 if mes < 1 else 0)
            self.mes = dt.date(anio, (mes - 1) % 12 + 1, 1)
            self._pintar_mes()
            return
        if e.y < 40:
            return
        fila, col = (e.y - 36) // 26, int(e.x // (self.ANCHO / 7))
        fecha = self._celdas.get((fila, col))
        if fecha:
            self.app.ir_a_fecha(fecha)
            self.destroy()

    # -- busqueda -----------------------------------------------------------

    def _al_escribir(self, _=None):
        texto = self.entrada.get()
        fecha = interpretar_fecha(texto, self.hoy)
        if fecha:
            self.pista.config(text=f"Intro para ir al {fecha:%d/%m/%Y}")
            self.mes = fecha.replace(day=1)
            self._pintar_mes()
            self.lista.pack_forget()
            self.resultados = []
            return

        self.resultados = buscar_eventos(self.app._mostrables(), texto)
        self.lista.delete(0, "end")
        if not texto.strip():
            self.pista.config(text="Busca un evento, o escribe una fecha "
                                   "(18/9, 18 sep, hoy)")
            self.lista.pack_forget()
            return
        if not self.resultados:
            self.pista.config(text="Sin coincidencias en los próximos meses")
            self.lista.pack_forget()
            return
        self.pista.config(text=f"{len(self.resultados)} coincidencias · "
                               f"Intro o doble clic para ir")
        for ev in self.resultados:
            hora = "todo el día" if ev.todo_dia else f"{ev.inicio:%H:%M}"
            self.lista.insert("end",
                              f" {fecha_corta(ev.inicio)}  {hora}  {ev.titulo}")
        self.lista.pack(fill="x", pady=(6, 0))
        self.lista.selection_set(0)

    def _aceptar(self, _=None):
        fecha = interpretar_fecha(self.entrada.get(), self.hoy)
        if not fecha and self.resultados:
            sel = self.lista.curselection()
            fecha = self.resultados[sel[0] if sel else 0].inicio.date()
        if fecha:
            self.app.ir_a_fecha(fecha)
            self.destroy()
        return "break"


class Widget(tk.Tk):
    def __init__(self):
        super().__init__()
        self.eventos = leer_cache()
        self._firma_postit = None        # firma de las notas de Post-It ya unidas a self.eventos
        self._carpeta_sync = None
        self.avisados = set()
        self.desplazamiento = 0          # dias de desfase respecto a la ventana de hoy
        self.actualizado = None          # hora de la ultima descarga correcta
        self.error_red = ""              # ultimo fallo al descargar
        self._id_refresco = None
        self.correo_carpetas = []        # [{"nombre","color","total","no_leidos"}]
        self.correo_mensajes = {}        # nombre de carpeta -> [Correo, ...]
        self.correo_error = ""
        self._id_refresco_correo = None
        self._panel_correo = None
        self._drag = None
        self._redim = None
        self._divisor = None
        self._divisor_t = None        # arrastre del divisor de la columna to-dos
        self._divisor_c = None        # arrastre del divisor de la columna de correo
        self._moviendo = None         # evento que se esta arrastrando
        self._pintando = None         # rato que se esta dibujando con el raton
        self._off_rejilla = (0, 0)    # a donde se ha movido la rejilla
        self._vistas_aparte = []      # cada vista abierta en su propia ventana
        self._panel_control = None    # Panel de control (extras.py)
        self._paleta = None           # Paleta de comandos (extras.py)
        self._destino = None          # ventana a la que se esta pintando
        self._lienzo_activo = None    # lienzo que manda al buscar eventos
        self._hueco = None            # hueco de la rejilla al repartir los dias
        self.hwnd = None
        self._hwnd_escritorio = None
        self._alto_banda = BANDA_MINIMA
        self._altos_linea = {}
        self._fuentes = {}
        self._anchos = {}             # (fuente, texto) -> ancho en pixeles
        # a partir de aqui el guardado agrupado ya tiene ventana donde programarse
        _PENDIENTE["raiz"] = self
        self._tam = None
        self._pos = (CFG["posicion_x"], CFG["posicion_y"])
        self.vista = CFG.get("vista", "rejilla")
        self.orientacion = CFG.get("orientacion", "horizontal")
        self.visible = True
        self.bandeja = None
        self.nivel = CFG.get("nivel_ventana", "fondo")
        self.filtro = CFG.get("filtro")
        self._btn_activo = {}
        self._ayuda_win = None
        self._ayuda_id = None
        self._ayuda_tag = None
        self._claro = tema_claro_activo()
        self._festivos = []
        self._vars_festivos = {}
        self._info_eventos = {}     # tag del lienzo -> evento (para el detalle)
        self._n_evento = 0
        self._resumen_hecho = False

        self.overrideredirect(True)
        self.configure(bg=C["fondo"])
        self.attributes("-alpha", CFG["opacidad"])

        self.lienzo = tk.Canvas(self, bg=C["fondo"], highlightthickness=0)
        self.lienzo.pack(fill="both", expand=True)

        self.lienzo.bind("<Button-1>", self._inicio_arrastre)
        self.lienzo.bind("<B1-Motion>", self._arrastrar)
        self.lienzo.bind("<ButtonRelease-1>", self._fin_arrastre)
        self.lienzo.bind("<Motion>", self._cursor_borde)
        self.lienzo.bind("<Button-3>", self._menu)
        self.lienzo.bind("<Double-Button-1>", self._doble_clic)
        # paleta de comandos: vale desde cualquier ventana del widget
        self.bind_all("<Control-k>", self._tecla_paleta)
        self.bind_all("<Control-K>", self._tecla_paleta)
        # evento nuevo desde el teclado (Ctrl+N), en el siguiente hueco libre
        self.bind_all("<Control-n>", lambda e: self._nuevo_evento())
        self.bind_all("<Control-N>", lambda e: self._nuevo_evento())
        # rueda: un dia (o el horario si esta sobre la columna de horas)
        self.lienzo.bind("<MouseWheel>",
                         lambda e: self._rueda(e, e.delta > 0))
        self.lienzo.bind("<Shift-MouseWheel>",
                         lambda e: self._mover(-CFG["dias_visibles"] if e.delta > 0
                                               else CFG["dias_visibles"]))
        self.lienzo.bind("<Control-MouseWheel>",
                         lambda e: self._cambiar_escala(0.1 if e.delta > 0 else -0.1))
        self.lienzo.bind("<Control-Shift-MouseWheel>",
                         lambda e: self._cambiar_letra(0.1 if e.delta > 0 else -0.1))
        self.lienzo.bind("<Button-4>", lambda e: self._rueda(e, True))   # X11
        self.lienzo.bind("<Button-5>", lambda e: self._rueda(e, False))
        # detalle del evento al pasar el raton
        self.lienzo.tag_bind("divisor", "<Button-1>", self._inicio_divisor)
        self.lienzo.tag_bind("divisor", "<Double-Button-1>",
                             self._reparto_por_defecto)
        self.lienzo.tag_bind("divisor", "<Button-3>", self._menu_divisor)
        # el cursor depende de hacia donde se mueve la separacion
        self.lienzo.tag_bind("divisor_v", "<Enter>",
                             lambda e: self._cursor_divisor(e, vertical=True))
        self.lienzo.tag_bind("divisor_h", "<Enter>",
                             lambda e: self._cursor_divisor(e, vertical=False))
        self.lienzo.tag_bind("divisor", "<Leave>",
                             lambda e: (self.lienzo.config(cursor=""),
                                        self._resaltar_divisor(False)))
        # divisor propio de la columna de to-dos (cambia solo su ancho)
        self.lienzo.tag_bind("divisor_todos", "<Button-1>",
                             self._inicio_divisor_todos)
        self.lienzo.tag_bind("divisor_todos", "<Double-Button-1>",
                             lambda e: self._reset_ancho_todos())
        self.lienzo.tag_bind("divisor_todos", "<Button-3>",
                             self._menu_ancho_todos)
        self.lienzo.tag_bind("divisor_todos", "<Enter>",
                             lambda e: self._cursor_columna("divisor_todos", True))
        self.lienzo.tag_bind("divisor_todos", "<Leave>",
                             lambda e: self._cursor_columna("divisor_todos", False))
        # divisor propio de la columna de correo
        self.lienzo.tag_bind("divisor_correo", "<Button-1>",
                             self._inicio_divisor_correo)
        self.lienzo.tag_bind("divisor_correo", "<Double-Button-1>",
                             lambda e: self._reset_ancho_correo())
        self.lienzo.tag_bind("divisor_correo", "<Button-3>",
                             self._menu_ancho_correo)
        self.lienzo.tag_bind("divisor_correo", "<Enter>",
                             lambda e: self._cursor_columna("divisor_correo", True))
        self.lienzo.tag_bind("divisor_correo", "<Leave>",
                             lambda e: self._cursor_columna("divisor_correo", False))
        self.lienzo.tag_bind("evento", "<Button-1>", self._inicio_mover)
        self.lienzo.tag_bind("evento", "<Enter>", self._mostrar_detalle)
        self.lienzo.tag_bind("evento", "<Button-3>", self._menu_evento)
        self.lienzo.tag_bind("evento", "<Leave>", lambda e: self._quitar_detalle())
        # botones de la barra superior
        acciones = {
            "nav_ant": lambda e: self._mover(-CFG["dias_visibles"]),
            "nav_sig": lambda e: self._mover(CFG["dias_visibles"]),
            "nav_hoy": lambda e: self._ir_a_hoy(),
            "btn_dias": self._menu_dias,
            "btn_festivos": self._menu_festivos,
            "btn_horario": self._menu_horario,
            "btn_ir": self._abrir_panel_ir,
            "btn_nuevo": self._nuevo_evento,
            "btn_nivel": lambda e: self._siguiente_nivel(),
            "btn_horizontal": lambda e: self._cambiar_orientacion("horizontal"),
            "btn_vertical": lambda e: self._cambiar_orientacion("vertical"),
            "btn_vista": self._abrir_accesos,
            "btn_panel": self._abrir_panel_control,
            "btn_paleta": self._abrir_paleta,
            "btn_avisos": self._alternar_notificaciones,
            "btn_todos": self._abrir_todos,
            "btn_correo_col": self._abrir_correo_columna,
            "btn_cuenta": self._menu_cuentas,
            "btn_perfil": lambda e: self._abrir_menu(
                self._menu_perfiles(self.lienzo), e),
            "btn_cerrar": lambda e: self._salir(),
            "btn_mas": lambda e: self._menu_desbordados(e),
            "btn_refrescar": lambda e: self.refrescar(),
            "btn_ocultar": lambda e: self._alternar_visible(False),
        }
        self._acciones_barra = acciones
        self._tags_interactivos = set(acciones) | {"divisor", "divisor_correo",
                                                   "correo_col", "chip_cuenta"}
        self.lienzo.tag_bind("chip_cuenta", "<Button-1>", self._menu_cuentas)
        for tag, accion in acciones.items():
            self.lienzo.tag_bind(tag, "<Button-1>", lambda e, f=accion: f(e))
            self.lienzo.tag_bind(tag, "<Enter>",
                                 lambda e, t=tag: self._resaltar(t, True))
            self.lienzo.tag_bind(tag, "<Leave>",
                                 lambda e, t=tag: self._resaltar(t, False))
            if tag in self.NOMBRES_BOTONES:
                self._vincular_ayuda(tag, self.NOMBRES_BOTONES[tag])
        # el botón de vista: clic izquierdo pasa a la siguiente, clic derecho
        # abre el menú con todas (y con el tablero de plantillas)
        self.lienzo.tag_bind("btn_vista", "<Button-3>", self._menu_vistas)
        self.lienzo.tag_bind("btn_avisos", "<Button-3>", self._menu_campana)

        self.after(500, lambda: self._cambiar_nivel(self.nivel, guardar=False))
        if CFG.get("clics_atraviesan"):
            self.after(700, lambda: self._bloquear_clics(True))
        self.after(400, self._configurar_ventana)
        self.after(600, self._al_fondo)
        self.after(800, self._arrancar_bandeja)
        self.after(900, self._arrancar_atajo)
        self.after(25000, self._comprobar_actualizacion_periodica)
        self._vigilar_senal()
        self._vigilar_tema()
        self._en_otro_escritorio = False
        self._vigilar_escritorio()
        self._firma = self.firma_pantallas()
        self.after(700, self._restaurar_posicion)
        self.after(20_000, self._vigilar_pantallas)
        self.refrescar()
        self.after(2000, self._sincronizar)   # notas de Post-It, cada 15 s
        self.after(1200, self.refrescar_correo)
        self.tic()
        self.dibujar()
        self.after(1500, lambda: log(
            f"Ventana en {self.geometry()}  nivel={self.nivel}  "
            f"visible={self.visible}  pantalla="
            f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}"))

    # -- mostrar / ocultar ---------------------------------------------------

    def _vigilar_senal(self):
        """Otra ejecucion del programa (acceso directo) pide mostrar u ocultar."""
        if consumir_alternar():
            self._alternar_visible()
        self.after(700, self._vigilar_senal)

    def _alternar_visible(self, mostrar=None):
        self.visible = (not self.visible) if mostrar is None else mostrar
        if self.visible:
            self.deiconify()
            self.after(120, self._configurar_ventana)   # recupera el estilo de widget
            self.after(240, self.dibujar)
        else:
            self.withdraw()
        log(f"Widget {'mostrado' if self.visible else 'oculto'} "
            f"({self.geometry()})")

    def _nuevo_rapido(self):
        """Ctrl+Alt+N desde cualquier sitio: muestra el widget y abre la ficha."""
        if not self.visible:
            self._alternar_visible(True)
        self._nuevo_evento()

    def report_callback_exception(self, tipo, valor, tb):
        # sin consola (pythonw) Tk se tragaba estos errores; ahora quedan en el registro
        import traceback
        log("ERROR en un evento de la ventana: "
            + " | ".join("".join(traceback.format_exception(tipo, valor, tb))
                         .splitlines()))

    def _salir(self):
        volcar_config()          # lo que quedara pendiente, al disco ya
        if self.bandeja:
            try:
                self.bandeja.stop()
            except Exception:
                pass
        self.destroy()

    def destroy(self):
        # cerrar por la X, por Alt+F4 o desde el codigo pasa siempre por aqui
        try:
            volcar_config()
        except Exception:
            pass
        super().destroy()

    def _arrancar_atajo(self):
        """Registra el atajo global de Windows (ctrl+alt+c por defecto)."""
        texto = CFG.get("atajo_global")
        if not texto:
            return
        threading.Thread(target=self._hilo_atajo, args=(str(texto),),
                         daemon=True).start()

    def _hilo_atajo(self, texto):
        import ctypes
        from ctypes import wintypes
        MODS = {"ctrl": 0x0002, "control": 0x0002, "alt": 0x0001,
                "shift": 0x0004, "win": 0x0008}
        ESPECIALES = {"space": 0x20, "espacio": 0x20, "f1": 0x70, "f2": 0x71,
                      "f3": 0x72, "f4": 0x73, "f5": 0x74, "f6": 0x75,
                      "f7": 0x76, "f8": 0x77, "f9": 0x78, "f10": 0x79,
                      "f11": 0x7A, "f12": 0x7B}
        partes = [p.strip().lower() for p in texto.split("+") if p.strip()]
        if not partes:
            return
        tecla, mods = partes[-1], 0
        for p in partes[:-1]:
            mods |= MODS.get(p, 0)
        vk = ESPECIALES.get(tecla, ord(tecla.upper()) if len(tecla) == 1 else 0)
        if not vk or not mods:
            log(f"Atajo '{texto}' no válido; se ignora")
            return
        try:
            user32 = ctypes.windll.user32
            if not user32.RegisterHotKey(None, 1, mods | 0x4000, vk):  # NOREPEAT
                log(f"No se pudo registrar el atajo '{texto}' (¿lo usa otra app?)")
                return
            log(f"Atajo global activo: {texto}")
            if user32.RegisterHotKey(None, 2, 0x0003 | 0x4000, ord("N")):
                log("Atajo global activo: ctrl+alt+n (evento nuevo)")
            if user32.RegisterHotKey(None, 3, 0x0003 | 0x4000, ord("B")):
                log("Atajo global activo: ctrl+alt+b (clics atraviesan)")
            acciones = {2: self._nuevo_rapido, 3: self._bloquear_clics}
            msg = wintypes.MSG()
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                if msg.message == 0x0312:            # WM_HOTKEY
                    self.after(0, acciones.get(msg.wParam, self._alternar_visible))
        except Exception as exc:
            log(f"Atajo global no disponible: {exc}")

    def _arrancar_bandeja(self):
        """Icono junto al reloj: clic izquierdo muestra/oculta, derecho abre menu."""
        try:
            import pystray
            from PIL import Image
        except Exception as exc:
            log(f"Sin icono de bandeja (falta pystray o Pillow): {exc}")
            return

        try:
            imagen = (Image.open(ICONO_PATH) if os.path.exists(ICONO_PATH)
                      else Image.new("RGB", (64, 64), C["evento"]))
        except Exception:
            imagen = Image.new("RGB", (64, 64), C["evento"])

        def en_hilo_tk(funcion):
            return lambda icono=None, item=None: self.after(0, funcion)

        menu = pystray.Menu(
            pystray.MenuItem("Mostrar / ocultar",
                             en_hilo_tk(self._alternar_visible), default=True),
            pystray.MenuItem("No la encuentro: traerla al centro",
                             en_hilo_tk(self._rescatar)),
            pystray.MenuItem("Nuevo evento en Outlook…",
                             en_hilo_tk(self._nuevo_evento)),
            pystray.MenuItem("Buscar o ir a una fecha…",
                             en_hilo_tk(self._abrir_panel_ir)),
            pystray.MenuItem("Actualizar ahora", en_hilo_tk(self.refrescar)),
            pystray.MenuItem("Ayuda y atajos…", en_hilo_tk(self._abrir_ayuda_general)),
            pystray.MenuItem("Buscar actualizaciones…", en_hilo_tk(lambda: self._buscar_actualizacion(True))),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Calendario",
                             en_hilo_tk(lambda: self._cambiar_vista("rejilla")),
                             checked=lambda i: self.vista == "rejilla", radio=True),
            pystray.MenuItem("Lista",
                             en_hilo_tk(lambda: self._cambiar_vista("lista")),
                             checked=lambda i: self.vista == "lista", radio=True),
            pystray.MenuItem("Las dos",
                             en_hilo_tk(lambda: self._cambiar_vista("ambas")),
                             checked=lambda i: self.vista == "ambas", radio=True),
            pystray.MenuItem("Mes completo",
                             en_hilo_tk(lambda: self._cambiar_vista("mes")),
                             checked=lambda i: self.vista == "mes", radio=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Próximos 7 días",
                             en_hilo_tk(lambda: self._cambiar_dias(7)),
                             checked=lambda i: CFG["dias_visibles"] == 7, radio=True),
            pystray.MenuItem("Próximos 14 días",
                             en_hilo_tk(lambda: self._cambiar_dias(14)),
                             checked=lambda i: CFG["dias_visibles"] == 14, radio=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Siempre encima de todo",
                             en_hilo_tk(lambda: self._cambiar_nivel("encima")),
                             checked=lambda i: self.nivel == "encima", radio=True),
            pystray.MenuItem("Ventana normal",
                             en_hilo_tk(lambda: self._cambiar_nivel("normal")),
                             checked=lambda i: self.nivel == "normal", radio=True),
            pystray.MenuItem("Al fondo, sobre el escritorio",
                             en_hilo_tk(lambda: self._cambiar_nivel("fondo")),
                             checked=lambda i: self.nivel == "fondo", radio=True),
            pystray.MenuItem("Clics atraviesan (modo fondo)",
                             en_hilo_tk(self._bloquear_clics),
                             checked=lambda i: CFG.get("clics_atraviesan", False)),
            pystray.MenuItem("Centrar en la pantalla", en_hilo_tk(self._centrar)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Salir", en_hilo_tk(self._salir)),
        )
        try:
            self.bandeja = pystray.Icon("calendario", imagen, "Calendario", menu)
            threading.Thread(target=self.bandeja.run, daemon=True).start()
            log("Icono de bandeja activo")
        except Exception as exc:
            log(f"No se pudo crear el icono de bandeja: {exc}")

    # -- geometria -----------------------------------------------------------

    def _ancho_columna(self):
        """Con 14 dias las columnas se estrechan para no ensanchar el widget."""
        forzado = getattr(self, "_ov_col", None)
        if forzado:
            return forzado
        dias = max(CFG["dias_visibles"], 1)
        if CFG.get("ancho_fijo", True):
            base = CFG["ancho_columna"]
            # ancho minimo: se estrecha cuanto mas dias se muestran a la vez
            piso = max(22, min(52, round(1300 / dias)))
            ancho = max(piso, min(base * 3, round(base * 7 / dias)))
        else:
            ancho = CFG["ancho_columna"]
        return round(ancho * ESCALA)

    def _alto_hora(self):
        return round(CFG["alto_hora"] * ESCALA)

    @staticmethod
    def _division():
        """Cada cuántos minutos se marca una raya en la rejilla."""
        try:
            n = int(CFG.get("division_minutos", 60) or 60)
        except (TypeError, ValueError):
            n = 60
        return max(5, min(240, n))

    def _marcas_horarias(self):
        """[(minutos desde el principio, es hora en punto)] de la rejilla."""
        paso = self._division()
        total = (CFG["hora_fin"] - CFG["hora_inicio"]) * 60
        marcas = [(m, m % 60 == 0) for m in range(0, total, paso)]
        marcas.append((total, total % 60 == 0))
        return marcas

    @staticmethod
    def _etiqueta_marca(minutos):
        """'9' en las horas en punto, ':30' en las de en medio."""
        h, m = divmod(CFG["hora_inicio"] * 60 + minutos, 60)
        return f"{h}" if m == 0 else f":{m:02d}"

    def _ancho_hora(self):
        forzado = getattr(self, "_ov_hora", None)
        if forzado:
            return forzado
        return round(CFG["ancho_hora"] * ESCALA)

    def _alto_fila_v(self):
        return round(CFG["alto_fila_v"] * ESCALA)

    def _margen_dias(self):
        base = round(86 * ESCALA)
        if getattr(self, "_hay_especiales_v", False):
            base += self._ancho_especial_v() + 4
        return base

    def _ancho_especial_v(self):
        """Ancho reservado para el recuadro de festivo/cumpleaños en modo
        vertical, pegado al lado del nombre del dia, antes de que empiecen
        las horas. Con este ancho la mayoria de nombres caben en una linea;
        si no caben, se envuelven a una segunda antes de recortarse."""
        return round(108 * ESCALA)

    def _cab_horas(self):
        return BARRA + round(20 * ESCALA)

    def _medidas_vertical(self):
        horas = CFG["hora_fin"] - CFG["hora_inicio"]
        ancho = self._margen_dias() + horas * self._ancho_hora() + 8
        alto = (self._cab_horas() + CFG["dias_visibles"] * self._alto_fila_v()
                + ALTO_ESTADO + 2)
        return ancho, alto

    def _cambiar_orientacion(self, orientacion=None):
        if orientacion is None:
            orientacion = ("vertical" if self.orientacion == "horizontal"
                           else "horizontal")
        self.orientacion = CFG["orientacion"] = orientacion
        self._hueco = None
        guardar_config()
        self._tam = None
        log(f"Orientación: {orientacion}")
        self.dibujar()
        return "break"

    def _ancho_lista(self):
        forzado = getattr(self, "_ov_lista", None)
        if forzado:
            return forzado
        return round(CFG["ancho_lista"] * ESCALA)

    # -- columnas (to-dos, correo) dentro del tamaño actual del widget --------

    def _anchos_columnas(self):
        """(ancho de to-dos, ancho de correo) de lo que esté a la vista."""
        if self.vista == "mes":
            return 0, 0
        return (self._ancho_todos() if CFG.get("todos_visible") else 0,
                self._ancho_correo() if CFG.get("correo_visible") else 0)

    def _quitar_reserva(self):
        self._ov_col = self._ov_hora = self._ov_lista = None
        self._rec = (0, 0)
        self._sobra_col = 0

    def _reservar_para_columnas(self, reserva, con_rejilla, con_lista,
                                vertical, apilada):
        """Estrecha el calendario (y la lista) para que las columnas de
        to-dos y correo quepan dentro del ancho que ya tiene el widget, en
        vez de ensancharlo. Si ya no se puede estrechar más, el widget crece
        solo lo que falte."""
        falta = reserva
        rec_rej = rec_lista = 0
        if con_rejilla:
            if vertical:
                n = max(CFG["hora_fin"] - CFG["hora_inicio"], 1)
                actual, minimo = self._ancho_hora(), round(18 * ESCALA)
            else:
                n = max(CFG["dias_visibles"], 1)
                actual, minimo = self._ancho_columna(), round(30 * ESCALA)
            quita = min(falta, max(0, (actual - minimo) * n))
            if quita > 0:
                nuevo = max(minimo, actual - int(quita // n) - (1 if quita % n else 0))
                if vertical:
                    self._ov_hora = nuevo
                else:
                    self._ov_col = nuevo
                rec_rej = (actual - nuevo) * n
                falta -= rec_rej
        if con_lista and not apilada and falta > 0:
            actual, minimo = self._ancho_lista(), round(170 * ESCALA)
            quita = min(falta, max(0, actual - minimo))
            if quita > 0:
                self._ov_lista = actual - int(quita)
                rec_lista = int(quita)
        self._rec = (rec_rej, rec_lista)
        # al redondear a pixeles se puede quitar de mas: lo absorben las columnas
        self._sobra_col = max(0, rec_rej + rec_lista - reserva)

    def _ancho_rejilla(self):
        return MARGEN_IZQ + CFG["dias_visibles"] * self._ancho_columna() + 8

    def _alto_rejilla(self):
        horas = CFG["hora_fin"] - CFG["hora_inicio"]
        return (CABECERA + self._alto_banda + horas * self._alto_hora()
                + ALTO_MESES + ALTO_ESTADO + 2)

    def _ajustar(self, ancho, alto):
        """Redimensiona solo si cambia (evita parpadeos en cada repintado).

        Si se está pintando una vista en ventana aparte, el tamaño va a esa
        ventana y no a la principal.
        """
        destino = getattr(self, "_destino", None)
        if destino is not None:
            if (ancho, alto) != self._tam:
                self._tam = (ancho, alto)
                destino.geometry(f"{ancho}x{alto}")
            return
        if (ancho, alto) != self._tam:
            self._tam = (ancho, alto)
            self.geometry(f"{ancho}x{alto}+{self._pos[0]}+{self._pos[1]}")
            self._reposicionar_paneles_acoplados()

    def _reposicionar_paneles_acoplados(self):
        """Los paneles anclados a un lado siguen a la ventana principal al
        moverse o cambiar de tamaño (los que están en ventana aparte, no)."""
        ubicaciones = CFG.get("ubicacion_vistas", {})
        for clave, atributo in (("correo", "_panel_correo"),
                                ("habitos", "_panel_habitos"),
                                ("cumpleanos", "_panel_cumpleanos"),
                                ("plantillas", "_panel_plantillas"),
                                ("tareas", "_panel_tareas"),
                                ("control", "_panel_control"),
                                ("paleta", "_paleta")):
            if ubicaciones.get(clave, "flotante") == "flotante":
                continue
            panel = getattr(self, atributo, None)
            if panel is None:
                continue
            try:
                if panel.winfo_exists():
                    panel.actualizar_posicion()
            except Exception:
                pass

    # -- comportamiento de ventana -----------------------------------------

    def _configurar_ventana(self):
        """Fuera de alt-tab y de la barra de tareas."""
        try:
            import win32con
            import win32gui
            hwnd = win32gui.GetParent(self.winfo_id()) or self.winfo_id()
            estilo = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            estilo |= win32con.WS_EX_TOOLWINDOW
            estilo &= ~win32con.WS_EX_APPWINDOW
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, estilo)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWNA)
            self.hwnd = hwnd
        except Exception as exc:
            print("pywin32 no disponible, el widget se comportara como ventana normal:", exc)
            self.hwnd = None

    def _buscar_escritorio(self):
        """Handle del fondo de escritorio (WorkerW si hay presentacion, si no Progman)."""
        import win32gui
        encontrado = []

        def cb(h, _):
            if win32gui.FindWindowEx(h, 0, "SHELLDLL_DefView", None):
                w = win32gui.FindWindowEx(0, h, "WorkerW", None)
                if w:
                    encontrado.append(w)
            return True

        try:
            win32gui.EnumWindows(cb, None)
        except Exception:
            pass
        if encontrado:
            return encontrado[0]
        return win32gui.FindWindow("Progman", None)

    @staticmethod
    def _posicion_z(hwnd, tope=500):
        """Posicion en el orden Z (0 = la ventana mas al frente)."""
        import win32con
        import win32gui
        h = win32gui.GetWindow(win32gui.GetDesktopWindow(), win32con.GW_CHILD)
        i = 0
        while h and i < tope:
            if h == hwnd:
                return i
            h = win32gui.GetWindow(h, win32con.GW_HWNDNEXT)
            i += 1
        return None

    def _cambiar_nivel(self, nivel, guardar=True):
        """fondo / normal / encima, en caliente."""
        self.nivel = nivel
        if guardar:
            CFG["nivel_ventana"] = nivel
            guardar_config()
        try:
            self.attributes("-topmost", nivel == "encima")
        except Exception:
            pass
        if self.hwnd:   # en Windows se fuerza ademas por API, mas fiable
            try:
                import win32con
                import win32gui
                win32gui.SetWindowPos(
                    self.hwnd,
                    win32con.HWND_TOPMOST if nivel == "encima"
                    else win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
                    | win32con.SWP_NOACTIVATE)
            except Exception:
                pass
        if nivel == "encima":
            self.lift()
        log(f"Nivel de ventana: {nivel}")
        if self.visible:
            self.dibujar()

    def _bloquear_clics(self, activar=None):
        """Los clics atraviesan el widget (como un fondo de escritorio).

        Se sale con ctrl+alt+b, desde el icono de la bandeja o con ctrl+alt+c.
        """
        if activar is None:
            activar = not CFG.get("clics_atraviesan", False)
        CFG["clics_atraviesan"] = bool(activar)
        guardar_config()
        if not self.hwnd:
            return
        try:
            import win32con
            import win32gui
            estilo = win32gui.GetWindowLong(self.hwnd, win32con.GWL_EXSTYLE)
            marcas = win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT
            estilo = (estilo | marcas) if activar else (estilo & ~win32con.WS_EX_TRANSPARENT)
            win32gui.SetWindowLong(self.hwnd, win32con.GWL_EXSTYLE, estilo)
            log(f"Clics {'atraviesan' if activar else 'normales'}")
        except Exception as exc:
            log(f"No se pudo cambiar el modo de clics: {exc}")

    def _siguiente_nivel(self):
        orden = ["fondo", "normal", "encima"]
        self._cambiar_nivel(orden[(orden.index(self.nivel) + 1) % len(orden)]
                            if self.nivel in orden else "encima")

    def _alternar_arranque(self):
        """Enciende o apaga el arranque automático con Windows."""
        activar = not arranque_activo()
        hecho, detalle = poner_arranque(activar)
        AvisoFlotante(self,
                      "Arranque con Windows" if hecho else "No se ha podido",
                      detalle, C["evento"] if hecho else C["ahora"], 8)
        return "break"

    def _centrar(self):
        """Por si la ventana se queda en un monitor que ya no existe."""
        ancho, alto = self._tam or (600, 400)
        x = max(0, (self.winfo_screenwidth() - ancho) // 2)
        y = max(0, (self.winfo_screenheight() - alto) // 3)
        self._pos = (x, y)
        CFG["posicion_x"], CFG["posicion_y"] = x, y
        guardar_config()
        self.geometry(f"+{x}+{y}")
        self._alternar_visible(True)
        log(f"Ventana centrada en {x},{y}")

    def _rescatar(self):
        """Cuando no se encuentra la ventana: la trae al centro y por delante."""
        try:
            CFG["opacidad"] = max(CFG.get("opacidad", 0.92), 0.92)
            self.attributes("-alpha", CFG["opacidad"])
        except Exception:
            pass
        self._cambiar_nivel("encima")
        self._alternar_visible(True)
        self._centrar()
        self.lift()
        log(f"RESCATE -> {self.geometry()}  nivel={self.nivel}")

    def _al_fondo(self):
        """Manda la ventana al fondo, pero nunca por detras del escritorio.

        HWND_BOTTOM a secas acaba dejando el widget debajo del fondo de
        pantalla y deja de verse. Por eso, si el escritorio queda por delante,
        se le empuja a el hacia atras.
        """
        if self.nivel != "fondo":
            # con "normal" o "encima" no se toca el orden Z
            self.after(3000, self._al_fondo)
            return
        if self.hwnd:
            try:
                import win32con
                import win32gui
                flags = (win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
                         | win32con.SWP_NOACTIVATE)
                win32gui.SetWindowPos(self.hwnd, win32con.HWND_BOTTOM, 0, 0, 0, 0, flags)

                if self._hwnd_escritorio is None or not win32gui.IsWindow(
                        self._hwnd_escritorio):
                    self._hwnd_escritorio = self._buscar_escritorio()

                desk = self._hwnd_escritorio
                if desk:
                    z_widget = self._posicion_z(self.hwnd)
                    z_desk = self._posicion_z(desk)
                    # el escritorio esta por delante del widget -> lo tapa
                    if z_widget is None or (z_desk is not None and z_desk < z_widget):
                        win32gui.SetWindowPos(desk, self.hwnd, 0, 0, 0, 0, flags)
            except Exception:
                pass
        else:
            self.lower()
        self.after(3000, self._al_fondo)

    # -- redimensionar arrastrando los bordes --------------------------------

    def _zona_borde(self, x, y):
        """'izq', 'der', 'arr', 'aba' o sus combinaciones en las esquinas."""
        ancho, alto = self._tam or (self.winfo_width(), self.winfo_height())
        zona = ""
        if x <= BORDE:
            zona += "izq"
        elif x >= ancho - BORDE:
            zona += "der"
        if y <= BORDE:
            zona += "arr"
        elif y >= alto - BORDE:
            zona += "aba"
        return zona

    def _cursor_borde(self, e):
        if self._redim or self._drag:
            return
        zona = self._zona_borde(e.x, e.y)
        cursores = {"izq": "sb_h_double_arrow", "der": "sb_h_double_arrow",
                    "arr": "sb_v_double_arrow", "aba": "sb_v_double_arrow",
                    "izqarr": "size_nw_se", "derarr": "size_ne_sw",
                    "izqaba": "size_ne_sw", "deraba": "size_nw_se"}
        try:
            self.lienzo.config(cursor=cursores.get(zona, ""))
        except tk.TclError:
            self.lienzo.config(cursor="sizing" if zona else "")

    def _inicio_arrastre(self, e):
        if e.state & 0x20000:     # Alt: mueve la ventana desde cualquier sitio
            self._drag = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())
            return "break"
        try:                      # clic sobre la columna de correo: no arrastra
            if "correo_col" in self.lienzo.gettags("current"):
                return
        except Exception:
            pass
        zona = self._zona_borde(e.x, e.y)
        if zona:
            rec_rej, rec_lista = getattr(self, "_rec", (0, 0))
            dias_r = max(CFG["dias_visibles"], 1)
            horas_r = max(CFG["hora_fin"] - CFG["hora_inicio"], 1)
            self._redim = {
                "zona": zona, "x0": e.x_root, "y0": e.y_root,
                # los valores «base», sin lo que se cede a las columnas
                "col": self._ancho_columna() + rec_rej / dias_r,
                "hora": self._alto_hora(),
                "lista": self._ancho_lista() + rec_lista,
                "alto_lista": CFG.get("alto_max_lista", 620),
                "celda": round(CFG.get("alto_celda_mes", 88) * ESCALA),
                "ancho_hora": self._ancho_hora() + rec_rej / horas_r,
                "fila_v": self._alto_fila_v(),
                "pos": self._pos,
            }
            return
        if self._evento_bajo_raton() is not None:
            return self._inicio_mover(e)      # sobre un evento manda el evento
        dy_rej = getattr(self, "_off_rejilla", (0, 0))[1]
        en_horas = (e.y >= CABECERA + self._alto_banda + dy_rej
                    if self.orientacion == "horizontal"
                    else e.y >= self._cab_horas() + dy_rej)
        hueco = (self._hueco_en(e.x, e.y)
                 if CFG.get("dibujar_huecos", True) and en_horas else None)
        if hueco is not None and self.vista in ("rejilla", "ambas"):
            # pintar un rato libre para crear ahí un evento con esa duración
            self._pintando = {"desde": hueco, "hasta": hueco, "activo": False,
                              "x0": e.x, "y0": e.y}
            return
        self._drag = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    # -- mover eventos arrastrandolos ---------------------------------------

    def _inicio_mover(self, e):
        """Prepara el arrastre de un evento (aun no mueve nada)."""
        if not CFG.get("arrastrar_eventos", True) or self.vista in ("lista", "mes"):
            return None
        ev = self._evento_bajo_raton()
        if ev is None or ev.todo_dia or ev.parte or es_festivo(ev.categoria):
            return None
        # agarrado por el borde final -> se estira, no se mueve
        caja = self._caja_horaria(ev.inicio, ev.fin)
        estirar = False
        if caja:
            x1, y1, x2, y2 = caja
            estirar = (abs(e.x - x2) <= 6 if self.orientacion == "vertical"
                       else abs(e.y - y2) <= 6)
        self._moviendo = {"ev": ev, "x0": e.x, "y0": e.y, "activo": False,
                          "destino": None, "agarre": self._hueco_en(e.x, e.y),
                          "estirar": estirar}
        return "break"

    def _caja_horaria(self, inicio, fin):
        """Rectangulo del lienzo que ocuparia un evento a esa hora."""
        dx, dy = getattr(self, "_off_rejilla", (0, 0))
        ini, _ = limites_ventana(desplazamiento=self.desplazamiento)
        dia = (inicio.date() - ini.date()).days
        if not 0 <= dia < CFG["dias_visibles"]:
            return None

        def dec(m):
            return m.hour + m.minute / 60

        h0, h1 = CFG["hora_inicio"], CFG["hora_fin"]
        a = min(max(dec(inicio), h0), h1)
        b = min(max(dec(fin) if fin.date() == inicio.date() else h1, h0), h1)
        if self.orientacion == "vertical":
            y1 = self._cab_horas() + dia * self._alto_fila_v() + 2 + dy
            x1 = self._margen_dias() + (a - h0) * self._ancho_hora() + dx
            return (x1, y1, x1 + max((b - a) * self._ancho_hora(), 6),
                    y1 + self._alto_fila_v() - 4)
        ancho_col = self._ancho_columna()
        x1 = MARGEN_IZQ + dia * ancho_col + 2 + dx
        y_rej = CABECERA + self._alto_banda + dy
        y1 = y_rej + (a - h0) * self._alto_hora()
        return (x1, y1, x1 + ancho_col - 4,
                y1 + max((b - a) * self._alto_hora(), 8))

    def _pintar_fantasma_hueco(self, desde, hasta):
        """Recuadro del rato que se está pintando con el ratón."""
        c = self.lienzo
        c.delete("fantasma")
        caja = self._caja_horaria(desde, hasta)
        if not caja:
            return
        x1, y1, x2, y2 = caja
        c.create_rectangle(x1, y1, x2, y2, outline=C["evento"], width=2,
                           dash=(4, 3), tags="fantasma")
        minutos = int((hasta - desde).total_seconds() // 60)
        c.create_text(x1 + 4, y1 + 2, anchor="nw", tags="fantasma",
                      text=f"{desde:%H:%M} · {minutos} min",
                      fill=C["evento"], font=letra(8, "bold"))

    def _pintar_fantasma(self, ev, destino):
        """Dibuja donde caeria el evento mientras se arrastra."""
        c = self.lienzo
        c.delete("fantasma")
        caja = self._caja_horaria(destino, destino + (ev.fin - ev.inicio))
        if not caja:
            return
        x1, y1, x2, y2 = caja
        c.create_rectangle(x1, y1, x2, y2, outline=C["ahora"], width=2,
                           dash=(4, 3), tags="fantasma")
        c.create_text(x1 + 4, y1 + 2, anchor="nw", tags="fantasma",
                      text=f"{destino:%H:%M}", fill=C["ahora"],
                      font=letra(8, "bold"))

    def _estirar_evento(self, ev, nuevo_fin):
        """Cambia solo cuánto dura, arrastrando su borde final."""
        if nuevo_fin == ev.fin:
            return False
        antes = ev.fin
        datos = self._datos_de(ev)
        datos["fin"] = nuevo_fin
        ev.fin = nuevo_fin
        repartir_columnas([e for e in self.eventos if not e.todo_dia])
        self.dibujar()

        def deshacer():
            ev.fin = antes
            repartir_columnas([e for e in self.eventos if not e.todo_dia])

        self._escribir_en_outlook(lambda: guardar_cita_outlook(datos, ev),
                                  deshacer, "No se ha podido alargar")
        minutos = int((nuevo_fin - ev.inicio).total_seconds() // 60)
        AvisoFlotante(self, f"{ev.titulo[:34]} · {minutos} min",
                      "Pulsa aquí para dejarlo como estaba",
                      self._color_evento(ev), 9,
                      al_pulsar=lambda: self._estirar_evento(ev, antes))
        return True

    def _escribir_en_outlook(self, trabajo, deshacer, aviso="No se ha podido guardar"):
        """Escribe en Outlook por detrás; la pantalla ya enseña el cambio.

        Si Outlook dice que no, se deshace lo que se había pintado y se avisa.
        Con 'guardar_en_segundo_plano' en falso, escribe al momento (lo usan
        las pruebas).
        """
        def terminar(hecho, motivo):
            if hecho:
                guardar_cache(self.eventos)
                self.after(3000, self.refrescar)
                return True
            try:
                deshacer()
            except Exception as exc:
                log(f"No se pudo deshacer tras el fallo: {exc}")
            self.dibujar()
            AvisoFlotante(self, aviso, motivo, C["ahora"], 10)
            return False

        if not CFG.get("guardar_en_segundo_plano", True):
            try:
                hecho, motivo = trabajo()
            except Exception as exc:
                hecho, motivo = False, str(exc)[:70]
            return terminar(hecho, motivo)

        def hilo():
            try:
                hecho, motivo = trabajo()
            except Exception as exc:
                hecho, motivo = False, str(exc)[:70]
            self.after(0, lambda: terminar(hecho, motivo))

        threading.Thread(target=hilo, daemon=True).start()
        return True

    def _mover_evento(self, ev, nuevo_inicio, avisar=True):
        """Cambia la hora del evento en Outlook y en la pantalla."""
        antes, antes_fin = ev.inicio, ev.fin
        duracion = ev.fin - ev.inicio
        ev.inicio, ev.fin = nuevo_inicio, nuevo_inicio + duracion
        repartir_columnas([e for e in self.eventos if not e.todo_dia])
        self._seguir_cuentas()
        self.dibujar()

        def deshacer():
            ev.inicio, ev.fin = antes, antes_fin
            repartir_columnas([e for e in self.eventos if not e.todo_dia])
            self._seguir_cuentas()

        self._escribir_en_outlook(lambda: mover_en_outlook(ev, nuevo_inicio),
                                  deshacer, "No se ha podido mover")
        if avisar:
            AvisoFlotante(
                self, f"{ev.titulo[:34]} → {fecha_corta(nuevo_inicio)} "
                      f"{nuevo_inicio:%H:%M}", "Pulsa aquí para deshacerlo",
                self._color_evento(ev), 9,
                al_pulsar=lambda: self._mover_evento(ev, antes, avisar=False))
        return True

    def _inicio_divisor_todos(self, e):
        """Empieza a arrastrar el borde izquierdo de la columna de to-dos.

        A diferencia del divisor de la lista, aqui cambia solo el ancho de la
        columna (la ventana entera se ensancha o encoge)."""
        self._divisor_t = {"x0": e.x_root, "ancho": self._ancho_todos()}
        self.lienzo.config(cursor="sb_h_double_arrow")
        return "break"

    def _reset_ancho_todos(self):
        CFG["ancho_todos"] = DEFAULTS["ancho_todos"]
        guardar_config()
        self._tam = None
        self.dibujar()

    def _cambiar_ancho_todos(self, px):
        CFG["ancho_todos"] = max(120, min(600, int(px)))
        guardar_config()
        self._tam = None
        self.dibujar()

    def _construir_menu_ancho_todos(self, padre):
        m = tk.Menu(padre, tearoff=0)
        actual = CFG.get("ancho_todos", DEFAULTS["ancho_todos"])
        for px in (150, 180, 210, 260, 320, 400):
            m.add_command(
                label=("● " if px == actual else "     ") + f"{px} px",
                command=lambda v=px: self._cambiar_ancho_todos(v))
        m.add_separator()
        m.add_command(label="Restablecer",
                      command=self._reset_ancho_todos)
        return m

    def _menu_ancho_todos(self, e=None):
        self._abrir_menu(self._construir_menu_ancho_todos(self.lienzo), e)
        return "break"

    def _inicio_divisor(self, e):
        """Empieza a mover la separacion entre el calendario y la lista."""
        lado = self._lado_lista()
        # si aun no hay tamaño calculado, el de la ventana real: con 0 el
        # reparto saldria negativo y la lista se iria al minimo de golpe
        total = (self._tam or (0, 0))[0] or self.winfo_width()
        rec_rej, rec_lista = getattr(self, "_rec", (0, 0))
        # el total que se reparte es el del calendario + lista, sin las
        # columnas de to-dos y correo; los valores son los «base»
        total = total - sum(self._anchos_columnas()) + rec_rej + rec_lista
        self._divisor = {"x0": e.x_root, "y0": e.y_root, "lado": lado,
                         "ancho": self._ancho_lista() + rec_lista,
                         "alto": CFG.get("alto_max_lista", 620),
                         "total": total}
        self.lienzo.config(
            cursor="sb_v_double_arrow" if lado in ("arriba", "abajo")
            else "sb_h_double_arrow")
        return "break"

    def _ajustar_ancho_rejilla(self, ancho_objetivo):
        """Cambia la columna (o la hora) para que la rejilla mida justo eso."""
        if self.orientacion == "vertical":
            horas = max(CFG["hora_fin"] - CFG["hora_inicio"], 1)
            util = ancho_objetivo - self._margen_dias() - 8
            CFG["ancho_hora"] = max(20, round(util / horas / ESCALA))
        else:
            dias = max(CFG["dias_visibles"], 1)
            util = ancho_objetivo - MARGEN_IZQ - 8
            CFG["ancho_fijo"] = False
            CFG["ancho_columna"] = max(24, round(util / dias / ESCALA))

    def _arrastrar(self, e):
        if getattr(self, "_pintando", None):
            p = self._pintando
            if not p["activo"] and abs(e.x - p["x0"]) + abs(e.y - p["y0"]) < 10:
                return "break"
            punto = self._hueco_en(e.x, e.y)
            if punto is None:
                return "break"
            p["activo"] = True
            paso = dt.timedelta(minutes=self._division())
            p["hasta"] = max(punto + paso, p["desde"] + paso)
            self._pintar_fantasma_hueco(p["desde"], p["hasta"])
            return "break"
        if self._moviendo:
            m = self._moviendo
            if not m["activo"] and abs(e.x - m["x0"]) + abs(e.y - m["y0"]) < 12:
                return "break"          # un temblor no mueve nada
            m["activo"] = True
            self._quitar_detalle()
            if m.get("estirar"):
                punto = self._hueco_en(e.x, e.y)
                if punto:
                    paso = dt.timedelta(minutes=self._division())
                    m["fin"] = max(punto + paso, m["ev"].inicio + paso)
                    self._pintar_fantasma_hueco(m["ev"].inicio, m["fin"])
                return "break"
            punto = self._hueco_en(e.x, e.y)
            # se mueve lo mismo que el raton: si lo agarras por el medio,
            # el evento no salta a la hora del cursor
            destino = (m["ev"].inicio + (punto - m["agarre"])
                       if punto and m["agarre"] else punto)
            if destino and destino != m["ev"].inicio:
                m["destino"] = destino
                self._pintar_fantasma(m["ev"], destino)
            return "break"
        if self._divisor_c:
            dx = self._divisor_c["x0"] - e.x_root
            nuevo = round((self._divisor_c["ancho"] + dx) / ESCALA)
            CFG["ancho_correo"] = max(160, min(700, nuevo))
            self._tam = None
            return self._dibujar_arrastre()
        if self._divisor_t:
            # arrastrar el borde izquierdo de la columna de to-dos: hacia la
            # izquierda la ensancha, hacia la derecha la estrecha
            dx = self._divisor_t["x0"] - e.x_root
            nuevo = round((self._divisor_t["ancho"] + dx) / ESCALA)
            CFG["ancho_todos"] = max(120, min(600, nuevo))
            self._tam = None
            return self._dibujar_arrastre()
        if self._divisor:
            lado = self._divisor.get("lado", "derecha")
            if lado in ("arriba", "abajo"):
                # apilada: la barra decide cuanto alto se lleva la lista
                dy = e.y_root - self._divisor["y0"]
                CFG["alto_max_lista"] = max(
                    60, min(900, self._divisor["alto"]
                            + (dy if lado == "abajo" else -dy)))
                self._tam = None
                return self._dibujar_arrastre()
            # el ancho total no cambia: lo que gana uno lo pierde el otro
            total = self._divisor["total"]
            self._quitar_reserva()      # se trabaja con los valores base
            señal = -1 if lado == "izquierda" else 1
            nuevo = self._divisor["ancho"] + señal * (
                self._divisor["x0"] - e.x_root)
            # la rejilla solo crece a saltos de un dia (o de una hora): se deja
            # ese margen en los topes para que el minimo se respete de verdad
            paso = (CFG["hora_fin"] - CFG["hora_inicio"]
                    if self.orientacion == "vertical" else CFG["dias_visibles"])
            nuevo = max(160 + paso, min(max(total - 220 - paso, 200), nuevo))
            self._ajustar_ancho_rejilla(total - nuevo)
            # la rejilla solo admite anchos "redondos": la lista se queda el resto
            # exacto, asi la ventana mide siempre lo mismo
            real = (self._medidas_vertical()[0] if self.orientacion == "vertical"
                    else self._ancho_rejilla())
            CFG["ancho_lista"] = max(160, round((total - real) / ESCALA))
            self._tam = None
            return self._dibujar_arrastre()
        if self._redim:
            return self._redimensionar(e)
        if self._drag:
            self.geometry(f"+{e.x_root - self._drag[0]}+{e.y_root - self._drag[1]}")
            self._reposicionar_paneles_acoplados()

    def _redimensionar(self, e):
        """Convierte el arrastre del borde en ancho de columna y alto de hora."""
        r = self._redim
        zona = r["zona"]
        dx, dy = e.x_root - r["x0"], e.y_root - r["y0"]
        x, y = r["pos"]

        if "izq" in zona:                 # el borde izquierdo mueve la ventana
            x = r["pos"][0] + dx
            dx = -dx
        if "arr" in zona:
            y = r["pos"][1] + dy
            dy = -dy

        if self.vista == "lista":
            if "izq" in zona or "der" in zona:
                CFG["ancho_lista"] = max(180, round((r["lista"] + dx) / ESCALA))
            if "arr" in zona or "aba" in zona:
                CFG["alto_max_lista"] = max(140, round(r["alto_lista"] + dy))
        elif self.vista == "mes":
            if "izq" in zona or "der" in zona:
                CFG["ancho_fijo"] = False
                CFG["ancho_columna"] = max(40, round((r["col"] + dx / 7) / ESCALA))
            if "arr" in zona or "aba" in zona:
                CFG["alto_celda_mes"] = max(40, round((r["celda"] + dy / 6) / ESCALA))
        elif self.orientacion == "vertical":
            dias = max(CFG["dias_visibles"], 1)
            horas = max(CFG["hora_fin"] - CFG["hora_inicio"], 1)
            if "izq" in zona or "der" in zona:
                CFG["ancho_hora"] = max(20, round(
                    (r["ancho_hora"] + dx / horas) / ESCALA))
            if "arr" in zona or "aba" in zona:
                CFG["alto_fila_v"] = max(26, round(
                    (r["fila_v"] + dy / dias) / ESCALA))
        else:
            dias = max(CFG["dias_visibles"], 1)
            horas = max(CFG["hora_fin"] - CFG["hora_inicio"], 1)
            if "izq" in zona or "der" in zona:
                CFG["ancho_fijo"] = False
                CFG["ancho_columna"] = max(24, round((r["col"] + dx / dias) / ESCALA))
            if "arr" in zona or "aba" in zona:
                CFG["alto_hora"] = max(12, round((r["hora"] + dy / horas) / ESCALA))

        self._pos = (int(x), int(y))
        self._tam = None                  # obliga a recalcular la geometria
        self._dibujar_arrastre()

    def _fin_redimension(self):
        self._redim = None
        self._hueco = None            # el usuario manda: ese es el nuevo hueco
        CFG["posicion_x"], CFG["posicion_y"] = self._pos
        CFG.setdefault("posiciones", {})[self.firma_pantallas()] = list(self._pos)
        guardar_config()
        log(f"Tamaño ajustado: columna {CFG['ancho_columna']}, "
            f"hora {CFG['alto_hora']}")

    def _fin_arrastre(self, _):
        if getattr(self, "_pintando", None):
            p, self._pintando = self._pintando, None
            self.lienzo.delete("fantasma")
            if p["activo"]:
                self._ficha(plantilla={"inicio": p["desde"], "fin": p["hasta"]})
            return
        if getattr(self, "_moviendo", None):
            m, self._moviendo = self._moviendo, None
            self.lienzo.delete("fantasma")
            if m["activo"] and m.get("estirar") and m.get("fin"):
                self._estirar_evento(m["ev"], m["fin"])
            elif m["activo"] and m["destino"] and m["destino"] != m["ev"].inicio:
                self._mover_evento(m["ev"], m["destino"])
            return
        if getattr(self, "_divisor_c", None):
            self._divisor_c = None
            self.lienzo.config(cursor="")
            guardar_config()
            log(f"Columna de correo: {CFG['ancho_correo']} px de ancho")
            return
        if getattr(self, "_divisor_t", None):
            self._divisor_t = None
            self.lienzo.config(cursor="")
            guardar_config()
            log(f"Columna to-dos: {CFG['ancho_todos']} px de ancho")
            return
        if getattr(self, "_divisor", None):
            apilada = self._divisor.get("lado") in ("arriba", "abajo")
            self._divisor = None
            self._hueco = None        # la rejilla tiene otro sitio a partir de ahora
            self.lienzo.config(cursor="")
            self._resaltar_divisor(False)
            guardar_config()
            log(f"Reparto: lista {CFG['alto_max_lista']} px de alto" if apilada
                else f"Reparto: lista {CFG['ancho_lista']} px de ancho")
            return
        if self._redim:
            return self._fin_redimension()
        self._drag = None
        self._pos = (self.winfo_x(), self.winfo_y())
        CFG["posicion_x"], CFG["posicion_y"] = self._pos
        CFG.setdefault("posiciones", {})[self.firma_pantallas()] = list(self._pos)
        guardar_config()

    # -- posicion segun las pantallas conectadas -----------------------------

    def firma_pantallas(self):
        """Identifica la combinacion de monitores actual."""
        try:
            import ctypes
            u = ctypes.windll.user32
            return (f"{u.GetSystemMetrics(78)}x{u.GetSystemMetrics(79)}"
                    f"@{u.GetSystemMetrics(80)}")
        except Exception:
            return f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}"

    def _area_visible(self):
        """(x, y, ancho, alto) del escritorio completo, con todos los monitores."""
        try:
            import ctypes
            u = ctypes.windll.user32
            return (u.GetSystemMetrics(76), u.GetSystemMetrics(77),
                    u.GetSystemMetrics(78), u.GetSystemMetrics(79))
        except Exception:
            return (0, 0, self.winfo_screenwidth(), self.winfo_screenheight())

    def _restaurar_posicion(self, avisar=False):
        """Recupera la posicion guardada para estas pantallas, o una visible."""
        firma = self.firma_pantallas()
        guardada = (CFG.get("posiciones") or {}).get(firma)
        x, y = (guardada if guardada
                else (CFG["posicion_x"], CFG["posicion_y"]))
        vx, vy, va, vh = self._area_visible()
        ancho, alto = self._tam or (400, 300)
        if not (vx - 20 <= x <= vx + va - 60 and vy - 20 <= y <= vy + vh - 40):
            x, y = vx + 40, vy + 40          # se habia quedado fuera de pantalla
            log(f"Posición fuera de las pantallas actuales; se recoloca en {x},{y}")
        self._pos = (int(x), int(y))
        self.geometry(f"+{self._pos[0]}+{self._pos[1]}")
        self._firma = firma
        if avisar:
            log(f"Pantallas {firma}: ventana en {self._pos[0]},{self._pos[1]}")

    def _vigilar_pantallas(self):
        """Si conectas o quitas un monitor, recoloca la ventana."""
        if self.firma_pantallas() != getattr(self, "_firma", None):
            log("Cambio de pantallas detectado")
            self._restaurar_posicion(avisar=True)
        self.after(20_000, self._vigilar_pantallas)

    # -- navegacion ----------------------------------------------------------

    def _mover_mes(self, n):
        """En la vista de mes, las flechas saltan de mes en mes."""
        base, _ = limites_ventana()
        actual = (base + dt.timedelta(days=self.desplazamiento)).date().replace(day=1)
        mes = actual.month + n
        anio = actual.year + (mes - 1) // 12
        self.ir_a_fecha(dt.date(anio, (mes - 1) % 12 + 1, 1))
        return "break"

    def _mover(self, dias):
        """Desplaza la ventana visible, descargando mas si se sale de lo que hay."""
        if self.vista == "mes" and abs(dias) >= 7:
            return self._mover_mes(-1 if dias < 0 else 1)
        self.desplazamiento = max(-TOPE_NAVEGACION, min(
            TOPE_NAVEGACION, self.desplazamiento + dias))
        self._asegurar_datos()
        self.dibujar()
        return "break"

    def _asegurar_datos(self):
        """Si la ventana sale de lo descargado, amplia el rango y vuelve a bajar."""
        ver = max(CFG["dias_visibles"], 42)      # el mes pinta hasta 6 semanas
        pide = {"adelante": self.desplazamiento + ver,
                "atras": -self.desplazamiento}
        base = {"adelante": max(CFG["dias_cache_adelante"],
                                CFG["dias_visibles"] + 14),
                "atras": CFG["dias_cache_atras"]}
        salio = False
        for lado, necesito in pide.items():
            if necesito > max(base[lado], ALCANCE[lado]):
                ALCANCE[lado] = necesito + 60    # margen para seguir navegando
                salio = True
        if salio:
            self.refrescar()

    def ir_a_fecha(self, fecha):
        """Coloca la ventana visible en esa fecha (si hace falta, descarga mas)."""
        base, _ = limites_ventana()
        self._mover(0)      # normaliza los topes
        objetivo = (fecha - base.date()).days
        self.desplazamiento = max(-TOPE_NAVEGACION,
                                  min(TOPE_NAVEGACION, objetivo))
        self._asegurar_datos()
        self.dibujar()
        log(f"Ir a {fecha} (desplazamiento {self.desplazamiento})")
        return self.desplazamiento == objetivo

    def _abrir_panel_ir(self, e=None):
        try:
            if getattr(self, "_panel", None) and self._panel.winfo_exists():
                self._panel.destroy()
        except Exception:
            pass
        self._panel = PanelIr(self)
        return "break"

    def _ir_a_hoy(self):
        self.desplazamiento = 0
        ALCANCE.update(atras=0, adelante=0)     # el proximo refresco vuelve a lo normal
        self.dibujar()
        return "break"

    def _construir_menu_festivos(self, padre):
        m = tk.Menu(padre, tearoff=0)
        for clave, conf in grupos_de_festivos():
            if clave.startswith("municipio:"):
                etiqueta = conf.get("nombre") or clave.split(":", 1)[1]
            else:
                etiqueta = conf.get("nombre") or clave.capitalize()
                if clave == "local" and not conf.get("fijos"):
                    continue          # el hueco viejo, sin rellenar: no estorba
            var = self._vars_festivos.setdefault(clave, tk.BooleanVar())
            var.set(conf.get("activo", True))
            m.add_checkbutton(
                label=etiqueta, variable=var,
                foreground=conf.get("color", C["texto"]),
                command=lambda c=clave, v=var: self._cambiar_festivo(c, v.get()))
        m.add_separator()
        m.add_command(label="Añadir otro municipio…",
                      command=self._anadir_municipio)
        municipios = (CFG.get("festivos") or {}).get("municipios") or []
        if municipios:
            quitar = tk.Menu(m, tearoff=0)
            for mu in municipios:
                quitar.add_command(
                    label=mu.get("nombre", ""),
                    command=lambda k=mu.get("clave"): self._quitar_municipio(k))
            m.add_cascade(label="Quitar un municipio", menu=quitar)
        m.add_separator()
        m.add_command(label="Ver todos",
                      command=lambda: self._cambiar_festivo(None, True))
        m.add_command(label="Quitar todos",
                      command=lambda: self._cambiar_festivo(None, False))
        return m

    def _anadir_municipio(self):
        """Alta rápida: nombre y las dos fiestas locales en dd/mm."""
        try:
            from tkinter import simpledialog
            nombre = simpledialog.askstring(
                "Municipio nuevo", "Nombre del municipio:", parent=self)
            if not (nombre or "").strip():
                return
            fechas = simpledialog.askstring(
                f"Fiestas de {nombre.strip()}",
                "Sus dos fiestas locales en dd/mm, separadas por coma\n"
                "(por ejemplo: 06/08, 09/10):", parent=self)
        except Exception as exc:
            log(f"No se pudo pedir el municipio: {exc}")
            return
        fijos = {}
        for trozo in (fechas or "").split(","):
            m = re.fullmatch(r"\s*(\d{1,2})[/\-.](\d{1,2})\s*", trozo)
            if not m:
                continue
            dia, mes = int(m[1]), int(m[2])
            try:
                dt.date(2024, mes, dia)      # 2024 valida el 29 de febrero
            except ValueError:
                continue
            fijos[f"{mes:02d}-{dia:02d}"] = "Fiesta local"
        if not fijos:
            AvisoFlotante(self, "No he entendido las fechas",
                          "Escríbelas como 06/08, 09/10", C["ahora"], 8)
            return
        clave = sin_acentos(nombre.strip().lower()).replace(" ", "-")[:24]
        lista = CFG.setdefault("festivos", {}).setdefault("municipios", [])
        lista = [m for m in lista if m.get("clave") != clave]
        lista.append({"clave": clave, "nombre": nombre.strip(), "activo": True,
                      "color": COLORES_HABITO[len(lista) % len(COLORES_HABITO)],
                      "fijos": fijos, "extra": {}})
        CFG["festivos"]["municipios"] = lista
        guardar_config()
        log(f"Municipio añadido: {nombre.strip()} ({len(fijos)} fiestas)")
        self.dibujar()

    def _quitar_municipio(self, clave):
        fest = CFG.setdefault("festivos", {})
        fest["municipios"] = [m for m in (fest.get("municipios") or [])
                              if m.get("clave") != clave]
        guardar_config()
        self.dibujar()

    def _menu_festivos(self, e=None):
        m = self._construir_menu_festivos(self)
        x = e.x_root if e is not None else self.winfo_rootx() + 40
        y = e.y_root if e is not None else self.winfo_rooty() + BARRA
        m.tk_popup(x, y)

    def _cambiar_festivo(self, clave, activo):
        fest = CFG.setdefault("festivos", {})
        for c, conf in grupos_de_festivos():
            if clave is None or c == clave:
                if c.startswith("municipio:"):
                    conf["activo"] = activo      # es el dict de la lista
                else:
                    fest.setdefault(c, {})["activo"] = activo
        guardar_config()
        self.dibujar()

    # -- horario visible (se cambia en caliente) ----------------------------

    def _rueda(self, e, arriba):
        """Sobre la columna de horas mueve el horario; en el resto, los dias."""
        if self.vista == "mes":
            self._mover_mes(-1 if arriba else 1)
        elif self.vista != "lista" and e.x < MARGEN_IZQ:
            self._cambiar_horario(mover=-1 if arriba else 1)
        else:
            self._mover(-1 if arriba else 1)
        return "break"

    def _cambiar_horario(self, inicio=0, fin=0, mover=0, fijo=None):
        if fijo:
            h0, h1 = fijo
        else:
            h0 = CFG["hora_inicio"] + inicio + mover
            h1 = CFG["hora_fin"] + fin + mover
        h0 = max(0, min(h0, 22))
        h1 = max(h0 + 2, min(h1, 24))
        if (h0, h1) == (CFG["hora_inicio"], CFG["hora_fin"]):
            return
        CFG["hora_inicio"], CFG["hora_fin"] = h0, h1
        guardar_config()
        self.dibujar()

    def _ajustar_horario(self):
        """Encaja el horario visible con los eventos de la ventana."""
        ini, _ = limites_ventana(desplazamiento=self.desplazamiento)
        dias = CFG["dias_visibles"]
        horas = []
        for ev in self._mostrables():
            if ev.todo_dia or not 0 <= (ev.inicio.date() - ini.date()).days < dias:
                continue
            horas.append((ev.inicio.hour,
                          ev.fin.hour + (1 if ev.fin.minute else 0)))
        if not horas:
            return
        self._cambiar_horario(fijo=(min(h for h, _ in horas),
                                    max(f for _, f in horas)))

    def _alternar_lineas(self):
        CFG["lineas_rejilla"] = not CFG.get("lineas_rejilla", True)
        guardar_config()
        self.dibujar()

    def _alternar_ajuste(self, clave):
        """Enciende o apaga un ajuste de si/no y repinta."""
        CFG[clave] = not CFG.get(clave, True)
        guardar_config()
        self._tam = None
        log(f"{clave}: {CFG[clave]}")
        self.dibujar()

    def _cambiar_lineas_lista(self, n):
        """Cuantas lineas puede ocupar cada titulo de la lista."""
        CFG["lineas_lista"] = max(1, min(4, int(n)))
        guardar_config()
        self._tam = None
        log(f"Texto de la lista: hasta {CFG['lineas_lista']} líneas")
        self.dibujar()

    def _cambiar_tema(self, tema):
        CFG["tema"] = tema
        aplicar_tema()
        guardar_config()
        try:
            self.configure(bg=C["fondo"])
            self.lienzo.configure(bg=C["fondo"])
        except Exception:
            pass
        log(f"Tema: {tema} ({'claro' if tema_claro_activo() else 'oscuro'})")
        self.dibujar()

    def _vigilar_tema(self):
        """Con el tema en automatico, sigue los cambios de Windows."""
        if CFG.get("tema") == "auto" and tema_claro_activo() != self._claro:
            self._claro = tema_claro_activo()
            self._cambiar_tema("auto")
        self.after(60_000, self._vigilar_tema)

    def _vigilar_escritorio(self):
        """Con «solo en este escritorio», oculta el widget en los demás."""
        if self.hwnd and self.visible:
            actual, todos = escritorios_virtuales()
            fijo = CFG.get("escritorio_id")
            # sin escritorio guardado (o ya no existe, o es de otro equipo): adopta el actual
            if (CFG.get("solo_un_escritorio") and actual
                    and (not fijo or fijo not in todos)):
                fijo = CFG["escritorio_id"] = actual
                guardar_config()
            fuera = bool(CFG.get("solo_un_escritorio") and actual and fijo
                         and actual != fijo
                         and time.time() > getattr(self, "_alarma_hasta", 0))
            if fuera != self._en_otro_escritorio:
                self._en_otro_escritorio = fuera
                try:
                    import win32con
                    import win32gui
                    win32gui.ShowWindow(self.hwnd, win32con.SW_HIDE if fuera
                                        else win32con.SW_SHOWNA)
                except Exception:
                    pass
                # los paneles y fichas (sin botón en la barra) también salen en todos los escritorios
                if fuera:
                    self._ocultas = [h for h in self.winfo_children()
                                     if isinstance(h, tk.Toplevel) and h.state() == "normal"]
                    for h in self._ocultas:
                        h.withdraw()
                else:
                    for h in getattr(self, "_ocultas", []):
                        try:
                            h.deiconify()
                        except Exception:
                            pass
                    self._ocultas = []
        self.after(300, self._vigilar_escritorio)

    def _alternar_un_escritorio(self):
        activar = not CFG.get("solo_un_escritorio", False)
        CFG["solo_un_escritorio"] = activar
        CFG["escritorio_id"] = escritorios_virtuales()[0] if activar else None
        guardar_config()
        log(f"Solo en un escritorio: {activar} ({CFG['escritorio_id']})")

    def _cambiar_escala(self, delta=0.0, fijo=None):
        nueva = max(0.7, min(2.0, fijo if fijo else round(ESCALA + delta, 2)))
        if nueva == ESCALA:
            return "break"
        CFG["escala"] = nueva
        self._hueco = None
        aplicar_escala()
        guardar_config()
        self._tam = None          # obliga a recalcular la geometria
        self.dibujar()
        return "break"

    def _cambiar_lado_lista(self, lado):
        """Pone la lista a la derecha, a la izquierda, arriba o abajo."""
        CFG["lado_lista"] = lado
        if self.vista != "ambas":
            self.vista = CFG["vista"] = "ambas"
        guardar_config()
        self._hueco = None
        self._tam = None
        log(f"La lista va: {lado}")
        self.dibujar()
        return "break"

    def _cambiar_opacidad(self, valor):
        """0.3 a 1.0: cuanto se ve el escritorio a través del widget."""
        valor = max(0.3, min(1.0, round(float(valor), 2)))
        CFG["opacidad"] = valor
        guardar_config()
        try:
            self.attributes("-alpha", valor)
        except Exception as exc:
            log(f"No se pudo cambiar la transparencia: {exc}")
        log(f"Transparencia: {int(valor * 100)} %")
        return "break"

    def _opacidad_a_medida(self):
        try:
            from tkinter import simpledialog
            n = simpledialog.askinteger(
                "Transparencia", "¿Cuánto se ve el widget? (30 a 100)",
                parent=self, initialvalue=round(CFG.get("opacidad", 1.0) * 100),
                minvalue=30, maxvalue=100)
        except Exception as exc:
            log(f"No se pudo pedir la transparencia: {exc}")
            return
        if n:
            self._cambiar_opacidad(n / 100)

    def _cambiar_letra(self, delta=0.0, fijo=None):
        """Sube o baja solo el tamaño del texto, sin tocar el de la ventana."""
        nueva = max(0.7, min(1.8, fijo if fijo else round(LETRA + delta, 2)))
        if nueva == LETRA:
            return "break"
        CFG["escala_letra"] = nueva
        aplicar_escala()
        guardar_config()
        self._altos_linea.clear()      # las medidas de texto cambian
        self._fuentes.clear()
        self._tam = None
        log(f"Tamaño de la letra: {int(nueva * 100)} %")
        self.dibujar()
        return "break"

    def _cambiar_division(self, minutos):
        """Cada cuántos minutos se divide la rejilla."""
        CFG["division_minutos"] = max(5, min(240, int(minutos)))
        guardar_config()
        log(f"Divisiones de la rejilla: {CFG['division_minutos']} min")
        self.dibujar()

    def _division_a_medida(self):
        """Pide los minutos a mano, entre 5 y 240."""
        try:
            from tkinter import simpledialog
            n = simpledialog.askinteger(
                "Divisiones de la rejilla",
                "¿Cada cuántos minutos quieres una raya? (5 a 240)",
                parent=self, initialvalue=self._division(),
                minvalue=5, maxvalue=240)
        except Exception as exc:
            log(f"No se pudo pedir la división: {exc}")
            return
        if n:
            self._cambiar_division(n)

    def _elegir_hora(self, cual, h):
        """Fija la hora de inicio o la de fin; la otra se aparta si estorba."""
        if cual == "inicio":
            self._cambiar_horario(fijo=(h, max(CFG["hora_fin"], h + 2)))
        else:
            self._cambiar_horario(fijo=(min(CFG["hora_inicio"], h - 2), h))

    def _submenu_horas(self, padre, cual):
        """Lista de horas en punto para elegir el principio o el final."""
        m = tk.Menu(padre, tearoff=0)
        horas = range(0, 23) if cual == "inicio" else range(2, 25)
        actual = CFG["hora_inicio"] if cual == "inicio" else CFG["hora_fin"]
        for i, h in enumerate(horas):
            m.add_command(label=("● " if h == actual else "     ") + f"{h:02d}:00",
                          columnbreak=1 if i and i % 8 == 0 else 0,
                          command=lambda v=h, c=cual: self._elegir_hora(c, v))
        return m

    def _menu_horario(self, e=None):
        m = tk.Menu(self, tearoff=0)
        m.add_command(label=f"Ahora se ve de {CFG['hora_inicio']}:00 "
                            f"a {CFG['hora_fin']}:00", state="disabled")
        m.add_separator()
        m.add_command(label="Empezar 1 h antes",
                      command=lambda: self._cambiar_horario(inicio=-1))
        m.add_command(label="Empezar 1 h después",
                      command=lambda: self._cambiar_horario(inicio=1))
        m.add_command(label="Terminar 1 h antes",
                      command=lambda: self._cambiar_horario(fin=-1))
        m.add_command(label="Terminar 1 h después",
                      command=lambda: self._cambiar_horario(fin=1))
        m.add_separator()
        m.add_cascade(label=f"Empieza a las…  ({CFG['hora_inicio']:02d}:00)",
                      menu=self._submenu_horas(m, "inicio"))
        m.add_cascade(label=f"Termina a las…  ({CFG['hora_fin']:02d}:00)",
                      menu=self._submenu_horas(m, "fin"))
        m.add_separator()
        division = tk.Menu(m, tearoff=0)
        for n, etiqueta in ((10, "Cada 10 minutos"), (15, "Cada cuarto de hora"),
                            (20, "Cada 20 minutos"), (30, "Cada media hora"),
                            (60, "Cada hora"), (120, "Cada dos horas")):
            division.add_command(
                label=("● " if self._division() == n else "     ") + etiqueta,
                command=lambda v=n: self._cambiar_division(v))
        division.add_separator()
        division.add_command(label="Escribir los minutos…",
                             command=self._division_a_medida)
        m.add_cascade(label=f"Rayas cada…  ({self._division()} min)",
                      menu=division)
        m.add_separator()
        m.add_command(label="Ajustar a mis eventos", command=self._ajustar_horario)
        m.add_command(label="Jornada (8 – 20)",
                      command=lambda: self._cambiar_horario(fijo=(8, 20)))
        m.add_command(label="Día completo (0 – 24)",
                      command=lambda: self._cambiar_horario(fijo=(0, 24)))
        m.add_separator()
        m.add_command(label="(rueda del ratón sobre las horas para desplazarlo)",
                      state="disabled")
        self._abrir_menu(m, e)

    def _abrir_menu(self, m, e=None):
        x = e.x_root if e is not None else self.winfo_rootx() + 40
        y = e.y_root if e is not None else self.winfo_rooty() + BARRA
        m.tk_popup(x, y)

    # -- varios calendarios --------------------------------------------------

    def _construir_menu_calendarios(self, padre):
        m = tk.Menu(padre, tearoff=0)
        for i, conf in enumerate(CFG["calendarios"]):
            var = self._vars_festivos.setdefault(f"cal{i}", tk.BooleanVar())
            var.set(conf.get("activo", True))
            m.add_checkbutton(label=conf.get("nombre", f"Calendario {i + 1}"),
                              variable=var, foreground=conf.get("color"),
                              command=lambda n=i, v=var: self._activar_calendario(n, v.get()))
        return m

    def _activar_calendario(self, indice, activo):
        CFG["calendarios"][indice]["activo"] = activo
        guardar_config()
        self.refrescar()

    NOMBRES_VISTA = {"rejilla": "Calendario", "lista": "Lista",
                     "ambas": "Cal + lista", "mes": "Mes"}

    def _etiqueta_vista(self):
        """Lo que se lee en el botón único de vistas."""
        return self.NOMBRES_VISTA.get(self.vista, "Vista") + " ▾"

    def _siguiente_vista(self):
        """Un clic en el botón pasa a la siguiente vista, sin abrir menús."""
        orden = ["rejilla", "ambas", "lista", "mes"]
        try:
            siguiente = orden[(orden.index(self.vista) + 1) % len(orden)]
        except ValueError:
            siguiente = "ambas"
        return self._cambiar_vista(siguiente)

    def _menu_vistas(self, e=None):
        """Todas las vistas en un solo sitio, para no llenar la barra."""
        m = tk.Menu(self, tearoff=0)
        for clave in ("rejilla", "lista", "ambas", "mes"):
            m.add_command(
                label=("● " if self.vista == clave else "     ")
                      + self.NOMBRES_VISTA[clave],
                command=lambda c=clave: self._cambiar_vista(c))
        m.add_separator()
        abierto = bool(getattr(self, "_panel_plantillas", None))
        m.add_command(label=("✓ " if abierto else "     ")
                            + "Plantillas",
                      command=lambda: self._abrir_plantillas("plantillas"))
        abierto_tareas = bool(getattr(self, "_panel_tareas", None))
        m.add_command(label=("✓ " if abierto_tareas else "     ")
                            + "Tareas permanentes",
                      command=lambda: self._abrir_plantillas("tareas"))
        abierto_hab = bool(getattr(self, "_panel_habitos", None))
        m.add_command(label=("✓ " if abierto_hab else "     ") + "Hábitos",
                      command=self._abrir_habitos)
        abierto_correo = bool(getattr(self, "_panel_correo", None))
        m.add_command(label=("✓ " if abierto_correo else "     ") + "Correo",
                      command=self._abrir_correo)
        sitio = tk.Menu(m, tearoff=0)
        for clave, etiqueta in (("derecha", "A la derecha"),
                                ("izquierda", "A la izquierda"),
                                ("arriba", "Arriba"), ("abajo", "Abajo")):
            sitio.add_command(
                label=("● " if self._lado_lista() == clave else "     ")
                      + etiqueta,
                command=lambda c=clave: self._cambiar_lado_lista(c))
        m.add_cascade(label="La lista va…", menu=sitio)
        m.add_separator()
        m.add_command(label="Horizontal", command=lambda:
                      self._cambiar_orientacion("horizontal"))
        m.add_command(label="Vertical", command=lambda:
                      self._cambiar_orientacion("vertical"))
        self._abrir_menu(m, e)
        return "break"

    def _conservar_ancho_al_cambiar(self, antes, vista):
        """Al pasar de Calendario a Cal + lista (y al volver) el widget no
        cambia de ancho: la lista se reparte el sitio con el calendario."""
        self._quitar_reserva()
        r = (self._medidas_vertical()[0] if self.orientacion == "vertical"
             else self._ancho_rejilla())
        lista = round(CFG["ancho_lista"] * ESCALA)
        total = r if antes == "rejilla" else r + lista
        self._ajustar_ancho_rejilla(total - lista if vista == "ambas" else total)
        self._tam = None

    def _cambiar_vista(self, vista):
        antes = self.vista
        if (CFG.get("columnas_dentro", True) and antes != vista
                and {antes, vista} == {"rejilla", "ambas"}
                and self._lado_lista() in ("izquierda", "derecha")):
            try:
                self._conservar_ancho_al_cambiar(antes, vista)
            except Exception as exc:
                log(f"Conservar ancho al cambiar de vista: {exc!r}")
        self.vista = CFG["vista"] = vista
        self._hueco = None
        guardar_config()
        self.dibujar()

    def _menu_dias(self, e=None):
        """Los días a la vista: valores de siempre y uno a medida."""
        m = tk.Menu(self, tearoff=0)
        for n, texto in ((1, "Solo hoy"), (3, "3 días"),
                         (5, "5 días (semana laboral)"), (7, "7 días"),
                         (10, "10 días"), (14, "14 días"), (21, "21 días"),
                         (30, "30 días")):
            m.add_command(
                label=("● " if CFG["dias_visibles"] == n else "     ") + texto,
                command=lambda d=n: self._cambiar_dias(d))
        m.add_separator()
        m.add_command(
            label=("Empezar en lunes" if CFG.get("inicio_ventana", "hoy") == "hoy"
                   else "Empezar en hoy"),
            command=self._cambiar_inicio)
        m.add_command(label=f"Escribir un número (1 a {MAX_DIAS})…",
                      command=self._dias_a_medida)
        self._abrir_menu(m, e)
        return "break"

    def _dias_a_medida(self):
        """Pide un numero de dias entre 1 y MAX_DIAS."""
        try:
            from tkinter import simpledialog
            n = simpledialog.askinteger(
                "Días a la vista", f"¿Cuántos días quieres ver? (1 a {MAX_DIAS})",
                parent=self, initialvalue=CFG["dias_visibles"],
                minvalue=1, maxvalue=MAX_DIAS)
        except Exception as exc:
            log(f"No se pudo pedir el número de días: {exc}")
            return
        if n:
            self._cambiar_dias(n)

    def _hueco_rejilla(self):
        """Lo que ocupa ahora la rejilla en el sentido en que crecen los dias."""
        return (self._medidas_vertical()[1] if self.orientacion == "vertical"
                else self._ancho_rejilla())

    def _encajar_dias(self, dias):
        """Mete los dias nuevos en el mismo sitio que la rejilla ya ocupa."""
        if self.vista not in ("rejilla", "ambas") or dias == CFG["dias_visibles"]:
            return
        if self._hueco is None:       # se recuerda para no ir encogiendo a cada paso
            self._hueco = self._hueco_rejilla()
        if self.orientacion == "vertical":
            util = self._hueco - self._cab_horas() - ALTO_ESTADO - 2
            CFG["alto_fila_v"] = max(26, round(util / dias / ESCALA))
        else:
            util = self._hueco - MARGEN_IZQ - 8
            CFG["ancho_fijo"] = False
            CFG["ancho_columna"] = max(18, round(util / dias / ESCALA))

    def _cambiar_dias(self, dias):
        dias = max(1, min(MAX_DIAS, int(dias)))
        self._encajar_dias(dias)
        CFG["dias_visibles"] = dias
        # el desplazamiento se recorta si la nueva ventana se sale del cache
        self.desplazamiento = min(self.desplazamiento, TOPE_NAVEGACION - dias)
        guardar_config()
        self.dibujar()

    def _menu(self, e):
        """Clic derecho: lo más usado a la vista y el resto agrupado."""
        # el «break» de un menú propio de un elemento no frena este, que es el
        # de la ventana: sin esto salía el general encima del del evento
        propios = {"evento", "btn_vista", "btn_avisos", "divisor",
                   "divisor_todos", "divisor_correo", "habitos", "correo"}
        for item in self.lienzo.find_withtag("current"):
            if propios & set(self.lienzo.gettags(item)):
                return "break"
        m = tk.Menu(self, tearoff=0)

        def marca(activo, etiqueta, radio=False):
            return (("● " if radio else "✓ ") if activo else "     ") + etiqueta

        # -- lo de cada día
        m.add_command(label="Nuevo evento…", command=lambda: self._nuevo_evento(e))
        m.add_command(label="Buscar o ir a una fecha…",
                      command=self._abrir_panel_ir)
        m.add_command(label="Volver a hoy", command=self._ir_a_hoy)
        m.add_command(label="Actualizar ahora", command=self.refrescar)
        m.add_separator()
        m.add_cascade(label="Correo en el widget",
                      menu=self._construir_menu_correo_columna(m))
        m.add_command(label="Comandos…", command=self._abrir_paleta,
                      accelerator="Ctrl+K")
        m.add_command(label="Panel de control…",
                      command=self._abrir_panel_control)
        m.add_command(label="Ayuda y atajos…", command=self._abrir_ayuda_general)
        m.add_command(label="Buscar actualizaciones…", command=lambda: self._buscar_actualizacion(True))
        m.add_separator()

        # -- Vista: cómo se ve el calendario
        vista = tk.Menu(m, tearoff=0)
        for etiqueta, clave in (("Calendario", "rejilla"), ("Lista", "lista"),
                                ("Las dos", "ambas"), ("Mes completo", "mes")):
            vista.add_command(label=marca(self.vista == clave, etiqueta, True),
                              command=lambda v=clave: self._cambiar_vista(v))
        vista.add_separator()
        etiquetas_dias = {1: "Solo el día de hoy", 3: "Próximos 3 días",
                          5: "Próximos 5 días (semana laboral)"}
        for n in (1, 3, 5, 7, 14, 21):
            vista.add_command(
                label=marca(CFG["dias_visibles"] == n,
                            etiquetas_dias.get(n, f"Próximos {n} días"), True),
                command=lambda d=n: self._cambiar_dias(d))
        vista.add_command(label=f"Otro número de días (1 a {MAX_DIAS})…",
                          command=self._dias_a_medida)
        vista.add_command(
            label=("Empezar en lunes" if CFG.get("inicio_ventana", "hoy") == "hoy"
                   else "Empezar en hoy"),
            command=self._cambiar_inicio)
        vista.add_separator()
        for clave, etiqueta in (("horizontal", "Calendario horizontal"),
                                ("vertical", "Calendario vertical")):
            vista.add_command(
                label=marca(self.orientacion == clave, etiqueta, True),
                command=lambda c=clave: self._cambiar_orientacion(c))
        sitio = tk.Menu(vista, tearoff=0)
        for clave, etiqueta in (("derecha", "A la derecha"),
                                ("izquierda", "A la izquierda"),
                                ("arriba", "Arriba"), ("abajo", "Abajo")):
            sitio.add_command(
                label=marca(self._lado_lista() == clave, etiqueta, True),
                command=lambda c=clave: self._cambiar_lado_lista(c))
        vista.add_cascade(label="La lista va…", menu=sitio)
        vista.add_command(label="Horario visible…", command=self._menu_horario)
        vista.add_command(
            label=marca(CFG.get("lineas_rejilla", True), "Líneas de la rejilla"),
            command=self._alternar_lineas)
        lineas = tk.Menu(vista, tearoff=0)
        for n, etiqueta in ((1, "1 línea (más compacto)"), (2, "Hasta 2 líneas"),
                            (3, "Hasta 3 líneas"), (4, "Hasta 4 líneas")):
            lineas.add_command(
                label=marca(self._lineas_max() == n, etiqueta, True),
                command=lambda v=n: self._cambiar_lineas_lista(v))
        vista.add_cascade(label="Texto de la lista", menu=lineas)
        m.add_cascade(label="Vista", menu=vista)

        # -- Qué se muestra: calendarios, festivos, filtros
        datos = tk.Menu(m, tearoff=0)
        datos.add_cascade(label="Festivos",
                          menu=self._construir_menu_festivos(datos))
        datos.add_cascade(label="Filtrar",
                          menu=self._construir_menu_filtro(datos))
        if len(CFG["calendarios"]) > 1:
            datos.add_cascade(label="Calendarios",
                              menu=self._construir_menu_calendarios(datos))
        datos.add_cascade(label=f"Perfil ({CFG.get('perfil') or 'sin perfil'})",
                          menu=self._menu_perfiles(datos))
        origen = tk.Menu(datos, tearoff=0)
        for clave, etiqueta in (
                ("ics", "Enlaces ICS publicados"),
                ("outlook", "Outlook instalado (con etiquetas)"),
                ("auto", "Outlook si responde, si no los ICS")):
            origen.add_command(
                label=marca(CFG.get("origen", "ics") == clave, etiqueta, True),
                command=lambda c=clave: self._cambiar_origen(c))
        origen.add_separator()
        origen.add_command(
            label=marca(CFG.get("leer_notas", False),
                        "Leer también las notas (Outlook puede pedir permiso)"),
            command=self._alternar_notas)
        datos.add_cascade(label="Leer los eventos de", menu=origen)
        datos.add_cascade(label="Avisos y barra de abajo",
                          menu=self._menu_avisos(datos))
        m.add_cascade(label="Datos y avisos", menu=datos)

        # -- Aspecto
        aspecto = tk.Menu(m, tearoff=0)
        tema = tk.Menu(aspecto, tearoff=0)
        for clave, etiqueta in (("oscuro", "Oscuro"), ("claro", "Claro"),
                                ("auto", "Según Windows")):
            tema.add_command(
                label=marca(CFG.get("tema", "oscuro") == clave, etiqueta, True),
                command=lambda c=clave: self._cambiar_tema(c))
        aspecto.add_cascade(label="Tema", menu=tema)
        tam = tk.Menu(aspecto, tearoff=0)
        tam.add_command(label="Más grande",
                        command=lambda: self._cambiar_escala(0.1))
        tam.add_command(label="Más pequeño",
                        command=lambda: self._cambiar_escala(-0.1))
        tam.add_command(label="Tamaño normal",
                        command=lambda: self._cambiar_escala(fijo=1.0))
        tam.add_separator()
        tam.add_command(label="(Ctrl + rueda del ratón)", state="disabled")
        aspecto.add_cascade(label=f"Tamaño ({int(ESCALA * 100)} %)", menu=tam)
        tipo = tk.Menu(aspecto, tearoff=0)
        tipo.add_command(label="Letra más grande",
                         command=lambda: self._cambiar_letra(0.1))
        tipo.add_command(label="Letra más pequeña",
                         command=lambda: self._cambiar_letra(-0.1))
        tipo.add_command(label="Letra normal",
                         command=lambda: self._cambiar_letra(fijo=1.0))
        tipo.add_separator()
        tipo.add_command(label="(Ctrl + Mayús + rueda del ratón)",
                         state="disabled")
        aspecto.add_cascade(label=f"Tamaño de la letra ({int(LETRA * 100)} %)",
                            menu=tipo)
        barra = tk.Menu(aspecto, tearoff=0)
        for etiqueta, v in (("Pequeña", 1.0), ("Media", 1.5), ("Grande", 2.0)):
            barra.add_command(
                label=marca(abs(float(CFG.get("barra_zoom", 1.5)) - v) < 0.05,
                            etiqueta, True),
                command=lambda z=v: self._cambiar_zoom_barra(z))
        aspecto.add_cascade(label="Tamaño de la barra de arriba", menu=barra)
        opacidad = tk.Menu(aspecto, tearoff=0)
        actual_op = round(float(CFG.get("opacidad", 1.0)) * 100)
        for valor in (100, 95, 90, 85, 80, 70, 60, 50):
            opacidad.add_command(
                label=marca(actual_op == valor, f"{valor} %", True),
                command=lambda v=valor: self._cambiar_opacidad(v / 100))
        opacidad.add_separator()
        opacidad.add_command(label="Escribir el porcentaje…",
                             command=self._opacidad_a_medida)
        aspecto.add_cascade(label=f"Transparencia ({actual_op} %)",
                            menu=opacidad)
        aspecto.add_cascade(
            label=f"Barra separadora ({CFG.get('grosor_divisor', 11)} px)",
            menu=self._construir_menu_divisor(aspecto))
        aspecto.add_command(
            label=marca(CFG.get("menu_moderno", True),
                        "Menú de clic derecho moderno"),
            command=lambda: self._alternar_ajuste("menu_moderno"))
        m.add_cascade(label="Aspecto", menu=aspecto)

        # -- Comportamiento al crear y mover eventos
        hacer = tk.Menu(m, tearoff=0)
        hacer.add_command(
            label=marca(CFG.get("doble_clic_crea", False),
                        "Crear también con doble clic en un hueco"),
            command=self._alternar_doble_clic)
        hacer.add_command(
            label=marca(CFG.get("arrastrar_eventos", True),
                        "Mover eventos arrastrándolos"),
            command=lambda: self._alternar_ajuste("arrastrar_eventos"))
        hacer.add_command(
            label=marca(CFG.get("dibujar_huecos", True),
                        "Crear pintando un rato con el ratón"),
            command=lambda: self._alternar_ajuste("dibujar_huecos"))
        hacer.add_command(label="Crear desde una plantilla…",
                          command=self._menu_plantillas)
        donde = tk.Menu(hacer, tearoff=0)
        for clave, etiqueta in (
                ("ficha", "Aquí mismo, sin abrir Outlook"),
                ("auto", "Outlook web, ya rellenado"),
                ("nuevo", "App del Outlook nuevo (sin rellenar)"),
                ("outlook", "Outlook clásico, ya rellenado")):
            donde.add_command(
                label=marca(CFG.get("crear_con", "ficha") == clave, etiqueta,
                            True),
                command=lambda c=clave: self._donde_crear(c))
        hacer.add_cascade(label="Crear los eventos en", menu=donde)
        tachados = tk.Menu(hacer, tearoff=0)
        tachados.add_command(
            label=marca(CFG.get("tachar_pasados"),
                        "Tachar solo lo que ya ha pasado"),
            command=lambda: self._alternar_ajuste("tachar_pasados"))
        tachados.add_command(label="Quitar todos los tachados",
                             command=self._limpiar_tachados)
        hacer.add_cascade(label="Tachados", menu=tachados)
        m.add_cascade(label="Al crear y mover", menu=hacer)

        # -- La ventana
        ventana = tk.Menu(m, tearoff=0)
        for clave, etiqueta in (("encima", "Siempre encima de todo"),
                                ("normal", "Ventana normal"),
                                ("fondo", "Al fondo, sobre el escritorio")):
            ventana.add_command(
                label=marca(self.nivel == clave, etiqueta, True),
                command=lambda c=clave: self._cambiar_nivel(c))
        ventana.add_separator()
        ventana.add_command(
            label=marca(CFG.get("clics_atraviesan", False),
                        "Clics atraviesan (ctrl+alt+b)"),
            command=self._bloquear_clics)
        ventana.add_command(
            label=marca(CFG.get("solo_un_escritorio", False),
                        "Solo en este escritorio virtual"),
            command=self._alternar_un_escritorio)
        ventana.add_command(
            label=marca(CFG.get("semana_blanca", True),
                        "Raya blanca al empezar la semana (lunes)"),
            command=lambda: self._alternar_ajuste("semana_blanca"))
        ventana.add_command(
            label=marca(arranque_activo(), "Arrancar con Windows"),
            command=self._alternar_arranque)
        ventana.add_command(label="Centrar en la pantalla",
                            command=self._centrar)
        ventana.add_command(label="Rescatar (centrar y poner encima)",
                            command=self._rescatar)
        ventana.add_separator()
        ventana.add_command(label="Guardar copia de la configuración",
                            command=self._copia_config)
        ventana.add_command(label="Exportar configuración…",
                            command=self._exportar_config)
        ventana.add_command(label="Importar configuración…",
                            command=self._importar_config)
        m.add_cascade(label="Ventana y copias", menu=ventana)

        m.add_separator()
        m.add_command(label="Ocultar (sigue en la bandeja)",
                      command=lambda: self._alternar_visible(False))
        m.add_command(label="Salir del todo", command=self._salir)
        m.tk_popup(e.x_root, e.y_root)

    def _cambiar_zoom_barra(self, valor):
        CFG["barra_zoom"] = valor
        guardar_config()
        aplicar_escala()
        self._tam = None
        self.dibujar()

    def _cambiar_inicio(self):
        CFG["inicio_ventana"] = "lunes" if CFG.get("inicio_ventana", "hoy") == "hoy" \
            else "hoy"
        guardar_config()
        self.refrescar()
        self.dibujar()

    # -- ciclo de datos ------------------------------------------------------

    def refrescar(self, minutos=None):
        """Descarga ya, y deja programado el siguiente intento (solo uno)."""
        if self._id_refresco:
            try:
                self.after_cancel(self._id_refresco)
            except Exception:
                pass
            self._id_refresco = None
        threading.Thread(target=self._descargar, daemon=True).start()
        espera = minutos or CFG["refresco_minutos"]
        self._id_refresco = self.after(int(espera * 60_000), self.refrescar)
        return "break"

    def _descargar(self):
        try:
            eventos = descargar_eventos()
            repartir_columnas([e for e in eventos if not e.todo_dia])
            self.eventos = eventos
            self._firma_postit = None
            self.actualizado = dt.datetime.now(TZ)
            guardar_cache(eventos)
            self.error_red = ""
        except Exception as exc:
            import traceback
            print("Error al descargar el calendario:", exc)
            log(str(exc) if isinstance(exc, OutlookCerrado) else "ERROR al descargar: "
                + traceback.format_exc().strip().replace("\n", " | "))
            self.error_red = str(exc)[:60] or "sin conexión"
            # sin red se reintenta cada poco, no dentro de media hora
            self.after(0, lambda: self.refrescar(
                minutos=CFG.get("reintento_minutos", 1)))
        self.after(0, self._tras_descargar)
        self.after(50, self._resumen_del_dia)

    def _tras_descargar(self):
        """Ya con los eventos nuevos: las cuentas atrás siguen a los eventos que se hayan movido, y se redibuja."""
        self._sincronizar(reprogramar=False)
        # sin red (o si no ha llegado nada) no se sabe qué se ha borrado: solo se siguen los movidos
        self._seguir_cuentas(rango_descarga() if self.eventos and not self.error_red else None)
        self.dibujar()

    def _seguir_cuentas(self, rango=None):
        """Las cuentas atrás y el evento fijado, al día con self.eventos (tras descargar o tras cambiar un evento aquí)."""
        try:
            if seguir_eventos_movidos(self.eventos, rango):
                self._tam = None             # la barra puede cambiar de alto si aparece o se va una cuenta atrás
        except Exception as exc:
            log(f"No se pudieron actualizar las cuentas atrás: {exc}")

    # ---- Post-It <-> Calendario por una carpeta compartida (sincronizar.py)
    def _carpeta_compartida(self):
        if self._carpeta_sync is None:
            raiz = os.path.dirname(BASE) if os.path.basename(BASE).lower() == "app" else BASE
            self._carpeta_sync = sincronizar.resolver_carpeta(CFG.get("carpeta_compartida") or None, raiz, local=LOCAL)[0] or ""
        return self._carpeta_sync

    def _sincronizar(self, reprogramar=True):
        """Une las notas de Post-It a self.eventos (si cambiaron) y escribe la agenda para Post-It.
        Cada 15 s y sin red: solo lee y escribe unos ficheros pequeños."""
        if reprogramar:
            self.after(15000, self._sincronizar)
        try:
            carpeta = self._carpeta_compartida() if CFG.get("sincronizar_postit", True) else ""
            sin_notas = [e for e in self.eventos if e.categoria != "postit"]
            if not carpeta:
                if len(sin_notas) != len(self.eventos):
                    self.eventos = sin_notas
                    self.dibujar()
                return
            firma = sincronizar.firma_notas(carpeta)
            if firma != self._firma_postit:
                self._firma_postit = firma
                notas = sincronizar.leer_notas(carpeta, TZ, Evento, CFG.get("color_postit"))
                self.eventos = sorted(sin_notas + notas, key=lambda e: e.inicio)
                repartir_columnas([e for e in self.eventos if not e.todo_dia])
                self.dibujar()
            if self.actualizado:     # antes de la primera descarga correcta (Outlook cerrado, OneDrive sin bajar) la agenda no se pisa con casi nada
                sincronizar.escribir_agenda(carpeta, self.eventos + list(self._festivos), dt.datetime.now(TZ))
            sincronizar.escribir_estado(carpeta, {         # el latido: cada equipo dice que está vivo, su versión y de dónde lee
                "app_version": VERSION, "outlook": ESTADO["origen"],
                "ultima_descarga": self.actualizado.replace(tzinfo=None, microsecond=0).isoformat() if self.actualizado else None,
                "eventos": sum(1 for e in self.eventos if e.categoria != "postit")}, dt.datetime.now(TZ))
        except Exception as exc:
            log(f"Sincronización con Post-It: {exc}")

    def _alternar_sync_postit(self):
        CFG["sincronizar_postit"] = not CFG.get("sincronizar_postit", True)
        guardar_config()
        self._firma_postit = None
        self._sincronizar(reprogramar=False)
        return "break"

    def _elegir_carpeta_sync(self):
        from tkinter import filedialog, messagebox
        ruta = filedialog.askdirectory(parent=self, title="Carpeta de sincronización (la misma que Post-It, en OneDrive)",
                                       initialdir=self._carpeta_compartida() or None)
        if ruta:
            # la elección es de ESTE equipo (otro tiene otra ruta de OneDrive): va a sync.json local, no al config.json compartido
            sincronizar.guardar_carpeta_elegida(LOCAL, os.path.normpath(ruta))
            if CFG.pop("carpeta_compartida", None) is not None:
                guardar_config()
            self._carpeta_sync = None
            self._firma_postit = None
            self._sincronizar(reprogramar=False)
            messagebox.showinfo("Carpeta de sincronización",
                                "La agenda y Post-It usan ya esa carpeta.\n\nLa configuración y los perfiles se cargarán de "
                                "«Calendario» dentro de ella la próxima vez que abras el Calendario (si aún no están, se copian los de ahora).",
                                parent=self)

    # ---- ayuda y atajos (todo lo que hay en el widget, en un sitio)
    AYUDA_GENERAL = (
        ("Moverte por el calendario", (
            ("Rueda del ratón", "pasa de día en día (sobre la columna de horas, mueve el horario visible)"),
            ("Mayús + rueda", "salta tantos días como tengas a la vista"),
            ("Ctrl + rueda", "cambia el tamaño de todo el widget"),
            ("Ctrl + Mayús + rueda", "cambia solo el tamaño de la letra"),
            ("Doble clic en el divisor", "devuelve el reparto por defecto entre calendario y lista"),
        )),
        ("Comandos y paneles", (
            ("Ctrl + K  ·  botón ⌘", "paleta de comandos: escribe lo que quieras hacer (cambiar de vista, ir a una fecha como «18/9» o «mañana», abrir un panel, saltar a un evento por su título) y pulsa Intro"),
            ("Botón ⚙  ·  clic derecho › Panel de control", "siete pestañas: Vista, Barra, Avisos, Agenda, Correo, Eventos y Sistema"),
            ("Ctrl + Alt + C", "mostrar u ocultar el widget desde cualquier programa (el atajo se cambia en el config; «null» lo desactiva)"),
            ("Esc", "cierra el panel o el diálogo que tengas abierto"),
            ("Ctrl + Intro", "guarda el evento que estás escribiendo"),
        )),
        ("Eventos", (
            ("Clic derecho en un evento", "editarlo, abrirlo en Outlook, posponerlo, eliminarlo o ponerle una «Alarma antes…» (5 a 60 minutos, o la que quieras; puede haber varias)"),
            ("Pasar el ratón", "enseña el detalle del evento"),
            ("Panel de control › Eventos", "los próximos 30 días con buscador y filtro por calendario, y las listas de solapes y de eventos repetidos"),
        )),
        ("Ventana", (
            ("Clic derecho › Ventana y copias", "siempre encima, normal o al fondo sobre el escritorio, clics que atraviesan (modo fondo) y «Solo en este escritorio virtual»"),
            ("Botón de la cabecera de un panel", "lo cambia entre ventana aparte y anclado a la izquierda o a la derecha del widget"),
            ("Icono de la bandeja", "clic para mostrar u ocultar; clic derecho para cambiar la vista, buscar una fecha, rescatar la ventana si no la encuentras o salir"),
        )),
        ("Datos y varios equipos", (
            ("Carpeta de sincronización", "Panel de control › Sistema › Post-It. Es una carpeta de OneDrive (la misma en cada equipo) donde Calendario y Post-It dejan cada uno sus ficheros; así no hay copias en conflicto. La configuración y los perfiles viven dentro, en «Calendario»"),
            ("Esta configuración es de cada equipo", "posición, tamaño, monitores y atajo no se comparten: mover la ventana no toca la configuración común"),
            ("Perfiles", "botón 👤 o clic derecho › Perfil: cada persona guarda sus datos de cuenta (origen de los eventos, ICS, festivos) y se cambia de uno a otro sin volver a escribirlos"),
            ("Post-It", "las notas unidas (botón 📅 de la nota) mandan aquí sus líneas con hora como eventos, y reciben los eventos de cada fecha"),
        )),
        ("Actualizaciones", (
            ("Aviso al abrir", "si hay una versión nueva en GitHub te lo dice una vez; «Sí» descarga el instalador y lo abre. Tu configuración y tus perfiles se conservan"),
            ("Clic derecho › Buscar actualizaciones…", "lo comprueba cuando quieras"),
        )),
    )

    def _abrir_ayuda_general(self):
        if getattr(self, "_ayuda_general", None) is not None:
            try:
                self._ayuda_general.deiconify()
                self._ayuda_general.lift()
                return
            except tk.TclError:
                self._ayuda_general = None
        v = tk.Toplevel(self)
        self._ayuda_general = v
        v.title("Calendario · ayuda y atajos")
        v.configure(bg=C["fondo"])
        v.geometry("720x560")
        v.bind("<Escape>", lambda e: v.destroy())
        v.protocol("WM_DELETE_WINDOW", v.destroy)
        marco = tk.Frame(v, bg=C["fondo"])
        marco.pack(fill="both", expand=True, padx=10, pady=10)
        barra = tk.Scrollbar(marco)
        barra.pack(side="right", fill="y")
        t = tk.Text(marco, wrap="word", bg=C["fondo"], fg=C["texto"], relief="flat", padx=8, pady=6,
                    font=letra(10), yscrollcommand=barra.set, cursor="arrow")
        t.pack(side="left", fill="both", expand=True)
        barra.config(command=t.yview)
        t.tag_configure("seccion", foreground=C["texto_tenue"], font=letra(9, "bold"), spacing1=14, spacing3=4)
        t.tag_configure("clave", font=letra(10, "bold"))
        t.tag_configure("tenue", foreground=C["texto_tenue"])
        t.insert("end", "Ayuda y atajos  ", "clave")
        t.insert("end", "(Esc para cerrar)\n", "tenue")
        for titulo, filas in self.AYUDA_GENERAL:
            t.insert("end", titulo.upper() + "\n", "seccion")
            for clave, texto in filas:
                t.insert("end", clave, "clave")
                t.insert("end", "  " + texto + "\n")
        t.insert("end", "\nVersión " + VERSION + "  ·  github.com/Kuerb0/Calendario-App\n", "tenue")
        t.config(state="disabled")

    def _comprobar_actualizacion_periodica(self):
        self._buscar_actualizacion(False)
        self.after(6 * 3600 * 1000, self._comprobar_actualizacion_periodica)

    # ---- aviso de versión nueva: se mira la VERSION.txt publicada en GitHub (no se manda ningún dato). «Sí» baja el instalador de
    # esa versión y lo abre; él conserva datos y perfiles. Se pregunta una vez por versión (o siempre que se pida a mano).
    def _buscar_actualizacion(self, a_mano=False):
        import threading

        def trabajo():
            try:
                with requests.get("https://raw.githubusercontent.com/Kuerb0/Calendario-App/main/app/VERSION.txt", timeout=15) as r:
                    r.raise_for_status()
                    nueva = r.text.strip()
                self.after(0, lambda: self._ofrecer_actualizacion(nueva, a_mano))
            except Exception as exc:
                if a_mano:
                    self.after(0, lambda: self._avisar("No se pudo comprobar las actualizaciones.\n\n%s" % exc))

        threading.Thread(target=trabajo, daemon=True).start()

    def _avisar(self, texto):
        from tkinter import messagebox
        messagebox.showinfo("Calendario", texto, parent=self)

    def _ofrecer_actualizacion(self, nueva, a_mano):
        from tkinter import messagebox
        if not sincronizar.version_mayor(nueva, VERSION):
            if a_mano:
                self._avisar("Tienes la última versión (%s)." % VERSION)
            return
        marca = os.path.join(LOCAL, "aviso_version.txt")        # la última versión por la que ya se preguntó (de este equipo)
        try:
            with open(marca, encoding="utf-8") as f:
                ya = f.read().strip()
        except OSError:
            ya = ""
        if not a_mano and ya == nueva:
            return
        try:
            with open(marca, "w", encoding="utf-8") as f:
                f.write(nueva)
        except OSError:
            pass
        if not messagebox.askyesno("Calendario", "Hay una versión nueva del Calendario (%s); tienes la %s.\n\n"
                                   "¿Descargar el instalador y abrirlo ahora? Tu configuración y tus perfiles se conservan."
                                   % (nueva, VERSION), parent=self):
            return
        try:
            nombre = "Calendario %s - Instalador.exe" % nueva
            url = "https://github.com/Kuerb0/Calendario-App/raw/main/instaladores/" + requests.utils.quote(nombre)
            destino = os.path.join(os.environ.get("TEMP", LOCAL), nombre)
            with requests.get(url, timeout=120, stream=True) as r:
                r.raise_for_status()
                with open(destino, "wb") as f:
                    for trozo in r.iter_content(65536):
                        f.write(trozo)
            os.startfile(destino)
        except Exception as exc:
            self._avisar("No se pudo descargar el instalador.\n\n%s" % exc)

    def refrescar_correo(self, minutos=None):
        """Igual que refrescar(), pero para el correo (si está activo)."""
        if self._id_refresco_correo:
            try:
                self.after_cancel(self._id_refresco_correo)
            except Exception:
                pass
            self._id_refresco_correo = None
        if not CFG.get("correo_activo", True):
            return "break"
        self._correo_ahora()
        espera = minutos or CFG.get("correo_refresco_minutos", 3)
        self._id_refresco_correo = self.after(
            int(espera * 60_000), self.refrescar_correo)
        return "break"

    def _correo_ahora(self):
        """Descarga el correo ya mismo, sin tocar el temporizador periódico."""
        threading.Thread(target=self._descargar_correo, daemon=True).start()

    def _descargar_correo(self):
        habia_datos = bool(self.correo_carpetas)
        antes = sum(c.get("no_leidos", 0) for c in self.correo_carpetas)
        try:
            self.correo_carpetas, self.correo_mensajes = descargar_correo()
            self.correo_error = ""
        except Exception as exc:
            import traceback
            log(str(exc) if isinstance(exc, OutlookCerrado) else "ERROR al descargar correo: "
                + traceback.format_exc().strip().replace("\n", " | "))
            self.correo_error = str(exc)[:60] or "sin Outlook"
        self.after(0, self.dibujar)
        panel = getattr(self, "_panel_correo", None)
        if panel is not None:
            self.after(0, lambda: panel.pintar() if panel.winfo_exists()
                       else None)
        if habia_datos and CFG.get("correo_avisos", True):
            despues = sum(c.get("no_leidos", 0) for c in self.correo_carpetas)
            nuevos = despues - antes
            if nuevos > 0:
                self.after(0, lambda: self._avisar_correo_nuevo(nuevos))

    def _avisar_correo_nuevo(self, nuevos):
        texto = f"{nuevos} correo{'s' if nuevos != 1 else ''} nuevo{'s' if nuevos != 1 else ''}"
        notificar("Correo nuevo", texto)
        AvisoFlotante(self, "✉ Correo nuevo", texto, "#3a86ff", 8,
                     al_pulsar=lambda: self._abrir_correo(modo="bandeja"))

    def tic(self):
        """Cada 30 s: repinta la linea de ahora y comprueba avisos."""
        ahora = dt.datetime.now(TZ)
        if config_cambio_fuera() and sincronizar_avisos():
            log("Avisos actualizados desde config.json (otro ordenador)")
        if CFG.get("notificaciones_activas", True):
            for ev in self.eventos:
                if ev.todo_dia or ev.parte:  # los trozos de continuacion no avisan
                    continue
                faltan = (ev.inicio - ahora).total_seconds() / 60
                for umbral in CFG["avisos_minutos"]:
                    clave = (ev.uid, ev.inicio.isoformat(), umbral)
                    if (clave not in self.avisados and 0 < faltan <= umbral
                            and self._debe_avisar(ev)):
                        propio = sonido_de_aviso()
                        notificar(ev.titulo, f"Empieza a las {ev.inicio:%H:%M} "
                                             f"(en {int(faltan)} min)", mudo=propio)
                        self._aviso_flotante(ev, int(faltan))
                        self.avisados.add(clave)
            self._aviso_de_todo_el_dia(ahora)
            self._habitos_tic(ahora)
        self._alarmas_tic(ahora)      # son explícitas: suenan aunque la campana esté apagada
        if len(self.avisados) > 500:
            self.avisados.clear()
        self.dibujar()
        self.after(30_000, self.tic)

    def _debe_avisar(self, ev):
        """Quien avisa y quien no.

        Por defecto NINGUN evento avisa: solo los que hayas marcado a mano.
        Si enciendes 'avisos_eventos' vuelven los avisos de todos, con el
        filtro de etiqueta de siempre. Los tachados nunca avisan.
        """
        if self.esta_hecho(ev):
            return False
        if self._avisa_siempre(ev):
            return True               # marcado a mano: manda sobre todo
        if not CFG.get("avisos_eventos", False):
            return False
        solo = (CFG.get("avisos_solo_etiqueta") or "").strip()
        if not solo:
            return True
        etiquetas = [x.strip().lower()
                     for x in str(ev.categoria or "").split(";")]
        return solo.lower() in etiquetas

    @staticmethod
    def _clave_aviso(ev):
        """Identifica al evento en la lista de «avisarme de este»."""
        return f"{ev.uid}|{ev.inicio:%Y-%m-%d}"

    def _avisa_siempre(self, ev):
        return self._clave_aviso(ev) in set(CFG.get("avisar_de") or [])

    def _alternar_aviso(self, ev):
        """Marca (o desmarca) este evento como «de este sí avísame»."""
        clave = self._clave_aviso(ev)
        lista = [c for c in (CFG.get("avisar_de") or []) if c != clave]
        puesto = len(lista) == len(CFG.get("avisar_de") or [])
        if puesto:
            lista.append(clave)
        CFG["avisar_de"] = lista[-200:]      # no crece sin fin
        guardar_config()
        AvisoFlotante(
            self, ev.titulo[:34],
            "Te avisaré antes de empezar" if puesto else "Ya no te avisaré",
            self._color_evento(ev), 5)
        self.dibujar()
        return "break"

    # -- alarmas con sonido, puestas a mano en cada evento -------------------

    def _alarmas_de(self, ev):
        return sorted(CFG.get("alarmas", {}).get(self._clave_aviso(ev), []))

    def _alternar_alarma(self, ev, minutos):
        """Pone (o quita) una alarma `minutos` antes de este evento."""
        alarmas = CFG.setdefault("alarmas", {})
        clave = self._clave_aviso(ev)
        lista = set(alarmas.get(clave, []))
        puesta = minutos not in lista
        lista.symmetric_difference_update({minutos})
        if lista:
            alarmas[clave] = sorted(lista)
        else:
            alarmas.pop(clave, None)
        guardar_config()
        hora = (ev.inicio - dt.timedelta(minutes=minutos)).strftime("%H:%M")
        AvisoFlotante(self, ev.titulo[:34],
                      f"Sonará a las {hora} ({minutos} min antes)" if puesta
                      else f"Alarma de {minutos} min quitada",
                      self._color_evento(ev), 5)
        return "break"

    def _quitar_alarmas(self, ev):
        CFG.get("alarmas", {}).pop(self._clave_aviso(ev), None)
        guardar_config()

    def _alarma_a_medida(self, ev):
        from tkinter import simpledialog
        n = simpledialog.askinteger(
            "Alarma", "¿Cuántos minutos antes quieres que suene?",
            parent=self, initialvalue=15, minvalue=1, maxvalue=10080)
        if n:
            self._alternar_alarma(ev, n)

    def _menu_alarmas(self, padre, ev):
        sub = tk.Menu(padre, tearoff=0)
        puestas = self._alarmas_de(ev)
        for n in sorted({5, 10, 20, 30, 60} | set(puestas)):
            sub.add_command(
                label=("✓ " if n in puestas else "     ") + f"{n} minutos antes",
                command=lambda n=n: self._alternar_alarma(ev, n))
        sub.add_command(label="Otro número de minutos…",
                        command=lambda: self._alarma_a_medida(ev))
        if puestas:
            sub.add_separator()
            sub.add_command(label="Quitar todas las alarmas",
                            command=lambda: self._quitar_alarmas(ev))
        return sub

    def _alarmas_tic(self, ahora):
        """Suena la alarma que toque y deja el calendario al frente."""
        alarmas = CFG.get("alarmas")
        if not alarmas:
            return
        ayer = (ahora - dt.timedelta(days=1)).strftime("%Y-%m-%d")
        for clave in [c for c in alarmas if c[-10:] < ayer]:   # las ya pasadas
            del alarmas[clave]
        for ev in self.eventos:
            if ev.todo_dia or ev.parte:
                continue
            faltan = (ev.inicio - ahora).total_seconds() / 60
            for minutos in alarmas.get(self._clave_aviso(ev), []):
                marca = ("alarma", self._clave_aviso(ev), minutos)
                if marca not in self.avisados and 0 < faltan <= minutos:
                    self.avisados.add(marca)
                    self._sonar_alarma(ev, int(faltan))

    def _sonar_alarma(self, ev, faltan):
        """Sonido, calendario encima de todo (un minuto) y aviso grande."""
        self._alarma_hasta = time.time() + 90      # ni el «solo un escritorio» la esconde
        sonido_alarma()
        self._alternar_visible(True)
        self.ir_a_fecha(ev.inicio.date())
        self._cambiar_nivel("encima", guardar=False)
        self.after(60_000, lambda: self._cambiar_nivel(
            CFG.get("nivel_ventana", "fondo"), guardar=False))
        detalle = f"{ev.inicio:%H:%M} · empieza en {max(faltan, 1)} min"
        if ev.lugar:
            detalle += f" · {ev.lugar[:34]}"
        AvisoFlotante(self, "⏰ " + ev.titulo, detalle, C["ahora"], 60,
                      al_pulsar=lambda: self._ir_al_evento(ev))
        log(f"Alarma: {ev.titulo} en {faltan} min")

    def _avisar_solo_de(self, etiqueta):
        CFG["avisos_solo_etiqueta"] = etiqueta or ""
        guardar_config()
        log(f"Avisos solo de la etiqueta: {etiqueta or '(todas)'}")

    def _aviso_flotante(self, ev, faltan=None):
        """Recuadro propio en la esquina, sin depender de las notificaciones."""
        if not CFG.get("avisos_flotantes", True):
            return None
        if faltan is None:
            faltan = max(0, int((ev.inicio - dt.datetime.now(TZ)).total_seconds() / 60))
        detalle = f"{ev.inicio:%H:%M} · empieza en {max(faltan, 0)} min"
        if ev.lugar:
            detalle += f" · {ev.lugar[:34]}"
        try:
            return AvisoFlotante(self, ev.titulo, detalle, self._color_evento(ev),
                                 CFG.get("aviso_flotante_segundos", 12),
                                 al_pulsar=lambda: self._ir_al_evento(ev))
        except Exception as exc:
            log(f"No se pudo mostrar el aviso flotante: {exc}")
            return None

    def _ir_al_evento(self, ev):
        """Trae la ventana al frente y salta al dia del evento."""
        self._alternar_visible(True)
        self.ir_a_fecha(ev.inicio.date())

    def _menu_avisos(self, padre):
        """Construye (y devuelve) el menú de configuración de avisos, para
        usarlo tanto desde la campana de la barra como desde el menú grande."""
        avisos = tk.Menu(padre, tearoff=0)
        activas = CFG.get("notificaciones_activas", True)
        avisos.add_command(
            label=("✓ " if activas else "     ") + "Notificaciones activadas",
            command=self._alternar_notificaciones)
        avisos.add_separator()
        for clave, etiqueta in (
                ("avisos_flotantes", "Aviso propio en la esquina"),
                ("barra_proximos", "«Ahora» y «luego» abajo"),
                ("linea_ahora_lista", "Raya de la hora en la lista")):
            avisos.add_command(
                label=("✓ " if CFG.get(clave, True) else "     ") + etiqueta,
                command=lambda c=clave: self._alternar_ajuste(c))
        avisos.add_separator()
        avisos.add_command(label="— Antes de cada evento —", state="disabled")
        avisos.add_command(
            label=("✓ " if CFG.get("avisos_eventos", False) else "     ")
                  + "Avisar de todos los eventos",
            command=lambda: self._alternar_ajuste("avisos_eventos"))
        marcados = len(CFG.get("avisar_de") or [])
        avisos.add_command(
            label=f"(apagado: solo avisan los marcados a mano · {marcados})",
            state="disabled")
        solo = tk.Menu(avisos, tearoff=0)
        actual = (CFG.get("avisos_solo_etiqueta") or "").strip()
        solo.add_command(label=("● " if not actual else "     ")
                               + "Avisar de todos",
                         command=lambda: self._avisar_solo_de(""))
        for nombre in self.etiquetas_usadas():
            solo.add_command(
                label=("● " if actual == nombre else "     ") + f"Solo «{nombre}»",
                command=lambda n=nombre: self._avisar_solo_de(n))
        if actual and actual not in self.etiquetas_usadas():
            solo.add_command(label=f"● Solo «{actual}»", state="disabled")
        avisos.add_cascade(label="De qué eventos avisar", menu=solo)
        if marcados:
            avisos.add_command(label="Olvidar los eventos marcados",
                               command=self._limpiar_avisos_marcados)
        avisos.add_command(label="(los tachados nunca avisan)", state="disabled")

        avisos.add_separator()
        hora = self._hora_recordatorio()
        avisos.add_command(label=f"— Recordatorio de las {hora}:00 —",
                           state="disabled")
        avisos.add_command(
            label=("✓ " if CFG.get("recordar_habitos", True) else "     ")
            + "Marcar los hábitos",
            command=lambda: self._alternar_ajuste("recordar_habitos"))
        horas = tk.Menu(avisos, tearoff=0)
        for h in (18, 19, 20, 21, 22, 23):
            horas.add_command(
                label=("● " if hora == h else "     ") + f"{h}:00",
                command=lambda v=h: self._cambiar_hora_recordatorio(v))
        avisos.add_cascade(label="A qué hora", menu=horas)

        avisos.add_separator()
        avisos.add_cascade(label="Sonido de los avisos",
                           menu=self._menu_sonidos(avisos))
        avisos.add_command(label="Ver cómo queda el aviso",
                           command=self._probar_aviso)
        return avisos

    def _menu_sonidos(self, padre):
        """Sonido de las alarmas y de los avisos antes de cada evento."""
        actual = CFG.get("sonido_aviso", "")
        menu = tk.Menu(padre, tearoff=0)

        def marca(activo, texto):
            return ("● " if activo else "     ") + texto

        menu.add_command(label=marca(actual == "", "El de siempre"),
                         command=lambda: self._elegir_sonido(""))
        menu.add_command(label=marca(actual == "ninguno", "Sin sonido"),
                         command=lambda: self._elegir_sonido("ninguno"))
        grupos = (("Alarmas", [f"Alarm{n:02d}.wav" for n in range(1, 11)]),
                  ("Tonos de llamada", [f"Ring{n:02d}.wav" for n in range(1, 11)]),
                  ("Notificaciones", ["Windows Notify Calendar.wav",
                                      "Windows Notify.wav", "Windows Ding.wav",
                                      "Windows Message Nudge.wav"]))
        for titulo, nombres in grupos:
            sub = tk.Menu(menu, tearoff=0)
            for nombre in nombres:
                ruta = os.path.join(CARPETA_SONIDOS, nombre)
                if os.path.isfile(ruta):
                    sub.add_command(
                        label=marca(actual == ruta, nombre[:-4]),
                        command=lambda r=ruta: self._elegir_sonido(r))
            menu.add_cascade(label=titulo, menu=sub)
        menu.add_separator()
        propio = actual and actual != "ninguno" and not actual.startswith(CARPETA_SONIDOS)
        menu.add_command(
            label=marca(bool(propio), "Elegir un archivo .wav…"
                        + (f" ({os.path.basename(actual)})" if propio else "")),
            command=self._elegir_wav)
        return menu

    def _elegir_sonido(self, ruta):
        CFG["sonido_aviso"] = ruta
        guardar_config()
        if ruta and ruta != "ninguno":
            reproducir_sonido(ruta)          # para oírlo al elegirlo

    def _elegir_wav(self):
        from tkinter import filedialog
        ruta = filedialog.askopenfilename(
            parent=self, title="Sonido de los avisos",
            filetypes=[("Sonido WAV", "*.wav")])
        if ruta:
            self._elegir_sonido(os.path.normpath(ruta))

    def _cambiar_hora_recordatorio(self, hora):
        CFG["hora_recordatorio"] = int(hora)
        self._habitos_avisado = None
        guardar_config()
        log(f"Recordatorio diario a las {hora}:00")

    def _limpiar_avisos_marcados(self):
        CFG["avisar_de"] = []
        guardar_config()
        log("Lista de eventos con aviso vaciada")
        self.dibujar()

    def _alternar_notificaciones(self, e=None):
        """La campana: apaga o enciende TODOS los avisos de golpe."""
        CFG["notificaciones_activas"] = not CFG.get("notificaciones_activas", True)
        guardar_config()
        self.dibujar()
        return "break"

    def _menu_campana(self, e=None):
        """Clic derecho sobre la campana: el menú completo de avisos."""
        self._abrir_menu(self._menu_avisos(self), e)
        return "break"

    def _probar_aviso(self):
        """Enseña como se ve el aviso flotante, con el proximo evento."""
        sonido_de_aviso()
        _, siguiente = self._ahora_y_siguiente()
        if siguiente is None:
            AvisoFlotante(self, "Aviso de prueba",
                          "Así se verán los avisos antes de cada evento",
                          C["evento"], CFG.get("aviso_flotante_segundos", 12))
        else:
            self._aviso_flotante(siguiente)

    def _aviso_de_todo_el_dia(self, ahora):
        """Un unico aviso al llegar la hora configurada, con lo del dia entero."""
        cuando = CFG.get("aviso_todo_dia")
        if not cuando:
            return
        try:
            hora, minuto = (int(x) for x in str(cuando).split(":"))
        except ValueError:
            return
        clave = ("todo-el-dia", ahora.date().isoformat())
        if clave in self.avisados or (ahora.hour, ahora.minute) < (hora, minuto):
            return
        self.avisados.add(clave)
        hoy = ahora.date()
        titulos = [ev.titulo for ev in self._mostrables()
                   if ev.todo_dia and ev.inicio.date() <= hoy < ev.fin.date()]
        if titulos:
            notificar("Hoy", " · ".join(titulos[:4]))

    def _resumen_del_dia(self):
        """Al arrancar: cuantas cosas hay hoy y cuales son las primeras."""
        if self._resumen_hecho or not CFG.get("resumen_al_arrancar", True):
            return
        self._resumen_hecho = True
        ahora = dt.datetime.now(TZ)
        del_dia = [ev for ev in self._eventos_del_dia(ahora.date())
                   if ev.todo_dia or ev.fin >= ahora]
        if not del_dia:
            notificar("Hoy no te queda nada", "Día libre en el calendario")
            return
        partes = []
        for ev in del_dia[:4]:
            partes.append(ev.titulo if ev.todo_dia
                          else f"{ev.inicio:%H:%M} {ev.titulo}")
        resto = f"  (+{len(del_dia) - 4})" if len(del_dia) > 4 else ""
        notificar(f"Hoy: {len(del_dia)} cosas", " · ".join(partes) + resto)

    # -- pintado -------------------------------------------------------------

    def _alto_linea(self, fuente):
        """Alto de una linea de texto, cacheado por fuente."""
        if fuente not in self._altos_linea:
            import tkinter.font as tkfont
            self._altos_linea[fuente] = tkfont.Font(font=fuente).metrics("linespace")
        return self._altos_linea[fuente]

    def _ancho_texto(self, texto, fuente=None):
        """Ancho en pixeles de un texto, cacheando la medida.

        Cachear solo el objeto Font no bastaba: cada .measure() es un viaje de
        ida y vuelta a Tcl, y al calcular altos de fila se miden las mismas
        palabras cientos de veces por repintado. La clave lleva la fuente
        dentro, asi que al cambiar de tamaño o de letra no hay medidas viejas.
        """
        fuente = fuente or letra(8)
        clave = (fuente, texto)
        medida = self._anchos.get(clave)
        if medida is not None:
            return medida
        if fuente not in self._fuentes:
            import tkinter.font as tkfont
            self._fuentes[fuente] = tkfont.Font(font=fuente)
        medida = self._fuentes[fuente].measure(texto)
        if len(self._anchos) > 20_000:      # no crece sin fin
            self._anchos.clear()
        self._anchos[clave] = medida
        return medida

    def _texto_en_caja(self, x, y, ancho, alto, texto, font=None, **kw):
        """Escribe el texto recortandolo con puntos suspensivos si no cabe.

        Antes se tanteaba: se dibujaba, se medía con bbox(), se acortaba y
        vuelta a empezar, hasta 24 veces por texto. Cada bbox() obliga a Tk a
        recalcular la caja, asi que con muchos eventos se iba el tiempo ahi.
        Ahora se calcula cuantas lineas caben con las medidas ya cacheadas y
        se dibuja una sola vez.
        """
        c = self.lienzo
        lh = self._alto_linea(font)
        caben = max(1, int((alto + lh * 0.35) // lh))
        limite = caben * lh + 2
        util = max(ancho, 12)

        def encaja(item):
            """Ultima palabra: la que dice Tk, no la estimacion."""
            caja = c.bbox(item)
            return not caja or caja[3] - caja[1] <= limite

        if self._lineas_envueltas(texto, util, font) <= caben:
            item = c.create_text(x, y, text=texto, width=util, anchor="nw",
                                 font=font, **kw)
        else:
            # no cabe: biseccion para el trozo mas largo que si quepa
            bajo, tope, mejor = 0, len(texto), ""
            while bajo <= tope:
                medio = (bajo + tope) // 2
                trozo = texto[:medio].rstrip()
                if self._lineas_envueltas(trozo + "…", util, font) <= caben:
                    mejor, bajo = trozo, medio + 1
                else:
                    tope = medio - 1
            item = c.create_text(x, y, width=util, anchor="nw", font=font,
                                 text=(mejor.rstrip() + "…") if mejor else "…",
                                 **kw)
        # red de seguridad: con palabras larguisimas sin espacios Tk parte de
        # otra forma y la estimacion se queda corta. Se comprueba una sola vez
        # (antes eran hasta 24 bbox por texto) y se recorta si hiciera falta.
        recorte = c.itemcget(item, "text")
        for _ in range(6):
            if encaja(item) or len(recorte) <= 2:
                break
            recorte = recorte.rstrip("…")
            recorte = recorte[:max(1, int(len(recorte) * 0.8))].rstrip()
            c.itemconfig(item, text=recorte + "…")
        return item

    def _medir_especial(self, icono, texto, ancho_max):
        """Cuanto alto ocupara la cajita de un festivo/cumpleaños con este
        ancho disponible (1 o 2 lineas), sin llegar a dibujarla."""
        fuente = letra(8, "bold")
        etiqueta = f"{icono} {texto}"
        ancho_txto = self._ancho_texto(etiqueta, fuente)
        ancho_util = max(ancho_max, 20)
        if ancho_txto + 10 <= ancho_util:
            lineas = 1
        else:
            lineas = min(2, self._lineas_envueltas(etiqueta, ancho_util - 8, fuente))
        return lineas * self._alto_linea(fuente) + 6

    def _dibujar_especial(self, x, y, ancho_max, icono, texto, color, ev=None):
        """Cajita de festivo/cumpleaños: icono + nombre. Usa el ancho
        disponible y, si el nombre no cabe en una linea, se envuelve a una
        segunda linea (en vez de recortarlo enseguida con puntos suspensivos).
        Devuelve (ancho_usado, alto_usado) para que quien la coloque sepa
        cuanto sitio ha ocupado de verdad."""
        c = self.lienzo
        fuente = letra(8, "bold")
        etiqueta = f"{icono} {texto}"
        ancho_txto = self._ancho_texto(etiqueta, fuente)
        ancho_util = max(ancho_max, 20)
        ancho = min(ancho_txto + 10, ancho_util)
        alto = self._medir_especial(icono, texto, ancho_max)
        if ancho_txto + 10 > ancho_util:
            ancho = ancho_util
        marca = ("evento", self._registrar(ev)) if ev is not None else ()
        c.create_rectangle(x, y, x + ancho, y + alto - 2, fill=color, outline="",
                           tags=marca)
        self._texto_en_caja(x + 4, y + 1, ancho - 8, alto - 4,
                            etiqueta, font=fuente, fill=C["evento_texto"],
                            tags=marca)
        return ancho, alto

    # -- botones de la barra superior ---------------------------------------

    # Glifos de «Segoe MDL2 Assets» (viene con Windows 10/11). Si la fuente no
    # está, se usa el texto de siempre.
    GLIFOS_MDL2 = {
        "btn_panel": "\uE713", "btn_correo_col": "\uE715",
        "btn_avisos": "\uEA8F", "btn_todos": "\uE73E",
        "btn_paleta": "\uE721", "btn_horario": "\uE121",
        "btn_ir": "\uE787", "btn_vista": "\uE700",
    }

    def _tiene_mdl2(self):
        if not hasattr(self, "_mdl2"):
            try:
                import tkinter.font as tkfont
                self._mdl2 = "Segoe MDL2 Assets" in tkfont.families(self)
            except Exception:
                self._mdl2 = False
        return self._mdl2

    def _dibujar_icono(self, nombre, cx, cy, tam, tinta, tag):
        """Iconos vectoriales: se ven nítidos a cualquier tamaño."""
        c = self.lienzo
        g = max(2, round(1.7 * BZ))          # grosor del trazo
        k = tam * 0.5
        kw = dict(fill=tinta, width=g, capstyle="round", joinstyle="round",
                  tags=(tag,))
        if nombre == "cerrar":
            c.create_line(cx - k, cy - k, cx + k, cy + k, **kw)
            c.create_line(cx - k, cy + k, cx + k, cy - k, **kw)
        elif nombre == "izq":
            c.create_line(cx + k * 0.5, cy - k, cx - k * 0.5, cy,
                          cx + k * 0.5, cy + k, **kw)
        elif nombre == "der":
            c.create_line(cx - k * 0.5, cy - k, cx + k * 0.5, cy,
                          cx - k * 0.5, cy + k, **kw)
        elif nombre == "mas":
            c.create_line(cx - k, cy, cx + k, cy, **kw)
            c.create_line(cx, cy - k, cx, cy + k, **kw)
        elif nombre == "ocultar":
            c.create_line(cx - k, cy, cx + k, cy, **kw)
        elif nombre == "refrescar":
            r = k
            c.create_arc(cx - r, cy - r, cx + r, cy + r, start=40, extent=280,
                         style="arc", outline=tinta, width=g, tags=(tag,))
            th = math.radians(320)
            px, py = cx + r * math.cos(th), cy - r * math.sin(th)
            tx, ty = -math.sin(th), -math.cos(th)        # sentido del giro
            nx, ny = ty, -tx
            a = r * 0.85
            c.create_polygon(px + tx * a, py + ty * a,
                             px + nx * a * 0.8, py + ny * a * 0.8,
                             px - nx * a * 0.8, py - ny * a * 0.8,
                             fill=tinta, outline="", tags=(tag,))

    def _btn(self, x, y, ancho, texto, tag, activo=False, alto=None, icono=None,
             fuente=None):
        """Dibuja un boton y devuelve la x donde acaba."""
        c = self.lienzo
        t = ESCALA * max(1.0, LETRA)
        alto = alto or round(16 * t * BZ)
        fondo = C["evento"] if activo else C.get("boton", "#2f2f33")
        tinta = C["evento_texto"] if activo else C["texto"]
        c.create_rectangle(x, y, x + ancho, y + alto, fill=fondo, outline="",
                           tags=(tag,))
        cx, cy = x + ancho / 2, y + alto / 2
        glifo = self.GLIFOS_MDL2.get(tag)
        if tag == "btn_avisos" and texto == "🔕":
            glifo = "\uE74F"
        if icono:
            self._dibujar_icono(icono, cx, cy, min(alto, ancho) * 0.34,
                                tinta, tag)
        elif glifo and self._tiene_mdl2():
            c.create_text(cx, cy, text=glifo, fill=tinta, tags=(tag,),
                          font=("Segoe MDL2 Assets",
                                max(8, round(9 * t * BZ))))
        else:
            c.create_text(cx, cy, text=texto, fill=tinta,
                          font=fuente or letra(8 * BZ, "bold"), tags=(tag,))
        self._btn_activo[tag] = activo
        return x + ancho

    def _vincular_ayuda(self, tag, texto):
        """Explicación breve al dejar el ratón quieto sobre un botón."""
        c = self.lienzo
        c.tag_bind(tag, "<Enter>",
                  lambda e, t=tag, x=texto: self._programar_ayuda(t, x, e),
                  add="+")
        c.tag_bind(tag, "<Motion>",
                  lambda e, t=tag, x=texto: self._programar_ayuda(t, x, e),
                  add="+")
        c.tag_bind(tag, "<Leave>", lambda e, t=tag: self._cancelar_ayuda(t),
                  add="+")
        c.tag_bind(tag, "<Button-1>", lambda e, t=tag: self._cancelar_ayuda(t),
                  add="+")

    def _programar_ayuda(self, tag, texto, e):
        if self._ayuda_tag != tag:
            self._cancelar_ayuda(self._ayuda_tag)
        self._ayuda_tag = tag
        self._ayuda_pos = (e.x_root, e.y_root)
        if self._ayuda_win is None and self._ayuda_id is None:
            self._ayuda_id = self.after(550, lambda: self._mostrar_ayuda(texto))

    def _mostrar_ayuda(self, texto):
        self._ayuda_id = None
        if self._ayuda_win is not None:
            try:
                self._ayuda_win.destroy()
            except tk.TclError:
                pass
        win = tk.Toplevel(self)
        win.overrideredirect(True)
        try:
            win.attributes("-topmost", True)
        except tk.TclError:
            pass
        tk.Label(win, text=texto, bg="#1c1c1f", fg="#f2f2f2",
                font=letra(8), padx=7, pady=3, relief="solid",
                borderwidth=1).pack()
        x, y = getattr(self, "_ayuda_pos", (self.winfo_pointerx(),
                                            self.winfo_pointery()))
        win.geometry(f"+{x + 14}+{y + 18}")
        self._ayuda_win = win

    def _cancelar_ayuda(self, tag):
        if self._ayuda_tag != tag and tag is not None:
            return
        self._ayuda_tag = None
        if self._ayuda_id is not None:
            try:
                self.after_cancel(self._ayuda_id)
            except tk.TclError:
                pass
            self._ayuda_id = None
        if self._ayuda_win is not None:
            try:
                self._ayuda_win.destroy()
            except tk.TclError:
                pass
            self._ayuda_win = None

    def _resaltar(self, tag, dentro):
        """Aclara el boton al pasar el raton por encima."""
        self.lienzo.config(cursor="hand2" if dentro else "")
        if self._btn_activo.get(tag):
            return
        items = self.lienzo.find_withtag(tag)
        if items:
            self.lienzo.itemconfig(
                items[0], fill=C.get("boton_raton", "#40404a") if dentro
                else C.get("boton", "#2f2f33"))

    def _dibujar_barra(self, ancho, titulo=None):
        c = self.lienzo
        c.create_rectangle(0, 0, ancho, BARRA - 2, fill=C["fondo"], outline="")

        m = ESCALA * max(1.0, LETRA) * BZ
        y_b = max(2, (BARRA - 2 - round(16 * m)) // 2)
        # ✕ arriba a la izquierda: cierra el widget del todo (el «—» de la
        # derecha solo lo oculta)
        x = self._btn(round(6 * m), y_b, round(18 * m), "✕", "btn_cerrar",
                      icono="cerrar")
        x = self._btn(x + round(8 * m), y_b, round(16 * m), "‹", "nav_ant", icono="izq")
        x = self._btn(x + 2, y_b, round(16 * m), "›", "nav_sig", icono="der")
        # «hoy» siempre ocupa su sitio (resaltado solo si te has desplazado):
        # así nada de lo que hay a su derecha se mueve
        x = self._btn(x + 4, y_b, round(30 * m), "hoy", "nav_hoy",
                      activo=bool(self.desplazamiento))
        alto_chip = round(16 * m)
        x = self._pintar_cuentas(x + round(8 * m), y_b, ancho - 8 - round(270 * m))
        x_izq = x

        algun_festivo = any(f.get("activo", True)
                            for f in CFG["festivos"].values())
        notif_activas = CFG.get("notificaciones_activas", True)
        # el numero de dias a la vista, tal cual esté ahora mismo
        etiqueta_dias = f"{CFG['dias_visibles']} d ▾"
        ancho_dias = 30 + 10 * 6             # cabe hasta «45 d ▾», en letra grande
        # (prioridad, botones). Prioridad: número alto = se quita antes si no
        # cabe; 0 = no se quita nunca. Los días no se quitan: sin ese botón no
        # se podría volver desde la vista de 1 día, que es muy estrecha.
        catalogo = [
            (0, [(etiqueta_dias, "btn_dias", False,
                  ancho_dias, None)]),
            (3, [("H", "btn_horizontal", self.orientacion == "horizontal",
                  18, None),
                 ("V", "btn_vertical", self.orientacion == "vertical",
                  18, None)]),
            (3, [(NIVELES.get(self.nivel, "Fondo"), "btn_nivel",
                  self.nivel == "encima", 46, None)]),
            (2, [("🔔" if notif_activas else "🔕", "btn_avisos",
                  not notif_activas, 26, None)]),
            (0, [("⏳", "btn_cuenta", bool(cuentas_vigentes()), 26, None)]),   # prioridad 0: nunca pasa al menú ⋯
            (0, [("", "btn_ocultar", False, 22, "ocultar")]),
        ]
        # botones que el usuario ha quitado (Panel de control > Barra)
        quitados = set(CFG.get("barra_ocultos") or [])
        fijos = ("btn_dias", "btn_ocultar")
        catalogo = [(pr, g) for pr, g in catalogo
                    if g[0][1] in fijos or
                    ("btn_orient" if g[0][1] == "btn_horizontal"
                     else g[0][1]) not in quitados]

        grupos = [(pr, [(t, tg, ac, round(w * m), ic) for t, tg, ac, w, ic in g])
                  for pr, g in catalogo]

        def medir(gs):
            n = sum(len(g) for g in gs)
            return (sum(b[3] for g in gs for b in g)
                    + 2 * (n - len(gs))        # separacion dentro del grupo
                    + 8 * (len(gs) - 1))       # separacion entre grupos

        # si la ventana es estrecha se quitan los menos importantes
        # (siguen estando en el menu del clic derecho)
        # (lo que se quita no desaparece: pasa al menú del botón «⋯»)
        self._barra_desbordados = []
        mas = (0, [("⋯", "btn_mas", False, round(24 * m), None)])
        while len(grupos) > 2:
            if ancho - 8 - medir([g for _, g in grupos]) >= x_izq + 10:
                break
            peor = max(range(len(grupos)), key=lambda i: grupos[i][0])
            if grupos[peor][0] == 0:
                break
            self._barra_desbordados.extend(
                (t, tg) for t, tg, *_ in grupos[peor][1])
            grupos.pop(peor)
            if mas not in grupos:
                # va justo antes de ocultar (el último)
                grupos.insert(len(grupos) - 1, mas)

        grupos = [g for _, g in grupos]
        x = x_der = ancho - 8 - medir(grupos)
        for grupo in grupos:
            for texto, tag, activo, w, icono in grupo:
                grande = ({"alto": round(22 * m),
                           "fuente": letra(11 * BZ, "bold")}
                          if tag == "btn_dias" else {})
                x = self._btn(x, y_b - (grande["alto"] - round(16 * m)) // 2
                              if grande else y_b, w, texto, tag, activo,
                              icono=icono, **grande) + 2
            x += 6

        if titulo and x_der - x_izq > 96:
            c.create_text((x_izq + x_der) / 2, BARRA / 2, text=titulo,
                          fill=C["texto"],
                          font=letra(9 * max(1.0, BZ * 0.8), "bold"))

    def _dibujar_agarradera(self, ancho, alto):
        """Rayitas en la esquina: aviso de que se puede redimensionar."""
        c = self.lienzo
        for i in (0, 4, 8):
            c.create_line(ancho - 4 - i, alto - 3, ancho - 3, alto - 4 - i,
                          fill=C["texto_tenue"])

    def _chip_estado(self, x, y, texto, fondo, tag, tinta=None):
        """Etiqueta de color en la franja de abajo; devuelve donde acaba."""
        c = self.lienzo
        item = c.create_text(x + 5, y, text=texto, anchor="nw", tags=tag,
                             fill=tinta or C["evento_texto"],
                             font=letra(7, "bold"))
        x1, y1, x2, y2 = c.bbox(item)
        caja = c.create_rectangle(x, y1 - 1, x2 + 5, y2 + 1, fill=fondo,
                                  outline="", tags=tag)
        c.tag_lower(caja, item)
        return x2 + 5

    def _pintar_proximos(self, x, y, ancho):
        """«Ahora» y «luego» con su cuenta atras, al lado del estado."""
        if not CFG.get("barra_proximos", True):
            return x
        c = self.lienzo
        en_curso, siguiente = self._ahora_y_siguiente()
        for ev, tag in ((en_curso, "ev_ahora"), (siguiente, "ev_siguiente")):
            if ev is None or (ancho and x > ancho - round(120 * ESCALA)):
                continue
            titulo = ev.titulo[:22]
            if ev is en_curso:
                x = self._chip_estado(x, y, f"▶ {titulo} · hasta {ev.fin:%H:%M}",
                                      self._color_evento(ev), tag) + 10
            else:
                item = c.create_text(x, y, anchor="nw", tags=tag,
                                     text=f"→ {titulo} · {cuenta_atras(ev.inicio)}",
                                     fill=C["texto_tenue"], font=letra(7))
                x = c.bbox(item)[2] + 12
            c.tag_bind(tag, "<Button-1>",
                       lambda e, v=ev: self._pintar_detalle(v, e.x, e.y - 40))
            c.tag_bind(tag, "<Enter>", lambda e: c.config(cursor="hand2"))
            c.tag_bind(tag, "<Leave>", lambda e: c.config(cursor=""))
        return x

    def _w_chip_habitos(self, fuente):
        return (self._ancho_texto("Hábitos 88/88", fuente)
                + round(16 * ESCALA * BZ))

    def _w_chip_correo(self, fuente):
        return (self._ancho_texto("✉ 99+", fuente)
                + round(16 * ESCALA * BZ))

    def _pintar_cuentas(self, x, y, limite):
        """Chips de cuenta atrás de dos líneas: nombre y días que faltan
        (rojo 0-5 días, ámbar 6-10, verde más)."""
        import extras
        c = self.lienzo
        cuentas = cuentas_vigentes()
        hoy = dt.date.today()
        alto = BARRA - 6
        y0 = 2
        f_nombre, f_dias = letra(7, "bold"), letra(10, "bold")
        for i, cu in enumerate(cuentas):
            n = extras.dias_hasta(dt.date.fromisoformat(cu["fecha"]), hoy)
            nombre = cu["nombre"][:18]
            dias = "hoy" if n == 0 else f"{n} d"
            ancho = max(self._ancho_texto(nombre, f_nombre),
                        self._ancho_texto(dias, f_dias)) + 14
            fondo, tinta = extras.color_cuenta(n), "#101010"
            if x + ancho > limite:
                nombre, dias = "más", f"+{len(cuentas) - i}"
                fondo, tinta = C.get("boton", "#2f2f33"), C["texto"]
                ancho = max(self._ancho_texto(nombre, f_nombre),
                            self._ancho_texto(dias, f_dias)) + 14
            c.create_rectangle(x, y0, x + ancho, y0 + alto, fill=fondo,
                               outline="", tags="chip_cuenta")
            c.create_text(x + ancho / 2, y0 + alto * 0.28, text=nombre,
                          fill=tinta, font=f_nombre, tags="chip_cuenta")
            c.create_text(x + ancho / 2, y0 + alto * 0.70, text=dias,
                          fill=tinta, font=f_dias, tags="chip_cuenta")
            x += ancho + 4
            if nombre == "más":
                break
        return x

    def _clave_cuenta(self, ev):
        return (ev.uid or ev.titulo, ev.inicio.date().isoformat())

    def _cuenta_de(self, ev):
        for c in CFG.get("cuentas_atras") or []:
            if (c.get("uid"), c.get("fecha")) == self._clave_cuenta(ev):
                return c
        return None

    def _alternar_cuenta_evento(self, ev):
        c = self._cuenta_de(ev)
        if c:
            CFG["cuentas_atras"].remove(c)
        else:
            CFG.setdefault("cuentas_atras", []).append(
                {"nombre": ev.titulo, "fecha": ev.inicio.date().isoformat(),
                 "uid": ev.uid or ev.titulo, "hasta": ev.fin.isoformat()})
        guardar_config()
        self._rehacer_barra()

    def _rehacer_barra(self):
        aplicar_escala()
        self._tam = None
        self.dibujar()

    def _nueva_cuenta(self):
        try:
            from tkinter import simpledialog
            nombre = simpledialog.askstring("Cuenta atrás", "¿Qué evento?",
                                            parent=self)
            if not nombre:
                return
            import extras
            fecha = extras.leer_fecha(simpledialog.askstring(
                "Cuenta atrás", "¿Qué día? (DD/MM/AAAA)", parent=self) or "")
        except ValueError:
            return
        except Exception as exc:
            log(f"No se pudo pedir la cuenta atrás: {exc}")
            return
        CFG.setdefault("cuentas_atras", []).append(
            {"nombre": nombre.strip(), "fecha": fecha.isoformat()})
        guardar_config()
        self._rehacer_barra()

    def _menu_cuentas(self, e=None):
        """Cuenta atrás: ver, quitar y añadir eventos del calendario."""
        import extras
        m = tk.Menu(self, tearoff=0)
        hoy = dt.date.today()
        vigentes = cuentas_vigentes()
        for c in vigentes:
            n = extras.dias_hasta(dt.date.fromisoformat(c["fecha"]), hoy)
            m.add_command(
                label=f"{'hoy' if n == 0 else f'{n} d'} · {c['nombre'][:40]}   (quitar)",
                command=lambda c=c: (CFG["cuentas_atras"].remove(c),
                                     guardar_config(), self._rehacer_barra()))
        if vigentes:
            m.add_separator()
        proximos = tk.Menu(m, tearoff=0)
        ahora = dt.datetime.now(TZ)
        visto = 0
        for ev in sorted(list(self.eventos) + list(self._festivos),
                         key=lambda x: x.inicio):
            if ev.fin <= ahora or ev.parte or self._cuenta_de(ev):
                continue
            proximos.add_command(
                label=f"{ev.inicio:%d/%m} · {ev.titulo[:40]}",
                command=lambda ev=ev: self._alternar_cuenta_evento(ev))
            visto += 1
            if visto >= 25:
                break
        m.add_cascade(label="Añadir un evento del calendario",
                      menu=proximos, state="normal" if visto else "disabled")
        m.add_command(label="Nueva fecha…", command=self._nueva_cuenta)
        self._abrir_menu(m, e)
        return "break"

    def _pintar_habitos(self, x, y, alto):
        """Chip «Hábitos 2/4» en la barra de arriba. Devuelve donde acaba."""
        info = habitos_estado(self.eventos)
        if not info["activo"]:
            return x
        sin_marcar = info["total"] - info["hechos"] - info["fallados"]
        if not info["total"]:
            # todavia no hay ninguno: el chip es la puerta para crear el primero
            texto = "Hábitos +"
            fondo, tinta = C.get("boton", "#2f2f33"), C["texto"]
        else:
            texto = f"Hábitos {info['hechos']}/{info['total']}"
            if info["fallados"]:
                texto += f" · {info['fallados']}✕"
            if info["vencidos"]:
                texto += f" · {info['vencidos']}{MARCA_PENDIENTE}"
            if info["al_dia"]:
                # nada fallado ni pasado de hora: verde aunque queden huecos,
                # porque un hueco es «aún no toca», no un descuido
                if not sin_marcar and info["racha"] > 1:
                    texto += f" · {info['racha']}"
                fondo, tinta = "#00d98b", "#06281c"
            elif info["fallados"] and not info["vencidos"]:
                fondo, tinta = "#ff2d55", "#ffffff"   # fallado a conciencia
            else:
                fondo, tinta = "#ffb020", "#2b1c00"   # se te ha pasado la hora

        c = self.lienzo
        fuente = letra(8 * BZ, "bold")
        pad = round(16 * ESCALA * BZ)
        ancho = self._w_chip_habitos(fuente)
        # ancho fijo: si el texto no cabe se le quitan los detalles del final,
        # así el chip no cambia de tamaño (ni empuja a los botones) con cada dato
        while self._ancho_texto(texto, fuente) + pad > ancho and " · " in texto:
            texto = texto.rsplit(" · ", 1)[0]
        c.create_rectangle(x, y, x + ancho, y + alto, fill=fondo, outline="",
                           tags=("habitos",))
        c.create_text(x + ancho / 2, y + alto / 2, text=texto, fill=tinta,
                      font=fuente, tags=("habitos",))
        # asi _resaltar no le pisa el color al pasar el raton por encima
        self._btn_activo["habitos"] = True
        c.tag_bind("habitos", "<Button-1>", lambda e: self._abrir_habitos())
        c.tag_bind("habitos", "<Button-3>", lambda e: self._menu_habitos_rapido(e))
        c.tag_bind("habitos", "<Enter>", lambda e: c.config(cursor="hand2"))
        c.tag_bind("habitos", "<Leave>", lambda e: c.config(cursor=""))
        return x + ancho

    def _menu_habitos_rapido(self, e=None):
        """Clic derecho en el chip: marcar los de hoy sin abrir el panel."""
        hoy = dt.date.today()
        m = tk.Menu(self.lienzo, tearoff=0)
        lista = habitos_lista()
        if not lista:
            m.add_command(label="Aún no hay hábitos", state="disabled")
        for h in lista:
            estado = estado_habito(h["id"], hoy)
            nombre = f"{h['icono']} {h['nombre']}".strip()
            sub = tk.Menu(m, tearoff=0)
            for clave, texto in (("ok", "Cumplido ✓"), ("no", "Fallido ✕"),
                                 ("", "Sin marcar")):
                sub.add_command(
                    label=("● " if estado == clave else "     ") + texto,
                    command=lambda x=h["id"], s=clave: self._habito_hoy(x, s))
            m.add_cascade(
                label=f"{MARCA_HABITO.get(estado, '·')}  {nombre}", menu=sub)
        m.add_separator()
        m.add_command(label="Abrir el panel de hábitos…",
                      command=self._abrir_habitos)
        self._abrir_menu(m, e)
        return "break"

    def _pintar_correo(self, x, y, alto):
        """Chip «✉ 3» con los correos sin leer, en la barra de arriba."""
        if not CFG.get("correo_activo", True):
            return x
        no_leidos = sum(c.get("no_leidos", 0) for c in self.correo_carpetas)
        if self.correo_error:
            texto, fondo, tinta = "✉ ⚠", C["ahora"], "#ffffff"
        elif no_leidos:
            cifra = f"{no_leidos}" if no_leidos < 100 else "99+"
            texto, fondo, tinta = f"✉ {cifra}", "#ff453a", "#ffffff"
        else:
            texto = "✉"
            fondo, tinta = C.get("boton", "#2f2f33"), C["texto"]
        c = self.lienzo
        fuente = letra(8 * BZ, "bold")
        ancho = self._w_chip_correo(fuente)      # fijo: no baila con el número
        c.create_rectangle(x, y, x + ancho, y + alto, fill=fondo, outline="",
                           tags=("correo",))
        c.create_text(x + ancho / 2, y + alto / 2, text=texto, fill=tinta,
                      font=fuente, tags=("correo",))
        self._btn_activo["correo"] = True
        c.tag_bind("correo", "<Button-1>", lambda e: self._abrir_correo())
        c.tag_bind("correo", "<Button-3>", lambda e: self._menu_correo_rapido(e))
        c.tag_bind("correo", "<Enter>", lambda e: c.config(cursor="hand2"))
        c.tag_bind("correo", "<Leave>", lambda e: c.config(cursor=""))
        return x + ancho + round(8 * ESCALA * BZ)

    def _menu_correo_rapido(self, e=None):
        """Clic derecho en el chip: ir directo a la bandeja o a las reglas."""
        m = tk.Menu(self.lienzo, tearoff=0)
        m.add_command(label="Abrir el correo…",
                      command=lambda: self._abrir_correo(modo="bandeja"))
        m.add_command(label="Ver reglas de correo…",
                      command=lambda: self._abrir_correo(modo="reglas"))
        m.add_separator()
        m.add_command(label="Actualizar correo ahora", command=self._correo_ahora)
        self._abrir_menu(m, e)
        return "break"

    def _abrir_correo(self, e=None, modo="bandeja"):
        """Abre (o cierra) el panel de correo."""
        panel = getattr(self, "_panel_correo", None)
        try:
            if panel is not None and panel.winfo_exists():
                if modo != panel.modo:
                    panel.cambiar_modo(modo)
                else:
                    panel.destroy()
                    self._panel_correo = None
                return "break"
        except Exception:
            pass
        self._panel_correo = PanelCorreo(self, modo=modo)
        try:
            self._panel_correo.lift()
        except Exception:
            pass
        if not self.correo_carpetas and not self.correo_error:
            self.refrescar_correo()
        return "break"

    def _puntos_habitos(self, x, y, ancho_col, fecha, hoy):
        """Un puntito por hábito en la esquina de la celda del mes.

        Relleno = cumplido, hueco = fallido, tenue = tocaba y sin marcar.
        Los días que no tocan no pintan nada, para no ensuciar el mes.
        """
        if not CFG.get("habitos_en_mes", True) or fecha > hoy:
            return
        lista = [h for h in habitos_lista() if toca_habito(h, fecha)]
        if not lista:
            return
        c = self.lienzo
        r = max(1.5, 2 * ESCALA)
        paso = r * 3
        # se pintan de derecha a izquierda, bajo el numero del dia
        px = x + ancho_col - 6 - r
        py = y + round(13 * ESCALA)
        for h in lista[:6]:
            estado = estado_habito(h["id"], fecha)
            if estado == "ok":
                c.create_oval(px - r, py - r, px + r, py + r,
                              fill=h["color"], outline="")
            elif estado == "no":
                c.create_oval(px - r, py - r, px + r, py + r,
                              fill="", outline=C["ahora"])
            else:
                c.create_oval(px - r, py - r, px + r, py + r,
                              fill=mezclar(h["color"], C["fondo"], 0.3),
                              outline="")
            px -= paso

    def _marcar_habito_en(self, hid, fecha, estado):
        """Marca un hábito en una fecha, desde donde sea (evento, chip, panel)."""
        poner_estado_habito(hid, fecha, estado)
        panel = getattr(self, "_panel_habitos", None)
        try:
            if panel is not None and panel.winfo_exists():
                panel.pintar()
                panel.actualizar_posicion()
        except Exception:
            pass
        self.dibujar()
        return "break"

    def _habito_hoy(self, hid, estado):
        poner_estado_habito(hid, dt.date.today(), estado)
        panel = getattr(self, "_panel_habitos", None)
        try:
            if panel is not None and panel.winfo_exists():
                panel.pintar()
                panel.actualizar_posicion()
        except Exception:
            pass
        self.dibujar()

    def _abrir_habitos(self, e=None):
        """Abre (o cierra) el panel de habitos."""
        panel = getattr(self, "_panel_habitos", None)
        try:
            if panel is not None and panel.winfo_exists():
                panel.destroy()
                self._panel_habitos = None
                return "break"
        except Exception:
            pass
        self._panel_habitos = PanelHabitos(self)
        try:
            self._panel_habitos.lift()
        except Exception:
            pass
        return "break"

    def _habitos_tic(self, ahora):
        """Un aviso al dia, a la hora del recordatorio, con lo que falte."""
        if not CFG.get("recordar_habitos", True):
            return
        hoy = ahora.date()
        if getattr(self, "_habitos_avisado", None) == hoy:
            return
        if ahora.hour < self._hora_recordatorio():
            return
        info = habitos_estado(self.eventos)
        if not info["activo"] or not info["total"]:
            return
        # a la hora del recordatorio ya se acaba el dia, asi que aqui si
        # cuentan todos los huecos, no solo los que se hayan pasado de hora
        faltan = info["total"] - info["hechos"] - info["fallados"]
        if not faltan:
            return                    # el día ya está cerrado, nada que pedir
        self._habitos_avisado = hoy
        cuerpo = f"Te queda{'n' if faltan > 1 else ''} {faltan} por marcar"
        if info["vencidos"]:
            cuerpo += f" ({info['vencidos']} estaban planificados)"
        notificar("Hábitos", cuerpo)
        if CFG.get("avisos_flotantes", True):
            AvisoFlotante(self, "Hábitos de hoy", cuerpo, C["evento"], 12,
                          al_pulsar=self._abrir_habitos)

    @staticmethod
    def _hora_recordatorio():
        try:
            return max(0, min(23, int(CFG.get("hora_recordatorio", 21))))
        except (TypeError, ValueError):
            return 21

    def _dibujar_estado(self, y, ancho=0):
        c = self.lienzo
        estado = (f"actualizado {self.actualizado:%H:%M}" if self.actualizado
                  else "sin conexión todavía")
        item = c.create_text(4, y, text=estado, anchor="nw",
                             fill=C["texto_tenue"], font=letra(7))
        x = c.bbox(item)[2] + 10
        if self.error_red:
            x = self._chip_estado(x, y, "⚠ sin red · reintentar", C["ahora"],
                                  "reintentar") + 10
            c.tag_bind("reintentar", "<Button-1>", lambda e: self.refrescar())
            c.tag_bind("reintentar", "<Enter>",
                       lambda e: c.config(cursor="hand2"))
            c.tag_bind("reintentar", "<Leave>", lambda e: c.config(cursor=""))
        x = self._pintar_proximos(x, y, ancho)
        if self.filtro:
            x = self._chip_estado(x, y, f"solo «{self.filtro[:16]}» ✕",
                                  C["evento"], "quitar_filtro") + 12
            c.tag_bind("quitar_filtro", "<Button-1>",
                       lambda e: self._filtrar(None))
        anclado = CFG.get("anclado")
        if anclado:
            try:
                falta = cuenta_atras(dt.datetime.fromisoformat(anclado["inicio"]))
            except Exception:
                falta = ""
            texto = c.create_text(
                x + 5, y, anchor="nw", tags="anclado",
                text=f"★ {anclado.get('titulo', '')[:22]} · {falta}",
                fill=anclado.get("color") or C["evento"], font=letra(7, "bold"))
            x1, y1, x2, y2 = c.bbox(texto)
            equis = c.create_text(x2 + 8, y, text="✕", anchor="nw",
                                  tags="quitar_anclado", fill=C["texto_tenue"],
                                  font=letra(7, "bold"))
            c.tag_bind("anclado", "<Button-1>", lambda e: self._ir_al_anclado())
            c.tag_bind("quitar_anclado", "<Button-1>",
                       lambda e: self._desanclar())
            x = c.bbox(equis)[2] + round(12 * ESCALA)

        # leyenda de calendarios (clic para encender o apagar cada uno)
        if len(CFG["calendarios"]) < 2:
            return
        for i, conf in enumerate(CFG["calendarios"]):
            nombre = conf.get("nombre", f"Cal {i + 1}")
            activo = conf.get("activo", True)
            if ancho and x + 14 + 6 * len(nombre) > ancho - 6:
                break
            tag = f"leyenda{i}"
            c.create_oval(x, y + 3, x + 6, y + 9, outline="",
                          fill=conf.get("color") if activo else C["texto_tenue"],
                          tags=(tag,))
            item = c.create_text(x + 10, y, text=nombre, anchor="nw",
                                 fill=C["texto"] if activo else C["texto_tenue"],
                                 font=letra(7), tags=(tag,))
            c.tag_bind(tag, "<Button-1>",
                       lambda e, n=i: self._activar_calendario(
                           n, not CFG["calendarios"][n].get("activo", True)))
            x = c.bbox(item)[2] + 12

    # -- detalle del evento al pasar el raton --------------------------------

    def _registrar(self, ev):
        """Etiqueta unica para poder identificar el evento bajo el raton."""
        self._n_evento += 1
        tag = f"ev{self._n_evento}"
        self._info_eventos[tag] = ev
        return tag

    def _quitar_detalle(self):
        self.lienzo.delete("detalle")

    def _evento_bajo_raton(self):
        c = getattr(self, "_lienzo_activo", None) or self.lienzo
        actual = c.find_withtag("current")
        if not actual:
            return None
        for t in c.gettags(actual[0]):
            if t in self._info_eventos:
                return self._info_eventos[t]
        return None

    def _mostrar_detalle(self, e):
        ev = self._evento_bajo_raton()
        if ev is not None:
            self._pintar_detalle(ev, e.x, e.y)

    # -- crear eventos -------------------------------------------------------

    # -- perfiles (una cuenta por persona) ----------------------------------

    NOMBRES_BOTONES = {
        "btn_festivos": "Festivos", "btn_nuevo": "Nuevo evento",
        "btn_ir": "Ir a una fecha", "btn_paleta": "Comandos (Ctrl+K)",
        "btn_horario": "Horario visible", "btn_horizontal": "Orientación horizontal",
        "btn_vertical": "Orientación vertical",
        "btn_cuenta": "Cuenta atrás de eventos", "btn_nivel": "Nivel (fondo / normal / encima)",
        "btn_vista": "Vistas", "btn_avisos": "Avisos", "btn_todos": "To-dos",
        "btn_correo_col": "Correo en el widget", "btn_perfil": "Perfil",
        "btn_panel": "Panel de control", "btn_dias": "Días a la vista",
        "nav_ant": "Ventana anterior", "nav_sig": "Ventana siguiente",
        "nav_hoy": "Volver a hoy", "btn_mas": "Más botones",
        "btn_refrescar": "Actualizar ahora", "btn_ocultar": "Ocultar (minimizar)",
        "btn_cerrar": "Cerrar el widget",
    }

    def _menu_desbordados(self, e=None):
        """Los botones que no caben en la barra, en un menú."""
        m = tk.Menu(self.lienzo, tearoff=0)
        pendientes = getattr(self, "_barra_desbordados", [])
        for _texto, tag in pendientes:
            accion = self._acciones_barra.get(tag)
            if accion is None:
                continue
            m.add_command(label=self.NOMBRES_BOTONES.get(tag, tag),
                          command=lambda f=accion: f(e))
        if not pendientes:
            m.add_command(label="(no falta ningún botón)", state="disabled")
        self._abrir_menu(m, e)
        return "break"

    @staticmethod
    def _etiqueta_perfil():
        """Lo que se lee en el botón de la barra: el perfil de ahora."""
        nombre = (CFG.get("perfil") or "").strip()
        if len(nombre) > 8:
            nombre = nombre[:7] + "…"
        return f"👤 {nombre}" if nombre else "👤 Perfil"

    def _menu_perfiles(self, padre):
        m = tk.Menu(padre, tearoff=0)
        actual = CFG.get("perfil") or ""
        perfiles = listar_perfiles()
        if not perfiles:
            m.add_command(label="Todavía no hay perfiles", state="disabled")
        for nombre in perfiles:
            m.add_command(
                label=("● " if nombre == actual else "     ") + nombre,
                command=lambda n=nombre: self._cambiar_perfil(n))
        m.add_separator()
        m.add_command(label="Guardar estos datos como perfil nuevo…",
                      command=self._perfil_nuevo)
        if actual:
            m.add_command(label=f"Actualizar «{actual}» con lo de ahora",
                          command=lambda: self._perfil_actualizar(actual))
            m.add_command(label=f"Dejar de usar «{actual}»",
                          command=lambda: self._cambiar_perfil(""))
        if perfiles:
            quitar = tk.Menu(m, tearoff=0)
            for nombre in perfiles:
                quitar.add_command(label=nombre,
                                   command=lambda n=nombre: self._perfil_borrar(n))
            m.add_cascade(label="Borrar un perfil", menu=quitar)
        m.add_separator()
        m.add_command(label="Rellenar los datos de este perfil…",
                      command=lambda: self._abrir_asistente(
                          self._recoger_del_asistente))
        return m

    def _cambiar_perfil(self, nombre):
        """Guarda lo del perfil de ahora y se pasa al otro."""
        anterior = CFG.get("perfil") or ""
        if anterior:
            guardar_perfil(anterior)
        if nombre:
            datos = leer_perfil(nombre)
            if not datos:
                AvisoFlotante(self, "No encuentro ese perfil",
                              f"Falta el fichero de «{nombre}»", C["ahora"], 8)
                return "break"
            for clave in CLAVES_PERFIL:
                if clave in datos:
                    CFG[clave] = datos[clave]
        CFG["perfil"] = nombre
        guardar_config()
        global TZ
        try:
            TZ = ZoneInfo(CFG["zona_horaria"])
        except Exception:
            pass
        self.eventos = []
        self.avisados.clear()
        log(f"Perfil activo: {nombre or '(ninguno)'}")
        self.dibujar()
        if nombre and self._perfil_sin_datos():
            AvisoFlotante(self, f"Perfil: {nombre}",
                          "Le faltan los datos de la cuenta: abro el asistente",
                          C["evento"], 7)
            self._abrir_asistente(self._recoger_del_asistente)
        else:
            AvisoFlotante(self, f"Perfil: {nombre or 'sin perfil'}",
                          "Leyendo el calendario de esa cuenta…", C["evento"], 6)
            self.refrescar()
        return "break"

    def _perfil_nuevo(self):
        try:
            from tkinter import simpledialog
            nombre = simpledialog.askstring(
                "Perfil nuevo",
                "Nombre del perfil (por ejemplo, tu nombre):", parent=self)
        except Exception as exc:
            log(f"No se pudo pedir el nombre del perfil: {exc}")
            return
        nombre = (nombre or "").strip()
        if not nombre:
            return
        anterior = CFG.get("perfil") or ""
        if anterior:
            guardar_perfil(anterior)          # lo de la otra persona, a salvo
        if guardar_perfil(nombre):
            CFG["perfil"] = nombre
            guardar_config()
            AvisoFlotante(self, f"Perfil «{nombre}» creado",
                          "Rellena sus datos en el asistente", C["evento"], 7)
            self._abrir_asistente(self._recoger_del_asistente)

    def _perfil_actualizar(self, nombre):
        if guardar_perfil(nombre):
            AvisoFlotante(self, f"Perfil «{nombre}» actualizado",
                          "Con los datos de cuenta de ahora mismo",
                          C["evento"], 6)

    def _perfil_borrar(self, nombre):
        try:
            os.remove(ruta_perfil(nombre))
            log(f"Perfil borrado: {nombre}")
        except Exception as exc:
            log(f"No se pudo borrar el perfil '{nombre}': {exc}")
        if CFG.get("perfil") == nombre:
            CFG["perfil"] = ""
            guardar_config()

    def _abrir_asistente(self, al_terminar=None):
        """Abre el asistente de configuración sin ventana de comandos.

        Si se pasa 'al_terminar', se llama cuando el asistente se cierra, para
        recoger lo que se haya escrito.
        """
        guion = os.path.join(BASE, "configurar.py")
        if not os.path.exists(guion):
            AvisoFlotante(self, "No encuentro el asistente",
                          "Falta configurar.py en la carpeta", C["ahora"], 8)
            return None
        import subprocess
        ejecutable = sys.executable
        if ejecutable.lower().endswith("python.exe"):
            alterno = ejecutable[:-10] + "pythonw.exe"
            ejecutable = alterno if os.path.exists(alterno) else ejecutable
        try:
            proceso = subprocess.Popen(
                [ejecutable, guion], cwd=BASE,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception as exc:
            log(f"No se pudo abrir el asistente: {exc}")
            return None
        if al_terminar:
            self._vigilar_asistente(proceso, al_terminar)
        return proceso

    def _vigilar_asistente(self, proceso, al_terminar, intentos=0):
        """Espera a que se cierre el asistente (sin bloquear la ventana)."""
        if proceso.poll() is None and intentos < 1800:      # media hora de tope
            self.after(1000, lambda: self._vigilar_asistente(
                proceso, al_terminar, intentos + 1))
            return
        try:
            al_terminar()
        except Exception as exc:
            log(f"Al volver del asistente: {exc}")

    def _recoger_del_asistente(self):
        """Lo que se acaba de escribir en el asistente pasa al perfil activo."""
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                nuevo = json.load(f)
        except Exception as exc:
            log(f"No se pudo releer la configuración: {exc}")
            return
        for clave in CLAVES_PERFIL:
            if clave in nuevo:
                CFG[clave] = nuevo[clave]
        global TZ
        try:
            TZ = ZoneInfo(CFG["zona_horaria"])
        except Exception:
            pass
        perfil = CFG.get("perfil") or ""
        if perfil:
            guardar_perfil(perfil)
        guardar_config()
        self.eventos = []
        log("Datos recogidos del asistente para el perfil "
            f"'{perfil or '-'}'")
        AvisoFlotante(self, f"Perfil {perfil or 'actualizado'}",
                      "Datos guardados; leyendo ese calendario…",
                      C["evento"], 6)
        self.dibujar()
        self.refrescar()

    @staticmethod
    def _perfil_sin_datos():
        """True si el perfil de ahora no tiene de dónde leer los eventos."""
        if CFG.get("origen") in ("outlook", "auto") and (
                CFG.get("carpetas_outlook") or CFG.get("carpeta_outlook")):
            return False
        return not any(c.get("url") for c in CFG.get("calendarios") or [])

    def _alternar_notas(self):
        """Las notas obligan a tocar el cuerpo de la cita, y Outlook avisa."""
        CFG["leer_notas"] = not CFG.get("leer_notas", False)
        guardar_config()
        log(f"Leer notas de Outlook: {CFG['leer_notas']}")
        self.refrescar()

    def _cambiar_origen(self, origen):
        CFG["origen"] = origen
        guardar_config()
        log(f"Origen de los eventos: {origen}")
        self.refrescar()

    def _copia_config(self):
        try:
            destino = copia_fechada()
            notificar("Configuración guardada", os.path.basename(destino))
        except Exception as exc:
            log(f"No se pudo copiar la configuración: {exc}")

    def _exportar_config(self):
        """Guarda config.json donde el usuario elija (para llevarlo a otro
        PC, o como copia aparte de la automática en la propia carpeta)."""
        from tkinter import filedialog, messagebox
        guardar_config()
        destino = filedialog.asksaveasfilename(
            parent=self, title="Exportar configuración",
            defaultextension=".json", initialfile="calendario_config.json",
            filetypes=[("Configuración de Calendario", "*.json")])
        if not destino:
            return
        try:
            shutil.copy(CONFIG_PATH, destino)
            log(f"Configuración exportada a {destino}")
            messagebox.showinfo(
                "Exportado", f"Configuración guardada en:\n{destino}",
                parent=self)
        except Exception as exc:
            log(f"No se pudo exportar la configuración: {exc}")
            messagebox.showerror("No se pudo exportar", str(exc), parent=self)

    def _importar_config(self):
        """Reemplaza config.json por uno elegido por el usuario (para traer
        la configuración de otro PC). Deja una copia de la actual antes,
        por si acaso, y pide reiniciar para que se note el cambio."""
        from tkinter import filedialog, messagebox
        origen = filedialog.askopenfilename(
            parent=self, title="Importar configuración",
            filetypes=[("Configuración de Calendario", "*.json")])
        if not origen:
            return
        try:
            with open(origen, encoding="utf-8") as f:
                nueva = json.load(f)
            if not isinstance(nueva, dict):
                raise ValueError("el archivo no tiene la forma esperada")
        except Exception as exc:
            messagebox.showerror(
                "Archivo no válido",
                f"No parece un archivo de configuración de Calendario:\n{exc}",
                parent=self)
            return
        if not messagebox.askyesno(
                "Importar configuración",
                "Esto reemplaza tu configuración actual (antes se guarda "
                "una copia por si quieres deshacerlo) y hay que reiniciar "
                "el programa para que se note. ¿Continuar?", parent=self):
            return
        try:
            copia_fechada()      # por si hay que volver atrás
            shutil.copy(origen, CONFIG_PATH)
            log(f"Configuración importada desde {origen}")
            messagebox.showinfo(
                "Importado",
                "Configuración importada. Cierra y vuelve a abrir el "
                "programa para que se aplique.", parent=self)
        except Exception as exc:
            log(f"No se pudo importar la configuración: {exc}")
            messagebox.showerror("No se pudo importar", str(exc), parent=self)

    def _donde_crear(self, modo):
        CFG["crear_con"] = modo
        guardar_config()
        log(f"Los eventos se crearán con: {modo}")

    def _ajustar_a_division(self, decimal):
        """Redondea una hora decimal al hueco elegido (30 min, 1 h, lo que sea)."""
        paso = self._division()
        minutos = int(max(0, decimal * 60) // paso) * paso
        hora, minuto = divmod(int(minutos), 60)
        return max(0, min(23, hora)), minuto

    def _hueco_en(self, x, y):
        """Fecha y hora del punto del lienzo donde se ha hecho doble clic."""
        dx, dy = getattr(self, "_off_rejilla", (0, 0))
        x, y = x - dx, y - dy
        ini, _ = limites_ventana(desplazamiento=self.desplazamiento)
        if self.vista == "mes":
            fecha = getattr(self, "_celdas_mes", {}).get(
                self._celda_mes_en(x, y))
            if not fecha:
                return None
            return dt.datetime.combine(fecha, dt.time(9), tzinfo=TZ)
        if self.vista == "lista":
            return None

        if self.orientacion == "vertical" and self.vista != "mes":
            i = int((y - self._cab_horas()) // self._alto_fila_v())
            if not 0 <= i < CFG["dias_visibles"] or x < self._margen_dias():
                return None
            decimal = ((x - self._margen_dias()) / self._ancho_hora()
                       + CFG["hora_inicio"])
            hora, minuto = self._ajustar_a_division(decimal)
            return dt.datetime.combine((ini + dt.timedelta(days=i)).date(),
                                       dt.time(hora, minuto), tzinfo=TZ)

        ancho_col = self._ancho_columna()
        col = int((x - MARGEN_IZQ) // ancho_col)
        if x < MARGEN_IZQ or not 0 <= col < CFG["dias_visibles"]:
            return None
        y_rejilla = CABECERA + self._alto_banda
        fecha = (ini + dt.timedelta(days=col)).date()
        if y < y_rejilla:                      # banda de todo el dia
            return dt.datetime.combine(fecha, dt.time(9), tzinfo=TZ)
        decimal = (y - y_rejilla) / self._alto_hora() + CFG["hora_inicio"]
        hora, minuto = self._ajustar_a_division(decimal)
        return dt.datetime.combine(fecha, dt.time(hora, minuto), tzinfo=TZ)

    def _celda_mes_en(self, x, y):
        ancho_col = self._ancho_columna()
        alto_celda = round(CFG.get("alto_celda_mes", 88) * ESCALA)
        if y < CABECERA:
            return None
        return ((y - CABECERA) // alto_celda, int((x - 4) // ancho_col))

    def _doble_clic(self, e):
        """Sobre un evento lo abre para cambiarlo; en un hueco vacio (o con
        sitio libre) crea uno nuevo, si esta activado. Sobre botones,
        flechas u otros controles, no hace nada, se pulse las veces que se
        pulse."""
        objetivo = self.lienzo.find_overlapping(e.x, e.y, e.x, e.y)
        etiquetas = set()
        for item in objetivo:
            etiquetas.update(self.lienzo.gettags(item))
        if etiquetas & getattr(self, "_tags_interactivos", set()):
            return None
        ev = self._evento_bajo_raton()
        if ev is not None:
            return self._editar_evento(ev)
        if not CFG.get("doble_clic_crea", False):
            return None
        if self._hueco_en(e.x, e.y) is None:
            return None
        return self._nuevo_evento(e)

    def _editar_evento(self, ev):
        """Abre la ficha con el evento ya relleno."""
        if ev.todo_dia and (es_festivo(ev.categoria)
                            or ev.categoria == "cumpleanos"):
            return "break"          # festivos y cumpleaños no son citas de nadie
        if not ev.id_outlook:
            AvisoFlotante(self, "Solo de lectura",
                          "Este evento viene del enlace ICS: para cambiarlo "
                          "hay que leer del Outlook instalado", C["ahora"], 9)
            return "break"
        self._quitar_detalle()
        self._ficha(ev=ev)
        return "break"

    def _ficha(self, ev=None, inicio=None, plantilla=None):
        """Abre el formulario propio, cerrando el que hubiera."""
        try:
            if getattr(self, "_panel_ficha", None):
                self._panel_ficha.destroy()
        except Exception:
            pass
        self._panel_ficha = Ficha(self, ev=ev, inicio=inicio,
                                  plantilla=plantilla)
        return self._panel_ficha

    def aplicar_ficha(self, datos, ev=None, id_outlook="", calendario=None):
        """Pinta ya el evento y lo guarda en Outlook por detrás.

        Así la ficha se cierra al instante: si Outlook fallara, se quita lo
        pintado y sale un aviso.
        """
        nuevo = ev is None
        if nuevo:
            ev = Evento(datos["inicio"], datos["fin"], datos["titulo"],
                        datos["todo_dia"],
                        f"ficha-{id_outlook or os.urandom(4).hex()}",
                        id_outlook=id_outlook)
            self.eventos.append(ev)
            copia = None
        else:
            copia = self._datos_de(ev)        # por si hay que dar marcha atrás
            ev.inicio, ev.fin = datos["inicio"], datos["fin"]
            ev.titulo, ev.todo_dia = datos["titulo"], datos["todo_dia"]
        ev.lugar, ev.notas = datos.get("lugar", ""), datos.get("notas", "")
        ev.categoria = datos.get("categoria") or None
        self._recordar_lugar(ev.lugar)
        self._recordar_etiqueta(ev.categoria)
        if ev.categoria:
            ev.color = self._color_de_categoria(ev.categoria) or ev.color
        repartir_columnas([e for e in self.eventos if not e.todo_dia])
        self._seguir_cuentas()
        self.dibujar()
        AvisoFlotante(self, ev.titulo[:36],
                      f"Guardado · {fecha_corta(ev.inicio)}"
                      + ("" if ev.todo_dia else f" {ev.inicio:%H:%M}"),
                      self._color_evento(ev), 6)

        def escribir():
            hecho, detalle = guardar_cita_outlook(datos, None if nuevo else ev,
                                                  calendario)
            if hecho and nuevo:
                ev.id_outlook = detalle       # el EntryID que ha dado Outlook
            return hecho, detalle

        def deshacer():
            if nuevo:
                if ev in self.eventos:
                    self.eventos.remove(ev)
            elif copia:
                ev.inicio, ev.fin = copia["inicio"], copia["fin"]
                ev.titulo, ev.todo_dia = copia["titulo"], copia["todo_dia"]
                ev.lugar, ev.notas = copia["lugar"], copia["notas"]
                ev.categoria = copia["categoria"] or None
            repartir_columnas([e for e in self.eventos if not e.todo_dia])
            self._seguir_cuentas()

        self._escribir_en_outlook(escribir, deshacer,
                                  "No se ha podido guardar el evento")
        return ev

    def plantillas_recientes(self, tope=8):
        """Lo que has hecho hace poco, listo para repetirlo."""
        vistos, salida = set(), []
        ahora = dt.datetime.now(TZ)
        for ev in sorted((e for e in self.eventos
                          if not e.parte and e.inicio <= ahora
                          and not es_festivo(e.categoria)),
                         key=lambda e: e.inicio, reverse=True):
            clave = ev.titulo.strip().lower()
            if not clave or clave in vistos:
                continue
            vistos.add(clave)
            salida.append({"titulo": ev.titulo,
                           "minutos": max(int((ev.fin - ev.inicio).total_seconds()
                                              // 60), 15),
                           "lugar": ev.lugar or "", "notas": "",
                           "categoria": ev.categoria or "",
                           "todo_dia": bool(ev.todo_dia)})
            if len(salida) >= tope:
                break
        return salida

    def crear_desde_plantilla(self, plantilla, inicio):
        """Suelta una copia de la plantilla en ese hueco del calendario."""
        minutos = int(plantilla.get("minutos", 60) or 60)
        datos = {"titulo": plantilla.get("titulo", "(sin título)"),
                 "inicio": inicio, "fin": inicio + dt.timedelta(minutes=minutos),
                 "todo_dia": bool(plantilla.get("todo_dia")),
                 "lugar": plantilla.get("lugar", ""),
                 "notas": plantilla.get("notas", ""),
                 "categoria": plantilla.get("categoria", "")}
        return self.aplicar_ficha(datos)

    def abrir_vista_aparte(self, vista, dias=None):
        """Saca esa vista en su propia ventana, sin tocar la principal."""
        for v in list(self._vistas_aparte):
            try:
                if v.vista == vista and v.winfo_exists():
                    v.lift()
                    v.pintar()
                    return v
            except Exception:
                pass
        try:
            return VistaAparte(self, vista, dias)
        except Exception as exc:
            log(f"No se pudo abrir la vista '{vista}' aparte: {exc}")
            return None

    def _refrescar_vistas_aparte(self):
        """Cuando cambian los datos, se repintan las ventanas abiertas."""
        for v in list(getattr(self, "_vistas_aparte", [])):
            try:
                if v.winfo_exists():
                    v.pintar()
                else:
                    self._vistas_aparte.remove(v)
            except Exception:
                pass

    def _abrir_accesos(self, e=None):
        """Abre (o cierra) la barra de accesos de la derecha."""
        panel = getattr(self, "_panel_accesos", None)
        try:
            if panel is not None and panel.winfo_exists():
                panel.destroy()
                self._panel_accesos = None
                return "break"
        except Exception:
            pass
        self._panel_accesos = PanelAccesos(self)
        return "break"

    def _abrir_panel_control(self, e=None, pestana=None, forzar=False):
        """Abre (o cierra) el Panel de control. Con 'pestana' salta a esa.

        Es una ventana aparte o va anclada a un lado, según lo que se elija en
        la cabecera del propio panel (igual que Correo, Hábitos...).
        """
        try:
            import extras
        except Exception as exc:
            log(f"No se pudo cargar extras.py: {exc!r}")
            return "break"
        panel = getattr(self, "_panel_control", None)
        try:
            if panel is not None and panel.winfo_exists():
                if pestana and pestana != panel.pestana:
                    panel.cambiar_pestana(pestana)
                elif forzar:
                    panel.lift()
                else:
                    panel.destroy()
                return "break"
        except Exception:
            pass
        self._panel_control = extras.PanelControl(
            self, sys.modules[__name__], pestana)
        try:
            self._panel_control.lift()
        except Exception:
            pass
        self.dibujar()
        return "break"

    def _tecla_paleta(self, e):
        """Ctrl+K, salvo dentro de un cuadro de texto (allí ya borra hasta el
        final de la línea, y no hay que pisarlo)."""
        try:
            if e.widget.winfo_class() in ("Entry", "Text", "TEntry",
                                          "TCombobox", "Spinbox"):
                return None
        except Exception:
            pass
        return self._abrir_paleta()

    def _abrir_paleta(self, e=None):
        """Abre (o cierra) la paleta de comandos (Ctrl+K)."""
        try:
            import extras
        except Exception as exc:
            log(f"No se pudo cargar extras.py: {exc!r}")
            return "break"
        paleta = getattr(self, "_paleta", None)
        try:
            if paleta is not None and paleta.winfo_exists():
                paleta.destroy()
                return "break"
        except Exception:
            pass
        self._paleta = extras.PaletaComandos(self, sys.modules[__name__])
        self.dibujar()
        return "break"

    def _accesos_al_dia(self):
        """Mantiene la barra pegada y con la vista de ahora marcada."""
        panel = getattr(self, "_panel_accesos", None)
        try:
            if panel is not None and panel.winfo_exists():
                panel.refrescar_tablero()
                panel.actualizar_posicion()
        except Exception:
            self._panel_accesos = None

    def _abrir_plantillas(self, modo="plantillas", e=None):
        """Abre (o cierra) el tablero de plantillas o el de tareas permanentes."""
        atributo = "_panel_plantillas" if modo == "plantillas" else "_panel_tareas"
        panel = getattr(self, atributo, None)
        try:
            if panel is not None and panel.winfo_exists():
                panel.destroy()
                setattr(self, atributo, None)
                return "break"
        except Exception:
            pass
        nuevo = PanelPlantillas(self, modo=modo)
        setattr(self, atributo, nuevo)
        try:
            nuevo.lift()
        except Exception:
            pass
        return "break"

    def _abrir_cumpleanos(self, e=None):
        """Abre (o cierra) el panel de cumpleaños."""
        panel = getattr(self, "_panel_cumpleanos", None)
        try:
            if panel is not None and panel.winfo_exists():
                panel.destroy()
                self._panel_cumpleanos = None
                return "break"
        except Exception:
            pass
        self._panel_cumpleanos = PanelCumpleanos(self)
        try:
            self._panel_cumpleanos.lift()
        except Exception:
            pass
        return "break"

    def _abrir_todos(self, e=None):
        """Enseña o esconde la columna de to-dos acoplada a la derecha."""
        CFG["todos_visible"] = not CFG.get("todos_visible", False)
        guardar_config()
        self.dibujar()
        return "break"

    # -- datos de los to-dos (persisten en config.json, a salvo de refrescos)

    def _todos_datos(self):
        lista = CFG.get("todos")
        if not isinstance(lista, list):
            lista = []
            CFG["todos"] = lista
        return lista

    def _todo_añadir(self, texto):
        texto = (texto or "").strip()
        if not texto:
            return
        self._todos_datos().append({"texto": texto, "hecho": False,
                                    "etiqueta": ""})
        guardar_config()
        self.dibujar()

    def _todo_alternar(self, n):
        datos = self._todos_datos()
        if 0 <= n < len(datos):
            datos[n]["hecho"] = not datos[n].get("hecho")
            guardar_config()
            self.dibujar()

    def _todo_quitar(self, n):
        datos = self._todos_datos()
        if 0 <= n < len(datos):
            datos.pop(n)
            guardar_config()
            self.dibujar()

    def _editar_todo_inline(self, n, x, y, ancho, alto):
        """Edita el texto de un to-do en el sitio, con un campo temporal."""
        datos = self._todos_datos()
        if not 0 <= n < len(datos):
            return
        c = self.lienzo
        editor = tk.Entry(c, bg=C.get("boton", "#2f2f33"), fg=C["texto"],
                          insertbackground=C["texto"], relief="flat",
                          font=letra(8))
        editor.insert(0, datos[n].get("texto", ""))
        editor.select_range(0, "end")
        ventana = c.create_window(x, y + 1, anchor="nw", window=editor,
                                  width=max(ancho, 30), height=alto - 2)

        def cerrar(guardar):
            texto = editor.get().strip()
            try:
                c.delete(ventana)
                editor.destroy()
            except Exception:
                pass
            if guardar and 0 <= n < len(datos):
                if texto:
                    datos[n]["texto"] = texto
                else:
                    datos.pop(n)
                guardar_config()
            self.dibujar()

        editor.bind("<Return>", lambda e: cerrar(True))
        editor.bind("<Escape>", lambda e: cerrar(False))
        editor.bind("<FocusOut>", lambda e: cerrar(True))
        editor.focus_set()

    def _todo_menu_etiqueta(self, e, n):
        datos = self._todos_datos()
        if not 0 <= n < len(datos):
            return
        m = tk.Menu(self, tearoff=0)
        actual = datos[n].get("etiqueta") or ""
        m.add_command(label=("● " if not actual else "     ") + "Sin etiqueta",
                      command=lambda: self._todo_poner_etiqueta(n, ""))
        for nombre in self.etiquetas_usadas():
            marca = "●" if nombre == actual else "○"
            m.add_command(label=f"{marca} {nombre}",
                          command=lambda nom=nombre:
                              self._todo_poner_etiqueta(n, nom))
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()

    def _todo_poner_etiqueta(self, n, etiqueta):
        datos = self._todos_datos()
        if 0 <= n < len(datos):
            datos[n]["etiqueta"] = etiqueta
            guardar_config()
            self.dibujar()

    def _todo_limpiar_hechos(self):
        CFG["todos"] = [t for t in self._todos_datos() if not t.get("hecho")]
        guardar_config()
        self.dibujar()

    def lugares_usados(self):
        """Sitios que ya has puesto antes, para el desplegable de la ficha."""
        cuenta = {}
        for e in self.eventos:
            sitio = (e.lugar or "").strip()
            if sitio:
                cuenta[sitio] = cuenta.get(sitio, 0) + 1
        for guardado in CFG.get("lugares_recientes") or []:
            cuenta.setdefault(guardado, 0)
        return sorted(cuenta, key=lambda k: (-cuenta[k], k.lower()))[:12]

    def _recordar_lugar(self, lugar):
        if not lugar:
            return
        recientes = [x for x in (CFG.get("lugares_recientes") or []) if x != lugar]
        CFG["lugares_recientes"] = ([lugar] + recientes)[:12]

    def etiquetas_usadas(self):
        """Etiquetas que ya aparecen en el calendario, las más usadas primero."""
        cuenta = {}
        for e in self.eventos:
            if e.categoria and not es_festivo(e.categoria):
                for parte in str(e.categoria).split(";"):
                    parte = parte.strip()
                    if parte:
                        cuenta[parte] = cuenta.get(parte, 0) + 1
        for guardada in CFG.get("etiquetas_recientes") or []:
            cuenta.setdefault(guardada, 0)
        return sorted(cuenta, key=lambda k: (-cuenta[k], k.lower()))[:15]

    def _recordar_etiqueta(self, categoria):
        """Guarda la etiqueta para poder ofrecerla aunque no haya eventos con ella."""
        if not categoria:
            return
        recientes = [e for e in (CFG.get("etiquetas_recientes") or [])
                     if e != categoria]
        CFG["etiquetas_recientes"] = ([categoria] + recientes)[:15]
        guardar_config()

    # -- cambios rapidos desde el clic derecho ------------------------------

    def _datos_de(self, ev):
        """Los campos del evento en el formato que usa la ficha."""
        return {"titulo": ev.titulo, "inicio": ev.inicio, "fin": ev.fin,
                "todo_dia": ev.todo_dia, "lugar": ev.lugar or "",
                "notas": ev.notas or "", "categoria": ev.categoria or ""}

    # -- plantillas ----------------------------------------------------------

    def _guardar_plantilla(self, ev):
        """Guarda el evento como plantilla para crearlo luego en dos clics."""
        minutos = max(int((ev.fin - ev.inicio).total_seconds() // 60), 15)
        nueva = {"titulo": ev.titulo, "minutos": minutos,
                 "lugar": ev.lugar or "", "categoria": ev.categoria or "",
                 "notas": ev.notas or "", "todo_dia": bool(ev.todo_dia)}
        otras = [p for p in (CFG.get("plantillas") or [])
                 if p.get("titulo") != nueva["titulo"]]
        CFG["plantillas"] = ([nueva] + otras)[:12]
        guardar_config()
        AvisoFlotante(self, "Plantilla guardada",
                      f"«{ev.titulo[:30]}» está ya en el botón +",
                      self._color_evento(ev), 6)
        return "break"

    def _desde_plantilla(self, plantilla, inicio=None):
        """Abre la ficha rellena con una plantilla guardada."""
        arranque = inicio or redondear_media_hora(dt.datetime.now(TZ))
        datos = dict(plantilla)
        datos["inicio"] = arranque
        datos["fin"] = arranque + dt.timedelta(
            minutes=int(plantilla.get("minutos", 60)))
        self._ficha(plantilla=datos)
        return "break"

    def _borrar_plantilla(self, titulo):
        CFG["plantillas"] = [p for p in (CFG.get("plantillas") or [])
                             if p.get("titulo") != titulo]
        guardar_config()

    def _menu_plantillas(self, e=None, inicio=None):
        """Lista de plantillas; si no hay ninguna, lo explica."""
        m = tk.Menu(self, tearoff=0)
        plantillas = CFG.get("plantillas") or []
        if not plantillas:
            m.add_command(label="Aún no hay plantillas", state="disabled")
            m.add_command(label="(clic derecho en un evento → Guardar como "
                                "plantilla)", state="disabled")
        for p in plantillas:
            minutos = int(p.get("minutos", 60))
            texto = f"{p.get('titulo', '')[:28]}  ·  {minutos} min"
            if p.get("categoria"):
                texto += f"  ·  {p['categoria'][:14]}"
            m.add_command(label=texto,
                          command=lambda q=p: self._desde_plantilla(q, inicio))
        if plantillas:
            quitar = tk.Menu(m, tearoff=0)
            for p in plantillas:
                quitar.add_command(
                    label=p.get("titulo", "")[:30],
                    command=lambda t=p.get("titulo"): self._borrar_plantilla(t))
            m.add_separator()
            m.add_cascade(label="Borrar una plantilla", menu=quitar)
        self._abrir_menu(m, e)
        return "break"

    def _duplicar(self, ev):
        """Abre la ficha con una copia del evento, para elegir el día nuevo."""
        copia = self._datos_de(ev)
        copia["inicio"] = ev.inicio + dt.timedelta(days=1)
        copia["fin"] = ev.fin + dt.timedelta(days=1)
        self._ficha(plantilla=copia)
        return "break"

    def _poner_etiqueta(self, ev, categoria):
        """Cambia la etiqueta (y con ella el color) sin abrir la ficha."""
        antes_cat, antes_color = ev.categoria, ev.color
        ev.categoria = categoria or None
        ev.color = self._color_de_categoria(categoria) if categoria else None
        self._recordar_etiqueta(categoria)
        self.dibujar()

        def deshacer():
            ev.categoria, ev.color = antes_cat, antes_color

        self._escribir_en_outlook(
            lambda: cambiar_en_outlook(ev, categoria=categoria), deshacer,
            "No se ha podido cambiar la etiqueta")
        return "break"

    def _poner_estado(self, ev, estado):
        """Ocupado / provisional / libre: cambia cómo se ve y cómo lo ve Outlook."""
        antes = ev.estado
        ev.estado = estado
        self.dibujar()

        def deshacer():
            ev.estado = antes

        self._escribir_en_outlook(
            lambda: cambiar_en_outlook(ev, estado=estado), deshacer,
            "No se ha podido cambiar cómo se muestra")
        return "break"

    def _borrar_evento(self, ev, preguntar=True):
        """Borra el evento, siempre preguntando antes. Va a la papelera de Outlook."""
        cuando = (fecha_corta(ev.inicio) if ev.todo_dia
                  else f"{fecha_corta(ev.inicio)} {ev.inicio:%H:%M}")
        if preguntar:
            try:
                from tkinter import messagebox
                if not messagebox.askyesno(
                        "Eliminar el evento",
                        f"¿Eliminar «{ev.titulo}»?\n{cuando}\n\n"
                        "Va a Elementos eliminados de Outlook, así que se puede "
                        "recuperar desde allí.", parent=self, icon="warning",
                        default="no"):
                    return False
            except Exception as exc:
                log(f"No se pudo preguntar antes de borrar: {exc}")
                return False
        hecho, motivo = borrar_en_outlook(ev)
        if not hecho:
            AvisoFlotante(self, "No se ha podido eliminar", motivo, C["ahora"], 10)
            return False
        self.eventos = [e for e in self.eventos
                        if not (e.uid == ev.uid and e.inicio == ev.inicio)]
        guardar_cache(self.eventos)
        # solo cuenta como borrado lo de ese día (otras fechas de una serie siguen ahí)
        self._seguir_cuentas((ev.inicio - dt.timedelta(hours=12), ev.inicio + dt.timedelta(hours=36)))
        self.dibujar()
        AvisoFlotante(self, f"Eliminado: {ev.titulo[:30]}",
                      "Está en Elementos eliminados de Outlook por si acaso",
                      C["texto_tenue"], 8)
        self.after(4000, self.refrescar)
        return True

    def _posponer(self, ev, minutos=None, dias=0):
        """Lo retrasa un rato o lo pasa a otro día."""
        salto = dt.timedelta(minutes=minutos or 0, days=dias)
        return self._mover_evento(ev, ev.inicio + salto)

    def _color_de_categoria(self, categoria):
        """El color que ya usan otros eventos con esa misma etiqueta."""
        for e in self.eventos:
            if e.categoria == categoria and e.color:
                return e.color
        return None

    def _alternar_doble_clic(self):
        CFG["doble_clic_crea"] = not CFG.get("doble_clic_crea", False)
        guardar_config()
        log(f"Doble clic crea eventos: {CFG['doble_clic_crea']}")

    def _nuevo_evento(self, e=None):
        """Ficha propia o, si se prefiere, Outlook (el ICS es de solo lectura)."""
        inicio = None
        if e is not None and getattr(e, "x", None) is not None:
            inicio = self._hueco_en(e.x, e.y)
        if inicio is None:
            inicio = redondear_media_hora(dt.datetime.now(TZ))
        if CFG.get("crear_con", "ficha") == "ficha":
            self._ficha(inicio=inicio)
            return "break"
        try:
            via = crear_evento(inicio)
            log(f"Nuevo evento {inicio:%d/%m %H:%M} (por {via})")
        except Exception as exc:
            log(f"No se pudo abrir Outlook: {exc}")
        return "break"

    # -- copiar y abrir ------------------------------------------------------

    def _texto_evento(self, ev, completo=True):
        if not completo:
            return ev.titulo
        partes = [ev.titulo, fecha_corta(ev.inicio)]
        partes[-1] += ("  todo el día" if ev.todo_dia
                       else f"  {ev.inicio:%H:%M} – {ev.fin:%H:%M}")
        if ev.lugar:
            partes.append("Lugar: " + ev.lugar)
        if ev.calendario and len(CFG["calendarios"]) > 1:
            partes.append("Calendario: " + ev.calendario)
        if ev.notas:
            partes.append(ev.notas)
        return "\n".join(partes)

    def _anclar(self, ev):
        CFG["anclado"] = {"uid": ev.uid, "titulo": ev.titulo,
                          "inicio": ev.inicio.isoformat(),
                          "color": ev.color or C["evento"]}
        guardar_config()
        log(f"Fijado: {ev.titulo} ({ev.inicio:%d/%m %H:%M})")
        self.dibujar()

    def _desanclar(self):
        CFG["anclado"] = None
        guardar_config()
        self.dibujar()

    def _ir_al_anclado(self):
        try:
            fecha = dt.datetime.fromisoformat(CFG["anclado"]["inicio"]).date()
            self.ir_a_fecha(fecha)
        except Exception:
            pass

    def _copiar(self, ev, completo=True):
        texto = self._texto_evento(ev, completo)
        try:
            self.clipboard_clear()
            self.clipboard_append(texto)
            self.update()
            notificar("Copiado", texto.splitlines()[0][:60])
        except Exception as exc:
            log(f"No se pudo copiar: {exc}")

    def _menu_evento(self, e):
        ev = self._evento_bajo_raton()
        if ev is None:
            return None
        self._quitar_detalle()
        m = tk.Menu(self, tearoff=0)
        m.add_command(label=ev.titulo[:45], state="disabled")
        m.add_separator()
        m.add_command(
            label=("⏳ Quitar de la cuenta atrás" if self._cuenta_de(ev)
                   else "⏳ Añadir a la cuenta atrás"),
            command=lambda: self._alternar_cuenta_evento(ev))
        m.add_command(label=("Quitar el tachado" if self.esta_hecho(ev)
                             else "Tacharlo (hecho)"),
                      command=lambda: self._alternar_hecho(ev))
        if not ev.todo_dia and not es_festivo(ev.categoria):
            m.add_command(
                label=("🔔 No avisarme de este" if self._avisa_siempre(ev)
                       else "🔔 Avisarme de este evento"),
                command=lambda: self._alternar_aviso(ev))
            m.add_cascade(label=("⏰ Alarma antes… (" + ", ".join(
                f"{n} min" for n in self._alarmas_de(ev)) + ")")
                if self._alarmas_de(ev) else "⏰ Alarma antes…",
                menu=self._menu_alarmas(m, ev))
        # si el titulo del evento habla de algun habito, se marca desde aqui:
        # es el sitio natural, sin tener que abrir el panel
        fecha_ev = ev.inicio.date()
        for h in habitos_lista():
            if not dias_planificados([ev], h, fecha_ev, fecha_ev):
                continue
            estado = estado_habito(h["id"], fecha_ev)
            sub = tk.Menu(m, tearoff=0)
            for clave, texto in (("ok", "Cumplido ✓"), ("no", "Fallido ✕"),
                                 ("", "Sin marcar")):
                sub.add_command(
                    label=("● " if estado == clave else "     ") + texto,
                    command=lambda i=h["id"], f=fecha_ev, s=clave:
                    self._marcar_habito_en(i, f, s))
            etiqueta = f"{h['icono']} {h['nombre']}".strip()
            m.add_cascade(
                label=f"{MARCA_HABITO.get(estado) or MARCA_PENDIENTE}  {etiqueta}",
                menu=sub)
        if ev.id_outlook:
            posponer = tk.Menu(m, tearoff=0)
            for texto, kw in (("+15 minutos", {"minutos": 15}),
                              ("+30 minutos", {"minutos": 30}),
                              ("+1 hora", {"minutos": 60}),
                              ("A mañana, misma hora", {"dias": 1}),
                              ("A la semana que viene", {"dias": 7}),
                              ("Al día de antes", {"dias": -1})):
                posponer.add_command(label=texto,
                                     command=lambda k=kw: self._posponer(ev, **k))
            m.add_cascade(label="Posponer", menu=posponer)

            etiquetas = tk.Menu(m, tearoff=0)
            for nombre in self.etiquetas_usadas():
                etiquetas.add_command(
                    label=("● " if ev.categoria == nombre else "     ") + nombre,
                    command=lambda n=nombre: self._poner_etiqueta(ev, n))
            etiquetas.add_separator()
            etiquetas.add_command(label="Sin etiqueta",
                                  command=lambda: self._poner_etiqueta(ev, ""))
            m.add_cascade(label="Etiqueta y color", menu=etiquetas)

            estados = tk.Menu(m, tearoff=0)
            for clave, texto in (("ocupado", "Ocupado"),
                                 ("provisional", "Provisional (a rayas)"),
                                 ("libre", "Libre (apagado)")):
                estados.add_command(
                    label=("● " if ev.estado == clave else "     ") + texto,
                    command=lambda c=clave: self._poner_estado(ev, c))
            m.add_cascade(label="Mostrarlo como", menu=estados)
            m.add_command(label="Duplicarlo…", command=lambda: self._duplicar(ev))
            m.add_command(label="Guardar como plantilla",
                          command=lambda: self._guardar_plantilla(ev))
            m.add_separator()
        if ev.id_outlook:
            m.add_command(label="Cambiar este evento…",
                          command=lambda: self._editar_evento(ev))
        m.add_command(label="Copiar todos los datos",
                      command=lambda: self._copiar(ev, True))
        m.add_command(label="Copiar solo el título",
                      command=lambda: self._copiar(ev, False))
        if ev.id_outlook:
            m.add_command(label="Abrir este evento en Outlook",
                          command=lambda: (abrir_cita_en_outlook(ev)
                                           or webbrowser.open(
                                               "https://outlook.live.com/calendar/")))
        else:
            m.add_command(label="Abrir el calendario de Outlook",
                          command=lambda: webbrowser.open(
                              "https://outlook.live.com/calendar/"))
        anclado = (CFG.get("anclado") or {}).get("uid")
        if anclado and anclado == ev.uid:
            m.add_command(label="Quitar de la barra", command=self._desanclar)
        else:
            m.add_command(label="Fijar abajo con cuenta atrás",
                          command=lambda: self._anclar(ev))
        for valor in (ev.categoria, ev.calendario):
            if valor and not es_festivo(valor) and valor != "cumpleanos":
                m.add_separator()
                m.add_command(label=f"Ver solo «{valor}»",
                              command=lambda v=valor: self._filtrar(v))
                break
        if ev.id_outlook:
            m.add_separator()
            m.add_command(label="Eliminar el evento…",
                          command=lambda: self._borrar_evento(ev))
        m.tk_popup(e.x_root, e.y_root)
        return "break"

    def _pintar_detalle(self, ev, x_raton, y_raton):
        lineas = [ev.titulo]
        if ev.todo_dia:
            dias = (ev.fin.date() - ev.inicio.date()).days
            cuando = f"{fecha_corta(ev.inicio)} · todo el día"
            if dias > 1:
                cuando = (f"{fecha_corta(ev.inicio)} – "
                          f"{fecha_corta(ev.fin - dt.timedelta(days=1))}"
                          f" · todo el día")
            lineas.append(cuando)
        else:
            lineas.append(f"{fecha_corta(ev.inicio)} · {ev.inicio:%H:%M} – "
                          f"{ev.fin:%H:%M}")
        if ev.estado and ev.estado != "ocupado":
            lineas.append({"provisional": "Provisional", "libre": "Libre",
                           "fuera": "Fuera de la oficina"}[ev.estado])
        if ev.lugar:
            lineas.append("En " + ev.lugar)
        if ev.calendario and len(CFG["calendarios"]) > 1:
            lineas.append(ev.calendario)
        if ev.categoria and not es_festivo(ev.categoria) and ev.categoria not in (
                                                 "cumpleanos"):
            lineas.append("Etiqueta: " + ev.categoria)
        if ev.notas:
            lineas.append(ev.notas)

        self._quitar_detalle()
        c = self.lienzo
        texto = c.create_text(0, 0, text="\n".join(lineas), anchor="nw",
                              width=230, justify="left", fill=C["texto"],
                              font=letra(8), tags=("detalle",),
                              state="disabled")
        x1, y1, x2, y2 = c.bbox(texto)
        ancho_caja, alto_caja = x2 - x1 + 16, y2 - y1 + 12
        # se coloca junto al raton pero sin salirse de la ventana
        x = min(x_raton + 14, max(4, c.winfo_width() - ancho_caja - 4))
        y = min(y_raton + 14, max(4, c.winfo_height() - alto_caja - 4))
        c.coords(texto, x + 8, y + 6)
        fondo = c.create_rectangle(x, y, x + ancho_caja, y + alto_caja,
                                   fill="#26262b", outline=ev.color or C["evento"],
                                   tags=("detalle",), state="disabled")
        c.tag_lower(fondo, texto)

    def _titulo_rango(self, ini):
        fin = ini + dt.timedelta(days=CFG["dias_visibles"] - 1)
        if ini.month == fin.month:
            return f"{ini.day} – {fin.day} {MESES[fin.month - 1]}"
        return (f"{ini.day} {MESES[ini.month - 1]} – "
                f"{fin.day} {MESES[fin.month - 1]}")

    def _mostrables(self):
        """Eventos descargados mas los festivos, ya filtrados."""
        eventos = self.eventos + self._festivos
        if self.filtro:
            eventos = [e for e in eventos if self.filtro in
                       (e.categoria, e.calendario)]
        return eventos

    def _ahora_y_siguiente(self, ahora=None):
        """El evento que esta pasando y el proximo que empieza."""
        ahora = ahora or dt.datetime.now(TZ)
        en_curso = siguiente = None
        for ev in sorted((e for e in self._mostrables() if not e.todo_dia),
                         key=lambda e: e.inicio):
            if ev.inicio <= ahora < ev.fin:
                if en_curso is None or ev.fin < en_curso.fin:
                    en_curso = ev
            elif ev.inicio > ahora and siguiente is None:
                siguiente = ev
        return en_curso, siguiente

    def _filtrar(self, valor):
        self.filtro = CFG["filtro"] = valor
        guardar_config()
        self.dibujar()

    def _construir_menu_filtro(self, padre):
        m = tk.Menu(padre, tearoff=0)
        m.add_command(label=("● " if not self.filtro else "     ") + "Ver todo",
                      command=lambda: self._filtrar(None))
        etiquetas = sorted({e.categoria for e in self.eventos if e.categoria})
        calendarios = sorted({e.calendario for e in self.eventos if e.calendario})
        for grupo in (etiquetas, calendarios if len(calendarios) > 1 else []):
            if not grupo:
                continue
            m.add_separator()
            for valor in grupo:
                m.add_command(
                    label=("● " if self.filtro == valor else "     ")
                          + f"Solo «{valor}»",
                    command=lambda v=valor: self._filtrar(v))
        return m

    def _eventos_del_dia(self, fecha):
        """Eventos que tocan esa fecha, de todo el dia primero y luego por hora."""
        del_dia = []
        for ev in self._mostrables():
            if ev.todo_dia:
                if ev.inicio.date() <= fecha < ev.fin.date():
                    del_dia.append(ev)
            elif ev.inicio.date() == fecha:
                del_dia.append(ev)
        del_dia.sort(key=lambda e: (not e.todo_dia, e.inicio))
        return del_dia

    def _medidas_especiales_horizontal(self, ini):
        """Festivos/cumpleaños agrupados por dia visible, y la altura maxima
        que ocupan (para reservar el sitio justo arriba de la rejilla)."""
        ancho_col = self._ancho_columna()
        especiales = self._especiales_por_dia(ini)
        ancho_disp = ancho_col - 4
        alto_max = 0
        for items in especiales.values():
            alto = sum(self._medir_especial(icono, texto, ancho_disp)
                      for icono, texto, *_ in items)
            alto_max = max(alto_max, alto)
        return especiales, alto_max

    def _colocar_todo_dia(self, ini):
        """Reparte los eventos de todo el dia en filas, ocupando todos sus dias.

        Festivos y cumpleaños NO entran aquí: tienen su propio recuadro
        compacto por día (ver '_especiales_por_dia'), no una barra que se
        estira a lo ancho de la columna como un evento real.
        """
        dias = CFG["dias_visibles"]
        filas = []      # cada fila: lista de (col_ini, col_fin) ocupados
        colocados = []
        for ev in self._mostrables():
            if not ev.todo_dia or ev.categoria in ("nacional", "regional",
                                                    "local", "cumpleanos"):
                continue
            c1 = (ev.inicio.date() - ini.date()).days
            c2 = (ev.fin.date() - ini.date()).days - 1  # DTEND es exclusivo
            c2 = max(c2, c1)
            if c2 < 0 or c1 >= dias:
                continue
            c1, c2 = max(c1, 0), min(c2, dias - 1)
            for fila, ocupado in enumerate(filas):
                if all(c2 < a or c1 > b for a, b in ocupado):
                    ocupado.append((c1, c2))
                    break
            else:
                filas.append([(c1, c2)])
                fila = len(filas) - 1
            colocados.append((ev, c1, c2, fila))
        return colocados, len(filas)

    def _especiales_por_dia(self, ini):
        """Festivos y cumpleaños de los días visibles, agrupados por columna.

        Devuelve {indice_de_dia: [(icono, texto, color), ...]}. Cada uno se
        pinta luego como una cajita compacta (icono + nombre), del tamaño
        de su propio texto, no de la columna entera.
        """
        dias = CFG["dias_visibles"]
        porcol = {}
        for f in self._festivos:
            i = (f.inicio.date() - ini.date()).days
            if not 0 <= i < dias:
                continue
            if f.categoria == "cumpleanos":
                icono = "🎂"
                texto = f.titulo[2:].strip() if f.titulo.startswith("🎂") \
                    else f.titulo
            else:
                icono = "🎉"
                texto = f.titulo
            color = f.color or C["texto_tenue"]
            porcol.setdefault(i, []).append((icono, texto, color, f))
        return porcol

    # -- vista de lista ------------------------------------------------------

    @staticmethod
    def _lineas_max():
        """Cuantas lineas como mucho puede ocupar un titulo en la lista."""
        return max(1, min(4, int(CFG.get("lineas_lista", 2))))

    @staticmethod
    def _hora_lista(ev):
        hora = f"← {ev.fin:%H:%M}" if ev.viene else f"{ev.inicio:%H:%M}"
        return hora + (" →" if ev.sigue else "")

    def _sangria_lista(self, filas):
        """x del titulo: justo detras de la hora mas ancha que haya que escribir."""
        horas = [self._ancho_texto(self._hora_lista(ev))
                 for t, ev in filas if t == "ev" and not ev.todo_dia]
        return 19 + (max(horas) + 8 if horas else 0)

    def _lineas_envueltas(self, texto, ancho, fuente=None):
        """Lineas que ocupa el texto al partirlo por palabras, como hace Tk."""
        if ancho < 12 or not texto:
            return 1
        lineas, actual = 1, ""
        for palabra in texto.split():
            prueba = (actual + " " + palabra).strip()
            if actual and self._ancho_texto(prueba, fuente) > ancho:
                lineas, actual = lineas + 1, palabra
            else:
                actual = prueba
            while len(actual) > 1 and self._ancho_texto(actual, fuente) > ancho:
                # una palabra sola mas larga que la caja: Tk tambien la parte
                cabe = max(1, int(len(actual) * ancho
                                  / max(self._ancho_texto(actual, fuente), 1)))
                actual = actual[cabe:]
                lineas += 1
        return lineas

    def _alto_minimo_fila(self):
        """Alto de una fila de la lista de una sola línea, con la letra de ahora."""
        return max(ALTO_FILA_EV, self._alto_linea(letra(8)) + round(6 * ESCALA))

    def _alto_fila(self, tipo, dato, ancho, sangria):
        """Alto de una fila: crece si el titulo necesita varias lineas."""
        if tipo == "dia":
            return ALTO_FILA_DIA
        minimo = self._alto_minimo_fila()
        if self._lineas_max() == 1:
            return minimo
        fuente = letra(8, "bold") if dato.todo_dia else letra(8)
        # el mismo ancho que usa el pintado, para no dejar lineas vacias
        util = max(ancho - (29 if dato.todo_dia else sangria + 8), 24)
        lineas = min(self._lineas_max(),
                     self._lineas_envueltas(dato.titulo, util, fuente))
        if lineas == 1:
            return minimo
        return 6 + lineas * self._alto_linea(fuente)

    def _alto_filas(self, filas, ancho):
        sangria = self._sangria_lista(filas)
        return sum(self._alto_fila(t, d, ancho, sangria) for t, d in filas)

    def _filas_lista(self, ini):
        """Cabecera + eventos de cada dia, saltando los dias vacios."""
        filas = []
        for i in range(CFG["dias_visibles"]):
            fecha = (ini + dt.timedelta(days=i)).date()
            del_dia = self._eventos_del_dia(fecha)
            if not del_dia:
                continue
            filas.append(("dia", fecha))
            filas.extend(("ev", ev) for ev in del_dia)
        return filas

    def _recortar_filas(self, filas, disponible, ancho=None):
        """Corta la lista por lo que cabe y dice cuantos eventos se quedan fuera."""
        ancho = ancho or self._ancho_lista()
        sangria = self._sangria_lista(filas)
        if self._alto_filas(filas, ancho) <= disponible:
            return filas, 0
        acumulado, corte = 0, len(filas)
        for n, (t, d) in enumerate(filas):
            salto = self._alto_fila(t, d, ancho, sangria)
            if acumulado + salto > disponible - 12:   # deja sitio al "+N mas"
                corte = n
                break
            acumulado += salto
        return filas[:corte], sum(1 for t, _ in filas[corte:] if t == "ev")

    def _linea_ahora_lista(self, x0, y, ancho, ahora):
        """Raya con la hora actual: separa lo que ya ha pasado de lo que queda."""
        c = self.lienzo
        y += 4
        fuente = letra(6, "bold")
        c.create_text(x0 + 8, y, text=f"{ahora:%H:%M}", anchor="w",
                      fill=C["ahora"], font=fuente)
        x = x0 + 12 + self._ancho_texto(f"{ahora:%H:%M}", fuente)
        c.create_oval(x, y - 2, x + 4, y + 2, fill=C["ahora"], outline="")
        c.create_line(x + 4, y, x0 + ancho - 8, y, fill=C["ahora"])
        return y + 5

    def _pintar_lista(self, ini, filas, x0, ancho, alto_total):
        c = self.lienzo
        hoy = dt.datetime.now(TZ).date()
        if x0:
            self._divisor_vertical(x0, alto_total)

        disponible = alto_total - ALTO_ESTADO - Y_LISTA - 4
        if CFG.get("linea_ahora_lista", True) and any(t == "dia" and d == hoy
                                                      for t, d in filas):
            disponible -= 10          # el sitio de la raya de la hora actual
        filas, recorte = self._recortar_filas(filas, disponible, ancho)
        sangria = self._sangria_lista(filas)

        ahora = dt.datetime.now(TZ)
        marca_ahora = CFG.get("linea_ahora_lista", True)
        # ultima fila del bloque de hoy, por si todo lo de hoy ya ha pasado
        fin_de_hoy = max((n for n, (t, d) in enumerate(filas)
                          if t == "ev" and not d.todo_dia
                          and d.inicio.date() == hoy), default=None)

        y = Y_LISTA
        mes_escrito = None
        for n, (tipo, dato) in enumerate(filas):
            if marca_ahora and (
                    (tipo == "ev" and not dato.todo_dia
                     and dato.inicio.date() == hoy and dato.inicio > ahora)
                    or (fin_de_hoy is not None and n == fin_de_hoy + 1)):
                marca_ahora = False           # solo se pinta una vez
                y = self._linea_ahora_lista(x0, y, ancho, ahora)
            if tipo == "dia":
                es_hoy = dato == hoy
                etiqueta = f"{DIAS[dato.weekday()]} {dato.day}"
                if dato.month != mes_escrito:      # el mes solo cuando cambia
                    etiqueta += f" {MESES[dato.month - 1]}"
                    mes_escrito = dato.month
                if es_hoy:
                    etiqueta += "  · hoy"
                c.create_line(x0 + 8, y + 1, x0 + ancho - 8, y + 1,
                              fill=C["rejilla"])
                c.create_text(x0 + 10, y + 11, text=etiqueta, anchor="w",
                              font=letra(8, "bold"),
                              fill=C["evento"] if es_hoy else C["texto"])
                y += ALTO_FILA_DIA
                continue

            ev = dato
            alto_f = self._alto_fila("ev", ev, ancho, sangria)
            marca = ("evento", self._registrar(ev))
            c.create_rectangle(x0 + 2, y, x0 + ancho - 4, y + alto_f,
                               fill=C["fondo"], outline="", tags=marca)
            c.create_rectangle(x0 + 10, y + 3, x0 + 13, y + alto_f - 3,
                               fill=self._color_evento(ev), outline="")
            if ev.todo_dia:
                # sin hora: el titulo ocupa toda la fila
                self._texto_en_caja(x0 + 19, y + 2, ancho - 29, alto_f - 4,
                                    ev.titulo, fill=ev.color or C["texto"],
                                    font=self._fuente_evento(ev, 8, "bold"),
                                    tags=marca)
            else:
                c.create_text(x0 + 19, y + 3, text=self._hora_lista(ev), anchor="nw",
                              fill=C["texto_tenue"], font=letra(8))
                self._texto_en_caja(x0 + sangria, y + 3, ancho - sangria - 8,
                                    alto_f - 5, ev.titulo, fill=C["texto"],
                                    font=self._fuente_evento(ev), tags=marca)
            y += alto_f

        if marca_ahora and fin_de_hoy is not None and fin_de_hoy == len(filas) - 1:
            y = self._linea_ahora_lista(x0, y, ancho, ahora)

        if not filas:
            c.create_text(x0 + ancho / 2, Y_LISTA + 20,
                          text="Sin eventos en estos días",
                          fill=C["texto_tenue"], font=letra(9))
        elif recorte:
            c.create_text(x0 + ancho / 2, y + 7, text=f"+{recorte} eventos más",
                          fill=C["texto_tenue"], font=letra(7))

    # -- vista de calendario -------------------------------------------------

    @staticmethod
    def _linea_rejilla():
        """Color de las lineas finas, o None si estan desactivadas."""
        return C["rejilla"] if CFG.get("lineas_rejilla", True) else None

    def _etiqueta_evento(self, ev, ancho):
        """Con poco sitio, se cae la hora y se deja solo el titulo."""
        titulo = ev.titulo
        cola = " →" if ev.sigue else ""
        if ev.viene:
            return "← " + titulo + cola
        hora = f"{ev.inicio:%H:%M} "
        # se queda la hora solo si detras cabe un trozo util del titulo
        if self._ancho_texto(hora + titulo[:7]) > ancho - 8:
            return titulo + cola
        return hora + titulo + cola

    def _fuente_evento(self, ev, tam=8, estilo=None):
        """La letra del evento, tachada si está marcado como hecho."""
        if self.esta_hecho(ev):
            estilo = (estilo + " overstrike") if estilo else "overstrike"
        return letra(tam, estilo)

    @staticmethod
    def _clave_hecho(ev):
        return f"{ev.uid}|{ev.inicio:%Y-%m-%dT%H:%M}"

    def esta_hecho(self, ev):
        """True si el evento se ha tachado a mano, o si ya pasó y así se pidió."""
        if self._clave_hecho(ev) in set(CFG.get("hechos") or []):
            return True
        return bool(CFG.get("tachar_pasados")) and not ev.todo_dia \
            and ev.fin < dt.datetime.now(TZ)

    def _limpiar_tachados(self):
        CFG["hechos"] = []
        guardar_config()
        log("Tachados borrados")
        self.dibujar()

    def _alternar_hecho(self, ev):
        """Tacha o destacha el evento; se guarda aquí, no toca Outlook."""
        clave = self._clave_hecho(ev)
        hechos = [h for h in (CFG.get("hechos") or []) if h != clave]
        if len(hechos) == len(CFG.get("hechos") or []):
            hechos.append(clave)
        CFG["hechos"] = hechos[-400:]        # no crece sin fin
        guardar_config_pronto(self)
        self.dibujar()
        return "break"

    def _color_evento(self, ev, ahora=None):
        """Color del evento, apagado si ya ha terminado o si está tachado."""
        color = ev.color or C["evento"]
        if self.esta_hecho(ev):
            return mezclar(color, C["fondo"], 0.72)
        if CFG.get("atenuar_pasado", True):
            ahora = ahora or dt.datetime.now(TZ)
            if not ev.todo_dia and ev.fin <= ahora:
                return mezclar(color, C["fondo"], 0.42)
        return color

    def _pintar_rejilla(self, ini, todo_dia):
        c = self.lienzo
        dias = CFG["dias_visibles"]
        ancho_col = self._ancho_columna()
        alto_hora = self._alto_hora()
        y_rejilla = CABECERA + self._alto_banda
        hoy = dt.datetime.now(TZ).date()
        h0, h1 = CFG["hora_inicio"], CFG["hora_fin"]
        horas = h1 - h0
        y_fin = y_rejilla + horas * alto_hora

        # fondo de las columnas (hoy, fines de semana y festivos)
        dias_festivos = {f.inicio.date() for f in self._festivos
                         if f.categoria != "cumpleanos"}
        dias_cumple = {f.inicio.date() for f in self._festivos
                      if f.categoria == "cumpleanos"}
        for i in range(dias):
            fecha = (ini + dt.timedelta(days=i)).date()
            color = fondo_dia(i)
            if fecha == hoy:
                color = mezclar(C["evento"], color, 0.22)
            elif fecha in dias_cumple:
                color = mezclar(CFG.get("color_cumpleanos", "#a855f7"),
                               color, 0.18)
            x = MARGEN_IZQ + i * ancho_col
            c.create_rectangle(x, CABECERA - 4, x + ancho_col, y_fin,
                               fill=color, outline="")

        # cabecera de dias (festivos y cumpleaños van en su propia cajita,
        # justo debajo, no aquí)
        for i in range(dias):
            fecha = ini + dt.timedelta(days=i)
            x = MARGEN_IZQ + i * ancho_col
            resalta = fecha.date() == hoy
            if ancho_col < 46:                     # columnas muy estrechas
                etiqueta = f"{INICIALES_DIA[fecha.weekday()]}{fecha.day}"
            else:
                etiqueta = f"{DIAS[fecha.weekday()]} {fecha.day}"
            if (i == 0 or fecha.day == 1) and ancho_col >= 90:
                etiqueta += f" {MESES[fecha.month - 1]}"
            c.create_text(
                x + ancho_col / 2, BARRA + 14, text=etiqueta,
                fill=C["evento"] if resalta else C["texto"],
                font=letra(11, "bold" if resalta else "normal"),
            )

        # rejilla (las lineas finas se pueden quitar desde el menu)
        # horas: casi transparentes; días: barra marcada entre columnas
        linea = mezclar(C["texto"], C["fondo"], 0.10)
        suave = mezclar(C["texto"], C["fondo"], 0.05)
        barra_dia = mezclar(C["texto"], C["fondo"], 0.40)
        alto_paso = self._division() / 60 * alto_hora
        for minutos, en_punto in self._marcas_horarias():
            y = y_rejilla + minutos / 60 * alto_hora
            if en_punto or alto_paso >= 8:
                c.create_line(MARGEN_IZQ, y, MARGEN_IZQ + dias * ancho_col, y,
                              fill=linea if en_punto else suave)
            if en_punto or alto_paso >= round(13 * ESCALA):
                c.create_text(MARGEN_IZQ - 8, y, anchor="e",
                              text=self._etiqueta_marca(minutos),
                              fill=C["texto_tenue"],
                              font=letra(8 if en_punto else 7))
        for i in range(dias + 1):
            x = MARGEN_IZQ + i * ancho_col
            c.create_line(x, CABECERA - 4, x, y_fin, fill=barra_dia,
                          width=max(2, round(2 * ESCALA)))

        # festivos y cumpleaños: cajita propia, apiladas encima de los
        # eventos reales de todo el dia
        especiales, alto_esp_max = self._medidas_especiales_horizontal(ini)
        for i, items in especiales.items():
            x = MARGEN_IZQ + i * ancho_col + 2
            ancho_disp = ancho_col - 4
            y_cursor = CABECERA
            for icono, texto, color, ev_esp in items:
                _, alto_item = self._dibujar_especial(x, y_cursor, ancho_disp,
                                                       icono, texto, color,
                                                       ev_esp)
                y_cursor += alto_item

        # eventos de todo el dia (banda superior, varias filas y varios dias,
        # debajo de los festivos/cumpleaños si los hay)
        for ev, c1, c2, fila in todo_dia:
            x1 = MARGEN_IZQ + c1 * ancho_col + 2
            x2 = MARGEN_IZQ + (c2 + 1) * ancho_col - 2
            y1 = CABECERA + alto_esp_max + fila * ALTO_CHIP
            marca_chip = ("evento", self._registrar(ev))
            c.create_rectangle(x1, y1, x2, y1 + ALTO_CHIP - 3,
                               fill=ev.color or C["evento"], outline="",
                               tags=marca_chip)
            self._texto_en_caja(x1 + 5, y1 + 2, x2 - x1 - 10, ALTO_CHIP - 6,
                                ev.titulo, fill=C["evento_texto"],
                                font=self._fuente_evento(ev), tags=marca_chip)

        # eventos con hora
        fuera = {}   # col -> [antes del horario, despues del horario]
        for ev in self.eventos:
            if ev.todo_dia:
                continue
            col = (ev.inicio.date() - ini.date()).days
            if not 0 <= col < dias:
                continue

            h_ini = ev.inicio.hour + ev.inicio.minute / 60
            h_fin = ev.fin.hour + ev.fin.minute / 60
            if ev.fin.date() > ev.inicio.date() or h_fin == 0:
                h_fin = 24.0

            # fuera del horario visible: se anota para marcarlo en el margen
            if h_fin <= h0 or h_ini >= h1:
                antes, despues = fuera.setdefault(col, [[], []])
                (antes if h_fin <= h0 else despues).append(ev)
                continue

            v_ini = max(h_ini, h0)
            v_fin = min(h_fin, h1)
            ancho_ev = (ancho_col - 4) / ev.ncols
            x1 = MARGEN_IZQ + col * ancho_col + 2 + ev.col * ancho_ev
            y1 = y_rejilla + (v_ini - h0) * alto_hora
            y2 = y_rejilla + (v_fin - h0) * alto_hora
            y2 = max(y2, y1 + 15)
            pasado = (not ev.todo_dia
                      and ev.fin <= dt.datetime.now(TZ)
                      and CFG.get("atenuar_pasado", True))
            color = self._color_evento(ev)
            marca = ("evento", self._registrar(ev))
            if ev.estado == "provisional":      # solo el contorno, a rayas
                c.create_rectangle(x1 + 1, y1 + 1, x1 + ancho_ev - 2, y2 - 1,
                                   fill=C["fondo"], outline=color, width=2,
                                   dash=(4, 3), tags=marca)
            elif ev.estado == "libre":          # muy apagado, con barra lateral
                c.create_rectangle(x1, y1, x1 + ancho_ev - 1, y2,
                                   fill=mezclar(color, C["fondo"], 0.18),
                                   outline=C["fondo"], tags=marca)
                c.create_rectangle(x1, y1, x1 + 3, y2, fill=color, outline="")
            else:
                c.create_rectangle(x1, y1, x1 + ancho_ev - 1, y2,
                                   fill=color, outline=C["fondo"], tags=marca)

            claro = ev.estado in ("provisional", "libre")
            etiqueta = self._etiqueta_evento(ev, ancho_ev)
            self._texto_en_caja(x1 + 4, y1 + 2, ancho_ev - 8, y2 - y1 - 2,
                                etiqueta, font=self._fuente_evento(ev), tags=marca,
                                fill=(C["texto"] if claro
                                      else mezclar(C["evento_texto"], C["fondo"], 0.6)
                                      if pasado else C["evento_texto"]))

        # avisos de eventos fuera del horario visible
        for col, (antes, despues) in fuera.items():
            x = MARGEN_IZQ + (col + 1) * ancho_col - 4
            if antes:
                primero = min(antes, key=lambda e: e.inicio)
                extra = f" +{len(antes) - 1}" if len(antes) > 1 else ""
                c.create_text(x, y_rejilla + 7,
                              text=f"▲ {primero.inicio:%H:%M}{extra}",
                              anchor="e", fill=C["texto_tenue"],
                              font=letra(7))
            if despues:
                ultimo = min(despues, key=lambda e: e.inicio)
                extra = f" +{len(despues) - 1}" if len(despues) > 1 else ""
                c.create_text(x, y_fin - 7,
                              text=f"▼ {ultimo.inicio:%H:%M}{extra}",
                              anchor="e", fill=C["texto_tenue"],
                              font=letra(7))

        # barra tenue donde empieza cada semana
        for i in columnas_de_semana(ini, dias):
            x = MARGEN_IZQ + i * ancho_col
            c.create_line(x, CABECERA - 4, x, y_fin + ALTO_MESES,
                          fill=color_semana(),
                          width=max(1, round(1.5 * ESCALA)))

        # franja con el mes, centrado sobre sus columnas, y barra al cambiar
        anio_actual = hoy.year
        for n, (clave, c1, c2) in enumerate(grupos_de_mes(ini, dias)):
            x1 = MARGEN_IZQ + c1 * ancho_col
            x2 = MARGEN_IZQ + (c2 + 1) * ancho_col
            anio, mes = clave
            nombre = MESES_LARGOS[mes - 1]
            if (x2 - x1) < len(nombre) * round(7 * ESCALA) + 20:
                nombre = MESES[mes - 1]
            nombre = nombre.capitalize()
            if anio != anio_actual:
                nombre += f" {anio}"
            c.create_text((x1 + x2) / 2, y_fin + ALTO_MESES / 2, text=nombre,
                          fill=C["texto_tenue"], font=letra(8))
            if n:      # barra vertical justo donde empieza el mes nuevo
                c.create_line(x1, CABECERA - 4, x1, y_fin + ALTO_MESES,
                              fill=C["texto_tenue"], width=max(2, round(2 * ESCALA)))

        visibles = len(todo_dia) + sum(
            1 for ev in self._mostrables() if not ev.todo_dia
            and 0 <= (ev.inicio.date() - ini.date()).days < dias)
        if not visibles:
            c.create_text(MARGEN_IZQ + (dias * ancho_col) / 2,
                          (y_rejilla + y_fin) / 2,
                          text="Sin eventos en estos días",
                          fill=C["texto_tenue"], font=letra(9))

        # linea de la hora actual
        ahora = dt.datetime.now(TZ)
        if h0 <= ahora.hour < h1:
            col = (ahora.date() - ini.date()).days
            if 0 <= col < dias:
                y = y_rejilla + (ahora.hour + ahora.minute / 60 - h0) * alto_hora
                x = MARGEN_IZQ + col * ancho_col
                c.create_line(x, y, x + ancho_col, y, fill=C["ahora"], width=2)
                c.create_oval(x - 3, y - 3, x + 3, y + 3, fill=C["ahora"], outline="")

    # -- vista de mes --------------------------------------------------------

    def _semanas_del_mes(self, primero):
        """Lista de semanas (listas de 7 fechas) que cubren ese mes."""
        inicio = primero - dt.timedelta(days=primero.weekday())
        dias_mes = (primero.replace(day=28) + dt.timedelta(days=4))
        dias_mes = (dias_mes.replace(day=1) - dt.timedelta(days=1)).day
        total = (primero.weekday() + dias_mes + 6) // 7
        return [[inicio + dt.timedelta(days=f * 7 + c) for c in range(7)]
                for f in range(total)]

    def _dibujar_mes(self, ini):
        primero = ini.date().replace(day=1)
        semanas = self._semanas_del_mes(primero)
        ancho_col = self._ancho_columna()
        alto_celda = round(CFG.get("alto_celda_mes", 88) * ESCALA)
        ancho = 7 * ancho_col + 8
        alto = CABECERA + len(semanas) * alto_celda + ALTO_ESTADO + 2
        self._ajustar(ancho, alto)

        c = self.lienzo
        c.delete("all")
        self._info_eventos.clear()
        self._n_evento = 0
        self._dibujar_barra(
            ancho, f"{MESES_LARGOS[primero.month - 1].capitalize()} {primero.year}")

        hoy = dt.datetime.now(TZ).date()
        festivos = {f.inicio.date() for f in self._festivos
                   if f.categoria != "cumpleanos"}
        cumples = {f.inicio.date() for f in self._festivos
                  if f.categoria == "cumpleanos"}
        alto_fila = round(15 * ESCALA)
        for i, ini_dia in enumerate("LMXJVSD"):
            c.create_text(4 + ancho_col * (i + 0.5), BARRA + 14, text=ini_dia,
                          fill=C["texto_tenue"], font=letra(8))

        self._celdas_mes = {}
        for f, semana in enumerate(semanas):
            for col, fecha in enumerate(semana):
                self._celdas_mes[(f, col)] = fecha
                x = 4 + col * ancho_col
                y = CABECERA + f * alto_celda
                otro_mes = fecha.month != primero.month
                if fecha == hoy:
                    fondo = C["hoy"]
                elif fecha in festivos or fecha.weekday() >= 5:
                    fondo = C.get("finde", C["fondo"])
                else:
                    fondo = C["fondo"]
                c.create_rectangle(x, y, x + ancho_col, y + alto_celda,
                                   fill=fondo,
                                   outline=self._linea_rejilla() or fondo)
                c.create_text(x + 5, y + 4,
                              text=str(fecha.day), anchor="nw",
                              fill=(C["texto_tenue"] if otro_mes
                                    else C["evento"] if fecha == hoy
                                    else CFG.get("color_cumpleanos", "#a855f7")
                                    if fecha in cumples
                                    else C["texto"]),
                              font=letra(8, "bold" if (fecha == hoy
                                                       or fecha in cumples)
                                        else "normal"))
                if fecha in cumples:
                    c.create_text(x + ancho_col - 6, y + 4, text="🎂",
                                 anchor="ne", font=letra(8))
                self._puntos_habitos(x, y, ancho_col, fecha, hoy)

                del_dia = self._eventos_del_dia(fecha)
                caben = max(1, (alto_celda - round(20 * ESCALA)) // alto_fila)
                for n, ev in enumerate(del_dia[:caben]):
                    if len(del_dia) > caben and n == caben - 1:
                        c.create_text(x + 5, y + round(18 * ESCALA) + n * alto_fila,
                                      anchor="nw", text=f"+{len(del_dia) - n} más",
                                      fill=C["texto_tenue"], font=letra(7))
                        break
                    yf = y + round(18 * ESCALA) + n * alto_fila
                    marca_mes = ("evento", self._registrar(ev))
                    c.create_rectangle(x + 5, yf + 2, x + 8, yf + alto_fila - 3,
                                       fill=self._color_evento(ev), outline="",
                                       tags=marca_mes)
                    etiqueta = (ev.titulo if ev.todo_dia
                                else self._etiqueta_evento(ev, ancho_col - 18))
                    self._texto_en_caja(x + 11, yf, ancho_col - 18, alto_fila - 2,
                                        etiqueta, fill=C["texto"], tags=marca_mes,
                                        font=self._fuente_evento(ev, 7))

        self._dibujar_estado(alto - ALTO_ESTADO + 2, ancho)
        self._dibujar_agarradera(ancho, alto)
        self._accesos_al_dia()
        if not getattr(self, "_pintando_aparte", False):
            self._refrescar_vistas_aparte()

    # -- rejilla girada: dias en filas, horas hacia la derecha ---------------

    def _pintar_vertical(self, ini, todo_dia):
        c = self.lienzo
        dias = CFG["dias_visibles"]
        h0, h1 = CFG["hora_inicio"], CFG["hora_fin"]
        horas = h1 - h0
        ancho_hora = self._ancho_hora()
        alto_fila = self._alto_fila_v()
        margen = self._margen_dias()
        cab = self._cab_horas()
        hoy = dt.datetime.now(TZ).date()
        x_fin = margen + horas * ancho_hora
        alto_todo_dia = round(15 * ESCALA)
        festivos = {f.inicio.date() for f in self._festivos
                   if f.categoria != "cumpleanos"}
        cumples = {f.inicio.date() for f in self._festivos
                  if f.categoria == "cumpleanos"}
        especiales_v = self._especiales_por_dia(ini)

        # cabecera de horas
        linea = self._linea_rejilla()
        suave = mezclar(linea, C["fondo"], 0.55) if linea else None
        ancho_paso = self._division() / 60 * ancho_hora
        for minutos, en_punto in self._marcas_horarias():
            x = margen + minutos / 60 * ancho_hora
            if linea and (en_punto or ancho_paso >= 8):
                c.create_line(x, cab, x, cab + dias * alto_fila,
                              fill=linea if en_punto else suave)
            if minutos < horas * 60 and (en_punto
                                         or ancho_paso >= round(22 * ESCALA)):
                c.create_text(x + 3, cab - round(9 * ESCALA), anchor="w",
                              text=self._etiqueta_marca(minutos),
                              fill=C["texto_tenue"],
                              font=letra(8 if en_punto else 7))

        chips = {}
        for ev, c1, c2, _fila in todo_dia:     # cuantos por dia, para repartir
            for col in range(c1, c2 + 1):
                chips.setdefault(col, []).append(ev)

        margen_base = round(86 * ESCALA)
        zona_especial = margen - margen_base   # 0 si no hay festivos/cumples

        for i in range(dias):
            fecha = (ini + dt.timedelta(days=i)).date()
            y0 = cab + i * alto_fila
            fondo = fondo_dia(i)
            if fecha == hoy:
                fondo = mezclar(C["evento"], fondo, 0.22)
            elif fecha in cumples:
                fondo = mezclar(CFG.get("color_cumpleanos", "#a855f7"),
                                fondo, 0.18)
            c.create_rectangle(margen, y0, x_fin, y0 + alto_fila,
                               fill=fondo, outline="")

            # etiqueta del dia, con el mes cuando cambia (festivos y
            # cumpleaños van en su propia cajita, al lado, no en esta etiqueta)
            etiqueta = f"{DIAS[fecha.weekday()]} {fecha.day}"
            if (i == 0 or fecha.day == 1) and margen_base >= round(84 * ESCALA):
                etiqueta += f" {MESES[fecha.month - 1]}"
            resalta = fecha == hoy
            c.create_text(margen_base - 6, y0 + alto_fila / 2, text=etiqueta,
                          anchor="e", fill=C["evento"] if resalta else C["texto"],
                          font=letra(9, "bold" if resalta else "normal"))

            # separadores: linea normal, mas marcada al cambiar semana o mes
            grosor, color = 1, linea
            if i and fecha.day == 1:
                grosor, color = max(2, round(2 * ESCALA)), C["texto_tenue"]
            elif i and fecha.weekday() == 0:
                grosor, color = max(1, round(1.5 * ESCALA)), color_semana()
            if color:
                c.create_line(0, y0, x_fin, y0, fill=color, width=grosor)

            # festivos y cumpleaños: recuadro cuadrado, pegado al lado del
            # nombre del dia, en su propio hueco reservado, antes de que
            # empiecen las horas (no dentro de la franja de eventos)
            items_esp = especiales_v.get(i, [])
            if items_esp and zona_especial > 2:
                ancho_util_esp = zona_especial - 4
                alturas = [self._medir_especial(icono, texto, ancho_util_esp)
                          for icono, texto, *_ in items_esp]
                alto_total = sum(alturas)
                y_cursor = y0 + max(2, (alto_fila - alto_total) / 2)
                for (icono, texto, color_esp, ev_esp), alto_item in zip(items_esp, alturas):
                    self._dibujar_especial(margen_base + 2, y_cursor,
                                           ancho_util_esp, icono, texto,
                                           color_esp, ev_esp)
                    y_cursor += alto_item

            # eventos de todo el dia: barritas repartidas arriba de la fila
            del_dia = chips.get(i, [])
            if del_dia:
                ancho_chip = (x_fin - margen) / len(del_dia)
                for n, ev in enumerate(del_dia):
                    x1 = margen + n * ancho_chip
                    marca_chipv = ("evento", self._registrar(ev))
                    c.create_rectangle(x1 + 1, y0 + 2, x1 + ancho_chip - 1,
                                       y0 + alto_todo_dia, outline="",
                                       fill=ev.color or C["evento"],
                                       tags=marca_chipv)
                    self._texto_en_caja(x1 + 5, y0 + 2, ancho_chip - 10,
                                        alto_todo_dia - 3, ev.titulo,
                                        fill=C["evento_texto"], tags=marca_chipv,
                                        font=self._fuente_evento(ev))

            # eventos con hora
            arriba = y0 + (alto_todo_dia + 2 if del_dia else 2)
            util = alto_fila - (arriba - y0) - 3
            fuera = [[], []]
            for ev in self._mostrables():
                if ev.todo_dia or (ev.inicio.date() - ini.date()).days != i:
                    continue
                e_ini = ev.inicio.hour + ev.inicio.minute / 60
                e_fin = ev.fin.hour + ev.fin.minute / 60
                if ev.fin.date() > ev.inicio.date() or e_fin == 0:
                    e_fin = 24.0
                if e_fin <= h0 or e_ini >= h1:
                    fuera[0 if e_fin <= h0 else 1].append(ev)
                    continue
                x1 = margen + (max(e_ini, h0) - h0) * ancho_hora
                x2 = margen + (min(e_fin, h1) - h0) * ancho_hora
                x2 = max(x2, x1 + round(16 * ESCALA))
                alto_ev = util / max(ev.ncols, 1)
                y1 = arriba + ev.col * alto_ev
                color = self._color_evento(ev)
                marca = ("evento", self._registrar(ev))
                if ev.estado == "provisional":
                    c.create_rectangle(x1 + 1, y1 + 1, x2 - 1, y1 + alto_ev - 2,
                                       fill=C["fondo"], outline=color, width=2,
                                       dash=(4, 3), tags=marca)
                elif ev.estado == "libre":
                    c.create_rectangle(x1, y1, x2, y1 + alto_ev - 1,
                                       fill=mezclar(color, C["fondo"], 0.18),
                                       outline=C["fondo"], tags=marca)
                    c.create_rectangle(x1, y1, x1 + 3, y1 + alto_ev - 1,
                                       fill=color, outline="")
                else:
                    c.create_rectangle(x1, y1, x2, y1 + alto_ev - 1, fill=color,
                                       outline=C["fondo"], tags=marca)
                etiqueta = self._etiqueta_evento(ev, x2 - x1)
                claro = ev.estado in ("provisional", "libre")
                self._texto_en_caja(x1 + 4, y1 + 1, x2 - x1 - 8, alto_ev - 3,
                                    etiqueta, font=self._fuente_evento(ev),
                                    tags=marca,
                                    fill=C["texto"] if claro else C["evento_texto"])

            for lado, lista in enumerate(fuera):
                if not lista:
                    continue
                primero = min(lista, key=lambda x: x.inicio)
                extra = f" +{len(lista) - 1}" if len(lista) > 1 else ""
                c.create_text(margen + 3 if not lado else x_fin - 3,
                              y0 + alto_fila - round(7 * ESCALA),
                              anchor="w" if not lado else "e",
                              text=(f"◀ {primero.inicio:%H:%M}{extra}" if not lado
                                    else f"{primero.inicio:%H:%M}{extra} ▶"),
                              fill=C["texto_tenue"], font=letra(7))

        if linea:
            c.create_line(0, cab + dias * alto_fila, x_fin,
                          cab + dias * alto_fila, fill=linea)

        # linea de ahora, vertical, solo en la fila de hoy
        ahora = dt.datetime.now(TZ)
        col_hoy = (ahora.date() - ini.date()).days
        if 0 <= col_hoy < dias and h0 <= ahora.hour < h1:
            x = margen + (ahora.hour + ahora.minute / 60 - h0) * ancho_hora
            y0 = cab + col_hoy * alto_fila
            c.create_line(x, y0, x, y0 + alto_fila, fill=C["ahora"], width=2)
            c.create_oval(x - 3, y0 - 3, x + 3, y0 + 3, fill=C["ahora"],
                          outline="")

    # -- reparto de vistas ---------------------------------------------------

    def dibujar(self):
        ini, fin = limites_ventana(desplazamiento=self.desplazamiento)
        self._quitar_reserva()
        if self.vista == "mes":
            # el mes ocupa toda la ventana: sin desplazamiento de rejilla
            self._off_rejilla = (0, 0)
            # el mes necesita sus propios festivos, que abarcan mas dias
            ini_mes, fin_mes = ini - dt.timedelta(days=40), ini + dt.timedelta(days=50)
            self._festivos = (generar_festivos(ini_mes, fin_mes)
                              + generar_cumpleanos(ini_mes, fin_mes))
            self._dibujar_mes(ini)
            return
        self._festivos = generar_festivos(ini, fin) + generar_cumpleanos(ini, fin)
        con_rejilla = self.vista in ("rejilla", "ambas")
        con_lista = self.vista in ("lista", "ambas")

        todo_dia = []
        ancho_r = alto_r = ancho_l = alto_l = 0
        vertical = con_rejilla and self.orientacion == "vertical"
        self._hay_especiales_v = False
        # to-dos y correo dentro del ancho actual: el calendario cede sitio
        cols_ancho = sum(self._anchos_columnas())
        if cols_ancho and CFG.get("columnas_dentro", True):
            lado0 = self._lado_lista() if (con_rejilla and con_lista) \
                else "derecha"
            self._reservar_para_columnas(
                cols_ancho, con_rejilla, con_lista, vertical,
                lado0 in ("arriba", "abajo"))
        if con_rejilla:
            todo_dia, n_filas = self._colocar_todo_dia(ini)
            if vertical:
                self._hay_especiales_v = bool(self._especiales_por_dia(ini))
                self._alto_banda = n_filas * ALTO_CHIP + 4 if n_filas else BANDA_MINIMA
                ancho_r, alto_r = self._medidas_vertical()
            else:
                _, alto_esp_max = self._medidas_especiales_horizontal(ini)
                alto_todo = n_filas * ALTO_CHIP + alto_esp_max
                self._alto_banda = (alto_todo + 4 if alto_todo else BANDA_MINIMA)
                ancho_r, alto_r = self._ancho_rejilla(), self._alto_rejilla()

        filas = self._filas_lista(ini) if con_lista else []
        if con_lista:
            ancho_l = self._ancho_lista()
            alto_l = min(Y_LISTA + max(self._alto_filas(filas, ancho_l), 40)
                         + ALTO_ESTADO + 4, CFG["alto_max_lista"])

        lado = self._lado_lista() if (con_rejilla and con_lista) else "derecha"
        apilada = lado in ("arriba", "abajo")
        # columna de to-dos acoplada a la derecha (independiente de la vista)
        con_todos = bool(CFG.get("todos_visible")) and self.vista != "mes"
        ancho_t = self._ancho_todos() if con_todos else 0
        # columna de correo, a la derecha de todo (independiente de la vista)
        con_correo = bool(CFG.get("correo_visible")) and self.vista != "mes"
        ancho_c = self._ancho_correo() if con_correo else 0
        banda = 0
        if apilada:
            # la lista ocupa todo el ancho, encima o debajo del calendario
            banda = min(max(self._alto_filas(filas, ancho_r) + 10, 60),
                        CFG["alto_max_lista"])
            ancho = ancho_r
            alto = alto_r + banda + 6 + (ALTO_ESTADO if lado == "abajo" else 0)
            self._off_rejilla = (0, banda + 6 if lado == "arriba" else 0)
        else:
            ancho, alto = ancho_r + ancho_l, alto_r
            if self.vista != "ambas":
                alto = max(alto_r, alto_l)
            self._off_rejilla = (ancho_l if lado == "izquierda" else 0, 0)
        if self._sobra_col:
            if con_correo:
                ancho_c += self._sobra_col
            elif con_todos:
                ancho_t += self._sobra_col
        if con_todos:
            ancho += ancho_t
            alto = max(alto, self._alto_todos_columna())
        if con_correo:
            ancho += ancho_c
            alto = max(alto, self._alto_correo_columna())
        self._ajustar(ancho, alto)

        self.lienzo.delete("all")
        self._info_eventos.clear()
        self._n_evento = 0
        self._dibujar_barra(ancho, self._titulo_rango(ini))
        if con_rejilla:
            antes = set(self.lienzo.find_all())
            if vertical:
                self._pintar_vertical(ini, todo_dia)
            else:
                self._pintar_rejilla(ini, todo_dia)
            self._desplazar_nuevos(antes, *self._off_rejilla)
        if con_lista:
            if apilada:
                antes = set(self.lienzo.find_all())
                self._pintar_lista(ini, filas, 0, ancho - ancho_t - ancho_c,
                                   Y_LISTA + banda + ALTO_ESTADO)
                if lado == "abajo":
                    self._desplazar_nuevos(antes, 0, alto_r - Y_LISTA + 6)
                    self._divisor_horizontal(alto_r, ancho - ancho_t - ancho_c)
                else:
                    self._divisor_horizontal(banda, ancho - ancho_t - ancho_c)
            else:
                self._pintar_lista(ini, filas,
                                   0 if lado == "izquierda" else ancho_r,
                                   ancho_l, alto)
                if lado == "izquierda":
                    self._divisor_vertical(ancho_l, alto)
        if con_todos:
            self._pintar_todos_columna(ancho - ancho_c - ancho_t, ancho_t, alto)
        else:
            self._ocultar_entrada_todos()
        if con_correo:
            self._pintar_correo_columna(ancho - ancho_c, ancho_c, alto)
        self._dibujar_estado(alto - ALTO_ESTADO + 2, ancho)
        self._dibujar_agarradera(ancho, alto)

    @staticmethod
    def _lado_lista():
        lado = CFG.get("lado_lista", "derecha")
        return lado if lado in ("derecha", "izquierda", "arriba", "abajo") \
            else "derecha"

    def _desplazar_nuevos(self, antes, dx, dy):
        """Mueve lo recién dibujado (así el calendario puede ir a otro sitio)."""
        if not dx and not dy:
            return
        for item in self.lienzo.find_all():
            if item not in antes:
                self.lienzo.move(item, dx, dy)

    def _grosor_divisor(self):
        """Ancho de la franja agarrable del divisor, en pixeles de pantalla."""
        return max(5, round(CFG.get("grosor_divisor", 11) * ESCALA))

    def _ancho_todos(self):
        return round(CFG.get("ancho_todos", 210) * ESCALA)

    # -- columna de correo (correo_columna.py) --------------------------------

    def _ancho_correo(self):
        return round(CFG.get("ancho_correo", 260) * ESCALA)

    def _alto_correo_columna(self):
        try:
            import correo_columna
            return correo_columna.alto_deseado(self, sys.modules[__name__])
        except Exception as exc:
            log(f"Columna de correo (alto): {exc!r}")
            return 0

    def _pintar_correo_columna(self, x0, ancho, alto):
        try:
            import correo_columna
            correo_columna.pintar(self, sys.modules[__name__], x0, ancho, alto)
        except Exception as exc:
            import traceback
            log("Columna de correo: " + traceback.format_exc().strip()
                .replace("\n", " | "))

    def _abrir_correo_columna(self, e=None):
        """Enseña o esconde la columna de correo dentro del widget."""
        try:
            import correo_columna
            return correo_columna.alternar_columna(self, sys.modules[__name__])
        except Exception as exc:
            log(f"No se pudo cargar correo_columna.py: {exc!r}")
            return "break"

    def _construir_menu_correo_columna(self, padre):
        try:
            import correo_columna
            return correo_columna.construir_menu(
                self, sys.modules[__name__], padre)
        except Exception as exc:
            log(f"Menú de la columna de correo: {exc!r}")
            return tk.Menu(padre, tearoff=0)

    def _inicio_divisor_correo(self, e):
        self._divisor_c = {"x0": e.x_root, "ancho": self._ancho_correo()}
        self.lienzo.config(cursor="sb_h_double_arrow")
        return "break"

    def _reset_ancho_correo(self):
        CFG["ancho_correo"] = DEFAULTS["ancho_correo"]
        guardar_config()
        self._tam = None
        self.dibujar()

    def _menu_ancho_correo(self, e=None):
        self._abrir_menu(self._construir_menu_correo_columna(self.lienzo), e)
        return "break"

    def _alto_todos_columna(self):
        """Alto que le gustaria a la columna de to-dos, para que la ventana
        crezca si hace falta (como hace la lista)."""
        n = len(self._todos_datos())
        fila = round(20 * ESCALA)
        return Y_LISTA + (n + 1) * fila + round(10 * ESCALA) + ALTO_ESTADO + 4

    def _todos_ordenados(self):
        """Devuelve [(indice_real, td), ...] con los pendientes arriba y los
        ya hechos al fondo, conservando el orden dentro de cada grupo."""
        datos = self._todos_datos()
        pendientes = [(i, t) for i, t in enumerate(datos) if not t.get("hecho")]
        hechos = [(i, t) for i, t in enumerate(datos) if t.get("hecho")]
        return pendientes + hechos

    def _pintar_todos_columna(self, x0, ancho, alto_total):
        """Dibuja la columna de to-dos acoplada, con el mismo estilo que la
        lista. Cada to-do: punto de color (etiqueta) + casilla + texto + ✕.
        Los pendientes van arriba; los hechos, al fondo y en gris. Debajo (o
        arriba si no hay ninguno) hay un campo fijo para añadir con Enter."""
        c = self.lienzo
        if x0:
            self._divisor_todos(x0, alto_total)

        # cabecera de la columna
        c.create_text(x0 + 10, Y_LISTA + 2, text="To-dos", anchor="nw",
                      font=letra(9, "bold"), fill=C["texto"])
        hay_hechos = any(t.get("hecho") for t in self._todos_datos())
        if hay_hechos:
            limpiar = c.create_text(x0 + ancho - 10, Y_LISTA + 2, text="🧹",
                                    anchor="ne", font=letra(9),
                                    fill=C["texto_tenue"])
            c.tag_bind(limpiar, "<Button-1>",
                       lambda e: self._todo_limpiar_hechos())
        c.create_line(x0 + 8, Y_LISTA + 20, x0 + ancho - 8, Y_LISTA + 20,
                      fill=C["rejilla"])

        orden = self._todos_ordenados()
        y = Y_LISTA + 26
        alto_entrada = round(24 * ESCALA)
        tope = alto_total - ALTO_ESTADO - alto_entrada - 4
        recortados = 0
        for pos, (i, td) in enumerate(orden):
            alto_f = self._alto_todo_fila(td, ancho)
            if y + alto_f > tope:
                recortados = len(orden) - pos
                break
            self._pintar_todo_fila(x0, y, ancho, i, td, alto_f)
            y += alto_f

        if not orden:
            c.create_text(x0 + ancho / 2, y + 6,
                          text="Escribe abajo para añadir", fill=C["texto_tenue"],
                          font=letra(8))
            y += round(22 * ESCALA)
        if recortados:
            c.create_text(x0 + ancho / 2, y + 4,
                          text=f"(+{recortados} más)", fill=C["texto_tenue"],
                          font=letra(7))
            y += round(14 * ESCALA)

        # campo fijo para añadir: justo debajo del ultimo to-do, sin pasarse
        y_entrada = min(y + 2, tope + 2)
        self._colocar_entrada_todos(x0, y_entrada, ancho)

    def _colocar_entrada_todos(self, x0, y, ancho):
        """Coloca (creandolo la primera vez) el campo fijo de 'nuevo to-do'
        como un widget real sobre el lienzo, para escribir y pulsar Enter."""
        c = self.lienzo
        if getattr(self, "_entrada_todos", None) is None:
            fila = tk.Frame(c, bg=C["fondo"])
            signo = tk.Label(fila, text="＋", bg=C["fondo"], fg=C["texto_tenue"],
                             font=letra(10))
            signo.pack(side="left")
            entrada = tk.Entry(fila, bg=C.get("boton", "#2f2f33"),
                               fg=C["texto"], insertbackground=C["texto"],
                               relief="flat", font=letra(8))
            entrada.pack(side="left", fill="x", expand=True, ipady=2,
                         padx=(4, 0))
            entrada.bind("<Return>", lambda e: self._añadir_desde_entrada())
            self._entrada_todos = fila
            self._entrada_todos_campo = entrada
        self._entrada_todos_ventana = c.create_window(
            x0 + 6, y, anchor="nw", window=self._entrada_todos,
            width=ancho - 12)

    def _añadir_desde_entrada(self):
        campo = getattr(self, "_entrada_todos_campo", None)
        if campo is None:
            return
        texto = campo.get().strip()
        if not texto:
            return
        self._todos_datos().append({"texto": texto, "hecho": False,
                                    "etiqueta": ""})
        guardar_config()
        self.dibujar()
        # el campo persiste; lo vaciamos y dejamos el foco para seguir metiendo
        try:
            campo.delete(0, "end")
            campo.focus_set()
        except Exception:
            pass

    def _ocultar_entrada_todos(self):
        """Esconde el campo de añadir cuando la columna no se muestra."""
        fila = getattr(self, "_entrada_todos", None)
        if fila is not None:
            try:
                fila.place_forget()
                self.lienzo.delete(self._entrada_todos_ventana)
            except Exception:
                pass

    def _alto_todo_fila(self, td, ancho):
        """Alto de una fila segun cuanto ocupe su texto (una o dos lineas)."""
        fuente = letra(8)
        ancho_txt = ancho - round(70 * ESCALA)
        lineas = min(2, self._lineas_envueltas(td.get("texto", ""),
                                               ancho_txt, fuente))
        return max(round(20 * ESCALA), lineas * self._alto_linea(fuente) + 6)

    def _pintar_todo_fila(self, x0, y, ancho, i, td, alto_f):
        c = self.lienzo
        hecho = bool(td.get("hecho"))
        color = color_de_categoria(td.get("etiqueta")) or C["texto_tenue"]

        # punto de color = etiqueta (clic abre el menu de etiquetas)
        punto = c.create_text(x0 + 12, y + alto_f / 2, text="●", anchor="w",
                              fill=color, font=letra(9))
        c.tag_bind(punto, "<Button-1>",
                   lambda e, n=i: self._todo_menu_etiqueta(e, n))

        # casilla (clic marca/desmarca)
        casilla = c.create_text(x0 + 26, y + alto_f / 2,
                                text="☑" if hecho else "☐", anchor="w",
                                fill=C["texto"] if hecho else C["texto_tenue"],
                                font=letra(11))
        c.tag_bind(casilla, "<Button-1>", lambda e, n=i: self._todo_alternar(n))

        # texto (clic marca/desmarca; doble clic edita en linea)
        tipo = "overstrike" if hecho else "normal"
        tx = x0 + 56
        tw = ancho - round(70 * ESCALA)
        txt = self._texto_en_caja(
            tx, y + 2, tw, alto_f - 4,
            td.get("texto", ""), font=letra(8, tipo),
            fill=C["texto_tenue"] if hecho else C["texto"])
        c.tag_bind(txt, "<Button-1>", lambda e, n=i: self._todo_alternar(n))
        c.tag_bind(txt, "<Double-Button-1>",
                   lambda e, n=i, X=tx, Y=y, W=tw, H=alto_f:
                       self._editar_todo_inline(n, X, Y, W, H))

        # ✕ para borrar
        quita = c.create_text(x0 + ancho - 10, y + alto_f / 2, text="✕",
                              anchor="e", fill=C["texto_tenue"], font=letra(8))
        c.tag_bind(quita, "<Button-1>", lambda e, n=i: self._todo_quitar(n))

    def _divisor_todos(self, x, alto):
        """Borde izquierdo de la columna de to-dos, arrastrable para cambiar
        su ancho (con su propia etiqueta, distinta del divisor de la lista)."""
        c = self.lienzo
        y0, y1 = BARRA, alto - ALTO_ESTADO
        media = self._grosor_divisor() / 2
        c.create_rectangle(x - media, y0, x + media, y1,
                           fill=C["fondo"], outline="", tags=("divisor_todos",))
        c.create_line(x, y0, x, y1, fill=C["texto_tenue"],
                      tags=("divisor_todos", "divisor_todos_tinta"))
        self._tirador(x, (y0 + y1) / 2, ("divisor_todos", "divisor_todos_tinta"))

    def _tirador(self, x, medio, tags):
        """Tres puntitos en la raya: se ve que se puede agarrar y arrastrar."""
        c = self.lienzo
        paso = max(5, round(6 * ESCALA))
        radio = max(1, round(1.5 * ESCALA))
        for d in (-paso, 0, paso):
            c.create_oval(x - radio, medio + d - radio, x + radio,
                          medio + d + radio, fill=C["texto_tenue"], outline="",
                          tags=tags)

    def _divisor_vertical(self, x, alto, y0=None):
        """Separacion entre calendario y lista cuando estan uno al lado del otro."""
        c = self.lienzo
        y0 = BARRA if y0 is None else y0
        y1 = alto - ALTO_ESTADO
        media = self._grosor_divisor() / 2
        # franja invisible mas ancha que la raya, para poder agarrarla sin puntería
        c.create_rectangle(x - media, y0, x + media, y1,
                           fill=C["fondo"], outline="",
                           tags=("divisor", "divisor_v"))
        c.create_line(x, y0, x, y1, fill=C["texto_tenue"],
                      tags=("divisor", "divisor_v", "divisor_tinta"))
        medio = (y0 + y1) / 2
        paso = max(5, round(6 * ESCALA))
        radio = max(1, round(1.5 * ESCALA))
        for d in (-paso, 0, paso):            # tres puntitos, como un tirador
            c.create_oval(x - radio, medio + d - radio,
                          x + radio, medio + d + radio,
                          fill=C["texto_tenue"], outline="",
                          tags=("divisor", "divisor_v", "divisor_tinta"))

    def _divisor_horizontal(self, y, ancho):
        """Separacion cuando la lista va encima o debajo del calendario."""
        c = self.lienzo
        media = self._grosor_divisor() / 2
        c.create_rectangle(0, y - media, ancho, y + media,
                           fill=C["fondo"], outline="",
                           tags=("divisor", "divisor_h"))
        c.create_line(4, y, ancho - 4, y, fill=C["texto_tenue"],
                      tags=("divisor", "divisor_h", "divisor_tinta"))
        paso = max(5, round(6 * ESCALA))
        radio = max(1, round(1.5 * ESCALA))
        for d in (-paso, 0, paso):
            c.create_oval(ancho / 2 + d - radio, y - radio,
                          ancho / 2 + d + radio, y + radio,
                          fill=C["texto_tenue"], outline="",
                          tags=("divisor", "divisor_h", "divisor_tinta"))

    def _resaltar_divisor(self, dentro):
        """Al pasar por encima, el tirador se enciende: se ve que se agarra."""
        if self._divisor:
            return                    # ya se esta arrastrando: no tocar nada
        try:
            self.lienzo.itemconfig(
                "divisor_tinta", fill=C["texto"] if dentro else C["texto_tenue"])
        except tk.TclError:
            pass

    def _cursor_columna(self, tag, dentro):
        """Cursor de flechas y tirador encendido en el borde de una columna."""
        if self._divisor_c or self._divisor_t:
            return                    # arrastrando: no tocar nada
        self.lienzo.config(cursor="sb_h_double_arrow" if dentro else "")
        try:
            self.lienzo.itemconfig(
                tag + "_tinta", fill=C["evento"] if dentro else C["texto_tenue"])
        except tk.TclError:
            pass

    def _dibujar_arrastre(self):
        """Repinta mientras se arrastra un borde, pero sin pasarse: como mucho
        unas 30 veces por segundo (el resto de movimientos se juntan)."""
        ahora = time.monotonic()
        if ahora - getattr(self, "_t_arrastre", 0) >= 0.033:
            self._t_arrastre = ahora
            self._pend_arrastre = False
            self.dibujar()
        elif not getattr(self, "_pend_arrastre", False):
            self._pend_arrastre = True
            self.after(35, self._dibujar_arrastre_final)
        return "break"

    def _dibujar_arrastre_final(self):
        self._pend_arrastre = False
        self._t_arrastre = time.monotonic()
        self.dibujar()

    def _cursor_divisor(self, e=None, vertical=True):
        self.lienzo.config(
            cursor="sb_h_double_arrow" if vertical else "sb_v_double_arrow")
        self._resaltar_divisor(True)

    def _reparto_por_defecto(self, e=None):
        """Doble clic en el divisor: vuelve al reparto de fabrica."""
        if self._lado_lista() in ("arriba", "abajo"):
            CFG["alto_max_lista"] = DEFAULTS["alto_max_lista"]
        else:
            CFG["ancho_lista"] = DEFAULTS["ancho_lista"]
        self._divisor = None
        self._tam = None
        self._hueco = None
        guardar_config()
        self.dibujar()
        return "break"

    def _cambiar_grosor_divisor(self, valor):
        CFG["grosor_divisor"] = max(5, min(30, int(valor)))
        guardar_config()
        self.dibujar()

    def _construir_menu_divisor(self, padre):
        m = tk.Menu(padre, tearoff=0)
        actual = CFG.get("grosor_divisor", 11)
        for valor, etiqueta in ((7, "Fina"), (11, "Normal"),
                                (15, "Ancha"), (21, "Muy ancha")):
            m.add_command(
                label=("● " if actual == valor else "     ")
                      + f"{etiqueta} ({valor} px)",
                command=lambda v=valor: self._cambiar_grosor_divisor(v))
        m.add_separator()
        m.add_command(label="Volver al reparto por defecto",
                      command=self._reparto_por_defecto)
        m.add_command(label="(doble clic en la barra hace lo mismo)",
                      state="disabled")
        return m

    def _menu_divisor(self, e=None):
        """Clic derecho sobre la propia barra separadora."""
        self._abrir_menu(self._construir_menu_divisor(self.lienzo), e)
        return "break"


def instancia_unica():
    """Una sola ventana. Si ya hay una, este arranque solo la muestra u oculta."""
    try:
        import ctypes
        mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "WidgetCalendarioICS")
        if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            pedir_alternar()
            log("Ya habia una instancia: se le pide mostrar/ocultar y esta se cierra.")
            sys.exit(0)
        return mutex          # se mantiene vivo mientras dure el proceso
    except Exception:
        return None


# el menú de clic derecho, moderno y con el tema del widget (menus.py)
try:
    import menus
    menus.instalar(sys.modules[__name__])
except Exception as _exc:          # sin él, queda el menú normal de Windows
    print(f"Menú moderno no disponible: {_exc!r}")


if __name__ == "__main__":
    _mutex = instancia_unica()
    try:  # nitidez en pantallas con escalado distinto de 100%
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    consumir_alternar()          # descarta una senal vieja del arranque anterior
    log(f"Calendario {VERSION} arrancado")
    # los cierres sin aviso dejaban el registro mudo: ahora quedan sus rastros
    import faulthandler
    import traceback
    _fallos = open(os.path.join(LOCAL, "cierres.txt"), "a", encoding="utf-8",
                   buffering=1)       # aparte: log() rota registro.txt
    faulthandler.enable(_fallos, all_threads=True)    # cuelgue o cierre brusco
    _ventana = Widget()

    def _latido():
        # si la ventana tarda más de 4 s en volver aquí (se pone «blanquecina»),
        # se vuelca dónde está parada cada hilo; cada latido rearma el aviso
        faulthandler.dump_traceback_later(4, repeat=True, file=_fallos)
        _ventana.after(1000, _latido)
    _latido()
    sys.excepthook = lambda t, v, tb: log(
        "ERROR sin controlar: " + " | ".join(
            "".join(traceback.format_exception(t, v, tb)).splitlines()))
    threading.excepthook = lambda a: log(
        f"ERROR en un hilo ({a.thread.name if a.thread else '?'}): " + " | ".join(
            "".join(traceback.format_exception(
                a.exc_type, a.exc_value, a.exc_traceback)).splitlines()))
    _ventana.mainloop()
