"""Sincronización con Post-It por una carpeta compartida (la misma lógica que sync.js de Post-It).

  notas-<equipo>.json   lo escribe Post-It: las líneas de sus notas que llevan hora
  agenda-<equipo>.json  lo escribe este Calendario: los eventos de los próximos días

Cada equipo escribe solo su fichero (así OneDrive no crea copias en conflicto) y todos leen
todos los de su tipo. Las horas van como texto ISO local, sin zona.
Este módulo no importa calendario.py: recibe lo que necesita (zona horaria, la clase Evento).
"""
import datetime as dt
import json
import os
import re
import socket

NOMBRE_CARPETA = "AgendaCompartida"
PREFIJO_POSTIT = "postit:"      # los eventos que vienen de Post-It no se devuelven a Post-It
DIAS_AGENDA = 31
DIAS_ATRAS = 7                  # las notas unidas también muestran los eventos de días ya pasados


def equipo():
    return re.sub(r"[^A-Za-z0-9_-]", "_", socket.gethostname())[:40] or "PC"


def resolver_carpeta(guardada=None, raiz=None, entorno=None, hogar=None):
    """(ruta, origen). Orden: la guardada · AGENDA_COMPARTIDA · la hermana de la app, en OneDrive o en
    Documentos · se crea junto a la app. Ruta None si no se puede."""
    entorno = os.environ if entorno is None else entorno
    hogar = hogar or os.path.expanduser("~")
    if guardada and os.path.isdir(guardada):
        return guardada, "guardada"
    if entorno.get("AGENDA_COMPARTIDA") and os.path.isdir(entorno["AGENDA_COMPARTIDA"]):
        return entorno["AGENDA_COMPARTIDA"], "entorno"
    candidatas = []
    if raiz:
        padre = os.path.dirname(raiz)
        candidatas += [os.path.join(padre, NOMBRE_CARPETA), os.path.join(os.path.dirname(padre), NOMBRE_CARPETA)]
    candidatas += [os.path.join(entorno[k], NOMBRE_CARPETA)
                   for k in ("OneDrive", "OneDriveConsumer", "OneDriveCommercial") if entorno.get(k)]
    candidatas += [os.path.join(hogar, d, NOMBRE_CARPETA) for d in ("OneDrive", "Documents", "Documentos")]
    for c in candidatas:
        if os.path.isdir(c):
            return c, "detectada"
    nueva = os.path.join(os.path.dirname(raiz), NOMBRE_CARPETA) if raiz else os.path.join(hogar, NOMBRE_CARPETA)
    try:
        os.makedirs(nueva, exist_ok=True)
        return nueva, "creada"
    except OSError:
        return None, "ninguna"


def datos_compartidos(entorno=None, hogar=None):
    """Carpeta «Calendario/datos» de OneDrive si ya existe (con su config.json): por ahí se comparten la configuración y los
    perfiles entre equipos. Un .exe vive en local, así que sin esto sus datos no se sincronizarían. None si no hay."""
    entorno = os.environ if entorno is None else entorno
    hogar = hogar or os.path.expanduser("~")
    raices = [entorno[k] for k in ("OneDrive", "OneDriveConsumer", "OneDriveCommercial") if entorno.get(k)]
    for r in raices + [os.path.join(hogar, "OneDrive")]:
        c = os.path.join(r, "Calendario", "datos")
        if os.path.isfile(os.path.join(c, "config.json")):
            return c
    return None


def _ficheros(carpeta, prefijo):
    try:
        nombres = sorted(n for n in os.listdir(carpeta) if n.startswith(prefijo + "-") and n.endswith(".json"))
    except OSError:
        return []
    salida = []
    for n in nombres:
        try:
            with open(os.path.join(carpeta, n), encoding="utf-8") as f:
                datos = json.load(f)
            if isinstance(datos, dict):
                salida.append(datos)
        except (OSError, ValueError):      # a medio sincronizar o roto: se ignora y se lee la próxima vez
            pass
    return salida


def firma_notas(carpeta):
    """Cambia cuando cambia algún notas-*.json: para no releerlos si nada se ha movido."""
    try:
        return "|".join(n + str(os.stat(os.path.join(carpeta, n)).st_mtime_ns)
                        for n in sorted(os.listdir(carpeta)) if n.startswith("notas-"))
    except OSError:
        return ""


def _local(texto, tz):
    try:
        return dt.datetime.fromisoformat(texto).replace(tzinfo=tz)
    except (TypeError, ValueError):
        return None


def leer_notas(carpeta, tz, crear_evento, color=None):
    """Los items de Post-It como eventos. crear_evento(inicio, fin, titulo, todo_dia, uid, **extra)."""
    vistos, eventos = set(), []
    for f in _ficheros(carpeta, "notas"):
        for it in f.get("items") or []:
            ini = _local(it.get("inicio"), tz)
            uid = str(it.get("id") or "")
            if ini is None or not uid or uid in vistos or not str(it.get("titulo") or "").strip():
                continue
            vistos.add(uid)
            todo_dia = bool(it.get("todo_dia"))     # tarea apuntada bajo una fecha, sin hora
            fin = _local(it.get("fin"), tz)
            if todo_dia:
                ini = ini.replace(hour=0, minute=0, second=0)
                fin = ini + dt.timedelta(days=1)
            elif fin is None or fin <= ini:
                fin = ini + dt.timedelta(hours=1)
            titulo = ("✓ " if it.get("hecho") else "") + str(it["titulo"]).strip()
            eventos.append(crear_evento(ini, fin, titulo, todo_dia, uid,
                                        categoria="postit", color=color, notas="Viene de una nota de Post-It"))
    return eventos


def _iso_local(momento):
    return momento.replace(tzinfo=None, microsecond=0).isoformat()


def escribir_agenda(carpeta, eventos, ahora, nombre=None, dias=DIAS_AGENDA):
    """Escribe agenda-<equipo>.json con los eventos de los próximos días. Devuelve True si cambió el fichero.
    Se saltan los que vienen de Post-It y los trozos de continuación de un evento largo."""
    nombre = nombre or equipo()
    hoy = ahora.replace(hour=0, minute=0, second=0, microsecond=0)
    desde, hasta = hoy - dt.timedelta(days=DIAS_ATRAS), hoy + dt.timedelta(days=dias)
    lista, vistos = [], set()
    for e in sorted(eventos, key=lambda x: x.inicio):
        uid = str(e.uid or e.titulo)
        if uid.startswith(PREFIJO_POSTIT) or getattr(e, "parte", 0) or not desde <= e.inicio < hasta:
            continue
        id_ = f"{uid}@{_iso_local(e.inicio)}"        # un evento repetido tiene el mismo uid cada día
        if id_ in vistos:
            continue
        vistos.add(id_)
        lista.append({"id": id_, "titulo": e.titulo, "inicio": _iso_local(e.inicio), "fin": _iso_local(e.fin),
                      "todo_dia": bool(e.todo_dia), "origen": "calendario"})
    nuevo = {"version": 1, "equipo": nombre, "eventos": lista}
    ruta = os.path.join(carpeta, f"agenda-{nombre}.json")
    try:
        with open(ruta, encoding="utf-8") as f:
            if json.load(f) == nuevo:
                return False
    except (OSError, ValueError):
        pass
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(nuevo, f, ensure_ascii=False, indent=1)
    os.replace(tmp, ruta)
    return True
