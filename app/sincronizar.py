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
import time

NOMBRE_CARPETA = "Carpeta de Sincronización"     # una sola carpeta (OneDrive) para Calendario y Post-It; cada equipo escribe solo sus ficheros
NOMBRE_ANTIGUO = "AgendaCompartida"              # como se llamaba antes: se sigue usando si es la que ya existe
SUBCARPETA_DATOS = "Calendario"                  # dentro de la carpeta de sincronización: config.json y perfiles/ compartidos
PREFIJO_POSTIT = "postit:"      # los eventos que vienen de Post-It no se devuelven a Post-It
DIAS_AGENDA = 31
DIAS_ATRAS = 7                  # las notas unidas también muestran los eventos de días ya pasados


def version_mayor(a, b):
    """¿La versión «a» (2.4.1) es más nueva que «b» (2.4.0)? Compara número a número; lo que no sea número cuenta como 0."""
    def partes(v):
        return [int(x) if x.isdigit() else 0 for x in str(v).strip().split(".")]
    x, y = partes(a), partes(b)
    n = max(len(x), len(y))
    return x + [0] * (n - len(x)) > y + [0] * (n - len(y))


def equipo():
    return re.sub(r"[^A-Za-z0-9_-]", "_", socket.gethostname())[:40] or "PC"


def raices_onedrive(entorno=None, hogar=None):
    """Carpetas raíz de OneDrive de este equipo, la personal antes que la de la organización ($env:OneDrive apunta a la de la
    organización cuando hay una)."""
    entorno = os.environ if entorno is None else entorno
    hogar = hogar or os.path.expanduser("~")
    raices = [entorno[k] for k in ("OneDriveConsumer", "OneDrive", "OneDriveCommercial") if entorno.get(k)]
    raices.append(os.path.join(hogar, "OneDrive"))
    vistas, salida = set(), []
    for r in raices:
        if r.lower() not in vistas:
            vistas.add(r.lower())
            salida.append(r)
    return salida


def _leer_json(ruta):
    try:
        with open(ruta, encoding="utf-8") as f:
            datos = json.load(f)
        return datos if isinstance(datos, dict) else {}
    except (OSError, ValueError):
        return {}


def carpeta_elegida(local=None, entorno=None):
    """La carpeta de sincronización de este equipo: la que dejó el instalador o se eligió en la app (sync.json en la carpeta
    local), y si no, la de Post-It (%APPDATA%\\post-it\\sync.json), para que las dos apps compartan la misma. None si no hay."""
    entorno = os.environ if entorno is None else entorno
    fichas = []
    if local:
        fichas.append(os.path.join(local, "sync.json"))
    if entorno.get("APPDATA"):
        fichas.append(os.path.join(entorno["APPDATA"], "post-it", "sync.json"))
    for f in fichas:
        ruta = _leer_json(f).get("sync")
        if ruta and os.path.isdir(ruta):
            return ruta
    return None


def guardar_carpeta_elegida(local, ruta):
    """Recuerda (solo en este equipo) qué carpeta de sincronización usa; None = sin sincronización."""
    f = os.path.join(local, "sync.json")
    datos = _leer_json(f)
    datos["sync"] = ruta or None
    tmp = f + ".tmp"
    with open(tmp, "w", encoding="utf-8") as h:
        json.dump(datos, h, ensure_ascii=False)
    os.replace(tmp, f)


def resolver_carpeta(guardada=None, raiz=None, entorno=None, hogar=None, local=None):
    """(ruta, origen). Orden: la guardada en la configuración · AGENDA_COMPARTIDA · la elegida en este equipo (o la de Post-It) ·
    «Carpeta de Sincronización» en OneDrive · la antigua «AgendaCompartida» si es la que existe · se crea «Carpeta de
    Sincronización» en OneDrive (o junto a la app). Ruta None si no se puede."""
    entorno = os.environ if entorno is None else entorno
    hogar = hogar or os.path.expanduser("~")
    if guardada and os.path.isdir(guardada):
        return guardada, "guardada"
    if entorno.get("AGENDA_COMPARTIDA") and os.path.isdir(entorno["AGENDA_COMPARTIDA"]):
        return entorno["AGENDA_COMPARTIDA"], "entorno"
    elegida = carpeta_elegida(local, entorno) if local is not None else None
    if elegida:
        return elegida, "elegida"
    raices = raices_onedrive(entorno, hogar) + [os.path.join(hogar, d) for d in ("Documents", "Documentos")]
    hermanas = []                       # junto a la app y un nivel más arriba (la carpeta de la app puede ser Calendario\app)
    if raiz:
        padre = os.path.dirname(raiz)
        hermanas = [padre, os.path.dirname(padre)]
    for r in hermanas + raices:
        c = os.path.join(r, NOMBRE_CARPETA)
        if os.path.isdir(c):
            return c, "detectada"
    for r in hermanas + raices:
        c = os.path.join(r, NOMBRE_ANTIGUO)
        if os.path.isdir(c):
            return c, "antigua"
    onedrive = next((r for r in raices_onedrive(entorno, hogar) if os.path.isdir(r)), None)
    nueva = os.path.join(onedrive, NOMBRE_CARPETA) if onedrive else (
        os.path.join(os.path.dirname(raiz), NOMBRE_CARPETA) if raiz else os.path.join(hogar, NOMBRE_CARPETA))
    try:
        os.makedirs(nueva, exist_ok=True)
        return nueva, "creada"
    except OSError:
        return None, "ninguna"


def datos_en_sync(sync, origenes=()):
    """«<sync>/Calendario»: ahí viven config.json y perfiles/ cuando hay carpeta de sincronización, para que los vean todos los
    equipos. La primera vez, si aún no hay config.json, se llevan los de «origenes» (los de antes: junto a la app o en
    OneDrive\\Calendario\\datos), que no se tocan. None si no se puede."""
    import shutil
    destino = os.path.join(sync, SUBCARPETA_DATOS)
    try:
        os.makedirs(destino, exist_ok=True)
        if not os.path.isfile(os.path.join(destino, "config.json")):
            for o in origenes:
                if o and os.path.isfile(os.path.join(o, "config.json")) and os.path.abspath(o) != os.path.abspath(destino):
                    shutil.copy2(os.path.join(o, "config.json"), os.path.join(destino, "config.json"))
                    if os.path.isdir(os.path.join(o, "perfiles")) and not os.path.isdir(os.path.join(destino, "perfiles")):
                        shutil.copytree(os.path.join(o, "perfiles"), os.path.join(destino, "perfiles"))
                    break
        return destino
    except OSError:
        return None


def esperar(listo, segundos=40, pausa=1):
    """Espera a que listo() sea cierto, como mucho «segundos». True si lo consiguió; nunca bloquea más."""
    fin = time.monotonic() + segundos
    while not listo():
        if time.monotonic() >= fin:
            return False
        time.sleep(pausa)
    return True


def datos_compartidos(entorno=None, hogar=None, espera=0):
    """Carpeta «Calendario/datos» de OneDrive si ya existe (con su config.json): por ahí se comparten la configuración y los
    perfiles entre equipos. Un .exe vive en local, así que sin esto sus datos no se sincronizarían. None si no hay.
    Con «espera» (segundos): si OneDrive ya trajo la carpeta Calendario pero aún no su config.json, espera a que llegue."""
    entorno = os.environ if entorno is None else entorno
    hogar = hogar or os.path.expanduser("~")
    raices = [entorno[k] for k in ("OneDrive", "OneDriveConsumer", "OneDriveCommercial") if entorno.get(k)]
    raices.append(os.path.join(hogar, "OneDrive"))

    def buscar():
        for r in raices:
            c = os.path.join(r, "Calendario", "datos")
            if os.path.isfile(os.path.join(c, "config.json")):
                return c
        return None

    if buscar() is None and espera and any(os.path.isdir(os.path.join(r, "Calendario")) for r in raices):
        esperar(lambda: buscar() is not None, espera)
    return buscar()


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


APP = "calendario"
ESTADO_ACTIVO_MIN = 20          # un equipo que ha publicado hace menos de esto está «vivo»
FICHERO_LIMPIO = re.compile(r"(agenda|notas|estado-calendario|estado-postit)-([A-Za-z0-9_]+)\.json")


def _ahora_local(ahora):
    return ahora.replace(tzinfo=None)


def escribir_estado(carpeta, estado, ahora, nombre=None, cada=300):
    """estado-calendario-<equipo>.json: «sigo vivo», qué versión soy y de dónde leo los eventos. Solo se reescribe si
    cambia algo o pasan «cada» segundos (OneDrive no sincroniza en balde). Devuelve True si escribió."""
    nombre = nombre or equipo()
    ruta = os.path.join(carpeta, f"estado-{APP}-{nombre}.json")
    nuevo = {"version": 1, "app": APP, "equipo": nombre, **estado}
    try:
        with open(ruta, encoding="utf-8") as f:
            viejo = json.load(f)
        if {k: v for k, v in viejo.items() if k != "escrito"} == nuevo and ahora.timestamp() - os.path.getmtime(ruta) < cada:
            return False
    except (OSError, ValueError):
        pass
    nuevo["escrito"] = _iso_local(ahora)
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(nuevo, f, ensure_ascii=False, indent=1)
    os.replace(tmp, ruta)
    return True


def copias_en_conflicto(carpeta):
    """Ficheros que OneDrive ha duplicado al chocar dos equipos («agenda-A-B.json», «x (1).json», «...conflict...»)."""
    try:
        nombres = sorted(n for n in os.listdir(carpeta) if n.endswith(".json"))
    except OSError:
        return []
    conocidos = {(m.group(1), m.group(2)) for m in map(FICHERO_LIMPIO.fullmatch, nombres) if m}
    salida = []
    for n in nombres:
        if FICHERO_LIMPIO.fullmatch(n):
            continue
        resto = n[:-5]
        if " (" in n or "onflict" in n or any(resto.startswith(f"{p}-{k}-") for p, k in conocidos):
            salida.append(n)
    return salida


def _hace(segundos):
    m = int(segundos // 60)
    if m < 1:
        return "ahora mismo"
    if m < 90:
        return f"hace {m} min"
    if m < 48 * 60:
        return f"hace {m // 60} h"
    return f"hace {m // 1440} días"


def informe(carpeta, ahora, version=None, nombre=None, datos=None, onedrive=None):
    """Lo que se ve de la sincronización, como [(nivel, texto)]; nivel «ok», «aviso» o «mal».
    version: la de este programa · datos: carpeta del config.json compartido · onedrive: True/False si se sabe."""
    nombre = nombre or equipo()
    if not carpeta:
        return [("mal", "No hay carpeta compartida con Post-It")]
    sal = []
    if onedrive is False:
        sal.append(("mal", "OneDrive no está en marcha: nada de lo compartido se actualiza"))
    ahora = _ahora_local(ahora)
    estados = [e for e in _ficheros(carpeta, "estado-calendario") + _ficheros(carpeta, "estado-postit") if e.get("app")]
    if not any(e.get("equipo") == nombre and e.get("app") == APP for e in estados):
        sal.append(("aviso", "Este equipo aún no ha publicado su estado"))
    for e in sorted(estados, key=lambda x: (x["app"], str(x.get("equipo")))):
        try:
            edad = (ahora - dt.datetime.fromisoformat(e["escrito"])).total_seconds()
        except (KeyError, TypeError, ValueError):
            sal.append(("aviso", f"{e['app']} en {e.get('equipo')}: estado ilegible"))
            continue
        quien = f"{'Calendario' if e['app'] == APP else 'Post-It'} {e.get('app_version', '?')} en {e.get('equipo')}"
        if edad > ESTADO_ACTIVO_MIN * 60:
            sal.append(("aviso", f"{quien}: sin publicar desde {_hace(edad)} (¿cerrado o apagado?)"))
        else:
            sal.append(("ok", f"{quien}: al día ({_hace(edad)})"))
        if e.get("outlook") == "cerrado":
            sal.append(("aviso", f"En {e.get('equipo')} Outlook está cerrado: su agenda puede estar desactualizada"))
    for app in (APP, "postit"):
        versiones = {str(e.get("app_version")) for e in estados if e["app"] == app and e.get("app_version")}
        if len(versiones) > 1:
            sal.append(("aviso", f"{'Calendario' if app == APP else 'Post-It'}: versiones distintas entre equipos ({', '.join(sorted(versiones))})"))
    for n in copias_en_conflicto(carpeta):
        sal.append(("mal", f"Copia en conflicto de OneDrive en la carpeta compartida: {n}"))
    if datos:
        try:
            for n in sorted(os.listdir(datos)):
                if n.startswith("config") and n not in ("config.json", "config.json.bak"):
                    sal.append(("aviso", f"Copia de la configuración de otro equipo en datos: {n}"))
        except OSError:
            pass
    if not any(n != "ok" for n, _ in sal):
        sal.append(("ok", "Sin conflictos ni versiones mezcladas"))
    return sal


def onedrive_en_marcha():
    """True/False si OneDrive.exe está ejecutándose; None si no se puede saber."""
    import subprocess
    try:
        salida = subprocess.run(["tasklist", "/FI", "IMAGENAME eq OneDrive.exe", "/NH"], capture_output=True, text=True,
                                timeout=10, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout
        return "onedrive.exe" in salida.lower()
    except Exception:
        return None


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
