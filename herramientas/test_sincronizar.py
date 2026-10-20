"""python herramientas/test_sincronizar.py  ·  debe terminar en «TEST OK»"""
import datetime as dt, json, os, sys, tempfile
from types import SimpleNamespace
from zoneinfo import ZoneInfo

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "app"))
import sincronizar as S

TZ = ZoneInfo("Europe/Madrid")
def ev(ini, fin, titulo, todo=False, uid="u", parte=0, **kw):
    return SimpleNamespace(inicio=ini, fin=fin, titulo=titulo, todo_dia=todo, uid=uid, parte=parte, **kw)
def d(h, m=0, dia=19): return dt.datetime(2026, 10, dia, h, m, tzinfo=TZ)

carpeta = tempfile.mkdtemp()
# --- notas de Post-It -> eventos
json.dump({"version": 1, "equipo": "PC1", "items": [
    {"id": "postit:n1:a", "titulo": "Dentista", "inicio": "2026-10-20T18:30:00", "fin": None, "hecho": False, "nota": ""},
    {"id": "postit:n1:b", "titulo": "Reunión", "inicio": "2026-10-20T10:00:00", "fin": "2026-10-20T11:30:00", "hecho": True, "nota": "x"},
    {"id": "postit:n1:c", "titulo": "  ", "inicio": "2026-10-20T10:00:00", "fin": None, "hecho": False}]},
    open(os.path.join(carpeta, "notas-PC1.json"), "w", encoding="utf-8"))
json.dump({"version": 1, "equipo": "PC2", "items": [
    {"id": "postit:n1:a", "titulo": "Dentista", "inicio": "2026-10-20T18:30:00", "fin": None, "hecho": False}]},
    open(os.path.join(carpeta, "notas-PC2.json"), "w", encoding="utf-8"))
open(os.path.join(carpeta, "notas-PC3.json"), "w").write("{ roto")
crear = lambda i, f, t, td, u, **kw: SimpleNamespace(inicio=i, fin=f, titulo=t, todo_dia=td, uid=u, **kw)
evs = S.leer_notas(carpeta, TZ, crear)
assert [(e.titulo, e.inicio.hour, e.fin.hour) for e in evs] == [("Dentista", 18, 19), ("✓ Reunión", 10, 11)], evs
assert evs[0].inicio.tzinfo is TZ and evs[1].notas.startswith("Viene")
# --- tarea sin hora apuntada bajo una fecha: evento de todo el día
json.dump({"version": 1, "equipo": "PC4", "items": [
    {"id": "postit:n2:t", "titulo": "Cama", "inicio": "2026-10-21T12:00:00", "fin": None, "hecho": True, "todo_dia": True}]},
    open(os.path.join(carpeta, "notas-PC4.json"), "w", encoding="utf-8"))
td = [e for e in S.leer_notas(carpeta, TZ, crear) if e.todo_dia]
assert [(e.titulo, e.inicio.hour, (e.fin - e.inicio).days) for e in td] == [("✓ Cama", 0, 1)], td
os.remove(os.path.join(carpeta, "notas-PC4.json"))
assert S.firma_notas(carpeta) and S.firma_notas(os.path.join(carpeta, "no")) == ""

# --- el Calendario escribe su agenda: sin lo de Post-It, sin continuaciones, sin lo lejano, repetidos separados
lista = [ev(d(10), d(11), "Reunión", uid="r"), ev(d(10, dia=20), d(11, dia=20), "Reunión", uid="r"),
         ev(d(9), d(10), "De Post-It", uid="postit:n1:a"), ev(d(12), d(13), "Continuación", uid="c", parte=1),
         ev(d(0, dia=19), d(0, dia=20), "Cumple", todo=True, uid="cumple"), ev(d(9, dia=15), d(10, dia=15), "Pasado", uid="p"), ev(dt.datetime(2026, 9, 1, 10, tzinfo=TZ), dt.datetime(2026, 9, 1, 11, tzinfo=TZ), "Muy atrás", uid="a"), ev(dt.datetime(2026, 12, 20, 10, tzinfo=TZ), dt.datetime(2026, 12, 20, 11, tzinfo=TZ), "Lejos", uid="l")]
ahora = d(8)
assert S.escribir_agenda(carpeta, lista, ahora, "PC1") is True
assert S.escribir_agenda(carpeta, lista, ahora, "PC1") is False            # nada cambió: no se toca el fichero
datos = json.load(open(os.path.join(carpeta, "agenda-PC1.json"), encoding="utf-8"))
assert [e["titulo"] for e in datos["eventos"]] == ["Pasado", "Cumple", "Reunión", "Reunión"], datos
assert len({e["id"] for e in datos["eventos"]}) == 4 and datos["eventos"][1]["todo_dia"] and datos["eventos"][2]["inicio"] == "2026-10-19T10:00:00"

# --- carpeta: la guardada manda; si no, la hermana de la app; si no, se crea
raiz = tempfile.mkdtemp(); app = os.path.join(raiz, "Calendario"); os.makedirs(app)
assert S.resolver_carpeta(None, app, {}, raiz) == (os.path.join(raiz, "AgendaCompartida"), "creada")
assert S.resolver_carpeta(None, app, {}, raiz)[1] == "detectada"
assert S.resolver_carpeta(carpeta, app, {}, raiz) == (carpeta, "guardada")
# --- datos del .exe: los de OneDrive\Calendario\datos si existen (con config.json); si no, None (junto al .exe)
od = os.path.join(raiz, "OneDrive"); os.makedirs(os.path.join(od, "Calendario", "datos"))
assert S.datos_compartidos({}, raiz) is None                              # carpeta vacía, sin config.json: no vale
open(os.path.join(od, "Calendario", "datos", "config.json"), "w").write("{}")
assert S.datos_compartidos({}, raiz) == os.path.join(od, "Calendario", "datos")
assert S.datos_compartidos({"OneDrive": od}, tempfile.mkdtemp()) == os.path.join(od, "Calendario", "datos")
assert S.datos_compartidos({}, tempfile.mkdtemp()) is None
print("TEST OK")
