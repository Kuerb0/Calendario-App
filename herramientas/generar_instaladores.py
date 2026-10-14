"""Genera los instaladores autoextraibles a partir del codigo de app/.

Uso (desde cualquier sitio):

    python herramientas/generar_instaladores.py

Lee app/ (calendario.py, configurar.py, extras.py, icono.ico, ayuda/, requirements.txt,
VERSION.txt) y herramientas/descargar_python.ps1, y escribe en instaladores/:

    Calendario <version> - Instalador.bat    (instalacion completa)
    Calendario <version> - Actualizar.bat    (solo el programa y las librerias)

Cada .bat lleva dentro TODOS los ficheros del programa, asi que se puede
mandar suelto a otra persona: al ejecutarlo crea app/, datos/, etc. en su
propia carpeta. Las plantillas estan en herramientas/plantillas/.

Al terminar comprueba que lo empaquetado se extrae idéntico al original.
Ejecutalo siempre despues de cambiar algo en app/, antes de repartir nada.
"""
import base64
import os
import re
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLANTILLAS = os.path.join(RAIZ, "herramientas", "plantillas")
SALIDA = os.path.join(RAIZ, "instaladores")

# ruta relativa a la carpeta principal -> modo ("text" o "b64")
ARCHIVOS = [
    ("app/calendario.py", "text"),
    ("app/configurar.py", "text"),
    ("app/extras.py", "text"),
    ("app/correo_columna.py", "text"),
    ("app/menus.py", "text"),
    ("app/requirements.txt", "text"),
    ("app/VERSION.txt", "text"),
    ("app/icono.ico", "b64"),
    ("app/ayuda/1-enlace-ics.txt", "text"),
    ("app/ayuda/2-outlook-instalado.txt", "text"),
    ("app/ayuda/3-festivos.txt", "text"),
    ("app/ayuda/4-despues.txt", "text"),
    ("herramientas/descargar_python.ps1", "text"),
]

PLANTILLAS_BAT = [
    ("instalador.bat.txt", "Calendario {v} - Instalador.bat"),
    ("actualizar.bat.txt", "Calendario {v} - Actualizar.bat"),
]


def leer_bytes(rel):
    with open(os.path.join(RAIZ, *rel.split("/")), "rb") as f:
        return f.read()


def lineas_de_texto(rel):
    texto = leer_bytes(rel).decode("utf-8-sig")
    return texto.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n").split("\n")


def version_del_programa():
    codigo = leer_bytes("app/calendario.py").decode("utf-8-sig")
    m = re.search(r'^VERSION\s*=\s*"([^"]+)"', codigo, re.M)
    if not m:
        sys.exit("No encuentro VERSION = \"...\" en app/calendario.py")
    return m.group(1)


def bloque(rel, modo):
    """Las lineas de un fichero dentro del .bat."""
    salida = [f":::BEGIN {rel}|{modo}"]
    if modo == "b64":
        b64 = base64.b64encode(leer_bytes(rel)).decode("ascii")
        salida += [b64[i:i + 76] for i in range(0, len(b64), 76)]
    else:
        for linea in lineas_de_texto(rel):
            if linea.startswith(":::"):
                sys.exit(f"{rel}: una linea empieza por ':::' y rompe el empaquetado")
            salida.append(linea)
    salida.append(":::END")
    return salida


def carga_util():
    lineas = []
    for rel, modo in ARCHIVOS:
        lineas += bloque(rel, modo)
    return lineas


def extraer(contenido):
    """Misma logica que el extractor de PowerShell del .bat (para comprobar)."""
    lineas = contenido.decode("utf-8").replace("\r\n", "\n").split("\n")
    salida, i = {}, 0
    while i < len(lineas):
        if lineas[i].startswith(":::BEGIN "):
            rel, modo = lineas[i][9:].strip().split("|")
            j, buf = i + 1, []
            while j < len(lineas) and not lineas[j].startswith(":::END"):
                buf.append(lineas[j])
                j += 1
            if modo == "b64":
                salida[rel] = base64.b64decode("".join(buf))
            else:
                salida[rel] = ("\r\n".join(buf) + "\r\n").encode("utf-8")
            i = j
        i += 1
    return salida


def main():
    version = version_del_programa()
    os.makedirs(SALIDA, exist_ok=True)
    carga = carga_util()
    for plantilla, nombre in PLANTILLAS_BAT:
        with open(os.path.join(PLANTILLAS, plantilla), encoding="utf-8") as f:
            texto = f.read().replace("\r\n", "\n")
        assert "@@PAYLOAD@@" in texto, plantilla
        lineas = texto.replace("@@VERSION@@", version).split("\n")
        completo = []
        for linea in lineas:
            if linea.strip() == "@@PAYLOAD@@":
                completo += carga
            else:
                completo.append(linea)
        contenido = ("\r\n".join(completo).rstrip("\r\n") + "\r\n").encode("utf-8")

        # comprobacion: lo que sale al extraer == lo que hay en app/
        sacado = extraer(contenido)
        for rel, modo in ARCHIVOS:
            esperado = leer_bytes(rel)
            if modo == "text":
                esperado = ("\r\n".join(lineas_de_texto(rel)) + "\r\n").encode("utf-8")
            if sacado.get(rel) != esperado:
                sys.exit(f"{nombre}: {rel} no sale igual al extraerlo")

        destino = os.path.join(SALIDA, nombre.format(v=version))
        with open(destino, "wb") as f:
            f.write(contenido)
        print(f"OK  {os.path.relpath(destino, RAIZ)}  ({len(contenido) // 1024} KB)")


if __name__ == "__main__":
    main()
