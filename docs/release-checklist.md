# Checklist de release

## 1. Antes de generar el instalador
- [ ] El código está actualizado y sin cambios de prueba temporales.
- [ ] `python herramientas/tests_calendario.py` termina en "OK".
- [ ] `app/VERSION.txt` y `VERSION` en `app/calendario.py` coinciden con la versión a entregar.
- [ ] Ejecutado `python herramientas/generar_instaladores.py` (rehace los `.bat` de `instaladores/` y comprueba que se extraen idénticos).

## 2. Qué se entrega
- [ ] Solo `instaladores/Calendario 2.0 - Instalador.bat` (y `Calendario 2.0 - Actualizar.bat` para quien ya tenga la app).
- [ ] No se incluye `datos/`, `python/`, `exe/`, `cache.pkl`, `registro.txt`, `alternar.senal` ni `*.zip`.
- [ ] El `.gitignore` excluye esos ficheros, `__pycache__/` y `*.pyc`.

## 3. Instalación en una máquina limpia
- [ ] Ejecutar el instalador en una carpeta nueva con permiso de escritura.
- [ ] Detecta Python o instala uno portátil en la subcarpeta `python`.
- [ ] Crea `app/` y `datos/`, extrae el programa e instala las librerías de `app/requirements.txt`.
- [ ] Comprobación final de librerías sin avisos (zonas horarias y Outlook).
- [ ] Crea el acceso directo del Escritorio y pregunta por el arranque con Windows.
- [ ] Se abre el asistente de configuración y, al guardarlo, arranca el widget sin errores.

## 4. Actualización
- [ ] Ejecutar `Calendario 2.0 - Actualizar.bat` sobre una instalación existente.
- [ ] Se conservan `datos/config.json` y `datos/perfiles/`.
- [ ] Desde la estructura antigua: los datos se mueven a `datos/` y los accesos directos apuntan a `app/`.

## 5. Prueba final
- [ ] Los perfiles funcionan y no se mezclan con el código.
- [ ] El botón ⚙ abre el Panel de control y Ctrl+K la paleta de comandos (ventana aparte y anclados).
- [ ] El widget abre correctamente y muestra los eventos.
