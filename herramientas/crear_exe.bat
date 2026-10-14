@echo off
setlocal
cd /d "%~dp0.."
echo ============================================
echo   Compilar Calendario.exe
echo ============================================
echo.
echo Genera un unico Calendario.exe que no necesita Python instalado.
echo Queda en la carpeta "exe". Su config.json, cache y registro se guardan
echo junto al .exe.
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No se encuentra Python en el PATH.
    pause
    exit /b 1
)

python -m pip install --upgrade pyinstaller -r "app\requirements.txt" --quiet
if errorlevel 1 (
    echo [ERROR] Fallo al instalar PyInstaller o las dependencias.
    pause
    exit /b 1
)

echo Compilando ^(tarda un par de minutos^)...
python -m PyInstaller --noconfirm --clean --onefile --noconsole ^
  --name Calendario ^
  --icon "%CD%\app\icono.ico" ^
  --add-data "%CD%\app\icono.ico;." ^
  --collect-all tzdata ^
  --collect-all holidays ^
  --hidden-import pystray._win32 ^
  --hidden-import extras ^
  --hidden-import correo_columna ^
  --hidden-import menus ^
  --paths "%CD%\app" ^
  --distpath "%CD%\exe" ^
  --workpath "%CD%\build" ^
  --specpath "%CD%\build" ^
  "%CD%\app\calendario.py"
if errorlevel 1 (
    echo [ERROR] La compilacion ha fallado.
    pause
    exit /b 1
)

echo.
echo Listo: exe\Calendario.exe
echo Puedes borrar la carpeta "build".
echo.
echo Para que arranque con Windows con el .exe:
echo   1. Win+R  y escribe  shell:startup
echo   2. Arrastra Calendario.exe ahi con el boton derecho y elige "Crear acceso directo".
echo.
echo Abrir el .exe cuando ya esta abierto = mostrar u ocultar el widget.
pause
