@echo off
chcp 65001 >nul
setlocal EnableExtensions
title Calendario - limpiar restos de la estructura antigua
color 0E

set "RAIZ=%~dp0.."
for %%I in ("%RAIZ%") do set "RAIZ=%%~fI"

echo ============================================================
echo    Limpiar restos de la estructura antigua
echo ============================================================
echo.
echo Carpeta: %RAIZ%
echo.

if not exist "%RAIZ%\app\calendario.py" (
    echo No encuentro app\calendario.py: esta no es la estructura nueva.
    echo No se borra nada.
    pause
    exit /b 1
)

echo Ahora el programa vive en "app", tus datos en "datos" y los
echo instaladores en "instaladores". Este script BORRA de la carpeta
echo principal lo que ya no hace falta:
echo.
echo   - copias antiguas del programa ^(calendario.py, configurar.py, icono...^)
echo   - instaladores antiguos ^(instalar.bat, Instalar.pyw, los .bat sueltos^)
echo   - restos de pruebas y diagnosticos, el .zip viejo, cache y registro
echo   - carpetas __pycache__, tests, ayuda y _to_delete antiguas
echo.
echo NO toca: datos, app, instaladores, herramientas, docs, .git, tu
echo configuracion antigua ^(config.json y perfiles, o sus copias .migrado^)
echo ni README.md.
echo.
echo IMPORTANTE: antes debes haber ejecutado el instalador o el actualizador
echo nuevo ^(carpeta "instaladores"^), para que el acceso directo del
echo Escritorio apunte ya a la carpeta "app".
echo.
choice /c SN /n /m "Borrar ahora? [S/N] "
if errorlevel 2 (
    echo No se ha borrado nada.
    pause
    exit /b 0
)

for %%F in (
    "calendario.py" "configurar.py" "icono.ico" "VERSION.txt"
    "instalar.bat" "Instalar.pyw" "crear_exe.bat"
    "diagnostico.py" "Diagnostico.pyw" "diagnostico.txt"
    "diag_horas.pyw" "diag_horas.txt" "test_horas.py"
    "tests_calendario.py" "release-checklist.md"
    "WidgetCalendario.zip" "config.json.bak" "cache.pkl" "registro.txt"
    "Calendario 1.0 - Instalador.bat" "Calendario 1.0 - Actualizar.bat"
) do (
    if exist "%RAIZ%\%%~F" del /f /q "%RAIZ%\%%~F" >nul 2>&1
)

for %%D in ("ayuda" "__pycache__" "tests" "_to_delete") do (
    if exist "%RAIZ%\%%~D" rmdir /s /q "%RAIZ%\%%~D" >nul 2>&1
)

echo.
echo Hecho.
pause
exit /b 0
