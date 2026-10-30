# powershell -NoProfile -ExecutionPolicy Bypass -File herramientas\test_instalador.ps1   ·   debe terminar en «TEST OK»
# Recorre los caminos del asistente que ya fallaron en Post-It: instalar, ACTUALIZAR encima (las opciones no se pierden), carpeta en la raiz de una
# unidad y las erratas de PowerShell (Create(... -replace ...) y $o/$O). Todo en carpetas temporales, con -SinRed (no descarga ni instala
# librerias) y con LOCALAPPDATA aparte: no toca tu instalacion ni tu sync.json.
$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$version = (Get-Content (Join-Path $raiz 'app\VERSION.txt') -Raw).Trim()
$bat = Join-Path $raiz "instaladores\Calendario $version - Instalador.bat"
if (-not (Test-Path $bat)) { throw "Falta ${bat}: ejecuta antes herramientas\generar_instaladores.py" }
function Comprobar($ok, $texto) { if (-not $ok) { throw "FALLA: $texto" }; "  ok  $texto" }

# --- 1. erratas de PowerShell en la plantilla
$plantilla = Get-Content (Join-Path $raiz 'herramientas\plantillas\instalador.ps1.txt') -Raw -Encoding UTF8
Comprobar ($plantilla -notmatch 'Create\(\$\w+ -replace') 'ningun Create($x -replace ...) sin parentesis (se lee como dos argumentos)'
Comprobar ($plantilla -cnotmatch '\$o\s*=|foreach\s*\(\$o\s') 'ninguna variable $o en minuscula ($o y $O son la misma: pisaba las opciones)'

# --- 2. instalar y actualizar encima, en silencio
$t = Join-Path $env:TEMP ('calendario-test-instalador-' + [guid]::NewGuid().ToString().Substring(0, 6)); New-Item -ItemType Directory "$t\local" -Force | Out-Null
$real = $env:LOCALAPPDATA; $env:LOCALAPPDATA = "$t\local"
try {
    foreach ($vez in 1, 2) {
        Remove-Item "$t\local\WidgetCalendario" -Recurse -Force -ErrorAction SilentlyContinue
        $sync = "$t\Carpeta de Sincronizaci$([char]243)n"
        $sal = (& cmd /c "`"$bat`" -Silencioso -SinRed -SinAccesos -Destino `"$t\prog`" -SyncDir `"$sync`"" 2>&1) -join "`n"
        $cfg = Get-Content "$t\local\WidgetCalendario\sync.json" -Raw -Encoding UTF8 | ConvertFrom-Json
        $que = if ($vez -eq 1) { 'instalacion nueva' } else { 'ACTUALIZACION sobre la anterior' }
        Comprobar ($vez -eq 1 -or (Test-Path "$t\prog\anteriores")) "$que`: la copia de la version anterior existe"
        Comprobar ($cfg.sync -like '*Sincronizaci*n') "$que`: la carpeta de sincronizacion llega a sync.json (con tilde y espacios)"
        Comprobar (Test-Path "$t\prog\app\calendario.py") "$que`: el programa esta en la carpeta de instalacion"
    }
    # sin -SyncDir en una actualizacion no se toca la sincronizacion que ya habia
    $sal = (& cmd /c "`"$bat`" -Silencioso -SinRed -SinAccesos -Destino `"$t\prog`"" 2>&1) -join "`n"
    $cfg = Get-Content "$t\local\WidgetCalendario\sync.json" -Raw -Encoding UTF8 | ConvertFrom-Json
    Comprobar ($cfg.sync -like '*Sincronizaci*n') 'actualizar sin indicar carpeta deja la sincronizacion como estaba'

    # --- 3. la raiz de una unidad se convierte en subcarpeta (se prueba la funcion, no se instala en D:\)
    $ps1 = Join-Path $t 'asistente.ps1'
    & python -c "import sys; sys.path.insert(0, r'$PSScriptRoot'); import generar_instaladores as g; open(r'$ps1','w',encoding='utf-8-sig').write('\n'.join(g.asistente('$version')))"
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ps1, [ref]$null, [ref]$null)
    $App = @{ Nombre = 'Calendario' }
    $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Normalizar-Destino' }, $true)[0]
    . ([scriptblock]::Create($fn.Extent.Text))
    Comprobar ((Normalizar-Destino 'D:\') -eq 'D:\Calendario') 'D:\ -> D:\Calendario'
    Comprobar ((Normalizar-Destino '\\servidor\recurso') -eq '\\servidor\recurso\Calendario') 'un recurso de red -> subcarpeta'
    Comprobar ((Normalizar-Destino 'D:\Programas\Calendario') -eq 'D:\Programas\Calendario') 'una subcarpeta se respeta'
} finally { $env:LOCALAPPDATA = $real }
"TEST OK"
