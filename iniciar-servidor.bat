@echo off
setlocal
cd /d "%~dp0"

set PORT=8080
if not "%~1"=="" set PORT=%~1

echo ==========================================
echo  Gerador de Horario Escolar - Servidor
echo  Pasta: %CD%
echo  Porta: %PORT%
echo ==========================================

where python >nul 2>nul
if %errorlevel%==0 (
  start "" "http://localhost:%PORT%/"
  echo Usando Python (python -m http.server %PORT%)
  python -m http.server %PORT%
  goto :end
)

where py >nul 2>nul
if %errorlevel%==0 (
  start "" "http://localhost:%PORT%/"
  echo Usando Python Launcher (py -m http.server %PORT%)
  py -m http.server %PORT%
  goto :end
)

where npx >nul 2>nul
if %errorlevel%==0 (
  start "" "http://localhost:%PORT%/"
  echo Usando Node (npx --yes serve -l %PORT%)
  npx --yes serve -l %PORT%
  goto :end
)

echo.
echo Nao encontrei Python nem Node.js (npx) instalados.
echo Instale Python ou Node.js para iniciar o servidor local.
pause

:end
endlocal
