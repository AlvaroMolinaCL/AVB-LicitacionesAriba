@echo off
chcp 65001 > nul
title Instalador - Reporte Licitaciones Codelco
cd /d "%~dp0"

echo ========================================================
echo    INSTALACION INICIAL - LICITACIONES CODELCO (ARIBA)
echo ========================================================
echo.

:: Verificar si la carpeta esta en Descargas y advertir
echo "%~dp0" | findstr /i "Downloads Descargas" >nul
if not errorlevel 1 (
    echo ADVERTENCIA: Esta ejecutando el programa desde la carpeta Descargas.
    echo Se recomienda mover esta carpeta a Documentos o C:\ antes de instalar.
    echo Si la mueve despues, use la Opcion 12 de 2_PANEL_CONTROL.bat para reparar rutas.
    echo.
    pause
)

echo [1/6] Verificando instalacion de Python...
echo [2/6] Creando entorno virtual de Python...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\preparar_python.ps1" -Carpeta "%~dp0."
if errorlevel 1 (
    echo.
    echo ERROR: No se pudo preparar Python. La instalacion queda incompleta.
    pause
    exit /b 1
)

echo [3/6] Instalando librerias necesarias...
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: No se pudieron instalar las librerias. Verifique su conexion a Internet.
    pause
    exit /b 1
)

echo [4/6] Instalando navegador Chromium interno...
set "PLAYWRIGHT_BROWSERS_PATH=%~dp0.datos\navegadores"
.\.venv\Scripts\python.exe -m playwright install chromium
if errorlevel 1 (
    echo ERROR: No se pudo instalar Chromium. La instalacion queda incompleta.
    pause
    exit /b 1
)

echo.
echo ========================================================
echo   CONFIGURACION DE CREDENCIALES Y CORREO
echo ========================================================
echo.
echo [5/6] Configurando acceso a Ariba y correo electronico...
.\.venv\Scripts\python.exe principal.py --configurar-credenciales
if errorlevel 1 goto ERROR_CONFIGURACION
echo.
.\.venv\Scripts\python.exe principal.py --configurar-correo
if errorlevel 1 goto ERROR_CONFIGURACION
.\.venv\Scripts\python.exe principal.py --configurar-destinatarios
if errorlevel 1 goto ERROR_CONFIGURACION

echo.
echo ========================================================
echo   [6/6] TAREA PROGRAMADA AUTOMATICA (LUNES A VIERNES)
echo ========================================================
echo IMPORTANTE: Para evitar correos duplicados, active esta
echo opcion UNICAMENTE en el PC principal designado.
echo.
set /p activar_tarea="¿Desea activar la ejecucion automatica en este PC? (S/N): "

if /i "%activar_tarea%"=="S" goto CREAR_TAREA
echo.
echo Tarea automatica omitida. Podra activarla mas adelante
echo desde 2_PANEL_CONTROL.bat (Opcion 9 o 12). Una tarea previa no se modifica.
goto FIN

:CREAR_TAREA
echo.
set "hora="
set /p hora="Ingrese hora de ejecucion diaria (formato HH:MM, ej. 08:30): "
if "%hora%"=="" set hora=08:30

echo Creando tarea programada para ejecutarse de Lunes a Viernes a las %hora%...

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\configurar_tarea.ps1" -Carpeta "%~dp0." -Hora "%hora%"
if errorlevel 1 (
    echo.
    echo ERROR: No se confirmo la tarea. Use la opcion 12 del panel para reintentar.
    pause
    exit /b 1
)
goto FIN

:ERROR_CONFIGURACION
echo ERROR: No se completo la configuracion de credenciales o destinatarios.
pause
exit /b 1

:FIN
echo.
echo ========================================================
echo   INSTALACION COMPLETADA
echo ========================================================
echo Se recomienda ejecutar una validacion inicial abriendo
echo el archivo 2_PANEL_CONTROL.bat (Opcion 1).
echo.
pause
