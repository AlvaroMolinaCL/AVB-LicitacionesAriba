@echo off
chcp 65001 > nul
title Panel de Control - Licitaciones Codelco
cd /d "%~dp0"

set "NOMBRE_TAREA=Actualizar Licitaciones Codelco desde Ariba"

:MENU
set "opcion="
cls
echo ========================================================
echo     PANEL DE CONTROL - LICITACIONES CODELCO (ARIBA)
echo ========================================================
echo.
echo   --- EJECUCION Y PRUEBAS ---
echo   1. Validar conexion a Ariba (sin descargar ni enviar)
echo   2. Ejecutar proceso completo hoy (navegador visible)
echo   3. Reanudar ejecucion de una fecha anterior
echo   4. Enviar correo de prueba al remitente
echo.
echo   --- CONFIGURACION Y SOPORTE ---
echo   5. Cambiar credenciales de Ariba
echo   6. Cambiar correo remitente o destinatarios
echo   7. Resolver correo incierto (Marcar como ENVIADO)
echo   8. Resolver correo incierto (Permitir REINTENTAR)
echo.
echo   --- TAREA PROGRAMADA Y MANTENIMIENTO ---
echo   9. Activar (Habilitar) tarea programada en este PC
echo  10. Pausar (Deshabilitar) tarea programada en este PC
echo  11. Ver estado y proxima ejecucion de la tarea
echo  12. Reparar rutas (si movio esta carpeta de lugar)
echo   0. Salir
echo.
set /p opcion="Seleccione una opcion (0-12): "

if "%opcion%"=="1" goto VALIDAR
if "%opcion%"=="2" goto EJECUTAR
if "%opcion%"=="3" goto REANUDAR
if "%opcion%"=="4" goto PRUEBA_CORREO
if "%opcion%"=="5" goto CRED_ARIBA
if "%opcion%"=="6" goto CRED_CORREO
if "%opcion%"=="7" goto RES_ENVIADO
if "%opcion%"=="8" goto RES_REINTENTAR
if "%opcion%"=="9" goto ACTIVAR_TAREA
if "%opcion%"=="10" goto DESACTIVAR_TAREA
if "%opcion%"=="11" goto ESTADO_TAREA
if "%opcion%"=="12" goto REPARAR_RUTA
if "%opcion%"=="0" exit /b
goto MENU

:VALIDAR
echo.
.\.venv\Scripts\python.exe principal.py --visible --validar-ariba
pause
goto MENU

:EJECUTAR
echo.
.\.venv\Scripts\python.exe principal.py --visible
pause
goto MENU

:REANUDAR
echo.
set /p fecha_rea="Ingrese fecha a reanudar (AAAA-MM-DD, dentro de ultimos 7 dias): "
if "%fecha_rea%"=="" goto MENU
.\.venv\Scripts\python.exe principal.py --visible --fecha %fecha_rea%
pause
goto MENU

:PRUEBA_CORREO
echo.
set /p fecha="Ingrese fecha del resumen (AAAA-MM-DD) o presione Enter para hoy: "
if "%fecha%"=="" (
    .\.venv\Scripts\python.exe principal.py --probar-correo
) else (
    .\.venv\Scripts\python.exe principal.py --probar-correo --fecha %fecha%
)
pause
goto MENU

:CRED_ARIBA
echo.
.\.venv\Scripts\python.exe principal.py --configurar-credenciales
pause
goto MENU

:CRED_CORREO
echo.
.\.venv\Scripts\python.exe principal.py --configurar-correo
.\.venv\Scripts\python.exe principal.py --configurar-destinatarios
pause
goto MENU

:RES_ENVIADO
echo.
set /p fecha_env="Ingrese fecha de la ejecucion (AAAA-MM-DD): "
if "%fecha_env%"=="" goto MENU
.\.venv\Scripts\python.exe principal.py --resolver-correo enviado --fecha %fecha_env%
pause
goto MENU

:RES_REINTENTAR
echo.
set /p fecha_rei="Ingrese fecha de la ejecucion (AAAA-MM-DD): "
if "%fecha_rei%"=="" goto MENU
.\.venv\Scripts\python.exe principal.py --resolver-correo reintentar --fecha %fecha_rei%
if errorlevel 1 (
    echo ERROR: No se pudo resolver el correo; no se reanudara el envio.
    pause
    goto MENU
)
echo Reanudando envio...
.\.venv\Scripts\python.exe principal.py --fecha %fecha_rei%
pause
goto MENU

:ACTIVAR_TAREA
echo.
echo IMPORTANTE: Asegurese de que ningun otro PC tenga la tarea activa.
schtasks /change /tn "%NOMBRE_TAREA%" /enable
if errorlevel 1 (
    echo ERROR: No se encontro la tarea. Use la Opcion 12 para crearla o repararla.
) else (
    echo Tarea programada HABILITADA correctamente.
)
pause
goto MENU

:DESACTIVAR_TAREA
echo.
schtasks /change /tn "%NOMBRE_TAREA%" /disable
if errorlevel 1 (
    echo ERROR: No se encontro la tarea o faltan permisos.
) else (
    echo Tarea programada DESHABILITADA correctamente.
)
pause
goto MENU

:ESTADO_TAREA
echo.
schtasks /query /tn "%NOMBRE_TAREA%" /fo LIST /v
if errorlevel 1 (
    echo ERROR: No se pudo consultar la tarea. Revise permisos y Programador de tareas.
)
pause
goto MENU

:REPARAR_RUTA
echo.
echo ========================================================
echo   REPARADOR DE UBICACION Y TAREA PROGRAMADA
echo   (Conserva sus claves, correos y datos guardados)
echo ========================================================
echo.
echo [1/2] Actualizando entorno interno a la carpeta actual...
rmdir /s /q .venv >nul 2>&1
if exist ".venv" (
    echo ERROR: No se pudo eliminar el entorno anterior. Cierre los procesos que lo usen.
    pause
    goto MENU
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\preparar_python.ps1" -Carpeta "%~dp0."
if errorlevel 1 (
    echo ERROR: No se pudo preparar Python. La reparacion queda incompleta.
    pause
    goto MENU
)
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: No se pudieron instalar las librerias. La reparacion queda incompleta.
    pause
    goto MENU
)
set "PLAYWRIGHT_BROWSERS_PATH=%~dp0.datos\navegadores"
.\.venv\Scripts\python.exe -m playwright install chromium
if errorlevel 1 (
    echo ERROR: No se pudo instalar Chromium. La reparacion queda incompleta.
    pause
    goto MENU
)

echo.
echo [2/2] Reconfigurando tarea automatica en Windows...
set "crear_t="
set /p crear_t="¿Desea activar/actualizar la tarea automatica en este PC? (S/N): "
if /i not "%crear_t%"=="S" (
    echo Entorno reparado. La tarea existente, si la hay, no se modifico.
    echo Si movio la carpeta, compruebe sus rutas antes de ejecutarla.
    pause
    goto MENU
)

set "hora="
set /p hora="Ingrese hora de ejecucion diaria (formato HH:MM, ej. 08:30): "
if "%hora%"=="" set hora=08:30

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\configurar_tarea.ps1" -Carpeta "%~dp0." -Hora "%hora%"
if errorlevel 1 (
    echo ERROR: La reparacion de la tarea no se confirmo. Revise el mensaje anterior.
    pause
    goto MENU
)
echo.
echo Reparacion completada con exito. Sus claves y correos se mantuvieron intactos.
pause
goto MENU
