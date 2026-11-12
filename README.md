# Automatización de Reportes de Licitaciones de Codelco desde Ariba

Aplicación para Windows y Python 3.10 o posterior. Extrae los eventos abiertos de Codelco, los compara con el último reporte anterior, descarga los documentos nuevos y produce un consolidado.

## Preparación

En PowerShell, desde esta carpeta:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
\(env:PLAYWRIGHT_BROWSERS_PATH = (Join-Path\)PWD '.datos\navegadores')
.\.venv\Scripts\python.exe -m playwright install chromium
```

Registrar las credenciales de Ariba directamente en la terminal, sin enviarlas por chat:

```powershell
.\.venv\Scripts\python.exe principal.py --configurar-credenciales
```

El correo y la contraseña se guardan mediante `keyring` en el Administrador de credenciales de Windows del usuario que ejecuta el programa. La sesión de Chromium queda en `.datos\perfil-ariba`; debe protegerse como información sensible y no compartirse ni subirse al repositorio.

## Validación y Ejecución

Antes de ejecutar el proceso real, verificar el acceso y la exportación, sin descargar documentos individuales ni enviar correos:

```powershell
.\.venv\Scripts\python.exe principal.py --visible --validar-ariba
```

Para procesar las licitaciones:

```powershell
.\.venv\Scripts\python.exe principal.py --visible
```

Sin `--visible`, Chromium se ejecuta en segundo plano. Los reportes de comparación se guardan en `reportes_diarios` y los consolidados en `resumen_licitaciones`, dentro de la raíz del proyecto; `--raiz RUTA` permite elegir otra raíz para ambas carpetas. Los documentos permanecen en `.datos\ejecuciones\AAAA-MM-DD` dentro del proyecto.

Se necesita al menos un reporte anterior en `reportes_diarios`, con el nombre `AAAA-MM-DD_Licitaciones_Codelco_(Ariba).xlsx` o el formato anterior `LIC ARIBA DD-MM.xlsx`. Para este último se interpreta la fecha como la ocurrencia anterior más reciente respecto al día de ejecución.

El programa valida que la exportación tenga los encabezados y estados esperados, compara solo los ID actuales con los del reporte anterior y conserva los enlaces de cada título. Los eventos cuyo título comienza por `IMCSPOT` se registran como excluidos y no se descargan ni se incluyen en el consolidado; permanecen en el reporte diario como parte de la exportación histórica. Usa la fecha de vencimiento del documento impreso; si falta, el evento se registra como excluido.

Extrae ítems principales y anidados, incluidos los de documentos SR y RFP, usando sus campos y descripciones para distinguirlos de secciones administrativas. Los documentos sin ítems reconocidos se incluyen en el consolidado con la indicación `Revisar manualmente` y los nombres de los adjuntos visibles; el programa no descarga ni analiza esos adjuntos. Los resúmenes ya generados no se modifican automáticamente.

Los datos de cada ejecución se guardan en `.datos\ejecuciones\AAAA-MM-DD`. Ante fallos transitorios de navegación, sesión o respuestas HTTP 429 y 5xx, el programa recupera el acceso y vuelve a intentar. Si un evento falla tres veces, continúa con los demás y lo vuelve a intentar en rondas con pausas crecientes de hasta cinco minutos. Si la página se cierra, intenta reabrir el navegador y continuar. No publica el reporte ni el consolidado mientras queden eventos pendientes; si el fallo persiste, puede permanecer esperando indefinidamente.

Los errores de credenciales, datos inesperados o archivos diferentes requieren revisión y detienen el proceso: no es posible garantizar recuperación automática ante todos los problemas de Ariba. Volver a ejecutar el comando reanuda la ejecución del día. Para reanudar una ejecución anterior aún conservada, agregar `--fecha AAAA-MM-DD`.

Al finalizar, el reporte de comparación `AAAA-MM-DD_Licitaciones_Codelco_(Ariba).xlsx` se publica en `reportes_diarios` y el consolidado `AAAA-MM-DD_Resumen_Licitaciones_Codelco.xlsx` en `resumen_licitaciones`. No se sobrescriben archivos diferentes y se rechazan ejecuciones simultáneas. Si ya existe un reporte del día sin un estado de ejecución correspondiente, el proceso se detiene para evitar reemplazarlo.

Al iniciar una ejecución normal, se eliminan las carpetas con fecha de más de siete días en `.datos\ejecuciones`, incluidos los documentos y estados incompletos que contengan. La validación `--validar-ariba` y la configuración de credenciales no ejecutan esta limpieza. Los reportes publicados no se eliminan. La limpieza solo afecta carpetas con nombre de fecha válido y no sigue enlaces.

La ejecución del mismo día que ya finalizó no genera otro consolidado. Si no hay documentos aptos, no se crea consolidado. Además de la prueba de navegación simulada, se ha comprobado el acceso real, la exportación y el procesamiento aislado de un evento normal y uno SR. Valide los resultados con más ejemplos y en su entorno operativo antes de programar tareas desatendidas.

## Ejecutable de Respaldo

El ejecutable `Licitaciones_Codelco.exe` se genera en la raíz y no se incluye en Git. Por defecto muestra una consola de progreso y ejecuta Chromium sin ventana visible. Usa `.datos\navegadores` y `.datos\perfil-ariba` junto al ejecutable; para usarlo en otro equipo hay que instalar Chromium en esa ubicación y registrar las credenciales de ese usuario de Windows.

Si faltan credenciales y hay una consola interactiva, las solicita sin mostrarlas en pantalla. El ejecutable espera Enter al terminar; para una tarea desatendida, usar `--sin-pausa` y configurar antes las credenciales. Los errores y cambios inesperados de Ariba, los cortes de energía o de red y la ausencia de eventos aptos impiden garantizar que exista un resumen cada día; supervise la tarea y su salida.

Para generar nuevamente el ejecutable en Windows desde la raíz del proyecto:

```powershell
.\.venv\Scripts\python.exe -m pip install "pyinstaller>=6.16,<7"
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --onefile --console --name Licitaciones_Codelco --distpath . --workpath .datos/pyinstaller-trabajo --specpath .datos/pyinstaller-config --hidden-import keyring.backends.Windows --collect-all playwright principal.py
```

El comando reemplaza el ejecutable anterior. El navegador no se empaqueta dentro del archivo; siga los pasos de preparación para instalarlo.

## Pruebas

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -p 'test_*.py' -v
```

Para programar una tarea de lunes a viernes, usar el Programador de tareas de Windows con el mismo usuario que registró las credenciales, como programa la ruta absoluta a `.venv\Scripts\python.exe`, como argumento `principal.py` y como directorio inicial la raíz del proyecto. El equipo debe estar encendido y tener acceso a la red. Configurar la tarea solamente después de validar el flujo real.
