# Automatización de Reportes de Licitaciones de Codelco desde Ariba

Aplicación para Windows y Python 3.10 o posterior. Extrae los eventos de licitación abiertos de Codelco desde Ariba, los compara con el último reporte anterior, descarga los documentos nuevos y produce un consolidado que posteriormente se envía por correo a quien corresponda.

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

## Configuración del Correo

Para configurar el remitente, active la verificación en dos pasos y cree una contraseña de aplicación de Google. Desde la terminal, con el mismo usuario de Windows que ejecutará la tarea:

```powershell
.\.venv\Scripts\python.exe principal.py --configurar-correo
```

El comando solicita el correo remitente y la contraseña de aplicación. Escriba la contraseña directamente en la terminal, sin compartirla. Se guardan por separado de las credenciales de Ariba en el Administrador de credenciales de Windows. La cuenta debe tener habilitado este método por su administrador de Google Workspace. Para cambiar el remitente o la contraseña, vuelva a ejecutar el comando.

Configure al menos un destinatario y, opcionalmente, destinatarios en copia:

```powershell
.\.venv\Scripts\python.exe principal.py --configurar-destinatarios
```

Introduzca las direcciones separadas por comas; para no agregar copia, presione Enter en el campo CC. Al ejecutar nuevamente el comando se reemplazan ambas listas. Las direcciones se guardan en `.datos\destinatarios_correo.json`, excluido de Git; proteja también esa carpeta mediante los permisos del equipo. No se envía ningún correo al configurar.

El ejecutable de respaldo `Licitaciones_Codelco.exe` también admite ambos comandos, sustituyendo el comando de Python por el nombre del ejecutable y conservando los argumentos. Al abrirlo de forma interactiva para una ejecución normal, solicita el remitente, la contraseña o los destinatarios si aún no están configurados. En ejecuciones desatendidas con `--sin-pausa`, deben configurarse previamente.

Para enviar una prueba al propio remitente usando un resumen existente, sin acceder a Ariba:

```powershell
.\.venv\Scripts\python.exe principal.py --probar-correo --fecha 2026-11-12
```

La prueba se envía solo al remitente, sin usar los destinatarios configurados. La fecha debe corresponder al resumen existente en `resumen_licitaciones` y a una ejecución completada cuyos documentos y estado sigan disponibles en `.datos\ejecuciones`; si se omite, se usa la fecha actual. La prueba tiene su propio registro y no se repite al ejecutar el mismo comando otra vez. Para enviar otra prueba intencional de la misma fecha, agregue `--nueva-prueba`. Si el resultado anterior es incierto, debe resolverlo primero. La prueba es independiente del envío automático, por lo que realizar ambos para la misma fecha puede producir dos correos intencionales. Las ejecuciones finalizadas antes de activar el correo no se envían automáticamente.

En las nuevas ejecuciones, tras publicar el resumen y el reporte diario, se envían a los destinatarios y CC configurados dos adjuntos: el resumen y `AAAA-MM-DD_Documentos_Licitaciones_Codelco.zip`, con los Word descargados y registrados. El ZIP se conserva en `.datos\ejecuciones\AAAA-MM-DD` y está sujeto a la limpieza de siete días. El asunto y el cuerpo son `Reporte Licitaciones Codelco Ariba (DD-MM-AAAA)`. Si no hay documentos aptos y no se genera resumen, no se envía correo. Los fallos temporales anteriores al envío y los rechazos temporales explícitos se reintentan; errores de configuración o credenciales detienen el proceso con el correo pendiente. Volver a ejecutar el programa reanuda el envío pendiente. La aceptación por Gmail no garantiza que el mensaje llegue a la bandeja del destinatario, y Gmail puede rechazar mensajes demasiado grandes.

Si se pierde la conexión durante el envío, el resultado puede ser incierto y no se reenvía automáticamente. Revise la carpeta Enviados de la cuenta remitente. Para registrar que sí se envió, use `--resolver-correo enviado`; para permitir un nuevo intento tras confirmar que no se envió, use `--resolver-correo reintentar` y vuelva a ejecutar el proceso. Si Gmail acepta solo algunos destinatarios, el programa muestra los rechazados; tras revisarlos, `--resolver-correo reintentar` permite continuar solo con ellos. Las direcciones de un envío iniciado se conservan en su estado aunque se cambie la configuración general. Agregue `--fecha AAAA-MM-DD` si corresponde a otro día. Para resolver una prueba manual, agregue también `--correo-de-prueba`. El envío automático y el de prueba mantienen registros separados.

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

Se necesita al menos un reporte anterior en `reportes_diarios`, con el nombre `AAAA-MM-DD_Licitaciones_Codelco_(Ariba).xlsx` o el formato antiguo `LIC ARIBA DD-MM.xlsx`. Para este último se interpreta la fecha como la ocurrencia anterior más reciente respecto al día de ejecución.

El programa valida que la exportación tenga los encabezados y estados esperados, compara solo los ID actuales con los del reporte anterior y conserva los enlaces de cada título. Los eventos cuyo título comienza por `IMCSPOT` se registran como excluidos y no se descargan ni se incluyen en el consolidado; permanecen en el reporte diario como parte de la exportación histórica. Usa la fecha de vencimiento del documento impreso; si falta, el evento se registra como excluido.

Extrae ítems principales y anidados, incluidos los de documentos SR y RFP, usando sus campos y descripciones para distinguirlos de secciones administrativas. El consolidado incluye la fecha y hora de finalización de la exportación y un hipervínculo a Ariba sobre cada nombre de documento. Los documentos sin ítems reconocidos se incluyen en el consolidado con la indicación `Revisar manualmente` y los nombres de los adjuntos visibles; el programa no descarga ni analiza esos adjuntos. Los resúmenes ya generados no se modifican automáticamente.

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
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --onefile --console --name Licitaciones_Codelco --distpath . --workpath .datos/pyinstaller-produccion-trabajo --specpath .datos/pyinstaller-config --hidden-import keyring.backends.Windows --collect-all playwright principal.py
```

El comando reemplaza el ejecutable anterior. El navegador no se empaqueta dentro del archivo; siga los pasos de preparación para instalarlo.

## Pruebas

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -p 'test_*.py' -v
```

Para programar una tarea de lunes a viernes, usar el Programador de tareas de Windows con el mismo usuario que registró las credenciales, como programa la ruta absoluta a `.venv\Scripts\python.exe`, como argumento `principal.py` y como directorio inicial la raíz del proyecto. El equipo debe estar encendido y tener acceso a la red. Configurar la tarea solamente después de validar el flujo real.
