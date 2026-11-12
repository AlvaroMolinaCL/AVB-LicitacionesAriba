import getpass
import re
import tempfile
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from bs4 import BeautifulSoup
from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError


URL_INICIO = "https://service.ariba.com/Authenticator.aw/ad/ssoIDP"
SERVICIO = "AVB-LicitacionesAriba"


class SesionExpiradaError(RuntimeError):
    pass


class ServicioTemporalError(TimeoutError):
    pass


def fecha_vencimiento(ruta: Path) -> datetime | None:
    documento = BeautifulSoup(ruta.read_text(encoding="utf-8-sig"), "html.parser")
    if not documento.find(string=lambda texto: texto and "ariba.sourcing.rfxui.PrintRFXEncode" in texto):
        raise ValueError("El documento descargado no es el informe HTML de Ariba")
    etiqueta = documento.find(string=re.compile(r"^\s*Fecha de vencimiento\s*$", re.I))
    fila = etiqueta.find_parent("tr") if etiqueta else None
    celdas = fila.find_all("td", recursive=False) if fila else []
    if len(celdas) < 2:
        return None
    valor = celdas[-1].get_text(" ", strip=True).replace("\xa0", "").strip()
    if not valor or valor.casefold() in {"n/a", "ninguno", "no definido"}:
        return None
    for formato in ("%d/%m/%Y %H:%M", "%d/%m/%Y %I:%M %p", "%d/%m/%Y"):
        try:
            return datetime.strptime(valor, formato)
        except ValueError:
            continue
    raise ValueError(f"Fecha de vencimiento no reconocida: {valor}")


def nombre_documento(fecha: datetime, identificador: str, titulo: str) -> str:
    identificador = re.sub(r"[^A-Za-z0-9_-]", "", identificador)
    if not identificador:
        raise ValueError("El ID del evento no es válido")
    nombre = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", titulo)
    nombre = re.sub(r"\s+", " ", nombre).strip(" .")
    if not nombre:
        raise ValueError("El título del evento no es válido")
    prefijo = f"{fecha:%d-%m} {identificador} - "
    return prefijo + nombre[:160 - len(prefijo)].rstrip(" .") + ".doc"


def url_ariba(url: str) -> bool:
    destino = urlparse(url)
    dominio = (destino.hostname or "").lower()
    return destino.scheme == "https" and (dominio == "ariba.com" or dominio.endswith(".ariba.com"))


def es_impresion(url: str, identificador_control: str) -> bool:
    parametros = parse_qs(urlparse(url).query)
    return (
        url_ariba(url)
        and parametros.get("awsn") == [identificador_control]
        and parametros.get("awii") == ["AWRefreshFrame"]
    )


def capturar_impresion(pagina, boton, destino: Path) -> None:
    identificador_control = boton.get_attribute("id")
    if not identificador_control:
        with pagina.expect_download(timeout=60000) as descarga:
            boton.click()
        descarga.value.save_as(destino)
        return

    resultado = []

    def es_solicitud(url):
        return es_impresion(str(url), identificador_control)

    def interceptar(ruta):
        try:
            respuesta = ruta.fetch(timeout=60000)
            contenido = respuesta.body()
            tipo = respuesta.headers.get("content-type", "").casefold()
            if respuesta.status in (401, 403) or ("text/html" in tipo and re.search(r"login|sign.?in|inicio de sesi[oó]n|contrase(?:n|ñ)a|password", contenido.decode("utf-8", errors="replace"), re.I)):
                raise SesionExpiradaError("Ariba requiere iniciar sesión nuevamente")
            if respuesta.status == 429 or 500 <= respuesta.status < 600:
                raise ServicioTemporalError(f"Ariba respondió con HTTP {respuesta.status}")
            if respuesta.status != 200 or "application/msword" not in tipo:
                raise ValueError("Ariba no devolvió un documento Word válido")
            if b"ariba.sourcing.rfxui.PrintRFXEncode" not in contenido:
                raise ValueError("El documento recibido no corresponde a una impresión de Ariba")
            resultado.append(contenido)
        except Exception as error:
            resultado.append(error)
        ruta.fulfill(status=204, body="")

    pagina.route(es_solicitud, interceptar, times=1)
    try:
        with pagina.expect_response(lambda respuesta: es_solicitud(respuesta.url), timeout=60000):
            boton.click()
    finally:
        pagina.unroute(es_solicitud, interceptar)
    if not resultado:
        raise RuntimeError("No se recibió el documento de impresión")
    if isinstance(resultado[0], Exception):
        raise resultado[0]
    with destino.open("xb") as archivo:
        archivo.write(resultado[0])


def esperar_marco_eventos(pagina, segundos: int = 45):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        for marco in pagina.frames:
            try:
                if marco.get_by_role("heading", name="Filtros de búsqueda de eventos").is_visible():
                    return marco
            except PlaywrightError:
                if not marco.is_detached():
                    raise
        pagina.wait_for_timeout(250)
    raise TimeoutError("No apareció el panel de eventos de Codelco")


def iniciar_sesion(pagina, usuario: str | None, clave: str | None) -> None:
    print("Ariba: abriendo página de acceso.", flush=True)
    pagina.goto(URL_INICIO, wait_until="domcontentloaded", timeout=60000)
    if not url_ariba(pagina.url):
        raise ValueError("Ariba redirigió a un sitio no autorizado")
    pagina.wait_for_function(
        """() => [...document.querySelectorAll('button, input')].some(elemento =>
            elemento.getClientRects().length > 0 &&
            (elemento.type === 'password' || (elemento.tagName === 'INPUT' &&
            ['text', 'email'].includes(elemento.type)) ||
            (elemento.tagName === 'BUTTON' && /Business Network/i.test(elemento.textContent))))""",
        timeout=60000,
    )
    dashboard = pagina.get_by_role("button", name=re.compile("Business Network", re.I)).first
    contrasena = pagina.locator('input[type="password"]:visible').first
    campo_usuario = pagina.locator('input[type="text"]:visible, input[type="email"]:visible, input:not([type]):visible').first
    if contrasena.is_visible() or campo_usuario.is_visible() and not dashboard.is_visible() or pagina.locator('#userid:visible').count():
        print("Ariba: autenticando sesión.", flush=True)
        if not usuario or not clave:
            raise RuntimeError("La sesión ha expirado y no hay credenciales configuradas")
        if campo_usuario.is_visible():
            if not url_ariba(pagina.url):
                raise ValueError("No se introducirá el usuario fuera de Ariba")
            campo_usuario.fill(usuario)
        if not contrasena.is_visible():
            siguiente = pagina.get_by_role("link", name=re.compile(r"siguiente|next", re.I)).or_(pagina.get_by_role("button", name=re.compile(r"siguiente|next", re.I)))
            siguiente.first.click()
            pagina.wait_for_function(
                """() => [...document.querySelectorAll('input[type=password]')].some(elemento => elemento.getClientRects().length > 0) ||
                    (![...document.querySelectorAll('#userid')].some(elemento => elemento.getClientRects().length > 0) &&
                    [...document.querySelectorAll('button')].some(elemento => elemento.getClientRects().length > 0 && /Business Network/i.test(elemento.textContent)))""",
                timeout=60000,
            )
        if contrasena.is_visible():
            if not url_ariba(pagina.url):
                raise ValueError("No se introducirá la clave fuera de Ariba")
            contrasena.fill(clave)
            enlace_ingreso = pagina.get_by_role("link", name=re.compile(r"conectarse|iniciar sesión|sign in", re.I))
            boton_ingreso = pagina.get_by_role("button", name=re.compile(r"conectarse|iniciar sesión|login|sign in", re.I))
            if enlace_ingreso.count():
                enlace_ingreso.first.click()
            elif boton_ingreso.count():
                boton_ingreso.first.click()
            else:
                raise RuntimeError("No se encontró el control para iniciar sesión en Ariba")
    dashboard.wait_for(timeout=60000)
    if contrasena.is_visible() or pagina.locator('#userid:visible').count():
        raise RuntimeError("No se pudo completar el inicio de sesión en Ariba")
    print("Ariba: sesión iniciada.", flush=True)


def seleccionar_codelco(pagina) -> None:
    print("Ariba: abriendo propuestas y seleccionando Codelco.", flush=True)
    pagina.get_by_role("button", name=re.compile("Business Network", re.I)).first.click()
    pagina.get_by_role("menuitem", name="Ariba Proposals & Questionnaires").click()
    pagina.get_by_role("link", name="Más...").wait_for(timeout=60000)
    pagina.get_by_role("link", name="Más...").click()
    pagina.get_by_role("link", name="Corporación Nacional del Cobre (Ambiente Productivo)").click()
    esperar_marco_eventos(pagina)
    print("Ariba: tabla de eventos disponible.", flush=True)


def exportar_eventos(pagina, destino: Path) -> None:
    seleccionar_codelco(pagina)
    marco = esperar_marco_eventos(pagina)
    etiqueta = marco.get_by_text("Estado del evento:", exact=True)
    fila = marco.locator("tr").filter(has=etiqueta).last
    selector = fila.get_by_role("combobox")
    valor = selector.get_by_role("textbox").first
    estado = lambda: valor.evaluate("elemento => elemento.tagName === 'INPUT' ? elemento.value : elemento.textContent").strip().casefold()
    if estado() != "abierto":
        selector.click()
        marco.get_by_role("option", name="Abierto", exact=True).click()
    if estado() != "abierto":
        raise ValueError("No fue posible verificar el filtro Abierto")
    print("Ariba: buscando eventos abiertos y preparando exportación.", flush=True)
    marco.get_by_role("button", name="Buscar", exact=True).click()
    opciones = marco.get_by_title("Menú Opciones de tabla").first
    exportar = marco.get_by_role("menuitem", name="Exportar todas las filas")
    for intento in range(3):
        if not exportar.is_visible():
            opciones.click()
        try:
            exportar.wait_for(state="visible", timeout=2000)
            break
        except PlaywrightTimeoutError:
            if intento == 2:
                raise ServicioTemporalError("No se pudo abrir el menú de exportación de eventos")
    print("Ariba: descargando exportación de eventos.", flush=True)
    with pagina.expect_download(timeout=60000) as descarga:
        exportar.click()
    descarga.value.save_as(destino)
    print("Ariba: exportación descargada.", flush=True)


def descargar_evento(pagina, evento, carpeta: Path) -> tuple[Path | None, str | None]:
    if not url_ariba(evento.enlace):
        raise ValueError(f"Enlace no autorizado para {evento.identificador}")
    print(f"{evento.identificador}: abriendo detalle del evento.", flush=True)
    respuesta = pagina.goto(evento.enlace, wait_until="domcontentloaded", timeout=60000)
    if respuesta and respuesta.status in (401, 403):
        raise SesionExpiradaError("Ariba requiere iniciar sesión nuevamente")
    if respuesta and (respuesta.status == 429 or 500 <= respuesta.status < 600):
        raise ServicioTemporalError(f"Ariba respondió con HTTP {respuesta.status}")
    if not url_ariba(pagina.url):
        raise ValueError("El detalle redirigió a un sitio no autorizado")
    if pagina.locator('#userid:visible, input[type="password"]:visible').count():
        raise SesionExpiradaError("Ariba redirigió al inicio de sesión")
    boton = pagina.get_by_role("button", name="Imprimir información de evento")
    boton.wait_for(timeout=60000)
    if not url_ariba(pagina.url):
        raise ValueError("El detalle redirigió a un sitio no autorizado")
    encabezado = pagina.get_by_text(re.compile(rf"^{re.escape(evento.identificador)}\s*-"))
    encabezado.first.wait_for(timeout=30000)
    with tempfile.TemporaryDirectory(dir=carpeta) as directorio:
        temporal = Path(directorio) / "evento.descarga"
        print(f"{evento.identificador}: solicitando impresión del evento.", flush=True)
        capturar_impresion(pagina, boton, temporal)
        fecha = fecha_vencimiento(temporal)
        if fecha is None:
            return None, f"{evento.identificador}: sin fecha de vencimiento; excluida"
        print(f"{evento.identificador}: documento recibido y fecha verificada.", flush=True)
        destino = carpeta / nombre_documento(fecha, evento.identificador, evento.titulo)
        if destino.exists():
            if destino.read_bytes() != temporal.read_bytes():
                raise FileExistsError(f"Ya existe un documento diferente: {destino.name}")
            return destino, None
        temporal.rename(destino)
        return destino, None


class DescargaPendienteError(RuntimeError):
    pass


class NavegadorCerradoError(RuntimeError):
    pass


def comprobar_pagina(pagina) -> None:
    if pagina is not None and pagina.is_closed():
        raise NavegadorCerradoError("La página de Ariba se cerró")


def detalle_fallo(error: Exception) -> str:
    texto = str(error)
    operacion = re.match(r"([A-Za-z]+\.[A-Za-z_]+):", texto)
    estado_http = re.search(r"\bHTTP (?:429|5\d\d)\b", texto)
    detalle = operacion.group(1) if operacion else estado_http.group(0) if estado_http else "sin detalle adicional"
    return f"{type(error).__name__}: {detalle}"


def reintentar_navegacion(accion, descripcion: str, recuperar_sesion=None, pagina=None):
    espera = 30
    intento = 0
    while True:
        try:
            if intento and recuperar_sesion is not None:
                recuperar_sesion()
            return accion()
        except (PlaywrightError, TimeoutError, SesionExpiradaError) as error:
            comprobar_pagina(pagina)
            intento += 1
            print(f"{descripcion}: {detalle_fallo(error)}; nuevo intento en {espera} segundos.", flush=True)
            time.sleep(espera)
            espera = min(espera * 2, 300)


def descargar_con_reintentos(pagina, evento, carpeta: Path, recuperar_sesion=None):
    for intento in range(1, 4):
        try:
            if intento > 1 and recuperar_sesion is not None:
                recuperar_sesion()
            return descargar_evento(pagina, evento, carpeta)
        except (PlaywrightError, TimeoutError, SesionExpiradaError) as error:
            comprobar_pagina(pagina)
            if intento == 3 or recuperar_sesion is None:
                raise DescargaPendienteError(f"No se pudo procesar {evento.identificador}; queda pendiente para otra ronda") from error
            print(f"Fallo de navegación en {evento.identificador} ({detalle_fallo(error)}); reintentando ({intento + 1}/3).", flush=True)


def configurar_credenciales() -> None:
    import keyring

    usuario = input("Correo de Ariba: ").strip()
    clave = getpass.getpass("Contraseña de Ariba: ")
    if not usuario or not clave:
        raise ValueError("El correo y la contraseña son obligatorios")
    keyring.set_password(SERVICIO, "usuario", usuario)
    keyring.set_password(SERVICIO, "contrasena", clave)
    print("Credenciales guardadas en el almacén de Windows.")