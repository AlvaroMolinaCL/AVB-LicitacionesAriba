import argparse
import getpass
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import time
from contextlib import contextmanager
from datetime import date, timedelta
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError

from src.ariba import SesionExpiradaError, descargar_evento, exportar_eventos, iniciar_sesion, seleccionar_codelco
from src.consolidado import generar_consolidado
from src.reportes import buscar_reporte_anterior, eventos_nuevos, guardar_reporte, leer_exportacion, leer_ids_anteriores


BASE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
SERVICIO = "AVB-LicitacionesAriba"


class DescargaPendienteError(RuntimeError):
    pass


class NavegadorCerradoError(RuntimeError):
    pass


def comprobar_pagina(pagina) -> None:
    if pagina is not None and pagina.is_closed():
        raise NavegadorCerradoError("La pagina de Ariba se cerro")


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


def limpiar_ejecuciones(carpeta: Path, hoy: date) -> None:
    if not carpeta.exists():
        return
    limite = hoy - timedelta(days=7)
    for entrada in carpeta.iterdir():
        if entrada.is_symlink() or getattr(entrada, "is_junction", lambda: False)() or not entrada.is_dir():
            continue
        try:
            fecha = date.fromisoformat(entrada.name)
        except ValueError:
            continue
        if entrada.name != fecha.isoformat() or fecha >= limite:
            continue
        shutil.rmtree(entrada)
        print(f"Ejecucion antigua eliminada: {entrada.name}")


def guardar_estado(ruta: Path, estado: dict) -> None:
    temporal = ruta.with_suffix(".tmp")
    temporal.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")
    temporal.replace(ruta)


def hash_archivo(ruta: Path) -> str:
    resumen = hashlib.sha256()
    with ruta.open("rb") as archivo:
        for bloque in iter(lambda: archivo.read(1024 * 1024), b""):
            resumen.update(bloque)
    return resumen.hexdigest()


def copiar_sin_sobrescribir(origen: Path, destino: Path) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists():
        if hash_archivo(origen) != hash_archivo(destino):
            raise FileExistsError(f"El archivo de destino ya existe y es diferente: {destino}")
        return
    with origen.open("rb") as entrada, destino.open("xb") as salida:
        shutil.copyfileobj(entrada, salida)


@contextmanager
def bloqueo_ejecucion():
    import msvcrt

    carpeta = BASE / ".datos"
    carpeta.mkdir(parents=True, exist_ok=True)
    with (carpeta / "ejecucion.lock").open("a+b") as archivo:
        archivo.seek(0)
        try:
            msvcrt.locking(archivo.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as error:
            raise RuntimeError("Ya hay otra ejecucion en curso") from error
        try:
            yield
        finally:
            archivo.seek(0)
            msvcrt.locking(archivo.fileno(), msvcrt.LK_UNLCK, 1)


def asegurar_resultado(ruta: Path, clave: str, generar, estado: dict, estado_path: Path) -> None:
    if ruta.exists():
        if clave not in estado or hash_archivo(ruta) != estado[clave]:
            raise ValueError(f"El resultado pendiente requiere revision: {ruta}")
        return
    with tempfile.TemporaryDirectory(dir=ruta.parent) as temporal:
        generado = Path(temporal) / ruta.name
        generar(generado)
        copiar_sin_sobrescribir(generado, ruta)
    estado[clave] = hash_archivo(ruta)
    guardar_estado(estado_path, estado)


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
            print(f"Fallo de navegacion en {evento.identificador} ({detalle_fallo(error)}); reintentando ({intento + 1}/3).", flush=True)


def preparar_ejecucion(pagina, raiz: Path, hoy: date, recuperar_sesion=None) -> tuple[dict, Path, Path]:
    carpeta = BASE / ".datos" / "ejecuciones" / hoy.isoformat()
    carpeta.mkdir(parents=True, exist_ok=True)
    estado_path = carpeta / "estado.json"
    if estado_path.exists():
        estado = json.loads(estado_path.read_text(encoding="utf-8"))
        if estado["raiz"] != str(raiz.resolve()):
            raise ValueError("La ejecucion existente pertenece a otra carpeta de datos")
        print(f"Ejecucion {hoy.isoformat()}: recuperando estado guardado.", flush=True)
        return estado, carpeta, estado_path

    reportes = raiz / "reportes_diarios"
    for nombre in (f"{hoy.isoformat()}_Licitaciones_Codelco_(Ariba).xlsx", f"LIC ARIBA {hoy:%d-%m}.xlsx"):
        destino = reportes / nombre
        if destino.exists():
            raise FileExistsError(f"Ya existe el reporte de hoy; no se sobrescribira: {destino}")
    anterior = buscar_reporte_anterior(reportes, hoy)
    print(f"Comparando con reporte anterior: {anterior.name}.", flush=True)
    anteriores = leer_ids_anteriores(anterior)
    exportacion = carpeta / "exportacion.xls"
    if exportacion.exists():
        raise FileExistsError(f"Existe una exportacion sin estado; revisar antes de continuar: {exportacion}")
    with tempfile.TemporaryDirectory(dir=carpeta) as temporal:
        descarga = Path(temporal) / "exportacion.xls"
        intento = 0
        espera = 30
        while True:
            intento += 1
            try:
                if intento > 1 and recuperar_sesion is not None:
                    recuperar_sesion()
                exportar_eventos(pagina, descarga)
                eventos = leer_exportacion(descarga)
                break
            except (PlaywrightError, TimeoutError, SesionExpiradaError) as error:
                comprobar_pagina(pagina)
                if descarga.exists():
                    descarga.unlink()
                if recuperar_sesion is None:
                    raise RuntimeError("No se pudo exportar la lista; vuelva a ejecutar el programa") from error
                if intento % 3 == 0:
                    print(f"Fallo al exportar eventos ({detalle_fallo(error)}); nuevo intento en {espera} segundos.", flush=True)
                    time.sleep(espera)
                    espera = min(espera * 2, 300)
                else:
                    print(f"Fallo al exportar eventos ({detalle_fallo(error)}); reintentando.", flush=True)
        copiar_sin_sobrescribir(descarga, exportacion)
    nuevos = eventos_nuevos(eventos, anteriores)
    print(f"Exportacion validada: {len(eventos)} eventos abiertos, {len(nuevos)} nuevos.", flush=True)
    estado = {
        "fecha": hoy.isoformat(),
        "raiz": str(raiz.resolve()),
        "anterior": str(anterior.resolve()),
        "hash_anterior": hash_archivo(anterior),
        "hash_exportacion": hash_archivo(exportacion),
        "eventos": {evento.identificador: {"estado": "pendiente"} for evento in nuevos},
        "completado": False,
    }
    guardar_estado(estado_path, estado)
    return estado, carpeta, estado_path


def procesar_ejecucion(pagina, raiz: Path, hoy: date, recuperar_sesion=None, recuperar_exportacion=None) -> None:
    estado, carpeta, estado_path = preparar_ejecucion(pagina, raiz, hoy, recuperar_exportacion)
    if estado["completado"]:
        print("La ejecucion de hoy ya termino; no se enviara ni procesara nuevamente.")
        return
    anterior = Path(estado["anterior"])
    exportacion = carpeta / "exportacion.xls"
    if hash_archivo(anterior) != estado["hash_anterior"] or hash_archivo(exportacion) != estado["hash_exportacion"]:
        raise ValueError("Los datos de referencia cambiaron durante la ejecucion")
    anteriores = leer_ids_anteriores(anterior)
    eventos = leer_exportacion(exportacion)
    nuevos = eventos_nuevos(eventos, anteriores)
    if set(estado["eventos"]) != {evento.identificador for evento in nuevos}:
        raise ValueError("El estado no coincide con los eventos de la exportacion")

    espera = 30
    while True:
        pendientes = 0
        for numero, evento in enumerate(nuevos, start=1):
            registro = estado["eventos"][evento.identificador]
            if evento.titulo.strip().casefold().startswith("imcspot"):
                if registro["estado"] != "excluido":
                    if "hash_resumen" in estado:
                        raise ValueError("El consolidado existente requiere revision: puede contener eventos IMCSPOT")
                    aviso = f"{evento.identificador}: evento IMCSPOT excluido"
                    registro.update(estado="excluido", aviso=aviso)
                    guardar_estado(estado_path, estado)
                    print(aviso)
                continue
            if registro["estado"] != "pendiente":
                continue
            print(f"[{numero}/{len(nuevos)}] Procesando {evento.identificador}", flush=True)
            try:
                documento, aviso = descargar_con_reintentos(pagina, evento, carpeta, recuperar_sesion)
            except DescargaPendienteError as error:
                if recuperar_sesion is None:
                    raise
                pendientes += 1
                causa = error.__cause__
                detalle = detalle_fallo(causa) if causa else "causa desconocida"
                print(f"{evento.identificador}: pendiente ({detalle}); se intentara en otra ronda.", flush=True)
                continue
            if documento is None:
                registro.update(estado="excluido", aviso=aviso)
                print(aviso)
            else:
                registro.update(estado="descargado", archivo=documento.name, sha256=hash_archivo(documento))
                print(f"{evento.identificador}: documento guardado y registrado.", flush=True)
            guardar_estado(estado_path, estado)
        if not pendientes:
            break
        print(f"Quedan {pendientes} eventos pendientes; nuevo intento en {espera} segundos.", flush=True)
        time.sleep(espera)
        espera = min(espera * 2, 300)

    documentos = []
    for evento in nuevos:
        registro = estado["eventos"][evento.identificador]
        if registro["estado"] == "excluido":
            continue
        if registro["estado"] != "descargado":
            raise ValueError(f"Evento sin procesar: {evento.identificador}")
        nombre = registro["archivo"]
        if Path(nombre).name != nombre:
            raise ValueError("Ruta de documento invalida en el estado")
        documento = carpeta / nombre
        if hash_archivo(documento) != registro["sha256"]:
            raise ValueError(f"El documento cambio desde su descarga: {nombre}")
        es_sr = bool(re.search(r"\bSR(?:\d|\b)", evento.titulo, re.I))
        documentos.append((documento, es_sr))

    resumen_temporal = carpeta / f"{hoy.isoformat()}_Resumen_Licitaciones_Codelco.xlsx"
    if documentos:
        print(f"Generando resumen para {len(documentos)} documentos.", flush=True)
        asegurar_resultado(resumen_temporal, "hash_resumen", lambda ruta: generar_consolidado(ruta, documentos), estado, estado_path)
    reporte_temporal = carpeta / f"LIC ARIBA {hoy:%d-%m}.xlsx"
    asegurar_resultado(reporte_temporal, "hash_reporte", lambda ruta: guardar_reporte(ruta, eventos, anteriores), estado, estado_path)
    if documentos:
        destino_resumen = raiz / "resumen_licitaciones" / resumen_temporal.name
        copiar_sin_sobrescribir(resumen_temporal, destino_resumen)
        print(f"Consolidado: {destino_resumen}")
    else:
        print("No hay documentos aptos para generar un consolidado.")
    destino_reporte = raiz / "reportes_diarios" / f"{hoy.isoformat()}_Licitaciones_Codelco_(Ariba).xlsx"
    copiar_sin_sobrescribir(reporte_temporal, destino_reporte)
    estado["completado"] = True
    guardar_estado(estado_path, estado)
    print(f"Reporte: {destino_reporte}; nuevos: {len(nuevos)}; excluidos: {sum(registro['estado'] == 'excluido' for registro in estado['eventos'].values())}")
    print("El envio de correo queda pendiente de configuracion; no se ha enviado ningun mensaje.")


def configurar_credenciales() -> None:
    import keyring

    usuario = input("Correo de Ariba: ").strip()
    clave = getpass.getpass("Contraseña de Ariba: ")
    if not usuario or not clave:
        raise ValueError("El correo y la contraseña son obligatorios")
    keyring.set_password(SERVICIO, "usuario", usuario)
    keyring.set_password(SERVICIO, "contrasena", clave)
    print("Credenciales guardadas en el almacen de Windows.")


def pausar_al_finalizar() -> None:
    if getattr(sys, "frozen", False) and "--sin-pausa" not in sys.argv and sys.stdin and sys.stdin.isatty():
        try:
            input("\nProceso finalizado. Presione Enter para cerrar esta ventana...")
        except EOFError:
            pass


def main() -> int:
    analizador = argparse.ArgumentParser(description="Automatizacion de licitaciones de Ariba")
    analizador.add_argument("--configurar-credenciales", action="store_true")
    analizador.add_argument("--visible", action="store_true")
    analizador.add_argument("--validar-ariba", action="store_true", help="Comprueba la conexion y la exportacion sin procesar licitaciones")
    analizador.add_argument("--sin-pausa", action="store_true", help="Evita la pausa final del ejecutable en tareas desatendidas")
    analizador.add_argument("--raiz", type=Path, default=BASE)
    analizador.add_argument("--fecha", type=date.fromisoformat, help="Reanudar una ejecucion anterior en formato AAAA-MM-DD")
    opciones = analizador.parse_args()
    if opciones.configurar_credenciales:
        configurar_credenciales()
        return 0
    if getattr(sys, "frozen", False):
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(BASE / ".datos" / "navegadores")
    else:
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(BASE / ".datos" / "navegadores"))
    import keyring
    from playwright.sync_api import sync_playwright

    usuario = keyring.get_password(SERVICIO, "usuario")
    clave = keyring.get_password(SERVICIO, "contrasena")
    if getattr(sys, "frozen", False) and (not usuario or not clave):
        if opciones.sin_pausa or not sys.stdin or not sys.stdin.isatty():
            raise RuntimeError("Faltan credenciales de Ariba; configurelas en una consola interactiva")
        print("No hay credenciales de Ariba configuradas.", flush=True)
        configurar_credenciales()
        usuario = keyring.get_password(SERVICIO, "usuario")
        clave = keyring.get_password(SERVICIO, "contrasena")
    perfil = BASE / ".datos" / "perfil-ariba"
    perfil.mkdir(parents=True, exist_ok=True)
    with bloqueo_ejecucion(), sync_playwright() as playwright:
        if not opciones.validar_ariba:
            limpiar_ejecuciones(BASE / ".datos" / "ejecuciones", date.today())
        fecha_ejecucion = opciones.fecha or date.today()
        if opciones.fecha and not (BASE / ".datos" / "ejecuciones" / fecha_ejecucion.isoformat() / "estado.json").exists():
            raise FileNotFoundError("No existe una ejecucion guardada para esa fecha")
        while True:
            print("Iniciando navegador de Ariba.", flush=True)
            contexto = reintentar_navegacion(
                lambda: playwright.chromium.launch_persistent_context(str(perfil), headless=not opciones.visible, accept_downloads=True, locale="es-CL"),
                "Apertura del navegador",
            )
            pagina = None
            try:
                pagina = contexto.pages[0] if contexto.pages else contexto.new_page()
                reintentar_navegacion(lambda: iniciar_sesion(pagina, usuario, clave), "Inicio de sesion", pagina=pagina)
                if opciones.validar_ariba:
                    with tempfile.TemporaryDirectory(dir=BASE / ".datos") as temporal:
                        exportacion = Path(temporal) / "exportacion.xls"
                        exportar_eventos(pagina, exportacion)
                        print(f"Conexion y exportacion verificadas: {len(leer_exportacion(exportacion))} eventos abiertos.")
                else:
                    def recuperar_sesion():
                        iniciar_sesion(pagina, usuario, clave)
                        seleccionar_codelco(pagina)

                    if (BASE / ".datos" / "ejecuciones" / fecha_ejecucion.isoformat() / "estado.json").exists():
                        reintentar_navegacion(lambda: seleccionar_codelco(pagina), "Seleccion de Codelco", lambda: iniciar_sesion(pagina, usuario, clave), pagina)
                    procesar_ejecucion(pagina, opciones.raiz.resolve(), fecha_ejecucion, recuperar_sesion, lambda: iniciar_sesion(pagina, usuario, clave))
                break
            except NavegadorCerradoError:
                print("El navegador se cerro; se reabrira y se reanudara la ejecucion en 30 segundos.", flush=True)
                time.sleep(30)
            finally:
                try:
                    contexto.close()
                except PlaywrightError:
                    if pagina is None or not pagina.is_closed():
                        raise
    return 0


if __name__ == "__main__":
    codigo_salida = 0
    try:
        codigo_salida = main()
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        codigo_salida = 1
    finally:
        pausar_al_finalizar()
    sys.exit(codigo_salida)