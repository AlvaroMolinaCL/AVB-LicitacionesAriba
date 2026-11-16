import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from datetime import date, timedelta
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError

from src.actualizaciones import directorio_actual, publicar_actualizacion
from src.ariba import (
    DescargaPendienteError,
    NavegadorCerradoError,
    SesionExpiradaError,
    comprobar_pagina,
    descargar_con_reintentos,
    detalle_fallo,
    exportar_eventos,
    configurar_credenciales,
    iniciar_sesion,
    reintentar_navegacion,
    seleccionar_codelco,
)
from src.archivos import (
    bloqueo_ejecucion,
    copiar_sin_sobrescribir,
    crear_zip_documentos,
    guardar_estado,
    hash_archivo,
)
from src.consolidado import generar_consolidado, validar_formato_resumen
from src.pdf import convertir_documentos_pdf
from src.correo import (
    enviar_correo_con_reintentos,
    configurar_correo_interactivo_si_falta,
)
from src.reportes import (
    buscar_reporte_anterior,
    eventos_nuevos,
    guardar_reporte,
    leer_exportacion,
    leer_ids_anteriores,
)

BASE = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent.parent
)


def limpiar_ejecuciones(carpeta: Path, hoy: date) -> None:
    if not carpeta.exists():
        return
    limite = hoy - timedelta(days=7)
    for entrada in carpeta.iterdir():
        if (
            entrada.is_symlink()
            or getattr(entrada, "is_junction", lambda: False)()
            or not entrada.is_dir()
        ):
            continue
        try:
            fecha = date.fromisoformat(entrada.name)
        except ValueError:
            continue
        if entrada.name != fecha.isoformat() or fecha >= limite:
            continue
        shutil.rmtree(entrada)
        print(f"Ejecución antigua eliminada: {entrada.name}")


def asegurar_resultado(
    ruta: Path, clave: str, generar, estado: dict, estado_path: Path
) -> None:
    if ruta.exists():
        if clave not in estado or hash_archivo(ruta) != estado[clave]:
            raise ValueError(f"El resultado pendiente requiere revisión: {ruta}")
        return
    with tempfile.TemporaryDirectory(dir=ruta.parent) as temporal:
        generado = Path(temporal) / ruta.name
        generar(generado)
        copiar_sin_sobrescribir(generado, ruta)
    estado[clave] = hash_archivo(ruta)
    guardar_estado(estado_path, estado)


def preparar_zip_ejecucion(
    carpeta: Path, fecha: date, estado: dict, estado_path: Path
) -> Path:
    documentos = []
    for registro in estado["eventos"].values():
        if registro["estado"] != "descargado":
            continue
        nombre = registro["archivo"]
        if Path(nombre).name != nombre:
            raise ValueError("Ruta de documento inválida en el estado")
        documento = carpeta / nombre
        if hash_archivo(documento) != registro["sha256"]:
            raise ValueError(f"El documento cambió desde su descarga: {nombre}")
        documentos.append(documento)
    destino = carpeta / f"{fecha.isoformat()}_Documentos_Licitaciones_Codelco.zip"

    def generar_zip(ruta: Path) -> None:
        with tempfile.TemporaryDirectory(dir=ruta.parent) as temporal:
            pdfs = convertir_documentos_pdf(documentos, Path(temporal))
            crear_zip_documentos(ruta, pdfs)

    asegurar_resultado(
        destino,
        "hash_zip",
        generar_zip,
        estado,
        estado_path,
    )
    with zipfile.ZipFile(destino) as archivo:
        nombres = archivo.namelist()
        if (
            len(nombres) != len(documentos)
            or any(
                nombre not in {documento.name, documento.with_suffix(".pdf").name}
                for nombre, documento in zip(nombres, documentos)
            )
            or len({nombre.casefold() for nombre in nombres}) != len(nombres)
        ):
            raise ValueError(
                "El ZIP existente no contiene los documentos esperados; requiere revisión"
            )
    return destino


def preparar_zip_correo(fecha: date, estado: dict, estado_path: Path) -> Path | None:
    if estado["correo"]["estado"] != "pendiente":
        return None
    destino = preparar_zip_ejecucion(estado_path.parent, fecha, estado, estado_path)
    registro = estado["correo"]
    if "sha256_zip" not in registro:
        registro["sha256_zip"] = estado["hash_zip"]
        guardar_estado(estado_path, estado)
    return destino


def preparar_ejecucion(
    pagina, raiz: Path, hoy: date, recuperar_sesion=None
) -> tuple[dict, Path, Path]:
    carpeta = BASE / ".datos" / "ejecuciones" / hoy.isoformat()
    carpeta.mkdir(parents=True, exist_ok=True)
    carpeta, estado_previo = directorio_actual(carpeta)
    estado_path = carpeta / "estado.json"
    if estado_path.exists():
        estado = json.loads(estado_path.read_text(encoding="utf-8"))
        if estado["raiz"] != str(raiz.resolve()):
            raise ValueError("La ejecución existente pertenece a otra carpeta de datos")
        if not estado["completado"] or estado["eventos"] or "correo" in estado:
            print(
                f"Ejecución {hoy.isoformat()}: recuperando estado guardado.", flush=True
            )
            return estado, carpeta, estado_path
        estado["actualizacion"] = "actualizacion"
        guardar_estado(estado_path, estado)
        estado_previo = estado
        carpeta = carpeta / "actualizacion"
        carpeta.mkdir(exist_ok=True)
        estado_path = carpeta / "estado.json"
        print(
            "Sin eventos nuevos en la consulta anterior; actualizando exportación.",
            flush=True,
        )
    elif estado_previo is not None:
        if estado_previo["raiz"] != str(raiz.resolve()):
            raise ValueError("La ejecución existente pertenece a otra carpeta de datos")
        carpeta.mkdir(exist_ok=True)

    reportes = raiz / "reportes_diarios"
    for nombre in (
        f"{hoy.isoformat()}_Licitaciones_Codelco_(Ariba).xlsx",
        f"LIC ARIBA {hoy:%d-%m}.xlsx",
    ):
        destino = reportes / nombre
        if destino.exists() and not (
            estado_previo is not None
            and nombre == f"{hoy.isoformat()}_Licitaciones_Codelco_(Ariba).xlsx"
            and hash_archivo(destino) == estado_previo["hash_reporte"]
        ):
            raise FileExistsError(
                f"Ya existe el reporte de hoy; no se sobrescribirá: {destino}"
            )
    anterior = (
        Path(estado_previo["anterior"])
        if estado_previo and estado_previo["anterior"]
        else None if estado_previo else buscar_reporte_anterior(reportes, hoy)
    )
    if (
        anterior is not None
        and estado_previo
        and hash_archivo(anterior) != estado_previo["hash_anterior"]
    ):
        raise ValueError("El reporte anterior cambió durante la ejecución")
    if anterior is None:
        print(
            "Sin reporte anterior: todos los eventos se consideran nuevos.", flush=True
        )
        anteriores = []
    else:
        print(f"Comparando con reporte anterior: {anterior.name}.", flush=True)
        anteriores = leer_ids_anteriores(anterior)
    exportacion = carpeta / "exportacion.xls"
    if exportacion.exists():
        raise FileExistsError(
            f"Existe una exportación sin estado; revisar antes de continuar: {exportacion}"
        )
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
                    raise RuntimeError(
                        "No se pudo exportar la lista; vuelva a ejecutar el programa"
                    ) from error
                if intento % 3 == 0:
                    print(
                        f"Fallo al exportar eventos ({detalle_fallo(error)}); intento {intento}. "
                        f"El reporte y el correo siguen pendientes; nuevo intento en {espera} segundos. "
                        "Revise el último paso de Ariba indicado arriba; la causa aún no está confirmada.",
                        flush=True,
                    )
                    time.sleep(espera)
                    espera = min(espera * 2, 300)
                else:
                    print(
                        f"Fallo al exportar eventos ({detalle_fallo(error)}); reintentando.",
                        flush=True,
                    )
        copiar_sin_sobrescribir(descarga, exportacion)
    nuevos = eventos_nuevos(eventos, anteriores)
    print(
        f"Exportación validada: {len(eventos)} eventos abiertos, {len(nuevos)} nuevos.",
        flush=True,
    )
    estado = {
        "fecha": hoy.isoformat(),
        "raiz": str(raiz.resolve()),
        "anterior": str(anterior.resolve()) if anterior else None,
        "hash_anterior": hash_archivo(anterior) if anterior else None,
        "hash_exportacion": hash_archivo(exportacion),
        "eventos": {evento.identificador: {"estado": "pendiente"} for evento in nuevos},
        "completado": False,
    }
    if estado_previo is not None:
        estado["hash_reporte_previo"] = estado_previo["hash_reporte"]
    guardar_estado(estado_path, estado)
    return estado, carpeta, estado_path


def procesar_ejecucion(
    pagina,
    raiz: Path,
    hoy: date,
    recuperar_sesion=None,
    recuperar_exportacion=None,
    habilitar_correo=False,
) -> None:
    estado, carpeta, estado_path = preparar_ejecucion(
        pagina, raiz, hoy, recuperar_exportacion
    )
    if estado["completado"]:
        if habilitar_correo and "correo" in estado:
            resumen = (
                raiz
                / "resumen_licitaciones"
                / f"{hoy.isoformat()}_Resumen_Licitaciones_Codelco.xlsx"
            )
            documentos_zip = preparar_zip_correo(hoy, estado, estado_path)
            enviar_correo_con_reintentos(
                resumen, hoy, estado, estado_path, documentos_zip
            )
        else:
            print("La ejecución de hoy ya terminó; no se procesará nuevamente.")
        return
    anterior = Path(estado["anterior"]) if estado["anterior"] is not None else None
    exportacion = carpeta / "exportacion.xls"
    if (
        anterior is not None and hash_archivo(anterior) != estado["hash_anterior"]
    ) or hash_archivo(exportacion) != estado["hash_exportacion"]:
        raise ValueError("Los datos de referencia cambiaron durante la ejecución")
    anteriores = leer_ids_anteriores(anterior) if anterior is not None else []
    eventos = leer_exportacion(exportacion)
    nuevos = eventos_nuevos(eventos, anteriores)
    if set(estado["eventos"]) != {evento.identificador for evento in nuevos}:
        raise ValueError("El estado no coincide con los eventos de la exportación")

    espera = 30
    while True:
        pendientes = 0
        for numero, evento in enumerate(nuevos, start=1):
            registro = estado["eventos"][evento.identificador]
            if evento.titulo.strip().casefold().startswith("imcspot"):
                if registro["estado"] != "excluido":
                    if "hash_resumen" in estado:
                        raise ValueError(
                            "El consolidado existente requiere revisión: puede contener eventos IMCSPOT"
                        )
                    aviso = f"{evento.identificador}: evento IMCSPOT excluido"
                    registro.update(estado="excluido", aviso=aviso)
                    guardar_estado(estado_path, estado)
                    print(aviso)
                continue
            if registro["estado"] != "pendiente":
                continue
            print(
                f"[{numero}/{len(nuevos)}] Procesando {evento.identificador}",
                flush=True,
            )
            try:
                documento, aviso = descargar_con_reintentos(
                    pagina, evento, carpeta, recuperar_sesion
                )
            except DescargaPendienteError as error:
                if recuperar_sesion is None:
                    raise
                pendientes += 1
                causa = error.__cause__
                detalle = detalle_fallo(causa) if causa else "causa desconocida"
                print(
                    f"{evento.identificador}: pendiente ({detalle}); se intentará en otra ronda.",
                    flush=True,
                )
                continue
            if documento is None:
                registro.update(estado="excluido", aviso=aviso)
                print(aviso)
            else:
                registro.update(
                    estado="descargado",
                    archivo=documento.name,
                    sha256=hash_archivo(documento),
                )
                print(
                    f"{evento.identificador}: documento guardado y registrado.",
                    flush=True,
                )
            guardar_estado(estado_path, estado)
        if not pendientes:
            break
        print(
            f"Quedan {pendientes} eventos pendientes; nuevo intento en {espera} segundos.",
            flush=True,
        )
        time.sleep(espera)
        espera = min(espera * 2, 300)

    documentos = []
    eventos_por_documento = {}
    for evento in nuevos:
        registro = estado["eventos"][evento.identificador]
        if registro["estado"] == "excluido":
            continue
        if registro["estado"] != "descargado":
            raise ValueError(f"Evento sin procesar: {evento.identificador}")
        nombre = registro["archivo"]
        if Path(nombre).name != nombre:
            raise ValueError("Ruta de documento inválida en el estado")
        documento = carpeta / nombre
        if hash_archivo(documento) != registro["sha256"]:
            raise ValueError(f"El documento cambió desde su descarga: {nombre}")
        es_sr = bool(re.search(r"\bSR(?:\d|\b)", evento.titulo, re.I))
        documentos.append((documento, es_sr))
        eventos_por_documento[documento] = evento

    resumen_temporal = carpeta / f"{hoy.isoformat()}_Resumen_Licitaciones_Codelco.xlsx"
    if documentos:
        print(f"Generando resumen para {len(documentos)} documentos.", flush=True)
        asegurar_resultado(
            resumen_temporal,
            "hash_resumen",
            lambda ruta: generar_consolidado(ruta, documentos, eventos_por_documento),
            estado,
            estado_path,
        )
        validar_formato_resumen(resumen_temporal)
        print(f"Comprimiendo {len(documentos)} documentos.", flush=True)
        documentos_zip = preparar_zip_ejecucion(carpeta, hoy, estado, estado_path)
    reporte_temporal = carpeta / f"LIC ARIBA {hoy:%d-%m}.xlsx"
    asegurar_resultado(
        reporte_temporal,
        "hash_reporte",
        lambda ruta: guardar_reporte(ruta, eventos, anteriores),
        estado,
        estado_path,
    )
    if documentos:
        destino_resumen = raiz / "resumen_licitaciones" / resumen_temporal.name
        copiar_sin_sobrescribir(resumen_temporal, destino_resumen)
        print(f"Consolidado: {destino_resumen}")
    else:
        print("No hay documentos aptos para generar un consolidado.")
    destino_reporte = (
        raiz
        / "reportes_diarios"
        / f"{hoy.isoformat()}_Licitaciones_Codelco_(Ariba).xlsx"
    )
    if "hash_reporte_previo" in estado:
        publicar_actualizacion(
            reporte_temporal, destino_reporte, estado["hash_reporte_previo"]
        )
    else:
        copiar_sin_sobrescribir(reporte_temporal, destino_reporte)
    estado["completado"] = True
    if documentos and habilitar_correo:
        estado["correo"] = {
            "estado": "pendiente",
            "sha256": estado["hash_resumen"],
            "sha256_zip": estado["hash_zip"],
        }
    guardar_estado(estado_path, estado)
    print(
        f"Reporte: {destino_reporte}; nuevos: {len(nuevos)}; excluidos: {sum(registro['estado'] == 'excluido' for registro in estado['eventos'].values())}"
    )
    if documentos and habilitar_correo:
        enviar_correo_con_reintentos(
            destino_resumen, hoy, estado, estado_path, documentos_zip
        )


def ejecutar_flujo_principal(
    base: Path,
    servicio: str,
    raiz: Path,
    fecha: date | None,
    visible: bool,
    validar_ariba: bool,
    sin_pausa: bool,
) -> int:
    import keyring
    from playwright.sync_api import sync_playwright

    if getattr(sys, "frozen", False):
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(base / ".datos" / "navegadores")
    else:
        os.environ.setdefault(
            "PLAYWRIGHT_BROWSERS_PATH", str(base / ".datos" / "navegadores")
        )

    usuario = keyring.get_password(servicio, "usuario")
    clave = keyring.get_password(servicio, "contrasena")
    if getattr(sys, "frozen", False) and (not usuario or not clave):
        if sin_pausa or not sys.stdin or not sys.stdin.isatty():
            raise RuntimeError(
                "Faltan credenciales de Ariba; configúrelas en una consola interactiva"
            )
        print("No hay credenciales de Ariba configuradas.", flush=True)
        configurar_credenciales()
        usuario = keyring.get_password(servicio, "usuario")
        clave = keyring.get_password(servicio, "contrasena")

    perfil = base / ".datos" / "perfil-ariba"
    perfil.mkdir(parents=True, exist_ok=True)
    with bloqueo_ejecucion(), sync_playwright() as playwright:
        if (
            getattr(sys, "frozen", False)
            and not validar_ariba
            and not sin_pausa
            and sys.stdin
            and sys.stdin.isatty()
        ):
            configurar_correo_interactivo_si_falta()
        if not validar_ariba:
            limpiar_ejecuciones(base / ".datos" / "ejecuciones", date.today())
        fecha_ejecucion = fecha or date.today()
        estado_path = (
            base
            / ".datos"
            / "ejecuciones"
            / fecha_ejecucion.isoformat()
            / "estado.json"
        )
        if fecha and not estado_path.exists():
            raise FileNotFoundError("No existe una ejecución guardada para esa fecha")
        carpeta_actual, _ = directorio_actual(estado_path.parent)
        estado_actual_path = carpeta_actual / "estado.json"
        if not validar_ariba and estado_actual_path.exists():
            estado = json.loads(estado_actual_path.read_text(encoding="utf-8"))
            if estado.get("completado") and "correo" in estado:
                if estado["raiz"] != str(raiz.resolve()):
                    raise ValueError(
                        "La ejecución existente pertenece a otra carpeta de datos"
                    )
                resumen = (
                    raiz.resolve()
                    / "resumen_licitaciones"
                    / f"{fecha_ejecucion.isoformat()}_Resumen_Licitaciones_Codelco.xlsx"
                )
                documentos_zip = preparar_zip_correo(
                    fecha_ejecucion, estado, estado_actual_path
                )
                enviar_correo_con_reintentos(
                    resumen, fecha_ejecucion, estado, estado_actual_path, documentos_zip
                )
                return 0
        while True:
            print("Iniciando navegador de Ariba.", flush=True)
            contexto = reintentar_navegacion(
                lambda: playwright.chromium.launch_persistent_context(
                    str(perfil),
                    headless=not visible,
                    accept_downloads=True,
                    locale="es-CL",
                ),
                "Apertura del navegador",
            )
            pagina = None
            try:
                pagina = contexto.pages[0] if contexto.pages else contexto.new_page()
                reintentar_navegacion(
                    lambda: iniciar_sesion(pagina, usuario, clave),
                    "Inicio de sesión",
                    pagina=pagina,
                )
                if validar_ariba:
                    with tempfile.TemporaryDirectory(dir=base / ".datos") as temporal:
                        exportacion = Path(temporal) / "exportacion.xls"

                        def validar_exportacion():
                            exportacion.unlink(missing_ok=True)
                            exportar_eventos(pagina, exportacion)
                            return leer_exportacion(exportacion)

                        eventos = reintentar_navegacion(
                            validar_exportacion,
                            "Validación de exportación",
                            lambda: iniciar_sesion(pagina, usuario, clave),
                            pagina,
                        )
                        print(
                            f"Conexión y exportación verificadas: {len(eventos)} eventos abiertos."
                        )
                else:

                    def recuperar_sesion():
                        iniciar_sesion(pagina, usuario, clave)
                        seleccionar_codelco(pagina)

                    if estado_actual_path.exists() and not estado["completado"]:
                        reintentar_navegacion(
                            lambda: seleccionar_codelco(pagina),
                            "Selección de Codelco",
                            lambda: iniciar_sesion(pagina, usuario, clave),
                            pagina,
                        )
                    procesar_ejecucion(
                        pagina,
                        raiz.resolve(),
                        fecha_ejecucion,
                        recuperar_sesion,
                        lambda: iniciar_sesion(pagina, usuario, clave),
                        habilitar_correo=True,
                    )
                break
            except NavegadorCerradoError:
                print(
                    "El navegador se cerró; se reabrirá y se reanudará la ejecución en 30 segundos.",
                    flush=True,
                )
                time.sleep(30)
            finally:
                try:
                    contexto.close()
                except PlaywrightError:
                    if pagina is None or not pagina.is_closed():
                        raise
    return 0
