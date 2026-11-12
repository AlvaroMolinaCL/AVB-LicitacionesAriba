import getpass
import json
import smtplib
import ssl
import sys
import time
from datetime import date
from email.message import EmailMessage
from email.utils import parseaddr
from pathlib import Path
from src.archivos import bloqueo_ejecucion, guardar_estado, hash_archivo


BASE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent
SERVICIO_CORREO = "AVB-LicitacionesAriba-Correo"


def validar_remitente(remitente: str) -> str:
    remitente = remitente.strip()
    nombre, direccion = parseaddr(remitente)
    if nombre or direccion != remitente or remitente.count("@") != 1 or any(caracter.isspace() for caracter in remitente):
        raise ValueError("El remitente debe ser una dirección de correo válida")
    if not all(remitente.split("@")):
        raise ValueError("El remitente debe ser una dirección de correo válida")
    return remitente


def validar_destinatarios(destinatarios: list[str], copia: list[str]) -> tuple[list[str], list[str]]:
    if not isinstance(destinatarios, list) or not isinstance(copia, list) or not all(isinstance(direccion, str) for direccion in destinatarios + copia):
        raise ValueError("La configuración de destinatarios no es válida")
    if not destinatarios:
        raise ValueError("Debe configurar al menos un destinatario")
    destinatarios = [validar_remitente(direccion) for direccion in destinatarios]
    copia = [validar_remitente(direccion) for direccion in copia]
    direcciones = destinatarios + copia
    if len({direccion.casefold() for direccion in direcciones}) != len(direcciones):
        raise ValueError("No se permiten destinatarios duplicados")
    return destinatarios, copia


def crear_mensaje(remitente: str, fecha: date, resumen: Path, documentos_zip: Path, destinatarios: list[str], copia: list[str]) -> EmailMessage:
    remitente = validar_remitente(remitente)
    destinatarios, copia = validar_destinatarios(destinatarios, copia)
    contenido = resumen.read_bytes()
    texto = f"Reporte Licitaciones Codelco Ariba ({fecha:%d-%m-%Y})"
    mensaje = EmailMessage()
    mensaje["From"] = remitente
    mensaje["To"] = ", ".join(destinatarios)
    if copia:
        mensaje["Cc"] = ", ".join(copia)
    mensaje["Subject"] = texto
    mensaje.set_content(texto)
    mensaje.add_attachment(contenido, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=resumen.name)
    mensaje.add_attachment(documentos_zip.read_bytes(), maintype="application", subtype="zip", filename=documentos_zip.name)
    return mensaje


def conectar_correo(remitente: str, clave: str):
    servidor = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)
    try:
        servidor.ehlo()
        servidor.starttls(context=ssl.create_default_context())
        servidor.ehlo()
        servidor.login(validar_remitente(remitente), clave)
        return servidor
    except Exception:
        servidor.close()
        raise


def configurar_correo() -> None:
    with bloqueo_ejecucion():
        import keyring

        remitente = validar_remitente(input("Correo remitente: "))
        clave = getpass.getpass("Contraseña de aplicación de Google: ").replace(" ", "")
        if not clave:
            raise ValueError("La contraseña de aplicación es obligatoria")
        keyring.set_password(SERVICIO_CORREO, "remitente", remitente)
        keyring.set_password(SERVICIO_CORREO, "contrasena", clave)
        print("Credenciales de correo guardadas en el almacén de Windows.")


def configurar_destinatarios() -> None:
    with bloqueo_ejecucion():
        entrada_destinatarios = input("Destinatarios (separados por comas): ").strip()
        entrada_copia = input("Copia CC (separada por comas; Enter para ninguna): ").strip()
        destinatarios = [direccion.strip() for direccion in entrada_destinatarios.split(",")] if entrada_destinatarios else []
        copia = [direccion.strip() for direccion in entrada_copia.split(",")] if entrada_copia else []
        destinatarios, copia = validar_destinatarios(destinatarios, copia)
        ruta = BASE / ".datos" / "destinatarios_correo.json"
        ruta.parent.mkdir(parents=True, exist_ok=True)
        guardar_estado(ruta, {"destinatarios": destinatarios, "copia": copia})
        print("Destinatarios de correo guardados.")


def cargar_destinatarios() -> tuple[list[str], list[str]]:
    ruta = BASE / ".datos" / "destinatarios_correo.json"
    if not ruta.is_file():
        raise RuntimeError("Faltan destinatarios; ejecute --configurar-destinatarios")
    datos = json.loads(ruta.read_text(encoding="utf-8"))
    if not isinstance(datos, dict):
        raise ValueError("La configuración de destinatarios no es válida")
    return validar_destinatarios(datos.get("destinatarios"), datos.get("copia"))


def configurar_correo_interactivo_si_falta() -> None:
    import keyring

    remitente = keyring.get_password(SERVICIO_CORREO, "remitente")
    clave = keyring.get_password(SERVICIO_CORREO, "contrasena")
    if not remitente or not clave:
        print("Configure el remitente y la contraseña de aplicación.", flush=True)
        configurar_correo()
    if not (BASE / ".datos" / "destinatarios_correo.json").is_file():
        print("Configure los destinatarios y la copia del correo.", flush=True)
        configurar_destinatarios()


def enviar_resumen_guardado(resumen: Path, fecha: date, estado: dict, estado_path: Path, documentos_zip: Path | None = None) -> None:
    import keyring

    registro = estado["correo"]
    if registro["estado"] == "enviado":
        print("Correo ya enviado; no se enviará nuevamente.")
        return
    if registro["estado"] == "parcial":
        raise RuntimeError("El correo fue aceptado solo para algunos destinatarios; revise los rechazados antes de reintentar")
    if registro["estado"] != "pendiente":
        raise RuntimeError("El resultado del envío es incierto; revise el correo enviado antes de reintentar")
    if hash_archivo(resumen) != registro["sha256"]:
        raise ValueError("El resumen publicado cambió; no se enviará por correo")
    if documentos_zip is None or hash_archivo(documentos_zip) != registro.get("sha256_zip"):
        raise ValueError("El ZIP de documentos falta o cambió; no se enviará por correo")
    remitente = keyring.get_password(SERVICIO_CORREO, "remitente")
    clave = keyring.get_password(SERVICIO_CORREO, "contrasena")
    if not remitente or not clave:
        raise RuntimeError("Faltan credenciales de correo; ejecute --configurar-correo")
    remitente = validar_remitente(remitente)
    if "destinatarios" in registro:
        if registro.get("remitente") != remitente:
            raise ValueError("El remitente configurado cambió durante un envío pendiente")
        destinatarios, copia = validar_destinatarios(registro["destinatarios"], registro.get("copia"))
    elif registro.get("solo_remitente"):
        destinatarios, copia = [remitente], []
    else:
        destinatarios, copia = cargar_destinatarios()
    aceptados = registro.get("aceptados", [])
    if not isinstance(aceptados, list) or not all(isinstance(direccion, str) for direccion in aceptados):
        raise ValueError("El estado de destinatarios aceptados no es válido")
    pendientes = [direccion for direccion in destinatarios + copia if direccion not in aceptados]
    if not pendientes:
        raise ValueError("No hay destinatarios pendientes para el envío")
    mensaje = crear_mensaje(remitente, fecha, resumen, documentos_zip, destinatarios, copia)
    print("Correo: conectando con Gmail mediante TLS.", flush=True)
    servidor = conectar_correo(remitente, clave)
    try:
        registro.update(estado="enviando", remitente=remitente, destinatarios=destinatarios, copia=copia)
        guardar_estado(estado_path, estado)
        print("Correo: enviando resumen a los destinatarios configurados.", flush=True)
        try:
            rechazados = servidor.send_message(mensaje, from_addr=remitente, to_addrs=pendientes)
            if rechazados:
                registro["aceptados"] = aceptados + [direccion for direccion in pendientes if direccion not in rechazados]
                registro["rechazados"] = list(rechazados)
                registro["estado"] = "parcial"
                guardar_estado(estado_path, estado)
                print("Correo: destinatarios rechazados: " + ", ".join(registro["rechazados"]), flush=True)
                raise RuntimeError("Gmail aceptó el mensaje solo para algunos destinatarios; revise los rechazados antes de reintentar")
        except (smtplib.SMTPRecipientsRefused, smtplib.SMTPSenderRefused, smtplib.SMTPDataError):
            registro["estado"] = "pendiente"
            guardar_estado(estado_path, estado)
            raise
        registro["estado"] = "enviado"
        registro["aceptados"] = aceptados + pendientes
        registro.pop("rechazados", None)
        guardar_estado(estado_path, estado)
        print("Correo: Gmail aceptó el mensaje.", flush=True)
    finally:
        servidor.close()


def enviar_correo_con_reintentos(resumen: Path, fecha: date, estado: dict, estado_path: Path, documentos_zip: Path | None = None) -> None:
    espera = 30
    while True:
        try:
            enviar_resumen_guardado(resumen, fecha, estado, estado_path, documentos_zip)
            return
        except (OSError, smtplib.SMTPException) as error:
            if estado["correo"]["estado"] != "pendiente":
                raise
            if isinstance(error, smtplib.SMTPRecipientsRefused):
                temporal = bool(error.recipients) and all(400 <= codigo < 500 for codigo, _ in error.recipients.values())
            elif isinstance(error, smtplib.SMTPResponseException):
                temporal = 400 <= error.smtp_code < 500
            else:
                temporal = isinstance(error, (OSError, smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError))
            if not temporal:
                raise
            print(f"Correo: fallo temporal de conexión o rechazo; nuevo intento en {espera} segundos.", flush=True)
            time.sleep(espera)
            espera = min(espera * 2, 300)


def resolver_correo_incierto(estado: dict, estado_path: Path, decision: str) -> None:
    registro = estado.get("correo", {})
    if registro.get("estado") not in {"enviando", "parcial"}:
        raise ValueError("El correo no tiene un resultado incierto que resolver")
    if registro["estado"] == "parcial" and decision == "enviado":
        raise ValueError("Hay destinatarios rechazados; debe reintentar para completarlos")
    registro["estado"] = "enviado" if decision == "enviado" else "pendiente"
    guardar_estado(estado_path, estado)
    print(f"Correo: resultado registrado como {registro['estado']}.")


def ejecutar_resolucion_correo(base: Path, fecha: date, correo_de_prueba: bool, accion: str) -> None:
    from src.ejecucion import bloqueo_ejecucion

    if correo_de_prueba:
        estado_path = base / ".datos" / "pruebas-correo" / f"{fecha.isoformat()}.json"
    else:
        estado_path = base / ".datos" / "ejecuciones" / fecha.isoformat() / "estado.json"
    with bloqueo_ejecucion():
        estado = json.loads(estado_path.read_text(encoding="utf-8"))
        resolver_correo_incierto(estado, estado_path, accion)


def ejecutar_prueba_correo(base: Path, raiz: Path, fecha: date, nueva_prueba: bool) -> None:
    from src.ejecucion import bloqueo_ejecucion, preparar_zip_ejecucion

    raiz = raiz.resolve()
    resumen = raiz / "resumen_licitaciones" / f"{fecha.isoformat()}_Resumen_Licitaciones_Codelco.xlsx"
    with bloqueo_ejecucion():
        carpeta = base / ".datos" / "pruebas-correo"
        carpeta.mkdir(parents=True, exist_ok=True)
        estado_path = carpeta / f"{fecha.isoformat()}.json"
        if estado_path.exists():
            estado = json.loads(estado_path.read_text(encoding="utf-8"))
            if nueva_prueba:
                if estado["correo"]["estado"] != "enviado":
                    raise ValueError("La prueba anterior no está confirmada; resuelva su estado antes de iniciar otra")
                estado.setdefault("historial", []).append(estado["correo"].copy())
                estado["correo"] = {"estado": "pendiente", "sha256": hash_archivo(resumen)}
                guardar_estado(estado_path, estado)
        else:
            estado = {"correo": {"estado": "pendiente", "sha256": hash_archivo(resumen)}}
            guardar_estado(estado_path, estado)
        documentos_zip = None
        if estado["correo"]["estado"] == "pendiente":
            if "destinatarios" not in estado["correo"]:
                estado["correo"]["solo_remitente"] = True
                guardar_estado(estado_path, estado)
            estado_ejecucion_path = base / ".datos" / "ejecuciones" / fecha.isoformat() / "estado.json"
            estado_ejecucion = json.loads(estado_ejecucion_path.read_text(encoding="utf-8"))
            if not estado_ejecucion.get("completado") or estado_ejecucion.get("raiz") != str(raiz):
                raise ValueError("La ejecución del resumen no está completada o pertenece a otra carpeta")
            if estado_ejecucion.get("hash_resumen") != hash_archivo(resumen):
                raise ValueError("El resumen no coincide con la ejecución guardada")
            documentos_zip = preparar_zip_ejecucion(estado_ejecucion_path.parent, fecha, estado_ejecucion, estado_ejecucion_path)
            if "sha256_zip" not in estado["correo"]:
                estado["correo"]["sha256_zip"] = estado_ejecucion["hash_zip"]
                guardar_estado(estado_path, estado)
        enviar_correo_con_reintentos(resumen, fecha, estado, estado_path, documentos_zip)