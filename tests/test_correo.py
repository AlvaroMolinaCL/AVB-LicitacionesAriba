import tempfile
import unittest
import json
import smtplib
import zipfile
from datetime import date
from contextlib import nullcontext
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook
from src import archivos, correo, ejecucion as modulo_ejecucion
from src.correo import conectar_correo, crear_mensaje, validar_destinatarios, validar_remitente
import principal


class PruebasCorreo(unittest.TestCase):
    def test_primera_ejecucion_interactiva_configura_datos_faltantes(self):
        with tempfile.TemporaryDirectory() as carpeta, patch.object(correo, "BASE", Path(carpeta)), patch("keyring.get_password", return_value=None), patch.object(correo, "configurar_correo") as configurar, patch.object(correo, "configurar_destinatarios") as destinatarios:
            correo.configurar_correo_interactivo_si_falta()
        configurar.assert_called_once()
        destinatarios.assert_called_once()

    def test_comando_configura_remitente_sin_predeterminado(self):
        with patch.object(correo, "bloqueo_ejecucion", return_value=nullcontext()), patch("sys.argv", ["principal", "--configurar-correo"]), patch("builtins.input", return_value="remitente@example.com") as leer, patch("getpass.getpass", return_value="clave ficticia"), patch("keyring.set_password") as guardar:
            self.assertEqual(principal.main(), 0)
        leer.assert_called_once_with("Correo remitente: ")
        self.assertEqual(guardar.call_count, 2)
        self.assertEqual(guardar.call_args_list[0].args[1:], ("remitente", "remitente@example.com"))
        self.assertEqual(guardar.call_args_list[1].args[1:], ("contrasena", "claveficticia"))

    def test_comando_configura_destinatarios_y_cc(self):
        with tempfile.TemporaryDirectory() as carpeta, patch.object(correo, "BASE", Path(carpeta)), patch.object(correo, "bloqueo_ejecucion", return_value=nullcontext()), patch("sys.argv", ["principal", "--configurar-destinatarios"]), patch("builtins.input", side_effect=["principal@example.com", "copia@example.com"]):
            self.assertEqual(principal.main(), 0)
            self.assertEqual(correo.cargar_destinatarios(), (["principal@example.com"], ["copia@example.com"]))

    def test_sin_destinatarios_no_inicia_conexion(self):
        with tempfile.TemporaryDirectory() as carpeta, patch.object(correo, "BASE", Path(carpeta)):
            raiz = Path(carpeta)
            resumen = raiz / "resumen.xlsx"
            documentos_zip = raiz / "documentos.zip"
            resumen.write_bytes(b"resumen")
            documentos_zip.write_bytes(b"zip")
            estado = {"correo": {"estado": "pendiente", "sha256": archivos.hash_archivo(resumen), "sha256_zip": archivos.hash_archivo(documentos_zip)}}
            with patch("keyring.get_password", side_effect=["remitente@example.com", "clave-ficticia"]), patch.object(correo, "conectar_correo") as conectar:
                with self.assertRaisesRegex(RuntimeError, "configurar-destinatarios"):
                    correo.enviar_resumen_guardado(resumen, date(2030, 11, 12), estado, raiz / "estado.json", documentos_zip)
            conectar.assert_not_called()
            self.assertEqual(estado["correo"]["estado"], "pendiente")

    def test_configurar_destinatarios_actualiza_archivo_local(self):
        with tempfile.TemporaryDirectory() as carpeta, patch.object(correo, "BASE", Path(carpeta)):
            with patch("builtins.input", side_effect=["uno@example.com, dos@example.com", "copia@example.com"]):
                principal.configurar_destinatarios()
            self.assertEqual(correo.cargar_destinatarios(), (["uno@example.com", "dos@example.com"], ["copia@example.com"]))
            with patch("builtins.input", side_effect=["otro@example.com", ""]):
                principal.configurar_destinatarios()
            self.assertEqual(correo.cargar_destinatarios(), (["otro@example.com"], []))
            with patch("builtins.input", side_effect=["", ""]), self.assertRaises(ValueError):
                principal.configurar_destinatarios()
            self.assertEqual(correo.cargar_destinatarios(), (["otro@example.com"], []))

    def test_mensaje_usa_destinatarios_y_copia_configurados(self):
        with tempfile.TemporaryDirectory() as carpeta:
            resumen = Path(carpeta) / "resumen.xlsx"
            documentos_zip = Path(carpeta) / "documentos.zip"
            resumen.write_bytes(b"resumen")
            documentos_zip.write_bytes(b"zip")
            mensaje = crear_mensaje("remitente@example.com", date(2030, 11, 12), resumen, documentos_zip, ["destino@example.com"], ["copia@example.com"])
            self.assertEqual(mensaje["To"], "destino@example.com")
            self.assertEqual(mensaje["Cc"], "copia@example.com")
            self.assertEqual(len(list(mensaje.iter_attachments())), 2)

    def test_rechazo_parcial_reintenta_solo_destinatarios_rechazados(self):
        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            resumen = carpeta / "resumen.xlsx"
            documentos_zip = carpeta / "documentos.zip"
            resumen.write_bytes(b"resumen")
            documentos_zip.write_bytes(b"zip")
            estado_path = carpeta / "estado.json"
            estado = {"correo": {"estado": "pendiente", "sha256": archivos.hash_archivo(resumen), "sha256_zip": archivos.hash_archivo(documentos_zip)}}
            servidor = Mock()
            servidor.send_message.side_effect = [{"copia@example.com": (450, b"temporal")}, {}]
            with patch("keyring.get_password", side_effect=["remitente@example.com", "clave-ficticia"] * 2), patch.object(correo, "cargar_destinatarios", return_value=(["principal@example.com"], ["copia@example.com"])) as cargar, patch.object(correo, "conectar_correo", return_value=servidor):
                with self.assertRaisesRegex(RuntimeError, "algunos destinatarios"):
                    correo.enviar_resumen_guardado(resumen, date(2030, 11, 12), estado, estado_path, documentos_zip)
                self.assertEqual(estado["correo"]["estado"], "parcial")
                self.assertEqual(estado["correo"]["aceptados"], ["principal@example.com"])
                with self.assertRaises(ValueError):
                    correo.resolver_correo_incierto(estado, estado_path, "enviado")
                correo.resolver_correo_incierto(estado, estado_path, "reintentar")
                correo.enviar_resumen_guardado(resumen, date(2030, 11, 12), estado, estado_path, documentos_zip)
            self.assertEqual(servidor.send_message.call_args_list[0].kwargs["to_addrs"], ["principal@example.com", "copia@example.com"])
            self.assertEqual(servidor.send_message.call_args_list[1].kwargs["to_addrs"], ["copia@example.com"])
            self.assertEqual(estado["correo"]["estado"], "enviado")
            cargar.assert_called_once()

    def test_valida_destinatarios_obligatorios_y_sin_duplicados(self):
        with self.assertRaises(ValueError):
            validar_destinatarios([], [])
        with self.assertRaises(ValueError):
            validar_destinatarios(["destino@example.com"], ["DESTINO@example.com"])
        self.assertEqual(validar_destinatarios(["destino@example.com"], []), (["destino@example.com"], []))

    def test_zip_contiene_solo_los_documentos_indicados(self):
        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            primero = carpeta / "Doc1.doc"
            segundo = carpeta / "Doc2.doc"
            primero.write_bytes(b"documento uno")
            segundo.write_bytes(b"documento dos")
            (carpeta / "otro.doc").write_bytes(b"no incluido")
            destino = carpeta / "2030-11-12_Documentos_Licitaciones_Codelco.zip"
            archivos.crear_zip_documentos(destino, [primero, segundo])
            with zipfile.ZipFile(destino) as archivo:
                self.assertEqual(archivo.namelist(), ["Doc1.doc", "Doc2.doc"])
                self.assertEqual(archivo.read("Doc1.doc"), b"documento uno")
                self.assertEqual(archivo.read("Doc2.doc"), b"documento dos")
            with self.assertRaises(FileExistsError):
                archivos.crear_zip_documentos(destino, [primero])
            with self.assertRaises(ValueError):
                archivos.crear_zip_documentos(carpeta / "vacio.zip", [])

    def test_no_envia_si_el_zip_fue_modificado(self):
        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            resumen = carpeta / "resumen.xlsx"
            resumen.write_bytes(b"resumen")
            documentos_zip = carpeta / "documentos.zip"
            documentos_zip.write_bytes(b"original")
            estado = {"correo": {"estado": "pendiente", "sha256": archivos.hash_archivo(resumen), "sha256_zip": archivos.hash_archivo(documentos_zip)}}
            documentos_zip.write_bytes(b"modificado")
            with patch.object(correo, "conectar_correo") as conectar:
                with self.assertRaisesRegex(ValueError, "ZIP.*cambió"):
                    correo.enviar_resumen_guardado(resumen, date(2030, 11, 12), estado, carpeta / "estado.json", documentos_zip)
            conectar.assert_not_called()

    def test_no_comprime_un_documento_modificado(self):
        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            documento = carpeta / "evento.doc"
            documento.write_bytes(b"original")
            estado = {"eventos": {"Doc1": {"estado": "descargado", "archivo": documento.name, "sha256": archivos.hash_archivo(documento)}}}
            documento.write_bytes(b"modificado")
            with self.assertRaisesRegex(ValueError, "cambió"):
                modulo_ejecucion.preparar_zip_ejecucion(carpeta, date(2030, 11, 12), estado, carpeta / "estado.json")
            self.assertFalse((carpeta / "2030-11-12_Documentos_Licitaciones_Codelco.zip").exists())

    def test_fallo_temporal_antes_de_enviar_reintenta_sin_duplicar(self):
        estado = {"correo": {"estado": "pendiente"}}
        enviar = Mock(side_effect=[OSError("conexion interrumpida"), None])
        with patch.object(correo, "enviar_resumen_guardado", enviar), patch.object(correo.time, "sleep") as esperar:
            correo.enviar_correo_con_reintentos(Path("resumen.xlsx"), date(2030, 11, 12), estado, Path("estado.json"))
        self.assertEqual(enviar.call_count, 2)
        esperar.assert_called_once_with(30)

    def test_estado_incierto_solo_se_resuelve_explicitamente(self):
        with tempfile.TemporaryDirectory() as carpeta:
            estado_path = Path(carpeta) / "estado.json"
            estado = {"correo": {"estado": "enviando"}}
            correo.resolver_correo_incierto(estado, estado_path, "reintentar")
            self.assertEqual(json.loads(estado_path.read_text(encoding="utf-8"))["correo"]["estado"], "pendiente")
            with self.assertRaises(ValueError):
                correo.resolver_correo_incierto(estado, estado_path, "enviado")
            estado["correo"]["estado"] = "enviando"
            correo.resolver_correo_incierto(estado, estado_path, "enviado")
            self.assertEqual(json.loads(estado_path.read_text(encoding="utf-8"))["correo"]["estado"], "enviado")

    def test_ejecucion_publica_y_reanuda_el_envio_pendiente(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            reportes = raiz / "reportes_diarios"
            reportes.mkdir()
            libro = Workbook()
            libro.active.append(["Título", "ID"])
            libro.active.append(["Anterior", "Doc0"])
            libro.save(reportes / "2030-11-11_Licitaciones_Codelco_(Ariba).xlsx")
            libro.close()
            exportacion = '''<div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div><table class="tableBody"><tr><td><a href="https://s3.ariba.com/1">6001</a></td><td>Doc1</td><td></td><td>Abierto</td><td>RFP</td></tr></table>'''
            def exportar(pagina, destino):
                destino.write_text(exportacion, encoding="utf-8")
            def descargar(pagina, evento, destino):
                documento = destino / "evento.doc"
                documento.write_text('<html><table><tr><td class="titleFieldValue">7 Articulo</td><td></td></tr><tr><td>Cantidad</td><td>1 Unidad</td></tr></table></html>', encoding="utf-8")
                return documento, None
            fecha = date(2030, 11, 12)
            with patch.object(principal, "BASE", raiz), patch("src.ejecucion.BASE", raiz), patch.object(correo, "BASE", raiz), patch("src.ejecucion.exportar_eventos", side_effect=exportar), patch("src.ariba.descargar_evento", side_effect=descargar), patch("keyring.get_password", return_value=None):
                with self.assertRaisesRegex(RuntimeError, "credenciales de correo"):
                    modulo_ejecucion.procesar_ejecucion(None, raiz, fecha, habilitar_correo=True)
            estado_path = raiz / ".datos" / "ejecuciones" / "2030-11-12" / "estado.json"
            estado = json.loads(estado_path.read_text(encoding="utf-8"))
            self.assertTrue(estado["completado"])
            self.assertEqual(estado["correo"]["estado"], "pendiente")
            servidor = Mock()
            servidor.send_message.return_value = {}
            with patch.object(principal, "BASE", raiz), patch("src.ejecucion.BASE", raiz), patch.object(correo, "BASE", raiz), patch("src.ejecucion.exportar_eventos", side_effect=AssertionError("no debe exportar")), patch("keyring.get_password", side_effect=["remitente@example.com", "clave-ficticia"]), patch.object(correo, "cargar_destinatarios", return_value=(["destino@example.com"], ["copia@example.com"])), patch.object(correo, "conectar_correo", return_value=servidor):
                modulo_ejecucion.procesar_ejecucion(None, raiz, fecha, habilitar_correo=True)
                modulo_ejecucion.procesar_ejecucion(None, raiz, fecha, habilitar_correo=True)
            self.assertEqual(json.loads(estado_path.read_text(encoding="utf-8"))["correo"]["estado"], "enviado")
            servidor.send_message.assert_called_once()
            self.assertEqual(servidor.send_message.call_args.kwargs["to_addrs"], ["destino@example.com", "copia@example.com"])
            adjuntos = list(servidor.send_message.call_args.args[0].iter_attachments())
            self.assertEqual([adjunto.get_filename() for adjunto in adjuntos], ["2030-11-12_Resumen_Licitaciones_Codelco.xlsx", "2030-11-12_Documentos_Licitaciones_Codelco.zip"])
            archivo_zip = raiz / ".datos" / "ejecuciones" / "2030-11-12" / "2030-11-12_Documentos_Licitaciones_Codelco.zip"
            with zipfile.ZipFile(archivo_zip) as archivo:
                self.assertEqual(archivo.namelist(), ["evento.doc"])

    def test_prueba_manual_envia_resumen_existente_sin_ariba(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            resumenes = raiz / "resumen_licitaciones"
            resumenes.mkdir()
            resumen = resumenes / "2030-11-12_Resumen_Licitaciones_Codelco.xlsx"
            resumen.write_bytes(b"xlsx-ficticio")
            ejecucion = raiz / ".datos" / "ejecuciones" / "2030-11-12"
            ejecucion.mkdir(parents=True)
            documento = ejecucion / "evento.doc"
            documento.write_bytes(b"documento-ficticio")
            estado_ejecucion = {"completado": True, "raiz": str(raiz.resolve()), "hash_resumen": archivos.hash_archivo(resumen), "eventos": {"Doc1": {"estado": "descargado", "archivo": documento.name, "sha256": archivos.hash_archivo(documento)}}}
            archivos.guardar_estado(ejecucion / "estado.json", estado_ejecucion)
            servidor = Mock()
            servidor.send_message.return_value = {}
            with patch.object(principal, "BASE", raiz), patch("src.ejecucion.BASE", raiz), patch.object(correo, "BASE", raiz), patch.object(modulo_ejecucion, "bloqueo_ejecucion", return_value=nullcontext()), patch("sys.argv", ["principal", "--probar-correo", "--fecha", "2030-11-12", "--raiz", str(raiz)]), patch("keyring.get_password", side_effect=["remitente@example.com", "clave-ficticia"]), patch.object(correo, "conectar_correo", return_value=servidor), patch.object(modulo_ejecucion, "iniciar_sesion", side_effect=AssertionError("no debe acceder a Ariba")):
                self.assertEqual(principal.main(), 0)
                self.assertEqual(principal.main(), 0)
            servidor.send_message.assert_called_once()
            mensaje = servidor.send_message.call_args.args[0]
            self.assertEqual(mensaje["To"], "remitente@example.com")
            self.assertEqual(mensaje["Subject"], "Reporte Licitaciones Codelco Ariba (12-11-2030)")
            self.assertEqual([adjunto.get_filename() for adjunto in mensaje.iter_attachments()], [resumen.name, "2030-11-12_Documentos_Licitaciones_Codelco.zip"])
            with patch.object(principal, "BASE", raiz), patch("src.ejecucion.BASE", raiz), patch.object(correo, "BASE", raiz), patch.object(modulo_ejecucion, "bloqueo_ejecucion", return_value=nullcontext()), patch("sys.argv", ["principal", "--probar-correo", "--nueva-prueba", "--fecha", "2030-11-12", "--raiz", str(raiz)]), patch("keyring.get_password", side_effect=["remitente@example.com", "clave-ficticia"]), patch.object(correo, "conectar_correo", return_value=servidor):
                self.assertEqual(principal.main(), 0)
            self.assertEqual(servidor.send_message.call_count, 2)
            estado_prueba = json.loads((raiz / ".datos" / "pruebas-correo" / "2030-11-12.json").read_text(encoding="utf-8"))
            self.assertEqual(estado_prueba["historial"][0]["estado"], "enviado")
            self.assertEqual(estado_prueba["correo"]["estado"], "enviado")

    def test_envio_exitoso_se_registra_y_no_se_repite(self):
        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            resumen = carpeta / "resumen.xlsx"
            resumen.write_bytes(b"xlsx-ficticio")
            documentos_zip = carpeta / "documentos.zip"
            documentos_zip.write_bytes(b"zip-ficticio")
            estado_path = carpeta / "estado.json"
            estado = {"correo": {"estado": "pendiente", "sha256": archivos.hash_archivo(resumen), "sha256_zip": archivos.hash_archivo(documentos_zip)}}
            servidor = Mock()
            servidor.send_message.return_value = {}
            with patch("keyring.get_password", side_effect=["remitente@example.com", "clave-ficticia"]), patch.object(correo, "cargar_destinatarios", return_value=(["remitente@example.com"], [])), patch.object(correo, "conectar_correo", return_value=servidor):
                correo.enviar_resumen_guardado(resumen, date(2030, 11, 12), estado, estado_path, documentos_zip)
                correo.enviar_resumen_guardado(resumen, date(2030, 11, 12), estado, estado_path, documentos_zip)
            self.assertEqual(json.loads(estado_path.read_text(encoding="utf-8"))["correo"]["estado"], "enviado")
            servidor.send_message.assert_called_once()
            self.assertEqual(servidor.send_message.call_args.kwargs["to_addrs"], ["remitente@example.com"])
            servidor.close.assert_called_once()

    def test_fallo_incierto_no_se_reenvia_automaticamente(self):
        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            resumen = carpeta / "resumen.xlsx"
            resumen.write_bytes(b"xlsx-ficticio")
            documentos_zip = carpeta / "documentos.zip"
            documentos_zip.write_bytes(b"zip-ficticio")
            estado_path = carpeta / "estado.json"
            estado = {"correo": {"estado": "pendiente", "sha256": archivos.hash_archivo(resumen), "sha256_zip": archivos.hash_archivo(documentos_zip)}}
            servidor = Mock()
            servidor.send_message.side_effect = smtplib.SMTPServerDisconnected("conexion perdida")
            with patch("keyring.get_password", side_effect=["remitente@example.com", "clave-ficticia"]), patch.object(correo, "cargar_destinatarios", return_value=(["remitente@example.com"], [])), patch.object(correo, "conectar_correo", return_value=servidor):
                with self.assertRaises(smtplib.SMTPServerDisconnected):
                    correo.enviar_resumen_guardado(resumen, date(2030, 11, 12), estado, estado_path, documentos_zip)
                estado = json.loads(estado_path.read_text(encoding="utf-8"))
                with self.assertRaisesRegex(RuntimeError, "incierto"):
                    correo.enviar_resumen_guardado(resumen, date(2030, 11, 12), estado, estado_path, documentos_zip)
            self.assertEqual(estado["correo"]["estado"], "enviando")
            servidor.send_message.assert_called_once()

    def test_rechazo_explicito_permite_reintentar(self):
        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            resumen = carpeta / "resumen.xlsx"
            resumen.write_bytes(b"xlsx-ficticio")
            documentos_zip = carpeta / "documentos.zip"
            documentos_zip.write_bytes(b"zip-ficticio")
            estado_path = carpeta / "estado.json"
            estado = {"correo": {"estado": "pendiente", "sha256": archivos.hash_archivo(resumen), "sha256_zip": archivos.hash_archivo(documentos_zip)}}
            servidor = Mock()
            servidor.send_message.side_effect = smtplib.SMTPDataError(451, b"error temporal")
            with patch("keyring.get_password", side_effect=["remitente@example.com", "clave-ficticia"]), patch.object(correo, "cargar_destinatarios", return_value=(["remitente@example.com"], [])), patch.object(correo, "conectar_correo", return_value=servidor):
                with self.assertRaises(smtplib.SMTPDataError):
                    correo.enviar_resumen_guardado(resumen, date(2030, 11, 12), estado, estado_path, documentos_zip)
            self.assertEqual(json.loads(estado_path.read_text(encoding="utf-8"))["correo"]["estado"], "pendiente")

    def test_mensaje_solo_para_el_remitente_con_dos_adjuntos(self):
        with tempfile.TemporaryDirectory() as carpeta:
            resumen = Path(carpeta) / "2030-11-12_Resumen_Licitaciones_Codelco.xlsx"
            resumen.write_bytes(b"xlsx-ficticio")
            documentos_zip = Path(carpeta) / "2030-11-12_Documentos_Licitaciones_Codelco.zip"
            documentos_zip.write_bytes(b"zip-ficticio")
            mensaje = crear_mensaje("remitente@example.com", date(2030, 11, 12), resumen, documentos_zip, ["remitente@example.com"], [])
        self.assertEqual(mensaje["From"], "remitente@example.com")
        self.assertEqual(mensaje["To"], "remitente@example.com")
        self.assertIsNone(mensaje["Cc"])
        self.assertEqual(mensaje["Subject"], "Reporte Licitaciones Codelco Ariba (12-11-2030)")
        self.assertEqual(mensaje.get_body(preferencelist=("plain",)).get_content().strip(), mensaje["Subject"])
        adjuntos = list(mensaje.iter_attachments())
        self.assertEqual(len(adjuntos), 2)
        self.assertEqual(adjuntos[0].get_filename(), resumen.name)
        self.assertEqual(adjuntos[0].get_payload(decode=True), b"xlsx-ficticio")
        self.assertEqual(adjuntos[1].get_filename(), documentos_zip.name)
        self.assertEqual(adjuntos[1].get_payload(decode=True), b"zip-ficticio")

    def test_conexion_usa_tls_y_credenciales_configuradas(self):
        servidor = Mock()
        with patch("src.correo.smtplib.SMTP", return_value=servidor) as smtp:
            self.assertIs(conectar_correo("remitente@example.com", "clave-ficticia"), servidor)
        smtp.assert_called_once_with("smtp.gmail.com", 587, timeout=30)
        self.assertEqual(servidor.ehlo.call_count, 2)
        servidor.starttls.assert_called_once()
        servidor.login.assert_called_once_with("remitente@example.com", "clave-ficticia")

    def test_direccion_no_admite_varios_destinatarios(self):
        for invalida in ("", "uno@example.com, otro@example.com", "Nombre <uno@example.com>", "uno@example.com\nCc: otro@example.com"):
            with self.subTest(invalida=invalida), self.assertRaises(ValueError):
                validar_remitente(invalida)


if __name__ == "__main__":
    unittest.main()