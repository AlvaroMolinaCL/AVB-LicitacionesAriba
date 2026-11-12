import json
import os
import runpy
import subprocess
import sys
import tempfile
import unittest
import zipfile
from contextlib import nullcontext
from datetime import date
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

import principal
from src import ariba, archivos, consolidado, ejecucion as modulo_ejecucion


class PruebasFlujo(unittest.TestCase):
    @unittest.skipUnless(os.name == "nt", "Requiere Windows")
    def test_salida_redirigida_usa_utf8_aunque_el_entorno_indique_cp1252(self):
        ruta = Path(__file__).resolve().parent.parent / "principal.py"
        entorno = os.environ.copy()
        entorno["PYTHONIOENCODING"] = "cp1252"
        resultado = subprocess.run([sys.executable, str(ruta), "--help"], capture_output=True, env=entorno, check=True)
        salida = resultado.stdout.decode("utf-8")
        self.assertIn("Automatización", salida)
        self.assertIn("contraseña de aplicación", salida)

    def test_configurar_consola_windows_usa_utf8_en_ambos_flujos(self):
        salida = Mock()
        errores = Mock()
        with patch("sys.platform", "win32"), patch("sys.stdout", salida), patch("sys.stderr", errores), patch.object(principal.os, "system") as sistema:
            principal.configurar_consola()
        sistema.assert_called_once_with("chcp 65001 > nul")
        salida.reconfigure.assert_called_once_with(encoding="utf-8", errors="backslashreplace")
        errores.reconfigure.assert_called_once_with(encoding="utf-8", errors="backslashreplace")

    def test_configurar_consola_no_modifica_otras_plataformas(self):
        with patch("sys.platform", "linux"), patch.object(principal.os, "system") as sistema:
            principal.configurar_consola()
        sistema.assert_not_called()

    def test_resumen_antiguo_no_se_publica_como_actual(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "resumen.xlsx"
            libro = Workbook()
            libro.active.append(["Archivo", "Región", "Ítem Principal", "N° Material", "Cant.", "Entrega", "Texto largo"])
            libro.save(ruta)
            libro.close()
            with self.assertRaisesRegex(ValueError, "formato anterior"):
                consolidado.validar_formato_resumen(ruta)

    def test_base_es_la_raiz_del_proyecto(self):
        self.assertEqual(principal.BASE, Path(__file__).resolve().parent.parent)

    def test_base_del_ejecutable_es_su_carpeta(self):
        ejecutable = Path(tempfile.gettempdir()) / "aplicacion" / "respaldo.exe"
        with patch("sys.frozen", True, create=True), patch("sys.executable", str(ejecutable)):
            modulo = runpy.run_path(str(Path(__file__).resolve().parent.parent / "principal.py"))
        self.assertEqual(modulo["BASE"], ejecutable.resolve().parent)

    def test_pausa_solo_en_ejecutable_interactivo(self):
        entrada = Mock()
        entrada.isatty.return_value = True
        with patch("sys.frozen", True, create=True), patch("sys.stdin", entrada), patch("sys.argv", ["respaldo.exe"]), patch("builtins.input") as leer:
            principal.pausar_al_finalizar()
            leer.assert_called_once()
        with patch("sys.frozen", True, create=True), patch("sys.stdin", entrada), patch("sys.argv", ["respaldo.exe", "--sin-pausa"]), patch("builtins.input") as leer:
            principal.pausar_al_finalizar()
            leer.assert_not_called()

    def test_ejecutable_pide_credenciales_y_usa_navegador_oculto_local(self):
        pagina = Mock()
        contexto = Mock(pages=[pagina])
        playwright = Mock()
        playwright.chromium.launch_persistent_context.return_value = contexto
        entrada = Mock()
        entrada.isatty.return_value = True
        with tempfile.TemporaryDirectory() as carpeta, patch.object(principal, "BASE", Path(carpeta)), patch.object(modulo_ejecucion, "BASE", Path(carpeta)), patch.object(modulo_ejecucion, "bloqueo_ejecucion", return_value=nullcontext()), patch("playwright.sync_api.sync_playwright", return_value=nullcontext(playwright)), patch("keyring.get_password", side_effect=[None, None, "usuario-ficticio", "clave-ficticia"]), patch("sys.frozen", True, create=True), patch("sys.stdin", entrada), patch("sys.argv", ["principal"]), patch.dict(os.environ, {"PLAYWRIGHT_BROWSERS_PATH": "ruta-externa"}), patch.object(modulo_ejecucion, "configurar_credenciales") as configurar, patch.object(modulo_ejecucion, "configurar_correo_interactivo_si_falta"), patch.object(modulo_ejecucion, "iniciar_sesion") as login, patch.object(modulo_ejecucion, "procesar_ejecucion"):
            self.assertEqual(principal.main(), 0)
            self.assertEqual(os.environ["PLAYWRIGHT_BROWSERS_PATH"], str(Path(carpeta) / ".datos" / "navegadores"))
        configurar.assert_called_once()
        login.assert_called_once_with(pagina, "usuario-ficticio", "clave-ficticia")
        self.assertTrue(playwright.chromium.launch_persistent_context.call_args.kwargs["headless"])

    def test_ejecutable_sin_consola_no_pide_credenciales(self):
        entrada = Mock()
        entrada.isatty.return_value = False
        with patch("sys.frozen", True, create=True), patch("sys.stdin", entrada), patch("sys.argv", ["principal"]), patch("keyring.get_password", return_value=None), patch.object(modulo_ejecucion, "configurar_credenciales") as configurar:
            with self.assertRaisesRegex(RuntimeError, "Faltan credenciales"):
                principal.main()
        configurar.assert_not_called()

    def test_ejecutable_sin_pausa_no_pide_credenciales(self):
        entrada = Mock()
        entrada.isatty.return_value = True
        with patch("sys.frozen", True, create=True), patch("sys.stdin", entrada), patch("sys.argv", ["principal", "--sin-pausa"]), patch("keyring.get_password", return_value=None), patch.object(modulo_ejecucion, "configurar_credenciales") as configurar:
            with self.assertRaisesRegex(RuntimeError, "Faltan credenciales"):
                principal.main()
        configurar.assert_not_called()

    def test_acceso_reintenta_fallos_transitorios_sin_reintentar_errores_permanentes(self):
        accion = Mock(side_effect=[PlaywrightTimeoutError("Page.goto: timeout"), "correcto"])
        recuperar = Mock()
        with patch.object(modulo_ejecucion.time, "sleep") as esperar:
            self.assertEqual(ariba.reintentar_navegacion(accion, "Acceso", recuperar), "correcto")
        esperar.assert_called_once_with(30)
        recuperar.assert_called_once()
        with patch.object(modulo_ejecucion.time, "sleep") as esperar:
            with self.assertRaises(ValueError):
                ariba.reintentar_navegacion(Mock(side_effect=ValueError("datos invalidos")), "Acceso")
        esperar.assert_not_called()

    def test_pagina_cerrada_solicita_reabrir_el_navegador(self):
        pagina = Mock()
        pagina.is_closed.return_value = True
        evento = Mock(identificador="Doc1")
        with patch("src.ariba.descargar_evento", side_effect=PlaywrightTimeoutError("Page.goto: cerrado")):
            with self.assertRaises(ariba.NavegadorCerradoError):
                ariba.descargar_con_reintentos(pagina, evento, Path("."), Mock())

    def test_cierre_de_navegador_reabre_y_reanuda_el_proceso(self):
        primera_pagina = Mock()
        primera_pagina.is_closed.return_value = True
        segunda_pagina = Mock()
        segundo_contexto = Mock(pages=[segunda_pagina])
        primer_contexto = Mock(pages=[primera_pagina])
        playwright = Mock()
        playwright.chromium.launch_persistent_context.side_effect = [primer_contexto, segundo_contexto]
        with tempfile.TemporaryDirectory() as carpeta, patch.object(principal, "BASE", Path(carpeta)), patch.object(modulo_ejecucion, "BASE", Path(carpeta)), patch.object(modulo_ejecucion, "bloqueo_ejecucion", return_value=nullcontext()), patch("playwright.sync_api.sync_playwright", return_value=nullcontext(playwright)), patch("keyring.get_password", return_value="valor-ficticio"), patch("sys.argv", ["principal"]), patch.object(modulo_ejecucion.time, "sleep") as esperar, patch.object(modulo_ejecucion, "iniciar_sesion", side_effect=[PlaywrightTimeoutError("Page.goto: cerrado"), None]), patch.object(modulo_ejecucion, "procesar_ejecucion") as procesar:
            self.assertEqual(principal.main(), 0)
        self.assertEqual(playwright.chromium.launch_persistent_context.call_count, 2)
        primer_contexto.close.assert_called_once()
        segundo_contexto.close.assert_called_once()
        esperar.assert_called_once_with(30)
        self.assertIs(procesar.call_args.args[0], segunda_pagina)

    def test_evento_transitorio_se_reintenta_en_otra_ronda_sin_omitirlo(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            reportes = raiz / "reportes_diarios"
            reportes.mkdir()
            libro = Workbook()
            libro.active.append(["Título", "ID"])
            libro.active.append(["Anterior", "Doc0"])
            libro.save(reportes / "2026-11-10_Licitaciones_Codelco_(Ariba).xlsx")
            contenido = '''<div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div><table class="tableBody"><tr><td><a href="https://s3.ariba.com/1">6001</a></td><td>Doc1</td><td></td><td>Abierto</td><td>RFP</td></tr><tr><td><a href="https://s3.ariba.com/2">6002</a></td><td>Doc2</td><td></td><td>Abierto</td><td>RFP</td></tr></table>'''
            orden = []
            def descargar(pagina, evento, destino):
                orden.append(evento.identificador)
                if evento.identificador == "Doc1" and orden.count("Doc1") <= 3:
                    raise PlaywrightTimeoutError("Locator.wait_for: timeout")
                documento = destino / f"{evento.identificador}.doc"
                documento.write_text("<html></html>", encoding="utf-8")
                return documento, None
            def esperar(segundos):
                self.assertEqual(segundos, 30)
                estado = json.loads((raiz / ".datos" / "ejecuciones" / "2026-11-11" / "estado.json").read_text(encoding="utf-8"))
                self.assertEqual(estado["eventos"]["Doc1"]["estado"], "pendiente")
                self.assertEqual(estado["eventos"]["Doc2"]["estado"], "descargado")
                self.assertFalse((raiz / "resumen_licitaciones").exists())
            with patch.object(principal, "BASE", raiz), patch("src.ejecucion.BASE", raiz), patch("src.ejecucion.exportar_eventos", side_effect=lambda pagina, destino: destino.write_text(contenido, encoding="utf-8")), patch("src.ariba.descargar_evento", side_effect=descargar), patch.object(modulo_ejecucion.time, "sleep", side_effect=esperar):
                modulo_ejecucion.procesar_ejecucion(None, raiz, date(2026, 11, 11), Mock())
            self.assertEqual(orden, ["Doc1", "Doc1", "Doc1", "Doc2", "Doc1"])
            estado = json.loads((raiz / ".datos" / "ejecuciones" / "2026-11-11" / "estado.json").read_text(encoding="utf-8"))
            self.assertTrue(estado["completado"])
            self.assertTrue(all(registro["estado"] == "descargado" for registro in estado["eventos"].values()))

    def test_exportacion_incompleta_se_reintenta_sin_guardar_archivo_parcial(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            reportes = raiz / "reportes_diarios"
            reportes.mkdir()
            libro = Workbook()
            libro.active.append(["Título", "ID"])
            libro.active.append(["Anterior", "Doc0"])
            libro.save(reportes / "LIC ARIBA 10-11.xlsx")
            contenido = '''<html><div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div><table class="tableBody"><tr><td><a href="https://s3.ariba.com/1">6001</a></td><td>Doc1</td><td></td><td>Abierto</td><td>RFP</td></tr></table></html>'''
            intentos = []
            def exportar(pagina, destino):
                intentos.append(destino.exists())
                if len(intentos) <= 3:
                    destino.write_text("parcial", encoding="utf-8")
                    raise PlaywrightTimeoutError("timeout")
                destino.write_text(contenido, encoding="utf-8")
            recuperar = Mock()
            with patch.object(principal, "BASE", raiz), patch("src.ejecucion.BASE", raiz), patch("src.ejecucion.exportar_eventos", side_effect=exportar), patch.object(modulo_ejecucion.time, "sleep") as esperar:
                estado, ejecucion, _ = modulo_ejecucion.preparar_ejecucion(None, raiz, date(2026, 11, 11), recuperar)
            self.assertEqual(intentos, [False, False, False, False])
            self.assertEqual(recuperar.call_count, 3)
            esperar.assert_called_once_with(30)
            self.assertEqual(estado["eventos"]["Doc1"]["estado"], "pendiente")
            self.assertEqual((ejecucion / "exportacion.xls").read_text(encoding="utf-8"), contenido)

    def test_descarga_reintenta_despues_de_timeout_y_recupera_sesion(self):
        evento = Mock(identificador="Doc1")
        recuperar = Mock()
        with patch("src.ariba.descargar_evento", side_effect=[PlaywrightTimeoutError("timeout"), (Path("resultado.doc"), None)]) as descarga:
            resultado = ariba.descargar_con_reintentos(None, evento, Path("."), recuperar)
        self.assertEqual(resultado, (Path("resultado.doc"), None))
        self.assertEqual(descarga.call_count, 2)
        recuperar.assert_called_once()

    def test_descarga_agota_reintentos_y_conserva_error(self):
        evento = Mock(identificador="Doc1")
        recuperar = Mock()
        with patch("src.ariba.descargar_evento", side_effect=PlaywrightTimeoutError("timeout")) as descarga:
            with self.assertRaisesRegex(RuntimeError, "Doc1.*pendiente"):
                ariba.descargar_con_reintentos(None, evento, Path("."), recuperar)
        self.assertEqual(descarga.call_count, 3)
        self.assertEqual(recuperar.call_count, 2)

    def test_limpieza_solo_elimina_carpetas_con_fecha_anterior_a_siete_dias(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ejecuciones = Path(carpeta)
            antigua = ejecuciones / "2026-11-03"
            limite = ejecuciones / "2026-11-04"
            reciente = ejecuciones / "2026-11-11"
            otra = ejecuciones / "notas"
            for directorio in (antigua, limite, reciente, otra):
                directorio.mkdir()
                (directorio / "documento.doc").write_text("ejemplo")
            modulo_ejecucion.limpiar_ejecuciones(ejecuciones, date(2026, 11, 11))
            self.assertFalse(antigua.exists())
            self.assertTrue(limite.exists())
            self.assertTrue(reciente.exists())
            self.assertTrue(otra.exists())

    def test_no_permite_dos_ejecuciones_simultaneas(self):
        with tempfile.TemporaryDirectory() as carpeta, patch.object(principal, "BASE", Path(carpeta)), patch.object(modulo_ejecucion, "BASE", Path(carpeta)):
            with archivos.bloqueo_ejecucion():
                with self.assertRaisesRegex(RuntimeError, "otra ejecución"):
                    with archivos.bloqueo_ejecucion():
                        pass
            with archivos.bloqueo_ejecucion():
                pass

    def test_interrupcion_reanudacion_y_exclusion(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            codigo = raiz / "aplicacion"
            codigo.mkdir()
            reportes = raiz / "reportes_diarios"
            reportes.mkdir()
            anterior = reportes / "2026-11-10_Licitaciones_Codelco_(Ariba).xlsx"
            libro = Workbook()
            libro.active.append(["Título", "ID"])
            libro.active.append(["Existente", "Doc0"])
            libro.save(anterior)
            fecha = date(2026, 11, 11)
            exportacion = """<html><div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div><table class="tableBody">
                <tr><td><a href="https://s3.ariba.com/0">Existente</a></td><td>Doc0</td><td></td><td>Abierto</td><td>RFP</td></tr>
                <tr><td><a href="https://s3.ariba.com/1">6001</a></td><td>Doc1</td><td>12/11/2026 15:47</td><td>Abierto</td><td>RFP</td></tr>
                <tr><td><a href="https://s3.ariba.com/2">SR123</a></td><td>Doc2</td><td></td><td>Abierto</td><td>RFP</td></tr>
                <tr><td><a href="https://s3.ariba.com/3">ImcsPot_T2_02515</a></td><td>Doc3</td><td></td><td>Abierto</td><td>RFP</td></tr>
            </table></html>"""
            documento = """<html><table><tr><td class="titleFieldValue">7 Articulo</td><td></td></tr>
                <tr><td>Cantidad</td><td>2 Unidades</td></tr></table></html>"""

            def exportar(pagina, destino):
                destino.write_text(exportacion, encoding="utf-8")

            def descargar(pagina, evento, destino):
                if evento.identificador == "Doc2":
                    raise RuntimeError("fallo simulado")
                ruta = destino / "12-11 Doc1 - 6001.doc"
                ruta.write_text(documento, encoding="utf-8")
                return ruta, None

            with patch.object(principal, "BASE", codigo), patch("src.ejecucion.BASE", codigo), patch("src.ejecucion.exportar_eventos", side_effect=exportar) as exportar_mock, patch("src.ariba.descargar_evento", side_effect=descargar):
                with self.assertRaisesRegex(RuntimeError, "fallo simulado"):
                    modulo_ejecucion.procesar_ejecucion(None, raiz, fecha)
                self.assertFalse((reportes / "2026-11-11_Licitaciones_Codelco_(Ariba).xlsx").exists())
                self.assertEqual(exportar_mock.call_count, 1)

            def excluir(pagina, evento, destino):
                self.assertEqual(evento.identificador, "Doc2")
                return None, "Doc2: sin fecha de vencimiento; excluida"

            with patch.object(principal, "BASE", codigo), patch("src.ejecucion.BASE", codigo), patch("src.ejecucion.exportar_eventos", side_effect=AssertionError("no se debe exportar de nuevo")), patch("src.ariba.descargar_evento", side_effect=excluir):
                modulo_ejecucion.procesar_ejecucion(None, raiz, fecha)
                modulo_ejecucion.procesar_ejecucion(None, raiz, fecha)

            reporte = reportes / "2026-11-11_Licitaciones_Codelco_(Ariba).xlsx"
            self.assertTrue(reporte.exists())
            libro_reporte = load_workbook(reporte, read_only=True, data_only=True)
            self.assertIn("Doc3", [fila[0] for fila in libro_reporte.active.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True)])
            libro_reporte.close()
            self.assertTrue((codigo / ".datos" / "ejecuciones" / "2026-11-11" / "12-11 Doc1 - 6001.doc").exists())
            zip_documentos = codigo / ".datos" / "ejecuciones" / "2026-11-11" / "2026-11-11_Documentos_Licitaciones_Codelco.zip"
            with zipfile.ZipFile(zip_documentos) as archivo:
                self.assertEqual(archivo.namelist(), ["12-11 Doc1 - 6001.doc"])
            self.assertFalse((raiz / "Descargas").exists())
            resumen = raiz / "resumen_licitaciones" / "2026-11-11_Resumen_Licitaciones_Codelco.xlsx"
            generado = load_workbook(resumen, data_only=True)
            self.assertEqual(generado.active.max_row, 2)
            self.assertEqual(generado.active["F2"].value.strftime("%d/%m/%Y %H:%M"), "12/11/2026 15:47")
            self.assertEqual(generado.active["A2"].hyperlink.target, "https://s3.ariba.com/1")
            generado.close()
            estado = json.loads((codigo / ".datos" / "ejecuciones" / "2026-11-11" / "estado.json").read_text(encoding="utf-8"))
            self.assertEqual(estado["eventos"]["Doc2"]["estado"], "excluido")
            self.assertEqual(estado["eventos"]["Doc3"]["estado"], "excluido")
            self.assertIn("IMCSPOT", estado["eventos"]["Doc3"]["aviso"])
            self.assertTrue(estado["completado"])

    def test_imcspot_descargado_previamente_no_reutiliza_consolidado(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            ejecucion = raiz / "ejecucion"
            ejecucion.mkdir()
            anterior = raiz / "anterior.xlsx"
            libro = Workbook()
            libro.active.append(["Título", "ID"])
            libro.active.append(["Existente", "Doc0"])
            libro.save(anterior)
            exportacion = ejecucion / "exportacion.xls"
            exportacion.write_text('''<html><div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div><table class="tableBody">
                <tr><td><a href="https://s3.ariba.com/0">Existente</a></td><td>Doc0</td><td></td><td>Abierto</td><td>RFP</td></tr>
                <tr><td><a href="https://s3.ariba.com/3">IMCSPOT_T2_02515</a></td><td>Doc3</td><td></td><td>Abierto</td><td>RFP</td></tr>
            </table></html>''', encoding="utf-8")
            estado = {
                "anterior": str(anterior),
                "hash_anterior": archivos.hash_archivo(anterior),
                "hash_exportacion": archivos.hash_archivo(exportacion),
                "eventos": {"Doc3": {"estado": "descargado"}},
                "hash_resumen": "resumen_anterior",
                "completado": False,
            }
            estado_path = ejecucion / "estado.json"
            with patch.object(modulo_ejecucion, "preparar_ejecucion", return_value=(estado, ejecucion, estado_path)), patch("src.ariba.descargar_evento") as descargar:
                with self.assertRaisesRegex(ValueError, "puede contener eventos IMCSPOT"):
                    modulo_ejecucion.procesar_ejecucion(None, raiz, date(2026, 11, 11))
                descargar.assert_not_called()
            self.assertEqual(estado["eventos"]["Doc3"]["estado"], "descargado")


if __name__ == "__main__":
    unittest.main()