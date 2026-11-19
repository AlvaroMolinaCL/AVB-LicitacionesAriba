import json
import os
import runpy
import subprocess
import sys
import tempfile
import pytest
from contextlib import nullcontext
from datetime import date
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

import principal
from src import ariba, archivos, consolidado, ejecucion as modulo_ejecucion


class TestFlujo:
    def test_segunda_consulta_no_selecciona_codelco_antes_de_exportar(self):
        pagina = Mock()
        contexto = Mock(pages=[pagina])
        playwright = Mock()
        playwright.chromium.launch_persistent_context.return_value = contexto
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            fecha = date.today()
            estado_path = (
                raiz / ".datos" / "ejecuciones" / fecha.isoformat() / "estado.json"
            )
            estado_path.parent.mkdir(parents=True)
            estado_path.write_text(
                json.dumps(
                    {"raiz": str(raiz.resolve()), "completado": True, "eventos": {}}
                ),
                encoding="utf-8",
            )
            with patch.object(
                modulo_ejecucion, "bloqueo_ejecucion", return_value=nullcontext()
            ), patch(
                "playwright.sync_api.sync_playwright",
                return_value=nullcontext(playwright),
            ), patch(
                "keyring.get_password", return_value="valor-ficticio"
            ), patch.object(
                modulo_ejecucion, "iniciar_sesion"
            ), patch.object(
                modulo_ejecucion, "seleccionar_codelco"
            ) as seleccionar, patch.object(
                modulo_ejecucion, "procesar_ejecucion"
            ) as procesar:
                assert (
                    modulo_ejecucion.ejecutar_flujo_principal(
                        raiz, "servicio", raiz, fecha, False, False, True
                    )
                ) == (0)
            seleccionar.assert_not_called()
            procesar.assert_called_once()

    def test_actualizacion_no_sobrescribe_reporte_modificado(self):
        from src.actualizaciones import publicar_actualizacion

        with tempfile.TemporaryDirectory() as carpeta:
            origen = Path(carpeta) / "nuevo.xlsx"
            destino = Path(carpeta) / "reporte.xlsx"
            origen.write_bytes(b"nuevo")
            destino.write_bytes(b"original")
            hash_previo = archivos.hash_archivo(destino)
            destino.write_bytes(b"modificado")
            with pytest.raises(ValueError, match="cambió"):
                publicar_actualizacion(origen, destino, hash_previo)
            assert (destino.read_bytes()) == (b"modificado")

    def test_sin_novedades_se_vuelve_a_consultar_en_el_mismo_dia(self):
        cabecera = '<div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div>'
        fila = '<tr><td><a href="https://s3.ariba.com/1">6001</a></td><td>Doc1</td><td></td><td>Abierto</td><td>RFP</td></tr>'
        nueva = '<tr><td><a href="https://s3.ariba.com/2">6002</a></td><td>Doc2</td><td></td><td>Abierto</td><td>RFP</td></tr>'
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            reportes = raiz / "reportes_diarios"
            reportes.mkdir()
            libro = Workbook()
            libro.active.append(["Título", "ID"])
            libro.active.append(["Anterior", "Doc1"])
            libro.save(reportes / "2026-11-10_Licitaciones_Codelco_(Ariba).xlsx")
            libro.close()
            exportaciones = iter(
                [
                    cabecera + '<table class="tableBody">' + fila + "</table>",
                    cabecera + '<table class="tableBody">' + fila + nueva + "</table>",
                ]
            )

            def exportar(pagina, destino):
                destino.write_text(next(exportaciones), encoding="utf-8")

            def descargar(pagina, evento, destino, recuperar):
                if descargar.intentos == 0:
                    descargar.intentos += 1
                    raise RuntimeError("interrupción simulada")
                documento = destino / "Doc2.doc"
                documento.write_text("<html></html>", encoding="utf-8")
                return documento, None

            descargar.intentos = 0
            with patch.object(modulo_ejecucion, "BASE", raiz), patch.object(
                modulo_ejecucion, "exportar_eventos", side_effect=exportar
            ) as exportacion, patch.object(
                modulo_ejecucion, "descargar_con_reintentos", side_effect=descargar
            ) as descarga:
                modulo_ejecucion.procesar_ejecucion(None, raiz, date(2026, 11, 11))
                with pytest.raises(RuntimeError, match="interrupción simulada"):
                    modulo_ejecucion.procesar_ejecucion(None, raiz, date(2026, 11, 11))
                assert (
                    modulo_ejecucion.leer_ids_anteriores(
                        reportes / "2026-11-11_Licitaciones_Codelco_(Ariba).xlsx"
                    )
                ) == (["Doc1"])
                modulo_ejecucion.procesar_ejecucion(None, raiz, date(2026, 11, 11))
            assert (exportacion.call_count) == (2)
            assert (descarga.call_args.args[1].identificador) == ("Doc2")
            assert (
                modulo_ejecucion.leer_ids_anteriores(
                    reportes / "2026-11-11_Licitaciones_Codelco_(Ariba).xlsx"
                )
            ) == (["Doc1", "Doc2"])
            assert (
                raiz
                / "resumen_licitaciones"
                / "2026-11-11_Resumen_Licitaciones_Codelco.xlsx"
            ).exists()

    def test_primera_ejecucion_se_reanuda_y_conserva_exclusiones(self):
        contenido = """<div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div><table class="tableBody"><tr><td><a href="https://s3.ariba.com/1">6001</a></td><td>Doc1</td><td></td><td>Abierto</td><td>RFP</td></tr><tr><td><a href="https://s3.ariba.com/2">6002</a></td><td>Doc2</td><td></td><td>Abierto</td><td>RFP</td></tr><tr><td><a href="https://s3.ariba.com/3">IMCSPOT prueba</a></td><td>Doc3</td><td></td><td>Abierto</td><td>RFP</td></tr></table>"""
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            orden = []

            def descargar(pagina, evento, destino, recuperar):
                orden.append(evento.identificador)
                if evento.identificador == "Doc2":
                    if orden.count("Doc2") == 1:
                        raise RuntimeError("interrupción simulada")
                    return None, "Doc2: sin fecha de vencimiento; excluida"
                documento = destino / "Doc1.doc"
                documento.write_text("<html></html>", encoding="utf-8")
                return documento, None

            with patch.object(modulo_ejecucion, "BASE", raiz), patch.object(
                modulo_ejecucion,
                "exportar_eventos",
                side_effect=lambda pagina, destino: destino.write_text(
                    contenido, encoding="utf-8"
                ),
            ) as exportar, patch.object(
                modulo_ejecucion, "descargar_con_reintentos", side_effect=descargar
            ):
                with pytest.raises(RuntimeError, match="interrupción simulada"):
                    modulo_ejecucion.procesar_ejecucion(None, raiz, date(2026, 11, 11))
                modulo_ejecucion.procesar_ejecucion(None, raiz, date(2026, 11, 11))

            estado_path = raiz / ".datos" / "ejecuciones" / "2026-11-11" / "estado.json"
            estado = json.loads(estado_path.read_text(encoding="utf-8"))
            assert (estado["anterior"]) is None
            assert (estado["hash_anterior"]) is None
            assert estado["completado"]
            assert (orden) == (["Doc1", "Doc2", "Doc2"])
            assert (estado["eventos"]["Doc1"]["estado"]) == ("descargado")
            assert (estado["eventos"]["Doc2"]["estado"]) == ("excluido")
            assert (estado["eventos"]["Doc3"]["estado"]) == ("excluido")
            exportar.assert_called_once()
            reporte = (
                raiz
                / "reportes_diarios"
                / "2026-11-11_Licitaciones_Codelco_(Ariba).xlsx"
            )
            assert (modulo_ejecucion.leer_ids_anteriores(reporte)) == (
                ["Doc1", "Doc2", "Doc3"]
            )
            assert (
                raiz
                / "resumen_licitaciones"
                / "2026-11-11_Resumen_Licitaciones_Codelco.xlsx"
            ).exists()

    def test_validacion_reintenta_exportacion_y_elimina_descarga_parcial(self):
        pagina = Mock()
        pagina.is_closed.return_value = False
        contexto = Mock(pages=[pagina])
        playwright = Mock()
        playwright.chromium.launch_persistent_context.return_value = contexto
        intentos = []

        def exportar(pagina, destino):
            intentos.append(destino.exists())
            destino.write_text("datos", encoding="utf-8")
            if len(intentos) == 1:
                raise PlaywrightTimeoutError("Locator.click: timeout")

        with tempfile.TemporaryDirectory() as carpeta, patch.object(
            modulo_ejecucion, "bloqueo_ejecucion", return_value=nullcontext()
        ), patch(
            "playwright.sync_api.sync_playwright", return_value=nullcontext(playwright)
        ), patch(
            "keyring.get_password", return_value="valor-ficticio"
        ), patch.object(
            modulo_ejecucion, "iniciar_sesion"
        ) as login, patch.object(
            modulo_ejecucion, "exportar_eventos", side_effect=exportar
        ), patch.object(
            modulo_ejecucion, "leer_exportacion", return_value=[]
        ), patch.object(
            modulo_ejecucion.time, "sleep"
        ) as esperar:
            resultado = modulo_ejecucion.ejecutar_flujo_principal(
                Path(carpeta), "servicio", Path(carpeta), None, False, True, True
            )
        assert (resultado) == (0)
        assert (intentos) == ([False, False])
        assert (login.call_count) == (2)
        esperar.assert_called_once_with(30)

    @pytest.mark.skipif(os.name != "nt", reason="Requiere Windows")
    def test_salida_redirigida_usa_utf8_aunque_el_entorno_indique_cp1252(self):
        ruta = Path(__file__).resolve().parent.parent / "principal.py"
        entorno = os.environ.copy()
        entorno["PYTHONIOENCODING"] = "cp1252"
        resultado = subprocess.run(
            [sys.executable, str(ruta), "--help"],
            capture_output=True,
            env=entorno,
            check=True,
        )
        salida = resultado.stdout.decode("utf-8")
        assert ("Automatización") in (salida)
        assert ("contraseña de aplicación") in (salida)

    def test_configurar_consola_windows_usa_utf8_en_ambos_flujos(self):
        salida = Mock()
        errores = Mock()
        with patch("sys.platform", "win32"), patch("sys.stdout", salida), patch(
            "sys.stderr", errores
        ), patch.object(principal.os, "system") as sistema:
            principal.configurar_consola()
        sistema.assert_called_once_with("chcp 65001 > nul")
        salida.reconfigure.assert_called_once_with(
            encoding="utf-8", errors="backslashreplace"
        )
        errores.reconfigure.assert_called_once_with(
            encoding="utf-8", errors="backslashreplace"
        )

    def test_configurar_consola_no_modifica_otras_plataformas(self):
        with patch("sys.platform", "linux"), patch.object(
            principal.os, "system"
        ) as sistema:
            principal.configurar_consola()
        sistema.assert_not_called()

    def test_resumen_antiguo_no_se_publica_como_actual(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "resumen.xlsx"
            libro = Workbook()
            libro.active.append(
                [
                    "Archivo",
                    "Región",
                    "Ítem Principal",
                    "N° Material",
                    "Cant.",
                    "Entrega",
                    "Texto largo",
                ]
            )
            libro.save(ruta)
            libro.close()
            with pytest.raises(ValueError, match="formato anterior"):
                consolidado.validar_formato_resumen(ruta)

    def test_base_es_la_raiz_del_proyecto(self):
        assert (principal.BASE) == (Path(__file__).resolve().parent.parent)

    def test_base_del_ejecutable_es_su_carpeta(self):
        ejecutable = Path(tempfile.gettempdir()) / "aplicacion" / "respaldo.exe"
        with patch("sys.frozen", True, create=True), patch(
            "sys.executable", str(ejecutable)
        ):
            modulo = runpy.run_path(
                str(Path(__file__).resolve().parent.parent / "principal.py")
            )
        assert (modulo["BASE"]) == (ejecutable.resolve().parent)

    def test_pausa_solo_en_ejecutable_interactivo(self):
        entrada = Mock()
        entrada.isatty.return_value = True
        with patch("sys.frozen", True, create=True), patch("sys.stdin", entrada), patch(
            "sys.argv", ["respaldo.exe"]
        ), patch("builtins.input") as leer:
            principal.pausar_al_finalizar()
            leer.assert_called_once()
        with patch("sys.frozen", True, create=True), patch("sys.stdin", entrada), patch(
            "sys.argv", ["respaldo.exe", "--sin-pausa"]
        ), patch("builtins.input") as leer:
            principal.pausar_al_finalizar()
            leer.assert_not_called()

    def test_ejecutable_pide_credenciales_y_usa_navegador_oculto_local(self):
        pagina = Mock()
        contexto = Mock(pages=[pagina])
        playwright = Mock()
        playwright.chromium.launch_persistent_context.return_value = contexto
        entrada = Mock()
        entrada.isatty.return_value = True
        with tempfile.TemporaryDirectory() as carpeta, patch.object(
            principal, "BASE", Path(carpeta)
        ), patch.object(modulo_ejecucion, "BASE", Path(carpeta)), patch.object(
            modulo_ejecucion, "bloqueo_ejecucion", return_value=nullcontext()
        ), patch(
            "playwright.sync_api.sync_playwright", return_value=nullcontext(playwright)
        ), patch(
            "keyring.get_password",
            side_effect=[None, None, "usuario-ficticio", "clave-ficticia"],
        ), patch(
            "sys.frozen", True, create=True
        ), patch(
            "sys.stdin", entrada
        ), patch(
            "sys.argv", ["principal"]
        ), patch.dict(
            os.environ, {"PLAYWRIGHT_BROWSERS_PATH": "ruta-externa"}
        ), patch.object(
            modulo_ejecucion, "configurar_credenciales"
        ) as configurar, patch.object(
            modulo_ejecucion, "configurar_correo_interactivo_si_falta"
        ), patch.object(
            modulo_ejecucion, "iniciar_sesion"
        ) as login, patch.object(
            modulo_ejecucion, "procesar_ejecucion"
        ):
            assert (principal.main()) == (0)
            assert (os.environ["PLAYWRIGHT_BROWSERS_PATH"]) == (
                str(Path(carpeta) / ".datos" / "navegadores")
            )
        configurar.assert_called_once()
        login.assert_called_once_with(pagina, "usuario-ficticio", "clave-ficticia")
        assert playwright.chromium.launch_persistent_context.call_args.kwargs[
            "headless"
        ]

    def test_ejecutable_sin_consola_no_pide_credenciales(self):
        entrada = Mock()
        entrada.isatty.return_value = False
        with patch("sys.frozen", True, create=True), patch("sys.stdin", entrada), patch(
            "sys.argv", ["principal"]
        ), patch("keyring.get_password", return_value=None), patch.object(
            modulo_ejecucion, "configurar_credenciales"
        ) as configurar:
            with pytest.raises(RuntimeError, match="Faltan credenciales"):
                principal.main()
        configurar.assert_not_called()

    def test_ejecutable_sin_pausa_no_pide_credenciales(self):
        entrada = Mock()
        entrada.isatty.return_value = True
        with patch("sys.frozen", True, create=True), patch("sys.stdin", entrada), patch(
            "sys.argv", ["principal", "--sin-pausa"]
        ), patch("keyring.get_password", return_value=None), patch.object(
            modulo_ejecucion, "configurar_credenciales"
        ) as configurar:
            with pytest.raises(RuntimeError, match="Faltan credenciales"):
                principal.main()
        configurar.assert_not_called()

    def test_acceso_reintenta_fallos_transitorios_sin_reintentar_errores_permanentes(
        self,
    ):
        accion = Mock(
            side_effect=[PlaywrightTimeoutError("Page.goto: timeout"), "correcto"]
        )
        recuperar = Mock()
        with patch.object(modulo_ejecucion.time, "sleep") as esperar:
            assert (ariba.reintentar_navegacion(accion, "Acceso", recuperar)) == (
                "correcto"
            )
        esperar.assert_called_once_with(30)
        recuperar.assert_called_once()
        with patch.object(modulo_ejecucion.time, "sleep") as esperar:
            with pytest.raises(ValueError):
                ariba.reintentar_navegacion(
                    Mock(side_effect=ValueError("datos invalidos")), "Acceso"
                )
        esperar.assert_not_called()

    def test_pagina_cerrada_solicita_reabrir_el_navegador(self):
        pagina = Mock()
        pagina.is_closed.return_value = True
        evento = Mock(identificador="Doc1")
        with patch(
            "src.ariba.descargar_evento",
            side_effect=PlaywrightTimeoutError("Page.goto: cerrado"),
        ):
            with pytest.raises(ariba.NavegadorCerradoError):
                ariba.descargar_con_reintentos(pagina, evento, Path("."), Mock())

    def test_cierre_de_navegador_reabre_y_reanuda_el_proceso(self):
        primera_pagina = Mock()
        primera_pagina.is_closed.return_value = True
        segunda_pagina = Mock()
        segundo_contexto = Mock(pages=[segunda_pagina])
        primer_contexto = Mock(pages=[primera_pagina])
        playwright = Mock()
        playwright.chromium.launch_persistent_context.side_effect = [
            primer_contexto,
            segundo_contexto,
        ]
        with tempfile.TemporaryDirectory() as carpeta, patch.object(
            principal, "BASE", Path(carpeta)
        ), patch.object(modulo_ejecucion, "BASE", Path(carpeta)), patch.object(
            modulo_ejecucion, "bloqueo_ejecucion", return_value=nullcontext()
        ), patch(
            "playwright.sync_api.sync_playwright", return_value=nullcontext(playwright)
        ), patch(
            "keyring.get_password", return_value="valor-ficticio"
        ), patch(
            "sys.argv", ["principal"]
        ), patch.object(
            modulo_ejecucion.time, "sleep"
        ) as esperar, patch.object(
            modulo_ejecucion,
            "iniciar_sesion",
            side_effect=[PlaywrightTimeoutError("Page.goto: cerrado"), None],
        ), patch.object(
            modulo_ejecucion, "procesar_ejecucion"
        ) as procesar:
            assert (principal.main()) == (0)
        assert (playwright.chromium.launch_persistent_context.call_count) == (2)
        primer_contexto.close.assert_called_once()
        segundo_contexto.close.assert_called_once()
        esperar.assert_called_once_with(30)
        assert (procesar.call_args.args[0]) is (segunda_pagina)

    def test_evento_transitorio_se_reintenta_en_otra_ronda_sin_omitirlo(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            reportes = raiz / "reportes_diarios"
            reportes.mkdir()
            libro = Workbook()
            libro.active.append(["Título", "ID"])
            libro.active.append(["Anterior", "Doc0"])
            libro.save(reportes / "2026-11-10_Licitaciones_Codelco_(Ariba).xlsx")
            contenido = """<div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div><table class="tableBody"><tr><td><a href="https://s3.ariba.com/1">6001</a></td><td>Doc1</td><td></td><td>Abierto</td><td>RFP</td></tr><tr><td><a href="https://s3.ariba.com/2">6002</a></td><td>Doc2</td><td></td><td>Abierto</td><td>RFP</td></tr></table>"""
            orden = []

            def descargar(pagina, evento, destino):
                orden.append(evento.identificador)
                if evento.identificador == "Doc1" and orden.count("Doc1") <= 3:
                    raise PlaywrightTimeoutError("Locator.wait_for: timeout")
                documento = destino / f"{evento.identificador}.doc"
                documento.write_text("<html></html>", encoding="utf-8")
                return documento, None

            def esperar(segundos):
                assert (segundos) == (30)
                estado = json.loads(
                    (
                        raiz / ".datos" / "ejecuciones" / "2026-11-11" / "estado.json"
                    ).read_text(encoding="utf-8")
                )
                assert (estado["eventos"]["Doc1"]["estado"]) == ("pendiente")
                assert (estado["eventos"]["Doc2"]["estado"]) == ("descargado")
                assert not ((raiz / "resumen_licitaciones").exists())

            with patch.object(principal, "BASE", raiz), patch(
                "src.ejecucion.BASE", raiz
            ), patch(
                "src.ejecucion.exportar_eventos",
                side_effect=lambda pagina, destino: destino.write_text(
                    contenido, encoding="utf-8"
                ),
            ), patch(
                "src.ariba.descargar_evento", side_effect=descargar
            ), patch.object(
                modulo_ejecucion.time, "sleep", side_effect=esperar
            ):
                modulo_ejecucion.procesar_ejecucion(
                    None, raiz, date(2026, 11, 11), Mock()
                )
            assert (orden) == (["Doc1", "Doc1", "Doc1", "Doc2", "Doc1"])
            estado = json.loads(
                (
                    raiz / ".datos" / "ejecuciones" / "2026-11-11" / "estado.json"
                ).read_text(encoding="utf-8")
            )
            assert estado["completado"]
            assert all(
                registro["estado"] == "descargado"
                for registro in estado["eventos"].values()
            )

    def test_exportacion_incompleta_se_reintenta_sin_guardar_archivo_parcial(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            reportes = raiz / "reportes_diarios"
            reportes.mkdir()
            libro = Workbook()
            libro.active.append(["Título", "ID"])
            libro.active.append(["Anterior", "Doc0"])
            libro.save(reportes / "LIC ARIBA 10-11.xlsx")
            contenido = """<html><div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div><table class="tableBody"><tr><td><a href="https://s3.ariba.com/1">6001</a></td><td>Doc1</td><td></td><td>Abierto</td><td>RFP</td></tr></table></html>"""
            intentos = []

            def exportar(pagina, destino):
                intentos.append(destino.exists())
                if len(intentos) <= 3:
                    destino.write_text("parcial", encoding="utf-8")
                    raise PlaywrightTimeoutError("timeout")
                destino.write_text(contenido, encoding="utf-8")

            recuperar = Mock()
            with patch.object(principal, "BASE", raiz), patch(
                "src.ejecucion.BASE", raiz
            ), patch(
                "src.ejecucion.exportar_eventos", side_effect=exportar
            ), patch.object(
                modulo_ejecucion.time, "sleep"
            ) as esperar:
                estado, ejecucion, _ = modulo_ejecucion.preparar_ejecucion(
                    None, raiz, date(2026, 11, 11), recuperar
                )
            assert (intentos) == ([False, False, False, False])
            assert (recuperar.call_count) == (3)
            esperar.assert_called_once_with(30)
            assert (estado["eventos"]["Doc1"]["estado"]) == ("pendiente")
            assert ((ejecucion / "exportacion.xls").read_text(encoding="utf-8")) == (
                contenido
            )

    def test_descarga_reintenta_despues_de_timeout_y_recupera_sesion(self):
        evento = Mock(identificador="Doc1")
        recuperar = Mock()
        with patch(
            "src.ariba.descargar_evento",
            side_effect=[
                PlaywrightTimeoutError("timeout"),
                (Path("resultado.doc"), None),
            ],
        ) as descarga:
            resultado = ariba.descargar_con_reintentos(
                None, evento, Path("."), recuperar
            )
        assert (resultado) == ((Path("resultado.doc"), None))
        assert (descarga.call_count) == (2)
        recuperar.assert_called_once()

    def test_descarga_agota_reintentos_y_conserva_error(self):
        evento = Mock(identificador="Doc1")
        recuperar = Mock()
        with patch(
            "src.ariba.descargar_evento", side_effect=PlaywrightTimeoutError("timeout")
        ) as descarga:
            with pytest.raises(RuntimeError, match="Doc1.*pendiente"):
                ariba.descargar_con_reintentos(None, evento, Path("."), recuperar)
        assert (descarga.call_count) == (3)
        assert (recuperar.call_count) == (2)

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
            assert not (antigua.exists())
            assert limite.exists()
            assert reciente.exists()
            assert otra.exists()

    def test_no_permite_dos_ejecuciones_simultaneas(self):
        with tempfile.TemporaryDirectory() as carpeta, patch.object(
            principal, "BASE", Path(carpeta)
        ), patch.object(modulo_ejecucion, "BASE", Path(carpeta)):
            with archivos.bloqueo_ejecucion():
                with pytest.raises(RuntimeError, match="otra ejecución"):
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

            with patch.object(principal, "BASE", codigo), patch(
                "src.ejecucion.BASE", codigo
            ), patch(
                "src.ejecucion.exportar_eventos", side_effect=exportar
            ) as exportar_mock, patch(
                "src.ariba.descargar_evento", side_effect=descargar
            ):
                with pytest.raises(RuntimeError, match="fallo simulado"):
                    modulo_ejecucion.procesar_ejecucion(None, raiz, fecha)
                assert not (
                    (reportes / "2026-11-11_Licitaciones_Codelco_(Ariba).xlsx").exists()
                )
                assert (exportar_mock.call_count) == (1)

            def excluir(pagina, evento, destino):
                assert (evento.identificador) == ("Doc2")
                return None, "Doc2: sin fecha de vencimiento; excluida"

            with patch.object(principal, "BASE", codigo), patch(
                "src.ejecucion.BASE", codigo
            ), patch(
                "src.ejecucion.exportar_eventos",
                side_effect=AssertionError("no se debe exportar de nuevo"),
            ), patch(
                "src.ariba.descargar_evento", side_effect=excluir
            ):
                modulo_ejecucion.procesar_ejecucion(None, raiz, fecha)
                modulo_ejecucion.procesar_ejecucion(None, raiz, fecha)

            reporte = reportes / "2026-11-11_Licitaciones_Codelco_(Ariba).xlsx"
            assert reporte.exists()
            libro_reporte = load_workbook(reporte, read_only=True, data_only=True)
            assert ("Doc3") in (
                [
                    fila[0]
                    for fila in libro_reporte.active.iter_rows(
                        min_row=2, min_col=2, max_col=2, values_only=True
                    )
                ]
            )
            libro_reporte.close()
            assert (
                codigo
                / ".datos"
                / "ejecuciones"
                / "2026-11-11"
                / "12-11 Doc1 - 6001.doc"
            ).exists()
            documento_pdf = (
                codigo
                / ".datos"
                / "ejecuciones"
                / "2026-11-11"
                / "12-11 Doc1 - 6001.pdf"
            )
            assert documento_pdf.read_bytes().startswith(b"%PDF-")
            assert not ((raiz / "Descargas").exists())
            resumen = (
                raiz
                / "resumen_licitaciones"
                / "2026-11-11_Resumen_Licitaciones_Codelco.xlsx"
            )
            generado = load_workbook(resumen, data_only=True)
            assert (generado.active.max_row) == (2)
            assert (generado.active["F2"].value.strftime("%d/%m/%Y %H:%M")) == (
                "12/11/2026 15:47"
            )
            assert (generado.active["A2"].hyperlink.target) == (
                "https://s3.ariba.com/1"
            )
            generado.close()
            estado = json.loads(
                (
                    codigo / ".datos" / "ejecuciones" / "2026-11-11" / "estado.json"
                ).read_text(encoding="utf-8")
            )
            assert (estado["eventos"]["Doc2"]["estado"]) == ("excluido")
            assert (estado["eventos"]["Doc3"]["estado"]) == ("excluido")
            assert ("IMCSPOT") in (estado["eventos"]["Doc3"]["aviso"])
            assert estado["completado"]

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
            exportacion.write_text(
                """<html><div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div><table class="tableBody">
                <tr><td><a href="https://s3.ariba.com/0">Existente</a></td><td>Doc0</td><td></td><td>Abierto</td><td>RFP</td></tr>
                <tr><td><a href="https://s3.ariba.com/3">IMCSPOT_T2_02515</a></td><td>Doc3</td><td></td><td>Abierto</td><td>RFP</td></tr>
            </table></html>""",
                encoding="utf-8",
            )
            estado = {
                "anterior": str(anterior),
                "hash_anterior": archivos.hash_archivo(anterior),
                "hash_exportacion": archivos.hash_archivo(exportacion),
                "eventos": {"Doc3": {"estado": "descargado"}},
                "hash_resumen": "resumen_anterior",
                "completado": False,
            }
            estado_path = ejecucion / "estado.json"
            with patch.object(
                modulo_ejecucion,
                "preparar_ejecucion",
                return_value=(estado, ejecucion, estado_path),
            ), patch("src.ariba.descargar_evento") as descargar:
                with pytest.raises(ValueError, match="puede contener eventos IMCSPOT"):
                    modulo_ejecucion.procesar_ejecucion(None, raiz, date(2026, 11, 11))
                descargar.assert_not_called()
            assert (estado["eventos"]["Doc3"]["estado"]) == ("descargado")
