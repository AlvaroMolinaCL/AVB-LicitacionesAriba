import tempfile
import pytest
from contextlib import nullcontext
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from src.ariba import (
    ServicioTemporalError,
    SesionExpiradaError,
    capturar_impresion,
    descargar_evento,
    es_impresion,
    fecha_vencimiento,
    nombre_documento,
    seleccionar_codelco,
    url_ariba,
)


class TestAriba:
    def test_seleccion_informa_el_paso_que_agota_el_tiempo(self):
        pagina = Mock()
        pagina.get_by_role.return_value.wait_for.side_effect = TimeoutError()
        with patch("builtins.print") as imprimir:
            with pytest.raises(TimeoutError):
                seleccionar_codelco(pagina)
        mensajes = [llamada.args[0] for llamada in imprimir.call_args_list]
        assert ("Ariba: esperando la lista de clientes.") in (mensajes)
        assert ("Ariba: seleccionando Codelco.") not in (mensajes)

    def test_descarga_parcial_no_bloquea_el_siguiente_intento(self):
        pagina = Mock(url="https://s3.ariba.com/evento")
        pagina.goto.return_value = None
        pagina.locator.return_value.count.return_value = 0
        evento = SimpleNamespace(
            identificador="Doc1", titulo="6001", enlace="https://s3.ariba.com/evento"
        )
        contenido = "<!-- class: ariba.sourcing.rfxui.PrintRFXEncode --><html><table><tr><td>Fecha de vencimiento</td><td>24/11/2026</td></tr></table></html>"

        def descargar(pagina, boton, destino):
            destino.write_text(contenido, encoding="utf-8")
            raise TimeoutError("descarga incompleta")

        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            with patch("src.ariba.capturar_impresion", side_effect=descargar):
                with pytest.raises(TimeoutError):
                    descargar_evento(pagina, evento, carpeta)
            assert (list(carpeta.iterdir())) == ([])
            with patch(
                "src.ariba.capturar_impresion",
                side_effect=lambda pagina, boton, destino: destino.write_text(
                    contenido, encoding="utf-8"
                ),
            ):
                documento, aviso = descargar_evento(pagina, evento, carpeta)
            assert (aviso) is None
            assert (documento.name) == ("24-11 Doc1 - 6001.doc")
            with patch(
                "src.ariba.capturar_impresion",
                side_effect=lambda pagina, boton, destino: destino.write_text(
                    contenido, encoding="utf-8"
                ),
            ):
                recuperado, _ = descargar_evento(pagina, evento, carpeta)
            assert (recuperado) == (documento)
            documento.write_text("otro documento", encoding="utf-8")
            with patch(
                "src.ariba.capturar_impresion",
                side_effect=lambda pagina, boton, destino: destino.write_text(
                    contenido, encoding="utf-8"
                ),
            ):
                with pytest.raises(FileExistsError):
                    descargar_evento(pagina, evento, carpeta)
            assert (documento.read_text(encoding="utf-8")) == ("otro documento")

    def test_descarga_identifica_redireccion_de_login_y_error_temporal(self):
        pagina = Mock(url="https://s3.ariba.com/evento")
        evento = SimpleNamespace(
            identificador="Doc1", titulo="6001", enlace="https://s3.ariba.com/evento"
        )
        with tempfile.TemporaryDirectory() as carpeta:
            pagina.goto.return_value = Mock(status=503)
            with pytest.raises(ServicioTemporalError):
                descargar_evento(pagina, evento, Path(carpeta))
            pagina.goto.return_value = Mock(status=200)
            pagina.locator.return_value.count.return_value = 1
            with pytest.raises(SesionExpiradaError):
                descargar_evento(pagina, evento, Path(carpeta))

    def test_impresion_identifica_error_temporal(self):
        respuesta = Mock(status=503, headers={"content-type": "text/html"})
        respuesta.body.return_value = b"servicio no disponible"
        ruta = Mock()
        ruta.fetch.return_value = respuesta
        pagina = Mock()
        pagina.expect_response.return_value = nullcontext()
        boton = Mock()
        boton.get_attribute.return_value = "control"
        boton.click.side_effect = lambda: pagina.route.call_args.args[1](ruta)
        with tempfile.TemporaryDirectory() as carpeta:
            with pytest.raises(ServicioTemporalError):
                capturar_impresion(pagina, boton, Path(carpeta) / "evento.doc")

    def test_fecha_y_nombre_del_documento(self):
        with tempfile.TemporaryDirectory() as carpeta:
            documento = Path(carpeta) / "evento_ficticio.doc"
            documento.write_text(
                """<!-- class: ariba.sourcing.rfxui.PrintRFXEncode -->
                <html><table><tr><td class="fieldLabel">Fecha de vencimiento</td>
                <td class="fieldValue">04/12/2030 15:13</td></tr></table></html>""",
                encoding="utf-8",
            )
            fecha = fecha_vencimiento(documento)
            assert (fecha) == (datetime(2030, 12, 4, 15, 13))
            assert (nombre_documento(fecha, "Doc123", "6001/ARTICULO")) == (
                "04-12 Doc123 - 6001 ARTICULO.doc"
            )

    def test_fecha_ausente_y_fecha_desconocida(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "evento.doc"
            ruta.write_text(
                "<!-- class: ariba.sourcing.rfxui.PrintRFXEncode --><html><table></table></html>",
                encoding="utf-8",
            )
            assert (fecha_vencimiento(ruta)) is None
            ruta.write_text(
                "<!-- class: ariba.sourcing.rfxui.PrintRFXEncode --><html><table><tr><td>Fecha de vencimiento</td><td>fecha desconocida</td></tr></table></html>",
                encoding="utf-8",
            )
            with pytest.raises(ValueError):
                fecha_vencimiento(ruta)

    def test_fecha_de_vencimiento_con_formato_am_pm(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "evento.doc"
            ruta.write_text(
                "<!-- class: ariba.sourcing.rfxui.PrintRFXEncode --><html><table><tr><td>Fecha de vencimiento</td><td>24/11/2026 03:00 PM</td></tr></table></html>",
                encoding="utf-8",
            )
            assert (fecha_vencimiento(ruta)) == (datetime(2026, 11, 24, 15))
            ruta.write_text("<html>Sesion expirada</html>", encoding="utf-8")
            with pytest.raises(ValueError):
                fecha_vencimiento(ruta)

    def test_dominio_de_ariba(self):
        assert url_ariba("https://s3.ariba.com/ejemplo")
        assert not (url_ariba("https://ariba.com.ejemplo.com/"))
        assert not (url_ariba("http://s3.ariba.com/"))

    def test_solicitud_de_impresion_solo_acepta_el_control_y_marco_correspondientes(
        self,
    ):
        assert es_impresion(
            "https://s3.ariba.com/accion?awsn=control&awii=AWRefreshFrame",
            "control",
        )
        assert not (
            es_impresion(
                "https://s3.ariba.com/accion?awsn=otro&awii=AWRefreshFrame", "control"
            )
        )
        assert not (
            es_impresion(
                "https://s3.ariba.com/accion?awsn=control&awii=otro", "control"
            )
        )
        assert not (
            es_impresion(
                "https://ariba.com.ejemplo.com/?awsn=control&awii=AWRefreshFrame",
                "control",
            )
        )

    def test_captura_respuesta_del_marco_interno(self):
        documento = b"<!-- class: ariba.sourcing.rfxui.PrintRFXEncode --><html></html>"
        respuesta = Mock(status=200, headers={"content-type": "application/msword"})
        respuesta.body.return_value = documento
        ruta = Mock()
        ruta.fetch.return_value = respuesta
        pagina = Mock()
        pagina.expect_response.return_value = nullcontext()
        boton = Mock()
        boton.get_attribute.return_value = "control"
        boton.click.side_effect = lambda: pagina.route.call_args.args[1](ruta)
        with tempfile.TemporaryDirectory() as carpeta:
            destino = Path(carpeta) / "evento.doc"
            capturar_impresion(pagina, boton, destino)
            assert (destino.read_bytes()) == (documento)
        ruta.fulfill.assert_called_once_with(status=204, body="")
        pagina.unroute.assert_called_once()

    def test_captura_rechaza_respuesta_no_word(self):
        respuesta = Mock(status=200, headers={"content-type": "text/html"})
        respuesta.body.return_value = b"sesion expirada"
        ruta = Mock()
        ruta.fetch.return_value = respuesta
        pagina = Mock()
        pagina.expect_response.return_value = nullcontext()
        boton = Mock()
        boton.get_attribute.return_value = "control"
        boton.click.side_effect = lambda: pagina.route.call_args.args[1](ruta)
        with tempfile.TemporaryDirectory() as carpeta:
            destino = Path(carpeta) / "evento.doc"
            with pytest.raises(ValueError, match="Word válido"):
                capturar_impresion(pagina, boton, destino)
            assert not (destino.exists())
        pagina.unroute.assert_called_once()

    def test_captura_detecta_sesion_expirada(self):
        respuesta = Mock(status=200, headers={"content-type": "text/html"})
        respuesta.body.return_value = "<html>Inicio de sesión</html>".encode("utf-8")
        ruta = Mock()
        ruta.fetch.return_value = respuesta
        pagina = Mock()
        pagina.expect_response.return_value = nullcontext()
        boton = Mock()
        boton.get_attribute.return_value = "control"
        boton.click.side_effect = lambda: pagina.route.call_args.args[1](ruta)
        with tempfile.TemporaryDirectory() as carpeta:
            destino = Path(carpeta) / "evento.doc"
            with pytest.raises(SesionExpiradaError):
                capturar_impresion(pagina, boton, destino)
            assert not (destino.exists())
