import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from playwright.sync_api import sync_playwright

from src.ariba import descargar_evento, exportar_eventos, iniciar_sesion, seleccionar_codelco

LOGIN = """<html><body><p>Inicio de sesión de proveedor</p><input id="userid">
<a href="#" onclick="document.querySelector('#password').hidden=false; this.hidden=true">Siguiente</a>
<div id="password" hidden><input type="password"><a href="/dashboard">Conectarse</a></div>
</body></html>"""
DASHBOARD = """<html><body><a href="/sourcing">Más...</a></body></html>"""
SOURCING = """<html><body><a href="#" onclick="document.querySelector('#clientes').hidden=false">Más...</a>
<div id="clientes" hidden><a href="#" onclick="document.querySelector('iframe').src='/eventos'">Corporación Nacional del Cobre (Ambiente Productivo)</a></div>
<iframe></iframe></body></html>"""
EVENTOS = """<html><body><h2>Filtros de búsqueda de eventos</h2><table><tr><td>Estado del evento:</td><td>
<div role="combobox" onclick="this.dataset.intentos = Number(this.dataset.intentos || 0) + 1; if (this.dataset.intentos > 1) document.querySelector('#opciones').hidden=false"><input value=""><span>v</span></div>
<div role="listbox" id="opciones" hidden><div role="option" onclick="document.querySelector('input').value='Abierto'; document.querySelector('#opciones').hidden=true">Abierto</div></div>
</td></tr></table><button>Buscar</button><button onclick="this.dataset.intentos = Number(this.dataset.intentos || 0) + 1; if (this.dataset.intentos > 1) document.querySelector('#exportar').hidden=false"><span title="Menú Opciones de tabla">tabla</span></button>
<div id="exportar" hidden><a role="menuitem" href="/descarga.xls">Exportar todas las filas</a></div></body></html>"""
DETALLE = """<html><body><span>Doc1 - 6001</span><button onclick="location.href='/descarga.doc'">Imprimir información de evento</button></body></html>"""
DOCUMENTO = """<!-- class: ariba.sourcing.rfxui.PrintRFXEncode --><html><table><tr><td>Fecha de vencimiento</td><td>24/11/2026 15:00</td></tr></table></html>"""


class PruebasNavegacion(unittest.TestCase):
    def test_login_con_sesion_de_propuestas_existente(self):
        carpeta_programa = Path(__file__).resolve().parent.parent
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(carpeta_programa / ".datos" / "navegadores")
        with sync_playwright() as playwright:
            navegador = playwright.chromium.launch()
            pagina = navegador.new_page()
            pagina.route("**/*", lambda ruta: ruta.fulfill(
                status=200, content_type="text/html; charset=utf-8", body=SOURCING
            ))
            iniciar_sesion(pagina, None, None)
            self.assertTrue(pagina.get_by_role("link", name="Más...").is_visible())
            navegador.close()

    def test_seleccion_directa_de_codelco(self):
        carpeta_programa = Path(__file__).resolve().parent.parent
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(carpeta_programa / ".datos" / "navegadores")
        with sync_playwright() as playwright:
            navegador = playwright.chromium.launch()
            pagina = navegador.new_page()
            pagina.set_default_timeout(1000)
            pagina.set_content("""<a href="#" onclick="document.querySelector('#clientes').hidden=false">Más...</a>
                <div id="clientes" hidden>
                <a href="#" onclick="this.dataset.seleccionado='si'">Corporación Nacional del Cobre (Ambiente Productivo)</a></div>""")
            with patch("src.ariba.esperar_marco_eventos"):
                seleccionar_codelco(pagina)
            self.assertEqual(pagina.locator("[data-seleccionado]").count(), 1)
            navegador.close()

    def test_login_con_sesion_existente_tras_usuario(self):
        carpeta_programa = Path(__file__).resolve().parent.parent
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(
            carpeta_programa / ".datos" / "navegadores"
        )
        with sync_playwright() as playwright:
            navegador = playwright.chromium.launch()
            pagina = navegador.new_page()

            def respuesta(ruta):
                if "Sourcing.aw" in ruta.request.url:
                    contenido = '<html><body><input id="userid"><a href="/Supplier.aw/panel">Siguiente</a></body></html>'
                else:
                    contenido = DASHBOARD
                ruta.fulfill(status=200, content_type="text/html; charset=utf-8", body=contenido)

            pagina.route("**/*", respuesta)
            iniciar_sesion(pagina, "usuario@ejemplo.com", "clave-ficticia")
            self.assertTrue(pagina.url.endswith("/Supplier.aw/panel"))
            navegador.close()

    def test_login_con_usuario_y_contrasena_en_la_misma_pagina(self):
        carpeta_programa = Path(__file__).resolve().parent.parent
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(
            carpeta_programa / ".datos" / "navegadores"
        )
        with sync_playwright() as playwright:
            navegador = playwright.chromium.launch()
            pagina = navegador.new_page()

            def respuesta(ruta):
                if "Sourcing.aw" in ruta.request.url:
                    contenido = """<html><body><input id="usuario" type="email"><input id="clave" type="password">
                        <button onclick="if (document.querySelector('#usuario').value === 'usuario@ejemplo.com' && document.querySelector('#clave').value === 'clave-ficticia') location.href='/dashboard'">Inicio de sesión</button></body></html>"""
                else:
                    contenido = DASHBOARD
                ruta.fulfill(status=200, content_type="text/html; charset=utf-8", body=contenido)

            pagina.route("**/*", respuesta)
            iniciar_sesion(pagina, "usuario@ejemplo.com", "clave-ficticia")
            self.assertTrue(pagina.url.endswith("/dashboard"))
            navegador.close()

    def test_flujo_de_login_exportacion_y_descarga(self):
        carpeta_programa = Path(__file__).resolve().parent.parent
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(
            carpeta_programa / ".datos" / "navegadores"
        )
        with tempfile.TemporaryDirectory() as carpeta, sync_playwright() as playwright:
            navegador = playwright.chromium.launch()
            contexto = navegador.new_context(accept_downloads=True)
            pagina = contexto.new_page()

            def respuesta(ruta):
                url = ruta.request.url
                if "Sourcing.aw" in url:
                    contenido = LOGIN
                elif url.endswith("/dashboard"):
                    contenido = SOURCING
                elif url.endswith("/sourcing"):
                    contenido = SOURCING
                elif url.endswith("/eventos"):
                    contenido = EVENTOS
                elif url.endswith("/evento"):
                    contenido = DETALLE
                elif url.endswith("/descarga.xls"):
                    ruta.fulfill(
                        status=200,
                        headers={
                            "Content-Type": "application/octet-stream",
                            "Content-Disposition": 'attachment; filename="data.xls"',
                        },
                        body="exportacion",
                    )
                    return
                elif url.endswith("/descarga.doc"):
                    ruta.fulfill(
                        status=200,
                        headers={
                            "Content-Type": "application/octet-stream",
                            "Content-Disposition": 'attachment; filename="evento.doc"',
                        },
                        body=DOCUMENTO,
                    )
                    return
                else:
                    ruta.abort()
                    return
                ruta.fulfill(
                    status=200, content_type="text/html; charset=utf-8", body=contenido
                )

            pagina.route("**/*", respuesta)
            iniciar_sesion(pagina, "usuario@ejemplo.com", "clave-ficticia")
            archivo = Path(carpeta) / "data.xls"
            exportar_eventos(pagina, archivo)
            self.assertEqual(archivo.read_text(), "exportacion")
            evento = SimpleNamespace(
                identificador="Doc1",
                titulo="6001",
                enlace="https://service.ariba.com/evento",
            )
            documento, aviso = descargar_evento(pagina, evento, Path(carpeta))
            self.assertIsNone(aviso)
            self.assertEqual(documento.name, "24-11 Doc1 - 6001.doc")
            navegador.close()


if __name__ == "__main__":
    unittest.main()
