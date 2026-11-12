import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from src.consolidado import COLUMNAS, extraer_documento, generar_consolidado
from src.reportes import Evento


class PruebasConsolidado(unittest.TestCase):
    def test_fecha_hora_y_hipervinculos_corresponden_a_cada_documento(self):
        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            primero = carpeta / "Doc1.doc"
            segundo = carpeta / "Doc2.doc"
            contenido = '<html><table><tr><td class="titleFieldValue">7 Articulo</td><td></td></tr><tr><td>Cantidad</td><td>1 Unidad</td></tr></table></html>'
            for documento in (primero, segundo):
                documento.write_text(contenido, encoding="utf-8")
            eventos = {
                primero: Evento("Uno", "Doc1", datetime(2030, 11, 24, 15, 47), "RFP", "https://s3.ariba.com/evento1"),
                segundo: Evento("Dos", "Doc2", None, "RFP", "https://s3.ariba.com/evento2"),
            }
            destino = carpeta / "resumen.xlsx"
            generar_consolidado(destino, [(primero, False), (segundo, False)], eventos)
            libro = load_workbook(destino)
            hoja = libro.active
            self.assertEqual(hoja["F1"].value, "Fecha Finalización")
            self.assertEqual(hoja["G1"].value, "Fecha Entrega")
            self.assertEqual(hoja["F2"].value, datetime(2030, 11, 24, 15, 47))
            self.assertIn("hh:mm", hoja["F2"].number_format)
            self.assertEqual(hoja["F3"].value, "N/A")
            self.assertEqual((hoja["A2"].value, hoja["A2"].hyperlink.target), ("Doc1.doc", "https://s3.ariba.com/evento1"))
            self.assertEqual((hoja["A3"].value, hoja["A3"].hyperlink.target), ("Doc2.doc", "https://s3.ariba.com/evento2"))
            libro.close()
            with self.assertRaisesRegex(ValueError, "Faltan datos"):
                generar_consolidado(carpeta / "incompleto.xlsx", [(primero, False)], {})

    def test_descarta_secciones_numeradas_y_conserva_items_con_campos_vacios(self):
        with tempfile.TemporaryDirectory() as carpeta:
            documento = Path(carpeta) / "evento.doc"
            documento.write_text('''<html><table>
                <tr><td class="titleFieldValue">11 Articulo anterior</td><td></td></tr>
                <tr><td>Número de material</td><td>00042</td></tr><tr><td>Cantidad</td><td>2 Unidades</td></tr>
                <tr><td class="titleFieldValue">12 Información a considerar</td><td></td></tr>
                <tr><td class="titleFieldValue">13 Documentos de licitación</td><td><a href="#archivo">bases.pdf</a></td></tr>
                <tr><td class="titleFieldValue">16 Oferta Técnica</td><td></td></tr>
                <tr><td class="titleFieldValue">17 Oferta Económica</td><td></td></tr>
                <tr><td class="titleFieldValue">17.1.1 Pregunta administrativa</td><td><a href="#respuesta">Si</a></td></tr>
                <tr><td class="titleFieldValue">18 Articulo sin cantidad informada</td><td></td></tr>
                <tr><td>Cantidad</td><td></td></tr>
                <tr><td class="titleFieldValue">18.1.1 Servicio descrito</td><td></td></tr>
                <tr><td class="descriptionFieldValue">Detalle del servicio</td><td></td></tr></table></html>''', encoding="utf-8")
            filas = extraer_documento(documento, False)
            nombres = ["11 Articulo anterior", "18 Articulo sin cantidad informada", "18.1.1 Servicio descrito"]
            self.assertEqual([fila["Ítem Principal"] for fila in filas], nombres)
            self.assertEqual((filas[0]["N.° Material"], filas[1]["Cantidad"]), ("42", "N/A"))
            resumen = Path(carpeta) / "resumen.xlsx"
            generar_consolidado(resumen, [(documento, False)])
            libro = load_workbook(resumen, read_only=True, data_only=True)
            self.assertEqual([fila[2] for fila in list(libro.active.values)[1:]], nombres)
            libro.close()

    def test_documento_ficticio_coincide_con_consolidado(self):
        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            documento = carpeta / "evento_ficticio.doc"
            documento.write_text('''<html><table><tr><td>Regiones</td><td>Zona de prueba</td></tr></table>
                <table><tr><td class="titleFieldValue">7 Articulo ficticio</td><td></td></tr>
                <tr><td>Número de material</td><td><a href="#MaterialNumber">000001234</a></td></tr>
                <tr><td>Cantidad</td><td><a href="#QUANTITY">8 Unidades</a></td></tr>
                <tr><td>Texto Largo</td><td>Descripcion ficticia</td></tr>
                <tr><td>Fecha de entrega solicitada</td><td>15/12/2030</td></tr>
                <tr><td class="titleFieldValue">8 Otro articulo</td><td></td></tr>
                <tr><td>Cantidad</td><td>2 Cajas</td></tr></table></html>''', encoding="utf-8")
            esperadas = [
                ("evento_ficticio.doc", "Zona de prueba", "7 Articulo ficticio", "1234", "8 Unidades", "N/A", "15/12/2030", "Descripcion ficticia"),
                ("evento_ficticio.doc", "Zona de prueba", "8 Otro articulo", "N/A", "2 Cajas", "N/A", "N/A", "N/A"),
            ]
            filas = extraer_documento(documento, False)
            self.assertEqual([tuple(fila[columna] for columna in COLUMNAS) for fila in filas], esperadas)
            destino = carpeta / "resumen.xlsx"
            self.assertEqual(generar_consolidado(destino, [(documento, False)]), 2)
            generado = load_workbook(destino, read_only=True, data_only=True)
            self.assertEqual(list(generado.active.values), [COLUMNAS, *esperadas])
            generado.close()

    def test_documento_sin_items_se_marca_para_revision(self):
        with tempfile.TemporaryDirectory() as carpeta:
            documento = Path(carpeta) / "ejemplo.doc"
            documento.write_text("<html><body></body></html>", encoding="utf-8")
            self.assertEqual(extraer_documento(documento, False)[0]["Descripción"], "Revisar manualmente")

    def test_documento_sr_con_items_anidados(self):
        with tempfile.TemporaryDirectory() as carpeta:
            documento = Path(carpeta) / "SR123.doc"
            documento.write_text("""<html><table><tr><td>Regiones</td><td>TE01 El Teniente</td></tr></table>
                <table><tr><td class="titleFieldValue">3.1.1 Material</td><td></td></tr>
                <tr><td class="descriptionFieldValue">Descripcion detallada</td><td></td></tr>
                <tr><td>Cantidad</td><td>3 Unidades</td></tr>
                <tr><td>Fecha de entrega solicitada</td><td>10/12/2026</td></tr></table></html>""", encoding="utf-8")
            fila = extraer_documento(documento, True)[0]
            self.assertEqual((fila["Región"], fila["Ítem Principal"], fila["N.° Material"], fila["Cantidad"], fila["Fecha Entrega"], fila["Descripción"]),
                             ("TE01 El Teniente", "3.1.1 Material", "N/A", "3 Unidades", "10/12/2026", "Descripcion detallada"))

    def test_documento_sr_ficticio_con_varios_items(self):
        with tempfile.TemporaryDirectory() as carpeta:
            documento = Path(carpeta) / "SR_ficticio.doc"
            documento.write_text('''<html><table><tr><td>Regiones</td><td>Zona ficticia</td></tr></table>
                <table><tr><td class="titleFieldValue">3 Oferta</td><td></td></tr>
                <tr><td class="titleFieldValue">3.1 Itemizado</td><td></td></tr>
                <tr><td class="titleFieldValue">3.1.1 Equipo ficticio</td><td></td></tr>
                <tr><td class="descriptionFieldValue">Detalle del primer equipo</td><td></td></tr>
                <tr><td>Cantidad</td><td><a href="#QUANTITY">1 Unidad</a></td></tr>
                <tr><td>Fecha de entrega solicitada</td><td>10/12/2030</td></tr>
                <tr><td class="titleFieldValue">3.1.2 Segundo equipo</td><td></td></tr>
                <tr><td class="descriptionFieldValue">Detalle del segundo equipo</td><td></td></tr>
                <tr><td>Cantidad</td><td>4 Cajas</td></tr></table></html>''', encoding="utf-8")
            filas = extraer_documento(documento, True)
            self.assertEqual(len(filas), 2)
            self.assertEqual([(fila["Región"], fila["Ítem Principal"], fila["N.° Material"], fila["Cantidad"], fila["Fecha Entrega"], fila["Descripción"]) for fila in filas], [
                ("Zona ficticia", "3.1.1 Equipo ficticio", "N/A", "1 Unidad", "10/12/2030", "Detalle del primer equipo"),
                ("Zona ficticia", "3.1.2 Segundo equipo", "N/A", "4 Cajas", "N/A", "Detalle del segundo equipo"),
            ])

    def test_formatos_ficticios_y_consolidado(self):
        with tempfile.TemporaryDirectory() as directorio:
            carpeta = Path(directorio)
            contenidos = {
                "sct.doc": '''<tr><td class="titleFieldValue">7 Articulo</td><td></td></tr>
                    <tr><td>Número de material</td><td><a href="#MaterialNumber">000005432</a></td></tr>
                    <tr><td>Cantidad</td><td>1,200 Unidades</td></tr><tr><td>Texto Largo</td><td>Detalle ficticio</td></tr>''',
                "sr.doc": '''<tr><td class="titleFieldValue">3.1.1 Material A</td><td></td></tr>
                    <tr><td class="descriptionFieldValue">Descripcion A</td><td></td></tr><tr><td>Cantidad</td><td>3 Unidades</td></tr>
                    <tr><td class="titleFieldValue">3.1.2 Material B</td><td></td></tr>
                    <tr><td class="descriptionFieldValue">Descripcion B</td><td></td></tr><tr><td>Cantidad</td><td>5 Metros</td></tr>''',
                "rfp.doc": '''<tr><td class="titleFieldValue">6.3.1 Servicio A</td><td></td></tr>
                    <tr><td class="descriptionFieldValue"><div>Primera linea</div><div>Segunda linea</div></td><td></td></tr>
                    <tr><td>Cantidad</td><td>1 Global</td></tr><tr><td>Fecha de entrega solicitada</td><td>20/12/2030</td></tr>
                    <tr><td class="titleFieldValue">6.3.2 Servicio B</td><td></td></tr>
                    <tr><td class="descriptionFieldValue"><div></div></td><td></td></tr><tr><td>Cantidad</td><td>2 Global</td></tr>''',
                "adjuntos.doc": '''<tr><td class="titleFieldValue">1 Instrucciones</td><td><a href="#archivo">instrucciones_ficticias.pdf</a></td></tr>
                    <tr><td class="titleFieldValue">2 Oferta</td><td><a href="#oferta">oferta_ficticia.xlsx</a></td></tr>
                    <tr><td><a href="#duplicado">oferta_ficticia.xlsx</a></td></tr>''',
            }
            for nombre, contenido in contenidos.items():
                (carpeta / nombre).write_text(f'<html><table><tr><td>Regiones</td><td>Zona ficticia</td></tr></table><table>{contenido}</table></html>', encoding="utf-8")
            sct = extraer_documento(carpeta / "sct.doc", False)
            sr = extraer_documento(carpeta / "sr.doc", True)
            rfp = extraer_documento(carpeta / "rfp.doc", False)
            adjuntos = extraer_documento(carpeta / "adjuntos.doc", False)
            self.assertEqual((sct[0]["N.° Material"], sct[0]["Cantidad"], sct[0]["Descripción"]), ("5432", "1,200 Unidades", "Detalle ficticio"))
            self.assertEqual([(fila["Ítem Principal"], fila["Cantidad"]) for fila in sr], [("3.1.1 Material A", "3 Unidades"), ("3.1.2 Material B", "5 Metros")])
            self.assertEqual((rfp[0]["Descripción"], rfp[0]["Fecha Entrega"], rfp[1]["Descripción"]), ("Primera linea Segunda linea", "20/12/2030", "N/A"))
            self.assertEqual(adjuntos[0]["Descripción"], "Revisar manualmente. Adjuntos: instrucciones_ficticias.pdf; oferta_ficticia.xlsx")
            destino = carpeta / "consolidado.xlsx"
            documentos = [(carpeta / nombre, nombre == "sr.doc") for nombre in contenidos]
            self.assertEqual(generar_consolidado(destino, documentos), 6)
            libro = load_workbook(destino, read_only=True, data_only=True)
            filas = list(libro.active.values)
            libro.close()
            self.assertEqual(filas, [COLUMNAS, *[tuple(fila[columna] for columna in COLUMNAS) for grupo in (sct, sr, rfp, adjuntos) for fila in grupo]])


if __name__ == "__main__":
    unittest.main()