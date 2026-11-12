import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from src.reportes import (
    Evento,
    buscar_reporte_anterior,
    eventos_nuevos,
    guardar_reporte,
    leer_exportacion,
    leer_ids_anteriores,
)


class PruebasReportes(unittest.TestCase):
    def test_exportacion_acepta_horas_de_12_y_24_horas(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "exportacion.xls"
            ruta.write_text(
                """<div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div>
                <table class="tableBody">
                <tr><td><a href="https://s3.ariba.com/1">Primero</a></td><td>Doc1</td><td>13/11/2026 03:15 PM</td><td>Abierto</td><td>RFP</td></tr>
                <tr><td><a href="https://s3.ariba.com/2">Segundo</a></td><td>Doc2</td><td>14/11/2026 00:30</td><td>Abierto</td><td>RFP</td></tr>
                </table>""",
                encoding="utf-8",
            )
            eventos = leer_exportacion(ruta)
            self.assertEqual(eventos[0].finalizacion, datetime(2026, 11, 13, 15, 15))
            self.assertEqual(eventos[1].finalizacion, datetime(2026, 11, 14, 0, 30))

    def test_comparacion_reporte_e_hipervinculos(self):
        eventos = [
            Evento(
                "Existente",
                "Doc1",
                datetime(2026, 11, 12, 15),
                "RFP",
                "https://example.com/1",
            ),
            Evento("Nuevo", "Doc2", None, "RFP", "https://example.com/2"),
        ]
        anteriores = ["Doc1", "Doc3"]
        self.assertEqual(
            [evento.identificador for evento in eventos_nuevos(eventos, anteriores)],
            ["Doc2"],
        )
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "reporte.xlsx"
            guardar_reporte(ruta, eventos, anteriores)
            self.assertEqual(leer_ids_anteriores(ruta), ["Doc1", "Doc2"])
            libro = load_workbook(ruta)
            hoja = libro.active
            self.assertEqual(hoja["A3"].hyperlink.target, "https://example.com/2")
            self.assertEqual([hoja["E2"].value, hoja["E3"].value], anteriores)
            rango = next(iter(hoja.conditional_formatting))
            self.assertEqual(hoja.conditional_formatting[rango][0].type, "uniqueValues")
            libro.close()
            with self.assertRaises(FileExistsError):
                guardar_reporte(ruta, eventos, anteriores)

    def test_ultimo_reporte_considera_cambio_de_anio(self):
        with tempfile.TemporaryDirectory() as carpeta:
            directorio = Path(carpeta)
            (directorio / "LIC ARIBA 30-12.xlsx").touch()
            ultimo = directorio / "LIC ARIBA 31-12.xlsx"
            ultimo.touch()
            (directorio / "LIC ARIBA 01-01.xlsx").touch()
            self.assertEqual(
                buscar_reporte_anterior(directorio, date(2027, 1, 1)), ultimo
            )

    def test_busca_reportes_nuevos_y_antiguos_por_fecha(self):
        with tempfile.TemporaryDirectory() as carpeta:
            directorio = Path(carpeta)
            (directorio / "LIC ARIBA 30-12.xlsx").touch()
            ultimo = directorio / "2026-12-31_Licitaciones_Codelco_(Ariba).xlsx"
            ultimo.touch()
            (directorio / "2027-01-01_Licitaciones_Codelco_(Ariba).xlsx").touch()
            (directorio / "2025-12-31_Licitaciones_Codelco_(Ariba).xlsx").touch()
            (directorio / "2026-02-30_Licitaciones_Codelco_(Ariba).xlsx").touch()
            self.assertEqual(
                buscar_reporte_anterior(directorio, date(2027, 1, 1)), ultimo
            )
            (directorio / "LIC ARIBA 31-12.xlsx").touch()
            with self.assertRaisesRegex(ValueError, "misma fecha"):
                buscar_reporte_anterior(directorio, date(2027, 1, 1))

    def test_exportacion_y_reporte_anterior_ficticios(self):
        with tempfile.TemporaryDirectory() as carpeta:
            carpeta = Path(carpeta)
            anterior = carpeta / "2030-11-10_Licitaciones_Codelco_(Ariba).xlsx"
            libro = Workbook()
            hoja = libro.active
            hoja.append(
                ["Título", "ID", "Hora de finalización", "Tipo de evento", "ID"]
            )
            hoja.append(["Anterior", "Doc1", None, "RFP", "Doc3"])
            hoja.append(["Anterior 2", "Doc2", None, "RFP", None])
            libro.save(anterior)
            libro.close()
            exportacion = carpeta / "exportacion_ficticia.xls"
            exportacion.write_text(
                """<html><div class="tableHead"><table><tr><th>Título</th><th>ID</th><th>Hora de finalización</th><th>Estado</th><th>Tipo de evento</th></tr></table></div>
                <table class="tableBody"><tr><td><a href="https://s3.ariba.com/1">Existente</a></td><td>Doc1</td><td>11/11/2030 15:00</td><td>Abierto</td><td>RFP</td></tr>
                <tr><td><a href="https://s3.ariba.com/3">Nuevo</a></td><td>Doc3</td><td>12/11/2030 03:00 PM</td><td>Abierto</td><td>RFP</td></tr></table></html>""",
                encoding="utf-8",
            )
            self.assertEqual(
                buscar_reporte_anterior(carpeta, date(2030, 11, 11)), anterior
            )
            eventos = leer_exportacion(exportacion)
            anteriores = leer_ids_anteriores(anterior)
            self.assertEqual(anteriores, ["Doc1", "Doc2"])
            self.assertEqual(
                [
                    evento.identificador
                    for evento in eventos_nuevos(eventos, anteriores)
                ],
                ["Doc3"],
            )
            self.assertEqual(eventos[1].enlace, "https://s3.ariba.com/3")


if __name__ == "__main__":
    unittest.main()
