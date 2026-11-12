from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re

from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.differential import DifferentialStyle


@dataclass(frozen=True)
class Evento:
    titulo: str
    identificador: str
    finalizacion: datetime | None
    tipo: str
    enlace: str


def leer_exportacion(ruta: Path) -> list[Evento]:
    contenido = ruta.read_text(encoding="utf-8-sig")
    documento = BeautifulSoup(contenido, "html.parser")
    cabeceras = [
        celda.get_text(" ", strip=True).casefold()
        for celda in documento.select("div.tableHead th")
    ]
    if cabeceras != [
        "título",
        "id",
        "hora de finalización",
        "estado",
        "tipo de evento",
    ]:
        raise ValueError("La exportación tiene columnas diferentes a las esperadas")
    tabla = documento.select_one("table.tableBody")
    if tabla is None:
        raise ValueError(
            "La exportación de Ariba no contiene la tabla de eventos esperada"
        )

    eventos = []
    for fila in tabla.find_all("tr", recursive=False):
        celdas = fila.find_all("td", recursive=False)
        if not celdas:
            continue
        if len(celdas) != 5:
            raise ValueError(
                "La exportación contiene una fila con columnas inesperadas"
            )
        titulo, identificador, finalizacion, estado, tipo = [
            celda.get_text(" ", strip=True).replace("\xa0", "").strip()
            for celda in celdas
        ]
        if estado.casefold() != "abierto":
            raise ValueError(
                f"La exportación contiene un evento no abierto: {identificador}"
            )
        vinculo = celdas[0].find("a", href=True)
        if not titulo or not identificador or not vinculo:
            raise ValueError("Hay un evento sin título, ID o enlace")
        formato_fecha = (
            "%d/%m/%Y %I:%M %p"
            if re.search(r"\s[AP]M$", finalizacion, re.I)
            else "%d/%m/%Y %H:%M"
        )
        fecha = datetime.strptime(finalizacion, formato_fecha) if finalizacion else None
        eventos.append(Evento(titulo, identificador, fecha, tipo, vinculo["href"]))

    if not eventos:
        raise ValueError(
            "La exportación no contiene eventos abiertos; no se actualizará la referencia"
        )
    if len({evento.identificador for evento in eventos}) != len(eventos):
        raise ValueError("La exportación contiene ID duplicados")
    return eventos


def leer_ids_anteriores(ruta: Path) -> list[str]:
    libro = load_workbook(ruta, read_only=True, data_only=True)
    try:
        hoja = libro.active
        if hoja is None or str(hoja.cell(1, 2).value).strip().casefold() != "id":
            raise ValueError("El reporte anterior no tiene la columna ID esperada en B")
        identificadores = [
            str(fila[0]).strip()
            for fila in hoja.iter_rows(
                min_row=2, min_col=2, max_col=2, values_only=True
            )
            if fila[0] is not None and str(fila[0]).strip()
        ]
        if not identificadores or len(set(identificadores)) != len(identificadores):
            raise ValueError(
                "El reporte anterior no tiene ID válidos o contiene duplicados"
            )
        return identificadores
    finally:
        libro.close()


def eventos_nuevos(eventos: list[Evento], anteriores: list[str]) -> list[Evento]:
    conocidos = set(anteriores)
    return [evento for evento in eventos if evento.identificador not in conocidos]


def buscar_reporte_anterior(carpeta: Path, hoy: date) -> Path:
    candidatos = []
    for ruta in carpeta.glob("*.xlsx"):
        nuevo = re.fullmatch(
            r"(\d{4}-\d{2}-\d{2})_Licitaciones_Codelco_\(Ariba\)\.xlsx", ruta.name, re.I
        )
        if nuevo:
            try:
                fecha = date.fromisoformat(nuevo.group(1))
            except ValueError:
                continue
            if fecha < hoy:
                candidatos.append((fecha, ruta))
            continue
        coincidencia = re.fullmatch(
            r"LIC\s+ARIBA\s+(\d{1,2})-(\d{1,2})\.xlsx", ruta.name, re.I
        )
        if not coincidencia:
            continue
        dia, mes = map(int, coincidencia.groups())
        for anio in (hoy.year, hoy.year - 1):
            try:
                fecha = date(anio, mes, dia)
            except ValueError:
                continue
            if fecha < hoy:
                candidatos.append((fecha, ruta))
                break
    if not candidatos:
        raise FileNotFoundError("No se encontró un reporte anterior para comparar")
    fecha_mas_reciente = max(fecha for fecha, _ in candidatos)
    ultimos = [ruta for fecha, ruta in candidatos if fecha == fecha_mas_reciente]
    if len(ultimos) != 1:
        raise ValueError("Hay más de un reporte anterior con la misma fecha")
    return ultimos[0]


def guardar_reporte(ruta: Path, eventos: list[Evento], anteriores: list[str]) -> None:
    libro = Workbook()
    hoja = libro.active
    hoja.title = "data"
    hoja.append(["Título", "ID", "Hora de finalización", "Tipo de evento", "ID"])
    for evento in eventos:
        hoja.append(
            [evento.titulo, evento.identificador, evento.finalizacion, evento.tipo]
        )
        celda_titulo = hoja.cell(hoja.max_row, 1)
        celda_titulo.data_type = "s"
        celda_titulo.hyperlink = evento.enlace
        celda_titulo.style = "Hyperlink"
        hoja.cell(hoja.max_row, 3).number_format = "dd-mm-yyyy hh:mm"
    for numero, identificador in enumerate(anteriores, start=2):
        hoja.cell(numero, 5, identificador)
    ultima_fila = max(hoja.max_row, 2)
    formato = DifferentialStyle(
        fill=PatternFill(fill_type="solid", fgColor="FFC7CE"),
        font=Font(color="9C0006"),
    )
    hoja.conditional_formatting.add(
        f"B2:B{ultima_fila} E2:E{ultima_fila}",
        Rule(type="uniqueValues", dxf=formato),
    )
    for columna, ancho in {"A": 75, "B": 20, "C": 23, "D": 18, "E": 20}.items():
        hoja.column_dimensions[columna].width = ancho
    for celda in hoja[1]:
        celda.font = Font(bold=True)
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = f"A1:D{len(eventos) + 1}"
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with ruta.open("xb") as archivo:
        libro.save(archivo)
