import re
from pathlib import Path

from bs4 import BeautifulSoup
import xlsxwriter

from openpyxl import load_workbook

COLUMNAS = (
    "Licitación",
    "Región",
    "Ítem Principal",
    "N.° Material",
    "Cantidad",
    "Fecha Finalización",
    "Fecha Entrega",
    "Descripción",
)


def texto_limpio(texto: str) -> str:
    return re.sub(r"\s+", " ", texto).strip() or "N/A"


def extraer_documento(ruta: Path, es_sr: bool) -> list[dict[str, str]]:
    contenido = ruta.read_text(encoding="utf-8-sig")
    documento = BeautifulSoup(contenido, "html.parser")
    if not documento.find("html"):
        raise ValueError(f"El documento no contiene HTML de Ariba: {ruta.name}")
    etiqueta_region = documento.find(
        string=re.compile(r"Regiones|Organización de compras", re.I)
    )
    fila_region = etiqueta_region.find_parent("tr") if etiqueta_region else None
    celdas_region = fila_region.find_all("td") if fila_region else []
    region = (
        texto_limpio(celdas_region[-1].get_text(" ", strip=True))
        if celdas_region
        else "N/A"
    )
    resultado = []
    for titulo in documento.find_all("td", class_="titleFieldValue"):
        nombre = texto_limpio(titulo.get_text(" ", strip=True))
        normal = re.match(r"^(\d+)\s+", nombre)
        subitem = re.match(r"^\d+(?:\.\d+){2,}\s+", nombre)
        if not (subitem or (not es_sr and normal and int(normal.group(1)) >= 7)):
            continue
        tabla = titulo.find_parent("table")
        if tabla is None:
            continue
        filas = tabla.find_all("tr")
        inicio = next(
            (
                numero
                for numero, fila in enumerate(filas)
                if titulo in fila.find_all("td")
            ),
            None,
        )
        if inicio is None:
            continue
        material = cantidad = entrega = texto_largo = "N/A"
        tiene_campos_item = False
        for fila in filas[inicio + 1 :]:
            siguiente = fila.find("td", class_="titleFieldValue")
            if siguiente:
                break
            celdas = fila.find_all("td")
            if (
                subitem
                and celdas
                and "descriptionFieldValue" in celdas[0].get("class", [])
                and texto_largo == "N/A"
            ):
                texto_largo = texto_limpio(celdas[0].get_text(" ", strip=True))
            if len(celdas) < 2:
                continue
            etiqueta = texto_limpio(celdas[0].get_text(" ", strip=True)).casefold()
            valor = texto_limpio(celdas[-1].get_text(" ", strip=True))
            if "número de material" in etiqueta:
                tiene_campos_item = True
                material = "".join(filter(str.isdigit, valor)).lstrip("0") or "N/A"
            elif etiqueta == "cantidad" or (not subitem and "cantidad" in etiqueta):
                tiene_campos_item = True
                cantidad = valor
            elif "entrega solicitada" in etiqueta:
                tiene_campos_item = True
                entrega = valor
            elif "texto largo" in etiqueta:
                tiene_campos_item = True
                texto_largo = valor
        if tiene_campos_item or (subitem and texto_largo != "N/A"):
            resultado.append(
                dict(
                    zip(
                        COLUMNAS,
                        (
                            ruta.name,
                            region,
                            nombre,
                            material,
                            cantidad,
                            "N/A",
                            entrega,
                            texto_largo,
                        ),
                    )
                )
            )
    if not resultado:
        adjuntos = []
        for enlace in documento.select('a[href^="#"]'):
            nombre_adjunto = texto_limpio(enlace.get_text(" ", strip=True))
            if (
                re.search(r"\.(?:pdf|xlsx?|docx?|zip|txt)$", nombre_adjunto, re.I)
                and nombre_adjunto not in adjuntos
            ):
                adjuntos.append(nombre_adjunto)
        detalle = "Revisar manualmente"
        if adjuntos:
            detalle += ". Adjuntos: " + "; ".join(adjuntos)
        resultado.append(
            dict(
                zip(
                    COLUMNAS,
                    (
                        ruta.name,
                        region,
                        "Sin ítems detectados",
                        "N/A",
                        "N/A",
                        "N/A",
                        "N/A",
                        detalle,
                    ),
                )
            )
        )
    return resultado


def generar_consolidado(
    ruta: Path, documentos: list[tuple[Path, bool]], eventos_por_documento=None
) -> int:
    if not documentos:
        raise ValueError("No hay documentos para generar el consolidado")
    if eventos_por_documento is not None and any(
        documento not in eventos_por_documento for documento, _ in documentos
    ):
        raise ValueError("Faltan datos de eventos para el consolidado")
    filas = [
        (
            fila,
            es_sr,
            (
                eventos_por_documento[documento]
                if eventos_por_documento is not None
                else None
            ),
        )
        for documento, es_sr in documentos
        for fila in extraer_documento(documento, es_sr)
    ]
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with ruta.open("xb") as archivo:
        with xlsxwriter.Workbook(archivo) as libro:
            hoja = libro.add_worksheet("Consolidado")
            encabezado = libro.add_format(
                {
                    "bold": True,
                    "bg_color": "#004586",
                    "font_color": "white",
                    "border": 1,
                    "align": "center",
                    "valign": "vcenter",
                }
            )
            formatos = [
                libro.add_format(
                    {"bg_color": color, "border": 1, "text_wrap": True, "valign": "top"}
                )
                for color in ("#FFFFFF", "#BDD7EE", "#F4B183", "#ED7D31")
            ]
            formatos_fecha = [
                libro.add_format(
                    {
                        "bg_color": color,
                        "border": 1,
                        "text_wrap": True,
                        "valign": "top",
                        "num_format": "dd/mm/yyyy hh:mm",
                    }
                )
                for color in ("#FFFFFF", "#BDD7EE", "#F4B183", "#ED7D31")
            ]
            for numero, columna in enumerate(COLUMNAS):
                hoja.write_string(0, numero, columna, encabezado)
            ultimo = None
            alternancia = {False: False, True: False}
            for numero, (fila, es_sr, evento) in enumerate(filas, start=1):
                if fila["Licitación"] != ultimo:
                    ultimo = fila["Licitación"]
                    alternancia[es_sr] = not alternancia[es_sr]
                indice_formato = (2 if es_sr else 0) + int(alternancia[es_sr])
                formato = formatos[indice_formato]
                for columna, nombre in enumerate(COLUMNAS):
                    if nombre == "Licitación" and evento is not None:
                        if (
                            hoja.write_url(
                                numero,
                                columna,
                                evento.enlace,
                                formato,
                                string=fila[nombre],
                            )
                            != 0
                        ):
                            raise ValueError("No se pudo escribir el enlace del evento")
                        continue
                    if (
                        nombre == "Fecha Finalización"
                        and evento is not None
                        and evento.finalizacion is not None
                    ):
                        hoja.write_datetime(
                            numero,
                            columna,
                            evento.finalizacion,
                            formatos_fecha[indice_formato],
                        )
                        continue
                    valor = str(fila[nombre])
                    hoja.write_string(
                        numero,
                        columna,
                        valor if len(valor) <= 32000 else valor[:32000] + "...",
                        formato,
                    )
            for numero, ancho in enumerate((35, 25, 45, 18, 18, 22, 22, 90)):
                hoja.set_column(numero, numero, ancho)
            hoja.autofilter(0, 0, len(filas), len(COLUMNAS) - 1)
            hoja.freeze_panes(1, 0)
    return len(filas)


def validar_formato_resumen(ruta: Path) -> None:
    libro = load_workbook(ruta, read_only=True, data_only=True)
    try:
        encabezados = tuple(
            celda.value for celda in next(libro.active.iter_rows(min_row=1, max_row=1))
        )
        if encabezados != COLUMNAS:
            raise ValueError(
                "El resumen pendiente tiene un formato anterior; requiere revisión antes de enviarlo"
            )
    finally:
        libro.close()
