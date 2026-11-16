import os
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError, sync_playwright

from src.archivos import BASE


def convertir_documentos_pdf(documentos: list[Path], carpeta: Path) -> list[Path]:
    nombres = [documento.with_suffix(".pdf").name for documento in documentos]
    if not documentos or len({nombre.casefold() for nombre in nombres}) != len(nombres):
        raise ValueError("No hay documentos o los nombres PDF están duplicados")
    ruta_navegadores = str(BASE / ".datos" / "navegadores")
    if getattr(sys, "frozen", False):
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = ruta_navegadores
    else:
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", ruta_navegadores)
    resultados = list(documentos)
    try:
        with ThreadPoolExecutor(max_workers=1) as ejecutor:
            ejecutor.submit(_generar_pdfs, documentos, carpeta, nombres, resultados).result()
    except (PlaywrightError, OSError) as error:
        print(f"No se pudo completar la conversión a PDF: {error}. Se conservarán los documentos originales restantes.", flush=True)
    return resultados


def _generar_pdfs(
    documentos: list[Path], carpeta: Path, nombres: list[str], resultados: list[Path]
) -> None:
    with sync_playwright() as playwright:
        navegador = playwright.chromium.launch(headless=True)
        try:
            contexto = navegador.new_context(java_script_enabled=False)
            contexto.route("**/*", lambda ruta: ruta.abort())
            pagina = contexto.new_page()
            for indice, (documento, nombre) in enumerate(zip(documentos, nombres)):
                try:
                    pagina.set_content(documento.read_text(encoding="utf-8-sig"), wait_until="load")
                    pagina.pdf(path=str(carpeta / nombre), format="A4", print_background=True)
                except (PlaywrightError, OSError, UnicodeError) as error:
                    print(f"No se pudo convertir {documento.name} a PDF: {error}. Se incluirá el original.", flush=True)
                else:
                    resultados[indice] = carpeta / nombre
        finally:
            navegador.close()