import json
import os
import shutil
import tempfile
from pathlib import Path

from src.archivos import hash_archivo


def directorio_actual(carpeta: Path) -> tuple[Path, dict | None]:
    anterior = None
    while (carpeta / "estado.json").exists():
        estado = json.loads((carpeta / "estado.json").read_text(encoding="utf-8"))
        if estado.get("actualizacion") != "actualizacion":
            return carpeta, estado
        anterior = estado
        carpeta = carpeta / "actualizacion"
    return carpeta, anterior


def publicar_actualizacion(origen: Path, destino: Path, hash_previo: str) -> None:
    if not destino.exists():
        raise ValueError("El reporte anterior del día falta; requiere revisión")
    hash_destino = hash_archivo(destino)
    if hash_destino == hash_archivo(origen):
        return
    if hash_destino != hash_previo:
        raise ValueError("El reporte anterior del día cambió; requiere revisión")
    with tempfile.TemporaryDirectory(dir=destino.parent) as temporal:
        preparado = Path(temporal) / destino.name
        shutil.copyfile(origen, preparado)
        os.replace(preparado, destino)
