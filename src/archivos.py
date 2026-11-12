import hashlib
import json
import shutil
import sys
import zipfile
from contextlib import contextmanager
from pathlib import Path

BASE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent

def hash_archivo(ruta: Path) -> str:
    resumen = hashlib.sha256()
    with ruta.open("rb") as archivo:
        for bloque in iter(lambda: archivo.read(1024 * 1024), b""):
            resumen.update(bloque)
    return resumen.hexdigest()


def copiar_sin_sobrescribir(origen: Path, destino: Path) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists():
        if hash_archivo(origen) != hash_archivo(destino):
            raise FileExistsError(f"El archivo de destino ya existe y es diferente: {destino}")
        return
    with origen.open("rb") as entrada, destino.open("xb") as salida:
        shutil.copyfileobj(entrada, salida)


def crear_zip_documentos(destino: Path, documentos: list[Path]) -> None:
    if not documentos or any(documento.suffix.casefold() != ".doc" for documento in documentos):
        raise ValueError("El ZIP requiere documentos Word de la ejecución")
    if len({documento.name.casefold() for documento in documentos}) != len(documentos):
        raise ValueError("Hay documentos con nombres duplicados para el ZIP")
    with zipfile.ZipFile(destino, "x", compression=zipfile.ZIP_DEFLATED) as archivo:
        for documento in documentos:
            archivo.write(documento, arcname=documento.name)


def guardar_estado(ruta: Path, estado: dict) -> None:
    temporal = ruta.with_suffix(".tmp")
    temporal.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")
    temporal.replace(ruta)


@contextmanager
def bloqueo_ejecucion():
    import msvcrt

    carpeta = BASE / ".datos"
    carpeta.mkdir(parents=True, exist_ok=True)
    with (carpeta / "ejecucion.lock").open("a+b") as archivo:
        archivo.seek(0)
        try:
            msvcrt.locking(archivo.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as error:
            raise RuntimeError("Ya hay otra ejecución en curso") from error
        try:
            yield
        finally:
            archivo.seek(0)
            msvcrt.locking(archivo.fileno(), msvcrt.LK_UNLCK, 1)