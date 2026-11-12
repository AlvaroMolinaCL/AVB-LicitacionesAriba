import argparse
import os
import sys
from datetime import date
from pathlib import Path

from src.ariba import configurar_credenciales
from src.correo import configurar_correo, configurar_destinatarios, ejecutar_prueba_correo, ejecutar_resolucion_correo
from src.ejecucion import ejecutar_flujo_principal


BASE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
SERVICIO = "AVB-LicitacionesAriba"


def configurar_consola() -> None:
    if sys.platform != "win32":
        return
    os.system("chcp 65001 > nul")
    for flujo in (sys.stdout, sys.stderr):
        if flujo is not None and hasattr(flujo, "reconfigure"):
            flujo.reconfigure(encoding="utf-8", errors="backslashreplace")


def pausar_al_finalizar() -> None:
    if getattr(sys, "frozen", False) and "--sin-pausa" not in sys.argv and sys.stdin and sys.stdin.isatty():
        try:
            input("\nProceso finalizado. Presione Enter para cerrar esta ventana...")
        except EOFError:
            pass


def main() -> int:
    analizador = argparse.ArgumentParser(description="Automatización de licitaciones de Ariba")
    acciones = analizador.add_mutually_exclusive_group()
    acciones.add_argument("--configurar-credenciales", action="store_true")
    acciones.add_argument("--configurar-correo", action="store_true", help="Guarda el remitente y la contraseña de aplicación de Google")
    acciones.add_argument("--configurar-destinatarios", action="store_true", help="Configura los destinatarios y la copia del correo")
    acciones.add_argument("--probar-correo", action="store_true", help="Envía una prueba solo al remitente para la fecha indicada")
    acciones.add_argument("--resolver-correo", choices=("enviado", "reintentar"), help="Resuelve un envío incierto tras revisar Gmail")
    analizador.add_argument("--visible", action="store_true")
    acciones.add_argument("--validar-ariba", action="store_true", help="Comprueba la conexión y la exportación sin procesar licitaciones")
    analizador.add_argument("--sin-pausa", action="store_true", help="Evita la pausa final del ejecutable en tareas desatendidas")
    analizador.add_argument("--raiz", type=Path, default=BASE)
    analizador.add_argument("--fecha", type=date.fromisoformat, help="Reanudar una ejecución anterior en formato AAAA-MM-DD")
    analizador.add_argument("--correo-de-prueba", action="store_true", help="Resuelve el estado de un envío de prueba")
    analizador.add_argument("--nueva-prueba", action="store_true", help="Permite una nueva prueba después de un envío confirmado")
    opciones = analizador.parse_args()
    if opciones.correo_de_prueba and not opciones.resolver_correo:
        analizador.error("--correo-de-prueba requiere --resolver-correo")
    if opciones.nueva_prueba and not opciones.probar_correo:
        analizador.error("--nueva-prueba requiere --probar-correo")
    if opciones.configurar_credenciales:
        configurar_credenciales()
        return 0
    if opciones.configurar_correo:
        configurar_correo()
        return 0
    if opciones.configurar_destinatarios:
        configurar_destinatarios()
        return 0
    if opciones.resolver_correo:
        ejecutar_resolucion_correo(BASE, opciones.fecha or date.today(), opciones.correo_de_prueba, opciones.resolver_correo)
        return 0
    if opciones.probar_correo:
        ejecutar_prueba_correo(BASE, opciones.raiz, opciones.fecha or date.today(), opciones.nueva_prueba)
        return 0
    return ejecutar_flujo_principal(
        BASE, SERVICIO, opciones.raiz, opciones.fecha,
        opciones.visible, opciones.validar_ariba, opciones.sin_pausa,
    )


if __name__ == "__main__":
    configurar_consola()
    codigo_salida = 0
    try:
        codigo_salida = main()
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        codigo_salida = 1
    finally:
        pausar_al_finalizar()
    sys.exit(codigo_salida)