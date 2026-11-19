import os
import subprocess
import sys
import tempfile
import pytest
from pathlib import Path


@pytest.mark.skipif(os.name != "nt", reason="Requiere Windows")
class TestInstalador:
    def ejecutar_simulacion(
        self, fallo_descarga: bool, hash_valido: bool = True
    ) -> tuple[subprocess.CompletedProcess, bool, str]:
        raiz = Path(__file__).resolve().parent.parent
        contenido = (raiz / "scripts" / "preparar_python.ps1").read_text(
            encoding="utf-8"
        )
        simulacion = r"""
function Buscar-PythonCompatible { return $null }
function Es-PythonCompatible([string]$Ruta) { return (Test-Path -LiteralPath $Ruta -PathType Leaf) }
function Invoke-WebRequest { param([switch]$UseBasicParsing, $Uri, $OutFile)
    if ($env:AVB_FALLO_DESCARGA -eq '1') { throw 'descarga bloqueada' }
    New-Item -ItemType File -Path $OutFile -Force | Out-Null
}
function Start-Process { param($FilePath, $ArgumentList, [switch]$Wait, [switch]$PassThru)
    $ArgumentList | Set-Content -LiteralPath $env:AVB_ARGUMENTOS
    $destino = Join-Path $env:LOCALAPPDATA 'Programs\Python\AVB-Python314'
    New-Item -ItemType Directory -Path $destino -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $destino 'python.exe') -Force | Out-Null
    return @{ ExitCode = 0 }
}
"""
        contenido = contenido.replace(
            "\ntry {\n    $carpetaProyecto",
            simulacion + "\ntry {\n    $carpetaProyecto",
            1,
        )
        hash_descarga = (
            "92480239264744FE0386A1BD8275786FEFA126E3129E56A68E22B52C73329965"
            if hash_valido
            else "incorrecto"
        )
        contenido = contenido.replace(
            "(Get-FileHash -LiteralPath $instalador -Algorithm SHA256).Hash",
            f"'{hash_descarga}'",
        )
        contenido = contenido.replace(
            "& $python -m venv $entorno", f'& "{sys.executable}" -m venv $entorno'
        )
        with tempfile.TemporaryDirectory() as temporal:
            carpeta = Path(temporal)
            script = carpeta / "preparar_python.ps1"
            script.write_text(contenido, encoding="utf-8")
            entorno = os.environ.copy()
            entorno["LOCALAPPDATA"] = str(carpeta / "usuario")
            entorno["AVB_ARGUMENTOS"] = str(carpeta / "argumentos.txt")
            entorno["AVB_FALLO_DESCARGA"] = "1" if fallo_descarga else "0"
            resultado = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script),
                    "-Carpeta",
                    str(carpeta),
                ],
                env=entorno,
                capture_output=True,
                text=True,
                errors="replace",
            )
            existe_entorno = (carpeta / ".venv" / "Scripts" / "python.exe").exists()
            argumentos = (
                (carpeta / "argumentos.txt").read_text()
                if (carpeta / "argumentos.txt").exists()
                else ""
            )
            return resultado, existe_entorno, argumentos

    def test_instala_solo_para_el_usuario_sin_elevacion(self) -> None:
        resultado, existe_entorno, argumentos = self.ejecutar_simulacion(False)
        assert (resultado.returncode) == (0), resultado.stdout + resultado.stderr
        assert existe_entorno
        assert ("InstallAllUsers=0") in (argumentos)
        assert ("Include_launcher=0") in (argumentos)
        assert ("RunAs") not in (argumentos)

    def test_fallo_descarga_no_crea_entorno(self) -> None:
        resultado, existe_entorno, _ = self.ejecutar_simulacion(True)
        assert (resultado.returncode) != (0)
        assert ("descarga bloqueada") in (resultado.stdout)
        assert not (existe_entorno)

    def test_hash_incorrecto_no_instala(self) -> None:
        resultado, existe_entorno, argumentos = self.ejecutar_simulacion(False, False)
        assert (resultado.returncode) != (0)
        assert ("suma de seguridad") in (resultado.stdout)
        assert not (existe_entorno)
        assert (argumentos) == ("")

    def test_entorno_invalido_no_se_sobrescribe(self) -> None:
        raiz = Path(__file__).resolve().parent.parent
        with tempfile.TemporaryDirectory() as temporal:
            entorno = Path(temporal) / ".venv"
            entorno.mkdir()
            resultado = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(raiz / "scripts" / "preparar_python.ps1"),
                    "-Carpeta",
                    temporal,
                ],
                capture_output=True,
                text=True,
                errors="replace",
            )
            assert (resultado.returncode) != (0)
            assert ("entorno existente") in (resultado.stdout)
            assert entorno.exists()
