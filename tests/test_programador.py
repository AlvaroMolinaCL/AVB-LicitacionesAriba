import os
import subprocess
import tempfile
import uuid
from datetime import datetime, timedelta
from pathlib import Path

import pytest


@pytest.mark.skipif(os.name != "nt", reason="Requiere Windows")
class TestProgramador:
    def test_rechaza_hora_invalida(self) -> None:
        raiz = Path(__file__).resolve().parent.parent
        resultado = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(raiz / "scripts" / "configurar_tarea.ps1"),
                "-Carpeta",
                str(raiz),
                "-Hora",
                "99:99",
            ],
            capture_output=True,
            text=True,
            errors="replace",
        )
        assert resultado.returncode != 0
        assert "formato HH:MM" in resultado.stdout

    @pytest.mark.skipif(
        os.environ.get("AVB_PROBAR_PROGRAMADOR") != "1",
        reason="Prueba de integracion optativa que registra una tarea temporal",
    )
    def test_registra_y_verifica_tarea_temporal(self) -> None:
        raiz = Path(__file__).resolve().parent.parent
        nombre = f"PruebaLicitaciones_{uuid.uuid4().hex}"
        contenido = (raiz / "scripts" / "configurar_tarea.ps1").read_text(
            encoding="utf-8"
        )
        contenido = contenido.replace(
            "Actualizar Licitaciones Codelco desde Ariba", nombre
        )
        hora = (datetime.now() + timedelta(hours=2)).strftime("%H:%M")
        with tempfile.TemporaryDirectory() as temporal:
            script = Path(temporal) / "configurar_tarea.ps1"
            script.write_text(contenido, encoding="utf-8")
            carpeta = Path(temporal) / "Proyecto con espacios & caracteres"
            ejecutable = carpeta / ".venv" / "Scripts" / "python.exe"
            ejecutable.parent.mkdir(parents=True)
            ejecutable.touch()
            (carpeta / "principal.py").touch()
            try:
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
                        "-Hora",
                        hora,
                    ],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                )
                if resultado.returncode:
                    diagnostico = subprocess.run(
                        [
                            "powershell.exe",
                            "-NoProfile",
                            "-Command",
                            f"$t=Get-ScheduledTask -TaskName '{nombre}'; $t.State; $t.Actions | Select-Object Execute,Arguments,WorkingDirectory | Format-List; $t.Principal | Select-Object UserId,LogonType | Format-List; $t.Triggers | Select-Object Enabled,StartBoundary,DaysOfWeek,WeeksInterval | Format-List",
                        ],
                        capture_output=True,
                        text=True,
                        errors="replace",
                    )
                    pytest.fail(
                        resultado.stdout
                        + resultado.stderr
                        + diagnostico.stdout
                        + diagnostico.stderr
                    )
                consulta = subprocess.run(
                    ["schtasks", "/query", "/tn", nombre],
                    capture_output=True,
                    text=True,
                )
                assert consulta.returncode == 0, consulta.stdout + consulta.stderr
            finally:
                subprocess.run(
                    ["schtasks", "/delete", "/tn", nombre, "/f"],
                    capture_output=True,
                )
