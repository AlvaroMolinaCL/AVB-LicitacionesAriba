param([Parameter(Mandatory = $true)][string]$Carpeta)

$ErrorActionPreference = 'Stop'

function Es-PythonCompatible([string]$Ruta) {
    if (-not (Test-Path -LiteralPath $Ruta -PathType Leaf)) { return $false }
    try {
        & $Ruta -c 'import sys; sys.exit(0 if (3, 10) <= sys.version_info[:2] < (4, 0) else 1)' 2>$null
        return $LASTEXITCODE -eq 0
    } catch { return $false }
}

function Buscar-PythonCompatible {
    $candidatos = [System.Collections.Generic.List[string]]::new()
    $lanzador = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($lanzador) {
        try {
            $listado = & $lanzador.Source -0p 2>$null
            foreach ($linea in $listado) {
                if ($linea -match '([A-Za-z]:\\.*?python(?:w)?\.exe)\s*$') {
                    $candidatos.Add($Matches[1])
                }
            }
        } catch { Write-Host 'No se pudo consultar el lanzador de Python; se revisaran otras ubicaciones.' }
    }
    foreach ($registro in @('HKCU:\Software\Python\PythonCore\*\InstallPath', 'HKLM:\Software\Python\PythonCore\*\InstallPath', 'HKLM:\Software\WOW6432Node\Python\PythonCore\*\InstallPath')) {
        foreach ($clave in @(Get-ItemProperty -Path $registro -ErrorAction SilentlyContinue)) {
            if ($clave.ExecutablePath) { $candidatos.Add($clave.ExecutablePath) }
            if ($clave.'(default)') { $candidatos.Add((Join-Path $clave.'(default)' 'python.exe')) }
        }
    }
    foreach ($comando in @(Get-Command python.exe -All -ErrorAction SilentlyContinue)) {
        if ($comando.Source -and $comando.Source -notmatch '\\Microsoft\\WindowsApps\\') {
            $candidatos.Add($comando.Source)
        }
    }
    $candidatos.Add((Join-Path $env:LOCALAPPDATA 'Programs\Python\AVB-Python314\python.exe'))
    foreach ($candidato in $candidatos) {
        if (Es-PythonCompatible $candidato) { return $candidato }
    }
    return $null
}

try {
    $carpetaProyecto = [System.IO.Path]::GetFullPath($Carpeta)
    $entorno = Join-Path $carpetaProyecto '.venv'
    $pythonEntorno = Join-Path $entorno 'Scripts\python.exe'
    if (Test-Path -LiteralPath $entorno) {
        if (-not (Es-PythonCompatible $pythonEntorno)) {
            throw 'El entorno existente no funciona o usa Python anterior a 3.10. Use la opcion 12 del panel para repararlo.'
        }
        Write-Host 'Se usara el entorno virtual existente.'
        exit 0
    }

    $python = Buscar-PythonCompatible
    if (-not $python) {
        if (-not [Environment]::Is64BitOperatingSystem) {
            throw 'No se encontro Python 3.10 o posterior. En Windows de 32 bits instale manualmente una version compatible para este usuario.'
        }
        Write-Host 'No se encontro Python compatible. Se instalara Python 3.14 solo para este usuario, sin permisos de administrador.'
        $instalador = Join-Path ([System.IO.Path]::GetTempPath()) ("python-avb-{0}.exe" -f [guid]::NewGuid())
        $destino = Join-Path $env:LOCALAPPDATA 'Programs\Python\AVB-Python314'
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -UseBasicParsing -Uri 'https://www.python.org/ftp/python/3.14.10/python-3.14.10-amd64.exe' -OutFile $instalador
            if ((Get-FileHash -LiteralPath $instalador -Algorithm SHA256).Hash -ne '92480239264744FE0386A1BD8275786FEFA126E3129E56A68E22B52C73329965') {
                throw 'No coincide la suma de seguridad del instalador de Python.'
            }
            $argumentos = @('/quiet', 'InstallAllUsers=0', 'PrependPath=0', 'Include_launcher=0', 'Include_test=0', "TargetDir=`"$destino`"")
            $proceso = Start-Process -FilePath $instalador -ArgumentList $argumentos -Wait -PassThru
            if ($proceso.ExitCode -ne 0) { throw "El instalador de Python devolvio el codigo $($proceso.ExitCode)." }
        } finally {
            if (Test-Path -LiteralPath $instalador) { Remove-Item -LiteralPath $instalador -Force -ErrorAction SilentlyContinue }
        }
        $python = Join-Path $destino 'python.exe'
        if (-not (Es-PythonCompatible $python)) { throw 'No se pudo verificar el Python instalado para el usuario.' }
    }
    Write-Host "Python compatible: $python"
    & $python -m venv $entorno
    if ($LASTEXITCODE -ne 0 -or -not (Es-PythonCompatible $pythonEntorno)) {
        throw 'No se pudo crear o verificar el entorno virtual. Revise permisos de escritura y use la opcion 12 para repararlo.'
    }
    exit 0
} catch {
    Write-Host "ERROR al preparar Python: $($_.Exception.Message)"
    Write-Host 'Compruebe la conexion, los permisos de la carpeta y las politicas de su empresa; consulte a soporte si es necesario.'
    exit 1
}