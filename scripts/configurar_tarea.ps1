param(
    [Parameter(Mandatory = $true)][string]$Carpeta,
    [Parameter(Mandatory = $true)][string]$Hora
)

$ErrorActionPreference = 'Stop'
$nombre = 'Actualizar Licitaciones Codelco desde Ariba'

try {
    if ($Hora -notmatch '^([01][0-9]|2[0-3]):[0-5][0-9]$') {
        throw 'La hora debe tener formato HH:MM de 00:00 a 23:59.'
    }
    $carpetaProyecto = [System.IO.Path]::GetFullPath($Carpeta)
    if ($carpetaProyecto.Length -gt [System.IO.Path]::GetPathRoot($carpetaProyecto).Length) {
        $carpetaProyecto = $carpetaProyecto.TrimEnd('\')
    }
    $python = Join-Path $carpetaProyecto '.venv\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $python -PathType Leaf) -or
        -not (Test-Path -LiteralPath (Join-Path $carpetaProyecto 'principal.py') -PathType Leaf)) {
        throw 'Faltan Python del entorno virtual o principal.py en la carpeta indicada.'
    }
    $usuario = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $accion = New-ScheduledTaskAction -Execute $python -Argument 'principal.py --sin-pausa' -WorkingDirectory $carpetaProyecto
    $desencadenador = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At $Hora
    $principal = New-ScheduledTaskPrincipal -UserId $usuario -LogonType Interactive -RunLevel Limited
    $configuracion = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -WakeToRun -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 96 -RestartInterval (New-TimeSpan -Minutes 15)
    Register-ScheduledTask -TaskName $nombre -TaskPath '\' -Action $accion -Trigger $desencadenador -Principal $principal -Settings $configuracion -Force | Out-Null
    $tarea = Get-ScheduledTask -TaskName $nombre -TaskPath '\'
    if ($tarea.Principal.UserId -match '^S-\d-') {
        $usuarioRegistrado = $tarea.Principal.UserId
    } else {
        $usuarioRegistrado = ([System.Security.Principal.NTAccount]::new($tarea.Principal.UserId)).Translate([System.Security.Principal.SecurityIdentifier]).Value
    }
    if ($tarea.State -eq 'Disabled' -or $tarea.Actions.Count -ne 1 -or
        $tarea.Actions[0].Execute -ne $python -or
        $tarea.Actions[0].Arguments -ne 'principal.py --sin-pausa' -or
        $tarea.Actions[0].WorkingDirectory -ne $carpetaProyecto -or
        $usuarioRegistrado -ne $usuario -or
        $tarea.Principal.LogonType -ne 'Interactive' -or
        $tarea.Triggers.Count -ne 1 -or -not $tarea.Triggers[0].Enabled -or
        ([datetime]$tarea.Triggers[0].StartBoundary).ToString('HH:mm') -ne $Hora -or
        $tarea.Triggers[0].DaysOfWeek -ne 62 -or $tarea.Triggers[0].WeeksInterval -ne 1) {
        throw 'La tarea registrada no coincide con la configuracion solicitada. Revise el Programador de tareas.'
    }
    Write-Host "Tarea creada y verificada en la Biblioteca del Programador de tareas: $nombre"
    Write-Host "Usuario: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name). Hora: $Hora, lunes a viernes."
    Write-Host 'La tarea requiere que este usuario haya iniciado sesion en Windows.'
    exit 0
} catch {
    Write-Host "ERROR al configurar o verificar la tarea: $($_.Exception.Message)"
    Write-Host 'No se ha confirmado su configuracion. Revise la Biblioteca del Programador de tareas.'
    exit 1
}