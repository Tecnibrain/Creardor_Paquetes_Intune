<#
.SYNOPSIS
    Revierte Silenciar-Sonido.ps1: vuelve a activar el sonido del equipo.

.DESCRIPTION
    Rehabilita el servicio «Audio de Windows» (Audiosrv) si estaba deshabilitado, lo inicia
    y quita el silencio de todas las salidas de audio con el volumen indicado.

.PARAMETER Volumen
    Volumen de 0 a 100 que se deja en cada salida. Por defecto 50.
#>
[CmdletBinding()]
param(
    [ValidateRange(0, 100)]
    [int]$Volumen = 50
)

$ErrorActionPreference = 'Stop'

try {
    $servicio = Get-Service -Name 'Audiosrv' -ErrorAction SilentlyContinue
    if ($servicio) {
        $modo = (Get-CimInstance -ClassName Win32_Service -Filter "Name='Audiosrv'").StartMode
        if ($modo -eq 'Disabled') {
            Set-Service -Name 'Audiosrv' -StartupType Automatic
            Write-Output 'Servicio Audio de Windows rehabilitado (inicio automático).'
        }
        if ((Get-Service -Name 'Audiosrv').Status -ne 'Running') {
            Start-Service -Name 'Audiosrv'
            (Get-Service -Name 'Audiosrv').WaitForStatus('Running', [TimeSpan]::FromSeconds(30))
            Start-Sleep -Seconds 3   # dar tiempo a que aparezcan las salidas de audio
        }
    }

    # Reutiliza el código de Core Audio de Silenciar-Sonido.ps1 (misma carpeta del paquete).
    $principal = Join-Path -Path $PSScriptRoot -ChildPath 'Silenciar-Sonido.ps1'
    $texto = Get-Content -LiteralPath $principal -Raw -Encoding UTF8
    if ($texto -notmatch "(?s)\`$codigo = @'\r?\n(.*?)\r?\n'@") { throw 'No se encontró el código de audio en Silenciar-Sonido.ps1.' }
    if (-not ('FabricaAudio.Salidas' -as [type])) { Add-Type -TypeDefinition $Matches[1] -Language CSharp }

    $n = [FabricaAudio.Salidas]::Silenciar($false, [float]($Volumen / 100))
    Write-Output "Salidas de audio activadas: $n (volumen $Volumen %)"
    exit 0
}
catch {
    Write-Output "Error al activar el sonido: $($_.Exception.Message)"
    exit 1
}
